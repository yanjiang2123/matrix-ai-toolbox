#!/usr/bin/env python3
"""Excel / CSV 两文件比对引擎（纯 Python，不依赖 pandas）

比 webapp 那版多的一件事：主键不唯一时做「基于内容的最优配对」。
主键重复在业务明细表里很常见（同一订单包含多条子项），此时两侧各有 N、M 行，
按出现顺序硬配会造出一堆假差异——同一批数据只要行序不同就全红。
这里先用整行哈希配掉完全相同记录，小重复组再用匈牙利算法给出「总差异最小」
的配对；极大重复组使用有界、可告警的确定性配对，避免 JOIN 倍乘数据拖垮进程。

数据形态沿用工具箱其他模块的 headers/rows，读文件复用 matrix_core 的 openpyxl 通道。
"""

from __future__ import annotations

import csv
import io
import re
from collections import Counter, defaultdict, deque
from pathlib import Path

# 这些写法都当空值看待，否则「NULL」与空单元格会被判成差异
NULL_LIKE = {"(null)", "null", "none", "n/a", "na", "nan", "<null>", "#n/a", ""}
# 组内行数超过这个数就退化成贪心：匈牙利是 O(S³)，200 行已经上万次内循环
HUNGARIAN_LIMIT = 200
MAX_ROWS = 500_000            # 单文件行数上限，再多会吃满内存
MAX_EXCEL_DIFF_ROWS = 50_000  # 报告里差异行的上限，超出只在页面上看
CSV_ENCODINGS = ("utf-8-sig", "utf-8", "gbk", "gb2312", "gb18030", "latin-1")

DEFAULT_OPTIONS = {
    "case_sensitive": False,     # 区分大小写
    "trim_whitespace": True,     # 去首尾空格
    "null_equals_empty": True,   # NULL 与空值视为相同
    "strict_numeric": False,     # 严格比数值文本（关掉则 125.00 == 125）
}


# ── 读文件 ──────────────────────────────────────────────────────

def _decode(raw: bytes) -> str:
    """按常见编码依次试。国产系统导出的 CSV 多是 GBK 系，写死 utf-8 会直接乱码"""
    for enc in CSV_ENCODINGS:
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace")


def _read_delimited(path: Path, sep: str = "") -> tuple[list[str], list[list]]:
    text = _decode(Path(path).read_bytes())
    if not sep:
        # 制表符比逗号多就按 TSV 读：从数据库直接复制出来的多是制表符分隔
        head = text.splitlines()[:10]
        tabs = sum(l.count("\t") for l in head if l.strip())
        commas = sum(l.count(",") for l in head if l.strip())
        sep = "\t" if tabs > commas else ","
    rdr = csv.reader(io.StringIO(text), delimiter=sep)
    rows = [r for r in rdr]
    if not rows:
        return [], []
    headers = [str(h).strip() for h in rows[0]]
    return headers, [list(r) for r in rows[1:]]


def read_table(path: Path, sheet=None) -> tuple[list[str], list[list]]:
    """读一个文件的表头与数据行。支持 xlsx/xlsm/csv/txt/tsv。"""
    import matrix_core as core

    p = Path(path)
    ext = p.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        headers, rows = core.read_sheet(p, sheet)
        return [str(h).strip() for h in headers], rows
    if ext in (".csv",):
        return _read_delimited(p, ",")
    if ext in (".txt", ".tsv"):
        return _read_delimited(p)
    raise ValueError(
        f"不支持的文件类型：{ext}。支持 xlsx / xlsm / csv / txt / tsv；"
        "旧版 .xls 请先另存为 .xlsx"
    )


# ── 归一化 ──────────────────────────────────────────────────────

_NUM_RE = re.compile(r"^-?\d+(?:\.\d+)?$")


def _trim_number(s: str) -> str:
    """去掉小数点后多余的零：125.00000000 → 125、362948.98000000 → 362948.98

    只对带小数点的动手。整数 100 直接 rstrip('0') 会变成 1，那是灾难。
    """
    if "." not in s or not _NUM_RE.match(s):
        return s
    out = s.rstrip("0").rstrip(".")
    return out if out not in ("", "-") else "0"


def make_normalizer(options: dict):
    """按选项生成单元格归一化函数"""
    trim = options.get("trim_whitespace", True)
    fold = not options.get("case_sensitive", False)
    null_eq = options.get("null_equals_empty", True)
    loose_num = not options.get("strict_numeric", False)

    def norm(v) -> str:
        s = "" if v is None else str(v)
        if trim:
            s = s.strip()
        if fold:
            s = s.lower()
        if null_eq and s.lower() in NULL_LIKE:
            return ""
        if loose_num:
            s = _trim_number(s)
        return s

    return norm


# ── 匈牙利最优配对 ────────────────────────────────────────────────

def hungarian(cost: list[list[int]]) -> list[int]:
    """最小代价配对。cost[i][j] = A 第 i 行与 B 第 j 行有几个字段不同。

    返回 assign[i] = j（配到 B 的第 j 行），配不上记 -1。
    非方阵补成方阵，补出来的格子给一个大到不可能被选中的代价。
    """
    n = len(cost)
    m = len(cost[0]) if n else 0
    if not n or not m:
        return [-1] * n
    size = max(n, m)
    big = max(max(r) for r in cost) * 2 + 100
    sq = [[big] * size for _ in range(size)]
    for i in range(n):
        row = cost[i]
        for j in range(m):
            sq[i][j] = row[j]

    inf = float("inf")
    u = [0] * (size + 1)
    v = [0] * (size + 1)
    p = [0] * (size + 1)
    way = [0] * (size + 1)
    for i in range(1, size + 1):
        p[0] = i
        j0 = 0
        minv = [inf] * (size + 1)
        used = [False] * (size + 1)
        while True:
            used[j0] = True
            i0 = p[j0]
            delta, j1 = inf, 0
            srow = sq[i0 - 1]
            for j in range(1, size + 1):
                if used[j]:
                    continue
                cur = srow[j - 1] - u[i0] - v[j]
                if cur < minv[j]:
                    minv[j], way[j] = cur, j0
                if minv[j] < delta:
                    delta, j1 = minv[j], j
            for j in range(size + 1):
                if used[j]:
                    u[p[j]] += delta
                    v[j] -= delta
                else:
                    minv[j] -= delta
            j0 = j1
            if p[j0] == 0:
                break
        while True:
            j1 = way[j0]
            p[j0] = p[j1]
            j0 = j1
            if j0 == 0:
                break

    assign = [-1] * n
    for j in range(1, size + 1):
        r, c = p[j], j - 1
        if r != 0 and r <= n and c < m:
            assign[r - 1] = c
    return assign


def greedy_assign(cost: list[list[int]]) -> list[int]:
    """组内行数太多时的退路：差异最小的行先挑走"""
    n = len(cost)
    m = len(cost[0]) if n else 0
    assign, taken = [-1] * n, set()
    for i in sorted(range(n), key=lambda r: min(cost[r])):
        best, bj = None, -1
        for j in range(m):
            if j in taken:
                continue
            if best is None or cost[i][j] < best:
                best, bj = cost[i][j], j
        if bj >= 0:
            assign[i] = bj
            taken.add(bj)
    return assign


def match_group(na: list[list[str]], nb: list[list[str]],
                meta: dict | None = None) -> tuple[list, list, list]:
    """一个重复主键组内部的配对。na/nb 是两侧已归一化的比较列值。

    返回 (配上的下标对, A 侧没配上的下标, B 侧没配上的下标)。
    """
    if meta is not None:
        meta["approximate"] = False
    if not na or not nb:
        return [], list(range(len(na))), list(range(len(nb)))

    def token(row):
        values = tuple(row)
        try:
            hash(values)
            return values
        except TypeError:
            return tuple(repr(v) for v in row)

    # JOIN 倍乘时同一主键可能出现几千行。先按整行哈希配掉完全相同的内容，
    # 既是最优配对的一部分，也避免先构造 n×m 代价矩阵耗尽内存。
    buckets = defaultdict(deque)
    for j, row in enumerate(nb):
        buckets[token(row)].append(j)
    exact_pairs: list[tuple[int, int]] = []
    remaining_a: list[int] = []
    used_b: set[int] = set()
    for i, row in enumerate(na):
        matches = buckets.get(token(row))
        if matches:
            j = matches.popleft()
            exact_pairs.append((i, j))
            used_b.add(j)
        else:
            remaining_a.append(i)
    remaining_b = [j for j in range(len(nb)) if j not in used_b]
    if not remaining_a or not remaining_b:
        return exact_pairs, remaining_a, remaining_b

    if max(len(remaining_a), len(remaining_b)) <= HUNGARIAN_LIMIT:
        cost = [[sum(1 for x, y in zip(na[i], nb[j]) if x != y)
                 for j in remaining_b] for i in remaining_a]
        assign = hungarian(cost)
        residual_pairs = [(remaining_a[i], remaining_b[j])
                          for i, j in enumerate(assign) if j != -1]
        unmatched_a = [remaining_a[i] for i, j in enumerate(assign) if j == -1]
        paired_b = {j for _, j in residual_pairs}
        unmatched_b = [j for j in remaining_b if j not in paired_b]
    else:
        # 超大残余组本来就不是可靠主键。采用确定性的排序配对，保持 O(n log n)
        # 内存/时间上界；页面会同时报告重复主键，提示用户补充更细的业务键。
        ordered_a = sorted(remaining_a, key=lambda i: repr(token(na[i])))
        ordered_b = sorted(remaining_b, key=lambda j: repr(token(nb[j])))
        if meta is not None:
            meta["approximate"] = True
        pair_count = min(len(ordered_a), len(ordered_b))
        residual_pairs = list(zip(ordered_a[:pair_count], ordered_b[:pair_count]))
        unmatched_a = ordered_a[pair_count:]
        unmatched_b = ordered_b[pair_count:]
    return exact_pairs + residual_pairs, unmatched_a, unmatched_b


# ── 主流程 ──────────────────────────────────────────────────────

def _col_index(headers: list[str]) -> dict:
    return {h: i for i, h in enumerate(headers)}


def _cell(row: list, i: int) -> str:
    v = row[i] if i < len(row) else None
    return "" if v is None else str(v)


def common_columns(ha: list[str], hb: list[str]) -> list[str]:
    """两边同名的列，按 A 侧顺序。列名不同的部分没法逐格比，只能忽略"""
    sb = set(hb)
    return [h for h in ha if h in sb]


# 名字像主键的
KEY_HINT = ("id", "no", "code", "key", "编号", "主键", "单号", "流水", "号")
# 名字像度量的：这类列唯一率再高也不是主键，金额几乎不重复但拿它对齐纯属灾难
MEASURE_HINT = ("amount", "amt", "qty", "quantity", "num", "count", "sum", "total",
                "price", "fee", "cost", "rate", "金额", "数量", "价", "费", "率")
_DECIMAL_RE = re.compile(r"^-?\d+\.\d+$")


def _name_hit(col: str, hints) -> bool:
    """列名是否命中提示词。按整名或下划线分段匹配，避免 note 被当成 no"""
    low = col.lower()
    parts = re.split(r"[_\s\-]+", low)
    return any(low == h or h in parts or (len(h) > 3 and h in low) or
               (not h.isascii() and h in col) for h in hints)


def guess_keys(ha: list[str], ra: list[list], hb: list[str], rb: list[list],
               common: list[str], limit: int = 4) -> list[dict]:
    """猜主键：在两侧都尽量唯一、不为空、且名字不像度量列的排前面。

    选错主键是这类比对最常见的返工原因。光看唯一率会推荐「金额」——
    金额几乎不重复，唯一率能到 90% 以上，但拿它对齐两张表毫无意义。
    所以名字像 id/编号 的加分，像金额/数量的扣分，纯小数列直接判为度量。
    """
    ia, ib = _col_index(ha), _col_index(hb)
    out = []
    for pos, c in enumerate(common):
        va = [_cell(r, ia[c]).strip() for r in ra]
        vb = [_cell(r, ib[c]).strip() for r in rb]
        if not va or not vb:
            continue
        blank = sum(1 for v in va if not v) + sum(1 for v in vb if not v)
        rate = min(len(set(va)) / len(va), len(set(vb)) / len(vb))
        # 大半是带小数点的数字 → 这是度量列，不是标识
        dec = sum(1 for v in va[:500] if _DECIMAL_RE.match(v))
        looks_measure = (dec > len(va[:500]) * 0.5
                         or _name_hit(c, MEASURE_HINT))
        score = rate
        if _name_hit(c, KEY_HINT):
            score += 0.35
        if looks_measure:
            score -= 0.8
        score += max(0.0, 0.05 - pos * 0.01)     # 靠左的列更可能是主键
        out.append({"column": c, "unique_rate": round(rate, 4),
                    "blank": blank, "score": round(score, 4),
                    "measure": looks_measure,
                    "unique": rate >= 0.999 and blank == 0})
    out.sort(key=lambda x: (-x["score"], x["blank"]))
    return out[:limit]


def compare_tables(ha: list[str], ra: list[list],
                   hb: list[str], rb: list[list],
                   keys: list[str], options: dict | None = None) -> dict:
    """按主键比两张表。主键可以重复，重复组内做最优配对。"""
    opts = {**DEFAULT_OPTIONS, **(options or {})}
    if not keys:
        raise ValueError("请至少选一个主键列，否则两边的行没法对齐")
    if len(ra) > MAX_ROWS or len(rb) > MAX_ROWS:
        raise ValueError(f"单文件最多 {MAX_ROWS:,} 行"
                         f"（A {len(ra):,} 行、B {len(rb):,} 行）")

    def duplicate_headers(headers):
        counts = Counter(str(h).strip().lower() for h in headers)
        return [str(h) for h in headers
                if counts[str(h).strip().lower()] > 1]

    dup_a = list(dict.fromkeys(duplicate_headers(ha)))
    dup_b = list(dict.fromkeys(duplicate_headers(hb)))
    if dup_a or dup_b:
        raise ValueError(
            "文件表头必须唯一，否则主键和字段值无法可靠定位："
            f"A 侧重复 {dup_a or '无'}；B 侧重复 {dup_b or '无'}。"
            "请先修改重复表头后重试")
    ia, ib = _col_index(ha), _col_index(hb)
    missing = ([f"A 侧缺 {k}" for k in keys if k not in ia] +
               [f"B 侧缺 {k}" for k in keys if k not in ib])
    if missing:
        raise ValueError("主键列两边都得有：" + "；".join(missing))

    common = common_columns(ha, hb)
    cmp_cols = [c for c in common if c not in keys]
    norm = make_normalizer(opts)
    ka = [ia[k] for k in keys]
    kb = [ib[k] for k in keys]
    ca = [ia[c] for c in cmp_cols]
    cb = [ib[c] for c in cmp_cols]

    # 按主键分组。原始行留着出报告，归一化值只用于判等
    # disp 记每组主键的原始写法：分组必须用归一化值（否则 V001/v001 会分成两组），
    # 但报告里给用户看的得是他文件里的原样，不然拿着小写的主键回去搜不到
    ga, gb = defaultdict(list), defaultdict(list)
    disp: dict[tuple, list] = {}
    for n, row in enumerate(ra):
        k = tuple(norm(_cell(row, i)) for i in ka)
        ga[k].append(n)
        disp.setdefault(k, [_cell(row, i) for i in ka])
    for n, row in enumerate(rb):
        k = tuple(norm(_cell(row, i)) for i in kb)
        gb[k].append(n)
        disp.setdefault(k, [_cell(row, i) for i in kb])

    na = {n: [norm(_cell(ra[n], i)) for i in ca] for n in range(len(ra))}
    nb = {n: [norm(_cell(rb[n], i)) for i in cb] for n in range(len(rb))}

    diffs, only_a, only_b = [], [], []
    matched = uniq_keys = dup_keys = approximate_groups = 0
    dup_unmatched_a = dup_unmatched_b = 0
    dup_rows_a = dup_rows_b = 0

    def keyd(k):
        return dict(zip(keys, disp.get(k, list(k))))

    def row_out(side: str, n: int, k, kind: str) -> dict:
        src, idx = (ra, ca) if side == "a" else (rb, cb)
        d = {"type": kind, "key": keyd(k), "row": n + 2}   # +2：表头占一行
        for c, i in zip(cmp_cols, idx):
            d[f"val_{c}"] = _cell(src[n], i)
        return d

    for k, ias in ga.items():
        ibs = gb.get(k)
        if ibs is None:
            for n in ias:
                only_a.append(row_out("a", n, k, "only_in_a"))
            continue
        if len(ias) == 1 and len(ibs) == 1:
            uniq_keys += 1
            pairs = [(0, 0)]
            ua, ub = [], []
        else:
            dup_keys += 1
            dup_rows_a += len(ias)
            dup_rows_b += len(ibs)
            pair_meta: dict = {}
            pairs, ua, ub = match_group(
                [na[n] for n in ias], [nb[n] for n in ibs], pair_meta)
            approximate_groups += int(bool(pair_meta.get("approximate")))
            dup_unmatched_a += len(ua)
            dup_unmatched_b += len(ub)
        matched += len(pairs)
        for i, j in pairs:
            x, y = ias[i], ibs[j]
            for pos, col in enumerate(cmp_cols):
                if na[x][pos] == nb[y][pos]:
                    continue
                diffs.append({"type": "diff", "key": keyd(k), "column": col,
                              "value_a": _cell(ra[x], ca[pos]),
                              "value_b": _cell(rb[y], cb[pos]),
                              "row_a": x + 2, "row_b": y + 2})
        for i in ua:
            only_a.append(row_out("a", ias[i], k, "unmatched_a"))
        for j in ub:
            only_b.append(row_out("b", ibs[j], k, "unmatched_b"))

    for k, ibs in gb.items():
        if k in ga:
            continue
        for n in ibs:
            only_b.append(row_out("b", n, k, "only_in_b"))

    truly_a = sum(1 for d in only_a if d["type"] == "only_in_a")
    truly_b = sum(1 for d in only_b if d["type"] == "only_in_b")
    stats = {
        "rows_a": len(ra), "rows_b": len(rb),
        "cols_a": len(ha), "cols_b": len(hb),
        "keys": keys, "cmp_cols": cmp_cols,
        "common_cols": len(common),
        "only_col_a": [h for h in ha if h not in set(hb)],
        "only_col_b": [h for h in hb if h not in set(ha)],
        "unique_keys": uniq_keys, "dup_keys": dup_keys,
        "approximate_groups": approximate_groups,
        "dup_rows_a": dup_rows_a, "dup_rows_b": dup_rows_b,
        "matched": matched,
        "truly_only_a": truly_a, "truly_only_b": truly_b,
        "dup_unmatched_a": dup_unmatched_a, "dup_unmatched_b": dup_unmatched_b,
        "only_a": len(only_a), "only_b": len(only_b),
        "diff_cells": len(diffs),
    }
    return {"stats": stats, "diffs": diffs, "only_a": only_a, "only_b": only_b,
            "verdict": verdict(stats), "options": opts}


def verdict(s: dict) -> dict:
    """一句结论 + 下一步。跟 SQL 排查那边保持同一套话术"""
    row_diff = bool(s["only_a"] or s["only_b"] or s["diff_cells"])
    structure_diff = bool(s["only_col_a"] or s["only_col_b"])
    if not row_diff and structure_diff:
        return {
            "level": "diff",
            "text": (f"字段结构不一致：仅 A 有 {len(s['only_col_a'])} 列 / "
                     f"仅 B 有 {len(s['only_col_b'])} 列；共有字段当前值一致"),
            "next": "先统一字段清单和名称；单边字段没有参与值比较，不能判定整体一致",
        }
    if not row_diff and not s["cmp_cols"]:
        return {
            "level": "warn",
            "text": f"只对齐了 {s['matched']} 行主键，未比较任何非主键业务字段",
            "next": "至少保留一个金额、状态或时间等共有业务字段后再判断数据是否一致",
        }
    if not row_diff and s["dup_keys"]:
        return {
            "level": "warn",
            "text": (f"字段值当前一致，但有 {s['dup_keys']} 个重复主键组，"
                     + (f"其中 {s.get('approximate_groups', 0)} 组过大、采用有界近似配对"
                        if s.get("approximate_groups") else "已按内容配对")),
            "next": "确认重复是否符合业务粒度；若不是合法一对多，请补充子项序号等字段组成复合主键",
        }
    if not row_diff:
        return {"level": "ok",
                "text": f"两边 {s['matched']} 行完全一致",
                "next": f"比了 {len(s['cmp_cols'])} 个共有列。"
                        f"若两边列名不完全相同，只有同名列参与了对比"}
    bits = []
    if s["truly_only_a"]:
        bits.append(f"B 侧缺 {s['truly_only_a']} 行（主键在 B 侧压根没有）")
    if s["truly_only_b"]:
        bits.append(f"A 侧缺 {s['truly_only_b']} 行")
    if s["dup_unmatched_a"] or s["dup_unmatched_b"]:
        bits.append(f"重复主键组里配不上的：A {s['dup_unmatched_a']} 行 / "
                    f"B {s['dup_unmatched_b']} 行")
    if s["diff_cells"]:
        bits.append(f"字段值不同 {s['diff_cells']} 处")
    nxt = []
    if s["dup_keys"]:
        nxt.append(f"有 {s['dup_keys']} 个主键重复（A {s['dup_rows_a']} 行 / "
                   f"B {s['dup_rows_b']} 行），已按内容最优配对；"
                   f"主键选得更细一些能减少这类不确定")
    if s.get("approximate_groups"):
        nxt.append(f"其中 {s['approximate_groups']} 个超大重复组为避免卡死采用有界近似配对，"
                   "请补充更细主键后复核这些组")
    if s["only_col_a"] or s["only_col_b"]:
        nxt.append(f"列名不一致：只有 A 有 {len(s['only_col_a'])} 列、"
                   f"只有 B 有 {len(s['only_col_b'])} 列，这些列没参与对比")
    if not s["cmp_cols"]:
        nxt.append("两侧只有主键可比，不能据此判断其余业务字段")
    if s["diff_cells"]:
        nxt.append("字段值差异先看是不是格式问题（数值尾零、大小写、空格），"
                   "可以调上面的比对选项再跑一次")
    return {"level": "diff", "text": "发现差异：" + "；".join(bits),
            "next": "。".join(nxt) or "逐条核对差异明细"}


# ── 导出报告 ────────────────────────────────────────────────────

def build_report(res: dict, name_a: str = "A", name_b: str = "B") -> dict:
    """拼出多 sheet 报告的内容，交给 core.write_xlsx 落盘"""
    s = res["stats"]
    v = res["verdict"]
    keys, cols = s["keys"], s["cmp_cols"]
    exported_diffs = min(len(res["diffs"]), MAX_EXCEL_DIFF_ROWS)
    exported_only_a = min(len(res["only_a"]), MAX_EXCEL_DIFF_ROWS)
    exported_only_b = min(len(res["only_b"]), MAX_EXCEL_DIFF_ROWS)
    truncated = (exported_diffs < len(res["diffs"])
                 or exported_only_a < len(res["only_a"])
                 or exported_only_b < len(res["only_b"]))
    summary = [["结论", v["text"]], ["下一步", v["next"]],
               ["A 文件", name_a], ["B 文件", name_b],
               ["主键列", "、".join(keys)],
               ["A 侧行数", s["rows_a"]], ["B 侧行数", s["rows_b"]],
               ["参与对比的列数", len(cols)],
               ["唯一主键（1:1 对上）", s["unique_keys"]],
               ["重复主键（按内容配对）", s["dup_keys"]],
               ["超大组近似配对", s.get("approximate_groups", 0)],
               ["配上的行对", s["matched"]],
               ["仅 A 有（主键缺失）", s["truly_only_a"]],
               ["仅 B 有（主键缺失）", s["truly_only_b"]],
               ["重复组内 A 未配上", s["dup_unmatched_a"]],
               ["重复组内 B 未配上", s["dup_unmatched_b"]],
               ["字段值差异数", s["diff_cells"]],
               ["字段差异明细（导出/总数）", f"{exported_diffs}/{len(res['diffs'])}"],
               ["仅 A 明细（导出/总数）", f"{exported_only_a}/{len(res['only_a'])}"],
               ["仅 B 明细（导出/总数）", f"{exported_only_b}/{len(res['only_b'])}"],
               ["导出完整性", ("明细已按安全上限截断，统计总数仍为完整结果"
                              if truncated else "全部差异明细已导出")],
               ["只有 A 有的列", "、".join(s["only_col_a"]) or "无"],
               ["只有 B 有的列", "、".join(s["only_col_b"]) or "无"]]

    dh = keys + ["列名", f"{name_a} 的值", f"{name_b} 的值", "A 行号", "B 行号"]
    drows = [[d["key"].get(k, "") for k in keys] +
             [d["column"], d["value_a"], d["value_b"], d["row_a"], d["row_b"]]
             for d in res["diffs"][:MAX_EXCEL_DIFF_ROWS]]

    def side_sheet(items):
        h = keys + ["情况", "行号"] + cols
        rows = []
        for d in items[:MAX_EXCEL_DIFF_ROWS]:
            kind = ("主键在对侧不存在" if d["type"].startswith("only")
                    else "重复主键组内没配上")
            rows.append([d["key"].get(k, "") for k in keys] +
                        [kind, d["row"]] + [d.get(f"val_{c}", "") for c in cols])
        return h, rows

    ha, rows_a = side_sheet(res["only_a"])
    hb, rows_b = side_sheet(res["only_b"])
    return {"结论": (["项目", "内容"], summary),
            "字段值差异": (dh, drows),
            # 文件显示名可能完全相同；固定页名可避免字典键碰撞导致一侧数据丢失。
            "仅A有": (ha, rows_a),
            "仅B有": (hb, rows_b)}
