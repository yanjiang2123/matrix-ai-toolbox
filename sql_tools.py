#!/usr/bin/env python3
"""Matrix 工具箱 · SQL 排查工具层

把常见的数据核对与链路排查动作固化成可复用能力：

- parse_sql()        把 SQL 拆成结构：表 / JOIN / WHERE / SELECT / GROUP BY / CTE
- compare_logic()    第一层比较：只看逻辑，标出 JOIN 类型变更、限定条件增减等风险
- split_sql()        长 SQL 按 CTE / UNION / 子查询拆成可读片段
- find_time_columns() 识别业务时间字段（Matrix 单次只回 500 行，靠时间切片取数）
- build_slice_sql()  按业务时间范围包装取数 SQL
- compare_details()  第二层比较：拉两边明细按主键对齐，输出差异与排查日志
- table_lineage()    列出依赖表并按数仓分层归类，指向该查哪一层的脚本

刻意不引第三方 SQL parser：Matrix/StarRocks 方言与 @变量、中文表名会让通用
parser 频繁报错，这里用括号配对 + 正则做到「够用且不会崩」。
"""

from __future__ import annotations

import re
from collections import OrderedDict

# ── 通用数仓分层：库名前缀 → (层级, 出问题时该追查什么) ──────────
LAYERS = OrderedDict([
    ("raw", ("RAW 原始层", "检查源系统采集、文件到达和字段解析")),
    ("stg", ("STG 暂存层", "检查采集任务、清洗规则和增量条件")),
    ("ods", ("ODS 明细层", "检查源到目标同步、去重和标准化规则")),
    ("dwd", ("DWD 明细层", "检查明细加工、主键和状态过滤")),
    ("dws", ("DWS 汇总层", "检查聚合粒度、JOIN 和指标口径")),
    ("dim", ("DIM 维度层", "检查维度同步、版本和有效期")),
    ("ads", ("ADS 应用层", "检查面向应用的加工与过滤脚本")),
    ("tmp", ("临时层", "检查临时表是否过期或被重复使用")),
])

# 业务时间字段候选：命中即认为可用于时间切片
TIME_HINTS = ("date", "time", "_dt", "day", "month", "year",
              "日期", "时间", "时点")
# 这些是系统时间，不是业务时间，排在候选末尾
SYS_TIME = ("sys_load_time", "sys_update_time", "sys_create_time",
            "create_time", "update_time", "etl_time", "load_time")
MAX_ARCHIVE_SCOPE_KEYS = 20_000  # 避免超大结果把后台任务响应撑到数十 MB

JOIN_RE = re.compile(
    r"\b(LEFT\s+OUTER\s+JOIN|RIGHT\s+OUTER\s+JOIN|FULL\s+OUTER\s+JOIN|"
    r"LEFT\s+JOIN|RIGHT\s+JOIN|INNER\s+JOIN|CROSS\s+JOIN|FULL\s+JOIN|JOIN)\b",
    re.I)
# 库.表 或 表，允许反引号与中文
IDENT = r"`?[\w\u4e00-\u9fff]+`?(?:\.`?[\w\u4e00-\u9fff]+`?)?"


# ══════════════════════════════════════════════════════════════
# 预处理
# ══════════════════════════════════════════════════════════════


def strip_comments(sql: str) -> str:
    """去掉 -- 行注释与 /* */ 块注释，但保留字符串字面量里的内容"""
    out, i, n = [], 0, len(sql)
    while i < n:
        ch = sql[i]
        if ch in "'\"":                      # 字符串原样保留
            quote = ch
            out.append(ch)
            i += 1
            while i < n:
                out.append(sql[i])
                if sql[i] == quote and sql[i - 1] != "\\":
                    i += 1
                    break
                i += 1
            continue
        if sql.startswith("--", i):
            while i < n and sql[i] != "\n":
                i += 1
            continue
        if sql.startswith("/*", i):
            end = sql.find("*/", i + 2)
            i = n if end < 0 else end + 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def normalize(sql: str) -> str:
    """归一化：去注释、压空白、关键字大写，用于逻辑层比较"""
    s = strip_comments(sql)
    s = re.sub(r"\s+", " ", s).strip().rstrip(";")
    # 只大写独立的关键字，不动标识符与字符串
    kws = ("SELECT|FROM|WHERE|AND|OR|NOT|IN|IS|NULL|AS|ON|LEFT|RIGHT|INNER|"
           "FULL|OUTER|CROSS|JOIN|GROUP|ORDER|BY|HAVING|LIMIT|UNION|ALL|"
           "DISTINCT|CASE|WHEN|THEN|ELSE|END|WITH|INSERT|INTO|VALUES|"
           "BETWEEN|LIKE|EXISTS|DESC|ASC")
    return re.sub(rf"\b({kws})\b", lambda m: m.group(1).upper(), s, flags=re.I)


def split_top_level(text: str, sep_re: re.Pattern) -> list[str]:
    """在括号深度 0 处按 sep_re 切分，避免切到子查询内部"""
    parts, depth, last, i = [], 0, 0, 0
    n = len(text)
    while i < n:
        c = text[i]
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
        elif depth == 0:
            m = sep_re.match(text, i)
            if m:
                parts.append(text[last:i])
                last = i = m.end()
                continue
        i += 1
    parts.append(text[last:])
    return [p.strip() for p in parts if p.strip()]


def _find_kw(sql: str, kw_re: re.Pattern) -> int:
    """找关键字在括号深度 0 处的位置，找不到返回 -1"""
    depth, i, n = 0, 0, len(sql)
    while i < n:
        c = sql[i]
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
        elif depth == 0:
            m = kw_re.match(sql, i)
            if m:
                return i
        i += 1
    return -1


# ══════════════════════════════════════════════════════════════
# 结构解析
# ══════════════════════════════════════════════════════════════


def extract_ctes(sql: str) -> tuple[list[dict], str]:
    """剥离 WITH 子句，返回 ([{name, body}], 主查询)"""
    s = sql.lstrip()
    if not re.match(r"WITH\b", s, re.I):
        return [], sql
    i = len(re.match(r"WITH\s+", s, re.I).group(0))
    ctes = []
    while i < len(s):
        m = re.match(rf"({IDENT})\s+AS\s*\(", s[i:], re.I)
        if not m:
            break
        name = m.group(1).strip("`")
        start = i + m.end()          # 左括号之后
        depth, j = 1, start
        while j < len(s) and depth:
            if s[j] == "(":
                depth += 1
            elif s[j] == ")":
                depth -= 1
            j += 1
        ctes.append({"name": name, "body": s[start:j - 1].strip()})
        rest = s[j:].lstrip()
        if rest.startswith(","):
            i = len(s) - len(rest) + 1
            while i < len(s) and s[i] in " \t\n":
                i += 1
        else:
            return ctes, rest
    return ctes, s[i:]


def extract_tables(sql: str) -> list[dict]:
    """提取 FROM / JOIN 后的表，带别名。子查询记为 (子查询)"""
    tables, seen = [], set()
    for m in re.finditer(rf"\b(FROM|JOIN)\s+(\(|{IDENT})\s*(?:AS\s+)?(`?\w+`?)?",
                         sql, re.I):
        raw = m.group(2)
        if raw == "(":
            continue                       # 派生表，由 split_sql 单独呈现
        name = raw.strip("`")
        alias = (m.group(3) or "").strip("`")
        if alias.upper() in ("ON", "WHERE", "GROUP", "ORDER", "LEFT", "RIGHT",
                             "INNER", "JOIN", "UNION", "LIMIT", "HAVING", "AS"):
            alias = ""
        key = (name.lower(), alias.lower())
        if key in seen:
            continue
        seen.add(key)
        tables.append({"name": name, "alias": alias,
                       "db": name.split(".")[0] if "." in name else ""})
    return tables


def _skip_parens(sql: str, i: int) -> int:
    """i 指向 '('，返回其配对 ')' 之后的位置"""
    depth, n = 0, len(sql)
    while i < n:
        if sql[i] == "(":
            depth += 1
        elif sql[i] == ")":
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1
    return n


def extract_joins(sql: str) -> list[dict]:
    """提取 JOIN 类型、右表、ON 条件。

    右表是子查询时必须整体跳过括号再找 ON，否则会错取到子查询内部的 ON。
    """
    joins = []
    for m in JOIN_RE.finditer(sql):
        jtype = re.sub(r"\s+", " ", m.group(1).upper())
        jtype = jtype.replace(" OUTER", "")          # LEFT OUTER JOIN → LEFT JOIN
        pos = m.end()
        while pos < len(sql) and sql[pos] in " \t\n":
            pos += 1
        if pos < len(sql) and sql[pos] == "(":
            table = "(子查询)"
            after = _skip_parens(sql, pos)           # 跳过整个派生表
            am = re.match(r"\s*(?:AS\s+)?(`?\w+`?)", sql[after:], re.I)
            if am and am.group(1).upper() != "ON":
                table = f"(子查询 {am.group(1).strip('`')})"
                after += am.end()
        else:
            tm = re.match(rf"({IDENT})\s*(?:AS\s+)?(`?\w+`?)?", sql[pos:], re.I)
            table = tm.group(1).strip("`") if tm else "?"
            after = pos + (tm.end() if tm else 0)
        # 从右表之后开始找 ON，且只取本层（深度 0）的 ON
        tail = sql[after:]
        on = ""
        if re.match(r"\s*\bON\b", tail, re.I):
            body_start = re.match(r"\s*\bON\b", tail, re.I).end()
            depth, i = 0, body_start
            stop = re.compile(r"\b(LEFT|RIGHT|INNER|FULL|CROSS|JOIN|WHERE|"
                              r"GROUP\s+BY|ORDER\s+BY|HAVING|LIMIT|UNION)\b", re.I)
            while i < len(tail):
                c = tail[i]
                if c == "(":
                    depth += 1
                elif c == ")":
                    if depth == 0:
                        break
                    depth -= 1
                elif depth == 0 and stop.match(tail, i):
                    break
                i += 1
            on = re.sub(r"\s+", " ", tail[body_start:i]).strip()
        joins.append({"type": jtype, "table": table, "on": on})
    return joins


def extract_all_conditions(sql: str) -> list[str]:
    """收集所有层级的 WHERE 条件，包括子查询内部。

    排查场景下最外层往往只是 SELECT COUNT(1) FROM (...)，真正的限定条件
    全在子查询里。只看顶层 WHERE 会得到「0 个条件」这种没用的结论。
    """
    conds, seen = [], set()
    for m in re.finditer(r"\bWHERE\b", sql, re.I):
        start = m.end()
        depth, i = 0, start
        stop = re.compile(r"\b(GROUP\s+BY|ORDER\s+BY|HAVING|LIMIT|UNION)\b", re.I)
        while i < len(sql):
            c = sql[i]
            if c == "(":
                depth += 1
            elif c == ")":
                if depth == 0:          # 碰到包住自己的右括号，本层 WHERE 结束
                    break
                depth -= 1
            elif depth == 0 and stop.match(sql, i):
                break
            i += 1
        for c in split_conditions(sql[start:i].strip()):
            key = re.sub(r"\s+", "", c).lower()
            if key and key not in seen:
                seen.add(key)
                conds.append(c)
    return conds


def extract_all_select_fields(sql: str) -> list[str]:
    """取字段数最多的那个 SELECT 作为业务字段列表。

    外层常是 SELECT COUNT(1)，真正的字段清单在内层，取最长的那组更贴近意图。
    """
    best: list[str] = []
    for m in re.finditer(r"\bSELECT\b(\s+DISTINCT\b)?", sql, re.I):
        rest = sql[m.end():]
        depth, i = 0, 0
        while i < len(rest):
            c = rest[i]
            if c == "(":
                depth += 1
            elif c == ")":
                if depth == 0:
                    break
                depth -= 1
            elif depth == 0 and re.match(r"\bFROM\b", rest[i:], re.I):
                break
            i += 1
        fields = [re.sub(r"\s+", " ", f).strip()
                  for f in split_top_level(rest[:i], re.compile(r","))]
        if len(fields) > len(best):
            best = fields
    return best


def extract_clause(sql: str, kw: str) -> str:
    """取某个顶层子句的内容，如 WHERE / GROUP BY / HAVING"""
    kw_re = re.compile(rf"\b{kw}\b", re.I)
    pos = _find_kw(sql, kw_re)
    if pos < 0:
        return ""
    rest = sql[pos:]
    rest = rest[len(re.match(rf"\b{kw}\b", rest, re.I).group(0)):]
    stop = re.compile(r"\b(GROUP\s+BY|ORDER\s+BY|HAVING|LIMIT|UNION)\b", re.I)
    end = _find_kw(rest, stop)
    body = rest if end < 0 else rest[:end]
    return re.sub(r"\s+", " ", body).strip()


def _strip_wrapping_parens(text: str) -> str:
    """剥掉最外层成对括号，StarRocks 导出的 SQL 常见层层包裹"""
    body = text.strip()
    while body.startswith("(") and body.endswith(")"):
        depth, ok = 0, True
        for i, c in enumerate(body):
            depth += (c == "(") - (c == ")")
            if depth == 0 and i < len(body) - 1:
                ok = False
                break
        if not ok:
            break
        body = body[1:-1].strip()
    return body


def _split_and(text: str) -> list[str]:
    """按顶层 AND 切分条件，且不切开 BETWEEN x AND y 里的那个 AND"""
    parts, depth, last, i = [], 0, 0, 0
    pending_between = 0
    n = len(text)
    and_re = re.compile(r"\bAND\b", re.I)
    between_re = re.compile(r"\bBETWEEN\b", re.I)
    while i < n:
        c = text[i]
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
        elif depth == 0:
            if between_re.match(text, i):
                pending_between += 1
                i = between_re.match(text, i).end()
                continue
            m = and_re.match(text, i)
            if m:
                if pending_between:        # 这个 AND 属于 BETWEEN，不能切
                    pending_between -= 1
                    i = m.end()
                    continue
                parts.append(text[last:i])
                last = i = m.end()
                continue
        i += 1
    parts.append(text[last:])
    return [p.strip() for p in parts if p.strip()]


def split_conditions(where: str) -> list[str]:
    """把 WHERE 按顶层 AND 拆成条件列表，便于逐条比对限定条件"""
    if not where:
        return []
    conds = _split_and(_strip_wrapping_parens(where))
    return [c for c in (_strip_wrapping_parens(x) for x in conds) if c]


def extract_select_fields(sql: str) -> list[str]:
    """提取最外层 SELECT 的字段列表"""
    m = re.search(r"\bSELECT\b(\s+DISTINCT\b)?", sql, re.I)
    if not m:
        return []
    rest = sql[m.end():]
    end = _find_kw(rest, re.compile(r"\bFROM\b", re.I))
    body = rest if end < 0 else rest[:end]
    return [re.sub(r"\s+", " ", f).strip()
            for f in split_top_level(body, re.compile(r","))]


class ColumnScopeError(ValueError):
    """要引用的列在最外层查不到。

    在最外层写 DATE(某内层字段) 会被 StarRocks 直接判为
    `Column 'x' cannot be resolved`（Error 1064），必须提前拦下来并说清怎么改。
    """


def _select_alias(field: str) -> str:
    """取一个 SELECT 项对外暴露的列名。

    三种写法都要认：`x AS y`（显式别名）、`count(1) y`（隐式别名）、`t.x`（裸列名）。
    最外层能引用到的，只有这里返回的名字——这是判断作用域的唯一依据。
    """
    f = re.sub(r"\s+", " ", strip_comments(field)).strip().rstrip(",")
    if not f:
        return ""
    m = re.search(r"\bAS\s+(`?[\w\u4e00-\u9fff]+`?)\s*$", f, re.I)
    if m:
        return m.group(1).strip("`")
    # 隐式别名：别名前必须紧跟表达式的结尾（) 、引号或标识符），排除 CASE…END 这类尾关键字
    m = re.search(r"[\w`)'\"]\s+(`?[\w\u4e00-\u9fff]+`?)\s*$", f)
    if m and m.group(1).strip("`").upper() not in (
            "END", "ASC", "DESC", "NULL", "DISTINCT", "AND", "OR"):
        return m.group(1).strip("`")
    if re.fullmatch(IDENT, f):
        return f.replace("`", "").split(".")[-1]
    return ""


def _output_names(body: str) -> tuple[set[str], bool]:
    """一个查询块对外输出的列名集合。

    返回 (列名小写集合, 是否完整)。带 `*` 时不完整——静态判断不出实表有哪些列，
    这种情况一律放宽，不拦用户。
    """
    names: set[str] = set()
    complete = True
    parts = split_top_level(body, re.compile(r"\bUNION(\s+ALL)?\b", re.I))
    for part in parts or [body]:
        _, main = extract_ctes(part.strip())
        for f in extract_select_fields(main):
            if "*" in f:
                complete = False
                continue
            alias = _select_alias(f)
            if alias:
                names.add(alias.lower())
            else:
                complete = False
    return names, complete


def top_level_columns(sql: str) -> dict:
    """最外层 SELECT 能引用到的列名。

    只有一种情况能确定：最外层 FROM 紧跟一个派生表（或 CTE），且没再 JOIN 实表——
    此时最外层可见的列就是派生表输出的那几个别名。
    其余情况（FROM 实表、JOIN 实表、带 `*`）静态判断不了，known=False，一律放宽。
    """
    body = strip_comments(sql).strip().rstrip(";")
    ctes, main = extract_ctes(body)
    cte_map = {c["name"].lower(): c["body"] for c in ctes}
    m = re.search(r"\bSELECT\b(\s+DISTINCT\b)?", main, re.I)
    if not m:
        return {"names": set(), "known": False}
    rest = main[m.end():]
    fpos = _find_kw(rest, re.compile(r"\bFROM\b", re.I))
    if fpos < 0:
        return {"names": set(), "known": False}
    tail = rest[fpos + len("FROM"):]
    # 只取 FROM 到下一个顶层子句之间的「数据源」部分
    stop = _find_kw(tail, re.compile(
        r"\b(WHERE|GROUP\s+BY|ORDER\s+BY|HAVING|LIMIT|UNION)\b", re.I))
    src = (tail if stop < 0 else tail[:stop]).strip()
    # 最外层还 JOIN 或逗号连了别的数据源 → 可见列不止派生表那些，放宽
    if (_find_kw(src, JOIN_RE) >= 0
            or len(split_top_level(src, re.compile(r","))) > 1):
        return {"names": set(), "known": False}
    inner = ""
    if src.startswith("("):
        inner = src[1:_skip_parens(src, 0) - 1]
    else:
        ref = re.match(rf"({IDENT})", src)
        if ref and ref.group(1).replace("`", "").lower() in cte_map:
            inner = cte_map[ref.group(1).replace("`", "").lower()]
    if not inner.strip():
        return {"names": set(), "known": False}
    names, complete = _output_names(inner)
    return {"names": names, "known": bool(names) and complete}


def column_scope(sql: str, col: str) -> str:
    """这个列在最外层能不能被引用：top（能）/ inner（只在子查询里）/ unknown（判断不了）"""
    if not col:
        return "unknown"
    info = top_level_columns(sql)
    if not info["known"]:
        return "unknown"
    bare = col.replace("`", "").split(".")[-1].lower()
    return "top" if bare in info["names"] else "inner"


def top_level_time_candidates(sql: str) -> list[str]:
    """最外层可见的、像业务时间的列名，用于给「字段不可见」的报错提供替代方案"""
    info = top_level_columns(sql)
    return sorted(c for c in info["names"] if any(h in c for h in TIME_HINTS))


def ensure_top_level(sql: str, col: str, what: str = "业务时间字段") -> None:
    """改写前的作用域体检：列只在子查询里可见就直接报错，别把 1064 留给数据库"""
    if column_scope(sql, col) != "inner":
        return
    alts = top_level_time_candidates(sql)
    visible = sorted(top_level_columns(sql)["names"])
    raise ColumnScopeError(
        f"{what}「{col}」只存在于子查询内部，最外层查不到它。\n"
        f"本工具做时间分片 / 明细改写时要在最外层写 "
        f"DATE({col.split('.')[-1]})，数据库会直接报 "
        f"Column '{col.split('.')[-1]}' cannot be resolved（Error 1064）。\n"
        f"最外层实际可见的列：{('、'.join(visible[:12]) or '（识别不到）')}\n"
        f"三种改法任选一种：\n"
        f"  1. 换成最外层可见的时间字段"
        f"{'：' + '、'.join(alts[:5]) if alts else '（下拉里带「最外层」标记的才能用于分片）'}\n"
        f"  2. 在最内层子查询里把 {col.split('.')[-1]} 原样 select 出来、逐层往上传递\n"
        f"  3. 不勾「全量分片」，改用抽样模式——抽样只替换原 SQL 里已有的时间字面量，"
        f"不要求最外层可见")


def parse_sql(sql: str) -> dict:
    """把 SQL 解析成结构化描述。

    顶层子句与「全部层级」两套都给：顶层用于理解骨架，全部层级用于排查比对，
    因为真实脚本的限定条件与业务字段基本都藏在子查询里。
    """
    norm = normalize(sql)
    ctes, main = extract_ctes(norm)
    where = extract_clause(main, "WHERE")
    all_conds = extract_all_conditions(norm)
    all_fields = extract_all_select_fields(norm)
    return {
        "normalized": norm,
        "ctes": ctes,
        "tables": extract_tables(norm),
        "joins": extract_joins(norm),
        "where": where,
        "conditions": all_conds,               # 排查主用：含子查询内部
        "top_conditions": split_conditions(where),
        "group_by": extract_clause(main, "GROUP BY"),
        "having": extract_clause(main, "HAVING"),
        "order_by": extract_clause(main, "ORDER BY"),
        "select_fields": all_fields,           # 排查主用：最宽的那组字段
        "top_select_fields": extract_select_fields(main),
        "has_distinct": bool(re.search(r"\bSELECT\s+DISTINCT\b", norm, re.I)),
        "union_parts": len(split_top_level(
            main, re.compile(r"\bUNION(\s+ALL)?\b", re.I))),
        "where_blocks": len(re.findall(r"\bWHERE\b", norm, re.I)),
        "subquery_count": len(re.findall(r"\bFROM\s*\(|\bJOIN\s*\(", norm, re.I)),
    }


# ══════════════════════════════════════════════════════════════
# 第一层：逻辑比较
# ══════════════════════════════════════════════════════════════

AGG_RE = re.compile(r"\b(SUM|COUNT|AVG|MAX|MIN|GROUP_CONCAT)\s*\(", re.I)


def _risk(level: str, title: str, detail: str, advice: str) -> dict:
    return {"level": level, "title": title, "detail": detail, "advice": advice}


def compare_logic(sql_a: str, sql_b: str,
                  name_a: str = "旧脚本", name_b: str = "新脚本") -> dict:
    """只比逻辑不连库：列出结构差异，并按内置通用规则给出风险提示"""
    pa, pb = parse_sql(sql_a), parse_sql(sql_b)
    risks: list[dict] = []

    # ── 表差异 ──
    ta = {t["name"].lower() for t in pa["tables"]}
    tb = {t["name"].lower() for t in pb["tables"]}
    only_a, only_b = sorted(ta - tb), sorted(tb - ta)
    if only_a or only_b:
        risks.append(_risk(
            "high", "用到的表不一致",
            f"{name_a} 独有: {only_a or '无'}；{name_b} 独有: {only_b or '无'}",
            "确认换表是否有意为之。换表往往同时换了口径，需要核对两张表的过滤条件与数据范围"))

    # ── JOIN 类型差异（常见风险：LEFT→INNER 会丢数）──
    ja = {(j["table"].lower(), j["type"]) for j in pa["joins"]}
    jb = {(j["table"].lower(), j["type"]) for j in pb["joins"]}
    map_a = {t: ty for t, ty in ja}
    map_b = {t: ty for t, ty in jb}
    for tbl in sorted(set(map_a) & set(map_b)):
        if map_a[tbl] != map_b[tbl]:
            a_t, b_t = map_a[tbl], map_b[tbl]
            lvl, advice = "mid", "确认 JOIN 类型变更是否有意为之"
            if a_t.startswith("LEFT") and b_t.startswith("INNER"):
                lvl = "high"
                advice = ("LEFT JOIN 改成 INNER JOIN 会丢掉右表无匹配的记录，"
                          "这是数据变少最常见的原因。先用 LEFT JOIN 查右表为 NULL 的明细看丢了多少")
            elif a_t.startswith("INNER") and b_t.startswith("LEFT"):
                lvl = "high"
                advice = ("INNER 改成 LEFT 会让右表字段出现 NULL，若外层有 SUM/COUNT "
                          "或用该字段做过滤，结果会变。检查是否需要 ifnull 兜底")
            risks.append(_risk(lvl, f"JOIN 类型变更：{tbl}",
                               f"{name_a}: {a_t} → {name_b}: {b_t}", advice))

    # ── JOIN 关联键差异 ──
    on_a = {j["table"].lower(): j["on"] for j in pa["joins"]}
    on_b = {j["table"].lower(): j["on"] for j in pb["joins"]}
    for tbl in sorted(set(on_a) & set(on_b)):
        if on_a[tbl] != on_b[tbl] and on_a[tbl] and on_b[tbl]:
            risks.append(_risk(
                "high", f"JOIN 关联条件变更：{tbl}",
                f"{name_a}: {on_a[tbl]}\n{name_b}: {on_b[tbl]}",
                "关联键变了会直接改变匹配行数，可能引入一对多倍乘或丢数。"
                "分别用 GROUP BY 关联键 HAVING COUNT(1)>1 验证两边的一对多情况"))

    # ── WHERE 限定条件差异（限定条件不一致是高频原因）──
    ca = {c.lower(): c for c in pa["conditions"]}
    cb = {c.lower(): c for c in pb["conditions"]}
    cond_only_a = [ca[k] for k in ca if k not in cb]
    cond_only_b = [cb[k] for k in cb if k not in ca]
    if cond_only_a or cond_only_b:
        risks.append(_risk(
            "high", "WHERE 限定条件不一致",
            f"{name_a} 多出 {len(cond_only_a)} 条、{name_b} 多出 {len(cond_only_b)} 条",
            "逐条核对限定条件。少一个条件会多出数据，多一个条件会少出数据"))

    # ── 聚合与 GROUP BY 一致性 ──
    for tag, p, nm in (("a", pa, name_a), ("b", pb, name_b)):
        if AGG_RE.search(p["normalized"]) and not p["group_by"] and p["joins"]:
            risks.append(_risk(
                "mid", f"{nm}：有聚合和 JOIN 但没有 GROUP BY",
                "JOIN 之后直接聚合，若右表一对多会把指标放大",
                "先把右表按关联键聚合再 JOIN，或改用 COUNT(DISTINCT 主键)"))
    if pa["group_by"] != pb["group_by"]:
        risks.append(_risk(
            "mid", "GROUP BY 不一致",
            f"{name_a}: {pa['group_by'] or '无'}\n{name_b}: {pb['group_by'] or '无'}",
            "分组维度变化会改变结果行数与聚合口径"))
    if pa["has_distinct"] != pb["has_distinct"]:
        risks.append(_risk(
            "mid", "DISTINCT 使用不一致",
            f"{name_a}: {'有' if pa['has_distinct'] else '无'} / "
            f"{name_b}: {'有' if pb['has_distinct'] else '无'}",
            "一边去重一边不去重，存在一对多时两边行数必然不同"))

    # ── SELECT 字段差异 ──
    fa = {f.lower(): f for f in pa["select_fields"]}
    fb = {f.lower(): f for f in pb["select_fields"]}
    fld_only_a = [fa[k] for k in fa if k not in fb]
    fld_only_b = [fb[k] for k in fb if k not in fa]

    if pa["union_parts"] != pb["union_parts"]:
        risks.append(_risk(
            "mid", "UNION 分支数量不同",
            f"{name_a}: {pa['union_parts']} 段 / {name_b}: {pb['union_parts']} 段",
            "确认是否漏了某个分支。UNION ALL 不去重、UNION 去重，两者结果不同"))

    order = {"high": 0, "mid": 1, "low": 2}
    risks.sort(key=lambda r: order.get(r["level"], 9))
    return {
        "parsed": {"a": pa, "b": pb},
        "names": {"a": name_a, "b": name_b},
        "tables": {"only_a": only_a, "only_b": only_b,
                   "both": sorted(ta & tb)},
        "conditions": {"only_a": cond_only_a, "only_b": cond_only_b,
                       "same": [ca[k] for k in ca if k in cb]},
        "fields": {"only_a": fld_only_a, "only_b": fld_only_b},
        "joins": {"a": pa["joins"], "b": pb["joins"]},
        "risks": risks,
        "identical": (not risks and pa["normalized"] == pb["normalized"]),
    }


def text_diff(sql_a: str, sql_b: str) -> list[dict]:
    """归一化后逐行 diff，作为结构差异的补充细节"""
    import difflib

    def prep(s: str) -> list[str]:
        s = strip_comments(s)
        # 在主要子句前断行，让 diff 结果贴近人的阅读习惯
        s = re.sub(r"\s+", " ", s).strip()
        s = re.sub(r"\b(SELECT|FROM|WHERE|GROUP\s+BY|ORDER\s+BY|HAVING|LIMIT|"
                   r"LEFT\s+JOIN|RIGHT\s+JOIN|INNER\s+JOIN|CROSS\s+JOIN|JOIN|"
                   r"UNION\s+ALL|UNION|AND|OR)\b",
                   lambda m: "\n" + m.group(0).upper(), s, flags=re.I)
        return [l.strip() for l in s.split("\n") if l.strip()]

    la, lb = prep(sql_a), prep(sql_b)
    out = []
    for line in difflib.unified_diff(la, lb, lineterm="", n=1):
        if line.startswith(("---", "+++")):
            continue
        kind = ("hunk" if line.startswith("@@") else
                "add" if line.startswith("+") else
                "del" if line.startswith("-") else "same")
        out.append({"kind": kind, "text": line})
    return out


def _strip_alias(field: str) -> str:
    """去掉 AS 别名，返回表达式本体"""
    return re.sub(r"\s+AS\s+[`\"]?[\w\u4e00-\u9fff]+[`\"]?\s*$", "", field.strip(),
                  flags=re.I).strip()


def _bare_col(expr: str) -> str:
    """从 date(a.`x`) 这类表达式里剥出列名 a.x；剥不出就原样返回"""
    e = expr.strip()
    # 层层剥掉最外层函数
    while True:
        m = re.match(r"^\w+\s*\((.*)\)$", e, re.S)
        if not m:
            break
        e = m.group(1).strip()
        # CAST(x AS type) → x
        e = re.sub(r"\s+AS\s+[\w()\d,\s]+$", "", e, flags=re.I).strip()
    e = e.split(",")[0].strip()
    return e.replace("`", "")


def suggest_primary_keys(sql: str) -> list[dict]:
    """推荐明细主键，按可信度排序。

    使用通用命名约定：id、*_id、pk、pk_* 等字段优先；
    再兜底：SELECT 的第一个非聚合字段。
    """
    norm = normalize(sql)
    p = parse_sql(norm)
    found: dict[str, dict] = {}

    def add(col: str, source: str, score: int, note: str = ""):
        col = (col or "").strip().replace("`", "")
        if not col or " " in col or "(" in col or col.upper() in ("NULL", "*"):
            return
        bare = col.split(".")[-1].lower()
        if not bare or bare in ("and", "or", "on"):
            return
        cur = found.get(col.lower())
        if cur is None or score > cur["score"]:
            found[col.lower()] = {"column": col, "bare": bare, "source": source,
                                  "score": score, "note": note}

    # 1. SELECT 第一个非聚合字段 —— 用户口径：一般第一个就是主键
    for i, f in enumerate(p["select_fields"]):
        body = _strip_alias(f)
        if AGG_RE.search(body):
            continue
        col = _bare_col(body)
        base = 90 if i == 0 else max(70 - i * 3, 40)
        add(col, f"SELECT 第 {i+1} 个字段", base,
            "SELECT 首个非聚合字段，通常就是明细粒度" if i == 0 else "")

    # 2. JOIN 的 ON 条件里的字段：通用主键/ID 命名优先
    for j in p["joins"]:
        for m in re.finditer(rf"({IDENT})\s*=\s*({IDENT})", j["on"] or ""):
            for col in (m.group(1), m.group(2)):
                c = col.replace("`", "")
                bare = c.split(".")[-1].lower()
                if bare == "pk" or bare.startswith("pk_"):
                    add(c, f"JOIN 条件（{j['table']}）", 92, "字段名符合常见主键约定")
                elif bare.endswith("_id") or bare == "id":
                    add(c, f"JOIN 条件（{j['table']}）", 72, "字段名符合常见 ID 约定")

    out = sorted(found.values(), key=lambda x: -x["score"])
    if out:
        out[0]["default"] = True          # 默认取第一个
    return out[:12]


def build_detail_sql(sql: str, keys: list[str], time_col: str = "",
                     start: str = "", end: str = "", limit: int = 500) -> str:
    """把聚合口径的指标 SQL 改写成按主键取明细的 SQL。

    指标 SQL 通常只 SELECT date() 和 sum()，没法按主键对齐比较。
    这里外层套一次，把主键和时间列显式带出来。

    套壳只能引用最外层输出的列。主键或时间字段是最外层起的别名时，套壳后照样能用；
    真的埋在子查询里（外层压根没输出它）才没办法，那就从 WHERE / ORDER BY 里去掉并写明原因。
    """
    inner = strip_comments(sql).strip().rstrip(";")
    out_names = {a for a in outer_alias_map(sql)} | set(top_level_columns(sql)["names"])

    def usable(c):
        bare = c.replace("`", "").split(".")[-1].lower()
        return column_scope(sql, c) != "inner" or bare in out_names

    keep = [k for k in keys if usable(k)]
    hidden = [k for k in keys if k not in keep]
    cols = ", ".join(f"`{k.split('.')[-1].strip('`')}`" for k in keep)
    where = ""
    if time_col and start and end and usable(time_col):
        tc = f"`{time_col.split('.')[-1].strip('`')}`"
        where = f"\nWHERE DATE({tc}) BETWEEN '{start}' AND '{end}'"
    elif time_col:
        hidden.append(time_col)
    order = f"\nORDER BY {cols}" if cols else ""
    warn = ("" if not hidden else
            f"-- ⚠️ {'、'.join(dict.fromkeys(hidden))} 最外层没有输出，"
            f"套壳后引用会报 cannot be resolved，已从 WHERE/ORDER BY 中去掉；\n"
            f"--    要按它们过滤，请在子查询的 SELECT 里把这些字段带出来\n")
    return (f"-- 由指标 SQL 改写为明细口径：显式带出主键，便于按主键对齐比较\n"
            f"{warn}"
            f"SELECT * FROM (\n{inner}\n) AS _detail{where}{order}\nLIMIT {limit}")


# ══════════════════════════════════════════════════════════════
# 长 SQL 逻辑拆分
# ══════════════════════════════════════════════════════════════


def _describe(body: str) -> str:
    """给片段生成一句人话说明，帮助快速读懂长 SQL 在干什么"""
    p = parse_sql(body)
    bits = []
    names = [t["name"] for t in p["tables"][:3]]
    if names:
        bits.append("读 " + "、".join(names) + ("…" if len(p["tables"]) > 3 else ""))
    if p["joins"]:
        kinds = sorted({j["type"] for j in p["joins"]})
        bits.append(f"{len(p['joins'])} 个 JOIN（{'/'.join(kinds)}）")
    if p["conditions"]:
        bits.append(f"{len(p['conditions'])} 个过滤条件")
    if p["group_by"]:
        bits.append(f"按 {p['group_by'][:40]} 聚合")
    if AGG_RE.search(p["normalized"]):
        bits.append("含聚合函数")
    if p["has_distinct"]:
        bits.append("去重")
    return "；".join(bits) or "简单查询"


def split_sql(sql: str) -> dict:
    """把长 SQL 拆成可读片段：CTE → UNION 分支 → 派生子查询"""
    norm = normalize(sql)
    ctes, main = extract_ctes(norm)
    segments = []

    for i, c in enumerate(ctes, 1):
        segments.append({
            "kind": "CTE", "no": f"CTE-{i}", "name": c["name"],
            "desc": _describe(c["body"]), "sql": c["body"],
            "chars": len(c["body"]),
            "tables": [t["name"] for t in extract_tables(c["body"])],
        })

    parts = split_top_level(main, re.compile(r"\bUNION(\s+ALL)?\b", re.I))
    for i, part in enumerate(parts, 1):
        segments.append({
            "kind": "主查询" if len(parts) == 1 else f"UNION 分支{i}",
            "no": f"MAIN-{i}", "name": "" if len(parts) == 1 else f"分支{i}",
            "desc": _describe(part), "sql": part, "chars": len(part),
            "tables": [t["name"] for t in extract_tables(part)],
        })

    # 派生子查询：FROM ( ... ) 形式，单独列出便于逐段验证
    subs = []
    for m in re.finditer(r"\b(?:FROM|JOIN)\s*\(", main, re.I):
        start = m.end()
        depth, j = 1, start
        while j < len(main) and depth:
            if main[j] == "(":
                depth += 1
            elif main[j] == ")":
                depth -= 1
            j += 1
        body = main[start:j - 1].strip()
        if len(body) > 60 and re.match(r"SELECT\b", body, re.I):
            subs.append(body)
    for i, body in enumerate(subs, 1):
        am = re.search(r"\)\s*(?:AS\s+)?(`?\w+`?)", main[main.find(body) + len(body):], re.I)
        segments.append({
            "kind": "派生子查询", "no": f"SUB-{i}",
            "name": (am.group(1).strip("`") if am else f"子查询{i}"),
            "desc": _describe(body), "sql": body, "chars": len(body),
            "tables": [t["name"] for t in extract_tables(body)],
        })

    return {"total_chars": len(norm), "segment_count": len(segments),
            "segments": segments,
            "outline": [f"{s['no']} {s['name']}: {s['desc']}" for s in segments]}


# ══════════════════════════════════════════════════════════════
# 业务时间识别 + 500 行限制下的切片取数
# ══════════════════════════════════════════════════════════════


def find_time_columns(sql: str) -> list[dict]:
    """从 SQL 里找出可用于时间切片的字段。

    Matrix 单次查询最多回 500 行，排查明细必须靠业务时间收窄范围。
    优先级：WHERE 里已用于时间过滤的 > SELECT 出来的业务时间 > 系统时间。

    还要判断「最外层能不能引用到」：只在子查询里出现的字段，分片时会被数据库
    判为 cannot be resolved，这类一律降到候选末尾并标注出来。
    """
    norm = normalize(sql)
    found: dict[str, dict] = {}
    # 这些是函数名/关键字，不是字段，别混进候选
    not_column = {"date", "datetime", "timestamp", "now", "curdate", "curtime",
                  "current_date", "current_timestamp", "time", "year", "month",
                  "day", "date_format", "date_add", "date_sub", "datediff",
                  "timestampdiff", "str_to_date", "unix_timestamp", "between"}

    def add(col: str, source: str, score: int):
        # 用 replace 而不是 strip：strip 只去两端，a.`col` 会剩下中间那个反引号
        col = col.strip().replace("`", "")
        if not col or " " in col:
            return
        bare = col.split(".")[-1].lower()
        if bare in not_column:
            return
        if not any(h in bare for h in TIME_HINTS):
            return
        if bare in SYS_TIME:
            score -= 50
        cur = found.get(col.lower())
        if cur is None or score > cur["score"]:
            found[col.lower()] = {"column": col, "source": source, "score": score,
                                  "is_sys_time": bare in SYS_TIME}

    # 1. WHERE 中已参与时间比较的字段，最可信
    for m in re.finditer(rf"({IDENT})\s*(?:BETWEEN|>=|<=|>|<|=)", norm):
        add(m.group(1), "WHERE 时间过滤", 100)
    for m in re.finditer(rf"\bDATE\s*\(\s*({IDENT})\s*\)", norm, re.I):
        add(m.group(1), "WHERE DATE() 截断", 95)
    # 2. SELECT 字段
    for f in extract_select_fields(norm):
        alias = re.search(r"\bAS\s+(`?[\w\u4e00-\u9fff]+`?)$", f, re.I)
        add(alias.group(1) if alias else f, "SELECT 字段", 60)
    # 3. 兜底：全文出现的疑似时间字段
    for m in re.finditer(rf"({IDENT})", norm):
        add(m.group(1), "SQL 中出现", 20)

    # 4. 作用域标注：最外层引用不到的，看能不能换算/下推，真的不行才标成不可用
    aliases = outer_alias_map(sql)
    for item in found.values():
        col = item["column"]
        scope = column_scope(sql, col)
        item["scope"] = scope
        item["usable"] = True
        item["scope_label"] = {"top": "最外层可见", "unknown": ""}.get(scope, "")
        if scope != "inner":
            continue
        bare = col.replace("`", "").split(".")[-1].lower()
        expr = aliases.get(bare)
        if expr and expr.replace("`", "").split(".")[-1].lower() != bare \
                and _expr_top_level_ok(sql, expr):
            item["scope"] = "alias"
            item["scope_label"] = f"最外层别名·改写时换成 {expr}"
            item["score"] -= 5              # 能用，但不如直接可见的干净
        elif innermost_blocks_with(strip_comments(sql), col):
            item["scope"] = "pushdown"
            item["scope_label"] = "仅子查询内可见·时间条件会下推到子查询"
            item["score"] -= 10
        else:
            item["usable"] = False
            item["scope_label"] = "最外层引用不到·且换算不出来"
            item["score"] -= 200

    out = sorted(found.values(), key=lambda x: -x["score"])
    return out[:15]


def _col_pattern(time_col: str) -> str:
    """把 a.b 形式的列名变成能匹配 SQL 里各种反引号写法的正则"""
    parts = [p for p in time_col.replace("`", "").split(".") if p]
    return r"\.\s*".join(rf"`?{re.escape(p)}`?" for p in parts)


def _split_field_alias(field: str) -> tuple[str, str]:
    """把一个 SELECT 项拆成 (表达式, 别名)"""
    f = re.sub(r"\s+", " ", strip_comments(field)).strip().rstrip(",")
    alias = _select_alias(f)
    if not alias:
        return f, ""
    m = re.search(rf"\s+AS\s+`?{re.escape(alias)}`?\s*$", f, re.I)
    if m:
        return f[:m.start()].strip(), alias
    m = re.search(rf"\s+`?{re.escape(alias)}`?\s*$", f)
    if m and f[:m.start()].strip():
        return f[:m.start()].strip(), alias
    return f, alias                      # 裸列名：表达式与别名是同一个东西


def outer_alias_map(sql: str) -> dict[str, str]:
    """最外层 SELECT 列表的「别名 → 背后的表达式」。

    别名只是给结果集起的名字。改写最外层 SELECT 时整个列表都被换掉，别名随之消失，
    所以要按别名分片/过滤，必须换回它背后的表达式：
        SELECT t4.datetime AS year_month FROM (…) t4
    想按 year_month 分组，实际要写的是 DATE(t4.datetime)。
    """
    body = strip_comments(sql).strip().rstrip(";")
    _, main = extract_ctes(body)
    out: dict[str, str] = {}
    for f in extract_select_fields(main):
        expr, alias = _split_field_alias(f)
        if alias and expr:
            out.setdefault(alias.lower(), expr)
    return out


# 表达式里这些词不是列名，判断作用域时要跳过
_EXPR_NOISE = {"AS", "AND", "OR", "NOT", "CASE", "WHEN", "THEN", "ELSE", "END",
               "IS", "NULL", "IN", "LIKE", "BETWEEN", "DISTINCT", "INTERVAL",
               "DAY", "MONTH", "YEAR", "HOUR", "MINUTE", "SECOND", "DECIMAL",
               "INT", "BIGINT", "VARCHAR", "DATE", "DATETIME", "CHAR"}


def _expr_columns(expr: str) -> list[str]:
    """表达式里引用到的列名（函数名、数字、字符串字面量、SQL 关键字都不算）"""
    body = re.sub(r"'[^']*'", "''", expr)
    cols = []
    for m in re.finditer(rf"({IDENT})", body):
        tok = m.group(1)
        if body[m.end():m.end() + 1] == "(":          # 函数名
            continue
        bare = tok.replace("`", "").split(".")[-1]
        if re.fullmatch(r"\d+", bare) or bare.upper() in _EXPR_NOISE:
            continue
        cols.append(tok)
    return cols


def _expr_top_level_ok(sql: str, expr: str) -> bool:
    """表达式里引用的列是否都在最外层可见（函数名/数字/字符串字面量不算列）"""
    return all(column_scope(sql, tok) != "inner" for tok in _expr_columns(expr))


def resolve_top_level(sql: str, col: str, what: str = "业务时间字段") -> dict:
    """把列名换算成「最外层能直接引用」的表达式。

    返回 {expr, derived, note}：derived=True 表示换算过，expr 已不是原来那个名字。
    实在换算不出来才抛 ColumnScopeError（附带三条改法）。
    """
    if column_scope(sql, col) != "inner":
        return {"expr": col, "derived": False, "note": ""}
    bare = col.replace("`", "").split(".")[-1].lower()
    expr = outer_alias_map(sql).get(bare)
    if expr and expr.replace("`", "").split(".")[-1].lower() != bare \
            and _expr_top_level_ok(sql, expr):
        return {"expr": expr, "derived": True,
                "note": f"{what}「{col}」是最外层 SELECT 起的别名，"
                        f"改写时已自动换回它背后的表达式 {expr}"}
    ensure_top_level(sql, col, what)
    return {"expr": col, "derived": False, "note": ""}      # 不会走到


def _query_blocks(sql: str) -> list[tuple[int, int]]:
    """所有括号包起来的子查询块，返回 [(内容起, 内容止)]"""
    out = []
    for i, ch in enumerate(sql):
        if ch != "(":
            continue
        j = _skip_parens(sql, i)
        if re.match(r"\s*SELECT\b", sql[i + 1:j - 1], re.I):
            out.append((i + 1, j - 1))
    return out


def innermost_blocks_with(sql: str, col: str) -> list[tuple[int, int]]:
    """能看见这个列的最内层子查询块。

    时间条件要加在列真正可见的那一层——这就是「把 WHERE 下推到子查询」：
    外层看不见 event_time，但产生它的那个子查询看得见，
    条件写进去既合法，又能让引擎在扫表阶段就把数据过滤掉。
    """
    pat = re.compile(_col_pattern(col), re.I)
    hits = [(s, e) for s, e in _query_blocks(sql) if pat.search(sql[s:e])]
    # 只保留最内层：块里还套着另一个命中块的，交给里面那层
    return [(s, e) for s, e in hits
            if not any(s < s2 and e2 < e for s2, e2 in hits)]


def _inject_cond(block: str, cond: str) -> str:
    """把条件加进一个查询块：有 WHERE 就接在后面，没有就在 GROUP BY 之前新建"""
    wpos = _find_kw(block, re.compile(r"\bWHERE\b", re.I))
    if wpos >= 0:
        after = wpos + len(re.match(r"\bWHERE\b", block[wpos:], re.I).group(0))
        return block[:after] + f" {cond} AND " + block[after:]
    stop = _find_kw(block, re.compile(
        r"\b(GROUP\s+BY|ORDER\s+BY|HAVING|LIMIT|UNION)\b", re.I))
    if stop >= 0:
        return block[:stop] + f" WHERE {cond} " + block[stop:]
    return block + f" WHERE {cond}"


def _all_select_fields(sql: str) -> list[str]:
    """所有层次 SELECT 的字段项（不只最外层），用于跨层追列定义"""
    out = []
    for m in re.finditer(r"\bSELECT\b(\s+DISTINCT\b)?", sql, re.I):
        rest = sql[m.end():]
        depth, i = 0, 0
        while i < len(rest):
            c = rest[i]
            if c == "(":
                depth += 1
            elif c == ")":
                if depth == 0:
                    break
                depth -= 1
            elif depth == 0 and re.match(r"\bFROM\b", rest[i:], re.I):
                break
            i += 1
        out.extend(split_top_level(rest[:i], re.compile(r",")))
    return out


def _def_step(sql: str, name: str) -> str:
    """往里走一层：这个名字是谁的别名，返回它背后的表达式（没有就返回空）"""
    bare = name.replace("`", "").split(".")[-1]
    for f in _all_select_fields(strip_comments(sql)):
        e, a = _split_field_alias(f)
        if a and a.lower() == bare.lower() and e and \
                e.replace("`", "").split(".")[-1].lower() != bare.lower():
            return e
    return ""


def trace_definition(sql: str, col: str, depth: int = 3) -> str:
    """把列名一路追到它最初的定义表达式（可跨多层子查询与别名）。

    year_month → t4.datetime → DATE_FORMAT(event_time, '%Y%m%d')
    追到底才知道这列里装的到底是 datetime 还是 '20250701' 这样的字符串。
    """
    expr = _def_step(sql, col)
    if not expr:
        return col
    if depth > 0 and re.fullmatch(IDENT, expr.strip()):
        return trace_definition(sql, expr.strip(), depth - 1)
    return expr


def _def_chain(sql: str, name: str, depth: int = 5) -> list[str]:
    """一个最外层别名往里追的完整链条：[别名, 上一层表达式, 更里层表达式…]

    trace_definition 只要最里层那一站，这里要沿途每一站——
    用户填的列名可能停在任意一层（`t7.datetime` 是中间站，
    `event_time` 才是终点站），逐站比才对得上。
    """
    chain, cur = [name], name
    for _ in range(depth):
        nxt = _def_step(sql, cur)
        if not nxt or nxt in chain:
            break
        chain.append(nxt)
        cur = nxt.strip() if re.fullmatch(IDENT, nxt.strip()) else ""
        if not cur:
            break
    return chain


def alias_for_column(sql: str, col: str) -> dict:
    """反查：埋在子查询里的列，在最外层被起成了哪个别名。

    主键和时间字段平时不会原样 select 到最外层，外面只留一个别名：
        select concat(t7.datetime, t7.category_id) as id,
               t7.datetime as year_month
        from ( select DATE_FORMAT(event_time,'%Y%m%d') datetime … ) t7
    拿 `t7.datetime` 直接去结果集里找必然找不到，得顺着别名一层层追下来，
    认出它对外就叫 year_month。

    只认「这一层表达式整体就是该列」的情况：`concat(t7.datetime, …) as id`
    里 datetime 只是拼串的一块料，认成 id 会把两个不同的东西当同一个。
    返回 {alias, chain, exact}，找不到返回 {}。
    """
    bare = col.replace("`", "").split(".")[-1].lower()
    body = strip_comments(sql).strip().rstrip(";")
    _, main = extract_ctes(body)
    loose = {}
    for f in extract_select_fields(main):
        expr, alias = _split_field_alias(f)
        if not alias:
            continue
        for i, step in enumerate(_def_chain(sql, alias)):
            s = step.strip()
            if s.replace("`", "").split(".")[-1].lower() == bare:
                return {"alias": alias, "chain": _def_chain(sql, alias),
                        "exact": i == 0}
            cols = _expr_columns(s)
            # 这一层是个只包一列的函数（DATE_FORMAT(event_time,…)），
            # 算换算得出；包了多列的拼串表达式不算
            if len(cols) == 1 and \
                    cols[0].replace("`", "").split(".")[-1].lower() == bare:
                loose.setdefault(alias, _def_chain(sql, alias))
    if loose:
        alias = next(iter(loose))
        return {"alias": alias, "chain": loose[alias], "exact": False}
    return {}


def resolve_keys(keys: list[str], sql: str, headers: list[str]) -> dict:
    """把用户填的主键换算成结果集里真实存在的列名。

    返回 {resolved: {原名: 结果集列名}, notes: [...], missing: [...]}。
    """
    resolved, notes, missing = {}, [], []
    for k in keys:
        hit = _resolve_column(k, headers)
        if hit:
            resolved[k] = hit
            continue
        info = alias_for_column(sql, k) if sql else {}
        hit = _resolve_column(info.get("alias", ""), headers) if info else None
        if hit:
            resolved[k] = hit
            notes.append(f"主键「{k}」最外层没有原样输出，已按别名链换算为「{hit}」"
                         f"（{' → '.join(info['chain'])}）")
        else:
            missing.append(k)
    return {"resolved": resolved, "notes": notes, "missing": missing}


# 这些定义方式产出的是「日粒度」值：datetime 比较那一套用不了
_DAY_GRAIN_RE = re.compile(
    r"DATE_FORMAT\s*\([^,]+,\s*'%Y[-/]?%m[-/]?%d'\s*\)|"
    r"DATE_FORMAT\s*\([^,]+,\s*'%Y[-/]?%m'\s*\)|"
    r"\bTO_DATE\s*\(|\bDATE\s*\(", re.I)


def column_day_grain(sql: str, col: str) -> bool:
    """这列是不是只有日粒度（DATE_FORMAT('%Y%m%d') / DATE() 之类产出的）。

    关键差别：'20250705' 直接和 '2025-07-05 00:00:00' 比大小恒为 false——
    实测会一条数据都查不到，静默少数据比报错更难查。这种列必须两边都套 DATE()。
    """
    return bool(_DAY_GRAIN_RE.search(trace_definition(sql, col)))


def _range_cond(col: str, lo: str, hi: str, day_grain: bool) -> str:
    """生成时间范围条件。

    day_grain（DATE_FORMAT 出来的 '%Y%m%d' 字符串等）必须两边都套 DATE()：
    实测 '20250701' BETWEEN '2025-07-01 00:00:00' AND '2025-07-05 23:59:59' 返回 false。
    真 datetime 列用原值比较更精确，也不会挡住分区裁剪。
    """
    if day_grain:
        return f"DATE({col}) BETWEEN DATE('{lo}') AND DATE('{hi}')"
    return f"{col} BETWEEN '{lo}' AND '{hi}'"


def rewrite_time_filter(sql: str, time_col: str, start: str, end: str) -> dict:
    """在原 SQL 自己的层次上替换（或注入）时间条件。

    为什么不能在外面套一层 SELECT * FROM (原SQL) WHERE 时间…：
    聚合口径的 SQL 输出的列名是 `date(a.register_time)` 这种表达式，
    外层根本没有 register_time 这个列，一套壳就报 Column cannot be resolved。
    所以必须改原 SQL 自己的 WHERE。

    四种改法按优先级来：
      1. 已有 BETWEEN 字面量 → 换掉两个日期，最贴近原意
      2. 已有 >= / <= 范围 → 分别换掉
      3. 列只在子查询里可见 → 把条件下推到看得见它的那一层（可能是多个 UNION 分支）
      4. 列在最外层可见 → 加进最外层 WHERE

    start/end 可以只给日期（自动补 00:00:00 / 23:59:59），也可以给到秒。
    返回 {sql, mode, note}
    """
    body = strip_comments(sql).strip().rstrip(";")
    col_re = _col_pattern(time_col)
    lo_raw, hi_raw = _as_lo(start), _as_hi(end)
    lo, hi = f"'{lo_raw}'", f"'{hi_raw}'"

    # 1. 已有 BETWEEN 'x' AND 'y' → 直接换掉两个字面量，最贴近原意
    pat = re.compile(rf"({col_re}\s*)(BETWEEN\s*)'[^']*'(\s*AND\s*)'[^']*'", re.I)
    new, n = pat.subn(rf"\g<1>\g<2>{lo}\g<3>{hi}", body)
    if n:
        return {"sql": new, "mode": "replaced_between",
                "note": f"替换了原 SQL 里 {n} 处 BETWEEN 时间范围（在原层次，未套壳）"}

    # 2. 已有 >= / > 和 <= / < 的范围写法 → 分别替换
    ge = re.compile(rf"({col_re}\s*)(>=?\s*)'[^']*'", re.I)
    le = re.compile(rf"({col_re}\s*)(<=?\s*)'[^']*'", re.I)
    new, n1 = ge.subn(rf"\g<1>\g<2>{lo}", body)
    new, n2 = le.subn(rf"\g<1>\g<2>{hi}", new)
    if n1 or n2:
        return {"sql": new, "mode": "replaced_range",
                "note": f"替换了原 SQL 里的时间范围条件（>= {n1} 处、<= {n2} 处）"}

    # 3. 列只在子查询里可见 → 条件下推到看得见它的那一层
    if column_scope(sql, time_col) == "inner":
        blocks = innermost_blocks_with(body, time_col)
        if blocks:
            cond = _range_cond(time_col, lo_raw, hi_raw,
                               column_day_grain(sql, time_col))
            new = body
            for s, e in sorted(blocks, reverse=True):     # 从后往前改，偏移才不乱
                new = new[:s] + _inject_cond(new[s:e], cond) + new[e:]
            return {"sql": new, "mode": "pushed_down",
                    "note": f"「{time_col}」在最外层看不到，已把时间条件下推到"
                            f"看得见它的 {len(blocks)} 个子查询的 WHERE 里"}
        # 子查询里也找不到 → 试着换成最外层的等价表达式
        r = resolve_top_level(sql, time_col)
        cond = _range_cond(r["expr"], lo_raw, hi_raw,
                           column_day_grain(sql, time_col))
        return {"sql": _inject_cond(body, cond), "mode": "injected_resolved",
                "note": r["note"] + "；条件已加进最外层 WHERE"}

    # 4. 列在最外层可见 → 注入到最外层 WHERE
    cond = _range_cond(time_col, lo_raw, hi_raw, column_day_grain(sql, time_col))
    return {"sql": _inject_cond(body, cond), "mode": "injected",
            "note": "原 SQL 没有该字段的时间条件，已注入到最外层 WHERE"
                    + ("（该列是日粒度字符串，条件两边都套了 DATE() 才比得动）"
                       if column_day_grain(sql, time_col) else "")}


def _as_lo(s: str) -> str:
    """把边界补成起始时刻：只给日期就补 00:00:00"""
    s = str(s).strip()
    return s if len(s) > 10 else f"{s} 00:00:00"


def _as_hi(s: str) -> str:
    """把边界补成结束时刻：只给日期就补 23:59:59"""
    s = str(s).strip()
    if len(s) > 10:
        return s if len(s) > 16 else f"{s}:59"
    return f"{s} 23:59:59"


def build_count_sql(sql: str) -> str:
    """聚合后的输出行数：原 SQL 到底返回多少行。COUNT 不引用内层字段，套壳安全"""
    inner = strip_comments(sql).strip().rstrip(";")
    return f"SELECT COUNT(1) AS _cnt FROM (\n{inner}\n) AS _t"


def _replace_select_list(sql: str, new_list: str, keep_group: bool = False,
                         group_expr: str = "") -> str:
    """把最外层 SELECT 列表换成 new_list，并按需重设 GROUP BY。

    在原层次改写而不是外面套壳：聚合 SQL 输出的列名是表达式，
    外层引用原始字段会报 Column cannot be resolved。
    """
    body = strip_comments(sql).strip().rstrip(";")
    m = re.search(r"\bSELECT\b(\s+DISTINCT\b)?", body, re.I)
    if not m:
        raise ValueError("这段 SQL 里找不到 SELECT")
    rest = body[m.end():]
    fpos = _find_kw(rest, re.compile(r"\bFROM\b", re.I))
    if fpos < 0:
        raise ValueError("这段 SQL 里找不到顶层 FROM")
    tail = rest[fpos:]
    cut = _find_kw(tail, re.compile(r"\b(GROUP\s+BY|ORDER\s+BY|HAVING|LIMIT)\b", re.I))
    if cut >= 0:
        tail = tail[:cut]
    out = f"SELECT {new_list}\n{tail}"
    if keep_group and group_expr:
        out += f"\nGROUP BY {group_expr}\nORDER BY {group_expr}"
    return out


def build_detail_count_sql(sql: str) -> str:
    """明细层行数：GROUP BY 之前有多少行。

    和分布查询同口径，两者才能相互校验。
    套壳 COUNT 数的是聚合后的行数（比如按天聚合就只有 26 行），
    但排查明细时关心的是聚合前的真实行数。
    """
    return _replace_select_list(sql, "COUNT(1) AS _cnt")


def _agg_inner_fields(sql: str) -> list[str]:
    """把 SELECT 里聚合函数包着的字段剥出来。

    sum(cast(a.x as decimal(27,8))) → a.x
    聚合的度量字段是排查重点：两边对不上，多半就差在这个字段上。
    """
    out, seen = [], set()
    body = strip_comments(sql)
    for m in re.finditer(r"\b(SUM|COUNT|AVG|MAX|MIN|GROUP_CONCAT)\s*\(", body, re.I):
        start = m.end()
        depth, i = 1, start
        while i < len(body) and depth:
            if body[i] == "(":
                depth += 1
            elif body[i] == ")":
                depth -= 1
            i += 1
        col = _bare_col(body[start:i - 1])
        col = re.sub(r"^\s*DISTINCT\s+", "", col, flags=re.I).strip()
        if (col and col not in ("1", "*") and " " not in col
                and col.lower() not in seen):
            seen.add(col.lower())
            out.append(col)
    return out


def build_detail_select_sql(sql: str, keys: list[str], time_col: str = "",
                            extra: list[str] | None = None) -> dict:
    """把聚合口径的指标 SQL 改写成明细口径。

    指标 SQL 通常长这样：
        SELECT date(t.时间), sum(x) FROM … GROUP BY date(t.时间)
    这种结果没法按主键对齐比较——一天只有一行。改写成：
        SELECT 主键, 时间字段, 被聚合的度量字段 FROM …（去掉 GROUP BY）
    行数就和分片预估的明细行数对得上了。
    """
    cols: list[str] = []
    seen = set()
    dropped: list[str] = []
    resolved: list[str] = []

    def push(c):
        c = (c or "").strip()
        if c and c.lower() not in seen:
            seen.add(c.lower())
            cols.append(c)

    # 主键与时间字段是对齐比较的地基：最外层引用不到就先换算成等价表达式，
    # 并保留原名做别名，这样两侧结果的表头仍然对得上
    for k in keys:
        r = resolve_top_level(sql, k, "明细主键")
        if r["derived"]:
            resolved.append(r["note"])
            push(f"{r['expr']} AS `{k.split('.')[-1]}`")
        else:
            push(k)
    if time_col:
        # 时间字段只是顺带带出来看的：换算不出来就不带，别为它中断整个比对
        try:
            r = resolve_top_level(sql, time_col)
            if r["derived"]:
                resolved.append(r["note"])
                push(f"{r['expr']} AS `{time_col.split('.')[-1]}`")
            else:
                push(time_col)
        except ColumnScopeError:
            dropped.append(time_col)
    # 被聚合的度量字段是「顺带看一眼」，藏在子查询里就跳过，不值得为它中断整个比对
    for c in list(_agg_inner_fields(sql)) + list(extra or []):
        if column_scope(sql, c) == "inner":
            dropped.append(c)
            continue
        push(c)
    if not cols:
        raise ValueError("没有可用于明细口径的字段，请至少指定一个主键")

    note = ("已去掉 GROUP BY，改成按主键取明细；"
            "带出的字段是主键、时间字段和原来被聚合的度量字段")
    for r in resolved:
        note += f"。{r}"
    if dropped:
        note += (f"。{len(dropped)} 个度量字段只在子查询里可见（"
                 f"{'、'.join(dropped[:5])}），最外层引用会报 cannot be resolved，已跳过")
    return {"sql": _replace_select_list(sql, ", ".join(cols)),
            "columns": cols,
            "dropped": dropped,
            "note": note}


# 时间分片粒度：单个时段行数超上限时，逐级往下细分
GRAINS = [
    ("天", "DATE({col})"),
    ("小时", "DATE_FORMAT({col}, '%Y-%m-%d %H:00:00')"),
    ("分钟", "DATE_FORMAT({col}, '%Y-%m-%d %H:%i:00')"),
]


def build_distribution_sql(sql: str, time_col: str, grain: int = 0) -> str:
    """按时间字段查各时段的行数分布（在原层次改写 SELECT 与 GROUP BY）

    最外层引用不到 time_col 时，先换算成等价表达式（如别名 year_month → t4.datetime），
    换算不出来才报错——别让用户在「换个字段」和「又报错」之间来回打转。
    """
    col = resolve_top_level(sql, time_col)["expr"]
    expr = GRAINS[min(grain, len(GRAINS) - 1)][1].format(col=col)
    return _replace_select_list(sql, f"{expr} AS _seg, COUNT(1) AS _cnt",
                                keep_group=True, group_expr=expr)


def _seg_end(seg: str, grain_name: str) -> str:
    """把时段标签扩展成该时段的最后一刻。

    分组标签是时段的「起点」：小时段 07:00:00 实际覆盖 07:00:00~07:59:59。
    如果直接拿标签当查询上界，一小时的数据只能查到整点那一秒，
    实取行数会远小于预估——这个坑不填，分片核对永远对不上。
    """
    s = str(seg).strip()
    if grain_name == "天" or len(s) <= 10:
        return f"{s[:10]} 23:59:59"
    if grain_name == "小时":
        return f"{s[:13]}:59:59"
    if grain_name == "分钟":
        return f"{s[:16]}:59"
    return s


def plan_time_slices(dist: list[tuple], batch: int = 500,
                     grain_name: str = "天") -> list[dict]:
    """把「各时段行数」装箱成若干片，每片总行数不超过 batch。

    dist: [(时段字符串, 行数), …] 已按时段升序
    单个时段自己就超过 batch 的，标 over=True，交给调用方降粒度再切。
    """
    slices: list[dict] = []
    cur: list[tuple] = []
    total = 0
    for seg, cnt in dist:
        cnt = int(cnt or 0)
        if cnt <= 0:
            continue
        if cnt > batch:                      # 单个时段就超限，自己一片
            if cur:
                slices.append(_mk_slice(cur, total, batch, grain_name))
                cur, total = [], 0
            slices.append({"start": seg, "end": seg,
                           "end_bound": _seg_end(seg, grain_name),
                           "rows": cnt, "segs": 1,
                           "grain": grain_name, "over": True})
            continue
        if total + cnt > batch and cur:
            slices.append(_mk_slice(cur, total, batch, grain_name))
            cur, total = [], 0
        cur.append((seg, cnt))
        total += cnt
    if cur:
        slices.append(_mk_slice(cur, total, batch, grain_name))
    return slices


def _mk_slice(items: list[tuple], total: int, batch: int, grain: str) -> dict:
    return {"start": items[0][0], "end": items[-1][0],
            "end_bound": _seg_end(items[-1][0], grain),
            "rows": total, "segs": len(items), "grain": grain,
            "over": total > batch}


def _enumerate_segments(start: str, end: str, grain: int = 0) -> list[tuple]:
    """把时间范围按粒度铺成 [(段起, 段止, 段标签)]，标签与 _seg_end 的口径一致"""
    import datetime as dt

    def parse(s, tail):
        s = str(s).strip()
        if len(s) <= 10:
            s = f"{s} {tail}"
        if len(s) <= 16:
            s = f"{s}:00"
        return dt.datetime.strptime(s, "%Y-%m-%d %H:%M:%S")

    lo, hi = parse(start, "00:00:00"), parse(end, "23:59:59")
    step = [dt.timedelta(days=1), dt.timedelta(hours=1), dt.timedelta(minutes=1)][
        min(grain, 2)]
    fmt = ["%Y-%m-%d", "%Y-%m-%d %H:00:00", "%Y-%m-%d %H:%M:00"][min(grain, 2)]
    cur = lo.replace(hour=0, minute=0, second=0) if grain == 0 else \
        lo.replace(minute=0, second=0) if grain == 1 else lo.replace(second=0)
    out = []
    while cur <= hi and len(out) < 2000:          # 上限兜底，别把自己卡死
        nxt = cur + step
        out.append((cur.strftime("%Y-%m-%d %H:%M:%S"),
                    (nxt - dt.timedelta(seconds=1)).strftime("%Y-%m-%d %H:%M:%S"),
                    cur.strftime(fmt)))
        cur = nxt
    return out


def _dist_by_probe(run_sql, sql: str, time_col: str, start: str, end: str,
                   grain: int = 0, progress=None) -> list[tuple]:
    """逐时段各数一次行数，拼出行数分布。

    列只在子查询里可见时，最外层没法 GROUP BY 它，一条 SQL 拿不到分布。
    换个笨一点但一定对的办法：把时间范围铺开，每段用下推后的 WHERE 单独 COUNT 一次。
    段数不多（一个月 31 段、细分一天 24 段），换来的是完全不用在最外层引用这个列。
    """
    segs = _enumerate_segments(start, end, grain)
    out = []
    for i, (s, e, label) in enumerate(segs, 1):
        if progress:
            progress(f"逐段数行数 {i}/{len(segs)}：{label}")
        piece = rewrite_time_filter(sql, time_col, s, e)
        _, rs = run_sql(build_detail_count_sql(piece["sql"]))
        n = int(rs[0][0]) if rs and rs[0] and rs[0][0] is not None else 0
        if n:
            out.append((label, n))
    return out


def plan_full_scan(run_sql, sql: str, time_col: str, start: str, end: str,
                   batch: int = 500, progress=None) -> dict:
    """规划出能覆盖全时间段的分片方案，每片行数都不超过 batch。

    做法（对应「先查总数、再按时间分组、拼接成总时间和总量」的思路）：
      1. 在原层次把时间条件改写到 [start, end]，查明细总行数
      2. 按「天」查各时段行数分布
      3. 装箱：连续时段累加到接近 batch 就切一片
      4. 某一天自己就超过 batch → 对这一天单独降到「小时」，还超就降到「分钟」
      5. 汇总所有片，校验分片行数之和与总量是否一致

    run_sql(sql_text) → (headers, rows)，由调用方注入，方便复用连库客户端。
    """
    def note(msg):
        if progress:
            progress(msg)

    # 作用域体检：能在最外层引用（或换算成表达式）就走一条 GROUP BY 拿分布；
    # 只在子查询里可见的，靠「下推 WHERE + 逐段 COUNT」拿分布，全程不碰最外层
    scope_note, outer_ref = "", True
    try:
        scope_note = resolve_top_level(sql, time_col)["note"]
    except ColumnScopeError:
        outer_ref = False
        blocks = innermost_blocks_with(strip_comments(sql), time_col)
        if not blocks:
            raise
        scope_note = (f"「{time_col}」最外层引用不到，已改用「条件下推到子查询 + "
                      f"逐段数行数」的方式规划分片（涉及 {len(blocks)} 个子查询）")
    day_only = column_day_grain(sql, time_col)
    grand = rewrite_time_filter(sql, time_col, start, end)
    note("查明细层总行数（GROUP BY 之前的真实行数）…")
    _, rows = run_sql(build_detail_count_sql(grand["sql"]))
    total = int(rows[0][0]) if rows else 0
    note(f"明细层共 {total} 行，查原 SQL 聚合后的输出行数…")
    _, arows = run_sql(build_count_sql(grand["sql"]))
    agg_rows = int(arows[0][0]) if arows else 0

    def dist_of(rng_start, rng_end, grain):
        """取某个时间范围内、某个粒度下的行数分布"""
        if not outer_ref:
            return _dist_by_probe(run_sql, sql, time_col, rng_start, rng_end,
                                  grain, note)
        piece = rewrite_time_filter(sql, time_col, rng_start, rng_end)
        _, rs = run_sql(build_distribution_sql(piece["sql"], time_col, grain))
        return [(str(r[0]), int(r[1] or 0)) for r in rs if r[0] is not None]

    note("按天统计行数分布…")
    dist = dist_of(start, end, 0)
    slices = plan_time_slices(dist, batch, GRAINS[0][0])

    # 逐级细分：某片单独超限就换更细的粒度重切这一段。
    # 日粒度字符串列（DATE_FORMAT '%Y%m%d'）本身没有时分秒，再细分只会切出空段
    grain_range = range(1, 1 if day_only else len(GRAINS))
    for grain in grain_range:
        over = [s for s in slices if s.get("over")]
        if not over:
            break
        note(f"有 {len(over)} 段超过 {batch} 行，降到「{GRAINS[grain][0]}」粒度细分…")
        refined: list[dict] = []
        for s in slices:
            if not s.get("over"):
                refined.append(s)
                continue
            sub = dist_of(s["start"], s.get("end_bound") or s["end"], grain)
            refined.extend(plan_time_slices(sub, batch, GRAINS[grain][0])
                           if sub else [s])
        slices = refined

    planned = sum(s["rows"] for s in slices)
    still_over = [s for s in slices if s.get("over")]
    return {
        "total_rows": total,
        "agg_rows": agg_rows,
        "planned_rows": planned,
        "consistent": planned == total,
        "slices": slices,
        "slice_count": len(slices),
        "batch": batch,
        "range": {"start": start, "end": end},
        "rewrite_mode": grand["mode"],
        "rewrite_note": grand["note"] + (f"。{scope_note}" if scope_note else ""),
        "rewritten_sql": grand["sql"],
        "still_over": still_over,
        "is_aggregated": agg_rows != total,
        "warning": (
            f"仍有 {len(still_over)} 段单段超过 {batch} 行（已细到分钟仍超），"
            f"这些段只能取到前 {batch} 行，建议再加别的过滤条件收窄"
            if still_over else
            ("" if planned == total else
             f"分片行数合计 {planned} 与明细总量 {total} 不一致，"
             f"差 {abs(total - planned)} 行。常见原因是时间字段有 NULL，"
             f"这些行落不进任何时间段")),
    }


def fetch_full_scan(run_sql, sql: str, time_col: str, plan: dict,
                    keys: list[str] | None = None, detail: bool = True,
                    progress=None) -> dict:
    """按分片方案逐片取数并拼接成完整结果集。

    detail=True 时先把 SQL 改写成明细口径再逐片取：
    分片是按明细行数规划的，如果还按原聚合口径取，一天只回一行，
    实取行数会远小于预估，两边根本对不上。
    每片都在原 SQL 自己的层次上改写时间条件，不套壳，聚合口径也能跑。
    """
    base_sql, detail_note, detail_cols = sql, "", []
    if detail and plan.get("is_aggregated") and keys:
        d = build_detail_select_sql(sql, keys, time_col)
        base_sql, detail_note, detail_cols = d["sql"], d["note"], d["columns"]

    headers: list[str] = []
    all_rows: list[list] = []
    slice_detail: list[dict] = []
    slices = plan["slices"]
    for i, s in enumerate(slices, 1):
        if progress:
            progress(f"取第 {i}/{len(slices)} 片：{s['start']} ~ {s['end']}"
                     f"（预计 {s['rows']} 行）")
        piece = rewrite_time_filter(base_sql, time_col, s["start"],
                                    s.get("end_bound") or s["end"])
        h, rows = run_sql(piece["sql"])
        if h and not headers:
            headers = h
        all_rows.extend(rows)
        slice_detail.append({"start": s["start"], "end": s["end"],
                             "expect": s["rows"], "got": len(rows),
                             "grain": s.get("grain", ""),
                             "match": len(rows) == s["rows"]})
    mismatch = [d for d in slice_detail if not d["match"]]
    return {
        "headers": headers, "rows": all_rows,
        "fetched": len(all_rows),
        "expected": plan["planned_rows"],
        "slice_detail": slice_detail,
        "mismatch": mismatch,
        "detail_mode": bool(detail_note),
        "detail_note": detail_note,
        "detail_columns": detail_cols,
        "executed_sample": base_sql,
        "note": ("每片实际取到的行数都与预估一致" if not mismatch else
                 f"{len(mismatch)} 片实际行数与预估不一致，"
                 f"多半是取数期间数据还在变动，或该片撞上了 {plan['batch']} 行上限"),
    }



def suggest_time_range(sql: str) -> dict:
    """从 SQL 里已有的时间字面量猜一个默认范围，省得每次手填。

    要滤掉 1970-01-01 这类 NULL 兜底哨兵值——它们是占位符不是业务时间，
    直接拿来当查询范围会一条数据都查不到。
    """
    import datetime as dt
    today = dt.date.today()
    sentinel = {"1970-01-01", "1900-01-01", "0001-01-01",
                "9999-12-31", "2999-12-31", "1899-12-30"}
    raw = re.findall(r"'(\d{4}-\d{2}-\d{2})", strip_comments(sql))
    good = []
    for d in raw:
        if d in sentinel:
            continue
        try:
            y = int(d[:4])
        except ValueError:
            continue
        if 2000 <= y <= today.year + 5:          # 合理业务年份区间
            good.append(d)
    if good:
        ds = sorted(set(good))
        return {"start": ds[0], "end": ds[-1],
                "source": f"取自 SQL 中的日期字面量（已滤掉 {len(raw) - len(good)} 个哨兵值）"}
    hint = "SQL 里没有可用日期字面量" if not raw else \
           f"SQL 里的 {len(raw)} 个日期都是哨兵/越界值（如 1970-01-01）"
    return {"start": str(today.replace(day=1)), "end": str(today),
            "source": f"{hint}，默认本月至今"}


# ══════════════════════════════════════════════════════════════
# 第二层：数据层明细比较 + 排查日志
# ══════════════════════════════════════════════════════════════


def _resolve_column(name: str, headers: list[str]) -> str | None:
    """把用户填的字段名匹配到结果集里真实的列名。

    用户填的常带表别名（a.order_id），但数据库返回的表头是裸列名
    （order_id）；反过来也可能。这里按精确 → 去前缀 → 忽略大小写
    逐级放宽匹配，避免「主键字段不存在」这种明显能对上却报错的情况。
    """
    if not name:
        return None
    want = name.strip().replace("`", "")
    bare = want.split(".")[-1]
    lower = {h.lower(): h for h in headers}
    # 精确
    if want in headers:
        return want
    if want.lower() in lower:
        return lower[want.lower()]
    # 用户带前缀、结果集是裸名
    if bare in headers:
        return bare
    if bare.lower() in lower:
        return lower[bare.lower()]
    # 结果集带前缀、用户填裸名
    for h in headers:
        if h.replace("`", "").split(".")[-1].lower() == bare.lower():
            return h
    return None


def compare_details(headers_a: list[str], rows_a: list[list],
                    headers_b: list[str], rows_b: list[list],
                    keys: list[str], sql_a: str = "", sql_b: str = "") -> dict:
    """按主键对齐两边明细，逐字段找差异，并产出可读的排查日志。

    日志每条都带主键与出错原因，便于定位一条异常数据并逐层检查过滤条件。
    给了 sql_a/sql_b 时，主键可以填子查询里的原始写法（如 `t7.datetime`），
    自动顺着别名链换算成最外层真实输出的列名。
    """
    # 一侧完全查不到数据，这本身就是结论，不是配置错误。
    # 空结果集连列名都带不回来，走下面的主键解析必然报「主键字段不都存在」，
    # 等于把「这一侧一条都没有」这个最关键的发现说成了参数配错。
    if not headers_a or not headers_b:
        return _one_side_empty(headers_a, rows_a, headers_b, rows_b, keys)

    ia = {h: i for i, h in enumerate(headers_a)}
    ib = {h: i for i, h in enumerate(headers_b)}
    ka = resolve_keys(keys, sql_a, headers_a)
    kb = resolve_keys(keys, sql_b, headers_b)
    missing = sorted(set(ka["missing"]) | set(kb["missing"]))
    if missing:
        raise ValueError(_missing_key_msg(missing, headers_a, headers_b,
                                          sql_a, sql_b))
    keys_a = [ka["resolved"][k] for k in keys]
    keys_b = [kb["resolved"][k] for k in keys]
    key_notes = ka["notes"] + [n for n in kb["notes"] if n not in ka["notes"]]
    # 比较列：按裸名取两侧交集，排除主键
    def barename(h):
        return h.replace("`", "").split(".")[-1].lower()

    key_bare = {barename(k) for k in keys_a} | {barename(k) for k in keys_b}
    b_by_bare = {barename(h): h for h in headers_b}
    cmp_pairs = [(h, b_by_bare[barename(h)]) for h in headers_a
                 if barename(h) in b_by_bare and barename(h) not in key_bare]

    def norm_cell(v):
        if v is None:
            return ""
        s = str(v).strip()
        if s.upper() in ("NULL", "NONE"):
            return ""
        try:                                   # 1 与 1.0 视为相同
            return f"{float(s):.10g}"
        except ValueError:
            return s

    def index(rows, idx, key_cols):
        m: dict[tuple, list] = {}
        for r in rows:
            k = tuple(str(r[idx[c]]) if idx[c] < len(r) and r[idx[c]] is not None
                      else "" for c in key_cols)
            m.setdefault(k, []).append(r)
        return m

    ma, mb = index(rows_a, ia, keys_a), index(rows_b, ib, keys_b)
    logs: list[dict] = []

    # 主键重复 = 一对多倍乘的直接证据
    for tag, m in (("A", ma), ("B", mb)):
        for k, rs in m.items():
            if len(rs) > 1:
                logs.append({
                    "level": "high", "key": "+".join(k), "column": "",
                    "reason": f"{tag} 侧主键重复 {len(rs)} 行",
                    "detail": "同一主键出现多行，存在一对多倍乘。"
                              "先按该主键 GROUP BY … HAVING COUNT(1)>1 定位是哪个 JOIN 放大的",
                })

    only_a = [k for k in ma if k not in mb]
    only_b = [k for k in mb if k not in ma]
    for k in only_a[:200]:
        logs.append({"level": "high", "key": "+".join(k), "column": "",
                     "reason": "只在 A 侧出现",
                     "detail": "B 侧丢了这条。优先怀疑 INNER JOIN 未匹配、"
                               "或 B 侧多了过滤条件"})
    for k in only_b[:200]:
        logs.append({"level": "high", "key": "+".join(k), "column": "",
                     "reason": "只在 B 侧出现",
                     "detail": "A 侧没有这条。优先怀疑 A 侧过滤更严、"
                               "或 B 侧 JOIN 放大产生了新行"})

    diffs, matched = [], 0
    null_flip = 0
    for k in ma:
        if k not in mb:
            continue
        for ra_row, rb_row in zip(ma[k], mb[k]):
            matched += 1
            for ca, cb in cmp_pairs:
                va = ra_row[ia[ca]] if ia[ca] < len(ra_row) else None
                vb = rb_row[ib[cb]] if ib[cb] < len(rb_row) else None
                na, nb = norm_cell(va), norm_cell(vb)
                if na == nb:
                    continue
                diffs.append({"key": "+".join(k), "column": ca,
                              "a": "" if va is None else str(va),
                              "b": "" if vb is None else str(vb)})
                if na == "" or nb == "":
                    null_flip += 1
                    side = "B" if nb == "" else "A"
                    logs.append({
                        "level": "high", "key": "+".join(k), "column": ca,
                        "reason": f"{side} 侧该字段为空",
                        "detail": "一边有值一边为空，典型是 LEFT JOIN 右表没匹配上。"
                                  "检查该字段来源表的关联键，或补 ifnull 兜底",
                    })
                else:
                    logs.append({
                        "level": "mid", "key": "+".join(k), "column": ca,
                        "reason": "同主键字段值不同",
                        "detail": f"A={va} / B={vb}。若是金额或数量，"
                                  f"检查是否 JOIN 倍乘导致重复累加",
                    })

    order = {"high": 0, "mid": 1, "low": 2}
    logs.sort(key=lambda x: order.get(x["level"], 9))
    for n in key_notes:
        logs.append({"level": "low", "key": "", "column": "",
                     "reason": "主键已自动换算", "detail": n})
    verdict = _verdict(len(only_a), len(only_b), len(diffs), null_flip, matched)
    # 页面只回传有限数量的差异明细。若某主键还有被截断的差异，就不能把它
    # 放进档案复核范围，否则历史条目会因为“本轮未回传”而被误判为已修复。
    omitted_pks = {
        d["key"] for d in diffs[2000:]
    } | {
        "+".join(k) for k in only_a[500:]
    } | {
        "+".join(k) for k in only_b[500:]
    }
    all_pks = ["+".join(k) for k in dict.fromkeys([*ma, *mb])]
    safe_pks = [pk for pk in all_pks if pk not in omitted_pks]
    pks_in_scope = safe_pks[:MAX_ARCHIVE_SCOPE_KEYS]
    archive_scope_truncated = bool(omitted_pks) or len(safe_pks) > len(pks_in_scope)
    return {
        "stats": {"rows_a": len(rows_a), "rows_b": len(rows_b),
                  "keys": keys, "resolved_keys": {"a": keys_a, "b": keys_b},
                  "matched": matched,
                  "only_a": len(only_a), "only_b": len(only_b),
                  "cmp_cols": len(cmp_pairs), "diff_cells": len(diffs),
                  "null_flip": null_flip},
        "key_notes": key_notes,
        "only_a": ["+".join(k) for k in only_a[:500]],
        "only_b": ["+".join(k) for k in only_b[:500]],
        "pks_in_scope": pks_in_scope,
        "archive_scope_truncated": archive_scope_truncated,
        "diffs": diffs[:2000], "diffs_total": len(diffs),
        "logs": logs[:1000], "logs_total": len(logs),
        "verdict": verdict,
    }


def _one_side_empty(headers_a: list[str], rows_a: list[list],
                    headers_b: list[str], rows_b: list[list],
                    keys: list[str]) -> dict:
    """有一侧一行都没查到时的结论。

    照常给出「只在某侧出现」的主键清单，方便直接拿去核对；
    最容易撞上的两种原因是这一侧压根没这段数据（采集断了），
    或者时间条件下推后落到了不该落的那一层。
    """
    side = "A" if headers_a else "B"          # 有数据的那一侧
    live_h = headers_a or headers_b
    live_r = rows_a or rows_b
    pos = {h: i for i, h in enumerate(live_h)}
    resolved = [_resolve_column(k, live_h) for k in keys]   # 保持主键填写顺序
    idx = [pos[h] for h in resolved if h in pos]
    ks = ["+".join(str(r[i]) if i < len(r) and r[i] is not None else ""
                   for i in idx) for r in live_r] if idx else []
    empty = "B" if side == "A" else "A"
    if not live_r:
        text, nxt = "两侧都没有查到数据", ("时间范围可能整体落在数据之外，"
                                          "或时间条件下推到了错误的子查询层，先单独跑一侧确认")
        lv = "warn"
    else:
        text = f"{empty} 侧在该范围内没有任何数据（{side} 侧有 {len(live_r)} 行）"
        nxt = (f"这不是字段对不上，是 {empty} 侧真的空。"
               f"先查 {empty} 侧的采集任务在这个时间段是否停了；"
               f"再确认 {empty} 侧表里该范围的原始数据是否存在")
        lv = "diff"
    logs = [{"level": "high", "key": "", "column": "",
             "reason": f"{empty} 侧结果为空",
             "detail": f"{side} 侧 {len(live_r)} 行，{empty} 侧 0 行。"
                       f"两侧 SQL 若只差数据源，优先怀疑 {empty} 侧数据未同步"}]
    logs += [{"level": "high", "key": k, "column": "",
              "reason": f"只在 {side} 侧出现", "detail": f"{empty} 侧整段为空"}
             for k in ks[:200]]
    return {
        "stats": {"rows_a": len(rows_a), "rows_b": len(rows_b),
                  "keys": keys, "resolved_keys": {"a": [], "b": []},
                  "matched": 0,
                  "only_a": len(rows_a) if side == "A" else 0,
                  "only_b": len(rows_b) if side == "B" else 0,
                  "cmp_cols": 0, "diff_cells": 0, "null_flip": 0},
        "only_a": ks[:500] if side == "A" else [],
        "only_b": ks[:500] if side == "B" else [],
        "pks_in_scope": list(dict.fromkeys(ks[:500]))[:MAX_ARCHIVE_SCOPE_KEYS],
        "archive_scope_truncated": len(ks) > 500,
        "diffs": [], "diffs_total": 0,
        "logs": logs[:1000], "logs_total": len(logs),
        "verdict": {"level": lv, "text": text, "next": nxt},
    }


def _missing_key_msg(missing: list[str], headers_a: list[str],
                     headers_b: list[str], sql_a: str, sql_b: str) -> str:
    """主键换算不出来时，直接把能用的候选摆出来。

    只报「字段不都存在」等于把问题丢回给用户——他并不知道最外层输出了什么，
    更不知道该填哪个。这里把两侧共有的列、以及推荐的主键组合一起给出。
    """
    ba = {h.replace("`", "").split(".")[-1].lower(): h for h in headers_a}
    bb = {h.replace("`", "").split(".")[-1].lower() for h in headers_b}
    common = [ba[k] for k in ba if k in bb]
    lines = [f"主键「{'、'.join(missing)}」在最外层结果里找不到，"
             f"顺着别名链也没换算出来。"]
    if common:
        lines.append(f"两侧都有的列（可以直接填）：{'、'.join(common)}")
    try:
        # suggest_primary_keys 给的是 SQL 里的写法（t7.datetime），
        # 结果集里是别名（year_month），推荐前得先换算一遍，否则一个都对不上
        picks = []
        for k in suggest_primary_keys(sql_a)[:6]:
            r = resolve_keys([k["column"]], sql_a, headers_a)
            hit = r["resolved"].get(k["column"])
            if hit and _resolve_column(hit, headers_b) and hit not in picks:
                picks.append(hit)
        if picks:
            lines.append(f"推荐主键：{'，'.join(picks[:3])}")
    except Exception:
        pass
    only_a = [h for h in headers_a
              if h.replace("`", "").split(".")[-1].lower() not in bb]
    if only_a:
        lines.append(f"只有 A 侧有的列：{'、'.join(only_a)}")
    return " ".join(lines)


def _verdict(only_a: int, only_b: int, diff_cells: int,
             null_flip: int, matched: int) -> dict:
    """把比对结果归纳成一句结论 + 下一步建议"""
    if not (only_a or only_b or diff_cells):
        return {"level": "ok", "text": f"两侧 {matched} 行明细完全一致",
                "next": "该时间窗内逻辑等价。可换其他时间窗再抽验，"
                        "尤其覆盖月初月末与有撤销/重算记录的日期"}
    causes = []
    if only_a and not only_b:
        causes.append("B 侧丢数据：优先查 INNER JOIN 未匹配、或 B 多了过滤条件")
    if only_b and not only_a:
        causes.append("B 侧多数据：优先查 JOIN 一对多放大、或 A 多了过滤条件")
    if only_a and only_b:
        causes.append("两侧各有独有主键：先核对主键定义与过滤条件是否同口径")
    if null_flip:
        causes.append(f"{null_flip} 处一边为空：LEFT JOIN 右表未匹配的典型症状")
    if diff_cells and not null_flip:
        causes.append("字段值不同但都非空：核对计算口径、单位、聚合方式")
    return {"level": "diff",
            "text": f"发现差异：仅A {only_a} / 仅B {only_b} / 字段差异 {diff_cells} 处",
            "next": "；".join(causes)}


# ══════════════════════════════════════════════════════════════
# 排查辅助 SQL 生成
# ══════════════════════════════════════════════════════════════


def build_probe_sqls(sql: str, pk: str = "") -> list[dict]:
    """按内置通用排查规则，为这段 SQL 生成一组可直接执行的验证 SQL"""
    inner = strip_comments(sql).strip().rstrip(";")
    p = parse_sql(sql)
    probes: list[dict] = []

    probes.append({
        "no": "P0", "title": "总行数",
        "why": "先看整体量级，和对照口径比一下是多了还是少了",
        "sql": f"SELECT COUNT(1) AS cnt FROM (\n{inner}\n) AS t",
    })

    if pk:
        probes.append({
            "no": "P1", "title": f"主键是否重复（一对多倍乘检测）",
            "why": "同一主键出现多行说明某个 JOIN 把数据放大了，指标会虚高",
            "sql": (f"SELECT {pk}, COUNT(1) AS c FROM (\n{inner}\n) AS t\n"
                    f"GROUP BY {pk} HAVING COUNT(1) > 1\nORDER BY c DESC LIMIT 100"),
        })
        probes.append({
            "no": "P2", "title": "去重后行数对比",
            "why": "与 P0 对比：不一致即存在重复，差额就是被放大的行数",
            "sql": (f"SELECT COUNT(1) AS 总行数, COUNT(DISTINCT {pk}) AS 去重行数\n"
                    f"FROM (\n{inner}\n) AS t"),
        })

    # 逐个 JOIN 生成关联完整性检测
    for i, j in enumerate(p["joins"], 1):
        if not j["on"]:
            continue
        if j["type"].startswith("INNER"):
            probes.append({
                "no": f"J{i}", "title": f"INNER JOIN 丢数检测：{j['table']}",
                "why": "INNER JOIN 会丢掉右表匹配不上的记录，这是数据变少最常见原因。"
                       "把它临时改成 LEFT JOIN，统计右表为空的行数就是丢掉的量",
                "sql": (f"-- 手动操作：把下面这句里 {j['table']} 的 INNER JOIN 改成 LEFT JOIN，\n"
                        f"-- 再统计关联键为 NULL 的行数，即被 INNER JOIN 丢掉的记录\n"
                        f"-- 原 ON 条件: {j['on']}\n"
                        f"SELECT COUNT(1) AS 丢数行数 FROM (\n{inner}\n) AS t\n"
                        f"-- WHERE <{j['table']} 的关联字段> IS NULL"),
            })
        elif j["type"].startswith("LEFT"):
            probes.append({
                "no": f"J{i}", "title": f"LEFT JOIN 未匹配率：{j['table']}",
                "why": "LEFT JOIN 匹配不上会让右表字段为 NULL，"
                       "若外层拿它做过滤或聚合，结果会异常",
                "sql": (f"-- 右表: {j['table']}   ON: {j['on']}\n"
                        f"-- 把 <右表字段> 换成该表实际带出的任一字段\n"
                        f"SELECT COUNT(1) AS 总行数,\n"
                        f"       SUM(CASE WHEN <右表字段> IS NULL THEN 1 ELSE 0 END) AS 未匹配行数\n"
                        f"FROM (\n{inner}\n) AS t"),
            })
        probes.append({
            "no": f"M{i}", "title": f"右表一对多检测：{j['table']}",
            "why": "右表在关联键上不唯一，JOIN 后主表行数会成倍增加",
            "sql": (f"-- 关联条件: {j['on']}\n"
                    f"-- 把 <关联键> 换成 ON 里右表侧的字段\n"
                    f"SELECT <关联键>, COUNT(1) AS c\n"
                    f"FROM {j['table']}\nGROUP BY <关联键>\n"
                    f"HAVING COUNT(1) > 1 ORDER BY c DESC LIMIT 50"),
        })

    # 逐层剥离过滤条件，检查过滤条件造成的数据变化
    if p["conditions"]:
        lines = []
        for i, c in enumerate(p["conditions"], 1):
            lines.append(f"-- 第{i}层: {c}")
        probes.append({
            "no": "F1", "title": "逐层加过滤条件对数",
            "why": "从无条件开始，一次只加一个条件看行数怎么掉，"
                   "哪一步掉得离谱就是那个条件的问题",
            "sql": ("-- 依次执行，每次只加一个 WHERE 条件，记录行数变化\n"
                    + "\n".join(lines)
                    + f"\n\nSELECT COUNT(1) FROM (\n{inner}\n) AS t"),
        })

    return probes


# ══════════════════════════════════════════════════════════════
# 依赖表分层 + 脚本层归因
# ══════════════════════════════════════════════════════════════


def table_lineage(sql: str) -> dict:
    """列出依赖表并按数仓分层归类。

    把「表 → 所属层 → 建议检查的处理环节」直接给出来。
    """
    tables = extract_tables(normalize(sql))
    groups: dict[str, list[dict]] = {}
    unknown = []
    for t in tables:
        name = t["name"]
        db = t["db"].lower().rstrip("_v2").rstrip("_old").rstrip("_cold")
        bare = name.split(".")[-1].lower()
        layer_key = None
        if db and db in LAYERS:
            layer_key = db
        else:
            for k in LAYERS:                      # 无库名前缀时按表名前缀猜
                if bare.startswith(k + "_"):
                    layer_key = k
                    break
        if layer_key:
            label, advice = LAYERS[layer_key]
            groups.setdefault(label, []).append(
                {"table": name, "alias": t["alias"], "advice": advice})
        else:
            unknown.append({"table": name, "alias": t["alias"], "advice": ""})

    ordered = []
    for k, (label, advice) in LAYERS.items():
        if label in groups:
            ordered.append({"layer": label, "advice": advice,
                            "tables": groups[label]})
    if unknown:
        ordered.append({"layer": "未识别层级", "tables": unknown,
                        "advice": "可能是别名、CTE 名或自定义层级；请查看对应任务配置和上游依赖。"})

    return {
        "total": len(tables), "layers": ordered,
        "trace_order": [g["layer"] for g in ordered],
        "hint": ("排查顺序由下往上：先确认最底层（STG/ODS）有没有数据，"
                 "再逐层往上看是哪一层的脚本把数据弄丢或弄错。"
                 "某一层数据不对，就去查产出这一层的脚本"),
    }
