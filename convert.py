#!/usr/bin/env python3
"""Matrix 工具箱 · 格式转换层

- excel_to_insert()  Excel + 建表语句 → INSERT INTO 语句（按 DDL 字段类型正确加引号）
- excel_to_pdf()     Excel → PDF（fpdf2 + macOS 中文字体，不依赖 Office/LibreOffice）
- excel_to_word()    Excel → Word（python-docx）
- image_to_rows()    图片 → 表格二维数组（macOS Vision OCR，按文字坐标还原行列）

选型说明：本机没有 LibreOffice/Excel/Numbers，所以 PDF/Word 都用纯 Python 生成；
OCR 走 macOS 原生 Vision（中文识别质量好、零第三方依赖），Swift 源码随包分发，
首次使用时就地编译并缓存。
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

import matrix_core as core

# ══════════════════════════════════════════════════════════════
# Excel → INSERT INTO
# ══════════════════════════════════════════════════════════════

# 需要加引号的类型；数值/布尔类型直接裸写
QUOTED_TYPES = ("char", "varchar", "string", "text", "date", "datetime",
                "timestamp", "time", "json", "binary", "blob")
NUMERIC_TYPES = ("int", "bigint", "smallint", "tinyint", "largeint", "decimal",
                 "double", "float", "numeric", "boolean", "bool")


def parse_ddl(ddl: str) -> dict:
    """从建表语句里解析出表名与字段（名称 + 类型）。

    兼容 StarRocks/MySQL 的 CREATE TABLE，容忍反引号、中文名、COMMENT。
    """
    text = core_strip_comments(ddl)
    m = re.search(r"CREATE\s+(?:EXTERNAL\s+)?TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?"
                  r"([`\w\u4e00-\u9fff.]+)\s*\(", text, re.I)
    if not m:
        raise ValueError("没找到 CREATE TABLE ... ( ，请贴完整的建表语句")
    table = m.group(1).replace("`", "")
    # 取最外层括号内的字段定义区
    start = m.end()
    depth, i = 1, start
    while i < len(text) and depth:
        if text[i] == "(":
            depth += 1
        elif text[i] == ")":
            depth -= 1
        i += 1
    body = text[start:i - 1]

    cols = []
    for line in _split_top_commas(body):
        line = line.strip()
        if not line:
            continue
        # 跳过表级约束/索引定义
        if re.match(r"(PRIMARY|UNIQUE|DUPLICATE|AGGREGATE)\s+KEY\s*\(", line, re.I):
            continue
        if re.match(r"(INDEX|KEY|CONSTRAINT|FOREIGN)\b", line, re.I):
            continue
        cm = re.match(r"[`\"]?([\w\u4e00-\u9fff]+)[`\"]?\s+([A-Za-z]+)", line)
        if not cm:
            continue
        name, ctype = cm.group(1), cm.group(2).lower()
        cols.append({
            "name": name, "type": ctype,
            "quoted": _needs_quote(ctype),
            "nullable": not re.search(r"\bNOT\s+NULL\b", line, re.I),
        })
    if not cols:
        raise ValueError("建表语句里没解析出任何字段")
    return {"table": table, "columns": cols}


def _needs_quote(ctype: str) -> bool:
    if any(ctype.startswith(t) for t in NUMERIC_TYPES):
        return False
    if any(t in ctype for t in QUOTED_TYPES):
        return True
    return True                      # 认不出的类型按字符串处理，更安全


def _split_top_commas(text: str) -> list[str]:
    """按最外层逗号切分，忽略括号内与字符串内的逗号"""
    out, depth, last, i = [], 0, 0, 0
    n = len(text)
    while i < n:
        c = text[i]
        if c in "'\"":
            q = c
            i += 1
            while i < n and not (text[i] == q and text[i - 1] != "\\"):
                i += 1
        elif c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
        elif c == "," and depth == 0:
            out.append(text[last:i])
            last = i + 1
        i += 1
    out.append(text[last:])
    return out


def core_strip_comments(sql: str) -> str:
    """复用 SQL 层的注释清理，避免重复实现"""
    import sql_tools
    return sql_tools.strip_comments(sql)


def _sql_literal(value, col: dict) -> str:
    """把单元格值转成 SQL 字面量"""
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return "NULL"
    if isinstance(value, bool):
        return "1" if value else "0"
    s = str(value).strip()
    if s.upper() in ("NULL", "\\N"):
        return "NULL"
    if not col["quoted"]:
        # 数值列：Excel 常把整数读成 100.0，去掉多余小数
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        try:
            float(s)
            return s
        except ValueError:
            return "'" + s.replace("\\", "\\\\").replace("'", "\\'") + "'"
    if isinstance(value, float) and value.is_integer():
        s = str(int(value))
    return "'" + s.replace("\\", "\\\\").replace("'", "\\'") + "'"


def excel_to_insert(xlsx: Path, ddl: str, batch: int = 500,
                    only_matched: bool = True, sheet=None) -> dict:
    """Excel + 建表语句 → INSERT 语句

    Excel 表头与 DDL 字段按名称匹配（忽略大小写与首尾空格）。
    """
    meta_ddl = parse_ddl(ddl)
    headers, rows, meta = core.read_sheet_meta(xlsx, sheet)
    if not headers:
        raise ValueError("没读到表头。" + "；".join(meta["warnings"])
                         + f"\n工作表清单: {[s['name'] for s in meta['sheets']]}")
    col_by_name = {c["name"].lower(): c for c in meta_ddl["columns"]}

    used, unmatched_header, positions = [], [], []
    for idx, h in enumerate(headers):
        key = str(h).strip().lower()
        c = col_by_name.get(key)
        if c:
            used.append(c)
            positions.append(idx)
        else:
            unmatched_header.append(h)
    if not used:
        raise ValueError(f"Excel 表头与建表语句字段没有任何匹配。\n"
                         f"Excel 表头: {headers[:10]}\n"
                         f"DDL 字段: {[c['name'] for c in meta_ddl['columns']][:10]}")
    missing_cols = [c["name"] for c in meta_ddl["columns"]
                    if c["name"].lower() not in
                    {str(h).strip().lower() for h in headers}]

    col_list = ", ".join(f"`{c['name']}`" for c in used)
    statements, values_buf, skipped = [], [], 0
    for r in rows:
        if all(v is None or str(v).strip() == "" for v in r):
            skipped += 1
            continue
        vals = []
        for c, pos in zip(used, positions):
            v = r[pos] if pos < len(r) else None
            vals.append(_sql_literal(v, c))
        values_buf.append("(" + ", ".join(vals) + ")")
        if len(values_buf) >= batch:
            statements.append(_wrap_insert(meta_ddl["table"], col_list, values_buf))
            values_buf = []
    if values_buf:
        statements.append(_wrap_insert(meta_ddl["table"], col_list, values_buf))

    data_rows = len(rows) - skipped
    return {
        "table": meta_ddl["table"],
        "sql": "\n\n".join(statements),
        "stmt_count": len(statements),
        "row_count": data_rows,
        "skipped_empty": skipped,
        "matched_columns": [c["name"] for c in used],
        "unmatched_header": unmatched_header,
        "missing_columns": missing_cols,
        "batch": batch,
        "sheet": meta["sheet"],
        "sheets": [s["name"] for s in meta["sheets"]],
        "warnings": meta["warnings"],
    }


def _wrap_insert(table: str, col_list: str, values: list[str]) -> str:
    quoted_table = ".".join(
        "`" + part.replace("`", "``") + "`"
        for part in table.split(".") if part
    )
    if not quoted_table:
        raise ValueError("目标表名不能为空")
    return (f"INSERT INTO {quoted_table}\n  ({col_list})\nVALUES\n  "
            + ",\n  ".join(values) + ";")


# ══════════════════════════════════════════════════════════════
# Excel → PDF
# ══════════════════════════════════════════════════════════════

CJK_FONTS = ("/Library/Fonts/Arial Unicode.ttf",
             "/System/Library/Fonts/Supplemental/Songti.ttc",
             "/System/Library/Fonts/Hiragino Sans GB.ttc")


def _pick_font() -> str | None:
    for f in CJK_FONTS:
        if Path(f).exists():
            return f
    return None


def _prepare_font() -> tuple[str, bool]:
    """准备可用的中文字体，返回 (字体路径, 是否老版 pyfpdf 需要 uni=True)。

    老版 pyfpdf 用 uni=True 时会把解析缓存 .pkl 写到「字体文件所在目录」，
    而 /Library/Fonts 不可写，所以先把字体复制到自己的缓存目录再用。
    只接受 .ttf：老版 pyfpdf 不支持 .ttc 字体集合。
    """
    import inspect
    from fpdf import FPDF
    needs_uni = "uni" in inspect.signature(FPDF.add_font).parameters

    src = None
    for f in CJK_FONTS:
        if Path(f).exists() and (f.lower().endswith(".ttf") or not needs_uni):
            src = Path(f)
            break
    if src is None:
        raise RuntimeError(
            "找不到可用的中文字体。老版 pyfpdf 只支持 .ttf，"
            f"已尝试: {CJK_FONTS}")

    if not needs_uni:
        return str(src), False

    font_dir = core.CACHE_DIR / "fonts"
    font_dir.mkdir(parents=True, exist_ok=True)
    dst = font_dir / src.name
    if not dst.exists() or dst.stat().st_size != src.stat().st_size:
        import shutil
        # 用 copyfile 而非 copy2：系统字体带特殊 chflags，复制元数据会被拒
        shutil.copyfile(src, dst)
    return str(dst), True


def excel_to_pdf(xlsx: Path, out: Path, landscape: bool = True,
                 max_rows: int = 800, sheet=None) -> dict:
    """Excel → PDF。表格按列宽自适应分页，中文用系统字体嵌入。

    同时兼容老版 pyfpdf(1.7.x) 与 fpdf2：两者 add_font/cell 的签名不同。

    性能注意：老版 pyfpdf 是纯 Python 实现，每个单元格都要做字体子集映射，
    3000 行 × 21 列（6 万格）要跑两分半。默认只出 800 行，
    需要更多行请显式调大 max_rows 并接受相应耗时。
    """
    from fpdf import FPDF

    font_path, needs_uni = _prepare_font()

    headers, rows, meta = core.read_sheet_meta(xlsx, sheet)
    if not headers and not rows:
        raise ValueError("没读到任何内容。" + "；".join(meta["warnings"])
                         + f"\n工作表清单: {[s['name'] for s in meta['sheets']]}")
    total_rows = len(rows)
    truncated = total_rows > max_rows
    rows = rows[:max_rows]

    pdf = FPDF(orientation="L" if landscape else "P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)
    if needs_uni:
        pdf.add_font("cjk", "", font_path, uni=True)
    else:
        pdf.add_font("cjk", "", font_path)
    pdf.add_page()

    avail = pdf.w - 2 * pdf.l_margin

    def cell_text(v):
        return "" if v is None else str(v)

    # 列宽：按内容长度加权分配（中文算 2 个宽度），并限制上下限
    ncol = max(len(headers), max((len(r) for r in rows), default=0)) or 1
    weights = []
    for i in range(ncol):
        vals = [cell_text(headers[i]) if i < len(headers) else ""]
        vals += [cell_text(r[i]) for r in rows[:200] if i < len(r)]
        w = max((len(v) + sum(1 for ch in v if ord(ch) > 127) for v in vals),
                default=4)
        weights.append(min(max(w, 4), 40))
    total_w = sum(weights)
    widths = [max(avail * w / total_w, 8) for w in weights]
    scale = avail / sum(widths)
    widths = [w * scale for w in widths]

    pdf.set_font("cjk", size=8)
    line_h = 5.0
    fit_cache: dict = {}          # 表格里重复值多，缓存能省掉大量宽度测量

    def row_cells(values, height, fill, border=1):
        """逐格输出后手动换行，兼容两个版本的 cell 签名"""
        for i, w in enumerate(widths):
            txt = cell_text(values[i]) if i < len(values) else ""
            pdf.cell(w, height, _fit(pdf, txt, w, fit_cache), border, 0, "L", fill)
        pdf.ln(height)

    def draw_header():
        pdf.set_font("cjk", size=8)
        pdf.set_fill_color(47, 84, 150)
        pdf.set_text_color(255, 255, 255)
        for i, w in enumerate(widths):
            txt = cell_text(headers[i]) if i < len(headers) else ""
            pdf.cell(w, line_h + 1, _fit(pdf, txt, w, fit_cache), 1, 0, "C", True)
        pdf.ln(line_h + 1)
        pdf.set_text_color(0, 0, 0)

    if headers:
        draw_header()
    fill = False
    for r in rows:
        if pdf.get_y() + line_h > pdf.h - pdf.b_margin:
            pdf.add_page()
            if headers:
                draw_header()
        pdf.set_fill_color(245, 246, 248)
        row_cells(r, line_h, fill)
        fill = not fill

    out.parent.mkdir(parents=True, exist_ok=True)
    pdf.output(str(out))
    return {"file": out.name, "path": str(out), "rows": len(rows),
            "cols": ncol, "truncated": truncated,
            "size": f"{out.stat().st_size / 1024:.0f} KB",
            "font": Path(font_path).name,
            "sheet": meta["sheet"],
            "sheets": [s["name"] for s in meta["sheets"]],
            "total_rows": total_rows,
            "warnings": meta["warnings"] + (
                [f"源数据 {total_rows} 行，只出了前 {max_rows} 行。"
                 f"PDF 渲染是纯 Python 实现，每千行约 40 秒，"
                 f"行数多请直接用 Excel 或先筛选再转"] if truncated else [])}


def _fit(pdf, text: str, width: float, cache: dict | None = None) -> str:
    """超宽文本截断加省略号，避免串格。

    逐字符缩短会对每个单元格反复测量宽度，几万个格子能跑上两分钟。
    这里先按宽度比例估算该保留多少字符，再微调几次；
    并对「文本+列宽」组合做缓存——数据表里重复值很多（类别名、状态名），
    缓存命中率很高。
    """
    text = text.replace("\n", " ").replace("\r", " ")
    if not text:
        return ""
    key = (text, round(width, 2))
    if cache is not None and key in cache:
        return cache[key]
    limit = width - 2
    w = pdf.get_string_width(text)
    if w <= limit:
        out = text
    else:
        keep = max(1, int(len(text) * limit / w)) if w else 1
        out = text[:keep]
        # 估算偏长就收，偏短就放，正常一两次就收敛
        while keep > 1 and pdf.get_string_width(out + "…") > limit:
            keep -= 1
            out = text[:keep]
        while keep < len(text) and pdf.get_string_width(text[:keep + 1] + "…") <= limit:
            keep += 1
            out = text[:keep]
        out += "…"
    if cache is not None:
        cache[key] = out
    return out


# ══════════════════════════════════════════════════════════════
# Excel → Word
# ══════════════════════════════════════════════════════════════


def excel_to_word(xlsx: Path, out: Path, landscape: bool = True,
                  max_rows: int = 3000, title: str = "", sheet=None) -> dict:
    """Excel → Word（.docx），表头加底色，正文隔行浅色"""
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt

    headers, rows, meta = core.read_sheet_meta(xlsx, sheet)
    if not headers and not rows:
        raise ValueError("没读到任何内容。" + "；".join(meta["warnings"])
                         + f"\n工作表清单: {[s['name'] for s in meta['sheets']]}")
    truncated = len(rows) > max_rows
    rows = rows[:max_rows]

    doc = Document()
    sec = doc.sections[0]
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = sec.page_height, sec.page_width

    if title:
        h = doc.add_paragraph(title)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = h.runs[0]
        run.bold = True
        run.font.size = Pt(14)

    ncol = max(len(headers), max((len(r) for r in rows), default=0)) or 1
    table = doc.add_table(rows=1 if headers else 0, cols=ncol)
    table.style = "Table Grid"

    def shade(cell, color: str):
        el = OxmlElement("w:shd")
        el.set(qn("w:fill"), color)
        cell._tc.get_or_add_tcPr().append(el)

    if headers:
        hdr = table.rows[0].cells
        for i in range(ncol):
            txt = str(headers[i]) if i < len(headers) else ""
            cell = hdr[i]
            cell.text = txt
            shade(cell, "2F5496")
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.bold = True
                    r.font.size = Pt(9)
                    r.font.color.rgb = None
                    r.font.name = "微软雅黑"

    for ri, r in enumerate(rows):
        cells = table.add_row().cells
        for i in range(ncol):
            v = r[i] if i < len(r) else None
            cells[i].text = "" if v is None else str(v)
            if ri % 2:
                shade(cells[i], "F5F6F8")
            for p in cells[i].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8.5)

    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return {"file": out.name, "path": str(out), "rows": len(rows),
            "cols": ncol, "truncated": truncated,
            "size": f"{out.stat().st_size / 1024:.0f} KB",
            "sheet": meta["sheet"],
            "sheets": [s["name"] for s in meta["sheets"]],
            "warnings": meta["warnings"]}


# ══════════════════════════════════════════════════════════════
# 图片 → Excel（macOS Vision OCR）
# ══════════════════════════════════════════════════════════════

OCR_SWIFT = r'''
import Foundation
import Vision
import AppKit

let args = CommandLine.arguments
guard args.count > 1, let img = NSImage(contentsOfFile: args[1]),
      let tiff = img.tiffRepresentation,
      let bmp = NSBitmapImageRep(data: tiff),
      let cg = bmp.cgImage else {
    FileHandle.standardError.write("无法读取图片\n".data(using: .utf8)!)
    exit(2)
}
let req = VNRecognizeTextRequest()
req.recognitionLevel = .accurate
req.recognitionLanguages = ["zh-Hans", "en-US"]
req.usesLanguageCorrection = false
let handler = VNImageRequestHandler(cgImage: cg, options: [:])
do { try handler.perform([req]) } catch {
    FileHandle.standardError.write("OCR 失败: \(error)\n".data(using: .utf8)!)
    exit(3)
}
var out: [[String: Any]] = []
for ob in (req.results ?? []) {
    guard let top = ob.topCandidates(1).first else { continue }
    let b = ob.boundingBox            // 归一化坐标，原点在左下
    out.append(["text": top.string, "conf": top.confidence,
                "x": b.minX, "y": 1 - b.maxY, "w": b.width, "h": b.height])
}
print(String(data: try! JSONSerialization.data(withJSONObject: out),
             encoding: .utf8)!)
'''


def _ensure_ocr_binary() -> Path:
    """确保 OCR 二进制存在：优先用随包预编译的，否则就地编译并缓存"""
    bundled = core.res_path("vendor/ocr")
    if bundled.exists():
        return bundled
    import hashlib
    tag = hashlib.sha256(OCR_SWIFT.encode()).hexdigest()[:16]
    cache_dir = core.CACHE_DIR / f"ocr_{tag}"
    binary = cache_dir / "ocr"
    if binary.exists():
        return binary
    if not _has_swiftc():
        raise RuntimeError(
            "图片识别需要 macOS 的 Swift 编译器（Xcode Command Line Tools）。\n"
            "请在终端执行 xcode-select --install 后重试，"
            "或改用「文本转 Excel」手工粘贴表格内容。")
    cache_dir.mkdir(parents=True, exist_ok=True)
    src = cache_dir / "ocr.swift"
    src.write_text(OCR_SWIFT, encoding="utf-8")
    r = subprocess.run(["swiftc", "-O", str(src), "-o", str(binary)],
                       capture_output=True, text=True, timeout=300,
                       env=core.clean_subprocess_env())
    if r.returncode != 0 or not binary.exists():
        raise RuntimeError("编译 OCR 程序失败:\n" + (r.stderr or "")[:600])
    return binary


def _has_swiftc() -> bool:
    from shutil import which
    return which("swiftc") is not None


def ocr_blocks(image: Path) -> list[dict]:
    """调用 Vision 识别图片文字，返回带归一化坐标的文字块"""
    if not image.exists():
        raise FileNotFoundError(f"图片不存在: {image}")
    binary = _ensure_ocr_binary()
    r = subprocess.run([str(binary), str(image)], capture_output=True,
                       text=True, timeout=180, env=core.clean_subprocess_env())
    if r.returncode != 0:
        raise RuntimeError("图片识别失败: " + (r.stderr or "未知错误")[:400])
    line = next((l for l in r.stdout.splitlines() if l.startswith("[")), None)
    if line is None:
        raise RuntimeError("图片识别没有返回结果，可能图里没有可识别的文字")
    return json.loads(line)


def blocks_to_rows(blocks: list[dict], row_tol: float = 0.6,
                   col_gap: float = 0.025, min_fill: float = 0.15) -> list[list[str]]:
    """把带坐标的文字块还原成表格二维数组。

    行：按 y 坐标聚类，容差取字高的一定比例（同一行的文字 y 几乎相同）。
    列：把所有块的 x 起点聚类成列锚点，再把每个块归到最近的锚点。
    最后剔除填充率过低的列——截图里的水印、翻页控件会各自占一个稀疏列，
    不剔掉的话两列表格会变成七列。
    """
    if not blocks:
        return []
    items = sorted(blocks, key=lambda b: (b["y"], b["x"]))

    # ── 聚行 ──
    lines: list[list[dict]] = []
    for b in items:
        placed = False
        for ln in lines:
            ref = ln[0]
            tol = max(ref["h"], b["h"]) * row_tol
            if abs((b["y"] + b["h"] / 2) - (ref["y"] + ref["h"] / 2)) <= tol:
                ln.append(b)
                placed = True
                break
        if not placed:
            lines.append([b])
    for ln in lines:
        ln.sort(key=lambda b: b["x"])
    lines.sort(key=lambda ln: min(b["y"] for b in ln))

    # ── 聚列：所有 x 起点归并成锚点 ──
    xs = sorted(b["x"] for b in blocks)
    anchors: list[float] = []
    for x in xs:
        if not anchors or x - anchors[-1] > col_gap:
            anchors.append(x)
        else:
            anchors[-1] = (anchors[-1] + x) / 2
    if not anchors:
        return [[b["text"] for b in ln] for ln in lines]

    rows: list[list[str]] = []
    for ln in lines:
        cells = [""] * len(anchors)
        for b in ln:
            ci = min(range(len(anchors)), key=lambda i: abs(anchors[i] - b["x"]))
            cells[ci] = (cells[ci] + " " + b["text"]).strip() if cells[ci] else b["text"]
        rows.append(cells)

    if not rows:
        return []
    # 按填充率保留列：至少留 1 列，避免全被剔掉
    total = len(rows)
    fill_rate = [sum(1 for r in rows if r[i].strip()) / total
                 for i in range(len(anchors))]
    keep = [i for i, fr in enumerate(fill_rate) if fr >= min_fill]
    if not keep:
        keep = [max(range(len(fill_rate)), key=lambda i: fill_rate[i])]
    rows = [[r[i] for i in keep] for r in rows]
    # 去掉剔列后变成全空的行
    return [r for r in rows if any(c.strip() for c in r)]


def image_to_rows(image: Path, row_tol: float = 0.6, col_gap: float = 0.025,
                  min_fill: float = 0.15) -> dict:
    """图片 → 表格二维数组 + 识别质量信息"""
    blocks = ocr_blocks(image)
    rows = blocks_to_rows(blocks, row_tol, col_gap, min_fill)
    confs = [b.get("conf", 0) for b in blocks]
    low = [b["text"] for b in blocks if b.get("conf", 1) < 0.5]
    return {
        "rows": rows,
        "block_count": len(blocks),
        "row_count": len(rows),
        "col_count": len(rows[0]) if rows else 0,
        "avg_conf": round(sum(confs) / len(confs), 3) if confs else 0,
        "low_conf_texts": low[:20],
        "warning": ("部分文字识别置信度偏低，导入后请核对（尤其相似字符 f/t、l/1、0/O）"
                    if len(low) > 3 else ""),
    }
