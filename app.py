#!/usr/bin/env python3
"""Matrix AI 数据工具箱 · 独立小程序（内嵌 Web UI）

双击运行后：本地起服务 → 自动打开浏览器 → 常用工具一站式操作
  1. AI SQL 助手    自然语言生成只读 SQL、静态安全检查、差异解释
  2. SQL 排查工作台 新老脚本逻辑层 + 数据层对照，出排查日志
  3. SQL 查询       集群/库切换、结果表格、导出
  4. 表行数统计     一个库的全表行数，NULL 可兜底 COUNT(*)
  5. 库/表数据对比  表清单、字段、行数与主键明细比较
  6. 表对应关系     发现源库→目标库的表对应，复杂映射可手工补充
  7. 文件比对       两个 Excel/CSV 逐格比，主键可重复（组内最优配对）
  8. 文本/图片转 Excel
  9. Excel 转 INSERT / PDF / Word

只监听 127.0.0.1，不对外暴露。
"""

from __future__ import annotations

import socket
import hmac
import re
import secrets
import sys
import threading
import time
import webbrowser
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from flask import Flask, jsonify, render_template, request, send_from_directory

import matrix_core as core
import sql_tools as st
import convert as cv
import excel_diff as xd
import ai_tools as ait


def _boot_fail(msg: str):
    """启动阶段的致命错误：打包后没有终端，必须弹窗告知，否则双击毫无反应"""
    print(msg)
    if getattr(sys, "frozen", False) and sys.platform == "darwin":
        import subprocess
        body = msg.replace('"', "'")[:900]
        subprocess.run(["osascript", "-e",
                        f'display dialog "{body}" with title "Matrix AI 数据工具箱 启动失败" '
                        f'buttons {{"好"}} default button 1'],
                       capture_output=True, timeout=120)
    sys.exit(1)


app = Flask(__name__, template_folder=str(core.res_path("templates")))
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 上传 Excel 上限 200MB

try:
    CFG = core.load_config()
except Exception as e:  # noqa: BLE001
    _boot_fail(f"读取本机配置或示例配置失败：\n{e}")
CLIENT = core.MatrixClient(CFG)
UI = CFG.get("ui", {})
CLIENT_LOCK = threading.RLock()
SESSION_TOKEN = secrets.token_urlsafe(32)
UPLOAD_DIR = core.CACHE_DIR.parent / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
# 连库发现的表名对应存这里。打包后 .app 内部是只读的，必须落在家目录
BASE_DIR = core.CACHE_DIR.parent
AI_RUNTIME = ait.RuntimeAI()


def ok(**kw):
    return jsonify({"ok": True, **kw})


def fail(msg):
    return jsonify({"ok": False, "msg": str(msg)[:1500]})


def api(fn):
    """统一异常包装，避免每个接口重复 try/except"""
    def wrapper(*a, **kw):
        try:
            return fn(*a, **kw)
        except Exception as e:  # noqa: BLE001
            return fail(e)
    wrapper.__name__ = fn.__name__
    return wrapper


# ── 页面 ────────────────────────────────────────────────────────


@app.route("/")
def index():
    problems = CLIENT.check_env()
    CLIENT.data_dir.mkdir(parents=True, exist_ok=True)
    problems = [p for p in problems if "产物目录不存在" not in p]
    return render_template(
        "index.html",
        clusters=[v["label"] for v in CLIENT.clusters.values()],
        default_cluster=CLIENT.default_cluster_label(),
        common_dbs=UI.get("common_dbs", []),
        default_db=UI.get("default_db", ""),
        sql_max_rows=UI.get("sql_max_rows", 500),
        data_dir=str(CLIENT.data_dir),
        config_path=CFG.get("_config_path", ""),
        connection_state=CLIENT.public_state(),
        ai_state=AI_RUNTIME.public_state(),
        session_token=SESSION_TOKEN,
        problems=problems,
    )


# ── 数据库连接（前端运行时配置，不落盘）─────────────────────────


def _runtime_connection_config(p: dict) -> dict:
    """把前端参数转换为 MatrixClient 配置。

    返回值只用于当前进程；用户名和密码不会写回任何配置文件。
    """
    label = str(p.get("label") or "默认连接").strip()[:80]
    cluster_name = str(p.get("cluster") or label).strip()[:160]
    url_template = str(p.get("url_template") or "").strip()
    driver_class = str(p.get("driver_class") or "").strip()
    driver_jar = str(p.get("driver_jar") or "").strip()
    jdk_bin = str(p.get("jdk_bin") or "").strip()
    if not url_template.startswith("jdbc:"):
        raise core.ConfigError("JDBC URL 必须以 jdbc: 开头")
    if any(ch in url_template for ch in ("\r", "\n", "\x00")):
        raise core.ConfigError("JDBC URL 不能包含换行或空字符")
    if re.search(r"(?i)(?:[?;&]|^)(?:user(?:name)?|password|passwd|pwd)\s*=", url_template) \
            or re.search(r"(?i)jdbc:[^\s:]+://[^/@\s]+:[^/@\s]+@", url_template):
        raise core.ConfigError("请勿在 JDBC URL 中填写账号或密码，请使用独立的用户名和密码输入框")
    insecure = ("http://" in url_template.lower() or bool(re.search(
        r"(?i)(?:ssl|usessl|tls|encrypt)\s*=\s*(?:false|0|no)", url_template)))
    if insecure and not bool(p.get("allow_insecure")):
        raise core.ConfigError("检测到明文或显式关闭加密的连接；确认位于受控网络后再勾选允许不安全连接")
    if not driver_class or not driver_jar or not jdk_bin:
        raise core.ConfigError("JDBC 驱动类、驱动 JAR 路径和 JDK bin 均不能为空")
    timeout = max(5, min(int(p.get("timeout") or 180), 3600))
    workspace = str(p.get("workspace") or ".").strip()
    data_dir = str(p.get("data_dir") or "data").strip()
    clusters = {
        "runtime_1": {"label": label, "name": cluster_name, "default": True}
    }
    seen_labels = {label}
    raw_profiles = str(p.get("cluster_profiles") or "")
    for line in raw_profiles.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = re.split(r"\s*(?:=|\|)\s*", line, maxsplit=1)
        if len(parts) != 2 or not parts[0].strip():
            raise core.ConfigError("附加连接每行应写成：显示名称=URL中的cluster值")
        extra_label, extra_name = parts[0].strip()[:80], parts[1].strip()[:160]
        if extra_label in seen_labels:
            continue
        seen_labels.add(extra_label)
        clusters[f"runtime_{len(clusters) + 1}"] = {
            "label": extra_label, "name": extra_name, "default": False
        }
        if len(clusters) >= 20:
            break
    return {
        "server": dict(CFG.get("server") or {"port": 8765}),
        "paths": {
            "workspace": workspace,
            "jdk_bin": jdk_bin,
            "driver_jar": driver_jar,
            "data_dir": data_dir,
        },
        "matrix": {
            "driver_class": driver_class,
            "url_template": url_template,
            "username": str(p.get("username") or ""),
            "password": str(p.get("password") or ""),
            "clusters": clusters,
            "sql_timeout_seconds": timeout,
            "stat_timeout_seconds": max(timeout, 600),
        },
        "ui": {**dict(CFG.get("ui") or {}),
               "default_db": str(p.get("db") or "").strip()},
    }


def _require_session_token():
    supplied = request.headers.get("X-Toolbox-Token", "")
    if not hmac.compare_digest(supplied, SESSION_TOKEN):
        raise PermissionError("请求未通过本机页面校验")


@app.before_request
def protect_api_writes():
    """所有会执行查询、上传文件或写本地产物的 API 均要求本机页面令牌。"""
    if request.path.startswith("/api/") and request.method not in ("GET", "HEAD", "OPTIONS"):
        supplied = request.headers.get("X-Toolbox-Token", "")
        if not hmac.compare_digest(supplied, SESSION_TOKEN):
            return jsonify({"ok": False, "msg": "请求未通过本机页面校验"}), 403


@app.route("/api/connection")
@api
def api_connection_status():
    """只返回非敏感状态，不回传账号、密码或 JDBC URL。"""
    return ok(connection=CLIENT.public_state())


@app.route("/api/connection/test", methods=["POST"])
@api
def api_connection_test():
    """测试前端提供的 JDBC 连接；成功后可原子替换为当前活动连接。"""
    global CLIENT, CFG, UI
    _require_session_token()
    p = request.get_json(force=True)
    cfg = _runtime_connection_config(p)
    candidate = core.MatrixClient(cfg)
    candidate.data_dir.mkdir(parents=True, exist_ok=True)
    fatal = [x for x in candidate.check_env() if "产物目录不存在" not in x]
    if fatal:
        return fail("；".join(fatal))
    test_sql = str(p.get("test_sql") or "SELECT 1").strip()
    if not core.MatrixClient.is_readonly(test_sql):
        return fail("连接测试只允许 SELECT/SHOW/DESC/EXPLAIN/WITH 查询")
    db = str(p.get("db") or "").strip()
    result = candidate.query(test_sql, candidate.default_cluster_label(), db,
                             max_rows=10, timeout=min(candidate.timeout, 60))
    if p.get("apply", True):
        with CLIENT_LOCK:
            active = sum(1 for job in JOBS.values() if not job.get("done"))
            if active:
                return fail(f"当前有 {active} 个后台任务正在运行，请等待完成后再切换连接")
            CLIENT = candidate
            CFG = cfg
            UI = CFG.get("ui", {})
    return ok(message="连接测试成功，已作为当前活动连接",
              connection=candidate.public_state(),
              columns=result.get("headers", []),
              row_count=len(result.get("rows", [])))


@app.route("/api/connection/disconnect", methods=["POST"])
@api
def api_connection_disconnect():
    """从当前进程内存清除连接地址与凭据，不修改磁盘文件。"""
    global CLIENT, CFG, UI
    _require_session_token()
    blank = {
        "server": dict(CFG.get("server") or {"port": 8765}),
        "paths": {"workspace": ".", "jdk_bin": "", "driver_jar": "", "data_dir": "data"},
        "matrix": {"driver_class": "", "url_template": "", "username": "",
                   "password": "", "clusters": {}, "sql_timeout_seconds": 180,
                   "stat_timeout_seconds": 600},
        "ui": dict(CFG.get("ui") or {}),
    }
    with CLIENT_LOCK:
        active = sum(1 for job in JOBS.values() if not job.get("done"))
        if active:
            return fail(f"当前有 {active} 个后台任务正在运行，请等待完成后再断开连接")
        CLIENT = core.MatrixClient(blank)
        CFG = blank
        UI = CFG.get("ui", {})
    return ok(message="已断开连接并从内存清除凭据", connection=CLIENT.public_state())


# ── AI SQL 助手（模型配置同样只保存在当前进程内存）──────────────


@app.route("/api/ai/status")
@api
def api_ai_status():
    return ok(ai=AI_RUNTIME.public_state())


@app.route("/api/ai/config", methods=["POST"])
@api
def api_ai_config():
    global AI_RUNTIME
    p = request.get_json(force=True)
    # 候选配置先独立测试，失败时不覆盖当前可用配置，也不在全局对象里残留新 Key。
    candidate = ait.RuntimeAI()
    state = candidate.configure(p)
    probe = candidate.test() if p.get("test", True) else ""
    AI_RUNTIME = candidate
    return ok(message="AI 模型连接成功" if probe else "AI 模型配置已在内存中启用",
              probe=probe, ai=state)


@app.route("/api/ai/disconnect", methods=["POST"])
@api
def api_ai_disconnect():
    AI_RUNTIME.clear()
    return ok(message="已从内存清除 AI 服务地址和 API Key",
              ai=AI_RUNTIME.public_state())


@app.route("/api/ai/audit", methods=["POST"])
@api
def api_ai_audit():
    p = request.get_json(force=True)
    return ok(audit=ait.audit_readonly_sql(p.get("sql") or ""))


@app.route("/api/ai/schema", methods=["POST"])
@api
def api_ai_schema():
    """读取字段元数据供用户确认；不会自动发送给任何模型服务。"""
    p = request.get_json(force=True)
    cluster = str(p.get("cluster") or CLIENT.default_cluster_label()).strip()
    db = str(p.get("db") or "").strip()
    if not db:
        return fail("请填写要读取结构的库名")
    if len(db) > 180 or any(x in db for x in ("\x00", "\r", "\n")):
        return fail("库名格式不正确")
    table_limit = max(1, min(int(p.get("table_limit") or 30), 80))
    db_lit = db.replace("'", "''")
    sql = (
        "SELECT table_name, column_name, data_type, column_key, ordinal_position "
        "FROM information_schema.columns "
        f"WHERE table_schema = '{db_lit}' "
        "ORDER BY table_name, ordinal_position"
    )
    data = CLIENT.query(sql, cluster, db, max_rows=6000)
    grouped: dict[str, list[str]] = {}
    truncated_tables = False
    for row in data.get("rows", []):
        table = str(row[0])
        if table not in grouped and len(grouped) >= table_limit:
            truncated_tables = True
            continue
        dtype = str(row[2] or "")
        key = " PRIMARY KEY" if len(row) > 3 and str(row[3] or "").upper() == "PRI" else ""
        grouped.setdefault(table, []).append(f"{row[1]} {dtype}{key}".strip())
    lines = [f"database {db}"]
    lines.extend(f"table {name} (" + ", ".join(cols) + ")"
                 for name, cols in grouped.items())
    note = f"已读取 {len(grouped)} 张表的字段结构"
    if data.get("truncated") or truncated_tables:
        note += "（已按上限截断，请只保留本次查询相关表）"
    return ok(schema="\n".join(lines), tables=len(grouped), note=note,
              truncated=bool(data.get("truncated") or truncated_tables))


@app.route("/api/ai/generate", methods=["POST"])
@api
def api_ai_generate():
    p = request.get_json(force=True)
    result = AI_RUNTIME.generate_sql(p.get("task") or "", p.get("schema") or "",
                                     p.get("dialect") or "通用 SQL")
    return ok(**result)


@app.route("/api/ai/analyze", methods=["POST"])
@api
def api_ai_analyze():
    p = request.get_json(force=True)
    result = AI_RUNTIME.analyze_material(p.get("material") or "",
                                         p.get("question") or "")
    return ok(**result)


# ── 1. SQL 查询 ─────────────────────────────────────────────────


@app.route("/api/sql", methods=["POST"])
@api
def api_sql():
    p = request.get_json(force=True)
    sql = (p.get("sql") or "").strip()
    if not sql:
        return fail("SQL 不能为空")
    audit = ait.audit_readonly_sql(sql)
    if not audit["safe"]:
        reasons = "；".join(x["message"] for x in audit["issues"]
                           if x["level"] == "blocked")
        return fail("只读安全检查已拦截：" + reasons)
    limit = max(1, int(p.get("limit") or 500))
    data = CLIENT.query(sql, p["cluster"], (p.get("db") or "").strip(),
                        max_rows=limit)
    return ok(headers=data["headers"], rows=data["rows"], audit=audit,
              truncated=data.get("truncated", False))


# ── 2. 表行数统计 ───────────────────────────────────────────────


@app.route("/api/databases", methods=["POST"])
@api
def api_databases():
    """取出一个集群里的所有库（information_schema.schemata），供前端填充库名下拉"""
    p = request.get_json(force=True)
    cluster = p.get("cluster") or CLIENT.default_cluster_label()
    data = CLIENT.query("SELECT schema_name FROM information_schema.schemata "
                        "ORDER BY schema_name", cluster,
                        UI.get("default_db", ""), max_rows=500)
    names = [r[0] for r in data["rows"]]
    return ok(databases=names)


# 长任务用后台线程跑，前端轮询进度，避免 HTTP 超时
JOBS: dict[str, dict] = {}


def start_job(fn) -> str:
    jid = secrets.token_urlsafe(12)
    with CLIENT_LOCK:
        JOBS[jid] = {"done": False, "msg": "启动中…", "result": None, "error": None}

    def run():
        try:
            JOBS[jid]["result"] = fn(lambda m: JOBS[jid].update(msg=m))
        except Exception as e:  # noqa: BLE001
            JOBS[jid]["error"] = str(e)[:1200]
        finally:
            JOBS[jid]["done"] = True

    threading.Thread(target=run, daemon=True).start()
    return jid


@app.route("/api/job/<jid>")
def api_job(jid):
    job = JOBS.get(jid)
    if not job:
        return fail("任务不存在或已过期")
    if not job["done"]:
        return jsonify({"ok": True, "done": False, "msg": job["msg"]})
    if job["error"]:
        return jsonify({"ok": False, "done": True, "msg": job["error"]})
    return jsonify({"ok": True, "done": True, **job["result"]})


@app.route("/api/rowcount", methods=["POST"])
@api
def api_rowcount():
    p = request.get_json(force=True)
    cluster, db = p["cluster"], (p.get("db") or "").strip()
    if not db:
        return fail("请填写库名")
    fallback = bool(p.get("fallback"))
    limit = int(UI.get("stat_max_rows", 5000))

    def work(progress):
        progress(f"读取 {db} 的表清单…")
        items = core.fetch_table_rows(CLIENT, cluster, db, max_rows=limit)
        if fallback:
            # 视图的 table_rows 天然为空，跳过，免得白扫一遍
            nulls = [t for t in items if t["rows"] is None
                     and "VIEW" not in t["type"].upper()]
            for i, t in enumerate(nulls, 1):
                progress(f"兜底 COUNT(*) {i}/{len(nulls)}: {t['name']}")
                try:
                    t["rows"] = core.count_one(CLIENT, cluster, db, t["name"])
                except Exception:
                    t["rows"] = None
        total = sum(t["rows"] for t in items if t["rows"])
        return {"items": items, "db": db,
                "summary": {
                    "objects": len(items),
                    "views": sum(1 for t in items if "VIEW" in t["type"].upper()),
                    "empty": sum(1 for t in items if t["rows"] == 0),
                    "unknown": sum(1 for t in items if t["rows"] is None),
                    "total_rows": total}}

    return ok(job=start_job(work))


# ── 3. 库数据对比：库表对比 / 自定义对比 / 单表数据对比 ──────────


@app.route("/api/tables", methods=["POST"])
@api
def api_tables():
    """取一个库的表清单（带行数），供表名映射预览与单表对比的下拉框使用"""
    p = request.get_json(force=True)
    db = (p.get("db") or "").strip()
    if not db:
        return fail("请填写库名")
    items = core.fetch_table_rows(CLIENT, p["cluster"], db,
                                  max_rows=int(UI.get("stat_max_rows", 5000)))
    return ok(items=items, db=db, count=len(items))


@app.route("/api/mapping/modes")
def api_mapping_modes():
    return ok(modes=core.MAP_MODES)


@app.route("/api/mapping/preview", methods=["POST"])
@api
def api_mapping_preview():
    """映射规则试跑：先看几个例子对不对，再决定要不要连库跑全量"""
    p = request.get_json(force=True)
    rule = p.get("rule") or {}
    da, db_ = (p.get("dbA") or "").strip(), (p.get("dbB") or "").strip()
    names = p.get("names") or []
    if not names:
        if not da:
            return fail("请先填 A 侧库名，或直接传入表名清单")
        items = core.fetch_table_rows(CLIENT, p["clusterA"], da,
                                      max_rows=int(UI.get("stat_max_rows", 5000)))
        names = [t["name"] for t in items]
    return ok(**core.mapping_preview(names, rule, da, db_,
                                     limit=int(p.get("limit") or 40)))


@app.route("/api/mapping/saved", methods=["POST"])
@api
def api_mapping_saved():
    """查这组集群+库有没有存过表名对应"""
    p = request.get_json(force=True)
    store = core.load_pairs(BASE_DIR)
    key = core.pairs_key(p.get("clusterA") or "", (p.get("dbA") or "").strip(),
                         p.get("clusterB") or "", (p.get("dbB") or "").strip())
    hit = store.get(key) or {}
    return ok(key=key, found=bool(hit), count=hit.get("count", 0),
              saved_at=hit.get("saved_at", ""), rule_desc=hit.get("rule_desc", ""),
              pairs=hit.get("pairs", {}), groups=sorted(store.keys()))


@app.route("/api/mapping/discover", methods=["POST"])
@api
def api_mapping_discover():
    """连库把两边表清单拉下来算真实对应，结果存本地复用。

    规则只说明「B 侧应该叫什么」，两边表清单才说明「实际存不存在」。
    """
    p = request.get_json(force=True)
    da, db_ = (p.get("dbA") or "").strip(), (p.get("dbB") or "").strip()
    if not (da and db_):
        return fail("请把两侧库名都填上")
    ca = p.get("clusterA") or CLIENT.default_cluster_label()
    cb = p.get("clusterB") or ca

    def work(progress):
        r = core.discover_pairs(CLIENT, ca, da, cb, db_, p.get("rule") or {},
                                progress)
        if p.get("save", True) and r["pairs"]:
            r["saved"] = core.save_pairs(BASE_DIR, r["key"], r["pairs"],
                                         {"rule_desc": r["rule_desc"]})
        # 表名对太多时前端不需要全量，存盘的才是完整的
        r["sample"] = [{"a": a, "b": b}
                       for a, b in list(r["pairs"].items())[:60]]
        r["pairs_total"] = len(r["pairs"])
        r.pop("pairs", None)
        return r

    return ok(job=start_job(work))


@app.route("/api/mapping/discover-all", methods=["POST"])
@api
def api_mapping_discover_all():
    """按可配置前缀发现源库 → 目标库的表对应关系并存下来。"""
    p = request.get_json(force=True)
    source_cluster = p.get("sourceCluster") or CLIENT.default_cluster_label()
    target_cluster = p.get("targetCluster") or CLIENT.default_cluster_label()
    target_db = (p.get("targetDb") or "").strip()
    prefix = (p.get("prefix") or "target_").strip()
    if not (source_cluster and target_cluster and target_db):
        return fail("请填写源连接、目标连接和目标库名")

    def work(progress):
        r = core.discover_target_map(
            CLIENT, source_cluster, target_cluster, target_db, prefix, progress)
        if p.get("save", True) and r["pairs"]:
            r["saved"] = core.save_discovered_map(BASE_DIR, r)
        # 明细太多，页面只要按库汇总 + 几个样例
        r["summary"] = [{"db": k, "count": len(v),
                         "sample": [f"{i['source_table']} → {i['target_table']}"
                                    for i in v[:2]]}
                        for k, v in sorted(r["by_db"].items(),
                                           key=lambda x: -len(x[1]))]
        r.pop("by_db", None)
        r.pop("pairs", None)
        return r

    return ok(job=start_job(work))


@app.route("/api/mapping/groups")
@api
def api_mapping_groups():
    """已存的对应关系分组概览"""
    return ok(groups=core.pairs_groups(BASE_DIR),
              file=str(core.pairs_path(BASE_DIR)))


@app.route("/api/mapping/list", methods=["POST"])
@api
def api_mapping_list():
    """看一组对应关系的明细，可按库名/关键词筛"""
    p = request.get_json(force=True)
    return ok(**core.pairs_of_group(BASE_DIR, p.get("key") or "",
                                    (p.get("db") or "").strip(),
                                    (p.get("kw") or "").strip(),
                                    int(p.get("limit") or 500)))


@app.route("/api/mapping/edit", methods=["POST"])
@api
def api_mapping_edit():
    """手工维护对应关系：粘一批新增、删掉几条、或整组替换"""
    p = request.get_json(force=True)
    key = (p.get("key") or "").strip()
    if not key:
        return fail("请先选一组对应关系，或填一个新的分组名")
    add = core.parse_pairs(p.get("text") or "")
    drop = [x for x in (p.get("drop") or []) if x]
    if not (add or drop or p.get("replace")):
        return fail("没解析出任何表名对。每行写一对，如 "
                    "`source_table => target_source_table`"
                    "（支持 => -> → 逗号 制表符 竖线 分隔）")
    r = core.edit_pairs(BASE_DIR, key, add, drop, bool(p.get("replace")))
    return ok(**r, added=len(add), dropped=len(drop))


@app.route("/api/dbcompare", methods=["POST"])
@api
def api_dbcompare():
    p = request.get_json(force=True)
    ca, da = p["clusterA"], (p.get("dbA") or "").strip()
    cb, db_ = p["clusterB"], (p.get("dbB") or "").strip()
    if not (da and db_):
        return fail("请填写两侧库名")
    rule = p.get("rule") or {"mode": "same"}
    pair_src = ""
    if (rule.get("mode") or "") == "saved" and not rule.get("saved_pairs"):
        # 先找这一对库单独存过的，没有就从整套源库→目标库对应里切一段
        rule["saved_pairs"], pair_src = core.saved_pairs_for(
            BASE_DIR, ca, da, cb, db_)
    same_side = (ca, da) == (cb, db_)
    if same_side and (rule.get("mode") or "same") == "same":
        return fail("两侧是同一个库且按同名对比，结果必然全一致。\n"
                    "要在同库内比改了名的表（如 xxx 与 xxx_old），请选一种表名映射方式")
    # 规则先在本地校验，写错了不必等连库跑完才报
    mapper, desc = core.build_mapper(rule, da, db_)
    b_filter, scope_desc = core.build_b_filter(rule, da, db_)
    limit = int(UI.get("stat_max_rows", 5000))

    def work(progress):
        progress(f"读取 A 侧 {da} …")
        a = core.fetch_table_rows(CLIENT, ca, da, max_rows=limit)
        progress(f"A 侧 {len(a)} 个对象，读取 B 侧 {db_} …")
        b = core.fetch_table_rows(CLIENT, cb, db_, max_rows=limit)
        progress(f"按「{desc}」换算表名后对比…")
        r = core.compare_two_dbs(a, b, mapper, b_filter)
        return {"result": r, "dbA": da, "dbB": db_,
                "clusterA": ca, "clusterB": cb,
                "mapping": {"mode": rule.get("mode", "same"), "desc": desc,
                            "scope": scope_desc, "source": pair_src,
                            "out_of_scope": r.get("out_of_scope", 0)},
                "summary": {k: len(v) for k, v in r.items()
                            if k not in ("pairs", "out_of_scope")}}

    return ok(job=start_job(work))


@app.route("/api/tablecompare", methods=["POST"])
@api
def api_tablecompare():
    """单表数据对比：结构 → 行数 → 按主键对齐比明细"""
    p = request.get_json(force=True)
    ta, tb = (p.get("tableA") or "").strip(), (p.get("tableB") or "").strip()
    da, db_ = (p.get("dbA") or "").strip(), (p.get("dbB") or "").strip()
    if not (da and ta):
        return fail("请填写 A 侧库名与表名")
    if not db_:
        db_ = da
    if not tb:
        # 没填 B 侧表名就按映射规则推一个，避免手工输入长表名
        mapper, _ = core.build_mapper(p.get("rule") or {"mode": "same"}, da, db_)
        tb = mapper(ta)
    side_a = {"cluster": p["clusterA"], "db": da, "table": ta}
    side_b = {"cluster": p.get("clusterB") or p["clusterA"], "db": db_, "table": tb}
    if side_a == side_b:
        return fail("两侧是同一张表，无需对比")
    keys = [k.strip() for k in (p.get("keys") or "").split(",") if k.strip()]
    limit = max(1, min(int(p.get("limit") or 500), int(UI.get("stat_max_rows", 5000))))

    def work(progress):
        return {"result": core.compare_one_table(
            CLIENT, side_a, side_b, keys=keys, limit=limit,
            where_a=(p.get("whereA") or "").strip(),
            where_b=(p.get("whereB") or "").strip(),
            progress=progress)}

    return ok(job=start_job(work))


# ── 4. SQL 排查工作台 ───────────────────────────────────────────


@app.route("/api/sql/logic", methods=["POST"])
@api
def api_sql_logic():
    """第一层：只比逻辑，不连库"""
    p = request.get_json(force=True)
    a, b = (p.get("sql_a") or "").strip(), (p.get("sql_b") or "").strip()
    if not (a and b):
        return fail("请把两个脚本 SQL 都填上")
    na = p.get("name_a") or "脚本A"
    nb = p.get("name_b") or "脚本B"
    r = st.compare_logic(a, b, na, nb)
    pa, pb = r["parsed"]["a"], r["parsed"]["b"]
    return ok(
        names=r["names"], risks=r["risks"], identical=r["identical"],
        tables=r["tables"], conditions=r["conditions"], fields=r["fields"],
        joins=r["joins"],
        summary={
            "a": {"tables": len(pa["tables"]), "joins": len(pa["joins"]),
                  "conditions": len(pa["conditions"]),
                  "fields": len(pa["select_fields"]),
                  "subqueries": pa["subquery_count"],
                  "union_parts": pa["union_parts"], "chars": len(pa["normalized"])},
            "b": {"tables": len(pb["tables"]), "joins": len(pb["joins"]),
                  "conditions": len(pb["conditions"]),
                  "fields": len(pb["select_fields"]),
                  "subqueries": pb["subquery_count"],
                  "union_parts": pb["union_parts"], "chars": len(pb["normalized"])},
        },
        diff=st.text_diff(a, b)[:600])


@app.route("/api/sql/ai-logic", methods=["POST"])
@api
def api_sql_ai_logic():
    """在静态解析证据之上做 AI 语义分析；不连库、不自动执行模型建议。"""
    p = request.get_json(force=True)
    a, b = (p.get("sql_a") or "").strip(), (p.get("sql_b") or "").strip()
    if not (a and b):
        return fail("请把两个脚本 SQL 都填上")
    na, nb = p.get("name_a") or "脚本A", p.get("name_b") or "脚本B"
    static = st.compare_logic(a, b, na, nb)
    pa, pb = static["parsed"]["a"], static["parsed"]["b"]
    summary = {
        "a": {"tables": len(pa["tables"]), "joins": len(pa["joins"]),
              "conditions": len(pa["conditions"]), "fields": len(pa["select_fields"]),
              "subqueries": pa["subquery_count"], "union_parts": pa["union_parts"]},
        "b": {"tables": len(pb["tables"]), "joins": len(pb["joins"]),
              "conditions": len(pb["conditions"]), "fields": len(pb["select_fields"]),
              "subqueries": pb["subquery_count"], "union_parts": pb["union_parts"]},
    }
    evidence = {
        "static_risks": static["risks"],
        "summary": summary,
        "tables": static["tables"],
        "joins": static["joins"],
        "conditions": static["conditions"],
        "fields": static["fields"],
    }
    result = AI_RUNTIME.compare_sql(
        a, b, evidence,
        context=p.get("context") or "",
        name_a=na, name_b=nb,
        dialect=p.get("dialect") or "StarRocks / MySQL",
    )
    return ok(**result, static={"risks": static["risks"], "summary": summary})


@app.route("/api/sql/split", methods=["POST"])
@api
def api_sql_split():
    """长 SQL 逻辑拆分"""
    p = request.get_json(force=True)
    sql = (p.get("sql") or "").strip()
    if not sql:
        return fail("请填入 SQL")
    return ok(**st.split_sql(sql))


@app.route("/api/sql/prepare", methods=["POST"])
@api
def api_sql_prepare():
    """取数前的准备：按「先精确日期 → 再找主键」的顺序给出候选，
    并附依赖表分层与辅助排查 SQL"""
    p = request.get_json(force=True)
    sql = (p.get("sql") or "").strip()
    if not sql:
        return fail("请填入 SQL")
    times = st.find_time_columns(sql)
    keys = st.suggest_primary_keys(sql)
    suggest = st.suggest_time_range(sql)
    # 给了集群就连库读表结构：字段名像不像时间是猜的，data_type 和 PRI 是准的
    fields = {}
    if p.get("cluster"):
        real = [t["name"] for t in st.extract_tables(st.strip_comments(sql))
                if "(" not in t["name"]]
        if real:
            fields = core.table_field_hints(
                CLIENT, p["cluster"],
                (p.get("db") or UI.get("default_db", "")).strip(),
                sorted(set(real))[:8])
    detail_sql = ""
    if keys:
        detail_sql = st.build_detail_sql(
            sql, [keys[0]["column"]],
            times[0]["column"] if times else "",
            suggest["start"], suggest["end"], 500)
    return ok(time_columns=times,
              primary_keys=keys,
              default_key=keys[0]["column"] if keys else "",
              table_fields=fields,
              detail_sql=detail_sql,
              suggest=suggest,
              lineage=st.table_lineage(sql),
              probes=st.build_probe_sqls(sql, (p.get("pk") or "").strip()
                                        or (keys[0]["column"] if keys else "")))


@app.route("/api/sql/plan-scan", methods=["POST"])
@api
def api_sql_plan_scan():
    """规划全量分片：先查明细总量，再按时间分组统计，装箱成每片 ≤500 行"""
    p = request.get_json(force=True)
    sql = (p.get("sql") or "").strip()
    time_col = (p.get("time_col") or "").strip()
    if not sql:
        return fail("请填入 SQL")
    if not time_col:
        return fail("全量分片必须指定业务时间字段，否则无法切片")
    start, end = (p.get("start") or "").strip(), (p.get("end") or "").strip()
    if not (start and end):
        return fail("请给出起止日期")
    batch = max(1, min(int(p.get("batch") or 500), 500))
    cluster = p.get("cluster") or CLIENT.default_cluster_label()
    db = (p.get("db") or UI.get("default_db", "")).strip()

    def work(progress):
        def run(text):
            d = CLIENT.query(text, cluster, db, max_rows=batch)
            return d["headers"], d["rows"]
        plan = st.plan_full_scan(run, sql, time_col, start, end, batch, progress)
        plan["slices"] = plan["slices"][:400]      # 回传上限，避免响应过大
        return {"plan": plan}

    return ok(job=start_job(work))


@app.route("/api/sql/compare-data", methods=["POST"])
@api
def api_sql_compare_data():
    """第二层：两侧取数后按主键对齐比明细。

    两种取数模式：
    - 抽样（默认）：各取 ≤500 行，快，适合先看有没有明显差异
    - 全量分片：按业务时间把范围切成每片 ≤500 行，逐片取完再拼接，
      这样能突破 Matrix 单次 500 行的限制，比完一个时间段的全部数据
    """
    p = request.get_json(force=True)
    a, b = (p.get("sql_a") or "").strip(), (p.get("sql_b") or "").strip()
    if not (a and b):
        return fail("请把两个脚本 SQL 都填上")
    keys = [k.strip() for k in (p.get("keys") or "").split(",") if k.strip()]
    if not keys:
        return fail("请填明细主键字段（多个用英文逗号分隔），否则无法对齐比较")
    time_col = (p.get("time_col") or "").strip()
    start, end = (p.get("start") or "").strip(), (p.get("end") or "").strip()
    batch = max(1, min(int(p.get("limit") or 500), 500))
    full_scan = bool(p.get("full_scan"))
    cluster_a = p.get("cluster_a") or CLIENT.default_cluster_label()
    cluster_b = p.get("cluster_b") or cluster_a
    db_a = (p.get("db_a") or UI.get("default_db", "")).strip()
    db_b = (p.get("db_b") or db_a).strip()

    if time_col and not (start and end):
        return fail("选了业务时间字段就必须同时给出起止日期")
    if full_scan and not time_col:
        return fail("全量分片必须指定业务时间字段")

    def runner(cluster, db):
        def run(text):
            d = CLIENT.query(text, cluster, db, max_rows=batch)
            return d["headers"], d["rows"]
        return run

    def work(progress):
        info = {}
        if full_scan:
            # 两侧各自规划分片（数据分布可能不同，不能共用一套方案）
            progress("A 侧：规划时间分片…")
            ra = runner(cluster_a, db_a)
            plan_a = st.plan_full_scan(ra, a, time_col, start, end, batch,
                                       lambda m: progress(f"A 侧：{m}"))
            progress(f"A 侧：{plan_a['slice_count']} 片，逐片取数…")
            fa = st.fetch_full_scan(ra, a, time_col, plan_a, keys,
                                    progress=lambda m: progress(f"A 侧 {m}"))
            progress("B 侧：规划时间分片…")
            rb = runner(cluster_b, db_b)
            plan_b = st.plan_full_scan(rb, b, time_col, start, end, batch,
                                       lambda m: progress(f"B 侧：{m}"))
            progress(f"B 侧：{plan_b['slice_count']} 片，逐片取数…")
            fb = st.fetch_full_scan(rb, b, time_col, plan_b, keys,
                                    progress=lambda m: progress(f"B 侧 {m}"))
            ha, rows_a = fa["headers"], fa["rows"]
            hb, rows_b = fb["headers"], fb["rows"]
            info = {"mode": "full_scan",
                    "a": {"total": plan_a["total_rows"], "agg": plan_a["agg_rows"],
                          "slices": plan_a["slice_count"], "fetched": fa["fetched"],
                          "grains": sorted({s["grain"] for s in plan_a["slices"]}),
                          "consistent": plan_a["consistent"],
                          "slice_detail": fa["slice_detail"][:200],
                          "warning": plan_a["warning"], "note": fa["note"]},
                    "b": {"total": plan_b["total_rows"], "agg": plan_b["agg_rows"],
                          "slices": plan_b["slice_count"], "fetched": fb["fetched"],
                          "grains": sorted({s["grain"] for s in plan_b["slices"]}),
                          "consistent": plan_b["consistent"],
                          "slice_detail": fb["slice_detail"][:200],
                          "warning": plan_b["warning"], "note": fb["note"]},
                    "detail_mode": fa["detail_mode"] or fb["detail_mode"],
                    "detail_note": fa["detail_note"] or fb["detail_note"],
                    "rewrite_note": plan_a["rewrite_note"]}
            sql_a_exec, sql_b_exec = fa["executed_sample"], fb["executed_sample"]
        else:
            # 抽样模式：在原层次改写时间条件，不套壳
            sql_a_exec = (st.rewrite_time_filter(a, time_col, start, end)["sql"]
                          if time_col else a)
            sql_b_exec = (st.rewrite_time_filter(b, time_col, start, end)["sql"]
                          if time_col else b)
            progress(f"查 A 侧（{db_a}）…")
            da = CLIENT.query(sql_a_exec, cluster_a, db_a, max_rows=batch)
            progress(f"A 侧 {len(da['rows'])} 行，查 B 侧（{db_b}）…")
            db_res = CLIENT.query(sql_b_exec, cluster_b, db_b, max_rows=batch)
            ha, rows_a = da["headers"], da["rows"]
            hb, rows_b = db_res["headers"], db_res["rows"]
            info = {"mode": "sample",
                    "hint": f"抽样模式各取 ≤{batch} 行。要比完整个时间段的全部数据，"
                            f"请勾选「全量分片」"}

        progress("按主键对齐比对明细…")
        r = st.compare_details(ha, rows_a, hb, rows_b, keys, sql_a=a, sql_b=b)
        r["scan"] = info
        r["executed"] = {"sql_a": sql_a_exec, "sql_b": sql_b_exec,
                         "time_col": time_col, "start": start, "end": end,
                         "limit": batch, "full_scan": full_scan}
        r["headers"] = {"a": ha, "b": hb}
        return r

    return ok(job=start_job(work))


@app.route("/api/sql/accumulate-diffs", methods=["POST"])
@api
def api_sql_accumulate_diffs():
    """把一轮 SQL 对比的差异累入差异跟踪 Excel，跨时间片多轮排查用。

    按「主键 + 字段」去重合并：
      · 新差异 → 追加，标「未修复」
      · 历史差异本轮还在 → 更新值与最后出现时间
      · 历史未修复条目本轮没再出现（且 pk 在本轮复查范围内） → 标「已修复」
      · 曾经修过又冒出来 → 标「又出现」
    """
    p = request.get_json(force=True)
    name = core.safe_name((p.get("name") or "").strip() or "差异跟踪")
    new_diffs = p.get("diffs") or []
    time_slice = (p.get("time_slice") or "").strip()
    if not new_diffs:
        return fail("本轮没有可累入的差异")
    if not time_slice:
        return fail("请填写时间片标签，否则多轮累入时无法区分来源")
    archive_path = CLIENT.data_dir / f"{name}.xlsx"
    r = core.accumulate_diffs(archive_path, new_diffs, time_slice,
                              p.get("pks_in_scope") or [])
    return ok(
        file=archive_path.name,
        path=str(archive_path),
        size=f"{archive_path.stat().st_size / 1024:.0f} KB",
        added=r["added"], fixed=r["fixed"],
        kept_unfixed=r["kept_unfixed"], updated=r["updated"],
        total=r["total"], was_empty=r["was_empty"],
        time_slice=time_slice)


@app.route("/api/sql/export-log", methods=["POST"])
@api
def api_sql_export_log():
    """把排查日志与差异明细导出成 Excel，方便贴到交接文档"""
    p = request.get_json(force=True)
    logs = p.get("logs") or []
    diffs = p.get("diffs") or []
    stats = p.get("stats") or {}
    verdict = p.get("verdict") or {}
    executed = p.get("executed") or {}
    name = core.safe_name(p.get("name") or "SQL排查日志")
    out = CLIENT.data_dir / f"{name}.xlsx"
    core.write_xlsx(out, {
        "结论": (["项目", "内容"],
                 [["结论", verdict.get("text", "")],
                  ["下一步", verdict.get("next", "")],
                  ["主键", ", ".join(stats.get("keys") or [])],
                  ["A 侧行数", stats.get("rows_a")],
                  ["B 侧行数", stats.get("rows_b")],
                  ["匹配行对", stats.get("matched")],
                  ["仅 A 有", stats.get("only_a")],
                  ["仅 B 有", stats.get("only_b")],
                  ["字段差异数", stats.get("diff_cells")],
                  ["一边为空的差异", stats.get("null_flip")],
                  ["业务时间字段", executed.get("time_col", "")],
                  ["时间范围", f"{executed.get('start','')} ~ {executed.get('end','')}"]]),
        "排查日志": (["级别", "明细主键", "字段", "出错原因", "说明"],
                     [[l.get("level"), l.get("key"), l.get("column"),
                       l.get("reason"), l.get("detail")] for l in logs]),
        "差异明细": (["主键", "字段", "A 值", "B 值"],
                     [[d.get("key"), d.get("col"), d.get("a"), d.get("b")]
                      for d in diffs]),
        "实际执行SQL": (["侧", "SQL"],
                        [["A", executed.get("sql_a", "")],
                         ["B", executed.get("sql_b", "")]]),
    })
    return ok(file=out.name, path=str(out),
              size=f"{out.stat().st_size / 1024:.0f} KB")


# ── 5. Excel → INSERT / PDF / Word ─────────────────────────────


def _save_upload(field: str, suffix: str) -> Path:
    f = request.files.get(field)
    if not f:
        raise ValueError(f"请选择文件（{field}）")
    path = UPLOAD_DIR / f"{int(time.time() * 1000)}_{field}{suffix}"
    f.save(path)
    return path


@app.route("/api/excel/sheets", methods=["POST"])
@api
def api_excel_sheets():
    """上传后先看有哪些工作表，避免默认读错页导致转出空文件"""
    xlsx = _save_upload("file", ".xlsx")
    sheets = core.list_sheets(xlsx)
    headers, rows, meta = core.read_sheet_meta(xlsx)
    return ok(sheets=sheets, picked=meta["sheet"], warnings=meta["warnings"],
              headers=headers[:40], preview_rows=len(rows),
              token=xlsx.name)


@app.route("/api/excel2insert", methods=["POST"])
@api
def api_excel2insert():
    ddl = (request.form.get("ddl") or "").strip()
    if not ddl:
        return fail("请贴入目标表的建表语句（CREATE TABLE …）")
    batch = max(1, int(request.form.get("batch") or 500))
    sheet = (request.form.get("sheet") or "").strip() or None
    xlsx = _save_upload("file", ".xlsx")
    r = cv.excel_to_insert(xlsx, ddl, batch=batch, sheet=sheet)
    out = CLIENT.data_dir / f"{core.safe_name(r['table'])}_insert.sql"
    out.write_text(r["sql"], encoding="utf-8")
    r["file"] = out.name
    r["size"] = f"{out.stat().st_size / 1024:.0f} KB"
    r["preview"] = r["sql"][:4000]
    r["sql_truncated"] = len(r["sql"]) > 4000
    del r["sql"]
    return ok(**r)


@app.route("/api/excel2doc", methods=["POST"])
@api
def api_excel2doc():
    """Excel → PDF 或 Word"""
    fmt = (request.form.get("format") or "pdf").lower()
    landscape = (request.form.get("landscape") or "1") == "1"
    title = (request.form.get("title") or "").strip()
    sheet = (request.form.get("sheet") or "").strip() or None
    max_rows = max(1, int(request.form.get("max_rows") or 800))
    xlsx = _save_upload("file", ".xlsx")
    stem = core.safe_name(Path(request.files["file"].filename).stem or "导出")
    if fmt == "word":
        out = CLIENT.data_dir / f"{stem}.docx"
        r = cv.excel_to_word(xlsx, out, landscape=landscape, title=title,
                             sheet=sheet, max_rows=max_rows)
    else:
        out = CLIENT.data_dir / f"{stem}.pdf"
        r = cv.excel_to_pdf(xlsx, out, landscape=landscape, sheet=sheet,
                            max_rows=max_rows)
    return ok(**r)


@app.route("/api/image2rows", methods=["POST"])
@api
def api_image2rows():
    """图片 → 表格（macOS Vision OCR 按坐标还原行列）"""
    img = _save_upload("file", Path(request.files["file"].filename).suffix or ".png")
    r = cv.image_to_rows(img, min_fill=float(request.form.get("min_fill") or 0.15))
    r["rows"] = r["rows"][:500]
    return ok(**r)


# ── 5. 文本转 Excel ─────────────────────────────────────────────


@app.route("/api/text2rows", methods=["POST"])
@api
def api_text2rows():
    p = request.get_json(force=True)
    rows = core.text_to_rows(p.get("text") or "", p.get("delim") or "auto")
    return ok(rows=rows[:500], total=len(rows), cols=len(rows[0]))


# ── 导出 ────────────────────────────────────────────────────────


@app.route("/api/xdiff/upload", methods=["POST"])
@api
def api_xdiff_upload():
    """两个文件先传上来，回两边的列清单与共有列，供选主键"""
    pa = _save_upload("file_a", Path(request.files["file_a"].filename or "a").suffix
                      or ".xlsx")
    pb = _save_upload("file_b", Path(request.files["file_b"].filename or "b").suffix
                      or ".xlsx")
    ha, ra = xd.read_table(pa)
    hb, rb = xd.read_table(pb)
    common = xd.common_columns(ha, hb)
    if not common:
        return fail("两个文件没有一个同名列，没法逐格比。"
                    "先确认表头行是否一致（多一行标题会让表头读错）")
    return ok(token_a=pa.name, token_b=pb.name,
              name_a=request.files["file_a"].filename,
              name_b=request.files["file_b"].filename,
              headers_a=ha, headers_b=hb, common=common,
              only_a=[h for h in ha if h not in set(hb)],
              only_b=[h for h in hb if h not in set(ha)],
              rows_a=len(ra), rows_b=len(rb),
              key_guess=xd.guess_keys(ha, ra, hb, rb, common),
              options=xd.DEFAULT_OPTIONS)


@app.route("/api/xdiff/compare", methods=["POST"])
@api
def api_xdiff_compare():
    """按主键比两个文件。主键可以重复，重复组内按内容做最优配对"""
    p = request.get_json(force=True)
    pa, pb = UPLOAD_DIR / (p.get("token_a") or ""), UPLOAD_DIR / (p.get("token_b") or "")
    if not (pa.exists() and pb.exists()):
        return fail("上传的文件已失效，请重新上传")
    keys = [k for k in (p.get("keys") or []) if str(k).strip()]
    if not keys:
        return fail("请至少选一个主键列，否则两边的行没法对齐")
    opts = {k: bool(p.get(k, v)) for k, v in xd.DEFAULT_OPTIONS.items()}

    def work(progress):
        progress("读取两个文件…")
        ha, ra = xd.read_table(pa)
        hb, rb = xd.read_table(pb)
        progress(f"A {len(ra)} 行 / B {len(rb)} 行，按主键分组比对…")
        r = xd.compare_tables(ha, ra, hb, rb, keys, opts)
        s = r["stats"]
        if s["dup_keys"]:
            progress(f"{s['dup_keys']} 个主键重复，组内做最优配对…")
        r["names"] = {"a": p.get("name_a") or pa.name, "b": p.get("name_b") or pb.name}
        r["tokens"] = {"a": pa.name, "b": pb.name}
        # 页面只展示前若干条，完整结果导出到 Excel
        r["diffs"] = r["diffs"][:300]
        r["only_a"] = r["only_a"][:200]
        r["only_b"] = r["only_b"][:200]
        return r

    return ok(job=start_job(work))


@app.route("/api/xdiff/export", methods=["POST"])
@api
def api_xdiff_export():
    """重跑一遍并把完整结果导出成 Excel（页面上只有截断后的预览）"""
    p = request.get_json(force=True)
    pa, pb = UPLOAD_DIR / (p.get("token_a") or ""), UPLOAD_DIR / (p.get("token_b") or "")
    if not (pa.exists() and pb.exists()):
        return fail("上传的文件已失效，请重新上传后再导出")
    keys = [k for k in (p.get("keys") or []) if str(k).strip()]
    if not keys:
        return fail("请至少选一个主键列")
    opts = {k: bool(p.get(k, v)) for k, v in xd.DEFAULT_OPTIONS.items()}

    def work(progress):
        progress("重跑比对以拿到完整结果…")
        ha, ra = xd.read_table(pa)
        hb, rb = xd.read_table(pb)
        r = xd.compare_tables(ha, ra, hb, rb, keys, opts)
        na = core.safe_name(Path(p.get("name_a") or "A").stem)[:12] or "A"
        nb = core.safe_name(Path(p.get("name_b") or "B").stem)[:12] or "B"
        progress("写报告…")
        name = core.safe_name(p.get("name") or f"文件比对报告_{na}_vs_{nb}")
        out = CLIENT.data_dir / f"{name}.xlsx"
        core.write_xlsx(out, xd.build_report(r, na, nb))
        return {"file": out.name, "path": str(out),
                "size": f"{out.stat().st_size / 1024:.0f} KB",
                "stats": r["stats"]}

    return ok(job=start_job(work))


@app.route("/api/export", methods=["POST"])
@api
def api_export():
    """前端把要导出的 sheets 传回来，后端落盘到产物目录"""
    p = request.get_json(force=True)
    name = core.safe_name(p.get("name") or "导出")
    fmt = (p.get("format") or "xlsx").lower()
    sheets = p.get("sheets") or {}
    if not sheets:
        return fail("没有可导出的内容")
    out = CLIENT.data_dir / f"{name}.{'csv' if fmt == 'csv' else 'xlsx'}"
    if fmt == "csv":
        headers, rows = next(iter(sheets.values()))
        core.write_csv(out, headers, rows)
    else:
        core.write_xlsx(out, {k: (v[0], v[1]) for k, v in sheets.items()})
    return ok(file=out.name, path=str(out),
              size=f"{out.stat().st_size / 1024:.0f} KB")


@app.route("/api/files")
def api_files():
    d = CLIENT.data_dir
    files = sorted((f for f in d.glob("*") if f.suffix.lower() in (".xlsx", ".csv")
                    and not f.name.startswith(".")),
                   key=lambda p: p.stat().st_mtime, reverse=True)[:60]
    return ok(items=[{"name": f.name, "size": f"{f.stat().st_size / 1024:.0f} KB",
                      "time": time.strftime("%m-%d %H:%M",
                                            time.localtime(f.stat().st_mtime))}
                     for f in files])


@app.route("/download/<path:name>")
def download(name):
    return send_from_directory(CLIENT.data_dir, name, as_attachment=True)


@app.route("/api/reveal", methods=["POST"])
@api
def api_reveal():
    """在 Finder / 资源管理器里打开产物目录"""
    import subprocess
    opener = {"darwin": "open", "win32": "explorer"}.get(sys.platform, "xdg-open")
    subprocess.Popen([opener, str(CLIENT.data_dir)])
    return ok()


# ── 启动 ────────────────────────────────────────────────────────


# Chrome/Firefox 会以 ERR_UNSAFE_PORT 拒绝访问这些端口，自动开浏览器会直接失败。
# 名单取自 Chromium net/base/port_util.cc 的 kRestrictedPorts（只列 >1024 的部分）。
UNSAFE_PORTS = {
    1719, 1720, 1723, 2049, 3659, 4045, 4190, 5060, 5061, 6000, 6566,
    6665, 6666, 6667, 6668, 6669, 6679, 6697, 10080,
}


def pick_port(preferred: int) -> int:
    """挑一个可用且浏览器不会拒绝的端口，避免和已有的 webapp(5050) 撞车"""
    for port in range(preferred, preferred + 40):
        if port in UNSAFE_PORTS:
            continue
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return preferred


def notify(title: str, message: str):
    """打包成 .app 后没有终端窗口，用系统弹窗把问题告诉用户"""
    print(f"{title}: {message}")
    if sys.platform != "darwin":
        return
    import subprocess
    body = message.replace('"', "'")[:900]
    try:
        subprocess.run(["osascript", "-e",
                        f'display dialog "{body}" with title "{title}" buttons {{"好"}} default button 1'],
                       capture_output=True, timeout=120)
    except Exception:
        pass


def main():
    frozen = getattr(sys, "frozen", False)
    log_path = core.CACHE_DIR.parent / "运行日志.log"
    if frozen:
        # .app 里 stdout 会被丢弃，转存到用户目录方便排查
        log_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            handle = open(log_path, "a", buffering=1, encoding="utf-8")
            sys.stdout = sys.stderr = handle
        except OSError:
            pass

    port = pick_port(int(CFG.get("server", {}).get("port", 5060)))
    url = f"http://127.0.0.1:{port}"
    problems = CLIENT.check_env()
    print("=" * 60)
    print(f"  Matrix AI 数据工具箱  {url}   （{time.strftime('%Y-%m-%d %H:%M:%S')}）")
    print(f"  配置: {CFG.get('_config_path')}")
    print(f"  产物: {CLIENT.data_dir}")
    if problems:
        print("  ⚠️ 环境自检：")
        for p in problems:
            print(f"     - {p}")
        if frozen:
            notify("Matrix AI 数据工具箱 · 环境提醒",
                   "界面能打开，但连库功能会失败：\\n\\n"
                   + "\\n".join(f"• {p}" for p in problems)
                   + f"\\n\\n请检查配置: {CFG.get('_config_path')}")
    print("=" * 60)
    print("  关闭本窗口即退出程序" if not frozen else "  退出：菜单栏右上角图标 → 退出")
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    try:
        app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
    except Exception as e:  # noqa: BLE001
        import traceback
        traceback.print_exc()
        if frozen:
            notify("Matrix AI 数据工具箱 启动失败",
                   f"{e}\\n\\n详细日志: {log_path}")
        raise


if __name__ == "__main__":
    main()
