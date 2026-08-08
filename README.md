# Matrix AI 数据工具箱（脱敏版）

本项目是一个只监听 `127.0.0.1` 的本地 Flask 工具箱，提供自然语言生成只读 SQL、AI 差异解释、JDBC 只读查询、库表统计、SQL 分析与数据比较、Excel/CSV 比较、Excel 转 INSERT/PDF/Word、文本与图片转 Excel 等能力。

本版本已移除原项目中的组织名称、真实连接地址、账号密码、环境路径、历史产物和业务样例。数据库连接与 AI 模型连接均由用户在本机前端运行时填写，只保存在当前进程内存中。

## 技术栈

- Python 3.10+
- Flask + Jinja2 + 原生 JavaScript
- Java/JDBC（支持配置驱动 JAR、驱动类和 JDBC URL）
- OpenAI Chat Completions 兼容接口（云端或本地模型，使用 Python 标准库调用）
- openpyxl（Excel 读写）
- fpdf2（Excel 转 PDF）
- python-docx（Excel 转 Word）
- macOS Vision OCR + Swift（图片识别）
- PyInstaller（可选，桌面应用打包）

## 安装与启动

Windows 推荐直接双击 `start_windows.bat`。首次运行会创建独立环境并安装依赖，以后双击即可启动。

也可以手动安装：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

Windows 激活虚拟环境：

```powershell
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

启动后程序会选择安全的本地端口并打开浏览器。默认首选端口为 `8765`；如果被占用，会自动顺延。

## 使用 AI SQL 助手

1. 点击页面右上角“AI 模型”。
2. 填写兼容 Chat Completions 的服务地址、模型名称和 API Key。本地免鉴权服务可以不填 Key。
3. 点击“保存并测试”。配置只存在内存中，退出程序即清除。
4. 在“AI SQL 助手”描述查询需求，并手动粘贴脱敏表结构；连接数据库后也可读取字段元数据。
5. 点击“生成只读 SQL”，核对模型说明和本地安全检查结果。
6. 点击“放入 SQL 查询页”，再由使用者确认并手动执行；系统不会自动执行模型生成结果。

外部服务默认必须使用 HTTPS。本机 `127.0.0.1`/`localhost` 模型可使用 HTTP；其他内网 HTTP 服务必须由使用者手动确认。不要把未经许可的公司表结构、业务数据、账号、密码或内部地址发送到外部模型。

## AI 增强 SQL 排查

“SQL 排查工作台”保留原有静态解析器作为证据层，并增加“AI 推理分析”：

1. 粘贴两段 SQL，并补充业务背景或预期变化。
2. 点击“比较逻辑（不连库）”，先得到确定的表、JOIN、条件、字段及规则风险。
3. 点击“AI 推理分析”，模型基于这些静态证据说明可能的语义影响，而不是直接猜测结果。
4. 页面明确区分直接证据、可能影响、待验证假设和结论边界。
5. 模型建议的验证 SQL 会再次经过本地只读检查；未通过的内容自动隐藏，通过的内容也只能手动放入查询页，不会自动执行。

AI 推理不能替代真实数据验证。复杂 SQL 超过单段 30000 字时，应先使用“长 SQL 拆分”选取相关逻辑层，避免模型因为上下文截断给出不完整结论。

## 在前端连接数据库

1. 点击页面右上角“数据库连接”。
2. 填写连接显示名称、JDBC URL、驱动类、驱动 JAR、JDK bin、用户名和密码。
3. URL 可以使用 `{cluster}` 与 `{db}` 占位符，例如 `jdbc:vendor://host:port/{db}`；固定 URL 也可以直接填写。若同一模板需要连接多个环境，可在“附加连接”中按 `显示名称=cluster值` 每行添加一个，A/B 对比下拉框会同时显示这些连接。
4. 输入只读测试 SQL（默认 `SELECT 1`），点击“测试并使用”。
5. 测试成功后页面刷新，查询和比较页面使用该活动连接。

请不要把账号或密码拼进 JDBC URL；后端会拦截常见的 URL 凭据写法。账号密码应填写在独立输入框中，避免进入 Java 进程参数。
检测到明文 HTTP 或显式关闭 SSL/TLS 的连接参数时，程序默认拒绝连接；只有确认数据库位于受控网络并手动勾选后才允许继续。此确认不能替代数据库侧的 TLS 和网络访问控制。

安全行为：

- 前端调用查询、上传和写入类 API 时使用当前进程生成的随机令牌，降低其他网页调用本地服务的风险。
- 密码不会通过状态接口返回，也不会自动回填。
- JDBC URL、SQL、用户名和密码通过子进程标准输入传给 Java，不放在命令行参数或环境变量中。
- 连接信息不写回 `config.json`，重启或点击“断开并清除”后失效。
- 数据库账号仍应在数据库侧授予最小化的只读权限。
- 有后台查询或比较任务运行时，程序会暂时拒绝切换或断开连接，避免同一任务跨到另一个数据库环境。

## 页面与功能

| 页面 | 主要功能 |
|---|---|
| AI SQL 助手 | 自然语言生成单条只读 SQL、本地安全检查、脱敏差异摘要解释 |
| SQL 排查工作台 | SQL 静态逻辑比较、AI 证据推理、只读验证建议、结构拆分、时间字段和主键候选、双侧取数比较 |
| SQL 查询 | 只读 SQL 执行、结果展示与导出 |
| 表行数统计 | 读取元数据行数；未知时可逐表执行精确 `COUNT(*)` |
| 库数据对比 | 表清单、行数和可配置表名映射比较 |
| 单表数据对比 | 字段、行数及按主键抽样明细比较 |
| 表对应关系 | 规则发现、预览、保存及手工维护表映射 |
| 文件比对 | Excel/CSV/TXT/TSV 按主键逐格比较，重复键组内最优配对 |
| 转 Excel | 结构化文本或图片 OCR 结果导出为 Excel |
| Excel 转换 | Excel 转 INSERT SQL、PDF 或 Word |
| 产物文件 | 查看和下载本地生成文件 |

## 配置文件

仓库只跟踪空白的 `config.example.json`，不含真实连接信息。程序没有本机配置时会自动读取该文件；实际账号密码不要写入仓库，优先使用前端运行时连接。

如果需要修改默认界面参数，可以编辑：

```json
{
  "server": {"port": 8765},
  "paths": {"workspace": ".", "jdk_bin": "", "driver_jar": "", "data_dir": "data"},
  "matrix": {"driver_class": "", "url_template": "", "username": "", "password": "", "clusters": {}},
  "ui": {"default_db": "", "sql_max_rows": 500, "stat_max_rows": 5000, "common_dbs": []}
}
```

## 打包

macOS 可执行：

```bash
bash build_app.sh
```

现有 `dist/`、`build/`、`.selftest/`、`data/`、`config.json` 和 `config.local.json` 均不应提交或直接分发。打包时只嵌入空白示例配置；发布前仍应再次扫描敏感信息并重新构建。

## 自检

```bash
python -m unittest discover -s tests -v
python -m py_compile app.py ai_tools.py matrix_core.py sql_tools.py excel_diff.py convert.py
```

## 已知边界

- AI 生成内容可能错误；必须人工核对。本地静态检查只负责拦截明显写操作和危险结构，不能证明 SQL 业务口径正确。
- 调用云端模型会把用户主动提交的需求、表结构或差异摘要发送给对应服务；是否允许发送由使用者和所在组织决定。
- JDBC 查询器使用 `executeQuery()`，所有调用统一经过只读静态检查；数据库账号仍必须限制为只读。
- 元数据行数可能滞后，精确统计需启用 `COUNT(*)`。
- 库表清单、明细比较和文件处理存在行数上限；大数据量结果需要分片或分页。
- 单表明细和普通 SQL 数据比较可能是上限内抽样，不等同于全表证明。
- 图片转 Excel 依赖 macOS Vision；其他系统需要替换 OCR 实现。
- PDF/Word/Excel 的具体版式和 OCR 结果应在导出后人工复核。
- 当前实现是 **Excel 转 PDF/Word**，不是 PDF/Word 转 Excel。
