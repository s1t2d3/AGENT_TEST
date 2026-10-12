# AGENT_TEST# AGENT_TEST

基于 Python + Pytest 的**数据驱动接口自动化测试框架**。测试用例维护在 Excel 中，框架自动读取用例、渲染变量、发送 HTTP / JDBC 请求、执行断言与数据提取，并生成 Allure 可视化报告。项目已集成 Docker 与 GitHub Actions，支持 CI 自动执行并将报告发布到 GitHub Pages。

![Python](https://img.shields.io/badge/python-3.11-blue)
![Pytest](https://img.shields.io/badge/pytest-9.1.1-green)
![Allure](https://img.shields.io/badge/allure-2.27.0-orange)

---

## 核心特性

- **Excel 数据驱动**：用例统一维护在 `data/测试用例.xlsx`，无需改代码即可增删用例
- **变量传递与模板渲染**：支持 `${xxx}` 占位符，自动替换为前序用例提取的全局变量
- **接口关联**：支持 JSONPath 提取响应数据、JDBC 从数据库提取数据，供后续用例使用
- **HTTP / JDBC 双协议**：既可发 HTTP 请求，也可直接执行数据库操作
- **Session 管理**：支持登录态复用（`use_session=True`）与免登录请求（`use_session=False`）
- **Allure 报告**：自动生成 JSON 与 HTML 报告，报告美观、可追溯
- **日志记录**：终端实时日志 + 文件日志（`log/test.log`）
- **持续集成**：GitHub Actions 自动拉起被测服务、执行测试、发布报告到 `gh-pages`

---

## 技术栈

Python / Pytest / Allure / Requests / openpyxl / PyMySQL / JSONPath / Docker / GitHub Actions

---

## 项目结构

```
AGENT_TEST/
├── .github/workflows/
│   └── main.yml              # GitHub Actions 工作流
├── config/
│   └── config.py             # 全局配置：BASE_URL、Excel 路径、MySQL 连接
├── data/
│   └── 测试用例.xlsx          # 测试用例数据源
├── log/
│   └── test.log              # 运行日志
├── report/
│   ├── json_report/          # Allure 原始 JSON 数据
│   └── html_report/          # Allure HTML 报告
├── testcases/
│   └── test_runner.py        # 用例执行入口（核心调度）
├── utils/
│   ├── allure_util.py        # Allure 报告初始化
│   ├── asserts.py            # 断言工具
│   ├── case_data.py          # 用例数据处理
│   ├── data_extraction.py    # JSONPath / JDBC 数据提取
│   ├── excel_util.py         # Excel 读取工具
│   └── send_requests.py      # HTTP / JDBC 请求发送
├── Dockerfile
├── pytest.ini                # pytest 与日志配置
├── requirements.txt
└── run.py                    # 统一执行入口
```

---

## 环境要求

- Python 3.11（CI 使用版本，本地 3.8+ 亦可）
- pip
- Allure CLI 2.27.0（生成 HTML 报告需要）
- （可选）Docker

---

## 快速启动

### 1. 克隆项目

```bash
git clone https://github.com/s1t2d3/AGENT_TEST.git
cd AGENT_TEST
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 安装 Allure CLI

**macOS（Homebrew）**

```bash
brew install allure
```

**Linux**

```bash
curl -o allure.tgz -L https://github.com/allure-framework/allure2/releases/download/2.27.0/allure-2.27.0.tgz
tar -zxvf allure.tgz
sudo mv allure-2.27.0 /opt/allure
export PATH=$PATH:/opt/allure/bin
```

**Windows**：从 [Allure Releases](https://github.com/allure-framework/allure2/releases) 下载压缩包，解压后将 `bin` 目录加入系统 `PATH`。

### 4. 准备测试数据

- 用例文件：`data/测试用例.xlsx`，Sheet 名为 `测试用例`
- 被测服务地址：修改 `config/config.py` 中的 `BASE_URL`（默认 `http://127.0.0.1:5005`）

### 5. 配置数据库密码

JDBC 提取或数据库断言需要连接 MySQL，密码通过环境变量注入。

**Linux / macOS**

```bash
export DB_PASSWORD=your_db_password
```

**Windows PowerShell**

```powershell
$env:DB_PASSWORD="your_db_password"
```

数据库连接参数在 `config/config.py` 中配置（`HOST` / `PORT` / `USER` / `DB` / `CHARSET`）。

### 6. 运行测试

```bash
python run.py
```

`run.py` 实际执行的动作：

1. 运行 `testcases/test_runner.py`，输出 Allure 原始数据到 `report/json_report`
2. 调用 `allure generate` 生成 HTML 报告到 `report/html_report`

运行完成后，用浏览器打开 `report/html_report/index.html` 即可查看报告。

也可以直接用 pytest 执行：

```bash
pytest -vs testcases/test_runner.py --alluredir report/json_report --clean-alluredir
```


---

## Docker 运行

构建镜像：

```bash
docker build -t agent_test .
```

运行容器：

```bash
docker run --rm -e DB_PASSWORD=your_db_password agent_test
```

---

## 持续集成（GitHub Actions）

工作流文件：`.github/workflows/main.yml`，触发条件为 push / PR 到 `main` 分支。

执行流程：

1. 拉取并启动被测服务镜像 `ghcr.io/s1t2d3/std:86eaae...`（映射 5005 端口）
2. 检出代码、安装 Python 3.11 与依赖
3. 安装 Allure CLI 2.27.0
4. 执行 `python run.py`（`DB_PASSWORD` 从 Secrets 注入）
5. 无论测试是否失败，都将 `report/html_report` 发布到 `gh-pages` 分支

**使用前需配置**：在仓库 **Settings → Secrets and variables → Actions** 中添加 `DB_PASSWORD`。

报告发布后可通过 `https://s1t2d3.github.io/AGENT_TEST/` 访问（前提是已在仓库 Settings → Pages 中选择 `gh-pages` 分支）。

---

## 测试报告与日志

- **Allure 原始数据**：`report/json_report/`
- **Allure HTML 报告**：`report/html_report/`
- **运行日志**：`log/test.log`，格式为 `时间 - 级别 - 文件:行号 - 消息`，同时输出到终端

日志配置见 `pytest.ini`。

---

## 依赖清单

| 组件 | 版本 / 用途 |
| --- | --- |
| pytest | 9.1.1，测试框架 |
| allure-pytest | 2.16.2，Allure 报告集成 |
| requests | 2.34.2，HTTP 请求 |
| openpyxl | 3.1.5，读取 Excel 用例 |
| pymysql | 1.2.3，MySQL 连接与 JDBC 提取 |
| jsonpath | 0.82.2，响应数据提取 |

---

## 未来计划

- [ ] 支持更多协议（如 WebSocket、gRPC）
- [ ] 用例数据支持 YAML / JSON 格式
- [ ] 增加测试数据工厂与 Mock 能力
- [ ] 优化报告展示与失败重试机制
- [ ] 完善单元测试与文档

---

## 贡献

欢迎提交 Issue 与 Pull Request。

---

## 许可证

待添加
