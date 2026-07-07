# Orion 比价UPS RCA 归因系统 — 代码审查报告

> 审查日期：2026-07-07 | 审查范围：全部 6 个模块，16 个文件

---

## 一、总体概览

这是一个去哪儿网酒店价格比价 RCA（Root Cause Analysis）系统，用于分析用户认为去哪儿价格高于竞品（携程、美团、飞猪、同程）的原因。系统采用多模块 Python 流水线架构，通过 LLM（Claude API）驱动部分分析，Excel 文件作为数据交换媒介，`state.json` 作为编排状态管理。

**技术栈**：Python 3.14 / pandas / openpyxl / requests / anthropic SDK (Claude Sonnet 4.6)

### 模块清单

| 模块 | 主要文件 | 代码行数 | 模块职责 |
|------|---------|---------|---------|
| **attribution** | `attribute.py`, `eval.py` | 1205 + 292 | 归因主逻辑、测试框架 |
| **promo** | `validate.py`, `validate_fallback.py` | 674 + 975 | 促销校验（V1/V2）、兜底处理 |
| **fallback** | `execute.py` | 275 | 促销兜底子命令（LLM/Vision） |
| **im** | `io.py` | 247 | IM 归因数据读写 |
| **shared** | `field_map.json`, `field_definitions.md`, `io.md` | — | 字段映射配置、文档 |

---

## 二、严重问题

### 2.1 硬编码绝对路径（阻塞性）

以下路径硬编码了 `/Users/zhangwang/` 的 macOS 绝对路径，在 Windows 或其他用户环境下完全不可用：

| 文件 | 行号 | 硬编码路径 |
|------|------|-----------|
| `attribute.py` | 60 | `/Users/zhangwang/.claude/agents/orion-bijiaups-rca-shared` |
| `validate.py` | 40 | 同上 |
| `validate.py` | 41 | `/Users/zhangwang/.claude/agents/orion-bijiaups-rca-promo/code_tables/C_Q.xlsx` |
| `validate.py` | 42 | `/Users/zhangwang/.claude/agents/orion-bijiaups-rca-promo/code_tables/C_C.xlsx` |
| `validate.py` | 172 | `/Users/zhangwang/.claude/agents/orion-bijiaups-rca-promo/code_tables/E_QE.xlsx` |
| `validate_fallback.py` | 60-62 | 三条同上 |

**建议**：将所有路径改为基于 `__file__` 的相对路径，例如：
```python
RULE_ROOT = Path(__file__).parent.parent / "orion-bijiaups-rca-shared"
```

### 2.2 `io.py` 中 `field_map.json` 路径错误

`io.py:28`：`_FIELD_MAP_PATH = Path(__file__).parent / "field_map.json"` 在 `orion-bijiaups-rca-im/` 目录下查找 `field_map.json`，但实际文件在 `orion-bijiaups-rca-shared/` 中。`try/except` 静默吞掉 `FileNotFoundError`，导致 `field_map.json` 配置从未被 `io.py` 实际加载，始终使用硬编码的 `FIELDS` 回退值。

**建议**：修正路径指向 shared 目录，或将 `field_map.json` 复制到 im 目录。

### 2.3 无依赖管理文件

项目没有任何 `requirements.txt`、`pyproject.toml` 或 `setup.py`。依赖项（`pandas`、`openpyxl`、`requests`、`anthropic`）仅在代码中通过 `import` 或 `try/except` 提及。

**建议**：创建 `requirements.txt`，至少包含：
```
pandas
openpyxl
requests
anthropic
```

### 2.4 `--break-system-packages` 危险安装

多处使用 `pip3 install --break-system-packages openpyxl`（`attribute.py:1048`、`io.py:46`、`validate.py:112`），这是不安全的安装方式，暗示了环境管理问题。

**建议**：使用虚拟环境（venv/conda）管理依赖。

### 2.5 关键路径缺少异常处理

- `attribute.py:1178-1179`：`_copy_promo_sheets` 完成后直接 `os.remove(promo_file)` 删除源文件，无 try/except
- `validate.py` 中 `wb_input = openpyxl.load_workbook(input_file)` 无异常处理
- `validate_fallback.py:81-83`：`_fetch_image_base64` 有 try/except 但未处理 SSL 证书、超时等边界情况

**建议**：在文件 I/O、网络请求、外部 API 调用处添加 try/except 并记录错误。

### 2.6 `eval.py` 中 `q_df=None, c_df=None` 导致价格类测试无效

`attribute.py:1238-1239` 在 eval 中传递 `q_df=None, c_df=None`，所有价格类测试 case 的促销计算全部为 0，无法真正验证价格归因逻辑。

**建议**：为 eval 模式构造模拟促销数据，或从标准测试用例中定义促销输入。

---

## 三、架构问题

### 3.1 模块间耦合度高

```
orion-bijiaups-rca-shared (field_map.json)
    ↓ 硬编码绝对路径
orion-bijiaups-rca-im (io.py)           ← 独立模块
orion-bijiaups-rca-promo (validate.py)   ← 依赖 shared
    ↓ 写入 state.json
orion-bijiaups-rca-fallback (execute.py) ← 依赖 validate.py (import)
    ↓ 更新 state.json
orion-bijiaups-rca-attribution (attribute.py) ← 依赖所有上游
```

- Fallback 模块直接 `from validate import ...` 导入函数，存在循环依赖风险
- 只有 `state.json` 和 Excel 文件作为数据交换，错误传播难以追踪
- 无统一入口脚本，需手动按顺序执行各模块

**建议**：创建统一的 orchestrator 脚本串联流水线，使用明确的模块间 API（而非直接 import）。

### 3.2 字段映射机制重复且不一致

`FIELDS` 字典在 `attribute.py`、`validate.py`、`io.py` 三个文件中各自定义了一次默认值，然后通过 `load_field_map()` 从 `field_map.json` 加载覆盖：

| 文件 | 行号 | FIELDS 定义 |
|------|------|-------------|
| `attribute.py` | 64-84 | 完整定义（18 个字段） |
| `validate.py` | 60-71 | 部分定义（11 个字段） |
| `io.py` | 29-34 | 最小定义（5 个字段） |

这违反了 DRY 原则，且容易导致不同步。

**建议**：将 `FIELDS` 和 `load_field_map()` 统一定义在 shared 模块中，所有模块从 shared 导入。

### 3.3 无日志系统

所有模块使用 `print(f"[{datetime.now()}] ...", file=sys.stderr)` 手动格式化的日志，无统一的日志级别、无日志文件输出。

**建议**：用 Python 标准库 `logging` 替换，支持日志级别和文件输出。

### 3.4 无配置文件管理

`state.json` 动态写入运行状态，但无 schema 验证。字段随意增删，容易产生不一致。

**建议**：使用 Pydantic 或 dataclass 定义状态结构，读写时做 schema 校验。

### 3.5 无单元测试，仅有基于文档的集成测试

`eval.py` 从 Markdown 文档解析测试用例，不是真正的单元测试。无 pytest/unittest 框架，无法在 CI 中自动运行。

**建议**：引入 pytest，为核心函数（如 `route_level1()`、`attribute_price()` 等）编写单元测试。

---

## 四、代码质量问题

### 4.1 大量重复代码

以下代码在多个文件中出现完全或几乎完全相同的实现：

| 重复内容 | 出现位置 | 说明 |
|---------|---------|------|
| `FIELDS` 字典 | `attribute.py:64`, `validate.py:60`, `io.py:29` | 相同字段键，不同子集 |
| `load_field_map()` | `attribute.py:87`, `validate.py:74` | 逐行完全一致 |
| `normalize_platform()` | `attribute.py:113`, `validate.py:127` | 完全一致 |
| `require_platform()` | `attribute.py:118`, `validate.py:132` | 几乎一致 |
| `normalize_rule_set()` | `attribute.py:129`, `validate.py:143` | 完全一致 |
| `PLATFORM_ALIASES` | `attribute.py:48`, `validate.py:44` | 完全一致 |
| `VALID_PLATFORMS` | `attribute.py:46`, `validate.py:39` | 完全一致 |
| `_safe_float()` | `attribute.py:143`, `validate.py:344` | 完全一致 |
| `_log()` | 全部 5 个 .py 文件 | 相同模式 |
| `check_openpyxl()` | `io.py:42`, `validate.py:107` | 完全一致 |
| `load_full_state()` / `_load_state()` | `attribute.py:136`, `validate.py:161`, `validate_fallback.py:70` | 相似模式 |

**建议**：提取公共代码到 `orion-bijiaups-rca-shared/common.py`，所有模块统一导入。

### 4.2 函数过长

| 函数 | 文件 | 行数 | 建议 |
|------|------|------|------|
| `main()` | `attribute.py:1007` | 197 行 | 拆分为 init、process、output 三个子函数 |
| `main()` | `validate.py:442` | 231 行 | 同上 |
| `attribute_row()` | `attribute.py:722` | 104 行 | 按一级路由拆分，每个场景一个独立函数 |
| `run_fallback()` | `validate_fallback.py:804` | 146 行 | 按侧（Q/C）拆分 |
| `run_batch_fallback()` | `validate_fallback.py:722` | 77 行 | 可接受 |

### 4.3 魔法数字

| 数值 | 位置 | 含义 | 建议命名 |
|------|------|------|---------|
| `1.0` | `validate.py:361,367`, `validate_fallback.py:228,238` | V1 金额平衡容差（元） | `AMOUNT_BALANCE_TOLERANCE` |
| `0.1` | `attribute.py:401,407,606` | 划线价匹配容差（元） | `LIST_PRICE_MATCH_TOLERANCE` |
| `0.75` | `validate_fallback.py:248` | 模糊匹配阈值 | `FUZZY_MATCH_CUTOFF` |
| `50` | `attribute.py`, `validate.py` | 大额券阈值（元） | `LARGE_COUPON_THRESHOLD` |

### 4.4 中文硬编码分散

归因逻辑中的中文关键词分散在 `route_level1()`、`attribute_*()` 等函数中，没有统一的常量/枚举管理：

- `'民宿'`、`'团购'`、`'已下单'`、`'非C平台'` — 路由场景关键词
- `'酒店不一致'`、`'房型不一致'`、`'同质化权益'` — 用户误解子类
- `'同质化lose'`、`'价格高'`、`'去哪儿价格高'` — 价格 lose 子类
- `'酒店缺失'`、`'房型缺失'`、`'物理房型缺失'` — 库存 lose 子类

**建议**：定义枚举或常量类管理这些关键词。

### 4.5 类型提示缺失

所有函数缺少类型注解（Type Hints），降低了代码可维护性和 IDE 支持。

**建议**：逐步添加类型注解，使用 `dict[str, Any]`、`pd.DataFrame`、`Optional[str]` 等。

### 4.6 `attribute_row()` 返回字典包含隐式字段

函数返回 `dict` 包含 `_manual_reasons` 和 `_l3_l4_paths` 两个下划线前缀的隐式字段，在调用方通过 `pop()` 取出。这种设计不够清晰。

**建议**：使用 dataclass 或 NamedTuple 作为返回值，或将隐式字段合并到正式的返回结构中。

### 4.7 `eval.py` 的测试用例解析脆弱

`parse_cases_md()` 使用正则解析 Markdown 文件，依赖精确定义的格式，容易因格式变化而静默失败。

**建议**：使用 YAML 或 JSON 格式定义测试用例，或使用 `pytest` 参数化测试。

---

## 五、安全与数据问题

### 5.1 API Key 管理

`validate_fallback.py:822`：`os.environ.get("ANTHROPIC_API_KEY")`，虽然这是标准做法，但无验证或友好提示。如果 Key 未设置，错误信息不明确。

**建议**：在启动时检查 API Key 是否存在，给出清晰的错误提示。

### 5.2 Excel 文件无数据验证

所有模块从 Excel 读取数据时不做 schema 验证，依赖列名精确匹配。列名变化会导致静默错误（读取到空值）。

**建议**：在读取后验证关键列是否存在，缺失时给出明确错误提示。

### 5.3 `state.json` 无并发保护

多模块顺序读写 `state.json`，如果并发运行会丢失数据。

**建议**：引入文件锁（`fcntl` / `msvcrt`）或使用 SQLite 作为状态存储。

---

## 六、文档质量

**优点**：
- 每个模块的 `.md` 文件与代码有明确的同步规则声明
- `field_definitions.md` 完整记录了所有字段
- `attribution_logic.md`（459 行）归因树规则详尽
- 标准测试用例（C 携程 22 个 + E 同程 12 个）覆盖较好

**不足**：
- 无架构设计文档
- 无部署文档（环境要求、安装步骤）
- 无快速入门指南
- `io.md` 在 shared 目录但其内容仅针对 IM 模块，位置有误导性

---

## 七、改进建议优先级

| 优先级 | 问题 | 建议 |
|--------|------|------|
| **P0** | 硬编码绝对路径 | 改为基于 `__file__` 的相对路径 |
| **P0** | 无依赖文件 | 创建 `requirements.txt` |
| **P0** | `io.py` 中 `field_map.json` 路径错误 | 修正路径指向 shared 目录 |
| **P1** | 公共代码重复 | 提取到 `shared/common.py` |
| **P1** | 关键路径无异常处理 | 文件 I/O 和网络请求处添加 try/except |
| **P1** | 无日志系统 | 引入 `logging` 模块 |
| **P1** | `eval.py` 价格类测试无效 | 构造模拟促销数据 |
| **P2** | 函数过长 | 拆分 `main()` 和 `attribute_row()` |
| **P2** | 字段映射重复 | 统一由 shared 模块导出 |
| **P2** | 类型提示缺失 | 逐步添加 type hints |
| **P2** | 魔法数字 | 提取为命名常量 |
| **P3** | 无单元测试 | 引入 pytest |
| **P3** | `state.json` 并发保护 | 引入文件锁 |
| **P3** | 中文硬编码 | 提取为常量/枚举 |
| **P3** | `attribute_row()` 隐式字段 | 使用 dataclass |

---

## 八、模块维度总结

| 模块 | 代码量 | 质量评级 | 主要问题 |
|------|--------|---------|---------|
| **attribution** | 1205 行 | B | 函数过长、硬编码路径、重复代码、price eval 无效 |
| **promo** | 1649 行 | B- | 两个 main 函数过长、硬编码路径、重复代码最多 |
| **fallback** | 275 行 | B+ | 代码相对清晰，但依赖 validate.py 的 import |
| **im** | 247 行 | A- | 结构最清晰，但 `field_map.json` 路径有 bug |
| **shared** | 3 个文件 | C | 缺少可复用的 Python 代码，仅有 JSON 和 MD |
| **eval** | 292 行 | B | 测试框架设计合理但覆盖不全、解析脆弱 |

---

## 九、文件清单

```
orion-bijiaups-rca-attribution/
├── attribute.py          (1205 行) 归因主逻辑
├── attribution_logic.md  (459 行)  归因规则文档
├── eval.py              (292 行)  测试框架
├── validation_logic.md   (91 行)  校验规则文档
└── standard_cases/
    ├── C_standard_cases.md  (494 行) 携程标准用例 × 22
    └── E_standard_cases.md  (348 行) 同程标准用例 × 12

orion-bijiaups-rca-promo/
├── validate.py           (674 行)  促销校验主逻辑
├── validate_fallback.py  (975 行)  兜底处理
└── promo_logic.md        (218 行)  促销校验规则文档

orion-bijiaups-rca-fallback/
├── execute.py            (275 行)  兜底子命令
└── scenarios/
    └── promo_v2.md       (149 行)  兜底场景文档

orion-bijiaups-rca-im/
├── io.py                 (247 行)  IM 数据读写
└── references/
    └── logic.md          (250 行)  IM 归因逻辑文档

orion-bijiaups-rca-shared/
├── field_map.json         (26 行)  字段映射配置
├── field_definitions.md   (98 行)  字段定义文档
└── io.md                 (115 行)  IM 模块 I/O 文档
```