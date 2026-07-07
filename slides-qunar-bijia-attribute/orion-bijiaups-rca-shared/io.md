# IM归因 Excel I/O 接口文档

> ⚠️ **修改联动**：本文件与 `io.py`（同目录）强绑定。
> - 改接口行为 → 先更新本文件，再同步修改 `io.py`
> - 改代码 → 先确认本文件描述，修改后同步更新本文件
> 两者不同步 = 接口文档与代码行为不一致

---

## 功能概述

`io.py` 是 IM归因模块专用的 Excel I/O 工具，提供三个子命令：`init`、`read`、`write`。

所有操作依赖 `openpyxl`。缺少时报错并提示安装命令。

---

## 子命令：init

**用途**：从原始输入 Excel 创建主数据输出文件（仅含 Sheet1=主数据）。

**调用**：
```bash
python3 io.py init --input <原始.xlsx> [--output <输出.xlsx>]
```

**行为**：
1. 复制原始文件到输出路径
2. 确保存在名为「主数据」的 sheet（通过 case编号 列自动识别，否则取第一个 sheet 改名）
3. 清理所有其他 sheet，只保留「主数据」
4. 在 stderr 打印结构化日志，在 stdout 输出 `OUTPUT_PATH:<路径>` 和完成提示

**输出文件命名规则**：`原目录/原文件名_AI主数据_YYYYmmdd_HHMM.xlsx`

**参数**：
| 参数 | 必须 | 说明 |
|------|------|------|
| `--input` | 是 | 原始输入 Excel 文件路径 |
| `--output` | 否 | 指定输出路径（不填则自动生成） |

---

## 子命令：read

**用途**：从输入 Excel 读取 IM 对话数据，输出 JSON 数组到 stdout。

**调用**：
```bash
python3 io.py read --input <文件.xlsx>
```

**列识别逻辑**（字段名通过 `shared/field_map.json` 配置，精确匹配）：
| 列 | FIELDS key | 当前 Excel 列名 |
|----|-----------|----------------|
| IM对话列 | `im_text` | `客服与用户沟通记录` |
| case编号列 | `case_no` | `case编号` |
| 工单号列 | `order_no` | `工单号` |
| 客服判断列 | `kf_note` | `客服判断场景` |

**输出格式**：
```json
[
  {
    "row_num": 2,
    "case_id": "CASE001",
    "flow_no": "工单001",
    "im_text": "用户: 我觉得贵 → 客服: ...",
    "cs_judgment": "可对比-lose"
  }
]
```

**错误处理**：
- 未找到精确列名 `客服与用户沟通记录` → 打印所有列标题，退出 code 1

---

## 子命令：write

**用途**：将 IM 归因结果写回主数据文件的 Sheet1（主数据）。

**调用**：
```bash
python3 io.py write --output <主数据.xlsx> --results <结果.json>
```

**写入列**：`im-归因结果-AI` 和 `im-归因说明-AI`（若列不存在则自动追加到末尾）

**results JSON 格式**：
```json
[
  {
    "row_num": 2,
    "case_id": "CASE001",
    "short_labels": "入离不一致",
    "explanation": "客服指出用户对比日期不同，有效判断依据：..."
  }
]
```

**参数**：
| 参数 | 必须 | 说明 |
|------|------|------|
| `--output` | 是 | 主数据文件路径（IM归因模块 init 阶段创建的文件） |
| `--results` | 是 | 归因结果 JSON 文件路径 |

---

## 日志格式

所有日志输出到 stderr，格式：
```
[YYYY-MM-DD HH:MM:SS] 消息内容
```
