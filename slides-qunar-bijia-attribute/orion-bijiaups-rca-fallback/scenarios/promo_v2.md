# 场景：promo_v2_unmatched

## 触发模块
促销校验（validate.py）— V1金额不平衡 或 V2促销归类失败：促销名称在码表中未匹配

## 核心原则

**V1（金额平衡）和 V2（类型识别）是两个独立问题，各自独立判断、独立写入。**

- V2 解决 = 所有未匹配的促销名称能够被识别出类型，且通过 check-v2 码表验证
- V1 解决 = 重新解析的金额能通过平衡校验（划线价 - 促销合计 = 到手价，容差 ±1 元）
- Q侧和竞品侧独立处理，各自的优化建议只写各自分行表

## 读取字段（从原始 Excel）
- `Q优惠明细`：Q侧原始促销文本
- `促销明细（竞品）`：竞品侧原始促销文本
- `明细页-价格优惠明细页`：竞品促销截图 URL（可为空）
- `Q折后底价——同质化房型`/`Q折后底价——物理房型最低价`：Q划线价
- `Q到手价——同质化房型`/`Q到手价——物理房型最低价`：Q到手价
- `划线价（竞品）`：竞品划线价
- `到手价（竞品）`：竞品到手价
- `promo_code_set`：从 state.json 的 input.promo_code_set 读取

## Agent 任务

### 处理顺序

**竞品侧**：先兜底2（截图）→ 失败再兜底1（文本+搜索）
**Q侧**：仅兜底1（文本+搜索）

---

### 兜底2：截图识别竞品促销（优先）

若 `明细页-价格优惠明细页` URL 存在，用 Vision 读取截图，提取竞品每条促销的名称和金额。

**识别出内容后**：
1. 用 check-v2 验证竞品侧名称是否在码表中
2. 做 V1 平衡校验：`|竞品划线价 - (截图金额合计 + 竞品到手价)| ≤ 1.0`

**兜底2结果 → c_tier**：
- V1✅ V2✅ → `兜底2`
- V1❌ V2✅ → `兜底2-V1待处理`
- V1✅ V2❌ → `兜底2-V2待码表更新`
- V1❌ V2❌ / 截图无内容 → 进入兜底1

---

### 兜底1：文本推断（搜索 + 解析）

**步骤1a — 网络搜索（优先）**

从 `errors` 字段提取未匹配的促销名称，对每个名称搜索：
```
搜索词：<平台名称> <促销名称>
例：同程旅行 超值优惠
```
根据搜索结果判断促销类型（商促/平促/平券/积分/返现等）。

**步骤1b — 文本推断**

阅读原始促销文本，重新提取每条促销的名称和金额，判断类型。

**完成后对每侧分别**：
1. 用 check-v2 验证名称是否在码表中
2. 做 V1 平衡校验

**兜底1结果 → q_tier / c_tier**：
- V1✅ V2✅ → `兜底1`
- V1❌ V2✅ → `兜底1-V1待处理`
- V1✅ V2❌ → `兜底1-V2待码表更新`
- V1❌ V2❌ / 无法解析 → `需人工介入`

---

### check-v2 使用方法

解析出促销名称后，调用 check-v2 验证是否在码表中：

```bash
python3 ~/.claude/agents/orion-bijiaups-rca-fallback/execute.py check-v2 \
  --promo-names '["促销名1", "促销名2"]' \
  --side q|c \
  --promo-code-set tongcheng|xiecheng
```

返回：`{"all_matched": true/false, "matched": [...], "unmatched": [...]}`

- `all_matched=true` → V2 已解决
- `all_matched=false` → V2 未解决，`unmatched` 列表说明哪些名称仍未匹配

---

## 输出格式（JSON，Q侧和竞品侧独立）

```json
{
  "q_tier": "兜底1 | 兜底1-V1待处理 | 兜底1-V2待码表更新 | 需人工介入",
  "c_tier": "兜底2 | 兜底2-V1待处理 | 兜底2-V2待码表更新 | 兜底1 | 兜底1-V1待处理 | 兜底1-V2待码表更新 | 需人工介入",
  "q_optimization": "Q侧码表优化建议（仅与Q侧相关）",
  "c_optimization": "竞品侧码表优化建议（仅与竞品侧相关）",
  "q_amounts": [12.0, 5.0],
  "c_amounts": [15.0],
  "q_v1_resolved": true,
  "q_v2_resolved": true,
  "c_v1_resolved": false,
  "c_v2_resolved": true
}
```

字段说明：
- `q_tier` / `c_tier`：各侧兜底方式，独立判断
- `q_optimization` / `c_optimization`：各侧优化建议，不跨侧混用
- `q_amounts` / `c_amounts`：仅 v1_resolved=true 时提供（按分行表行顺序）
- `*_v1_resolved`：V1金额校验通过
- `*_v2_resolved`：V2码表匹配通过（由 check-v2 验证）

如果某侧原始 errors 中没有该侧的错误，对应 tier 输出 `""` 且不调用 write-promo。

---

## 写入方式

**每侧调用一次 write-promo，不混写**：

```bash
# Q侧（如有Q错误）
python3 ~/.claude/agents/orion-bijiaups-rca-fallback/execute.py write-promo \
  --promo-file "<promo_file>" \
  --case-id "<case_id>" \
  --tier "<q_tier>" \
  --optimization "<q_optimization>" \
  --q-amounts '<JSON数组>' \
  --v1-resolved \   # 仅 q_v1_resolved=true 时加
  --v2-resolved \   # 仅 q_v2_resolved=true 时加
  --side q

# 竞品侧（如有竞品错误）
python3 ~/.claude/agents/orion-bijiaups-rca-fallback/execute.py write-promo \
  --promo-file "<promo_file>" \
  --case-id "<case_id>" \
  --tier "<c_tier>" \
  --optimization "<c_optimization>" \
  --c-amounts '<JSON数组>' \
  --v1-resolved \   # 仅 c_v1_resolved=true 时加
  --v2-resolved \   # 仅 c_v2_resolved=true 时加
  --side c
```
