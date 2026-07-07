#!/usr/bin/env python3
"""
归因自校验脚本 — 对比 standard_cases/ 中的预期输出与 attribute.py 的实际输出。

用法：
  python3 eval.py --platform 携程
  python3 eval.py --platform 同程
  python3 eval.py                   # 运行全部平台
"""

import argparse
import re
import sys
from pathlib import Path

# ── 将 attribution/ 目录加入路径，直接复用 attribute.py 的函数 ──
_HERE = Path(__file__).parent
sys.path.insert(0, str(_HERE))

from attribute import (
    FIELDS,
    load_field_map,
    attribute_row,
    _safe_float,
)

CASES_DIR = _HERE / "standard_cases"

PLATFORM_FILE = {
    "携程": CASES_DIR / "C_standard_cases.md",
    "同程": CASES_DIR / "E_standard_cases.md",
}

PLATFORM_RULE_SET = {
    "携程": "xiecheng",
    "同程": "tongcheng",
}

OUTPUT_KEYS = ["一级_AI", "二级_AI", "三级_AI", "四级_AI"]


# ═══════════════════════════════════════════
# 解析 standard_cases/*.md
# ═══════════════════════════════════════════

def _strip_comment(line: str) -> str:
    return line.split("#")[0].strip()


def _parse_kv(line: str) -> tuple[str, str] | None:
    """解析 'Key: Value' 或 'Key：Value' 格式，返回 (key, value) 或 None。"""
    for sep in (": ", "：", ": "):
        if sep in line:
            k, _, v = line.partition(sep)
            return k.strip(), v.strip()
    return None


def parse_cases_md(filepath: Path) -> list[dict]:
    """
    从 standard_cases md 文件解析所有 case。
    每个 case 返回 {'title': str, 'input': dict, 'expected': dict}。
    输入字段直接用中文列名（与 FIELDS 值匹配）。
    """
    text = filepath.read_text(encoding="utf-8")
    # 按 ### Case 分割
    blocks = re.split(r'^### Case ', text, flags=re.MULTILINE)

    cases = []
    for block in blocks[1:]:
        lines = block.splitlines()
        title = lines[0].strip()

        # 提取 **输入** 和 **预期输出** 代码块
        input_section = _extract_code_block(block, "输入")
        expected_section = _extract_code_block(block, "预期输出")

        if not input_section or not expected_section:
            continue

        input_fields = _parse_input_section(input_section)
        expected = _parse_expected_section(expected_section)

        if expected:
            cases.append({"title": title, "input": input_fields, "expected": expected})

    return cases


def _extract_code_block(text: str, section_keyword: str) -> str | None:
    """提取指定小节下的第一个代码块内容。"""
    pattern = rf'\*\*{section_keyword}\*\*.*?```.*?\n(.*?)```'
    m = re.search(pattern, text, re.DOTALL)
    return m.group(1).strip() if m else None


def _parse_input_section(text: str) -> dict:
    """将输入代码块解析为 {列名: 值} 字典。"""
    row = {}
    current_key = None
    promo_lines = []
    in_promo = False

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        # 促销明细（多行子项）
        if line.startswith("Q促销:") or line.startswith("Q优惠明细:"):
            current_key = "Q优惠明细"
            in_promo = True
            promo_lines = []
            val = line.split(":", 1)[1].strip()
            if val:
                promo_lines.append(val)
            continue
        if line.startswith("E促销:") or line.startswith("C优惠明细:"):
            current_key = "促销明细（竞品）"
            in_promo = True
            promo_lines = []
            val = line.split(":", 1)[1].strip()
            if val:
                promo_lines.append(val)
            continue

        if in_promo and line.startswith("-") or (in_promo and "¥" in line and not ":" in line.split("¥")[0]):
            # 子项行
            promo_lines.append(line.lstrip("-").strip())
            continue

        # 非子项 → 结束促销收集
        if in_promo:
            row[current_key] = "\n".join(promo_lines)
            in_promo = False
            promo_lines = []

        # 普通 key: value（支持 / 分隔多个同行字段）
        parts = [p.strip() for p in line.split(" / ")]
        for part in parts:
            kv = _parse_kv(part)
            if kv:
                k, v = kv
                mapped = _map_input_key(k)
                if mapped:
                    row[mapped] = v

    if in_promo:
        row[current_key] = "\n".join(promo_lines)

    return row


_INPUT_KEY_MAP = {
    # 携程
    "客服判断结果": "客服判断场景",
    "客服判断场景": "客服判断场景",
    "Q划线价": "Q划线价——同质化房型",
    "Q划线价——同质化房型": "Q划线价——同质化房型",
    "竞品划线价": "划线价（竞品）",
    "划线价（竞品）": "划线价（竞品）",
    "Q支付价": "Q支付价——同质化房型",
    "Q支付价——同质化房型": "Q支付价——同质化房型",
    "竞品支付价": "到手价（竞品）",
    "到手价（竞品）": "到手价（竞品）",
    "Q RoomID": "同质化房型roomid",
    "同质化房型roomid": "同质化房型roomid",
    "C RoomID": "CRoomID",
    "CRoomID": "CRoomID",
    "是否分销": "是否分销",
    "是否露出": "是否露出",
    "领用券情况": "领用券情况",
    "im-归因结果-AI": "im-归因结果-AI",
    "比价平台": "比价平台",
    # 同程
    "备注": "客服判断场景",
    "是否多倍积分酒店": "是否多倍积分酒店",
}


def _map_input_key(k: str) -> str | None:
    return _INPUT_KEY_MAP.get(k)


def _parse_expected_section(text: str) -> dict:
    """解析预期输出行，返回 {一级_AI: ..., 二级_AI: ..., ...}。"""
    result = {}
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in line.split(" / ")]
        for part in parts:
            kv = _parse_kv(part)
            if kv:
                k, v = kv
                if k in OUTPUT_KEYS:
                    result[k] = v if v != "（无）" else ""
    return result


# ═══════════════════════════════════════════
# 执行单个平台的校验
# ═══════════════════════════════════════════

def run_platform(platform: str) -> tuple[int, int]:
    rule_set = PLATFORM_RULE_SET[platform]
    filepath = PLATFORM_FILE[platform]

    if not filepath.exists():
        print(f"[{platform}] ⚠ 找不到 standard_cases 文件: {filepath}")
        return 0, 0

    # 加载字段映射
    loaded = load_field_map(rule_set)
    if loaded:
        FIELDS.update(loaded)

    cases = parse_cases_md(filepath)
    print(f"\n{'═'*60}")
    print(f"平台: {platform}  |  共 {len(cases)} 个 case")
    print(f"{'═'*60}")

    passed = 0
    failed = 0

    for case in cases:
        row = case["input"]
        expected = case["expected"]
        title = case["title"]

        # 数值字段转换
        for fk in ("q_list_price", "q_pay_price", "c_list_price", "c_pay_price"):
            col = FIELDS.get(fk, "")
            if col in row:
                row[col] = _safe_float(row[col])

        actual_full = attribute_row(
            row,
            q_df=None,
            c_df=None,
            price_rule_set=rule_set,
            platform=platform,
            promo_code_set=rule_set,
            case_audit={},
        )
        actual_full.pop("_manual_reasons", None)

        mismatches = []
        for key in OUTPUT_KEYS:
            exp_val = expected.get(key, "")
            act_val = str(actual_full.get(key, "") or "").strip()
            exp_val = str(exp_val or "").strip()
            if exp_val and act_val != exp_val:
                mismatches.append(f"  {key}: 预期={exp_val!r}  实际={act_val!r}")

        if mismatches:
            failed += 1
            print(f"❌ Case {title}")
            for m in mismatches:
                print(m)
        else:
            passed += 1
            print(f"✅ Case {title}")

    print(f"\n结果: {passed} 通过 / {failed} 失败 / {len(cases)} 总计")
    return passed, failed


def main():
    parser = argparse.ArgumentParser(description="归因自校验脚本")
    parser.add_argument("--platform", choices=["携程", "同程"], default=None,
                        help="指定平台（不指定则运行全部）")
    args = parser.parse_args()

    platforms = [args.platform] if args.platform else list(PLATFORM_FILE.keys())

    total_pass = total_fail = 0
    for p in platforms:
        p_pass, p_fail = run_platform(p)
        total_pass += p_pass
        total_fail += p_fail

    print(f"\n{'═'*60}")
    print(f"全平台汇总: {total_pass} 通过 / {total_fail} 失败")
    if total_fail > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
