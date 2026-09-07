# -*- coding: utf-8 -*-
"""临时验证脚本：检查 verification_rules.yaml 结构完整性"""
import sys
from pathlib import Path
import yaml

cfg = Path("config/verification_rules.yaml")
data = yaml.safe_load(cfg.read_text(encoding="utf-8"))

print("YAML 解析成功:", cfg)
print("版本:", data["version"], "| 更新:", data["updated"])

person_types = data["person_types"]
print("人员类型:", ", ".join(person_types.keys()))
for key, v in person_types.items():
    docs = [d["name"] for d in v["docs"]]
    required = [d["name"] for d in v["docs"] if d.get("required")]
    print(f"\n[{v['name']}]")
    print("  资料清单(%d项):" % len(docs), "、".join(docs))
    print("  必附:", "、".join(required))
    print("  照片要求:", "、".join(v["photos_required"]))
    print("  证件页:", "、".join(v["id_pages_required"]))

print("\n--- 合同逻辑 ---")
rule = data["contractor_declaration_check"]["contract_logic"]["rule"]
print(rule.strip())

print("\n--- 身份证真伪记录 ---")
print("适用类型:", data["id_authenticity_check"]["applies_to"])

print("\n--- 蓝卡特殊 ---")
bc = person_types["blue_card"]["blue_card_special"]
print("启用:", bc["enabled"], "| 规则:", bc["rules"][0])

print("\n--- 劳工三表 ---")
lbc = person_types["blue_card"]["labor_docs_check"]
print("一致性要求:", lbc["consistency_required"])
for d in lbc["docs"]:
    print("  -", d)

print("\n--- 个人资料页检查项 ---")
for f in data["personal_info_page_check"]["fields_match_id_copy"]:
    print("  %s: %s" % (f["name"], f["check"]))

print("\n[OK] 结构完整，全部验证通过")
