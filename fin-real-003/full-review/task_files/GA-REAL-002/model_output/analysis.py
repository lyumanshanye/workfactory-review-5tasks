#!/usr/bin/env python3
import argparse
import json
import re
import zipfile
from collections import defaultdict, deque
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TERRAINS = ["flat", "forest", "village", "shallow_water", "deep_water"]
DAMAGE_TYPES = ["blade", "pierce", "impact", "fire", "cold", "arcane"]
UNIT_FIELDS = ["id", "level", "race", "cost", "hitpoints", "movement", "movement_type", "alignment", "experience", "advances_to"]


def rel(path, root):
    return path.relative_to(root).as_posix()


def clean_value(value):
    value = value.strip()
    if value.startswith("_"):
        value = value[1:].strip()
    if len(value) >= 2 and value[0] == value[-1] == '"':
        value = value[1:-1]
    return value.strip()


def int_or_none(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def split_csv(value):
    if not value or value.strip().lower() == "null":
        return []
    return [x.strip() for x in value.split(",") if x.strip()]


def parse_faction(path, root):
    lines = path.read_text(encoding="utf-8").splitlines()
    result = {"source_path": rel(path, root), "source_line": 1}
    for no, line in enumerate(lines, 1):
        m = re.match(r"\s*(id|leader|random_leader|recruit)\s*=\s*(.*?)\s*$", line)
        if m:
            key, value = m.groups()
            result[key] = clean_value(value)
            result[key + "_line"] = no
            result[key + "_source"] = f"{rel(path, root)}:{no}"
            result[key + "_source_record"] = {"path": rel(path, root), "tag": "multiplayer_side", "attribute": key, "line": no}
    for key in ("leader", "random_leader", "recruit"):
        result[key + "_ids"] = split_csv(result.get(key, ""))
    return result


def parse_cfg_units(path, root):
    lines = path.read_text(encoding="utf-8").splitlines()
    units = []
    stack = []
    current = None
    attack = None
    section = None
    variation_depth = 0
    for no, line in enumerate(lines, 1):
        stripped = line.strip()
        if stripped.startswith("#"):
            continue
        close = re.match(r"\[/([^]]+)\]", stripped)
        open_tag = re.match(r"\[([^]/+][^]]*)\]", stripped)
        if open_tag:
            tag = open_tag.group(1).strip()
            stack.append(tag)
            if tag == "unit_type":
                current = {
                    "source_path": rel(path, root), "source_line": no,
                    "fields": {}, "field_lines": {}, "attacks": [],
                    "overrides": {"movement_costs": {}, "defense": {}, "resistance": {}},
                    "override_lines": {"movement_costs": {}, "defense": {}, "resistance": {}},
                    "ability_macros": [],
                }
            elif current and tag == "variation":
                variation_depth += 1
            elif current and variation_depth == 0 and tag == "attack" and stack.count("unit_type") == 1:
                attack = {"source_path": rel(path, root), "source_line": no, "fields": {}, "field_lines": {}, "special_macros": []}
            elif current and variation_depth == 0 and tag in ("movement_costs", "defense", "resistance", "abilities"):
                section = tag
            continue
        if close:
            tag = close.group(1).strip()
            if current and tag == "attack" and attack is not None:
                current["attacks"].append(attack)
                attack = None
            elif current and tag == "variation":
                variation_depth = max(0, variation_depth - 1)
            elif current and tag in ("movement_costs", "defense", "resistance", "abilities"):
                section = None
            elif tag == "unit_type" and current:
                units.append(current)
                current = None
                section = None
                attack = None
            if stack and stack[-1] == tag:
                stack.pop()
            elif tag in stack:
                stack = stack[:len(stack) - 1 - stack[::-1].index(tag)]
            continue
        if not current or variation_depth:
            continue
        macro = re.match(r"\s*\{([A-Za-z0-9_:]+)(?:\s+.*?)?\}\s*$", line)
        if macro:
            name = macro.group(1)
            ref = {"macro": name, "source_path": rel(path, root), "source_line": no, "source": f"{rel(path, root)}:{no}"}
            if attack is not None and "specials" in stack:
                attack["special_macros"].append(ref)
            elif section == "abilities":
                current["ability_macros"].append(ref)
            continue
        attr = re.match(r"\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([^#]+?)\s*$", line)
        if not attr:
            continue
        key, value = attr.group(1), clean_value(attr.group(2))
        if attack is not None and stack and stack[-1] in ("attack", "specials"):
            if stack[-1] == "attack":
                attack["fields"][key] = value
                attack["field_lines"][key] = no
        elif section in ("movement_costs", "defense", "resistance") and stack and stack[-1] == section:
            current["overrides"][section][key] = value
            current["override_lines"][section][key] = no
        elif stack and stack[-1] == "unit_type":
            current["fields"][key] = value
            current["field_lines"][key] = no
    return units


def parse_movetypes(path, root):
    lines = path.read_text(encoding="utf-8").splitlines()
    result = {}
    current = None
    section = None
    for no, line in enumerate(lines, 1):
        stripped = line.strip()
        op = re.match(r"\[([^]/+][^]]*)\]", stripped)
        cl = re.match(r"\[/([^]]+)\]", stripped)
        if op:
            tag = op.group(1)
            if tag == "movetype":
                current = {"source_path": rel(path, root), "source_line": no, "values": defaultdict(dict), "lines": defaultdict(dict)}
            elif current and tag in ("movement_costs", "defense", "resistance"):
                section = tag
            continue
        if cl:
            tag = cl.group(1)
            if current and tag in ("movement_costs", "defense", "resistance"):
                section = None
            elif current and tag == "movetype":
                name = current.get("name")
                if name:
                    current["values"] = dict(current["values"])
                    current["lines"] = dict(current["lines"])
                    result[name] = current
                current = None
            continue
        if not current:
            continue
        attr = re.match(r"\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([^#]+?)\s*$", line)
        if not attr:
            continue
        key, value = attr.group(1), clean_value(attr.group(2))
        if section:
            current["values"][section][key] = value
            current["lines"][section][key] = no
        elif key == "name":
            current["name"] = value
            current["name_line"] = no
    return result


def parse_macros(paths, root):
    definitions = {}
    for path in paths:
        lines = path.read_text(encoding="utf-8").splitlines()
        i = 0
        while i < len(lines):
            m = re.match(r"\s*#define\s+([^\s]+)(?:\s+(.*))?$", lines[i])
            if not m:
                i += 1
                continue
            name = m.group(1)
            start = i + 1
            body = []
            i += 1
            while i < len(lines) and not re.match(r"\s*#enddef", lines[i]):
                body.append(lines[i])
                i += 1
            attrs = {}
            tags = []
            nested = []
            for offset, text in enumerate(body, start + 1):
                tm = re.match(r"\s*\[([^]/+][^]]*)\]", text)
                if tm:
                    tags.append({"tag": tm.group(1), "line": offset})
                am = re.match(r"\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$", text)
                if am and am.group(1) in ("id", "value", "multiply", "active_on", "apply_to", "cumulative", "max_value", "sub", "poison", "affect_self", "affect_allies"):
                    attrs[am.group(1)] = clean_value(am.group(2))
                nm = re.match(r"\s*\{([A-Za-z0-9_:]+)(?:\s+.*?)?\}", text)
                if nm:
                    nested.append(nm.group(1))
            definitions[name] = {
                "macro": name, "source_path": rel(path, root), "source_line": start,
                "source": f"{rel(path, root)}:{start}", "end_line": i + 1,
                "tags": tags, "attributes": attrs, "nested_macros": nested,
            }
            i += 1
    return definitions


def unit_public(unit):
    f = unit["fields"]
    return {
        "id": f.get("id"), "level": int_or_none(f.get("level")), "race": f.get("race"),
        "cost": int_or_none(f.get("cost")), "hp": int_or_none(f.get("hitpoints")),
        "movement": int_or_none(f.get("movement")), "movement_type": f.get("movement_type"),
        "alignment": f.get("alignment"), "experience": int_or_none(f.get("experience")),
        "advances_to": split_csv(f.get("advances_to", "")),
        "source_path": unit["source_path"], "source_line": unit["source_line"],
        "source": f"{unit['source_path']}:{unit['source_line']}", "source_tag": "unit_type",
        "field_sources": {k: {"path": unit["source_path"], "tag": "unit_type", "attribute": k, "line": line} for k, line in unit["field_lines"].items()},
    }


def build_graph(recruits, index):
    reachable = set()
    unresolved = []
    edges = []
    queue = deque(recruits)
    while queue:
        uid = queue.popleft()
        if uid in reachable:
            continue
        if uid not in index:
            unresolved.append({"id": uid, "reason": "招募或升级目标在扫描到的 [unit_type] ID 中不存在"})
            continue
        reachable.add(uid)
        src = index[uid]
        for target in split_csv(src["fields"].get("advances_to", "")):
            edge_source = (f"{src['source_path']}:{src['field_lines']['advances_to']}"
                           if 'advances_to' in src['field_lines'] else f"{src['source_path']}:{src['source_line']}")
            edge = {"from": uid, "to": target, "source": edge_source,
                    "source_record": {"path": src["source_path"], "tag": "unit_type", "attribute": "advances_to", "line": src["field_lines"].get("advances_to", src["source_line"])},
                    "resolved": target in index}
            edges.append(edge)
            if target in index:
                queue.append(target)
            else:
                unresolved.append({"id": target, "from": uid, "reason": "advances_to 目标未解析", "source": edge["source"]})
    adj = defaultdict(list)
    for e in edges:
        if e["resolved"]:
            adj[e["from"]].append(e["to"])
    cycles = []
    def dfs(node, path):
        if node in path:
            cycles.append(path[path.index(node):] + [node])
            return
        for nxt in adj[node]:
            dfs(nxt, path + [node])
    for r in recruits:
        if r in index:
            dfs(r, [])
    return reachable, edges, unresolved, cycles, adj


def attack_public(attack):
    f = attack["fields"]
    return {
        "name": f.get("name"), "range": f.get("range"), "type": f.get("type"),
        "damage": int_or_none(f.get("damage")), "number": int_or_none(f.get("number")),
        "source_path": attack["source_path"], "source_line": attack["source_line"],
        "source": f"{attack['source_path']}:{attack['source_line']}", "source_tag": "attack",
        "field_sources": {k: {"path": attack["source_path"], "tag": "attack", "attribute": k, "line": v} for k, v in attack["field_lines"].items()},
        "special_macros": attack["special_macros"],
    }


def terrain_rows(reachable, index, movetypes):
    rows = []
    unresolved = set()
    cpp_fallback = {
        "movement_costs": "src/movetype.cpp:79-86, src/movetype.hpp:173-176",
        "defense": "src/movetype.cpp:40-50,79-86; src/movetype.hpp:194-199",
        "resistance": "src/movetype.cpp:755-761",
    }
    for uid in sorted(reachable):
        unit = index[uid]
        mt_name = unit["fields"].get("movement_type")
        mt = movetypes.get(mt_name)
        if not mt:
            unresolved.add(mt_name or "<missing>")
        for category, keys in (("movement_costs", TERRAINS), ("defense", TERRAINS), ("resistance", DAMAGE_TYPES)):
            for key in keys:
                raw = None
                source_kind = "缺失"
                source = cpp_fallback[category]
                if mt and key in mt["values"].get(category, {}):
                    raw = mt["values"][category][key]
                    source_kind = "movement_type 公共值"
                    source = f"{mt['source_path']}:{mt['lines'][category][key]}"
                if key in unit["overrides"].get(category, {}):
                    raw = unit["overrides"][category][key]
                    source_kind = "单位覆盖"
                    source = f"{unit['source_path']}:{unit['override_lines'][category][key]}"
                numeric = int_or_none(raw)
                cap = category == "defense" and numeric is not None and numeric < 0
                if numeric is not None:
                    effective = abs(numeric) if category == "defense" else numeric
                    derivation = "单位覆盖公共值" if source_kind == "单位覆盖" else "公共 movement_type"
                else:
                    if category == "movement_costs":
                        effective = 99
                    elif category == "defense":
                        effective = 100
                    else:
                        effective = 100
                    derivation = "WML 未配置；按随包 C++ 默认解释，非 WML 明文值"
                rows.append({
                    "unit_id": uid, "movement_type": mt_name, "category": category, "key": key,
                    "raw_value": raw, "effective_value": effective, "source_kind": source_kind,
                    "source": source, "source_tag": category, "source_attribute": key,
                    "derivation": derivation, "defense_cap": cap,
                    "meaning": ("每次进入地形消耗；99 为引擎不可达哨兵" if category == "movement_costs" else
                                "命中率配置值；闪避率=100-有效值；负号还表示防御上限" if category == "defense" else
                                "承受伤害百分比；100=全额，80=承受80%，并非‘80%减伤’"),
                })
    return rows, sorted(unresolved)


def edge_details(edges, index):
    result = []
    for e in edges:
        row = dict(e)
        if not e["resolved"]:
            result.append(row)
            continue
        a, b = index[e["from"]], index[e["to"]]
        ap, bp = unit_public(a), unit_public(b)
        for key in ("hp", "movement", "cost"):
            av, bv = ap[key], bp[key]
            row[key + "_from"] = av
            row[key + "_to"] = bv
            row[key + "_delta"] = bv - av if av is not None and bv is not None else None
        row["alignment_from"] = ap["alignment"]
        row["alignment_to"] = bp["alignment"]
        row["alignment_changed"] = ap["alignment"] != bp["alignment"]
        row["hp_up_movement_down"] = (row["hp_delta"] or 0) > 0 and (row["movement_delta"] or 0) < 0
        aa = { (x["name"], x["range"], x["type"]): x for x in map(attack_public, a["attacks"]) }
        ba = { (x["name"], x["range"], x["type"]): x for x in map(attack_public, b["attacks"]) }
        changes = []
        for key in sorted(set(aa) | set(ba), key=lambda x: tuple(str(v) for v in x)):
            if key not in aa:
                changes.append({"status": "新增", "key": key, "to": ba[key], "source": ba[key]["source"]})
            elif key not in ba:
                changes.append({"status": "删除", "key": key, "from": aa[key], "source": aa[key]["source"]})
            else:
                x, y = aa[key], ba[key]
                changes.append({
                    "status": "保留", "key": key,
                    "damage_from": x["damage"], "damage_to": y["damage"],
                    "damage_delta": y["damage"] - x["damage"] if x["damage"] is not None and y["damage"] is not None else None,
                    "number_from": x["number"], "number_to": y["number"],
                    "number_delta": y["number"] - x["number"] if x["number"] is not None and y["number"] is not None else None,
                    "specials_from": [m["macro"] for m in x["special_macros"]],
                    "specials_to": [m["macro"] for m in y["special_macros"]],
                    "source": f"{x['source']} -> {y['source']}",
                })
        row["weapon_changes"] = changes
        result.append(row)
    return result


def enumerate_paths(recruits, reachable, adj, index):
    rows = []
    for root in recruits:
        if root not in reachable:
            continue
        paths_by_end = defaultdict(list)
        def walk(node, path):
            if node in path:
                return
            new_path = path + [node]
            if node != root:
                paths_by_end[node].append(new_path)
            for nxt in adj[node]:
                walk(nxt, new_path)
        walk(root, [])
        for endpoint, paths in sorted(paths_by_end.items()):
            values = []
            for p in paths:
                xp_parts = [int_or_none(index[x]["fields"].get("experience")) for x in p[:-1]]
                xp = sum(v for v in xp_parts if v is not None) if all(v is not None for v in xp_parts) else None
                values.append((p, xp))
            valid = [v for _, v in values if v is not None]
            minimum = min(valid) if valid else None
            for p, xp in values:
                rows.append({
                    "recruit": root, "endpoint": endpoint, "path": p, "path_text": " -> ".join(p),
                    "non_endpoint_units": p[:-1], "base_xp_sum": xp, "minimum_for_pair": minimum,
                    "is_minimum": xp == minimum if minimum is not None else False,
                    "sources": [(f"{index[x]['source_path']}:{index[x]['field_lines']['experience']}"
                                 if 'experience' in index[x]['field_lines'] else f"{index[x]['source_path']}:{index[x]['source_line']}")
                                for x in p[:-1]],
                    "caveat": "仅为途中各非终点单位 experience 原始值之和；不含击杀获益、经验倍率、traits 或 AMLA，不代表真实升级耗时。",
                })
    return rows


def capability_rows(reachable, index, macro_defs):
    uses = []
    unresolved = []
    for uid in sorted(reachable):
        unit = index[uid]
        for ref in unit["ability_macros"]:
            uses.append((uid, "unit_ability", None, ref))
        for attack in unit["attacks"]:
            ap = attack_public(attack)
            for ref in attack["special_macros"]:
                uses.append((uid, "weapon_special", f"{ap['name']}|{ap['range']}|{ap['type']}", ref))
    def dependency_chain(name, seen=None):
        seen = set() if seen is None else set(seen)
        if name in seen:
            return [{"macro": name, "cycle": True}]
        definition = macro_defs.get(name)
        if not definition:
            return [{"macro": name, "unresolved": True}]
        item = {
            "macro": name, "source": definition["source"], "tags": definition["tags"],
            "attributes": definition["attributes"], "dependencies": [],
        }
        for nested in definition["nested_macros"]:
            if nested.startswith("INTERNAL:SPECIAL_NOTES"):
                continue
            item["dependencies"].extend(dependency_chain(nested, seen | {name}))
        return [item]
    rows = []
    for uid, kind, weapon, ref in uses:
        definition = macro_defs.get(ref["macro"])
        chain = dependency_chain(ref["macro"])
        missing = []
        def collect_missing(items):
            for item in items:
                if item.get("unresolved"):
                    missing.append(item["macro"])
                collect_missing(item.get("dependencies", []))
        collect_missing(chain)
        if not definition or missing:
            unresolved.append({"unit_id": uid, "macro": ref["macro"], "use_source": ref["source"], "reason": "宏定义或业务嵌套定义未解析", "missing": missing})
        rows.append({
            "unit_id": uid, "kind": kind, "weapon": weapon, "macro": ref["macro"],
            "use_source": ref["source"], "use_tag": "abilities" if kind == "unit_ability" else "attack/specials",
            "definition_source": definition["source"] if definition else None,
            "definition": definition, "dependency_chain": chain,
        })
    return rows, unresolved


def flatten_weapon_changes(edge_rows):
    rows = []
    for e in edge_rows:
        for w in e.get("weapon_changes", []):
            key = w["key"]
            rows.append({"from": e["from"], "to": e["to"], "weapon_name": key[0], "range": key[1], "type": key[2], **{k: v for k, v in w.items() if k != "key"}})
    return rows


def add_sheet(wb, title, headers, rows):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for row in rows:
        ws.append([json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list, tuple)) else v for v in row])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col_idx, col in enumerate(ws.columns, 1):
        width = min(70, max(10, max(len(str(c.value or "")) for c in col) + 2))
        ws.column_dimensions[get_column_letter(col_idx)].width = width
        for c in col:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    return ws


def create_workbook(path, data):
    wb = Workbook()
    fixed = datetime(2000, 1, 1)
    wb.properties.created = fixed
    wb.properties.modified = fixed
    wb.properties.lastPrinted = None
    wb.remove(wb.active)
    f = data["faction"]
    faction_rows = []
    for kind, ids, source in (("recruit", f["recruit_ids"], f["recruit_source"]), ("leader", f["leader_ids"], f["leader_source"]), ("random_leader", f["random_leader_ids"], f["random_leader_source"])):
        for uid in ids:
            faction_rows.append([kind, uid, uid in data["reachable_ids"], uid in data["unit_ids"], source])
    add_sheet(wb, "阵营集合", ["集合", "unit_id", "升级闭包内", "全局单位已解析", "来源"], faction_rows)

    unit_rows = []
    for u in data["units"]:
        weapons = [f"{a['name']}|{a['range']}|{a['type']} {a['damage']}×{a['number']}" for a in u["attacks"]]
        unit_rows.append([u[k] for k in ("id", "level", "race", "cost", "hp", "movement", "movement_type", "alignment", "experience")] + [",".join(u["advances_to"]), "\n".join(weapons), u["source"], u["field_sources"]])
    add_sheet(wb, "单位目录", ["ID", "等级", "种族", "cost", "HP", "movement", "movement_type", "alignment", "experience", "advances_to", "直接武器", "单位来源", "字段来源"], unit_rows)

    attack_rows = []
    for u in data["units"]:
        for a in u["attacks"]:
            attack_rows.append([u["id"], a["name"], a["range"], a["type"], a["damage"], a["number"], [m["macro"] for m in a["special_macros"]], a["source"], a["field_sources"]])
    add_sheet(wb, "直接武器", ["unit_id", "名称", "距离", "类型", "伤害", "次数", "特殊宏", "attack来源", "字段来源"], attack_rows)

    edge_rows = []
    for e in data["edges"]:
        edge_rows.append([e.get(k) for k in ("from", "to", "resolved", "hp_from", "hp_to", "hp_delta", "movement_from", "movement_to", "movement_delta", "cost_from", "cost_to", "cost_delta", "alignment_from", "alignment_to", "alignment_changed", "hp_up_movement_down", "source")])
    add_sheet(wb, "升级边", ["起点", "终点", "已解析", "HP前", "HP后", "HP变化", "移动前", "移动后", "移动变化", "cost前", "cost后", "cost变化", "阵营前", "阵营后", "alignment改变", "HP升且移动降", "边来源"], edge_rows)

    wc_rows = []
    for w in data["weapon_changes"]:
        wc_rows.append([w.get(k) for k in ("from", "to", "status", "weapon_name", "range", "type", "damage_from", "damage_to", "damage_delta", "number_from", "number_to", "number_delta", "specials_from", "specials_to", "source")])
    add_sheet(wb, "武器变化", ["起点", "终点", "状态", "名称", "距离", "类型", "伤害前", "伤害后", "伤害变化", "次数前", "次数后", "次数变化", "特殊前", "特殊后", "来源"], wc_rows)

    path_rows = [[p["recruit"], p["endpoint"], p["path_text"], p["base_xp_sum"], p["minimum_for_pair"], p["is_minimum"], p["sources"], p["caveat"]] for p in data["paths"]]
    add_sheet(wb, "路径XP", ["招募起点", "终点", "简单路径", "非终点XP和", "该对最小值", "最小路径", "XP来源", "口径限制"], path_rows)

    tr_rows = [[r[k] for k in ("unit_id", "movement_type", "category", "key", "raw_value", "effective_value", "source_kind", "source_tag", "source_attribute", "defense_cap", "derivation", "meaning", "source")] for r in data["terrain_resistance"]]
    add_sheet(wb, "地形抗性", ["unit_id", "movement_type", "类别", "地形/伤害类型", "WML原始值", "静态有效值", "来源层", "来源标签", "来源属性", "负值上限", "推导", "解释", "来源"], tr_rows)

    cap_rows = []
    for c in data["capabilities"]:
        d = c.get("definition") or {}
        cap_rows.append([c["unit_id"], c["kind"], c["weapon"], c["macro"], c["use_tag"], c["use_source"], c["definition_source"], d.get("tags"), d.get("attributes"), d.get("nested_macros"), c.get("dependency_chain")])
    add_sheet(wb, "能力依赖", ["unit_id", "类别", "武器键", "宏", "使用标签", "使用来源", "定义来源", "定义标签", "关键属性", "嵌套宏", "完整定义链"], cap_rows)

    cov = data["coverage"]
    coverage_rows = [
        ["招募ID", cov["recruits_total"], cov["recruits_resolved"], cov["unresolved_recruits"], f["recruit_source"]],
        ["升级目标", cov["edge_targets_total"], cov["edge_targets_resolved"], cov["unresolved_edges"], "各单位 advances_to 字段"],
        ["movement_type", cov["movement_types_total"], cov["movement_types_resolved"], cov["unresolved_movement_types"], "data/core/units.cfg"],
        ["业务相关能力宏", cov["macros_total"], cov["macros_resolved"], cov["unresolved_macros"], "能力使用与宏定义"],
        ["升级环", len(cov["cycles"]), 0 if cov["cycles"] else len(cov["cycles"]), cov["cycles"], "递归简单路径检测"],
        ["多分支节点", len(cov["branch_nodes"]), len(cov["branch_nodes"]), cov["branch_nodes"], "升级图"],
        ["交汇节点", len(cov["merge_nodes"]), len(cov["merge_nodes"]), cov["merge_nodes"], "升级图"],
    ]
    add_sheet(wb, "完整性与覆盖", ["检查项", "总数", "已解析/结果数", "详情", "依据"], coverage_rows)
    wb.save(path)


def add_doc_table(doc, headers, rows):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = str(h)
    for row in rows:
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = "" if value is None else str(value)
    return table


def create_memo(path, data):
    doc = Document()
    styles = doc.styles
    styles["Normal"].font.name = "Microsoft YaHei"
    styles["Normal"].font.size = Pt(9)
    doc.add_heading("Wesnoth 1.18.0 Loyalists 招募与升级静态审计备忘录", 0)
    doc.add_paragraph("范围：官方 1.18.0 快照中的 Loyalists 默认多人阵营；仅做配置依赖审计，不推断完整战斗胜率，也不声称运行了游戏引擎。")

    f = data["faction"]
    doc.add_heading("1. 集合与升级闭包", level=1)
    doc.add_paragraph(f"真实 recruit 共 {len(f['recruit_ids'])} 个：{', '.join(f['recruit_ids'])}（{f['recruit_source']}）。")
    doc.add_paragraph(f"leader 共 {len(f['leader_ids'])} 个：{', '.join(f['leader_ids'])}（{f['leader_source']}）。leader 列表不是 recruit 列表；例如 Shock Trooper 可作领袖但不是基础招募。")
    doc.add_paragraph(f"从 recruit 递归 advances_to 得到 {len(data['reachable_ids'])} 个可达单位、{len(data['edges'])} 条升级边。单位来自 human-loyalists 之外的法师、骑士和 merfolk 路径也已纳入。")
    cov = data["coverage"]
    doc.add_paragraph(f"多分支：{', '.join(cov['branch_nodes']) or '无'}；交汇：{', '.join(cov['merge_nodes']) or '无'}；循环：{json.dumps(cov['cycles'], ensure_ascii=False) if cov['cycles'] else '无'}。未解析项见工作簿“完整性与覆盖”。")

    doc.add_heading("2. 口径与容易误读的引擎语义", level=1)
    for text in [
        "单位属性和武器只读取 [unit_type] 的直接属性与直接 [attack]；不把 [attack_anim]/[filter_attack]、文本引用或 [variation] 当基础武器。",
        "地形表按引擎加载顺序先继承 movement_type，再合并单位自身配置（src/units/types.cpp:323-335）。工作簿同时保留 WML 原始值、最终来源和静态有效值。",
        "[defense] 的配置值是受击命中率方向：有效值 60 对应 40% 闪避。负值先取绝对值，同时建立防御上限；mounted 的 forest=-70 不能只抄成普通 70（src/movetype.cpp:40-50,83-86；src/movetype.hpp:194-199）。",
        "[resistance] 数字是承受伤害比例，不是减伤率：80 表示承受 80%（约 20% 减伤），120 表示承受 120%（弱点）。战斗计算乘以 damage_from（src/actions/attack.cpp:182-203）。",
        "WML 缺失不写成 0。静态有效栏若显示 movement=99、defense=100 或 resistance=100，明确标为随包 C++ 默认解释：src/movetype.cpp:54-86,268-334,755-761；src/movetype.hpp:173-176。未运行引擎。",
        "路径 XP 是每条简单升级路径中所有非终点单位 experience 原始值之和，仅是基础配置门槛；不含击杀获益、经验倍率、traits 或 AMLA，不能称为真实升级耗时。",
    ]:
        doc.add_paragraph(text, style="List Bullet")

    doc.add_heading("3. 升级边中的隐性代价", level=1)
    special_edges = [e for e in data["edges"] if e.get("hp_up_movement_down") or e.get("alignment_changed")]
    rows = []
    for e in special_edges:
        flags = []
        if e.get("hp_up_movement_down"):
            flags.append("HP上升但movement下降")
        if e.get("alignment_changed"):
            flags.append("alignment改变")
        rows.append([f"{e['from']} → {e['to']}", f"{e.get('hp_from')}→{e.get('hp_to')} ({e.get('hp_delta'):+d})" if e.get('hp_delta') is not None else "", f"{e.get('movement_from')}→{e.get('movement_to')} ({e.get('movement_delta'):+d})" if e.get('movement_delta') is not None else "", f"{e.get('alignment_from')}→{e.get('alignment_to')}", "；".join(flags), e["source"]])
    add_doc_table(doc, ["升级边", "HP", "移动", "alignment", "标记", "来源"], rows or [["无", "", "", "", "", ""]])

    doc.add_heading("4. 为什么 damage×number÷cost 会失真", level=1)
    examples = data["distortion_examples"]
    for ex in examples:
        doc.add_heading(ex["title"], level=2)
        doc.add_paragraph(ex["text"])
        doc.add_paragraph("配置依据：" + "；".join(ex["sources"]))

    doc.add_heading("5. 建议的三个游戏内验证问题", level=1)
    for i, q in enumerate(data["test_questions"], 1):
        doc.add_heading(f"{i}. {q['question']}", level=2)
        doc.add_paragraph("分支与现有证据：" + q["evidence"])
        doc.add_paragraph("需实测条件：" + q["conditions"])
        doc.add_paragraph("静态来源：" + "；".join(q["sources"]))
    doc.add_paragraph("以上是待验证问题，不是现成测试日志；不预设敌军、权重或胜率。")

    doc.add_heading("6. 完整性结论", level=1)
    doc.add_paragraph(f"招募 ID {cov['recruits_resolved']}/{cov['recruits_total']} 已解析；升级目标 {cov['edge_targets_resolved']}/{cov['edge_targets_total']} 已解析；movement_type {cov['movement_types_resolved']}/{cov['movement_types_total']} 已解析；业务相关能力宏 {cov['macros_resolved']}/{cov['macros_total']} 已解析。详细来源、逐值覆盖、武器变化和路径均在 audit.xlsx 与 results.json。")
    doc.save(path)


def pick_unit(data, uid):
    return next(u for u in data["units"] if u["id"] == uid)


def build_narrative(data):
    macro_src = {c["macro"]: c["definition_source"] for c in data["capabilities"] if c["definition_source"]}
    def u(uid): return pick_unit(data, uid)
    def us(uid): return u(uid)["source"]
    data["distortion_examples"] = [
        {
            "title": "Charge：面板输出上升同时放大反击风险",
            "text": "Horseman → Lancer 的 spear 仍带 charge。该宏只在进攻时生效，将双方伤害乘 2；因此只按攻击者面板 damage×number÷cost 排序，会漏掉触发条件与承受双倍反击的代价。Knight/Paladin/Grand Knight 路线也保留 charge，但具体面板、HP、cost 与其他武器变化不同。",
            "sources": [us("Horseman"), us("Lancer"), macro_src.get("WEAPON_SPECIAL_CHARGE", "未解析")],
        },
        {
            "title": "Magical：命中率依赖无法由伤害乘次数表达",
            "text": "Mage 的 missile 及法师升级线相关法术带 magical，定义为命中率固定 70%，不取决于目标地形防御。Mage → Red Mage 还把 alignment 从 lawful 改为 neutral；静态面板比值既漏掉命中规则，也漏掉时段修正差异。",
            "sources": [us("Mage"), us("Red Mage"), macro_src.get("WEAPON_SPECIAL_MAGICAL", "未解析")],
        },
        {
            "title": "First strike：改变出手顺序，而非无条件增加面板伤害",
            "text": "Spearman 的近战 spear 带 first strike，Pikeman/Halberdier/Javelineer 分支也有对应特殊。它让该武器即使防守也先出手，价值取决于是否发生对应近战与双方存活状态，不能作为无条件数值奖励加到 damage×number。",
            "sources": [us("Spearman"), us("Pikeman"), us("Javelineer"), macro_src.get("WEAPON_SPECIAL_FIRSTSTRIKE", "未解析")],
        },
        {
            "title": "单位能力与机动：成本分母看不到治疗、照明、传送和防御态抗性",
            "text": "White Mage 有 cures，Mage of Light 另有 illuminates；Silver Mage 有 teleport；Merman Hoplite 有仅防守时生效的 steadfast（抗性翻倍且最高 50%，弱点不受影响）。这些依赖位置、回合、时段或攻守状态，均不在单武器面板比值中。",
            "sources": [us("White Mage"), us("Mage of Light"), us("Silver Mage"), us("Merman Hoplite"), macro_src.get("ABILITY_STEADFAST", "未解析")],
        },
    ]
    edge_map = {(e["from"], e["to"]): e for e in data["edges"]}
    def edge_ev(a, b):
        e = edge_map[(a, b)]
        return f"{a}→{b}: HP {e['hp_from']}→{e['hp_to']}，movement {e['movement_from']}→{e['movement_to']}，cost {e['cost_from']}→{e['cost_to']}"
    data["test_questions"] = [
        {
            "question": "Horseman 的 Knight 与 Lancer 分支，在什么进攻/防守节奏下 charge 的收益能覆盖双倍反击风险？",
            "evidence": edge_ev("Horseman", "Knight") + "；" + edge_ev("Horseman", "Lancer") + "。两线武器变化详见工作簿。",
            "conditions": "分别记录主动 charge、被动接战、是否有合法反击、地形命中率与昼夜；保持地图、资金和对局条件可复现，不预设敌军权重。",
            "sources": [us("Horseman"), us("Knight"), us("Lancer"), macro_src.get("WEAPON_SPECIAL_CHARGE", "未解析")],
        },
        {
            "question": "Mage 的 White Mage 与 Red Mage 分支，治疗/照明潜力和 neutral 输出在昼夜与高低防御地形中如何交换？",
            "evidence": edge_ev("Mage", "White Mage") + "；" + edge_ev("Mage", "Red Mage") + "；Mage→Red Mage 存在 alignment 改变，法术带 magical。",
            "conditions": "分别覆盖白天、黄昏、夜间，目标处于不同防御地形，并记录 cures/后续 illuminates 的实际触发机会与 magical 的固定命中；不把一次场景结果外推成胜率。",
            "sources": [us("Mage"), us("White Mage"), us("Red Mage"), us("Mage of Light"), macro_src.get("WEAPON_SPECIAL_MAGICAL", "未解析")],
        },
        {
            "question": "Merman Warrior 的 Triton 与 Hoplite 分支，在水域转陆地和防守态下，机动损失是否由抗性与 steadfast 补偿？",
            "evidence": edge_ev("Merman Warrior", "Merman Triton") + "；" + edge_ev("Merman Warrior", "Merman Hoplite") + "。Hoplite 有单位抗性覆盖及仅防守生效的 steadfast。",
            "conditions": "覆盖 deep/shallow water、village、flat、forest 的路径与驻防，区分主动进攻和被攻击，记录移动可达性、受击类型和 steadfast 是否满足防守条件。",
            "sources": [us("Merman Warrior"), us("Merman Triton"), us("Merman Hoplite"), macro_src.get("ABILITY_STEADFAST", "未解析")],
        },
    ]


def build(input_root):
    faction_path = input_root / "data/multiplayer/factions/loyalists-default.cfg"
    units_root = input_root / "data/core/units"
    units_cfg = input_root / "data/core/units.cfg"
    faction = parse_faction(faction_path, input_root)
    raw_units = []
    for path in sorted(units_root.rglob("*.cfg")):
        raw_units.extend(parse_cfg_units(path, input_root))
    index = {}
    duplicates = defaultdict(list)
    for raw in raw_units:
        uid = raw["fields"].get("id")
        if uid:
            duplicates[uid].append(raw["source_path"])
            index[uid] = raw
    reachable, graph_edges, unresolved_graph, cycles, adj = build_graph(faction["recruit_ids"], index)
    movetypes = parse_movetypes(units_cfg, input_root)
    macro_paths = [input_root / "data/core/macros/weapon_specials.cfg", input_root / "data/core/macros/abilities.cfg"]
    macro_defs = parse_macros(macro_paths, input_root)
    terrains, unresolved_mts = terrain_rows(reachable, index, movetypes)
    edges = edge_details(graph_edges, index)
    paths = enumerate_paths(faction["recruit_ids"], reachable, adj, index)
    capabilities, unresolved_macros = capability_rows(reachable, index, macro_defs)
    units = []
    for uid in sorted(reachable):
        p = unit_public(index[uid])
        p["attacks"] = [attack_public(a) for a in index[uid]["attacks"]]
        p["ability_macros"] = index[uid]["ability_macros"]
        p["overrides"] = index[uid]["overrides"]
        units.append(p)
    indegree = defaultdict(int)
    for e in edges:
        if e["resolved"]:
            indegree[e["to"]] += 1
    used_mts = sorted({index[u]["fields"].get("movement_type") for u in reachable})
    recruit_unresolved = [u for u in faction["recruit_ids"] if u not in index]
    unresolved_edges = [x for x in unresolved_graph if x.get("from")]
    coverage = {
        "recruits_total": len(faction["recruit_ids"]), "recruits_resolved": len(faction["recruit_ids"]) - len(recruit_unresolved), "unresolved_recruits": recruit_unresolved,
        "edge_targets_total": len(edges), "edge_targets_resolved": sum(1 for e in edges if e["resolved"]), "unresolved_edges": unresolved_edges,
        "movement_types_total": len(used_mts), "movement_types_resolved": len(used_mts) - len(unresolved_mts), "unresolved_movement_types": unresolved_mts,
        "macros_total": len(capabilities), "macros_resolved": sum(1 for c in capabilities if c["definition"]), "unresolved_macros": unresolved_macros,
        "cycles": cycles, "branch_nodes": sorted([u for u in reachable if len(adj[u]) > 1]), "merge_nodes": sorted([u for u, n in indegree.items() if n > 1]),
        "duplicate_unit_ids_in_snapshot": {k: v for k, v in duplicates.items() if len(v) > 1},
    }
    data = {
        "metadata": {
            "scope": "Battle for Wesnoth 1.18.0 Loyalists default faction static configuration audit",
            "input_root": str(input_root), "engine_executed": False,
            "parser_note": "analysis.py performs a line-aware constrained WML scan; official wmlparser3.py requires a wesnoth executable for preprocessing (data/tools/wesnoth/wmlparser3.py:3-7), so macro definitions are traced explicitly without executing unrelated scripts.",
            "source_entry": f"{faction['source_path']}:1",
        },
        "faction": faction, "unit_ids": sorted(index), "reachable_ids": sorted(reachable),
        "units": units, "edges": edges, "weapon_changes": flatten_weapon_changes(edges),
        "paths": paths, "terrain_resistance": terrains, "capabilities": capabilities,
        "coverage": coverage,
        "semantics_sources": {
            "movement_type_inheritance": ["src/units/types.cpp:323-335"],
            "defense_direction": ["src/actions/attack.cpp:167-180", "src/movetype.cpp:40-50", "src/movetype.hpp:194-199"],
            "missing_defaults": ["src/movetype.cpp:54-86", "src/movetype.cpp:268-334", "src/movetype.cpp:755-761", "src/movetype.hpp:173-176"],
            "resistance_damage_multiplier": ["src/actions/attack.cpp:182-203", "src/movetype.cpp:755-761"],
        },
    }
    build_narrative(data)
    return data


def validate(data, out):
    assert data["coverage"]["recruits_resolved"] == data["coverage"]["recruits_total"]
    assert data["coverage"]["edge_targets_resolved"] == data["coverage"]["edge_targets_total"]
    assert data["coverage"]["movement_types_resolved"] == data["coverage"]["movement_types_total"]
    assert data["coverage"]["macros_resolved"] == data["coverage"]["macros_total"]
    assert not data["coverage"]["cycles"]
    assert len(data["distortion_examples"]) >= 3
    for name in ("audit.xlsx", "memo.docx", "results.json", "analysis.py"):
        assert (out / name).exists(), name
    wb = load_workbook(out / "audit.xlsx", read_only=True)
    required = {"阵营集合", "单位目录", "直接武器", "升级边", "武器变化", "路径XP", "地形抗性", "能力依赖", "完整性与覆盖"}
    assert required.issubset(set(wb.sheetnames))
    wb.close()
    Document(out / "memo.docx")
    parsed = json.loads((out / "results.json").read_text(encoding="utf-8"))
    assert parsed["reachable_ids"] == data["reachable_ids"]


def normalize_ooxml(path):
    fixed_time = (2000, 1, 1, 0, 0, 0)
    temp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as src, zipfile.ZipFile(temp, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as dst:
        for name in sorted(src.namelist()):
            old = src.getinfo(name)
            payload = src.read(name)
            if name == "docProps/core.xml":
                payload = re.sub(
                    rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms:(?:created|modified)>)",
                    rb"\g<1>2000-01-01T00:00:00Z\g<2>", payload,
                )
            info = zipfile.ZipInfo(name, fixed_time)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = old.external_attr
            info.create_system = old.create_system
            dst.writestr(info, payload)
    temp.replace(path)


def main():
    parser = argparse.ArgumentParser(description="Rebuild the Wesnoth 1.18.0 Loyalists static audit")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    data = build(args.input.resolve())
    create_workbook(args.output / "audit.xlsx", data)
    create_memo(args.output / "memo.docx", data)
    normalize_ooxml(args.output / "audit.xlsx")
    normalize_ooxml(args.output / "memo.docx")
    (args.output / "results.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    validate(data, args.output)
    print(json.dumps({
        "reachable_units": len(data["reachable_ids"]), "edges": len(data["edges"]),
        "paths": len(data["paths"]), "capability_uses": len(data["capabilities"]),
        "coverage": data["coverage"],
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
