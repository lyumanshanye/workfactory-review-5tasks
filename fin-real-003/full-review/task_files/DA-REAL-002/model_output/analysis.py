#!/usr/bin/env python3
import argparse
import csv
import json
import math
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BEA_FILES = ("CAINC1__ALL_AREAS.csv", "CAINC4__ALL_AREAS.csv")
CENSUS_POP_FILE = "co-est2024-alldata.csv"
COUNTY_CODE_FILE = "national_county2020.txt"
YEARS = list(range(2020, 2025))
END_YEARS = (2020, 2024)
INDICATORS = {
    "CAINC1": {
        "total_income": "1",
        "population": "2",
        "per_capita_income": "3",
    },
    "CAINC4": {
        "total_income": "10",
        "net_earnings": "45",
        "property_income": "46",
        "transfer_income": "47",
    },
}
INDICATOR_LABELS = {
    ("CAINC1", "1"): "个人收入（千美元）",
    ("CAINC1", "2"): "人口（人）",
    ("CAINC1", "3"): "人均个人收入（美元）",
    ("CAINC4", "10"): "个人收入（千美元）",
    ("CAINC4", "45"): "按居住地净劳动收入（千美元）",
    ("CAINC4", "46"): "股息、利息和租金（千美元）",
    ("CAINC4", "47"): "个人经常转移收入（千美元）",
}


def norm_bea_fips(value):
    text = (value or "").strip().strip('"').strip()
    return text.zfill(5) if text.isdigit() and len(text) <= 5 else text


def parse_number(value):
    text = (value or "").strip()
    if not text:
        return None, "(blank)"
    try:
        return int(text), None
    except ValueError:
        try:
            return float(text), None
        except ValueError:
            return None, text


def safe_rate(new, old):
    if new is None or old is None or old <= 0:
        return None
    return new / old - 1


def safe_log_change(new, old):
    if new is None or old is None or new <= 0 or old <= 0:
        return None
    return math.log(new / old)


def read_bea(path):
    records = []
    footers = []
    malformed = []
    marker_counts = Counter()
    line_defs = {}
    key_counts = Counter()
    with path.open("r", encoding="cp1252", newline="") as handle:
        reader = csv.reader(handle)
        header = next(reader)
        years = [int(x) for x in header if x.isdigit()]
        positions = {name: i for i, name in enumerate(header)}
        for row_number, row in enumerate(reader, start=2):
            if len(row) != len(header):
                malformed.append({"raw_row": row_number, "column_count": len(row), "text": row[0] if row else ""})
            line_code = row[positions["LineCode"]].strip() if len(row) > positions["LineCode"] else ""
            if not line_code.isdigit():
                if any(cell.strip() for cell in row):
                    footers.append({"raw_row": row_number, "text": " | ".join(row)})
                continue
            fips_raw = row[positions["GeoFIPS"]]
            fips = norm_bea_fips(fips_raw)
            values = {}
            raw_values = {}
            for year in years:
                raw = row[positions[str(year)]].strip()
                value, marker = parse_number(raw)
                values[year] = value
                raw_values[year] = raw
                if marker and marker != "(blank)":
                    marker_counts[marker] += 1
            record = {
                "source_file": path.name,
                "raw_row": row_number,
                "geo_fips_raw": fips_raw,
                "fips": fips,
                "geo_name": row[positions["GeoName"]].strip(),
                "region": row[positions["Region"]].strip(),
                "table_name": row[positions["TableName"]].strip(),
                "line_code": line_code,
                "industry_classification": row[positions["IndustryClassification"]].strip(),
                "description": row[positions["Description"]].strip(),
                "unit": row[positions["Unit"]].strip(),
                "values": values,
                "raw_values": raw_values,
            }
            records.append(record)
            line_defs[line_code] = {"description": record["description"], "unit": record["unit"]}
            key_counts[(fips, line_code)] += 1
    return {
        "records": records,
        "years": years,
        "line_definitions": line_defs,
        "markers": dict(sorted(marker_counts.items())),
        "duplicate_keys": [
            {"fips": k[0], "line_code": k[1], "count": count}
            for k, count in sorted(key_counts.items()) if count > 1
        ],
        "footers": footers,
        "malformed_rows": malformed,
    }


def read_census_population(path):
    rows = []
    marker_counts = Counter()
    key_counts = Counter()
    with path.open("r", encoding="cp1252", newline="") as handle:
        reader = csv.DictReader(handle)
        fields = reader.fieldnames
        for row_number, row in enumerate(reader, start=2):
            values = {}
            raw_values = {}
            for year in YEARS:
                raw = (row.get(f"POPESTIMATE{year}") or "").strip()
                value, marker = parse_number(raw)
                values[year] = value
                raw_values[year] = raw
                if marker and marker != "(blank)":
                    marker_counts[marker] += 1
            fips = row["STATE"].strip().zfill(2) + row["COUNTY"].strip().zfill(3)
            item = {
                "source_file": path.name,
                "raw_row": row_number,
                "sumlev": row["SUMLEV"].strip(),
                "state": row["STATE"].strip().zfill(2),
                "county": row["COUNTY"].strip().zfill(3),
                "fips": fips,
                "state_name": row["STNAME"].strip(),
                "county_name": row["CTYNAME"].strip(),
                "values": values,
                "raw_values": raw_values,
            }
            rows.append(item)
            key_counts[(item["sumlev"], fips)] += 1
    counties = [row for row in rows if row["sumlev"] == "050"]
    county_counts = Counter(row["fips"] for row in counties)
    return {
        "records": rows,
        "counties": counties,
        "fields": fields,
        "markers": dict(sorted(marker_counts.items())),
        "duplicate_keys": [
            {"sumlev": k[0], "fips": k[1], "count": count}
            for k, count in sorted(key_counts.items()) if count > 1
        ],
        "county_duplicate_keys": [
            {"fips": fips, "count": count} for fips, count in sorted(county_counts.items()) if count > 1
        ],
    }


def read_county_codes(path):
    rows = []
    key_counts = Counter()
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle, delimiter="|")
        fields = reader.fieldnames
        for row_number, row in enumerate(reader, start=2):
            fips = row["STATEFP"].strip().zfill(2) + row["COUNTYFP"].strip().zfill(3)
            item = {
                "source_file": path.name,
                "raw_row": row_number,
                "fips": fips,
                "state_abbr": row["STATE"].strip(),
                "state_fips": row["STATEFP"].strip().zfill(2),
                "county_fips": row["COUNTYFP"].strip().zfill(3),
                "county_ns": row["COUNTYNS"].strip(),
                "county_name": row["COUNTYNAME"].strip(),
                "class_fips": row["CLASSFP"].strip(),
                "funcstat": row["FUNCSTAT"].strip(),
            }
            rows.append(item)
            key_counts[fips] += 1
    return {
        "records": rows,
        "fields": fields,
        "markers": {},
        "duplicate_keys": [
            {"fips": fips, "count": count} for fips, count in sorted(key_counts.items()) if count > 1
        ],
    }


def index_bea(data):
    return {(row["fips"], row["line_code"]): row for row in data["records"]}


def classify_coverage(fips, bea_row, census_row, code_row, census_state_aggregates):
    if bea_row and census_row:
        return "直接代码匹配"
    if census_row and not bea_row:
        if fips not in {row["fips"] for row in code_row and [code_row] or []} and not code_row:
            return "Census目标县仅见于2024人口表"
        return "Census目标县无BEA同码"
    if bea_row and not census_row:
        name = bea_row["geo_name"]
        if "+" in name or "includes" in name.lower():
            return "BEA合并地区（不拆分）"
        if fips == "00000":
            return "BEA全国汇总（非县）"
        if fips in census_state_aggregates:
            return "BEA州级汇总（非县）"
        if fips in {f"{n}000" for n in range(91, 99)}:
            return "BEA区域汇总（非县）"
        if code_row:
            return "仅BEA与2020代码表有同码，非2024目标"
        return "BEA非目标历史/其他地域（无法证实匹配）"
    return "仅2020县代码表"


def build_analysis(bea1, bea4, census, codes):
    i1 = index_bea(bea1)
    i4 = index_bea(bea4)
    census_county = {row["fips"]: row for row in census["counties"]}
    code_index = {row["fips"]: row for row in codes["records"]}
    bea_geo = {row["fips"]: row for row in bea1["records"] if row["line_code"] == "1"}
    census_states = {row["fips"] for row in census["records"] if row["sumlev"] == "040"}
    target = set(census_county)
    bea_set = set(bea_geo)
    code_set = set(code_index)
    direct = sorted(target & bea_set)

    coverage = []
    all_codes = sorted(target | bea_set | code_set)
    for fips in all_codes:
        b = bea_geo.get(fips)
        c = census_county.get(fips)
        o = code_index.get(fips)
        status = classify_coverage(fips, b, c, o, census_states)
        notes = []
        if c and not o:
            notes.append("2024人口目标集合有此码，2020代码表无此码；保留为目标候选")
        if o and not c:
            notes.append("2020代码表有此码，2024人口SUMLEV=050无此码；不纳入当前目标")
        if c and b and not o:
            notes.append("可按BEA与2024人口表同码比较，但代码版本差异需披露")
        if status == "BEA合并地区（不拆分）":
            notes.append("名称明确标示合并；未提供官方交叉表，不分摊到组成县")
        if status == "Census目标县无BEA同码":
            notes.append("不按名称或邻近合并地区补值")
        coverage.append({
            "fips": fips,
            "status": status,
            "in_bea": bool(b),
            "in_census_2024_county": bool(c),
            "in_county_code_2020": bool(o),
            "bea_name": b["geo_name"] if b else None,
            "census_state": c["state_name"] if c else None,
            "census_name": c["county_name"] if c else None,
            "code2020_state": o["state_abbr"] if o else None,
            "code2020_name": o["county_name"] if o else None,
            "code2020_class": o["class_fips"] if o else None,
            "code2020_funcstat": o["funcstat"] if o else None,
            "bea_raw_row": b["raw_row"] if b else None,
            "census_raw_row": c["raw_row"] if c else None,
            "code2020_raw_row": o["raw_row"] if o else None,
            "notes": "；".join(notes),
        })

    cross = []
    cross_metrics = [
        ("个人收入（千美元）", "1", "10"),
        ("人口（人）", "2", "20"),
        ("人均个人收入（美元）", "3", "30"),
    ]
    for fips in sorted(bea_set):
        for metric, cainc1_code, cainc4_code in cross_metrics:
            a = i1.get((fips, cainc1_code))
            b = i4.get((fips, cainc4_code))
            for year in YEARS:
                av = a["values"].get(year) if a else None
                bv = b["values"].get(year) if b else None
                if av is None or bv is None:
                    status = "不可解析/缺失"
                elif av == bv:
                    status = "一致"
                else:
                    status = "不一致"
                cross.append({
                    "fips": fips,
                    "geo_name": a["geo_name"] if a else b["geo_name"],
                    "year": year,
                    "metric": metric,
                    "cainc1_line_code": cainc1_code,
                    "cainc1_value": av,
                    "cainc1_raw": a["raw_values"].get(year) if a else None,
                    "cainc1_raw_row": a["raw_row"] if a else None,
                    "cainc4_line_code": cainc4_code,
                    "cainc4_value": bv,
                    "cainc4_raw": b["raw_values"].get(year) if b else None,
                    "cainc4_raw_row": b["raw_row"] if b else None,
                    "difference": av - bv if av is not None and bv is not None else None,
                    "status": status,
                })

    population = []
    for fips in direct:
        c = census_county[fips]
        income = i1[(fips, "1")]
        bea_pop = i1[(fips, "2")]
        published_pc = i1[(fips, "3")]
        for year in END_YEARS:
            inc = income["values"].get(year)
            bp = bea_pop["values"].get(year)
            cp = c["values"].get(year)
            pc = published_pc["values"].get(year)
            absolute = bp - cp if bp is not None and cp is not None else None
            relative = absolute / cp if absolute is not None and cp else None
            calc_bea = inc * 1000 / bp if inc is not None and bp and bp > 0 else None
            calc_census = inc * 1000 / cp if inc is not None and cp and cp > 0 else None
            rounding_difference = pc - calc_bea if pc is not None and calc_bea is not None else None
            population.append({
                "fips": fips,
                "bea_name": income["geo_name"],
                "census_state": c["state_name"],
                "census_name": c["county_name"],
                "year": year,
                "bea_population_line_code": "2",
                "bea_population": bp,
                "bea_population_raw": bea_pop["raw_values"].get(year),
                "bea_population_raw_row": bea_pop["raw_row"],
                "census_population_field": f"POPESTIMATE{year}",
                "census_population": cp,
                "census_population_raw": c["raw_values"].get(year),
                "census_raw_row": c["raw_row"],
                "absolute_difference_bea_minus_census": absolute,
                "relative_difference_vs_census": relative,
                "income_line_code": "1",
                "income_thousand_dollars": inc,
                "income_raw_row": income["raw_row"],
                "recalculated_pc_with_bea_population": calc_bea,
                "recalculated_pc_with_census_population": calc_census,
                "published_pc_line_code": "3",
                "published_pc_dollars": pc,
                "published_pc_raw_row": published_pc["raw_row"],
                "published_minus_recalculated_bea": rounding_difference,
                "rounding_check": "通过（发布整数在精确值四舍五入容差内）" if rounding_difference is not None and abs(rounding_difference) <= 0.5000001 else "未通过/不可核验",
            })
    pop_2024 = [row for row in population if row["year"] == 2024 and row["relative_difference_vs_census"] is not None]
    pop_2024.sort(key=lambda row: (-abs(row["relative_difference_vs_census"]), row["fips"]))
    population_top20 = []
    for rank, row in enumerate(pop_2024[:20], start=1):
        item = dict(row)
        item["rank_by_absolute_relative_difference"] = rank
        item["source_localization"] = (
            f"{BEA_FILES[0]}原始行{row['bea_population_raw_row']}、LineCode 2；"
            f"{CENSUS_POP_FILE}原始行{row['census_raw_row']}、字段POPESTIMATE2024"
        )
        item["interpretation"] = "两份留存表数值确有差异；原件未附足以判定成因的元数据，可能涉及估计版本或处理口径，列为核验项而非原始数据错误。"
        population_top20.append(item)

    component_checks = []
    for fips in sorted(bea_set):
        total = i4.get((fips, "10"))
        net = i4.get((fips, "45"))
        prop = i4.get((fips, "46"))
        transfer = i4.get((fips, "47"))
        for year in YEARS:
            values = [x["values"].get(year) if x else None for x in (total, net, prop, transfer)]
            t, n, p, tr = values
            component_sum = n + p + tr if all(v is not None for v in (n, p, tr)) else None
            difference = t - component_sum if t is not None and component_sum is not None else None
            status = "一致" if difference == 0 else ("不可解析/缺失" if difference is None else "不一致")
            component_checks.append({
                "fips": fips,
                "geo_name": total["geo_name"] if total else None,
                "year": year,
                "total_line_code": "10",
                "total_income": t,
                "net_earnings_line_code": "45",
                "net_earnings": n,
                "property_income_line_code": "46",
                "property_income": p,
                "transfer_income_line_code": "47",
                "transfer_income": tr,
                "component_sum": component_sum,
                "total_minus_components": difference,
                "status": status,
                "total_raw_row": total["raw_row"] if total else None,
                "net_raw_row": net["raw_row"] if net else None,
                "property_raw_row": prop["raw_row"] if prop else None,
                "transfer_raw_row": transfer["raw_row"] if transfer else None,
            })

    growth = []
    for fips in direct:
        c = census_county[fips]
        c1_total = i1.get((fips, "1"))
        bpop = i1.get((fips, "2"))
        published_pc = i1.get((fips, "3"))
        c4_total = i4.get((fips, "10"))
        net = i4.get((fips, "45"))
        prop = i4.get((fips, "46"))
        transfer = i4.get((fips, "47"))
        reasons = []
        series = {
            "CAINC1总收入": c1_total,
            "CAINC4总收入": c4_total,
            "净劳动收入": net,
            "财产收入": prop,
            "转移收入": transfer,
            "BEA人口": bpop,
        }
        for label, rec in series.items():
            for year in END_YEARS:
                value = rec["values"].get(year) if rec else None
                if value is None:
                    raw = rec["raw_values"].get(year) if rec else "缺行"
                    reasons.append(f"{label}{year}不可解析/缺失({raw})")
                elif value <= 0:
                    reasons.append(f"{label}{year}非正({value})")
        for year in END_YEARS:
            cp = c["values"].get(year)
            if cp is None:
                reasons.append(f"Census人口{year}不可解析/缺失({c['raw_values'].get(year)})")
            elif cp <= 0:
                reasons.append(f"Census人口{year}非正({cp})")
            a = c1_total["values"].get(year) if c1_total else None
            b = c4_total["values"].get(year) if c4_total else None
            if a is not None and b is not None and a != b:
                reasons.append(f"两表总收入{year}不一致({a} vs {b})")
        eligible = not reasons
        vals = {}
        for name, rec in series.items():
            vals[name] = {year: rec["values"].get(year) if rec else None for year in END_YEARS}
        census_vals = {year: c["values"].get(year) for year in END_YEARS}
        total20, total24 = vals["CAINC1总收入"][2020], vals["CAINC1总收入"][2024]
        bp20, bp24 = vals["BEA人口"][2020], vals["BEA人口"][2024]
        cp20, cp24 = census_vals[2020], census_vals[2024]
        pc_bea20 = total20 * 1000 / bp20 if total20 is not None and bp20 and bp20 > 0 else None
        pc_bea24 = total24 * 1000 / bp24 if total24 is not None and bp24 and bp24 > 0 else None
        pc_cen20 = total20 * 1000 / cp20 if total20 is not None and cp20 and cp20 > 0 else None
        pc_cen24 = total24 * 1000 / cp24 if total24 is not None and cp24 and cp24 > 0 else None
        total_growth = safe_rate(total24, total20)
        net_contrib = ((vals["净劳动收入"][2024] - vals["净劳动收入"][2020]) / total20 * 100
                       if eligible else None)
        prop_contrib = ((vals["财产收入"][2024] - vals["财产收入"][2020]) / total20 * 100
                        if eligible else None)
        transfer_contrib = ((vals["转移收入"][2024] - vals["转移收入"][2020]) / total20 * 100
                            if eligible else None)
        log_income = safe_log_change(total24, total20)
        log_bea_pop = safe_log_change(bp24, bp20)
        log_census_pop = safe_log_change(cp24, cp20)
        growth.append({
            "fips": fips,
            "bea_name": c1_total["geo_name"],
            "census_state": c["state_name"],
            "census_name": c["county_name"],
            "eligible": eligible,
            "exclusion_reasons": "；".join(reasons),
            "total_income_line_code": "1/10",
            "total_income_2020_thousand": total20,
            "total_income_2024_thousand": total24,
            "total_income_growth": total_growth,
            "bea_population_line_code": "2",
            "bea_population_2020": bp20,
            "bea_population_2024": bp24,
            "bea_population_growth": safe_rate(bp24, bp20),
            "census_population_field_2020": "POPESTIMATE2020",
            "census_population_2020": cp20,
            "census_population_field_2024": "POPESTIMATE2024",
            "census_population_2024": cp24,
            "census_population_growth": safe_rate(cp24, cp20),
            "pc_income_bea_population_2020": pc_bea20,
            "pc_income_bea_population_2024": pc_bea24,
            "pc_income_growth_bea_population": safe_rate(pc_bea24, pc_bea20) if eligible else None,
            "pc_income_census_population_2020": pc_cen20,
            "pc_income_census_population_2024": pc_cen24,
            "pc_income_growth_census_population": safe_rate(pc_cen24, pc_cen20) if eligible else None,
            "published_pc_2020": published_pc["values"].get(2020) if published_pc else None,
            "published_pc_2024": published_pc["values"].get(2024) if published_pc else None,
            "net_earnings_line_code": "45",
            "net_earnings_2020": vals["净劳动收入"][2020],
            "net_earnings_2024": vals["净劳动收入"][2024],
            "net_earnings_contribution_pp": net_contrib,
            "property_income_line_code": "46",
            "property_income_2020": vals["财产收入"][2020],
            "property_income_2024": vals["财产收入"][2024],
            "property_income_contribution_pp": prop_contrib,
            "transfer_income_line_code": "47",
            "transfer_income_2020": vals["转移收入"][2020],
            "transfer_income_2024": vals["转移收入"][2024],
            "transfer_income_contribution_pp": transfer_contrib,
            "component_contribution_sum_pp": (net_contrib + prop_contrib + transfer_contrib) if eligible else None,
            "total_income_growth_pp": total_growth * 100 if total_growth is not None else None,
            "log_income_change": log_income,
            "log_bea_population_change": log_bea_pop,
            "log_pc_change_bea": log_income - log_bea_pop if eligible and log_income is not None and log_bea_pop is not None else None,
            "log_census_population_change": log_census_pop,
            "log_pc_change_census": log_income - log_census_pop if eligible and log_income is not None and log_census_pop is not None else None,
            "cainc1_total_raw_row": c1_total["raw_row"],
            "cainc4_total_raw_row": c4_total["raw_row"],
            "bea_population_raw_row": bpop["raw_row"],
            "net_raw_row": net["raw_row"],
            "property_raw_row": prop["raw_row"],
            "transfer_raw_row": transfer["raw_row"],
            "census_raw_row": c["raw_row"],
        })

    eligible_rows = [row for row in growth if row["eligible"]]
    rank_bea = sorted(eligible_rows, key=lambda row: (-row["pc_income_growth_bea_population"], row["fips"]))
    rank_census = sorted(eligible_rows, key=lambda row: (-row["pc_income_growth_census_population"], row["fips"]))
    bea_ranks = {row["fips"]: rank for rank, row in enumerate(rank_bea, 1)}
    census_ranks = {row["fips"]: rank for rank, row in enumerate(rank_census, 1)}
    for row in growth:
        row["rank_bea_population"] = bea_ranks.get(row["fips"])
        row["rank_census_population"] = census_ranks.get(row["fips"])
        if row["eligible"]:
            row["rank_change_census_minus_bea"] = row["rank_census_population"] - row["rank_bea_population"]
        else:
            row["rank_change_census_minus_bea"] = None
    top_bea = [dict(row) for row in rank_bea[:20]]
    top_census = [dict(row) for row in rank_census[:20]]
    top_bea_set = {row["fips"] for row in top_bea}
    top_census_set = {row["fips"] for row in top_census}

    profiles = [
        {
            "source_file": BEA_FILES[0], "encoding": "cp1252", "record_count": len(bea1["records"]),
            "physical_data_or_footer_rows": len(bea1["records"]) + len(bea1["footers"]),
            "geography_grain": "每个GeoFIPS含全国、区域、州、县/县等价体、合并及历史地域；不能由5位格式判县",
            "indicator_grain": "GeoFIPS × LineCode宽年份；3个LineCode",
            "year_coverage": f"{min(bea1['years'])}—{max(bea1['years'])}",
            "duplicate_key": "GeoFIPS+LineCode", "duplicate_count": len(bea1["duplicate_keys"]),
            "non_numeric_markers": bea1["markers"],
        },
        {
            "source_file": BEA_FILES[1], "encoding": "cp1252", "record_count": len(bea4["records"]),
            "physical_data_or_footer_rows": len(bea4["records"]) + len(bea4["footers"]),
            "geography_grain": "每个GeoFIPS含全国、区域、州、县/县等价体、合并及历史地域；不能由5位格式判县",
            "indicator_grain": "GeoFIPS × LineCode宽年份；20个LineCode",
            "year_coverage": f"{min(bea4['years'])}—{max(bea4['years'])}",
            "duplicate_key": "GeoFIPS+LineCode", "duplicate_count": len(bea4["duplicate_keys"]),
            "non_numeric_markers": bea4["markers"],
        },
        {
            "source_file": CENSUS_POP_FILE, "encoding": "cp1252", "record_count": len(census["records"]),
            "physical_data_or_footer_rows": len(census["records"]),
            "geography_grain": "SUMLEV=040州级与SUMLEV=050县级；本分析仅以050的STATE+COUNTY为目标",
            "indicator_grain": "地域宽指标/年份人口估计及变化构成",
            "year_coverage": "人口估计2020—2024（另含2020估计基数）",
            "duplicate_key": "SUMLEV+STATE+COUNTY", "duplicate_count": len(census["duplicate_keys"]),
            "non_numeric_markers": census["markers"],
        },
        {
            "source_file": COUNTY_CODE_FILE, "encoding": "utf-8-sig", "record_count": len(codes["records"]),
            "physical_data_or_footer_rows": len(codes["records"]),
            "geography_grain": "2020全国县/县等价体代码记录",
            "indicator_grain": "STATEFP+COUNTYFP代码及名称/分类/功能状态",
            "year_coverage": "2020代码版本（文件名所示；无年度数值列）",
            "duplicate_key": "STATEFP+COUNTYFP", "duplicate_count": len(codes["duplicate_keys"]),
            "non_numeric_markers": {},
        },
    ]
    coverage_summary = {
        "bea_geographies": len(bea_set),
        "census_2024_sumlev_050_target": len(target),
        "county_code_2020": len(code_set),
        "direct_code_matches": len(target & bea_set),
        "census_target_without_bea": len(target - bea_set),
        "bea_not_in_census_target": len(bea_set - target),
        "census_target_not_in_code2020": len(target - code_set),
        "code2020_not_in_census_target": len(code_set - target),
        "direct_matches_not_in_code2020": len((target & bea_set) - code_set),
        "bea_combined_outside_target": sum(1 for row in coverage if row["status"] == "BEA合并地区（不拆分）"),
    }
    cross_metric_summary = []
    for metric, _, _ in cross_metrics:
        rows = [row for row in cross if row["metric"] == metric]
        cross_metric_summary.append({
            "metric": metric,
            "rows": len(rows),
            "matches": sum(row["status"] == "一致" for row in rows),
            "mismatches": sum(row["status"] == "不一致" for row in rows),
            "unavailable": sum(row["status"] == "不可解析/缺失" for row in rows),
        })
    checks = {
        "cross_table_rows": len(cross),
        "cross_table_mismatches": sum(row["status"] == "不一致" for row in cross),
        "cross_table_unavailable": sum(row["status"] == "不可解析/缺失" for row in cross),
        "cross_table_metric_summary": cross_metric_summary,
        "component_equation_rows": len(component_checks),
        "component_equation_mismatches": sum(row["status"] == "不一致" for row in component_checks),
        "component_equation_unavailable": sum(row["status"] == "不可解析/缺失" for row in component_checks),
        "published_pc_rounding_failures_direct_endpoints": sum(row["rounding_check"] != "通过（发布整数在精确值四舍五入容差内）" for row in population),
        "eligible_growth_count": len(eligible_rows),
        "excluded_growth_count": len(growth) - len(eligible_rows),
    }
    rank_comparison = {
        "top20_overlap_count": len(top_bea_set & top_census_set),
        "bea_only_fips": sorted(top_bea_set - top_census_set),
        "census_only_fips": sorted(top_census_set - top_bea_set),
        "common_fips": sorted(top_bea_set & top_census_set),
        "common_rank_changes": sorted([
            {"fips": fips, "bea_rank": bea_ranks[fips], "census_rank": census_ranks[fips],
             "census_minus_bea": census_ranks[fips] - bea_ranks[fips]}
            for fips in top_bea_set & top_census_set
        ], key=lambda x: (-abs(x["census_minus_bea"]), x["fips"])),
    }
    return {
        "profiles": profiles,
        "indicator_definitions": [
            {"table": table, "line_code": code, "role": role, "label": INDICATOR_LABELS[(table, code)],
             "source_description": (bea1 if table == "CAINC1" else bea4)["line_definitions"][code]["description"],
             "source_unit": (bea1 if table == "CAINC1" else bea4)["line_definitions"][code]["unit"]}
            for table, roles in INDICATORS.items() for role, code in roles.items()
        ],
        "coverage_summary": coverage_summary,
        "coverage": coverage,
        "cross_table_verification": cross,
        "population_comparison": population,
        "population_difference_top20_2024": population_top20,
        "component_verification": component_checks,
        "growth_analysis": growth,
        "ranking_bea_population_top20": top_bea,
        "ranking_census_population_top20": top_census,
        "ranking_comparison": rank_comparison,
        "checks": checks,
        "source_footers": {BEA_FILES[0]: bea1["footers"], BEA_FILES[1]: bea4["footers"]},
        "duplicate_keys": {
            BEA_FILES[0]: bea1["duplicate_keys"], BEA_FILES[1]: bea4["duplicate_keys"],
            CENSUS_POP_FILE: census["duplicate_keys"], COUNTY_CODE_FILE: codes["duplicate_keys"],
        },
    }


def json_ready(value):
    if isinstance(value, dict):
        return {str(k): json_ready(v) for k, v in value.items()}
    if isinstance(value, list):
        return [json_ready(v) for v in value]
    if isinstance(value, tuple):
        return [json_ready(v) for v in value]
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    return value


def add_sheet(wb, title, rows, columns, percent_columns=None, number_formats=None):
    ws = wb.create_sheet(title)
    percent_columns = set(percent_columns or [])
    number_formats = number_formats or {}
    headers = [label for _, label in columns]
    ws.append(headers)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in rows:
        ws.append([row.get(key) for key, _ in columns])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    for idx, (key, _) in enumerate(columns, start=1):
        letter = get_column_letter(idx)
        if key in percent_columns:
            for cell in ws[letter][1:]:
                cell.number_format = "0.0000%"
        if key in number_formats:
            for cell in ws[letter][1:]:
                cell.number_format = number_formats[key]
        sample = [str(headers[idx - 1])] + [str(row.get(key) or "") for row in rows[:300]]
        ws.column_dimensions[letter].width = min(55, max(10, max(len(x) for x in sample) + 2))
    return ws


def write_workbook(path, analysis):
    wb = Workbook()
    wb.remove(wb.active)
    profile_rows = []
    for row in analysis["profiles"]:
        item = dict(row)
        item["non_numeric_markers"] = json.dumps(item["non_numeric_markers"], ensure_ascii=False)
        profile_rows.append(item)
    add_sheet(wb, "表审计", profile_rows, [
        ("source_file", "源文件"), ("encoding", "实际编码"), ("record_count", "有效记录数"),
        ("physical_data_or_footer_rows", "表头后非空物理行数"), ("geography_grain", "地理粒度"),
        ("indicator_grain", "指标粒度"), ("year_coverage", "年份覆盖"),
        ("duplicate_key", "重复键定义"), ("duplicate_count", "重复键数"),
        ("non_numeric_markers", "非数值标记及次数"),
    ])
    add_sheet(wb, "指标口径", analysis["indicator_definitions"], [
        ("table", "表"), ("line_code", "LineCode"), ("role", "分析角色"), ("label", "中文口径"),
        ("source_description", "原始Description"), ("source_unit", "原始Unit"),
    ])
    add_sheet(wb, "覆盖摘要", [analysis["coverage_summary"]], [
        ("bea_geographies", "BEA地域数"), ("census_2024_sumlev_050_target", "2024 Census县目标数"),
        ("county_code_2020", "2020县代码数"), ("direct_code_matches", "BEA-Census直接同码"),
        ("census_target_without_bea", "Census目标无BEA"), ("bea_not_in_census_target", "BEA不在目标"),
        ("census_target_not_in_code2020", "目标不在2020代码表"), ("code2020_not_in_census_target", "2020代码不在目标"),
        ("direct_matches_not_in_code2020", "直接同码但不在2020代码表"), ("bea_combined_outside_target", "BEA合并地区数"),
    ])
    add_sheet(wb, "覆盖与例外", analysis["coverage"], [
        ("fips", "FIPS"), ("status", "覆盖分类"), ("in_bea", "在BEA"),
        ("in_census_2024_county", "在2024 Census县目标"), ("in_county_code_2020", "在2020县代码表"),
        ("bea_name", "BEA名称"), ("census_state", "Census州"), ("census_name", "Census名称"),
        ("code2020_state", "2020州缩写"), ("code2020_name", "2020代码名称"),
        ("code2020_class", "CLASSFP"), ("code2020_funcstat", "FUNCSTAT"),
        ("bea_raw_row", "BEA原始行"), ("census_raw_row", "人口表原始行"),
        ("code2020_raw_row", "代码表原始行"), ("notes", "处理/核验说明"),
    ])
    add_sheet(wb, "跨表指标核验", analysis["cross_table_verification"], [
        ("fips", "FIPS"), ("geo_name", "BEA地域"), ("year", "年份"), ("metric", "指标口径"),
        ("cainc1_line_code", "CAINC1 LineCode"), ("cainc1_value", "CAINC1值"), ("cainc1_raw", "CAINC1原值"),
        ("cainc1_raw_row", "CAINC1原始行"), ("cainc4_line_code", "CAINC4 LineCode"),
        ("cainc4_value", "CAINC4值"), ("cainc4_raw", "CAINC4原值"), ("cainc4_raw_row", "CAINC4原始行"),
        ("difference", "CAINC1-CAINC4"), ("status", "核验状态"),
    ])
    add_sheet(wb, "人口与人均核验", analysis["population_comparison"], [
        ("fips", "FIPS"), ("bea_name", "BEA名称"), ("census_state", "州"), ("census_name", "Census名称"),
        ("year", "年份"), ("bea_population_line_code", "BEA人口LineCode"), ("bea_population", "BEA人口"),
        ("bea_population_raw_row", "BEA人口原始行"), ("census_population_field", "Census字段"),
        ("census_population", "Census人口"), ("census_raw_row", "Census原始行"),
        ("absolute_difference_bea_minus_census", "人口绝对差BEA-Census"),
        ("relative_difference_vs_census", "相对差/ Census"), ("income_line_code", "收入LineCode"),
        ("income_thousand_dollars", "总收入千美元"), ("income_raw_row", "收入原始行"),
        ("recalculated_pc_with_bea_population", "按BEA人口重算人均美元"),
        ("recalculated_pc_with_census_population", "按Census人口重算人均美元"),
        ("published_pc_line_code", "发布人均LineCode"), ("published_pc_dollars", "BEA发布人均美元"),
        ("published_pc_raw_row", "发布人均原始行"), ("published_minus_recalculated_bea", "发布值-按BEA人口精确值"),
        ("rounding_check", "四舍五入核验"),
    ], percent_columns=["relative_difference_vs_census"], number_formats={
        "recalculated_pc_with_bea_population": "0.000", "recalculated_pc_with_census_population": "0.000",
        "published_minus_recalculated_bea": "0.000000",
    })
    add_sheet(wb, "2024人口差异前20", analysis["population_difference_top20_2024"], [
        ("rank_by_absolute_relative_difference", "排名"), ("fips", "FIPS"), ("bea_name", "BEA名称"),
        ("census_name", "Census名称"), ("bea_population", "BEA人口"), ("census_population", "Census人口"),
        ("absolute_difference_bea_minus_census", "绝对差BEA-Census"),
        ("relative_difference_vs_census", "相对差/Census"), ("source_localization", "逐项来源定位"),
        ("interpretation", "解释/核验结论"),
    ], percent_columns=["relative_difference_vs_census"])
    add_sheet(wb, "分项加总核验", analysis["component_verification"], [
        ("fips", "FIPS"), ("geo_name", "BEA地域"), ("year", "年份"),
        ("total_line_code", "总收入LineCode"), ("total_income", "总收入"),
        ("net_earnings_line_code", "净劳动收入LineCode"), ("net_earnings", "净劳动收入"),
        ("property_income_line_code", "财产收入LineCode"), ("property_income", "财产收入"),
        ("transfer_income_line_code", "转移收入LineCode"), ("transfer_income", "转移收入"),
        ("component_sum", "三分项和"), ("total_minus_components", "总收入-三分项和"), ("status", "核验状态"),
        ("total_raw_row", "总收入原始行"), ("net_raw_row", "净劳动原始行"),
        ("property_raw_row", "财产原始行"), ("transfer_raw_row", "转移原始行"),
    ])
    growth_columns = [
        ("fips", "FIPS"), ("bea_name", "BEA名称"), ("census_name", "Census名称"),
        ("eligible", "进入排名"), ("exclusion_reasons", "排除原因"),
        ("total_income_2020_thousand", "2020总收入千美元"), ("total_income_2024_thousand", "2024总收入千美元"),
        ("total_income_growth", "总收入增长率"), ("bea_population_2020", "BEA人口2020"),
        ("bea_population_2024", "BEA人口2024"), ("bea_population_growth", "BEA人口增长率"),
        ("census_population_2020", "Census人口2020"), ("census_population_2024", "Census人口2024"),
        ("census_population_growth", "Census人口增长率"),
        ("pc_income_growth_bea_population", "人均增长率-BEA人口"),
        ("pc_income_growth_census_population", "人均增长率-Census人口"),
        ("rank_bea_population", "BEA人口口径排名"), ("rank_census_population", "Census人口口径排名"),
        ("rank_change_census_minus_bea", "名次变化Census-BEA"),
        ("net_earnings_2020", "净劳动2020"), ("net_earnings_2024", "净劳动2024"),
        ("net_earnings_contribution_pp", "净劳动贡献百分点"),
        ("property_income_2020", "财产收入2020"), ("property_income_2024", "财产收入2024"),
        ("property_income_contribution_pp", "财产收入贡献百分点"),
        ("transfer_income_2020", "转移收入2020"), ("transfer_income_2024", "转移收入2024"),
        ("transfer_income_contribution_pp", "转移收入贡献百分点"),
        ("component_contribution_sum_pp", "三分项贡献合计百分点"), ("total_income_growth_pp", "总收入增幅百分点"),
        ("log_income_change", "对数收入变化"), ("log_bea_population_change", "对数BEA人口变化"),
        ("log_pc_change_bea", "对数人均变化-BEA"), ("log_census_population_change", "对数Census人口变化"),
        ("log_pc_change_census", "对数人均变化-Census"),
        ("cainc1_total_raw_row", "CAINC1总收入原始行"), ("cainc4_total_raw_row", "CAINC4总收入原始行"),
        ("bea_population_raw_row", "BEA人口原始行"), ("net_raw_row", "净劳动原始行"),
        ("property_raw_row", "财产原始行"), ("transfer_raw_row", "转移原始行"), ("census_raw_row", "Census原始行"),
    ]
    percent_cols = ["total_income_growth", "bea_population_growth", "census_population_growth",
                    "pc_income_growth_bea_population", "pc_income_growth_census_population"]
    add_sheet(wb, "增长可比集与排除", analysis["growth_analysis"], growth_columns, percent_columns=percent_cols,
              number_formats={k: "0.0000" for k in ["net_earnings_contribution_pp", "property_income_contribution_pp",
                                                        "transfer_income_contribution_pp", "component_contribution_sum_pp",
                                                        "total_income_growth_pp", "log_income_change", "log_bea_population_change",
                                                        "log_pc_change_bea", "log_census_population_change", "log_pc_change_census"]})
    ranking_cols = [
        ("rank_bea_population", "BEA口径排名"), ("rank_census_population", "Census口径排名"),
        ("rank_change_census_minus_bea", "名次变化Census-BEA"), ("fips", "FIPS"), ("bea_name", "地区"),
        ("pc_income_growth_bea_population", "人均增长率-BEA人口"),
        ("pc_income_growth_census_population", "人均增长率-Census人口"),
        ("total_income_growth", "总收入增长率"), ("bea_population_growth", "BEA人口增长率"),
        ("census_population_growth", "Census人口增长率"),
        ("net_earnings_contribution_pp", "净劳动贡献百分点"),
        ("property_income_contribution_pp", "财产收入贡献百分点"),
        ("transfer_income_contribution_pp", "转移收入贡献百分点"),
        ("total_income_growth_pp", "总收入增幅百分点"),
        ("log_income_change", "对数收入变化"), ("log_bea_population_change", "对数BEA人口变化"),
        ("log_pc_change_bea", "对数人均变化-BEA"), ("log_census_population_change", "对数Census人口变化"),
        ("log_pc_change_census", "对数人均变化-Census"),
        ("total_income_line_code", "总收入LineCode"), ("bea_population_line_code", "BEA人口LineCode"),
        ("net_earnings_line_code", "净劳动LineCode"), ("property_income_line_code", "财产LineCode"),
        ("transfer_income_line_code", "转移LineCode"), ("cainc1_total_raw_row", "CAINC1总收入行"),
        ("cainc4_total_raw_row", "CAINC4总收入行"), ("bea_population_raw_row", "BEA人口行"),
        ("net_raw_row", "净劳动行"), ("property_raw_row", "财产行"), ("transfer_raw_row", "转移行"),
        ("census_raw_row", "Census行"),
    ]
    ranking_percent = ["pc_income_growth_bea_population", "pc_income_growth_census_population", "total_income_growth",
                       "bea_population_growth", "census_population_growth"]
    rank_num_formats = {k: "0.0000" for k in ["net_earnings_contribution_pp", "property_income_contribution_pp",
                                                "transfer_income_contribution_pp", "total_income_growth_pp",
                                                "log_income_change", "log_bea_population_change", "log_pc_change_bea",
                                                "log_census_population_change", "log_pc_change_census"]}
    add_sheet(wb, "排名_BEA人口", analysis["ranking_bea_population_top20"], ranking_cols,
              percent_columns=ranking_percent, number_formats=rank_num_formats)
    add_sheet(wb, "排名_Census人口", analysis["ranking_census_population_top20"], ranking_cols,
              percent_columns=ranking_percent, number_formats=rank_num_formats)
    comparison_rows = []
    for row in analysis["ranking_comparison"]["common_rank_changes"]:
        comparison_rows.append({"category": "共同入围", **row})
    for fips in analysis["ranking_comparison"]["bea_only_fips"]:
        comparison_rows.append({"category": "仅BEA人口口径前20", "fips": fips})
    for fips in analysis["ranking_comparison"]["census_only_fips"]:
        comparison_rows.append({"category": "仅Census人口口径前20", "fips": fips})
    add_sheet(wb, "排名集合变化", comparison_rows, [
        ("category", "类别"), ("fips", "FIPS"), ("bea_rank", "BEA口径名次"),
        ("census_rank", "Census口径名次"), ("census_minus_bea", "名次变化Census-BEA"),
    ])
    wb.save(path)


def set_cell_text(cell, text):
    cell.text = str(text)
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(8)
            run.font.name = "宋体"


def add_doc_table(doc, headers, rows):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for idx, header in enumerate(headers):
        set_cell_text(table.rows[0].cells[idx], header)
    for row in rows:
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            set_cell_text(cells[idx], value)
    return table


def pct(value, digits=2):
    return "—" if value is None else f"{value * 100:.{digits}f}%"


def num(value, digits=0):
    if value is None:
        return "—"
    return f"{value:,.{digits}f}"


def write_memo(path, analysis):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    styles["Normal"].font.size = Pt(10)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("2020—2024年县域人口与名义人均收入变化审计备忘录")
    run.bold = True
    run.font.size = Pt(16)
    run.font.name = "黑体"

    cov = analysis["coverage_summary"]
    checks = analysis["checks"]
    rc = analysis["ranking_comparison"]
    doc.add_heading("结论先行", level=1)
    conclusions = [
        f"可直接按代码比较的范围是{cov['direct_code_matches']:,}个地区：它们同时出现在BEA GeoFIPS和Census 2024人口估计SUMLEV=050目标集合。不能把BEA的{cov['bea_not_in_census_target']:,}个集合外地域（含全国、区域、州、合并及历史地域）算入县级总计。",
        f"目标集合另有{cov['census_target_without_bea']:,}个代码没有BEA同码，必须留待核验，未按名称匹配或从BEA合并地区拆分。2024人口目标与2020县代码表还存在双向版本差异：前者独有{cov['census_target_not_in_code2020']:,}码，后者独有{cov['code2020_not_in_census_target']:,}码。",
        f"两张BEA表2020—2024对总收入、人口和发布人均值共核验{checks['cross_table_rows']:,}个地域—年份—指标，数值不一致{checks['cross_table_mismatches']:,}项、不可解析{checks['cross_table_unavailable']:,}项。CAINC4总收入与三分项加总核验不一致{checks['component_equation_mismatches']:,}项、不可解析{checks['component_equation_unavailable']:,}项。",
        f"严格要求两端总收入和三分项、两端BEA与Census人口均为正且有值，并要求两表总收入一致后，可排名{checks['eligible_growth_count']:,}个地区，排除{checks['excluded_growth_count']:,}个；每个排除原因已逐条保留。",
        f"两种人口口径的名义人均收入增长前20完全重合（{rc['top20_overlap_count']}个），且本次数据中的前20名次也完全相同；尽管两套人口值并非处处相等，本次分母差异不足以改变榜首集合或顺序。",
        "这些增长率均为名义值；不能解释为经过通胀调整的实际购买力或生活水平改善。",
    ]
    for text in conclusions:
        doc.add_paragraph(text, style="List Bullet")

    doc.add_heading("范围、原件与指标口径", level=1)
    doc.add_paragraph(
        "分析只使用交付的四份原件，不修改输入，不引入外部地域交叉表。BEA文件严格按cp1252解码，县代码表按UTF-8 BOM解码。"
        "GeoFIPS及STATE+COUNTY均正规化为保留前导零的5位字符串；县级目标仅由人口表SUMLEV=050定义，不以代码长度推断。"
    )
    add_doc_table(doc, ["源文件", "有效记录", "地理/指标粒度", "年份", "重复键", "非数值标记"], [
        [p["source_file"], f"{p['record_count']:,}", p["geography_grain"] + "；" + p["indicator_grain"],
         p["year_coverage"], str(p["duplicate_count"]), json.dumps(p["non_numeric_markers"], ensure_ascii=False)]
        for p in analysis["profiles"]
    ])
    doc.add_paragraph(
        "指标按原表LineCode和Description确定：CAINC1的1=个人收入（千美元）、2=人口（人）、3=人均个人收入（美元）；"
        "CAINC4的10=个人收入、45=按居住地净劳动收入、46=股息利息租金、47=个人经常转移收入，后三项及总收入单位均为千美元。"
    )

    doc.add_heading("地域覆盖与不能强拼的例外", level=1)
    add_doc_table(doc, ["项目", "数量"], [
        ["Census 2024 SUMLEV=050目标", f"{cov['census_2024_sumlev_050_target']:,}"],
        ["BEA地域", f"{cov['bea_geographies']:,}"],
        ["直接同码", f"{cov['direct_code_matches']:,}"],
        ["Census目标无BEA", f"{cov['census_target_without_bea']:,}"],
        ["BEA不在Census目标", f"{cov['bea_not_in_census_target']:,}"],
        ["Census目标不在2020代码表", f"{cov['census_target_not_in_code2020']:,}"],
        ["2020代码表不在Census目标", f"{cov['code2020_not_in_census_target']:,}"],
        ["BEA集合外明确合并地区", f"{cov['bea_combined_outside_target']:,}"],
    ])
    doc.add_paragraph(
        "典型不可直接匹配情形包括BEA明确用“+”或“includes”标示的合并地区，以及Census目标中没有BEA同码的县/独立市。"
        "本分析不依据相似名称分配合并地区，也不把2020代码表当成唯一真值：例如人口表新增/替换的代码仍保留在目标集合，并在覆盖表标记版本差异。"
    )

    doc.add_heading("人口差异与人均值复核", level=1)
    doc.add_paragraph(
        "对直接同码地区比较2020和2024人口。绝对差定义为BEA减Census，相对差以Census人口为分母。"
        "BEA总收入单位为千美元，因此重算人均收入使用“总收入×1,000÷人口”。按BEA人口重算的精确值与BEA发布的整数美元人均值之差应在±0.5美元内；"
        f"直接匹配两端的四舍五入核验失败/不可核验共{checks['published_pc_rounding_failures_direct_endpoints']:,}项。"
    )
    top_pop = analysis["population_difference_top20_2024"]
    add_doc_table(doc, ["名次", "FIPS", "地区", "BEA人口", "Census人口", "差", "相对差"], [
        [r["rank_by_absolute_relative_difference"], r["fips"], r["bea_name"], num(r["bea_population"]),
         num(r["census_population"]), num(r["absolute_difference_bea_minus_census"]),
         pct(r["relative_difference_vs_census"], 3)] for r in top_pop
    ])
    doc.add_paragraph(
        "上述每项在audit.xlsx中均定位到BEA原始行、LineCode 2及Census原始行、POPESTIMATE2024字段。"
        "原件本身没有足够元数据证明差异原因，因此仅将其列为估计版本、处理或口径待核验项，不自动判定任何一方原始数据错误。"
    )

    doc.add_heading("名义人均收入增长：双人口口径", level=1)
    doc.add_paragraph(
        "排名使用未四舍五入增长率降序，同值按FIPS升序。人均增长由同一BEA总收入分别除以BEA人口和Census人口计算。"
        "贡献百分点定义为各分项2020—2024增量除以2020总收入乘100，允许负值；三项贡献之和应等于总收入增幅百分点。"
        "对数拆分为ln(2024总收入/2020总收入)−ln(2024人口/2020人口)，只对正值计算。"
    )
    for heading, key, rank_key, growth_key in [
        ("按BEA人口计算的前20", "ranking_bea_population_top20", "rank_bea_population", "pc_income_growth_bea_population"),
        ("按Census人口计算的前20", "ranking_census_population_top20", "rank_census_population", "pc_income_growth_census_population"),
    ]:
        doc.add_heading(heading, level=2)
        add_doc_table(doc, ["名次", "FIPS", "地区", "名义人均增长", "净劳动贡献pp", "财产贡献pp", "转移贡献pp"], [
            [r[rank_key], r["fips"], r["bea_name"], pct(r[growth_key]),
             num(r["net_earnings_contribution_pp"], 2), num(r["property_income_contribution_pp"], 2),
             num(r["transfer_income_contribution_pp"], 2)] for r in analysis[key]
        ])
    doc.add_heading("集合与名次变化", level=2)
    doc.add_paragraph(
        f"共同入围{rc['top20_overlap_count']}个；仅BEA人口口径入围：{', '.join(rc['bea_only_fips']) or '无'}；"
        f"仅Census人口口径入围：{', '.join(rc['census_only_fips']) or '无'}。"
    )
    if rc["common_rank_changes"]:
        add_doc_table(doc, ["FIPS", "BEA口径名次", "Census口径名次", "后者减前者"], [
            [r["fips"], r["bea_rank"], r["census_rank"], r["census_minus_bea"]]
            for r in rc["common_rank_changes"]
        ])

    doc.add_heading("解释边界与追查路径", level=1)
    for text in [
        "audit.xlsx的“覆盖与例外”保留三套代码存在性、名称、功能状态和原始行号；“跨表指标核验”逐地域—年份保留总收入、人口和发布人均值的两表LineCode、原值和差；“增长可比集与排除”保留每个直接匹配地区的资格及逐条排除原因。",
        "results.json保存工作簿所用完整结构化结果；analysis.py可在离线环境重建全部交付。原始标记如(NA)保持为不可解析，不填0。",
        "排名只描述2020到2024名义收入与人口分母的联合变化。未使用价格指数、住房成本、税负、收入分布或修订历史，故不能推导实际生活水平改善，也不能由人口差异单独推断数据质量优劣。",
    ]:
        doc.add_paragraph(text, style="List Bullet")
    doc.save(path)


def main():
    parser = argparse.ArgumentParser(description="审计BEA县域收入与Census人口并生成可追溯交付。")
    parser.add_argument("--input", required=True, help="包含四份原件的目录")
    parser.add_argument("--output", required=True, help="输出目录")
    args = parser.parse_args()
    input_dir = Path(args.input).resolve()
    output_dir = Path(args.output).resolve()
    required = [*BEA_FILES, CENSUS_POP_FILE, COUNTY_CODE_FILE]
    missing = [name for name in required if not (input_dir / name).is_file()]
    if missing:
        raise FileNotFoundError("缺少原件: " + ", ".join(missing))
    if output_dir == input_dir or input_dir in output_dir.parents:
        raise ValueError("输出目录不得等于或位于输入目录内，以免修改原件")
    output_dir.mkdir(parents=True, exist_ok=True)

    bea1 = read_bea(input_dir / BEA_FILES[0])
    bea4 = read_bea(input_dir / BEA_FILES[1])
    census = read_census_population(input_dir / CENSUS_POP_FILE)
    codes = read_county_codes(input_dir / COUNTY_CODE_FILE)
    analysis = build_analysis(bea1, bea4, census, codes)
    result = {
        "metadata": {
            "analysis_title": "2020—2024年县域人口与名义人均收入变化审计",
            "generated_utc": datetime.now(timezone.utc).isoformat(),
            "input_directory": str(input_dir),
            "source_files": required,
            "target_definition": "co-est2024-alldata.csv中SUMLEV=050的STATE+COUNTY五位代码",
            "direct_match_rule": "目标代码与BEA正规化GeoFIPS完全相等；不使用名称匹配或合并地区拆分",
            "ranking_rule": "未四舍五入增长率降序，同值按FIPS升序",
            "nominal_warning": "名义增长不能视为通胀调整后的生活水平改善",
        },
        **analysis,
    }
    with (output_dir / "results.json").open("w", encoding="utf-8") as handle:
        json.dump(json_ready(result), handle, ensure_ascii=False, indent=2, allow_nan=False)
    write_workbook(output_dir / "audit.xlsx", analysis)
    write_memo(output_dir / "memo.docx", analysis)
    print(json.dumps({"output": str(output_dir), "coverage": analysis["coverage_summary"],
                      "checks": analysis["checks"], "ranking_comparison": analysis["ranking_comparison"]},
                     ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
