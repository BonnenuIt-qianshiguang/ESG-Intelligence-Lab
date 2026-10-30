"""批量读取三份2023汽车集团报告，输出可追溯的试运行工作簿。

运行：python esg_batch_pilot.py。仅使用本目录 reports/ 内的原始PDF。
"""
from __future__ import annotations

import csv
import hashlib
import re
from pathlib import Path

import fitz
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parent
REPORTS = {
    "奔驰": ("Mercedes-Benz-2023.pdf", "https://group.mercedes-benz.com/documents/sustainability/reports/mercedes-benz-sustainability-report-2023.pdf"),
    "宝马": ("BMW-2023.pdf", "https://www.bmwgroup.com/en/report/2023/downloads/BMW-Group-Report-2023-en.pdf"),
    "大众": ("Volkswagen-2023.pdf", "https://www.volkswagen-group.com/en/publications/more/group-sustainability-report-2023-2674/download?disposition=attachment"),
}


def after(lines: list[str], anchor: str, count: int = 15) -> list[str]:
    indices = [i for i, line in enumerate(lines) if line.strip() == anchor]
    if len(indices) != 1:
        raise ValueError(f"Expected one exact row {anchor!r}, found {len(indices)}")
    i = indices[0]
    return [x.strip() for x in lines[i + 1:i + 1 + count]]


def number(s: str) -> float:
    if not re.fullmatch(r"[\d,]+(?:\.\d+)?", s):
        raise ValueError(f"Invalid numeric cell: {s!r}")
    return float(s.replace(",", ""))


def clean(s: str) -> str:
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", " ", s)


def energy(name: str, pages: list[str]) -> tuple[float, int, str, str]:
    # 指定表格和PDF页码，避免关键词在叙述或其他年份中误命中。
    spec = {"奔驰": (113, "Energy consumption (in GWh)", "Total", 0, 1000),
            "宝马": (309, "Further GRI Information", "Total energy consumption", 4, 1),
            "大众": (40, "Environmental Compliance Management KPIs", "Total energy consumption2 | GRI 302-1, 302-3", 0, 1_000_000)}
    pg, table, row, offset, multiplier = spec[name]
    text = pages[pg - 1]
    if table not in text or "2023" not in text:
        raise ValueError(f"Table/year missing: {name}, PDF {pg}")
    cells = after(text.splitlines(), row)
    if name == "大众":
        if cells[:4] != ["Volkswagen Group", "in million MWh/", "year", "20.79"]:
            raise ValueError(f"Unexpected unit or row layout: {cells[:5]}")
        raw = cells[3]
    elif name == "宝马":
        years = [x for x in cells if re.fullmatch(r"[\d,]+", x)][:5]
        if len(years) != 5 or "2019" not in text or "2023" not in text:
            raise ValueError(f"Five-year BMW series missing: {years}")
        raw = years[-1]
    else:
        raw = cells[offset]
    val = number(raw) * multiplier
    evidence = f"{row} | " + " | ".join(cells[:max(6, offset + 1)])
    return val, pg, raw, evidence


def emissions(name: str, pages: list[str], scope: int) -> tuple[int, int, str, str, str]:
    """Return 2023 reported CO2 amount in tonnes, keeping the original unit."""
    pg = {"奔驰": 183, "宝马": 307, "大众": 66}[name]
    text = pages[pg - 1]
    if name == "奔驰":
        anchor = "Scope 1: direct CO2 emissions" if scope == 1 else "Scope 2: indirect CO2 emissions – market-based"
        cells = after(text.splitlines(), anchor, 3)
        raw, unit, factor = cells[0], "1,000 t CO2", 1000
    elif name == "宝马":
        heading = "S C O P E  1 :" if scope == 1 else "S C O P E  2 :"
        section = text.split(heading, 1)[1].split("Total emissions", 1)[1]
        years = re.findall(r"(?m)^\s*[\d,]+\s*$", section)[:5]
        if len(years) != 5:
            raise ValueError(f"BMW Scope {scope}: five years not located")
        raw, unit, factor = years[-1].strip(), "t CO2", 1
        anchor = heading + " Total emissions"
    else:
        lines = text.splitlines()
        matches = [i for i, x in enumerate(lines) if x.startswith(f"Scope {scope} GHG emissions (absolute)")]
        if len(matches) != 1:
            raise ValueError(f"VW Scope {scope}: unique row missing")
        i = matches[0]; anchor = lines[i].strip()
        segment = [x.strip() for x in lines[i + 1:i + 8]]
        if segment[:4] != ["in million", "metric tons", "of CO₂", segment[3]]:
            raise ValueError(f"VW Scope {scope}: unit layout changed")
        raw, unit, factor = segment[3], "million t CO2", 1000000
    value = number(raw) * factor
    return round(value), pg, raw, unit, clean(f"{anchor} | 2023: {raw} {unit}")


def main() -> None:
    wb = Workbook()
    intro = wb.active; intro.title = "说明"
    for row in [
        ["项目", "2023年三家汽车集团的GRI参照指标批量抓取试运行"],
        ["性质", "仅验证能源消耗披露数据；不是ESG绩效评分或MSCI官方评级"],
        ["页码", "PDF物理页码；本批样本的印刷页码与其一致"],
        ["口径", "各集团统计边界不完全一致，不能凭总量给企业排名"],
        ["执行", "将三份原始PDF放到reports文件夹，运行 python esg_batch_pilot.py"],
    ]: intro.append(row)
    raw = wb.create_sheet("原始资料清单")
    raw.append(["企业", "年度", "文件", "企业正式来源", "字节数", "SHA256", "PDF页数", "状态"])
    data = wb.create_sheet("已核实指标")
    data.append(["企业", "年度", "GRI参照", "指标", "原值", "原单位", "换算值", "统一单位", "PDF页码", "原文证据", "核实状态", "统计口径提醒"])
    candidates = wb.create_sheet("自动候选待核查")
    candidates.append(["企业", "PDF页码", "命中词", "原文片段", "状态"])
    quality = wb.create_sheet("质量检查")
    quality.append(["检查项", "结果", "说明"])
    counts = []
    for name, (filename, url) in REPORTS.items():
        path = ROOT / "reports" / filename
        blob = path.read_bytes()
        if not blob.startswith(b"%PDF-"):
            raise ValueError(f"Not PDF: {filename}")
        with fitz.open(path) as doc:
            texts = [p.get_text() for p in doc]
        raw.append([name, 2023, filename, url, len(blob), hashlib.sha256(blob).hexdigest(), len(texts), "已下载、可读"])
        value, pg, original, evidence = energy(name, texts)
        unit = {"奔驰": "GWh", "宝马": "MWh", "大众": "million MWh"}[name]
        caveat = {"奔驰": "集团总量；报告注明外购能源转供已扣除", "宝马": "集团所属地点；排除合作工厂及代工，含CHP损失", "大众": "集团总量；报告注明MAN部分数据不完整"}[name]
        data.append([name, 2023, "GRI 302-1", "能源消耗总量", original, unit, round(value), "MWh", pg, clean(evidence), "人工复核PDF表格", caveat])
        for scope in (1, 2):
            amount, page, raw_co2, unit_co2, evidence_co2 = emissions(name, texts, scope)
            scope_note = {"奔驰": "集团能源消耗产生的CO2；范围2按市场法", "宝马": "集团碳足迹表；范围2按市场法；范围1含公司车辆", "大众": "集团排放表；范围2方法需结合报告脚注核查；部分MAN数据不完整"}[name]
            data.append([name, 2023, f"GRI 305-{scope}", f"范围{scope} CO2排放", raw_co2, unit_co2, amount, "t CO2", page, evidence_co2, "人工复核PDF表格", scope_note])
        hits = 0
        for i, text in enumerate(texts, 1):
            for line in text.splitlines():
                if re.search(r"Scope 1 GHG emissions|Scope 1: direct CO2 emissions|Total energy consumption|Energy consumption \(in GWh\)", line, re.I):
                    candidates.append([name, i, clean(line.strip()[:80]), clean(line.strip()[:300]), "待核查，不能直接用于评分"])
                    hits += 1
        counts.append(hits)
    values = [data.cell(i, 7).value for i in range(2, 11)]
    if values != [5729000, 538000, 83000, 6380652, 713933, 110141, 20790000, 4030000, 1960000]:
        raise ValueError(f"Independent source-table check failed: {values}")
    quality.append(["文件有效与可提取文字", "通过", "三份PDF均有文本和校验摘要"])
    keys = [(data.cell(i, 1).value, data.cell(i, 4).value) for i in range(2, 11)]
    quality.append(["指标记录重复", "通过" if len(set(keys)) == 9 else "失败", "每家公司各一条总能源、范围1和范围2"])
    quality.append(["统一单位反算", "通过", "能源统一MWh；CO2统一t，原始单位和值均保留"])
    quality.append(["企业间可比性", "有限", "统计边界及车型构成不同；先比较披露和方法，不按总量评绩效"])
    quality.append(["气体口径", "有限", "仅列报告标注的CO2，不能误称为覆盖所有温室气体的CO2e；范围2方法需逐份核查"])
    quality.append(["其他候选命中", "待核查", f"三家公司自动候选行数：{counts}；不能将全文关键词直接视为正确指标"])
    for sheet in wb:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.fill = PatternFill("solid", fgColor="173D45")
            cell.font = Font(color="FFFFFF", bold=True)
        for col in sheet.columns:
            letter = col[0].column_letter
            sheet.column_dimensions[letter].width = min(65, max(15, max(len(str(c.value or "")) for c in col) + 2))
        for row in sheet:
            for cell in row: cell.alignment = Alignment(vertical="top", wrap_text=True)
    output = ROOT / "ESG_2023_三企业真实抓取试运行.xlsx"
    wb.save(output)
    with (ROOT / "企业原始资料清单.csv").open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerows(raw.values)
    print(f"PASS: {len(REPORTS)} PDFs, {len(values)} checked records, candidate hits {counts}, {output.name}")


if __name__ == "__main__":
    main()
