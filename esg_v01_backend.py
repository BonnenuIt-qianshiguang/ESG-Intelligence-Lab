"""V1.0 云端网页版后端：保留既有三样本规则，其他PDF仅生成待核查候选。"""
from __future__ import annotations

from dataclasses import dataclass, field
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import re

import fitz
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from esg_batch_pilot import REPORTS, after, clean, emissions, energy, number

KNOWN = {
    "fc182e2f9c363b7011abe38e0654e97ab8aad1c9e7ce4ba5c33b58b712b5617e": "奔驰",
    "07a384192eaec2286a221fc7d1c86d6f74a83ab0f0c31e90764764caf74694a9": "宝马",
    "b39ced16f7eca2cf1f7bb3273096c349308717f80e16fd0f69fb5e6a364cecde": "大众",
}
FIELDS = {
    "E": [("能源", r"energy consumption|total energy|能源消耗"),
          ("排放", r"scope\s*[123]|greenhouse gas emissions|温室气体排放")],
    "S": [("员工", r"total employees|number of employees|workforce|员工总数"),
          ("职业健康安全", r"work.related injur|occupational accident|职业健康|工伤")],
    "G": [("反腐败", r"anti.corruption|anti.bribery|反腐败|反贿赂"),
          ("董事会", r"supervisory board|board of directors|董事会")],
}
HEADERS = ["企业", "年度", "维度", "指标", "原值或内容", "原单位", "换算值", "统一单位",
           "报告文件", "PDF页码", "原文证据", "GRI参照", "提取方式", "核验状态", "口径提醒", "来源网址"]


@dataclass
class Batch:
    records: list[dict] = field(default_factory=list)
    candidates: list[dict] = field(default_factory=list)
    files: list[dict] = field(default_factory=list)


def _context(text: str, target: str, before: int = 0, after: int = 7) -> str:
    lines = [clean(x.strip()) for x in text.splitlines()]
    matches = [i for i, x in enumerate(lines) if target.lower() in x.lower()]
    if not matches:
        raise ValueError(f"原始证据行未找到: {target}")
    i = matches[0]
    return " | ".join(x for x in lines[max(0, i-before):i+after] if x)[:700]


def _known_rows(company: str, pages: list[str], filename: str) -> list[dict]:
    url = REPORTS[company][1]
    val, page, raw, _ = energy(company, pages)
    unit = {"奔驰": "GWh", "宝马": "MWh", "大众": "million MWh"}[company]
    key = {"奔驰": "Total", "宝马": "Total energy consumption", "大众": "Total energy consumption2"}[company]
    rows = [{"企业": company, "年度": 2023, "维度": "E", "指标": "能源消耗总量",
             "原值或内容": raw, "原单位": unit, "换算值": round(val), "统一单位": "MWh",
             "报告文件": filename, "PDF页码": page, "原文证据": _context(pages[page-1], key, after=19),
             "GRI参照": "GRI 302-1", "提取方式": "已知样本专用规则",
             "核验状态": "已与原报告人工核验；仅适用于同一文件摘要",
             "口径提醒": "集团统计边界和产品结构不同，不能以总量直接评价绩效", "来源网址": url}]
    for scope in (1, 2):
        amount, page, original, original_unit, _ = emissions(company, pages, scope)
        anchor = {"奔驰": f"Scope {scope}: ", "宝马": "S C O P E  " + str(scope) + " :", "大众": f"Scope {scope} GHG emissions (absolute)"}[company]
        rows.append({"企业": company, "年度": 2023, "维度": "E", "指标": f"范围{scope} CO2排放",
                     "原值或内容": original, "原单位": original_unit, "换算值": amount, "统一单位": "t CO2",
                     "报告文件": filename, "PDF页码": page,
                     "原文证据": _context(pages[page-1], anchor, after=18),
                     "GRI参照": f"GRI 305-{scope}", "提取方式": "已知样本专用规则",
                     "核验状态": "已与原报告人工核验；仅适用于同一文件摘要",
                     "口径提醒": "报告标注CO2，不能扩大解释为所有温室气体CO2e；范围2方法需按企业核查",
                     "来源网址": url})
    employee_page = {"奔驰": 187, "宝马": 314, "大众": 101}[company]
    employee_text = pages[employee_page-1]
    if company == "奔驰":
        if "Workforce1" not in employee_text or "GRI 2-7" not in employee_text:
            raise ValueError("奔驰员工表标题缺失")
        section = employee_text.split("Workforce1", 1)[1].split("Workforce by division1", 1)[0]
        parts = after(section.splitlines(), "Total", 3)
        employee_raw = parts[0]
        employee_evidence = _context(section, "Total", after=5)
    elif company == "宝马":
        section = employee_text.split("Employees at end of year1", 1)[1]
        segment = section.split("BMW Group", 1)[1]
        numbers = re.findall(r"(?m)^\s*[\d,]+\s*$", segment)[:5]
        if len(numbers) != 5 or "2023" not in employee_text:
            raise ValueError("宝马员工五年表结构变化")
        employee_raw = numbers[-1].strip()
        employee_evidence = _context(section, "BMW Group", after=15)
    else:
        segment = employee_text.split("Total workforce (of which Volkswagen AG)", 1)[1]
        numbers = re.findall(r"(?m)^\s*[\d,]+\s*$", segment)[:1]
        if len(numbers) != 1 or "GRI 2-7" not in segment[:60]:
            raise ValueError("大众员工表结构变化")
        employee_raw = numbers[0].strip()
        employee_evidence = _context(employee_text, "Total workforce (of which Volkswagen AG)", after=8)
    rows.append({"企业": company, "年度": 2023, "维度": "S", "指标": "期末员工人数",
                 "原值或内容": employee_raw, "原单位": "人", "换算值": round(number(employee_raw)),
                 "统一单位": "人", "报告文件": filename, "PDF页码": employee_page,
                 "原文证据": employee_evidence, "GRI参照": "GRI 2-7",
                 "提取方式": "已知样本专用规则", "核验状态": "人工核对原报告表格；仅适用于同一文件摘要",
                 "口径提醒": "各集团员工定义和并表范围不同；宝马图示近似数不能替代表格数",
                 "来源网址": url})
    governance_page = {"奔驰": 41, "宝马": 142, "大众": 111}[company]
    governance_text = pages[governance_page-1]
    anchor = {"奔驰": "The Mercedes-Benz Group’s corruption prevention",
              "宝马": "The CMS is tailored to the BMW Group’s risk situation",
              "大众": "The focuses of the Group Integrity & Compliance"}[company]
    start = governance_text.find(anchor)
    if start < 0:
        raise ValueError(f"反腐败原文锚点缺失：{company}")
    source_excerpt = clean(" ".join(governance_text[start:start+850].split()))
    if "corruption" not in source_excerpt.lower():
        raise ValueError(f"反腐败证据段未包含主题：{company}")
    rows.append({"企业": company, "年度": 2023, "维度": "G", "指标": "反腐败风险管理描述",
                 "原值或内容": "报告描述风险评估或合规管理举措", "原单位": "定性披露",
                 "换算值": "", "统一单位": "", "报告文件": filename, "PDF页码": governance_page,
                 "原文证据": source_excerpt, "GRI参照": "GRI 205-1（议题参照，非完整符合性判断）",
                 "提取方式": "已知样本专用规则", "核验状态": "人工核对原报告段落；仅表示企业披露",
                 "口径提醒": "企业管理举措的陈述不证明实施效果，也不等于没有腐败事件", "来源网址": url})
    return rows


def _candidates(pages: list[str], filename: str, company: str) -> list[dict]:
    found = []
    for page, text in enumerate(pages, 1):
        if not text.strip():
            continue
        lines = [clean(x.strip()) for x in text.splitlines()]
        for i, line in enumerate(lines):
            if not line or len(line) > 220:
                continue
            for dim, terms in FIELDS.items():
                for metric, pattern in terms:
                    if re.search(pattern, line, re.I):
                        excerpt = " | ".join(x for x in lines[max(0, i-1):i+5] if x)[:500]
                        found.append({"企业": company, "年度": 2023 if "2023" in text else "待核查",
                                      "维度": dim, "指标": metric + "（候选）", "原值或内容": "",
                                      "原单位": "", "换算值": "", "统一单位": "", "报告文件": filename,
                                      "PDF页码": page, "原文证据": excerpt, "GRI参照": "待核查",
                                      "提取方式": "全文关键词发现", "核验状态": "待人工复核，不能评分",
                                      "口径提醒": "关键词命中可能是目标、历史值或别家公司，不是已确认指标", "来源网址": ""})
                        break
                else:
                    continue
                break
        # 一个报告限制候选总量，避免大报告把界面和Excel淹没；优先保留各页前部证据。
        if len(found) >= 300:
            break
    keys = set(); unique = []
    for row in found:
        key = (row["PDF页码"], row["维度"], row["指标"], row["原文证据"])
        if key not in keys:
            keys.add(key); unique.append(row)
    return unique[:300]


def process_files(items: list[tuple[str, bytes]]) -> Batch:
    """All processing occurs in memory. Per-file failures do not abort other reports."""
    result = Batch()
    seen = set()
    for filename, content in items:
        basename = Path(filename).name
        digest = sha256(content).hexdigest()
        status = {"文件": basename, "大小（字节）": len(content), "SHA256": digest,
                  "识别企业": "", "PDF页数": "", "状态": "", "已核验记录": 0, "待复核候选": 0, "错误": ""}
        result.files.append(status)
        try:
            if digest in seen:
                status["状态"] = "重复文件，已跳过"
                continue
            seen.add(digest)
            if not content:
                raise ValueError("空文件")
            if not basename.lower().endswith(".pdf") or not content.startswith(b"%PDF-"):
                raise ValueError("不是有效PDF文件")
            if len(content) > 50 * 1024 * 1024:
                raise ValueError("单个文件超过50 MB限制")
            with fitz.open(stream=content, filetype="pdf") as doc:
                if doc.needs_pass:
                    raise ValueError("PDF受密码保护")
                if doc.page_count < 1:
                    raise ValueError("PDF没有页面")
                status["PDF页数"] = doc.page_count
                pages = [p.get_text() for p in doc]
            if not any(x.strip() for x in pages):
                raise ValueError("扫描图像PDF未检测到可提取文字；V1.0暂不支持OCR")
            company = KNOWN.get(digest, "未识别企业")
            status["识别企业"] = company
            if company != "未识别企业":
                try:
                    records = _known_rows(company, pages, basename)
                    result.records.extend(records)
                    status["已核验记录"] = len(records)
                except (ValueError, IndexError, KeyError) as err:
                    status["错误"] = f"已知报告专用规则未通过：{err}；仅提供待复核线索"[:250]
            candidates = _candidates(pages, basename, company)
            result.candidates.extend(candidates)
            status["待复核候选"] = len(candidates)
            status["状态"] = ("完成" if status["已核验记录"] or candidates else "可读，但未找到目标指标")
            if status["错误"] and status["状态"] == "完成":
                status["状态"] = "部分完成，专用规则未通过"
        except Exception as exc:
            status["状态"] = "处理失败"
            status["错误"] = f"{type(exc).__name__}: {exc}"[:250]
    return result


def excel_bytes(batch: Batch) -> bytes:
    wb = Workbook(); info = wb.active; info.title = "说明"
    for line in [
        ["ESG Intelligence Lab", "V1.0 云端展示版；2023汽车行业样本"],
        ["已核实数据", "仅原始PDF摘要与既有人工核验样本相同的文件；E指标不代表绩效排名"],
        ["待复核候选", "全文检索线索不代表正确数值；不可用于评分"],
        ["隐私", "上传文件由本程序在当前运行进程中处理；本程序不另存上传的PDF"],
    ]: info.append(line)
    for name, headers, rows in [("已核实数据", HEADERS, batch.records),
                                ("待复核候选", HEADERS, batch.candidates),
                                ("文件处理状态", list(batch.files[0]) if batch.files else ["文件", "状态"], batch.files)]:
        sheet = wb.create_sheet(name); sheet.append(headers)
        for row in rows:
            sheet.append([clean(str(row.get(h, ""))) if isinstance(row.get(h, ""), str) else row.get(h, "") for h in headers])
        sheet.freeze_panes = "A2"; sheet.auto_filter.ref = sheet.dimensions
    for sh in wb:
        for cell in sh[1]:
            cell.fill = PatternFill("solid", fgColor="173D45"); cell.font = Font(color="FFFFFF", bold=True)
        for col in sh.columns:
            sh.column_dimensions[col[0].column_letter].width = min(60, max(14, max(len(str(x.value or "")) for x in col)+2))
        for row in sh:
            for cell in row: cell.alignment = Alignment(vertical="top", wrap_text=True)
    out = BytesIO(); wb.save(out); return out.getvalue()
