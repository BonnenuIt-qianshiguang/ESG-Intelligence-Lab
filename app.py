"""ESG Intelligence Lab — V1.0 云端展示版。运行：streamlit run app.py"""
from __future__ import annotations

import pandas as pd
import streamlit as st
from pathlib import Path
from openpyxl import load_workbook

from esg_v01_backend import Batch, HEADERS, excel_bytes, process_files

st.set_page_config(page_title="ESG Intelligence Lab", page_icon="🌿", layout="wide")
st.markdown("""<style>
  .block-container {max-width: 1180px; padding-top: 2.1rem}
  h1,h2,h3 {color: #173D45}
  div[data-testid='stMetric'] {background:#edf4f1; border:1px solid #d3e3dc;
    border-radius:12px; padding:16px}
  .eyebrow {color:#628277;letter-spacing:.13em;font-size:.76rem;font-weight:700}
  .subtitle {color:#526a67; font-size:1.1rem}
</style>""", unsafe_allow_html=True)

st.markdown('<div class="eyebrow">2023 汽车行业 · 课程研究 / V1.0</div>', unsafe_allow_html=True)
st.title("ESG Intelligence Lab")
st.markdown('<div class="subtitle">多企业ESG信息自动提取与智能评价系统</div>', unsafe_allow_html=True)
st.caption("保存样本展示、PDF文本提取及Excel下载；已知原报告使用专用规则，其他PDF仅生成待人工复核线索。评分与多模态尚未实现。")

tab_overview, tab_upload, tab_data, tab_compare, tab_score, tab_media = st.tabs(
    ["项目总览", "批量PDF上传与提取", "结构化数据", "企业比较", "评价与评分", "多模态"]
)

@st.cache_data(show_spinner=False)
def verified_sample():
    path = Path(__file__).with_name("已核实样本_15条.xlsx")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        def rows(sheet):
            it = wb[sheet].values
            headers = next(it)
            return [dict(zip(headers, row)) for row in it]
        return Batch(records=rows("已核实数据"), files=rows("文件处理状态"))
    finally:
        wb.close()

with tab_overview:
    st.write("研究同一行业三家企业2023年的报告，展示ESG数据如何从原始PDF走到可核查的结构化记录。")
    c1, c2, c3 = st.columns(3)
    c1.metric("已核实样本企业", "3 家")
    c2.metric("已核实原始记录", "15 条")
    c3.metric("报告年度", "2023")
    st.info("九条E指标、三条S人数、三条G制度描述采用已知报告的专用规则。上传其他PDF时，只生成待复核的全文检索线索。")
    if st.button("载入15条保存样本（无需上传）"):
        sample = verified_sample()
        st.session_state["esg_batch"] = sample
        st.session_state["esg_excel"] = excel_bytes(sample)
        st.session_state["esg_sample_mode"] = True
    st.caption("样本结果来自此前真实报告的已核实提取记录；此按钮只加载保存的数据。上传 PDF 并点击“开始真实提取”才会重新运行程序。")

with tab_upload:
    st.subheader("上传企业报告")
    uploaded = st.file_uploader("选择一份或多份PDF", type=["pdf"], accept_multiple_files=True,
                                help="当前每个PDF限50 MB；文件只在本次运行内处理。")
    if uploaded:
        st.dataframe(pd.DataFrame([{"文件": f.name, "大小（MB）": round(f.size/1024/1024, 2)}
                                   for f in uploaded]), hide_index=True, width="stretch")
    if st.button("开始真实提取", type="primary", disabled=not uploaded):
        try:
            with st.spinner("正在逐页读取PDF并检索证据…"):
                batch = process_files([(f.name, f.getvalue()) for f in uploaded])
                output = excel_bytes(batch)
            st.session_state["esg_batch"] = batch
            st.session_state["esg_excel"] = output
            st.session_state["esg_sample_mode"] = False
            failures = sum(x["状态"] == "处理失败" for x in batch.files)
            if failures:
                st.warning(f"已检查 {len(batch.files)} 份文件，其中 {failures} 份处理失败。请查看下方状态表的错误原因。")
            else:
                st.success("文件处理结束。下方可查看结果与证据线索。")
        except Exception as exc:
            st.error(f"本次处理未完成：{type(exc).__name__}: {exc}")
    if "esg_batch" in st.session_state:
        batch = st.session_state["esg_batch"]
        if st.session_state.get("esg_sample_mode"):
            st.info("当前显示已保存的已核实样本结果；尚未在本次会话重新解析 PDF。")
        status_df = pd.DataFrame(batch.files)
        status_df["PDF页数"] = status_df["PDF页数"].fillna("").astype(str)
        st.dataframe(status_df, hide_index=True, width="stretch")
        label = "保存的既往已核实记录" if st.session_state.get("esg_sample_mode") else "匹配既有原报告摘要的记录"
        st.write(f"{label}：{len(batch.records)}；自动发现但待复核线索：{len(batch.candidates)}。")
        if batch.records:
            st.dataframe(pd.DataFrame(batch.records)[["企业", "指标", "原值或内容", "原单位", "PDF页码", "原文证据", "核验状态"]],
                         hide_index=True, width="stretch")
        st.download_button("下载本次结果 Excel", st.session_state["esg_excel"],
                           file_name=("ESG_15条保存样本.xlsx" if st.session_state.get("esg_sample_mode") else "ESG_本次上传提取结果.xlsx"),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab_data:
    batch = st.session_state.get("esg_batch")
    if batch is None:
        st.info("可先在“项目总览”载入15条保存样本，或上传PDF并运行提取。")
    else:
        choice = st.radio("查看数据", ["已核实数据", "待复核线索"], horizontal=True)
        rows = batch.records if choice == "已核实数据" else batch.candidates
        if rows:
            df = pd.DataFrame(rows, columns=HEADERS)
            for col in ("年度", "换算值"):
                df[col] = df[col].astype(str)
            a, b, c = st.columns(3)
            companies = a.multiselect("企业", sorted(df["企业"].unique()))
            dims = b.multiselect("维度", sorted(df["维度"].unique()))
            indicators = c.multiselect("指标", sorted(df["指标"].unique()))
            if companies: df = df[df["企业"].isin(companies)]
            if dims: df = df[df["维度"].isin(dims)]
            if indicators: df = df[df["指标"].isin(indicators)]
            st.dataframe(df, hide_index=True, width="stretch")
        else:
            st.info("当前没有此类记录。")

with tab_compare:
    batch = st.session_state.get("esg_batch")
    if batch and batch.records:
        df = pd.DataFrame(batch.records)
        coverage = df.groupby(["企业", "维度"], as_index=False).size().rename(columns={"size":"已核实记录数"})
        st.bar_chart(coverage.pivot(index="企业", columns="维度", values="已核实记录数").fillna(0))
        st.dataframe(coverage, hide_index=True, width="stretch")
        st.caption("图表只比较已核实记录覆盖情况，不按绝对能源和排放总量评判企业绩效。")
    else:
        st.info("可先载入保存样本，或上传PDF运行提取，再查看数据覆盖图。")

with tab_score:
    st.warning("自定义披露评分仍在开发。现有少量E/S/G披露记录不足以形成稳健的绩效评级；此页不生成虚假分数。")
    st.write("后续将公布指标、权重、缺失值处理和敏感性检查，再计算自定义披露评价。")

with tab_media:
    st.info("多模态提取开发中。尚无已核实视频、音频或图像提取结果。")
