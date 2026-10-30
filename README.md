# ESG Intelligence Lab V1.0 — Streamlit Community Cloud 发布包

以原在线部署预备包为基线进行最小修复。项目有 **15 条已保存的历史人工核实样本**；这不是本次自动重新抓取的结果。真实 PDF 上传调用后端；仅当上传文件 SHA256 与三份已知原报告完全匹配，才启用专用规则生成既有核验口径的 5 条记录/企业。其他 PDF 仅生成需要人工复核的关键词候选，不承诺自动数值抽取。暂无评级和多模态。

## 部署文件
- 入口：`app.py`
- 依赖：`requirements.txt`，部署 Python **3.12**
- 必须同时上传：`esg_v01_backend.py`、`esg_batch_pilot.py`、`已核实样本_15条.xlsx`
- 配置：`.streamlit/config.toml`
- 基础测试：`python smoke_test.py`（不需要三份原报告）
- 原报告集成测试：`python integration_test_original_pdfs.py`（需要 `reports/` 中三份特定文件，本包没有它们）

## 发布与核验
请将 **压缩包解压后的文件**（不是 ZIP 本身）上传 GitHub 仓库根目录；隐藏文件 `.gitignore`、目录 `.streamlit` 也要包括。登录 share.streamlit.io 后选 GitHub 仓库、main 分支和 `app.py`，Advanced settings 选 Python 3.12；无密钥需求。

部署完成后，需要实测：①外部浏览器打开公开网址；②载入保存样本查看15条；③下载样本 Excel；④上传真实可公开分享的文本型 PDF，查看待复核线索并下载 Excel；⑤如需证明历史15条原报告重跑，则再提供三份与已存 SHA256 匹配的原始 PDF 并运行集成测试。

只上传无保密要求、可公开分享的报告。上传文件由云端服务接收并在应用进程内处理；应用不主动将 PDF 保存到仓库。公开仓库将公开代码和保存样本。
