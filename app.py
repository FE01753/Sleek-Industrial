import io
import json
import os
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

# -----------------------------------------------------------------------------
# 1. 頁面基本配置
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="E&M Progress Logger",
    page_icon="📸",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# 草稿暫存檔路徑
DRAFT_FILE = "site_drafts.json"


# -----------------------------------------------------------------------------
# 2. 草稿管理讀寫函數
# -----------------------------------------------------------------------------
def load_all_drafts():
    if os.path.exists(DRAFT_FILE):
        try:
            with open(DRAFT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_draft(job_title, records):
    drafts = load_all_drafts()
    serializable_records = []
    for r in records:
        serializable_records.append(
            {
                "floor": r["floor"],
                "area": r["area"],
                "category": r["category"],
                "remark": r["remark"],
                "image_bytes": r["image_bytes"].getvalue(),
            }
        )
    drafts[job_title] = serializable_records
    with open(DRAFT_FILE, "w", encoding="utf-8") as f:
        json.dump(drafts, f, ensure_ascii=False, indent=2)


def delete_draft(job_title):
    drafts = load_all_drafts()
    if job_title in drafts:
        del drafts[job_title]
        with open(DRAFT_FILE, "w", encoding="utf-8") as f:
            json.dump(drafts, f, ensure_ascii=False, indent=2)


# -----------------------------------------------------------------------------
# 3. Word 報告生成器 (.docx)
# -----------------------------------------------------------------------------
def generate_word_report(job_title, records):
    doc = Document()

    # 設定標準窄邊框 (0.5 吋)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # 標頭 Title
    p_title = doc.add_paragraph()
    run_title = p_title.add_run(f"SITE PHOTO RECORD: {job_title}")
    run_title.font.name = "Arial"
    run_title.font.size = Pt(16)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(15, 23, 42)

    # 分類數據 (按 樓層 > 區域 > 類別)
    grouped = {}
    for r in records:
        key = f"{r['floor']} - {r['area']}"
        if key not in grouped:
            grouped[key] = []
        grouped[key].append(r)

    # 寫入內容
    for loc, items in grouped.items():
        p_loc = doc.add_paragraph()
        p_loc.paragraph_format.space_before = Pt(12)
        p_loc.paragraph_format.space_after = Pt(4)
        run_loc = p_loc.add_run(f"📍 位置: {loc}")
        run_loc.font.name = "Arial"
        run_loc.font.size = Pt(12)
        run_loc.font.bold = True
        run_loc.font.color.rgb = RGBColor(37, 99, 235)

        # 兩欄排版表格
        table = doc.add_table(rows=0, cols=2)
        table.autofit = False

        # 每兩張相排一列
        for i in range(0, len(items), 2):
            row_cells = table.add_row().cells

            for idx, item in enumerate(items[i : i + 2]):
                cell = row_cells[idx]
                cell.width = Inches(3.6)
                p_cell = cell.paragraphs[0]
                p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # 直接插入原圖 (無壓縮)
                img_stream = io.BytesIO(item["image_bytes"].getvalue())
                p_cell.add_run().add_picture(img_stream, width=Inches(3.4))

                # 備註說明
                p_desc = cell.add_paragraph()
                p_desc.alignment = WD_ALIGN_PARAGRAPH.LEFT
                desc_text = f"[{item['category']}] {item['remark']}" if item["remark"] else f"[{item['category']}]"
                run_desc = p_desc.add_run(desc_text)
                run_desc.font.name = "Arial"
                run_desc.font.size = Pt(9.5)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


# -----------------------------------------------------------------------------
# 4. 主程式 Session 初始化
# -----------------------------------------------------------------------------
if "records" not in st.session_state:
    st.session_state.records = []
if "job_title" not in st.session_state:
    st.session_state.job_title = ""

st.title("📸 E&M 現場進度記錄器")

# -----------------------------------------------------------------------------
# 5. 頂部 Tab：草稿暫存與載入
# -----------------------------------------------------------------------------
tab1, tab2 = st.tabs(["💾 新建 / 暫存", "📂 載入 / 管理草稿"])

with tab1:
    job_input = st.text_input("Job Title / 工程項目名稱", value=st.session_state.job_title, placeholder="例: Regent Hotel F3 改善工程")
    st.session_state.job_title = job_input

    if st.button("💾 暫存目前紀錄", use_container_width=True):
        if not st.session_state.job_title:
            st.error("請先輸入 Job Title 才可進行暫存！")
        elif not st.session_state.records:
            st.warning("目前沒有任何相片紀錄可供暫存。")
        else:
            save_draft(st.session_state.job_title, st.session_state.records)
            st.success(f"成功暫存草稿：{st.session_state.job_title}")

with tab2:
    all_drafts = load_all_drafts()
    if not all_drafts:
        st.info("目前沒有已儲存的草稿。")
    else:
        selected_draft = st.selectbox("選擇要載入的草稿", list(all_drafts.keys()))
        col_load, col_del = st.columns(2)

        with col_load:
            if st.button("📂 載入此草稿", use_container_width=True):
                draft_data = all_drafts[selected_draft]
                st.session_state.job_title = selected_draft
                st.session_state.records = []
                for item in draft_data:
                    st.session_state.records.append(
                        {
                            "floor": item["floor"],
                            "area": item["area"],
                            "category": item["category"],
                            "remark": item["remark"],
                            "image_bytes": io.BytesIO(bytes(item["image_bytes"])),
                        }
                    )
                st.success(f"已成功載入草稿：{selected_draft}")
                st.rerun()

        with col_del:
            if st.button("🗑️ 刪除此草稿", type="primary", use_container_width=True):
                delete_draft(selected_draft)
                st.success(f"已刪除草稿：{selected_draft}")
                st.rerun()

st.divider()

# -----------------------------------------------------------------------------
# 6. 現場輸入與相片上傳區
# -----------------------------------------------------------------------------
st.subheader("➕ 新增現場紀錄")

col1, col2 = st.columns(2)
with col1:
    floor = st.text_input("樓層 (Floor)", value="1F")
with col2:
    area = st.text_input("區域/位置 (Area)", value="Function Room")

category = st.selectbox("工程類別 (Category)", ["AC", "FS", "P&D", "EL", "OTHER"])
remark = st.text_input("備註說明 (Remark - 可不填)", placeholder="例: 完成管道鋪設")

uploaded_files = st.file_uploader(
    "上傳現場相片 (支援多張相片/即場拍攝)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if st.button("➕ 新增到今日清單", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("請先選取或拍攝至少一張相片！")
    else:
        for file in uploaded_files:
            # 直接讀取原圖 Bytes
            img_bytes = io.BytesIO(file.read())
            st.session_state.records.append(
                {
                    "floor": floor,
                    "area": area,
                    "category": category,
                    "remark": remark,
                    "image_bytes": img_bytes,
                }
            )
        st.success(f"成功新增 {len(uploaded_files)} 張相片！")
        st.rerun()

st.divider()

# -----------------------------------------------------------------------------
# 7. 已新增紀錄展示與刪除整理
# -----------------------------------------------------------------------------
st.subheader(f"📋 已記錄列表 (共 {len(st.session_state.records)} 張)")

if not st.session_state.records:
    st.info("目前尚未新增任何紀錄。")
else:
    for idx, rec in enumerate(st.session_state.records):
        with st.expander(f"#{idx+1} | [{rec['floor']} - {rec['area']}] [{rec['category']}] {rec['remark']}"):
            st.image(rec["image_bytes"], use_column_width=True)
            if st.button(f"🗑️ 刪除此張相片", key=f"del_{idx}"):
                st.session_state.records.pop(idx)
                st.rerun()

st.divider()

# -----------------------------------------------------------------------------
# 8. 匯出 Word 報告區
# -----------------------------------------------------------------------------
if st.session_state.records:
    if st.button("📥 一鍵生成 Word 報告 (.docx)", type="primary", use_container_width=True):
        if not st.session_state.job_title:
            st.error("請在頂部輸入 Job Title 後再生成報告！")
        else:
            with st.spinner("正在排版並生成 Word 報告..."):
                doc_file = generate_word_report(st.session_state.job_title, st.session_state.records)
                st.download_button(
                    label="💾 點擊下載 Word 報告 (.docx)",
                    data=doc_file,
                    file_name=f"{st.session_state.job_title}_Photo_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
