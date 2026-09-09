import io
import os
import glob
import pickle
from datetime import datetime
from collections import defaultdict
from PIL import Image
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="Sleek-Industrial 進度記錄器", page_icon="📸", layout="centered")

def set_cell_border(cell, color="000000", sz="6", val="single"):
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        '</w:tcBorders>'
    )
    tcPr.append(tcBorders)

def get_saved_drafts():
    """獲取目前所有已暫存的 Job 草稿清單"""
    draft_files = glob.glob("draft_*.pkl")
    draft_map = {}
    for filepath in draft_files:
        try:
            with open(filepath, "rb") as f:
                data = pickle.load(f)
                j_title = data.get("job_title", filepath)
                draft_map[j_title] = filepath
        except Exception:
            pass
    return draft_map

st.title("📸 Sleek-Industrial 進度記錄器")

# --- 1. 初始化 Session State ---
if 'records' not in st.session_state:
    st.session_state['records'] = []

if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

if 'job_title' not in st.session_state:
    st.session_state['job_title'] = ""

# --- 2. 項目基本資料 & 草稿管理 (極簡對稱 UI) ---
st.subheader("📌 項目基本資料 & 草稿管理")

tab_save, tab_load = st.tabs(["💾 新建 / 暫存目前 Job", "📂 載入 / 管理舊草稿"])

# --- TAB 1: 新建與暫存 ---
with tab_save:
    with st.container(border=True):
        job_title = st.text_input(
            "Job Title / 工程項目名稱", 
            value=st.session_state['job_title'], 
            placeholder="例如: Regent Hotel F3 改善工程",
            key="main_job_title_input"
        )
        st.session_state['job_title'] = job_title

        safe_job = "".join(c for c in job_title if c.isalnum() or c in ('_', '-')).strip()
        DRAFT_FILE = f"draft_{safe_job}.pkl" if safe_job else "draft_default.pkl"

        if st.button("💾 暫存此 Job 草稿", use_container_width=True, type="primary"):
            if not safe_job:
                st.warning("請先輸入 Job Title，先可以進行專屬暫存！")
            elif len(st.session_state['records']) > 0:
                with open(DRAFT_FILE, "wb") as f:
                    pickle.dump({
                        "job_title": job_title,
                        "records": st.session_state['records']
                    }, f)
                st.success(f"已成功暫存【{job_title}】！")
                st.rerun()
            else:
                st.warning("目前未有記錄可以暫存。")

# --- TAB 2: 載入與管理草稿 ---
with tab_load:
    with st.container(border=True):
        saved_drafts = get_saved_drafts()
        
        if saved_drafts:
            selected_draft_title = st.selectbox(
                "揀選要繼續或刪除的 Job 草稿", 
                options=list(saved_drafts.keys()),
                key="draft_selectbox"
            )
            
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                if st.button("📂 載入此草稿", use_container_width=True, type="primary"):
                    target_file = saved_drafts[selected_draft_title]
                    if os.path.exists(target_file):
                        with open(target_file, "rb") as f:
                            data = pickle.load(f)
                            st.session_state['records'] = data.get("records", [])
                            st.session_state['job_title'] = data.get("job_title", selected_draft_title)
                        st.success(f"成功載入【{selected_draft_title}】！")
                        st.rerun()
            
            with btn_col2:
                if st.button("🗑️ 刪除此草稿", use_container_width=True):
                    target_file = saved_drafts[selected_draft_title]
                    if os.path.exists(target_file):
                        os.remove(target_file)
                        st.success(f"已刪除【{selected_draft_title}】草稿！")
                        st.rerun()
        else:
            st.info("目前伺服器內未有任何暫存草稿。")

st.divider()

def get_grouped_records():
    grouped = defaultdict(list)
    for idx, rec in enumerate(st.session_state['records']):
        f_val = rec['floor'] if rec['floor'] else "未註明樓層"
        r_val = rec['room'] if rec['room'] else ""
        cat_val = rec['category']
        key = (f_val, r_val, cat_val)
        grouped[key].append((idx, rec))
    return grouped

# --- 3. 新增現場記錄 ---
st.subheader("1️⃣ 新增現場記錄")

floor = st.text_input("樓層 (Floor) [選填]", placeholder="例如: B2 / G/F / 1F (可留空)")
room = st.text_input("房間 / 區域 (Room / Area) [選填]", placeholder="例如: Function Room A / 掣房 (可留空)")
category = st.selectbox("工程類別", ["AC", "FS", "P&D", "EL", "OTHER"])
remarks = st.text_area("工作備忘", placeholder="請輸入工作內容或備忘...")

photos = st.file_uploader(
    "拍攝或上傳現場相片 (可一次選取多張)", 
    type=['jpg', 'jpeg', 'png', 'heic'], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}"
)

col_btn1, col_btn2 = st.columns([1, 1])

with col_btn1:
    if st.button("➕ 新增到今日清單", use_container_width=True, type="primary"):
        if photos:
            success_count = 0
            f_val = floor.strip() if floor else ""
            r_val = room.strip() if room else ""
            rem_val = remarks.strip() if remarks else ""
            
            for p in photos:
                try:
                    img = Image.open(p)
                    if img.mode in ('RGBA', 'LA', 'P'):
                        img = img.convert('RGB')
                    
                    photo_bytes = io.BytesIO()
                    img.save(photo_bytes, format='JPEG', quality=90)
                    photo_bytes.seek(0)
                    photo_bytes.name = "photo.jpg"
                    
                    st.session_state['records'].append({
                        "floor": f_val,
                        "room": r_val,
                        "category": category,
                        "remarks": rem_val,
                        "photo": photo_bytes
                    })
                    success_count += 1
                except Exception as e:
                    st.error(f"相片 {p.name} 處理失敗: {e}")
            
            if success_count > 0:
                st.session_state['uploader_key'] += 1
                st.success(f"成功新增 {success_count} 張相片！")
                st.rerun()
        else:
            st.warning("請上傳或拍攝至少一張現場相片！")

with col_btn2:
    if st.button("🧹 清空相片 / 重置", use_container_width=True):
        st.session_state['uploader_key'] += 1
        st.rerun()

st.divider()

# --- 4. 顯示已記錄清單 ---
st.subheader("📋 今日已記錄項目 (已自動分類)")
grouped_data = get_grouped_records()

if len(grouped_data) > 0:
    for (group_floor, group_room, group_cat), items in grouped_data.items():
        title_parts = [f"🏢 樓層: {group_floor}"]
        if group_room:
            title_parts.append(f"📍 區域: {group_room}")
        title_parts.append(f"🔧 類別: {group_cat} (共 {len(items)} 張)")
        
        title_str = " | ".join(title_parts)
        
        with st.expander(title_str, expanded=True):
            cols = st.columns(3)
            for sub_idx, (orig_idx, rec) in enumerate(items):
                with cols[sub_idx % 3]:
                    rec['photo'].seek(0)
                    st.image(rec['photo'], use_container_width=True)
                    if rec['remarks']:
                        st.write(f"📝 {rec['remarks']}")
                    if st.button("🗑️ 刪除", key=f"del_{orig_idx}"):
                        st.session_state['records'].pop(orig_idx)
                        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🗑️ 清空所有記錄"):
        st.session_state['records'] = []
        safe_job_name = "".join(c for c in st.session_state['job_title'] if c.isalnum() or c in ('_', '-')).strip()
        curr_draft = f"draft_{safe_job_name}.pkl" if safe_job_name else "draft_default.pkl"
        if os.path.exists(curr_draft):
            os.remove(curr_draft)
        st.rerun()
else:
    st.info("暫時未有記錄，請喺上面新增。")

st.divider()

# --- 5. 一鍵生成 Word 報告 ---
st.subheader("3️⃣ 匯出報告")
if st.button("📥 一鍵生成 Word 報告", type="primary", use_container_width=True):
    if len(st.session_state['records']) == 0:
        st.warning("請先新增至少一個記錄先可以出 Report！")
    else:
        doc = Document()

        for section in doc.sections:
            section.top_margin = Inches(0.4)
            section.bottom_margin = Inches(0.4)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

        display_title = st.session_state['job_title'].strip() if st.session_state['job_title'].strip() != "" else "Unnamed Project"

        header_p = doc.add_paragraph()
        header_p.paragraph_format.space_after = Pt(6)
        r_title = header_p.add_run(f"Job Title: {display_title}")
        r_title.bold = True
        r_title.font.size = Pt(12)

        r_date = header_p.add_run(f"  |  Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        r_date.font.size = Pt(10)
        r_date.font.color.rgb = RGBColor(100, 100, 100)

        cols_per_row = 2
        global_card_count = 0
        group_index = 0

        for (group_floor, group_room, group_cat), items in grouped_data.items():
            if group_index > 0:
                doc.add_page_break()
                global_card_count = 0
            
            group_index += 1

            group_p = doc.add_paragraph()
            group_p.paragraph_format.space_before = Pt(4)
            group_p.paragraph_format.space_after = Pt(4)
            
            header_text = f"【 樓層: {group_floor}"
            if group_room:
                header_text += f"  |  區域: {group_room}"
            header_text += f"  |  工程類別: {group_cat} 】"

            r_grp = group_p.add_run(header_text)
            r_grp.bold = True
            r_grp.font.size = Pt(11)
            r_grp.font.color.rgb = RGBColor(0, 51, 102)

            total_items = len(items)

            for r_idx in range(0, total_items, cols_per_row):
                if global_card_count > 0 and global_card_count % 6 == 0:
                    doc.add_page_break()

                table = doc.add_table(rows=1, cols=2)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                hdr_cells = table.rows[0].cells

                for c_idx in range(cols_per_row):
                    item_sub_idx = r_idx + c_idx
                    cell = hdr_cells[c_idx]

                    if item_sub_idx < total_items:
                        orig_idx, rec = items[item_sub_idx]
                        global_card_count += 1

                        set_cell_border(cell, color="000000", sz="6", val="single")

                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_before = Pt(2)
                        p.paragraph_format.space_after = Pt(2)

                        try:
                            rec['photo'].seek(0)
                            p.add_run().add_picture(rec['photo'], height=Inches(2.2))
                        except Exception:
                            p.add_run("[相片載入失敗]")

                        if rec['remarks']:
                            p_txt = cell.add_paragraph()
                            p_txt.paragraph_format.space_before = Pt(2)
                            p_txt.paragraph_format.space_after = Pt(4)
                            p_txt.paragraph_format.line_spacing = 1.0

                            r_rem = p_txt.add_run(f"備忘: {rec['remarks']}")
                            r_rem.font.size = Pt(8.5)
                            r_rem.font.color.rgb = RGBColor(50, 50, 50)

                    else:
                        set_cell_border(cell, color="FFFFFF", sz="0", val="none")

                spacer = doc.add_paragraph()
                spacer.paragraph_format.space_before = Pt(0)
                spacer.paragraph_format.space_after = Pt(4)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        safe_job_title = "".join(c for c in display_title if c.isalnum() or c in (' ', '_', '-')).strip().replace(' ', '_')

        st.download_button(
            label="💾 點擊下載 Word 報告 (.docx)",
            data=buffer,
            file_name=f"Report_{safe_job_title}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 14px;'>"
    "🛠️ <b>Design by nikki 💅</b>"
    "</div>", 
    unsafe_allow_html=True
)
