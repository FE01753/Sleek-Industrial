import io
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

# --- 頁面配置 ---
st.set_page_config(page_title="Sleek-Industrial 進度記錄器", page_icon="📸", layout="centered")

# --- XML 輔助函式：設置 Word 表格 Cell 黑/灰實線邊框 ---
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

st.title("📸 Sleek-Industrial 進度記錄器")

# --- 0. 設定 Job Title (工程項目名稱) ---
st.subheader("📌 項目基本資料")
job_title = st.text_input("Job Title / 工程項目名稱", value="", placeholder="例如: Regent Hotel F3 改善工程")

st.divider()

# --- 1. 初始化 Session State (暫存記錄 + Uploader Key) ---
if 'records' not in st.session_state:
    st.session_state['records'] = []

if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

# 輔助函式：將 records 依 (Floor, Category) 自動分組
def get_grouped_records():
    grouped = defaultdict(list)
    for idx, rec in enumerate(st.session_state['records']):
        f_val = rec['floor'] if rec['floor'] else "未註明樓層"
        cat_val = rec['category']
        key = (f_val, cat_val)
        grouped[key].append((idx, rec))
    return grouped

# --- 2. 新增現場記錄 ---
st.subheader("1️⃣ 新增現場記錄")

floor = st.text_input("樓層 (Floor) [選填]", placeholder="例如: B2 / G/F / 1F (可留空)")
room = st.text_input("房間 / 區域 (Room / Area) [選填]", placeholder="例如: Function Room A / 掣房 (可留空)")
category = st.selectbox("工程類別", ["AC", "FS", "P&D", "EL", "OTHER"])
remarks = st.text_area("工作備忘", placeholder="請輸入工作內容或備忘...")

# 📸 使用動態 Key 實現自動清空與手動重置
photos = st.file_uploader(
    "拍攝或上傳現場相片 (可一次選取多張)", 
    type=['jpg', 'jpeg', 'png', 'heic'], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}"
)

# 橫向排列按鈕
col_btn1, col_btn2 = st.columns([1, 1])

with col_btn1:
    if st.button("➕ 新增到今日清單", use_container_width=True):
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
                # 成功新增後自動更新 Key 以重置相片欄位
                st.session_state['uploader_key'] += 1
                st.success(f"成功新增 {success_count} 張相片！相片區域已自動清空。")
                st.rerun()
        else:
            st.warning("請上傳或拍攝至少一張現場相片！")

with col_btn2:
    if st.button("🧹 清空相片 / 重置", use_container_width=True):
        st.session_state['uploader_key'] += 1
        st.rerun()

st.divider()

# --- 3. 顯示已記錄清單 (自動依樓層/類別 Grouping) ---
st.subheader("📋 今日已記錄項目 (已自動分類)")
grouped_data = get_grouped_records()

if len(grouped_data) > 0:
    for (group_floor, group_cat), items in grouped_data.items():
        with st.expander(f"🏢 樓層: {group_floor} ({group_cat}) — 共 {len(items)} 張相片", expanded=True):
            cols = st.columns(3)
            for sub_idx, (orig_idx, rec) in enumerate(items):
                with cols[sub_idx % 3]:
                    rec['photo'].seek(0)
                    st.image(rec['photo'], use_container_width=True)
                    if rec['room']:
                        st.caption(f"📍 區域: {rec['room']}")
                    if rec['remarks']:
                        st.write(f"📝 {rec['remarks']}")
                    if st.button("🗑️ 刪除", key=f"del_{orig_idx}"):
                        st.session_state['records'].pop(orig_idx)
                        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🗑️ 清空所有記錄"):
        st.session_state['records'] = []
        st.rerun()
else:
    st.info("暫時未有記錄，請喺上面新增。")

st.divider()

# --- 4. 一鍵生成 Word 報告 (自動分組 + 2x3 Grid + 黑色相框) ---
st.subheader("3️⃣ 匯出報告")
if st.button("📥 一鍵生成 Word 報告"):
    if len(st.session_state['records']) == 0:
        st.warning("請先新增至少一個記錄先可以出 Report！")
    else:
        doc = Document()

        # 窄邊距設定 (Top/Bottom/Left/Right)
        for section in doc.sections:
            section.top_margin = Inches(0.4)
            section.bottom_margin = Inches(0.4)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

        display_title = job_title.strip() if job_title.strip() != "" else "Unnamed Project"

        # 精簡 Header
        header_p = doc.add_paragraph()
        header_p.paragraph_format.space_after = Pt(6)
        r_title = header_p.add_run(f"Job Title: {display_title}")
        r_title.bold = True
        r_title.font.size = Pt(12)

        r_date = header_p.add_run(f"  |  Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        r_date.font.size = Pt(10)
        r_date.font.color.rgb = RGBColor(100, 100, 100)

        # 依 Group 生成內容
        cols_per_row = 2
        global_card_count = 0  # 控制 6 張自動分頁
        group_index = 0        # 記錄分類組數

        for (group_floor, group_cat), items in grouped_data.items():
            # 第二個記錄項目 (Group) 開始自動換新頁，避免標題卡在上一頁底部
            if group_index > 0:
                doc.add_page_break()
                global_card_count = 0  # 換頁後計數重置
            
            group_index += 1

            group_p = doc.add_paragraph()
            group_p.paragraph_format.space_before = Pt(4)
            group_p.paragraph_format.space_after = Pt(4)
            r_grp = group_p.add_run(f"【 樓層: {group_floor}  |  工程類別: {group_cat} 】")
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

                        # 設置外邊框 (相框效果)
                        set_cell_border(cell, color="000000", sz="6", val="single")

                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_before = Pt(2)
                        p.paragraph_format.space_after = Pt(2)

                        # 1. 插入相片：固定高度 (2.2 Inches)，令直相與橫相高度一致
                        try:
                            rec['photo'].seek(0)
                            p.add_run().add_picture(rec['photo'], height=Inches(2.2))
                        except Exception:
                            p.add_run("[相片載入失敗]")

                        # 2. 插入文字 (無填寫則完全隱藏)
                        has_room = bool(rec['room'])
                        has_remarks = bool(rec['remarks'])

                        if has_room or has_remarks:
                            p_txt = cell.add_paragraph()
                            p_txt.paragraph_format.space_before = Pt(2)
                            p_txt.paragraph_format.space_after = Pt(4)
                            p_txt.paragraph_format.line_spacing = 1.0

                            if has_room:
                                r_room = p_txt.add_run(f"區域: {rec['room']}")
                                r_room.bold = True
                                r_room.font.size = Pt(9)
                                if has_remarks:
                                    p_txt.add_run("\n")

                            if has_remarks:
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
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# --- App 底部專屬水印 (Footer) ---
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 14px;'>"
    "🛠️ <b>Design by nikki 💅</b>"
    "</div>", 
    unsafe_allow_html=True
)
