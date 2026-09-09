import io
from datetime import datetime
from PIL import Image
import streamlit as st
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor


# XML 輔助函式：設置表格黑/灰實線邊框
def set_cell_border(cell, color='000000', sz='4', val='single'):
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


st.title('📸 Sleek-Industrial 進度記錄器')

# --- 0. 設定 Job Title (工程項目名稱) ---
st.subheader('📌 項目基本資料')
job_title = st.text_input(
    'Job Title / 工程項目名稱',
    value='',
    placeholder='例如: Regent Hotel F3 改善工程',
)

st.divider()

# 1. 初始化 Session State 黎暫存記錄
if 'records' not in st.session_state:
  st.session_state['records'] = []

# --- 輸入當前記錄 ---
st.subheader('1️⃣ 新增現場記錄')

floor = st.text_input(
    '樓層 (Floor) [選填]', placeholder='例如: B2 / G/F / 3/F (可留空)'
)
room = st.text_input(
    '房間 / 區域 (Room / Area) [選填]',
    placeholder='例如: Function Room A / 掣房 (可留空)',
)
category = st.selectbox('工程類別', ['AC', 'FS', 'P&D', 'EL', 'OTHER'])
remarks = st.text_area('工作備忘', placeholder='請輸入工作內容或備忘...')

# 拍照或上傳相片
photo = st.file_uploader(
    '拍攝或上傳現場相片', type=['jpg', 'jpeg', 'png', 'heic']
)

# 加入暫存清單按鈕
if st.button('➕ 新增到今日清單'):
  if photo is not None:
    try:
      img = Image.open(photo)
      if img.mode in ('RGBA', 'LA', 'P'):
        img = img.convert('RGB')

      photo_bytes = io.BytesIO()
      img.save(photo_bytes, format='JPEG', quality=90)
      photo_bytes.seek(0)
      photo_bytes.name = 'photo.jpg'

      f_val = floor.strip() if floor else ''
      r_val = room.strip() if room else ''

      st.session_state['records'].append({
          'floor': f_val,
          'room': r_val,
          'category': category,
          'remarks': remarks.strip(),
          'photo': photo_bytes,
      })
      st.success('成功新增現場記錄！')
    except Exception as e:
      st.error(f'相片處理失敗，請嘗試另一張相片。錯誤詳情: {e}')
  else:
    st.warning('請上傳或拍攝現場相片！')

st.divider()

# --- 2. 顯示已記錄嘅清單 ---
st.subheader('📋 今日已記錄項目')
if len(st.session_state['records']) > 0:
  for i, rec in enumerate(st.session_state['records']):
    loc_parts = []
    if rec['floor']:
      loc_parts.append(f"樓層: {rec['floor']}")
    if rec['room']:
      loc_parts.append(f"區域: {rec['room']}")
    loc_str = ' - '.join(loc_parts) if loc_parts else '未註明位置'

    with st.expander(f"項目 #{i+1}: {loc_str} ({rec['category']})"):
      st.write(f"備忘：{rec['remarks']}")
      rec['photo'].seek(0)
      st.image(rec['photo'], width=300)

      if st.button(f'刪除此項 #{i+1}', key=f'del_{i}'):
        st.session_state['records'].pop(i)
        st.rerun()

  if st.button('🗑️ 清空所有記錄'):
    st.session_state['records'] = []
    st.rerun()
else:
  st.info('暫時未有記錄，請喺上面新增。')

st.divider()

# --- 3. 最後一鍵生成 REPORT ---
st.subheader('3️⃣ 匯出報告')
if st.button('📥 一鍵生成 Word 報告'):
  if len(st.session_state['records']) == 0:
    st.warning('請先新增至少一個記錄先可以出 Report！')
  else:
    doc = Document()

    # 設置頁面邊距為窄邊距 (Top/Bottom/Left/Right 均為 0.5 吋)，確保 2x3 完美落入一頁
    sections = doc.sections
    for section in sections:
      section.top_margin = Inches(0.4)
      section.bottom_margin = Inches(0.4)
      section.left_margin = Inches(0.5)
      section.right_margin = Inches(0.5)

    display_title = (
        job_title.strip() if job_title.strip() != '' else 'Unnamed Project'
    )

    # 頂部精簡 Header (取消大標題與總項目數)
    header_p = doc.add_paragraph()
    header_p.paragraph_format.space_after = Pt(6)
    r_title = header_p.add_run(f'Job Title: {display_title}')
    r_title.bold = True
    r_title.font.size = Pt(12)

    r_date = header_p.add_run(
        f"  |  Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    )
    r_date.font.size = Pt(10)
    r_date.font.color.rgb = RGBColor(100, 100, 100)

    records = st.session_state['records']
    total_records = len(records)

    # 以 2 欄進行佈局，每頁最多 6 個項目 (3 行 x 2 欄)
    cols_per_row = 2
    for r_idx in range(0, total_records, cols_per_row):
      # 每 6 個項目自動分頁 (0, 6, 12...)
      if r_idx > 0 and r_idx % 6 == 0:
        doc.add_page_break()

      table = doc.add_table(rows=1, cols=2)
      table.alignment = WD_TABLE_ALIGNMENT.CENTER
      hdr_cells = table.rows[0].cells

      for c_idx in range(cols_per_row):
        item_idx = r_idx + c_idx
        cell = hdr_cells[c_idx]

        if item_idx < total_records:
          rec = records[item_idx]

          # 設定外框 (000000 為黑色邊框)
          set_cell_border(cell, color='000000', sz='6', val='single')

          p = cell.paragraphs[0]
          p.alignment = WD_ALIGN_PARAGRAPH.CENTER
          p.paragraph_format.space_before = Pt(2)
          p.paragraph_format.space_after = Pt(2)

          # 1. 插入圖片 (縮小寬度至 ~3.3 吋以適應 2 欄排版)
          try:
            rec['photo'].seek(0)
            p.add_run().add_picture(rec['photo'], width=Inches(3.2))
          except Exception:
            p.add_run('[相片載入失敗]')

          # 2. 插入文字說明 (單行打橫排版慳位)
          p_txt = cell.add_paragraph()
          p_txt.paragraph_format.space_before = Pt(2)
          p_txt.paragraph_format.space_after = Pt(4)
          p_txt.paragraph_format.line_spacing = 1.0

          loc_parts = []
          if rec['floor']:
            loc_parts.append(rec['floor'])
          if rec['room']:
            loc_parts.append(rec['room'])
          loc_str = ' / '.join(loc_parts) if loc_parts else '未註明位置'

          # 項目編號 + 位置 + 工程類別
          r_head = p_txt.add_run(
              f"#{item_idx+1} [{rec['category']}] {loc_str}\n"
          )
          r_head.bold = True
          r_head.font.size = Pt(9.5)
          r_head.font.color.rgb = RGBColor(0, 51, 102)

          # 備忘內容
          remark_text = (
              rec['remarks'] if rec['remarks'] else '無工作備忘'
          )
          r_rem = p_txt.add_run(f'備忘: {remark_text}')
          r_rem.font.size = Pt(8.5)
          r_rem.font.color.rgb = RGBColor(50, 50, 50)

        else:
          # 隱藏空白格邊框
          set_cell_border(cell, color='FFFFFF', sz='0', val='none')

      # 列與列之間的微小間隔
      spacer = doc.add_paragraph()
      spacer.paragraph_format.space_before = Pt(0)
      spacer.paragraph_format.space_after = Pt(4)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    safe_job_title = (
        ''.join(c for c in display_title if c.isalnum() or c in (' ', '_', '-'))
        .strip()
        .replace(' ', '_')
    )

    st.download_button(
        label='💾 點擊下載 Word 報告 (.docx)',
        data=buffer,
        file_name=(
            f"Report_{safe_job_title}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
        ),
        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )

# --- App 底部專屬水印 (Footer) ---
st.markdown('---')
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 14px;'>"
    '🛠️ <b>Design by nikki 💅</b>'
    '</div>',
    unsafe_allow_html=True,
)
