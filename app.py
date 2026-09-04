import streamlit as st
from docx import Document
import io
from datetime import datetime

st.title("📸 Sleek-Industrial 進度記錄器")

# --- 0. 設定 Job Title (工程項目名稱) ---
st.subheader("📌 項目基本資料")
job_title = st.text_input("Job Title / 工程項目名稱", value="", placeholder="例如: Regent Hotel F3 改善工程")

st.divider()

# 1. 初始化 Session State 黎暫存記錄
if 'records' not in st.session_state:
    st.session_state['records'] = []

# --- 輸入當前記錄 ---
st.subheader("1️⃣ 新增現場記錄")

# 樓層與房間全部改為選填
floor = st.text_input("樓層 (Floor) [選填]", placeholder="例如: B2 / G/F / 3/F (可留空)")
room = st.text_input("房間 / 區域 (Room / Area) [選填]", placeholder="例如: Function Room A / 掣房 (可留空)")
category = st.selectbox("工程類別", ["AC", "FS", "P&D", "EL", "OTHER"])
remarks = st.text_area("工作備忘", placeholder="請輸入工作內容或備忘...")

# 拍照或上傳相片
photo = st.file_uploader("拍攝或上傳現場相片", type=['jpg', 'jpeg', 'png'])

# 加入暫存清單按鈕
if st.button("➕ 新增到今日清單"):
    # 檢查條件改為：只要有相片即可（樓層同房間全變成選填）
    if photo is not None:
        # 將相片轉成獨立嘅 BytesIO 串流
        photo_bytes = io.BytesIO(photo.getvalue())
        
        # 處理顯示文字
        f_val = floor.strip() if floor else ""
        r_val = room.strip() if room else ""
        
        st.session_state['records'].append({
            "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "floor": f_val,
            "room": r_val,
            "category": category,
            "remarks": remarks,
            "photo": photo_bytes
        })
        st.success("成功新增現場記錄！")
    else:
        st.warning("請上傳或拍攝現場相片！")

st.divider()

# --- 2. 顯示已記錄嘅清單 ---
st.subheader("📋 今日已記錄項目")
if len(st.session_state['records']) > 0:
    for i, rec in enumerate(st.session_state['records']):
        # 組裝顯示標題
        loc_parts = []
        if rec['floor']: loc_parts.append(f"樓層: {rec['floor']}")
        if rec['room']: loc_parts.append(f"區域: {rec['room']}")
        loc_str = " - ".join(loc_parts) if loc_parts else "未註明位置"
        
        with st.expander(f"項目 #{i+1}: {loc_str} ({rec['category']})"):
            st.write(f"時間：{rec['time']}")
            st.write(f"備忘：{rec['remarks']}")
            
            # 預覽相片前先重置指標
            rec['photo'].seek(0)
            st.image(rec['photo'], width=300)
            
            if st.button(f"刪除此項 #{i+1}", key=f"del_{i}"):
                st.session_state['records'].pop(i)
                st.rerun()
                
    if st.button("🗑️ 清空所有記錄"):
        st.session_state['records'] = []
        st.rerun()
else:
    st.info("暫時未有記錄，請喺上面新增。")

st.divider()

# --- 3. 最後一鍵生成 REPORT ---
st.subheader("3️⃣ 匯出報告")
if st.button("📥 一鍵生成 Word 報告"):
    if len(st.session_state['records']) == 0:
        st.warning("請先新增至少一個記錄先可以出 Report！")
    else:
        doc = Document()
        
        doc.add_heading(f"工程進度巡檢報告", 0)
        display_title = job_title if job_title.strip() != "" else "Unnamed Project"
        doc.add_paragraph(f"Job Title: {display_title}")
        doc.add_paragraph(f"生成日期：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
        doc.add_paragraph(f"總記錄項目數：{len(st.session_state['records'])} 項\n")
        doc.add_paragraph("-" * 40)
        
        for i, rec in enumerate(st.session_state['records']):
            loc_parts = []
            if rec['floor']: loc_parts.append(f"樓層 {rec['floor']}")
            if rec['room']: loc_parts.append(f"區域 {rec['room']}")
            loc_title = f"項目 {i+1}: " + (" - ".join(loc_parts) if loc_parts else "未註明位置")
            
            doc.add_heading(loc_title, level=2)
            doc.add_paragraph(f"• 工程類別：{rec['category']}")
            doc.add_paragraph(f"• 記錄時間：{rec['time']}")
            doc.add_paragraph(f"• 工作備忘：{rec['remarks']}")
            
            try:
                # 插入 Word 前將指標歸零
                rec['photo'].seek(0)
                doc.add_picture(rec['photo'], width=docx.shared.Inches(4.5))
            except Exception as e:
                doc.add_paragraph("[相片載入失敗]")
            
            doc.add_paragraph("-" * 30)
            
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
