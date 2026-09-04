import streamlit as st
from docx import Document
import io
from datetime import datetime

st.title("📸 Sleek-Industrial 進度記錄器")

# --- 0. 設定 Job Title (工程項目名稱) ---
st.subheader("📌 項目基本資料")
job_title = st.text_input("Job Title / 工程項目名稱", value="Regent Hotel E&M Diversion Work", placeholder="例如: Regent Hotel F3 改善工程")

st.divider()

# 1. 初始化 Session State 黎暫存記錄
if 'records' not in st.session_state:
    st.session_state['records'] = []

# --- 輸入當前記錄 ---
st.subheader("1️⃣ 新增現場記錄")

# 樓層改成文字輸入框
floor = st.text_input("樓層 (Floor)", placeholder="例如: B2 / G/F / 3/F")
room = st.text_input("房間 / 區域 (Room / Area)", placeholder="例如: Function Room A / 掣房")

# 工程類別改成要求的選項
category = st.selectbox("工程類別", ["AC", "FS", "P&D", "EL", "OTHER"])

# 刪除「發現問題」字眼，改為純備忘
remarks = st.text_area("工作備忘", placeholder="請輸入工作內容或備忘...")
photo = st.camera_input("拍攝現場相片") # 或者用 st.file_uploader

# 加入暫存清單按鈕
if st.button("➕ 新增到今日清單"):
    if photo is not None and floor and room:
        # 將資料加入 session_state
        st.session_state['records'].append({
            "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "floor": floor,
            "room": room,
            "category": category,
            "remarks": remarks,
            "photo": photo
        })
        st.success(f"成功新增：{floor} - {room}")
    else:
        st.warning("請填寫樓層、房間名稱並拍攝相片！")

st.divider()

# --- 2. 顯示已記錄嘅清單 ---
st.subheader("📋 今日已記錄項目")
if len(st.session_state['records']) > 0:
    for i, rec in enumerate(st.session_state['records']):
        with st.expander(f"項目 #{i+1}: {rec['floor']} - {rec['room']} ({rec['category']})"):
            st.write(f"時間：{rec['time']}")
            st.write(f"備忘：{rec['remarks']}")
            st.image(rec['photo'], width=300)
            
            # 刪除單個記錄嘅按鈕
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
        # 建立 Word 文件
        doc = Document()
        
        # 融入 Job Title
        doc.add_heading(f"工程進度巡檢報告", 0)
        doc.add_paragraph(f"Job Title: {job_title}")
        doc.add_paragraph(f"生成日期：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
        doc.add_paragraph(f"總記錄項目數：{len(st.session_state['records'])} 項\n")
        doc.add_paragraph("-" * 40)
        
        # 循環寫入每一項記錄
        for i, rec in enumerate(st.session_state['records']):
            doc.add_heading(f"項目 {i+1}: {rec['floor']} - {rec['room']}", level=2)
            doc.add_paragraph(f"• 工程類別：{rec['category']}")
            doc.add_paragraph(f"• 記錄時間：{rec['time']}")
            doc.add_paragraph(f"• 工作備忘：{rec['remarks']}")
            
            # 插入相片
            try:
                doc.add_picture(rec['photo'], width=docx.shared.Inches(4.5))
            except Exception as e:
                doc.add_paragraph("[相片載入失敗]")
            
            doc.add_paragraph("-" * 30) # 分隔線
            
        # 儲存至記憶體
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        # 檔名自動帶埋 Job Title 方便辨識
        safe_job_title = "".join(c for c in job_title if c.isalnum() or c in (' ', '_', '-')).strip().replace(' ', '_')
        
        # 提供下載
        st.download_button(
            label="💾 點擊下載 Word 報告 (.docx)",
            data=buffer,
            file_name=f"Report_{safe_job_title}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
