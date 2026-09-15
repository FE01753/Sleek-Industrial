# --- 5. 一鍵生成 Word 報告 ---
st.subheader("3️⃣ 匯出報告")
if st.button("📥 一鍵生成 Word 報告", type="primary", use_container_width=True):
    if len(st.session_state["records"]) == 0:
        st.warning("請先新增至少一個記錄先可以出 Report！")
    else:
        doc = Document()

        for section in doc.sections:
            section.top_margin = Inches(0.4)
            section.bottom_margin = Inches(0.4)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

            # -------------------------------------------------------------
            # ✨ 頁首設置 (Header): 加入遠東工程 Logo
            # -------------------------------------------------------------
            header = section.header
            header_p = header.paragraphs[0]
            header_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            # 支援多種常見檔名與格式
            logo_candidates = ["logo.png", "logo.jpg", "logo.jpeg", "LOGO.PNG"]
            logo_path = None
            for cand in logo_candidates:
                if os.path.exists(cand):
                    logo_path = cand
                    break

            if logo_path:
                try:
                    header_run = header_p.add_run()
                    # 加載圖片並調整大小 (高度 0.45 吋)
                    header_run.add_picture(logo_path, height=Inches(0.45))
                except Exception as e:
                    st.error(f"頁首 Logo 載入失敗: {e}")
            else:
                # 提示使用者未放 Logo 檔
                st.info(
                    "💡 提示：專案目錄中找不到 `logo.png`，Word 頁首已自動跳過 Logo。"
                )

            # -------------------------------------------------------------
            # ✨ 頁尾設置 (Footer): 加入動態頁碼 (Page X)
            # -------------------------------------------------------------
            footer = section.footer
            footer_p = footer.paragraphs[0]
            footer_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            footer_run = footer_p.add_run("Page ")
            footer_run.font.size = Pt(9)
            footer_run.font.color.rgb = RGBColor(128, 128, 128)

            add_page_number(footer_run)
