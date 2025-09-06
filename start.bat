@echo off
echo 正在啟動SOP生成系統...
echo.
echo 如果這是第一次運行，請確保已安裝必要套件：
echo pip install streamlit python-docx Pillow
echo.
streamlit run sop_generator.py
