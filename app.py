import os
import streamlit as st
from google import genai
from google.genai import types
import pypdf
import docx
from pptx import Presentation
import openpyxl
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import pickle

# ================= 1. 系統基本設定 =================
st.set_page_config(page_title="工程專案智囊團", page_icon="🏗️", layout="centered")

TEAM_PASSWORD = st.secrets.get("TEAM_PASSWORD", "1234")
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")

DB_FILE = "engineering_knowledge_db.pkl"

# ================= 2. 密碼驗證機制 (手機介面友善) =================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🏗️ 工程專案智囊團")
    pwd = st.text_input("請輸入團隊通行密碼：", type="password")
    if st.button("登入系統", use_container_width=True):
        if pwd == TEAM_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("密碼錯誤，請重新輸入！")
    st.stop()

# ================= 3. 讀取與儲存工程資料庫 =================
def load_db():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "rb") as f:
            return pickle.load(f)
    return []

def save_db(data):
    with open(DB_FILE, "wb") as f:
        pickle.dump(data, f)

def extract_text_from_file(uploaded_file):
    text_chunks = []
    filename = uploaded_file.name
    lower_name = filename.lower()
    
    # 1. 讀取 PDF
    if lower_name.endswith(".pdf"):
        reader = pypdf.PdfReader(uploaded_file)
        for i, page in enumerate(reader.pages):
            page_text = page.extract_text()
            if page_text and page_text.strip():
                text_chunks.append({
                    "source": f"{filename} (第 {i+1} 頁)",
                    "content": page_text.strip()
                })
    # 2. 讀取 Word (.docx)
    elif lower_name.endswith(".docx"):
        doc = docx.Document(uploaded_file)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        # 連同 Word 內的表格一起讀取
        for table in doc.tables:
            for row in table.rows:
                row_text = " | ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])
                if row_text:
                    paragraphs.append(row_text)
        full_text = "\n".join(paragraphs)
        for i in range(0, len(full_text), 800):
            text_chunks.append({
                "source": f"{filename} (段落 {i//800 + 1})",
                "content": full_text[i:i+800]
            })
    # 3. 讀取 PowerPoint 簡報 (.pptx)
    elif lower_name.endswith(".pptx"):
        prs = Presentation(uploaded_file)
        for i, slide in enumerate(prs.slides):
            slide_texts = []
            for shape in slide.shapes:
                # 抓取文字框
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.strip():
                            slide_texts.append(paragraph.text.strip())
                # 抓取簡報內的表格數據
                if shape.has_table:
                    for row in shape.table.rows:
                        row_text = " | ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])
                        if row_text:
                            slide_texts.append(row_text)
            if slide_texts:
                text_chunks.append({
                    "source": f"{filename} (投影片第 {i+1} 頁)",
                    "content": "\n".join(slide_texts)
                })
    # 4. 讀取 Excel 表格 (.xlsx)
    elif lower_name.endswith(".xlsx"):
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            rows_text = []
            for row in ws.iter_rows(values_only=True):
                row_vals = [str(cell).strip() for cell in row if cell is not None and str(cell).strip()]
                if row_vals:
                    rows_text.append(" | ".join(row_vals))
            full_sheet_text = "\n".join(rows_text)
            for i in range(0, len(full_sheet_text), 800):
                if full_sheet_text[i:i+800].strip():
                    text_chunks.append({
                        "source": f"{filename} (工作表:{sheet} 區塊{i//800 + 1})",
                        "content": full_sheet_text[i:i+800]
                    })
    # 5. 讀取純文字 (.txt)
    elif lower_name.endswith(".txt"):
        full_text = uploaded_file.read().decode("utf-8", errors="ignore")
        for i in range(0, len(full_text), 800):
            text_chunks.append({
                "source": f"{filename} (段落 {i//800 + 1})",
                "content": full_text[i:i+800]
            })
    return text_chunks

def search_internal_db(query, db_data, top_k=6):
    if not db_data:
        return "（目前內部工程資料庫尚無檔案）"
    corpus = [item["content"] for item in db_data]
    vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
    tfidf_matrix = vectorizer.fit_transform(corpus + [query])
    cosine_sim = cosine_similarity(tfidf_matrix[-1], tfidf_matrix[:-1]).flatten()
    top_indices = cosine_sim.argsort()[-top_k:][::-1]
    
    results = []
    for idx in top_indices:
        if cosine_sim[idx] > 0.05:
            results.append(f"【來源：{db_data[idx]['source']}】\n{db_data[idx]['content']}")
    return "\n\n---\n\n".join(results) if results else "（內部工程資料庫中未找到高度相關段落）"

# ================= 4. 主畫面與檔案管理側邊欄 =================
st.title("🏗️ 工程文件問答＆法規智囊團")
st.caption("同時檢索「內部工程契約/規範/簡報/會議紀錄」與「Google 最新外部營建採購法規」")

db_data = load_db()

with st.sidebar:
    st.header("📂 工程資料庫管理")
    unique_files = sorted(list(set([item["source"].split(" (")[0] for item in db_data])))
    st.info(f"目前庫內共有 **{len(unique_files)}** 份文件（共 {len(db_data)} 個檢索區塊）")
    
    with st.expander("查看已匯入的文件清單"):
        for f in unique_files:
            st.write(f"📄 {f}")
            
    st.subheader("➕ 新增工程文件 (無上限)")
    uploaded_files = st.file_uploader(
        "上傳契約、規範、簡報或報表 (支援 PDF, Word, PPTX, Excel, TXT)", 
        type=["pdf", "docx", "pptx", "xlsx", "txt"], 
        accept_multiple_files=True
    )
    if st.button("匯入並更新知識庫", use_container_width=True) and uploaded_files:
        with st.spinner("正在解析文件並寫入資料庫..."):
            new_count = 0
            for uf in uploaded_files:
                db_data = [item for item in db_data if not item["source"].startswith(uf.name + " (")]
                chunks = extract_text_from_file(uf)
                db_data.extend(chunks)
                new_count += 1
            save_db(db_data)
            st.success(f"成功匯入 {new_count} 份工程文件！")
            st.rerun()

# ================= 5. 雙軌智慧問答區 =================
if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "assistant", "content": "你好！我是你們的工程專案智囊團。你可以直接問我**契約規範、簡報內容、驗收標準、歷次會議紀錄**，或是叫我幫你**上網比對最新的政府採購法與公共工程規範**！"}
    ]

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if user_query := st.chat_input("請輸入問題（例如：簡報中提到的施工進度為何？並查最新法規標準）"):
    st.session_state.messages.append({"role": "user", "content": user_query})
    with st.chat_message("user"):
        st.markdown(user_query)
        
    with st.chat_message("assistant"):
        with st.spinner("🔍 正在翻閱工程資料庫並同步搜尋最新外部法規..."):
            internal_context = search_internal_db(user_query, db_data, top_k=6)
            
            system_prompt = f"""
你是一位專業、嚴謹的「工程專案智囊團 AI 顧問」。
使用者是工程團隊成員，請根據以下【內部工程資料庫檢索結果】以及【Google 網路搜尋最新法規】來回答問題。

【回答嚴格守則】：
1. 內部資料引用：凡是引用內部工程資料的數據或規定，必須明確標註「【來源：檔名與頁碼/投影片頁數/段落】」。
2. 新舊版本覆蓋原則：若內部多份文件出現規定衝突（例如原契約 vs 變更設計、早期會議 vs 最新會議紀錄），請務必以「檔名日期最新、或屬於變更設計/最新核定公文」的內容為準，並主動提醒使用者新舊版本的差異。
3. 外部法規查詢：若問題涉及政府採購法、公共工程委員會規範、建築技術規則、職安法規或一般工程標準，請務必啟動 Google Search 查詢最新法規，並標明法規名稱與條號。
4. 嚴禁瞎掰：若內部資料庫沒有提到、網路上也查無相關標準，請誠實告知「目前內部文件未載明」，切勿自行編造工程數據。

=== 內部工程資料庫檢索結果 ===
{internal_context}
==============================

使用者問題：{user_query}
"""
            try:
                client = genai.Client(api_key=GEMINI_API_KEY)
                response = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents=system_prompt,
                    config=types.GenerateContentConfig(
                        tools=[{"google_search": {}}],
                        temperature=0.2,
                    )
                )
                answer = response.text
            except Exception as e:
                answer = f"⚠️ 系統呼叫失敗，請確認 API Key 是否正確設定。錯誤訊息：{e}"
                
            st.markdown(answer)
            st.session_state.messages.append({"role": "assistant", "content": answer})
