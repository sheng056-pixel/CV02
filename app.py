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
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "").strip()

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
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.strip():
                            slide_texts.append(paragraph.text.strip())
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

def search_internal_db(query, db_data, top_k=10):
    if not db_data:
        return "（目前內部工程資料庫尚無檔案，請先從左側邊欄上傳檔案並點選『匯入並更新知識庫』）"
    corpus = [item["content"] for item in db_data]
    vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
    tfidf_matrix = vectorizer.fit_transform(corpus + [query])
    cosine_sim = cosine_similarity(tfidf_matrix[-1], tfidf_matrix[:-1]).flatten()
    top_indices = cosine_sim.argsort()[-top_k:][::-1]
    
    results = []
    for idx in top_indices:
        if cosine_sim[idx] > 0.01:
            results.append(f"【來源：{db_data[idx]['source']}】\n{db_data[idx]['content']}")
    # 若關鍵字沒有完全命中，至少提供前 5 個段落供 AI 參考
    if not results:
        for idx in range(min(5, len(db_data))):
            results.append(f"【來源：{db_data[idx]['source']}】\n{db_data[idx]['content']}")
    return "\n\n---\n\n".join(results)

# ================= 4. 主畫面與檔案管理側邊欄 =================
st.title("🏗️ 工程文件問答＆法規智囊團")
st.caption("同時檢索「內部工程契約/規範/簡報/會議紀錄」與「營建採購法規」")

db_data = load_db()

with st.sidebar:
    st.header("📂 工程資料庫管理")
    unique_files = sorted(list(set([item["source"].split(" (")[0] for item in db_data])))
    st.info(f"目前庫內共有 **{len(unique_files)}** 份文件（共 {len(db_data)} 個檢索區塊）")
    
    with st.expander("查看已匯入的文件清單"):
        if unique_files:
            for f in unique_files:
                st.write(f"📄 {f}")
        else:
            st.write("尚無檔案")
            
    st.subheader("➕ 新增工程文件 (無上限)")
    uploaded_files = st.file_uploader(
        "1. 先選擇檔案 (支援 PDF, Word, PPTX, Excel, TXT)", 
        type=["pdf", "docx", "pptx", "xlsx", "txt"], 
        accept_multiple_files=True
    )
    if st.button("2. 點此匯入並更新知識庫", type="primary", use_container_width=True):
        if uploaded_files:
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
        else:
            st.warning("請先在上方選擇檔案後再按匯入！")
            
    st.divider()
    enable_web_search = st.checkbox("🌐 啟用 Google 即時聯網搜尋", value=False, help="若免費 API 額度受限，建議關閉此選項，僅使用內部資料庫與 AI 內建法規知識。")

# ================= 5. 雙軌智慧問答區 =================
if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "assistant", "content": "你好！我是你們的工程專案智囊團。請先從**左側邊欄上傳工程檔案並按下「匯入並更新知識庫」**，接著就可以直接問我簡報內容、預算金額、契約規範或採購法規囉！"}
    ]

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if user_query := st.chat_input("請輸入問題（例如：cv02工程預算大概是多少錢？）"):
    st.session_state.messages.append({"role": "user", "content": user_query})
    with st.chat_message("user"):
        st.markdown(user_query)
        
    with st.chat_message("assistant"):
        with st.spinner("🔍 正在翻閱工程資料庫並整理答案..."):
            internal_context = search_internal_db(user_query, db_data, top_k=10)
            
            system_prompt = f"""
你是一位專業、嚴謹的「工程專案智囊團 AI 顧問」。
使用者是工程團隊成員，請根據以下【內部工程資料庫檢索結果】以及你的營建與政府採購法規知識來回答問題。

【目前內部資料庫已上傳的文件清單】：
{", ".join(unique_files) if unique_files else "⚠️ 尚未匯入任何文件（使用者可能選了檔案但忘記按『匯入並更新知識庫』按鈕）"}

【回答嚴格守則】：
1. 內部資料引用：凡是引用內部工程資料的數據或規定，必須明確標註「【來源：檔名與頁碼/投影片頁數/段落】」。
2. 新舊版本覆蓋原則：若內部多份文件出現規定衝突（例如原契約 vs 變更設計、早期會議 vs 最新會議紀錄），請務必以「檔名日期最新、或屬於變更設計/最新核定公文」的內容為準，並主動提醒使用者新舊版本的差異。
3. 法規說明：若問題涉及政府採購法、公共工程委員會規範、建築技術規則或職安法規，請一併提供相關法規依據與條號。
4. 嚴禁瞎掰：若內部資料庫沒有提到該項工程數據，請誠實告知「目前內部文件未載明」，切勿自行編造工程預算或規範數據。

=== 內部工程資料庫檢索結果 ===
{internal_context}
==============================

使用者問題：{user_query}
"""
            client = genai.Client(api_key=GEMINI_API_KEY)
            answer = None
            
            # 自動嘗試多個可用模型，並自動處理免費版聯網搜尋配額限制 (429)
            models_to_try = ["gemini-3.8-flash", "gemini-2.5-flash-lite", "gemini-2.0-flash"]
            last_error = ""
            
            for model_name in models_to_try:
                try:
                    config_args = {"temperature": 0.2}
                    if enable_web_search:
                        config_args["tools"] = [{"google_search": {}}]
                        
                    response = client.models.generate_content(
                        model=model_name,
                        contents=system_prompt,
                        config=types.GenerateContentConfig(**config_args)
                    )
                    answer = response.text
                    break
                except Exception as e:
                    err_str = str(e)
                    last_error = err_str
                    # 如果是開啟聯網搜尋導致 429 額度不足，自動關閉聯網搜尋再重試一次該模型！
                    if enable_web_search and "429" in err_str:
                        try:
                            response = client.models.generate_content(
                                model=model_name,
                                contents=system_prompt,
                                config=types.GenerateContentConfig(temperature=0.2)
                            )
                            answer = response.text + "\n\n*(💡 系統提示：因免費版 API 之 Google 即時搜尋配額受限，本次回答已自動改為使用「內部工程資料庫 ＋ AI 內建法規知識庫」生成)*"
                            break
                        except Exception as e2:
                            last_error = str(e2)
                    continue
                    
            if not answer:
                answer = f"⚠️ 系統呼叫失敗。錯誤訊息：{last_error}\n\n**解決建議**：若持續出現 `429 RESOURCE_EXHAUSTED`，代表該組 API Key 的免費額度在 Google AI Studio 中被設為 0，請至 [aistudio.google.com/apikey](https://aistudio.google.com/apikey) 點選 **Create API key in new project** 重新產生一組 `AIzaSy` 開頭的金鑰，並更新至 Streamlit 後台的 Secrets 中。"
                
            st.markdown(answer)
            st.session_state.messages.append({"role": "assistant", "content": answer
