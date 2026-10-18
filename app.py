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

# 預設兩層資料夾結構 { "第一層主資料夾": ["第二層子資料夾A", "第二層子資料夾B"] }
DEFAULT_FOLDER_TREE = {
    "01_契約與補充規定": ["主契約條款", "投標須知與特訂條款"],
    "02_施工規範與圖說": ["土建工程規範", "機電工程規範", "職安衛與環保規範"],
    "03_設計變更與核定公文": ["契約變更設計", "重要往來公文"],
    "04_會議紀錄與簡報": ["工務會議紀錄", "專案進度簡報"],
    "05_材料送審與查驗紀錄": ["材料送審資料", "現場查驗與自主檢查"]
}

# ================= 2. 密碼驗證機制 =================
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

# ================= 3. 讀取與儲存資料庫 (自動無痛升級雙層結構) =================
def load_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "rb") as f:
                data = pickle.load(f)
            # 若已經是 v3.0 雙層結構字典
            if isinstance(data, dict) and "folder_tree" in data:
                return data
            # 若是 v2.0 單層結構，自動升格為雙層
            elif isinstance(data, dict) and "categories" in data:
                tree = {cat: ["一般文件"] for cat in data["categories"]}
                for k, v in DEFAULT_FOLDER_TREE.items():
                    if k not in tree:
                        tree[k] = v.copy()
                upgraded_chunks = []
                for c in data.get("chunks", []):
                    cat = c.get("category", "04_會議紀錄與簡報")
                    if cat not in tree:
                        tree[cat] = ["一般文件"]
                    upgraded_chunks.append({
                        "file_name": c["file_name"],
                        "cat1": cat,
                        "cat2": tree[cat][0],
                        "location": c.get("location", ""),
                        "content": c["content"]
                    })
                return {"folder_tree": tree, "chunks": upgraded_chunks}
            # 若是 v1.0 最早期列表結構
            elif isinstance(data, list):
                upgraded_chunks = []
                for item in data:
                    fname = item["source"].split(" (")[0]
                    loc = item["source"][len(fname):].strip()
                    upgraded_chunks.append({
                        "file_name": fname,
                        "cat1": "04_會議紀錄與簡報",
                        "cat2": "專案進度簡報",
                        "location": loc,
                        "content": item["content"]
                    })
                return {"folder_tree": DEFAULT_FOLDER_TREE.copy(), "chunks": upgraded_chunks}
        except Exception:
            pass
    return {"folder_tree": DEFAULT_FOLDER_TREE.copy(), "chunks": []}

def save_db(db_dict):
    with open(DB_FILE, "wb") as f:
        pickle.dump(db_dict, f)

def extract_text_from_file(uploaded_file, cat1, cat2):
    text_chunks = []
    filename = uploaded_file.name
    lower_name = filename.lower()
    
    if lower_name.endswith(".pdf"):
        reader = pypdf.PdfReader(uploaded_file)
        for i, page in enumerate(reader.pages):
            page_text = page.extract_text()
            if page_text and page_text.strip():
                text_chunks.append({
                    "file_name": filename, "cat1": cat1, "cat2": cat2,
                    "location": f"(第 {i+1} 頁)", "content": page_text.strip()
                })
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
                "file_name": filename, "cat1": cat1, "cat2": cat2,
                "location": f"(段落 {i//800 + 1})", "content": full_text[i:i+800]
            })
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
                    "file_name": filename, "cat1": cat1, "cat2": cat2,
                    "location": f"(投影片第 {i+1} 頁)", "content": "\n".join(slide_texts)
                })
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
                        "file_name": filename, "cat1": cat1, "cat2": cat2,
                        "location": f"(工作表:{sheet} 區塊{i//800 + 1})", "content": full_sheet_text[i:i+800]
                    })
    elif lower_name.endswith(".txt"):
        full_text = uploaded_file.read().decode("utf-8", errors="ignore")
        for i in range(0, len(full_text), 800):
            text_chunks.append({
                "file_name": filename, "cat1": cat1, "cat2": cat2,
                "location": f"(段落 {i//800 + 1})", "content": full_text[i:i+800]
            })
    return text_chunks

def search_internal_db(query, chunks, selected_cat1s, selected_cat2s, top_k=10):
    if not chunks:
        return "（目前內部工程資料庫尚無檔案，請先從左側邊欄上傳檔案）"
    
    filtered_chunks = [
        c for c in chunks 
        if c.get("cat1") in selected_cat1s and (not selected_cat2s or c.get("cat2") in selected_cat2s)
    ]
    if not filtered_chunks:
        return "（目前勾選的資料夾範圍內尚無任何文件）"
        
    corpus = [f"{item['file_name']} {item['content']}" for item in filtered_chunks]
    vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
    tfidf_matrix = vectorizer.fit_transform(corpus + [query])
    cosine_sim = cosine_similarity(tfidf_matrix[-1], tfidf_matrix[:-1]).flatten()
    top_indices = cosine_sim.argsort()[-top_k:][::-1]
    
    results = []
    for idx in top_indices:
        if cosine_sim[idx] > 0.01:
            item = filtered_chunks[idx]
            results.append(f"【路徑：{item['cat1']} > {item['cat2']}｜來源：{item['file_name']} {item['location']}】\n{item['content']}")
    if not results:
        for idx in range(min(5, len(filtered_chunks))):
            item = filtered_chunks[idx]
            results.append(f"【路徑：{item['cat1']} > {item['cat2']}｜來源：{item['file_name']} {item['location']}】\n{item['content']}")
    return "\n\n---\n\n".join(results)

# ================= 4. 側邊欄：雙層資料夾管理、檔案改名與上傳 =================
db = load_db()
folder_tree = db["folder_tree"]
chunks = db["chunks"]

# 整理目前檔案對照表: { file_name: (cat1, cat2) }
file_info_map = {}
for c in chunks:
    c1 = c.get("cat1", list(folder_tree.keys())[0])
    if c1 not in folder_tree:
        folder_tree[c1] = ["預設子夾"]
    c2 = c.get("cat2", folder_tree[c1][0])
    if c2 not in folder_tree[c1]:
        folder_tree[c1].append(c2)
    file_info_map[c["file_name"]] = (c1, c2)

with st.sidebar:
    st.header("📂 工程雙層資料庫管理")
    st.info(f"目前共 **{len(file_info_map)}** 份文件（共 {len(chunks)} 個檢索區塊）")
    
    # --- 功能 A：樹狀檢視兩層資料夾與檔案 ---
    with st.expander("📁 瀏覽兩層資料夾與檔案", expanded=True):
        if not file_info_map:
            st.caption("目前尚無檔案，請從下方選擇資料夾並上傳。")
        else:
            for c1, sub_list in folder_tree.items():
                c1_files = [f for f, (f_c1, _) in file_info_map.items() if f_c1 == c1]
                if c1_files:
                    st.markdown(f"**🗂️ {c1}**")
                    for c2 in sub_list:
                        c2_files = sorted([f for f, (f_c1, f_c2) in file_info_map.items() if f_c1 == c1 and f_c2 == c2])
                        if c2_files:
                            st.markdown(f" └ 📂 **{c2}** ({len(c2_files)})")
                            for f in c2_files:
                                st.caption(f"   📄 {f}")

    # --- 功能 B：直接在網頁管理「第一層」與「第二層」資料夾 ---
    with st.expander("⚙️ 管理兩層資料夾 (新增/改名/刪除)"):
        tab1, tab2 = st.tabs(["第一層(主夾)", "第二層(子夾)"])
        
        with tab1:
            new_c1 = st.text_input("➕ 新增第一層資料夾：", placeholder="例如：06_職安衛與環保")
            if st.button("新增主資料夾", use_container_width=True) and new_c1.strip():
                clean_c1 = new_c1.strip()
                if clean_c1 not in folder_tree:
                    folder_tree[clean_c1] = ["一般文件"]
                    save_db(db)
                    st.success(f"已建立主資料夾：{clean_c1}")
                    st.rerun()
            st.divider()
            target_c1 = st.selectbox("選擇要改名或刪除的主資料夾：", list(folder_tree.keys()), key="sel_c1")
            rename_c1 = st.text_input("修改主資料夾名稱：", value=target_c1, key="ren_c1")
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("儲存主夾名稱", use_container_width=True):
                    clean_r1 = rename_c1.strip()
                    if clean_r1 and clean_r1 != target_c1:
                        new_tree = {}
                        for k, v in folder_tree.items():
                            new_tree[clean_r1 if k == target_c1 else k] = v
                        db["folder_tree"] = new_tree
                        for c in chunks:
                            if c.get("cat1") == target_c1:
                                c["cat1"] = clean_r1
                        save_db(db)
                        st.rerun()
            with col_b:
                if st.button("🗑️ 刪除主夾", use_container_width=True):
                    if len(folder_tree) > 1:
                        del folder_tree[target_c1]
                        fallback_c1 = list(folder_tree.keys())[0]
                        fallback_c2 = folder_tree[fallback_c1][0]
                        for c in chunks:
                            if c.get("cat1") == target_c1:
                                c["cat1"] = fallback_c1
                                c["cat2"] = fallback_c2
                        save_db(db)
                        st.rerun()
                    else:
                        st.warning("至少需保留一個主資料夾！")

        with tab2:
            parent_c1 = st.selectbox("1. 先選擇所屬的「第一層主資料夾」：", list(folder_tree.keys()), key="p_c1")
            new_c2 = st.text_input("➕ 在此主夾下新增「第二層子資料夾」：", placeholder="例如：高空作業查驗")
            if st.button("新增子資料夾", use_container_width=True) and new_c2.strip():
                clean_c2 = new_c2.strip()
