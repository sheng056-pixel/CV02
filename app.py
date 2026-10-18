import os
import re
import hashlib
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
st.set_page_config(page_title="工程專案智囊團", page_icon="🏗️", layout="wide")

TEAM_PASSWORD = st.secrets.get("TEAM_PASSWORD", "1234")
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "").strip()

DB_FILE = "engineering_knowledge_db.pkl"

DEFAULT_FOLDER_TREE = {
    "01_契約與補充規定": ["主契約條款", "投標須知與特訂條款"],
    "02_施工規範與圖說": ["土建工程規範", "機電工程規範", "職安衛與環保規範"],
    "03_設計變更與核定公文": ["契約變更設計", "重要往來公文"],
    "04_會議紀錄與簡報": ["工務會議紀錄", "專案進度簡報"],
    "05_材料送審與查驗紀錄": ["材料送審資料", "現場查驗與自主檢查"]
}

# 預設組員帳號清單 (可在網頁後台直接新增/修改)
DEFAULT_USERS = {
    "聖引": {"password": TEAM_PASSWORD, "role": "管理員"},
    "組員A": {"password": TEAM_PASSWORD, "role": "一般組員"}
}

def make_token(username, password):
    raw = f"{username}:{password}:cv02_salt"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]

# ================= 2. 讀取與儲存資料庫 (含帳號系統) =================
def load_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "rb") as f:
                data = pickle.load(f)
            if isinstance(data, dict) and "folder_tree" in data:
                if "users" not in data:
                    data["users"] = DEFAULT_USERS.copy()
                return data
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
                        "content": c["content"],
                        "uploader": "系統管理員"
                    })
                return {"folder_tree": tree, "chunks": upgraded_chunks, "users": DEFAULT_USERS.copy()}
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
                        "content": item["content"],
                        "uploader": "系統管理員"
                    })
                return {"folder_tree": DEFAULT_FOLDER_TREE.copy(), "chunks": upgraded_chunks, "users": DEFAULT_USERS.copy()}
        except Exception:
            pass
    return {"folder_tree": DEFAULT_FOLDER_TREE.copy(), "chunks": [], "users": DEFAULT_USERS.copy()}

def save_db(db_dict):
    with open(DB_FILE, "wb") as f:
        pickle.dump(db_dict, f)

db = load_db()
users_db = db["users"]

# ================= 3. 組員帳號登入 & 同裝置自動記憶機制 =================
if "current_user" not in st.session_state:
    st.session_state.current_user = None

# 檢查網址是否有自動記憶憑證 (URL Query Params)
q_user = st.query_params.get("user", "")
q_token = st.query_params.get("token", "")

if not st.session_state.current_user and q_user in users_db:
    expected_token = make_token(q_user, users_db[q_user]["password"])
    if q_token == expected_token:
        st.session_state.current_user = q_user

if not st.session_state.current_user:
    st.title("🏗️ 工程專案智囊團 — 團隊登入")
    st.caption("請選擇您的組員帳號登入；勾選「記住我」後將網頁加入書籤或手機主畫面，下次點開即可免輸入密碼自動登入！")
    
    with st.container(border=True):
        user_list = list(users_db.keys())
        login_user = st.selectbox("👤 選擇組員帳號：", user_list)
        login_pwd = st.text_input("🔑 請輸入密碼：", type="password")
        remember_me = st.checkbox("📌 在此裝置記住我的登入狀態（下次自動登入）", value=True)
        
        if st.button("🚀 登入系統", type="primary", use_container_width=True):
            # 允許使用個人密碼或全域團隊密碼登入
            valid_pwd = users_db[login_user]["password"]
            if login_pwd == valid_pwd or login_pwd == TEAM_PASSWORD:
                st.session_state.current_user = login_user
                if remember_me:
                    st.query_params["user"] = login_user
                    st.query_params["token"] = make_token(login_user, valid_pwd)
                st.rerun()
            else:
                st.error("密碼錯誤，請重新輸入！")
    st.stop()

current_user = st.session_state.current_user
current_role = users_db.get(current_user, {}).get("role", "一般組員")

# ================= 4. 文件解析與搜尋核心函數 =================
def extract_text_from_file(uploaded_file, cat1, cat2, uploader):
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
                    "location": f"(第 {i+1} 頁)", "content": page_text.strip(), "uploader": uploader
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
                "location": f"(段落 {i//800 + 1})", "content": full_text[i:i+800], "uploader": uploader
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
                    "location": f"(投影片第 {i+1} 頁)", "content": "\n".join(slide_texts), "uploader": uploader
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
                        "location": f"(工作表:{sheet} 區塊{i//800 + 1})", "content": full_sheet_text[i:i+800], "uploader": uploader
                    })
    elif lower_name.endswith(".txt"):
        full_text = uploaded_file.read().decode("utf-8", errors="ignore")
        for i in range(0, len(full_text), 800):
            text_chunks.append({
                "file_name": filename, "cat1": cat1, "cat2": cat2,
                "location": f"(段落 {i//800 + 1})", "content": full_text[i:i+800], "uploader": uploader
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

def highlight_keywords(text, keywords):
    highlighted = text
    for kw in keywords:
        if kw:
            pattern = re.compile(re.escape(kw), re.IGNORECASE)
            highlighted = pattern.sub(
                lambda m: f"<mark style='background-color: #ffe066; padding: 2px 4px; border-radius: 4px; font-weight: bold;'>{m.group(0)}</mark>",
                highlighted
            )
    return highlighted

# ================= 5. 側邊欄：個人狀態、上傳置頂、資料夾與帳號管理 =================
folder_tree = db["folder_tree"]
chunks = db["chunks"]

file_info_map = {}
for c in chunks:
    c1 = c.get("cat1", list(folder_tree.keys())[0])
    if c1 not in folder_tree:
        folder_tree[c1] = ["預設子夾"]
    c2 = c.get("cat2", folder_tree[c1][0])
    if c2 not in folder_tree[c1]:
        folder_tree[c1].append(c2)
    file_info_map[c["file_name"]] = (c1, c2, c.get("uploader", "未知"))

with st.sidebar:
    # --- 顯示目前登入組員與登出按鈕 ---
    col_u1, col_u2 = st.columns([3, 2])
    with col_u1:
        st.markdown(f"👤 **{current_user}** `({current_role})`")
    with col_u2:
        if st.button("🚪 切換帳號", use_container_width=True):
            st.session_state.current_user = None
            st.query_params.clear()
            st.rerun()
            
    st.caption("💡 提示：已開啟同裝置自動記憶，請直接將此頁加入手機主畫面或書籤，下次點開免打密碼！")
    st.divider()
    
    st.header("📂 工程資料庫控制台")
    st.info(f"目前共 **{len(file_info_map)}** 份文件（共 {len(chunks)} 個檢索區塊）")
    
    # --- 【置頂第 1 順位】上傳工程文件 ---
    with st.expander("➕ 上傳新工程文件 (點此展開)", expanded=True):
        up_c1 = st.selectbox("1. 選擇第一層（主資料夾）：", list(folder_tree.keys()))
        up_c2 = st.selectbox("2. 選擇第二層（子資料夾）：", folder_tree[up_c1])
        uploaded_files = st.file_uploader(
            "3. 選擇檔案 (支援 PDF, Word, PPTX, Excel, TXT)", 
            type=["pdf", "docx", "pptx", "xlsx", "txt"], 
            accept_multiple_files=True
        )
        if st.button("🚀 點此匯入並更新知識庫", type="primary", use_container_width=True):
            if uploaded_files:
                with st.spinner("正在解析文件並寫入資料庫..."):
                    new_count = 0
                    for uf in uploaded_files:
                        db["chunks"] = [c for c in db["chunks"] if c["file_name"] != uf.name]
                        new_chunks = extract_text_from_file(uf, up_c1, up_c2, current_user)
                        db["chunks"].extend(new_chunks)
                        new_count += 1
                    save_db(db)
                    st.success(f"已由 {current_user} 匯入 {new_count} 份文件！")
                    st.rerun()
            else:
                st.warning("請先選擇檔案後再按匯入！")

    # --- 第 2 順位：瀏覽兩層資料夾 ---
    with st.expander("📁 瀏覽兩層資料夾與檔案", expanded=False):
        if not file_info_map:
            st.caption("目前尚無檔案。")
        else:
            for c1, sub_list in folder_tree.items():
                c1_files = [f for f, (f_c1, _, _) in file_info_map.items() if f_c1 == c1]
                if c1_files:
                    st.markdown(f"**🗂️ {c1}**")
                    for c2 in sub_list:
                        c2_files = sorted([f for f, (f_c1, f_c2, _) in file_info_map.items() if f_c1 == c1 and f_c2 == c2])
                        if c2_files:
                            st.markdown(f" └ 📂 **{c2}** ({len(c2_files)})")
                            for f in c2_files:
                                uploader_name = file_info_map[f][2]
                                st.caption(f"   📄 {f} *(by {uploader_name})*")

    # --- 第 3 順位：修改檔名 / 搬移資料夾 / 刪除檔案 ---
    if file_info_map:
        with st.expander("✏️ 修改檔名 / 搬移 / 刪除檔案", expanded=False):
            selected_file = st.selectbox("選擇要編輯的檔案：", sorted(list(file_info_map.keys())))
            new_file_name = st.text_input("修改檔案顯示名稱：", value=selected_file)
            
            curr_c1, curr_c2, _ = file_info_map[selected_file]
            c1_list = list(folder_tree.keys())
            c1_idx = c1_list.index(curr_c1) if curr_c1 in c1_list else 0
            move_c1 = st.selectbox("搬移至第一層主資料夾：", c1_list, index=c1_idx)
            
            c2_list = folder_tree[move_c1]
            c2_idx = c2_list.index(curr_c2) if curr_c2 in c2_list else 0
            move_c2 = st.selectbox("搬移至第二層子資料夾：", c2_list, index=c2_idx)
            
            b1, b2 = st.columns(2)
            with b1:
                if st.button("💾 儲存變更", use_container_width=True):
                    clean_fname = new_file_name.strip()
                    if clean_fname:
                        for c in chunks:
                            if c["file_name"] == selected_file:
                                c["file_name"] = clean_fname
                                c["cat1"] = move_c1
                                c["cat2"] = move_c2
                        save_db(db)
                        st.success("已更新！")
                        st.rerun()
            with b2:
                if st.button("🗑️ 刪除此檔案", use_container_width=True):
                    db["chunks"] = [c for c in chunks if c["file_name"] != selected_file]
                    save_db(db)
                    st.success("已刪除！")
                    st.rerun()

    # --- 第 4 順位：管理兩層資料夾 ---
    with st.expander("⚙️ 管理兩層資料夾 (新增/改名)", expanded=False):
        tab1, tab2 = st.tabs(["第一層(主夾)", "第二層(子夾)"])
        with tab1:
            new_c1 = st.text_input("➕ 新增第一層資料夾：", placeholder="例如：06_職安衛與環保")
            if st.button("新增主資料夾", use_container_width=True) and new_c1.strip():
                clean_c1 = new_c1.strip()
                if clean_c1 not in folder_tree:
                    folder_tree[clean_c1] = ["一般文件"]
                    save_db(db)
                    st.rerun()
            st.divider()
            target_c1 = st.selectbox("選擇要改名或刪除的主資料夾：", list(folder_tree.keys()), key="sel_c1")
            rename_c1 = st.text_input("修改主資料夾名稱：", value=target_c1, key="ren_c1")
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("儲存主夾名稱", use_container_width=True):
                    clean_r1 = rename_c1.strip()
                    if clean_r1 and clean_r1 != target_c1:
                        new_tree = { (clean_r1 if k == target_c1 else k): v for k, v in folder_tree.items() }
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
                        fb1 = list(folder_tree.keys())[0]
                        fb2 = folder_tree[fb1][0]
                        for c in chunks:
                            if c.get("cat1") == target_c1:
                                c["cat1"], c["cat2"] = fb1, fb2
                        save_db(db)
                        st.rerun()
        with tab2:
            parent_c1 = st.selectbox("1. 選擇所屬主資料夾：", list(folder_tree.keys()), key="p_c1")
            new_c2 = st.text_input("➕ 新增第二層子資料夾：", placeholder="例如：高空作業查驗")
            if st.button("新增子資料夾", use_container_width=True) and new_c2.strip():
                clean_c2 = new_c2.strip()
                if clean_c2 not in folder_tree[parent_c1]:
                    folder_tree[parent_c1].append(clean_c2)
                    save_db(db)
                    st.rerun()
            st.divider()
            target_c2 = st.selectbox("2. 選擇要改名或刪除的子夾：", folder_tree[parent_c1], key="sel_c2")
            rename_c2 = st.text_input("修改子資料夾名稱：", value=target_c2, key="ren_c2")
            col_c, col_d = st.columns(2)
            with col_c:
                if st.button("儲存子夾名稱", use_container_width=True):
                    clean_r2 = rename_c2.strip()
                    if clean_r2 and clean_r2 != target_c2:
                        idx = folder_tree[parent_c1].index(target_c2)
                        folder_tree[parent_c1][idx] = clean_r2
                        for c in chunks:
                            if c.get("cat1") == parent_c1 and c.get("cat2") == target_c2:
                                c["cat2"] = clean_r2
                        save_db(db)
                        st.rerun()
            with col_d:
                if st.button("🗑️ 刪除子夾", use_container_width=True):
                    if len(folder_tree[parent_c1]) > 1:
                        folder_tree[parent_c1].remove(target_c2)
                        fb_sub = folder_tree[parent_c1][0]
                        for c in chunks:
                            if c.get("cat1") == parent_c1 and c.get("cat2") == target_c2:
                                c["cat2"] = fb_sub
                        save_db(db)
                        st.rerun()

    # --- 第 5 順位：組員帳號管理 (新增/改密碼/刪除組員) ---
    with st.expander("👥 組員帳號管理 (新增/修改)", expanded=False):
        st.caption("可為每位組員建立專屬姓名與密碼：")
        new_u_name = st.text_input("➕ 新增組員姓名：", placeholder="例如：王小明")
        new_u_pwd = st.text_input("設定登入密碼：", value=TEAM_PASSWORD, key="new_u_pwd")
        new_u_role = st.selectbox("權限角色：", ["一般組員", "管理員"])
        if st.button("新增組員帳號", use_container_width=True) and new_u_name.strip():
            u_clean = new_u_name.strip()
            users_db[u_clean] = {"password": new_u_pwd.strip() or TEAM_PASSWORD, "role": new_u_role}
            save_db(db)
            st.success(f"已新增組員：{u_clean}")
            st.rerun()
            
        st.divider()
        edit_u = st.selectbox("選擇要修改或刪除的組員：", list(users_db.keys()))
        edit_pwd = st.text_input("修改該組員密碼：", value=users_db[edit_u]["password"], key="ed_pwd")
        u_col1, u_col2 = st.columns(2)
        with u_col1:
            if st.button("更新密碼", use_container_width=True):
                users_db[edit_u]["password"] = edit_pwd.strip() or TEAM_PASSWORD
                save_db(db)
                st.success("密碼已更新！")
                st.rerun()
        with u_col2:
            if st.button("🗑️ 刪除組員", use_container_width=True):
                if len(users_db) > 1 and edit_u != current_user:
                    del users_db[edit_u]
                    save_db(db)
                    st.rerun()
                else:
                    st.warning("無法刪除自己或最後一位成員！")

    st.divider()
    enable_web_search = st.checkbox("🌐 啟用 Google 即時聯網搜尋", value=False)
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "rb") as f:
            st.download_button(
                label="📥 下載知識庫備份檔 (.pkl)",
                data=f,
                file_name="engineering_knowledge_db.pkl",
                mime="application/octet-stream",
                use_container_width=True
            )

# ================= 6. 主畫面：三大功能分頁 =================
st.title("🏗️ 工程文件問答 ＆ 內部檔案搜尋系統")

with st.expander("🎯 點此篩選要檢索的「兩層資料夾範圍」（預設搜尋全部資料夾）", expanded=False):
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        selected_cat1s = st.multiselect(
            "1. 篩選第一層（主資料夾）：",
            options=list(folder_tree.keys()),
            default=list(folder_tree.keys())
        )
    with col_f2:
        available_cat2s = []
        for c1 in selected_cat1s:
            available_cat2s.extend(folder_tree.get(c1, []))
        selected_cat2s = st.multiselect(
            "2. 篩選第二層（子資料夾，留空代表全選）：",
            options=available_cat2s,
            default=[]
        )

main_tab1, main_tab2, main_tab3 = st.tabs([
    "💬 AI 智囊團問答", 
    "🔎 內部檔案搜尋引擎 (關鍵字秒查)", 
    "📂 檔案全文檢視器"
])

# ----------------- 分頁 1：AI 智囊團問答 -----------------
with main_tab1:
    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "assistant", "content": f"**{current_user}** 你好！我是你們的工程專案智囊團。你可以直接問我**簡報內容、預算金額、施工規範或採購法規**；若想單純用關鍵字翻找檔案原文，請切換到上方第二個**「🔎 內部檔案搜尋引擎」**分頁喔！"}
        ]

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if user_query := st.chat_input(f"以【{current_user}】身分提問（例如：台中二期議題討論會提到的重點有哪些？）"):
        st.session_state.messages.append({"role": "user", "content": user_query})
        with st.chat_message("user"):
            st.markdown(user_query)
            
        with st.chat_message("assistant"):
            with st.spinner("🔍 正在翻閱工程資料庫並整理答案..."):
                internal_context = search_internal_db(user_query, db["chunks"], selected_cat1s, selected_cat2s, top_k=10)
                file_summary = [
                    f"[{c1} > {c2}] {fname}" 
                    for fname, (c1, c2, _) in file_info_map.items() 
                    if c1 in selected_cat1s and (not selected_cat2s or c2 in selected_cat2s)
                ]
                
                system_prompt = f"""
你是一位專業、嚴謹的「工程專案智囊團 AI 顧問」。目前提問的團隊成員是：{current_user}。
請根據以下【內部工程資料庫檢索結果】以及你的營建與政府採購法規知識來回答問題。

【目前檢索範圍內的文件清單】：
{", ".join(file_summary) if file_summary else "⚠️ 尚未匯入任何文件"}

【回答嚴格守則】：
1. 內部資料引用：凡是引用內部工程資料的數據或規定，必須明確標註「【路徑：主資料夾 > 子資料夾｜來源：檔名與頁碼/投影片頁數】」。
2. 分類與新舊版本權重：『03_設計變更與核定公文』效力高於『01_契約與補充規定』；會議簡報請以檔名日期最新者為準。
3. 法規說明：若問題涉及政府採購法、公共工程委員會規範、建築技術規則或職安法規，請一併提供相關法規依據與條號。
4. 嚴禁瞎掰：若內部資料庫沒有提到該項工程數據，請誠實告知「目前內部文件未載明」，切勿自行編造。

=== 內部工程資料庫檢索結果 ===
{internal_context}
==============================

使用者問題：{user_query}
"""
                client = genai.Client(api_key=GEMINI_API_KEY)
                answer = None
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
                        if enable_web_search and "429" in err_str:
                            try:
                                response = client.models.generate_content(
                                    model=model_name,
                                    contents=system_prompt,
                                    config=types.GenerateContentConfig(temperature=0.2)
                                )
                                answer = response.text + "\n\n*(💡 系統提示：因免費版即時搜尋配額受限，已自動改用「內部工程資料庫 ＋ AI 內建法規知識」回答)*"
                                break
                            except Exception as e2:
                                last_error = str(e2)
                        continue
                        
                if not answer:
                    answer = f"⚠️ 系統呼叫失敗。錯誤訊息：{last_error}"
                    
                st.markdown(answer)
                st.session_state.messages.append({"role": "assistant", "content": answer})

# ----------------- 分頁 2：專屬內部檔案搜尋引擎 -----------------
with main_tab2:
    st.subheader("🔎 內部工程檔案專用搜尋引擎")
    st.caption("不需消耗 AI 額度，0.1 秒直接搜尋所有 PDF、Word、PPT、Excel 內文與檔名，並將關鍵字螢光筆標示！")
    
    col_s1, col_s2 = st.columns([3, 1])
    with col_s1:
        search_kw = st.text_input(
            "請輸入搜尋關鍵字（多個關鍵字可用空白隔開，例如：台中二期 水工 進度）：", 
            placeholder="輸入關鍵字後按 Enter 即可立即搜尋..."
        )
    with col_s2:
        match_mode = st.radio("多關鍵字匹配模式：", ["包含任一關鍵字 (OR)", "必須包含全部 (AND)"], horizontal=False)

    if search_kw.strip():
        keywords = [k.strip() for k in search_kw.strip().split() if k.strip()]
        matched_results = []
        
        for c in chunks:
            if c.get("cat1") not in selected_cat1s:
                continue
            if selected_cat2s and c.get("cat2") not in selected_cat2s:
                continue
                
            target_text = f"{c['file_name']} {c['content']}".lower()
            if match_mode == "必須包含全部 (AND)":
                is_match = all(k.lower() in target_text for k in keywords)
            else:
                is_match = any(k.lower() in target_text for k in keywords)
                
            if is_match:
                hit_count = sum(target_text.count(k.lower()) for k in keywords)
                matched_results.append((hit_count, c))
                
        matched_results.sort(key=lambda x: x[0], reverse=True)
        
        st.success(f"🎯 共找到 **{len(matched_results)}** 個包含關鍵字【{'、'.join(keywords)}】的文件段落！")
        
        for idx, (hits, item) in enumerate(matched_results, 1):
            with st.container(border=True):
                st.markdown(
                    f"**#{idx}｜🗂️ `{item['cat1']} > {item['cat2']}` ｜ 📄 `{item['file_name']}` {item['location']}** "
                    f"*(命中 {hits} 次｜上傳者：{item.get('uploader', '系統')})*"
                )
                highlighted_html = highlight_keywords(item["content"], keywords).replace("\n", "<br>")
                st.markdown(f"<div style='font-size: 0.95rem; line-height: 1.6;'>{highlighted_html}</div>", unsafe_allow_html=True)
    else:
        st.info("👆 請在上方輸入任何關鍵字（例如：工程名稱、材料規格、文號、金額），系統會立刻幫你把所有相關頁數翻出來！")

# ----------------- 分頁 3：檔案全文檢視器 -----------------
with main_tab3:
    st.subheader("📂 已匯入工程檔案全文檢視")
    if not file_info_map:
        st.warning("目前資料庫中尚無檔案，請從左側邊欄最上方上傳檔案。")
    else:
        view_file = st.selectbox("選擇要直接閱讀全文的檔案：", sorted(list(file_info_map.keys())))
        v_c1, v_c2, v_up = file_info_map[view_file]
        st.caption(f"目前所在資料夾路徑：**{v_c1} > {v_c2}** ｜ 上傳者：**{v_up}**")
        
        file_chunks = [c for c in chunks if c["file_name"] == view_file]
        for fc in file_chunks:
            with st.expander(f"📄 {fc['file_name']} — {fc['location']}", expanded=True):
                st.text(fc["content"])
