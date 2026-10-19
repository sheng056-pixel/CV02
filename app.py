import os
import io
import re
import hashlib
import pickle
import streamlit as st
from google import genai
from google.genai import types
import pypdf
import docx
from pptx import Presentation
import openpyxl
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# ================= 1. 系統基本設定與緊湊型 UI 樣式 =================
st.set_page_config(page_title="CV02 工程助手", page_icon="🏗️", layout="wide")

# 讓左側目錄按鈕與摺疊選單更緊湊俐落
st.markdown("""
<style>
div[data-testid="stSidebar"] div.stButton > button {
    padding-top: 0.15rem;
    padding-bottom: 0.15rem;
    font-size: 0.85rem;
    text-align: left;
    justify-content: flex-start;
}
</style>
""", unsafe_allow_html=True)

TEAM_PASSWORD = st.secrets.get("TEAM_PASSWORD", "1234")
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "").strip()
DB_FILE = "engineering_knowledge_db.pkl"

# 最新 Gemini 3 系列模型優先順序清單
AVAILABLE_MODELS = ["gemini-3.5-flash-lite", "gemini-3.8-flash", "gemini-3.5-flash"]

# 預設三層資料夾結構：{ 主夾(c1): { 次夾(c2): [細夾(c3), ...] } }
DEFAULT_FOLDER_TREE_V7 = {
    "01_招標文件與契約": {
        "第一次公告版本": ["契約條款", "投標與評選須知", "施工規範與圖說"],
        "第二次公告版本": ["契約條款", "投標與評選須知", "補充規定与附件"]
    },
    "02_施工規範與圖說": {
        "土建與水工": ["一般規範", "設計圖說"],
        "機電與弱電": ["設備規範", "界面圖說"],
        "職安衛與環保": ["職安衛規定", "環境保護計畫"]
    },
    "03_設計變更與核定公文": {
        "契約變更設計": ["變更設計圖說", "新增單價議價"],
        "重要往來公文": ["業主來文", "監造與承商發文"]
    },
    "04_會議紀錄與簡報": {
        "專案進度簡報": ["週會與月會簡報", "專案議題討論會"],
        "工務會議紀錄": ["歷次會議紀錄", "現勘結論"]
    },
    "05_材料送審與查驗紀錄": {
        "施工日報與查驗": ["施工日報表", "現場自主檢查表"],
        "材料送審資料": ["材料送審核定", "試驗報告"]
    }
}

DEFAULT_USERS = {
    "聖引": {"password": TEAM_PASSWORD, "role": "管理員"},
    "組員A": {"password": TEAM_PASSWORD, "role": "一般組員"}
}

DEFAULT_SHORTCUTS = [
    "📌 最新進度與預算",
    "⏳ 關鍵要徑與浮時",
    "📅 鋼筋工項施工日期統計",
    "📋 歷次會議決議追蹤",
    "⚠️ 職安衛查驗標準"
]

ENGINEERING_SYNONYMS = {
    "浮實": ["浮時", "總浮時", "自由浮時", "浮動", "Total Float", "TF"],
    "浮時": ["浮實", "總浮時", "自由浮時", "浮動", "Total Float", "TF"],
    "要徑": ["關鍵路徑", "關鍵要徑", "Critical Path", "CP", "進度", "落後", "趲趕"],
    "關鍵路徑": ["要徑", "關鍵要徑", "浮時"],
    "決議": ["結論", "裁示", "會議紀錄", "討論會", "拍板", "定案"],
    "預算": ["經費", "金額", "契約價金", "總價", "發包", "變更設計"],
    "施工日期": ["日報", "週報", "出工", "施作", "綁紮", "澆置", "開挖", "日曆天"],
    "里程碑": ["分項", "節點", "竣工", "期限", "工期"]
}

def make_token(username, password):
    raw = f"{username}:{password}:cv02_salt"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]

# ================= 2. 資料庫讀取與自動無痛升級三層結構 =================
def ensure_three_level_tree(raw_tree):
    new_tree = {}
    for c1, val in raw_tree.items():
        if isinstance(val, dict):
            new_tree[c1] = {}
            for c2, c3_list in val.items():
                new_tree[c1][c2] = list(c3_list) if c3_list else ["預設細項"]
        elif isinstance(val, list):
            new_tree[c1] = {}
            for c2 in val:
                new_tree[c1][c2] = ["一般文件", "補充附件"]
    if not new_tree:
        return DEFAULT_FOLDER_TREE_V7.copy()
    return new_tree

def load_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "rb") as f:
                data = pickle.load(f)
            if isinstance(data, dict):
                raw_tree = data.get("folder_tree", DEFAULT_FOLDER_TREE_V7.copy())
                tree_3d = ensure_three_level_tree(raw_tree)
                chunks = data.get("chunks", [])
                for c in chunks:
                    c1 = c.get("cat1", list(tree_3d.keys())[0])
                    if c1 not in tree_3d:
                        tree_3d[c1] = {"預設次夾": ["一般文件"]}
                    c2 = c.get("cat2", list(tree_3d[c1].keys())[0])
                    if c2 not in tree_3d[c1]:
                        tree_3d[c1][c2] = ["一般文件"]
                    if not tree_3d[c1][c2]:
                        tree_3d[c1][c2] = ["一般文件"]
                    c3 = c.get("cat3", tree_3d[c1][c2][0])
                    if c3 not in tree_3d[c1][c2]:
                        tree_3d[c1][c2].append(c3)
                    c["cat1"], c["cat2"], c["cat3"] = c1, c2, c3
                    if "uploader" not in c:
                        c["uploader"] = "聖引"
                return {
                    "folder_tree": tree_3d,
                    "chunks": chunks,
                    "users": data.get("users", DEFAULT_USERS.copy()),
                    "shortcuts": data.get("shortcuts", DEFAULT_SHORTCUTS.copy()),
                    "max_mb": data.get("max_mb", 80)
                }
        except Exception:
            pass
    return {
        "folder_tree": DEFAULT_FOLDER_TREE_V7.copy(),
        "chunks": [],
        "users": DEFAULT_USERS.copy(),
        "shortcuts": DEFAULT_SHORTCUTS.copy(),
        "max_mb": 80
    }

def save_db(db_dict):
    with open(DB_FILE, "wb") as f:
        pickle.dump(db_dict, f)

db = load_db()
users_db = db["users"]

# ================= 3. 組員登入 & 同裝置自動記憶 =================
if "current_user" not in st.session_state:
    st.session_state.current_user = None

q_user = st.query_params.get("user", "")
q_token = st.query_params.get("token", "")

if not st.session_state.current_user and q_user in users_db:
    expected_token = make_token(q_user, users_db[q_user]["password"])
    if q_token == expected_token:
        st.session_state.current_user = q_user

if not st.session_state.current_user:
    st.title("🏗️ CV02 工程助手 — 登入")
    st.caption("勾選「記住我」後將網頁加入手機主畫面或書籤，下次點開即可直接進入系統！")
    with st.container(border=True):
        login_user = st.selectbox("👤 選擇帳號：", list(users_db.keys()))
        login_pwd = st.text_input("🔑 輸入密碼：", type="password")
        remember_me = st.checkbox("📌 在此裝置記住我的登入狀態", value=True)
        if st.button("🚀 登入", type="primary", use_container_width=True):
            valid_pwd = users_db[login_user]["password"]
            if login_pwd == valid_pwd or login_pwd == TEAM_PASSWORD:
                st.session_state.current_user = login_user
                if remember_me:
                    st.query_params["user"] = login_user
                    st.query_params["token"] = make_token(login_user, valid_pwd)
                st.rerun()
            else:
                st.error("密碼錯誤！")
    st.stop()

current_user = st.session_state.current_user

# ================= 4. 匯入成功彈出式通知視窗 (Modal Dialog) =================
@st.dialog("🎉 檔案匯入成功！")
def show_upload_success_dialog(info):
    st.success(f"✅ 已成功將 **{info['count']}** 份檔案寫入工程資料庫！")
    st.markdown(f"**📂 歸檔三層路徑：** `{info['path']}`")
    if info.get("ocr_pages", 0) > 0:
        st.info(f"👁️ 內含 **{info['ocr_pages']}** 頁掃描影像，已自動透過 AI 視覺辨識轉為可搜尋文字！")
    st.markdown("**📄 本次匯入檔案清單：**")
    for fn in info.get("files", []):
        st.markdown(f"- `{fn}`")
    if st.button("👌 我知道了", type="primary", use_container_width=True):
        st.session_state.upload_notice = None
        st.rerun()

if st.session_state.get("upload_notice"):
    show_upload_success_dialog(st.session_state.upload_notice)

# ================= 5. 視覺 OCR 辨識與文件解析引擎 (含三層路徑 & 表頭保留) =================
def ocr_with_gemini(file_bytes, mime_type):
    if not GEMINI_API_KEY:
        return ""
    client = genai.Client(api_key=GEMINI_API_KEY)
    prompt = (
        "請完整、精準地辨識並轉錄這份工程公文、契約、圖表或照片中的所有繁體中文文字、文號、日期、表格數據與印章文字。"
        "若有表格請以『欄位1 | 欄位2』格式呈現，不要遺漏任何細節。"
    )
    for model_name in AVAILABLE_MODELS:
        try:
            resp = client.models.generate_content(
                model=model_name,
                contents=[
                    types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
                    prompt
                ],
                config=types.GenerateContentConfig(temperature=0.1)
            )
            if resp and resp.text:
                return resp.text.strip()
        except Exception:
            continue
    return ""

def extract_text_from_file(uploaded_file, cat1, cat2, cat3, uploader, enable_ocr=True):
    text_chunks = []
    filename = uploaded_file.name
    lower_name = filename.lower()
    raw_bytes = uploaded_file.getvalue()

    if lower_name.endswith((".jpg", ".jpeg", ".png")):
        mime = "image/png" if lower_name.endswith(".png") else "image/jpeg"
        ocr_text = ocr_with_gemini(raw_bytes, mime) if enable_ocr else ""
        if ocr_text:
            text_chunks.append({
                "file_name": filename, "cat1": cat1, "cat2": cat2, "cat3": cat3,
                "location": "(影像OCR辨識)", "content": ocr_text, "uploader": uploader
            })
        return text_chunks, (1 if ocr_text else 0)

    elif lower_name.endswith(".pdf"):
        reader = pypdf.PdfReader(io.BytesIO(raw_bytes))
        scan_pages = []
        for i, page in enumerate(reader.pages):
            page_text = (page.extract_text() or "").strip()
            if len(page_text) >= 25:
                text_chunks.append({
                    "file_name": filename, "cat1": cat1, "cat2": cat2, "cat3": cat3,
                    "location": f"(第 {i+1} 頁)", "content": page_text, "uploader": uploader
                })
            else:
                scan_pages.append(i)

        ocr_count = 0
        if scan_pages and enable_ocr:
            max_ocr_pages = min(len(scan_pages), 25)
            prog = st.progress(0, text=f"🔍 偵測到《{filename}》含掃描頁，正在啟動 AI 光學辨識 (共 {max_ocr_pages} 頁)...")
            for idx, page_num in enumerate(scan_pages[:max_ocr_pages]):
                writer = pypdf.PdfWriter()
                writer.add_page(reader.pages[page_num])
                buf = io.BytesIO()
                writer.write(buf)
                ocr_res = ocr_with_gemini(buf.getvalue(), "application/pdf")
                if ocr_res:
                    text_chunks.append({
                        "file_name": filename, "cat1": cat1, "cat2": cat2, "cat3": cat3,
                        "location": f"(第 {page_num+1} 頁-OCR掃描辨識)", "content": ocr_res, "uploader": uploader
                    })
                    ocr_count += 1
                prog.progress((idx + 1) / max_ocr_pages, text=f"🔍 正在辨識《{filename}》第 {page_num+1} 頁 ({idx+1}/{max_ocr_pages})...")
            prog.empty()
        return text_chunks, ocr_count

    elif lower_name.endswith(".docx"):
        doc = docx.Document(io.BytesIO(raw_bytes))
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            header_str = ""
            for r_idx, row in enumerate(table.rows):
                row_cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
                row_text = " | ".join([c for c in row_cells if c])
                if not row_text:
                    continue
                if r_idx == 0:
                    header_str = row_text
                    paragraphs.append(f"【表格表頭：{header_str}】")
                else:
                    paragraphs.append(f"[{header_str}] ➔ {row_text}" if header_str else row_text)
        full_text = "\n".join(paragraphs)
        for i in range(0, len(full_text), 900):
            if full_text[i:i+900].strip():
                text_chunks.append({
                    "file_name": filename, "cat1": cat1, "cat2": cat2, "cat3": cat3,
                    "location": f"(段落 {i//900 + 1})", "content": full_text[i:i+900], "uploader": uploader
                })

    elif lower_name.endswith(".pptx"):
        prs = Presentation(io.BytesIO(raw_bytes))
        for i, slide in enumerate(prs.slides):
            slide_texts = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.strip():
                            slide_texts.append(paragraph.text.strip())
                if shape.has_table:
                    header_str = ""
                    for r_idx, row in enumerate(shape.table.rows):
                        row_cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
                        row_text = " | ".join([c for c in row_cells if c])
                        if not row_text:
                            continue
                        if r_idx == 0:
                            header_str = row_text
                            slide_texts.append(f"【表頭：{header_str}】")
                        else:
                            slide_texts.append(f"[{header_str}] ➔ {row_text}" if header_str else row_text)
            if slide_texts:
                text_chunks.append({
                    "file_name": filename, "cat1": cat1, "cat2": cat2, "cat3": cat3,
                    "location": f"(投影片第 {i+1} 頁)", "content": "\n".join(slide_texts), "uploader": uploader
                })

    elif lower_name.endswith(".xlsx"):
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            rows_text = []
            header_str = ""
            for row in ws.iter_rows(values_only=True):
                row_vals = [str(cell).strip().replace("\n", " ") for cell in row if cell is not None and str(cell).strip()]
                if not row_vals:
                    continue
                line = " | ".join(row_vals)
                if not header_str:
                    header_str = line
                    rows_text.append(f"【工作表：{sheet}｜表頭欄位：{header_str}】")
                else:
                    rows_text.append(f"[{sheet}] {line}")
            for i in range(0, len(rows_text), 25):
                batch = rows_text[i:i+25]
                if i > 0 and header_str:
                    batch.insert(0, f"【表頭欄位：{header_str}】")
                chunk_str = "\n".join(batch)
                if chunk_str.strip():
                    text_chunks.append({
                        "file_name": filename, "cat1": cat1, "cat2": cat2, "cat3": cat3,
                        "location": f"(工作表:{sheet} 第{i+1}~{min(i+25, len(rows_text))}列)",
                        "content": chunk_str, "uploader": uploader
                    })

    elif lower_name.endswith(".txt"):
        full_text = raw_bytes.decode("utf-8", errors="ignore")
        for i in range(0, len(full_text), 900):
            if full_text[i:i+900].strip():
                text_chunks.append({
                    "file_name": filename, "cat1": cat1, "cat2": cat2, "cat3": cat3,
                    "location": f"(段落 {i//900 + 1})", "content": full_text[i:i+900], "uploader": uploader
                })

    return text_chunks, 0

# ================= 6. 智慧檢索演算法 (含同義詞擴充 & 全量統計模式) =================
def expand_query_with_synonyms(query):
    expanded = query
    for key, syns in ENGINEERING_SYNONYMS.items():
        if key in query:
            expanded += " " + " ".join(syns)
    return expanded

def search_internal_db(query, chunks, selected_cat1s, selected_cat2s, selected_cat3s):
    if not chunks:
        return "（目前內部目錄尚無檔案，請先從左側『➕ 上傳』新增檔案）"

    filtered_chunks = [
        c for c in chunks
        if c.get("cat1") in selected_cat1s
        and (not selected_cat2s or c.get("cat2") in selected_cat2s)
        and (not selected_cat3s or c.get("cat3") in selected_cat3s)
    ]
    if not filtered_chunks:
        return "（目前勾選的資料夾範圍內尚無文件）"

    deep_keywords = ["日期", "天數", "幾天", "所有", "全部", "統計", "歷次", "決議", "何時", "什麼時候", "要徑", "浮時", "浮實", "日報", "里程碑", "分項"]
    is_deep_search = any(k in query for k in deep_keywords)
    top_k = 100 if is_deep_search else 20

    expanded_query = expand_query_with_synonyms(query)
    corpus = [f"{item['file_name']} {item['cat2']} {item.get('cat3', '')} {item['content']}" for item in filtered_chunks]
    vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
    tfidf_matrix = vectorizer.fit_transform(corpus + [expanded_query])
    cosine_sim = cosine_similarity(tfidf_matrix[-1], tfidf_matrix[:-1]).flatten()
    top_indices = cosine_sim.argsort()[-top_k:][::-1]

    selected_items = []
    for idx in top_indices:
        if cosine_sim[idx] > 0.008:
            selected_items.append(filtered_chunks[idx])

    if not selected_items:
        selected_items = filtered_chunks[:min(12, len(filtered_chunks))]

    if is_deep_search:
        selected_items.sort(key=lambda x: (x["file_name"], x["location"]))

    results = [
        f"【目錄：{item['cat1']} > {item['cat2']} > {item.get('cat3', '一般')}｜來源：{item['file_name']} {item['location']}】\n{item['content']}"
        for item in selected_items
    ]
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

# ================= 7. 側邊欄：三層可收合目錄與上傳自動清空 =================
folder_tree = db["folder_tree"]
chunks = db["chunks"]
max_mb_limit = db.get("max_mb", 80)

# 整理檔案三層對照表: { file_name: (c1, c2, c3, uploader) }
file_info_map = {}
for c in chunks:
    c1 = c.get("cat1", list(folder_tree.keys())[0])
    if c1 not in folder_tree:
        folder_tree[c1] = {"預設次夾": ["一般文件"]}
    c2 = c.get("cat2", list(folder_tree[c1].keys())[0])
    if c2 not in folder_tree[c1]:
        folder_tree[c1][c2] = ["一般文件"]
    c3 = c.get("cat3", folder_tree[c1][c2][0])
    if c3 not in folder_tree[c1][c2]:
        folder_tree[c1][c2].append(c3)
    file_info_map[c["file_name"]] = (c1, c2, c3, c.get("uploader", "系統"))

db_size_mb = os.path.getsize(DB_FILE) / (1024 * 1024) if os.path.exists(DB_FILE) else 0.0

if "jump_to_file" not in st.session_state:
    st.session_state.jump_to_file = None
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

with st.sidebar:
    col_u1, col_u2 = st.columns([3, 2])
    with col_u1:
        st.markdown(f"👤 **{current_user}**")
    with col_u2:
        if st.button("🚪 登出", use_container_width=True):
            st.session_state.current_user = None
            st.query_params.clear()
            st.rerun()

    st.caption(f"📊 檔案：**{len(file_info_map)}** 份 ｜ 容量：**{db_size_mb:.2f} MB**")

    # --- 1. ➕ 上傳 (含自動清空佇列與彈窗通知) ---
    with st.expander("➕ 上傳", expanded=False):
        up_c1 = st.selectbox("1. 第一層 (主夾)：", list(folder_tree.keys()))
        up_c2 = st.selectbox("2. 第二層 (次夾)：", list(folder_tree[up_c1].keys()))
        up_c3 = st.selectbox("3. 第三層 (細夾)：", folder_tree[up_c1][up_c2])
        enable_ocr_upload = st.checkbox("👁️ 掃描檔自動 AI 光學辨識 (OCR)", value=True)
        uploaded_files = st.file_uploader(
            f"4. 選擇檔案 (上限 {max_mb_limit}MB)",
            type=["pdf", "docx", "pptx", "xlsx", "txt", "jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key=f"file_up_{st.session_state.uploader_key}"
        )
        if st.button("🚀 匯入資料庫", type="primary", use_container_width=True):
            if not uploaded_files:
                st.warning("請先選擇檔案！")
            else:
                oversized = [uf.name for uf in uploaded_files if (uf.size / (1024 * 1024)) > max_mb_limit]
                img_count = sum(1 for uf in uploaded_files if uf.name.lower().endswith((".jpg", ".jpeg", ".png")))
                if oversized:
                    st.error(f"⚠️ 超過 {max_mb_limit}MB 上限：\n" + "、".join(oversized))
                elif img_count > 10:
                    st.warning(f"⚠️ 一次選擇了 {img_count} 張圖片，建議單次 10 張以內！")
                else:
                    with st.spinner("正在解析內容並建立索引..."):
                        imported_names, total_ocr = [], 0
                        for uf in uploaded_files:
                            db["chunks"] = [c for c in db["chunks"] if c["file_name"] != uf.name]
                            new_chunks, ocr_cnt = extract_text_from_file(uf, up_c1, up_c2, up_c3, current_user, enable_ocr=enable_ocr_upload)
                            if new_chunks:
                                db["chunks"].extend(new_chunks)
                                imported_names.append(uf.name)
                                total_ocr += ocr_cnt
                        save_db(db)
                        # 自動遞增 key 清空上傳框，並設定彈出通知視窗
                        st.session_state.uploader_key += 1
                        st.session_state.upload_notice = {
                            "count": len(imported_names),
                            "files": imported_names,
                            "path": f"{up_c1} > {up_c2} > {up_c3}",
                            "ocr_pages": total_ocr
                        }
                        st.rerun()

    # --- 2. 📁 目錄 (三層可逐級展開/收合樹狀結構) ---
    with st.expander("📁 目錄", expanded=True):
        if not file_info_map:
            st.caption("目前尚無檔案，請點上方「➕ 上傳」。")
        else:
            st.caption("💡 點擊資料夾可展開/收合，點檔名直接閱讀：")
            for c1, c2_dict in folder_tree.items():
                c1_files = [f for f, (fc1, _, _, _) in file_info_map.items() if fc1 == c1]
                if not c1_files:
                    continue
                # 第一層可收合
                with st.expander(f"🗂️ {c1} ({len(c1_files)})", expanded=False):
                    for c2, c3_list in c2_dict.items():
                        c2_files = [f for f, (fc1, fc2, _, _) in file_info_map.items() if fc1 == c1 and fc2 == c2]
                        if not c2_files:
                            continue
                        # 第二層可收合
                        with st.expander(f"📂 {c2} ({len(c2_files)})", expanded=False):
                            for c3 in c3_list:
                                c3_files = sorted([f for f, (fc1, fc2, fc3, _) in file_info_map.items() if fc1 == c1 and fc2 == c2 and fc3 == c3])
                                if not c3_files:
                                    continue
                                st.markdown(f"**🏷️ {c3}** `({len(c3_files)})`")
                                for f in c3_files:
                                    if st.button(f"📄 {f}", key=f"dir_{c1}_{c2}_{c3}_{f}", use_container_width=True):
                                        st.session_state.jump_to_file = f

    # --- 3. ⚙️ 三層資料夾與防呆設定 ---
    with st.expander("⚙️ 資料夾設定", expanded=False):
        t1, t2, t3, t4 = st.tabs(["第1層", "第2層", "第3層", "防呆/快捷"])
        with t1:
            new_c1 = st.text_input("➕ 新增第1層主夾：")
            if st.button("新增主夾", use_container_width=True) and new_c1.strip():
                if new_c1.strip() not in folder_tree:
                    folder_tree[new_c1.strip()] = {"預設次夾": ["一般文件"]}
                    save_db(db)
                    st.rerun()
            target_c1 = st.selectbox("選擇第1層主夾：", list(folder_tree.keys()), key="s_c1")
            rename_c1 = st.text_input("改名主夾：", value=target_c1)
            ca, cb = st.columns(2)
            with ca:
                if st.button("儲存主夾", use_container_width=True) and rename_c1.strip():
                    db["folder_tree"] = {(rename_c1.strip() if k == target_c1 else k): v for k, v in folder_tree.items()}
                    for c in chunks:
                        if c.get("cat1") == target_c1:
                            c["cat1"] = rename_c1.strip()
                    save_db(db)
                    st.rerun()
            with cb:
                if st.button("🗑️ 刪主夾", use_container_width=True) and len(folder_tree) > 1:
                    del folder_tree[target_c1]
                    fb1 = list(folder_tree.keys())[0]
                    fb2 = list(folder_tree[fb1].keys())[0]
                    fb3 = folder_tree[fb1][fb2][0]
                    for c in chunks:
                        if c.get("cat1") == target_c1:
                            c["cat1"], c["cat2"], c["cat3"] = fb1, fb2, fb3
                    save_db(db)
                    st.rerun()

        with t2:
            p_c1 = st.selectbox("所屬第1層主夾：", list(folder_tree.keys()), key="p2_c1")
            new_c2 = st.text_input("➕ 新增第2層次夾：")
            if st.button("新增次夾", use_container_width=True) and new_c2.strip():
                if new_c2.strip() not in folder_tree[p_c1]:
                    folder_tree[p_c1][new_c2.strip()] = ["一般文件"]
                    save_db(db)
                    st.rerun()
            target_c2 = st.selectbox("選擇第2層次夾：", list(folder_tree[p_c1].keys()), key="s_c2")
            rename_c2 = st.text_input("改名次夾：", value=target_c2)
            cc, cd = st.columns(2)
            with cc:
                if st.button("儲存次夾", use_container_width=True) and rename_c2.strip():
                    folder_tree[p_c1] = {(rename_c2.strip() if k == target_c2 else k): v for k, v in folder_tree[p_c1].items()}
                    for c in chunks:
                        if c.get("cat1") == p_c1 and c.get("cat2") == target_c2:
                            c["cat2"] = rename_c2.strip()
                    save_db(db)
                    st.rerun()
            with cd:
                if st.button("🗑️ 刪次夾", use_container_width=True) and len(folder_tree[p_c1]) > 1:
                    del folder_tree[p_c1][target_c2]
                    fb2 = list(folder_tree[p_c1].keys())[0]
                    fb3 = folder_tree[p_c1][fb2][0]
                    for c in chunks:
                        if c.get("cat1") == p_c1 and c.get("cat2") == target_c2:
                            c["cat2"], c["cat3"] = fb2, fb3
                    save_db(db)
                    st.rerun()

        with t3:
            p3_c1 = st.selectbox("所屬第1層：", list(folder_tree.keys()), key="p3_c1")
            p3_c2 = st.selectbox("所屬第2層：", list(folder_tree[p3_c1].keys()), key="p3_c2")
            new_c3 = st.text_input("➕ 新增第3層細夾：")
            if st.button("新增細夾", use_container_width=True) and new_c3.strip():
                if new_c3.strip() not in folder_tree[p3_c1][p3_c2]:
                    folder_tree[p3_c1][p3_c2].append(new_c3.strip())
                    save_db(db)
                    st.rerun()
            target_c3 = st.selectbox("選擇第3層細夾：", folder_tree[p3_c1][p3_c2], key="s_c3")
            rename_c3 = st.text_input("改名細夾：", value=target_c3)
            ce, cf = st.columns(2)
            with ce:
                if st.button("儲存細夾", use_container_width=True) and rename_c3.strip():
                    idx = folder_tree[p3_c1][p3_c2].index(target_c3)
                    folder_tree[p3_c1][p3_c2][idx] = rename_c3.strip()
                    for c in chunks:
                        if c.get("cat1") == p3_c1 and c.get("cat2") == p3_c2 and c.get("cat3") == target_c3:
                            c["cat3"] = rename_c3.strip()
                    save_db(db)
                    st.rerun()
            with cf:
                if st.button("🗑️ 刪細夾", use_container_width=True) and len(folder_tree[p3_c1][p3_c2]) > 1:
                    folder_tree[p3_c1][p3_c2].remove(target_c3)
                    fb3 = folder_tree[p3_c1][p3_c2][0]
                    for c in chunks:
                        if c.get("cat1") == p3_c1 and c.get("cat2") == p3_c2 and c.get("cat3") == target_c3:
                            c["cat3"] = fb3
                    save_db(db)
                    st.rerun()

        with t4:
            new_max_mb = st.slider("單檔防呆上限 (MB)：", 20, 200, int(max_mb_limit), 10)
            if new_max_mb != max_mb_limit:
                db["max_mb"] = new_max_mb
                save_db(db)
                st.rerun()
            sc_str = st.text_area("自訂首頁快捷鈕（一行一個）：", value="\n".join(db.get("shortcuts", DEFAULT_SHORTCUTS)))
            if st.button("儲存快捷鈕", use_container_width=True):
                db["shortcuts"] = [s.strip() for s in sc_str.split("\n") if s.strip()]
                save_db(db)
                st.rerun()

    # --- 4. 👥 帳號管理 (聖引專屬) ---
    if current_user == "聖引":
        with st.expander("👥 帳號管理 (聖引專屬)", expanded=False):
            new_u = st.text_input("➕ 新增組員姓名：")
            new_p = st.text_input("設定密碼：", value=TEAM_PASSWORD)
            if st.button("新增組員", use_container_width=True) and new_u.strip():
                users_db[new_u.strip()] = {"password": new_p.strip() or TEAM_PASSWORD, "role": "一般組員"}
                save_db(db)
                st.success(f"已新增：{new_u.strip()}")
                st.rerun()
            st.divider()
            ed_u = st.selectbox("選擇組員：", list(users_db.keys()))
            ed_p = st.text_input("修改密碼：", value=users_db[ed_u]["password"])
            u1, u2 = st.columns(2)
            with u1:
                if st.button("更新密碼", use_container_width=True):
                    users_db[ed_u]["password"] = ed_p.strip() or TEAM_PASSWORD
                    save_db(db)
                    st.success("已更新！")
                    st.rerun()
            with u2:
                if st.button("🗑️ 刪除帳號", use_container_width=True):
                    if ed_u != "聖引":
                        del users_db[ed_u]
                        save_db(db)
                        st.rerun()
                    else:
                        st.warning("最高管理員「聖引」不可刪除！")

    # --- 5. 💾 備份與還原 ---
    with st.expander("💾 備份與還原", expanded=False):
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as f:
                st.download_button("📥 下載備份 (.pkl)", data=f, file_name="engineering_knowledge_db.pkl", use_container_width=True)
        restore_file = st.file_uploader("📤 上傳備份檔還原：", type=["pkl"])
        if restore_file and st.button("⚡ 立即還原知識庫", use_container_width=True):
            with open(DB_FILE, "wb") as f:
                f.write(restore_file.getvalue())
            st.success("已還原！")
            st.rerun()

    enable_web_search = st.checkbox("🌐 Google 聯網搜尋", value=False)

# ================= 8. 主畫面：極簡頂部 & 三層範圍篩選 =================
col_title, col_filter = st.columns([3, 2])
with col_title:
    st.subheader("🏗️ CV02 工程助手")
with col_filter:
    with st.popover("🎯 篩選檢索目錄範圍（預設全選）", use_container_width=True):
        selected_cat1s = st.multiselect("1. 第一層 (主夾)：", list(folder_tree.keys()), default=list(folder_tree.keys()))
        avail_c2 = [c2 for c1 in selected_cat1s for c2 in folder_tree.get(c1, {}).keys()]
        selected_cat2s = st.multiselect("2. 第二層 (次夾，留空代表全選)：", avail_c2, default=[])
        avail_c3 = [
            c3 for c1 in selected_cat1s
            for c2, c3_list in folder_tree.get(c1, {}).items()
            if (not selected_cat2s or c2 in selected_cat2s)
            for c3 in c3_list
        ]
        selected_cat3s = st.multiselect("3. 第三層 (細夾，留空代表全選)：", list(dict.fromkeys(avail_c3)), default=[])

# 若從左側目錄點擊了某個檔案，顯示頂部快速閱覽卡片 (內建一鍵改名與換資料夾)
if st.session_state.jump_to_file and st.session_state.jump_to_file in file_info_map:
    j_file = st.session_state.jump_to_file
    jc1, jc2, jc3, jup = file_info_map[j_file]
    with st.container(border=True):
        jh1, jh2 = st.columns([5, 1])
        with jh1:
            st.markdown(f"### 📖 快速閱覽：`{j_file}`  `({jc1} > {jc2} > {jc3}｜上傳：{jup})`")
        with jh2:
            if st.button("✖️ 關閉閱覽", use_container_width=True):
                st.session_state.jump_to_file = None
                st.rerun()

        # 就地快速移動與改名面板
        with st.expander("✏️ 就地改名 / 換資料夾 / 刪除此檔案", expanded=False):
            q_col1, q_col2, q_col3, q_col4 = st.columns([2, 2, 2, 2])
            with q_col1:
                q_new_name = st.text_input("檔案名稱：", value=j_file, key="q_ren")
            with q_col2:
                c1_opts = list(folder_tree.keys())
                q_c1 = st.selectbox("移至第1層：", c1_opts, index=c1_opts.index(jc1) if jc1 in c1_opts else 0, key="q_c1")
            with q_col3:
                c2_opts = list(folder_tree[q_c1].keys())
                q_c2 = st.selectbox("移至第2層：", c2_opts, index=c2_opts.index(jc2) if jc2 in c2_opts else 0, key="q_c2")
            with q_col4:
                c3_opts = folder_tree[q_c1][q_c2]
                q_c3 = st.selectbox("移至第3層：", c3_opts, index=c3_opts.index(jc3) if jc3 in c3_opts else 0, key="q_c3")
            qb1, qb2 = st.columns(2)
            with qb1:
                if st.button("💾 儲存變更", key="q_save", type="primary", use_container_width=True):
                    clean_qn = q_new_name.strip() or j_file
                    for c in chunks:
                        if c["file_name"] == j_file:
                            c["file_name"] = clean_qn
                            c["cat1"], c["cat2"], c["cat3"] = q_c1, q_c2, q_c3
                    save_db(db)
                    st.session_state.jump_to_file = clean_qn
                    st.rerun()
            with qb2:
                if st.button("🗑️ 刪除此檔案", key="q_del", use_container_width=True):
                    db["chunks"] = [c for c in chunks if c["file_name"] != j_file]
                    save_db(db)
                    st.session_state.jump_to_file = None
                    st.rerun()

        for fc in [c for c in chunks if c["file_name"] == j_file]:
            with st.expander(f"📄 {fc['location']}", expanded=True):
                st.text(fc["content"])

# 建立四大主畫面分頁 (新增「📦 批次搬移與整理」)
tab_ai, tab_search, tab_batch, tab_read = st.tabs([
    "💬 AI 問答",
    "🔎 搜檔案",
    "📦 批次搬移與整理",
    "📖 讀全文"
])

# ----------------- 分頁 1：💬 AI 問答 (支援 Ctrl+V 貼圖/貼文字 & 隨附檔案) -----------------
with tab_ai:
    shortcuts = db.get("shortcuts", DEFAULT_SHORTCUTS)
    triggered_query = None
    if shortcuts:
        sc_cols = st.columns(len(shortcuts))
        for idx, sc_text in enumerate(shortcuts):
            if sc_cols[idx].button(sc_text, key=f"sc_{idx}", use_container_width=True):
                triggered_query = sc_text

    with st.expander("📋 直接複製貼上文字／表格 ＆ 📎 隨附臨時檔案問 AI（不弄亂永久目錄，亦可一鍵歸檔）", expanded=False):
        p_col1, p_col2 = st.columns(2)
        with p_col1:
            pasted_context = st.text_area(
                "📋 直接貼上臨時文字 / 公文草稿 / LINE 訊息 / 表格 (Ctrl+V)：",
                height=115,
                placeholder="在此直接 Ctrl+V 貼上任何文字或表格，下方提問時 AI 會自動納入比對..."
            )
        with p_col2:
            temp_files = st.file_uploader(
                "📎 或選擇臨時檔案 / 螢幕截圖 (亦可在最下方打字框直接 Ctrl+V 貼截圖)：",
                type=["pdf", "docx", "pptx", "xlsx", "txt", "jpg", "jpeg", "png"],
                accept_multiple_files=True,
                key="chat_temp_files"
            )

        if (pasted_context and pasted_context.strip()) or temp_files:
            st.markdown("##### 💾 覺得上面的臨時內容很重要？可一鍵轉存進永久目錄：")
            tc1, tc2, tc3, tc4, tc5 = st.columns([2, 2, 2, 2, 2])
            with tc1:
                save_t_c1 = st.selectbox("主夾：", list(folder_tree.keys()), key="st_c1")
            with tc2:
                save_t_c2 = st.selectbox("次夾：", list(folder_tree[save_t_c1].keys()), key="st_c2")
            with tc3:
                save_t_c3 = st.selectbox("細夾：", folder_tree[save_t_c1][save_t_c2], key="st_c3")
            with tc4:
                pasted_fname = st.text_input("若存貼上文字之檔名：", value="臨時貼上資料.txt", key="st_fn")
            with tc5:
                st.write("")
                if st.button("💾 存入永久目錄", type="primary", use_container_width=True):
                    saved_names = []
                    if pasted_context and pasted_context.strip():
                        fn_clean = pasted_fname.strip() or "臨時貼上資料.txt"
                        if not fn_clean.endswith(".txt"):
                            fn_clean += ".txt"
                        db["chunks"] = [c for c in db["chunks"] if c["file_name"] != fn_clean]
                        txt_val = pasted_context.strip()
                        for i in range(0, len(txt_val), 900):
                            db["chunks"].append({
                                "file_name": fn_clean, "cat1": save_t_c1, "cat2": save_t_c2, "cat3": save_t_c3,
                                "location": f"(段落 {i//900 + 1})", "content": txt_val[i:i+900], "uploader": current_user
                            })
                        saved_names.append(fn_clean)
                    if temp_files:
                        for tf in temp_files:
                            db["chunks"] = [c for c in db["chunks"] if c["file_name"] != tf.name]
                            new_c, _ = extract_text_from_file(tf, save_t_c1, save_t_c2, save_t_c3, current_user, enable_ocr=True)
                            db["chunks"].extend(new_c)
                            saved_names.append(tf.name)
                    save_db(db)
                    st.session_state.upload_notice = {
                        "count": len(saved_names),
                        "files": saved_names,
                        "path": f"{save_t_c1} > {save_t_c2} > {save_t_c3}",
                        "ocr_pages": 0
                    }
                    st.rerun()

    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "assistant", "content": f"**{current_user}** 您好！您可以在下方輸入框**直接打字、按 `Ctrl+V` 直接貼上螢幕截圖**，或展開上方**貼上大段文字/檔案**讓我幫您與資料庫交叉比對！"}
        ]

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # 支援原生直接貼圖 (Ctrl+V) 與夾帶檔案的對話輸入框 (若環境為舊版 Streamlit 則自動優雅降級)
    inline_files = []
    user_query = None
    try:
        chat_val = st.chat_input(
            "請輸入問題（支援在此直接 Ctrl+V 貼上截圖，或點左側 ➕ 夾帶檔案）...",
            accept_file="multiple",
            file_type=["pdf", "docx", "pptx", "xlsx", "txt", "jpg", "jpeg", "png"]
        )
        if chat_val:
            if isinstance(chat_val, str):
                user_query = chat_val
            else:
                user_query = getattr(chat_val, "text", "") or ""
                inline_files = list(getattr(chat_val, "files", []) or [])
                if not user_query and inline_files:
                    user_query = "請幫我分析隨附的圖片或檔案內容，並與內部工程資料庫比對。"
    except TypeError:
        chat_val = st.chat_input("請輸入問題（如：本案契約里程碑分項、鋼筋施工日期統計、關鍵要徑浮時...）")
        if chat_val:
            user_query = chat_val

    if triggered_query:
        user_query = triggered_query

    if user_query:
        all_temp_files = list(temp_files or []) + inline_files
        extra_tags = []
        if all_temp_files:
            extra_tags.append(f"📎 隨附 {len(all_temp_files)} 份檔案/截圖")
        if pasted_context and pasted_context.strip():
            extra_tags.append("📋 隨附貼上文字")
        display_q = f"{user_query} *({'、'.join(extra_tags)})*" if extra_tags else user_query

        st.session_state.messages.append({"role": "user", "content": display_q})
        with st.chat_message("user"):
            st.markdown(display_q)

        with st.chat_message("assistant"):
            with st.spinner("🔍 正在深度檢索三層資料庫並分析中..."):
                internal_context = search_internal_db(user_query, db["chunks"], selected_cat1s, selected_cat2s, selected_cat3s)
                file_summary = [
                    f"[{c1}>{c2}>{c3}] {fname}"
                    for fname, (c1, c2, c3, _) in file_info_map.items()
                    if c1 in selected_cat1s
                    and (not selected_cat2s or c2 in selected_cat2s)
                    and (not selected_cat3s or c3 in selected_cat3s)
                ]

                temp_text_context = ""
                if pasted_context and pasted_context.strip():
                    temp_text_context += f"\n=== 使用者直接複製貼上之臨時文字/表格內容 ===\n{pasted_context.strip()}\n====================================\n"

                multimodal_parts = []
                for tf in all_temp_files:
                    lname = tf.name.lower()
                    if lname.endswith((".jpg", ".jpeg", ".png")):
                        mime = "image/png" if lname.endswith(".png") else "image/jpeg"
                        multimodal_parts.append(types.Part.from_bytes(data=tf.getvalue(), mime_type=mime))
                        temp_text_context += f"\n【隨附影像/截圖：{tf.name}（請直接觀看影像內容並對照內部資料庫回答）】\n"
                    elif lname.endswith(".pdf"):
                        t_chunks, _ = extract_text_from_file(tf, "臨時", "臨時", "臨時", current_user, enable_ocr=False)
                        if t_chunks:
                            joined_t = "\n".join([f"{c['location']} {c['content']}" for c in t_chunks])
                            temp_text_context += f"\n=== 隨附臨時檔案：{tf.name} ===\n{joined_t[:15000]}\n"
                        else:
                            multimodal_parts.append(types.Part.from_bytes(data=tf.getvalue(), mime_type="application/pdf"))
                            temp_text_context += f"\n【隨附掃描PDF：{tf.name}（請直接視覺閱讀此PDF內容）】\n"
                    else:
                        t_chunks, _ = extract_text_from_file(tf, "臨時", "臨時", "臨時", current_user, enable_ocr=False)
                        joined_t = "\n".join([f"{c['location']} {c['content']}" for c in t_chunks])
                        temp_text_context += f"\n=== 隨附臨時檔案：{tf.name} ===\n{joined_t[:15000]}\n"

                system_prompt = f"""
你是一位專業、嚴謹的「工程專案主任工程師 AI 顧問」。提問組員：{current_user}。
請根據以下【內部工程資料庫檢索結果】、【使用者貼上/隨附之臨時資料】以及營建與採購法規知識回答。

【目前檢索範圍內的文件清單】：
{", ".join(file_summary) if file_summary else "尚未匯入文件"}

【回答嚴格守則】：
1. 明確標註出處：凡引用內部資料，必須標註「【目錄：第1層 > 第2層 > 第3層｜來源：檔名與頁碼/投影片頁數/工作表列數】」。
2. 會議決議時間軸追蹤：若詢問「某議題何時決議/討論過程」，請依會議日期先後順序列出「歷次討論演進時間軸」並標示最終拍板之會議與頁數。
3. 跨檔案全量日期/天數統計：若詢問「某工項施工日期或累計天數」，請將檢索結果中所有出現該工項之日報/週報日期逐一列出，並計算總天數。
4. 工程術語與表格精準度：注意「浮時(Float/浮實)」、「關鍵要徑(Critical Path)」、「里程碑分項」與表格表頭欄位之對應。
5. 嚴禁瞎掰：若資料庫與隨附資料皆未載明，請誠實告知「目前內部文件未載明」。

{temp_text_context}

=== 內部工程資料庫檢索結果 ===
{internal_context}
==============================

使用者問題：{user_query}
"""
                client = genai.Client(api_key=GEMINI_API_KEY)
                answer = None
                last_error = ""
                contents_payload = multimodal_parts + [system_prompt] if multimodal_parts else system_prompt

                for model_name in AVAILABLE_MODELS:
                    try:
                        config_args = {"temperature": 0.2}
                        if enable_web_search:
                            config_args["tools"] = [{"google_search": {}}]
                        response = client.models.generate_content(
                            model=model_name,
                            contents=contents_payload,
                            config=types.GenerateContentConfig(**config_args)
                        )
                        answer = response.text
                        break
                    except Exception as e:
                        last_error = str(e)
                        if enable_web_search and ("429" in last_error or "RESOURCE_EXHAUSTED" in last_error):
                            try:
                                response = client.models.generate_content(
                                    model=model_name,
                                    contents=contents_payload,
                                    config=types.GenerateContentConfig(temperature=0.2)
                                )
                                answer = response.text
                                break
                            except Exception as e2:
                                last_error = str(e2)
                        continue

                if not answer:
                    answer = f"⚠️ 呼叫失敗，請稍後再試。({last_error})"

                st.markdown(answer)
                st.session_state.messages.append({"role": "assistant", "content": answer})

# ----------------- 分頁 2：🔎 搜檔案 (0.1秒螢光筆關鍵字搜尋引擎) -----------------
with tab_search:
    col_s1, col_s2 = st.columns([3, 1])
    with col_s1:
        search_kw = st.text_input("🔎 輸入關鍵字秒查檔案內文（多個詞用空白隔開，如：里程碑 罰則 預算）：", placeholder="輸入後按 Enter 立即列出所有頁數...")
    with col_s2:
        match_mode = st.radio("匹配模式：", ["包含任一 (OR)", "包含全部 (AND)"], horizontal=True)

    if search_kw.strip():
        raw_kws = [k.strip() for k in search_kw.strip().split() if k.strip()]
        matched_results = []
        for c in chunks:
            if c.get("cat1") not in selected_cat1s:
                continue
            if selected_cat2s and c.get("cat2") not in selected_cat2s:
                continue
            if selected_cat3s and c.get("cat3") not in selected_cat3s:
                continue
            target_text = f"{c['file_name']} {c['content']}".lower()
            if match_mode == "包含全部 (AND)":
                is_match = all(k.lower() in target_text for k in raw_kws)
            else:
                is_match = any(k.lower() in target_text for k in raw_kws)
            if is_match:
                hits = sum(target_text.count(k.lower()) for k in raw_kws)
                matched_results.append((hits, c))

        matched_results.sort(key=lambda x: x[0], reverse=True)
        st.success(f"🎯 共找到 **{len(matched_results)}** 個段落！")

        for idx, (hits, item) in enumerate(matched_results, 1):
            with st.container(border=True):
                st.markdown(
                    f"**#{idx}｜🗂️ `{item['cat1']} > {item['cat2']} > {item.get('cat3', '一般')}` ｜ 📄 `{item['file_name']}` {item['location']}** "
                    f"*(命中 {hits} 次｜上傳：{item.get('uploader', '系統')})*"
                )
                hl_html = highlight_keywords(item["content"], raw_kws).replace("\n", "<br>")
                st.markdown(f"<div style='font-size: 0.95rem; line-height: 1.6;'>{hl_html}</div>", unsafe_allow_html=True)

# ----------------- 分頁 3：📦 批次搬移與整理 (多選打勾一鍵搬家/刪除) -----------------
with tab_batch:
    st.markdown("#### 📦 檔案批次搬移、改名與刪除總管")
    if not file_info_map:
        st.info("目前資料庫尚無檔案。")
    else:
        # 篩選要顯示的來源資料夾以便快速勾選
        b_col1, b_col2 = st.columns([2, 2])
        with b_col1:
            filter_b_c1 = st.selectbox("🔍 篩選檢視的來源主夾：", ["(顯示全部檔案)"] + list(folder_tree.keys()))
        with b_col2:
            kw_filter_file = st.text_input("🔍 或輸入檔名關鍵字過濾：", placeholder="例如：0d版、契約...")

        display_files = sorted(list(file_info_map.keys()))
        if filter_b_c1 != "(顯示全部檔案)":
            display_files = [f for f in display_files if file_info_map[f][0] == filter_b_c1]
        if kw_filter_file.strip():
            display_files = [f for f in display_files if kw_filter_file.strip().lower() in f.lower()]

        st.caption("👇 請在下方表格最左側 **打勾 `☑️` 選擇多份檔案**（亦可直接雙擊「檔案名稱」欄位修改檔名）：")
        table_data = []
        for f in display_files:
            fc1, fc2, fc3, fup = file_info_map[f]
            table_data.append({
                "勾選": False,
                "檔案名稱 (雙擊可直接改名)": f,
                "目前所在三層路徑": f"{fc1} > {fc2} > {fc3}",
                "上傳者": fup,
                "_orig_name": f
            })

        edited_table = st.data_editor(
            table_data,
            column_config={
                "勾選": st.column_config.CheckboxColumn("☑️ 勾選", default=False),
                "檔案名稱 (雙擊可直接改名)": st.column_config.TextColumn("📄 檔案名稱 (雙擊可直接改名)", width="large"),
                "目前所在三層路徑": st.column_config.TextColumn("🗂️ 目前所在三層路徑", disabled=True, width="large"),
                "上傳者": st.column_config.TextColumn("👤 上傳者", disabled=True, width="small"),
                "_orig_name": None
            },
            hide_index=True,
            use_container_width=True,
            key="batch_editor"
        )

        selected_rows = [r for r in edited_table if r["勾選"]]
        renamed_rows = [r for r in edited_table if r["檔案名稱 (雙擊可直接改名)"].strip() != r["_orig_name"]]

        if renamed_rows:
            if st.button(f"✏️ 立即儲存 {len(renamed_rows)} 份已修改的檔案名稱", type="primary"):
                for r in renamed_rows:
                    old_n = r["_orig_name"]
                    new_n = r["檔案名稱 (雙擊可直接改名)"].strip()
                    if new_n:
                        for c in chunks:
                            if c["file_name"] == old_n:
                                c["file_name"] = new_n
                save_db(db)
                st.success("✅ 檔案名稱已批次更新！")
                st.rerun()

        st.divider()
        st.markdown(f"##### 🚀 將已勾選的 **{len(selected_rows)}** 份檔案批次搬移或刪除：")
        m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns([2, 2, 2, 2, 2])
        with m_col1:
            dest_c1 = st.selectbox("目標第1層 (主夾)：", list(folder_tree.keys()), key="b_c1")
        with m_col2:
            dest_c2 = st.selectbox("目標第2層 (次夾)：", list(folder_tree[dest_c1].keys()), key="b_c2")
        with m_col3:
            dest_c3 = st.selectbox("目標第3層 (細夾)：", folder_tree[dest_c1][dest_c2], key="b_c3")
        with m_col4:
            st.write("")
            if st.button(f"📦 一鍵搬移 ({len(selected_rows)} 份)", type="primary", use_container_width=True, disabled=(len(selected_rows) == 0)):
                sel_names = {r["_orig_name"] for r in selected_rows}
                for c in chunks:
                    if c["file_name"] in sel_names:
                        c["cat1"], c["cat2"], c["cat3"] = dest_c1, dest_c2, dest_c3
                save_db(db)
                st.success(f"✅ 已將 {len(sel_names)} 份檔案成功搬移至【{dest_c1} > {dest_c2} > {dest_c3}】！")
                st.rerun()
        with m_col5:
            st.write("")
            if st.button(f"🗑️ 批次刪除 ({len(selected_rows)} 份)", use_container_width=True, disabled=(len(selected_rows) == 0)):
                sel_names = {r["_orig_name"] for r in selected_rows}
                db["chunks"] = [c for c in chunks if c["file_name"] not in sel_names]
                save_db(db)
                st.success(f"🗑️ 已批次刪除 {len(sel_names)} 份檔案！")
                st.rerun()

# ----------------- 分頁 4：📖 讀全文 -----------------
with tab_read:
    if not file_info_map:
        st.info("目前尚無檔案。")
    else:
        sorted_files = sorted(list(file_info_map.keys()))
        default_idx = sorted_files.index(st.session_state.jump_to_file) if st.session_state.jump_to_file in sorted_files else 0
        view_file = st.selectbox("選擇要閱讀的檔案：", sorted_files, index=default_idx)
        vc1, vc2, vc3, vup = file_info_map[view_file]
        st.caption(f"📂 三層路徑：**{vc1} > {vc2} > {vc3}** ｜ 👤 上傳者：**{vup}**")
        for fc in [c for c in chunks if c["file_name"] == view_file]:
            with st.expander(f"📄 {fc['file_name']} — {fc['location']}", expanded=True):
                st.text(fc["content"])
