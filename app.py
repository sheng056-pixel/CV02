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

# ================= 1. 系統基本設定 =================
st.set_page_config(page_title="CV02 工程助手", page_icon="🏗️", layout="wide")

TEAM_PASSWORD = st.secrets.get("TEAM_PASSWORD", "1234")
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "").strip()
DB_FILE = "engineering_knowledge_db.pkl"

DEFAULT_FOLDER_TREE = {
    "01_契約與補充規定": ["主契約條款", "投標須知與特訂條款"],
    "02_施工規範與圖說": ["土建工程規範", "機電工程規範", "職安衛與環保規範"],
    "03_設計變更與核定公文": ["契約變更設計", "重要往來公文"],
    "04_會議紀錄與簡報": ["工務會議紀錄", "專案進度簡報"],
    "05_材料送審與查驗紀錄": ["材料送審資料", "施工日報與查驗"]
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

# 工程同義詞與常見錯別字擴充字典
ENGINEERING_SYNONYMS = {
    "浮實": ["浮時", "總浮時", "自由浮時", "浮動", "Total Float", "TF"],
    "浮時": ["浮實", "總浮時", "自由浮時", "浮動", "Total Float", "TF"],
    "要徑": ["關鍵路徑", "關鍵要徑", "Critical Path", "CP", "進度", "落後", "趲趕"],
    "關鍵路徑": ["要徑", "關鍵要徑", "浮時"],
    "決議": ["結論", "裁示", "會議紀錄", "討論會", "拍板", "定案"],
    "預算": ["經費", "金額", "契約價金", "總價", "發包", "變更設計"],
    "施工日期": ["日報", "週報", "出工", "施作", "綁紮", "澆置", "開挖", "日曆天"]
}

def make_token(username, password):
    raw = f"{username}:{password}:cv02_salt"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]

# ================= 2. 資料庫讀取與無痛升級 =================
def load_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "rb") as f:
                data = pickle.load(f)
            if isinstance(data, dict) and "folder_tree" in data:
                if "users" not in data:
                    data["users"] = DEFAULT_USERS.copy()
                if "shortcuts" not in data:
                    data["shortcuts"] = DEFAULT_SHORTCUTS.copy()
                if "max_mb" not in data:
                    data["max_mb"] = 80
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
                        "uploader": "聖引"
                    })
                return {
                    "folder_tree": tree,
                    "chunks": upgraded_chunks,
                    "users": DEFAULT_USERS.copy(),
                    "shortcuts": DEFAULT_SHORTCUTS.copy(),
                    "max_mb": 80
                }
        except Exception:
            pass
    return {
        "folder_tree": DEFAULT_FOLDER_TREE.copy(),
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

# ================= 4. 視覺 OCR 辨識與文件解析引擎 (含表頭保留) =================
def ocr_with_gemini(file_bytes, mime_type, desc=""):
    if not GEMINI_API_KEY:
        return ""
    client = genai.Client(api_key=GEMINI_API_KEY)
    prompt = (
        "請完整、精準地辨識並轉錄這份工程公文、圖表或照片中的所有繁體中文文字、文號、日期、表格數據與印章文字。"
        "若有表格請以『欄位1 | 欄位2』格式呈現，不要遺漏任何細節。"
    )
    for model_name in ["gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-3.5-flash"]:
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

def extract_text_from_file(uploaded_file, cat1, cat2, uploader, enable_ocr=True):
    text_chunks = []
    filename = uploaded_file.name
    lower_name = filename.lower()
    raw_bytes = uploaded_file.getvalue()

    # 1. 圖片格式 (JPG / PNG) -> 直接啟動視覺 OCR
    if lower_name.endswith((".jpg", ".jpeg", ".png")):
        mime = "image/png" if lower_name.endswith(".png") else "image/jpeg"
        ocr_text = ocr_with_gemini(raw_bytes, mime, filename) if enable_ocr else ""
        if ocr_text:
            text_chunks.append({
                "file_name": filename, "cat1": cat1, "cat2": cat2,
                "location": "(影像OCR辨識)", "content": ocr_text, "uploader": uploader
            })
        return text_chunks, (1 if ocr_text else 0)

    # 2. PDF 格式 (自動判斷一般文字或掃描死圖 OCR)
    elif lower_name.endswith(".pdf"):
        reader = pypdf.PdfReader(io.BytesIO(raw_bytes))
        scan_pages = []
        for i, page in enumerate(reader.pages):
            page_text = (page.extract_text() or "").strip()
            if len(page_text) >= 25:
                text_chunks.append({
                    "file_name": filename, "cat1": cat1, "cat2": cat2,
                    "location": f"(第 {i+1} 頁)", "content": page_text, "uploader": uploader
                })
            else:
                scan_pages.append(i)

        ocr_count = 0
        # 若有掃描頁且開啟 OCR，分批將掃描頁打包成單頁 PDF 送交 Gemini 視覺辨識 (單次最多處理前 25 頁防卡死)
        if scan_pages and enable_ocr:
            max_ocr_pages = min(len(scan_pages), 25)
            prog = st.progress(0, text=f"🔍 偵測到《{filename}》含掃描頁面，正在啟動 AI 光學辨識 (共 {max_ocr_pages} 頁)...")
            for idx, page_num in enumerate(scan_pages[:max_ocr_pages]):
                writer = pypdf.PdfWriter()
                writer.add_page(reader.pages[page_num])
                buf = io.BytesIO()
                writer.write(buf)
                ocr_res = ocr_with_gemini(buf.getvalue(), "application/pdf", f"{filename} P.{page_num+1}")
                if ocr_res:
                    text_chunks.append({
                        "file_name": filename, "cat1": cat1, "cat2": cat2,
                        "location": f"(第 {page_num+1} 頁-OCR掃描辨識)", "content": ocr_res, "uploader": uploader
                    })
                    ocr_count += 1
                prog.progress((idx + 1) / max_ocr_pages, text=f"🔍 正在辨識《{filename}》第 {page_num+1} 頁 ({idx+1}/{max_ocr_pages})...")
            prog.empty()
        return text_chunks, ocr_count

    # 3. Word 格式 (表格自動保留第一列表頭)
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
                    "file_name": filename, "cat1": cat1, "cat2": cat2,
                    "location": f"(段落 {i//900 + 1})", "content": full_text[i:i+900], "uploader": uploader
                })

    # 4. PPTX 簡報格式 (表格自動綁定表頭)
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
                    "file_name": filename, "cat1": cat1, "cat2": cat2,
                    "location": f"(投影片第 {i+1} 頁)", "content": "\n".join(slide_texts), "uploader": uploader
                })

    # 5. Excel 格式 (自動保留每張工作表第一列表頭，確保工項、浮時、日期不脫鉤)
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
            # 每 25 列打包成一個區塊，並自動冠上表頭名稱
            for i in range(0, len(rows_text), 25):
                batch = rows_text[i:i+25]
                if i > 0 and header_str:
                    batch.insert(0, f"【表頭欄位：{header_str}】")
                chunk_str = "\n".join(batch)
                if chunk_str.strip():
                    text_chunks.append({
                        "file_name": filename, "cat1": cat1, "cat2": cat2,
                        "location": f"(工作表:{sheet} 第{i+1}~{min(i+25, len(rows_text))}列)",
                        "content": chunk_str, "uploader": uploader
                    })

    # 6. TXT 純文字
    elif lower_name.endswith(".txt"):
        full_text = raw_bytes.decode("utf-8", errors="ignore")
        for i in range(0, len(full_text), 900):
            if full_text[i:i+900].strip():
                text_chunks.append({
                    "file_name": filename, "cat1": cat1, "cat2": cat2,
                    "location": f"(段落 {i//900 + 1})", "content": full_text[i:i+900], "uploader": uploader
                })

    return text_chunks, 0

# ================= 5. 智慧檢索演算法 (含同義詞擴充 & 全量統計模式) =================
def expand_query_with_synonyms(query):
    expanded = query
    for key, syns in ENGINEERING_SYNONYMS.items():
        if key in query:
            expanded += " " + " ".join(syns)
    return expanded

def search_internal_db(query, chunks, selected_cat1s, selected_cat2s):
    if not chunks:
        return "（目前內部目錄尚無檔案，請先從左側『➕ 上傳』新增檔案）"

    filtered_chunks = [
        c for c in chunks
        if c.get("cat1") in selected_cat1s and (not selected_cat2s or c.get("cat2") in selected_cat2s)
    ]
    if not filtered_chunks:
        return "（目前勾選的資料夾範圍內尚無文件）"

    # 判斷是否為「全量統計 / 日期 / 歷次決議 / 要徑浮時」類問題，若是則自動擴大抓取深度至 100 筆
    deep_keywords = ["日期", "天數", "幾天", "所有", "全部", "統計", "歷次", "決議", "何時", "什麼時候", "要徑", "浮時", "浮實", "日報"]
    is_deep_search = any(k in query for k in deep_keywords)
    top_k = 100 if is_deep_search else 18

    expanded_query = expand_query_with_synonyms(query)
    corpus = [f"{item['file_name']} {item['cat2']} {item['content']}" for item in filtered_chunks]
    vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
    tfidf_matrix = vectorizer.fit_transform(corpus + [expanded_query])
    cosine_sim = cosine_similarity(tfidf_matrix[-1], tfidf_matrix[:-1]).flatten()
    top_indices = cosine_sim.argsort()[-top_k:][::-1]

    selected_items = []
    for idx in top_indices:
        if cosine_sim[idx] > 0.008:
            selected_items.append(filtered_chunks[idx])

    if not selected_items:
        selected_items = filtered_chunks[:min(10, len(filtered_chunks))]

    # 若為日期統計或會議決議追蹤，依檔案名稱排序以利呈現時間軸
    if is_deep_search:
        selected_items.sort(key=lambda x: (x["file_name"], x["location"]))

    results = [
        f"【目錄：{item['cat1']} > {item['cat2']}｜來源：{item['file_name']} {item['location']}】\n{item['content']}"
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

# ================= 6. 側邊欄：極簡控制台與權限分流 =================
folder_tree = db["folder_tree"]
chunks = db["chunks"]
max_mb_limit = db.get("max_mb", 80)

file_info_map = {}
for c in chunks:
    c1 = c.get("cat1", list(folder_tree.keys())[0])
    if c1 not in folder_tree:
        folder_tree[c1] = ["預設子夾"]
    c2 = c.get("cat2", folder_tree[c1][0])
    if c2 not in folder_tree[c1]:
        folder_tree[c1].append(c2)
    file_info_map[c["file_name"]] = (c1, c2, c.get("uploader", "系統"))

db_size_mb = os.path.getsize(DB_FILE) / (1024 * 1024) if os.path.exists(DB_FILE) else 0.0

if "jump_to_file" not in st.session_state:
    st.session_state.jump_to_file = None

with st.sidebar:
    col_u1, col_u2 = st.columns([3, 2])
    with col_u1:
        st.markdown(f"👤 **{current_user}**")
    with col_u2:
        if st.button("🚪 登出", use_container_width=True):
            st.session_state.current_user = None
            st.query_params.clear()
            st.rerun()

    st.caption(f"📊 檔案數：**{len(file_info_map)}** 份 ｜ 容量：**{db_size_mb:.2f} MB** / 50 MB")

    # --- 1. ➕ 上傳 (含防呆檢查與自動 OCR) ---
    with st.expander("➕ 上傳", expanded=False):
        up_c1 = st.selectbox("1. 主資料夾：", list(folder_tree.keys()))
        up_c2 = st.selectbox("2. 子資料夾：", folder_tree[up_c1])
        enable_ocr_upload = st.checkbox("👁️ 掃描檔/圖片自動啟動 AI 光學辨識 (OCR)", value=True)
        uploaded_files = st.file_uploader(
            f"3. 選擇檔案 (單檔上限 {max_mb_limit}MB，支援 PDF/Word/PPT/Excel/圖片)",
            type=["pdf", "docx", "pptx", "xlsx", "txt", "jpg", "jpeg", "png"],
            accept_multiple_files=True
        )
        if st.button("🚀 匯入資料庫", type="primary", use_container_width=True):
            if not uploaded_files:
                st.warning("請先選擇檔案！")
            else:
                # 防呆檢查 1：檢查單一檔案大小
                oversized = [uf.name for uf in uploaded_files if (uf.size / (1024 * 1024)) > max_mb_limit]
                # 防呆檢查 2：檢查單次圖片數量
                img_count = sum(1 for uf in uploaded_files if uf.name.lower().endswith((".jpg", ".jpeg", ".png")))
                if oversized:
                    st.error(f"⚠️ 上傳防呆警示：以下檔案超過 {max_mb_limit}MB 上限，請先壓縮或分割後再傳（或至下方調高上限）：\n" + "、".join(oversized))
                elif img_count > 10:
                    st.warning(f"⚠️ 您一次選擇了 {img_count} 張圖片，為避免光學辨識逾時，建議一次上傳 10 張以內！")
                else:
                    with st.spinner("正在解析內容並建立索引..."):
                        new_count, total_ocr = 0, 0
                        for uf in uploaded_files:
                            db["chunks"] = [c for c in db["chunks"] if c["file_name"] != uf.name]
                            new_chunks, ocr_cnt = extract_text_from_file(uf, up_c1, up_c2, current_user, enable_ocr=enable_ocr_upload)
                            if new_chunks:
                                db["chunks"].extend(new_chunks)
                                new_count += 1
                                total_ocr += ocr_cnt
                            else:
                                st.warning(f"⚠️《{uf.name}》未偵測到文字內容。")
                        save_db(db)
                        msg_extra = f"（含 {total_ocr} 頁掃描檔 OCR 辨識）" if total_ocr > 0 else ""
                        st.success(f"✅ 成功匯入 {new_count} 份檔案{msg_extra}！")
                        st.rerun()

    # --- 2. 📁 目錄 (點擊檔名一鍵即開) ---
    with st.expander("📁 目錄", expanded=True):
        if not file_info_map:
            st.caption("目前尚無檔案，請點上方「➕ 上傳」。")
        else:
            st.caption("💡 點擊檔名可於右側「📖 讀全文」直接檢視：")
            for c1, sub_list in folder_tree.items():
                c1_files = [f for f, (f_c1, _, _) in file_info_map.items() if f_c1 == c1]
                if c1_files:
                    st.markdown(f"**🗂️ {c1}**")
                    for c2 in sub_list:
                        c2_files = sorted([f for f, (f_c1, f_c2, _) in file_info_map.items() if f_c1 == c1 and f_c2 == c2])
                        if c2_files:
                            st.markdown(f" └ 📂 **{c2}** `({len(c2_files)})`")
                            for f in c2_files:
                                if st.button(f"📄 {f}", key=f"dir_{f}", use_container_width=True):
                                    st.session_state.jump_to_file = f

    # --- 3. ✏️ 編輯檔案 ---
    if file_info_map:
        with st.expander("✏️ 編輯檔案", expanded=False):
            selected_file = st.selectbox("選擇檔案：", sorted(list(file_info_map.keys())))
            new_file_name = st.text_input("修改檔名：", value=selected_file)
            curr_c1, curr_c2, _ = file_info_map[selected_file]
            c1_list = list(folder_tree.keys())
            move_c1 = st.selectbox("主資料夾：", c1_list, index=c1_list.index(curr_c1) if curr_c1 in c1_list else 0)
            c2_list = folder_tree[move_c1]
            move_c2 = st.selectbox("子資料夾：", c2_list, index=c2_list.index(curr_c2) if curr_c2 in c2_list else 0)
            b1, b2 = st.columns(2)
            with b1:
                if st.button("💾 儲存", use_container_width=True):
                    if new_file_name.strip():
                        for c in chunks:
                            if c["file_name"] == selected_file:
                                c["file_name"] = new_file_name.strip()
                                c["cat1"], c["cat2"] = move_c1, move_c2
                        save_db(db)
                        st.rerun()
            with b2:
                if st.button("🗑️ 刪除", use_container_width=True):
                    db["chunks"] = [c for c in chunks if c["file_name"] != selected_file]
                    save_db(db)
                    st.rerun()

    # --- 4. ⚙️ 資料夾與防呆設定 ---
    with st.expander("⚙️ 資料夾設定", expanded=False):
        t1, t2, t3 = st.tabs(["主夾", "子夾", "防呆與快捷"])
        with t1:
            new_c1 = st.text_input("➕ 新增主夾：")
            if st.button("新增主夾", use_container_width=True) and new_c1.strip():
                if new_c1.strip() not in folder_tree:
                    folder_tree[new_c1.strip()] = ["一般文件"]
                    save_db(db)
                    st.rerun()
            target_c1 = st.selectbox("選擇主夾：", list(folder_tree.keys()), key="s_c1")
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
                    fb2 = folder_tree[fb1][0]
                    for c in chunks:
                        if c.get("cat1") == target_c1:
                            c["cat1"], c["cat2"] = fb1, fb2
                    save_db(db)
                    st.rerun()
        with t2:
            p_c1 = st.selectbox("所屬主夾：", list(folder_tree.keys()))
            new_c2 = st.text_input("➕ 新增子夾：")
            if st.button("新增子夾", use_container_width=True) and new_c2.strip():
                if new_c2.strip() not in folder_tree[p_c1]:
                    folder_tree[p_c1].append(new_c2.strip())
                    save_db(db)
                    st.rerun()
            target_c2 = st.selectbox("選擇子夾：", folder_tree[p_c1])
            rename_c2 = st.text_input("改名子夾：", value=target_c2)
            cc, cd = st.columns(2)
            with cc:
                if st.button("儲存子夾", use_container_width=True) and rename_c2.strip():
                    idx = folder_tree[p_c1].index(target_c2)
                    folder_tree[p_c1][idx] = rename_c2.strip()
                    for c in chunks:
                        if c.get("cat1") == p_c1 and c.get("cat2") == target_c2:
                            c["cat2"] = rename_c2.strip()
                    save_db(db)
                    st.rerun()
            with cd:
                if st.button("🗑️ 刪子夾", use_container_width=True) and len(folder_tree[p_c1]) > 1:
                    folder_tree[p_c1].remove(target_c2)
                    for c in chunks:
                        if c.get("cat1") == p_c1 and c.get("cat2") == target_c2:
                            c["cat2"] = folder_tree[p_c1][0]
                    save_db(db)
                    st.rerun()
        with t3:
            new_max_mb = st.slider("單檔上傳防呆上限 (MB)：", 20, 200, int(max_mb_limit), 10)
            if new_max_mb != max_mb_limit:
                db["max_mb"] = new_max_mb
                save_db(db)
                st.rerun()
            sc_str = st.text_area("自訂首頁快捷按鈕（一行一個）：", value="\n".join(db.get("shortcuts", DEFAULT_SHORTCUTS)))
            if st.button("儲存快捷鈕", use_container_width=True):
                db["shortcuts"] = [s.strip() for s in sc_str.split("\n") if s.strip()]
                save_db(db)
                st.rerun()

    # --- 5. 👥 帳號管理 (嚴格限定「聖引」專屬顯示與執行) ---
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

    # --- 6. 💾 備份與一鍵還原 ---
    with st.expander("💾 備份與還原", expanded=False):
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as f:
                st.download_button("📥 下載備份 (.pkl)", data=f, file_name="engineering_knowledge_db.pkl", use_container_width=True)
        restore_file = st.file_uploader("📤 上傳備份檔一鍵還原：", type=["pkl"])
        if restore_file and st.button("⚡ 立即還原知識庫", use_container_width=True):
            with open(DB_FILE, "wb") as f:
                f.write(restore_file.getvalue())
            st.success("知識庫已成功還原！")
            st.rerun()

    enable_web_search = st.checkbox("🌐 Google 聯網搜尋", value=False)

# ================= 7. 主畫面：極簡頂部 & 範圍篩選 =================
col_title, col_filter = st.columns([3, 2])
with col_title:
    st.subheader("🏗️ CV02 工程助手")
with col_filter:
    with st.popover("🎯 篩選檢索目錄範圍（預設全選）", use_container_width=True):
        selected_cat1s = st.multiselect("1. 主資料夾：", list(folder_tree.keys()), default=list(folder_tree.keys()))
        avail_c2 = [c2 for c1 in selected_cat1s for c2 in folder_tree.get(c1, [])]
        selected_cat2s = st.multiselect("2. 子資料夾（留空代表全選）：", avail_c2, default=[])

# 若從左側目錄點擊了某個檔案，顯示頂部快速閱覽卡片
if st.session_state.jump_to_file and st.session_state.jump_to_file in file_info_map:
    j_file = st.session_state.jump_to_file
    jc1, jc2, jup = file_info_map[j_file]
    with st.container(border=True):
        jh1, jh2 = st.columns([5, 1])
        with jh1:
            st.markdown(f"### 📖 快速閱覽：`{j_file}`  `({jc1} > {jc2}｜上傳：{jup})`")
        with jh2:
            if st.button("✖️ 關閉閱覽", use_container_width=True):
                st.session_state.jump_to_file = None
                st.rerun()
        for fc in [c for c in chunks if c["file_name"] == j_file]:
            with st.expander(f"📄 {fc['location']}", expanded=True):
                st.text(fc["content"])

# 建立極簡三大分頁
tab_ai, tab_search, tab_read = st.tabs(["💬 AI 問答", "🔎 搜檔案", "📖 讀全文"])

# ----------------- 分頁 1：💬 AI 問答 (含一鍵快查鈕 + 隨附檔案/照片提問) -----------------
with tab_ai:
    # 一鍵快查膠囊按鈕列
    shortcuts = db.get("shortcuts", DEFAULT_SHORTCUTS)
    triggered_query = None
    if shortcuts:
        sc_cols = st.columns(len(shortcuts))
        for idx, sc_text in enumerate(shortcuts):
            if sc_cols[idx].button(sc_text, key=f"sc_{idx}", use_container_width=True):
                triggered_query = sc_text

    # 對話框隨附臨時檔案 / 拍照片區
    with st.expander("📎 隨附臨時檔案或工地照片一起問 AI（不弄亂永久目錄，亦可一鍵轉存）", expanded=False):
        temp_files = st.file_uploader(
            "選擇要讓 AI 立即比對的臨時檔案或照片 (支援 PDF, Word, PPTX, Excel, JPG, PNG)：",
            type=["pdf", "docx", "pptx", "xlsx", "txt", "jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key="chat_temp_files"
        )
        if temp_files:
            tc1, tc2, tc3 = st.columns([2, 2, 2])
            with tc1:
                save_t_c1 = st.selectbox("若欲轉存之主夾：", list(folder_tree.keys()), key="st_c1")
            with tc2:
                save_t_c2 = st.selectbox("若欲轉存之子夾：", folder_tree[save_t_c1], key="st_c2")
            with tc3:
                st.write("")
                if st.button("💾 將附件存入永久目錄", use_container_width=True):
                    for tf in temp_files:
                        db["chunks"] = [c for c in db["chunks"] if c["file_name"] != tf.name]
                        new_c, _ = extract_text_from_file(tf, save_t_c1, save_t_c2, current_user, enable_ocr=True)
                        db["chunks"].extend(new_c)
                    save_db(db)
                    st.success("已將附件正式歸檔至永久目錄！")
                    st.rerun()

    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "assistant", "content": f"**{current_user}** 您好！可直接點上方**快捷鈕**、打字提問，或展開上方 **「📎 隨附臨時檔案或照片」** 讓我幫您跟資料庫交叉比對！"}
        ]

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    chat_input_val = st.chat_input("請輸入問題（如：鋼筋施工日期統計、某議題何時決議、關鍵要徑浮時...）")
    user_query = triggered_query or chat_input_val

    if user_query:
        display_q = f"{user_query} *(📎 隨附 {len(temp_files)} 份臨時附件)*" if temp_files else user_query
        st.session_state.messages.append({"role": "user", "content": display_q})
        with st.chat_message("user"):
            st.markdown(display_q)

        with st.chat_message("assistant"):
            with st.spinner("🔍 正在深度檢索資料庫並分析中..."):
                internal_context = search_internal_db(user_query, db["chunks"], selected_cat1s, selected_cat2s)
                file_summary = [
                    f"[{c1}>{c2}] {fname}"
                    for fname, (c1, c2, _) in file_info_map.items()
                    if c1 in selected_cat1s and (not selected_cat2s or c2 in selected_cat2s)
                ]

                # 解析對話框隨附的臨時文件或圖片
                temp_text_context = ""
                multimodal_parts = []
                if temp_files:
                    for tf in temp_files:
                        lname = tf.name.lower()
                        if lname.endswith((".jpg", ".jpeg", ".png")):
                            mime = "image/png" if lname.endswith(".png") else "image/jpeg"
                            multimodal_parts.append(types.Part.from_bytes(data=tf.getvalue(), mime_type=mime))
                            temp_text_context += f"\n【隨附圖片：{tf.name}（請直接觀看影像內容並對照內部資料庫回答）】\n"
                        elif lname.endswith(".pdf"):
                            # 同時抽取文字，若為掃描PDF則直接將PDF二進位送交模型視覺閱讀
                            t_chunks, _ = extract_text_from_file(tf, "臨時附件", "臨時附件", current_user, enable_ocr=False)
                            if t_chunks:
                                joined_t = "\n".join([f"{c['location']} {c['content']}" for c in t_chunks])
                                temp_text_context += f"\n=== 隨附臨時檔案：{tf.name} ===\n{joined_t[:15000]}\n"
                            else:
                                multimodal_parts.append(types.Part.from_bytes(data=tf.getvalue(), mime_type="application/pdf"))
                                temp_text_context += f"\n【隨附掃描PDF：{tf.name}（請直接視覺閱讀此PDF內容）】\n"
                        else:
                            t_chunks, _ = extract_text_from_file(tf, "臨時附件", "臨時附件", current_user, enable_ocr=False)
                            joined_t = "\n".join([f"{c['location']} {c['content']}" for c in t_chunks])
                            temp_text_context += f"\n=== 隨附臨時檔案：{tf.name} ===\n{joined_t[:15000]}\n"

                system_prompt = f"""
你是一位專業、嚴謹的「工程專案主任工程師 AI 顧問」。提問組員：{current_user}。
請根據以下【內部工程資料庫檢索結果】、【隨附臨時檔案/照片】以及營建與採購法規知識回答。

【目前檢索範圍內的文件清單】：
{", ".join(file_summary) if file_summary else "尚未匯入文件"}

【回答嚴格守則】：
1. 明確標註出處：凡引用內部資料，必須標註「【來源：檔名與頁碼/投影片頁數/工作表列數】」。
2. 會議決議時間軸追蹤：若詢問「某議題何時決議/討論過程」，請務必依會議日期先後順序列出「歷次討論演進時間軸」並標示最終拍板定案之會議與頁數。
3. 跨檔案全量日期/天數統計：若詢問「某工項施工日期或累計天數」，請將檢索結果中所有出現該工項之日報/週報日期逐一列出，並幫使用者計算總計天數。
4. 工程術語與表格精準度：注意「浮時(Float/浮實)」、「關鍵要徑(Critical Path)」與表格表頭欄位之對應，切勿張冠李戴。
5. 嚴禁瞎掰：若資料庫未載明，請誠實告知「目前內部文件未載明」。

{temp_text_context}

=== 內部工程資料庫檢索結果 ===
{internal_context}
==============================

使用者問題：{user_query}
"""
                client = genai.Client(api_key=GEMINI_API_KEY)
                answer = None
                contents_payload = multimodal_parts + [system_prompt] if multimodal_parts else system_prompt

                for model_name in ["gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-3.5-flash"]:
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
                        if enable_web_search and "429" in last_error:
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
        search_kw = st.text_input("🔎 輸入關鍵字秒查檔案內文（多個詞用空白隔開，如：二期 水工 預算）：", placeholder="輸入後按 Enter 立即列出所有頁數...")
    with col_s2:
        match_mode = st.radio("匹配模式：", ["包含任一 (OR)", "包含全部 (AND)"], horizontal=True)

    if search_kw.strip():
        raw_kws = [k.strip() for k in search_kw.strip().split() if k.strip()]
        # 連同工程同義詞一起高亮與輔助檢索
        matched_results = []
        for c in chunks:
            if c.get("cat1") not in selected_cat1s:
                continue
            if selected_cat2s and c.get("cat2") not in selected_cat2s:
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
                    f"**#{idx}｜🗂️ `{item['cat1']} > {item['cat2']}` ｜ 📄 `{item['file_name']}` {item['location']}** "
                    f"*(命中 {hits} 次｜上傳：{item.get('uploader', '系統')})*"
                )
                hl_html = highlight_keywords(item["content"], raw_kws).replace("\n", "<br>")
                st.markdown(f"<div style='font-size: 0.95rem; line-height: 1.6;'>{hl_html}</div>", unsafe_allow_html=True)

# ----------------- 分頁 3：📖 讀全文 -----------------
with tab_read:
    if not file_info_map:
        st.info("目前尚無檔案。")
    else:
        default_idx = 0
        sorted_files = sorted(list(file_info_map.keys()))
        if st.session_state.jump_to_file in sorted_files:
            default_idx = sorted_files.index(st.session_state.jump_to_file)
        view_file = st.selectbox("選擇要閱讀的檔案：", sorted_files, index=default_idx)
        vc1, vc2, vup = file_info_map[view_file]
        st.caption(f"📂 路徑：**{vc1} > {vc2}** ｜ 👤 上傳者：**{vup}**")
        for fc in [c for c in chunks if c["file_name"] == view_file]:
            with st.expander(f"📄 {fc['file_name']} — {fc['location']}", expanded=True):
                st.text(fc["content"])
