import os
import io
import re
import hashlib
import pickle
import base64
import urllib.parse
import requests
import streamlit as st
import streamlit.components.v1 as components
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
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "").strip()
GITHUB_REPO = st.secrets.get("GITHUB_REPO", "").strip()
DB_FILE = "engineering_knowledge_db.pkl"

AVAILABLE_MODELS = ["gemini-3.5-flash-lite", "gemini-3.8-flash", "gemini-3.5-flash"]

DEFAULT_FOLDER_TREE_V7 = {
    "01_招標文件與契約": {
        "第一次公告版本": ["契約條款", "投標與評選須知", "施工規範與圖說"],
        "第二次公告版本": ["契約條款", "投標與評選須知", "補充規定與附件"]
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
    "聖引": {"password": TEAM_PASSWORD, "role": "管理員", "display_name": "聖引"},
    "668436": {"password": TEAM_PASSWORD, "role": "一般組員", "display_name": "淳凱"},
    "077343": {"password": TEAM_PASSWORD, "role": "一般組員", "display_name": "有朋"},
    "479701": {"password": TEAM_PASSWORD, "role": "一般組員", "display_name": "旭暉"},
    "331165": {"password": TEAM_PASSWORD, "role": "一般組員", "display_name": "宣頷"},
    "組員A": {"password": TEAM_PASSWORD, "role": "一般組員", "display_name": "組員A"}
}

DEFAULT_SHORTCUTS = [
    "📌 最新進度與預算",
    "⏳ 關鍵要徑與浮時",
    "📅 鋼筋工項施工日期統計",
    "📋 歷次會議決議追蹤",
    "⚠️ 職安衛查驗標準"
]

DEFAULT_CLOUD_LINKS = [
    {"title": "📁 每日施工照片雲端夾", "url": "https://drive.google.com/drive/folders/1MFrIjPZB3cBWBb2Nt_q2eaTRDw7zzno9?usp=drive_link"},
    {"title": "📁 每日日報更新雲端夾", "url": "https://drive.google.com/drive/folders/1vmmaXftLfCTT3fd5_x_sIl6jc9KrEATX?usp=drive_link"}
]

DEFAULT_FIGMA_BOARDS = []

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

# ================= 自動 GitHub 同步備份機制 =================
def pull_db_from_github():
    if not GITHUB_TOKEN or not GITHUB_REPO:
        return None
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{DB_FILE}?ref=main"
    headers = {"Authorization": f"Bearer {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
    try:
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            content_b64 = r.json().get("content", "")
            if content_b64:
                return base64.b64decode(content_b64)
    except Exception:
        pass
    return None

def push_db_to_github(db_bytes):
    if not GITHUB_TOKEN or not GITHUB_REPO:
        return
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{DB_FILE}"
    headers = {"Authorization": f"Bearer {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
    try:
        r_get = requests.get(url, headers=headers, timeout=10)
        sha = r_get.json().get("sha", "") if r_get.status_code == 200 else ""
        data = {
            "message": "🔄 Auto-sync CV02 Database from Streamlit",
            "content": base64.b64encode(db_bytes).decode("utf-8"),
            "branch": "main"
        }
        if sha:
            data["sha"] = sha
        requests.put(url, headers=headers, json=data, timeout=15)
    except Exception:
        pass

def make_token(username, password):
    raw = f"{username}:{password}:cv02_salt"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]

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
    if not os.path.exists(DB_FILE) and GITHUB_TOKEN and GITHUB_REPO:
        remote_bytes = pull_db_from_github()
        if remote_bytes:
            with open(DB_FILE, "wb") as f:
                f.write(remote_bytes)

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
                    c3 = c.get("cat3", "")
                    c["cat1"], c["cat2"], c["cat3"] = c1, c2, c3
                    if "uploader" not in c:
                        c["uploader"] = "聖引"
                return {
                    "folder_tree": tree_3d,
                    "chunks": chunks,
                    "users": data.get("users", DEFAULT_USERS.copy()),
                    "shortcuts": data.get("shortcuts", DEFAULT_SHORTCUTS.copy()),
                    "cloud_links": data.get("cloud_links", [dict(x) for x in DEFAULT_CLOUD_LINKS]),
                    "figma_boards": data.get("figma_boards", [dict(x) for x in DEFAULT_FIGMA_BOARDS]),
                    "max_mb": data.get("max_mb", 80)
                }
        except Exception:
            pass
    return {
        "folder_tree": DEFAULT_FOLDER_TREE_V7.copy(),
        "chunks": [],
        "users": DEFAULT_USERS.copy(),
        "shortcuts": DEFAULT_SHORTCUTS.copy(),
        "cloud_links": [dict(x) for x in DEFAULT_CLOUD_LINKS],
        "figma_boards": [dict(x) for x in DEFAULT_FIGMA_BOARDS],
        "max_mb": 80
    }

def save_db(db_dict):
    with open(DB_FILE, "wb") as f:
        pickle.dump(db_dict, f)
    if GITHUB_TOKEN and GITHUB_REPO:
        with open(DB_FILE, "rb") as f:
            push_db_to_github(f.read())

def to_figma_embed_url(raw_url):
    u = raw_url.strip()
    if "iframe" in u and "src=" in u:
        m = re.search(r'src=["\']([^"\']+)["\']', u)
        if m:
            return m.group(1)
    if "figma.com/embed" in u:
        return u
    encoded = urllib.parse.quote(u, safe="")
    return f"https://www.figma.com/embed?embed_host=share&url={encoded}"

db = load_db()
users_db = db["users"]
folder_tree = db["folder_tree"]
chunks = db["chunks"]
max_mb_limit = db.get("max_mb", 80)

# ================= AI 智慧自動分類路徑預測引擎 =================
def auto_classify_folder(filename, folder_tree):
    fn = filename.lower()
    for c1, c2_dict in folder_tree.items():
        for c2, c3_list in c2_dict.items():
            for c3 in c3_list:
                if c3.lower() in fn:
                    return c1, c2, c3
            c2_clean = re.sub(r"^\d+_", "", c2).lower()
            if c2_clean and c2_clean in fn:
                return c1, c2, (c3_list[0] if c3_list else "")

    rules = [
        (["日報", "自主檢查", "查驗"], "05_材料送審與查驗紀錄", "施工日報與查驗", "施工日報表"),
        (["送審", "試驗", "試體", "出廠", "型錄"], "05_材料送審與查驗紀錄", "材料送審資料", "材料送審核定"),
        (["會議", "現勘", "紀錄", "討論會"], "04_會議紀錄與簡報", "工務會議紀錄", "歷次會議紀錄"),
        (["簡報", "週會", "月會", "進度", "ppt"], "04_會議紀錄與簡報", "專案進度簡報", "週會與月會簡報"),
        (["變更", "議價", "新增單價", "修正契約"], "03_設計變更與核定公文", "契約變更設計", "變更設計圖說"),
        (["函", "公文", "來文", "發文", "備忘錄", "簽"], "03_設計變更與核定公文", "重要往來公文", "業主來文"),
        (["職安", "安衛", "環保", "安全"], "02_施工規範與圖說", "職安衛與環保", "職安衛規定"),
        (["機電", "弱電", "電氣", "空調", "消防"], "02_施工規範與圖說", "機電與弱電", "設備規範"),
        (["規範", "圖說", "施工", "結構", "土建", "水工"], "02_施工規範與圖說", "土建與水工", "一般規範"),
        (["契約", "投標", "招標", "須知", "價目表", "預算", "單價"], "01_招標文件與契約", "第二次公告版本", "契約條款"),
    ]
    for kws, target_c1, target_c2, target_c3 in rules:
        if any(k in fn for k in kws):
            if target_c1 in folder_tree:
                c2_dict = folder_tree[target_c1]
                c2_use = target_c2 if target_c2 in c2_dict else list(c2_dict.keys())[0]
                c3_list = c2_dict[c2_use]
                c3_use = target_c3 if target_c3 in c3_list else (c3_list[0] if c3_list else "")
                return target_c1, c2_use, c3_use

    def_c1 = list(folder_tree.keys())[0]
    def_c2 = list(folder_tree[def_c1].keys())[0]
    def_c3 = folder_tree[def_c1][def_c2][0] if folder_tree[def_c1][def_c2] else ""
    return def_c1, def_c2, def_c3

# ================= 登入驗證 =================
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
        login_options = list(users_db.keys())
        def format_user_label(u):
            if u in users_db and "display_name" in users_db[u]:
                return f"{u} ({users_db[u]['display_name']})"
            return u

        login_user = st.selectbox("👤 選擇帳號 / 輸入工號：", login_options, format_func=format_user_label)
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

raw_current_user = st.session_state.current_user
current_user = users_db.get(raw_current_user, {}).get("display_name", raw_current_user)

# ================= 檔案解析與匯入核心函數 =================
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

def process_and_save_uploads(uploaded_files, mode="manual", target_c1="", target_c2="", target_c3="", enable_ocr=True):
    oversized = [uf.name for uf in uploaded_files if (uf.size / (1024 * 1024)) > max_mb_limit]
    img_count = sum(1 for uf in uploaded_files if uf.name.lower().endswith((".jpg", ".jpeg", ".png")))
    if oversized:
        st.error(f"⚠️ 超過 {max_mb_limit}MB 上限：\n" + "、".join(oversized))
        return False
    if img_count > 10:
        st.warning(f"⚠️ 一次選擇了 {img_count} 張圖片，建議單次 10 張以內！")
        return False

    with st.spinner("⚡ 正在極速解析內容並自動同步備份至 GitHub..."):
        imported_details = []
        total_ocr = 0
        for uf in uploaded_files:
            if mode == "smart":
                c1, c2, c3 = auto_classify_folder(uf.name, folder_tree)
            else:
                c1, c2, c3 = target_c1, target_c2, target_c3

            db["chunks"] = [c for c in db["chunks"] if c["file_name"] != uf.name]
            new_chunks, ocr_cnt = extract_text_from_file(uf, c1, c2, c3, current_user, enable_ocr=enable_ocr)
            if new_chunks:
                db["chunks"].extend(new_chunks)
                path_label = f"{c1} > {c2}" + (f" > {c3}" if c3 else " (第二層)")
                imported_details.append(f"`{uf.name}` ➔ 📂 **{path_label}**")
                total_ocr += ocr_cnt

        save_db(db)
        st.session_state.uploader_key += 1
        st.session_state.upload_notice = {
            "count": len(imported_details),
            "files": imported_details,
            "ocr_pages": total_ocr
        }
    return True

# ================= 彈出視窗：匯入通知、指定快傳、第三層細夾編輯 =================
@st.dialog("🎉 檔案極速匯入成功！")
def show_upload_success_dialog(info):
    st.success(f"✅ 已成功將 **{info['count']}** 份檔案寫入資料庫並自動同步至 GitHub！")
    if info.get("ocr_pages", 0) > 0:
        st.info(f"👁️ 內含 **{info['ocr_pages']}** 頁掃描影像，已透過 AI 視覺辨識轉為可搜尋文字！")
    st.markdown("**📄 本次歸檔清單與路徑（若想換位置可於目錄點擊改名/搬移）：**")
    for item_str in info.get("files", []):
        st.markdown(f"- {item_str}")
    if st.button("👌 太棒了，我知道了", type="primary", use_container_width=True):
        st.session_state.upload_notice = None
        st.rerun()

@st.dialog("📥 直接傳入此資料夾（拖拉放開即自動匯入）")
def show_quick_folder_upload_dialog(target_info):
    tc1, tc2, tc3 = target_info
    path_str = f"{tc1} > {tc2}" + (f" > {tc3}" if tc3 else " (直接放第二層)")
    st.markdown(f"📂 **目標鎖定路徑：** `{path_str}`")
    st.caption("💡 直接將電腦中的檔案拖拉進下方框內，**手放開就會立刻自動開始匯入**，免按確認鈕！")
    q_ocr = st.checkbox("👁️ 掃描檔自動 AI 光學辨識 (OCR)", value=True, key="q_dlg_ocr")
    q_files = st.file_uploader(
        "請拖曳或選擇檔案：",
        type=["pdf", "docx", "pptx", "xlsx", "txt", "jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"q_dlg_up_{st.session_state.uploader_key}"
    )
    if q_files:
        if process_and_save_uploads(q_files, mode="manual", target_c1=tc1, target_c2=tc2, target_c3=tc3, enable_ocr=q_ocr):
            st.session_state.quick_upload_target = None
            st.rerun()
    if st.button("✖️ 取消關閉", use_container_width=True):
        st.session_state.quick_upload_target = None
        st.rerun()

@st.dialog("✏️ 編輯第三層細夾名稱")
def show_edit_c3_dialog(edit_info):
    ec1, ec2, ec3 = edit_info
    st.caption(f"📍 所屬位置：`{ec1} > {ec2}`")
    new_c3_val = st.text_input("✏️ 修改細夾名稱：", value=ec3)
    col_e1, col_e2, col_e3 = st.columns(3)
    with col_e1:
        if st.button("💾 儲存名稱", type="primary", use_container_width=True):
            clean_v = new_c3_val.strip()
            if clean_v and clean_v != ec3:
                idx = folder_tree[ec1][ec2].index(ec3)
                folder_tree[ec1][ec2][idx] = clean_v
                for c in chunks:
                    if c.get("cat1") == ec1 and c.get("cat2") == ec2 and c.get("cat3") == ec3:
                        c["cat3"] = clean_v
                save_db(db)
            st.session_state.edit_c3_target = None
            st.rerun()
    with col_e2:
        if st.button("🗑️ 刪除此細夾", use_container_width=True):
            if len(folder_tree[ec1][ec2]) > 1:
                folder_tree[ec1][ec2].remove(ec3)
                for c in chunks:
                    if c.get("cat1") == ec1 and c.get("cat2") == ec2 and c.get("cat3") == ec3:
                        c["cat3"] = ""
                save_db(db)
                st.session_state.edit_c3_target = None
                st.rerun()
            else:
                st.warning("至少需保留一個細夾！")
    with col_e3:
        if st.button("✖️ 取消", use_container_width=True):
            st.session_state.edit_c3_target = None
            st.rerun()

if "jump_to_file" not in st.session_state:
    st.session_state.jump_to_file = None
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "quick_upload_target" not in st.session_state:
    st.session_state.quick_upload_target = None
if "edit_c3_target" not in st.session_state:
    st.session_state.edit_c3_target = None

if st.session_state.get("upload_notice"):
    show_upload_success_dialog(st.session_state.upload_notice)
elif st.session_state.get("quick_upload_target"):
    show_quick_folder_upload_dialog(st.session_state.quick_upload_target)
elif st.session_state.get("edit_c3_target"):
    show_edit_c3_dialog(st.session_state.edit_c3_target)

# ================= 6. 全工項動態智慧聯想引擎 =================
def expand_query_with_synonyms(query):
    expanded = query
    for key, syns in ENGINEERING_SYNONYMS.items():
        if key in query:
            expanded += " " + " ".join(syns)
    
    if GEMINI_API_KEY and len(query.strip()) <= 20:
        try:
            client = genai.Client(api_key=GEMINI_API_KEY)
            res = client.models.generate_content(
                model="gemini-3.5-flash-lite",
                contents=f"請針對工程合約與詳細價目表中的查詢詞『{query}』，列出 3 至 5 個常見的工程同義詞、細部規格或相關工項名稱（例如查管子可聯想 PVC, PE, 鋼管, 排水管；查拆除可聯想打除, 拔除, 既有結構）。請直接用空白隔開輸出這些詞彙，不要輸出其他多餘文字。",
                config=types.GenerateContentConfig(temperature=0.1)
            )
            if res and res.text:
                expanded += " " + res.text.strip().replace("\n", " ")
        except Exception:
            pass
    return expanded

def search_internal_db(query, chunks, selected_cat1s, selected_cat2s, selected_cat3s):
    if not chunks:
        return "（目前內部目錄尚無檔案，請先從左側上傳檔案）"

    filtered_chunks = [
        c for c in chunks
        if c.get("cat1") in selected_cat1s
        and (not selected_cat2s or c.get("cat2") in selected_cat2s)
        and (not selected_cat3s or (not c.get("cat3") or c.get("cat3") in selected_cat3s))
    ]
    if not filtered_chunks:
        return "（目前勾選的資料夾範圍內尚無文件）"

    deep_keywords = ["日期", "天數", "幾天", "所有", "全部", "統計", "歷次", "決議", "何時", "什麼時候", "要徑", "浮時", "浮實", "日報", "里程碑", "分項", "價目表", "工項", "單價", "拔除", "管"]
    is_deep_search = any(k in query for k in deep_keywords)
    top_k = 120 if is_deep_search else 25

    expanded_query = expand_query_with_synonyms(query)
    corpus = [f"{item['file_name']} {item['cat2']} {item.get('cat3', '')} {item['content']}" for item in filtered_chunks]
    vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
    tfidf_matrix = vectorizer.fit_transform(corpus + [expanded_query])
    cosine_sim = cosine_similarity(tfidf_matrix[-1], tfidf_matrix[:-1]).flatten()
    top_indices = cosine_sim.argsort()[-top_k:][::-1]

    selected_items = []
    for idx in top_indices:
        if cosine_sim[idx] > 0.005:
            selected_items.append(filtered_chunks[idx])

    if not selected_items:
        selected_items = filtered_chunks[:min(15, len(filtered_chunks))]

    if is_deep_search:
        selected_items.sort(key=lambda x: (x["file_name"], x["location"]))

    results = [
        f"【目錄：{item['cat1']} > {item['cat2']}{' > ' + item['cat3'] if item.get('cat3') else ''}｜來源：{item['file_name']} {item['location']}】\n{item['content']}"
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

file_info_map = {}
for c in chunks:
    c1 = c.get("cat1", list(folder_tree.keys())[0])
    if c1 not in folder_tree:
        folder_tree[c1] = {"預設次夾": ["一般文件"]}
    c2 = c.get("cat2", list(folder_tree[c1].keys())[0])
    if c2 not in folder_tree[c1]:
        folder_tree[c1][c2] = ["一般文件"]
    c3 = c.get("cat3", "")
    file_info_map[c["file_name"]] = (c1, c2, c3, c.get("uploader", "系統"))

db_size_mb = os.path.getsize(DB_FILE) / (1024 * 1024) if os.path.exists(DB_FILE) else 0.0

# ================= 左側導覽：極速上傳 ＋ 展開就地編輯目錄 ＋ 動態捷徑 =================
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

    # ★ 黃金組合 1：AI 智慧極速拖曳區
    with st.expander("⚡ 智慧極速上傳 (AI 自動歸檔)", expanded=True):
        st.caption("💡 直接將多份檔案拖進下方框內，**手放開即自動依檔名分類並存入 GitHub**，免選資料夾、免按確認！")
        smart_ocr = st.checkbox("👁️ 掃描檔自動 AI 辨識 (OCR)", value=True, key="smart_ocr_chk")
        smart_files = st.file_uploader(
            f"拖拉檔案至此 (單檔上限 {max_mb_limit}MB)",
            type=["pdf", "docx", "pptx", "xlsx", "txt", "jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key=f"smart_up_{st.session_state.uploader_key}"
        )
        if smart_files:
            if process_and_save_uploads(smart_files, mode="smart", enable_ocr=smart_ocr):
                st.rerun()

    # ★ 黃金組合 3：手動指定資料夾上傳
    with st.expander("➕ 手動指定資料夾上傳", expanded=False):
        up_c1 = st.selectbox("1. 第一層 (主夾)：", list(folder_tree.keys()))
        up_c2 = st.selectbox("2. 第二層 (次夾)：", list(folder_tree[up_c1].keys()))
        raw_c3_opts = folder_tree[up_c1][up_c2]
        c3_opts = ["(無細夾 /直接放第二層)"] + (raw_c3_opts if isinstance(raw_c3_opts, list) else [])
        up_c3_selected = st.selectbox("3. 第三層 (細夾，可選無)：", c3_opts)
        up_c3 = "" if up_c3_selected == "(無細夾 /直接放第二層)" else up_c3_selected

        c_chk1, c_chk2 = st.columns(2)
        with c_chk1:
            enable_ocr_upload = st.checkbox("👁️ 自動 OCR", value=True)
        with c_chk2:
            auto_submit_manual = st.checkbox("⚡ 選完即傳", value=True)

        uploaded_files = st.file_uploader(
            f"4. 選擇檔案 (上限 {max_mb_limit}MB)",
            type=["pdf", "docx", "pptx", "xlsx", "txt", "jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key=f"file_up_{st.session_state.uploader_key}"
        )
        if uploaded_files and auto_submit_manual:
            if process_and_save_uploads(uploaded_files, mode="manual", target_c1=up_c1, target_c2=up_c2, target_c3=up_c3, enable_ocr=enable_ocr_upload):
                st.rerun()
        elif not auto_submit_manual:
            if st.button("🚀 匯入資料庫", type="primary", use_container_width=True):
                if not uploaded_files:
                    st.warning("請先選擇檔案！")
                else:
                    if process_and_save_uploads(uploaded_files, mode="manual", target_c1=up_c1, target_c2=up_c2, target_c3=up_c3, enable_ocr=enable_ocr_upload):
                        st.rerun()

    # ★ 展開目錄即可「直接原地改名、新增子夾、刪除」＋ 點 📥 直接傳
    with st.expander("📁 目錄 (展開可直接改名 / 點 📥 快傳)", expanded=True):
        st.caption("💡 展開資料夾可**直接改名或新增子夾**；點 `📥` 直接傳檔：")
        
        c_top1, c_top2 = st.columns([3, 2])
        with c_top1:
            show_empty_dirs = st.checkbox("顯示空資料夾", value=True)
        with c_top2:
            with st.popover("➕ 新增主夾", use_container_width=True):
                quick_new_c1 = st.text_input("輸入新第1層主夾名稱：", key="q_add_c1")
                if st.button("確認新增", key="btn_q_add_c1", type="primary", use_container_width=True) and quick_new_c1.strip():
                    if quick_new_c1.strip() not in folder_tree:
                        folder_tree[quick_new_c1.strip()] = {"預設次夾": ["一般文件"]}
                        save_db(db)
                        st.rerun()

        for c1, c2_dict in list(folder_tree.items()):
            c1_files = [f for f, (fc1, _, _, _) in file_info_map.items() if fc1 == c1]
            if not show_empty_dirs and not c1_files:
                continue
            with st.expander(f"🗂️ {c1} ({len(c1_files)})", expanded=False):
                with st.popover(f"✏️ 管理「{c1}」(改名 / 新增次夾 / 刪除)", use_container_width=True):
                    ren_c1_val = st.text_input("✏️ 修改第1層主夾名稱：", value=c1, key=f"in_ren_c1_{c1}")
                    if st.button("💾 儲存主夾名稱", key=f"btn_ren_c1_{c1}", type="primary", use_container_width=True):
                        clean_c1 = ren_c1_val.strip()
                        if clean_c1 and clean_c1 != c1:
                            db["folder_tree"] = {(clean_c1 if k == c1 else k): v for k, v in folder_tree.items()}
                            for c in chunks:
                                if c.get("cat1") == c1:
                                    c["cat1"] = clean_c1
                            save_db(db)
                            st.rerun()
                    st.divider()
                    add_c2_val = st.text_input("➕ 在此主夾下新增「第2層次夾」：", key=f"in_add_c2_{c1}")
                    if st.button("➕ 新增次夾", key=f"btn_add_c2_{c1}", use_container_width=True) and add_c2_val.strip():
                        if add_c2_val.strip() not in folder_tree[c1]:
                            folder_tree[c1][add_c2_val.strip()] = ["一般文件"]
                            save_db(db)
                            st.rerun()
                    st.divider()
                    if st.button("🗑️ 刪除此主夾", key=f"btn_del_c1_{c1}", use_container_width=True):
                        if len(folder_tree) > 1:
                            del folder_tree[c1]
                            fb1 = list(folder_tree.keys())[0]
                            fb2 = list(folder_tree[fb1].keys())[0]
                            for c in chunks:
                                if c.get("cat1") == c1:
                                    c["cat1"], c["cat2"], c["cat3"] = fb1, fb2, ""
                            save_db(db)
                            st.rerun()
                        else:
                            st.warning("至少需保留一個主夾！")

                for c2, c3_list in list(c2_dict.items()):
                    c2_files = [f for f, (fc1, fc2, _, _) in file_info_map.items() if fc1 == c1 and fc2 == c2]
                    if not show_empty_dirs and not c2_files:
                        continue
                    with st.expander(f"📂 {c2} ({len(c2_files)})", expanded=False):
                        with st.popover(f"✏️ 管理「{c2}」(改名 / 新增細夾 / 刪除)", use_container_width=True):
                            ren_c2_val = st.text_input("✏️ 修改第2層次夾名稱：", value=c2, key=f"in_ren_c2_{c1}_{c2}")
                            if st.button("💾 儲存次夾名稱", key=f"btn_ren_c2_{c1}_{c2}", type="primary", use_container_width=True):
                                clean_c2 = ren_c2_val.strip()
                                if clean_c2 and clean_c2 != c2:
                                    folder_tree[c1] = {(clean_c2 if k == c2 else k): v for k, v in folder_tree[c1].items()}
                                    for c in chunks:
                                        if c.get("cat1") == c1 and c.get("cat2") == c2:
                                            c["cat2"] = clean_c2
                                    save_db(db)
                                    st.rerun()
                            st.divider()
                            add_c3_val = st.text_input("➕ 在此次夾下新增「第3層細夾」：", key=f"in_add_c3_{c1}_{c2}")
                            if st.button("➕ 新增細夾", key=f"btn_add_c3_{c1}_{c2}", use_container_width=True) and add_c3_val.strip():
                                if add_c3_val.strip() not in folder_tree[c1][c2]:
                                    folder_tree[c1][c2].append(add_c3_val.strip())
                                    save_db(db)
                                    st.rerun()
                            st.divider()
                            if st.button("🗑️ 刪除此次夾", key=f"btn_del_c2_{c1}_{c2}", use_container_width=True):
                                if len(folder_tree[c1]) > 1:
                                    del folder_tree[c1][c2]
                                    fb2 = list(folder_tree[c1].keys())[0]
                                    for c in chunks:
                                        if c.get("cat1") == c1 and c.get("cat2") == c2:
                                            c["cat2"], c["cat3"] = fb2, ""
                                    save_db(db)
                                    st.rerun()
                                else:
                                    st.warning("至少需保留一個次夾！")

                        col_d2_title, col_d2_btn = st.columns([4, 1])
                        with col_d2_title:
                            no_c3_files = sorted([f for f, (fc1, fc2, fc3, _) in file_info_map.items() if fc1 == c1 and fc2 == c2 and not fc3])
                            st.markdown(f"**🏷️ (直接放第二層)** `({len(no_c3_files)})`")
                        with col_d2_btn:
                            if st.button("📥", key=f"qup_{c1}_{c2}_none", help=f"直接上傳檔案至：{c1} > {c2}"):
                                st.session_state.quick_upload_target = (c1, c2, "")
                                st.rerun()
                        for f in no_c3_files:
                            if st.button(f"📄 {f}", key=f"dir_{c1}_{c2}_none_{f}", use_container_width=True):
                                st.session_state.jump_to_file = f

                        for c3 in list(c3_list):
                            c3_files = sorted([f for f, (fc1, fc2, fc3, _) in file_info_map.items() if fc1 == c1 and fc2 == c2 and fc3 == c3])
                            if not show_empty_dirs and not c3_files:
                                continue
                            col_d3_title, col_d3_edit, col_d3_up = st.columns([3.2, 0.9, 0.9])
                            with col_d3_title:
                                st.markdown(f"**🏷️ {c3}** `({len(c3_files)})`")
                            with col_d3_edit:
                                if st.button("✏️", key=f"edc3_{c1}_{c2}_{c3}", help=f"改名或刪除細夾：{c3}"):
                                    st.session_state.edit_c3_target = (c1, c2, c3)
                                    st.rerun()
                            with col_d3_up:
                                if st.button("📥", key=f"qup_{c1}_{c2}_{c3}", help=f"直接上傳檔案至：{c1} > {c2} > {c3}"):
                                    st.session_state.quick_upload_target = (c1, c2, c3)
                                    st.rerun()
                            for f in c3_files:
                                if st.button(f"📄 {f}", key=f"dir_{c1}_{c2}_{c3}_{f}", use_container_width=True):
                                    st.session_state.jump_to_file = f

    # ★ 升級：可在網頁上自由新增/管理的雲端與外部捷徑
    with st.expander("☁️ 專案雲端與 Figma 捷徑", expanded=True):
        cloud_links = db.get("cloud_links", DEFAULT_CLOUD_LINKS)
        for lk in cloud_links:
            st.markdown(f"[{lk['title']}]({lk['url']})")
        for fb in db.get("figma_boards", []):
            st.markdown(f"[🎨 {fb['title']} (開新視窗)]({fb['url']})")

        with st.popover("✏️ 管理捷徑連結", use_container_width=True):
            st.markdown("##### ➕ 新增外部捷徑連結")
            nl_title = st.text_input("連結顯示名稱：", placeholder="如：📁 變更設計雲端夾", key="nl_t")
            nl_url = st.text_input("網址 (URL)：", placeholder="https://...", key="nl_u")
            if st.button("➕ 新增捷徑", type="primary", use_container_width=True):
                if nl_title.strip() and nl_url.strip():
                    db.setdefault("cloud_links", []).append({"title": nl_title.strip(), "url": nl_url.strip()})
                    save_db(db)
                    st.rerun()
            if cloud_links:
                st.divider()
                del_lk_idx = st.selectbox("選擇要移除的捷徑：", range(len(cloud_links)), format_func=lambda i: cloud_links[i]["title"], key="del_lk")
                if st.button("🗑️ 移除選中捷徑", use_container_width=True):
                    db["cloud_links"].pop(del_lk_idx)
                    save_db(db)
                    st.rerun()

    with st.expander("⚙️ 防呆上限與首頁快捷設定", expanded=False):
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

    if raw_current_user == "聖引":
        with st.expander("👥 帳號管理 (聖引專屬)", expanded=False):
            new_u = st.text_input("➕ 新增工號/帳號：")
            new_name = st.text_input("設定中文姓名 (如：小明)：")
            new_p = st.text_input("設定密碼：", value=TEAM_PASSWORD)
            if st.button("新增帳號", use_container_width=True) and new_u.strip():
                users_db[new_u.strip()] = {
                    "password": new_p.strip() or TEAM_PASSWORD, 
                    "role": "一般組員", 
                    "display_name": new_name.strip() or new_u.strip()
                }
                save_db(db)
                st.success(f"已新增：{new_u.strip()} ({new_name})")
                st.rerun()
            st.divider()
            ed_u = st.selectbox("選擇帳號：", list(users_db.keys()))
            ed_name = st.text_input("修改中文姓名：", value=users_db[ed_u].get("display_name", ed_u))
            ed_p = st.text_input("修改密碼：", value=users_db[ed_u]["password"])
            u1, u2 = st.columns(2)
            with u1:
                if st.button("更新設定", use_container_width=True):
                    users_db[ed_u]["password"] = ed_p.strip() or TEAM_PASSWORD
                    users_db[ed_u]["display_name"] = ed_name.strip() or ed_u
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

    with st.expander("💾 手動備份與還原", expanded=False):
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as f:
                st.download_button("📥 下載備份 (.pkl)", data=f, file_name="engineering_knowledge_db.pkl", use_container_width=True)
        restore_file = st.file_uploader("📤 上傳舊有備份檔還原：", type=["pkl"])
        if restore_file and st.button("⚡ 立即還原並同步至 GitHub", use_container_width=True):
            with open(DB_FILE, "wb") as f:
                f.write(restore_file.getvalue())
            db = load_db()
            save_db(db)
            st.success("已還原並同步！")
            st.rerun()

    enable_web_search = st.checkbox("🌐 Google 聯網搜尋", value=False)

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

if st.session_state.jump_to_file and st.session_state.jump_to_file in file_info_map:
    j_file = st.session_state.jump_to_file
    jc1, jc2, jc3, jup = file_info_map[j_file]
    with st.container(border=True):
        jh1, jh2 = st.columns([5, 1])
        with jh1:
            path_str = f"{jc1} > {jc2}" + (f" > {jc3}" if jc3 else "")
            st.markdown(f"### 📖 快速閱覽：`{j_file}`  `({path_str}｜上傳：{jup})`")
        with jh2:
            if st.button("✖️ 關閉閱覽", use_container_width=True):
                st.session_state.jump_to_file = None
                st.rerun()

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
                raw_q3_opts = folder_tree[q_c1][q_c2]
                q3_opts_all = ["(無細夾 / 直接放第二層)"] + (raw_q3_opts if isinstance(raw_q3_opts, list) else [])
                curr_q3_idx = q3_opts_all.index(jc3) if jc3 in q3_opts_all else 0
                q_c3_sel = st.selectbox("移至第3層：", q3_opts_all, index=curr_q3_idx, key="q_c3_sel")
                q_c3 = "" if q_c3_sel == "(無細夾 / 直接放第二層)" else q_c3_sel

            qb1, qb2 = st.columns(2)
            with qb1:
                if st.button("💾 儲存變更", key="q_save", type="primary", use_container_width=True):
                    clean_qn = q_new_name.strip() or j_file
                    for c in chunks:
                        if c["file_name"] == j_file:
                            c["file_name"] = clean_qn
                            c["cat1"], c["cat2"], c["cat3"] = q_c1, q_c2, q_c3
                    with st.spinner("💾 正在儲存並同步至 GitHub..."):
                        save_db(db)
                    st.session_state.jump_to_file = clean_qn
                    st.rerun()
            with qb2:
                if st.button("🗑️ 刪除此檔案", key="q_del", use_container_width=True):
                    db["chunks"] = [c for c in chunks if c["file_name"] != j_file]
                    with st.spinner("🗑️ 正在刪除並同步至 GitHub..."):
                        save_db(db)
                    st.session_state.jump_to_file = None
                    st.rerun()

        for fc in [c for c in chunks if c["file_name"] == j_file]:
            with st.expander(f"📄 {fc['location']}", expanded=True):
                st.text(fc["content"])

# ================= 核心六大分頁 =================
tab_ai, tab_figma, tab_memo, tab_search, tab_batch, tab_read = st.tabs([
    "💬 AI 問答",
    "🎨 流程圖 (Figma)",
    "📝 公文生成",
    "🔎 搜檔案",
    "📦 批次搬移與整理",
    "📖 讀全文"
])

with tab_ai:
    shortcuts = db.get("shortcuts", DEFAULT_SHORTCUTS)
    triggered_query = None
    if shortcuts:
        sc_cols = st.columns(len(shortcuts))
        for idx, sc_text in enumerate(shortcuts):
            if sc_cols[idx].button(sc_text, key=f"sc_{idx}", use_container_width=True):
                triggered_query = sc_text

    with st.expander("📋 直接複製貼上文字／表格 ＆ 📎 隨附臨時檔案問 AI（送出後自動清空絕不沾黏）", expanded=False):
        p_col1, p_col2 = st.columns(2)
        with p_col1:
            pasted_context = st.text_area(
                "📋 直接貼上臨時文字 / 公文草稿 / LINE 訊息 / 表格 (Ctrl+V)：",
                height=115,
                placeholder="在此直接 Ctrl+V 貼上任何文字或表格，下方提問時 AI 會自動納入比對..."
            )
        with p_col2:
            temp_files = st.file_uploader(
                "📎 或選擇臨時檔案 / 螢幕截圖：",
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
                raw_st_c3 = folder_tree[save_t_c1][save_t_c2]
                st_c3_opts = ["(無細夾 / 直接放第二層)"] + (raw_st_c3 if isinstance(raw_st_c3, list) else [])
                st_c3_sel = st.selectbox("細夾：", st_c3_opts, key="st_c3_sel")
                save_t_c3 = "" if st_c3_sel == "(無細夾 / 直接放第二層)" else st_c3_sel
            with tc4:
                pasted_fname = st.text_input("若存貼上文字之檔名：", value="臨時貼上資料.txt", key="st_fn")
            with tc5:
                st.write("")
                if st.button("💾 存入永久目錄", type="primary", use_container_width=True):
                    with st.spinner("💾 正在存入資料庫並同步備份至 GitHub..."):
                        saved_details = []
                        path_lbl = f"{save_t_c1} > {save_t_c2}" + (f" > {save_t_c3}" if save_t_c3 else "")
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
                            saved_details.append(f"`{fn_clean}` ➔ 📂 **{path_lbl}**")
                        if temp_files:
                            for tf in temp_files:
                                db["chunks"] = [c for c in db["chunks"] if c["file_name"] != tf.name]
                                new_c, _ = extract_text_from_file(tf, save_t_c1, save_t_c2, save_t_c3, current_user, enable_ocr=True)
                                db["chunks"].extend(new_c)
                                saved_details.append(f"`{tf.name}` ➔ 📂 **{path_lbl}**")
                        save_db(db)
                        st.session_state.upload_notice = {
                            "count": len(saved_details),
                            "files": saved_details,
                            "ocr_pages": 0
                        }
                        st.rerun()

    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "assistant", "content": f"**{current_user}** 您好！您可以在下方輸入框**直接打字、按 `Ctrl+V` 直接貼上螢幕截圖**，或直接把檔案拉進左側最上方的「⚡ 智慧極速上傳」讓 AI 自動幫您歸檔！"}
        ]

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    inline_files = []
    user_query = None
    try:
        chat_val = st.chat_input(
            "請輸入問題（支援在此直接 Ctrl+V 貼上截圖，或夾帶檔案）...",
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
        chat_val = st.chat_input("請輸入問題（如：詳細價目表有沒有設計舊樁拔除、管子、里程碑分項...）")
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
            with st.spinner("🔍 正在透過 AI 智慧動態同義詞擴展，深度檢索詳細價目表與規範中..."):
                internal_context = search_internal_db(user_query, db["chunks"], selected_cat1s, selected_cat2s, selected_cat3s)
                file_summary = [
                    f"[{c1}>{c2}{'>' + c3 if c3 else ''}] {fname}"
                    for fname, (c1, c2, c3, _) in file_info_map.items()
                    if c1 in selected_cat1s
                    and (not selected_cat2s or c2 in selected_cat2s)
                    and (not selected_cat3s or (not c3 or c3 in selected_cat3s))
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
                        t_chunks, _ = extract_text_from_file(tf, "臨時", "臨時", "", current_user, enable_ocr=False)
                        if t_chunks:
                            joined_t = "\n".join([f"{c['location']} {c['content']}" for c in t_chunks])
                            temp_text_context += f"\n=== 隨附臨時檔案：{tf.name} ===\n{joined_t[:15000]}\n"
                        else:
                            multimodal_parts.append(types.Part.from_bytes(data=tf.getvalue(), mime_type="application/pdf"))
                            temp_text_context += f"\n【隨附掃描PDF：{tf.name}（請直接視覺閱讀此PDF內容）】\n"
                    else:
                        t_chunks, _ = extract_text_from_file(tf, "臨時", "臨時", "", current_user, enable_ocr=False)
                        joined_t = "\n".join([f"{c['location']} {c['content']}" for c in t_chunks])
                        temp_text_context += f"\n=== 隨附臨時檔案：{tf.name} ===\n{joined_t[:15000]}\n"

                system_prompt = f"""
你是一位專業、嚴謹的「工程專案主任工程師 AI 顧問」。提問組員：{current_user}。
請根據以下【內部工程資料庫檢索結果】、【使用者貼上/隨附之臨時資料】以及營建與採購法規知識回答。

【目前檢索範圍內的文件清單】：
{", ".join(file_summary) if file_summary else "尚未匯入文件"}

【回答嚴格守則】：
1. 明確標註出處：凡引用內部資料，必須標註「【目錄：第1層 > 第2層 [> 第3層]｜來源：檔名與頁碼/投影片頁數/工作表列數】」。
2. 工程詳細價目表/工項智能對應：面對使用者查詢特定工項（如管子、舊樁拔除、防水等），請務必主動將檢索結果中所有「同義詞、細部規格（如PVC管、既有結構拆除）」納入綜合比對。若價目表或契約中確實有包含該性質之項目，絕對不可直接斷言未編列！
3. 會議決議時間軸追蹤：若詢問「某議題何時決議/討論過程」，請依會議日期先後順序列出「歷次討論演進時間軸」並標示最終拍板之會議與頁數。
4. 跨檔案全量日期/天數統計：若詢問「某工項施工日期或累計天數」，請將檢索結果中所有出現該工項之日報/週報日期逐一列出，並計算總天數。
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
                st.rerun()

# ================= 🎨 流程圖 (Figma 互動內嵌與管理) 專屬分頁 =================
with tab_figma:
    f_head1, f_head2 = st.columns([4, 1.5])
    with f_head1:
        st.markdown("### 🎨 專案工程流程圖與架構看板 (Figma 互動內嵌)")
    with f_head2:
        with st.popover("➕ 新增 / 管理 Figma 流程圖", use_container_width=True):
            st.markdown("##### 📌 新增 Figma 流程圖")
            st.caption("💡 請確認 Figma 右上角 **Share** 權限已設為 `Anyone with the link can view`，直接貼上網址即可自動內嵌！")
            fg_title = st.text_input("流程圖名稱：", placeholder="如：施工查驗與送審流程圖")
            fg_url = st.text_input("Figma 分享連結 (URL)：", placeholder="https://www.figma.com/...")
            if st.button("💾 儲存並內嵌此流程圖", type="primary", use_container_width=True):
                if fg_title.strip() and fg_url.strip():
                    db.setdefault("figma_boards", []).append({
                        "title": fg_title.strip(),
                        "url": fg_url.strip()
                    })
                    with st.spinner("💾 正在儲存並同步至 GitHub..."):
                        save_db(db)
                    st.rerun()
                else:
                    st.warning("請填寫名稱與網址！")

            figma_list = db.get("figma_boards", [])
            if figma_list:
                st.divider()
                st.markdown("##### 🗑️ 移除現有流程圖")
                del_fg_idx = st.selectbox("選擇要移除的項目：", range(len(figma_list)), format_func=lambda i: figma_list[i]["title"])
                if st.button("🗑️ 刪除此流程圖", use_container_width=True):
                    db["figma_boards"].pop(del_fg_idx)
                    save_db(db)
                    st.rerun()

    figma_boards = db.get("figma_boards", [])
    if not figma_boards:
        st.info("👆 目前尚未加入 Figma 流程圖！請點擊右上角的 **「➕ 新增 / 管理 Figma 流程圖」**，將您的 Figma 網址貼上，即可在此直接縮放、拖曳瀏覽互動式流程圖！")
    else:
        ctrl1, ctrl2, ctrl3 = st.columns([2.5, 1.5, 1.5])
        with ctrl1:
            sel_fg_idx = st.selectbox(
                "📂 選擇要檢視的流程圖看板：",
                range(len(figma_boards)),
                format_func=lambda i: f"🎨 {figma_boards[i]['title']}"
            )
        selected_board = figma_boards[sel_fg_idx]
        with ctrl2:
            canvas_height = st.slider("📐 畫布顯示高度 (px)：", 500, 1200, 750, 50)
        with ctrl3:
            st.write("")
            st.link_button("🚀 在 Figma 新分頁開啟編輯", selected_board["url"], use_container_width=True)

        embed_src = to_figma_embed_url(selected_board["url"])
        components.iframe(embed_src, height=canvas_height, scrolling=True)

# ================= 📝 公文生成專屬分頁 =================
with tab_memo:
    st.markdown("### 📝 智慧公文生成與自動排版室")
    st.caption("AI 將自動學習您上傳在資料庫中的歷史公文慣用語法、敘述風格與陳核對象，產出標準排版公文，並可一鍵下載 Word (.docx) 檔！")

    memo_topic = st.text_area("📌 請輸入擬稿需求或事件摘要（例如：針對現場排水管線繞道變更設計，擬函請監造單位核定...）：", height=100)
    
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        memo_type = st.selectbox("📄 文別選擇：", ["函", "書函", "簽呈", "開會通知單", "紀錄"])
    with m_col2:
        memo_target = st.text_input("🏢 預定受文對象（如未填，AI 將自動從歷史公文學習判斷）：", placeholder="例如：台灣世曦工程顧問股份有限公司 / 監造單位")

    if st.button("🚀 開始智慧擬稿與排版", type="primary", use_container_width=True):
        if not memo_topic.strip():
            st.warning("請先輸入擬稿需求！")
        else:
            with st.spinner("🤖 AI 正在深度參考您的歷史公文風格、慣用語法、受文對象與內部資料庫契約規範進行擬稿中..."):
                memo_history_chunks = [
                    c for c in db["chunks"] 
                    if "公文" in c.get("cat1", "") or "公文" in c.get("cat2", "") or "往來" in c.get("cat2", "")
                ]
                history_text = "\n".join([f"【參考範本：{c['file_name']}】\n{c['content']}" for c in memo_history_chunks[:10]])
                contract_context = search_internal_db(memo_topic, db["chunks"], selected_cat1s, selected_cat2s, selected_cat3s)

                memo_prompt = f"""
你是一位專業、嚴謹的台電新建工程處資深工程祕書與主任工程師。擬稿人：{current_user}。
請根據以下【歷史公文參考範本】（請徹底學習其慣用語法、敘述風格、排版結構與陳核對象習慣）以及【內部合約與資料庫檢索結果】，幫使用者產出一份標準的工程公文。

【文別】：{memo_type}
【受文對象】：{memo_target if memo_target.strip() else "請依歷史公文習慣自動判定"}
【擬稿核心需求】：{memo_topic}

【歷史公文參考範本（請模仿其風格與慣用語）】：
{history_text if history_text.strip() else "（目前尚無歷史公文，請依標準台電工程公文格式產出）"}

【內部合約與資料庫檢索結果】：
{contract_context}

【公文嚴格排版與內容規範】：
1. 必須嚴格遵守標準公文格式，包含【受文者】、【主旨】（開頭空兩格，簡明扼要）、【說明】（分項條列（一）、（二）、（三）並自動空行換行）、【辦法】（明確指示後續辦理動作）。
2. 絕對不允許文字全部擠在一起，段落之間必須有清晰的空行與縮排。
3. 語氣必須符合工程實務之嚴謹、專業、具說服力，並符合擬稿人平常慣用的公文敘述口氣。
"""
                client = genai.Client(api_key=GEMINI_API_KEY)
                memo_res = client.models.generate_content(
                    model="gemini-3.5-flash",
                    contents=memo_prompt,
                    config=types.GenerateContentConfig(temperature=0.2)
                )
                generated_memo_text = memo_res.text if memo_res and memo_res.text else "⚠️ 生成失敗，請稍後再試。"
                st.session_state.last_generated_memo = generated_memo_text
                st.rerun()

    if st.session_state.get("last_generated_memo"):
        st.divider()
        st.markdown("### 📄 產出之標準排版公文預覽：")
        st.markdown(st.session_state.last_generated_memo)

        doc_gen = docx.Document()
        doc_gen.add_heading(f"工程公文草稿 — {current_user} 擬", level=1)
        for line in st.session_state.last_generated_memo.split("\n"):
            if line.strip().startswith("###") or line.strip().startswith("【"):
                doc_gen.add_heading(line.replace("#", "").strip(), level=2)
            else:
                doc_gen.add_paragraph(line)
        
        docx_buf = io.BytesIO()
        doc_gen.save(docx_buf)
        docx_bytes = docx_buf.getvalue()

        st.download_button(
            label="📥 一鍵下載排版完整的 Word 檔 (.docx)",
            data=docx_bytes,
            file_name=f"工程公文草稿_{current_user}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

# ================= 🔎 搜尋分頁（清單總覽 ＋ 點擊折疊展開） =================
with tab_search:
    col_s1, col_s2 = st.columns([3, 1])
    with col_s1:
        search_kw = st.text_input("🔎 輸入關鍵字秒查檔案內文（多個詞用空白隔開，如：管 拔除 預算）：", placeholder="輸入後按 Enter 立即列出檔案清單...")
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
            c3_val = c.get("cat3", "")
            if selected_cat3s and c3_val and c3_val not in selected_cat3s:
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
        st.success(f"🎯 共找到符合條件的段落總計 **{len(matched_results)}** 處！以下為檔案清單總覽：")

        file_grouped_hits = {}
        for hits, item in matched_results:
            fname = item["file_name"]
            if fname not in file_grouped_hits:
                file_grouped_hits[fname] = {
                    "cat1": item["cat1"], "cat2": item["cat2"], "cat3": item.get("cat3", ""),
                    "uploader": item.get("uploader", "系統"), "total_hits": 0, "chunks": []
                }
            file_grouped_hits[fname]["total_hits"] += hits
            file_grouped_hits[fname]["chunks"].append((hits, item))

        st.caption("💡 點擊下方各檔案卡片即可展開／收合檢視詳細命中文內與螢光標記：")

        for fname, gdata in file_grouped_hits.items():
            c3_str = f" > {gdata['cat3']}" if gdata['cat3'] else " (直接放第二層)"
            card_title = f"📄 {fname} ｜ 🗂️ {gdata['cat1']} > {gdata['cat2']}{c3_str} ｜ 命中段落：{len(gdata['chunks'])} 處 (總計 {gdata['total_hits']} 次)"
            
            with st.expander(card_title, expanded=False):
                for idx, (h_cnt, chk) in enumerate(gdata["chunks"], 1):
                    st.markdown(f"**📍 命中位置：`{chk['location']}`** *(本段命中 {h_cnt} 次)*")
                    hl_html = highlight_keywords(chk["content"], raw_kws).replace("\n", "<br>")
                    st.markdown(f"<div style='background-color: #f8f9fa; padding: 10px; border-radius: 6px; font-size: 0.95rem; line-height: 1.6; border-left: 3px solid #ffc107;'>{hl_html}</div>", unsafe_allow_html=True)
                    if idx < len(gdata["chunks"]):
                        st.divider()

with tab_batch:
    st.markdown("#### 📦 檔案批次搬移、改名與刪除總管")
    if not file_info_map:
        st.info("目前資料庫尚無檔案。")
    else:
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
            path_display = f"{fc1} > {fc2}" + (f" > {fc3}" if fc3 else " (直接放第二層)")
            table_data.append({
                "勾選": False,
                "檔案名稱 (雙擊可直接改名)": f,
                "目前所在三層路徑": path_display,
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
                with st.spinner("💾 正在儲存並同步至 GitHub..."):
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
            raw_b_c3 = folder_tree[dest_c1][dest_c2]
            b_c3_opts = ["(無細夾 / 直接放第二層)"] + (raw_b_c3 if isinstance(raw_b_c3, list) else [])
            b_c3_sel = st.selectbox("目標第3層 (可選無)：", b_c3_opts, key="b_c3_sel")
            dest_c3 = "" if b_c3_sel == "(無細夾 / 直接放第二層)" else b_c3_sel
        with m_col4:
            st.write("")
            if st.button(f"📦 一鍵搬移 ({len(selected_rows)} 份)", type="primary", use_container_width=True, disabled=(len(selected_rows) == 0)):
                sel_names = {r["_orig_name"] for r in selected_rows}
                for c in chunks:
                    if c["file_name"] in sel_names:
                        c["cat1"], c["cat2"], c["cat3"] = dest_c1, dest_c2, dest_c3
                with st.spinner("📦 正在搬移並同步至 GitHub..."):
                    save_db(db)
                dest_path_str = f"{dest_c1} > {dest_c2}" + (f" > {dest_c3}" if dest_c3 else "")
                st.success(f"✅ 已將 {len(sel_names)} 份檔案成功搬移至【{dest_path_str}】！")
                st.rerun()
        with m_col5:
            st.write("")
            if st.button(f"🗑️ 批次刪除 ({len(selected_rows)} 份)", use_container_width=True, disabled=(len(selected_rows) == 0)):
                sel_names = {r["_orig_name"] for r in selected_rows}
                db["chunks"] = [c for c in chunks if c["file_name"] not in sel_names]
                with st.spinner("🗑️ 正在刪除並同步至 GitHub..."):
                    save_db(db)
                st.success(f"🗑️ 已批次刪除 {len(sel_names)} 份檔案！")
                st.rerun()

with tab_read:
    if not file_info_map:
        st.info("目前尚無檔案。")
    else:
        sorted_files = sorted(list(file_info_map.keys()))
        default_idx = sorted_files.index(st.session_state.jump_to_file) if st.session_state.jump_to_file in sorted_files else 0
        view_file = st.selectbox("選擇要閱讀的檔案：", sorted_files, index=default_idx)
        vc1, vc2, vc3, vup = file_info_map[view_file]
        path_read_str = f"{vc1} > {vc2}" + (f" > {fc['cat3']}" if vc3 else " (直接放第二層)")
        st.caption(f"📂 三層路徑：**{vc1} > {vc2}{' > ' + vc3 if vc3 else ''}** ｜ 👤 上傳者：**{vup}**")
        for fc in [c for c in chunks if c["file_name"] == view_file]:
            with st.expander(f"📄 {fc['file_name']} — {fc['location']}", expanded=True):
                st.text(fc["content"])
