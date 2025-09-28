import os
import re
import pickle
import pandas as pd
import win32com.client
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from collections import defaultdict

# ===== 基本設定 =====
fit_folder = "合適履歷"
unfit_folder = "不合適履歷"
excel_path = "履歷篩選.xlsx"
output_path = excel_path
time_file = "lasttime.txt"

os.makedirs(fit_folder, exist_ok=True)
os.makedirs(unfit_folder, exist_ok=True)

# ===== 時間工具 =====
def get_last_processed_time():
    if os.path.exists(time_file):
        with open(time_file, "r") as f:
            ts = f.read().strip()
            return datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
    else:
        return datetime.now() - timedelta(days=60)

def save_current_time():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(time_file, "w") as f:
        f.write(now)

# ===== 粗篩條件 / 學歷 bucket =====
def is_fit_candidate(text):
    accepted_genders = ["男", "女"]
    accepted_degrees = ["碩士", "博士", "大學"]
    min_age = 20
    max_age = 35

    age_match = re.search(r"(\d{1,2})歲", text)
    gender_match = re.search(r"(男|女)", text)
    degree_match = re.search(r"(國中\(含\)以下|高中|高職|五專|三專|二專|二技|四技|大學|碩士|博士)", text)
    age = int(age_match.group(1)) if age_match else 99
    gender = gender_match.group(1) if gender_match else ""
    degree = degree_match.group(1) if degree_match else ""

    return (
        (min_age <= age <= max_age) and
        (gender in accepted_genders) and
        (degree in accepted_degrees)
    )

def degree_to_bucket(degree_str: str) -> str:
    if not degree_str:
        return "其他"
    m = re.search(r"(國中\(含\)以下|高中|高職|五專|三專|二專|二技|四技|大學|碩士|博士)", str(degree_str))
    base = m.group(1) if m else ""
    if base in {"高中", "高職"}:
        return "高中職"
    if base in {"二專", "三專", "五專"}:
        return "專科"
    if base in {"二技", "四技", "大學"}:
        return "大學"
    if base in {"碩士"}:
        return "碩士"
    if base in {"博士"}:
        return "博士"
    return "其他"

# ===== 文字處理 / 學校科系解析：正則與工具 =====
_SCHOOL_RE = re.compile(
    r"(?:" 
    r"大學|科技大學|科大|大學校院|專科學校|技術學院|學院|學校|高中|高職|中學|商工|家商|高工|國中|空中大學"
    r"|University|College|Institute(?: of Technology)?|Polytechnic|Academy|Conservatory"
    r"|Business School|Law School|Medical School|Graduate School"
    r"|High School|Secondary School|school|Secondary|Junior College|Community College"
    r")",
    re.I
)

_MAJOR_RE  = re.compile(
    r"(?:"  # 科系關鍵詞（中英）
    r"學位學程|學士班|碩士班|博士班|學系|系所|系|科|所|學程|研究所|學院|日間部|夜間部|進修部|在職專班|學群|學部"
    r"|Department of|School of|Program in|Graduate Institute of|MBA|EMBA|M\.?S\.?|Ph\.?D\.?"
    r")",
    re.I
)

# 日期／期間樣式（給教育背景篩掉日期欄位用）
DATE_LIKE = re.compile(
    r"(\d{2,4}[./-]\d{1,2}(?:[./-]\d{1,2})?|民國\s*\d{2,3}\s*年|\d{2,4}年|\d{1,2}月|至|~|起|止|學年度|期間)"
)
PROGRAM_MAJOR_RE = re.compile(
    r'\b(?:EMBA|E[-\s]?MBA|Executive\s+MBA|MBA|M\.?S\.?|MSc|MSF|MFin|M\.?A\.?|MA|'
    r'M\.?Eng\.?|MEng|MPA|MPP|MPH|MAcc|LL\.?M\.?|LLM|J\.?D\.?|JD|M\.?D\.?|MD|'
    r'Ph\.?D\.?|DPhil|DBA|EdD|DVM|DDS)\b',
    re.I
)

def _move_program_tokens_from_school_to_major(school: str, major: str) -> tuple[str, str]:
    if not school:
        return school, (major or "")
    tokens = PROGRAM_MAJOR_RE.findall(school)
    if tokens:
        # 從學校移除學位詞
        school = PROGRAM_MAJOR_RE.sub("", school)
        # 清掉因移除留下的多空白與空括號
        school = re.sub(r"\(\s*\)", "", school)
        school = re.sub(r"\s{2,}", " ", school).strip()
        # 加到科系（避免重複）
        for t in tokens:
            if not re.search(fr'\b{re.escape(t)}\b', major or "", flags=re.I):
                major = (major + " " + t).strip()
    major = re.sub(r"\s{2,}", " ", (major or "")).strip()
    return school, major

def _looks_like_school(s: str) -> bool:
    return bool(_SCHOOL_RE.search(s or ""))

def _looks_like_major(s: str) -> bool:
    return bool(_MAJOR_RE.search(s or ""))

def _compact_zh_spaces(s: str) -> str:
    if not s:
        return s
    s = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", s)
    s = re.sub(r"^(國立|私立)\s+", r"\1", s)
    s = re.sub(r"\s+(科技大學|師範大學|醫學大學|交通大學|清華大學|海洋大學|體育大學|藝術大學|中醫藥大學|管理學院)$", r"\1", s)
    return s.strip()

def strip_period_parens(s: str) -> str:
    if not s:
        return s

    PAREN  = re.compile(r"[（(]([^）)]*)[）)]")
    KEEP   = re.compile(r"(在職|日間|夜間|假日|進修|專班|學士班|碩士班|博士班|EMBA|MBA)", re.I)
    CREDIT = re.compile(r"(碩士學分|學分班|學分課程|學分)", re.I)
    DATE   = re.compile(r"(\d{2,4}\s*[./-年/]\s*\d{1,2}(?:\s*[./-月/]\s*\d{1,2})?|民國\d{2,3}年|\d{1,2}月|至|~|學年度|期間)")
    END    = re.compile(r"(畢業|肄業)")
    # 位置/地名樣式（僅字母或僅中文字、長度合理）
    LOC    = re.compile(r"^(?:[A-Za-z .,'-]{2,30}|[\u4e00-\u9fff]{1,6})$")

    def _repl(m):
        content = re.sub(r"\s+", "", m.group(1))

        # 只有就讀 / 在學 → 去括號
        if re.fullmatch(r"(就讀|在學)", content):
            return ""
        has_keep   = bool(KEEP.search(content))
        has_credit = bool(CREDIT.search(content))
        has_date   = bool(DATE.search(content))
        has_end    = bool(END.search(content))

        # 純日期/畢業 → 拿掉
        if (has_date or has_end) and not (has_keep or has_credit):
            return ""
        # 地名樣式 → 只有在「不是學分/學分班/學分課程，也不是 KEEP/日期/畢業」時才拿掉
        if LOC.match(content) and not (has_keep or has_credit or has_date or has_end):
            return ""
        # 就讀/在學 + 日期 → 簡化
        if has_date and re.search(r"(就讀|在學)", content):
            return "(就讀中)"
        # 有保留關鍵詞 → 保留（學分類也一起保留）
        if has_keep:
            return f"({KEEP.search(content).group(0)})"
        if has_credit:
            return f"({CREDIT.search(content).group(0)})"
        # 其他未知內容 → 保留原樣
        return m.group(0)

    return PAREN.sub(_repl, s).strip()

def node_text_no_cjk_space(node) -> str:
    s = "".join(node.stripped_strings).replace("\xa0", " ")
    s = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", s)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip()

def parse_school_major(line: str):
    if not line:
        return "", ""

    raw = str(line)
    raw = re.sub(r"[\u3000\xa0]+", " ", raw)
    raw = re.sub(r"[／/、,，\-\–\—\|│•‧・．\.]+", " ", raw)
    raw = re.sub(r"\s{2,}", " ", raw).strip()

    # 先把所有括號做「簡化/去除」後，再來切（會拿掉 (就讀)、(英國) 等）
    def _strip_all_parens_keep_space(s: str) -> str:
        out = strip_period_parens(s)
        return re.sub(r"\s{2,}", " ", out).strip()

    simplified = _strip_all_parens_keep_space(raw)

    # 英文學校尾字（擴充 Business School / Secondary）
    en_school_tail = (
        r"(?:University|Institute of Technology|Institute|Polytechnic|College|Academy|Conservatory|"
        r"Business School|Law School|Medical School|Graduate School|"
        r"High School|Secondary School|Secondary|Junior College|Community College)"
    )

    # e.g., "Henley Business School Marketing and International Management"
    m_en = re.match(rf"^(.*?\b{en_school_tail}\b)\s+(.*)$", simplified, flags=re.I)
    if m_en:
        school = m_en.group(1).strip()
        major  = m_en.group(2).strip()

        # EMBA/MBA 訊號（保險再補一次）
        if re.search(r"\bEMBA\b", simplified, flags=re.I) and "EMBA" not in major:
            major = (major + " EMBA").strip()
        if re.search(r"\bMBA\b",  simplified, flags=re.I) and not re.search(r"\bEMBA\b", simplified, flags=re.I) and "MBA" not in major:
            major = (major + " MBA").strip()

        # 「碩士學分/學分班/學分課程」→ 歸學校欄
        m_credit = re.search(r"(碩士學分|學分班|學分課程)", simplified)
        if m_credit and m_credit.group(1) not in school:
            school = f"{school}（{m_credit.group(1)}）"

        # 標準化大小寫（Business School 等）
        school = re.sub(r"\b(Business|Law|Medical|Graduate)\s+school\b", lambda m: m.group(0).title(), school, flags=re.I)

        # —— 關鍵：把 EMBA/MBA/MS/PhD… 從學校搬去科系 —— #
        school, major = _move_program_tokens_from_school_to_major(school, major)

        school = _compact_zh_spaces(strip_period_parens(school))
        major  = _compact_zh_spaces(strip_period_parens(major))
        return school, major

    # —— 中文/混合情況 —— #
    clean = simplified
    major_tail = r"(?:學位學程|學士班|碩士班|博士班|學系|系所|系|科|所|學程|研究所|學院|日間部|夜間部|進修部|在職專班)"
    m = re.match(rf"^(.*?)\s+(.+?{major_tail})$", clean) or re.match(rf"^(.+?)(.+?{major_tail})$", clean)
    if m:
        school = m.group(1).strip()
        major  = m.group(2).strip()
    else:
        # 允許「學校尾字」後面**沒有空白**也能切（重點改這行）
        school_tail = (
            r"(?:科技大學|師範大學|醫學大學|體育大學|藝術大學|交通大學|清華大學|中醫藥大學|海洋大學|警察大學|"
            r"大學|技術學院|專科學校|學校|高中|高職|中學|商工|家商|工研院|高工|"
            r"University|College|Institute(?: of Technology)?|Polytechnic|Academy|Conservatory|BCIT|"
            r"Business School|Law School|Medical School|Graduate School|"
            r"High School|Secondary School|Secondary|Junior College|Community College)"
        )
        m2 = re.match(rf"^(.*?{school_tail})(.*)$", clean, flags=re.I)  # ← 新：允許零空白
        if m2:
            school = m2.group(1).strip()
            major  = (m2.group(2) or "").strip()
        else:
            m3 = re.match(r"^(.*?(高中|高職|商工|家商|高工|中學))\s*(.+科)?$", clean)
            if m3:
                school = m3.group(1).strip()
                major  = (m3.group(3) or "").strip()
            else:
                school, major = clean, ""

    # EMBA/MBA 訊號（保險再補一次）
    if re.search(r"\bEMBA\b", simplified, flags=re.I) and "EMBA" not in major:
        major = (major + " EMBA").strip()
    if re.search(r"\bMBA\b", simplified, flags=re.I) and not re.search(r"\bEMBA\b", simplified, flags=re.I) and "MBA" not in major:
        major = (major + " MBA").strip()

    # 「碩士學分/學分班/學分課程」→ 學校
    m_credit = re.search(r"(碩士學分|學分班|學分課程)", simplified)
    if m_credit:
        tag = m_credit.group(1)
        if tag in major:
            major = major.replace(tag, "").strip()
        if tag not in school:
            school = re.sub(r"\s{2,}", " ", f"{school} {tag}").strip()

    # —— 關鍵：把 EMBA/MBA/MS/PhD… 從學校搬去科系 —— #
    school, major = _move_program_tokens_from_school_to_major(school, major)

    return _compact_zh_spaces(strip_period_parens(school)), _compact_zh_spaces(strip_period_parens(major))

# ===== 語文能力區塊 + 證照成績 =====
def get_lang_section(full_text: str) -> str:
    text = re.sub(r"[\u3000\xa0]+", " ", full_text)
    m = re.search(r"語文能力\s*[:：]?", text)
    if not m:
        return ""
    start = m.end()
    nxt = re.search(
        r"(技能專長|教育背景|自傳|其他擅長工具|個人資料|求職者希望條件|工作經歷|專長|證照|專業證照|電腦技能|專業技能)",
        text[start:]
    )
    end = start + (nxt.start() if nxt else 4000)
    return text[start:end]

_PROFICIENCY_WORDS = r"(?:精通|流利|母語|中上|中等|中階|初階|初等|基礎|略懂|普通|良好|待加強|A1|A2|B1|B2|C1|C2)"

def _clean_cert_name(name: str) -> str :
    name = re.sub(r"^(?:語文能力|語言能力|中文|國語|華語|英文|英語|日文|日語|韓文|韓語|德文|德語|法文|法語|西班牙文|西語|越文|越語|粵語|客語|台語)\s*[:：\-]?\s*", "", name)
    name = re.sub(rf"^(?:{_PROFICIENCY_WORDS})\s*", "", name)
    name = re.sub(rf"\b{_PROFICIENCY_WORDS}\b", "", name)
    name = re.sub(r"[、,，;；\-–—]+", " ", name)
    name = re.sub(r"\s{2,}", " ", name).strip()
    name = re.sub(r"^[（(]\s*[）)]$", "", name)
    return name

def extract_cert_scores_from_lang(lang_text: str) -> str:
    if not lang_text:
        return ""

    s = re.sub(r"[\u3000\xa0]+", " ", lang_text)
    s = re.sub(r"\s{2,}", " ", s)

    found = []
    seen = set()

    # ---- 英文檢定（保留原則）----
    for m in re.finditer(r"(?:TOEIC|多益)[^\d]{0,8}(\d{3,4})", s, flags=re.I):
        score = m.group(1)
        k = f"toeic {score}"
        if k not in seen:
            seen.add(k); found.append(f"TOEIC {score}")

    for m in re.finditer(r"\bTOEFL(?:\s*iBT)?[^\d]{0,8}(\d{1,3})\b", s, flags=re.I):
        v = int(m.group(1))
        if 0 <= v <= 120:
            k = f"toefl ibt {v}"
            if k not in seen:
                seen.add(k); found.append(f"TOEFL iBT {v}")

    for m in re.finditer(r"\bTOEFL\s*ITP[^\d]{0,8}([3-6]\d{2}|677)\b", s, flags=re.I):
        v = int(m.group(1))
        if 310 <= v <= 677:
            k = f"toefl itp {v}"
            if k not in seen:
                seen.add(k); found.append(f"TOEFL ITP {v}")

    for m in re.finditer(r"\bIELTS[^\d]{0,8}([1-9](?:\.5|\.0)?)\b", s, flags=re.I):
        try:
            v = float(m.group(1))
            if 1.0 <= v <= 9.0:
                k = f"ielts {m.group(1)}"
                if k not in seen:
                    seen.add(k); found.append(f"IELTS {m.group(1)}")
        except:
            pass

    for m in re.finditer(r"(?:\bGEPT\b|全民英檢)\s*[:：\-]?\s*(初級|中級|中高級|高級|初|中|中高|高)\b", s, flags=re.I):
        lvl = m.group(1)
        if lvl in {"初","中","中高","高"}:
            lvl = {"初":"初級","中":"中級","中高":"中高級","高":"高級"}[lvl]
        k = f"gept {lvl}"
        if k not in seen:
            seen.add(k); found.append(f"GEPT {lvl}")

    for m in re.finditer(r"(?:\bJLPT\b|日檢)\s*[:：\-]?\s*(N[1-5])\b", s, flags=re.I):
        lvl = m.group(1).upper()
        k = f"jlpt {lvl}"
        if k not in seen:
            seen.add(k); found.append(f"JLPT {lvl}")

    # ---- 泛用中文證照（切段 → 取冒號後 → 比對）----
    EXCLUDE_NAMES = {"中文","國語","華語","英文","英語","日文","日語","韓文","韓語","德文","德語",
                     "法文","法語","西班牙文","西語","越文","越語","粵語","客語","台語","泰文","泰語","其他外文"}

    tokens = re.split(r"[、,，;；\n\r]+", s)
    LEVEL_PAT = r"(?:[A-Za-z]?\d(?:\.\d)?|[ABC][+\-]?|[Nn][1-5]|[1-9]\d{2,3}|初等|中等|高等|優等)"

    for t in tokens:
        if not t.strip():
            continue
        # 只看最後一個冒號（去掉「中文：」「英文：」這類前綴）
        if "：" in t or ":" in t:
            t = re.split(r"[：:]", t)[-1].strip()

        if "證照" not in t and "檢定" not in t:
            continue

        m = re.search(
            rf"([一-龥A-Za-z0-9（）()《》〈〉\-\s]{{2,50}}?)(?:證照|檢定)\s*(?:[:：\-]?\s*({LEVEL_PAT}))?",
            t
        )
        if not m:
            continue

        raw_name = m.group(1)
        lvl = (m.group(2) or "").upper()

        name = _clean_cert_name(raw_name)
        if not name or name in EXCLUDE_NAMES:
            continue
        if re.fullmatch(_PROFICIENCY_WORDS, name):
            continue

        item = f"{name}證照" + (f" {lvl}" if lvl else "")
        k = item.lower()
        if k not in seen:
            seen.add(k); found.append(item)

    return "、".join(found)

# ===== 教育背景抽取 =====
def extract_education_fields_from_html(html: str):

    def looks_school(s: str) -> bool:
        return _looks_like_school(s)

    def looks_major(s: str) -> bool:
        return _looks_like_major(s)

    def is_date_like(s: str) -> bool:
        return bool(DATE_LIKE.search(s or ""))

    def degree_rank(text: str) -> int:
        if not text: return 0
        t = str(text)
        if re.search(r"博士", t): return 5
        if re.search(r"碩士", t): return 4
        if re.search(r"(大學|學士|四技)", t): return 3
        if re.search(r"(二專|三專|五專|二技)", t): return 2
        if re.search(r"(高中|高職)", t): return 1
        return 0

    def normalize_degree(text: str) -> str:
        if not text: return ""
        t = re.sub(r"\s+", " ", str(text)).strip()
        m = re.search(r"(國中\(含\)以下|高中|高職|五專|三專|二專|二技|四技|大學|學士|碩士|博士)\s*(畢業|肄業|就學中)?", t)
        if not m: return ""
        base = m.group(1).replace("學士", "大學")
        stat = m.group(2) or ""
        return base + stat

    soup = BeautifulSoup(html, "html.parser")

    # 1) 錨定《教育背景》
    anchor = soup.find(string=re.compile("教育背景"))
    if not anchor:
        return ("", "", "", "", "")

    # 2) 找到第一張同時包含「最高學歷／最高／次高」的表
    STOP_RE = re.compile(r"(個人資料|求職者希望條件|工作經歷|技能專長|語文能力|自傳|其他)")
    target_table = None
    node = anchor
    for _ in range(2000):
        node = node.find_next()
        if not node: break
        if getattr(node, "name", "") == "table":
            txt = node_text_no_cjk_space(node)
            if re.search(r"(最高學歷|最高|次高)", txt):
                target_table = node
                break
            if STOP_RE.search(txt):
                break

    degree_text = ""
    hi_cells, se_cells = [], []

    if target_table:
        for tr in target_table.find_all("tr"):
            cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th","td"])]
            if not cells:
                continue
            head = cells[0]
            if re.search(r"最高學歷", head):
                degree_text = cells[-1]
            elif re.fullmatch(r"\s*最高\s*", head) or ("最高" in head and not hi_cells):
                hi_cells = cells[1:]  # 去掉 label
            elif re.fullmatch(r"\s*次高\s*", head) or ("次高" in head and not se_cells):
                se_cells = cells[1:]

    # —— 核心：從值欄中挑學校/科系（忽略日期類欄位） —— #
    def pick_school_major(value_cells):
        # 先清乾淨、去掉日期類欄位
        cells = [v.strip() for v in value_cells if v and v.strip()]
        non_date = [v for v in cells if not is_date_like(v)]

        # 1) 先用 parser 吃「整段」：同一格同時有學校與科系時，這一步會正確切開
        joined = " ".join(non_date) if non_date else " ".join(cells)
        s, m = parse_school_major(joined)
        s = _compact_zh_spaces(s)
        m = _compact_zh_spaces(m)

        # 針對偶發殘留：若學校與科系同時含有 EMBA/MBA，從學校再清一次
        if re.search(r"\b(?:EMBA|MBA)\b", s, flags=re.I) and re.search(r"\b(?:EMBA|MBA)\b", m, flags=re.I):
            s = re.sub(r"\b(?:EMBA|MBA)\b", "", s, flags=re.I)
            s = re.sub(r"\s{2,}", " ", s).strip()

        if s or m:
            return s, m

        # 2) Parser 沒吃到才用舊的保底：A. 學校在前 + 下一格當科系
        for i, v in enumerate(non_date):
            if looks_school(v) and i + 1 < len(non_date):
                s = _compact_zh_spaces(v)
                m = _compact_zh_spaces(non_date[i + 1])
                return s, m

        # 3) 再不行用 B. 「像學校 + 像科系」
        school_cands = [v for v in non_date if looks_school(v)]
        major_cands  = [v for v in non_date if looks_major(v)]
        if school_cands and major_cands:
            return _compact_zh_spaces(school_cands[0]), _compact_zh_spaces(major_cands[-1])

        # 4) 最後保底
        return _compact_zh_spaces(joined), ""

    hi_school = hi_major = se_school = se_major = ""
    if hi_cells:
        hi_school, hi_major = pick_school_major(hi_cells)
    if se_cells:
        se_school, se_major = pick_school_major(se_cells)

    # 3) Fallback（仍限《教育背景》區塊）：挑學歷最高的兩筆
    if not degree_text or not hi_school:
        section_tables = []
        node = anchor
        for _ in range(2000):
            node = node.find_next()
            if not node: break
            if getattr(node, "name", "") == "table":
                txt = node_text_no_cjk_space(node)
                if STOP_RE.search(txt): break
                section_tables.append(node)

        candidates = []
        for table in section_tables:
            for tr in table.find_all("tr"):
                cells = [node_text_no_cjk_space(c) for c in tr.find_all(["th","td"])]
                if not cells: continue
                if len(cells) == 1 and "教育背景" in cells[0]: continue
                line = " ".join(cells)
                if re.search(r"(博士|碩士|大學|學士|四技|五專|三專|二專|二技|高中|高職)", line) and \
                   re.search(r"(學系|系所|系|科|所|學程|研究所|學院)", line):
                    deg = normalize_degree(line)
                    school, major = pick_school_major(cells[1:] if len(cells)>1 else cells)
                    rk = degree_rank(deg or line)
                    if school:
                        candidates.append({"degree": deg, "school": school, "major": major, "rank": rk})

        if candidates:
            candidates.sort(key=lambda x: (-x["rank"]))
            if not degree_text:
                degree_text = candidates[0]["degree"]
            if not hi_school:
                hi_school, hi_major = candidates[0]["school"], candidates[0]["major"]
            if len(candidates) >= 2 and not se_school:
                se_school, se_major = candidates[1]["school"], candidates[1]["major"]

    degree_text = re.sub(r"\s+", " ", degree_text or "").strip()
    hi_school = strip_period_parens(_compact_zh_spaces(hi_school))
    hi_major  = strip_period_parens(_compact_zh_spaces(hi_major))
    se_school = strip_period_parens(_compact_zh_spaces(se_school))
    se_major  = strip_period_parens(_compact_zh_spaces(se_major))

    return (degree_text, hi_school.strip(), hi_major.strip(), se_school.strip(), se_major.strip())

# ===== 解析單一 HTML 檔（抽欄位） =====
def extract_fields_from_html(file_path, folder_label):
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            html = f.read()
        soup = BeautifulSoup(html, "html.parser")
        full_text = soup.get_text(" ", strip=True)
        full_text = full_text.replace("\xa0", " ").replace("\u3000", " ")
        full_text = re.sub(r"[\s\u3000]+", " ", full_text).strip()
        tds = soup.find_all("td")
    except Exception as e:
        print(f"❌ 無法解析 HTML 檔：{file_path} → {e}")
        return None

    # 基本資料
    name, age, gender = "", "", ""
    m = re.search(r"((?:[A-Za-z]{2,20}\s*){1,5}|[\u4e00-\u9fa5]{2,4})\s*(\d{1,2})\s*歲\s+(男|女)", full_text)
    if m:
        name = m.group(1).strip()
        age = m.group(2)
        gender = m.group(3)
    print(f"✅ 抓到：name='{name}', age='{age}', gender='{gender}'")

    # 應徵快照 / 應徵日期
    snapshot_match = re.search(r"應徵快照[:：]?\s*(\d{4}[-/]\d{2}[-/]\d{2})\s*(\d{2}:\d{2})", full_text)
    apply_date_match = re.search(r"應徵日期[:：]?\s*(\d{4}[-/]\d{2}[-/]\d{2})\s*(\d{2}:\d{2})", full_text)
    snapshot = pd.to_datetime(" ".join(snapshot_match.groups()), errors="coerce") if snapshot_match else pd.NaT
    apply_date = pd.to_datetime(" ".join(apply_date_match.groups()), errors="coerce") if apply_date_match else pd.NaT

    # 應徵職務與代碼
    job_match = re.search(r"應徵職務[:：]?\s*(.{2,40}?)\s*(自我推薦|應徵問題|希望職稱|工作經歷|居住地|E-mail|聯絡電話|聯絡方式|$)", full_text)
    job = job_match.group(1).strip() if job_match else ""
    code_match = re.search(r"代碼[:：]?\s*(\d{8,15})", full_text)
    code = code_match.group(1) if code_match else ""

    # 教育背景（**用我們的函式**）
    highest_degree, hi_school, hi_major, se_school, se_major = extract_education_fields_from_html(html)
    degree_bucket = degree_to_bucket(highest_degree)

    # 總年資
    total_exp = ""
    for i in range(len(tds) - 1):
        label = tds[i].get_text(strip=True)
        value = tds[i + 1].get_text(strip=True)
        if label == "總年資":
            if "年次" in value:
                continue
            match = re.search(r"([0-9]{1,2}(?:~[0-9]{1,2})?年(?:[（(]?\s*含\s*[）)]?)?\s*(?:以上|以下)?)", value)
            total_exp = match.group(1) if match else value
            break

    if not total_exp:
        for tr in soup.find_all("tr"):
            th = tr.find("th")
            td = tr.find("td")
            if th and "總年資" in th.get_text(strip=True):
                value = td.get_text(strip=True)
                if "年次" in value:
                    continue
                match = re.search(r"([0-9]{1,2}(?:~[0-9]{1,2})?年(?:[（(]?\s*含\s*[）)]?)?\s*(?:以上|以下)?)", value)
                total_exp = match.group(1) if match else value
                break

    if not total_exp or re.search(r"無.*工作經驗", total_exp):
        total_exp = "無"

    # 工作經歷年數
    work_exp = ""
    for i in range(len(tds) - 1):
        label = tds[i].get_text(strip=True)
        value = tds[i + 1].get_text(strip=True)
        if label == "工作經歷":
            if "年次" in value:
                continue
            match = re.search(r"([0-9]{1,2}(?:~[0-9]{1,2})?年(?:[（(]?\s*含\s*[）)]?)?\s*(?:以上|以下)?)", value)
            work_exp = match.group(1) if match else value
            break

    if not work_exp:
        for tr in soup.find_all("tr"):
            th = tr.find("th")
            td = tr.find("td")
            if th and "工作經歷" in th.get_text(strip=True):
                value = td.get_text(strip=True)
                if "年次" in value:
                    continue
                match = re.search(r"([0-9]{1,2}(?:~[0-9]{1,2})?年(?:[（(]?\s*含\s*[）)]?)?\s*(?:以上|以下)?)", value)
                work_exp = match.group(1) if match else value
                break

    if not work_exp or re.search(r"無.*工作經驗", work_exp):
        work_exp = "無"

    # —— 語文能力 + 證照成績 —— 
    lang_section = get_lang_section(full_text)   

    langs = [lang for lang in ["中文","英文","日文","台語","德文","法文","西班牙文","越文",
                            "泰文","韓文","粵語","客語","其他外文"] if lang in lang_section]

    cert_scores = extract_cert_scores_from_lang(lang_section) 

    # 其他欄位
    address_match = re.search(r"居住地[:：]?\s*(\d{3,5}\s*\S{3,}.*?)\s+(E-mail|聯絡電話|聯絡方式)", full_text)
    address = address_match.group(1).strip() if address_match else ""
    job_status_match = re.search(r"就業狀態[:：]?\s*(仍在職|待業中)", full_text)
    job_status = job_status_match.group(1) if job_status_match else ""

    # 待遇
    salary_text, salary_number = "", ""
    salary_match = re.search(r"希望待遇[:：]?\s*(.{1,80})", full_text)
    if salary_match:
        raw_salary = salary_match.group(1).strip()
        first_part = re.split(r"(可上班日|上班時段|工作經歷|總年資|聯絡電話|聯絡方式|E-mail)", raw_salary)[0]
        number_match = re.search(r"\d{4,6}", first_part.replace(",", ""))
        if number_match:
            salary_number = number_match.group(0)
        if "面議" in first_part:
            salary_text = "面議"
        elif "依公司規定" in first_part:
            salary_text = "依公司規定"
        elif "月薪" in first_part:
            salary_text = "月薪"
        elif "時薪" in first_part:
            salary_text = "時薪"
        elif "年薪" in first_part:
            salary_text = "年薪"

    # 連絡資訊
    email = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", full_text)
    phone = re.search(r"聯絡(電話|方式)[:：]?\s*(09\d{2}-?\d{3}-?\d{3})", full_text)

    # 主旨與應徵方式來源
    subject_match = re.search(r"<!--subject:(.*?)-->", html)
    subject = subject_match.group(1).strip() if subject_match else ""
    apply_text = subject + " " + full_text
    if "主動應徵履歷" in apply_text:
        apply_method = "主動應徵"
    elif any(k in apply_text for k in ["邀請您應徵", "誠摯邀請", "邀請您加入", "主動邀請"]):
        apply_method = "企業邀約"
    elif any(k in apply_text for k in ["回覆邀請", "接受邀請", "回覆面試邀請"]):
        apply_method = "回覆邀約"
    else:
        apply_method = "其他"

    source_channel = "104人力銀行" if "104" in subject else "1111人力銀行" if "1111" in subject else "其他"

    return {
        "資料夾名稱": folder_label,
        "應徵日期": snapshot.date() if not pd.isna(snapshot) else (apply_date.date() if not pd.isna(apply_date) else pd.NaT),
        "應徵管道": source_channel,
        "應徵方式": apply_method,
        "姓名": name,
        "代碼": code,
        "應徵職務": job,
        "性別": gender,
        "年齡": age,
        "學歷": degree_bucket,
        "最高學歷": highest_degree,
        "最高學歷學校": hi_school,
        "最高學歷科系": hi_major,
        "次高學歷學校": se_school,
        "次高學歷科系": se_major,
        "工作經歷": work_exp,
        "總年資": total_exp,
        "居住地": address,
        "就業狀態": job_status,
        "希望待遇": salary_text,
        "薪水": salary_number,
        "手機": phone.group(2) if phone else "",
        "E-mail": email.group(0) if email else "",
        "語言能力": "、".join(langs),
        "證照成績": cert_scores
    }

# ===== 批次載入檔案（兩資料夾） =====
def load_all_resumes(folder, label):
    results = []
    failed = 0
    for fname in os.listdir(folder):
        if fname.lower().endswith(".html"):
            path = os.path.join(folder, fname)
            fields = extract_fields_from_html(path, label)
            if fields:
                results.append(fields)
            else:
                failed += 1
                print(f"❌ 無法解析：{fname}")
    print(f"📂 資料夾「{label}」共載入：✅ {len(results)} 筆成功，❌ {failed} 筆失敗")
    return results

# ===== 匯出 Excel（保留最舊：keep='first'） =====
def update_excel_from_folder():

    fit = load_all_resumes(fit_folder, "合適履歷")
    unfit = load_all_resumes(unfit_folder, "不合適履歷")
    all_new = fit + unfit

    if not all_new:
        print("❌ 沒有讀到任何履歷資料")
        return

    new_df = pd.DataFrame(all_new)
    new_df["應徵日期"] = pd.to_datetime(new_df["應徵日期"], errors="coerce")

    if os.path.exists(excel_path):
        old_df = pd.read_excel(excel_path)

        # ★ 舊欄位同步改名，避免同時出現「多益成績」與「證照成績」
        if "多益成績" in old_df.columns and "證照成績" not in old_df.columns:
            old_df = old_df.rename(columns={"多益成績": "證照成績"})

        old_df["應徵日期"] = pd.to_datetime(old_df["應徵日期"], errors="coerce")
    else:
        old_df = pd.DataFrame(columns=new_df.columns)

    combined = pd.concat([old_df, new_df], ignore_index=True)

    combined["代碼"] = combined["代碼"].astype(str).str.strip()
    combined["應徵職務"] = combined["應徵職務"].astype(str).str.replace(r"[\u3000\s]+", "", regex=True)

    try:
        with open("apply_count.pkl", "rb") as f:
            count_dict = pickle.load(f)
    except Exception:
        count_dict = defaultdict(int)

    combined.sort_values("應徵日期", ascending=True, inplace=True)
    final_df = combined.drop_duplicates(subset=["代碼", "應徵職務"], keep="first").copy()

    final_df["應徵次數"] = final_df.apply(
        lambda r: str(count_dict.get((r["代碼"], r["應徵職務"]), 1)),
        axis=1
    )

    final_df = final_df.fillna("").astype(str)
    final_df["應徵日期"] = pd.to_datetime(final_df["應徵日期"], errors="coerce").dt.date
    final_df.sort_values("應徵日期", ascending=True, inplace=True)

    final_df.to_excel(output_path, index=False)
    print(f"✅ 輸出完成：{output_path}")

# ===== 讀取 Outlook 並分類（保留最舊檔） =====
def fetch_and_classify_emails():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    last_time = get_last_processed_time()
    restriction = f"[ReceivedTime] >= '{last_time.strftime('%m/%d/%Y %I:%M %p')}'"
    messages = messages.Restrict(restriction)
    apply_count_dict = defaultdict(int) 

    for msg in messages:
        if not hasattr(msg, "MessageClass") or msg.MessageClass != "IPM.Note":
            continue

        subject = msg.Subject or ""
        try:
            received_time = msg.ReceivedTime.replace(tzinfo=None)
        except Exception as e:
            print(f"⚠️ 無法讀取時間：{subject} → {e}")
            continue

        if not ("104應徵履歷" in subject or "104轉寄履歷" in subject):
            continue
        if received_time <= last_time:
            print(f"⏭ 跳過舊信件：{subject}（{received_time} <= {last_time}）")
            continue

        try:
            html = msg.HTMLBody if hasattr(msg, "HTMLBody") else ""
            soup = BeautifulSoup(html, "html.parser")
            text = soup.get_text(" ", strip=True)
            fit = is_fit_candidate(text)

            folder = fit_folder if fit else unfit_folder
            safe_subject = re.sub(r"[\\/:*?\"<>|]", "_", subject).strip()
            file_name = f"{safe_subject}.html"
            save_path = os.path.join(folder, file_name)

            if "<!--subject:" not in html:
                html = f"<!--subject:{subject}-->" + html

            code_match = re.search(r"代碼[:：]?\s*(\d{8,15})", text)
            job_match = re.search(r"應徵職務[:：]?\s*(.{2,40}?)\s*(自我推薦|應徵問題|希望職稱|工作經歷|居住地|E-mail|聯絡電話|聯絡方式|$)", text)
            code = code_match.group(1) if code_match else ""
            job = job_match.group(1).strip() if job_match else ""
            apply_count_dict[(code, job)] += 1

            def extract_apply_time(txt):
                snapshot = re.search(r"應徵快照[:：]?\s*(\d{4}[-/]\d{2}[-/]\d{2})\s*(\d{2}:\d{2})", txt)
                apply = re.search(r"應徵日期[:：]?\s*(\d{4}[-/]\d{2}[-/]\d{2})\s*(\d{2}:\d{2})", txt)
                return pd.to_datetime(" ".join(snapshot.groups()), errors="coerce") if snapshot else (
                    pd.to_datetime(" ".join(apply.groups()), errors="coerce") if apply else pd.NaT
                )
            new_time = extract_apply_time(text)

            # 儲存或更新（保留最舊）
            if os.path.exists(save_path):
                with open(save_path, "r", encoding="utf-8") as f:
                    old_html = f.read()
                old_text = BeautifulSoup(old_html, "html.parser").get_text(" ", strip=True)
                old_time = extract_apply_time(old_text)

                # 只在新進來的履歷「更舊」時才覆蓋 → 永遠保留最舊的那封
                should_replace = False
                if pd.isna(old_time) and not pd.isna(new_time):
                    # 舊檔看不出時間，但新的有時間 → 視情況保留舊或換成更舊的；這裡保守：不換
                    should_replace = False
                elif not pd.isna(new_time) and not pd.isna(old_time):
                    should_replace = new_time < old_time
                else:
                    # 兩邊都 NaT 或新的是 NaT → 不動
                    should_replace = False

                if should_replace:
                    with open(save_path, "w", encoding="utf-8") as f:
                        f.write(html)
                    print(f"🔁 以更舊的履歷覆蓋：{file_name}（{new_time} < {old_time}）")
                else:
                    print(f"⏭ 保留既有履歷（較舊版本優先）：{file_name}")
            else:
                with open(save_path, "w", encoding="utf-8") as f:
                    f.write(html)
                print(f"✅ 成功儲存 .html 檔：{file_name}")

        except Exception as e:
            print(f"❌ 錯誤處理信件：{subject} → {e}")
    save_current_time()

    with open("apply_count.pkl", "wb") as f:
        pickle.dump(apply_count_dict, f)

# ===== 入口 =====
if __name__ == "__main__":
    fetch_and_classify_emails()
    update_excel_from_folder()
