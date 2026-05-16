"""
瞎检索 — 文献核查与格式化工具
漫画手绘风格 GUI

打包命令:
  python -m PyInstaller --clean --onefile --windowed --icon xia_jiansuo_avatar.ico --name 瞎检索 xia_jiansuo_fixed_v5.py

依赖:
  pip install habanero python-docx pyinstaller
"""

import re
import time
import csv
import math
import html
import base64
import difflib
import unicodedata
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path

try:
    from habanero import Crossref
    HAS_CROSSREF = True
except ImportError:
    HAS_CROSSREF = False

try:
    from docx import Document as DocxDocument
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ══════════════════════════════════════════════════════════════
# 漫画风格色板  —  纸白 / 墨黑 / 铅笔灰
# ══════════════════════════════════════════════════════════════
INK        = "#0D0D0D"   # 接近纯黑的墨水色
INK_SOFT   = "#2A2A2A"   # 稍浅的墨
PAPER      = "#FAF9F6"   # 手稿纸白（微黄）
PAPER_DARK = "#F0EDE6"   # 略深一点的纸色，用于面板区分
PENCIL     = "#888888"   # 铅笔灰（次要文字）
PENCIL_LT  = "#BBBBBB"   # 浅灰（占位符 / disabled）
RULE_LINE  = "#CCCCCC"   # 淡灰线（边框）
OK_INK     = "#1A1A1A"   # 核实：深墨 + ✓ 符号
WARN_INK   = "#555555"   # 待确认：中灰
ERR_INK    = "#999999"   # 未找到：浅灰（淡出）

# 字体：优先用系统上有的手写/漫画感字体，兜底回中文宋体/黑体
def _font(*names, size=10, bold=False):
    weight = "bold" if bold else "normal"
    return (names[0], size, weight)

F_TITLE  = _font("DFKai-SB", "STKaiti", "KaiTi", "SimSun", size=18, bold=True)
F_SUB    = _font("Microsoft YaHei UI", "SimHei", size=9)
F_H2     = _font("Microsoft YaHei UI", "SimHei", size=11, bold=True)
F_BODY   = _font("Microsoft YaHei UI", "SimSun", size=10)
F_MONO   = _font("Consolas", "Courier New", size=9)
F_LOG    = _font("Consolas", "Courier New", size=8)
F_BTN    = _font("Microsoft YaHei UI", "SimHei", size=10, bold=True)
F_BIG    = _font("Microsoft YaHei UI", "SimHei", size=12, bold=True)

# 内置头像 Logo：用于标题栏和窗口图标，打包后不需要额外图片文件
APP_ICON_PNG_BASE64 = """
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
"""

# ══════════════════════════════════════════════════════════════
# 核心逻辑（英文 CrossRef + 中文本地解析）
# ══════════════════════════════════════════════════════════════

def clean_text(s: str) -> str:
    """清洗 CrossRef/网页返回的 HTML、乱码空格、上下标标签、PDF 断词等。"""
    if not s:
        return ""
    s = html.unescape(str(s))
    s = re.sub(r"<\s*sub\s*>\s*(.*?)\s*<\s*/\s*sub\s*>", r"\1", s, flags=re.I)
    s = re.sub(r"<\s*sup\s*>\s*(.*?)\s*<\s*/\s*sup\s*>", r"\1", s, flags=re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u00a0", " ").replace("\ufeff", " ")
    s = s.replace("ﬁ", "fi").replace("ﬂ", "fl").replace("ﬀ", "ff").replace("ﬃ", "ffi").replace("ﬄ", "ffl")
    s = re.sub(r"[‐‑‒–—―]", "-", s)
    # PDF/Word 复制出来常见的断词：expos- ure -> exposure, attrib- utable -> attributable
    s = re.sub(r"([A-Za-z])[-]\s+([a-z])", r"\1\2", s)
    # 少量常见粘连词/化学式，改善题名相似度
    s = re.sub(r"\bofthe\b", "of the", s, flags=re.I)
    s = re.sub(r"\boflow\b", "of low", s, flags=re.I)
    s = re.sub(r"\bheath\b", "health", s, flags=re.I)
    s = re.sub(r"\bCO\s*2\b", "CO2", s, flags=re.I)
    s = re.sub(r"\bSO\s*2\b", "SO2", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip()
    return s

STOPWORDS = {
    "the", "a", "an", "of", "and", "or", "for", "to", "in", "on", "with", "by",
    "from", "using", "based", "study", "analysis", "article", "journal", "review"
}

def _norm_title(s: str) -> str:
    s = clean_text(s).lower()
    s = re.sub(r"https?://\S+|doi\s*:?\s*\S+", " ", s)
    s = re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _tokens(s: str) -> list:
    s = _norm_title(s)
    toks = re.findall(r"[a-z0-9]+|[\u4e00-\u9fff]", s)
    return [t for t in toks if t not in STOPWORDS and (len(t) > 1 or re.match(r"[\u4e00-\u9fff]", t))]

def _smart_sentences(s: str) -> list:
    """按句号拆，但尽量不把 J. / X. 这种作者名缩写拆开。"""
    s = clean_text(s)
    parts = re.split(r"(?<=[a-z0-9\)\]])\.\s+(?=[A-Z\u4e00-\u9fff])", s)
    return [p.strip(" .;，,：:") for p in parts if p.strip(" .;，,：:")]

def extract_doi(ref: str) -> str:
    ref = clean_text(ref)
    m = re.search(r"\b(10\.\d{4,9}/[-._;()/:A-Z0-9]+)\b", ref, flags=re.I)
    if not m:
        return ""
    doi = m.group(1).rstrip('.。;,，)]"')
    doi = re.sub(r"^doi\s*:?\s*", "", doi, flags=re.I)
    return doi.lower()

def _extract_years(ref: str) -> list:
    return re.findall(r"(?:19|20)\d{2}", ref or "")

def _extract_ref_year(ref: str) -> str:
    """
    提取“出版年份”，不能简单取第一个年份。
    例如题名里有 1990-2021 / 2010 to 2100 时，真正出版年份通常在 Journal 2024; 403 这种位置。
    """
    ref = clean_text(ref)

    # Vancouver/期刊缩写格式：Lancet 2024; 403: 2162-2203
    m = re.search(r"\b((?:19|20)\d{2})\s*[;,]\s*\d", ref)
    if m:
        return m.group(1)

    # APA 格式：(2020). Title
    m = re.search(r"\(\s*((?:19|20)\d{2})[a-z]?\s*\)", ref)
    if m:
        return m.group(1)

    # Chicago/作者年格式：Author. 2020. “Title”
    m = re.search(r"\b((?:19|20)\d{2})[a-z]?\b\s*[\.。]\s*[“\"A-Z]", ref)
    if m:
        return m.group(1)

    # 兜底：若有多个年份，倾向靠后的年份；题名里出现年份区间时通常不是出版年
    years = _extract_years(ref)
    if not years:
        return ""
    return years[-1] if len(years) >= 2 else years[0]

def _looks_like_author_segment(seg: str) -> bool:
    """判断一个句段是否像作者列表。"""
    seg = clean_text(seg)
    if not seg:
        return False
    low = seg.lower()
    if "et al" in low:
        return True
    # 姓 + 首字母，或 APA 的 姓, 名缩写
    if re.search(r"\b[A-Z][A-Za-z'\-]{2,}\s+[A-Z]{1,4}\b", seg):
        return True
    if re.search(r"\b[A-Z][A-Za-z'\-]{2,}\s*,\s*[A-Z]\.", seg):
        return True
    if ("," in seg or " and " in low or "&" in seg) and len(_tokens(seg)) <= 22:
        return True
    return False


def _author_head(ref: str) -> str:
    """截出参考文献开头的作者段。核心原则：作者段必须在出版年份之前。"""
    ref = clean_text(ref)
    ref = re.sub(r"^\s*(?:\[\s*\d+\s*\]|\(?\d+\)?[\.)])\s+", "", ref)
    # Vancouver: Bistline JE and Rose SK. Title. Journal Year;...
    vm = re.match(r"^(.{5,180}?\b[A-Z]{1,4}\.)\s+[A-Z][a-z]", ref)
    if vm and _looks_like_author_segment(vm.group(1)):
        return vm.group(1).strip(" .;，,：:")
    m = re.search(r"\(?\b(?:19|20)\d{2}[a-z]?\b\)?", ref)
    if m and m.start() > 2:
        return ref[:m.start()].strip(" .;，,：:")
    segs = _smart_sentences(ref)
    if segs and _looks_like_author_segment(segs[0]):
        return segs[0][:220].strip(" .;，,：:")
    return ref[:160].strip(" .;，,：:")


def _extract_author_names(ref: str) -> set:
    """粗略提取参考文献开头处的英文姓氏，用于二次校验，避免错配。"""
    head = _author_head(ref)
    head = re.split(r"\bet\s+al\b|\.\.\.", head, flags=re.I)[0]
    names = set()
    head = re.sub(r"\s+(?:and|&)\s+", ", ", head, flags=re.I)
    parts = [p.strip(" .;，") for p in re.split(r",", head) if p.strip(" .;，")]

    i = 0
    while i < len(parts):
        part = parts[i]
        if re.fullmatch(r"[A-Z][A-Za-z'\-]{1,}", part):
            if i + 1 < len(parts) and re.search(r"\b[A-Z]\.?(?:\s*[A-Z]\.?)?\b", parts[i + 1]):
                if part.lower() not in {"and", "the", "press", "university", "journal"}:
                    names.add(part.lower())
                i += 2
                continue
        i += 1

    for part in parts:
        if re.match(r"^[A-Z][A-Za-z'\-]{1,}\s+(?:[A-Z]{1,4}\.?\s*){1,4}$", part):
            fam = re.match(r"^([A-Z][A-Za-z'\-]{1,})", part).group(1)
            if fam.lower() not in {"and", "the", "press", "university", "journal"}:
                names.add(fam.lower())

    if not names:
        for cand in re.findall(r"\b[A-Z][A-Za-z'\-]{2,}\b", head):
            if cand.lower() not in {"and", "the", "press", "university", "journal", "policy", "retrieved"}:
                names.add(cand.lower())
    return names


def _extract_author_objects_from_ref(ref: str) -> list:
    """从原始文献中提取作者对象，用于 CrossRef 候选缺作者时回填，避免 Unknown Author。"""
    head = _author_head(ref)
    head = re.split(r"\bet\s+al\b|\.\.\.", head, flags=re.I)[0]
    head = re.sub(r"\s+(?:and|&)\s+", ", ", head, flags=re.I)
    parts = [p.strip(" .;，") for p in re.split(r",", head) if p.strip(" .;，")]
    authors = []
    i = 0
    while i < len(parts):
        part = parts[i]
        if re.fullmatch(r"[A-Z][A-Za-z'\-]{1,}", part) and i + 1 < len(parts):
            nxt = parts[i + 1]
            if re.search(r"\b[A-Z]\.?(?:\s*[A-Z]\.?)?\b", nxt):
                given = " ".join(re.findall(r"[A-Z]", nxt))
                authors.append({"family": part, "given": given})
                i += 2
                continue
        m = re.match(r"^([A-Z][A-Za-z'\-]{1,})\s+([A-Z]{1,4})\.?$", part)
        if m:
            authors.append({"family": m.group(1), "given": " ".join(m.group(2))})
        i += 1
    if not authors:
        for fam in sorted(_extract_author_names(ref))[:20]:
            authors.append({"family": fam[:1].upper() + fam[1:], "given": ""})
    return authors[:20]


def _is_supplement_doi(doi: str) -> bool:
    doi = (doi or "").lower().strip()
    return bool(re.search(r"(?:\.s\d{3,}$|/suppl_file|supplement|supporting)", doi))


def _main_doi_from_supplement(doi: str) -> str:
    """ACS 常见补充材料 DOI: 10.1021/xxx.s001 -> 主文献 DOI: 10.1021/xxx。"""
    doi = (doi or "").lower().strip()
    return re.sub(r"\.s\d{3,}$", "", doi)


def _looks_like_non_crossref_work(ref: str) -> bool:
    """报告、网页、政策简报等通常不适合强行 CrossRef 匹配。"""
    text = clean_text(ref).lower()
    if extract_doi(text):
        return False
    hard_patterns = [
        "policy brief", "retrieved on", "cems", "china leadership monitor",
        "leadership monitor", "working paper", "world poverty institute working paper",
        "http://", "https://"
    ]
    return any(p in text for p in hard_patterns)
def has_chinese(s: str) -> bool:
    return bool(re.search(r"[\u4e00-\u9fff]", s or ""))

def detect_reference_type(ref: str) -> str:
    """粗略判断参考文献类型：用于选择检索/解析通道。"""
    ref = clean_text(ref)
    doi = extract_doi(ref)
    cn = has_chinese(ref)
    if cn and doi:
        return "中文+DOI"
    if doi:
        return "DOI"
    if cn:
        if re.search(r"\[[JMDCRN]\]", ref, flags=re.I):
            return "中文GB/T"
        return "中文"
    return "英文/其他"

def _split_authors_cn(authors: str) -> list:
    authors = clean_text(authors)
    authors = re.sub(r"^\s*(?:\[\s*\d+\s*\]|\(?\d+\)?[\.)])\s+", "", authors)
    authors = authors.replace("，", ",").replace("、", ",").replace(";", ",").replace("；", ",")
    authors = re.sub(r"\s*(等|et\s+al\.?)\s*$", "", authors, flags=re.I)
    parts = [a.strip() for a in authors.split(",") if a.strip()]
    # 中文姓名一般 2-4 字；如果切得过碎，保留原段
    good = []
    for a in parts:
        a = re.sub(r"\s+", "", a)
        if a:
            good.append(a)
    return good[:20]

def _clean_cn_field(s: str) -> str:
    return clean_text(s).strip(" .。;；,，:：")

def parse_chinese_reference(ref: str) -> dict:
    """
    中文/GB/T 7714 本地解析。目标不是数据库核验，而是稳定拆字段，避免乱匹配英文文献。
    支持：作者. 题名[J]. 期刊, 年, 卷(期): 页码.
          作者. 题名[M]. 出版地: 出版社, 年.
          作者. 题名[D]. 学校, 年.
          作者. 题名[C]//会议论文集...
    """
    original = clean_text(ref)
    text = original
    text = re.sub(r"^\s*(?:\[\s*\d+\s*\]|\(?\d+\)?[\.)])\s+", "", text)
    doi = extract_doi(text)
    url_m = re.search(r"https?://\S+", text, flags=re.I)
    url = url_m.group(0).rstrip(".。;,，") if url_m else ""
    # 去掉 DOI/URL，避免干扰字段解析
    work = re.sub(r"https?://\S+", "", text, flags=re.I)
    work = re.sub(r"(?:doi\s*[:：]?\s*)?10\.\d{4,9}/[-._;()/:A-Z0-9]+", "", work, flags=re.I)
    work = _clean_cn_field(work)

    result = {
        "original": original,
        "source": "本地解析",
        "ref_type": detect_reference_type(original),
        "authors_text": "",
        "authors_list": [],
        "title": "",
        "doc_type": "",
        "journal": "",
        "booktitle": "",
        "publisher": "",
        "place": "",
        "school": "",
        "year": "",
        "volume": "",
        "issue": "",
        "pages": "",
        "doi": doi,
        "url": url,
    }

    # 标准 GB/T：作者. 题名[J]. 期刊, 年, 卷(期): 页码
    m = re.match(r"^(.+?)[\.。]\s*(.+?)\s*\[\s*([JMDCRN])\s*\]\s*(?://)?\s*[\.。]?\s*(.*)$", work, flags=re.I)
    if m:
        authors, title, dtype, rest = m.groups()
        result["authors_text"] = _clean_cn_field(authors)
        result["authors_list"] = _split_authors_cn(authors)
        result["title"] = _clean_cn_field(title)
        result["doc_type"] = dtype.upper()
        rest = _clean_cn_field(rest)

        if dtype.upper() == "J":
            jm = re.search(
                r"^(.+?)[,，]\s*((?:19|20)\d{2}[a-z]?)\s*(?:[,，]\s*([0-9A-Za-z]+)\s*(?:[\(（]\s*([^\)）]+)\s*[\)）])?)?\s*(?:[:：]\s*([0-9A-Za-z\-–—]+))?.*$",
                rest
            )
            if jm:
                result["journal"] = _clean_cn_field(jm.group(1))
                result["year"] = _clean_cn_field(jm.group(2))
                result["volume"] = _clean_cn_field(jm.group(3) or "")
                result["issue"] = _clean_cn_field(jm.group(4) or "")
                result["pages"] = _clean_cn_field(jm.group(5) or "")
            else:
                result["journal"] = rest
                result["year"] = _extract_ref_year(rest)
        elif dtype.upper() == "M":
            mm = re.search(r"^([^:：,，。]+)\s*[:：]\s*([^,，。]+)[,，]\s*((?:19|20)\d{2})", rest)
            if mm:
                result["place"] = _clean_cn_field(mm.group(1))
                result["publisher"] = _clean_cn_field(mm.group(2))
                result["year"] = _clean_cn_field(mm.group(3))
            else:
                result["publisher"] = rest
                result["year"] = _extract_ref_year(rest)
        elif dtype.upper() == "D":
            dm = re.search(r"^(.+?)[,，]\s*((?:19|20)\d{2})", rest)
            if dm:
                result["school"] = _clean_cn_field(dm.group(1))
                result["year"] = _clean_cn_field(dm.group(2))
            else:
                result["school"] = rest
                result["year"] = _extract_ref_year(rest)
        elif dtype.upper() == "C":
            cm = re.search(r"^(.+?)[,，]\s*((?:19|20)\d{2})", rest)
            if cm:
                result["booktitle"] = _clean_cn_field(cm.group(1))
                result["year"] = _clean_cn_field(cm.group(2))
            else:
                result["booktitle"] = rest
                result["year"] = _extract_ref_year(rest)
        else:
            result["journal"] = rest
            result["year"] = _extract_ref_year(rest)

        if not result["year"]:
            result["year"] = _extract_ref_year(original)
        return result

    # 非标准中文：作者. 题名. 期刊, 年, 卷(期): 页码
    parts = re.split(r"[\.。]\s*", work, maxsplit=2)
    if len(parts) >= 2:
        authors = parts[0]
        title = parts[1]
        rest = parts[2] if len(parts) >= 3 else ""
        if has_chinese(title) and len(title) >= 4:
            result["authors_text"] = _clean_cn_field(authors)
            result["authors_list"] = _split_authors_cn(authors)
            result["title"] = _clean_cn_field(title)
            result["doc_type"] = "J" if re.search(r"[,，]\s*(?:19|20)\d{2}", rest) else ""
            jm = re.search(
                r"^(.+?)[,，]\s*((?:19|20)\d{2}[a-z]?)\s*(?:[,，]\s*([0-9A-Za-z]+)\s*(?:[\(（]\s*([^\)）]+)\s*[\)）])?)?\s*(?:[:：]\s*([0-9A-Za-z\-–—]+))?.*$",
                rest
            )
            if jm:
                result["journal"] = _clean_cn_field(jm.group(1))
                result["year"] = _clean_cn_field(jm.group(2))
                result["volume"] = _clean_cn_field(jm.group(3) or "")
                result["issue"] = _clean_cn_field(jm.group(4) or "")
                result["pages"] = _clean_cn_field(jm.group(5) or "")
            else:
                result["journal"] = _clean_cn_field(rest)
                result["year"] = _extract_ref_year(original)
            return result

    # 兜底：至少给出年份和可能题名，不乱查 CrossRef
    result["year"] = _extract_ref_year(original)
    # 找一个带中文、长度较像题名的片段
    segs = re.split(r"[\.。]", work)
    cands = [seg.strip() for seg in segs if has_chinese(seg) and 4 <= len(seg.strip()) <= 120]
    if cands:
        # 避免第一段作者过短，优先第二段
        result["title"] = _clean_cn_field(cands[1] if len(cands) > 1 else cands[0])
    return result

def _local_chinese_result(parsed: dict, confidence: str = "uncertain", reason: str = "中文文献本地解析，未数据库核验") -> dict:
    authors_list = parsed.get("authors_list") or []
    authors = [{"family": a, "given": ""} for a in authors_list]
    return {
        "confidence": confidence,
        "source": "本地解析",
        "ref_type": parsed.get("ref_type", "中文"),
        "reason": reason,
        "title": parsed.get("title", ""),
        "authors": authors,
        "authors_text": parsed.get("authors_text", ""),
        "year": parsed.get("year", ""),
        "journal": parsed.get("journal", ""),
        "booktitle": parsed.get("booktitle", ""),
        "publisher": parsed.get("publisher", ""),
        "place": parsed.get("place", ""),
        "school": parsed.get("school", ""),
        "doc_type": parsed.get("doc_type", ""),
        "volume": parsed.get("volume", ""),
        "issue": parsed.get("issue", ""),
        "pages": parsed.get("pages", ""),
        "doi": parsed.get("doi", ""),
        "url": parsed.get("url", ""),
        "type": "local-chinese",
        "original": parsed.get("original", ""),
    }

def process_reference(cr, ref: str) -> dict:
    """统一处理一条参考文献：中文无 DOI 走本地解析；英文/有 DOI 走 CrossRef。"""
    ref = clean_text(ref)
    rtype = detect_reference_type(ref)
    doi = extract_doi(ref)

    # 中文无 DOI：不要拿中文题名去 CrossRef 瞎匹配英文结果，直接本地解析并标 ⚠
    if has_chinese(ref) and not doi:
        parsed = parse_chinese_reference(ref)
        if parsed.get("title"):
            return _local_chinese_result(parsed, "uncertain", "中文文献无 DOI：已按 GB/T/中文格式本地解析，未数据库核验")
        return {
            "confidence": "not_found",
            "source": "人工核查",
            "ref_type": rtype,
            "original": ref,
            "title": "",
            "reason": "中文文献无 DOI，且本地字段解析失败；建议补充完整题名、期刊、年份或 DOI"
        }

    # 没有 CrossRef 依赖时：仍允许中文本地解析，英文则提示依赖缺失
    if cr is None:
        if has_chinese(ref):
            parsed = parse_chinese_reference(ref)
            if parsed.get("title"):
                return _local_chinese_result(parsed, "uncertain", "缺少 habanero：仅完成中文本地解析，未联网核验")
        return {"confidence": "not_found", "source": "人工核查", "ref_type": rtype, "original": ref, "reason": "缺少 habanero，无法使用 CrossRef 检索"}

    title = extract_title(ref)
    result = query_crossref(cr, title, ref)
    result.setdefault("ref_type", rtype)
    result.setdefault("source", "CrossRef" if result.get("confidence") != "not_found" else "人工核查")

    # 中文有 DOI：CrossRef 失败时 fallback 到本地解析，避免完全丢失字段
    if has_chinese(ref) and result.get("confidence") == "not_found":
        parsed = parse_chinese_reference(ref)
        if parsed.get("title"):
            local = _local_chinese_result(parsed, "uncertain", "中文 DOI/CrossRef 未可靠命中：已回退为本地解析，需人工核查")
            local["found_title"] = result.get("found_title", "")
            return local
    return result

def format_chinese_apa(ref: dict) -> str:
    """中文本地解析结果的 APA-like 输出：保留中文字段，并明确未核验。"""
    authors = clean_text(ref.get("authors_text") or "、".join(a.get("family", "") for a in ref.get("authors", [])) or "佚名")
    year = clean_text(ref.get("year") or "n.d.")
    title = clean_text(ref.get("title") or ref.get("original", ""))
    dtype = ref.get("doc_type", "")
    apa = f"[本地解析/待核查] {authors} ({year}). {title}."
    if dtype == "J" or ref.get("journal"):
        journal = clean_text(ref.get("journal", ""))
        volume = clean_text(ref.get("volume", ""))
        issue = clean_text(ref.get("issue", ""))
        pages = clean_text(ref.get("pages", ""))
        if journal:
            apa += f" {journal}"
            if volume: apa += f", {volume}"
            if issue: apa += f"({issue})"
            if pages: apa += f", {pages}"
            apa += "."
    elif dtype == "M" or ref.get("publisher"):
        place = clean_text(ref.get("place", ""))
        publisher = clean_text(ref.get("publisher", ""))
        if place or publisher:
            apa += f" {place + ': ' if place else ''}{publisher}."
    elif dtype == "D" or ref.get("school"):
        school = clean_text(ref.get("school", ""))
        if school:
            apa += f" {school}."
    elif dtype == "C" or ref.get("booktitle"):
        booktitle = clean_text(ref.get("booktitle", ""))
        if booktitle:
            apa += f" In {booktitle}."
    doi = clean_text(ref.get("doi", ""))
    url = clean_text(ref.get("url", ""))
    if doi:
        doi = doi.replace("https://doi.org/", "").replace("http://doi.org/", "")
        apa += f" https://doi.org/{doi}"
    elif url:
        apa += f" {url}"
    if ref.get("reason"):
        apa += f" ｜说明：{ref.get('reason')}"
    return apa

def split_references(raw: str) -> list:
    """
    更稳的拆分：
    1）优先按 [1] / 1. / 1) 编号拆；
    2）再按空行拆；
    3）最后按“作者 + 年份”识别新条目。
    这样可以避免把 Smith, J. 这种作者缩写误拆成单独文献。
    """
    raw = raw.replace("\r\n", "\n").replace("\r", "\n").strip()
    raw = re.sub(r"\n{3,}", "\n\n", raw)
    if not raw:
        return []

    # 编号参考文献：一行一个编号开头最可靠
    marker = re.compile(r"(?m)^\s*(?:\[\s*\d+\s*\]|\(?\d+\)?[\.)])\s+")
    ms = list(marker.finditer(raw))
    if len(ms) >= 2:
        refs = []
        for idx, m in enumerate(ms):
            start = m.end()
            end = ms[idx + 1].start() if idx + 1 < len(ms) else len(raw)
            item = raw[start:end].strip()
            item = re.sub(r"\s*\n\s*", " ", item)
            item = clean_text(item)
            if len(item) > 15:
                refs.append(item)
        return refs

    # 空行分隔
    paragraphs = [clean_text(p.replace("\n", " ")) for p in re.split(r"\n\s*\n", raw)]
    paragraphs = [p for p in paragraphs if len(p) > 15]
    if len(paragraphs) >= 2:
        return paragraphs

    # 无空行、无编号：根据作者-年份模式合并换行
    lines = [ln.strip() for ln in raw.split("\n") if ln.strip()]
    refs, current = [], []

    def is_new_ref(line: str) -> bool:
        # 行首像作者，且前 240 字符内有年份，才认为是新文献
        if not re.search(r"\(?\b(?:19|20)\d{2}[a-z]?\b\)?", line[:240]):
            return False
        if re.match(r"^[A-Z\u4e00-\u9fff][^,，]{1,40}[,，]", line):
            return True
        if re.match(r"^[A-Z][A-Za-z'\-]+\s+(?:&|and)\s+[A-Z][A-Za-z'\-]+", line):
            return True
        return False

    for line in lines:
        if current and is_new_ref(line):
            refs.append(clean_text(" ".join(current)))
            current = [line]
        else:
            current.append(line)
    if current:
        refs.append(clean_text(" ".join(current)))
    return [r for r in refs if len(r) > 15]

def extract_title(ref: str) -> str:
    """从不同参考文献格式中尽量提取真正题名，而不是作者名、期刊名或页码。"""
    if has_chinese(ref):
        parsed = parse_chinese_reference(ref)
        if parsed.get("title"):
            return parsed["title"]
    ref0 = clean_text(ref)
    ref0 = re.sub(r"^\s*(?:\[\s*\d+\s*\]|\(?\d+\)?[\.)])\s+", "", ref0)
    ref0 = re.sub(r"https?://doi\.org/\S+", "", ref0, flags=re.I)
    ref0 = re.sub(r"\bdoi\s*:?\s*10\.\S+", "", ref0, flags=re.I)

    m = re.search(r"\b(?:19|20)\d{2}[a-z]?\b\s*[\.。]?\s*[“\"]([^”\"]{4,260})[”\"]", ref0)
    if m:
        return clean_text(m.group(1))

    # APA 样式必须优先于 Vancouver 缩写规则，否则 Driscoll, C. T. 会被误切。
    m = re.search(r"\(\s*(?:19|20)\d{2}[a-z]?\s*\)\s*[\.。,，]?\s+(.+)", ref0)
    if m:
        tail = m.group(1).strip()
        for seg in _smart_sentences(tail):
            toks = _tokens(seg)
            if len(toks) >= 3 and not re.search(r"\b(vol|volume|issue|pp|press|publisher|retrieved|doi|http)\b", seg, re.I):
                return clean_text(seg)
        return clean_text(tail[:220])

    # Vancouver 作者缩写结尾：Bistline JE and Rose SK. Title. Journal...
    # 只在没有 APA 年份括号时使用，并且必须放在智能分句之前。
    if not re.search(r"\(\s*(?:19|20)\d{2}[a-z]?\s*\)", ref0):
        vm = re.match(r"^(.{5,180}?\b[A-Z]{1,4}\.)\s+(.+)$", ref0)
        if vm and _looks_like_author_segment(vm.group(1)):
            tail = vm.group(2).strip()
            for seg in _smart_sentences(tail):
                toks = _tokens(seg)
                if len(toks) >= 3 and not re.search(r"\b(journal|press|university|vol|volume|issue|pp|doi|http)\b", seg, re.I):
                    return clean_text(seg)
            return clean_text(tail[:260])

    segs = _smart_sentences(ref0)
    if len(segs) >= 2 and _looks_like_author_segment(segs[0]):
        second = clean_text(segs[1])
        if len(_tokens(second)) >= 3:
            return second[:260]

    if not re.search(r"\(\s*(?:19|20)\d{2}[a-z]?\s*\)", ref0):
        m = re.match(r"^(.{5,180}?\b[A-Z]{1,4}\.)\s+(.+)$", ref0)
        if m and _looks_like_author_segment(m.group(1)):
            tail = m.group(2).strip()
            for seg in _smart_sentences(tail):
                toks = _tokens(seg)
                if len(toks) >= 3 and not re.search(r"\b(journal|press|university|vol|volume|issue|pp|doi|http)\b", seg, re.I):
                    return clean_text(seg)
            return clean_text(tail[:260])

    m = re.search(r"\b(?:19|20)\d{2}[a-z]?\b[\.,]?\s+(.+)", ref0)
    if m:
        tail = m.group(1).strip()
        for seg in _smart_sentences(tail):
            toks = _tokens(seg)
            if len(toks) >= 3 and not re.search(r"\b(vol|volume|issue|pp|press|publisher|retrieved|doi|http)\b", seg, re.I):
                return clean_text(seg)
        return clean_text(tail[:220])

    m = re.search(r"[\.。]\s*([^\.。]{6,160}?)\s*\[[JMDCP]\]", ref0, flags=re.I)
    if m:
        return clean_text(m.group(1))

    candidates = []
    for seg in segs:
        toks = _tokens(seg)
        if len(toks) < 3:
            continue
        penalty = 0
        if _looks_like_author_segment(seg):
            penalty += 5
        if re.search(r"\b(?:journal|proceedings|press|university|vol|volume|issue|pp|doi|http|lancet|energy policy)\b", seg, re.I):
            penalty += 3
        if re.search(r"\b(?:19|20)\d{2}\b", seg):
            penalty += 1
        score = len(toks) - penalty
        candidates.append((score, seg))
    if candidates:
        candidates.sort(reverse=True, key=lambda x: x[0])
        return clean_text(candidates[0][1])[:260]

    return clean_text(ref0[:260])

def get_year(item: dict) -> str:
    for key in ("published-print", "published-online", "published", "issued", "created"):
        parts = item.get(key, {}).get("date-parts", [[""]])
        if parts and parts[0] and parts[0][0]:
            return str(parts[0][0])
    return ""

def fmt_author_apa(a: dict) -> str:
    family = clean_text(a.get("family", ""))
    given  = clean_text(a.get("given", ""))
    if not family:
        return clean_text(a.get("name", "Unknown"))
    initials = ". ".join(n[0] for n in given.replace("-", " ").split() if n) + "." if given else ""
    return f"{family}, {initials}".strip(", ")

def title_sim(a: str, b: str) -> float:
    a1, b1 = _norm_title(a), _norm_title(b)
    if not a1 or not b1:
        return 0.0
    ta, tb = set(_tokens(a1)), set(_tokens(b1))
    seq = difflib.SequenceMatcher(None, a1, b1).ratio()
    if not ta or not tb:
        return round(seq, 3)
    inter = len(ta & tb)
    jaccard = inter / max(1, len(ta | tb))
    containment = inter / max(1, min(len(ta), len(tb)))
    return round(0.45 * jaccard + 0.35 * containment + 0.20 * seq, 3)

def _crossref_items(resp) -> list:
    msg = resp.get("message", {}) if isinstance(resp, dict) else {}
    if isinstance(msg, dict) and "items" in msg:
        return msg.get("items") or []
    if isinstance(msg, dict) and ("title" in msg or "DOI" in msg):
        return [msg]
    return []

def _candidate_key(item: dict) -> str:
    doi = (item.get("DOI") or "").lower().strip()
    if doi:
        return "doi:" + doi
    title = clean_text((item.get("title") or [""])[0]).lower()
    year = get_year(item)
    return f"title:{title}|{year}"

def _score_candidate(item: dict, title: str, original: str, doi: str = "") -> tuple:
    found = clean_text((item.get("title") or [""])[0])
    if not found:
        return 0.0, "无题名"
    item_doi = (item.get("DOI") or "").lower().strip()
    item_type = clean_text(item.get("type", "")).lower()
    doi_norm = (doi or "").lower().strip()

    if doi_norm and item_doi and doi_norm == item_doi and not _is_supplement_doi(item_doi):
        return 1.0, "DOI 完全匹配"

    ts = title_sim(title, found)
    score = ts
    reasons = [f"题名相似度 {ts:.2f}"]

    if len(_tokens(title)) < 3:
        score -= 0.35
        reasons.append("题名过短/疑似残缺")

    ref_year = _extract_ref_year(original)
    item_year = get_year(item)
    if ref_year and item_year:
        if ref_year == item_year:
            score += 0.10
            reasons.append("年份一致")
        else:
            score -= 0.18
            reasons.append(f"年份不一致({ref_year}/{item_year})")

    ref_authors = _extract_author_names(original)
    item_authors = {clean_text(a.get("family", "")).lower() for a in item.get("author", []) if a.get("family")}
    if ref_authors and item_authors:
        hit = len(ref_authors & item_authors)
        if hit:
            score += min(0.16, 0.06 * hit)
            reasons.append(f"作者姓氏命中{hit}个")
        else:
            score -= 0.25
            reasons.append("作者姓氏未命中")
    elif ref_authors and not item_authors:
        score -= 0.18
        reasons.append("候选缺少作者，不能高置信")

    if item_doi and _is_supplement_doi(item_doi):
        score -= 0.35
        reasons.append("候选是补充材料 DOI")
    if item_type in {"component", "dataset", "posted-content"}:
        score -= 0.20
        reasons.append(f"候选类型为 {item_type}")

    return max(0.0, min(1.0, score)), "；".join(reasons)

def query_crossref(cr, title: str, original: str) -> dict:
    """
    v5 检索策略：
    1）有 DOI 先按 DOI 查；补充材料 DOI 会自动尝试主 DOI；
    2）报告/网页/政策简报类不强行 CrossRef，避免错配；
    3）同时用 query_title、query_bibliographic、宽检索；
    4）候选按题名+年份+作者+候选类型综合打分；
    5）拒绝 Unknown Author 高置信输出、补充材料 DOI、作者不匹配候选。
    """
    original = clean_text(original)
    title = clean_text(title)
    doi = extract_doi(original)

    if _looks_like_non_crossref_work(original):
        return {
            "confidence": "not_found",
            "source": "人工核查",
            "original": original,
            "query_title": title,
            "doi": doi,
            "reason": "疑似报告/网页/政策简报/非 DOI 来源，未强行 CrossRef 匹配"
        }

    try:
        candidates = []
        seen = set()

        def add_items(resp):
            for it in _crossref_items(resp):
                k = _candidate_key(it)
                if k not in seen:
                    candidates.append(it)
                    seen.add(k)

        doi_queries = []
        if doi:
            doi_queries.append(doi)
            main_doi = _main_doi_from_supplement(doi)
            if main_doi and main_doi != doi:
                doi_queries.insert(0, main_doi)
        for dq in doi_queries:
            try:
                add_items(cr.works(ids=dq))
            except Exception:
                pass

        if len(_tokens(title)) >= 4:
            try:
                add_items(cr.works(query_title=title, limit=12))
            except Exception:
                pass

        try:
            add_items(cr.works(query_bibliographic=original, limit=15))
        except Exception:
            pass

        ref_year = _extract_ref_year(original)
        ref_authors_sorted = sorted(_extract_author_names(original))
        first_author = ref_authors_sorted[0] if ref_authors_sorted else ""
        if len(_tokens(title)) >= 3:
            broad_query = " ".join(x for x in [title, ref_year, first_author] if x)
            try:
                add_items(cr.works(query=broad_query, limit=15))
            except Exception:
                pass

        for it in list(candidates):
            item_doi0 = (it.get("DOI") or "").lower().strip()
            main_doi = _main_doi_from_supplement(item_doi0)
            if item_doi0 and main_doi and main_doi != item_doi0:
                try:
                    add_items(cr.works(ids=main_doi))
                except Exception:
                    pass

        if not candidates:
            return {"confidence": "not_found", "original": original, "query_title": title, "doi": doi, "reason": "CrossRef 无候选结果"}

        scored = []
        for item in candidates:
            score, reason = _score_candidate(item, title, original, doi)
            scored.append((score, reason, item))
        scored.sort(key=lambda x: x[0], reverse=True)
        score, reason, item = scored[0]
        found = clean_text((item.get("title") or [""])[0])
        sim = title_sim(title, found)
        item_doi = (item.get("DOI") or "").lower().strip()
        item_authors = {clean_text(a.get("family", "")).lower() for a in item.get("author", []) if a.get("family")}
        ref_authors = _extract_author_names(original)
        exact_doi = bool(doi and item_doi and item_doi == doi.lower().strip())

        if _is_supplement_doi(item_doi) and not exact_doi:
            return {"confidence": "not_found", "original": original, "query_title": title, "doi": doi, "similarity": round(sim, 2), "score": round(score, 2), "found_title": found, "reason": "最佳候选是补充材料 DOI，已拒绝当作主文献：" + reason}

        if ref_authors and not item_authors and not (exact_doi and not _is_supplement_doi(item_doi)):
            return {"confidence": "not_found", "original": original, "query_title": title, "doi": doi, "similarity": round(sim, 2), "score": round(score, 2), "found_title": found, "reason": "候选缺少作者，已拒绝 Unknown Author 输出：" + reason}

        if ref_authors and item_authors and not (ref_authors & item_authors) and not exact_doi:
            return {"confidence": "not_found", "original": original, "query_title": title, "doi": doi, "similarity": round(sim, 2), "score": round(score, 2), "found_title": found, "reason": "作者姓氏不匹配，已拒绝错配：" + reason}

        if exact_doi and not _is_supplement_doi(item_doi):
            confidence = "verified"
        elif score >= 0.82 and sim >= 0.70:
            confidence = "verified"
        elif score >= 0.58 and sim >= 0.45:
            confidence = "uncertain"
        else:
            return {"confidence": "not_found", "original": original, "query_title": title, "doi": doi, "similarity": round(sim, 2), "score": round(score, 2), "found_title": found, "reason": "最佳候选分数过低，已拒绝错配：" + reason}

        authors = item.get("author", [])
        authors_text = ""
        if not authors and ref_authors:
            authors = _extract_author_objects_from_ref(original)
            authors_text = "从原始文献回填作者"

        return {
            "confidence": confidence,
            "similarity": round(sim, 2),
            "score": round(score, 2),
            "reason": reason,
            "query_title": title,
            "title": found,
            "authors": authors,
            "authors_text": authors_text,
            "year": get_year(item),
            "journal": clean_text((item.get("container-title") or [""])[0]),
            "volume": clean_text(item.get("volume", "")),
            "issue": clean_text(item.get("issue", "")),
            "pages": clean_text(item.get("page", "")),
            "doi": clean_text(item.get("DOI", "")),
            "type": item.get("type", "journal-article"),
            "original": original,
        }
    except Exception as e:
        return {"confidence": "not_found", "original": original, "query_title": title, "doi": doi, "error": str(e), "reason": "查询异常"}

def to_apa7(ref: dict) -> str:
    if ref.get("source") == "本地解析" or ref.get("type") == "local-chinese":
        return format_chinese_apa(ref)
    if ref["confidence"] == "not_found":
        reason = ref.get("reason", "")
        if reason:
            return f"[未找到/需人工核查] {ref['original']}  ｜原因：{reason}"
        return f"[未找到/需人工核查] {ref['original']}"
    authors = ref.get("authors", [])
    if not authors:
        astr = "Unknown Author"
    elif len(authors) == 1:
        astr = fmt_author_apa(authors[0])
    elif len(authors) <= 20:
        fmts = [fmt_author_apa(a) for a in authors]
        astr = ", ".join(fmts[:-1]) + ", & " + fmts[-1]
    else:
        fmts = [fmt_author_apa(a) for a in authors]
        astr = ", ".join(fmts[:19]) + ", ... " + fmts[-1]
    year, title = ref.get("year", "n.d."), clean_text(ref.get("title", ""))
    journal, volume = clean_text(ref.get("journal", "")), clean_text(ref.get("volume", ""))
    issue, pages, doi = clean_text(ref.get("issue", "")), clean_text(ref.get("pages", "")), clean_text(ref.get("doi", ""))
    apa = f"{astr} ({year}). {title}."
    if journal:
        apa += f" {journal}"
        if volume: apa += f", {volume}"
        if issue:  apa += f"({issue})"
        if pages:  apa += f", {pages}"
        apa += "."
    if doi:
        doi = doi.replace("https://doi.org/", "").replace("http://doi.org/", "")
        apa += f" https://doi.org/{doi}"
    return apa

def to_bibtex(ref: dict, key: str) -> str:
    if ref.get("source") == "本地解析" or ref.get("type") == "local-chinese":
        etype = {"J": "article", "M": "book", "D": "phdthesis", "C": "inproceedings"}.get(ref.get("doc_type", ""), "misc")
        authors_text = ref.get("authors_text") or " and ".join(a.get("family", "") for a in ref.get("authors", [])) or "Unknown"
        lines = [f"@{etype}{{{key},"]
        fields = [
            ("author", authors_text), ("title", ref.get("title", "")),
            ("journal", ref.get("journal", "")), ("booktitle", ref.get("booktitle", "")),
            ("publisher", ref.get("publisher", "")), ("school", ref.get("school", "")),
            ("year", ref.get("year", "")), ("volume", ref.get("volume", "")),
            ("number", ref.get("issue", "")), ("pages", ref.get("pages", "")),
            ("doi", ref.get("doi", "")), ("url", ref.get("url", "")),
            ("note", "本地解析，未数据库核验"),
        ]
        for k, v in fields:
            v = clean_text(v)
            if v:
                lines.append(f"  {k:<10} = {{{v}}},")
        lines.append("}")
        return "\n".join(lines)
    if ref["confidence"] == "not_found":
        return f"% [未找到/需人工核查] {ref['original']}\n"
    authors = ref.get("authors", [])
    astr = " and ".join(f"{clean_text(a.get('family',''))}, {clean_text(a.get('given',''))}" for a in authors) or "Unknown"
    etype = {"book": "book", "monograph": "book", "proceedings-article": "inproceedings"}.get(
        ref.get("type", ""), "article")
    lines = [f"@{etype}{{{key},"]
    for k, v in [("author", astr), ("title", clean_text(ref.get("title",""))),
                 ("journal", clean_text(ref.get("journal",""))), ("year", clean_text(ref.get("year",""))),
                 ("volume", clean_text(ref.get("volume",""))), ("number", clean_text(ref.get("issue",""))),
                 ("pages", clean_text(ref.get("pages",""))), ("doi", clean_text(ref.get("doi","")))]:
        if v:
            lines.append(f"  {k:<10} = {{{v}}},")
    lines.append("}")
    return "\n".join(lines)

def make_key(ref: dict, i: int) -> str:
    authors = ref.get("authors", [])
    family_raw = authors[0].get("family", "Unknown") if authors else (ref.get("authors_text") or "Unknown")
    # 中文作者转成可用的 ASCII key；保底用 Ref
    family = re.sub(r'[^a-zA-Z0-9]', '', family_raw)
    if not family:
        family = "Ref"
    return f"{family}{ref.get('year','0000')}_{i}"

# ══════════════════════════════════════════════════════════════
# 漫画风格自定义组件
# ══════════════════════════════════════════════════════════════

class InkFrame(tk.Frame):
    """带手绘风格粗边框的 Frame，用 Canvas 画不规则边框模拟漫画格"""
    def __init__(self, parent, thickness=2, **kw):
        super().__init__(parent, bg=PAPER, bd=0, **kw)
        self._thickness = thickness

class InkButton(tk.Canvas):
    """手绘风格按钮：黑色填充+白字 或 纸色+黑框"""
    def __init__(self, parent, text, command, filled=True, width_px=120, height_px=34, **kw):
        super().__init__(parent, width=width_px, height=height_px,
                         bg=PAPER, highlightthickness=0, cursor="hand2")
        self._text    = text
        self._command = command
        self._filled  = filled
        self._width_px = width_px
        self._height_px = height_px
        self._enabled = True
        self._draw(hover=False)
        self.bind("<Enter>",    lambda e: self._draw(hover=True))
        self.bind("<Leave>",    lambda e: self._draw(hover=False))
        self.bind("<Button-1>", lambda e: self._on_click())

    def _draw(self, hover=False):
        self.delete("all")
        w, h = self._width_px, self._height_px
        pad = 3
        # 手绘感：稍微不规则的圆角矩形（用多边形近似）
        if self._filled:
            fill   = INK_SOFT if hover else INK
            outline= INK
            fg     = PAPER
        else:
            fill   = PAPER_DARK if hover else PAPER
            outline= INK
            fg     = INK

        if not self._enabled:
            fill    = PAPER_DARK
            outline = PENCIL_LT
            fg      = PENCIL_LT

        # 画略微抖动的矩形边框（漫画手绘感）
        pts = [
            pad+2, pad,
            w-pad-1, pad+1,
            w-pad, pad+2,
            w-pad+1, h-pad-2,
            w-pad-2, h-pad,
            pad+1, h-pad+1,
            pad, h-pad-1,
            pad-1, pad+3,
        ]
        self.create_polygon(pts, fill=fill, outline=outline, width=2, smooth=False)
        self.create_text(w//2, h//2, text=self._text, fill=fg,
                         font=F_BTN, anchor="center")

    def _on_click(self):
        if self._enabled and self._command:
            self._command()

    def set_enabled(self, val: bool):
        self._enabled = val
        self._draw(hover=False)

    def set_text(self, text: str):
        self._text = text
        self._draw(hover=False)


class MangaProgressBar(tk.Canvas):
    """漫画风格进度条：黑色填充矩形，带手绘边框"""
    def __init__(self, parent, width=260, height=14, **kw):
        super().__init__(parent, width=width, height=height,
                         bg=PAPER, highlightthickness=0)
        self._width_px = width
        self._height_px = height
        self._val = 0.0
        self._draw()

    def set(self, ratio: float):
        self._val = max(0.0, min(1.0, ratio))
        self._draw()

    def _draw(self):
        self.delete("all")
        w, h = self._width_px, self._height_px
        # 外框（手绘感，略微不规则）
        self.create_rectangle(1, 1, w-1, h-1, outline=INK, width=2, fill=PAPER)
        # 填充进度
        if self._val > 0:
            fill_w = max(0, int((w - 6) * self._val))
            self.create_rectangle(3, 3, 3 + fill_w, h-3, fill=INK, outline="")
        # 百分比文字
        pct = int(self._val * 100)
        self.create_text(w//2, h//2, text=f"{pct}%",
                         fill=PAPER if self._val > 0.5 else INK,
                         font=("Consolas", 7))


class LogoCanvas(tk.Canvas):
    """在标题栏左侧画一个简笔漫画 logo（用 tkinter 原生绘图）"""
    def __init__(self, parent, size=48, **kw):
        super().__init__(parent, width=size, height=size,
                         bg=INK, highlightthickness=0)
        self._s = size
        self._draw()

    def _draw(self):
        s = self._s
        c = self
        # 黑色背景已经是 canvas bg
        # 画简笔眼镜男：圆脸 + 眼镜 + 猫耳
        # 脸
        c.create_oval(8, 10, s-8, s-6, fill=PAPER, outline=PAPER, width=0)
        # 眼镜框（两个小矩形）
        c.create_rectangle(11, 22, 22, 29, outline=PAPER, width=2, fill="")
        c.create_rectangle(26, 22, 37, 29, outline=PAPER, width=2, fill="")
        c.create_line(22, 25, 26, 25, fill=PAPER, width=1)
        # 鼻子
        c.create_oval(22, 30, 26, 33, fill=PAPER, outline="")
        # 嘴（小弧线=微笑）
        c.create_arc(18, 32, 30, 40, start=200, extent=140, style="arc",
                     outline=PAPER, width=2)
        # 猫耳（右上角）
        c.create_polygon(36, 10, 42, 2, 46, 12, fill=PAPER, outline="")
        # 问号气泡（左上）
        c.create_text(6, 8, text="?", fill=PAPER,
                      font=("Arial", 9, "bold"), anchor="nw")


class AvatarLogo(tk.Label):
    """标题栏头像 Logo。优先使用用户头像；失败时由调用处回退到旧手绘 Logo。"""
    def __init__(self, parent, image_obj, **kw):
        super().__init__(parent, image=image_obj, bg=INK, bd=0, highlightthickness=0, **kw)
        self.image = image_obj


# ══════════════════════════════════════════════════════════════
# 主界面
# ══════════════════════════════════════════════════════════════

class XiaJiansuoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("瞎检索")
        self.geometry("960x720")
        self.minsize(800, 580)
        self.configure(bg=PAPER)
        self.results = []
        self._anim_dots = 0
        self._anim_job  = None
        self._init_app_icon()
        self._build_ui()
        self._check_deps()

    def _init_app_icon(self):
        """设置窗口左上角/任务栏图标；打包后仍然有效。"""
        self._app_icon_img = None
        try:
            self._app_icon_img = tk.PhotoImage(data=APP_ICON_PNG_BASE64)
            self.iconphoto(True, self._app_icon_img)
        except Exception:
            self._app_icon_img = None

    # ── 构建 UI ──────────────────────────────────────────────
    def _build_ui(self):
        self._build_header()
        self._build_divider()
        self._build_body()
        self._build_run_bar()
        self._build_log()

    def _build_header(self):
        hdr = tk.Frame(self, bg=INK, pady=0)
        hdr.pack(fill="x")

        # 头像 Logo：替换原来的圆圈 + 问号
        if self._app_icon_img:
            AvatarLogo(hdr, self._app_icon_img).pack(side="left", padx=(10, 0), pady=6)
        else:
            LogoCanvas(hdr, size=52).pack(side="left", padx=(10, 0), pady=6)

        # 标题文字
        title_block = tk.Frame(hdr, bg=INK)
        title_block.pack(side="left", padx=10, pady=8)
        tk.Label(title_block, text="瞎检索", font=F_TITLE,
                 bg=INK, fg=PAPER).pack(anchor="w")
        tk.Label(title_block, text="CrossRef · 强化匹配 · APA 7 · BibTeX",
                 font=F_SUB, bg=INK, fg=PENCIL).pack(anchor="w")

        # 右侧状态气泡
        self.lbl_status = tk.Label(hdr, text="就绪", font=F_SUB,
                                   bg=INK, fg=PENCIL_LT)
        self.lbl_status.pack(side="right", padx=18)

    def _build_divider(self):
        # 模拟漫画格之间的粗线
        tk.Frame(self, bg=INK, height=3).pack(fill="x")

    def _build_body(self):
        body = tk.Frame(self, bg=PAPER)
        body.pack(fill="both", expand=True, padx=0, pady=0)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=0)   # 分割线
        body.columnconfigure(2, weight=1)
        body.rowconfigure(0, weight=1)

        # ── 左栏：输入 ──
        left = tk.Frame(body, bg=PAPER)
        left.grid(row=0, column=0, sticky="nsew", padx=12, pady=10)
        left.rowconfigure(1, weight=1)

        self._section_label(left, "① 粘贴 / 导入文献").grid(row=0, column=0, sticky="w", pady=(0,6))

        self.txt_input = scrolledtext.ScrolledText(
            left, font=F_MONO, wrap="word",
            bg=PAPER, fg=INK, insertbackground=INK,
            relief="flat", bd=0,
            highlightthickness=2, highlightbackground=INK, highlightcolor=INK,
            selectbackground=INK, selectforeground=PAPER)
        self.txt_input.grid(row=1, column=0, sticky="nsew")
        self._set_placeholder()
        self.txt_input.bind("<FocusIn>", self._clear_placeholder)

        btn_row = tk.Frame(left, bg=PAPER, pady=6)
        btn_row.grid(row=2, column=0, sticky="w")
        InkButton(btn_row, "导入文件", self._import_file,
                  filled=True, width_px=100).pack(side="left")
        InkButton(btn_row, "清空", self._clear_input,
                  filled=False, width_px=64).pack(side="left", padx=(8,0))

        # ── 分割线（漫画格间黑线）──
        tk.Frame(body, bg=INK, width=3).grid(row=0, column=1, sticky="ns")

        # ── 右栏：结果 ──
        right = tk.Frame(body, bg=PAPER)
        right.grid(row=0, column=2, sticky="nsew", padx=12, pady=10)
        right.rowconfigure(1, weight=1)

        self._section_label(right, "③ APA 7 结果").grid(row=0, column=0, sticky="w", pady=(0,6))

        self.txt_output = scrolledtext.ScrolledText(
            right, font=F_MONO, wrap="word",
            bg=PAPER_DARK, fg=INK,
            relief="flat", bd=0,
            highlightthickness=2, highlightbackground=INK, highlightcolor=INK,
            selectbackground=INK, selectforeground=PAPER,
            state="disabled")
        self.txt_output.grid(row=1, column=0, sticky="nsew")
        # 颜色 tag
        self.txt_output.tag_config("verified",  foreground=INK)
        self.txt_output.tag_config("uncertain", foreground=PENCIL)
        self.txt_output.tag_config("not_found", foreground=PENCIL_LT)
        self.txt_output.tag_config("num",       foreground=PENCIL, font=F_MONO)

        out_btns = tk.Frame(right, bg=PAPER, pady=6)
        out_btns.grid(row=2, column=0, sticky="w")
        InkButton(out_btns, "保存 APA",    self._save_apa,  width_px=90).pack(side="left")
        InkButton(out_btns, "保存 BibTeX", self._save_bib,  width_px=100).pack(side="left", padx=(6,0))
        InkButton(out_btns, "复制 APA",    self._copy_apa,  filled=False, width_px=90).pack(side="left", padx=(6,0))
        InkButton(out_btns, "导出报告",    self._save_csv,  filled=False, width_px=90).pack(side="left", padx=(6,0))

    def _build_run_bar(self):
        tk.Frame(self, bg=INK, height=3).pack(fill="x")
        bar = tk.Frame(self, bg=PAPER_DARK, pady=8)
        bar.pack(fill="x", padx=0)

        self.btn_run = InkButton(bar, "▶  开始核查", self._run,
                                  filled=True, width_px=140, height_px=38)
        self.btn_run.pack(side="left", padx=14)

        self.progress_bar = MangaProgressBar(bar, width=240, height=16)
        self.progress_bar.pack(side="left", padx=8)

        self.lbl_pct = tk.Label(bar, text="", font=F_MONO, bg=PAPER_DARK, fg=PENCIL)
        self.lbl_pct.pack(side="left", padx=4)

        # 统计 badge（右侧）
        self.lbl_stats = tk.Label(bar, text="", font=F_BIG, bg=PAPER_DARK, fg=INK)
        self.lbl_stats.pack(side="right", padx=18)

    def _build_log(self):
        tk.Frame(self, bg=INK, height=2).pack(fill="x")
        log_frame = tk.Frame(self, bg=INK, pady=0)
        log_frame.pack(fill="x")
        self.txt_log = scrolledtext.ScrolledText(
            log_frame, height=5, font=F_LOG,
            bg=INK, fg=PAPER_DARK,
            relief="flat", bd=0,
            highlightthickness=0,
            state="disabled")
        self.txt_log.pack(fill="x", padx=0)
        self.txt_log.tag_config("ok",    foreground="#DDDDDD")
        self.txt_log.tag_config("warn",  foreground="#AAAAAA")
        self.txt_log.tag_config("error", foreground="#777777")
        self.txt_log.tag_config("head",  foreground=PAPER, font=F_LOG)

    # ── 辅助组件 ─────────────────────────────────────────────

    def _section_label(self, parent, text: str) -> tk.Label:
        """章节标题：带下划线横线，漫画对话框感觉"""
        f = tk.Frame(parent, bg=PAPER)
        tk.Label(f, text=text, font=F_H2, bg=PAPER, fg=INK).pack(side="left")
        tk.Frame(f, bg=INK, height=2, width=60).pack(side="left", padx=6, pady=0, anchor="s")
        return f

    def _set_placeholder(self):
        ph = ("在此粘贴文献列表（任意格式均可）\n\n"
              "支持：\n"
              "  [1] Author... 编号列表\n"
              "  空行分隔的段落\n"
              "  APA / MLA / Vancouver / GB/T 7714 / 混乱格式\n"
              "  中文文献无 DOI 时会本地解析并标记待核查\n\n"
              "或点击「导入文件」读取 .docx / .txt")
        self.txt_input.insert("1.0", ph)
        self.txt_input.config(fg=PENCIL_LT)

    def _clear_placeholder(self, _=None):
        content = self.txt_input.get("1.0", "end-1c")
        if content.startswith("在此粘贴"):
            self.txt_input.delete("1.0", "end")
            self.txt_input.config(fg=INK)

    def _clear_input(self):
        self.txt_input.delete("1.0", "end")
        self._set_placeholder()

    # ── 日志 / 状态 ───────────────────────────────────────────

    def _log(self, msg: str, tag: str = ""):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", msg + "\n", tag or "")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def _set_status(self, text: str):
        self.lbl_status.config(text=text)

    def _set_output_colored(self, results):
        """按核查状态给每条结果上色"""
        self.txt_output.config(state="normal")
        self.txt_output.delete("1.0", "end")
        for i, r in enumerate(results, 1):
            conf = r.get("confidence", "not_found")
            prefix = {"verified": "✓ ", "uncertain": "⚠ ", "not_found": "✗ "}.get(conf, "")
            num_text = f"{i}. "
            self.txt_output.insert("end", num_text, "num")
            self.txt_output.insert("end", prefix + to_apa7(r) + "\n\n", conf)
        self.txt_output.config(state="disabled")

    # ── 动画 ──────────────────────────────────────────────────

    def _start_anim(self):
        self._anim_dots = 0
        self._tick_anim()

    def _tick_anim(self):
        dots = "・" * (self._anim_dots % 4)
        self.btn_run.set_text(f"检索中{dots}")
        self._anim_dots += 1
        self._anim_job = self.after(400, self._tick_anim)

    def _stop_anim(self):
        if self._anim_job:
            self.after_cancel(self._anim_job)
            self._anim_job = None
        self.btn_run.set_text("▶  开始核查")

    # ── 依赖检查 ──────────────────────────────────────────────

    def _check_deps(self):
        missing = []
        if not HAS_CROSSREF: missing.append("habanero")
        if not HAS_DOCX:     missing.append("python-docx")
        if missing:
            self._log(f"[ 提示 ] 缺少依赖，请运行: pip install {' '.join(missing)}", "warn")

    # ── 导入文件 ──────────────────────────────────────────────

    def _import_file(self):
        types = [("支持的文件", "*.docx *.txt *.md"), ("所有文件", "*.*")]
        path = filedialog.askopenfilename(filetypes=types)
        if not path:
            return
        p = Path(path)
        try:
            if p.suffix.lower() == ".docx":
                if not HAS_DOCX:
                    messagebox.showerror("缺少依赖", "请先安装: pip install python-docx")
                    return
                doc  = DocxDocument(str(p))
                text = "\n".join(para.text for para in doc.paragraphs if para.text.strip())
            else:
                text = p.read_text(encoding="utf-8", errors="replace")
            self.txt_input.delete("1.0", "end")
            self.txt_input.config(fg=INK)
            self.txt_input.insert("1.0", text)
            self._log(f"[ 导入 ] {p.name}  ({len(text)} 字符)", "ok")
        except Exception as e:
            messagebox.showerror("导入失败", str(e))

    # ── 核查主流程 ────────────────────────────────────────────

    def _run(self):
        if not HAS_CROSSREF:
            messagebox.showwarning("缺少依赖", "未安装 habanero：英文/CrossRef 检索不可用，但中文本地解析仍可运行。")
        raw = self.txt_input.get("1.0", "end-1c").strip()
        if not raw or raw.startswith("在此粘贴"):
            messagebox.showwarning("瞎检索", "还没有文献呢，先粘贴或者导入吧")
            return
        self.results = []
        self.btn_run.set_enabled(False)
        self._start_anim()
        self.progress_bar.set(0)
        self.lbl_stats.config(text="")
        threading.Thread(target=self._worker, args=(raw,), daemon=True).start()

    def _worker(self, raw: str):
        cr   = Crossref(mailto="xiajiannsuo@tool.local") if HAS_CROSSREF else None
        refs = split_references(raw)
        total = len(refs)
        self.after(0, lambda: self._set_status(f"检测到 {total} 条，查询中..."))
        self._log(f"[ 开始 ] 共 {total} 条文献", "head")

        results = []
        for i, ref in enumerate(refs, 1):
            title = extract_title(ref)
            self._log(f"[{i:2d}/{total}] {title[:52]}...")
            result = process_reference(cr, ref)
            results.append(result)

            conf = result.get("confidence", "not_found")
            icon = {"verified": "✓", "uncertain": "⚠", "not_found": "✗"}.get(conf, "?")
            tag  = {"verified": "ok", "uncertain": "warn", "not_found": "error"}.get(conf, "")
            msg  = result.get("title") or result.get("found_title") or result.get("error", "")
            extra = ""
            if result.get("score") is not None:
                extra = f"  score={result.get('score')}"
            self._log(f"        {icon} {msg[:56]}{extra}", tag)

            ratio = i / total
            self.after(0, lambda r=ratio, _i=i, _t=total: (
                self.progress_bar.set(r),
                self.lbl_pct.config(text=f"{_i}/{_t}")
            ))
            if i < total:
                time.sleep(0.4)

        self.results = results
        self.after(0, self._on_done)

    def _on_done(self):
        results   = self.results
        verified  = sum(1 for r in results if r["confidence"] == "verified")
        uncertain = sum(1 for r in results if r["confidence"] == "uncertain")
        not_found = sum(1 for r in results if r["confidence"] == "not_found")

        self._set_output_colored(results)
        self._stop_anim()
        self.btn_run.set_enabled(True)
        self.progress_bar.set(1.0)
        self._set_status("核查完毕")
        self.lbl_stats.config(text=f"✓{verified}  ⚠{uncertain}  ✗{not_found}")
        self._log(f"[ 完成 ] ✓{verified}  ⚠{uncertain}  ✗{not_found}", "head")

    # ── 导出 ──────────────────────────────────────────────────

    def _save_apa(self):
        if not self.results:
            messagebox.showwarning("瞎检索", "请先运行核查"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".txt", filetypes=[("文本文件", "*.txt")],
            initialfile="references_apa.txt")
        if not path: return
        lines = [f"{i}. {to_apa7(r)}" for i, r in enumerate(self.results, 1)]
        Path(path).write_text("\n\n".join(lines), encoding="utf-8")
        messagebox.showinfo("瞎检索", f"APA 列表已保存！\n{path}")

    def _save_bib(self):
        if not self.results:
            messagebox.showwarning("瞎检索", "请先运行核查"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".bib", filetypes=[("BibTeX 文件", "*.bib")],
            initialfile="references.bib")
        if not path: return
        entries = [to_bibtex(r, make_key(r, i)) for i, r in enumerate(self.results, 1)]
        Path(path).write_text("\n\n".join(entries), encoding="utf-8")
        messagebox.showinfo("瞎检索", f"BibTeX 已保存，可直接拖入 Zotero！\n{path}")

    def _copy_apa(self):
        if not self.results:
            messagebox.showwarning("瞎检索", "请先运行核查"); return
        lines = [f"{i}. {to_apa7(r)}" for i, r in enumerate(self.results, 1)]
        self.clipboard_clear()
        self.clipboard_append("\n\n".join(lines))
        self._set_status("已复制到剪贴板 ✓")

    def _save_csv(self):
        if not self.results:
            messagebox.showwarning("瞎检索", "请先运行核查"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV 文件", "*.csv")],
            initialfile="xia_jiansuo_report.csv")
        if not path: return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["状态","文献类型","数据来源","标题","年份","作者","期刊/来源","DOI","相似度","原因","原始文本"])
            w.writeheader()
            for r in self.results:
                w.writerow({
                    "状态":   {"verified":"✓","uncertain":"⚠","not_found":"✗"}.get(r.get("confidence"),""),
                    "文献类型": r.get("ref_type", ""),
                    "数据来源": r.get("source", ""),
                    "标题":   r.get("title",""),
                    "年份":   r.get("year",""),
                    "作者":   r.get("authors_text") or "; ".join(a.get("family","") for a in r.get("authors",[])),
                    "期刊/来源":   r.get("journal") or r.get("booktitle") or r.get("publisher") or r.get("school") or "",
                    "DOI":    r.get("doi",""),
                    "相似度": f"{r.get('similarity',0):.0%}" if r.get("similarity") else "",
                    "原因": r.get("reason", ""),
                    "原始文本": r.get("original","")[:160],
                })
        messagebox.showinfo("瞎检索", f"报告已保存！\n{path}")


# ══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = XiaJiansuoApp()
    app.mainloop()
