# -*- coding: utf-8 -*-
"""
Katta .xlsx fayllarni oqim (streaming) usulida o'qish.

openpyxl/pandas bilan 100 MB+ fayllar juda sekin va ko'p xotira oladi, shuning uchun
zip ichidagi sheet XML ni bo'laklab o'qiymiz. Tashqi kutubxona kerak emas.
"""
import codecs
import html
import re
import zipfile
import xml.etree.ElementTree as ET

_ROW_RE = re.compile(r"<row\b([^>]*?)(?:/>|>(.*?)</row>)", re.S)
_CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
_REF_RE = re.compile(r'\br="([A-Z]+)(\d+)"')
_TYPE_RE = re.compile(r'\bt="([^"]+)"')
_V_RE = re.compile(r"<v>(.*?)</v>", re.S)
_T_RE = re.compile(r"<t\b[^>]*>(.*?)</t>", re.S)

_NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_CHUNK = 8 * 1024 * 1024


def col_to_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _first_sheet_path(z):
    try:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheet = wb.find("%ssheets/%ssheet" % (_NS_MAIN, _NS_MAIN))
        rid = sheet.get(_NS_REL + "id")
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        for rel in rels:
            if rel.get("Id") == rid:
                target = rel.get("Target").lstrip("/")
                return target if target.startswith("xl/") else "xl/" + target
    except Exception:
        pass
    return "xl/worksheets/sheet1.xml"


def _load_shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    result = []
    with z.open("xl/sharedStrings.xml") as f:
        for _event, elem in ET.iterparse(f, events=("end",)):
            if elem.tag == _NS_MAIN + "si":
                result.append("".join(t.text or "" for t in elem.iter(_NS_MAIN + "t")))
                elem.clear()
    return result


def _cell_value(attrs, inner, shared):
    if not inner:
        return None
    m = _TYPE_RE.search(attrs)
    ctype = m.group(1) if m else "n"
    if ctype == "inlineStr":
        parts = _T_RE.findall(inner)
        return html.unescape("".join(parts)) if parts else None
    v = _V_RE.search(inner)
    if not v:
        return None
    raw = v.group(1)
    if ctype == "s":
        try:
            return shared[int(raw)]
        except (ValueError, IndexError):
            return None
    if ctype in ("str", "e"):
        return html.unescape(raw)
    if ctype == "b":
        return raw == "1"
    try:
        return float(raw)
    except ValueError:
        return html.unescape(raw)


def iter_rows(path, wanted=None):
    """
    (qator_raqami, {ustun_harfi: qiymat}) juftliklarini ketma-ket qaytaradi.
    `wanted` - ustun harflari to'plami (o'zgaruvchan): bo'sh bo'lsa hamma ustun o'qiladi,
    to'ldirilsa faqat shu ustunlar o'qiladi (tezroq).
    """
    try:
        zf = zipfile.ZipFile(str(path))
    except zipfile.BadZipFile:
        raise ValueError("%s: fayl buzilgan yoki to'liq yuklanmagan (haqiqiy .xlsx emas). "
                         "CAS dan qaytadan yuklab oling" % getattr(path, "name", path))
    with zf as z:
        shared = _load_shared_strings(z)
        sheet_path = _first_sheet_path(z)
        decoder = codecs.getincrementaldecoder("utf-8")()
        buf = ""
        with z.open(sheet_path) as f:
            while True:
                data = f.read(_CHUNK)
                if not data:
                    buf += decoder.decode(b"", final=True)
                    portion, buf = buf, ""
                else:
                    buf += decoder.decode(data)
                    cut = buf.rfind("</row>")
                    if cut < 0:
                        continue
                    cut += len("</row>")
                    portion, buf = buf[:cut], buf[cut:]
                for rm in _ROW_RE.finditer(portion):
                    rnum = re.search(r'\br="(\d+)"', rm.group(1))
                    cells = {}
                    for cm in _CELL_RE.finditer(rm.group(2) or ""):
                        ref = _REF_RE.search(cm.group(1))
                        if not ref:
                            continue
                        if wanted and ref.group(1) not in wanted:
                            continue
                        val = _cell_value(cm.group(1), cm.group(2), shared)
                        if val is not None:
                            cells[ref.group(1)] = val
                    yield (int(rnum.group(1)) if rnum else 0), cells
                if not data:
                    break


def clean_text(v):
    return v.strip() if isinstance(v, str) else v


def meter_key(v):
    """Hisoblagich raqamini solishtirish uchun yagona matn ko'rinishiga keltiradi."""
    if v is None:
        return None
    if isinstance(v, float):
        v = str(int(v)) if v.is_integer() else repr(v)
    v = str(v).strip()
    return v or None


def read_header(rows, required, preface=None):
    """
    `rows` generatoridan sarlavha qatorini topadi (required nomlarning hammasi bor qator).
    Qaytaradi: [(nom, ustun_harfi), ...] ustun tartibida. Generator sarlavhadan keyin qoladi.
    """
    need = set(required)
    for rnum, cells in rows:
        if preface is not None:
            preface.append((rnum, cells))
        names = {clean_text(v): c for c, v in cells.items() if isinstance(v, str)}
        if need.issubset(names):
            return sorted(((n, c) for n, c in names.items() if n), key=lambda x: col_to_index(x[1]))
        if rnum > 30:
            break
    raise ValueError("Sarlavha qatori topilmadi (kerakli ustunlar: %s)" % ", ".join(required))
