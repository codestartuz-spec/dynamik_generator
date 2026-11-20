# -*- coding: utf-8 -*-
"""
Hisobot jadvallarini TILGA BOG'LIQ BO'LMAGAN holda o'qish.

CAS platformasi tili (ru / en / uz ...) o'zgarsa, ustun sarlavhalari ham o'zgaradi.
Shuning uchun ustunlar nomi bo'yicha EMAS, quyidagi tartibda aniqlanadi:

 1) sarlavha matni - ko'p tilli sinonimlar ro'yxati (kichik/katta harf, apostrof, bo'sh joy, tinish
    belgilari e'tiborga olinmaydi);
 2) sinonim topilmasa - ustunning joylashuvi (CAS hisobotlarida ustun tartibi tilga bog'liq emas);
 3) joylashuv bo'yicha topilgan ustunning ichidagi ma'lumot tekshiriladi (hisoblagich raqami, sana,
    son...) - noto'g'ri ustun olinib qolmasligi uchun. Mos kelmasa dastur aniq xato bilan to'xtaydi.
"""
import itertools
import re
from datetime import datetime, timedelta

from xlsx_reader import iter_rows, col_to_index, clean_text, meter_key


# ----------------------------------------------------------------- yordamchi --
def norm(v):
    """Solishtirish uchun: kichik harf, ё=е, apostroflar va tinish belgilarisiz."""
    if v is None:
        return ""
    s = str(v).casefold().replace("ё", "е")
    s = re.sub(r"[`'‘’ʻʼ´′‛ʹʿ]", "", s)
    return re.sub(r"[\W_]+", "", s)


def index_to_col(n):
    """1 -> A, 27 -> AA"""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


_EXCEL_EPOCH = datetime(1899, 12, 30)
_TEXT_FORMATS = ("%d/%m/%Y %H:%M:%S", "%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M",
                 "%Y/%m/%d %H:%M:%S", "%d-%m-%Y %H:%M:%S", "%d.%m.%Y %H:%M")


def to_datetime(v):
    """Excel sana raqami yoki matnni datetime ga aylantiradi; bo'lmasa None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if v <= 0:
            return None
        return _EXCEL_EPOCH + timedelta(days=float(v))
    s = str(v).strip()
    for fmt in _TEXT_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def _is_number(v):
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        try:
            float(v.strip().replace(",", ".").rstrip("%"))
            return True
        except ValueError:
            return False
    return False


_VALIDATORS = {
    "id": lambda v: meter_key(v) is not None and len(meter_key(v)) >= 3,
    "datetime": lambda v: to_datetime(v) is not None,
    "number": _is_number,
    "text": lambda v: isinstance(v, str) and v.strip() != "",
}


# ---------------------------------------------------------------- tavsif ------
class Field(object):
    """key - ichki nom; name - chiqish fayldagi (kanonik) ustun nomi; pos - ustunning 0 dan boshlangan joyi."""

    def __init__(self, key, name, pos, aliases, kind=None, required=False, drop=False):
        self.key, self.name, self.pos, self.kind = key, name, pos, kind
        self.required, self.drop = required, drop
        self.aliases = set(norm(a) for a in ([name] + list(aliases)))


class Spec(object):
    def __init__(self, label, fields, min_cells):
        self.label, self.fields, self.min_cells = label, fields, min_cells
        self.by_key = dict((f.key, f) for f in fields)


class Table(object):
    """open_table natijasi."""
    pass


def _clean_strings(cells):
    return dict((c, clean_text(v)) for c, v in cells.items() if isinstance(v, str) and v.strip())


def _match_header(cells, spec):
    """
    Shu qator sarlavhami? Qaytaradi (header, cols, how) yoki None.
    header: [(nom, harf)] ustun tartibida; cols: {key: harf}; how: {key: 'nom'|'joy'}
    """
    strings = _clean_strings(cells)
    if not strings:
        return None
    ordered = sorted(strings, key=col_to_index)
    cols, how = {}, {}
    for f in spec.fields:
        for c in ordered:
            if norm(strings[c]) in f.aliases:
                cols[f.key], how[f.key] = c, "nom"
                break
    by_name_ok = (all(f.key in cols for f in spec.fields if f.required)
                  and len(cols) >= min(2, len(spec.fields)))
    # birlashtirilgan sarlavha (bir xil matn ko'p katakda) emasligini tekshirish: asosan har xil matnlar
    distinct = len(set(s.casefold() for s in strings.values())) >= 0.8 * len(strings)
    structural = len(strings) >= spec.min_cells and distinct
    if not (by_name_ok or structural):
        return None
    # nomi bo'yicha topilmaganlar - joylashuv bo'yicha (faqat joylashuv mos bo'lsa)
    consistent = all(col_to_index(cols[k]) - 1 == spec.by_key[k].pos for k in cols)
    for f in spec.fields:
        if f.key in cols:
            continue
        if consistent:
            cols[f.key], how[f.key] = index_to_col(f.pos + 1), "joy"
        elif f.required:
            return ("xato", "ustun joylashuvi kutilganidan farq qiladi va '%s' nomi bo'yicha topilmadi" % f.name)
    missing = [f.name for f in spec.fields if f.required and f.key not in cols]
    if missing:
        return None
    header = sorted(((clean_text(v), c) for c, v in strings.items()), key=lambda x: col_to_index(x[1]))
    return header, cols, how


def open_table(path, spec, wanted=None, sample_rows=60):
    """
    Hisobot faylini ochadi, sarlavha qatorini va kerakli ustunlarni topadi.
    Qaytaradi Table: preface, header, cols{key:harf}, how{key:'nom'|'joy'}, rows (generator), note, letter_field.
    """
    rows = iter_rows(path, wanted)
    preface = []
    found = None
    for rnum, cells in rows:
        preface.append((rnum, cells))
        res = _match_header(cells, spec)
        if res is not None:
            if res[0] == "xato":
                raise ValueError("%s: %s hisobotiga o'xshamaydi: %s" % (path.name, spec.label, res[1]))
            found = res
            break
        if rnum > 30:
            break
    if not found:
        names = ", ".join(f.name for f in spec.fields if f.required)
        raise ValueError("%s: %s hisoboti sarlavhasi topilmadi (kerakli ustunlar: %s). Boshqa hisobot yoki "
                         "ustun tartibi o'zgargan bo'lishi mumkin." % (path.name, spec.label, names))
    header, cols, how = found

    # joylashuv bo'yicha topilgan ustunlar ichidagi ma'lumotni tekshirish
    sample = []
    for item in rows:
        sample.append(item)
        if len(sample) >= sample_rows:
            break
    for key, h in how.items():
        f = spec.by_key[key]
        if h != "joy" or not f.kind:
            continue
        vals = [cells.get(cols[key]) for _r, cells in sample if cols[key] in cells]
        if not vals:
            continue
        ok = sum(1 for v in vals if _VALIDATORS[f.kind](v))
        if ok * 2 < len(vals):
            raise ValueError("%s: '%s' ustuni (%s ustun) nomi bo'yicha topilmadi, joylashuv bo'yicha olingan "
                             "ustundagi ma'lumot mos emas. Hisobot ustunlari tartibi o'zgargan bo'lishi mumkin."
                             % (path.name, f.name, cols[key]))

    t = Table()
    t.preface, t.header, t.cols, t.how = preface, header, cols, how
    t.rows = itertools.chain(sample, rows)
    t.letter_field = dict((c, spec.by_key[k]) for k, c in cols.items())
    pos_keys = [spec.by_key[k].name for k, h in how.items() if h == "joy"]
    t.note = ("ustunlar nomi mos kelmadi, joylashuv bo'yicha aniqlandi: %s" % ", ".join(pos_keys)) if pos_keys else None
    return t
