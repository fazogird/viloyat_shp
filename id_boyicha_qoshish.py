"""Asl shp ga boshqa (yordamchi) fayldan unique_id bo'yicha kerakli ustunlarni qo'shish / to'ldirish.

Odatda atribut_otkazish.py va qoshnidan_toldirish.py dan keyin ishlatiladi: masalan suv hajmi
ustunlari (uw3_m3 ... uw10_m3, uwt_m3) yoki fermer ma'lumoti boshqa faylda turadi va ularni ID
orqali olib kelish kerak.

Qoida:
  1. ASL — ma'lumot qabul qiladigan shp (o'zgarmaydi). YORDAMCHI — ustunlar olinadigan fayl
     (shp, gpkg, dbf, csv yoki xlsx; faqat o'qiladi).
  2. Bog'lash ID_ASL va ID_YORDAMCHI ustunlari bo'yicha. ID_TOZALASH = True bo'lsa IDlar
     solishtirishdan oldin tozalanadi: kichik harf, bo'sh joy, "-", "{", "}" olib tashlanadi
     (masalan "5E47-...-CA" va "5e47...ca" bir xil hisoblanadi). Fayllardagi IDlar o'zgarmaydi.
  3. USTUNLAR — yordamchidan olinadigan ustunlar: {yordamchidagi nomi: natijadagi nomi}.
     Natijadagi nom ASL da yo'q bo'lsa — yangi ustun qo'shiladi.
     ASL da bor bo'lsa — BOR_USTUN bo'yicha:
       "toldir"    — faqat ASL dagi bo'sh kataklar to'ldiriladi, bor qiymatlar o'zgarmaydi;
       "almashtir" — yordamchida qiymat bor joyda ASL qiymati almashtiriladi;
       "xato"      — skript to'xtaydi.
  4. BIRGA — bir-biriga bog'liq ustunlar guruhi (masalan fermer nomi va INN). "toldir" da bo'sh
     katak faqat guruhdagi boshqa ustunlar ASL da bo'sh yoki yordamchi bilan bir xil bo'lsa
     to'ldiriladi — bir fermer nomi yoniga boshqa fermer INN i tushib qolmasligi uchun.
  5. Yordamchida bir ID bir necha marta uchrasa — birinchi qatori olinadi; qiymatlari farqli
     takrorlar hisobotda alohida ko'rsatiladi.
  6. Natija CHIQISH ga yoziladi: ASL ning hamma qatorlari va geometriyasi o'zgarmaydi. ID
     topilmagan qatorlarda yangi ustunlar bo'sh qoladi.
  7. Hisobot Excel ga yoziladi (HISOBOT): nechta ID bor edi, nechtasi o'tdi, nechtasi o'tmadi,
     har ustun nechta katak to'ldi.

Ishga tushirish (repo papkasidan), avval pastdagi sozlamalarni to'g'rilang:
    .\\run.bat id_boyicha_qoshish.py
"""

from __future__ import annotations

# =====================================================================================
#  FOYDALANUVCHI SOZLAMALARI — har bir viloyat uchun shu yerni o'zgartiring
# =====================================================================================

# Ma'lumot qabul qiladigan shp (masalan qoshnidan_toldirish.py natijasi). O'zgarmaydi.
ASL = r"D:\natija\X_viloyati_v3.shp"

# Ustunlar olinadigan fayl (shp / gpkg / dbf / csv / xlsx). O'zgarmaydi (faqat o'qiladi).
YORDAMCHI = r"D:\malumot\X_viloyati_yordamchi.shp"

# Natija shp (yangi fayl; ASL yoki YORDAMCHI bilan bir xil bo'lishi mumkin emas).
CHIQISH = r"D:\natija\X_viloyati_v4.shp"

# Excel hisobot. None — CHIQISH yonida "<chiqish nomi>_hisobot.xlsx".
HISOBOT = None

# ID ustunlari nomi: ASL dagi va YORDAMCHI dagi (har xil bo'lishi mumkin).
ID_ASL = "unique_id"
ID_YORDAMCHI = "uniqueid"

# IDlarni solishtirishdan oldin tozalash (kichik harf; bo'sh joy, "-", "{", "}" olib tashlanadi).
ID_TOZALASH = True

# Yordamchidan olinadigan ustunlar: {yordamchidagi nomi: natijadagi nomi} (shp: 10 belgigacha).
# Nomini o'zgartirmaslik uchun ikkalasini bir xil yozing.
USTUNLAR = {
    "farmer_nam": "full_name",
    "farmer_inn": "tax_number",
}

# Natijadagi nom ASL da allaqachon bor bo'lsa:
#   "toldir"    — faqat bo'sh kataklarni to'ldirish (bor qiymat o'zgarmaydi);
#   "almashtir" — yordamchida qiymat bor joyda almashtirish;
#   "xato"      — to'xtash.
BOR_USTUN = "toldir"

# Birga to'ldiriladigan ustunlar guruhlari (natijadagi nomlar). Bo'sh katak faqat guruhdagi
# boshqa ustunlar ASL da bo'sh yoki yordamchi bilan bir xil bo'lsa to'ldiriladi. Kerak bo'lmasa [].
BIRGA = [["full_name", "tax_number"]]

# Guruhda nomlarni solishtirish: "yumshoq" — yuridik shakl (FX, MCHJ, fermer xo'jaligi,
# ФЕРМЕРСКОЕ ХОЗЯЙСТВО ...), kirill/lotin, apostrof va kesilgan nom farqi hisobga olinmaydi;
# "qattiq" — faqat harf va raqam bir xil bo'lsa.
NOM_SOLISHTIRISH = "yumshoq"

# Yordamchi xlsx bo'lsa — varaq nomi (None — birinchi varaq). csv bo'lsa — ajratuvchi.
XLSX_VARAQ = None
CSV_AJRATUVCHI = ","

# CHIQISH (va hisobot) allaqachon bor bo'lsa ustiga yozilsinmi.
USTIGA_YOZ = False

# =====================================================================================

import sys  # noqa: E402
from pathlib import Path  # noqa: E402

# ---- conda muhiti: 'activate' qilinmagan bo'lsa ham ishlashi uchun (numpy dan OLDIN) ----
# python.exe to'g'ridan-to'g'ri chaqirilsa Library\bin PATH da bo'lmaydi va boshqa joydagi mos kelmaydigan
# DLL yuklanadi (numpy / GDAL 0xc06d007f bilan qulaydi). Oddiy (conda emas) Python da hech narsa qilmaydi.
import os  # noqa: E402

_p = sys.prefix
_dirs = [_p, os.path.join(_p, "Library", "mingw-w64", "bin"), os.path.join(_p, "Library", "usr", "bin"),
         os.path.join(_p, "Library", "bin"), os.path.join(_p, "Scripts"), os.path.join(_p, "bin")]
os.environ["PATH"] = os.pathsep.join([d for d in _dirs if os.path.isdir(d)] + [os.environ.get("PATH", "")])
for _d in _dirs:
    if os.path.isdir(_d) and hasattr(os, "add_dll_directory"):
        try:
            os.add_dll_directory(_d)
        except OSError:
            pass
for _var, _sub in (("GDAL_DATA", "gdal"), ("PROJ_LIB", "proj"), ("PROJ_DATA", "proj")):
    _d = os.path.join(_p, "Library", "share", _sub)
    if os.path.isdir(_d):
        os.environ[_var] = _d
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

import hashlib  # noqa: E402
import re  # noqa: E402
import time  # noqa: E402

import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
import pyogrio  # noqa: E402
import shapely  # noqa: E402

# ---- yordamchi funksiyalar (fayl boshqa fayllarsiz ishlashi uchun shu yerda) --------------------
SHP_QOSHIMCHA = (".shp", ".shx", ".dbf", ".prj", ".cpg", ".qix", ".sbn", ".sbx")


def log(msg: str) -> None:
    satr = f"[{time.strftime('%H:%M:%S')}] {msg}"
    kod = getattr(sys.stdout, "encoding", None) or "utf-8"
    # Windows konsoli (cp1251 va h.k.) chiqara olmaydigan belgilar "?" bo'ladi — skript to'xtamaydi
    print(satr.encode(kod, errors="replace").decode(kod, errors="replace"), flush=True)


def shp_xesh(asos: Path) -> dict[str, str]:
    """shp ning hamma yon fayllari xeshi — fayl o'zgarmaganini tekshirish uchun."""
    out = {}
    for q in SHP_QOSHIMCHA:
        p = asos.with_suffix(q)
        if p.exists():
            out[p.name] = hashlib.sha256(p.read_bytes()).hexdigest()
    return out


def shp_ochir(asos: Path) -> None:
    for q in SHP_QOSHIMCHA:
        p = asos.with_suffix(q)
        if p.exists():
            p.unlink()


def crs_matn(crs) -> str:
    """EPSG:3857 (WGS 84 / Pseudo-Mercator) ko'rinishida."""
    if crs is None:
        return "yo'q"
    return f"{crs.to_string()} ({crs.name})"


def foiz(qism: float, jami: float) -> float:
    return round(qism / jami, 4) if jami else 0.0


def excel_yoz(yol: Path, varaqlar: dict[str, pd.DataFrame], foiz_ustunlar: set[str]) -> None:
    """Varaqlarni Excel ga yozish: qalin sarlavha, ustun kengligi, muzlatilgan birinchi qator."""
    from openpyxl.styles import Alignment, Font, PatternFill

    sarlavha_fon = PatternFill("solid", fgColor="DDEBF7")
    with pd.ExcelWriter(yol, engine="openpyxl") as w:
        for nom, df in varaqlar.items():
            df.to_excel(w, sheet_name=nom, index=False)
            ws = w.sheets[nom]
            for c in ws[1]:
                c.font = Font(bold=True)
                c.fill = sarlavha_fon
                c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.freeze_panes = "A2"
            for j, ustun in enumerate(df.columns, start=1):
                harf = ws.cell(row=1, column=j).column_letter
                uz = max([len(str(ustun))] + [len(str(v)) for v in df[ustun].head(500)])
                ws.column_dimensions[harf].width = min(max(uz + 2, 10), 70)
                if ustun in foiz_ustunlar:
                    for qator in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                        qator[0].number_format = "0.0%"

GEO_KENG = {".shp", ".gpkg", ".dbf", ".geojson", ".json"}


def fayl_xesh(p: Path) -> dict[str, str]:
    """shp bo'lsa yon fayllari bilan, aks holda faylning o'zi."""
    if p.suffix.lower() == ".shp":
        return shp_xesh(p)
    return {p.name: hashlib.sha256(p.read_bytes()).hexdigest()}


def matnga(v):
    """Qiymatni matnga: 308337301.0 -> "308337301"; bo'sh -> None."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, (float, np.floating)) and float(v).is_integer():
        return str(int(v))
    s = str(v).strip()
    return s or None


def id_tayyorla(s: pd.Series) -> pd.Series:
    """Solishtirish uchun ID: matn, bo'sh joysiz; ID_TOZALASH bo'lsa kichik harf, -{} siz.
    Bo'sh / NaN — None."""
    t = s.astype(object).map(matnga)
    if ID_TOZALASH:
        t = t.str.lower().str.replace(r"[\s{}\-]", "", regex=True)
    return t.where(t.notna() & (t != ""), None)


def bosh(s: pd.Series) -> np.ndarray:
    """None / NaN / bo'sh matn / "nan" — bo'sh."""
    return np.array([matnga(v) is None or matnga(v).lower() in ("nan", "none") for v in s.values])


def solishtir(v) -> str:
    """Bir xilligini tekshirish uchun: katta harf, faqat harf va raqam (apostrof, probel, "-" siz)."""
    m = matnga(v)
    return "" if m is None else "".join(ch for ch in m.upper() if ch.isalnum())


# ---- yumshoq nom solishtirish (fermer / tashkilot nomlari) -------------------------------
KIRILL = {"А": "A", "Б": "B", "В": "V", "Г": "G", "Д": "D", "Е": "E", "Ё": "YO", "Ж": "J", "З": "Z", "И": "I",
          "Й": "Y", "К": "K", "Л": "L", "М": "M", "Н": "N", "О": "O", "П": "P", "Р": "R", "С": "S", "Т": "T",
          "У": "U", "Ф": "F", "Х": "X", "Ц": "TS", "Ч": "CH", "Ш": "SH", "Щ": "SH", "Ъ": "", "Ы": "I", "Ь": "",
          "Э": "E", "Ю": "YU", "Я": "YA", "Ў": "O", "Қ": "Q", "Ғ": "G", "Ҳ": "H"}
YURIDIK_SHAKL = re.compile("|".join([
    r"MAS.?ULIYATI\s+CHEKLANGAN\s+JAMIYATI?", r"OBSHESTVO\s+S\s+OGRANICHENNOY\s+OTVETSTVENNOSTYU",
    r"MASULIYATI\s+CHEKLANGAN\s+JAMIYATI?", r"FERMER\s+XO.?JALIGI", r"FERMER\s+XUJALIGI",
    r"FERMERSKOE\s+XOZYAYSTVO", r"DEHQON\s+XO.?JALIGI", r"DEXKANSKOE\s+XOZYAYSTVO", r"XUSUSIY\s+KORXONA(SI)?",
    r"CHASTNOE\s+PREDPRIYATIE", r"QISHLOQ\s+XO.?JALIK\s+KOOPERATIVI?", r"\bMCHJ\b", r"\bOOO\b",
    r"\bF\s*/?\s*X\b", r"\bD\s*X\b", r"\bXK\b", r"\bCHP\b", r"\bQK\b", r"\bAJ\b"]))


def nom_yadrosi(v) -> str:
    """Nomning "yadrosi": kirill -> lotin, apostrofsiz, yuridik shaklsiz (FX, MCHJ, fermer xo'jaligi,
    ФЕРМЕРСКОЕ ХОЗЯЙСТВО ...), H/X bir xil, faqat harf va raqam."""
    m = matnga(v)
    if m is None:
        return ""
    t = "".join(KIRILL.get(ch, ch) for ch in m.upper())
    t = re.sub(r"[`'‘’ʻʼ´]", "", t)
    t = YURIDIK_SHAKL.sub(" ", t).replace("H", "X")
    return "".join(ch for ch in t if ch.isalnum())


def bir_xil(x, y) -> bool:
    """Ikki qiymat bir narsami: aynan (harf-raqam bo'yicha) yoki NOM_SOLISHTIRISH = "yumshoq" da
    nom yadrosi bir xil / biri ikkinchisining kesilgan boshi (kamida 8 belgi va 60%)."""
    if solishtir(x) == solishtir(y):
        return True
    if NOM_SOLISHTIRISH != "yumshoq":
        return False
    a, b = nom_yadrosi(x), nom_yadrosi(y)
    if not a or not b:
        return False
    if a == b:
        return True
    q, u = sorted((a, b), key=len)
    return len(q) >= 8 and len(q) >= 0.6 * len(u) and u.startswith(q)


def yordamchi_oqi(p: Path, ustunlar: list[str]) -> pd.DataFrame:
    kerak = list(dict.fromkeys([ID_YORDAMCHI] + ustunlar))
    k = p.suffix.lower()
    if k in GEO_KENG:
        bor = list(pyogrio.read_info(p)["fields"])
        yoq = [c for c in kerak if c not in bor]
        if yoq:
            sys.exit(f"YORDAMCHI da ustun yo'q: {yoq}\n  Bor ustunlar: {bor}")
        return pyogrio.read_dataframe(p, columns=kerak, read_geometry=False)
    if k == ".csv":
        df = pd.read_csv(p, sep=CSV_AJRATUVCHI, dtype={ID_YORDAMCHI: str}, encoding="utf-8-sig")
    elif k in (".xlsx", ".xls"):
        df = pd.read_excel(p, sheet_name=XLSX_VARAQ or 0, dtype={ID_YORDAMCHI: str})
    else:
        sys.exit(f"YORDAMCHI turi qo'llanmaydi: {k} (shp, gpkg, dbf, csv, xlsx)")
    yoq = [c for c in kerak if c not in df.columns]
    if yoq:
        sys.exit(f"YORDAMCHI da ustun yo'q: {yoq}\n  Bor ustunlar: {list(df.columns)}")
    return df[kerak]


def asl_turiga(qiymat: pd.Series, namuna: pd.Series) -> pd.Series:
    """Yordamchi qiymatini ASL ustuni turiga keltirish (matn ustunga 308337301.0 -> "308337301")."""
    if pd.api.types.is_numeric_dtype(namuna):
        return pd.to_numeric(qiymat, errors="coerce")
    return qiymat.astype(object).map(matnga)


def main() -> int:
    boshlanish = time.strftime("%Y-%m-%d %H:%M:%S")
    asl_yol, yor_yol, chiqish = Path(ASL), Path(YORDAMCHI), Path(CHIQISH)
    hisobot = Path(HISOBOT) if HISOBOT else chiqish.with_name(f"{chiqish.stem}_hisobot.xlsx")
    for p in (asl_yol, yor_yol):
        if not p.exists():
            sys.exit(f"Fayl topilmadi: {p}")
    if chiqish.resolve() in {asl_yol.resolve(), yor_yol.resolve()}:
        sys.exit("CHIQISH ASL yoki YORDAMCHI bilan bir xil bo'lmasin — natija alohida faylga yoziladi.")
    for p in (chiqish, hisobot):
        if p.exists() and not USTIGA_YOZ:
            sys.exit(f"Natija allaqachon bor: {p}\n  Ustiga yozish uchun USTIGA_YOZ = True qiling.")
    if NOM_SOLISHTIRISH not in ("yumshoq", "qattiq"):
        sys.exit(f'NOM_SOLISHTIRISH "yumshoq" yoki "qattiq" bo\'lsin: {NOM_SOLISHTIRISH}')
    if BOR_USTUN not in ("toldir", "almashtir", "xato"):
        sys.exit(f'BOR_USTUN "toldir", "almashtir" yoki "xato" bo\'lsin: {BOR_USTUN}')
    nomlar = dict(USTUNLAR) if isinstance(USTUNLAR, dict) else {c: c for c in USTUNLAR}
    if not nomlar:
        sys.exit("USTUNLAR bo'sh — kamida bitta ustun kiriting.")
    uzun = [y for y in nomlar.values() if len(y) > 10]
    if uzun:
        sys.exit(f"shp ustun nomi 10 belgidan oshmasin: {uzun}")
    if len(set(v.lower() for v in nomlar.values())) != len(nomlar):
        sys.exit(f"Natijadagi ustun nomlari takrorlanadi: {list(nomlar.values())}")
    teskari = {y: e for e, y in nomlar.items()}  # natijadagi nom -> yordamchidagi nom

    xesh_asl, xesh_yor = fayl_xesh(asl_yol), fayl_xesh(yor_yol)

    # ---- o'qish -------------------------------------------------------------------
    asl = pyogrio.read_dataframe(asl_yol)
    n = len(asl)
    log(f"ASL:       {asl_yol}  ({n} qator, {asl.crs})")
    if ID_ASL not in asl.columns:
        sys.exit(f"ASL da '{ID_ASL}' ustuni yo'q. Bor ustunlar: {list(asl.columns)}")
    asl_nomi = {c.lower(): c for c in asl.columns}
    mavjud = {y: asl_nomi[y.lower()] for y in nomlar.values() if y.lower() in asl_nomi}
    if mavjud and BOR_USTUN == "xato":
        sys.exit(f"Bu nomlar ASL da allaqachon bor: {list(mavjud)}\n"
                 f"  USTUNLAR da nomini o'zgartiring yoki BOR_USTUN = \"toldir\" qiling.")
    if mavjud:
        log(f"ASL da bor ustunlar ({BOR_USTUN}): {list(mavjud.values())}")
    guruhlar = [[g for g in grp if g in teskari] for grp in (BIRGA or [])]
    guruhlar = [g for g in guruhlar if len(g) > 1]
    yor = yordamchi_oqi(yor_yol, list(nomlar))
    log(f"YORDAMCHI: {yor_yol}  ({len(yor)} qator)")

    a_id = id_tayyorla(asl[ID_ASL])
    y_id = id_tayyorla(yor[ID_YORDAMCHI])

    # ---- yordamchi: takrorlar ------------------------------------------------------
    yv = yor[list(nomlar)].copy()
    yv["_id"] = y_id.values
    yv = yv[yv["_id"].notna()]
    takror_mask = yv["_id"].duplicated(keep=False)
    takror_id = yv.loc[takror_mask, "_id"].unique()
    farqli = []
    if len(takror_id):
        t = yv[takror_mask]
        farqli = [i for i, g in t.groupby("_id") if len(g.drop(columns="_id").drop_duplicates()) > 1]
    ybir = yv.drop_duplicates("_id", keep="first").set_index("_id")

    # ---- bog'lash ------------------------------------------------------------------
    natija = asl.copy()
    qiy = ybir.reindex(a_id.values)
    qiy.index = asl.index
    topildi = a_id.notna().values & a_id.isin(ybir.index).values

    # guruh sharti: ASL dagi qiymat bo'sh yoki yordamchi bilan bir xil
    def guruh_ok(nom: str) -> np.ndarray:
        ok = np.ones(n, bool)
        for grp in guruhlar:
            if nom not in grp:
                continue
            for boshqa in grp:
                if boshqa == nom or boshqa not in mavjud:
                    continue
                a_v = asl[mavjud[boshqa]]
                y_v = qiy[teskari[boshqa]]
                teng = np.array([bir_xil(x, y) for x, y in zip(a_v.values, y_v.values)])
                ok &= bosh(a_v) | teng
        return ok

    ustun_rows = []
    nomuvofiq = np.zeros(n, bool)
    yozildi_maska = {}
    for eski, yangi in nomlar.items():
        y_v = qiy[eski]
        y_bor = ~bosh(y_v) & topildi
        if yangi not in mavjud:
            natija[yangi] = y_v.values
            bor = int(natija[yangi].notna().sum())
            ustun_rows.append({
                "Yordamchidagi ustun": eski, "Natijadagi ustun": yangi, "Holat": "yangi ustun",
                "Avval bo'sh edi": n, "To'ldirildi": bor, "Bo'sh qoldi": n - bor,
                "Avvaldan bor, o'zgarmadi": 0, "Avvaldan bor, almashtirildi": 0,
                "Bor qiymat yordamchidan farqli (o'zgarmadi)": 0, "Guruh mos emas — to'ldirilmadi": 0,
                "To'ldi, %": foiz(bor, n),
            })
            log(f"  {eski} -> {yangi} (yangi ustun): {bor} qator")
            continue
        a_nom = mavjud[yangi]
        a_v = asl[a_nom]
        a_bosh = bosh(a_v)
        yv_t = asl_turiga(y_v, a_v)
        farq = ~a_bosh & y_bor & ~np.array([bir_xil(x, y) for x, y in zip(a_v.values, y_v.values)])
        if BOR_USTUN == "toldir":
            g_ok = guruh_ok(yangi)
            yoz = a_bosh & y_bor & g_ok
            guruh_tusdi = a_bosh & y_bor & ~g_ok
            nomuvofiq |= guruh_tusdi
            almash = np.zeros(n, bool)
        else:  # almashtir
            yoz = y_bor
            guruh_tusdi = np.zeros(n, bool)
            almash = ~a_bosh & y_bor & farq
        ustun = a_v.astype(object).copy()
        ustun[yoz] = yv_t[yoz].values
        natija[a_nom] = ustun if not pd.api.types.is_numeric_dtype(a_v) else pd.to_numeric(ustun, errors="coerce")
        yozildi_maska[a_nom] = yoz
        bosh_edi = int(a_bosh.sum())
        toldi = int((yoz & a_bosh).sum())
        ustun_rows.append({
            "Yordamchidagi ustun": eski, "Natijadagi ustun": a_nom,
            "Holat": "ASL da bor: faqat bo'sh kataklar" if BOR_USTUN == "toldir" else "ASL da bor: almashtirildi",
            "Avval bo'sh edi": bosh_edi, "To'ldirildi": toldi, "Bo'sh qoldi": bosh_edi - toldi,
            "Avvaldan bor, o'zgarmadi": int((~a_bosh).sum() - almash.sum()),
            "Avvaldan bor, almashtirildi": int(almash.sum()),
            "Bor qiymat yordamchidan farqli (o'zgarmadi)": int(farq.sum()) if BOR_USTUN == "toldir" else 0,
            "Guruh mos emas — to'ldirilmadi": int(guruh_tusdi.sum()),
            "To'ldi, %": foiz(toldi, bosh_edi),
        })
        log(f"  {eski} -> {a_nom} ({BOR_USTUN}): bo'sh edi {bosh_edi}, to'ldirildi {toldi}, "
            f"guruh mos emas {int(guruh_tusdi.sum())}, almashtirildi {int(almash.sum())}")

    asl_ust = [c for c in asl.columns if c != "geometry"]
    yangi_ust = [y for y in nomlar.values() if y not in mavjud]
    natija = natija[asl_ust + yangi_ust + ["geometry"]]

    a_bosh_id = a_id.isna().values
    a_uniq = pd.Series(a_id.dropna().unique())
    y_uniq = pd.Series(ybir.index)
    ortiqcha = y_uniq[~y_uniq.isin(a_uniq)]
    otdi_id = int(a_uniq.isin(ybir.index).sum())
    log(f"ID: ASL da {len(a_uniq)} ta noyob (qator {n}, bo'sh {int(a_bosh_id.sum())}); "
        f"YORDAMCHI da {len(y_uniq)} ta noyob (qator {len(yor)}, bo'sh {int(y_id.isna().sum())}, "
        f"takror ID {len(takror_id)}, shundan qiymati farqli {len(farqli)})")
    log(f"  ID orqali o'tdi: {otdi_id} ID ({int(topildi.sum())} qator);  o'tmadi: {len(a_uniq) - otdi_id} ID "
        f"({int((~topildi & ~a_bosh_id).sum())} qator) + bo'sh ID {int(a_bosh_id.sum())} qator;  "
        f"yordamchida ortiqcha: {len(ortiqcha)} ID")

    # ---- yozish va tekshiruv ------------------------------------------------------------
    chiqish.parent.mkdir(parents=True, exist_ok=True)
    shp_ochir(chiqish)
    pyogrio.write_dataframe(natija, chiqish, driver="ESRI Shapefile", encoding="UTF-8")
    log(f"Yozildi: {chiqish}")
    asl_ok = fayl_xesh(asl_yol) == xesh_asl
    yor_ok = fayl_xesh(yor_yol) == xesh_yor
    qayta = pyogrio.read_dataframe(chiqish)
    soni_ok = len(qayta) == n
    geom_ok = soni_ok and bool(shapely.equals(qayta.geometry.values, asl.geometry.values).all())
    crs_ok = qayta.crs == asl.crs
    # o'zgarmasligi kerak: to'ldirilmagan ustunlar to'liq; to'ldirilganlarda — yozilmagan kataklar
    eski_ok = soni_ok and all(c in qayta.columns for c in asl_ust)

    def matn_qator(s: pd.Series) -> np.ndarray:  # bo'sh -> "" (NaN != NaN muammosisiz)
        return s.astype(object).where(s.notna(), "").astype(str).values

    if eski_ok:
        for c in asl_ust:
            q, a = matn_qator(qayta[c]), matn_qator(asl[c])
            m = ~yozildi_maska.get(c, np.zeros(n, bool))
            if not (q[m] == a[m]).all():
                eski_ok = False
                log(f"XATO: {c} ustunida yozilmagan kataklar o'zgargan")
    for ok, matn in [(asl_ok, "ASL fayli o'zgargan!"), (yor_ok, "YORDAMCHI fayli o'zgargan!"),
                     (soni_ok, "qator soni farq qiladi"), (geom_ok, "geometriya farq qiladi"),
                     (crs_ok, "CRS farq qiladi"), (eski_ok, "ASL qiymatlari o'zgargan")]:
        if not ok:
            log(f"XATO: {matn}")
    xato = not (asl_ok and yor_ok and soni_ok and geom_ok and crs_ok and eski_ok)
    if not xato:
        log("Tekshiruv: ASL va YORDAMCHI o'zgarmagan; qator soni, geometriya, CRS bir xil; "
            "ASL dagi bor qiymatlar o'zgarmagan.")

    # ---- Excel hisobot ------------------------------------------------------------------
    def ha(ok: bool) -> str:
        return "Ha ✔" if ok else "YO'Q ✘"

    may = asl["maydon"].astype(float).values if "maydon" in asl.columns else None

    def ga(m: np.ndarray) -> str:
        return "" if may is None else f"{np.nansum(may[m]):,.2f} ga".replace(",", " ")

    umumiy = [
        ("Sana", boshlanish, ""),
        ("— FAYLLAR —", "", ""),
        ("ASL (ma'lumot qabul qilgan)", str(asl_yol), "o'zgarmadi, faqat o'qildi"),
        ("YORDAMCHI (ustunlar olingan)", str(yor_yol), "o'zgarmadi, faqat o'qildi"),
        ("CHIQISH (natija)", str(chiqish), "yangi fayl"),
        ("CRS (ASL = CHIQISH)", crs_matn(asl.crs), "geometriya o'zgarmadi"),
        ("— ID —", "", ""),
        ("ID ustuni: ASL / YORDAMCHI", f"{ID_ASL} / {ID_YORDAMCHI}", ""),
        ("ID tozalash", "ha" if ID_TOZALASH else "yo'q",
         "kichik harf; bo'sh joy, '-', '{', '}' olib tashlanadi (fayllar o'zgarmaydi)" if ID_TOZALASH else ""),
        ("— ASL —", "", ""),
        ("Qatorlar", n, ga(np.ones(n, bool))),
        ("Noyob ID", len(a_uniq), ""),
        ("ID bo'sh qatorlar", int(a_bosh_id.sum()), "bog'lanmaydi"),
        ("— YORDAMCHI —", "", ""),
        ("Qatorlar", len(yor), ""),
        ("Noyob ID", len(y_uniq), ""),
        ("ID bo'sh qatorlar", int(y_id.isna().sum()), ""),
        ("Takrorlangan ID", len(takror_id), "birinchi qatori olindi"),
        ("   shundan qiymatlari farqli", len(farqli), "'Takrorlar' varag'ida"),
        ("— ID BO'YICHA —", "", ""),
        ("ID orqali o'tdi: noyob ID", otdi_id, f"ASL dagi noyob IDlarning {foiz(otdi_id, len(a_uniq)):.1%}"),
        ("ID orqali o'tdi: qator", int(topildi.sum()), ga(topildi)),
        ("O'tmadi: noyob ID (yordamchida yo'q)", len(a_uniq) - otdi_id, "'O'tmaganlar' varag'ida"),
        ("O'tmadi: qator (ID bo'shlari bilan)", int((~topildi).sum()), ga(~topildi)),
        ("Yordamchida bor, ASL da yo'q: noyob ID", len(ortiqcha), "'Ortiqcha IDlar' varag'ida"),
        ("— USTUNLAR —", "", ""),
        ("Ustunlar", ", ".join(f"{e} → {y}" if e != y else e for e, y in nomlar.items()), "batafsil 'Ustunlar' varag'ida"),
        ("ASL da bor ustunlar bilan", BOR_USTUN if mavjud else "—",
         {"toldir": "faqat bo'sh kataklar to'ldirildi, bor qiymatlar o'zgarmadi",
          "almashtir": "yordamchida qiymat bor joyda almashtirildi"}.get(BOR_USTUN, "") if mavjud else ""),
    ]
    if guruhlar:
        umumiy.append(("Birga to'ldiriladigan guruhlar", "; ".join(" + ".join(g) for g in guruhlar),
                       "guruhdagi boshqa ustun ASL da bo'sh yoki yordamchi bilan bir xil bo'lsa to'ldiriladi"))
        umumiy.append(("Nomlarni solishtirish", NOM_SOLISHTIRISH,
                       "yuridik shakl (FX, MCHJ, fermer xo'jaligi ...), kirill/lotin, apostrof, kesilgan nom "
                       "farqi hisobga olinmaydi" if NOM_SOLISHTIRISH == "yumshoq" else "faqat harf va raqam bir xil bo'lsa"))
        umumiy.append(("Guruh mos emas — to'ldirilmagan qatorlar", int(nomuvofiq.sum()),
                       "'Guruh mos emas' varag'ida (masalan nom boshqa fermerniki)"))
    umumiy += [
        ("— TEKSHIRUV —", "", ""),
        ("ASL o'zgarmadi (fayl xeshi bir xil)", ha(asl_ok), ""),
        ("YORDAMCHI o'zgarmadi (fayl xeshi bir xil)", ha(yor_ok), ""),
        ("Qator soni bir xil", ha(soni_ok), ""),
        ("Geometriya bir xil", ha(geom_ok), ""),
        ("CRS bir xil", ha(crs_ok), ""),
        ("ASL dagi bor qiymatlar o'zgarmadi", ha(eski_ok), "faqat bo'sh kataklar / yangi ustunlar yozildi"
         if BOR_USTUN == "toldir" else "almashtirilganlardan tashqari"),
    ]

    varaqlar = {"Umumiy": pd.DataFrame(umumiy, columns=["Ko'rsatkich", "Qiymat", "Izoh"]),
                "Ustunlar": pd.DataFrame(ustun_rows)}

    if "tuman" in asl.columns:
        t = pd.DataFrame({"Tuman": asl["tuman"].fillna("(bo'sh)").values, "jami": 1, "otdi": topildi})
        for a_nom, m in yozildi_maska.items():
            t[f"{a_nom}: to'ldirildi"] = m
        g = t.groupby("Tuman").sum(numeric_only=True)
        g.loc["JAMI"] = g.sum()
        tum = pd.DataFrame({"Tuman": g.index, "Qatorlar": g.jami.values, "ID orqali o'tdi": g.otdi.values,
                            "O'tmadi": (g.jami - g.otdi).values, "O'tdi, %": (g.otdi / g.jami).round(4).values})
        for a_nom in yozildi_maska:
            tum[f"{a_nom}: to'ldirildi"] = g[f"{a_nom}: to'ldirildi"].values
        varaqlar["Tumanlar"] = tum

    if nomuvofiq.any():
        ix = np.flatnonzero(nomuvofiq)
        d = pd.DataFrame({"FID (qator, 0 dan)": ix, ID_ASL: asl[ID_ASL].values[ix]})
        if "tuman" in asl.columns:
            d["tuman"] = asl["tuman"].values[ix]
        for grp in guruhlar:
            for nom in grp:
                if nom in mavjud:
                    d[f"{nom} (ASL)"] = asl[mavjud[nom]].values[ix]
                d[f"{teskari[nom]} (yordamchi)"] = qiy[teskari[nom]].values[ix]
        varaqlar["Guruh mos emas"] = d

    ix = np.flatnonzero(~topildi)
    otm = pd.DataFrame({"FID (qator, 0 dan)": ix, ID_ASL: asl[ID_ASL].values[ix]})
    for c in ("tuman", "maydon", "crop"):
        if c in asl.columns:
            otm[c] = asl[c].values[ix]
    otm["Sabab"] = np.where(a_bosh_id[ix], "ASL da ID bo'sh", "ID yordamchida yo'q")
    varaqlar["O'tmaganlar"] = otm

    ort = pd.DataFrame({"Tozalangan ID": ortiqcha.values})
    ilk = pd.Series(yor[ID_YORDAMCHI].astype(str).values, index=y_id.values)
    ilk = ilk[~ilk.index.duplicated()]
    ort[f"{ID_YORDAMCHI} (yordamchidagi)"] = ilk.reindex(ortiqcha.values).values
    varaqlar["Ortiqcha IDlar"] = ort

    if len(farqli):
        tk = yv[yv["_id"].isin(farqli)].rename(columns={"_id": "Tozalangan ID"})
        varaqlar["Takrorlar"] = tk[["Tozalangan ID"] + list(nomlar)]

    try:
        hisobot.parent.mkdir(parents=True, exist_ok=True)
        excel_yoz(hisobot, varaqlar, {"To'ldi, %", "O'tdi, %"})
        log(f"Hisobot: {hisobot}")
    except PermissionError:
        log(f"XATO: hisobotni yozib bo'lmadi (Excel da ochiq bo'lishi mumkin): {hisobot}")
        xato = True
    return 1 if xato else 0


if __name__ == "__main__":
    sys.exit(main())
