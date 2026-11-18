"""Kadastrdan ekin yer shp ga egasi (fermer nomi), INN, kadastr raqami va yer turini o'tkazish.

Ekin dalasi va kadastr parseli bir narsa emas: ko'p dala bitta parsel ichida yotadi, lekin bir qismi bir
nechta parselga yoki ko'p kichik dehqon parsellariga bo'lingan. Shuning uchun har bir dala uchun:

  1. Dala ichidagi har bir kadastr parselining maydoni hisoblanadi (UTM da).
  2. Kadastrning o'zida ust-ust tushish tozalanadi (kadastr fayli o'zgarmaydi): zaxira parseli
     ("Bo'sh turgan ..." — hokimiyat) aniq egasi bor parsel bilan ust-ust tushsa, o'sha joy zaxiradan
     olib tashlanadi — auksion orqali ajratilgan yer ustun.
  3. Parsellar EGA bo'yicha yig'iladi (INN, bo'lmasa JShShIR, bo'lmasa nom yadrosi): bir egasining
     bir nechta parseli bitta ega.
  4. Eng katta eganing dala maydonidagi ulushi bo'yicha holat (kad_holat):
        aniq      >= 90%      — full_name, tax_number, cad_numb, yer_tur yoziladi
        asosiy    50–90%      — xuddi shu + "Ko'rib chiqish" varag'ida
        ko'p_ega  < 50%, bir nechta ega   — ega yozilmaydi (egalari "Bog'lanish" varag'ida)
        qisman    < 50%, kadastr dalaning ozgina qismini qoplaydi — ega yozilmaydi
        yo'q      kadastr dalaning MIN_QOPLANISH dan kamini qoplaydi — tegilmaydi
     cad_numb — tanlangan eganing dala ichidagi eng katta parseli raqami.
     yer_tur — kadastr yer_fondi (yer toifasi) dan YER_TURI_MOS jadvali bo'yicha; dalaning yarmidan
     ko'pi bitta toifada bo'lsa (ko'p_ega / qisman dalalarda ham).
  5. Ekin fayli o'zgarmaydi — natija yangi shp ga (hamma ustunlar va geometriya o'sha-o'sha, faqat shu
     4 ustun yangilanadi va kad_holat, kad_ulush, kad_soni, ega_soni qo'shiladi). Excel hisobot.

Ishga tushirish (repo papkasidan), avval pastdagi sozlamalarni to'g'rilang:
    .\\run.bat kadastrdan_ekinga.py
"""

from __future__ import annotations

# =====================================================================================
#  FOYDALANUVCHI SOZLAMALARI — har bir viloyat uchun shu yerni o'zgartiring
# =====================================================================================

# Ekin yerlar shp (o'zgarmaydi).
EKIN_FAYL = r"D:\malumot\X_viloyati_v4.shp"

# Birlashgan kadastr shp (kadastr_birlashtirish.py natijasi; o'zgarmaydi).
KADASTR_FAYL = r"D:\natija\X_viloyati_kadastr_birlashgan.shp"

# Natija shp (yangi fayl).
CHIQISH = r"D:\natija\X_viloyati_v5.shp"

# Excel hisobot. None — CHIQISH yonida "<chiqish nomi>_kadastr_hisobot.xlsx".
HISOBOT = None

# Ekin shp ustunlari
E_ID = "unique_id"
E_TUMAN = "tuman"            # hisobotdagi tumanlar jadvali uchun (bo'lmasa None)
E_NOM = "full_name"          # fermer / ega nomi
E_INN = "tax_number"         # INN (jismoniy shaxsda JShShIR — JSHSHIR_INN_ORNIGA)
E_KAD = "cad_numb"           # kadastr raqami
E_YER_TURI = "yer_tur"       # yer turi

# Kadastr shp ustunlari (kadastr_birlashtirish.py natijasidagi nomlar)
K_KAD = "kad_raqam"
K_ESKI_KAD = "eski_kad"
K_EGASI = "egasi"
K_EGA_TURI = "egasi_turi"
K_INN = "inn"
K_JSHSHIR = "jshshir"
K_YER_FONDI = "yer_fondi"
K_MULK = "mulk_turi"

# Zaxira parsellari (mulk_turi). Aniq egasi bor parsel bilan ust-ust tushgan joyi hisobga olinmaydi.
ZAXIRA_MULK = ["Bo‘sh turgan qishloq xo‘jaligiga mo‘ljallangan yer uchastkasi"]

# Chegaralar (eng katta eganing dala maydonidagi ulushi, 0..1)
ANIQ = 0.9                   # >= shu — "aniq"
ASOSIY = 0.5                 # >= shu — "asosiy" (ega yoziladi); kam bo'lsa — ega yozilmaydi
MIN_QOPLANISH = 0.1          # kadastr dalaning shundan kamini qoplasa — "yo'q" (tegilmaydi)
SEZILARLI = 0.05             # kad_soni / ega_soni da hisoblanadigan eng kichik ulush
BOGLANISH_MIN = 0.01         # "Bog'lanish" varag'iga yoziladigan eng kichik ulush

# Mavjud qiymatlar: "almashtir" — kadastr ustun (eski qiymat "O'zgargan qiymatlar" varag'ida);
#                   "toldir"    — faqat bo'sh kataklar to'ldiriladi.
MAVJUD = "almashtir"

# Egasida INN bo'lmasa (jismoniy shaxs) tax_number ga JShShIR yozilsinmi.
JSHSHIR_INN_ORNIGA = True

# Kadastr yer_fondi -> ekin yer_tur. Jadvalda yo'q toifa uchun kadastrdagi nomning o'zi yoziladi.
# (Apostrof turlari farq qilmaydi: ' ‘ ’ ʻ ʼ.)
YER_TURI_MOS = {
    "Fermer xo'jaliklariga berilgan yerlar": "Fermer xo‘jaligi",
    "Dehqon xo'jaliklarining yerlari": "Dehqon xo‘jaligi",
    "Ipakchilik va jun sanoati yerlari": "Ipak va jun sanoati yerlari",
    "Boshqa qishloq xo'jaligi korxonalari yerlari (klaster, agrofirma, korxona va tashkilotlar, jamiyatlar, "
    "shirkatlar)": "Korxonalarning qishloq xo‘jaligi yerlari",
    "Qishloq xo'jaligiga mo'ljallangan ijaraga berilmagan (zaxira) yerlar":
        "Zaxiradagi qishloq xo‘jaligiga mo‘ljallangan yer uchastkasi",
    "Qishloq xo'jaligi ilmiy-tadqiqot tashkilotlari va o'quv-tajriba xo'jaliklarining yerlari":
        "Qishloq xo‘jaligi ilmiy-tadqiqot tashkilotlari va o‘quv-tajriba xo‘jaliklari yerlari",
    "Yordamchi qishloq xo'jaliklarining yerlari": "Yordamchi qishloq xo‘jaliklarining yerlari",
}
# mulk_turi bo'yicha ustun turadigan yer turi (yer_fondi dan oldin qaraladi)
MULK_YER_TURI = {
    "Fermer xo'jaligi (auksion)": "Fermer xo‘jaligi (auksion)",
}

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
import warnings  # noqa: E402

import geopandas as gpd  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
import pyogrio  # noqa: E402
import shapely  # noqa: E402

warnings.filterwarnings("ignore", category=RuntimeWarning, module="pyogrio")

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
            h = hashlib.sha256()
            with open(p, "rb") as f:
                for blok in iter(lambda: f.read(1 << 24), b""):
                    h.update(blok)
            out[p.name] = h.hexdigest()
    return out


def shp_ochir(asos: Path) -> None:
    for q in SHP_QOSHIMCHA + (".shp.xml",):
        p = asos.with_name(asos.stem + q) if q == ".shp.xml" else asos.with_suffix(q)
        if p.exists():
            p.unlink()


def crs_matn(crs) -> str:
    if crs is None:
        return "yo'q"
    return f"{crs.to_string()} ({crs.name})"


def excel_yoz(yol: Path, varaqlar: dict, foiz_ustunlar: set[str]) -> None:
    """Varaqlarni Excel ga yozish. Varaq qiymati — DataFrame yoki [(sarlavha, DataFrame), ...] (bir
    varaqda bir nechta jadval ketma-ket). Qalin sarlavha, ustun kengligi, foiz formati."""
    from openpyxl.styles import Alignment, Font, PatternFill

    sarlavha_fon = PatternFill("solid", fgColor="DDEBF7")
    with pd.ExcelWriter(yol, engine="openpyxl") as w:
        for nom, qiymat in varaqlar.items():
            bloklar = [(None, qiymat)] if isinstance(qiymat, pd.DataFrame) else qiymat
            qator = 0
            kenglik: dict[int, int] = {}
            for sarlavha, df in bloklar:
                if sarlavha:
                    pd.DataFrame([[sarlavha]]).to_excel(w, sheet_name=nom, index=False, header=False, startrow=qator)
                    w.sheets[nom].cell(row=qator + 1, column=1).font = Font(bold=True, size=12)
                    qator += 1
                df.to_excel(w, sheet_name=nom, index=False, startrow=qator)
                ws = w.sheets[nom]
                for c in ws[qator + 1]:
                    if c.value is not None:
                        c.font = Font(bold=True)
                        c.fill = sarlavha_fon
                        c.alignment = Alignment(wrap_text=True, vertical="center")
                for j, ustun in enumerate(df.columns, start=1):
                    uz = max([len(str(ustun))] + [len(str(v)) for v in df[ustun].head(500)])
                    kenglik[j] = max(kenglik.get(j, 0), uz)
                    if ustun in foiz_ustunlar:
                        for r in ws.iter_rows(min_row=qator + 2, max_row=qator + 1 + len(df), min_col=j, max_col=j):
                            r[0].number_format = "0.0%"
                qator += len(df) + 3
            ws = w.sheets[nom]
            if len(bloklar) == 1 and not bloklar[0][0]:
                ws.freeze_panes = "A2"
            for j, uz in kenglik.items():
                ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = min(max(uz + 2, 10), 70)


# ---- nomlarni yumshoq solishtirish (fermer / tashkilot nomlari) ----------------------------------
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


def matnga(v):
    """Qiymatni matnga: 308337301.0 -> "308337301"; bo'sh -> None."""
    if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NA:
        return None
    if isinstance(v, (float, np.floating)) and float(v).is_integer():
        return str(int(v))
    s = str(v).strip()
    return None if s == "" or s.lower() in ("nan", "none", "null", "<na>") else s


def solishtir(v) -> str:
    m = matnga(v)
    return "" if m is None else "".join(ch for ch in m.upper() if ch.isalnum())


def nom_yadrosi(v) -> str:
    """Nomning "yadrosi": kirill -> lotin, apostrofsiz, yuridik shaklsiz, H/X bir xil, faqat harf va raqam."""
    m = matnga(v)
    if m is None:
        return ""
    t = "".join(KIRILL.get(ch, ch) for ch in m.upper())
    t = re.sub(r"[`'‘’ʻʼ´]", "", t)
    t = YURIDIK_SHAKL.sub(" ", t).replace("H", "X")
    return "".join(ch for ch in t if ch.isalnum())


def bir_xil(x, y) -> bool:
    """Ikki nom bir egami: aynan yoki yumshoq (nom yadrosi bir xil / kesilgan boshi, 8+ belgi va 60%)."""
    if solishtir(x) == solishtir(y):
        return True
    a, b = nom_yadrosi(x), nom_yadrosi(y)
    if not a or not b:
        return False
    if a == b:
        return True
    q, u = sorted((a, b), key=len)
    return len(q) >= 8 and len(q) >= 0.6 * len(u) and u.startswith(q)


def apos(v) -> str:
    """Lug'at kaliti uchun: apostrof turlari bir xil, ortiqcha bo'shliqsiz."""
    m = matnga(v)
    return "" if m is None else re.sub(r"\s+", " ", re.sub(r"[`‘’ʻʼ´]", "'", m)).strip()


def raqam(v):
    """INN / JShShIR — faqat raqamlar bo'lsa matn, aks holda None."""
    m = matnga(v)
    if m is None:
        return None
    m = m.replace(" ", "")
    return m if m.isdigit() and int(m) != 0 else None


def ustun(s, fn) -> np.ndarray:
    """Ustun qiymatlarini fn orqali object massivga (bo'sh -> None, NaN emas)."""
    return np.array([fn(v) for v in s.values], dtype=object)


def kad_asos(v):
    """Kadastr raqami: bo'shliqsiz, katta harf, '/RENT-...' qo'shimchasisiz."""
    m = matnga(v)
    if m is None:
        return None
    return re.sub(r"/RENT.*$", "", re.sub(r"\s+", "", m.upper()))


def xavfsiz(fn, a, b):
    """shapely ikki argumentli amal; topologiya xatosi bo'lsa — bittalab, kerak bo'lsa 1 sm aniqlikda."""
    try:
        return fn(a, b)
    except shapely.errors.GEOSException:
        out = np.empty(len(a), dtype=object)
        bb = b if isinstance(b, np.ndarray) else np.full(len(a), b, dtype=object)
        for i in range(len(a)):
            try:
                out[i] = fn(a[i], bb[i])
            except shapely.errors.GEOSException:
                out[i] = fn(shapely.set_precision(a[i], 0.01), shapely.set_precision(bb[i], 0.01))
        return out


# Ko'rib chiqish sabablari: kod -> (ustuvorlik, qisqa sabab, nima qilish kerak). 1 — eng muhim.
SABABLAR = {
    "ziddiyat": (1, "Kadastrda ikki xil ega ust-ust: ikkalasi ham dalaning yarmidan ko'pini qoplaydi",
                 "Kattaroq ulushli ega yozildi. Xaritada ikki parselni solishtiring, amaldagisini aniqlang "
                 "('Bog'lanish' varag'ida shu unique_id)."),
    "ega_ozgardi": (2, "Egasi o'zgardi: eski full_name boshqa ega edi",
                    "Kadastr bo'yicha yangi ega yozildi, eskisi 'O'zgargan qiymatlar' varag'ida. Yer haqiqatan "
                    "yangi egaga o'tganini tekshiring."),
    "kop_ega": (2, "Dalada bir nechta ega, hech biri yarmidan ko'p emas — egasi yozilmadi",
                "'Bog'lanish' varag'ida shu unique_id egalarini ko'ring. Dala bitta xo'jalikka tegishli bo'lsa — "
                "qo'lda yozing."),
    "eski_yoq": (3, "Eski full_name egasi bu dalada kadastrda topilmadi (eski qiymat qoldi)",
                 "Egasi yozilmagan dala; eski nom kadastrdagi egalardan birortasiga mos emas — tekshiring."),
    "asosiy": (3, "Asosiy ega dalaning 50–90% ini egallaydi",
               "Egasi yozildi. Chetdagi boshqa parsellarni xaritada tekshiring (kontur siljigan bo'lishi mumkin)."),
    "inn_yoq": (4, "Kadastrda egasining INN / JShShIR i yo'q — tax_number yozilmadi",
                "tax_number ni qo'lda kiriting."),
}

YANGI_UST = ["kad_holat", "kad_ulush", "kad_soni", "ega_soni"]
HOLATLAR = ["aniq", "asosiy", "ko'p_ega", "qisman", "yo'q"]


def main() -> int:
    boshlanish = time.strftime("%Y-%m-%d %H:%M:%S")
    e_yol, k_yol, chiqish = Path(EKIN_FAYL), Path(KADASTR_FAYL), Path(CHIQISH)
    hisobot = Path(HISOBOT) if HISOBOT else chiqish.with_name(f"{chiqish.stem}_kadastr_hisobot.xlsx")
    if MAVJUD not in ("almashtir", "toldir"):
        sys.exit("MAVJUD — 'almashtir' yoki 'toldir'.")
    for p in (e_yol, k_yol):
        if not p.exists():
            sys.exit(f"Fayl topilmadi: {p}")
    if chiqish.resolve() in {e_yol.resolve(), k_yol.resolve()}:
        sys.exit("CHIQISH ekin yoki kadastr fayli bilan bir xil bo'lmasin — natija alohida faylga yoziladi.")
    for p in (chiqish, hisobot):
        if p.exists() and not USTIGA_YOZ:
            sys.exit(f"Natija allaqachon bor: {p}\n  Ustiga yozish uchun USTIGA_YOZ = True qiling.")

    t0 = time.time()
    xesh_e, xesh_k = shp_xesh(e_yol), shp_xesh(k_yol)
    E = pyogrio.read_dataframe(e_yol)
    k_ust = [K_KAD, K_ESKI_KAD, K_EGASI, K_EGA_TURI, K_INN, K_JSHSHIR, K_YER_FONDI, K_MULK]
    K = pyogrio.read_dataframe(k_yol, columns=k_ust)
    log(f"Ekin: {e_yol}  ({len(E)} dala, {E.crs})")
    log(f"Kadastr: {k_yol}  ({len(K)} parsel, {K.crs})")
    e_kerak = [E_ID, E_NOM, E_INN, E_KAD, E_YER_TURI] + ([E_TUMAN] if E_TUMAN else [])
    yoq = [c for c in e_kerak if c not in E.columns]
    if yoq:
        sys.exit(f"Ekin shp da ustun yo'q: {yoq}\n  Bor ustunlar: {list(E.columns)}")
    tuqnash = [c for c in YANGI_UST if c.lower() in {x.lower() for x in E.columns}]
    if tuqnash:
        sys.exit(f"Ekin shp da bu ustun nomlari allaqachon bor: {tuqnash}")
    if K.crs != E.crs:
        K = K.to_crs(E.crs)

    # ---- 1. geometriya va kesishmalar ---------------------------------------------------------
    hisob_crs = E.estimate_utm_crs()
    eg = shapely.make_valid(E.geometry.to_crs(hisob_crs).values)
    kg = shapely.make_valid(K.geometry.to_crs(hisob_crs).values)
    e_ok = ~shapely.is_missing(eg) & ~shapely.is_empty(eg)
    k_ok = np.flatnonzero(~shapely.is_missing(kg) & ~shapely.is_empty(kg))
    ea = np.where(e_ok, shapely.area(eg), 0.0)
    ka = shapely.area(kg)
    tr = shapely.STRtree(kg[k_ok])
    ei, kj = tr.query(eg, predicate="intersects")
    kj = k_ok[kj]
    ia = np.zeros(len(ei))
    B = 100_000
    for s in range(0, len(ei), B):
        ia[s:s + B] = shapely.area(xavfsiz(shapely.intersection, eg[ei[s:s + B]], kg[kj[s:s + B]]))
    p = pd.DataFrame({"e": ei, "k": kj, "ia_asl": ia})
    p = p[p.ia_asl > 0].reset_index(drop=True)
    log(f"Dala–parsel kesishmalari: {len(p)} ({time.time() - t0:.0f} s)")

    # ---- 2. zaxira ust-ustini tozalash ----------------------------------------------------------
    zax_set = {apos(x) for x in ZAXIRA_MULK}
    k_zax = K[K_MULK].map(apos).isin(zax_set).values
    p["zax"] = k_zax[p.k.values]
    p["ia"] = p["ia_asl"]
    ikkisi = p.groupby("e").zax.agg(["any", "all"])
    f_ix = ikkisi.index[ikkisi["any"] & ~ikkisi["all"]]
    zax_kes, zax_dala = 0, 0
    if len(f_ix):
        boshqa = p[p.e.isin(f_ix) & ~p.zax]
        U = {e: shapely.union_all(kg[g.values]) for e, g in boshqa.groupby("e").k}
        zm = (p.e.isin(f_ix) & p.zax).values
        zi = np.flatnonzero(zm)
        ze = p.e.values[zi]
        kes = xavfsiz(shapely.intersection, eg[ze], kg[p.k.values[zi]])
        u_arr = np.empty(len(zi), dtype=object)
        u_arr[:] = [U[x] for x in ze]
        qoldi = shapely.area(xavfsiz(shapely.difference, kes, u_arr))
        p.loc[zm, "ia"] = qoldi
        zax_kes = int(np.sum(qoldi < p.loc[zm, "ia_asl"].values - 1))
        zax_dala = int(pd.Series(ze[qoldi < p.loc[zm, "ia_asl"].values - 1]).nunique())
    p["ul"] = p["ia"] / np.maximum(ea[p.e.values], 1e-9)
    p = p[p.ia > 0.01].reset_index(drop=True)
    log(f"Zaxira ust-ust tozalandi: {zax_dala} dalada ({zax_kes} kesishma kichraydi)")

    # ---- 3. ega bo'yicha yig'ish -----------------------------------------------------------------
    k_kad = ustun(K[K_KAD], matnga)
    k_eski = ustun(K[K_ESKI_KAD], kad_asos)
    k_ega = ustun(K[K_EGASI], matnga)
    k_tur = ustun(K[K_EGA_TURI], matnga)
    k_inn = ustun(K[K_INN], raqam)
    k_pin = ustun(K[K_JSHSHIR], raqam)
    k_yad = ustun(K[K_EGASI], nom_yadrosi)
    k_yf = ustun(K[K_YER_FONDI], matnga)
    mos = {apos(a): b for a, b in YER_TURI_MOS.items()}
    mulk_mos = {apos(a): b for a, b in MULK_YER_TURI.items()}
    k_yt = np.array([mulk_mos.get(apos(m)) or (mos.get(apos(f)) if f else None) or f
                     for m, f in zip(K[K_MULK].values, k_yf)], dtype=object)
    k_kalit = np.array([("I" + i) if i else ("P" + j) if j else ("N" + y) if y else f"K{n}"
                        for n, (i, j, y) in enumerate(zip(k_inn, k_pin, k_yad))], dtype=object)
    p["ega"] = k_kalit[p.k.values]

    N = len(E)
    qop = np.minimum(p.groupby("e").ul.sum().reindex(range(N), fill_value=0).values, 1.0)
    kad_soni = p[p.ul >= SEZILARLI].groupby("e").size().reindex(range(N), fill_value=0).values
    og = p.groupby(["e", "ega"], sort=False).ul.sum().reset_index()
    og["ul"] = og.ul.clip(upper=1.0)
    og = og.sort_values(["e", "ul"], ascending=[True, False])
    top = og.drop_duplicates("e").set_index("e")
    ikkinchi = og[og.duplicated("e")].drop_duplicates("e").set_index("e")
    top_ul = top.ul.reindex(range(N), fill_value=0).values
    top2_ul = ikkinchi.ul.reindex(range(N), fill_value=0).values
    ega_soni = og[og.ul >= SEZILARLI].groupby("e").size().reindex(range(N), fill_value=0).values

    holat = np.full(N, "yo'q", dtype=object)
    bor = qop >= MIN_QOPLANISH
    holat[bor & (top_ul >= ANIQ)] = "aniq"
    holat[bor & (top_ul >= ASOSIY) & (top_ul < ANIQ)] = "asosiy"
    past = bor & (top_ul < ASOSIY)
    holat[past & (ega_soni >= 2)] = "ko'p_ega"
    holat[past & (ega_soni < 2)] = "qisman"
    yoziladi = np.isin(holat, ["aniq", "asosiy"])
    ziddiyat = yoziladi & (top2_ul >= ASOSIY)

    # tanlangan eganing qiymatlari: eng katta parseli (nom, kadastr raqami), birinchi bor INN / JShShIR
    pc = p.merge(top[["ega"]].reset_index(), on=["e", "ega"]).sort_values(["e", "ia"], ascending=[True, False])
    eng = pc.drop_duplicates("e").set_index("e").k
    y_nom = np.full(N, None, dtype=object)
    y_kad = np.full(N, None, dtype=object)
    y_inn = np.full(N, None, dtype=object)
    y_pin = np.full(N, None, dtype=object)
    ix = eng.index.values
    y_nom[ix] = k_ega[eng.values]
    y_kad[ix] = k_kad[eng.values]
    pc["inn"] = k_inn[pc.k.values]
    pc["pin"] = k_pin[pc.k.values]
    t = pc.dropna(subset=["inn"]).drop_duplicates("e").set_index("e").inn
    y_inn[t.index.values] = t.values
    t = pc.dropna(subset=["pin"]).drop_duplicates("e").set_index("e").pin
    y_pin[t.index.values] = t.values
    for a in (y_nom, y_kad, y_inn, y_pin):
        a[~yoziladi] = None
    y_tax = np.where(pd.notna(y_inn), y_inn, y_pin if JSHSHIR_INN_ORNIGA else None)
    y_tax = np.array([x if isinstance(x, str) else None for x in y_tax], dtype=object)

    # yer turi: dalaning yarmidan ko'pi bitta toifada bo'lsa
    p["yt"] = k_yt[p.k.values]
    yg = p.dropna(subset=["yt"]).groupby(["e", "yt"]).ul.sum().reset_index().sort_values(
        ["e", "ul"], ascending=[True, False]).drop_duplicates("e").set_index("e")
    yg = yg[yg.ul >= ASOSIY]
    y_yer = np.full(N, None, dtype=object)
    y_yer[yg.index.values] = yg.yt.values
    y_yer[holat == "yo'q"] = None
    log("Holatlar: " + ", ".join(f"{h} {int((holat == h).sum())}" for h in HOLATLAR) +
        f"; ikki ega ust-ust {int(ziddiyat.sum())}")

    # ---- 4. eski qiymatlar bilan solishtirib yozish ----------------------------------------------
    o_nom = ustun(E[E_NOM], matnga)
    o_tax = ustun(E[E_INN], raqam)
    o_kad = ustun(E[E_KAD], matnga)
    o_yer = ustun(E[E_YER_TURI], matnga)
    ega_ozgardi = np.array([bool(yz and o and n and not bir_xil(o, n) and not (ot and ot == yt))
                            for yz, o, n, ot, yt in zip(yoziladi, o_nom, y_nom, o_tax, y_tax)])
    almashtir = MAVJUD == "almashtir"

    def qaror(eski, yangi, teng):
        """Har bir katak uchun: yozilsinmi va holati (to'ldirildi / almashtirildi / bir xil / '')."""
        bor_y = pd.notna(yangi) & np.array([v is not None for v in yangi])
        bosh_e = np.array([v is None for v in eski])
        bir = np.array([bool(y is not None and o is not None and teng(o, y)) for o, y in zip(eski, yangi)])
        yoz = bor_y & ~bir & (bosh_e | almashtir)
        h = np.full(len(eski), "", dtype=object)
        h[bor_y & bir] = "bir xil"
        h[yoz & bosh_e] = "to'ldirildi"
        h[yoz & ~bosh_e] = "almashtirildi"
        if not almashtir:
            h[bor_y & ~bir & ~bosh_e] = "qoldi (toldir rejimi)"
        return yoz, h

    yoz_nom, h_nom = qaror(o_nom, y_nom, lambda o, y: o == y)
    yoz_tax, h_tax = qaror(o_tax, y_tax, lambda o, y: o == y)
    yoz_kad, h_kad = qaror(o_kad, y_kad, lambda o, y: kad_asos(o) == kad_asos(y))
    yoz_yer, h_yer = qaror(o_yer, y_yer, lambda o, y: apos(o) == apos(y))
    # egasi boshqa bo'lib, yangi egasining INN / JShShIR i yo'q — eski INN boshqa odamniki, tozalanadi
    tozala_tax = almashtir & ega_ozgardi & np.array([v is None for v in y_tax]) & np.array([v is not None for v in o_tax])
    h_tax[tozala_tax] = "tozalandi"

    # eski egasi kadastrda yo'q (ega yozilmagan dalalarda)
    eski_yoq = np.zeros(N, dtype=bool)
    tek = np.flatnonzero(np.isin(holat, ["ko'p_ega", "qisman"]) & np.array([v is not None for v in o_nom]))
    if len(tek):
        pt = p[p.e.isin(tek)]
        yad = pt.groupby("e").k.agg(lambda s: {k_yad[x] for x in s} - {""})
        inns = pt.groupby("e").k.agg(lambda s: {k_inn[x] for x in s} | {k_pin[x] for x in s})
        for i in tek:
            on = nom_yadrosi(o_nom[i])
            eski_yoq[i] = not ((on and on in yad.get(i, set())) or (o_tax[i] and o_tax[i] in inns.get(i, set())))

    natija = E.copy()
    nom_v = natija[E_NOM].astype(object).to_numpy(copy=True)
    nom_v[yoz_nom] = y_nom[yoz_nom]
    natija[E_NOM] = nom_v
    kad_v = natija[E_KAD].astype(object).to_numpy(copy=True)
    kad_v[yoz_kad] = y_kad[yoz_kad]
    natija[E_KAD] = kad_v
    yer_v = natija[E_YER_TURI].astype(object).to_numpy(copy=True)
    yer_v[yoz_yer] = y_yer[yoz_yer]
    natija[E_YER_TURI] = yer_v
    tax_dt = natija[E_INN].dtype
    if pd.api.types.is_integer_dtype(tax_dt):
        tv = natija[E_INN].to_numpy(dtype="int64", copy=True)
        tv[yoz_tax] = [int(x) for x in y_tax[yoz_tax]]
        tv[tozala_tax] = 0
        natija[E_INN] = tv
    elif pd.api.types.is_float_dtype(tax_dt):
        tv = natija[E_INN].to_numpy(dtype="float64", copy=True)
        tv[yoz_tax] = [float(x) for x in y_tax[yoz_tax]]
        tv[tozala_tax] = np.nan
        natija[E_INN] = tv
    else:
        tv = natija[E_INN].astype(object).to_numpy(copy=True)
        tv[yoz_tax] = y_tax[yoz_tax]
        tv[tozala_tax] = None
        natija[E_INN] = tv
    natija["kad_holat"] = holat
    natija["kad_ulush"] = np.round(top_ul, 3)
    natija["kad_soni"] = kad_soni.astype("int32")
    natija["ega_soni"] = ega_soni.astype("int32")
    natija = natija[[c for c in natija.columns if c != "geometry"] + ["geometry"]]

    # ko'rib chiqish sabablari
    belgilar = {
        "ziddiyat": ziddiyat,
        "ega_ozgardi": ega_ozgardi & almashtir,
        "kop_ega": holat == "ko'p_ega",
        "eski_yoq": eski_yoq,
        "asosiy": holat == "asosiy",
        "inn_yoq": yoziladi & np.array([v is None for v in y_tax]),
    }
    sabab_qator = [[] for _ in range(N)]
    for kod, mask in belgilar.items():
        for i in np.flatnonzero(mask):
            sabab_qator[i].append(kod)
    log(f"Yozildi (katak): full_name {int(yoz_nom.sum())}, tax_number {int(yoz_tax.sum())} "
        f"(+tozalandi {int(tozala_tax.sum())}), cad_numb {int(yoz_kad.sum())}, yer_tur {int(yoz_yer.sum())}; "
        f"ko'rib chiqish {sum(bool(s) for s in sabab_qator)} dala ({time.time() - t0:.0f} s)")

    # ---- 5. yozish va tekshiruv ------------------------------------------------------------------
    chiqish.parent.mkdir(parents=True, exist_ok=True)
    shp_ochir(chiqish)
    pyogrio.write_dataframe(natija, chiqish, driver="ESRI Shapefile", encoding="UTF-8")
    log(f"Yozildi: {chiqish}")
    e_ok_x = shp_xesh(e_yol) == xesh_e
    k_ok_x = shp_xesh(k_yol) == xesh_k
    q = pyogrio.read_dataframe(chiqish, read_geometry=False, columns=[E_ID, E_NOM, E_INN, E_KAD, E_YER_TURI, "kad_holat"])
    soni_ok = len(q) == N
    id_ok = soni_ok and bool((q[E_ID].astype(str).values == E[E_ID].astype(str).values).all())

    def teng_ustun(a, b) -> bool:
        a = pd.Series(a).map(matnga).fillna("").values
        b = pd.Series(b).map(matnga).fillna("").values
        return bool((a == b).all())

    qiymat_ok = soni_ok and all([
        teng_ustun(q[E_NOM].values, nom_v), teng_ustun(q[E_KAD].values, kad_v),
        teng_ustun(q[E_YER_TURI].values, yer_v),
        teng_ustun(pd.Series(q[E_INN]).map(raqam).values, pd.Series(natija[E_INN]).map(raqam).values),
        teng_ustun(q["kad_holat"].values, holat)])
    xato = False
    for ok, matn in [(e_ok_x, "ekin fayli o'zgargan (boshqa dasturda tahrir qilinyaptimi?)"),
                     (k_ok_x, "kadastr fayli o'zgargan (boshqa dasturda tahrir qilinyaptimi?)"),
                     (soni_ok, "natijadagi dala soni ekin faylidan farq qiladi"),
                     (id_ok, "natijada unique_id tartibi o'zgargan"),
                     (qiymat_ok, "qayta o'qilgan qiymatlar yozilganidan farq qiladi (matn kesilganmi?)")]:
        if not ok:
            log(f"XATO: {matn}")
            xato = True
    if not xato:
        log("Tekshiruv: ekin va kadastr o'zgarmagan; natijada hamma dala, qiymatlar to'g'ri yozilgan.")

    # ---- 6. Excel hisobot ------------------------------------------------------------------------
    def ha(ok: bool) -> str:
        return "Ha ✔" if ok else "YO'Q ✘"

    ga = ea / 1e4
    jami_ga = float(ga.sum())
    umumiy = [
        ("Sana", boshlanish, ""),
        ("— FAYLLAR —", "", ""),
        ("Ekin yerlar", str(e_yol), "o'zgarmadi, faqat o'qildi"),
        ("Kadastr", str(k_yol), "o'zgarmadi, faqat o'qildi"),
        ("CHIQISH", str(chiqish), "yangi fayl: ekin fayli + yangilangan ustunlar"),
        ("CRS (natija)", crs_matn(E.crs), "ekin bilan bir xil, geometriya o'zgarmagan"),
        ("Maydon qaysi CRS da", crs_matn(hisob_crs), ""),
        ("Mavjud qiymatlar", MAVJUD, "almashtir — kadastr ustun; toldir — faqat bo'sh kataklar"),
        ("tax_number ga JShShIR (INN bo'lmasa)", "ha" if JSHSHIR_INN_ORNIGA else "yo'q", ""),
        ("— KIRISH —", "", ""),
        ("Ekin dalalari", N, f"{round(jami_ga, 2)} ga"),
        ("Kadastr parsellari", len(K), ""),
        ("Dala–parsel kesishmalari", len(p), ""),
        ("Zaxira ust-ust tozalangan dalalar", zax_dala, "zaxira parseli aniq egasi bor parsel bilan ust-ust"),
        ("Ikki xil ega ust-ust (ikkalasi ham >= 50%)", int(ziddiyat.sum()), "'Ko'rib chiqish' varag'ida, 1-ustuvorlik"),
    ]

    # Dalaning necha foizi bitta egaga tegishli
    def qator(nom, mask, holat_n, manosi, nima):
        return {"Eng katta ega ulushi": nom, "kad_holat": holat_n, "Dala": int(mask.sum()),
                "Maydon (ga)": round(float(ga[mask].sum()), 1), "Dala ulushi": float(mask.sum()) / N if N else 0,
                "Ma'nosi": manosi, "Nima yozildi": nima}

    m_a, m_s = holat == "aniq", holat == "asosiy"
    m_k, m_q, m_y = holat == "ko'p_ega", holat == "qisman", holat == "yo'q"
    ulush_jad = pd.DataFrame([
        qator(f"≥ {ANIQ:.0%}", m_a, "aniq", "Toza holat, bemalol yoziladi",
              "full_name, tax_number, cad_numb, yer_tur"),
        qator(f"{ASOSIY:.0%}–{ANIQ:.0%}", m_s, "asosiy", "Asosiy ega bor, chetiga boshqa parsel tegadi",
              "xuddi shu + 'Ko'rib chiqish' varag'ida"),
        qator(f"< {ASOSIY:.0%}", m_k | m_q, "ko'p_ega + qisman", "Bitta egaga tegishli emas",
              "ega yozilmadi; yer_tur — yarmidan ko'pi bitta toifada bo'lsa"),
        qator("   shundan ko'p egali", m_k, "ko'p_ega", "Dalada bir nechta ega (ko'pincha kichik dehqon parsellari)",
              "egalari 'Bog'lanish' varag'ida"),
        qator("   shundan kadastr qisman", m_q, "qisman", "Kadastr dalaning ozgina qismini qoplaydi", ""),
        qator(f"kadastr yo'q (< {MIN_QOPLANISH:.0%} qoplangan)", m_y, "yo'q", "Kadastrdan to'ldirib bo'lmaydi",
              "tegilmadi"),
        qator("JAMI", np.ones(N, dtype=bool), "", "", ""),
    ])

    def ust_qator(nom, eski, yangi_qiymat, h, toz=None):
        oldin = int(sum(v is not None for v in eski))
        keyin = int(pd.Series(yangi_qiymat).map(matnga if nom != E_INN else raqam).notna().sum())
        return {"Ustun": nom, "Oldin to'la": oldin, "Keyin to'la": keyin,
                "To'ldirildi (bo'sh edi)": int((h == "to'ldirildi").sum()),
                "Almashtirildi (boshqa qiymat edi)": int((h == "almashtirildi").sum()),
                "Tozalandi": int(toz.sum()) if toz is not None else 0,
                "Kadastr bilan bir xil edi": int((h == "bir xil").sum()),
                "Qoldi (toldir rejimi)": int((h == "qoldi (toldir rejimi)").sum())}

    ust_jad = pd.DataFrame([
        ust_qator(E_NOM, o_nom, nom_v, h_nom), ust_qator(E_INN, o_tax, natija[E_INN].values, h_tax, tozala_tax),
        ust_qator(E_KAD, o_kad, kad_v, h_kad), ust_qator(E_YER_TURI, o_yer, yer_v, h_yer)])

    soni = {k: int(m.sum()) for k, m in belgilar.items()}
    kr_umumiy = pd.DataFrame([{"Ustuvorlik": u, "Sabab": s, "Dala": soni[k], "Nima qilish kerak": n}
                              for k, (u, s, n) in sorted(SABABLAR.items(), key=lambda x: x[1][0]) if soni.get(k)])
    tekshiruv = pd.DataFrame([
        ("Ekin fayli o'zgarmadi (fayl xeshi bir xil)", ha(e_ok_x)),
        ("Kadastr fayli o'zgarmadi (fayl xeshi bir xil)", ha(k_ok_x)),
        ("Natijada dala soni = ekin fayli", ha(soni_ok)),
        ("unique_id tartibi bir xil", ha(id_ok)),
        ("Qayta o'qilgan qiymatlar = yozilgan", ha(qiymat_ok)),
    ], columns=["Tekshiruv", "Natija"])
    varaqlar = {"Umumiy": [
        (None, pd.DataFrame(umumiy, columns=["Ko'rsatkich", "Qiymat", "Izoh"])),
        ("Dalaning necha foizi bitta egaga tegishli (bir egasining bir nechta parseli qo'shib hisoblangan)",
         ulush_jad),
        ("Ustunlar: oldin va keyin", ust_jad),
        ("Ko'rib chiqish kerak (2-varaq)", kr_umumiy),
        ("Tekshiruv", tekshiruv),
    ]}

    # 2-varaq: ko'rib chiqiladigan dalalar
    eid = E[E_ID].values
    tum = E[E_TUMAN].values if E_TUMAN else np.full(N, None)
    kq = np.flatnonzero([bool(s) for s in sabab_qator])
    varaqlar["Ko'rib chiqish"] = pd.DataFrame({
        "Ustuvorlik": [min(SABABLAR[k][0] for k in sabab_qator[i]) for i in kq],
        "Sabab": [" | ".join(SABABLAR[k][1] for k in sabab_qator[i]) for i in kq],
        "Nima qilish kerak": [" | ".join(SABABLAR[k][2] for k in sabab_qator[i]) for i in kq],
        E_ID: eid[kq], "tuman": tum[kq], "maydon_ga": np.round(ga[kq], 3),
        "kad_holat": holat[kq], "Eng katta ega ulushi": np.round(top_ul[kq], 3),
        "Ikkinchi ega ulushi": np.round(top2_ul[kq], 3), "ega_soni": ega_soni[kq],
        "Eski full_name": o_nom[kq], "Yangi full_name": nom_v[kq],
        "Eski tax_number": o_tax[kq], "Yangi tax_number": pd.Series(natija[E_INN].values[kq]).map(raqam).values,
        "Eski cad_numb": o_kad[kq], "Yangi cad_numb": kad_v[kq],
    }).sort_values(["Ustuvorlik", "tuman", E_ID]).reset_index(drop=True)

    # O'zgargan qiymatlar (eski -> yangi)
    oz = []
    for ust_n, eski, yangi_q, h in ((E_NOM, o_nom, nom_v, h_nom), (E_INN, o_tax, y_tax, h_tax),
                                     (E_KAD, o_kad, kad_v, h_kad), (E_YER_TURI, o_yer, yer_v, h_yer)):
        ix2 = np.flatnonzero(np.isin(h, ["almashtirildi", "tozalandi"]))
        if not len(ix2):
            continue
        izoh = np.full(len(ix2), "", dtype=object)
        if ust_n == E_NOM:
            izoh = np.where(ega_ozgardi[ix2], "boshqa ega", "faqat yozilishi")
        elif ust_n == E_KAD:
            ek = np.array([k_eski[x] if x >= 0 else None for x in eng.reindex(ix2).fillna(-1).astype(int).values],
                          dtype=object)
            izoh = np.where([bool(o and kad_asos(o) == e_) for o, e_ in zip(o_kad[ix2], ek)],
                            "eski raqam yangilandi", "boshqa parsel")
        elif ust_n == E_INN:
            izoh = np.where(h[ix2] == "tozalandi", "ega o'zgardi, yangisining INN / JShShIR i yo'q", "")
        oz.append(pd.DataFrame({E_ID: eid[ix2], "tuman": tum[ix2], "Ustun": ust_n, "Eski": eski[ix2],
                                "Yangi": [None if h[i] == "tozalandi" else yangi_q[i] for i in ix2],
                                "Holat": h[ix2], "Izoh": izoh, "kad_holat": holat[ix2]}))
    if oz:
        varaqlar["O'zgargan qiymatlar"] = pd.concat(oz, ignore_index=True)

    # Ko'p egali dalalar
    kp = np.flatnonzero(m_k)
    if len(kp):
        pk = p[p.e.isin(kp)].copy()
        pk["nom"] = k_ega[pk.k.values]
        egl = pk.groupby(["e", "ega"]).agg(ul=("ul", "sum"), nom=("nom", "first")).reset_index().sort_values(
            ["e", "ul"], ascending=[True, False])
        roy = egl.groupby("e").apply(
            lambda g: "; ".join(f"{n or '(nomsiz)'} {u:.0%}" for n, u in zip(g.nom.head(5), g.ul.head(5))) +
            (f"; ... yana {len(g) - 5}" if len(g) > 5 else ""), include_groups=False)
        varaqlar["Ko'p egali dalalar"] = pd.DataFrame({
            E_ID: eid[kp], "tuman": tum[kp], "maydon_ga": np.round(ga[kp], 3), "ega_soni": ega_soni[kp],
            "kad_soni": kad_soni[kp], "Eng katta ega ulushi": np.round(top_ul[kp], 3),
            "Kadastr qoplagan": np.round(qop[kp], 3), "yer_tur": yer_v[kp],
            "Egalari (ulushi bilan, eng kattalari)": roy.reindex(kp).values}).sort_values(["tuman", E_ID])

    # Bog'lanish: dala <-> parsel
    pb = p[p.ul >= BOGLANISH_MIN].copy()
    tanl = set(zip(top.index.values, top.ega.values))
    varaqlar["Bog'lanish"] = pd.DataFrame({
        E_ID: eid[pb.e.values], "tuman": tum[pb.e.values], "kad_holat": holat[pb.e.values],
        "kad_raqam": k_kad[pb.k.values], "egasi": k_ega[pb.k.values], "egasi_turi": k_tur[pb.k.values],
        "inn": k_inn[pb.k.values], "jshshir": k_pin[pb.k.values], "yer_fondi": k_yf[pb.k.values],
        "Dala ichida, ga": np.round(pb.ia.values / 1e4, 4), "Dala ulushi": np.round(pb.ul.values, 4),
        "Parsel ulushi": np.round(pb.ia.values / np.maximum(ka[pb.k.values], 1e-9), 4),
        "Zaxira, ust-ust kesildi": np.where(pb.zax.values & (pb.ia.values < pb.ia_asl.values - 1), "ha", ""),
        "Tanlangan ega": ["ha" if (yoziladi[e_] and (e_, g_) in tanl) else ""
                          for e_, g_ in zip(pb.e.values, pb.ega.values)],
    }).sort_values([E_ID, "Dala ulushi"], ascending=[True, False])

    # Yer turi moslash
    ym = pd.DataFrame({"yf": k_yf, "mulk": ustun(K[K_MULK], matnga), "yt": k_yt})
    ym["qoida"] = [("mulk_turi" if mulk_mos.get(apos(m)) else "yer_fondi jadvali" if mos.get(apos(f)) else
                    "jadvalda yo'q — kadastr nomi yozildi" if f else "") for m, f in zip(ym.mulk, ym.yf)]
    yer_soni = pd.Series(y_yer[yoz_yer | (h_yer == "bir xil")]).value_counts()
    ymj = ym.groupby(["yf", "yt", "qoida"], dropna=False).size().reset_index(name="Kadastr parsellari")
    ymj["Natijada shu yer_tur li dalalar"] = ymj.yt.map(yer_soni).fillna(0).astype(int)
    varaqlar["Yer turi moslash"] = ymj.rename(columns={"yf": "Kadastr yer_fondi", "yt": "Natijadagi yer_tur",
                                                       "qoida": "Qoida"}).sort_values("Kadastr parsellari",
                                                                                      ascending=False)

    if E_TUMAN:
        td = pd.DataFrame({"Tuman": pd.Series(tum).fillna("(bo'sh)").values, "ga": ga, "holat": holat})
        tt = pd.crosstab(td.Tuman, td.holat).reindex(columns=HOLATLAR, fill_value=0)
        tt.columns.name = None
        tt.insert(0, "Dala", tt.sum(axis=1))
        tt.insert(1, "Maydoni, ga", td.groupby("Tuman").ga.sum().round(1))
        tt["Ko'rib chiqish"] = pd.Series(np.array([bool(s) for s in sabab_qator]), index=td.index).groupby(
            td.Tuman).sum()
        tt.loc["JAMI"] = tt.sum()
        varaqlar["Tumanlar"] = tt.reset_index()

    try:
        excel_yoz(hisobot, varaqlar, {"Dala ulushi", "Eng katta ega ulushi", "Ikkinchi ega ulushi",
                                      "Kadastr qoplagan", "Parsel ulushi"})
        log(f"Hisobot: {hisobot}")
    except PermissionError:
        log(f"XATO: hisobotni yozib bo'lmadi (Excel da ochiq bo'lishi mumkin): {hisobot}")
        xato = True
    log(f"Tugadi ({time.time() - t0:.0f} s)")
    return 1 if xato else 0


if __name__ == "__main__":
    sys.exit(main())
