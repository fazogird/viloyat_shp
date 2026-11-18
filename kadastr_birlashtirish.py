"""Ikki kadastr shp ni bitta to'liq kadastr fayliga birlashtirish (egasi, INN, JShShIR, kadastr raqami).

A — yangi uzkad (asos): egalar ma'lumoti to'liq (familiya / ism / otasining ismi alohida ustunlarda),
    INN (yuridik shaxs, fermer xo'jaligi) va JShShIR (jismoniy shaxs) bor, lekin ba'zi parsellar yo'q.
B — eski fayl: parsellar ko'proq, lekin eski; faqat kadastr raqami va to'liq nom (INN yo'q).

Qoida:
  1. Asos — A ning hamma parsellari va hamma ustunlari. Ustiga yagona ustunlar qo'shiladi:
     kad_raqam, eski_kad, egasi (jismoniy shaxsda familiya + ism + otasining ismi), egasi_turi,
     inn, jshshir, viloyat, tuman, tuman_kod, yer_fondi, huquq, mulk_turi, maydon_ga, manba.
  2. A da geometriyasi yo'q parsel — geometriya B dan (kadastr raqami bir xil bo'lsa).
  3. A da egasi bo'sh bo'lsa — B dagi shu kadastr raqamli parsel nomidan.
  4. B dagi parsel natijaga qo'shiladi, faqat kadastr raqami A da yo'q va A poligonlari uni
     B_QOPLANISH (0.1) dan kam qoplasa — ya'ni A da umuman yo'q yer. A bilan qoplangan B parsellari
     (qayta ro'yxatdan o'tgan) qo'shilmaydi — bir yer ikki egaga yozilmasligi uchun; ular va qisman
     qoplanganlar hisobotda.
  5. B dan qo'shilgan parselga INN / JShShIR: shu tumanda xuddi shu ega (nom yadrosi bo'yicha) A da
     bitta INN / JShShIR bilan uchrasa — o'sha olinadi (inn_manba = "nom (A)").
  6. Buzuq geometriyalar tuzatiladi; takror kadastr raqamlari belgilanadi (takror = 1).
  7. A va B fayllari o'zgarmaydi (fayl xeshi tekshiriladi). Natija yangi shp va Excel hisobot.

Ishga tushirish (repo papkasidan), avval pastdagi sozlamalarni to'g'rilang:
    .\\run.bat kadastr_birlashtirish.py
"""

from __future__ import annotations

# =====================================================================================
#  FOYDALANUVCHI SOZLAMALARI — har bir viloyat uchun shu yerni o'zgartiring
# =====================================================================================

# A — yangi uzkad kadastr shp (asos). O'zgarmaydi.
A_FAYL = r"D:\malumot\X_viloyati_kadastr_yangi.shp"

# B — eski, lekin to'liqroq kadastr shp. O'zgarmaydi.
B_FAYL = r"D:\malumot\X_viloyati_kadastr_eski.shp"

# Natija shp (yangi fayl).
CHIQISH = r"D:\natija\X_viloyati_kadastr_birlashgan.shp"

# Excel hisobot. None — CHIQISH yonida "<chiqish nomi>_hisobot.xlsx".
HISOBOT = None

# A ustunlari
A_KAD = "CADNUM"
A_ESKI_KAD = "OLDCADNUM"
A_FAMILIYA = "HOST_FNAME"     # jismoniy shaxsda familiya, boshqalarda to'liq nom
A_ISM = "HOST_NAME"
A_OTASI = "HOST_SNAME"
A_EGA_TURI = "HOST_PTYPE"
A_INN = "HOST_TIN"
A_JSHSHIR = "HOST_PIN"
A_VILOYAT = "REG_NAME"
A_TUMAN = "DIST_NAME"
A_TUMAN_KOD = "DIST_SOAT"
A_YER_FONDI = "FUND_TYPE"
A_HUQUQ = "LEGAL_TYPE"
A_MULK = "PROP_KIND"

# B ustunlari
B_KAD = "cadastral_"
B_EGASI = "full_name"
B_VILOYAT = "region_nam"
B_TUMAN = "district_n"
B_TUMAN_KOD = "soato_dist"
B_YER_FONDI = "land_fund1"

# Jismoniy shaxs turi (A_EGA_TURI da shu so'z bilan boshlansa — F.I.Sh birlashtiriladi).
JISMONIY = "O‘zR fuqarosi"

# B parsel natijaga qo'shiladi, agar A poligonlari uni shundan kam qoplasa (0..1).
B_QOPLANISH = 0.1

# JShShIR (PINFL) natijaga yozilsinmi.
JSHSHIR_YOZ = True

# CHIQISH (va hisobot) allaqachon bor bo'lsa ustiga yozilsinmi.
USTIGA_YOZ = False

# =====================================================================================

import sys  # noqa: E402
from pathlib import Path  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
import muhit  # noqa: F401,E402  (GDAL_DATA/PROJ va DLL yo'llari, numpy dan oldin)

import time  # noqa: E402
import warnings  # noqa: E402

import geopandas as gpd  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
import pyogrio  # noqa: E402
import shapely  # noqa: E402

from atribut_otkazish import crs_matn, excel_yoz, foiz, log, shp_ochir, shp_xesh  # noqa: E402
from id_boyicha_qoshish import bir_xil, matnga, nom_yadrosi  # noqa: E402

warnings.filterwarnings("ignore", category=RuntimeWarning, module="pyogrio")

# Natijadagi yagona ustunlar (shp: 10 belgigacha)
YANGI = ["kad_raqam", "eski_kad", "egasi", "egasi_turi", "inn", "jshshir", "inn_manba", "viloyat",
         "tuman", "tuman_kod", "yer_fondi", "huquq", "mulk_turi", "maydon_ga", "b_egasi", "ega_ozgar",
         "manba", "takror", "tekshir"]


def kalit(s: pd.Series) -> pd.Series:
    """Kadastr raqami kaliti: bo'shliqsiz, katta harf; bo'sh -> None."""
    t = s.astype(object).map(matnga)
    return t.str.upper().str.replace(r"\s+", "", regex=True).where(t.notna(), None)


def oqi(p: Path) -> gpd.GeoDataFrame:
    """Buzuq halqali poligonlarni ham o'qish (tuzatib; tuzatib bo'lmasa geometriya bo'sh)."""
    try:
        return pyogrio.read_dataframe(p, on_invalid="fix")
    except Exception:  # noqa: BLE001
        return pyogrio.read_dataframe(p, on_invalid="ignore")


def poligonga(g):
    """make_valid natijasidan faqat poligon qismlari (shp ga yoziladigan)."""
    if g is None or shapely.is_empty(g):
        return g
    if shapely.get_type_id(g) in (3, 6):
        return g
    qism = [q for q in shapely.get_parts(g) if shapely.get_type_id(q) in (3, 6)]
    if not qism:
        return None
    return shapely.union_all(qism) if len(qism) > 1 else qism[0]


def main() -> int:
    boshlanish = time.strftime("%Y-%m-%d %H:%M:%S")
    a_yol, b_yol, chiqish = Path(A_FAYL), Path(B_FAYL), Path(CHIQISH)
    hisobot = Path(HISOBOT) if HISOBOT else chiqish.with_name(f"{chiqish.stem}_hisobot.xlsx")
    for p in (a_yol, b_yol):
        if not p.exists():
            sys.exit(f"Fayl topilmadi: {p}")
    if chiqish.resolve() in {a_yol.resolve(), b_yol.resolve()}:
        sys.exit("CHIQISH A yoki B bilan bir xil bo'lmasin — natija alohida faylga yoziladi.")
    for p in (chiqish, hisobot):
        if p.exists() and not USTIGA_YOZ:
            sys.exit(f"Natija allaqachon bor: {p}\n  Ustiga yozish uchun USTIGA_YOZ = True qiling.")

    t0 = time.time()
    xesh_a, xesh_b = shp_xesh(a_yol), shp_xesh(b_yol)
    A = oqi(a_yol)
    B = oqi(b_yol)
    log(f"A: {a_yol}  ({len(A)} parsel, {A.crs})")
    log(f"B: {b_yol}  ({len(B)} parsel, {B.crs})")
    a_kerak = [A_KAD, A_ESKI_KAD, A_FAMILIYA, A_ISM, A_OTASI, A_EGA_TURI, A_INN, A_JSHSHIR, A_VILOYAT,
               A_TUMAN, A_TUMAN_KOD, A_YER_FONDI, A_HUQUQ, A_MULK]
    b_kerak = [B_KAD, B_EGASI, B_VILOYAT, B_TUMAN, B_TUMAN_KOD, B_YER_FONDI]
    for nom, df, kerak in (("A", A, a_kerak), ("B", B, b_kerak)):
        yoq = [c for c in kerak if c not in df.columns]
        if yoq:
            sys.exit(f"{nom} da ustun yo'q: {yoq}\n  Bor ustunlar: {list(df.columns)}")
    tuqnash = [c for c in YANGI if c.lower() in {x.lower() for x in A.columns}]
    if tuqnash:
        sys.exit(f"A da bu ustun nomlari allaqachon bor: {tuqnash}")
    if B.crs != A.crs:
        B = B.to_crs(A.crs)

    hisob_crs = A.estimate_utm_crs()
    a_geom_yoq = shapely.is_missing(A.geometry.values) | shapely.is_empty(A.geometry.values)
    A["_k"] = kalit(A[A_KAD]).values
    B["_k"] = kalit(B[B_KAD]).values
    b_idx = pd.Series(np.arange(len(B)), index=B["_k"].values)
    b_idx = b_idx[b_idx.index.notna() & ~b_idx.index.duplicated()]

    # ---- 1. A asos: yagona ustunlar -----------------------------------------------------
    nat = A.copy()
    fam = A[A_FAMILIYA].map(matnga)
    jismoniy = A[A_EGA_TURI].astype(str).str.startswith(JISMONIY).values
    fio = (fam.fillna("") + " " + A[A_ISM].map(matnga).fillna("") + " " + A[A_OTASI].map(matnga).fillna(""))
    fio = fio.str.replace(r"\s+", " ", regex=True).str.strip()
    egasi = np.where(jismoniy, fio.values, fam.values).astype(object)
    egasi = np.where(pd.Series(egasi).map(matnga).isna(), None, egasi)
    b_sh = b_idx.reindex(A["_k"].values)
    b_bor = b_sh.notna().values
    b_ega = np.full(len(A), None, dtype=object)
    b_ega[b_bor] = B[B_EGASI].map(matnga).values[b_sh.values[b_bor].astype(int)]
    ega_bosh = pd.Series(egasi).isna().values
    ega_b_dan = ega_bosh & b_bor & pd.Series(b_ega).notna().values
    egasi[ega_b_dan] = b_ega[ega_b_dan]
    ozgar = np.array([("ha" if (e and b and not bir_xil(e, b)) else ("" if not b else "yo'q"))
                      for e, b in zip(egasi, b_ega)], dtype=object)

    nat["kad_raqam"] = A[A_KAD].map(matnga).values
    nat["eski_kad"] = A[A_ESKI_KAD].map(matnga).values
    nat["egasi"] = egasi
    nat["egasi_turi"] = A[A_EGA_TURI].map(matnga).values
    nat["inn"] = A[A_INN].map(matnga).values
    nat["jshshir"] = A[A_JSHSHIR].map(matnga).values if JSHSHIR_YOZ else None
    nat["inn_manba"] = np.where(pd.Series(nat["inn"]).notna().values | pd.Series(nat["jshshir"]).notna().values,
                                "A", "")
    nat["viloyat"] = A[A_VILOYAT].map(matnga).values
    nat["tuman"] = A[A_TUMAN].map(matnga).values
    nat["tuman_kod"] = A[A_TUMAN_KOD].map(matnga).values
    nat["yer_fondi"] = A[A_YER_FONDI].map(matnga).values
    nat["huquq"] = A[A_HUQUQ].map(matnga).values
    nat["mulk_turi"] = A[A_MULK].map(matnga).values
    nat["b_egasi"] = b_ega
    nat["ega_ozgar"] = ozgar
    manba = np.full(len(A), "A", dtype=object)
    manba[ega_b_dan] = "A, egasi B"

    # A da geometriya yo'q -> B dan
    g = nat.geometry.values.copy()
    geom_b = a_geom_yoq & b_bor
    g[geom_b] = B.geometry.values[b_sh.values[geom_b].astype(int)]
    manba[geom_b] = np.where(ega_b_dan[geom_b], "A, egasi va geometriya B", "A, geometriya B")
    nat = nat.set_geometry(gpd.GeoSeries(g, crs=A.crs, index=nat.index))
    nat["manba"] = manba
    nat["tekshir"] = np.where(a_geom_yoq & ~b_bor, "geometriya yo'q", "")
    log(f"A asos: {len(A)} parsel; egasi B dan {int(ega_b_dan.sum())}; geometriya B dan {int(geom_b.sum())}; "
        f"geometriyasiz qoldi {int((a_geom_yoq & ~b_bor).sum())}")

    # ---- 2. B parsellarining A bilan qoplanishi ------------------------------------------
    ag = shapely.make_valid(gpd.GeoSeries(g, crs=A.crs).to_crs(hisob_crs).values)
    bg = shapely.make_valid(B.geometry.to_crs(hisob_crs).values)
    a_kalit = set(A["_k"].dropna())
    faqat_b = ~B["_k"].isin(a_kalit).values
    fb = np.flatnonzero(faqat_b)
    ok_a = ~shapely.is_missing(ag) & ~shapely.is_empty(ag)
    a_ok_idx = np.flatnonzero(ok_a)
    tr = shapely.STRtree(ag[a_ok_idx])
    i, j = tr.query(bg[fb], predicate="intersects")
    kes = shapely.area(shapely.intersection(bg[fb][i], ag[a_ok_idx][j]))
    juft = pd.DataFrame({"b": fb[i], "a": a_ok_idx[j], "kes": kes})
    juft = juft[juft.kes > 0]
    qop = juft.groupby("b").kes.sum().reindex(fb, fill_value=0).values / np.maximum(shapely.area(bg[fb]), 1e-9)
    qop = np.minimum(qop, 1.0)
    qosh = fb[qop < B_QOPLANISH]
    qisman = fb[(qop >= B_QOPLANISH) & (qop < 0.9)]
    qayta = fb[qop >= 0.9]
    eski_set = set(kalit(A[A_ESKI_KAD]).dropna())
    log(f"Faqat B: {len(fb)} parsel -> A da yo'q (qo'shiladi) {len(qosh)}; qisman qoplangan {len(qisman)}; "
        f"A bilan qoplangan (qayta ro'yxat) {len(qayta)}, shundan eski raqami A da "
        f"{int(B['_k'].iloc[qayta].isin(eski_set).sum())}")

    # ---- 3. B dan qo'shiladigan parsellar --------------------------------------------------
    Bq = B.iloc[qosh]
    q = gpd.GeoDataFrame({c: pd.Series([None] * len(Bq), dtype=object) for c in A.columns if c not in ("geometry", "_k")},
                         geometry=Bq.geometry.values, crs=A.crs)
    q["kad_raqam"] = Bq[B_KAD].map(matnga).values
    q["eski_kad"] = None
    q["egasi"] = Bq[B_EGASI].map(matnga).values
    q["egasi_turi"] = None
    q["inn"] = None
    q["jshshir"] = None
    q["inn_manba"] = ""
    q["viloyat"] = Bq[B_VILOYAT].map(matnga).values
    q["tuman"] = Bq[B_TUMAN].map(matnga).values
    q["tuman_kod"] = Bq[B_TUMAN_KOD].map(matnga).values
    q["yer_fondi"] = Bq[B_YER_FONDI].map(matnga).values
    q["huquq"] = None
    q["mulk_turi"] = None
    q["b_egasi"] = q["egasi"].values
    q["ega_ozgar"] = ""
    q["manba"] = "B (eski)"
    eski_raqam_a_da = Bq["_k"].isin(eski_set).values
    q["tekshir"] = np.where(eski_raqam_a_da,
                            "B dan qo'shildi: raqami A da eski raqam (OLDCADNUM) — takror bo'lishi mumkin",
                            "B dan qo'shildi: A da yo'q")

    # INN / JShShIR nom bo'yicha: shu tumanda shu ega A da bitta INN / JShShIR bilan bo'lsa
    lug = pd.DataFrame({"t": pd.Series(nat["tuman"]).map(nom_yadrosi).values,
                        "n": pd.Series(nat["egasi"]).map(nom_yadrosi).values,
                        "inn": nat["inn"].values, "pin": nat["jshshir"].values, "tur": nat["egasi_turi"].values})
    lug = lug[lug.n != ""]
    topildi_inn = 0
    for kol, ust in (("inn", "inn"), ("pin", "jshshir")):
        if ust == "jshshir" and not JSHSHIR_YOZ:
            continue
        g2 = lug[lug[kol].notna()].groupby(["t", "n"])[kol].agg(lambda s: set(s))
        bir = g2[g2.map(len) == 1].map(lambda s: next(iter(s)))
        kalitlar = list(zip(pd.Series(q["tuman"]).map(nom_yadrosi), pd.Series(q["egasi"]).map(nom_yadrosi)))
        v = [bir.get(k) for k in kalitlar]
        q[ust] = v
        topildi_inn += sum(x is not None for x in v)
    tur = lug.dropna(subset=["tur"]).groupby(["t", "n"]).tur.agg(lambda s: s.mode().iloc[0])
    kalitlar = list(zip(pd.Series(q["tuman"]).map(nom_yadrosi), pd.Series(q["egasi"]).map(nom_yadrosi)))
    q["egasi_turi"] = [tur.get(k) for k in kalitlar]
    inn_bor = pd.Series(q["inn"]).notna().values | pd.Series(q["jshshir"]).notna().values
    q["inn_manba"] = np.where(inn_bor, "nom (A)", "")
    log(f"B dan qo'shilgan {len(q)} parsel: INN / JShShIR nom bo'yicha topildi {int(inn_bor.sum())}")

    # ---- 4. birlashtirish, geometriya, takror, maydon ----------------------------------------
    nat = nat.drop(columns=["_k"])
    asl_ust = [c for c in A.columns if c not in ("geometry", "_k")]
    ust = [c for c in YANGI if c not in ("maydon_ga", "takror")]  # bu ikkisi pastda hisoblanadi
    nat = pd.concat([nat[ust + asl_ust + ["geometry"]], q[ust + asl_ust + ["geometry"]]], ignore_index=True)
    nat = gpd.GeoDataFrame(nat, geometry="geometry", crs=A.crs)
    gv = nat.geometry.values
    notogri = ~shapely.is_valid(gv) & ~shapely.is_missing(gv)
    if notogri.any():
        gv = gv.copy()
        gv[notogri] = [poligonga(x) for x in shapely.make_valid(gv[notogri])]
        nat = nat.set_geometry(gpd.GeoSeries(gv, crs=A.crs, index=nat.index))
    k_all = kalit(nat["kad_raqam"])
    nat["takror"] = (k_all.notna() & k_all.duplicated(keep=False)).astype(int).values
    nat["maydon_ga"] = np.round(shapely.area(shapely.make_valid(nat.geometry.to_crs(hisob_crs).values)) / 1e4, 4)
    nat = nat[YANGI + asl_ust + ["geometry"]]
    log(f"Natija: {len(nat)} parsel; buzuq geometriya tuzatildi {int(notogri.sum())}; "
        f"takror kadastr qatorlari {int(nat['takror'].sum())} ({time.time() - t0:.0f} s)")

    # ---- yozish va tekshiruv -----------------------------------------------------------
    chiqish.parent.mkdir(parents=True, exist_ok=True)
    shp_ochir(chiqish)
    pyogrio.write_dataframe(nat, chiqish, driver="ESRI Shapefile", encoding="UTF-8")
    log(f"Yozildi: {chiqish}")
    a_ok = shp_xesh(a_yol) == xesh_a
    b_ok = shp_xesh(b_yol) == xesh_b
    qayta_o = pyogrio.read_dataframe(chiqish, read_geometry=False, columns=["kad_raqam", "manba"])
    soni_ok = len(qayta_o) == len(nat) == len(A) + len(q)
    for ok, matn in [(a_ok, "A fayli o'zgargan (boshqa dasturda tahrir qilinyaptimi?)"),
                     (b_ok, "B fayli o'zgargan (boshqa dasturda tahrir qilinyaptimi?)"),
                     (soni_ok, "natijadagi parsel soni kutilganidan farq qiladi")]:
        if not ok:
            log(f"XATO: {matn}")
    xato = not (a_ok and b_ok and soni_ok)
    if not xato:
        log(f"Tekshiruv: A va B o'zgarmagan; natijada A ning {len(A)} + B dan {len(q)} = {len(nat)} parsel.")

    # ---- Excel hisobot -------------------------------------------------------------------
    def ha(ok: bool) -> str:
        return "Ha ✔" if ok else "YO'Q ✘"

    ga = nat["maydon_ga"].values
    m_a = nat["manba"].astype(str).str.startswith("A").values
    m_b = ~m_a
    b_ga = shapely.area(bg) / 1e4
    umumiy = [
        ("Sana", boshlanish, ""),
        ("— FAYLLAR —", "", ""),
        ("A (yangi uzkad, asos)", str(a_yol), "o'zgarmadi, faqat o'qildi"),
        ("B (eski)", str(b_yol), "o'zgarmadi, faqat o'qildi"),
        ("CHIQISH", str(chiqish), "yangi fayl"),
        ("CRS (natija)", crs_matn(nat.crs), "A bilan bir xil"),
        ("Maydon qaysi CRS da", crs_matn(hisob_crs), "maydon_ga"),
        ("— KIRISH —", "", ""),
        ("A: parsellar", len(A), ""),
        ("A: geometriyasi yo'q", int(a_geom_yoq.sum()), ""),
        ("B: parsellar", len(B), ""),
        ("Kadastr raqami ikkalasida bor", int(len(a_kalit & set(B["_k"].dropna()))), ""),
        ("Faqat A da", int(len(a_kalit - set(B["_k"].dropna()))), "yangi yoki qayta ro'yxatdan o'tgan parsellar"),
        ("Faqat B da", len(fb), ""),
        ("   A bilan qoplangan (qayta ro'yxat) — qo'shilmadi", len(qayta), "'Qayta ro'yxat' varag'ida"),
        ("   qisman qoplangan — qo'shilmadi", len(qisman), "'Ko'rib chiqish' varag'ida"),
        ("   A da yo'q — qo'shildi", len(qosh), "'B dan qo'shildi' varag'ida"),
        ("      shundan raqami A da eski raqam (OLDCADNUM)", int(eski_raqam_a_da.sum()),
         "joyi siljigan bo'lishi mumkin — tekshir ustunida belgilangan"),
        ("— NATIJA —", "", ""),
        ("Parsellar: jami", len(nat), f"{round(float(np.nansum(ga)), 2)} ga"),
        ("   A dan", int(m_a.sum()), f"{round(float(np.nansum(ga[m_a])), 2)} ga"),
        ("   B dan qo'shilgan", int(m_b.sum()), f"{round(float(np.nansum(ga[m_b])), 2)} ga"),
        ("Egasi bor", int(pd.Series(nat["egasi"]).notna().sum()), f"{foiz(int(pd.Series(nat['egasi']).notna().sum()), len(nat)):.1%}"),
        ("   egasi B dan to'ldirildi (A da bo'sh edi)", int(ega_b_dan.sum()), ""),
        ("INN bor", int(pd.Series(nat["inn"]).notna().sum()), "yuridik shaxs, fermer xo'jaligi va h.k."),
        ("JShShIR bor", int(pd.Series(nat["jshshir"]).notna().sum()) if JSHSHIR_YOZ else "yozilmadi", "jismoniy shaxs"),
        ("INN / JShShIR nom bo'yicha topildi (B dan qo'shilganlarda)", int(inn_bor.sum()), "inn_manba = 'nom (A)'"),
        ("Egasi B dagidan farqli (ega_ozgar = ha)", int((nat["ega_ozgar"] == "ha").sum()),
         "kadastr raqami bir xil, nom boshqa — yangi egaga o'tgan bo'lishi mumkin"),
        ("Geometriya B dan olindi", int(geom_b.sum()), ""),
        ("Buzuq geometriya tuzatildi", int(notogri.sum()), ""),
        ("Takror kadastr raqamli qatorlar", int(nat["takror"].sum()), "takror = 1"),
        ("— TEKSHIRUV —", "", ""),
        ("A o'zgarmadi (fayl xeshi bir xil)", ha(a_ok), ""),
        ("B o'zgarmadi (fayl xeshi bir xil)", ha(b_ok), ""),
        ("Natijada parsel soni = A + B dan qo'shilgan", ha(soni_ok), ""),
    ]
    varaqlar = {"Umumiy": pd.DataFrame(umumiy, columns=["Ko'rsatkich", "Qiymat", "Izoh"])}

    t = pd.DataFrame({"Tuman": pd.Series(nat["tuman"]).fillna("(bo'sh)").values, "jami": 1, "ga": np.nan_to_num(ga),
                      "a": m_a, "b": m_b, "ega": pd.Series(nat["egasi"]).notna().values,
                      "inn": pd.Series(nat["inn"]).notna().values, "pin": pd.Series(nat["jshshir"]).notna().values,
                      "oz": (nat["ega_ozgar"] == "ha").values})
    gg = t.groupby("Tuman").sum(numeric_only=True)
    gg.loc["JAMI"] = gg.sum()
    varaqlar["Tumanlar"] = pd.DataFrame({
        "Tuman": gg.index, "Parsellar": gg.jami.values, "Maydoni, ga": gg.ga.round(2).values,
        "A dan": gg.a.values, "B dan qo'shilgan": gg.b.values, "Egasi bor": gg.ega.values,
        "INN bor": gg.inn.values, "JShShIR bor": gg.pin.values, "Egasi B dan farqli": gg.oz.values})

    et = pd.DataFrame({"Egasi turi": pd.Series(nat["egasi_turi"]).fillna("(noma'lum)").values, "jami": 1,
                       "ga": np.nan_to_num(ga), "inn": pd.Series(nat["inn"]).notna().values,
                       "pin": pd.Series(nat["jshshir"]).notna().values})
    eg = et.groupby("Egasi turi").sum(numeric_only=True).sort_values("jami", ascending=False)
    varaqlar["Ega turlari"] = pd.DataFrame({"Egasi turi": eg.index, "Parsellar": eg.jami.values,
                                            "Maydoni, ga": eg.ga.round(2).values, "INN bor": eg.inn.values,
                                            "JShShIR bor": eg.pin.values})

    mb = pd.Series(nat["manba"]).value_counts()
    varaqlar["Manbalar"] = pd.DataFrame({"manba": mb.index, "Parsellar": mb.values,
                                         "Maydoni, ga": [round(float(np.nansum(ga[(nat["manba"] == k).values])), 2) for k in mb.index]})

    def b_jadval(ix: np.ndarray, qop_v: np.ndarray) -> pd.DataFrame:
        d = pd.DataFrame({"B kadastr": B[B_KAD].values[ix], "B egasi": B[B_EGASI].values[ix],
                          "Tuman": B[B_TUMAN].values[ix], "Yer fondi": B[B_YER_FONDI].values[ix],
                          "Maydoni, ga": np.round(b_ga[ix], 4), "A bilan qoplangan": np.round(qop_v, 3)})
        return d

    qop_s = pd.Series(qop, index=fb)
    if len(qayta):
        d = b_jadval(qayta, qop_s.reindex(qayta).values)
        ja = juft[juft.b.isin(qayta)]
        ja = ja.assign(r=ja.groupby("b").kes.rank(ascending=False, method="first")).sort_values(["b", "r"])
        a_kad = A[A_KAD].values
        lst = ja.groupby("b").a.agg(lambda s: ", ".join(str(a_kad[x]) for x in list(s)[:10]))
        d["A kadastr (yangi)"] = lst.reindex(qayta).values
        d["Eski raqami A da (OLDCADNUM)"] = ["ha" if k in eski_set else "" for k in B["_k"].values[qayta]]
        varaqlar["Qayta ro'yxat"] = d
    if len(qisman):
        varaqlar["Ko'rib chiqish"] = b_jadval(qisman, qop_s.reindex(qisman).values)
    if len(q):
        d = b_jadval(qosh, qop_s.reindex(qosh).values)
        d["INN topildi"] = q["inn"].values
        d["JShShIR topildi"] = q["jshshir"].values
        d["tekshir"] = q["tekshir"].values
        varaqlar["B dan qo'shildi"] = d
    oz = np.flatnonzero((nat["ega_ozgar"] == "ha").values)
    if len(oz):
        varaqlar["Egasi farqli"] = pd.DataFrame({
            "kad_raqam": nat["kad_raqam"].values[oz], "tuman": nat["tuman"].values[oz],
            "Egasi (A, yangi)": nat["egasi"].values[oz], "Egasi (B, eski)": nat["b_egasi"].values[oz],
            "egasi_turi": nat["egasi_turi"].values[oz], "inn": nat["inn"].values[oz]})
    tk = np.flatnonzero(nat["takror"].values == 1)
    if len(tk):
        varaqlar["Takrorlar"] = pd.DataFrame({
            "kad_raqam": nat["kad_raqam"].values[tk], "manba": nat["manba"].values[tk],
            "egasi": nat["egasi"].values[tk], "tuman": nat["tuman"].values[tk],
            "maydon_ga": nat["maydon_ga"].values[tk]}).sort_values("kad_raqam")

    try:
        excel_yoz(hisobot, varaqlar, set())
        log(f"Hisobot: {hisobot}")
    except PermissionError:
        log(f"XATO: hisobotni yozib bo'lmadi (Excel da ochiq bo'lishi mumkin): {hisobot}")
        xato = True
    return 1 if xato else 0


if __name__ == "__main__":
    sys.exit(main())
