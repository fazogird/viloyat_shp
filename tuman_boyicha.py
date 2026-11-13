"""Ekin yerlariga tuman chegaralaridan viloyat va tuman kodlarini berish (fazoviy).

Qoida:
  1. ASL — ekin yerlar shp (o'zgarmaydi). CHEGARA — tuman chegaralari shp (faqat o'qiladi).
  2. Har bir dala uchun u eng ko'p maydon bilan yotgan tuman topiladi (hisob UTM da, metrda).
     Dala ikki tuman chegarasida bo'lsa — kattaroq qismi qaysi tumanda bo'lsa, o'sha olinadi.
  3. USTUNLAR — chegaradan olinadigan ustunlar: {chegaradagi nomi: natijadagi nomi},
     masalan parent_cod -> region_id, code -> district_id.
  4. Nom tekshiruvi (NOM_TEKSHIRUV): ASL dagi viloyat / tuman nomlari chegaradagi region_nam /
     district bilan solishtiriladi (apostrof turi va katta-kichik harf hisobga olinmaydi):
       mos           — nomlar bir xil;
       qo'shimchasiz — farq faqat "tumani" / "shahri" so'zida (masalan "Guliston" va
                        "Guliston shahri"); "Guliston tumani" va "Guliston shahri" esa farqli;
       farqli        — boshqa tuman.
     NOMLARNI_YANGILASH = True bo'lsa natijada viloyat / tuman chegaradagi nomlar bilan almashadi.
  5. Natija CHIQISH ga yoziladi: ASL ning qatorlari, geometriyasi, CRS i va boshqa ustunlari
     o'zgarmaydi. Hech qaysi tumanga tegmagan dalada yangi ustunlar bo'sh qoladi.
  6. Hisobot Excel ga yoziladi (HISOBOT).

Ishga tushirish (repo papkasidan), avval pastdagi sozlamalarni to'g'rilang:
    .\\run.bat tuman_boyicha.py
"""

from __future__ import annotations

# =====================================================================================
#  FOYDALANUVCHI SOZLAMALARI — har bir viloyat uchun shu yerni o'zgartiring
# =====================================================================================

# Ekin yerlar shp (masalan id_boyicha_qoshish.py yoki qoshnidan_toldirish.py natijasi). O'zgarmaydi.
ASL = r"D:\natija\X_viloyati_v4.shp"

# Tuman chegaralari shp. O'zgarmaydi (faqat o'qiladi).
CHEGARA = r"D:\malumot\tuman_chegaralari.shp"

# Natija shp (yangi fayl; ASL yoki CHEGARA bilan bir xil bo'lishi mumkin emas).
CHIQISH = r"D:\natija\X_viloyati_v5.shp"

# Excel hisobot. None — CHIQISH yonida "<chiqish nomi>_hisobot.xlsx".
HISOBOT = None

# Chegaradan olinadigan ustunlar: {chegaradagi nomi: natijadagi nomi} (shp: 10 belgigacha,
# shuning uchun "district_id" (11) emas, "distrct_id").
USTUNLAR = {
    "parent_cod": "region_id",
    "code": "distrct_id",
}

# Nom tekshiruvi: {ASL dagi ustun: chegaradagi ustun}. Tekshirmaslik uchun {} qiling.
NOM_TEKSHIRUV = {
    "viloyat": "region_nam",
    "tuman": "district",
}

# True — natijada NOM_TEKSHIRUV dagi ASL ustunlari chegaradagi nomlar bilan almashtiriladi.
NOMLARNI_YANGILASH = False

# Dala tumanga qancha qismi bilan tushgani (0..1) uchun ustun nomi; kerak bo'lmasa None.
ULUSH_USTUNI = None

# CHIQISH (va hisobot) allaqachon bor bo'lsa ustiga yozilsinmi.
USTIGA_YOZ = False

# =====================================================================================

import sys  # noqa: E402
from pathlib import Path  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
import muhit  # noqa: F401,E402  (GDAL_DATA/PROJ va DLL yo'llari, numpy dan oldin)

import re  # noqa: E402
import time  # noqa: E402

import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
import pyogrio  # noqa: E402
import shapely  # noqa: E402

from atribut_otkazish import crs_matn, excel_yoz, foiz, log, shp_ochir, shp_xesh  # noqa: E402

QOSHIMCHA = re.compile(r"\s+(tumani|shahri|viloyati|tuman|shahar|viloyat)$")
APOSTROF = re.compile(r"[\u2018\u2019\u02bb\u02bc`'\u00b4]")


SINONIM = {"shahar": "shahri", "tuman": "tumani", "viloyat": "viloyati"}


def nom_toza(s) -> str:
    """Solishtirish uchun: kichik harf, apostrof turlari olib tashlangan, ortiqcha bo'shliqsiz;
    oxiridagi "shahar" / "tuman" / "viloyat" -> "shahri" / "tumani" / "viloyati"."""
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return ""
    t = re.sub(r"\s+", " ", APOSTROF.sub("", str(s))).strip().lower()
    soz = t.rsplit(" ", 1)
    if len(soz) == 2 and soz[1] in SINONIM:
        t = f"{soz[0]} {SINONIM[soz[1]]}"
    return t


def nom_holati(asl_nom, chegara_nom) -> str:
    a, c = nom_toza(asl_nom), nom_toza(chegara_nom)
    if not a:
        return "ASL da bo'sh"
    if a == c:
        return "mos"
    a0, c0 = QOSHIMCHA.sub("", a), QOSHIMCHA.sub("", c)
    # faqat bittasida qo'shimcha bo'lsa (Guliston ~ Guliston shahri); ikkalasida har xil bo'lsa — farqli
    if a0 == c0 and (a0 == a or c0 == c):
        return "qo'shimchasiz mos"
    return "farqli"


def main() -> int:
    boshlanish = time.strftime("%Y-%m-%d %H:%M:%S")
    asl_yol, ch_yol, chiqish = Path(ASL), Path(CHEGARA), Path(CHIQISH)
    hisobot = Path(HISOBOT) if HISOBOT else chiqish.with_name(f"{chiqish.stem}_hisobot.xlsx")
    for p in (asl_yol, ch_yol):
        if not p.exists():
            sys.exit(f"Fayl topilmadi: {p}")
    if chiqish.resolve() in {asl_yol.resolve(), ch_yol.resolve()}:
        sys.exit("CHIQISH ASL yoki CHEGARA bilan bir xil bo'lmasin — natija alohida faylga yoziladi.")
    for p in (chiqish, hisobot):
        if p.exists() and not USTIGA_YOZ:
            sys.exit(f"Natija allaqachon bor: {p}\n  Ustiga yozish uchun USTIGA_YOZ = True qiling.")
    yangi = list(USTUNLAR.values()) + ([ULUSH_USTUNI] if ULUSH_USTUNI else [])
    uzun = [c for c in yangi if len(c) > 10]
    if uzun:
        sys.exit(f"shp ustun nomi 10 belgidan oshmasin: {uzun}")

    xesh_asl, xesh_ch = shp_xesh(asl_yol), shp_xesh(ch_yol)

    # ---- o'qish -------------------------------------------------------------------
    ch_info = pyogrio.read_info(ch_yol)
    ch_bor = list(ch_info["fields"])
    kerak = list(dict.fromkeys(list(USTUNLAR) + list(NOM_TEKSHIRUV.values())))
    yoq = [c for c in kerak if c not in ch_bor]
    if yoq:
        sys.exit(f"CHEGARA da ustun yo'q: {yoq}\n  Bor ustunlar: {ch_bor}")
    ch = pyogrio.read_dataframe(ch_yol, columns=kerak)
    asl = pyogrio.read_dataframe(asl_yol)
    n = len(asl)
    log(f"ASL:     {asl_yol}  ({n} dala, {asl.crs})")
    log(f"CHEGARA: {ch_yol}  ({len(ch)} tuman, {ch.crs})")
    if asl.crs is None or ch.crs is None:
        sys.exit("ASL yoki CHEGARA da CRS (.prj) yo'q.")
    yoq = [c for c in NOM_TEKSHIRUV if c not in asl.columns]
    if yoq:
        sys.exit(f"ASL da nom tekshiruvi ustuni yo'q: {yoq}\n  Bor ustunlar: {list(asl.columns)}")
    tuqnash = [c for c in yangi if c.lower() in {x.lower() for x in asl.columns}]
    if tuqnash:
        sys.exit(f"Bu nomlar ASL da allaqachon bor: {tuqnash}\n  USTUNLAR da natijadagi nomini o'zgartiring.")

    # ---- fazoviy: har dala uchun eng katta ulushli tuman -----------------------------
    t0 = time.time()
    hisob_crs = asl.estimate_utm_crs()
    ag = shapely.make_valid(asl.geometry.to_crs(hisob_crs).values)
    chg = ch.geometry.to_crs(hisob_crs).values
    notogri = int((~shapely.is_valid(chg)).sum())
    chg = shapely.make_valid(chg)
    ai, ci = shapely.STRtree(chg).query(ag, predicate="intersects")
    a_may = shapely.area(ag)
    ul = shapely.area(shapely.intersection(ag[ai], chg[ci])) / np.maximum(a_may[ai], 1e-9)
    j = pd.DataFrame({"a": ai, "c": ci, "u": ul})
    j = j[j.u > 0].sort_values(["a", "u", "c"], ascending=[True, False, True], kind="mergesort")
    soni = j.groupby("a").size()
    top = j.drop_duplicates("a").set_index("a")
    ikkinchi = j[j.groupby("a").cumcount() == 1].set_index("a")
    tumanli = np.zeros(n, bool)
    tumanli[top.index.values] = True
    c_idx = np.full(n, -1)
    c_idx[top.index.values] = top.c.values
    u_arr = np.zeros(n)
    u_arr[top.index.values] = top.u.values
    kop = np.zeros(n, bool)
    kop[soni.index.values[soni.values > 1]] = True
    log(f"Moslash ({hisob_crs.to_string()}, {time.time() - t0:.0f} s): tumanga tushdi {int(tumanli.sum())} / {n}; "
        f"2+ tumanga tegadi {int(kop.sum())}; eng katta ulush < 0.5: {int((tumanli & (u_arr < 0.5)).sum())}; "
        f"tumanga tegmadi {int((~tumanli).sum())}" + (f"; chegarada noto'g'ri geometriya {notogri} (tuzatildi)" if notogri else ""))

    # ---- natija -----------------------------------------------------------------------
    natija = asl.copy()
    for eski, nom in USTUNLAR.items():
        v = ch[eski].values[np.where(c_idx >= 0, c_idx, 0)]
        s = pd.Series(v, index=natija.index).where(tumanli)
        if pd.api.types.is_numeric_dtype(ch[eski]) and tumanli.all():
            s = s.astype(ch[eski].dtype)
        natija[nom] = s
        log(f"  {eski} -> {nom}: {int(s.notna().sum())} dala")
    if ULUSH_USTUNI:
        natija[ULUSH_USTUNI] = np.round(u_arr, 3)

    nom_hol = {}
    for a_ust, c_ust in NOM_TEKSHIRUV.items():
        cn = np.where(tumanli, ch[c_ust].values[np.where(c_idx >= 0, c_idx, 0)], None)
        hol = np.array([nom_holati(x, y) if t else "tumanga tegmadi"
                        for x, y, t in zip(asl[a_ust].values, cn, tumanli)], dtype=object)
        nom_hol[a_ust] = (c_ust, cn, hol)
        vc = pd.Series(hol).value_counts()
        log(f"  nom {a_ust} ~ {c_ust}: " + ", ".join(f"{k} {v}" for k, v in vc.items()))
        if NOMLARNI_YANGILASH:
            natija[a_ust] = np.where(tumanli, cn, natija[a_ust].values)

    asl_ust = [c for c in asl.columns if c != "geometry"]
    natija = natija[asl_ust + yangi + ["geometry"]]

    # ---- yozish va tekshiruv ------------------------------------------------------------
    chiqish.parent.mkdir(parents=True, exist_ok=True)
    shp_ochir(chiqish)
    pyogrio.write_dataframe(natija, chiqish, driver="ESRI Shapefile", encoding="UTF-8")
    log(f"Yozildi: {chiqish}")
    asl_ok = shp_xesh(asl_yol) == xesh_asl
    ch_ok = shp_xesh(ch_yol) == xesh_ch
    qayta = pyogrio.read_dataframe(chiqish)
    soni_ok = len(qayta) == n
    geom_ok = soni_ok and bool(shapely.equals(qayta.geometry.values, asl.geometry.values).all())
    crs_ok = qayta.crs == asl.crs
    tekshir_ust = [c for c in asl_ust if not (NOMLARNI_YANGILASH and c in NOM_TEKSHIRUV)]
    eski_ok = soni_ok and all(c in qayta.columns for c in asl_ust) and all(
        qayta[c].astype(str).equals(asl[c].astype(str)) for c in tekshir_ust)
    for ok, matn in [(asl_ok, "ASL fayli o'zgargan!"), (ch_ok, "CHEGARA fayli o'zgargan!"),
                     (soni_ok, "dala soni farq qiladi"), (geom_ok, "geometriya farq qiladi"),
                     (crs_ok, "CRS farq qiladi"), (eski_ok, "ASL ustunlari o'zgargan")]:
        if not ok:
            log(f"XATO: {matn}")
    xato = not (asl_ok and ch_ok and soni_ok and geom_ok and crs_ok and eski_ok)
    if not xato:
        log("Tekshiruv: ASL va CHEGARA o'zgarmagan; dala soni, geometriya, CRS va ASL ustunlari bir xil.")

    # ---- Excel hisobot ------------------------------------------------------------------
    def ha(ok: bool) -> str:
        return "Ha ✔" if ok else "YO'Q ✘"

    may = asl["maydon"].astype(float).values if "maydon" in asl.columns else a_may / 1e4
    may_izoh = "'maydon' ustunidan" if "maydon" in asl.columns else "geometriyadan"

    def ga(m: np.ndarray) -> float:
        return round(float(np.nansum(may[m])), 2)

    umumiy = [
        ("Sana", boshlanish, ""),
        ("— FAYLLAR —", "", ""),
        ("ASL (ekin yerlar)", str(asl_yol), "o'zgarmadi, faqat o'qildi"),
        ("CHEGARA (tumanlar)", str(ch_yol), "o'zgarmadi, faqat o'qildi"),
        ("CHIQISH (natija)", str(chiqish), "yangi fayl"),
        ("— CRS —", "", ""),
        ("ASL CRS i", crs_matn(asl.crs), ""),
        ("CHEGARA CRS i", crs_matn(ch.crs), ""),
        ("Moslash qaysi CRS da hisoblandi", crs_matn(hisob_crs), "maydon ulushi metrda"),
        ("CHIQISH CRS i", crs_matn(qayta.crs), "ASL bilan bir xil"),
        ("— QOIDA —", "", ""),
        ("Olingan ustunlar", ", ".join(f"{e} → {y}" for e, y in USTUNLAR.items()), ""),
        ("Tuman tanlash", "dala eng ko'p maydon bilan yotgan tuman", "ikki tuman chegarasida — kattaroq qismi"),
        ("Nomlar yangilandimi", "ha" if NOMLARNI_YANGILASH else "yo'q (faqat tekshirildi)", ""),
        ("— NATIJA —", "", ""),
        ("Dalalar: soni", n, ""),
        ("Dalalar: maydoni, ga", ga(np.ones(n, bool)), may_izoh),
        ("Tumanga tushdi: soni", int(tumanli.sum()), f"{foiz(int(tumanli.sum()), n):.1%}"),
        ("Tumanga tushdi: maydoni, ga", ga(tumanli), ""),
        ("   shundan 2+ tuman chegarasida", int(kop.sum()), "'Chegaradagi dalalar' varag'ida"),
        ("   shundan eng katta ulush < 0.5", int((tumanli & (u_arr < 0.5)).sum()), ""),
        ("Hech qaysi tumanga tegmadi: soni", int((~tumanli).sum()), "yangi ustunlar bo'sh; 'Tegmaganlar' varag'ida"),
        ("Hech qaysi tumanga tegmadi: maydoni, ga", ga(~tumanli), ""),
        ("Chegaradagi tumanlar soni (fayl)", len(ch), ""),
        ("Dalasi bor tumanlar soni", int(pd.Series(c_idx[c_idx >= 0]).nunique()), "'Tumanlar' varag'ida"),
    ]
    if notogri:
        umumiy.append(("Chegarada noto'g'ri geometriya", notogri, "hisob uchun tuzatildi (fayl o'zgarmadi)"))
    if nom_hol:
        umumiy.append(("— NOM TEKSHIRUVI —", "", ""))
        for a_ust, (c_ust, _, hol) in nom_hol.items():
            vc = pd.Series(hol).value_counts()
            for k in ["mos", "qo'shimchasiz mos", "farqli", "ASL da bo'sh", "tumanga tegmadi"]:
                if k in vc:
                    umumiy.append((f"{a_ust} ~ {c_ust}: {k}", int(vc[k]),
                                   "'Nom farqlari' varag'ida" if k in ("farqli", "qo'shimchasiz mos") else ""))
    umumiy += [
        ("— TEKSHIRUV —", "", ""),
        ("ASL o'zgarmadi (fayl xeshi bir xil)", ha(asl_ok), ""),
        ("CHEGARA o'zgarmadi (fayl xeshi bir xil)", ha(ch_ok), ""),
        ("Dala soni bir xil", ha(soni_ok), ""),
        ("Geometriya bir xil", ha(geom_ok), ""),
        ("CRS bir xil", ha(crs_ok), ""),
        ("ASL ustunlari o'zgarmadi", ha(eski_ok), "nomlar yangilangan bo'lsa, ular bundan tashqari"),
    ]
    varaqlar = {"Umumiy": pd.DataFrame(umumiy, columns=["Ko'rsatkich", "Qiymat", "Izoh"])}

    # Tumanlar: chegaradagi har tuman bo'yicha
    tt = pd.DataFrame({"c": c_idx[tumanli], "may": np.nan_to_num(may[tumanli]), "kop": kop[tumanli]})
    g = tt.groupby("c").agg(dala=("may", "size"), maydon=("may", "sum"), chegarada=("kop", "sum"))
    tum = ch.drop(columns="geometry").iloc[g.index.values].reset_index(drop=True)
    tum.insert(0, "Dalalar soni", g.dala.values)
    tum.insert(1, "Maydoni, ga", g.maydon.round(2).values)
    tum.insert(2, "Chegaradagi dalalar", g.chegarada.values)
    tum = tum.sort_values(list(USTUNLAR)).reset_index(drop=True)
    varaqlar["Tumanlar"] = tum

    if nom_hol:
        rows = []
        for a_ust, (c_ust, cn, hol) in nom_hol.items():
            m = np.isin(hol, ["farqli", "qo'shimchasiz mos", "ASL da bo'sh"])
            if m.any():
                d = pd.DataFrame({"ASL ustuni": a_ust, "ASL dagi nom": asl[a_ust].values[m],
                                  "Chegara ustuni": c_ust, "Chegaradagi nom": cn[m], "Holat": hol[m],
                                  "may": np.nan_to_num(may[m])})
                d = (d.groupby(["ASL ustuni", "ASL dagi nom", "Chegara ustuni", "Chegaradagi nom", "Holat"],
                               dropna=False).agg(**{"Dalalar soni": ("may", "size"), "Maydoni, ga": ("may", "sum")})
                     .reset_index())
                d["Maydoni, ga"] = d["Maydoni, ga"].round(2)
                rows.append(d)
        if rows:
            varaqlar["Nom farqlari"] = pd.concat(rows, ignore_index=True).sort_values(
                ["Holat", "Dalalar soni"], ascending=[True, False])

    def dala_jadval(ix: np.ndarray) -> pd.DataFrame:
        d = pd.DataFrame({"FID (qator, 0 dan)": ix})
        for c in ("unique_id", "viloyat", "tuman", "maydon", "crop"):
            if c in asl.columns:
                d[c] = asl[c].values[ix]
        return d

    ix = np.flatnonzero(kop)
    if len(ix):
        d = dala_jadval(ix)
        t1 = ch.drop(columns="geometry").iloc[c_idx[ix]].reset_index(drop=True)
        d["Olingan tuman"] = t1[NOM_TEKSHIRUV.get("tuman", list(USTUNLAR)[-1])].values \
            if "tuman" in NOM_TEKSHIRUV else t1[list(USTUNLAR)[-1]].values
        d["Ulushi"] = np.round(u_arr[ix], 3)
        i2 = ikkinchi.reindex(ix)
        ok2 = i2.c.notna().values
        ikk_nom = np.full(len(ix), "", dtype=object)
        if ok2.any():
            src = NOM_TEKSHIRUV.get("tuman", list(USTUNLAR)[-1])
            ikk_nom[ok2] = ch[src].values[i2.c.values[ok2].astype(int)]
        d["2-tuman"] = ikk_nom
        d["2-tuman ulushi"] = np.round(i2.u.fillna(0).values, 3)
        varaqlar["Chegaradagi dalalar"] = d
    ix = np.flatnonzero(~tumanli)
    if len(ix):
        varaqlar["Tegmaganlar"] = dala_jadval(ix)

    try:
        hisobot.parent.mkdir(parents=True, exist_ok=True)
        excel_yoz(hisobot, varaqlar, {"Ulushi", "2-tuman ulushi"})
        log(f"Hisobot: {hisobot}")
    except PermissionError:
        log(f"XATO: hisobotni yozib bo'lmadi (Excel da ochiq bo'lishi mumkin): {hisobot}")
        xato = True
    return 1 if xato else 0


if __name__ == "__main__":
    sys.exit(main())
