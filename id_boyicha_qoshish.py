"""Asl shp ga boshqa (yordamchi) fayldan unique_id bo'yicha kerakli ustunlarni qo'shish.

Odatda atribut_otkazish.py va qoshnidan_toldirish.py dan keyin ishlatiladi: masalan suv hajmi
ustunlari (uw3_m3 ... uw10_m3, uwt_m3) boshqa faylda turadi va ularni ID orqali olib kelish kerak.

Qoida:
  1. ASL — ma'lumot qabul qiladigan shp (o'zgarmaydi). YORDAMCHI — ustunlar olinadigan fayl
     (shp, gpkg, dbf, csv yoki xlsx; faqat o'qiladi).
  2. Bog'lash ID_ASL va ID_YORDAMCHI ustunlari bo'yicha. ID_TOZALASH = True bo'lsa IDlar
     solishtirishdan oldin tozalanadi: kichik harf, bo'sh joy, "-", "{", "}" olib tashlanadi
     (masalan "5E47-...-CA" va "5e47...ca" bir xil hisoblanadi). Fayllardagi IDlar o'zgarmaydi.
  3. USTUNLAR — yordamchidan olinadigan ustunlar: {yordamchidagi nomi: natijadagi nomi}.
     Natijadagi nom ASL da allaqachon bo'lsa — skript to'xtaydi (nomini o'zgartiring).
  4. Yordamchida bir ID bir necha marta uchrasa — birinchi qatori olinadi; qiymatlari farqli
     takrorlar hisobotda alohida ko'rsatiladi.
  5. Natija CHIQISH ga yoziladi: ASL ning hamma qatorlari, ustunlari va geometriyasi o'zgarmaydi,
     faqat yangi ustunlar qo'shiladi. ID topilmagan qatorlarda yangi ustunlar bo'sh qoladi.
  6. Hisobot Excel ga yoziladi (HISOBOT): nechta ID bor edi, nechtasi o'tdi, nechtasi o'tmadi.

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
YORDAMCHI = r"D:\malumot\X_viloyati_suv.shp"

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
    "uw3_m3": "uw3_m3",
    "uw4_m3": "uw4_m3",
    "uw5_m3": "uw5_m3",
    "uw6_m3": "uw6_m3",
    "uw7_m3": "uw7_m3",
    "uw8_m3": "uw8_m3",
    "uw9_m3": "uw9_m3",
    "uw10_m3": "uw10_m3",
    "uwt_m3": "uwt_m3",
}

# Yordamchi xlsx bo'lsa — varaq nomi (None — birinchi varaq). csv bo'lsa — ajratuvchi.
XLSX_VARAQ = None
CSV_AJRATUVCHI = ","

# CHIQISH (va hisobot) allaqachon bor bo'lsa ustiga yozilsinmi.
USTIGA_YOZ = False

# =====================================================================================

import sys  # noqa: E402
from pathlib import Path  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
import muhit  # noqa: F401,E402  (GDAL_DATA/PROJ va DLL yo'llari, numpy dan oldin)

import hashlib  # noqa: E402
import time  # noqa: E402

import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
import pyogrio  # noqa: E402
import shapely  # noqa: E402

from atribut_otkazish import crs_matn, excel_yoz, foiz, log, shp_ochir, shp_xesh  # noqa: E402

GEO_KENG = {".shp", ".gpkg", ".dbf", ".geojson", ".json"}


def fayl_xesh(p: Path) -> dict[str, str]:
    """shp bo'lsa yon fayllari bilan, aks holda faylning o'zi."""
    if p.suffix.lower() == ".shp":
        return shp_xesh(p)
    return {p.name: hashlib.sha256(p.read_bytes()).hexdigest()}


def id_tayyorla(s: pd.Series) -> pd.Series:
    """Solishtirish uchun ID: matn, bo'sh joysiz; ID_TOZALASH bo'lsa kichik harf, -{} siz.
    Bo'sh / NaN — None."""
    t = s.astype(object).where(s.notna(), None)
    t = t.map(lambda v: None if v is None else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v)))
    t = t.str.strip()
    if ID_TOZALASH:
        t = t.str.lower().str.replace(r"[\s{}\-]", "", regex=True)
    return t.where(t.notna() & (t != ""), None)


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
    if isinstance(USTUNLAR, dict):
        nomlar = dict(USTUNLAR)
    else:
        nomlar = {c: c for c in USTUNLAR}
    if not nomlar:
        sys.exit("USTUNLAR bo'sh — kamida bitta ustun kiriting.")
    uzun = [y for y in nomlar.values() if len(y) > 10]
    if uzun:
        sys.exit(f"shp ustun nomi 10 belgidan oshmasin: {uzun}")
    if len(set(v.lower() for v in nomlar.values())) != len(nomlar):
        sys.exit(f"Natijadagi ustun nomlari takrorlanadi: {list(nomlar.values())}")

    xesh_asl, xesh_yor = fayl_xesh(asl_yol), fayl_xesh(yor_yol)

    # ---- o'qish -------------------------------------------------------------------
    asl = pyogrio.read_dataframe(asl_yol)
    n = len(asl)
    log(f"ASL:       {asl_yol}  ({n} qator, {asl.crs})")
    if ID_ASL not in asl.columns:
        sys.exit(f"ASL da '{ID_ASL}' ustuni yo'q. Bor ustunlar: {list(asl.columns)}")
    tuqnash = [y for y in nomlar.values() if y.lower() in {c.lower() for c in asl.columns}]
    if tuqnash:
        sys.exit(f"Bu nomlar ASL da allaqachon bor: {tuqnash}\n  USTUNLAR da natijadagi nomini o'zgartiring.")
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
    topildi = a_id.notna().values & a_id.isin(ybir.index).values
    for eski, yangi in nomlar.items():
        natija[yangi] = qiy[eski].values
    asl_ust = [c for c in asl.columns if c != "geometry"]
    natija = natija[asl_ust + list(nomlar.values()) + ["geometry"]]

    a_bosh = a_id.isna().values
    a_uniq = pd.Series(a_id.dropna().unique())
    y_uniq = pd.Series(ybir.index)
    ortiqcha = y_uniq[~y_uniq.isin(a_uniq)]
    otdi_id = int(a_uniq.isin(ybir.index).sum())
    log(f"ID: ASL da {len(a_uniq)} ta noyob (qator {n}, bo'sh {int(a_bosh.sum())}); "
        f"YORDAMCHI da {len(y_uniq)} ta noyob (qator {len(yor)}, bo'sh {int(y_id.isna().sum())}, "
        f"takror ID {len(takror_id)}, shundan qiymati farqli {len(farqli)})")
    log(f"  o'tdi: {otdi_id} ID ({int(topildi.sum())} qator);  o'tmadi: {len(a_uniq) - otdi_id} ID "
        f"({int((~topildi & ~a_bosh).sum())} qator) + bo'sh ID {int(a_bosh.sum())} qator;  "
        f"yordamchida ortiqcha: {len(ortiqcha)} ID")

    ustun_rows = []
    for eski, yangi in nomlar.items():
        v = natija[yangi]
        bor = int(v.notna().sum())
        ustun_rows.append({
            "Yordamchidagi ustun": eski, "Natijadagi ustun": yangi, "Turi": str(yor[eski].dtype),
            "Qiymat yozildi (qator)": bor,
            "ID topildi, lekin qiymat bo'sh": int((topildi & v.isna().values).sum()),
            "Bo'sh (ID topilmadi)": int((~topildi).sum()),
            "To'ldi, %": foiz(bor, n),
        })
        log(f"  {eski} -> {yangi}: {bor} qator")

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
    eski_ok = soni_ok and all(
        qayta[c].astype(str).fillna("").equals(asl[c].astype(str).fillna("")) for c in asl_ust if c in qayta.columns
    ) and all(c in qayta.columns for c in asl_ust)
    for ok, matn in [(asl_ok, "ASL fayli o'zgargan!"), (yor_ok, "YORDAMCHI fayli o'zgargan!"),
                     (soni_ok, "qator soni farq qiladi"), (geom_ok, "geometriya farq qiladi"),
                     (crs_ok, "CRS farq qiladi"), (eski_ok, "ASL ustunlari o'zgargan")]:
        if not ok:
            log(f"XATO: {matn}")
    xato = not (asl_ok and yor_ok and soni_ok and geom_ok and crs_ok and eski_ok)
    if not xato:
        log("Tekshiruv: ASL va YORDAMCHI o'zgarmagan; qator soni, geometriya, CRS va ASL ustunlari bir xil.")

    # ---- Excel hisobot ------------------------------------------------------------------
    def ha(ok: bool) -> str:
        return "Ha ✔" if ok else "YO'Q ✘"

    may = asl["maydon"].astype(float).values if "maydon" in asl.columns else None

    def ga(m: np.ndarray) -> str:
        return "" if may is None else f"{np.nansum(may[m]):,.2f} ga".replace(",", " ")

    umumiy = pd.DataFrame([
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
        ("ID bo'sh qatorlar", int(a_bosh.sum()), "bog'lanmaydi"),
        ("— YORDAMCHI —", "", ""),
        ("Qatorlar", len(yor), ""),
        ("Noyob ID", len(y_uniq), ""),
        ("ID bo'sh qatorlar", int(y_id.isna().sum()), ""),
        ("Takrorlangan ID", len(takror_id), "birinchi qatori olindi"),
        ("   shundan qiymatlari farqli", len(farqli), "'Takrorlar' varag'ida"),
        ("— NATIJA —", "", ""),
        ("ID orqali o'tdi: noyob ID", otdi_id, f"ASL dagi noyob IDlarning {foiz(otdi_id, len(a_uniq)):.1%}"),
        ("ID orqali o'tdi: qator", int(topildi.sum()), ga(topildi)),
        ("O'tmadi: noyob ID (yordamchida yo'q)", len(a_uniq) - otdi_id, "'O'tmaganlar' varag'ida"),
        ("O'tmadi: qator (ID bo'shlari bilan)", int((~topildi).sum()), ga(~topildi)),
        ("Yordamchida bor, ASL da yo'q: noyob ID", len(ortiqcha), "'Ortiqcha IDlar' varag'ida"),
        ("Qo'shilgan ustunlar", ", ".join(f"{e} → {y}" if e != y else e for e, y in nomlar.items()), ""),
        ("— TEKSHIRUV —", "", ""),
        ("ASL o'zgarmadi (fayl xeshi bir xil)", ha(asl_ok), ""),
        ("YORDAMCHI o'zgarmadi (fayl xeshi bir xil)", ha(yor_ok), ""),
        ("Qator soni bir xil", ha(soni_ok), ""),
        ("Geometriya bir xil", ha(geom_ok), ""),
        ("CRS bir xil", ha(crs_ok), ""),
        ("ASL ustunlari o'zgarmadi", ha(eski_ok), ""),
    ], columns=["Ko'rsatkich", "Qiymat", "Izoh"])

    varaqlar = {"Umumiy": umumiy, "Ustunlar": pd.DataFrame(ustun_rows)}

    if "tuman" in asl.columns:
        t = pd.DataFrame({"Tuman": asl["tuman"].fillna("(bo'sh)").values, "jami": 1, "otdi": topildi})
        g = t.groupby("Tuman").sum(numeric_only=True)
        g.loc["JAMI"] = g.sum()
        varaqlar["Tumanlar"] = pd.DataFrame({
            "Tuman": g.index, "Qatorlar": g.jami.values, "ID orqali o'tdi": g.otdi.values,
            "O'tmadi": (g.jami - g.otdi).values, "O'tdi, %": (g.otdi / g.jami).round(4).values,
        })

    ix = np.flatnonzero(~topildi)
    otm = pd.DataFrame({"FID (qator, 0 dan)": ix, ID_ASL: asl[ID_ASL].values[ix]})
    for c in ("tuman", "maydon", "crop"):
        if c in asl.columns:
            otm[c] = asl[c].values[ix]
    otm["Sabab"] = np.where(a_bosh[ix], "ASL da ID bo'sh", "ID yordamchida yo'q")
    varaqlar["O'tmaganlar"] = otm

    yor_asl_id = yor[ID_YORDAMCHI].astype(str).values
    ort = pd.DataFrame({"Tozalangan ID": ortiqcha.values})
    ilk = pd.Series(yor_asl_id, index=y_id.values)
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
