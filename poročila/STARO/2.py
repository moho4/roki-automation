import pandas as pd
from datetime import datetime, date

# =============================
# KONFIGURACIJA
# =============================

INPUT_FILE = "izpitni_roki_25_26_prilagojen_v5.xlsx"  # vnosni excel
SHEET_NAME = "Profesorji_pripravljeno"

# izpitna obdobja 2025/26
ZIMSKO_OD = date(2026, 1, 19)
ZIMSKO_DO = date(2026, 2, 13)

POLETNO_OD = date(2026, 6, 1)   # tvoj "spomladanski rok"
POLETNO_DO = date(2026, 7, 3)

JESENSKO_OD = date(2026, 8, 19)
JESENSKO_DO = date(2026, 9, 15)

# prazniki znotraj obdobij, kjer izpitov nočeš
PRAZNIKI = {
    date(2026, 2, 8),   # Prešernov dan
    date(2026, 6, 25),  # Dan državnosti
}

# =============================
# KONFLIKTNE SKUPINE – OBVEZNI
# =============================

# 1. stopnja – zimsko
KONFLIKTNE_SKUPINE_ZIMSKO = [
    # 1. letnik
    ["Antična filozofija 1", "Logika in argumentacija", "Novoveška filozofija 1", "Politična filozofija"],

    # 2. letnik E
    ["Srednjeveška in renesančna filozofija 1", "Filozofska antropologija", "Etika",
     "Azijske filozofije", "Slovenska filozofija in filozofska terminologija",
     "Marksizem in kritična teorija", "Filozofija in zgodovina znanosti"],

    # 2. letnik D – Filozofija splošno
    ["Srednjeveška in renesančna filozofija 1", "Filozofska antropologija",
     "Strukturalizem, psihoanaliza, filozofija"],

    # 2. letnik D – Kultura in etika
    ["Etika", "Moralna filozofija"],

    # 3. letnik E
    ["Fenomenologija 1", "Praktična filozofija", "Nemška klasična filozofija E",
     "Hermenevtika 2", "Filozofija narave", "Sodobna analitična filozofija", "Metafizika"],

    # 3. letnik D – Filozofija splošno
    ["Fenomenologija 1", "Nemška klasična filozofija D", "Filozofija zavesti in življenja"],

    # 3. letnik D – Kultura in etika
    ["Praktična filozofija", "Filozofija in humanistika", "Filozofija religije"],

    # 2. stopnja pedagoška – zimsko
    ["Kritična teorija družbe", "Didaktika filozofskih praks"],
    ["Praktična etika", "Didaktika filozofije in etike"],
]

# 1. stopnja – poletno (tvoj “spomladanski rok”)
KONFLIKTNE_SKUPINE_POLETNO = [
    # 1. letnik
    ["Estetika", "Antična filozofija 2", "Simbolna logika",
     "Spoznavna teorija", "Novoveška filozofija 2", "Ontologija"],

    # 2. letnik E
    ["Hermenevtika 1", "Osnove analitične filozofije", "Srednjeveška in renesančna filozofija 2",
     "Uvod v psihoanalizo", "Filozofija duha"],

    # 2. letnik D – Filozofija splošno
    ["Hermenevtika 1", "Osnove analitične filozofije", "Filozofija zgodovine"],

    # 2. letnik D – Kultura in etika
    ["Estetika", "Azijske filozofije, religije in kulture", "Socialna filozofija"],

    # 3. letnik E
    ["Normativna etika in teorija delovanja", "Fenomenologija 2",
     "Filozofija jezika", "Antropologija simbolnih form"],

    # 3. letnik D – Filozofija splošno
    ["Normativna etika in teorija delovanja", "Semiotika"],

    # 3. letnik D – Kultura in etika
    ["Praktična filozofija med Kantom in Heglom", "Človek in kozmos v renesansi"],
]

KONFLIKTNE_SKUPINE_PO_ROKU = {
    "zimsko": KONFLIKTNE_SKUPINE_ZIMSKO,
    "poletno": KONFLIKTNE_SKUPINE_POLETNO,
    # "jesensko": []  # trenutno nič, ker ga nisi navedel
}

# =============================
# POMOŽNE FUNKCIJE
# =============================

def get_rok(d: date) -> str | None:
    """Vrni ime izpitnega roka ali None, če datum ni v nobenem."""
    if d is None:
        return None
    if ZIMSKO_OD <= d <= ZIMSKO_DO:
        return "zimsko"
    if POLETNO_OD <= d <= POLETNO_DO:
        return "poletno"
    if JESENSKO_OD <= d <= JESENSKO_DO:
        return "jesensko"
    return None


def normalize_date(value):
    """Pretvori vrednost iz Excela v datetime.date ali None."""
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%d. %m. %Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def expand_predmeti_row(row):
    """
    Iz ene vrstice (kjer je 'Predmeti' = 'A. B. C') naredi več vrstic:
    A, B, C – vse z istim datumom, profesorjem itd.
    """
    predmeti_raw = row["Predmeti"]
    if pd.isna(predmeti_raw):
        return []

    parts = [p.strip() for p in str(predmeti_raw).split(".") if p.strip()]
    expanded = []
    for predmet in parts:
        new_row = row.copy()
        new_row["Predmet"] = predmet
        expanded.append(new_row)
    return expanded


# =============================
# 1. BRANJE IN NORMALIZACIJA
# =============================

df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

df = df.copy()
df["Datum_norm"] = df["Datum"].apply(normalize_date)
df["Rok"] = df["Datum_norm"].apply(get_rok)

expanded_rows = []
for _, row in df.iterrows():
    expanded = expand_predmeti_row(row)
    expanded_rows.extend(expanded)

if not expanded_rows:
    raise SystemExit("Ni nobenih podatkov za obdelavo – preveri stolpec 'Predmeti'.")

df_norm = pd.DataFrame(expanded_rows)

df_norm = df_norm[
    [
        "Prof.",
        "Predmet",
        "Datum_norm",
        "Rok",
        "Ura",
        "Predavalnica",
        "Dejanska predavalnica",
        "Opomba",
    ]
].rename(columns={"Prof.": "Prof", "Datum_norm": "Datum"})

# =============================
# 2. OSNOVNA PREVERJANJA
# =============================

def is_weekend(d: date) -> bool:
    if d is None:
        return False
    return d.weekday() >= 5  # 5 = sobota, 6 = nedelja

df_vikend = df_norm[df_norm["Datum"].apply(is_weekend)].copy()

df_izven = df_norm[df_norm["Rok"].isna() & df_norm["Datum"].notna()].copy()

df_prazniki = df_norm[df_norm["Datum"].isin(PRAZNIKI)].copy()

# 10-dnevno pravilo po predmetu (konzervativna, po imenu)
rows_10dni = []
for predmet, grp in df_norm[df_norm["Datum"].notna()].groupby("Predmet"):
    grp_sorted = grp.sort_values("Datum")
    dates = grp_sorted["Datum"].tolist()
    for i in range(len(dates) - 1):
        d1 = dates[i]
        d2 = dates[i + 1]
        delta = (d2 - d1).days
        if delta < 10:
            rows_10dni.append(
                {
                    "Predmet": predmet,
                    "Prof_1": grp_sorted.iloc[i]["Prof"],
                    "Datum_1": d1,
                    "Prof_2": grp_sorted.iloc[i + 1]["Prof"],
                    "Datum_2": d2,
                    "Razlika_dni": delta,
                }
            )

df_10dni = pd.DataFrame(rows_10dni)

# =============================
# 3. PREVERJANJE PREKRIVANJA OBVEZNIH PREDMETOV
# =============================

conflict_records = []
seen_pairs = set()  # za deduplikacijo

# gremo po rokih posebej (zimsko, poletno)
for rok, skupine in KONFLIKTNE_SKUPINE_PO_ROKU.items():
    df_r = df_norm[(df_norm["Rok"] == rok) & df_norm["Datum"].notna()].copy()
    if df_r.empty:
        continue

    # po datumih
    for datum, df_dan in df_r.groupby("Datum"):
        # slovar: predmet -> vrstice tistega dne
        # (v praksi večinoma ena vrstica, lahko pa tudi več)
        for idx_skupine, skupina in enumerate(skupine, start=1):
            # filtriraj vrstice, kjer je predmet v tej skupini
            df_conf = df_dan[df_dan["Predmet"].isin(skupina)].copy()
            if len(df_conf) < 2:
                continue  # nič spornega

            # vse pare v tej skupini za ta datum
            rows = df_conf.reset_index(drop=True)
            for i in range(len(rows)):
                for j in range(i + 1, len(rows)):
                    a = rows.iloc[i]
                    b = rows.iloc[j]

                    key = (rok, datum, a["Predmet"], b["Predmet"], a["Prof"], b["Prof"])
                    key_rev = (rok, datum, b["Predmet"], a["Predmet"], b["Prof"], a["Prof"])

                    if key in seen_pairs or key_rev in seen_pairs:
                        continue
                    seen_pairs.add(key)

                    conflict_records.append(
                        {
                            "Rok": rok,
                            "Datum": datum,
                            "Skupina_id": f"{rok}_{idx_skupine}",
                            "Predmet_A": a["Predmet"],
                            "Prof_A": a["Prof"],
                            "Ura_A": a["Ura"],
                            "Predavalnica_A": a["Predavalnica"],
                            "Predmet_B": b["Predmet"],
                            "Prof_B": b["Prof"],
                            "Ura_B": b["Ura"],
                            "Predavalnica_B": b["Predavalnica"],
                        }
                    )

df_prekrivanja = pd.DataFrame(conflict_records)

# =============================
# 4. ZAPIS POROČILA
# =============================

output_file = "porocilo_izpiti.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df_norm.to_excel(writer, sheet_name="IZPITI_NORMALIZIRANO", index=False)
    if not df_vikend.empty:
        df_vikend.to_excel(writer, sheet_name="NAPAKE_VIKEND", index=False)
    if not df_izven.empty:
        df_izven.to_excel(writer, sheet_name="NAPAKE_IZVEN_OBDOBJA", index=False)
    if not df_prazniki.empty:
        df_prazniki.to_excel(writer, sheet_name="NAPAKE_PRAZNIK", index=False)
    if not df_10dni.empty:
        df_10dni.to_excel(writer, sheet_name="NAPAKE_10_DNI", index=False)
    if not df_prekrivanja.empty:
        df_prekrivanja.to_excel(writer, sheet_name="NAPAKE_PREKRIVANJA", index=False)

print(f"Končano. Poročilo zapisano v: {output_file}")
