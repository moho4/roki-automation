import pandas as pd
from datetime import datetime, date

# =============================
# KONFIGURACIJA
# =============================

INPUT_FILE = "izpitni_roki_25_26_prilagojen_v5.xlsx"  # prilagodi po potrebi
SHEET_NAME = "Profesorji_pripravljeno"

# izpitna obdobja 2025/26
ZIMSKO_OD = date(2026, 1, 19)
ZIMSKO_DO = date(2026, 2, 13)

POLETNO_OD = date(2026, 6, 1)
POLETNO_DO = date(2026, 7, 3)

JESENSKO_OD = date(2026, 8, 19)
JESENSKO_DO = date(2026, 9, 15)

# prazniki znotraj obdobij, kjer izpitov nočeš
PRAZNIKI = {
    date(2026, 2, 8),   # Prešernov dan (nedelja, ampak vseeno eksplicitno)
    date(2026, 6, 25),  # Dan državnosti
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
    # string: poskusi nekaj formatov
    s = str(value).strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%d. %m. %Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    # če ne gre, pusti None – to lahko kasneje posebej obravnavamo
    return None


def expand_predmeti_row(row):
    """
    Iz ene vrstice (kjer je 'Predmeti' = 'A. B. C') naredi več vrstic:
    A, B, C – vse z istim datumom, profesorjem itd.
    """
    predmeti_raw = row["Predmeti"]
    if pd.isna(predmeti_raw):
        return []

    # razbij po piki
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

# pričakujemo stolpce:
# 'Prof.', 'Predmeti', 'Datum', 'Ura', 'Predavalnica', 'Dejanska predavalnica', 'Opomba'

# kopija za delo
df = df.copy()

# normaliziraj datum v datetime.date
df["Datum_norm"] = df["Datum"].apply(normalize_date)

# dodaj stolpec 'Rok'
df["Rok"] = df["Datum_norm"].apply(get_rok)

# naredi "ploščato" tabelo: ena vrstica = en predmet + en datum
expanded_rows = []
for _, row in df.iterrows():
    expanded = expand_predmeti_row(row)
    expanded_rows.extend(expanded)

if not expanded_rows:
    raise SystemExit("Ni nobenih podatkov za obdelavo – preveri stolpec 'Predmeti'.")

df_norm = pd.DataFrame(expanded_rows)

# ohranimo samo relevantne stolpce, plus 'Predmet'
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
# 2. PREVERJANJA
# =============================

# 2.1 Vikendi
def is_weekend(d: date) -> bool:
    if d is None:
        return False
    # Monday = 0, Sunday = 6
    return d.weekday() >= 5


df_vikend = df_norm[df_norm["Datum"].apply(is_weekend)].copy()

# 2.2 Izven izpitnega obdobja
df_izven = df_norm[df_norm["Rok"].isna() & df_norm["Datum"].notna()].copy()

# 2.3 Prazniki
df_prazniki = df_norm[df_norm["Datum"].isin(PRAZNIKI)].copy()

# 2.4 10-dnevno pravilo po predmetu
rows_10dni = []
for predmet, grp in df_norm[df_norm["Datum"].notna()].groupby("Predmet"):
    grp_sorted = grp.sort_values("Datum")
    dates = grp_sorted["Datum"].tolist()
    for i in range(len(dates) - 1):
        d1 = dates[i]
        d2 = dates[i + 1]
        delta = (d2 - d1).days
        if delta < 10:
            # zapišemo obe vrstici, kjer je kršitev
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
# 3. ZAPIS POROČILA
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

print(f"Končano. Poročilo zapisano v: {output_file}")
