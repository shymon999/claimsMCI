import streamlit as st
import pandas as pd
from datetime import date
import io

# Konfiguracja strony Streamlit
st.set_page_config(page_title="Procesor Roszczeń (Claims)", layout="wide")

# ── Stałe ─────────────────────────────────────────────────────────────────────
TODAY           = date.today().strftime("%d.%m.%Y")
ASSIGNED_CHC    = "005Ts000006oQdNIAU"
ASSIGNED_NL     = "005Vk00000HBVrVIAX"
TEAM_SCHENKER   = "Claims Schenker Legacy"
TEAM_CHC_DOC    = "CHC Doc Team"
TEAM_MCI        = "Claims MCI"
TEAM_PORTUGAL   = "Claims Portugal"
TEAM_CHC_GLOBAL = "CHC Global"

# ── Nazwy kolumn w pliku ───────────────────────────────────────────────────────
COL_COUNTRY    = "DSV Country (Lookup)"
COL_SHIPMENT   = "Shipment number"
COL_AMOUNT     = "Claim amount EUR"
COL_DATE_LOSS  = "Date of Loss"
COL_LEGACY     = "Schenker Legacy Claim"
COL_TEAM       = "Team Name"
COL_ASSIGNED   = "Assigned Name"
COL_INITIAL    = "Initial assignment"
COL_REQUESTER  = "Requester name"

# ── Pomocnicze ────────────────────────────────────────────────────────────────
def is_without_dash(val):
    if pd.isna(val) or str(val).strip() == "":
        return False
    return "-" not in str(val)

def has_value(val):
    return pd.notna(val) and str(val).strip() not in ("", "nan")

def get_year(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    s = str(val).strip()
    parts = s.split(".")
    if len(parts) == 3:
        try:
            return int(parts[2])
        except ValueError:
            pass
    try:
        return pd.to_datetime(s, dayfirst=True).year
    except Exception:
        return None

def to_amount(val):
    try:
        return float(str(val).replace(",", ".").replace(" ", ""))
    except (ValueError, TypeError):
        return 0.0

def process_row(row, country):
    result = {}
    if has_value(row.get(COL_REQUESTER)):
        result[COL_TEAM] = TEAM_CHC_GLOBAL
        return result

    if country == "France":
        legacy = "Yes" if is_without_dash(row.get(COL_SHIPMENT)) else "No"
        result[COL_LEGACY] = legacy
        if legacy == "Yes":
            result[COL_TEAM]    = TEAM_SCHENKER
            result[COL_ASSIGNED] = ""
        else:
            result[COL_TEAM]    = TEAM_CHC_DOC
            result[COL_ASSIGNED] = ASSIGNED_CHC
            result[COL_INITIAL]  = TODAY

    elif country == "Netherlands":
        result[COL_LEGACY]   = "No"
        result[COL_TEAM]     = TEAM_CHC_DOC
        result[COL_ASSIGNED] = ASSIGNED_NL
        result[COL_INITIAL]  = TODAY

    elif country in ("Spain", "Portugal"):
        year = get_year(row.get(COL_DATE_LOSS))
        if year == 2025:
            legacy = "Yes"
        elif year == 2026:
            legacy = "No"
        else:
            legacy = ""
        result[COL_LEGACY] = legacy

        if legacy == "No":
            amount = to_amount(row.get(COL_AMOUNT, 0))
            if amount < 500:
                result[COL_TEAM]    = TEAM_PORTUGAL if country == "Portugal" else TEAM_MCI
                result[COL_ASSIGNED] = ""
            else:
                result[COL_TEAM]    = TEAM_CHC_DOC
                result[COL_ASSIGNED] = ASSIGNED_CHC
                result[COL_INITIAL]  = TODAY
        else:
            result[COL_TEAM]    = TEAM_CHC_DOC
            result[COL_ASSIGNED] = ASSIGNED_CHC
            result[COL_INITIAL]  = TODAY

    return result

# ── Interfejs Streamlit ───────────────────────────────────────────────────────
def main():
    st.title("📊 Procesor Plików Claims")
    st.markdown("Wgraj plik Excel, aby automatycznie przypisać zespoły, zmienić nagłówki i zaktualizować statusy.")

    # Widget do wgrywania plików
    uploaded_file = st.file_uploader("Wybierz plik Excel (.xlsx, .xls)", type=["xlsx", "xls"])

    if uploaded_file is not None:
        st.info("Przetwarzanie pliku...")
        
        try:
            all_sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype=str)
            processed = {}
            
            for sheet_name, df in all_sheets.items():
                
                # 1. Zmiana nagłówka kolumny
                if "Claim: Claim Number" in df.columns:
                    df.rename(columns={"Claim: Claim Number": "claim import id"}, inplace=True)
                
                # 2. Ustawienie statusu wewnętrznego (POPRAWKA WIELKOŚCI LITER)
                # Standaryzujemy nazwę, jeśli jakimś trafem była napisana małymi literami
                if "internal status" in df.columns:
                    df.rename(columns={"internal status": "Internal Status"}, inplace=True)
                
                # Tworzy lub nadpisuje istniejącą kolumnę "Internal Status" w każdym wierszu
                df["Internal Status"] = "Awaiting own process"

                # Stara logika przypisywania (wymaga kolumny Claim amount EUR!)
                if COL_COUNTRY in df.columns:
                    for col in [COL_LEGACY, COL_TEAM, COL_ASSIGNED, COL_INITIAL]:
                        if col not in df.columns:
                            df[col] = ""

                    for idx, row in df.iterrows():
                        country = str(row.get(COL_COUNTRY, "")).strip()
                        if country not in ("France", "Netherlands", "Spain", "Portugal"):
                            continue

                        changes = process_row(row, country)
                        for col, val in changes.items():
                            df.at[idx, col] = val
                
                # 3. Usunięcie zbędnych kolumn (wykonane NA SAMYM KOŃCU)
                cols_to_remove = ["Claim amount EUR", "Total liability EUR"]
                df.drop(columns=[c for c in cols_to_remove if c in df.columns], inplace=True, errors='ignore')
                
                processed[sheet_name] = df

            # Zapis do pamięci
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for sheet_name, df in processed.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            output.seek(0)
            st.success("✅ Plik został pomyślnie przetworzony!")
            
            # Przycisk do pobrania
            st.download_button(
                label="📥 Pobierz przetworzony plik Excel",
                data=output,
                file_name=f"Processed_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Wystąpił błąd podczas przetwarzania: {e}")

if __name__ == "__main__":
    main()
