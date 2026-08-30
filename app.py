import streamlit as st
import pandas as pd
from datetime import date
import gspread
from google.oauth2.service_account import Credentials

# --- IMPOSTAZIONI DELLA PAGINA E CSS ---
st.set_page_config(page_title="Gestione Ore", layout="wide")
st.markdown("""
<style>
    header {visibility: hidden;}
    footer {visibility: hidden;}
    .block-container { padding: 1rem 3rem !important; }
    div[data-testid="stForm"] { padding: 1rem 1.5rem !important; }
    .st-emotion-cache-16idsys p { margin-bottom: 0px; }
</style>
""", unsafe_allow_html=True)

NOME_FOGLIO_GOOGLE = "Registro_ore_PES"

LISTA_JOB = ["0421088", "0804976", "8847761", "8847703", "8847704", "8847777", "8847801", "1655566", "8847683"]
LISTA_DESCRIZIONI = [
    "Active participation in ITO/OTR Hand Off meetings (General Scope of supply) and CLDR issue",
    "ECR Management, BoM creation", "Check and comments on bom draft structures",
    "Check and comments on First Plannning dates", "Selection of real codes for BoM substitutions",
    "Dummy-real substitutions", "Spare Parts Assessment and BoM issuing",
    "Action Items (DDR/FMEA) on IEP", "Flange to Flange Sign Off (FFSO) Case Management/Expediting",
    "Flange to Flange Sign Off (FFSO) for Upgrades & New Units: Case Resolution (no-GEV GTs)",
    "Flange to Flange Sign Off (FFSO) for CRE: Case Resolution (no-GEV GTs)",
    "Parts Evaluation Process (PEP): Case Management", "Parts Evaluation Process (PEP) Case Resolution (no-GEV GTs)",
    "Spare Parts Make FIR (no-GEV GTs)", "Spare Parts Make NOR (no-GEV GTs)",
    "Master BoM Revision (Frame 5)", "Master BoM Revision (Frame 3/2, PGT 5/2, PGT 10)",
    "Revision of costification Jobs 170AA** (Frame 5)", "Standard Data sheet revision",
    "OTRDR/FMEA metrics check and expediting", "Metrics results sharing (ppt with ITO/OTR, FFSO, etc…)",
    "IDM", "Service Request & Triage", "PCB support (data gathering, expediting, data upload in Clarity)",
    "Remaining Rebranding activities", "Management of Recurrent Issue on Other NON-GEV", "Formazione"
]
LISTA_PRODUCT = ["HDGT"]
LISTA_COMP = ["Costagliola S.", "Ermini R."]
COLONNE_DF = ["Data", "Mese_Anno", "JOB", "Alternative Job", "Requestor", "PRODUCT", "Comp", "Description", "Detail", "DOC", "REV", "HRS"]

def get_mese_anno(data_str):
    try:
        giorno, mese, anno = str(data_str).split('/')
        mesi_inglesi = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        return f"{mesi_inglesi[int(mese)-1]}-{anno[-2:]}"
    except:
        oggi = date.today()
        mesi_inglesi = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        return f"{mesi_inglesi[oggi.month-1]}-{str(oggi.year)[-2:]}"

# --- CONNESSIONE GOOGLE SHEETS ---
@st.cache_resource
def connetti_gsheets():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds_dict = st.secrets["gcp_service_account"]
    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(credentials)

def carica_dati():
    client = connetti_gsheets()
    sheet = client.open(NOME_FOGLIO_GOOGLE)
    worksheet = sheet.get_worksheet(0)
    
    try:
        dati = worksheet.get_all_records(head=1, numericise_ignore=['all'])
    except Exception:
        dati = worksheet.get_all_records(head=1)
    
    if not dati:
        df = pd.DataFrame(columns=COLONNE_DF)
    else:
        df = pd.DataFrame(dati)
        for col in COLONNE_DF:
            if col not in df.columns:
                df[col] = ""
        df = df[COLONNE_DF]
        
    df['Data'] = df['Data'].astype(str)
    df['JOB'] = df['JOB'].astype(str).str.strip()
    df.insert(0, "Seleziona", False)
    return df

def rigenera_report():
    """Rilegge i dati dal Foglio1, calcola l'aggregato e aggiorna il foglio Report."""
    try:
        client = connetti_gsheets()
        sheet = client.open(NOME_FOGLIO_GOOGLE)
        worksheet_1 = sheet.get_worksheet(0)
        
        # 1. Caricamento dati reali da Foglio1
        try:
            dati = worksheet_1.get_all_records(head=1, numericise_ignore=['all'])
        except Exception:
            dati = worksheet_1.get_all_records(head=1)
            
        if not dati:
            df_base = pd.DataFrame(columns=COLONNE_DF)
        else:
            df_base = pd.DataFrame(dati)
            for col in COLONNE_DF:
                if col not in df_base.columns:
                    df_base[col] = ""
            df_base = df_base[COLONNE_DF]

        # Aggiorna lo stato session_state di Streamlit
        df_memoria = df_base.copy()
        df_memoria['Data'] = df_memoria['Data'].astype(str)
        df_memoria['JOB'] = df_memoria['JOB'].astype(str).str.strip()
        df_memoria.insert(0, "Seleziona", False)
        st.session_state.dati = df_memoria

        # 2. Calcolo dati aggregati
        df_rep = df_base.copy()
        colonne_raggruppamento = ["Mese_Anno", "JOB", "Alternative Job", "Requestor", "PRODUCT", "Comp", "Description", "DOC", "REV"]
        
        for col in colonne_raggruppamento:
            df_rep[col] = df_rep[col].fillna("").astype(str).str.strip()
            
        df_rep["HRS"] = pd.to_numeric(df_rep["HRS"], errors='coerce').fillna(0)
        
        def unisci_dettagli(x):
            dettagli_validi = []
            for i in x:
                if pd.notna(i):
                    val = str(i).strip()
                    if val != "" and val not in dettagli_validi:
                        dettagli_validi.append(val)
            return " | ".join(dettagli_validi) if dettagli_validi else ""

        colonne_report = [col for col in COLONNE_DF if col != "Data"]
        
        if not df_rep.empty:
            df_report = df_rep.groupby(colonne_raggruppamento, as_index=False, dropna=False).agg({
                "Detail": unisci_dettagli,
                "HRS": "sum"
            })
            df_report = df_report[colonne_report]
        else:
            df_report = pd.DataFrame(columns=colonne_report)

        for col in df_report.columns:
            if col != "HRS":
                df_report[col] = df_report[col].fillna("").astype(str)
            else:
                df_report[col] = pd.to_numeric(df_report[col], errors='coerce').fillna(0)
            
        try:
            worksheet_rep = sheet.worksheet("Report")
        except gspread.exceptions.WorksheetNotFound:
            worksheet_rep = sheet.add_worksheet(title="Report", rows="1000", cols="20")
            
        worksheet_rep.clear()

        # Preparazione dati e inserimento partendo da "A1"
        valori = [df_report.columns.values.tolist()] + df_report.values.tolist()
        worksheet_rep.update("A1", valori, value_input_option="RAW")

        try:
            worksheet_rep.format("B:B", {"numberFormat": {"type": "TEXT"}})
        except Exception:
            pass

        return True, "Report aggiornato con successo!"
    except Exception as e:
        return False, f"Errore durante l'aggiornamento: {str(e)}"

def salva_dati(df):
    client = connetti_gsheets()
    sheet = client.open(NOME_FOGLIO_GOOGLE)
    
    # 1. SALVATAGGIO FOGLIO 1
    worksheet_1 = sheet.get_worksheet(0)
    if worksheet_1.title != "Foglio1":
        worksheet_1.update_title("Foglio1")
        
    df_base = df.drop(columns=["Seleziona"], errors='ignore').fillna("")
    df_to_save = df_base.copy()
    for col in df_to_save.columns:
        if col == "HRS":
            df_to_save[col] = pd.to_numeric(df_to_save[col], errors='coerce').fillna(0)
        else:
            df_to_save[col] = df_to_save[col].astype(str)
        
    worksheet_1.clear()
    valori_foglio1 = [df_to_save.columns.values.tolist()] + df_to_save.values.tolist()
    worksheet_1.update("A1", valori_foglio1, value_input_option="RAW")

    try:
        worksheet_1.format("C:C", {"numberFormat": {"type": "TEXT"}})
    except Exception:
        pass

    # 2. RIGENERA IL REPORT
    rigenera_report()

# --- INTERFACCIA UTENTE ---
if 'dati' not in st.session_state:
    st.session_state.dati = carica_dati()
        
if 'form_reset_key' not in st.session_state:
    st.session_state.form_reset_key = 0

oggi_str = date.today().strftime("%d/%m/%Y")
df_oggi = st.session_state.dati[st.session_state.dati["Data"] == oggi_str]
ore_gia_inserite = pd.to_numeric(df_oggi["HRS"], errors='coerce').fillna(0).sum()
ore_rimaste = 8 - ore_gia_inserite

col_data, col_btn_reset, col_ore = st.columns([2, 2, 2])
with col_data:
    st.markdown(f"### Data: {oggi_str}")
with col_btn_reset:
    if st.button("🔄 Resetta/Aggiorna Foglio Report", use_container_width=True):
        with st.spinner("Aggiornamento del Report in corso..."):
            esito, messaggio = rigenera_report()
        if esito:
            st.success(messaggio)
            st.rerun()
        else:
            st.error(messaggio)

with col_ore:
    st.markdown(f"<div style='text-align: right; margin-top: 10px; font-size: 18px;'><b>Ore registrate oggi: {int(ore_gia_inserite)} / 8</b></div>", unsafe_allow_html=True)

with st.form(f"form_inserimento_{st.session_state.form_reset_key}", clear_on_submit=False):
    col1, col2 = st.columns(2)
    with col1:
        job = st.selectbox("JOB", LISTA_JOB, index=None, placeholder="Seleziona un JOB...")
        alt_job = st.text_input("Alternative Job (Opzionale)")
        product = st.selectbox("PRODUCT (Opzionale)", LISTA_PRODUCT, index=None)
        comp = st.selectbox("Comp (Opzionale)", LISTA_COMP, index=None)
        
    with col2:
        descrizione = st.selectbox("Description", LISTA_DESCRIZIONI, index=None)
        dettaglio = st.text_input("Detail (Opzionale)")
        opzioni_ore = list(range(1, int(ore_rimaste) + 1)) if ore_rimaste > 0 else []
        hrs = st.selectbox("HRS", opzioni_ore, index=None)

    col_vuota, col_bottone = st.columns([5, 1])
    with col_bottone:
        submit = st.form_submit_button("INSERISCI RIGA", use_container_width=True)

    if submit:
        campi_mancanti = [c for c, v in zip(["JOB", "Description", "HRS"], [job, descrizione, hrs]) if not v]
        if campi_mancanti:
            st.error(f"⚠️ Compila i campi obbligatori: **{', '.join(campi_mancanti)}**")
        else:
            nuova_riga = pd.DataFrame([{
                "Seleziona": False, "Data": oggi_str, "Mese_Anno": get_mese_anno(oggi_str),
                "JOB": str(job), "Alternative Job": alt_job, "Requestor": "", 
                "PRODUCT": product if product else "", "Comp": comp if comp else "",
                "Description": descrizione, "Detail": dettaglio, "DOC": "", "REV": "", "HRS": hrs
            }])
            nuova_riga = nuova_riga[["Seleziona"] + COLONNE_DF]
            st.session_state.dati = pd.concat([st.session_state.dati, nuova_riga], ignore_index=True)
            salva_dati(st.session_state.dati)
            st.session_state.form_reset_key += 1
            st.rerun()

# --- TABELLA MODIFICA ORE ODIERNE ---
df_passato = st.session_state.dati[st.session_state.dati["Data"] != oggi_str].reset_index(drop=True)
df_oggi_edit = st.session_state.dati[st.session_state.dati["Data"] == oggi_str].reset_index(drop=True)

if not df_oggi_edit.empty:
    st.write("💡 *Modifica le celle cliccandoci sopra. Spunta 'Seleziona' per abilitare l'eliminazione.*")
    dati_modificati_oggi = st.data_editor(
        df_oggi_edit, use_container_width=True, hide_index=True, 
        column_config={
            "Mese_Anno": None, "Requestor": None, "DOC": None, "REV": None,
            "Seleziona": st.column_config.CheckboxColumn("Seleziona", default=False)
        }, key="editor_oggi"
    )

    if not df_oggi_edit.equals(dati_modificati_oggi):
        dati_modificati_oggi["Data"] = oggi_str
        dati_modificati_oggi["Mese_Anno"] = dati_modificati_oggi["Data"].apply(get_mese_anno)
        st.session_state.dati = pd.concat([df_passato, dati_modificati_oggi], ignore_index=True)
        salva_dati(st.session_state.dati)
        st.rerun()

    if dati_modificati_oggi["Seleziona"].any():
        col_space, col_btn = st.columns([5, 2])
        with col_btn:
            if st.button("🗑️ Elimina Righe Selezionate", use_container_width=True, type="primary"):
                righe_da_tenere = dati_modificati_oggi[dati_modificati_oggi["Seleziona"] == False]
                st.session_state.dati = pd.concat([df_passato, righe_da_tenere], ignore_index=True)
                salva_dati(st.session_state.dati)
                st.rerun()
else:
    st.markdown("<p style='text-align: center; color: gray;'>Nessun dato inserito oggi.</p>", unsafe_allow_html=True)
