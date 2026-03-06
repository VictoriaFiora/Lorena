import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import requests
import time

# --- CONFIGURACIÓN ---
SPREADSHEET_ID = "1A_3qY1Vz6HkZxX_Your_Actual_ID_Here" # Reemplazar con el ID real
SHEET_SUCURSALES = "Tabla Sucursales Localizacion"
SHEET_MAESTRO = "Maestro Localidades - Corregido"
CREDS_FILE = "creds.json"

# --- LOGIN ---
def connect_to_sheet():
    """
    Intenta conectar usando creds.json (Service Account) 
    o mediante OAuth (abre navegador) si no existe el archivo.
    """
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    try:
        if os.path.exists(CREDS_FILE):
            print(f"Usando archivo de credenciales: {CREDS_FILE}")
            creds = Credentials.from_service_account_file(CREDS_FILE, scopes=scope)
            return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)
        else:
            print("No se encontró creds.json. Iniciando flujo OAuth (se abrirá el navegador)...")
            return gspread.oauth(scopes=scope).open_by_key(SPREADSHEET_ID)
    except Exception as e:
        print(f"Error de conexión: {e}")
        return None

# --- LÓGICA DE AUDITORÍA ---
def auditoria_sucursales(df):
    """
    Fase 1: Completar datos de localización en Sucursales.
    """
    print(f"Analizando {len(df)} sucursales...")
    # Aquí irá la lógica de Georef AR (requests)
    return df

if __name__ == "__main__":
    import os
    print("🚀 Proyecto Lorena iniciado.")
    
    if SPREADSHEET_ID == "1A_3qY1Vz6HkZxX_Your_Actual_ID_Here":
        print("⚠️  ERROR: Debes poner el ID de tu Google Sheet en la variable SPREADSHEET_ID")
    else:
        ss = connect_to_sheet()
        if ss:
            print(f"✅ Conectado a: {ss.title}")
            # Ejemplo: Leer sucursales
            ws_sucs = ss.worksheet(SHEET_SUCURSALES)
            df_sucs = pd.DataFrame(ws_sucs.get_all_records())
            
            df_cleaned = auditoria_sucursales(df_sucs)
            print("Proceso finalizado.")
