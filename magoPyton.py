# -*- coding: utf-8 -*-
# automazione_copia_incolla.py

import openpyxl
import pyautogui
import time
import pyperclip # Per una gestione più affidabile degli appunti
import traceback # Per stampare errori dettagliati
from datetime import datetime # Per gestire le date lette da Excel
import sys # Per flushare l'output della percentuale
import os # Per ottenere il percorso assoluto del file
import keyboard # Per attendere input da tastiera e hotkeys
import subprocess # Per un riavvio più robusto dello script
import logging

# --- CONFIGURAZIONE LOGGING ---
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_file = 'automazione.log'

# Logger per file
file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
file_handler.setFormatter(log_formatter)
file_handler.setLevel(logging.INFO)

# Logger per console
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(log_formatter)
console_handler.setLevel(logging.INFO)

# Creazione del logger globale
logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Sovrascrive la funzione print per usare il logger
def print_and_log(messaggio=""):
    logger.info(messaggio)

# Rimpiazza il print globale
print = print_and_log


# --- NUOVO: Import per Screenshot e OCR ---
try:
    from PIL import Image
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    print("ATTENZIONE: Librerie 'Pillow' o 'pytesseract' non trovate.")
    print("Installa con: pip install Pillow pytesseract")
    OCR_AVAILABLE = False
# --- FINE NUOVO ---


# Import necessario per l'automazione di Excel
try:
    import win32com.client as win32
    WIN32_AVAILABLE = True
except ImportError:
    print("ATTENZIONE: Libreria 'pywin32' non trovata. Impossibile forzare il ricalcolo di Excel.")
    print("Installa con: pip install pywin32")
    WIN32_AVAILABLE = False

# --- CONFIGURAZIONE INIZIALE (MODIFICA QUESTI VALORI!) ---

# NUOVO: Flag per disabilitare la funzione di ricalcolo di Excel con pywin32,
# che a volte può essere instabile o causare crash.
# Impostare su True solo se si è sicuri che funzioni correttamente sul proprio sistema.
FORZARE_RICALCOLO_EXCEL = False

NOME_FILE_EXCEL_REL = r'C:\Users\Coemi\Desktop\PERFETTO\Dataease_ALLEGRETTI_02-2025.xlsm' # Esempio, modifica!
NOME_FILE_EXCEL = os.path.abspath(NOME_FILE_EXCEL_REL)

NOME_FOGLIO_DATI = 'DataEasePyton'
FOGLIO_PARAMETRI = 'ParametriPyton'
FOGLIO_SCARICO_ORE = 'SCARICO ORE'

COLONNA_STATO_PARAMETRI = 'D'
COLONNA_RIGA_INIZIO_DATI_PARAMETRI = 'B'
COLONNA_RIGA_FINE_DATI_PARAMETRI = 'C'
RIGA_INIZIO_PARAMETRI = 2
RIGA_FINE_PARAMETRI = 200
COLONNA_DATA_RAPPORTO_PARAMETRI = 'A'
STATO_DA_CERCARE = 'DA COMPLETARE'

COLONNA_PER_SCRIVERE_OK_DATI = 'A'
CELLA_CANTIERE_SCARICO_ORE = 'D1'

COORDINATE_DATA_RAPPORTO = (212, 172)
COORDINATE_CANTIERE = (280, 300)
COORDINATE_APERTURA_NUOVO_FOGLIO = (25, 88)
COORDINATE_SALVATAGGIO_FOGLIO = (85, 85)
COORDINATE_ELIMINA_FLAG_STAMPA = (777, 473)
COORDINATE_PULSANTE_OK = (846, 590)

# AGGIUNTA COLONNA T per poterla leggere durante la procedura ODC
COLONNE_EXCEL_DA_LEGGERE = ['C', 'D', 'I', 'J', 'K', 'L', 'M', 'N', 'P', 'Q', 'R', 'T']
# Le coordinate X per le colonne da incollare rimangono le stesse, 'T' non viene incollata direttamente.
TARGET_X_REMOTO_PER_COLONNA = [55, 138, 596, 730, 795, 860, 925, 985, 1110, 1175, 1260]

TARGET_Y_INIZIALE_REMOTO_ASSOLUTO = 375
INCREMENTO_Y_REMOTO = 19

MAX_RIGHE_DATI_PER_BLOCCO_TAB = 31
PRESSIONI_TASTO_TAB = 362
INTERVALLO_PRESSIONI_TAB = 0.01
PAUSA_DOPO_BLOCCO_TAB = 3.0

RITARDO_DOPO_MOVIMENTO_MOUSE = 0.2
RITARDO_TRA_CLICK_SINGOLO = 0.25
RITARDO_PRIMA_INCOLLA = 0.4
RITARDO_DOPO_COPIA_DA_EXCEL_PER_VERIFICA = 0.5
TENTATIVI_MAX_COPIA_DA_EXCEL = 3
RITARDO_TRA_TENTATIVI_COPIA_EXCEL = 0.6
DURATA_MOVIMENTO_MOUSE = 0.05
RITARDO_PRIMA_DI_COPIARE_FINALE_PRE_INCOLLA = 0.1
RITARDO_PRIMA_DI_COPIARE_IN_VERIFICA_EXCEL = 0.2
PAUSA_PRIMA_SALVATAGGIO_OPENPYXL = 0.5
PAUSA_TRA_AZIONI_PRELIMINARI_E_FINALI = 1.0
RITARDO_DOPO_SELECT_ALL = 0.2
RITARDO_DOPO_TAB = 0.15

TENTATIVI_MAX_SALVATAGGIO_EXCEL = 3
RITARDO_TRA_TENTATIVI_SALVATAGGIO = 2.0

# CORREZIONE: Reinserite le variabili mancanti
PULIZIA_APPUNTI_PAUSA_1 = 0.1
PULIZIA_APPUNTI_PAUSA_2 = 0.1
PULIZIA_APPUNTI_PAUSA_3 = 0.1
RICONFERMA_COPIA_PAUSA = 0.08

TASTO_PER_PROSEGUIRE = 'k'
TASTO_PER_RIAVVIARE = 'n'


# --- CONFIGURAZIONE OCR E PROCEDURA REGISTRAZIONE ODC ---

PATH_TESSERACT_CMD = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
REGIONE_SCREENSHOT_POPUP = (836, 459, 252, 173)
TESTO_DA_CERCARE_NEL_POPUP = "non trovato"
PAUSA_ATTESA_POPUP = 0.7

# --- NUOVE COORDINATE PER PROCEDURA ODC (DA COMPILARE) ---
COORDINATE_POPUP_ODC_TASTO_NO = (1042, 600)
COORDINATE_BARRA_TASTO_PREFERITI = (68, 34)
COORDINATE_BARRA_COMMESSE_DI_VARIANTE = (145, 128)
COORDINATE_CASELLA_COMMESSA = (59, 169)
COORDINATE_CASELLA_DESCRIZIONE = (334, 169)
COORDINATE_CASELLA_COMMESSA_PADRE = (535, 420)
COORDINATE_CASELLA_NS_RIF = (533, 522)
COORDINATE_APRI_ELENCO_GRUPPO = (616, 559)
COORDINATE_CASELLA_GRUPPO_GENERICO = (588, 592)
COORDINATE_CASELLA_RESP_COMMESSA = (682, 706)
COORDINATE_CASELLA_RESP_CANTIERE = (685, 730)
COORDINATE_BARRA_CAUSALI_DEPOSITI = (278, 243)
COORDINATE_APRI_ELENCO_DEPOSITO = (626, 587)
COORDINATE_CASELLA_DEPOSITO = (589, 619)
COORDINATE_CHIUDI_RAPPORTINO_COMMESSA = (333, 60)

# --- NUOVE COORDINATE E TESTI PER POPUP POST-COMMESSA PADRE ---
REGIONE_SCREENSHOT_POPUP_ATTENZIONE_SAL = (721, 443, 1204 - 721, 651 - 443)
# MODIFICATO: Usiamo una stringa più affidabile per il riconoscimento OCR
TESTO_POPUP_ATTENZIONE_SAL = "Nei SAL generati per questa commessa verranno"
COORDINATE_CLICK_POPUP_ATTENZIONE_SAL = (1147, 618)

REGIONE_SCREENSHOT_POPUP_CODICE_IVA = (727, 452, 1198 - 727, 637 - 452)
# MODIFICATO: Usiamo la stringa specifica richiesta
TESTO_POPUP_CODICE_IVA = "Il codice IVA del cliente viene impostato solo nella scheda"
COORDINATE_CLICK_POPUP_CODICE_IVA = (1145, 603)


# --- FINE CONFIGURAZIONE ---

def restart_script():
    current_pid = os.getpid()
    print(f"\n--- RICHIESTA DI RIAVVIO (TASTO '{TASTO_PER_RIAVVIARE}' PREMUTO) (PID: {current_pid}) ---")
    try:
        keyboard.remove_all_hotkeys()
    except Exception:
        pass
    python_executable = sys.executable
    script_to_run = os.path.abspath(sys.argv[0])
    subprocess.Popen([python_executable, script_to_run] + sys.argv[1:])
    sys.exit(0)

def force_excel_recalculation(filepath):
    if not WIN32_AVAILABLE:
        print("ERRORE: Impossibile forzare il ricalcolo, pywin32 non disponibile.")
        return False
    if not os.path.exists(filepath):
        print(f"ERRORE: File non trovato per il ricalcolo: {filepath}")
        return False

    print(f"\nTentativo di forzare il ricalcolo di Excel per: {os.path.basename(filepath)} (PID: {os.getpid()})...")
    excel_instance = None
    wb_instance = None
    try:
        excel_instance = win32.Dispatch('Excel.Application')
        excel_instance.Visible = False
        excel_instance.DisplayAlerts = False
        excel_instance.EnableEvents = False
        wb_instance = excel_instance.Workbooks.Open(filepath)
        print("  Forzatura ricalcolo in corso...")
        wb_instance.Application.CalculateFullRebuild()
        print("  Ricalcolo completato.")
        print("  Salvataggio del file...")
        wb_instance.Save()
        print("  File salvato.")
        wb_instance.Close(SaveChanges=False)
        print("  Workbook chiuso.")
        wb_instance = None
        excel_instance.EnableEvents = True
        excel_instance.Quit()
        print("  Istanza Excel chiusa.")
        excel_instance = None
        return True
    except Exception as e:
        print(f"\nERRORE durante il ricalcolo forzato di Excel (PID: {os.getpid()}):")
        print(e)
        traceback.print_exc()
        return False
    finally:
        if wb_instance is not None:
            try:
                wb_instance.Close(SaveChanges=False)
            except Exception: pass
        if excel_instance is not None:
            try:
                excel_instance.EnableEvents = True
                excel_instance.Quit()
            except Exception: pass

def gestisci_popup_con_ocr(regione_screenshot, testo_da_cercare, coordinate_click, nome_popup_log):
    """
    Funzione generica per gestire un popup: screenshot, OCR e click.
    """
    global OCR_AVAILABLE
    if not OCR_AVAILABLE:
        print(f"ATTENZIONE: OCR non disponibile, impossibile gestire il popup '{nome_popup_log}'.")
        return False
    try:
        pytesseract.pytesseract.tesseract_cmd = PATH_TESSERACT_CMD
        time.sleep(PAUSA_ATTESA_POPUP)
        screenshot = pyautogui.screenshot(region=regione_screenshot)
        screenshot.save(f"debug_popup_{nome_popup_log}.png")
        testo_estratto = pytesseract.image_to_string(screenshot, lang='ita')
        print(f"\n[DEBUG OCR - {nome_popup_log}] Testo estratto: '{testo_estratto.strip()}'")

        if testo_da_cercare.lower() in testo_estratto.lower():
            print(f"  Testo '{testo_da_cercare}' trovato nel popup '{nome_popup_log}'. Click in corso.")
            pyautogui.click(*coordinate_click)
            time.sleep(1.0)
            return True
        else:
            print(f"  ATTENZIONE: Testo '{testo_da_cercare}' NON trovato nel popup '{nome_popup_log}'.")
            return False
    except Exception as e:
        print(f"\nERRORE durante la gestione del popup '{nome_popup_log}' con OCR: {e}")
        traceback.print_exc()
        return False

def esegui_procedura_registrazione_odc(valore_odc_fallito, dati_riga_completa):
    """
    Contiene la sequenza di azioni per registrare una nuova commessa (ODC).
    """
    print("\n--- INIZIO PROCEDURA REGISTRAZIONE ODC ---")
    try:
        # Step 1: Click sul tasto "No" del popup iniziale
        print("  Azione 1: Click su TASTO_NO del popup.")
        pyautogui.click(*COORDINATE_POPUP_ODC_TASTO_NO)
        time.sleep(1.0)

        # Step 2: Click sulla barra dei preferiti
        print("  Azione 2: Click su BARRA_TASTO_PREFERITI.")
        pyautogui.click(*COORDINATE_BARRA_TASTO_PREFERITI)
        time.sleep(1.0)

        # Step 3: Click su commesse di variante
        print("  Azione 3: Click su BARRA_COMMESSE_DI_VARIANTE.")
        pyautogui.click(*COORDINATE_BARRA_COMMESSE_DI_VARIANTE)
        time.sleep(1.5)

        # Step 4: Click per aprire un nuovo foglio
        print("  Azione 4: Click su NUOVO_FOGLIO.")
        pyautogui.click(*COORDINATE_APERTURA_NUOVO_FOGLIO)
        time.sleep(1.0)

        # Step 5: Inserimento del codice commessa che ha causato l'errore
        print(f"  Azione 5: Inserimento codice commessa '{valore_odc_fallito}'.")
        pyautogui.click(*COORDINATE_CASELLA_COMMESSA)
        pyautogui.write(str(valore_odc_fallito), interval=0.05)
        pyautogui.press('tab')
        time.sleep(0.5)

        # Step 6: Inserimento della descrizione dalla colonna 'T'
        descrizione = dati_riga_completa.get('T', '')
        print(f"  Azione 6: Inserimento descrizione '{descrizione}'.")
        pyautogui.click(*COORDINATE_CASELLA_DESCRIZIONE)
        if descrizione:
            pyperclip.copy(str(descrizione))
            pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('tab')
        time.sleep(0.5)

        # Step 7: Inserimento commessa padre
        print("  Azione 7: Inserimento COMMESSA_PADRE '23/134'.")
        pyautogui.click(*COORDINATE_CASELLA_COMMESSA_PADRE)
        pyautogui.write("23/134", interval=0.05)
        pyautogui.press('tab'); time.sleep(0.5) # Aumentato leggermente il tempo per far apparire il popup

        # --- NUOVA GESTIONE POPUP ---
        print("  Azione 7.1: Gestione popup 'Attenzione SAL'.")
        gestisci_popup_con_ocr(
            REGIONE_SCREENSHOT_POPUP_ATTENZIONE_SAL,
            TESTO_POPUP_ATTENZIONE_SAL,
            COORDINATE_CLICK_POPUP_ATTENZIONE_SAL,
            "Attenzione_SAL"
        )

        print("  Azione 7.2: Gestione popup 'Codice IVA'.")
        gestisci_popup_con_ocr(
            REGIONE_SCREENSHOT_POPUP_CODICE_IVA,
            TESTO_POPUP_CODICE_IVA,
            COORDINATE_CLICK_POPUP_CODICE_IVA,
            "Codice_IVA"
        )
        # --- FINE NUOVA GESTIONE POPUP ---

        pyautogui.press('tab'); time.sleep(0.3)
        pyautogui.press('tab'); time.sleep(0.3)

        # Step 8: Inserimento riferimento
        print("  Azione 8: Inserimento NS_RIF '1765'.")
        pyautogui.click(*COORDINATE_CASELLA_NS_RIF)
        pyautogui.write("1765", interval=0.05)
        pyautogui.press('tab')
        time.sleep(0.5)

        # Step 9: Apertura elenco gruppo
        print("  Azione 9: Click su APRI_ELENCO_GRUPPO.")
        pyautogui.click(*COORDINATE_APRI_ELENCO_GRUPPO)
        time.sleep(1.0)

        # Step 10: Selezione gruppo generico
        print("  Azione 10: Click su CASELLA_GRUPPO_GENERICO.")
        pyautogui.click(*COORDINATE_CASELLA_GRUPPO_GENERICO)
        time.sleep(0.5)

        # Step 11: Inserimento responsabile commessa
        print("  Azione 11: Inserimento RESP_COMMESSA '1765'.")
        pyautogui.click(*COORDINATE_CASELLA_RESP_COMMESSA)
        pyautogui.write("1765", interval=0.05)
        pyautogui.press('tab')
        time.sleep(0.5)

        # Step 12: Inserimento responsabile cantiere
        print("  Azione 12: Inserimento RESP_CANTIERE '1765'.")
        pyautogui.click(*COORDINATE_CASELLA_RESP_CANTIERE)
        pyautogui.write("1765", interval=0.05)
        pyautogui.press('tab')
        time.sleep(0.5)

        # Step 13: Click su barra causali depositi
        print("  Azione 13: Click su BARRA_CAUSALI_DEPOSITI.")
        pyautogui.click(*COORDINATE_BARRA_CAUSALI_DEPOSITI)
        time.sleep(1.0)

        # Step 14: Apertura elenco deposito
        print("  Azione 14: Click su APRI_ELENCO_DEPOSITO.")
        pyautogui.click(*COORDINATE_APRI_ELENCO_DEPOSITO)
        time.sleep(1.0)

        # Step 15: Selezione deposito
        print("  Azione 15: Click su CASELLA_DEPOSITO.")
        pyautogui.click(*COORDINATE_CASELLA_DEPOSITO)
        time.sleep(0.5)

        # Step 16: Click su salvataggio ODC (ATTIVO)
        print("  Azione 16: Click su SALVATAGGIO_ODC.")
        pyautogui.click(*COORDINATE_SALVATAGGIO_FOGLIO) # Azione di salvataggio riattivata
        time.sleep(2.0)

        # Step 17: Ritorno al rapportino commessa
        print("  Azione 17: Click su RITORNA_RAPPORTINO_COMMESSA.")
        pyautogui.click(*COORDINATE_CHIUDI_RAPPORTINO_COMMESSA)
        time.sleep(1.5)

        print("--- PROCEDURA REGISTRAZIONE ODC COMPLETATA ---")
        return True

    except Exception as e:
        print(f"\nERRORE durante l'esecuzione della procedura di registrazione ODC: {e}")
        traceback.print_exc()
        return False

def controlla_e_gestisci_popup_odc(valore_odc_fallito, dati_riga_completa):
    """
    Cattura uno screenshot di una regione specifica, esegue l'OCR e
    se trova il testo target, lancia la procedura di gestione.
    """
    global OCR_AVAILABLE
    if not OCR_AVAILABLE:
        return False
    try:
        pytesseract.pytesseract.tesseract_cmd = PATH_TESSERACT_CMD
        time.sleep(PAUSA_ATTESA_POPUP)
        screenshot = pyautogui.screenshot(region=REGIONE_SCREENSHOT_POPUP)
        screenshot.save("debug_popup_screenshot.png")
        testo_estratto = pytesseract.image_to_string(screenshot, lang='ita')
        print(f"\n[DEBUG OCR] Testo estratto: '{testo_estratto.strip()}'")
        if TESTO_DA_CERCARE_NEL_POPUP.lower() in testo_estratto.lower():
            esegui_procedura_registrazione_odc(valore_odc_fallito, dati_riga_completa)
            return True
        else:
            return False
    except Exception as e:
        print(f"\nERRORE durante il controllo OCR del popup: {e}")
        traceback.print_exc()
        return False

def esegui_copia_da_buffer_e_verifica(valore_da_copiare_str, cella_excel_id_sorgente, _operazione_descr=""):
    for tentativo in range(TENTATIVI_MAX_COPIA_DA_EXCEL):
        try:
            pyperclip.copy('')
            time.sleep(PULIZIA_APPUNTI_PAUSA_1)
            pyperclip.copy('---PULIZIA BUFFER---')
            time.sleep(PULIZIA_APPUNTI_PAUSA_2)
            pyperclip.copy('')
            time.sleep(PULIZIA_APPUNTI_PAUSA_3)
            time.sleep(RITARDO_PRIMA_DI_COPIARE_IN_VERIFICA_EXCEL)
            pyperclip.copy(valore_da_copiare_str)
            time.sleep(RITARDO_DOPO_COPIA_DA_EXCEL_PER_VERIFICA)
            contenuto_appunti = pyperclip.paste()
            if contenuto_appunti == valore_da_copiare_str:
                pyperclip.copy(valore_da_copiare_str)
                time.sleep(RICONFERMA_COPIA_PAUSA)
                return True
            else:
                if tentativo == TENTATIVI_MAX_COPIA_DA_EXCEL - 1:
                    print(f"\nATTENZIONE: Errore verifica copia da buffer dopo {TENTATIVI_MAX_COPIA_DA_EXCEL} tentativi per {cella_excel_id_sorgente}. Atteso: '{valore_da_copiare_str}', Trovato: '{contenuto_appunti}'")
        except pyperclip.PyperclipException as pyperclip_err:
            print(f"\nECCEZIONE Pyperclip durante copia da buffer (Tentativo {tentativo + 1}): {pyperclip_err}")
        except Exception as e_copia:
            print(f"\nERRORE INATTESO durante copia da buffer (Tentativo {tentativo + 1}): {e_copia}")
        if tentativo < TENTATIVI_MAX_COPIA_DA_EXCEL - 1:
            time.sleep(RITARDO_TRA_TENTATIVI_COPIA_EXCEL)
    print(f"\nFALLIMENTO COPIA DA BUFFER DEFINITIVO per {cella_excel_id_sorgente}.")
    return False

def esegui_incolla_singolo_click(valore_incollato_str, target_x, target_y, descrizione_azione="", select_all_before_paste=False):
    pyperclip.copy(valore_incollato_str)
    time.sleep(RITARDO_PRIMA_DI_COPIARE_FINALE_PRE_INCOLLA)
    pyautogui.moveTo(target_x, target_y, duration=DURATA_MOVIMENTO_MOUSE)
    time.sleep(RITARDO_DOPO_MOVIMENTO_MOUSE)
    pyautogui.click()
    time.sleep(RITARDO_TRA_CLICK_SINGOLO)
    if select_all_before_paste:
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(RITARDO_DOPO_SELECT_ALL)
    time.sleep(RITARDO_PRIMA_INCOLLA)
    contenuto_attuale_appunti = pyperclip.paste()
    if contenuto_attuale_appunti != valore_incollato_str:
        print(f"\nATTENZIONE: Appunti cambiati PRIMA di Ctrl+V per '{descrizione_azione}'! Atteso:'{valore_incollato_str}', Trovato:'{contenuto_attuale_appunti}'. Ricopio...")
        pyperclip.copy(valore_incollato_str)
        time.sleep(0.15)
    pyautogui.hotkey('ctrl', 'v')

def esegui_incolla_e_log_click_ctrlA_tab(valore_incollato_str, target_x, target_y, cella_excel_id_sorgente, _descr_col_log_arg):
    pyperclip.copy(valore_incollato_str)
    time.sleep(RITARDO_PRIMA_DI_COPIARE_FINALE_PRE_INCOLLA)
    pyautogui.moveTo(target_x, target_y, duration=DURATA_MOVIMENTO_MOUSE)
    time.sleep(RITARDO_DOPO_MOVIMENTO_MOUSE)
    pyautogui.click()
    time.sleep(RITARDO_TRA_CLICK_SINGOLO)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(RITARDO_DOPO_SELECT_ALL)
    time.sleep(RITARDO_PRIMA_INCOLLA)
    contenuto_attuale_appunti = pyperclip.paste()
    if contenuto_attuale_appunti != valore_incollato_str:
        print(f"\nATTENZIONE: Appunti cambiati PRIMA di Ctrl+V per {cella_excel_id_sorgente}! Atteso:'{valore_incollato_str}', Trovato:'{contenuto_attuale_appunti}'. Ricopio...")
        pyperclip.copy(valore_incollato_str)
        time.sleep(0.15)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    time.sleep(RITARDO_DOPO_TAB)

# --- FUNZIONE PRINCIPALE ---
def main():
    timestamp_inizio_main = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
    current_pid = os.getpid()
    print(f"\n--- MAIN FUNCTION STARTED (PID: {current_pid}) at {timestamp_inizio_main} ---")
    try:
        keyboard.add_hotkey(TASTO_PER_RIAVVIARE, restart_script, suppress=False)
        print(f"INFO (PID: {current_pid}): Premi '{TASTO_PER_RIAVVIARE}' in qualsiasi momento per riavviare lo script.")
    except Exception as e_hotkey_setup:
        print(f"ATTENZIONE (PID: {current_pid}): Impossibile impostare l'hotkey di riavvio '{TASTO_PER_RIAVVIARE}': {e_hotkey_setup}")
    print(f"Avvio automazione... (PID: {current_pid})")
    if FORZARE_RICALCOLO_EXCEL:
        if WIN32_AVAILABLE:
            if not force_excel_recalculation(NOME_FILE_EXCEL):
                print(f"\nERRORE CRITICO (PID: {current_pid}): Impossibile forzare il ricalcolo del file Excel.")
                # Decidi se uscire o continuare comunque. Per ora, usciamo.
                return
            else:
                print(f"Ricalcolo forzato completato con successo. (PID: {current_pid})")
                time.sleep(1)
        else:
            print(f"\nATTENZIONE (PID: {current_pid}): Il ricalcolo e' abilitato (FORZARE_RICALCOLO_EXCEL=True) ma la libreria pywin32 non e' disponibile.")
    else:
        print("\nINFO: Il ricalcolo forzato di Excel e' disabilitato tramite flag (FORZARE_RICALCOLO_EXCEL=False).")
    print(f"\nFASE 1: Lettura parametri e buffer dati da '{os.path.basename(NOME_FILE_EXCEL)}'... (PID: {current_pid})")
    colonne_config = []
    # Associa solo le colonne con una X target. 'T' viene letta ma non ha una X, quindi non verrà inclusa qui.
    for i in range(len(TARGET_X_REMOTO_PER_COLONNA)):
        col_lettera = COLONNE_EXCEL_DA_LEGGERE[i]
        col_target_x = TARGET_X_REMOTO_PER_COLONNA[i]
        col_descr = f"Colonna {i+1} ({col_lettera})"
        colonne_config.append((col_lettera, col_target_x, col_descr))

    task_da_eseguire_trovato = False
    riga_param_task_trovato = -1
    dati_bufferizzati_per_task = []
    valore_data_rapporto_da_param = None
    valore_cantiere_da_param = None
    cella_data_rapporto_sorgente = ""
    cella_cantiere_sorgente = ""
    workbook_leggi = None
    try:
        workbook_leggi = openpyxl.load_workbook(NOME_FILE_EXCEL, data_only=True, keep_vba=True)
        sheet_parametri_leggi = workbook_leggi[FOGLIO_PARAMETRI]
        sheet_dati_leggi = workbook_leggi[NOME_FOGLIO_DATI]
        sheet_scarico_ore_leggi = workbook_leggi[FOGLIO_SCARICO_ORE]
        try:
            valore_cantiere_da_param = sheet_scarico_ore_leggi[CELLA_CANTIERE_SCARICO_ORE].value
            cella_cantiere_sorgente = f"{FOGLIO_SCARICO_ORE}!{CELLA_CANTIERE_SCARICO_ORE}"
            if valore_cantiere_da_param is not None: print(f"  Cantiere letto da '{cella_cantiere_sorgente}': '{valore_cantiere_da_param}'")
            else: print(f"  ATTENZIONE: Nessun valore Cantiere in '{cella_cantiere_sorgente}'.")
        except Exception as e_leggi_cantiere:
            print(f"  ERRORE lettura Cantiere da '{FOGLIO_SCARICO_ORE}!{CELLA_CANTIERE_SCARICO_ORE}': {e_leggi_cantiere}")
            valore_cantiere_da_param = None
        for r_param in range(RIGA_INIZIO_PARAMETRI, RIGA_FINE_PARAMETRI + 1):
            valore_cella_stato_raw = sheet_parametri_leggi[f"{COLONNA_STATO_PARAMETRI}{r_param}"].value
            stato_cella_parametri_processato = str(valore_cella_stato_raw).strip().upper() if valore_cella_stato_raw is not None else ""
            if stato_cella_parametri_processato == STATO_DA_CERCARE.upper():
                print(f"\n  TROVATO TASK VALIDO! riga {r_param} foglio '{FOGLIO_PARAMETRI}'.")
                riga_param_task_trovato = r_param
                try:
                    riga_inizio_dati_task = sheet_parametri_leggi[f"{COLONNA_RIGA_INIZIO_DATI_PARAMETRI}{r_param}"].value
                    riga_fine_dati_task = sheet_parametri_leggi[f"{COLONNA_RIGA_FINE_DATI_PARAMETRI}{r_param}"].value
                    valore_data_raw = sheet_parametri_leggi[f"{COLONNA_DATA_RAPPORTO_PARAMETRI}{r_param}"].value
                    cella_data_rapporto_sorgente = f"{FOGLIO_PARAMETRI}!{COLONNA_DATA_RAPPORTO_PARAMETRI}{r_param}"
                    if isinstance(valore_data_raw, datetime):
                        valore_data_rapporto_da_param = valore_data_raw.strftime('%d/%m/%Y')
                    elif valore_data_raw is not None:
                        valore_data_rapporto_da_param = str(valore_data_raw)
                    else:
                        valore_data_rapporto_da_param = None
                    if valore_data_rapporto_da_param: print(f"  Data Rapporto letta: '{valore_data_rapporto_da_param}'")
                    else: print(f"  ATTENZIONE: Nessun valore Data Rapporto in '{cella_data_rapporto_sorgente}'.")
                    if not (isinstance(riga_inizio_dati_task, int) and isinstance(riga_fine_dati_task, int) and riga_inizio_dati_task > 0 and riga_fine_dati_task >= riga_inizio_dati_task):
                        print(f"  ERRORE: Range righe dati non valido ({riga_inizio_dati_task}-{riga_fine_dati_task}). Salto task.")
                        continue
                    print(f"  Task valido. Righe dati da '{NOME_FOGLIO_DATI}': {riga_inizio_dati_task}-{riga_fine_dati_task}.")
                    for r_dati in range(riga_inizio_dati_task, riga_fine_dati_task + 1):
                        dati_riga_corrente = {}
                        almeno_un_valore_non_nullo = False
                        for col_excel_lettera in COLONNE_EXCEL_DA_LEGGERE:
                            valore_cella = sheet_dati_leggi[f"{col_excel_lettera}{r_dati}"].value
                            dati_riga_corrente[col_excel_lettera] = valore_cella
                            if valore_cella is not None and str(valore_cella).strip() != "":
                                almeno_un_valore_non_nullo = True
                        if almeno_un_valore_non_nullo:
                            dati_bufferizzati_per_task.append({'riga_excel_num': r_dati, 'dati_colonne': dati_riga_corrente})
                        else:
                            print(f"    (Riga {r_dati} saltata: tutte colonne target vuote)")
                    if not dati_bufferizzati_per_task:
                        print(f"  ATTENZIONE: Range ({riga_inizio_dati_task}-{riga_fine_dati_task}) non contiene dati validi. Salto task.")
                        riga_param_task_trovato = -1; continue
                    print(f"  Bufferizzate {len(dati_bufferizzati_per_task)} righe di dati.")
                    task_da_eseguire_trovato = True; break
                except Exception as e_leggi_param:
                    print(f"  ERRORE INATTESO lettura parametri/buffer riga {r_param}: {e_leggi_param}"); traceback.print_exc(); continue
    except Exception as e_glob_fase1:
        print(f"\nERRORE CRITICO FASE 1 (PID: {current_pid}): {e_glob_fase1}"); traceback.print_exc()
        return
    finally:
        if workbook_leggi: workbook_leggi.close()
    if not task_da_eseguire_trovato or not dati_bufferizzati_per_task:
        print(f"\n--- NESSUN TASK '{STATO_DA_CERCARE}' VALIDO O DATI DA PROCESSARE (PID: {current_pid}) ---")
        return
    print(f"\nHai 2 secondi per preparare le finestre... (Premi '{TASTO_PER_RIAVVIARE}' per riavviare)")
    for i in range(2, 0, -1): print(f"{i}...", end=" ", flush=True); time.sleep(1)
    print("Via!")
    workbook_scrivi_openpyxl = None
    sheet_dati_scrivi = None
    errore_azioni_gui_fase2 = False
    file_salvato_con_ok = False
    azioni_finali_gui_eseguite_con_successo = False
    errore_scrittura_ok_fase4 = False
    temp_file_excel_path = NOME_FILE_EXCEL + ".temp_script_py"
    righe_dati_processate_nel_blocco_globale = 0
    try:
        try:
            print(f"\nFASE 2: Esecuzione azioni preliminari e copia/incolla dati GUI (PID: {current_pid})")
            workbook_scrivi_openpyxl = openpyxl.load_workbook(NOME_FILE_EXCEL, keep_vba=True)
            sheet_dati_scrivi = workbook_scrivi_openpyxl[NOME_FOGLIO_DATI]
            print("\n--- INIZIO AZIONI PRELIMINARI (GUI) ---")
            pyautogui.click(*COORDINATE_APERTURA_NUOVO_FOGLIO); time.sleep(PAUSA_TRA_AZIONI_PRELIMINARI_E_FINALI); print("  Azione 1: Click apertura nuovo foglio ESEGUITO.")
            if valore_data_rapporto_da_param:
                data_str = str(valore_data_rapporto_da_param)
                if esegui_copia_da_buffer_e_verifica(data_str, cella_data_rapporto_sorgente, "Data Rapporto"):
                    esegui_incolla_singolo_click(data_str, *COORDINATE_DATA_RAPPORTO, "Data Rapporto", select_all_before_paste=True)
                    time.sleep(PAUSA_TRA_AZIONI_PRELIMINARI_E_FINALI); print(f"  Azione 2: Incolla Data Rapporto '{data_str}' ESEGUITO.")
                else:
                    errore_azioni_gui_fase2 = True; raise Exception(f"Errore copia Data Rapporto ({cella_data_rapporto_sorgente}).")
            else:
                print("  Azione 2: Valore data rapporto non disponibile. Salto."); time.sleep(PAUSA_TRA_AZIONI_PRELIMINARI_E_FINALI)
            if valore_cantiere_da_param:
                cantiere_str = str(valore_cantiere_da_param)
                if esegui_copia_da_buffer_e_verifica(cantiere_str, cella_cantiere_sorgente, "Cantiere"):
                    esegui_incolla_singolo_click(cantiere_str, *COORDINATE_CANTIERE, "Cantiere", select_all_before_paste=False)
                    pyautogui.press('tab'); time.sleep(RITARDO_DOPO_TAB); print(f"  Azione 3: Incolla Cantiere '{cantiere_str}' e Tab ESEGUITO.")
                else:
                    errore_azioni_gui_fase2 = True; raise Exception(f"Errore copia Cantiere ({cella_cantiere_sorgente}).")
            else:
                print("  Azione 3: Valore cantiere non disponibile. Salto."); time.sleep(PAUSA_TRA_AZIONI_PRELIMINARI_E_FINALI)
            print("--- FINE AZIONI PRELIMINARI ---")
            print("\n--- INIZIO COPIA/INCOLLA DATI (SOLO GUI) ---")
            for i_riga_buffer, dati_riga_obj in enumerate(dati_bufferizzati_per_task):
                if errore_azioni_gui_fase2: break
                riga_excel_effettiva = dati_riga_obj['riga_excel_num']
                dati_colonne_riga = dati_riga_obj['dati_colonne']
                y_target_corrente_per_riga = TARGET_Y_INIZIALE_REMOTO_ASSOLUTO + (righe_dati_processate_nel_blocco_globale * INCREMENTO_Y_REMOTO)
                percentuale_completamento = ((i_riga_buffer + 1) / len(dati_bufferizzati_per_task)) * 100
                print(f"Processando GUI riga Excel {riga_excel_effettiva} ({i_riga_buffer + 1}/{len(dati_bufferizzati_per_task)}) - Avanzamento: {percentuale_completamento:.0f}%", end='\r'); sys.stdout.flush()
                for idx_col, (col_excel_lettera, target_x_col, descr_col_log_val) in enumerate(colonne_config):
                    valore_da_copiare = dati_colonne_riga.get(col_excel_lettera)
                    cella_excel_sorgente_id = f"{NOME_FOGLIO_DATI}!{col_excel_lettera}{riga_excel_effettiva}"
                    if valore_da_copiare is not None:
                        val_str = str(valore_da_copiare)
                        if not esegui_copia_da_buffer_e_verifica(val_str, cella_excel_sorgente_id, descr_col_log_val):
                            errore_azioni_gui_fase2 = True; print(f"\nERRORE: Copia GUI fallita per {cella_excel_sorgente_id}."); break
                        esegui_incolla_e_log_click_ctrlA_tab(val_str, target_x_col, y_target_corrente_per_riga, cella_excel_sorgente_id, descr_col_log_val)
                        if target_x_col == TARGET_X_REMOTO_PER_COLONNA[0]:
                            popup_gestito = controlla_e_gestisci_popup_odc(val_str, dati_colonne_riga)
                            if popup_gestito:
                                print("  Re-inserimento del dato originale dopo gestione popup...")
                                if esegui_copia_da_buffer_e_verifica(val_str, cella_excel_sorgente_id, "Re-invio dato"):
                                    esegui_incolla_e_log_click_ctrlA_tab(val_str, target_x_col, y_target_corrente_per_riga, cella_excel_sorgente_id, "Re-invio dato")
                                else:
                                    print(f"\nERRORE CRITICO: Fallito il re-inserimento del dato {val_str}. Interruzione.")
                                    errore_azioni_gui_fase2 = True; break
                    else:
                        pyautogui.press('tab'); time.sleep(RITARDO_DOPO_TAB)
                if errore_azioni_gui_fase2: break
                righe_dati_processate_nel_blocco_globale += 1
                if righe_dati_processate_nel_blocco_globale >= MAX_RIGHE_DATI_PER_BLOCCO_TAB:
                    print(f"\n  Raggiunto limite blocco ({MAX_RIGHE_DATI_PER_BLOCCO_TAB} righe), premo Tab {PRESSIONI_TASTO_TAB} volte...")
                    for _ in range(PRESSIONI_TASTO_TAB): pyautogui.press('tab'); time.sleep(INTERVALLO_PRESSIONI_TAB)
                    print("  Pressioni Tab completate."); righe_dati_processate_nel_blocco_globale = 0; time.sleep(PAUSA_DOPO_BLOCCO_TAB)
            print()
            if errore_azioni_gui_fase2:
                raise Exception("Errore durante le azioni GUI di copia/incolla in FASE 2.")
            print("--- FINE AZIONI GUI DI COPIA/INCOLLA DATI (FASE 2) ---")
            print("--- LE MODIFICHE 'OK' NON SONO ANCORA STATE SCRITTE NEL FILE EXCEL ---")
        except Exception as e_fase2_gui_exc:
            if not errore_azioni_gui_fase2:
               print(f"\nERRORE CRITICO DURANTE FASE 2 (Azioni GUI): {e_fase2_gui_exc}")
               traceback.print_exc()
               errore_azioni_gui_fase2 = True
        if task_da_eseguire_trovato and not errore_azioni_gui_fase2:
            print(f"\nFASE 4: INIZIO SCRITTURA 'OK', SALVATAGGIO EXCEL E AZIONI FINALI (GUI) (PID: {current_pid})")
            print(f"Premi '{TASTO_PER_PROSEGUIRE}' per SCRIVERE 'OK' in Excel, SALVARE e poi eseguire TUTTE le azioni finali GUI...")
            print(f"(Oppure premi '{TASTO_PER_RIAVVIARE}' per riavviare l'intero script)")
            try:
                keyboard.wait(TASTO_PER_PROSEGUIRE)
                print(f"'{TASTO_PER_PROSEGUIRE}' premuto. (PID: {current_pid})")
                print(f"\nFASE 4.1: Scrittura degli 'OK' in memoria nel foglio Excel...")
                try:
                    if sheet_dati_scrivi is None:
                        raise Exception("ERRORE INTERNO: sheet_dati_scrivi non è inizializzato per scrittura OK.")
                    for dati_riga_obj in dati_bufferizzati_per_task:
                        riga_excel_effettiva = dati_riga_obj['riga_excel_num']
                        cella_ok_target = f"{COLONNA_PER_SCRIVERE_OK_DATI}{riga_excel_effettiva}"
                        sheet_dati_scrivi[cella_ok_target] = "OK"
                    print("  Scrittura 'OK' in memoria completata.")
                except Exception as e_write_ok_fase4:
                    print(f"\nERRORE durante la scrittura degli 'OK' in memoria (FASE 4.1): {e_write_ok_fase4}")
                    traceback.print_exc()
                    errore_scrittura_ok_fase4 = True
                if not errore_scrittura_ok_fase4:
                    print(f"\nFASE 4.2: Tentativo di salvataggio del file Excel '{os.path.basename(NOME_FILE_EXCEL)}'...")
                    time.sleep(PAUSA_PRIMA_SALVATAGGIO_OPENPYXL)
                    for tentativo_salvataggio in range(TENTATIVI_MAX_SALVATAGGIO_EXCEL):
                        try:
                            if workbook_scrivi_openpyxl is None:
                                raise Exception("ERRORE INTERNO: workbook_scrivi_openpyxl è None prima del salvataggio.")
                            workbook_scrivi_openpyxl.save(temp_file_excel_path)
                            if os.path.exists(NOME_FILE_EXCEL):
                                try:
                                    os.remove(NOME_FILE_EXCEL)
                                except Exception as e_remove:
                                    raise Exception(f"Errore rimozione file originale: {e_remove}")
                            os.rename(temp_file_excel_path, NOME_FILE_EXCEL)
                            file_salvato_con_ok = True
                            print(f"  File Excel '{os.path.basename(NOME_FILE_EXCEL)}' salvato e aggiornato.")
                            break
                        except Exception as e_save_attempt:
                            print(f"  ERRORE salvataggio (Tentativo {tentativo_salvataggio + 1}): {e_save_attempt}")
                            if os.path.exists(temp_file_excel_path):
                                try:
                                    os.remove(temp_file_excel_path)
                                except Exception:
                                    pass
                            if tentativo_salvataggio < TENTATIVI_MAX_SALVATAGGIO_EXCEL - 1:
                                time.sleep(RITARDO_TRA_TENTATIVI_SALVATAGGIO)
                            else:
                                print("  Max tentativi salvataggio raggiunti.")
                    if not file_salvato_con_ok:
                        print("\nERRORE CRITICO: File Excel NON salvato dopo tentativi.")
                else:
                    print("Scrittura 'OK' fallita. Salvataggio Excel e azioni GUI successive saltate.")
                if not errore_scrittura_ok_fase4 and file_salvato_con_ok:
                    print(f"\nFASE 4.3: Esecuzione azioni finali GUI...")
                    azioni_finali_config = [
                        {"coords": COORDINATE_SALVATAGGIO_FOGLIO, "descr": "salvataggio foglio applicazione"},
                        {"coords": COORDINATE_ELIMINA_FLAG_STAMPA, "descr": "elimina flag stampa"},
                        {"coords": COORDINATE_PULSANTE_OK, "descr": "pulsante OK applicazione"} ]
                    for i, azione_info in enumerate(azioni_finali_config):
                        print(f"  Azione GUI finale ({i+1}/{len(azioni_finali_config)}): {azione_info['descr']}...")
                        pyautogui.click(*azione_info["coords"]); time.sleep(PAUSA_TRA_AZIONI_PRELIMINARI_E_FINALI)
                    azioni_finali_gui_eseguite_con_successo = True
                    print(f"--- FINE AZIONI FINALI GUI --- (PID: {current_pid})")
                elif not errore_scrittura_ok_fase4 and not file_salvato_con_ok:
                    print("Salvataggio Excel fallito. Azioni finali GUI saltate.")
            except RuntimeError:
                print(f"\nAttesa tasto '{TASTO_PER_PROSEGUIRE}' interrotta (PID: {current_pid}).")
            except Exception as e_fase4_ops:
                print(f"\nErrore durante FASE 4 (post pressione '{TASTO_PER_PROSEGUIRE}'): {e_fase4_ops}"); traceback.print_exc()
            if not errore_scrittura_ok_fase4 and file_salvato_con_ok:
                print(f"\nORA PREMI '{TASTO_PER_RIAVVIARE}' per rieseguire l'intero script.")
            else:
                print(f"\nProcesso non completato con successo. Premi '{TASTO_PER_RIAVVIARE}' per riprovare o Ctrl+C per uscire.")
            try:
                while True:
                    time.sleep(0.5)
            except KeyboardInterrupt:
                print(f"\nScript interrotto (Ctrl+C). (PID: {current_pid})")
        elif task_da_eseguire_trovato and errore_azioni_gui_fase2:
            print(f"\nFASE 4 SALTATA (PID: {current_pid}): Causa: errori durante FASE 2 (azioni GUI).")
            print(f"\nPremi '{TASTO_PER_RIAVVIARE}' per riprovare o Ctrl+C per uscire.")
            try:
                while True:
                    time.sleep(0.5)
            except KeyboardInterrupt:
                print(f"\nScript interrotto (Ctrl+C). (PID: {current_pid})")
    finally:
        if workbook_scrivi_openpyxl is not None:
            try:
                workbook_scrivi_openpyxl.close()
                print("  Workbook scrittura (openpyxl) chiuso.")
            except Exception as e_close_final_wb:
                print(f"  Attenzione: Errore chiusura workbook: {e_close_final_wb}")
        if os.path.exists(temp_file_excel_path) and not file_salvato_con_ok :
            try:
                os.remove(temp_file_excel_path)
                print(f"  File temporaneo '{os.path.basename(temp_file_excel_path)}' rimosso (pulizia finale).")
            except Exception as e_temp_clean_final_extra:
                print(f"  Attenzione: Impossibile eliminare temp '{os.path.basename(temp_file_excel_path)}' (pulizia finale): {e_temp_clean_final_extra}")
    print(f"\n--- RIEPILOGO ESECUZIONE (PID: {current_pid}) ---")
    if not task_da_eseguire_trovato: pass
    elif errore_azioni_gui_fase2:
        print(f"ERRORE: Task (riga {riga_param_task_trovato}) trovato, ma errori durante azioni GUI (FASE 2).")
    elif errore_scrittura_ok_fase4:
        print(f"ERRORE: Task (riga {riga_param_task_trovato}), azioni GUI FASE 2 OK, ma errore scrittura 'OK' (FASE 4.1).")
    elif not file_salvato_con_ok:
        print(f"ERRORE: Task (riga {riga_param_task_trovato}), scrittura 'OK' OK, ma salvataggio Excel fallito (FASE 4.2).")
    elif not azioni_finali_gui_eseguite_con_successo:
        print(f"ATTENZIONE: Task (riga {riga_param_task_trovato}), file Excel salvato, ma azioni GUI finali (FASE 4.3) non completate.")
    else:
        print(f"SUCCESSO: Task (riga {riga_param_task_trovato}) processato, file Excel salvato, azioni GUI finali completate.")
    if task_da_eseguire_trovato:
        print("\nIMPORTANTE: Se riavvii con 'n', verrà forzato il ricalcolo di Excel.")
    print(f"\nScript terminato (o in attesa di Ctrl+C / '{TASTO_PER_RIAVVIARE}'). (PID: {current_pid})")

if __name__ == "__main__":
    main_pid = os.getpid()
    print(f"--- Script __main__ avviato (PID: {main_pid}) ---")
    if not WIN32_AVAILABLE and os.name != 'nt':
        print("ERRORE CRITICO: Richiede Windows e 'pywin32'."); sys.exit(1)
    elif not WIN32_AVAILABLE and os.name == 'nt':
        print("ATTENZIONE: 'pywin32' non trovata. Ricalcolo Excel non disponibile.")
    script_exit_code = 0
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n--- Script interrotto (Ctrl+C) (PID: {main_pid}) ---"); script_exit_code = 130
    except SystemExit as e:
        script_exit_code = e.code if isinstance(e.code, int) else 0
        if script_exit_code != 0: print(f"\n--- Script terminato con SystemExit (Code: {script_exit_code}) (PID: {main_pid}) ---")
    except Exception as e_global:
        print(f"\n--- ERRORE GLOBALE NON GESTITO (PID: {main_pid}) ---"); traceback.print_exc(); script_exit_code = 1
    finally:
        final_pid = os.getpid()
        print(f"\nPulizia finale hotkeys... (PID: {final_pid})")
        try:
            keyboard.remove_all_hotkeys(); print("  Hotkeys rimossi.")
        except Exception as e_remove_final:
            if not (isinstance(e_remove_final, RuntimeError) and "library not initialized" in str(e_remove_final)):
                print(f"  Attenzione: Errore rimozione hotkey: {e_remove_final}")
        print(f"--- Fine esecuzione (PID: {final_pid}) --- Exit Code: {script_exit_code}")