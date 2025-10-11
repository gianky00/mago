# -*- coding: utf-8 -*-
import openpyxl
import pyautogui
import time
import pyperclip
import traceback
from datetime import datetime
import sys
import os
import keyboard
import subprocess
import logging
import builtins
import json

# --- CONFIGURAZIONE LOGGING ---
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_file = 'automazione.log'
file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
file_handler.setFormatter(log_formatter)
logger = logging.getLogger()
logger.setLevel(logging.INFO)
if logger.hasHandlers():
    logger.handlers.clear()
logger.addHandler(file_handler)
original_print = builtins.print

def custom_print(*args, **kwargs):
    original_print(*args, **kwargs)
    sep = kwargs.get('sep', ' ')
    message = sep.join(map(str, args))
    if message.strip():
        logger.info(message)

builtins.print = custom_print

# --- LIBRERIE OPZIONALI ---
try:
    from PIL import Image
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    print("ATTENZIONE: Librerie 'Pillow' o 'pytesseract' non trovate.")
    OCR_AVAILABLE = False

try:
    import win32com.client as win32
    WIN32_AVAILABLE = True
except ImportError:
    print("ATTENZIONE: Libreria 'pywin32' non trovata.")
    WIN32_AVAILABLE = False

# --- FUNZIONI DI CONFIGURAZIONE ---
def load_config():
    """Carica la configurazione dal file config.json."""
    try:
        with open('config.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        print("ERRORE CRITICO: File 'config.json' non trovato.")
        return None
    except json.JSONDecodeError:
        print("ERRORE CRITICO: Il file 'config.json' non è formattato correttamente.")
        return None

# --- FUNZIONI DI AUTOMAZIONE ---
def restart_script(config):
    """Riavvia lo script corrente."""
    TASTO_PER_RIAVVIARE = config['tasti_rapidi']['riavvia']
    current_pid = os.getpid()
    print(f"\n--- RICHIESTA DI RIAVVIAO (TASTO '{TASTO_PER_RIAVVIARE}' PREMUTO) (PID: {current_pid}) ---")
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

    print(f"\nTentativo di forzare il ricalcolo di Excel per: {os.path.basename(filepath)}...")
    excel_instance = None
    try:
        excel_instance = win32.Dispatch('Excel.Application')
        excel_instance.Visible = False
        excel_instance.DisplayAlerts = False
        excel_instance.EnableEvents = False
        wb = excel_instance.Workbooks.Open(filepath)
        wb.Application.CalculateFullRebuild()
        wb.Save()
        wb.Close(SaveChanges=False)
        excel_instance.Quit()
        print("Ricalcolo e salvataggio completati.")
        return True
    except Exception as e:
        print(f"\nERRORE durante il ricalcolo forzato di Excel: {e}")
        traceback.print_exc()
        return False
    finally:
        if excel_instance:
            excel_instance.Quit()

def gestisci_popup_con_ocr(config, regione_screenshot, testo_da_cercare, coordinate_click, nome_popup_log):
    if not OCR_AVAILABLE:
        print(f"ATTENZIONE: OCR non disponibile, impossibile gestire il popup '{nome_popup_log}'.")
        return False
    try:
        pytesseract.pytesseract.tesseract_cmd = config['generali']['path_tesseract_cmd']
        time.sleep(config['coordinate']['odc']['pausa_attesa_popup'])
        screenshot = pyautogui.screenshot(region=regione_screenshot)
        screenshot.save(f"debug_popup_{nome_popup_log}.png")
        testo_estratto = pytesseract.image_to_string(screenshot, lang='ita')
        print(f"\n[DEBUG OCR - {nome_popup_log}] Testo estratto: '{testo_estratto.strip()}'")

        if testo_da_cercare.lower() in testo_estratto.lower():
            print(f"  Testo '{testo_da_cercare}' trovato. Click in corso.")
            pyautogui.click(coordinate_click)
            time.sleep(1.0)
            return True
        else:
            print(f"  ATTENZIONE: Testo '{testo_da_cercare}' NON trovato nel popup '{nome_popup_log}'.")
            return False
    except Exception as e:
        print(f"\nERRORE durante la gestione del popup '{nome_popup_log}' con OCR: {e}")
        traceback.print_exc()
        return False

def esegui_procedura_registrazione_odc(config, valore_odc_fallito, dati_riga_completa):
    print("\n--- INIZIO PROCEDURA REGISTRAZIONE ODC ---")
    odc_cfg = config['coordinate']['odc']
    gui_cfg = config['coordinate']['gui']
    try:
        pyautogui.click(odc_cfg['coordinate_popup_tasto_no']); time.sleep(1.0)
        pyautogui.click(odc_cfg['coordinate_barra_preferiti']); time.sleep(1.0)
        pyautogui.click(odc_cfg['coordinate_commesse_variante']); time.sleep(1.5)
        pyautogui.click(gui_cfg['coordinate_apertura_nuovo_foglio']); time.sleep(1.0)

        pyautogui.click(odc_cfg['coordinate_casella_commessa'])
        pyautogui.write(str(valore_odc_fallito), interval=0.05)
        pyautogui.press('tab'); time.sleep(0.5)

        descrizione = dati_riga_completa.get('T', '')
        pyautogui.click(odc_cfg['coordinate_casella_descrizione'])
        if descrizione:
            pyperclip.copy(str(descrizione))
            pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('tab'); time.sleep(0.5)

        pyautogui.click(odc_cfg['coordinate_casella_commessa_padre'])
        pyautogui.write("23/134", interval=0.05)
        pyautogui.press('tab'); time.sleep(0.5)

        gestisci_popup_con_ocr(config, tuple(odc_cfg['regione_popup_attenzione_sal']), odc_cfg['testo_popup_attenzione_sal'], tuple(odc_cfg['coordinate_click_popup_attenzione_sal']), "Attenzione_SAL")
        gestisci_popup_con_ocr(config, tuple(odc_cfg['regione_popup_codice_iva']), odc_cfg['testo_popup_codice_iva'], tuple(odc_cfg['coordinate_click_popup_codice_iva']), "Codice_IVA")

        pyautogui.press('tab'); time.sleep(0.3)
        pyautogui.press('tab'); time.sleep(0.3)

        pyautogui.click(odc_cfg['coordinate_casella_ns_rif'])
        pyautogui.write("1765", interval=0.05)
        pyautogui.press('tab'); time.sleep(0.5)

        pyautogui.click(odc_cfg['coordinate_apri_elenco_gruppo']); time.sleep(1.0)
        pyautogui.click(odc_cfg['coordinate_casella_gruppo_generico']); time.sleep(0.5)

        pyautogui.click(odc_cfg['coordinate_casella_resp_commessa'])
        pyautogui.write("1765", interval=0.05)
        pyautogui.press('tab'); time.sleep(0.5)

        pyautogui.click(odc_cfg['coordinate_casella_resp_cantiere'])
        pyautogui.write("1765", interval=0.05)
        pyautogui.press('tab'); time.sleep(0.5)

        pyautogui.click(odc_cfg['coordinate_barra_causali_depositi']); time.sleep(1.0)
        pyautogui.click(odc_cfg['coordinate_apri_elenco_deposito']); time.sleep(1.0)
        pyautogui.click(odc_cfg['coordinate_casella_deposito']); time.sleep(0.5)

        pyautogui.click(gui_cfg['coordinate_salvataggio_foglio']); time.sleep(2.0)
        pyautogui.click(odc_cfg['coordinate_chiudi_rapportino_commessa']); time.sleep(1.5)

        print("--- PROCEDURA REGISTRAZIONE ODC COMPLETATA ---")
        return True
    except Exception as e:
        print(f"\nERRORE durante la procedura ODC: {e}")
        traceback.print_exc()
        return False

def controlla_e_gestisci_popup_odc(config, valore_odc_fallito, dati_riga_completa):
    if not OCR_AVAILABLE: return False
    odc_cfg = config['coordinate']['odc']
    try:
        pytesseract.pytesseract.tesseract_cmd = config['generali']['path_tesseract_cmd']
        time.sleep(odc_cfg['pausa_attesa_popup'])
        screenshot = pyautogui.screenshot(region=tuple(odc_cfg['regione_screenshot_popup_generico']))
        screenshot.save("debug_popup_screenshot.png")
        testo_estratto = pytesseract.image_to_string(screenshot, lang='ita')
        print(f"\n[DEBUG OCR] Testo estratto: '{testo_estratto.strip()}'")
        if odc_cfg['testo_da_cercare_popup_generico'].lower() in testo_estratto.lower():
            esegui_procedura_registrazione_odc(config, valore_odc_fallito, dati_riga_completa)
            return True
        return False
    except Exception as e:
        print(f"\nERRORE durante il controllo OCR del popup: {e}")
        traceback.print_exc()
        return False

def esegui_copia_da_buffer_e_verifica(config, valore_da_copiare_str, cella_excel_id_sorgente):
    timing = config['timing_e_ritardi']
    pulizia = config['pulizia_appunti']
    for _ in range(timing['tentativi_max_copia_excel']):
        try:
            pyperclip.copy('')
            time.sleep(pulizia['pausa_1'])
            pyperclip.copy('---PULIZIA---')
            time.sleep(pulizia['pausa_2'])
            pyperclip.copy('')
            time.sleep(pulizia['pausa_3'])
            time.sleep(timing['ritardo_prima_copia_verifica'])
            pyperclip.copy(valore_da_copiare_str)
            time.sleep(timing['ritardo_dopo_copia_excel'])
            if pyperclip.paste() == valore_da_copiare_str:
                pyperclip.copy(valore_da_copiare_str)
                time.sleep(pulizia['riconferma_copia_pausa'])
                return True
        except Exception as e:
            print(f"ERRORE copia buffer: {e}")
        time.sleep(timing['ritardo_tra_tentativi_copia'])
    print(f"FALLIMENTO COPIA DA BUFFER per {cella_excel_id_sorgente}.")
    return False

def esegui_incolla_singolo_click(config, valore_incollato_str, target_coords, select_all=False):
    timing = config['timing_e_ritardi']
    pyperclip.copy(valore_incollato_str)
    time.sleep(timing['ritardo_prima_copia_finale'])
    pyautogui.moveTo(target_coords[0], target_coords[1], duration=timing['ritardo_movimento_mouse'])
    time.sleep(timing['ritardo_dopo_movimento'])
    pyautogui.click()
    time.sleep(timing['ritardo_click_singolo'])
    if select_all:
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(timing['ritardo_dopo_select_all'])
    time.sleep(timing['ritardo_prima_incolla'])
    pyautogui.hotkey('ctrl', 'v')

def esegui_incolla_e_tab(config, valore_incollato_str, target_coords, cella_excel_id):
    timing = config['timing_e_ritardi']
    pyperclip.copy(valore_incollato_str)
    time.sleep(timing['ritardo_prima_copia_finale'])
    pyautogui.moveTo(target_coords[0], target_coords[1], duration=timing['ritardo_movimento_mouse'])
    time.sleep(timing['ritardo_dopo_movimento'])
    pyautogui.click()
    time.sleep(timing['ritardo_click_singolo'])
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(timing['ritardo_dopo_select_all'])
    time.sleep(timing['ritardo_prima_incolla'])
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    time.sleep(timing['ritardo_dopo_tab'])

# --- FUNZIONE PRINCIPALE DI AUTOMAZIONE ---
def run_automation(config):
    # Setup delle variabili dalla nuova configurazione
    gen_cfg = config['generali']
    file_cfg = config['file_e_fogli_excel']['impostazioni_file']
    param_cfg = config['file_e_fogli_excel']['mappature_colonne_foglio_avanzamento']
    col_mapping = config['file_e_fogli_excel']['mappatura_colonne_foglio_dati']
    gui_cfg = config['coordinate_e_dati']['gui']
    odc_cfg = config['coordinate_e_dati']['odc']
    timing_cfg = config['timing_e_ritardi']
    tasti_cfg = config['tasti_rapidi']

    NOME_FILE_EXCEL = os.path.abspath(file_cfg['percorso_file_excel'])

    print(f"--- AVVIO AUTOMAZIONE (PID: {os.getpid()}) ---")
    try:
        keyboard.add_hotkey(tasti_cfg['riavvia'], lambda: restart_script(config), suppress=False)
        print(f"INFO: Premi '{tasti_cfg['riavvia']}' per riavviare.")
    except Exception as e:
        print(f"ATTENZIONE: Impossibile impostare hotkey: {e}")

    if gen_cfg['forzare_ricalcolo_excel']:
        if WIN32_AVAILABLE:
            if not force_excel_recalculation(NOME_FILE_EXCEL):
                print("ERRORE CRITICO: Ricalcolo Excel fallito.")
                return
        else:
            print("ATTENZIONE: Ricalcolo abilitato ma pywin32 non disponibile.")
    else:
        print("INFO: Ricalcolo forzato di Excel disabilitato.")

    print(f"\nFASE 1: Lettura dati da '{os.path.basename(NOME_FILE_EXCEL)}'...")

    colonne_da_leggere = [item['colonna_excel'] for item in col_mapping]
    STATO_DA_CERCARE = "DA COMPLETARE"

    task_da_eseguire = None
    try:
        wb_leggi = openpyxl.load_workbook(NOME_FILE_EXCEL, data_only=True, keep_vba=True)
        sheet_parametri = wb_leggi[file_cfg['nome_foglio_avanzamento']]
        sheet_dati = wb_leggi[file_cfg['nome_foglio_dati']]

        # cantiere = param_cfg.get('ns_riferimento', '') # This seems to be unused now, keeping it commented for safety

        col_stato = sheet_parametri[param_cfg['colonna_stato']]
        for cella_stato in col_stato:
            stato_raw = cella_stato.value
            stato_pulito = str(stato_raw).strip().upper() if stato_raw is not None else ""

            if stato_pulito == STATO_DA_CERCARE:
                r = cella_stato.row
                print(f"  --> TASK TROVATO ALLA RIGA {r}!")
                r_inizio = sheet_parametri[f"{param_cfg['colonna_riga_inizio']}{r}"].value
                r_fine = sheet_parametri[f"{param_cfg['colonna_riga_fine']}{r}"].value
                data_raw = sheet_parametri[f"{param_cfg['colonna_data_rapporto']}{r}"].value
                data_rapporto = data_raw.strftime('%d/%m/%Y') if isinstance(data_raw, datetime) else str(data_raw)

                dati_buffer = []
                for r_dati in range(r_inizio, r_fine + 1):
                    dati_riga = {col: sheet_dati[f"{col}{r_dati}"].value for col in colonne_da_leggere}
                    if any(dati_riga.values()):
                        dati_buffer.append({'riga_excel_num': r_dati, 'dati_colonne': dati_riga})

                if dati_buffer:
                    task_da_eseguire = {
                        'riga_param_task': r,
                        'data_rapporto': data_rapporto,
                        'dati_buffer': dati_buffer
                    }
                    print(f"  Trovato task valido (riga {r}), bufferizzate {len(dati_buffer)} righe.")
                    break
        wb_leggi.close()
    except Exception as e:
        print(f"ERRORE CRITICO FASE 1: {e}"); traceback.print_exc()
        return

    if not task_da_eseguire:
        print("\n--- NESSUN TASK VALIDO TROVATO ---")
        return

    print(f"\nHai 2 secondi per preparare le finestre... (Premi '{tasti_cfg['riavvia']}' per riavviare)")
    for i in range(2, 0, -1): print(f"{i}...", end=" ", flush=True); time.sleep(1)
    print("Via!")

    errore_gui = False
    try:
        # Azioni preliminari
        pyautogui.click(gui_cfg['coordinate_apertura_nuovo_foglio']); time.sleep(timing_cfg['pausa_tra_azioni_preliminari_finali'])
        if task_da_eseguire['data_rapporto']:
            esegui_incolla_singolo_click(config, task_da_eseguire['data_rapporto'], gui_cfg['coordinate_data_rapporto'], select_all=True)
            time.sleep(timing_cfg['pausa_tra_azioni_preliminari_finali'])

        if gui_cfg.get('dato_ns_riferimento'):
            esegui_incolla_singolo_click(config, str(gui_cfg['dato_ns_riferimento']), gui_cfg['coordinate_cantiere'])
            pyautogui.press('tab'); time.sleep(timing_cfg['ritardo_dopo_tab'])

        # Ciclo sui dati
        righe_processate_blocco = 0

        descrizione_mapping = next((item for item in col_mapping if item["nome"] == "Descrizione Attività"), None)
        matricola_mapping = next((item for item in col_mapping if item["nome"] == "Matricola"), None)

        for i, riga_obj in enumerate(task_da_eseguire['dati_buffer']):
            y_target = timing_cfg['y_iniziale_remoto'] + (righe_processate_blocco * timing_cfg['incremento_y_remoto'])
            print(f"Processando riga {i+1}/{len(task_da_eseguire['dati_buffer'])}...", end='\r'); sys.stdout.flush()

            for mapping_item in col_mapping:
                if mapping_item['nome'] == "Descrizione Attività":
                    continue

                col_lettera = mapping_item['colonna_excel']
                target_x = mapping_item['target_x_remoto']
                val = riga_obj['dati_colonne'].get(col_lettera)

                if val is not None:
                    val_str = str(val)
                    cella_id = f"{file_cfg['nome_foglio_dati']}!{col_lettera}{riga_obj['riga_excel_num']}"
                    if not esegui_copia_da_buffer_e_verifica(config, val_str, cella_id):
                        raise Exception(f"Errore copia GUI per {cella_id}")

                    if target_x > 0:
                        esegui_incolla_e_tab(config, val_str, (target_x, y_target), cella_id)

                        if col_lettera == col_mapping[0]['colonna_excel']:
                            if controlla_e_gestisci_popup_odc(config, val_str, riga_obj['dati_colonne']):
                                print("  Re-inserimento dato dopo gestione popup...")
                                esegui_incolla_e_tab(config, val_str, (target_x, y_target), cella_id)
                else:
                     if target_x > 0:
                        pyautogui.press('tab'); time.sleep(timing_cfg['ritardo_dopo_tab'])

                if matricola_mapping and mapping_item['nome'] == matricola_mapping['nome'] and descrizione_mapping:
                    val_desc = riga_obj['dati_colonne'].get(descrizione_mapping['colonna_excel'])
                    if val_desc is not None:
                        val_str_desc = str(val_desc)
                        cella_id_desc = f"{file_cfg['nome_foglio_dati']}!{descrizione_mapping['colonna_excel']}{riga_obj['riga_excel_num']}"
                        if not esegui_copia_da_buffer_e_verifica(config, val_str_desc, cella_id_desc):
                            raise Exception(f"Errore copia GUI per {cella_id_desc}")
                        if descrizione_mapping['target_x_remoto'] > 0:
                            esegui_incolla_e_tab(config, val_str_desc, (descrizione_mapping['target_x_remoto'], y_target), cella_id_desc)

            righe_processate_blocco += 1
            if righe_processate_blocco >= timing_cfg['max_righe_per_blocco_tab']:
                print("\n  Raggiunto limite blocco, premo Tab...")
                for _ in range(timing_cfg['pressioni_tasto_tab']): pyautogui.press('tab', interval=timing_cfg['intervallo_pressioni_tab'])
                righe_processate_blocco = 0
                time.sleep(timing_cfg['pausa_dopo_blocco_tab'])
        print("\n--- FINE AZIONI GUI ---")

    except Exception as e:
        print(f"\nERRORE CRITICO DURANTE AZIONI GUI: {e}"); traceback.print_exc()
        errore_gui = True

    if not errore_gui:
        print(f"\nPremi '{tasti_cfg['prosegui']}' per salvare 'OK' in Excel ed eseguire le azioni finali...")
        keyboard.wait(tasti_cfg['prosegui'])

        file_salvato = False
        try:
            wb_scrivi = openpyxl.load_workbook(NOME_FILE_EXCEL, keep_vba=True)
            sheet_dati_scrivi = wb_scrivi[file_cfg['nome_foglio_dati']]
            colonna_ok_dati = config.get('mapping_colonne', {}).get('colonna_ok_dati', 'A') # Fallback to 'A'
            for riga_obj in task_da_eseguire['dati_buffer']:
                sheet_dati_scrivi[f"{colonna_ok_dati}{riga_obj['riga_excel_num']}"] = "OK"

            for _ in range(timing_cfg['tentativi_max_salvataggio_excel']):
                try:
                    wb_scrivi.save(NOME_FILE_EXCEL)
                    file_salvato = True
                    print("File Excel salvato con successo.")
                    break
                except Exception as e_save:
                    print(f"Tentativo di salvataggio fallito: {e_save}")
                    time.sleep(timing_cfg['ritardo_tra_tentativi_salvataggio'])
            wb_scrivi.close()
        except Exception as e:
            print(f"ERRORE CRITICO SCRITTURA/SALVATAGGIO EXCEL: {e}"); traceback.print_exc()

        if file_salvato:
            print("Esecuzione azioni finali GUI...")
            pyautogui.click(gui_cfg['coordinate_salvataggio_foglio']); time.sleep(timing_cfg['pausa_tra_azioni_preliminari_finali'])
            pyautogui.click(gui_cfg['coordinate_elimina_flag_stampa']); time.sleep(timing_cfg['pausa_tra_azioni_preliminari_finali'])
            pyautogui.click(gui_cfg['coordinate_pulsante_ok']); time.sleep(timing_cfg['pausa_tra_azioni_preliminari_finali'])
            print("--- AUTOMAZIONE COMPLETATA CON SUCCESSO ---")
        else:
            print("--- AUTOMAZIONE INTERROTTA: SALVATAGGIO EXCEL FALLITO ---")
    else:
        print("--- AUTOMAZIONE INTERROTTA A CAUSA DI ERRORI GUI ---")

    print(f"\nPremi '{tasti_cfg['riavvia']}' per un nuovo ciclo o chiudi la finestra.")
    try:
        while True: time.sleep(1)
    except KeyboardInterrupt:
        print("\nScript terminato.")

# This block is intentionally left empty so this file can be used as a library
if __name__ == "__main__":
    pass