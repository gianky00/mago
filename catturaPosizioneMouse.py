# trova_coordinate_avanzato.py
import pyautogui
import screeninfo
import time
import sys

# --- Funzioni Ausiliarie ---

def rileva_e_stampa_info_monitor():
    """
    Rileva tutti i monitor collegati, stampa le loro informazioni dettagliate
    e restituisce la lista di oggetti Monitor.
    """
    print("1. Rilevamento delle informazioni dei monitor collegati...")
    time.sleep(1)

    try:
        monitors = screeninfo.get_monitors()

        if not monitors:
            print("\nERRORE: Nessun monitor rilevato. Impossibile continuare.")
            return None
        
        print(f"\nSono stati rilevati {len(monitors)} monitor.\n")
        for i, monitor in enumerate(monitors):
            print(f"--- Monitor {i + 1} ({'Primario' if monitor.is_primary else 'Secondario'}) ---")
            print(f"  Nome (o ID generico): {monitor.name}")
            print(f"  Risoluzione: {monitor.width}px x {monitor.height}px")
            print(f"  Offset Globale (origine virtuale): X={monitor.x}, Y={monitor.y}")
            print("-" * (25 + len(str(i+1))))
        
        print("\nSpiegazione dell'Offset Globale:")
        print(" - L'offset (X,Y) rappresenta la posizione dell'angolo in alto a sinistra di un monitor")
        print("   rispetto all'angolo in alto a sinistra del monitor PRIMARIO, che è l'origine (0,0).")
        print("-" * 50)
        return monitors

    except Exception as e:
        print(f"\nERRORE CRITICO durante il rilevamento dei monitor: {e}")
        print("Assicurati di aver installato la libreria 'screeninfo' (pip install screeninfo).")
        return None

def trova_monitor_corrente(x_globale, y_globale, lista_monitor):
    """
    Data una coordinata globale (x, y) e una lista di monitor,
    restituisce il monitor su cui si trova il cursore e il suo indice.
    """
    for i, monitor in enumerate(lista_monitor):
        # Controlla se la coordinata globale rientra nei bordi di questo monitor
        if (monitor.x <= x_globale < monitor.x + monitor.width) and \
           (monitor.y <= y_globale < monitor.y + monitor.height):
            return monitor, i
    return None, -1 # Non trovato

# --- Script Principale ---

# 1. Rileva i monitor all'avvio
lista_monitor = rileva_e_stampa_info_monitor()

if not lista_monitor:
    sys.exit() # Esce dallo script se non sono stati trovati monitor

# 2. Avvia il tracciamento del mouse
print("\n2. Avvio del tracciamento delle coordinate del mouse...")
print("Muovi il mouse sullo schermo. Per fermare, premi Ctrl+C nella console.")
print("-" * 50)

try:
    while True:
        # Ottieni le coordinate GLOBALI correnti del mouse
        x_globale, y_globale = pyautogui.position()

        # Trova su quale monitor si trova il cursore
        monitor_attivo, indice_monitor = trova_monitor_corrente(x_globale, y_globale, lista_monitor)

        output_str = f"Coordinate Globali: X={str(x_globale).rjust(5)}, Y={str(y_globale).rjust(5)}"

        if monitor_attivo:
            # Calcola le coordinate LOCALI relative al monitor attivo
            x_locale = x_globale - monitor_attivo.x
            y_locale = y_globale - monitor_attivo.y
            
            nome_monitor = f"Monitor {indice_monitor + 1}"
            output_str += f"  |  {nome_monitor}: X={str(x_locale).rjust(5)}, Y={str(y_locale).rjust(5)} (Locali)"
        else:
            # Il cursore potrebbe trovarsi in uno spazio non mappato tra i monitor
            output_str += "  |  (Nessun monitor attivo in questa posizione)"

        # Stampa la stringa e usa padding per cancellare la riga precedente completamente
        print(output_str + " " * 10, end="\r", flush=True)

        time.sleep(0.1)

except KeyboardInterrupt:
    # L'utente ha premuto Ctrl+C
    print("\n\n" + "-" * 50)
    print("Script terminato dall'utente.")
    print("Ora conosci la differenza tra coordinate globali e locali!")
    print("-" * 50)

except Exception as e:
    print(f"\n\nSi è verificato un errore inaspettato: {e}")