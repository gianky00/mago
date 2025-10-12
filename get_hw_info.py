import tkinter as tk
from tkinter import messagebox

def get_hardware_info():
    """
    Ottiene l'ID hardware utilizzando la funzione runtime di PyArmor.
    Questa funzione è disponibile solo quando lo script è stato offuscato.
    """
    try:
        # PyArmor sostituisce questo import con il riferimento corretto al runtime
        # durante il processo di offuscamento.
        from pyarmor_runtime import get_hd_info
        # hd_key=2 usa un mix di identificatori hardware (scelta robusta)
        return get_hd_info(2)
    except ImportError:
        # Questo errore viene mostrato se si esegue lo script sorgente
        # direttamente o se il pacchetto è incompleto.
        return "ERRORE: Impossibile trovare il runtime di PyArmor. Pacchetto corrotto o incompleto."
    except Exception as e:
        return f"Si è verificato un errore imprevisto: {e}"

def main():
    """
    Crea la GUI per mostrare l'ID hardware.
    """
    hw_info = get_hardware_info()

    root = tk.Tk()
    root.title("Utility Info Hardware")
    root.geometry("650x180")
    root.resizable(False, False)

    # Centra la finestra
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    main_frame = tk.Frame(root, padx=20, pady=20)
    main_frame.pack(fill=tk.BOTH, expand=True)

    label = tk.Label(main_frame, text="Copia il seguente ID Hardware e invialo all'amministratore:", font=("Helvetica", 12))
    label.pack(pady=(0, 10))

    info_entry = tk.Entry(main_frame, font=("Courier", 10), width=80)
    info_entry.insert(0, hw_info)
    info_entry.config(state='readonly')
    info_entry.pack(pady=5)

    def copy_to_clipboard():
        root.clipboard_clear()
        root.clipboard_append(hw_info)
        messagebox.showinfo("Copiato", "ID Hardware copiato con successo negli appunti!")

    copy_button = tk.Button(main_frame, text="Copia ID", command=copy_to_clipboard, font=("Helvetica", 11, "bold"), height=2, width=15)
    copy_button.pack(pady=10)

    # Disabilita il pulsante se non è stato possibile ottenere l'ID
    if "ERRORE" in hw_info:
        copy_button.config(state=tk.DISABLED)

    root.mainloop()

if __name__ == "__main__":
    main()