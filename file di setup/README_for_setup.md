# Guida Tecnica all'Installazione e Configurazione di MagoPyton

Questo documento fornisce tutte le istruzioni necessarie per installare, configurare e avviare il progetto di automazione MagoPyton.

## Indice

1.  [Prerequisiti Software](#1-prerequisiti-software)
2.  [Procedura di Setup](#2-procedura-di-setup)
3.  [Prima Configurazione](#3-prima-configurazione)
4.  [Avvio dell'Applicazione](#4-avvio-dellapplicazione)
5.  [Risoluzione dei Problemi](#5-risoluzione-dei-problemi)

---

### 1. Prerequisiti Software

Prima di procedere, assicurati che i seguenti software siano installati sul tuo sistema.

#### a. Python
È fondamentale avere una versione di Python installata (preferibilmente 3.8 o successiva).
- **Download**: Puoi scaricarlo dal [sito ufficiale di Python](https://www.python.org/downloads/).
- **Importante**: Durante l'installazione, assicurati di spuntare la casella **"Add Python to PATH"**. Questo è essenziale per il corretto funzionamento degli script.

#### b. Tesseract OCR
Tesseract è un motore di riconoscimento ottico dei caratteri (OCR) utilizzato per leggere il testo da popup e finestre di dialogo durante l'automazione.
- **Download**: Scarica il setup da [questo link ufficiale](https://github.com/UB-Mannheim/tesseract/wiki).
- **Installazione**:
  - Esegui il file di installazione.
  - **Fondamentale**: Installa Tesseract nel percorso predefinito: `C:\Program Files\Tesseract-OCR`. Lo script di configurazione cercherà l'eseguibile in questa cartella. Se scegli un percorso diverso, dovrai inserirlo manualmente nell'interfaccia di configurazione.

---

### 2. Procedura di Setup

Lo script `setup.bat` è progettato per preparare l'ambiente di esecuzione in modo automatico.

1.  **Posizionamento della Cartella**: Assicurati che l'intera cartella del progetto (quella che contiene questo file) si trovi in un percorso senza restrizioni, ad esempio sul Desktop o in `C:\Progetti`. Evita cartelle di sistema come `Program Files`.

2.  **Esecuzione dello Script**:
    - Naviga nella cartella `file di setup`.
    - Fai clic con il pulsante destro del mouse su `setup.bat` e seleziona **"Esegui come amministratore"**.
    - Lo script eseguirà le seguenti operazioni:
        - Richiederà i privilegi di amministratore.
        - Cercherà un'installazione valida di Python nel sistema.
        - Creerà un **ambiente virtuale** (`venv`) nella cartella **a monte** del progetto, per mantenere i sorgenti puliti.
        - Installerà tutte le librerie Python necessarie (`requirements.txt`).

3.  **Verifica**: Al termine, una nuova cartella `venv` sarà apparsa accanto alla cartella del tuo progetto. Non dentro.

---

### 3. Prima Configurazione

Dopo il setup, il primo passo è configurare l'applicazione tramite la sua interfaccia grafica.

1.  **Avvia l'applicazione**: Esegui `avvio.bat` (non richiede privilegi di amministratore).
2.  **Apri le Impostazioni**: All'avvio, vedrai la finestra principale. Clicca sulla tab **"Impostazioni"**.
3.  **Configurazioni Essenziali**:
    - **Impostazioni File**:
        - `percorso_file_excel`: Clicca su "Sfoglia..." e seleziona il file `.xlsm` che l'automazione dovrà utilizzare.
        - `path_tesseract_cmd`: Se non hai installato Tesseract nel percorso predefinito, clicca "Sfoglia..." e seleziona `tesseract.exe`. Puoi anche usare il pulsante di rilevamento automatico.
    - **Mappatura Colonne e Coordinate**:
        - Naviga tra le varie sezioni per adattare le coordinate e le mappature delle colonne Excel al tuo ambiente specifico. Usa i pulsanti **"Cattura"** per registrare le coordinate direttamente dallo schermo.

4.  **Salva la Configurazione**: Una volta terminate le modifiche, clicca su **"Salva Configurazione"**. Le impostazioni verranno salvate nel file `config.json`.

---

### 4. Avvio dell'Applicazione

Per l'uso quotidiano, utilizza lo script `avvio.bat`:
- Fai doppio clic su `avvio.bat`. Questo script attiva l'ambiente virtuale corretto e avvia l'interfaccia grafica principale (`gui.py`) senza mostrare una finestra del prompt dei comandi.

---

### 5. Risoluzione dei Problemi

- **Log degli Errori**: Qualsiasi problema o errore durante l'esecuzione dell'automazione viene registrato nel file `automazione.log`, che si trova nella cartella principale del progetto.
- **Permessi**: Lo script `setup.bat` necessita dei permessi di amministratore. `avvio.bat` no.
- **Librerie mancanti**: Se ricevi errori relativi a moduli Python mancanti, esegui nuovamente `setup.bat` come amministratore per reinstallare le dipendenze.