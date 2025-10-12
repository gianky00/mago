# Guida alla Distribuzione e Licenziamento

Questo documento spiega il processo per creare pacchetti licenziati dell'applicazione e come distribuirli agli utenti finali.

---

## Sezione 1: Per l'Amministratore (Tu)

Questa sezione descrive come generare un pacchetto software protetto e licenziato per un collega.

### Prerequisiti

1.  Assicurati di avere l'ambiente di sviluppo configurato (Python e le dipendenze da `requirements.txt` installate).
2.  Il tuo collega deve averti fornito il suo **ID Hardware**.

### Come Generare un Pacchetto Licenziato

1.  **Avvia il Generatore di Licenze**:
    Esegui lo script `license_generator.py` dalla cartella principale del progetto:
    ```bash
    python license_generator.py
    ```

2.  **Compila i Campi**:
    *   **ID Hardware**: Incolla il codice identificativo che ti ha fornito il tuo collega.
    *   **Data Scadenza**: Imposta la data di scadenza della licenza nel formato `YYYY-MM-DD`.

3.  **Genera il Pacchetto**:
    *   Clicca su **"Genera Pacchetto Licenziato"**.
    *   Ti verrà chiesto di scegliere una **cartella di destinazione**. È qui che verrà salvato il pacchetto pronto per la distribuzione. Scegli una cartella vuota o creane una nuova (es. `distribuzione_mario_rossi`).

4.  **Prepara lo ZIP**:
    *   Al termine del processo, la cartella di destinazione conterrà gli script Python offuscati (`.py`) e una cartella `pyarmor_runtime_000000`.
    *   **Copia i file ausiliari necessari** nella stessa cartella. Questi file **non** vengono inclusi automaticamente e sono essenziali per il funzionamento del programma:
        *   `config.json`
        *   `tooltips.json`
        *   `Dataease_ALLEGRETTI_02-2025.xlsm` (o il file di dati corrente)
        *   `avvio.bat`
        *   La cartella `file di setup/` se contiene elementi necessari all'avvio.
    *   Comprimi l'intera cartella di destinazione in un unico file `.zip`.

5.  **Distribuisci**:
    *   Invia il file `.zip` al tuo collega.

---

## Sezione 2: Per l'Utente Finale (I tuoi Colleghi)

Questa sezione contiene le istruzioni da fornire ai tuoi colleghi.

### A. Come Trovare il tuo ID Hardware

Per poterti generare una licenza, ho bisogno del codice identificativo del tuo PC.

1.  Apri un **Prompt dei Comandi (cmd)**.
2.  Esegui questo comando:
    ```bash
    pyarmor-7 hdinfo
    ```
    *(Nota: Se non hai `pyarmor-7` installato, puoi semplicemente installarlo con `pip install pyarmor`)*
3.  Copia l'output che ti viene mostrato (è una lunga stringa di testo) e inviamelo.

### B. Come Installare e Avviare l'Applicazione

Una volta che ricevi il file `.zip`:

1.  **Scompatta lo ZIP**: Estrai tutto il contenuto del file `.zip` in una cartella sul tuo computer (es. sul Desktop).
2.  **Avvia l'Applicazione**:
    *   Entra nella cartella appena creata.
    *   Fai doppio clic sul file `avvio.bat` per lanciare il programma.

Non tentare di copiare i file su un altro computer, perché la licenza è legata specificamente a questo PC e non funzionerebbe altrove.