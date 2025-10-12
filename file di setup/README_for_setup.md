# Guida all'Installazione di MagoPyton

## Prerequisiti

Prima di iniziare, assicurati di avere installato il seguente software:

1.  **Python**: È necessario che Python sia installato sul tuo sistema. Puoi scaricarlo dal [sito ufficiale di Python](https://www.python.org/downloads/).
2.  **Tesseract OCR**: Questo strumento è fondamentale per il riconoscimento del testo nei popup.
    *   Scarica il programma di installazione di Tesseract da [questo link](https://github.com/UB-Mannheim/tesseract/wiki).
    *   Durante l'installazione, assicurati di installarlo nel percorso predefinito: `C:\Program Files\Tesseract-OCR`.

## Procedura di Installazione

Per installare e configurare l'ambiente di MagoPyton, segui questi passaggi:

1.  **Esegui lo script di installazione con privilegi di amministratore**:
    *   Fai clic con il pulsante destro del mouse su `setup.bat` e seleziona **Esegui come amministratore**.
    *   Lo script eseguirà automaticamente le seguenti azioni:
        *   Creerà un ambiente virtuale Python (`venv`) nella cartella principale del progetto.
        *   Installerà tutte le dipendenze necessarie dal file `requirements.txt`.
        *   Configurerà tutto il necessario per il corretto funzionamento dell'applicazione.

Una volta completata l'installazione, puoi avviare l'applicazione eseguendo il file `avvio.bat` che si trova nella cartella principale del progetto.