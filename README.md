# Progetto di Automazione RPA - MagoPyton

## Descrizione del Progetto

Questo progetto è uno script di automazione dei processi robotici (RPA) scritto in Python. Lo script è progettato per automatizzare le attività di inserimento dati interagendo con un'applicazione desktop e un file Excel. Utilizza il riconoscimento ottico dei caratteri (OCR) per gestire i popup e può essere configurato per diverse macchine.

## Prerequisiti

Prima di eseguire il progetto, assicurati di avere installato il seguente software:

1.  **Python**: Assicurati che Python sia installato sul tuo sistema. Puoi scaricarlo dal [sito ufficiale di Python](https://www.python.org/downloads/).
2.  **Tesseract OCR**: Questo strumento è necessario per il riconoscimento del testo nei popup.
    *   Scarica il programma di installazione di Tesseract da [questo link](https://github.com/UB-Mannheim/tesseract/wiki).
    *   Durante l'installazione, assicurati di installarlo nel percorso predefinito: `C:\Program Files\Tesseract-OCR`.

## Installazione

Per installare e configurare il progetto, segui questi passaggi:

1.  **Esegui lo script di installazione con privilegi di amministratore**:
    *   Fai clic con il pulsante destro del mouse su `avvio_sicuro.bat` e seleziona **Esegui come amministratore**.
    *   Lo script eseguirà le seguenti azioni:
        *   Creerà un ambiente virtuale Python nella cartella `venv`.
        *   Installerà tutte le dipendenze richieste dal file `requirements.txt`.
        *   Avvierà l'applicazione principale.

## Utilizzo

Dopo aver completato l'installazione e la configurazione, esegui lo script `avvio_sicuro.bat` come amministratore per avviare l'automazione. Lo script gestirà tutte le dipendenze e avvierà l'applicazione.

## Troubleshooting

*   **Log**: I log vengono salvati nel file `automazione.log`. Se riscontri problemi, controlla questo file per messaggi di errore dettagliati.
*   **Permessi**: Assicurati sempre di eseguire `avvio_sicuro.bat` come amministratore per evitare problemi di permessi.