# Guida alla Distribuzione e Licenziamento (Metodo Semplificato)

Questo documento spiega il processo per creare e distribuire pacchetti software licenziati, con un flusso di lavoro semplificato sia per l'amministratore che per l'utente finale.

---

## Flusso di Lavoro Generale

Il processo si divide in 3 semplici macro-fasi:

1.  **Tu (Amministratore):** Crei una piccola utility per ottenere l'ID Hardware e la invii al collega.
2.  **Il tuo Collega (Utente):** Esegue l'utility, copia l'ID e te lo invia.
3.  **Tu (Amministratore):** Salvi l'ID, lo associ a un nome e generi il pacchetto software finale e licenziato da inviare al collega.

---

## Sezione 1: Istruzioni per l'Amministratore (Tu)

Usa lo strumento `license_generator.py` per gestire tutto. Avvialo con:
```bash
python license_generator.py
```

### FASE A: Ottenere l'ID Hardware dal Collega

1.  Nel generatore, clicca il pulsante **"Crea Utility Info Hardware..."**.
2.  Ti verrà chiesto dove salvare un file `.zip`. Salvalo con un nome chiaro (es. `utility-info-hardware.zip`).
3.  Invia questo file `.zip` al tuo collega, insieme alle istruzioni che trovi nella "Sezione 2" di questo README.

### FASE B: Salvare la Licenza e Generare il Pacchetto Finale

1.  Una volta che il collega ti ha inviato il suo ID Hardware:
    *   Nel generatore, vai alla **FASE 2**.
    *   **ID Licenza**: Inserisci un nome descrittivo per questa licenza (es. "PC Fisso Mario Rossi", "Laptop Sede Milano").
    *   **ID Hardware**: Incolla il codice ricevuto dal collega.
    *   Clicca su **"Salva Licenza"**. La licenza verrà salvata e sarà disponibile nel menu a tendina per usi futuri.

2.  Per creare il software da inviare al collega:
    *   Vai alla **FASE 3**.
    *   **Seleziona Licenza Salvata**: Scegli dal menu a tendina la licenza che hai appena salvato.
    *   **Data Scadenza**: Imposta la data di scadenza desiderata.
    *   Clicca su **"Genera Pacchetto Finale..."**.
    *   Scegli una cartella vuota dove salvare il pacchetto completo.

3.  Al termine, la cartella di destinazione conterrà **tutto il necessario**. Comprimila in un file `.zip` e inviala al tuo collega.

---

## Sezione 2: Istruzioni per l'Utente Finale (I tuoi Colleghi)

*Queste sono le istruzioni da inviare ai tuoi colleghi.*

### Parte 1: Trovare il tuo ID Hardware

Per poterti fornire una licenza software, ho bisogno di un codice identificativo del tuo PC.

1.  Riceverai da me un file `.zip` (es. `utility-info-hardware.zip`).
2.  Scompatta il file in una cartella qualsiasi.
3.  Entra nella cartella e fai doppio clic su **`AVVIA_PER_ID_HARDWARE.bat`**.
4.  Apparirà una finestra con il tuo ID Hardware.
5.  Clicca sul pulsante **"Copia ID"** e inviami via email o chat il codice che hai appena copiato.

### Parte 2: Installare l'Applicazione Finale

Una volta che mi avrai inviato l'ID, ti manderò un nuovo file `.zip` con il programma completo.

1.  Scompatta questo secondo file `.zip` in una cartella permanente sul tuo computer (es. sul Desktop o in `Documenti`).
2.  Entra nella cartella e fai doppio clic su **`avvio.bat`** per lanciare il programma.

Il software è legato a questo specifico PC e non funzionerà se copiato altrove.