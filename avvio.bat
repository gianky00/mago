@echo off
:: Imposta la directory di lavoro e il percorso di ricerca di Python
cd /d "%~dp0"
set PYTHONPATH=%~dp0

:: --- Controllo dei File Essenziali ---
:: Controlla la presenza dei file Python principali e del runtime di PyArmor
FOR %%F IN ("gui.py", "magoPyton.py", "config_ui.py") DO (
    IF NOT EXIST "%%~F" (
        echo ERRORE: File di programma essenziale '%%~F' non trovato.
        echo Assicurati di aver scompattato tutti i file dallo ZIP nella stessa cartella.
        pause
        exit /b
    )
)

IF NOT EXIST "pyarmor_runtime_*" (
    echo ERRORE: Cartella di runtime 'pyarmor_runtime_*' non trovata.
    echo Assicurati di aver scompattato tutti i file dallo ZIP nella stessa cartella.
    pause
    exit /b
)

:: Controlla la presenza dei file di configurazione
IF NOT EXIST "config.json" (
    echo ATTENZIONE: File 'config.json' non trovato.
    echo Il programma potrebbe non funzionare correttamente senza le configurazioni.
)

echo Avvio dell'applicazione in corso...

:: Esegui lo script principale della GUI senza mostrare la finestra del prompt
start "MagoPyton" /B pythonw.exe gui.py