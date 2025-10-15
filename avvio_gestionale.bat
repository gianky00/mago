@echo off
REM Imposta la cartella di lavoro corrente su quella dello script.
cd /d "%~dp0"

echo Configurazione ambiente temporaneo per Python 3.9...

REM Modifica il PATH *SOLO PER QUESTA FINESTRA DEL TERMINALE*.
REM Mette il percorso di Python 3.9 all'inizio, in modo che
REM qualsiasi sottoprogramma (come pyarmor-7) trovi questa versione per prima.
REM Questa modifica svanisce appena lo script si chiude e NON altera il tuo sistema.
set PATH=C:\Users\Coemi\AppData\Local\Programs\Python\Python39;%PATH%

echo Avvio del Gestionale con l'interprete Python corretto...

REM Usiamo la chiamata diretta al python del venv per avviare il nostro script.
..\venv_gestionale\Scripts\python.exe license_manager.py

echo Il programma e' stato chiuso.
pause