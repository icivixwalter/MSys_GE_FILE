@REM  cd.. devi posizionarti su git_desktop
CD c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\
dir 



@ECHO   fai un controllo con status, poi aggiungi con add tutto l'albero + il PULL + IL  COMMITT e il PUSH
git status
git pull

@ECHO pausa  per il controllo status ed il pull
@TIMEOUT /T 5 /NOBREAK




@CLS
@ECHO CICLO FOR I E II CASO PER LA GESTIONE DELLA DATA
@ECHO ========================================================================================================================

@REM Spiegazione:
@REM wmic os get localdatetime: Ottiene la data e l'ora correnti in formato YYYYMMDDHHMMSS.milliseconds.
@REM for /f: Elabora l'output per estrarre la data e l'ora.
@REM Saltiamo la prima riga (skip=1) perché contiene solo l'intestazione.
@REM La variabile %MyDate% viene impostata con il valore YYYYMMDDHHMMSS.
@REM Parsing dei dati:
@REM Anno: %MyDate:~0,4%
@REM Mese: %MyDate:~4,2%
@REM Giorno: %MyDate:~6,2%
@REM Ora: %MyDate:~8,2%
@REM Minuti: %MyDate:~10,2%
@REM Secondi: %MyDate:~12,2%
@REM Output Finale:
@REM Combina anno, mese, giorno, ora, minuti e secondi in un formato leggibile: YYYY_MM_DD_HH_MM_SS.
@REM Esempio di Output:
@REM Se la data e l'ora corrente sono 16 gennaio 2025, 14:30:45, l'output sarà:

@REM Copia
@REM Modifica
@REM 2025_01_16_14_30_45
@REM Nota:
@REM Questo codice funziona in ambienti Windows con WMIC disponibile.
@REM In versioni più recenti di Windows (come Windows 11), 
@REM WMIC potrebbe essere deprecato. In tal caso, è necessario utilizzare un'alternativa come PowerShell 
@REM per ottenere la data e l'ora. 

@echo off
@ECHO.
@ECHO.
for /f "skip=1" %%x in ('wmic os get localdatetime') do if not defined MyDate set MyDate=%%x
set today=%MyDate:~0,4%_%MyDate:~4,2%_%MyDate:~6,2%_H%MyDate:~8,2%_M%MyDate:~10,2%_S%MyDate:~12,2%
echo %today%



:----------------------------CICLO FOR I CASO per la gestione della data con le sottostringhe ma aggiunge lo 0 se <10
REM for /f "skip=1" %%x in ('wmic os get localdatetime') do if not defined MyDate set MyDate=%%x
REM set today=%MyDate:~0,4%_%MyDate:~4,2%_%MyDate:~6,2%
REM echo %today%


					@REM SOSPENSIONE TEMPO IMPOSTATA CON T
					@REM ---->	timeout /t 2 /nobreak > NUL
					@REM ---->	
					REM @TIMEOUT /T 2 /NOBREAK
					@REM Pause

:----------------------------CICLO FOR  II CASO cicolo for per per la data AAA MM GG senza sottostringhe


ECHO.
ECHO. 
git add .
git commit -m "%today%_aggiornamento_MSYS"


ECHO. 
ECHO. 
@ECHO ========================================================================================================================





@REM 01)
@REM sospensione per 1 secondo con crononometro 
@REM ---->	timeout /t 2 /nobreak > NUL
@REM ---->
ECHO
ECHO
ECHO  pausa  per consentire il commit
ECHO
@TIMEOUT /T 5 /NOBREAK

git push


@rem pausa di 5 secondi per il controllo
ECHO

ECHO pausa di 5 secondi per il controllo
@TIMEOUT /T 5 /NOBREAK
