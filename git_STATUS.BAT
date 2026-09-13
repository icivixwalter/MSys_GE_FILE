@REM  cd.. devi posizionarti su git_desktop
echo off
cls 
CD c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\
dir 




			@REM CICLO FOR I E II CASO PER LA GESTIONE DELLA DATA
			@REM ========================================================================================================================
			
			:----------------------------CICLO FOR I CASO per la gestione della data con le sottostringhe ma aggiunge lo 0 se <10
			for /f "skip=1" %%x in ('wmic os get localdatetime') do if not defined MyDate set MyDate=%%x
			REM echo solo il giorno:
			REM echo %MyDate:~6,2%

			REM echo I CASO la data con le sottostringhe con separataore l'undescore (_)
			set today=%MyDate:~0,4%_%MyDate:~4,2%_%MyDate:~6,2%

			REM  echo.
			REM echo today in formato stringa: 
			REM echo %today%
			REM echo.

			:----------------------------CICLO FOR  II CASO cicolo for per per la data AAA MM GG senza sottostringhe

			REM ECHO E' possibile ottenere la data corrente in modo indipendente dalle impostazioni locali utilizzando
			REM ECHO SENZA armeggiare con le sottostringhe
			REM echo vedi il link: https://qastack.it/programming/10945572/windows-batch-formatted-date-into-variable
			REM echo .

			REM echo. II CASO la data senza le sottostringhe solo numerico e con separatore il trattino (-)

			@REM for /f %%x in ('wmic path win32_localtime get /format:list ^| findstr "="') do set %%x
			@REM set today=%Year%_%Month%_%Day%

			REM echo.
			REM ECHO %TODAY%
			REM echo.

					@REM SOSPENSIONE TEMPO IMPOSTATA CON T
					@REM ---->	timeout /t 2 /nobreak > NUL
					@REM ---->	
					REM @TIMEOUT /T 2 /NOBREAK
					@REM Pause


			@REM CICLO FOR I E II CASO PER LA GESTIONE DELLA DATA  *** FINE ***
			@REM ========================================================================================================================



ECHO.  
ECHO ........................................................................
ECHO.
ECHO      		CONTROLLO STATUS IN DATA %TODAY%
ECHO ........................................................................
ECHO. 
git status  
ECHO 
ECHO ........................................................................
@rem pausa per consentire il commit

@REM 01)
@REM sospensione per 1 secondo con crononometro 
@REM ---->	timeout /t 2 /nobreak > NUL
@REM ---->
ECHO
ECHO
ECHO  pausa di 5 secondi per consentire il commit
ECHO
@TIMEOUT /T 5 /NOBREAK

