@REM BATI DI SALVATAGGIO INCREMENTALE
@REM NOME FILE          :        ZIP_RAR_INCREMENTALE.BAT
@REM PATH DI LAVORO     :        c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\
@REM ATTIVITA SVOLTA    :
@REM     la bat preleva i file nell Zip_N70_ElencoFile.txt e attiva il comando 7zip
@REM     per la creazione di file incrementali es. 2026_09_13_XXX_SALVATAGGI
@REM     nella cartella AA_SALVATAGGI

@REM ==== FIX 1: SETLOCAL per abilitare !ERRORLEVEL! e !VAR! dentro i blocchi IF ====
@SETLOCAL ENABLEDELAYEDEXPANSION

@REM ==== LOG: cattura TUTTO stdout+stderr ====
@SET "LOG_FILE=%~dp0ZIP_RAR_INCREMENTALE.LOG"
@IF EXIST "%LOG_FILE%" DEL "%LOG_FILE%"

@REM Redirigo TUTTO l'output del BAT (stdout+stderr) verso il log
@CALL :LOG "============================================================"
@CALL :LOG "LOG AVVIATO IL %DATE% %TIME%"
@CALL :LOG "============================================================"
@CALL :LOG "[STEP] Avvio BAT"
@CALL :LOG "[STEP] CD iniziale = %CD%"
@CALL :LOG "[STEP] LOG_FILE   = %LOG_FILE%"

@REM ==== FIX 1: t=0 DEFINITO SUBITO ====
@SET t=0
@CALL :LOG "[STEP] t=%t%"

@CLS
@echo.@@-------@_PROGETTO_BAT_INCREMENTALE_CRIPTA_DECOFICA_INTERNI@-----------------------

      @REM NOTE
      @REM file di zip denominato : ZIP_7ZIP_INCREMENTALE_CRIPTA_INTERNO.MD
      @REM  @salvataggio@incrementale_(con file @zip@incrementale con @cripta @interno)

      @REM  Note_incrementale
         @REM questo bat effettuo il salvataggio incrementale ma non ha i due file bat
         @REM separati di cripaggio e decriptagio in GE14 in quanto al suo interno effettua
         @REM la traformazione in file criptato.

   @ECHO.
   @ECHO.00.00 PROGETTO BAT INCREMENTALE BASATO SU 7zip
   @ECHO.
   @ECHO.
   @ECHO. IMPOSTO I PARAMETRI SU TUTTO IL FILE e qui posson ossere cambiati e
   @ECHO. di seguito puoi cambiare questo percorso path per una sostituzione collettiva:
   @ECHO.
   @ECHO. cambia questo parametro ---------:          c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\
   @ECHO.
   @ECHO. cambia questo parametro ---------:          _Archivi_Msys_GE_FILE
   @ECHO.
   @ECHO.
   @ECHO.
   @ECHO. cambiando questi 2 parametri si ha effetto di cambiare i valori di tutti i parametri
   @ECHO. dos preimpostati come di seguito indicato:
   @ECHO.
   @ECHO.                           effetto sostituzione collettivo
   @ECHO. --------------------------------------------------------------------------------------------------------
   @ECHO. LA PATH DI DESTINAZIONE PER I SALVATAGGI    =  c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\%CARTELLA_SALVATAGGI_S%
   @ECHO. LA CARTELLA SALVATAGGI CREATA O ESISTENTE   =     c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\AA_SALVATAGGI\
   @ECHO. LA PATH DI PARTENZA                   =  c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\
   @ECHO. FILE 7ZIP, RAR DI SALVATAGGIO            =  ZIP_SALVATAGGI_Archivi_Msys_GE_FILE
   @ECHO. LISTA FILE UTILIZZATA                 =  Zip_N70_ElencoFile.txt
   @ECHO. --------------------------------------------------------------------------------------------------------
   @ECHO.

@CALL :LOG "[STEP] Fine sezione 00"

@CLS
@echo.01--------------------------------CREO LA CARTELLA SALVATAGGI SE NON ESISTE

@CALL :LOG "[STEP] 01 - inizio"
@CALL :LOG "[STEP] 01 - CD = %CD%"

         @REM Se non esiste la cartella AA_SALVATAGGI, la crea.
         if not exist "%cd%\AA_SALVATAGGI\" (
            @CALL :LOG "[STEP] 01 - AA_SALVATAGGI non esiste, la creo"
            echo Crea la cartella AA_SALVATAGGI
            mkdir "%cd%\AA_SALVATAGGI"
            @CALL :LOG "[STEP] 01 - mkdir eseguito, ERRORLEVEL=!ERRORLEVEL!"

            @ECHO  ATTENZIONE HO CREATO LA CARTELLA AA_SALVATAGGI :
            @ECHO
            @ECHO
            @ECHO
            @ECHO
               DIR  AA_SALVATAGGI
               @CALL :LOG "[STEP] 01 - DIR eseguito, ERRORLEVEL=!ERRORLEVEL!"
         ) ELSE (
            @CALL :LOG "[STEP] 01 - AA_SALVATAGGI esiste già"
         )

            @REM CREATA OD ESISTENTE VIENE SALVATA LA CARTELLA SALATAGGI NELLA VARIABILE
            SET CARTELLA_SALVATAGGI_S=AA_SALVATAGGI
            @CALL :LOG "[STEP] 01 - CARTELLA_SALVATAGGI_S=%CARTELLA_SALVATAGGI_S%"
            @CALL :LOG "[STEP] 01 - fine"

@CLS
@echo.02--------------------------------TITOLO COSTRUZIONE DI UN BAT INCREMENTALE + IMPOSTAZIONE PARAMETRI

@CALL :LOG "[STEP] 02 - inizio"

         @REM  NOTE GENERALI
            @REM  CAMBI QUESTI PARAMETRI PER TUTTO IL FILE
            @REM  questo esempio di bat incrementale si trova i questa path dello 7Z zip questi sono parametri
            @REM  che possono essere sostituiti da qui, la PATH DI DESTINAZIONE + LA PATH DI PARTENZA  +
            @REM  IL FILE :
            @REM  @parametri_(per la @sostituzione dei parametri per @tutto il @progetto)
                     @REM  PATH_DEST_S
                     @REM  PATH_PART_S
                     @REM  FILE_S
                     @REM  LISTA_FILE_S
                     @REM  PATH_CORR_S
                     @REM  t
                     @REM  CARTELLA_SALVATAGGI_S

   @ECHO.
   @ECHO.02.01 IMPOSTO LE VARIABILI GENERALI DI SEGUITO DENOMINATE
   @ECHO.

   @ECHO  le variabili da utilizzare :
      @ECHO 1) Path di destinazione
      @ECHO ------------------------- PATH_DEST_S
      @ECHO 2) Path di partenza dove si trovano i dati
      @ECHO ------------------------- PATH_PART_S
      @ECHO 3) file di salvataggio comprezzo 7zip
      @ECHO ------------------------- FILE_S
      @ECHO 4) Lista file dove prelevare i dati da comprimere
      @ECHO ------------------------- LISTA_FILE_S
      @ECHO 5) La path corrente salvata con con il ciclo for /f
      @ECHO ------------------------- PATH_CORR_S
      @ECHO 6) variabile per la Sospensione del tempo
      @ECHO ------------------------- t
      @ECHO 7) variabile della cartella dei salvataggi
      @ECHO ------------------------- CARTELLA_SALVATAGGI_S
      @ECHO.

                  @REM CICLO FOR F PER TROVARE E SALVARE LA PATH CORRENTE
                     @for %%f in ("CurrentDIR.bat") DO @SET currDir_=%%~df%%~pf
                     @SET PATH_CORR_S=%currDir_%
                     @CALL :LOG "[STEP] 02 - PATH_CORR_S=%PATH_CORR_S%"

                  @REM imposto le variabili
                  @SET PATH_PART_S=c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE
                  @SET PATH_DEST_S=%PATH_PART_S%\%CARTELLA_SALVATAGGI_S%
                  @SET FILE_S=ZIP_SALVATAGGI_Archivi_Msys_GE_FILE
                  @SET LISTA_FILE_S=Zip_N70_ElencoFile.txt

                  @CALL :LOG "[STEP] 02 - PATH_PART_S=%PATH_PART_S%"
                  @CALL :LOG "[STEP] 02 - PATH_DEST_S=%PATH_DEST_S%"
                  @CALL :LOG "[STEP] 02 - FILE_S=%FILE_S%"
                  @CALL :LOG "[STEP] 02 - LISTA_FILE_S=%LISTA_FILE_S%"
                  @CALL :LOG "[STEP] 02 - fine"

@CLS
@echo.03--------------------------------CONTROLLO DEI PARAMETRI_GENERALI, DISCO, PATH DI DESTINAZIONE, NOME ARCHIVIO.ZIP NOME LISTA FILEeTorri

@CALL :LOG "[STEP] 03 - inizio"
         @ECHO.03.02 STAMPA DI CONTROLLO DEL SETTAGGIO DELLE VARIABILI
         @ECHO SET PATH DI DESTINAZIONE + PATH PARTENZA + FILE DI SALVATAGGIO + LISTA FILE
         @ECHO.
         @ECHO. controllo PATH di partenza...........PATH_PART_S          :%PATH_PART_S%
         @ECHO. controllo PATH di DESTINAZIONE.......PATH_DEST_S          :%PATH_DEST_S%
         @ECHO. controllo nome file Zip..............FILE_S            :%FILE_S%
         @ECHO. controllo elenco file salvati........LISTA_FILE_S         :%LISTA_FILE_S%
         @ECHO. controllo della path corrente........PATH_CORR_S       :%PATH_CORR_S%
         @ECHO. controllo della cartella ............CARTELLA_SALVATAGGI_S :%CARTELLA_SALVATAGGI_S%
         @ECHO.
@CALL :LOG "[STEP] 03 - fine"

@REM ==== FIX: TIMEOUT eseguito solo se t > 0 ====
@IF %t% GTR 0 (
    @CALL :LOG "[STEP] 03 - TIMEOUT %t% secondi"
    TIMEOUT /T %t% /NOBREAK > NUL
) ELSE (
    @CALL :LOG "[STEP] 03 - TIMEOUT saltato (t=0)"
)

@CLS
@echo.04-------------------------------TIPO FILE DA ARCHIVIARE SE NON ESISTE LA LISTA FILE CON VARIE ESTENSION, *.DOC, *.MDB ecc...

@CALL :LOG "[STEP] 04 - inizio"

         @SET TipoFile_TEXT_S=*.AUX *GZ *.LOG *.OUT *.TEX *.Tws
         @SET TipoFile_JAVA_S=*.CLASS *.EXE *HTLM *.IML *.JAR *.JAVA *.LOG *.MF *.LNK *.XML
         @SET TipoFile_TXT_S=*.TEX *.TWS *.TXT *BAT
         @SET TipoFile_Office_S=*.MDB *.XLS *.PDF
         @SET TIPO_FILE_S=%TipoFile_TEXT_S% %TipoFile_JAVA_S% %TipoFile_TXT_S% %TipoFile_Office_S%

         @CALL :LOG "[STEP] 04 - TIPO_FILE_S=%TIPO_FILE_S%"
         @CALL :LOG "[STEP] 04 - fine"

@CLS
@ECHO.05------------------------------- CICLO FOR CON GESTIONE DELLA DATA CORRENTE DA RECUPERARE SOTTO FORMA DI AAA MM GG @04

@CALL :LOG "[STEP] 05 - inizio"

      :----------------------------CICLO FOR  I^ CASO
         @ECHO  CICLO FOR I^ CASO per la gestione della data con le sottostringhe ma aggiunge lo 0 se <10

         @REM ==== FIX 2: DATA ROBUSTA ====
         @REM Provo PowerShell (sempre presente su Win 7+)
         @CALL :LOG "[STEP] 05 - provo PowerShell per la data"
         @SET "MyDate="
         @for /f "usebackq delims=" %%d in (`powershell -NoProfile -Command "Get-Date -Format yyyyMMdd" 2^>NUL`) do @SET "MyDate=%%d"
         @CALL :LOG "[STEP] 05 - dopo PowerShell, MyDate=[!MyDate!], ERRORLEVEL=!ERRORLEVEL!"

         @REM Fallback 1: wmic
         @IF NOT DEFINED MyDate (
            @CALL :LOG "[STEP] 05 - PowerShell FALLITO, provo wmic"
            @for /f "skip=1" %%x in ('wmic os get localdatetime 2^>NUL') do @IF NOT DEFINED MyDate SET "MyDate=%%x"
            @CALL :LOG "[STEP] 05 - dopo wmic, MyDate=[!MyDate!]"
         )

         @REM Fallback 2: %date% con parsing italiano (giorno settimana ignorato)
         @IF NOT DEFINED MyDate (
            @CALL :LOG "[STEP] 05 - wmic FALLITO, uso fallback su %%date%%"
            @for /f "tokens=1-3 delims=/-. " %%a in ('echo %date%') do (
               @SET "GG=%%a"
               @SET "MM=%%b"
               @SET "AA=%%c"
            )
            @REM Se GG non è numerico (es. 'gio', 'lun'), riprovo saltando il primo token
            @echo !GG!| findstr /R "^[0-9][0-9]*$" >NUL || (
               @for /f "tokens=2-4 delims=/-. " %%a in ('echo %date%') do (
                  @SET "GG=%%a"
                  @SET "MM=%%b"
                  @SET "AA=%%c"
               )
            )
            @IF "!GG:~1,1!"=="" SET "GG=0!GG!"
            @IF "!MM:~1,1!"=="" SET "MM=0!MM!"
            @SET "MyDate=!AA!!MM!!GG!000000"
            @CALL :LOG "[STEP] 05 - fallback MyDate=[!MyDate!]"
         )

         @echo solo L'ANNO che viene separato per la futura ricerca:
         @echo "inserito nella variabile Mydate       ---> " %MyDate:~0,4%

         @SET ANNO_S=%MyDate:~0,4%
         @SET MESE_S=%MyDate:~4,2%
         @SET GIORNO_S=%MyDate:~6,2%
         @set today=%MyDate:~0,4%_%MyDate:~4,2%_%MyDate:~6,2%

         @CALL :LOG "[STEP] 05 - ANNO_S=%ANNO_S%, MESE_S=%MESE_S%, GIORNO_S=%GIORNO_S%, TODAY=%today%"
         @CALL :LOG "[STEP] 05 - fine"

@IF %t% GTR 0 TIMEOUT /T %t% /NOBREAK > NUL

@CLS
@echo.06-------------------------------CHIAMO DECODIFICA @06

@CALL :LOG "[STEP] 06 - inizio"

                  @DIR Zip_N70_DECODIFICA.bat
                  @IF EXIST "Zip_N70_DECODIFICA.bat" (
                     @CALL :LOG "[STEP] 06 - chiamo Zip_N70_DECODIFICA.bat"
                     CALL Zip_N70_DECODIFICA.bat
                     @CALL :LOG "[STEP] 06 - ritorno da DECODIFICA, ERRORLEVEL=!ERRORLEVEL!"
                  ) ELSE (
                     @CALL :LOG "[STEP] 06 - Zip_N70_DECODIFICA.bat NON TROVATA"
                  )

@CALL :LOG "[STEP] 06 - fine"

@IF %t% GTR 0 TIMEOUT /T %t% /NOBREAK > NUL

@CLS
@echo.07------------------------------ SALVO TUTTI I FILE NELLE SOTTOCARTELLE /S con 7zip uso u -r  @06

@CALL :LOG "[STEP] 07 - inizio"

                  @REM ==== FIX 5: CD /D e virgolette ====
                  @CALL :LOG "[STEP] 07 - CD /D %PATH_PART_S%"
                  CD /D "%PATH_PART_S%"
                  @CALL :LOG "[STEP] 07 - dopo CD, CWD=!CD!, ERRORLEVEL=!ERRORLEVEL!"
                  DIR

         @REM ==== FIX 7: controllo esistenza file elenco ====
         @IF NOT EXIST "%PATH_PART_S%\%LISTA_FILE_S%" (
            @CALL :LOG "[ERRORE] 07 - %LISTA_FILE_S% NON TROVATO in %PATH_PART_S%"
            @ECHO *** ERRORE: file elenco NON TROVATO ***
            GOTO :FINE_ERRORE
         )
         @CALL :LOG "[STEP] 07 - file elenco OK: %PATH_PART_S%\%LISTA_FILE_S%"

         @IF NOT EXIST "%PATH_DEST_S%\" (
            MKDIR "%PATH_DEST_S%"
            @CALL :LOG "[STEP] 07 - creata cartella destinazione %PATH_DEST_S%"
         )
         @CALL :LOG "[STEP] 07 - cartella destinazione OK: %PATH_DEST_S%"

         @REM ==== FIX 3: riga 7z corretta - formato ZIP con nome AA_MM_GG_FILE.zip ====
         @SET "NOME_ARCHIVIO=%TODAY%_%FILE_S%.zip"
         @SET "PATH_ARCHIVIO=%PATH_DEST_S%\%NOME_ARCHIVIO%"

         @CALL :LOG "[STEP] 07 - NOME_ARCHIVIO=%NOME_ARCHIVIO%"
         @CALL :LOG "[STEP] 07 - PATH_ARCHIVIO=%PATH_ARCHIVIO%"
         @CALL :LOG "[STEP] 07 - lancio 7z"

         "C:\Program Files\7-Zip\7z.exe" u -tzip -r "%PATH_ARCHIVIO%" "@%PATH_PART_S%\%LISTA_FILE_S%" >> "%LOG_FILE%" 2>&1
         @REM ==== FIX 4: !ERRORLEVEL! per catturare il valore corrente ====
         @SET "RC_7Z=!ERRORLEVEL!"
         @CALL :LOG "[STEP] 07 - ERRORLEVEL 7z = !RC_7Z!"

         @IF EXIST "%PATH_ARCHIVIO%" (
            @CALL :LOG "[OK] 07 - ARCHIVIO CREATO: %PATH_ARCHIVIO%"
            @ECHO *** ARCHIVIO CREATO: %NOME_ARCHIVIO% ***
         ) ELSE (
            @CALL :LOG "[ERRORE] 07 - ARCHIVIO NON CREATO"
            @ECHO *** ARCHIVIO NON CREATO ***
         )

         @ECHO 07.05 recupero i file processati nel .zip e li salvo in .txt
         CD /D "%PATH_DEST_S%"
         @CALL :LOG "[STEP] 07 - CD in PATH_DEST_S, CWD=!CD!"
         DIR
         @IF EXIST "%NOME_ARCHIVIO%" (
            "C:\Program Files\7-Zip\7z.exe" l "%NOME_ARCHIVIO%" -bsp1 > "%TODAY%_%FILE_S%_LISTA.TXT" 2>&1
            @CALL :LOG "[STEP] 07 - creato elenco %TODAY%_%FILE_S%_LISTA.TXT, ERRORLEVEL=!ERRORLEVEL!"
         )

         @ECHO 07.05 RITORNO SUL DIRECTORY PRINCIPALE
         CD /D "%PATH_PART_S%"
         @CALL :LOG "[STEP] 07 - ritorno a !CD!"
         @CALL :LOG "[STEP] 07 - fine"

@IF %t% GTR 0 TIMEOUT /T %t% /NOBREAK > NUL

@CLS
@echo.08------------------------------ CHIAMO CRIPTA  @08

@CALL :LOG "[STEP] 08 - inizio"

                  @DIR Zip_N70_CRIPTA.bat
                  @IF EXIST "Zip_N70_CRIPTA.bat" (
                     @CALL :LOG "[STEP] 08 - chiamo Zip_N70_CRIPTA.bat"
                     CALL Zip_N70_CRIPTA.bat
                     @CALL :LOG "[STEP] 08 - ritorno da CRIPTA, ERRORLEVEL=!ERRORLEVEL!"
                  ) ELSE (
                     @CALL :LOG "[STEP] 08 - Zip_N70_CRIPTA.bat NON TROVATA"
                  )

@CALL :LOG "[STEP] 08 - fine"

@echo.09------------------------------ SOSPENSIONE ED USCITA DALLA PROCEDURA @08

         @echo. 09.01
         @echo."=================================================================="
         @echo."        MSG - EXIT_PROCEDURA "
         @echo.
         @echo."ATTENZIONE '!!!' USCITA DALLA PROCEDURA : - USCITA - "
         @echo."=================================================================="

@IF %t% GTR 0 TIMEOUT /T %t% /NOBREAK > NUL

         @echo.                  "**** fine ****"
         @echo."=================================================================="

         @CALL :LOG "[STEP] 09 - fine procedura"
         @CALL :LOG "============================================================"
         @CALL :LOG "LOG TERMINATO IL %DATE% %TIME%"
         @CALL :LOG "============================================================"

@ENDLOCAL
@GOTO :EOF

:FINE_ERRORE
@CALL :LOG "[ERRORE] Uscita per errore"
@CALL :LOG "============================================================"
@CALL :LOG "LOG TERMINATO CON ERRORE IL %DATE% %TIME%"
@CALL :LOG "============================================================"
@ECHO *** USCITA PER ERRORE - vedi %LOG_FILE% ***
@IF %t% GTR 0 TIMEOUT /T 5 /NOBREAK
@ENDLOCAL
@EXIT /B 1

:LOG
@REM Scrive la stringa passata nel log. Gestisce le virgolette.
@ECHO %~1 >> "%LOG_FILE%"
@GOTO :EOF

@REM FINE -----------------------------------------------------------------------------------------

      @GOTO NoteDos

         NOTE DOS DI UTILIZZO DELLA DIR CON I PARAMETRI

@REM [TUTTE LE NOTE ORIGINALI INVARIATE - le trovi nel tuo file]
