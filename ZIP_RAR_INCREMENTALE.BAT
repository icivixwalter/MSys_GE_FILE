@REM BATI DI SALVATAGGIO INCREMENTALE
@REM NOME FILE 			: 			ZIP_RAR_INCREMENTALE.BAT
@REM PATH DI LAVORO  	: 			c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\
@REM ATTIVITA SVOLTA		:
@REM 		la bat preleva i file nell Zip_N70_ElencoFile.txt e attiva il comando 7zip
@REM     per la creazione di file incrementali es. 2026_09_13_XXX_SALVATAGGI
@REM     nella cartella AA_SALVATAGGI
		

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
		@ECHO. cambia questo parametro ---------: 			c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\
		@ECHO. 
		@ECHO. cambia questo parametro ---------: 			_Archivi_Msys_GE_FILE
		@ECHO. 
		@ECHO. 
		@ECHO. 
		@ECHO. cambiando questi 2 parametri si ha effetto di cambiare i valori di tutti i parametri
		@ECHO. dos preimpostati come di seguito indicato:
		@ECHO. 
		@ECHO. 									effetto sostituzione collettivo
		@ECHO. --------------------------------------------------------------------------------------------------------
		@ECHO. LA PATH DI DESTINAZIONE PER I SALVATAGGI 	=	c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\%CARTELLA_SALVATAGGI_S%
		@ECHO. LA CARTELLA SALVATAGGI CREATA O ESISTENTE 	=  	c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\AA_SALVATAGGI\
		@ECHO. LA PATH DI PARTENZA							= 	c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\
		@ECHO. FILE 7ZIP, RAR DI SALVATAGGIO 				= 	ZIP_SALVATAGGI_Archivi_Msys_GE_FILE
		@ECHO. LISTA FILE UTILIZZATA 						= 	Zip_N70_ElencoFile.txt
		@ECHO. --------------------------------------------------------------------------------------------------------
		@ECHO.


	@CLS
	@echo.01--------------------------------CREO LA CARTELLA SALVATAGGI SE NON ESISTE
			
				@REM Se non esiste la cartella AA_SALVATAGGI, la crea.
				if not exist "%cd%\AA_SALVATAGGI\" (
					echo Crea la cartella AA_SALVATAGGI
					mkdir %cd%\AA_SALVATAGGI

					@ECHO  ATTENZIONE HO CREATO LA CARTELLA AA_SALVATAGGI : 
					@ECHO 
					@ECHO 
					@ECHO 
					@ECHO 
						DIR  AA_SALVATAGGI


						@REM SOSPENSIONE TEMPO IMPOSTATA CON T
										@REM ---->	timeout /t 2 /nobreak > NUL
										@REM ---->	
										@TIMEOUT /T %t% /NOBREAK
										@REM Pause

				
				)

					@REM CREATA OD ESISTENTE VIENE SALVATA LA CARTELLA SALATAGGI NELLA VARIABILE
					SET CARTELLA_SALVATAGGI_S=AA_SALVATAGGI


	@CLS
	@echo.02--------------------------------TITOLO COSTRUZIONE DI UN BAT INCREMENTALE + IMPOSTAZIONE PARAMETRI
	
				@REM 	NOTE GENERALI
					@REM	CAMBI QUESTI PARAMETRI PER TUTTO IL FILE
					@REM	questo esempio di bat incrementale si trova i questa path dello 7Z zip questi sono parametri
					@REM	che possono essere sostituiti da qui, la PATH DI DESTINAZIONE + LA PATH DI PARTENZA  + 
					@REM	IL FILE :	
					@REM	@parametri_(per la @sostituzione dei parametri per @tutto il @progetto)
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


				@REM IMPOSTA I PARAMETRI DELLE VARIABILI DI AMBIENTE PRINCIPALI		
					@REM ................................................................................
					@REM  @imposto@i@parametri_(IMPOSTA I PARAMETRI TEMPO, PATH ec...)

							@REM IMPOSTO LA VARIABILE PER LA SOSPENSIONE DEL TEMPO
								@REM .......................................................		
									@REM ---->	MODELLO timeout /t 2 /nobreak > NUL
									@REM        TIMEOUT /T %t% /NOBREAK
									@REM ---->	
									@SET t=0
									@echo Sospensione tempo per secondi: %t%
								@REM .......................................................		

								@REM VISUALIZZO LA CARTELLA SALVATAGGI IMPOSTATATA AL PUNTO 1)
								@REM .......................................................		
									@ECHO. 
									@ECHO. La cartella salvataggi è la seguente
									@echo --------------------------------------> : %CARTELLA_SALVATAGGI_S%
								@REM .......................................................		
				
									

							@REM CICLO FOR F PER TROVARE E SALVARE LA PATH CORRENTE
								@REM .......................................................		
									@for %%f in ("CurrentDIR.bat") DO @SET currDir_=%%~df%%~pf
									@REM alla fine del ciclo salvo la path corrente
									@SET PATH_CORR_S=%currDir_%
								@REM .......................................................


							@REM imposto le variabili

							@SET PATH_DEST_S=c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\%CARTELLA_SALVATAGGI_S%
							@REM @SET PATH_DEST_S=c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\AA_SALVATAGGI\
							@SET PATH_PART_S=c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\
							
							@SET FILE_S=ZIP_SALVATAGGI_Archivi_Msys_GE_FILE
							@SET LISTA_FILE_S=Zip_N70_ElencoFile.txt

							
					@REM ................................................................................
						
					@REM NOTE DI UTILITA
						@REM ................................................................................
						
							@REM PROGETTO ZIP E 7ZIP VEDI DEFINIZIONI .BAT IN TUTTI I PROGETTI

								@REM PER COSTRUIRE UN BAT INCREMENTALE DENOMINATO 7ZIP_INCREMENTALE.BAT 
								@REM occorre fare riferimento a questo progetto .BAT che si trova nel TUTORIAL ZIP E 7ZIP.

								@REM Per individuare il progetto del bat incrementale occorre far riferimento al progetto 
								@REM .MD  ZIP che si trova in questa path. Ogni modifica del progetto BAT incrementale 
								@REM nelle sue istruzioni occorre riportarla in questa progettazione:


								@REM TUTORIAL ZIP RAR O 7ZIP 
									@REM @DOVE@SI@TROVA_(@7zip dove si trova il 7z con il tutorial)_@7z
									@REM	apro il dos con il power shell 
									@REM 	  @APRI_@COMANDI_@7z_(la path ed il comando per aprire il progetto 7zip)
									@REM	START "apri il progetto dei comandi PER IL TUTORIAL DI 7z" "c:\\CASA\\LINGUAGGI\ZIP\\ZIP.sublime-project" 
										
					
									@REM Apro tutti i progetti con powershell
									@REM  START "apri il progetto dei comandi di TUTTI@PROGETTI" "c:\\CASA\\LINGUAGGI\\TROVA_PROGETTI\\TROVA_PROGETTI.sublime-project"


							@REM	SUBLIMETEXT 

								@REM	Per avere un tutorial veloce in sublimetex utiliza questo link
								@REM	https://www.w3schools.io/editor/sublime-text-terminal/
								

								@REM SOSPENSIONE TEMPO IMPOSTATA CON T
								@REM ---->	timeout /t 2 /nobreak > NUL
								@REM ---->	
								@TIMEOUT /T %t% /NOBREAK
								@REM Pause

						@REM ................................................................................			

    @CLS
	@echo.03--------------------------------CONTROLLO DEI PARAMETRI_GENERALI, DISCO, PATH DI DESTINAZIONE, NOME ARCHIVIO.ZIP NOME LISTA FILEeTorri
				@ECHO.03.02 STAMPA DI CONTROLLO DEL SETTAGGIO DELLE VARIABILI
				@ECHO SET PATH DI DESTINAZIONE + PATH PARTENZA + FILE DI SALVATAGGIO + LISTA FILE 				
				@ECHO. 
				@ECHO. controllo PATH di partenza...........PATH_PART_S  		  :%PATH_PART_S%
				@ECHO. controllo PATH di DESTINAZIONE.......PATH_DEST_S  		  :%PATH_DEST_S%
				@ECHO. controllo nome file Zip..............FILE_S  			  :%FILE_S%
				@ECHO. controllo elenco file salvati........LISTA_FILE_S 		  :%LISTA_FILE_S%
				@ECHO. controllo della path corrente........PATH_CORR_S 		  :%PATH_CORR_S%
				@ECHO. controllo della cartella ............CARTELLA_SALVATAGGI_S :%CARTELLA_SALVATAGGI_S%
				@ECHO. 


								@REM SOSPENSIONE TEMPO IMPOSTATA CON T
								@REM ---->	timeout /t 2 /nobreak > NUL
								@REM ---->	
								@TIMEOUT /T %t% /NOBREAK
								@REM Pause
    
	
    @CLS
	@echo.04-------------------------------TIPO FILE DA ARCHIVIARE SE NON ESISTE LA LISTA FILE CON VARIE ESTENSION, *.DOC, *.MDB ecc...

				@ECHO.04.01                      RAGGRUPPAMENTO DI TIPI DI FILE *.DOC, *.XLS EC...
				@ECHO.
				@ECHO. @tipi@di@file_(da salvare raggruppati in una variabile per ora SONO BLOCCATI )
				@ECHO.
				@ECHO. imposta i file da salvare su diverse variabili (ATTENZIONE AGLI SPAZI DOPO VARIABILE=*.XLS ... NO CI DEVONO ESSERE
				@ECHO. ATTENZIONE ALL'ERRRORE NELLA VARIABILE inserendo lo spazio tra uguale e variabile (es. ====== TipoFile_Office_S =*.MDB)
				@ECHO. deve essere tutto contiguo senza spazi (es. corretto ====== TipoFile_Office_S=*.MDB )
				@ECHO. 
				@ECHO. 
				
				@SET TipoFile_TEXT_S=*.AUX *GZ *.LOG *.OUT *.TEX *.Tws
				@SET TipoFile_JAVA_S=*.CLASS *.EXE *HTLM *.IML *.JAR *.JAVA *.LOG *.MF *.LNK *.XML
				@SET TipoFile_TXT_S=*.TEX *.TWS *.TXT *BAT
				@SET TipoFile_Office_S=*.MDB *.XLS *.PDF

				@echo 04.1 ------------------------CONCATENA controlla ED ESEGUI
				@echo 
				@SET TIPO_FILE_S=%TipoFile_TEXT_S% %TipoFile_JAVA_S% %TipoFile_TXT_S% %TipoFile_Office_S%

				@ECHO.	
				@ECHO. "--------------------------------------------------"
				@ECHO. "ELENCO FILE CON ESTENSIONE DA SALVARE "
				@ECHO. controllo lista file della variabile salvata in tipo file al momento non utilizzata
				@ECHO. %TIPO_FILE_S%
				@ECHO. 
				@ECHO. 
				@ECHO. 

				

								@REM SOSPENSIONE TEMPO IMPOSTATA CON T
								@REM ---->	timeout /t 2 /nobreak > NUL
								@REM ---->	
								@TIMEOUT /T %t% /NOBREAK
								@REM Pause
					
	@CLS 
	@ECHO.05------------------------------- CICLO FOR CON GESTIONE DELLA DATA CORRENTE DA RECUPERARE SOTTO FORMA DI AAA MM GG @04
			@ECHO.
			@ECHO.
			@ECHO.
			@ECHO.05.01 		CICLO FOR I^ E II^ CASO PER LA GESTIONE DELLA DATA
			@ECHO.
			@ECHO ========================================================================================================================
			
			:----------------------------CICLO FOR  I^ CASO 
				@ECHO  CICLO FOR I^ CASO per la gestione della data con le sottostringhe ma aggiunge lo 0 se <10
				
				@for /f "skip=1" %%x in ('wmic os get localdatetime') do if not defined MyDate set MyDate=%%x
				
				@echo solo L'ANNO che viene separato per la futura ricerca:
				@echo "inserito nella variabile Mydate  		---> " %MyDate:~0,4%
				
				:---------------------------imposto la variabile dell'anno + MESE + GIORNO
				@ECHO
				@ECHO.                IMPOSTO LE VARIABILI ANNO, MESE GIORNO
				@ECHO
				@ECHO
				@SET ANNO_S=%MyDate:~0,4%
				@SET MESE_S=%MyDate:~4,2%
				@SET GIORNO_S=%MyDate:~6,2%
				@echo "controllo nella variabile ANNO_S 		---> %ANNO_S%"
				@echo "controllo nella variabile MESE_S 		---> %MESE_S%"
				@echo "controllo nella variabile GIORNO_S 		---> %GIORNO_S%"
				@echo 
				@echo solo il giorno:
				@echo %MyDate:~6,2%
				@echo DATA RICOSTRUITA con le sottostringhe con separatore l'undescore (_)
				@set today=%MyDate:~0,4%_%MyDate:~4,2%_%MyDate:~6,2%
				@echo.
				@echo today in formato stringa: 
				@echo %today%
				@echo.

				@REM -----> PAUSE


								@REM SOSPENSIONE TEMPO IMPOSTATA CON T
								@REM ---->	timeout /t 2 /nobreak > NUL
								@REM ---->	
								@TIMEOUT /T %t% /NOBREAK
								@REM Pause

			

			@CLS		
			:----------------------------CICLO FOR  II^ CASO ciclo for per per la data AAA MM GG senza sottostringhe

				@ECHO E' possibile ottenere la data corrente in modo indipendente dalle impostazioni locali utilizzando
				@ECHO SENZA armeggiare con le sottostringhe
				@echo vedi il link: https://qastack.it/programming/10945572/windows-batch-formatted-date-into-variable
				@echo .

				@echo. II CASO la data senza le sottostringhe solo numerico e con separatore il trattino (-)

				@REM for /f %%x in ('wmic path win32_localtime get /format:list ^| findstr "="') do set %%x
				@REM set today=%Year%_%Month%_%Day%

				@echo. IL GIORNO DEFINITO IN VARIABILE
				@ECHO %TODAY%
				@echo.

			@ECHO 	CICLO FOR I E II CASO PER LA GESTIONE DELLA DATA  *** FINE ***
			@ECHO ========================================================================================================================
			@ECHO. 
			@ECHO. 
			@ECHO. 
			@ECHO. 

			   
								@REM SOSPENSIONE TEMPO IMPOSTATA CON T
								@REM ---->	timeout /t 2 /nobreak > NUL
								@REM ---->	
								@TIMEOUT /T %t% /NOBREAK
								@REM Pause

	@CLS 
	@echo.06-------------------------------CHIAMO DECODIFICA @06
		
						 
						
			
						

						@CLS
						@ECHO.
						@ECHO. 06.02 CONTROLLO VARIABILE ANNO_S
						@echo "controllo VALORE anno nella variabile ANNO_S 	--->  %ANNO_S%"
						@echo .

							@REM @ECHO. 05.03 FACCIO IL RENAME se esistono i file e controllo con dir
							@REM @echo  chiamando la procedura bat decodifica 
								@rem @RENAME "%ANNO_S%*.GE614" "%ANNO_S%*.ZIP"
								@rem @RENAME "%ANNO_S%*.GE614" "%ANNO_S%*.RAR"

							 
						@CLS
						@ECHO. ===================================================================	
						@ECHO. 			
						@ECHO. 06.04 CHIAMO LA BAT DECODIFICA
						@ECHO.  se esiste la .bat  la eseguo., mi sposto nella path SALVATAGGI
						@ECHO. 

							
							@DIR Zip_N70_DECODIFICA.bat
							@CALL Zip_N70_DECODIFICA.bat
							

				
						@ECHO. 
						@ECHO. 
						@ECHO. 
						@ECHO. ===================================================================	
						
						

								@REM SOSPENSIONE TEMPO IMPOSTATA CON T
								@REM ---->	timeout /t 2 /nobreak > NUL
								
								@TIMEOUT /T %t% /NOBREAK
								@REM 
								@REM ----> PAUSE	
    
    							@ECHO. 

    @CLS
	@echo.07------------------------------ SALVO TUTTI I FILE NELLE SOTTOCARTELLE /S con 7zip uso u -r  @06
		

				@ECHO. 
				@ECHO. 
				@ECHO.					CONTROLLO CHE SIAMO NELLA PATH DI PARTENZA @06
				@ECHO. 
				@ECHO. "@echo.07.01--------------------------------------------------"
				@ECHO. 
				@ECHO. "faccio il dir sulla directory corrente per vedere in quale"
				@ECHO. "posizione mi trovo per vedere se siamo nella path di destinazione"
				@ECHO. "In questo caso torno indieto di una posizione  ed effettuo un controllo"
				@ECHO. 
				@ECHO  "attivo questi comandi: DIR W per controllare la cartella dei salvataggi"	
				@ECHO. 
				@ECHO.   CONTROLLO DELLA VARIABILE PATH DI DESTINAZIONE E DELLA PATH DI PARTENZA
				@ECHO.   E MI POSIZIONO NELLA PATH DI PARTENZA PER EFFETTUARE I SALVATAGGI E CONTROLLO 
				@ECHO.   LA DIR CORRENTE
				@ECHO. 	 %PATH_DEST_S%
				@ECHO. 	 %PATH_PART_S%
			
							CD %PATH_PART_S%
							DIR  
							
					
				@ECHO.  
				@ECHO. 
				@ECHO. 
				@ECHO. 


				@ECHO.					SALVATAGGI CON Z7.ZIP
				@ECHO. "@echo.07.02--------------------------------------------------"
				@ECHO. 
				@ECHO qui effettuo il salvataggio del .zip con il nome DATA CORENTE in formato  AAA MM GG
				@ECHO. %PATH_DEST_S%
				@ECHO. %TODAY%
				@ECHO. _%FILE_S%
				

				@ECHO. ATTIVO IL SALVATAGGIO CON 7ZIP 
				@ECHO.  
				"C:\Program Files\7-Zip\7z.exe" u -tzip -r %PATH_DEST_S%\%TODAY%_%FILE_S% @Zip_N70_ElencoFile.txt
				@ECHO. 
				@ECHO. 
				
			
				@ECHO.					SALVATAGGI ELENCO FILE IN .TXT
				@ECHO. "@echo.07.03--------------------------------------------------"
				@ECHO  
				@ECHO  qui salvo i file processati NEL .ZIP E CHE SI TROVA NELLA CARTELLA SALVATAGGI 
				@ECHO  quindi 01) devo prima spostarmi di directory dei salvataggi; 
				@ECHO  quindi 02) recupero la lista dei file salvati nel .zip con il commando "l TODAY_FILE_S " e 
				@ECHO  l'ozione -bsp1  ">" lo invio nel file .txt ricostruito.

				@ECHO attenzione se 7z.exe non ha la variabile di ambiente riportare il percorso completo 
				@ECHO tendo presente che è stato salvato in Program File ed ha lo spazio in mezzo per cui 
				@ECHO occorre inserire tutto il percorso assoluto dell'eseguibile  tra "" altrimenti non lo
				@ECHO legge l'eseguibile e da errore. CONTROLLARE SEMPRE che l'eseguibile si trovi in quella path.
				@ECHO 

				
					
					@ECHO 07.05 recupero i file processati nel .zip e li salvo in .txt
					@ECHO VADO NELLA PATH DI DESTINAZIONE E SALVO L'ELENCO DEI FILE 
					CD %PATH_DEST_S%
					@ECHO IL FILE %TODAY%_%FILE_S%
					DIR
					"c:\Program Files\7-Zip\7z.exe" l %TODAY%_%FILE_S%.* -bsp1 > %TODAY%_%FILE_S%_LISTA.TXT
					

				@ECHO.  					
				@ECHO.  					
				@ECHO.  	

				@ECHO 07.05 RITORNO SUL DIRECTORY PRINCIPALE
					@ECHO 
					@CD ..
									

						
		   
								@REM SOSPENSIONE TEMPO IMPOSTATA CON T
								@REM ---->	timeout /t 2 /nobreak > NUL
								@REM ---->	
								@TIMEOUT /T %t% /NOBREAK
								@REM Pause

								@ECHO.
	
	@CLS
	@echo.08------------------------------ CHIAMO CRIPTA  @08
					
					@ECHO  					ESEGUO IL CRIPTA.BAT
					@ECHO .  -----------------------------------------------------------------------
					@ECHO . 
					@ECHO . 08.01. E ATTIVO IL CRIPTAGGIO !!!
					@ECHO . 
					

							@DIR Zip_N70_CRIPTA.bat
							@CALL Zip_N70_CRIPTA.bat
							

	
	@echo.09------------------------------ SOSPENSIONE ED USCITA DALLA PROCEDURA @08

				
				
				@echo.
				@echo.
				@echo. 09.01
				@echo."=================================================================="
				@echo."			MSG - EXIT_PROCEDURA "
				@echo.
				@echo."ATTENZIONE '!!!' USCITA DALLA PROCEDURA : - USCITA - "			
				@echo."=================================================================="

				
		   
								@REM SOSPENSIONE TEMPO IMPOSTATA CON T
								@REM ---->	timeout /t 2 /nobreak > NUL
								@REM ---->	
								@TIMEOUT /T %t% /NOBREAK
								@REM Pause

				@echo.								
				@echo.								
				@echo.						"**** fine ****"
				@echo."=================================================================="
				

				EXIT

		






@REM FINE -----------------------------------------------------------------------------------------

		@GOTO NoteDos

			NOTE DOS DI UTILIZZO DELLA DIR CON I PARAMETRI


				@REM PARAMETRI DOS
				@REM ------------------------------------------------------------------------------------------- 
				
					@REM DIR in dos ha diversi parametri per visualizzare le cartelle ed i file che sono:
					@REM 		:  /A:-d mostra solo i file    
					@REM 		/O:n  mostra i file in ordine alfabetico   
					@REM		/W  per vedere l’elenco formato solo dai nomi dei file 
					@REM		BLOCCATO -----> DIR /S /A:-D /O:n /W %TIPO_FILE_S%> %PATH_DEST_S%%today%_ZIP_SALVATAGGI_ARCHIVI_ARCHIVIO_TIPO_FILE_SALVATI.TXT


					@REM -tzip 	=	parametro -t per tipo file processato cioè il file zip.
					@REM -r 	=	iterazione delle sottocartelle
					@REM @file 	=	lista file da processare
					@REM 		 	attenzione nella listafile.txt indicare la path e al posto di *.* occore solo * es. c:\dos\* pero per riportare la path completa
					@REM 		 	della sottodirectory occorre il parametro -r = sottodirectory e nel file .txt indicare la path completa senza * es.
					@REM 		 	c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\
					@REM path   =	la path di salvataggio, per posizionare l'archivio in un'altra cartella occore scrivere il nome 
					@REM 		 	Dell'archivio preceduto dalla path estesa es.
					@REM 			C:\CASA\SALVATAGGI\archive.zip
					@REM -bsp1	= swithc che permette il reindirizzamento dell'output su file per l'elenco di tutti 
					@REM 		quelli processati e salvati su un file .txt, ma deve essere eseguito su doppio comando.
					@REM NOME DATA CORRENTE : per il nome sulla data corrente occorre utilizzare la variabile tmp
					@REM 			che contiene il nome in AAMMGG_, Attenzione per espandere la variabile occorre utilizzare ! ! altrimenti non funziona;
					@REM 			per ora la espanzione delle variabile con (!!) non è stata utilizzata 
					@REM 			es. %!PATH_DEST_S!%%today%_ZIP_SALVATAGGI_ARCHIVI.RAR %TIPO_FILE_S% %PATH_DEST_S%

					@REM  ATTENZIONE in @Zip_N70_ElencoFile.txt  devi inserire solo le directory senza *.* e no le cartelle interne perchè altrimenti
					@REM 				salva solo le sottocartelle e non la cartella madre. Quindi nel file.txt indicare le cartelle ORIGINE in questo
					@REM  				modo: c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS\MSys_GE_FILE\
					@REM  				modo: c:\CASA\LINGUAGGI\CPP\  ... ecc.




				@REM DIR @ORDINAMENTO_ALFABETICO_DECRESCENTE_PER_DATA_DI_CREAZIONE
				@REM ------------------------------------------------------------------------------------------- 
				@REM Sintassi 	DIR  [unità:][percorso][nomefile] [/P] [/W] [/O[[:]ordinamento]] [/S]
				@REM      	[/B] [/L] [/V]
					

				@REM  /A          Visualizza i file con gli attributi specificati.
				@REM  attributo	a:-d  (-d) = non le directory
				@REM		-h = non i file nascosti
				@REM /O		 Elenca file ordinandoli.
				@REM attributo	o:-n (-n) = escluso ordine alfabetico
				@REM /T		Controlla campo data visualizzato o utilizzato per ordinare.
				@REM attributo   t:c (:c)= ordina per creazione

					
					@REM ORDINA tutti i file escluso le directory ed i file nascosti in ordine alfabetico
					@REM decrescente e per data di creazione e salva nel file di testo.
					dir "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\" /a:-d -h /o:-n /t:c >>DOS_DIR_N01.txt

				@REM DIR @ORDINAMENTO_ALFABETICO_DECRESCENTE_PER_DATA_DI_CREAZIONE *** fine ***
				@REM ----------------------------------------------------






				@REM				PARAMETRI DOSE CON VARI ESEMPI DI COMANDO 
				@REM *************************************************************************************************************************
						@OPZIONI@DOS@DI_(ELENCO DI OPZIONI del comando DIR in dos)


					le opzioni principali sono: 
					/p per vedere l’elenco a schermate 
					/w per vedere l’elenco formato solo dai nomi dei file 
					/a:h mostra i file nascosti 
					/a:s mostra i file di sistema 
					/a:d mostra solo le subdirectory 
					/a:-d mostra solo i file 
					/o:n mostra in ordine alfabetico 
					/o:-n mostra in ordine alfabetico inverso 
					/o:d mostra in ordine di data 
					/o:-d mostra in ordine di data inverso 
					/o:s mostra in ordine di grandezza 
					/o:-s mostra in ordine di grandezza inverso 
					/s mostra tutto il disco 
					>prn stampa l’elenco 
						Esempio: 
						c:\> DIR c:\LETTERE\ROSSI /a:-d/o:n >prn

						stampa l’elenco, in ordine alfabetico, di tutti i file contenuti nella 
						subdirectory \LETTERE\ROSSI, omettendo i nomi delle eventuali subdirectory presenti. 




						@REM DIR @ORDINAMENTO_ALFABETICO_DECRESCENTE_PER_DATA_DI_CREAZIONE
						@REM ------------------------------------------------------------------------------------------- 
						@REM Sintassi 	DIR  [unità:][percorso][nomefile] [/P] [/W] [/O[[:]ordinamento]] [/S]
						@REM      	[/B] [/L] [/V]


						@REM  /A          Visualizza i file con gli attributi specificati.
						@REM  attributo	a:-d  (-d) = non le directory
						@REM			-h = non i file nascosti
						@REM /O		 Elenca file ordinandoli.
						@REM attributo	o:-n (-n) = escluso ordine alfabetico
						@REM
						@REM /T		Controlla campo data visualizzato o utilizzato per ordinare.
						@REM attributo   t:c (:c)= ordina per creazione


							@REM ORDINA tutti i file escluso le directory ed i file nascosti in ordine alfabetico
							@REM decrescente e per data di creazione e salva nel file di testo.
							dir "c:\GESTIONI\GESTIONE_LLPP\02_SCANNER\" /a:-d -h /o:-n /t:c >>DOS_DIR_N01.txt

						@REM DIR @ORDINAMENTO_ALFABETICO_DECRESCENTE_PER_DATA_DI_CREAZIONE *** fine ***
						@REM ----------------------------------------------------





						@REM DOS STRINGHE CONCATENA  @ESEMPIO@DOS@CONCATENA_(concatena le stringhe in dos)
						@REM ----------------------------------------------------

							:-------------------------------imposto le variabili set
							set str1="Hello"
							set str2="World"
							set str3= "1"

							:-------------------------------unisco le stringhe su una variabile 
							set "str3=%str1% %str3% %str2%"
							set "str4=%str1% %str2%"

							@REM qui scrivo str1 + la parola dos  + str2
							set "str1=%str1% DOS %str2%"

							:-------------------------------controllo concatenamento
							echo "str3 concatenata ->" %str3%
							echo "str4 concatenata ->" %str4%
							echo "str1 concatenata ->" %str1%




						@REM DOS STRINGHE CONCATENA	*** fine ***
						@REM ----------------------------------------------------





				:NoteDos

				@REM *************************************************************************************************************************

