'==========================================================
'00_NAME_MODULO = Msys_GEFILE_Mdl02_ESPORTA_OGGETTI
   '//PATH_PROGETTO: _
      'START "APRI PROGETTO CLASSE" "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_GE_FILE\MDB\MSys_GE_FILE.mdb"
      'ATTIVITA SVOLTE'
      '01_IMPOSTAZIONI_MODULO
      '02_ESPORTAZIONE_OGGETTI_GEFILE_INIZIO
      '03_CARICAMENTO_PERCORSI_DB_ESTERNI
      '04_CARICAMENTO_OGGETTI_DA_ESPORTARE
      '05_CONTROLLO_ESISTENZA_OGGETTO_DB_ESTERNO
      '06_CANCELLAZIONE_OGGETTO_DB_ESTERNO
      '07_CONTROLLO_ESISTENZA_OGGETTO_DB_ORIGINE
      '08_LOG_CSV_SEI_CAMPI
      '09_CANCELLA_OGGETTI_DB_ESTERNO_FUNCT
'==========================================================

'==========================================================
'01_IMPOSTAZIONI_MODULO
      Option Compare Database
      Option Explicit

       '//dim variabili generali per il log di cancellazione
       Private logFilePath As String
'==========================================================

'==========================================================
'02_ESPORTAZIONE_OGGETTI_GEFILE_INIZIO      
'==========================================================
         '// ATTIVITA ESEGUITE:
               '// - gestisce più percorsi (sxPath1, sxPath2, ecc.)
               '// - usa Collection numeriche per i tipi oggetto (1, 2, 3, …)
               '// - verifica esistenza oggetto prima dell’esportazione
               '// - cancella preventivamente gli oggetti già presenti nel DB esterno
               '// - scrive tutto su un file di log (oggetti OK, errori, mancanti)
               '// - salta in sicurezza se una Collection è vuota
               '// - mostra un messaggio riepilogativo finale
               '//----------------------------------------------------------------------------'
               '// FILTRO NUMERICO APPLICATO PER I DIVERSI OGGETTI:
               '// Tipo numerico  Costante VBA   Oggetto Access
               '// 1  acQuery    Query
               '// 2  acForm     Form
               '// 3  acReport   Report
               '// 4  acMacro    Macro
               '// 5  acModule   Modulo
               '// 6  acTable    Tabella

   Public Function Msys_GEFILE_Mdl01_ESPORTA_OGGETTI_Pfnc()
          On Error GoTo GestErr
          
          '//DIM VARIABILI **** INIZIO ****
          '/collection
          Dim colOggetti As Collection
          Dim colPercorsi As Collection
          Dim itm As Variant, Obj_v As Variant
          '//PER IL FILE LOG
          Dim fso As Object, logFile As Object
          Dim logFileOggetti As Object
          Dim exportNro As Long
          Dim logPath As String
          Dim startTime As Date
          Dim totaleOk As Long, totaleErr As Long
          Dim tipo As Integer, nome As String
          Dim Path_s As String
          Dim dbDest As DAO.Database
          Dim NomeDB As String
          
          MsgBox "ESPORTAZIONE OGGETTI GE_FILE IN TUTTI I DB ACCESS", vbExclamation, "ESPORTO OGGETTI" 
                   
          '=== Creo il FileSystemObject
          Set fso = CreateObject("Scripting.FileSystemObject")

          '=== Creo la cartella Logs se non esiste
          If Not fso.FolderExists(CurrentProject.Path & "\Logs") Then
              fso.CreateFolder CurrentProject.Path & "\Logs"
          End If

          '=== Log dettagliato oggetti esportati
          Set logFileOggetti = fso.CreateTextFile(CurrentProject.Path & "\Logs\ExportOggetti.txt", True)
          '=== CORREZIONE: intestazione a 6 campi con Timestamp
          logFileOggetti.WriteLine "Nro;NomeOggetto;Tipo;PathDestinazione;DBDestinazione;Timestamp"
          exportNro = 0

          '=== Log principale
          DoEvents  ' per il controllo del ciclo'
          startTime = Now()
          logPath = CurrentProject.Path & "\Logs\ExportLog.txt"
          Set logFile = fso.CreateTextFile(logPath, True)
          logFile.WriteLine String(60, "=")
          logFile.WriteLine "LOG ESPORTAZIONE OGGETTI ACCESS"
          logFile.WriteLine "Database origine: " & CurrentDb.Name
          logFile.WriteLine "Data/Ora inizio: " & startTime
          logFile.WriteLine String(60, "=")
          logFile.WriteLine ""

          '=== Precarico le path dei DB esterni
          Set colPercorsi = CaricaPercorsiEsportazione()

          '=== Precarico gli oggetti da esportare
          Set colOggetti = CaricaOggettiDaEsportare()

          '=== Verifiche preliminari
          If colPercorsi.Count = 0 Then
              MsgBox "Nessun percorso di destinazione definito. Esportazione annullata.", vbExclamation
              logFile.WriteLine "?? Nessun percorso di destinazione definito. Esportazione annullata."
              GoTo Esci
          End If

          If colOggetti.Count = 0 Then
              MsgBox "Nessun oggetto da esportare. Esportazione annullata.", vbExclamation
              logFile.WriteLine "?? Nessun oggetto da esportare. Collection vuota."
              GoTo Esci
          End If

          '==========================================================
          'CICLO PRINCIPALE DI ESPORTAZIONE
               
          For Each Obj_v In colPercorsi
              'Dim + Set variabili'  
              Path_s = Obj_v
              logFile.WriteLine vbCrLf & String(50, "-")
              logFile.WriteLine "DESTINAZIONE: " & Path_s
              logFile.WriteLine String(50, "-")
              If Not fso.FileExists(Path_s) Then
                  logFile.WriteLine "? ATTENZIONE: file non trovato -> " & Path_s
                  Debug.Print "ATTENZIONE: file non trovato -> " & Path_s
                  GoTo NextPath
              End If
              '=== Apro il database esterno
              Set dbDest = DBEngine.OpenDatabase(Path_s)
              '=== CORREZIONE: assegno NomeDB per log dettagliato
              NomeDB = dbDest.Name

              '=== Cancellazione preventiva degli oggetti già presenti
              For Each itm In colOggetti
                  tipo = CLng(Split(itm, "|")(0))
                  nome = Split(itm, "|")(1)
                  If OggettoEsisteNelDB(dbDest, tipo, nome) Then
                      On Error Resume Next
                      CancellaOggetto dbDest, Path_s, tipo, nome
                      If Err.Number <> 0 Then
                          logFile.WriteLine "ERRORE cancellazione oggetto: " & nome & " (" & Err.Description & ")"
                          Err.Clear
                      End If
                      On Error GoTo GestErr
                  End If
              Next itm

              '=== Esporto gli oggetti
              For Each itm In colOggetti
                  tipo = CLng(Split(itm, "|")(0))
                  nome = Split(itm, "|")(1)

                  '=== Controllo se l'oggetto esiste nel DB origine
                  If OggettoEsiste(tipo, nome) Then
                      On Error Resume Next
                      '=== Copia oggetto nel DB esterno
                      DoCmd.CopyObject Path_s, "", tipo, nome
                      If Err.Number = 0 Then
                          totaleOk = totaleOk + 1
                          exportNro = exportNro + 1

                          '=== Scrittura nel log generale
                          Debug.Print "OK -> " & nome
                          logFile.WriteLine "OK  -> " & nome

                          '=== Scrittura nel log dettagliato ExportOggetti.txt
                          logFileOggetti.WriteLine exportNro & ";" & nome & ";" & tipo & ";" & Path_s & ";" & NomeDB & ";" & Format(Now, "yyyy-mm-dd hh:nn:ss")
                      Else
                          totaleErr = totaleErr + 1
                          Debug.Print "ERRORE -> " & nome & " (" & Err.Description & ")"
                          logFile.WriteLine "ERRORE -> " & nome & " (" & Err.Description & ")"

                          '=== Scrittura nel log dettagliato anche in caso di errore
                          exportNro = exportNro + 1
                          logFileOggetti.WriteLine exportNro & ";" & nome & ";" & tipo & ";" & Path_s & ";" & NomeDB & " - ERRORE: " & Err.Description & ";" & Format(Now, "yyyy-mm-dd hh:nn:ss")
                      End If
                      Err.Clear
                      On Error GoTo GestErr
                  Else
                      totaleErr = totaleErr + 1
                      Debug.Print "NON TROVATO -> " & nome
                      logFile.WriteLine "NON TROVATO -> " & nome

                      '=== Scrittura nel log dettagliato anche per oggetti non trovati
                      exportNro = exportNro + 1
                      logFileOggetti.WriteLine exportNro & ";" & nome & ";" & tipo & ";" & Path_s & ";" & NomeDB & " - NON TROVATO;" & Format(Now, "yyyy-mm-dd hh:nn:ss")
                  End If
              Next itm

              dbDest.Close
      NextPath:
          Next Obj_v

          '=== Chiusura log dettagliato
          logFileOggetti.Close

      Esci:
          logFile.WriteLine vbCrLf & "Totale OK: " & totaleOk & " - Errori: " & totaleErr
          logFile.WriteLine "Data/Ora fine: " & Now()
          logFile.WriteLine String(60, "=")
          logFile.Close

          MsgBox "? Esportazione completata!" & vbCrLf & _
                 "Oggetti OK: " & totaleOk & vbCrLf & _
                 "Errori o non trovati: " & totaleErr & vbCrLf & _
                 "Log dettagliato: " & CurrentProject.Path & "\Logs\ExportOggetti.txt", vbInformation, "RISULTATO ESPORTAZIONE"

          Exit Function

      GestErr:
          '=== Chiusura sicura dei log in caso di errore
          If Not logFile Is Nothing Then
              logFile.WriteLine "ERRORE FATALE: " & Err.Description
              logFile.WriteLine "Data/Ora: " & Now()
              logFile.WriteLine String(60, "=")
              logFile.Close
          End If
          If Not logFileOggetti Is Nothing Then
              logFileOggetti.WriteLine "ERRORE FATALE: " & Err.Description
              logFileOggetti.WriteLine "Data/Ora: " & Now()
              logFileOggetti.Close
          End If
          MsgBox "Errore durante l’esportazione: " & Err.Description, vbCritical, "ERRORE"
   End Function

'==========================================================
'03_CARICAMENTO_PERCORSI_DB_ESTERNI: Caricamento percorsi DB esterni 
      Private Function CaricaPercorsiEsportazione() As Collection
             On Error GoTo GestErr
                 
             Dim colPercorsi As New Collection
             doEvents  ' per il controllo del ciclo'
             '=== Aggiungere qui i percorsi dei database di destinazione
             colPercorsi.Add "C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_ATTIVA_GEST_OGGETTI\MDB\MSYS_ATTIVA_GEST_OGGETTI.mdb"
             colPercorsi.Add "C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_BASE\MDB\MSYS_BASE_MODELLO.mdb"
             '... aggiungere altri percorsi come necessario

             Set CaricaPercorsiEsportazione = colPercorsi
             Exit Function

            GestErr:
             MsgBox "Errore in CaricaPercorsiEsportazione: " & Err.Description, vbCritical
             Set CaricaPercorsiEsportazione = Nothing
      End Function
'==========================================================

'==========================================================
'04_CARICAMENTO_OGGETTI_DA_ESPORTARE 
      Private Function CaricaOggettiDaEsportare() As Collection
          On Error GoTo GestErr
          Dim colOggetti As New Collection

          '=== Formato: (tipo, nome)
          ' 1 = Query, 2 = Form, 3 = Report, 4 = Macro, 5 = Modulo, 6 = Tabella
          
           'QUERY codice = 1
              colOggetti.Add "1|UTILITA_Msys_Qry21_00_{@======================================@}"
              colOggetti.Add "1|UTILITA_Msys_Qry21_01_GRUPPO_OGGETTI_CORRENTI"
              colOggetti.Add "1|UTILITA_Msys_Qry21_02_Oggetti_Tutti"
              colOggetti.Add "1|UTILITA_Msys_Qry21_02_Oggetti_Tutti_RICERCA"
              colOggetti.Add "1|UTILITA_Msys_Qry21_03_Oggetti_Form"
              colOggetti.Add "1|UTILITA_Msys_Qry21_03_Oggetti_Form_RICERCA"
              colOggetti.Add "1|UTILITA_Msys_Qry21_04_Oggetti_Macro"
              colOggetti.Add "1|UTILITA_Msys_Qry21_04_Oggetti_Macro_RICERCA"
              colOggetti.Add "1|UTILITA_Msys_Qry21_05_Oggetti_Report"
              colOggetti.Add "1|UTILITA_Msys_Qry21_05_Oggetti_Report_RICERCA"
              colOggetti.Add "1|UTILITA_Msys_Qry21_06_Oggetti_Moduli"
              colOggetti.Add "1|UTILITA_Msys_Qry21_06_Oggetti_Moduli_RICERCA"
              colOggetti.Add "1|UTILITA_Msys_Qry21_07_Oggetti_TABELLE"
              colOggetti.Add "1|UTILITA_Msys_Qry21_07_Oggetti_TABELLE_Link"
              colOggetti.Add "1|UTILITA_Msys_Qry21_07_Oggetti_TABELLE_Link_RICERCA"
              colOggetti.Add "1|UTILITA_Msys_Qry21_07_Oggetti_TABELLE_RICERCA"
              colOggetti.Add "1|UTILITA_Msys_Qry21_08_Oggetti_Query"
              colOggetti.Add "1|UTILITA_Msys_Qry21_08_Oggetti_Query_RICERCA"
              
              
              'MODULI codice = 5
             colOggetti.Add "5|CLASSE_GESTIONE_OGGETTI_}--------------------------------------@"
             colOggetti.Add "5|CLASSE_GESTIONE_OGGETTI_DATABASE"
             colOggetti.Add "5|Msys_GEFILE_CARICA_Mdl00_}-------------------------------------@"
             colOggetti.Add "5|Msys_GEFILE_CARICA_Mdl01_LIBRERIE"
         


          Set CaricaOggettiDaEsportare = colOggetti
          Exit Function

          GestErr:
          MsgBox "Errore in CaricaOggettiDaEsportare: " & Err.Description, vbCritical
          Set CaricaOggettiDaEsportare = Nothing
      End Function
'==========================================================

'==========================================================
'05_CONTROLLO_ESISTENZA_OGGETTO_DB_ESTERNO
      Private Function OggettoEsisteNelDB(db As DAO.Database, tipoOggetto As Integer, _
                                          nomeOggetto As String) As Boolean
                 On Error GoTo GestErr
                 Dim rs As DAO.Recordset
                 Dim sql As String
                 Dim tipoFiltro As Long

                 '--- Mappatura tipo Access ? tipo MSysObjects
                 doEvents  ' per il controllo del ciclo'
                 
                 Select Case tipoOggetto
                     Case 1: tipoFiltro = 5          ' Query
                     Case 2: tipoFiltro = -32758     ' Form
                     Case 3: tipoFiltro = -32764     ' Report
                     Case 4: tipoFiltro = -32766     ' Macro
                     Case 5: tipoFiltro = -32761     ' Modulo
                     Case 6: tipoFiltro = -32768     ' Tabella
                     Case Else
                         OggettoEsisteNelDB = False
                         Exit Function
                 End Select

                 sql = "SELECT Name FROM MSysObjects WHERE Type=" & tipoFiltro & _
                       " AND Name='" & nomeOggetto & "'"
                 Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
                 OggettoEsisteNelDB = (Not rs.EOF)
                 rs.Close
                 Set rs = Nothing
                 Exit Function

                GestErr:
                 Debug.Print "Errore in OggettoEsisteNelDB: " & Err.Description
                 OggettoEsisteNelDB = False
      End Function
'==========================================================

'==========================================================
'06_CANCELLAZIONE_OGGETTO_DB_ESTERNO 
      Private Function CancellaOggetto(dbDest As DAO.Database, sxPath As String, _
                                  tipo As Integer, nome As String)
             On Error GoTo GestErr
             doEvents  ' per il controllo del ciclo'

             '=== INIZIO MIA MODIFICA: Assicuro intestazione CSV log se non esiste ===
             If Dir(logFilePath) = "" Then
                 Dim f As Integer
                 f = FreeFile
                 Open logFilePath For Output As #f
                 ' Scrivo intestazione coerente con il formato ScriviLog
                 Print #f, "Timestamp;Routine;Testo;Livello;Dettaglio"
                 Close #f
             End If
             '=== FINE MIA MODIFICA ===

             Select Case tipo
                 Case 1
                     dbDest.QueryDefs.Delete nome
                 Case 6
                     dbDest.TableDefs.Delete nome
                 Case Else
                     '--- Per Form, Report, Macro, Moduli serve Access.Application
                     Dim accApp As Object
                     Set accApp = CreateObject("Access.Application")
                     accApp.OpenCurrentDatabase sxPath
                     accApp.DoCmd.DeleteObject tipo, nome
                     accApp.CloseCurrentDatabase
                     accApp.Quit
                     Set accApp = Nothing
             End Select

             ScriviLog "GestioneOggetti", "CancellaOggetto", "INFO", _
                     "Cancellato oggetto esistente: " & nome & " nel DB: " & sxPath
             Exit Function

             GestErr:
             ScriviLog "GestioneOggetti", "CancellaOggetto", "ERRORE", _
                        "Errore in CancellaOggetto (" & nome & "): " & Err.Description
      End Function
'=========================================================

'==========================================================
'07_CONTROLLO_ESISTENZA_OGGETTO_DB_ORIGINE 
      Private Function OggettoEsiste(tipoOggetto As Integer, nomeOggetto As String) As Boolean
                On Error Resume Next
                doEvents  ' per il controllo del ciclo'
                Select Case tipoOggetto
                    Case 1: OggettoEsiste = Not IsNull(CurrentData.AllQueries(nomeOggetto).Name)
                    Case 2: OggettoEsiste = Not IsNull(CurrentProject.AllForms(nomeOggetto).Name)
                    Case 3: OggettoEsiste = Not IsNull(CurrentProject.AllReports(nomeOggetto).Name)
                    Case 4: OggettoEsiste = Not IsNull(CurrentProject.AllMacros(nomeOggetto).Name)
                    Case 5: OggettoEsiste = Not IsNull(CurrentProject.AllModules(nomeOggetto).Name)
                    Case 6: OggettoEsiste = Not IsNull(CurrentData.AllTables(nomeOggetto).Name)
                    Case Else: OggettoEsiste = False
                End Select
      End Function
'==========================================================

'==========================================================
'08_LOG_CSV_SEI_CAMPI

   Private Sub ScriviLog(ByVal testo As String, Optional ByVal NomeRoutine As String = "", _
                         Optional ByVal Livello As String = "", Optional ByVal Dettaglio As String = "")

       On Error GoTo ErrHandler
       
       Dim f As Integer
       Dim logsFolder As String
       Dim nuovoFile As Boolean
       
       ' Percorso CSV di default
       If Trim$(logFilePath) = "" Then
           logFilePath = CurrentProject.Path & "\Logs\ExportLog.csv"
       End If
       
       logsFolder = CurrentProject.Path & "\Logs\"
       
       ' Assicuro esistenza cartella Logs
       If Dir(logsFolder, vbDirectory) = "" Then
           MkDir logsFolder
       End If
       
       ' Verifico se il file NON esiste → servirà testata
       nuovoFile = (Dir(logFilePath) = "")
       
       f = FreeFile
       Open logFilePath For Append As #f
       
       ' ================================
       ' TESTATA CSV se nuovo file
       ' ================================
       If nuovoFile Then
           Print #f, "DATA_ORA;NOME_ROUTINE;TESTO;LIVELLO;DETTAGLIO;DATABASE"
       End If
       
       ' ================================
       ' Scrivo riga CSV (6 campi)
       ' ================================
       Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & ";" & _
                  IIf(NomeRoutine = "", "GENERICA", NomeRoutine) & ";" & _
                  testo & ";" & _
                  Livello & ";" & _
                  Replace(Dettaglio, vbCrLf, " ") & ";" & _
                  CurrentProject.FullName   ' <<< 6° campo
       
       Close #f
       Exit Sub
       
   ErrHandler:
       On Error Resume Next
       Debug.Print "Errore ScriviLog: " & Err.Description
       If f <> 0 Then Close #f
   End Sub
'==========================================================

'==========================================================
'09_CANCELLA_OGGETTI_DB_ESTERNO 
         'NOTE: per la cancellazione utilizzo DUE ROUTINE '
          'LA PRIMA CANCELLA GLI OGGETTI NEL DB ESTERNO riutilizzando le routine del modulo 
          'LA SECONDA IMPOSTO IL FILE LOG'
          'LE ATTIVITA SVOLTE SONO:'
                  'Caricare la collection di oggetti da procedura_01 (CaricaOggettiDaEsportare)
                  'Caricare la collection di percorsi DB esterni da procedura_02 (CaricaPercorsiEsportazione)
                  'Iterare per ogni DB esterno e per ogni oggetto da cancellare
                  'Controllare se esiste nel DB usando OggettoEsisteNelDB
                  'Se esiste ? cancellarlo con CancellaOggetto
                  'Loggare ogni azione in CancellazioneOggettiLog.txt
                  'Gestire errori senza bloccare l’intero processo
          '==========================================================

   Public Function CANCELLA_OGGETTI_DB_ESTERNO_Funct ()
              
            '=== 0.1 On error + Dim variabiliYYY
                  On Error GoTo GestErr

                  Dim colOggetti As Collection
                  Dim colPercorsi As Collection
                  Dim dbDest As DAO.Database
                  Dim sxPath As String
                  Dim itm As Variant
                  Dim itmPerc As Variant
                  Dim tipo As Integer
                  Dim nome As String
                  Dim parts() As String
                  Dim logPath As String
               '----------------------------------------------------------------------------------'

            '=== 0.2 Imposta percorso file di log
                '----------------------------------------------------------------------------------'
                  msgbox "ATTIVA CANCELLA OGGETTI",vbExclamation, "INIZIO CANCELLA OGGETTI"
                  logPath = CurrentProject.Path & "\Logs\CancellazioneOggettiLog.txt"

                  'Crea cartella Logs se non esiste
                  If Dir(CurrentProject.Path & "\Logs\", vbDirectory) = "" Then MkDir CurrentProject.Path & "\Logs\"

                  'Reindirizza ScriviLog al nuovo file
                   ImpostaLogFile(logPath)
                    
                    ScriviLog "=== AVVIO CANCELLAZIONE OGGETTI DB ESTERNI ==="
                '----------------------------------------------------------------------------------'

            '=== 1.0 Carica gli oggetti da cancellare
                  Set colOggetti = CaricaOggettiDaEsportare()
                  If colOggetti Is Nothing Then
                      ScriviLog "Errore: impossibile caricare collezione oggetti."
                      GoTo Uscita
                  End If
                  '----------------------------------------------------------------------------------'

            '=== 2.0 Carica i percorsi dei DB esterni
               '----------------------------------------------------------------------------------'
                      Set colPercorsi = CaricaPercorsiEsportazione()
                      If colPercorsi Is Nothing Then
                          ScriviLog "Errore: impossibile caricare percorsi DB esterni."
                          GoTo Uscita
                      End If

            '=== 3.0 Loop sui DB esterni
               '----------------------------------------------------------------------------------'
                    For Each itmPerc In colPercorsi
                          doEvents  ' per il controllo del ciclo'
                          sxPath = CStr(itmPerc)
                          ScriviLog "Connessione a DB esterno: " & sxPath

                          If Dir(sxPath) = "" Then
                              ScriviLog "ATTENZIONE: DB non trovato - " & sxPath
                              GoTo NextPercorso
                          End If

                          Set dbDest = DBEngine.OpenDatabase(sxPath)

                          '=== 4. Loop su oggetti da cancellare
                              For Each itm In colOggetti
                                  parts = Split(itm, "|")
                                  If UBound(parts) = 1 Then
                                      tipo = CInt(parts(0))
                                      nome = parts(1)

                                      'Controlla se oggetto esiste
                                      If OggettoEsisteNelDB(dbDest, tipo, nome) Then
                                          CancellaOggetto dbDest, sxPath, tipo, nome
                                      Else
                                          ScriviLog "Oggetto NON trovato: " & nome & " nel DB: " & sxPath
                                      End If
                                  End If
                              Next itm

                          '----------------------------------------------------------------------------------'
                      dbDest.Close
                      Set dbDest = Nothing

                NextPercorso:
                  Next itmPerc

                           msgbox "FINE CANCELLA OGGETTI",vbExclamation, "FINE CANCELLA OGGETTI"
                '----------------------------------------------------------------------------------'

            '=== 4.0 Uscita + gestine errori
                    Uscita:
                      ScriviLog "=== FINE CANCELLAZIONE OGGETTI DB ESTERNI ==="
                      Exit Function

                GestErr:
                      ScriviLog "ERRORE GENERALE in CANCELLA_OGGETTI_DB_ESTERNO_Funct(): " & Err.Description
                      Resume Uscita
   End Function

   Private Function ImpostaLogFile(ByVal nuovoPercorso As String) '// imposta file log'
    If LCase$(Right$(nuovoPercorso, 4)) = ".txt" Then
        nuovoPercorso = Left$(nuovoPercorso, Len(nuovoPercorso) - 4) & ".csv"
    End If
    logFilePath = nuovoPercorso
   End Function
'==========================================================




