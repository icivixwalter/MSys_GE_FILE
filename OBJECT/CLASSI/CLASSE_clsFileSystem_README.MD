# COME UTILIZZARE LA CLASSE SYSTEM
il file di gestione della classe si chiama : CLASSE_clsFileSystem_README.MD


## $CLASSE$SYSTEM
la classe che gestisce i file e le cartelle oltre i 260 caratteri.

codice classe ----> Attribute VB_Name = "CLASSE_clsFileSystem"_{codice per chiamare la classe}



## CODICE PER APPLICARE LA CLASSE



' ============================================================================
' §7.2  MODELLO DI CHIAMATA DAL FORM — apertura path con controllo > 260
' ============================================================================
' $MODELLO$DI$APERTURA$CARTELLA
' $MODELLO$ROUTINE$DOPPIO_CLICK
' $MODELLO$APRI$PATHDOPPIO_CLICK$APRI$PATH
' $FAQ = $COME$APRIRE$UNA$PATH
' $FAQ = $COME$APRIRE$UNA$CARTELLA

'
' Questo blocco è SOLO DOCUMENTAZIONE: è interamente commentato e non
' viene eseguito. Serve come modello da copiare nella routine del form
' (es. PathFileAtto_s_Txt_DblClick) per aprire la path usando la classe
' CLASSE_clsFileSystem e il metodo ApriPathConControllo (§4.18 / §7.1).
'
' PREREQUISITI:
'   - Il form contiene i controlli:
'       PathFileAtto_s_Txt   (TextBox con la path da aprire)
'       FileAtto_s_Txt       (TextBox opzionale, non usato qui)
'       IDGestione_TXT       (ID del record corrente)
'       RICERCA_FileAtto_s_Txt (TextBox dove scrivere l'esito)
'       DATAAGG_TXT          (TextBox data aggiornamento)
'       TIMEOPER_Txt         (TextBox ora operazione)
'   - La classe CLASSE_clsFileSystem è presente nel progetto VBA.
'
' FLUSSO:
'   1) Legge la path dal TextBox.
'   2) Istanzia la classe.
'   3) Chiama ApriPathConControllo(path, True, False):
'        - avvisaPercorsoLungo = True  -> se path > 260 mostra MsgBox Sì/No
'        - annullaSeLungo      = False -> se l'utente dice Sì, tenta comunque
'   4) Scrive l'esito nei campi del form tramite Select Case.
'   5) Riposiziona il record corrente con il bookmark.
'
' ----------------------------------------------------------------------------
' ESEMPIO DA COPIARE NEL FORM (routine DblClick del TextBox PathFileAtto_s_Txt)
' ----------------------------------------------------------------------------
'
 Private Sub PathFileAtto_s_Txt_DblClick(Cancel As Integer)
     On Error GoTo ErrHandler

     '// Dichiarazione variabili
     Dim MyPath_s As String
     Dim IDGestione_lng As Long
     Dim Messaggio_s As String
     Dim rst As DAO.Recordset
     Dim clsFileSystem As CLASSE_clsFileSystem

     '// Uscita immediata se il record è vuoto
     If IsNull(Me.IDGestione_TXT) Or Me.IDGestione_TXT = "" Then Exit Sub

     '// Lettura path dal TextBox
     If Me.PathFileAtto_s_Txt.Value > "" Then
         MyPath_s = Me.PathFileAtto_s_Txt.Value
         Debug.Print "Path da aprire: " & MyPath_s
     End If

     '// Lettura IDGestione (per il bookmark finale)
     If Me!IDGestione_TXT.Value > 0 Then
         IDGestione_lng = Me!IDGestione_TXT.Value
     Else
        MsgBox "ERRORE_CODICE ID NON TROVATO", vbCritical
        Exit Sub
    End If
    '// Istanza della classe
     Set clsFileSystem = New CLASSE_clsFileSystem

     '// Chiamata al modello di apertura con controllo > 260
     '//   Parametri: (path, avvisaPercorsoLungo, annullaSeLungo)
     '//   Avvisa se > 260 e, se l'utente conferma, tenta comunque l'apertura.
     Messaggio_s = clsFileSystem.ApriPathConControllo(MyPath_s, True, False)

     '// Rilascio dell'istanza
     Set clsFileSystem = Nothing

     '// Scrittura esito nei campi del form
     Select Case Messaggio_s
         Case "FILE NON TROVATO"
             Me.RICERCA_FileAtto_s_Txt = "FILE NON TROVATO"
             Me.DATAAGG_TXT = Date
             Me.TIMEOPER_Txt = Now

         Case "PATH TROVATA"
             Me.RICERCA_FileAtto_s_Txt = "PATH TROVATA"
             Me.DATAAGG_TXT = Date
             Me.TIMEOPER_Txt = Now

         Case "FILE TROVATO"
             Me.RICERCA_FileAtto_s_Txt = "FILE TROVATO"
             Me.DATAAGG_TXT = Date
            Me.TIMEOPER_Txt = Now
        Case "PATH > 260: APERTURA ANNULLATA"
            Me.RICERCA_FileAtto_s_Txt = "PATH > 260: APERTURA ANNULLATA"
            Me.DATAAGG_TXT = Date
            Me.TIMEOPER_Txt = Now
        Case "ERRORE: impossibile aprire la cartella (path > 260)"
            Me.RICERCA_FileAtto_s_Txt = "ERRORE APERTURA (> 260)"
            Me.DATAAGG_TXT = Date
            Me.TIMEOPER_Txt = Now
        Case "ERRORE: impossibile aprire il file (path > 260)"
            Me.RICERCA_FileAtto_s_Txt = "ERRORE APERTURA (> 260)"
            Me.DATAAGG_TXT = Date
            Me.TIMEOPER_Txt = Now
        Case "ERRORE: impossibile aprire la cartella"
            Me.RICERCA_FileAtto_s_Txt = "ERRORE APERTURA"
            Me.DATAAGG_TXT = Date
            Me.TIMEOPER_Txt = Now
        Case "ERRORE: impossibile aprire il file"
            Me.RICERCA_FileAtto_s_Txt = "ERRORE APERTURA"
            Me.DATAAGG_TXT = Date
             Me.TIMEOPER_Txt = Now

         Case Else
             MsgBox "Si è verificato un errore imprevisto: " & Messaggio_s, vbCritical
     End Select

     '// Refresh del campo esito
     Me.RICERCA_FileAtto_s_Txt.Requery

     '// Refresh form + riposizionamento sul record corrente
     Me.Refresh
     Set rst = Me.RecordsetClone
     rst.FindFirst "IDGestione = " & IDGestione_lng
     If Not rst.NoMatch Then
         Me.Bookmark = rst.Bookmark
     End If
     rst.Close
     Set rst = Nothing

     '// Pulizia variabili
     MyPath_s = vbNullString
     IDGestione_lng = 0
     Exit Sub

 ErrHandler:
     MsgBox "Errore durante l'apertura della cartella o file: " & Err.Description, vbCritical
     Set clsFileSystem = Nothing
 End Sub

' ----------------------------------------------------------------------------
' COMBINAZIONI DEI PARAMETRI DI ApriPathConControllo
' ----------------------------------------------------------------------------
'   ApriPathConControllo(path, True,  False)  ' avvisa se > 260, poi tenta
'   ApriPathConControllo(path, False, False)  ' tenta senza avvisare
'   ApriPathConControllo(path, False, True)   ' annulla sempre se > 260
'   ApriPathConControllo(path, True,  True)   ' avvisa; se l'utente conferma,
'                                             ' annulla comunque (sconsigliato)
'
' ----------------------------------------------------------------------------
' VALORI DI RITORNO GESTITI DAL Select Case
' ----------------------------------------------------------------------------
'   "PATH TROVATA"                              cartella esistente e aperta
'   "FILE TROVATO"                              file esistente e aperto
'   "FILE NON TROVATO"                          percorso inesistente o vuoto
'   "PATH > 260: APERTURA ANNULLATA"            path lungo, apertura annullata
'   "ERRORE: impossibile aprire la cartella"    apertura cartella fallita
'   "ERRORE: impossibile aprire il file"        apertura file fallita
'   "ERRORE: impossibile aprire la cartella (path > 260)"
'   "ERRORE: impossibile aprire il file (path > 260)"
'
' ============================================================================


'// MODELLI DI UTILIZZO DELLA CLASSE SYSTEM  *** fine ***
'//===============================================================================================================//





