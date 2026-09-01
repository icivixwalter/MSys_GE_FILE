'//- CLASSE_ATTI_TB01_GESTIONE_(classe atti @tb01@gestione che si trova in questo progetto) - NOTE+ATTIVITA+AGGIORNAMENTI _
    '//PATH_PROGETTO: _
            START "APRI PROGETTO CLASSE" "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_GE_FILE\MDB\MSys_GE_FILE.mdb"

    '//MODELLO ISTANZIAMENTO CLASSE GENERALE
        ''//.....................................................................................................//
            '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
                'Dim MyCls_CLASSE_ATTI_Tb01_Gestione As CLASSE_ATTI_Tb01_Gestione

            '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_FORM assegnando _
               l 'oggetto di classe alla variabile e creo creo l'oggetto di classe
                'Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione

            '//PATH O FILE = Inserire o una path oppure una path+file con percoso esteso.
                'Str1 = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_SERVIZI\MDB"
                'MyCls_CLASSE_ATTI_Tb01_Gestione.APRO_PATH_O_FILE_GENERICI_pSub (Str1)

            '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
                'Set MyCls_CLASSE_ATTI_Tb01_Gestione = Nothing
         ''//.....................................................................................................//
    
    '//LE @LIBRERIE NECESSARIE PER IL FUNZIONAMENTO DELLA CLASSE'

         '//'   |NAME_s                           |LIBRERIA_DLL_s  PATH_s
         '//'   |---------------------------------|-----------------------------------------------------------------------
         '//'   |VBA VBE6.DLL                     |C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6.DLL
         '//'   |Access  MSACC9.OLB               |C:\Program Files (x86)\Microsoft Office\Office\MSACC9.OLB
         '//'   |Office  MSO9.DLL                 |C:\Program Files (x86)\Microsoft Office\Office\MSO9.DLL
         '//'   |ADODB   msado21.tlb              |C:\Program Files (x86)\Common Files\System\ado\msado21.tlb
         '//'   |DAO dao360.dll                   |C:\Program Files (x86)\Common Files\Microsoft Shared\DAO\dao360.dll
         '//'   |---------------------------------|-----------------------------------------------------------------------
    
     ''//.....................................................................................................//

    '//GLI AGGIORNAMENTI EFFETTUATI'
      '//.....................................................................................................//
      '// 01_2025_10_29_AGGIORNAMENTO= RIALLINEATO TUTTE LE FUNZIONI E LE ROUTINE'
      '//.....................................................................................................//


'//LE OPTION DELLA CLASSE'
  Option Compare Database



    
    '//1° PARTE LE VARIABILI E LE DICHIARAZIONI API
    
    'DAO
    Dim DaoDB As DAO.Database           '//occorre il riferimento alla libreria  Microsoft DAO 3.6 Object Libray
    Dim DaoWks As DAO.Workspace
    Dim DaoRs As DAO.Recordset
    
    
    'ADO
    Dim ADODB_Conn As ADODB.Connection
    Dim AdodaoRs As ADODB.Recordset
    
    
    'record di lettura
    Dim DaoRs_Lett As DAO.Recordset
    Dim Lng_id_Lett As Long
    
    'record da cancellare
    Dim DaoRs_Canc As DAO.Recordset
    
    'record oggetti
    Dim DaoRs_Oggetti As DAO.Recordset
    Dim vV1_Oggetti As Variant
    
    Dim Ctr_Label As Label
    Dim Ctr_TextBox As TextBox
    Dim ctr As Control
    
    Dim objx_Forms As AllForms
    Dim objx_Form As Form
    Dim ix_Count_Control As Integer
    
    Dim obj As AccessObject
    Dim obj_Oggetto As AccessObject
    Dim dbs As Object
    
    
    
    
    
    'Contatori
    Dim icount As Integer
    Dim dbl_count As Double
    
    'Le variabili generiche
    Dim sSql As String                                          '//Stringa di estrazione SQL
    
    'Variabili generali
    Dim str1 As String
    Dim Int1 As Integer
    Dim Lng1 As Long
    Dim Dbl1 As Double
    Dim Bln1 As Boolean
    Dim Vv1 As Variant
    Dim obj1 As Object
    
    
    '//variabili del gestione del file
    '//...............................................//
        
        '//Path e file selezionato
        Dim result As Variant
        '//dim percorso per la casella di testo
        Dim percorso As Variant
        
         ' Variabile per il percorso del file selezionato
        Dim filePath As String
        
        ' Crea una nuova istanza di FileDialog (per aprire un file)
        Dim fileDialog As Object
        
        ' comando shell
        Dim officePath As String
        Dim command As String
        
        

        Dim Bool1 As Boolean
        Dim NameForm_s As String
        Dim Msg_s   As String
        
        
    '//...............................................//
    
            '//VARIABILI LINK
                Dim Disco_s As String
                Dim PATH_s As String
                Dim DATABASE_s As String
                Dim Table_s As String
                Dim Cancella_b As Boolean
                
    'Variabili dei parametri
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    




    '//************************************************************************************************************//
    '//
    '// DICHIARAZIONI API WINDOWS PER LA GESTIONE DEI FILE - VALIDE PER ACCESS A 32 E 64 BIT  *** INIZIO ****
    '//
    '//************************************************************************************************************//
    
    
    '//08_CLASSE_APRI_FILE_CON_SHELL_API.dim
    Private Const SW_SHOWNORMAL As Long = 1                             '//costante per la funzione ApriFileShellApiWindows sheell delle api windows
   
    '------------------------------------------
    ' API: GETOPENFILENAME - Mostra finestra selezione file
    '------------------------------------------
    ' ============================================================
    ' Classe: Apertura file con API di Windows (compatibile 32/64 bit)
    ' ============================================================


        '------------------------------------------
        ' API: SHELLEXECUTE - Apre file con app predefinita
        '------------------------------------------
        ' Usa l’associazione file di Windows per aprire un file o eseguire un'operazione (es. "open", "print")
        ' Parametri:
        ' - hwnd: handle finestra (normalmente 0)
        ' - lpOperation: "open", "edit", "print"
        ' - lpFile: percorso completo del file
        ' - lpParameters: parametri della riga di comando (normalmente vbNullString)
        ' - lpDirectory: directory di lavoro (normalmente vbNullString)
        ' - nShowCmd: modalità di visualizzazione finestra (es. vbNormalFocus = 1)

'//IF VB7 O OFFICE 32/64'
 #If VBA7 Then
    ' --- Office 2010 o successivo (VBA7 a 32/64 bit) ---

    ' API finestra di dialogo "Apri file"
    Private Declare PtrSafe Function GetOpenFileNameA Lib "comdlg32.dll" ( _
        ByRef pOpenfilename As OPENFILENAME) As Long

    ' API ShellExecute per aprire file/cartelle
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr

 #Else
    ' --- Office 2007 o precedente (VBA6, solo 32 bit) ---

    Private Declare Function GetOpenFileNameA Lib "comdlg32.dll" ( _
        ByRef pOpenfilename As OPENFILENAME) As Long

    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
 #End If

'//LE COSTANTI per la finestra di dialogo file
    '------------------------------------------
    '------------------------------------------
    Private Const OFN_FILEMUSTEXIST = &H1000       ' Il file selezionato deve esistere
    Private Const OFN_PATHMUSTEXIST = &H800        ' La cartella specificata deve esistere
    Private Const OFN_ALLOWMULTISELECT = &H200     ' Permette la selezione multipla
    Private Const OFN_EXPLORER = &H80000           ' Usa lo stile Explorer (moderno)

    '------------------------------------------
    ' STRUTTURA richiesta dall'API GetOpenFileNameA
    '------------------------------------------
  #If VBA7 Then
        Private Type OPENFILENAME
            lStructSize As Long                     ' Dimensione della struttura (in byte)
            hwndOwner As LongPtr                    ' Handle della finestra chiamante (0 = nessuna)
            hInstance As LongPtr                    ' Handle dell'applicazione (spesso 0)
            lpstrFilter As String                   ' Filtro dei file (es. "MDB Files" & Chr(0) & "*.mdb" & Chr(0) & Chr(0))
            lpstrCustomFilter As String             ' Filtro personalizzato (facoltativo)
            nMaxCustFilter As Long                  ' Dimensione massima filtro personalizzato
            nFilterIndex As Long                    ' Indice del filtro attivo (1 = primo)
            lpstrFile As String                     ' Buffer contenente file selezionato (output)
            nMaxFile As Long                        ' Lunghezza massima del buffer file
            lpstrFileTitle As String                ' Nome del file (senza path)
            nMaxFileTitle As Long                   ' Lunghezza massima del file title
            lpstrInitialDir As String               ' Cartella iniziale mostrata
            lpstrTitle As String                    ' Titolo finestra di dialogo
            flags As Long                           ' Combinazione delle costanti OFN_*
            nFileOffset As Integer                  ' Offset del file selezionato nel buffer
            nFileExtension As Integer               ' Offset dell'estensione nel file selezionato
            lpstrDefExt As String                   ' Estensione predefinita se utente non la specifica
            lCustData As LongPtr                    ' Dato utente personalizzato (non usato)
            lpfnHook As LongPtr                     ' Puntatore a funzione hook (non usato)
            lpTemplateName As String                ' Nome template finestra (non usato)
        End Type
        #Else
            Private Type OPENFILENAME
                lStructSize As Long
                hwndOwner As Long
                hInstance As Long
                lpstrFilter As String
                lpstrCustomFilter As String
                nMaxCustFilter As Long
                nFilterIndex As Long
                lpstrFile As String
                nMaxFile As Long
                lpstrFileTitle As String
                nMaxFileTitle As Long
                lpstrInitialDir As String
                lpstrTitle As String
                flags As Long
                nFileOffset As Integer
                nFileExtension As Integer
                lpstrDefExt As String
                lCustData As Long
                lpfnHook As Long
                lpTemplateName As String
            End Type
        #End If


    '//Note importanti: _
    Il parametro pOpenfilename in GetOpenFileNameA deve essere passato ByRef per far funzionare correttamente la chiamata. _
    La struttura OPENFILENAME è stata dichiarata con LongPtr per i puntatori/handle in 64 bit, e con Long per valori interi. _
    La stringa filtro (lpstrFilter) deve usare Chr(0) come separatore (ad es. "Testo filtro" & Chr(0) & "*.ext" & Chr(0) & Chr(0)). _
    La funzione GetOpenFileNameA restituisce un valore Long anche in 64 bit (non LongPtr). _
    ShellExecute ritorna un handle, quindi in 64 bit è LongPtr.

    '//*************************************************************************************************//



    '//************************************************************************************************************//
    '//
    '// DICHIARAZIONI API WINDOWS PER LA GESTIONE DEI FILE - VALIDE PER ACCESS A 32 E 64 BIT  *** FINE ****
    '//
    '//************************************************************************************************************//

'//EVENTI CLASSE - INITIALIZE + TERMINATE
   '______________________________________________________________________________


    'INIZIALIZZO LA CLASSE

    Private Sub Class_Initialize()

        On Error GoTo Err_Class_Initialize

                '//SE MANCA LA LIBRERIA CARICARE c:\CASA\LINGUAGGI\ACCESS\DLL_LIBRERIE\MSCOMCTL.OCX
                'Reset Variabili Oggetti Form
                m_sxTIPOGGETTO = ""
                m_sxPROPRIETA = ""
                m_sxMETODO = ""
                m_sxEVENTO = ""

    'USCITA ED ERRORI
    '..............................................................
    Exit_Class_Initialize:
        Exit Sub

    Err_Class_Initialize:
        MsgBox Err.Description
        Resume Exit_Class_Initialize


    End Sub



    'TERMINO LA CLASSE
    Private Sub Class_Terminate()

        On Error GoTo Err_Class_Terminate

        Dim TITOLO_Get As String


    'USCITA ED ERRORI
    '..............................................................
    Exit_Class_Terminate:
        Exit Sub

    Err_Class_Terminate:
        MsgBox Err.Description
        Resume Exit_Class_Terminate


    End Sub

'//08_CLASSE_APRI_FILE_CON_SHELL_API = LA FUNZIONE APRI I FILE E LA PATH CON LA SHELL DELLE API WINDOWS *** INIZIO ***
    '//==================================================================================================================//
    '//Note: funzione della @classe gestione file che apre i file di qualsiasi tipo e le cartelle _
        Gestire correttamente percorsi completi con file (es. C:\tmp\file.txt) _
        @APRI@FILE@APRI@PATH_(Funzione delle classe gestione file che consente di apire la PAHT oppure il FILE SE ESISTONO)
    ' Parametri:
    ' - MyPath_s       = STRINGA DELLA PATH
    ' - MyFile_s       = STRINGA NOME DEL FILE SENZA ESTENSIONE OPPURE CON ESTENSIONE
    ' - IDGestione_lng = Id del record tabella per ricordarsi in quale record è avvenuto l'evento
    '  -MyPath_b       = VARIABILE BOOLEAN con True = PATH + FILE senza estensione, FALSE = SOLO PATH

    '------------------------------------------------------------------------------
    ' ?? Ritorni possibili:
    ' - "FILE TROVATO"
    ' - "PATH TROVATA MA FILE NON SPECIFICATO"
    ' - "FILE NON TROVATO"
    ' - "ERRORE"
    '------------------------------------------------------------------------------
    '==================================================================================
    ' Funzione: ApriFileShellApiWindows
    ' Scopo: cerca e apre un file o cartella con ShellExecute API di Windows
    ' Supporta path > 255 caratteri tramite prefisso \\?\
    ' Parametri:
    '   parPath       -> percorso cartella o file completo se parFileName = ""
    '   parFileName   -> nome file senza estensione (può essere vuoto)
    '   parIDGestione -> opzionale, per logging/gestione (non usato qui)
    '   soloPath     -> se True apre solo la cartella anche se file non trovato
    ' Ritorna:
    '   Stringa di stato ("FILE TROVATO", "PATH TROVATA MA FILE NON SPECIFICATO", ecc.)
    '==================================================================================

  Public Function ApriFileShellApiWindows(parPath As String, parFileName As String, _
                                        parIDGestione As Long, _
                                        soloPath As Boolean) As String
    On Error GoTo ErrHandler

    Dim fileCompleto As String
    Dim fileTrovato As String
    Dim fileEstensione As String
    Dim nomeFunzione As String: nomeFunzione = "ApriFileShellApiWindows"

    Debug.Print ">>> Avvio funzione ApriFileShellApiWindows"
    Debug.Print "Path ricevuto: " & parPath
    Debug.Print "File ricevuto (senza estensione): " & parFileName
    Debug.Print "Flag soloPath: " & soloPath

    ' Aggiunge "\" se mancante
    If Right(parPath, 1) <> "\" Then parPath = parPath & "\"

    ' I° CONTROLLO = Esistenza path
    If Dir(parPath, vbDirectory) = "" Then
        MsgBox "Percorso non trovato: " & vbCrLf & parPath, vbExclamation
        Debug.Print ">>> Percorso non trovato"

        Call ScriviLogErrore(nomeFunzione, 1001, "Percorso non trovato: " & parPath, Erl)
        ApriFileShellApiWindows = "PERCORSO NON TROVATO"
        Exit Function
    End If

    ' II° CONTROLLO = Solo path
    If soloPath Then
        Debug.Print ">>> Apertura cartella: " & parPath
        Call EseguiShell(CStr(parPath))
        ApriFileShellApiWindows = "PATH TROVATA"
        Exit Function
    End If

    ' III° CONTROLLO = FILE da aprire

    ' CASO A: parFileName contiene già l'estensione
    If InStrRev(parFileName, ".") > 0 Then
        fileCompleto = parPath & parFileName
        If Dir(fileCompleto, vbNormal) <> "" Then
            Debug.Print ">>> File con estensione fornita trovato: " & fileCompleto
            Call EseguiShell(CStr(fileCompleto))
            ApriFileShellApiWindows = "FILE TROVATO"
            Exit Function
        End If
    End If

    ' CASO B: Estensione da cercare
    fileTrovato = Dir(parPath & parFileName & ".*", vbNormal)

    If fileTrovato <> "" Then
        fileEstensione = Mid(fileTrovato, InStrRev(fileTrovato, "."))
        fileCompleto = parPath & parFileName & fileEstensione
        Debug.Print ">>> File trovato: " & fileTrovato
        Debug.Print ">>> Percorso completo: " & fileCompleto

        If Len(Trim(fileCompleto)) > 0 And Dir(fileCompleto, vbNormal) <> "" Then
            Call EseguiShell(CStr(fileCompleto))
            ApriFileShellApiWindows = "FILE TROVATO"
        Else
            MsgBox "File non trovato nel percorso: " & vbCrLf & fileCompleto, vbExclamation
            Debug.Print ">>> File non trovato nonostante Dir iniziale: " & fileCompleto
            Call ScriviLogErrore(nomeFunzione, 1003, "File non trovato al controllo finale: " & fileCompleto, Erl)
            ApriFileShellApiWindows = "FILE NON TROVATO"
        End If
    Else
        MsgBox "File non trovato con nome: " & vbCrLf & parFileName, vbExclamation
        Debug.Print ">>> Nessun file con nome base trovato: " & parFileName
        Call ScriviLogErrore(nomeFunzione, 1002, "File non trovato: " & parPath & parFileName, Erl)
        ApriFileShellApiWindows = "FILE NON TROVATO"
    End If

    Exit Function

    ' Gestione errori generici
    ErrHandler:
        Debug.Print ">>> ERRORE: " & Err.Number & " - " & Err.Description
        Call ScriviLogErrore(nomeFunzione, Err.Number, Err.Description, Erl)
        MsgBox "Errore durante l'apertura file o cartella: " & Err.Description, vbCritical
        ApriFileShellApiWindows = "ERRORE"
    End Function


    '==================================================================================
    ' Sub: EseguiShell
    ' Scopo: apre file o cartella tramite API ShellExecute, gestendo path > 255 caratteri
    ' Parametri:
    '   percorsoCompleto (String) -> percorso file o cartella da aprire _
        Fa logging coerente con la tua struttura _
        Compatibile con file > 255 caratteri _
        Fa distinzione tra file e cartella
    '==================================================================================

    '//08_CLASSE_APRI_FILE_CON_SHELL_API.esegui.shell
    ' --- DICHIARAZIONE API COMPATIBILE 32/64 bit ---


    ' EseguiShell --- PROCEDURA PRINCIPALE ---
        Private Sub EseguiShell(ByVal percorsoCompleto As String)
        On Error GoTo ErrHandler

        ' --- Variabili ---
        #If VBA7 Then
            Dim risultato As LongPtr
        #Else
            Dim risultato As Long
        #End If
            Dim pathSolo As String
            Dim esisteFile As Boolean
            Dim esistePath As Boolean
            Dim percorsoShell As String
            Dim nomeFunzione As String: nomeFunzione = "EseguiShell"

            ' --- Pulizia input ---
            percorsoCompleto = Trim(percorsoCompleto)
            If Len(percorsoCompleto) = 0 Then
                MsgBox "Percorso o file non specificato.", vbCritical
                Exit Sub
            End If

            ' --- Aggiunge prefisso \\?\ se percorso > 255 caratteri e non lo ha già ---
            If Len(percorsoCompleto) > 255 Then
                If Left(percorsoCompleto, 4) <> "\\?\" Then
                    percorsoShell = "\\?\" & percorsoCompleto
                Else
                    percorsoShell = percorsoCompleto
                End If
            Else
                percorsoShell = percorsoCompleto
            End If

            ' --- Controllo se il file esiste ---
            esisteFile = (Dir(percorsoCompleto, vbNormal) <> "")

            If Not esisteFile Then
                ' Se non esiste file, controllo la cartella (path padre)
                pathSolo = percorsoCompleto
                If Right(pathSolo, 1) = "\" Then pathSolo = Left(pathSolo, Len(pathSolo) - 1)
                If InStrRev(pathSolo, "\") > 0 Then
                    pathSolo = Left(pathSolo, InStrRev(pathSolo, "\") - 1)
                End If
                esistePath = (Dir(pathSolo, vbDirectory) <> "")

                If esistePath Then
                    risultato = ShellExecute(0, "open", "\\?\" & pathSolo, vbNullString, vbNullString, 1)
                Else
                    MsgBox "Né file né cartella esistono nel percorso specificato:" & vbCrLf & percorsoCompleto, vbCritical
                    Exit Sub
                End If
            Else
                ' --- Apro il file ---
                risultato = ShellExecute(0, "open", percorsoShell, vbNullString, vbNullString, 1)
            End If

            ' --- Controllo codice di errore restituito da ShellExecute ---
            If risultato <= 32 Then
                MsgBox "Errore nell'apertura con ShellExecute (codice " & risultato & ").", vbCritical
                Call ScriviLogErrore(nomeFunzione, 3001, "Errore ShellExecute codice: " & risultato & " su: " & percorsoShell, Erl)
            Else
                ' (facoltativo) Log di successo
                ' Call ScriviLogErrore(nomeFunzione, 0, "Apertura completata correttamente: " & percorsoShell, 0)
            End If

            Exit Sub

            ErrHandler:
                MsgBox "Errore durante l'esecuzione di ShellExecute: " & Err.Description, vbCritical
                Call ScriviLogErrore(nomeFunzione, Err.Number, Err.Description, Erl)
            End Sub



            '//08_CLASSE_APRI_FILE_CON_SHELL_API = LA FUNZIONE APRI I FILE E LA PATH CON LA SHELL DELLE API WINDOWS *** FINE ***
            '//==================================================================================================================//




    '//APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub - APRO FILE ATTI DI GESTIONE -
    '//###########################################################################################################################//
    '//NOTE: apro il file .mdb _
        non ci sono parametri _
        AGGIORNAMENTO DEL = 13.07.2020
    '//CODICE----------->APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub

    '//AGGIORNAMENTO_DEL_13_07_2020
    '//APRI_ATTI_DI_GESTIONE
    Public Sub APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub()

    '//RECUPERO PARAMETRO DA TABELLA OGGETTI
    '//File da ricercare e la path
    Dim MyFile_s As String
    Dim MyPath_s As String


    On Error GoTo Err_APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub





        '//APRO SORGENTE FORM FILTRO 66 CAMPI
        '//.....................................................................................................
        '//Note           : Tramite una Select vengono individuati i valori da restiuire.

            MyFile_s = "LLPP_ATTI_GESTIONE_Analisi_provisoria.txt"
            '//ATTENZIONE occorre inserire uno spazio finale altrimenti non riconosce il percorso. Se invece è _
               un programma oppure un eseguibile non occorre lo spazio.
            MyPath_s = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\ATTI_TB01_GESTIONE\"


            '//Imposto i parametri
            ProceduraAttivaEseguita_s = "APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub"
            ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s




            '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
            '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
            '// prima la sessione dos e poi il pdf..
            '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
            Set WshPDF = CreateObject("wscript.shell")

            '// Unisco il comando pdf start + path definitiva
            Stringa1 = "%comspec% /c start " & MyPath_s & MyFile_s                               '//Comando di apertura
            '//Attivo il comando
            WshPDF.Run Stringa1


        '//*** fine ***
        '//APRO SORGENTE FORM FILTRO 66 CAMPI
        '//.....................................................................................................










    'USCITA  E GESTIONE ERRORI
    '..............................................................................................................

    Exit_APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub:
        Exit Sub

    Err_APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub:
        MsgBox Err.Description
        Resume Exit_APRI_MDB_Tb1_ATTI_DI_GESTIONE_AnalisiProvvisoria_pSub

    End Sub

    '//*** fine ***
    '//APRO FILE ATTI DI GESTIONE -
    '//###########################################################################################################################//

'//APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub = APRO LA PATH DEGLI ATTI DI GESTIONE - AGGIORNAMENTO_DEL_15_07_2020
 '//###########################################################################################################################//
    '//NOTE: apro il file .mdb _
        non ci sono parametri _
        AGGIORNAMENTO DEL = 13.07.2020
    '//CODICE----------->APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub

    '//AGGIORNAMENTO_DEL_13_07_2020
    '//APRI_ATTI_DI_GESTIONE
    Public Sub APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub()

        '//RECUPERO PARAMETRO DA TABELLA OGGETTI
        '//File da ricercare e la path
        Dim MyFile_s As String
        Dim MyPath_s As String


        On Error GoTo Err_APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub





            '//APRO SORGENTE FORM FILTRO 66 CAMPI
            '//.....................................................................................................
            '//Note           : Tramite una Select vengono individuati i valori da restiuire.

                MyFile_s = ""
                '//ATTENZIONE occorre inserire uno spazio finale altrimenti non riconosce il percorso. Se invece è _
                   un programma oppure un eseguibile non occorre lo spazio.
                MyPath_s = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_ELENCO\"


                '//Imposto i parametri
                ProceduraAttivaEseguita_s = ""
                ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s




                '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
                '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
                '// prima la sessione dos e poi il pdf..
                '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
                Set WshPDF = CreateObject("wscript.shell")

                '// Unisco il comando pdf start + path definitiva
                Stringa1 = "%comspec% /c start " & MyPath_s & MyFile_s                               '//Comando di apertura
                '//Attivo il comando
                WshPDF.Run Stringa1


            '//*** fine ***
            '//APRO SORGENTE FORM FILTRO 66 CAMPI
            '//.....................................................................................................




        'USCITA  E GESTIONE ERRORI
        '..............................................................................................................

        Exit_APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub:
            Exit Sub

        Err_APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub:
            MsgBox Err.Description
            Resume Exit_APRO_PATH_Tb01_ATTI_DI_GESTIONE_pSub

        End Sub

        '//*** fine ***
        '//APRO LA PATH DEGLI ATTI DI GESTIONE -
        '//###########################################################################################################################//


        '//AGGIORNAMENTO_DEL_15_08_2020
        '//APRO PATH GENERICA = Apre la cartella generica passata come parametro stringa. _
                    Routine utilzizabile su tutte le cartelle del progetto, occorre sempre dare la path esplicita _
                    e completa di percorso per la sua apertura. Se viene passata la path+file  allora puo essere _
                    utilizzata anche per aprime un file generico.
        Public Sub APRO_PATH_O_FILE_GENERICI_pSub(par_MyPath_s As String)

        '//RECUPERO PARAMETRO DA TABELLA OGGETTI
        '//File da ricercare e la path
        Dim MyFile_s As String
        Dim MyPath_s As String


        On Error GoTo Err_APRO_PATH_O_FILE_GENERICI_pSub


            '//APRO PATH O FILE GENERICI
            '//.....................................................................................................
            '//Passo un parametro stringa che può rappresentare una path da aprire esempio : _
               "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_ELENCO". _
               Se voglio aprire un file  occore passare la PATH+FILE con estensione es. _
               "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_ELENCO\file.xls"

                '//File da ricercare e la path
                'Dim MyFile_s As String
                'Dim MyPath_s As String

                'MyFile_s = "LLPP_ATTI_ALLEGATI_ARCHIVIO.RAR"
                'MyPath_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\LPP_ALLEGATI\"


                 '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
                    'Dim MyCls_CLASSE_ATTI_Tb01_Gestione As CLASSE_ATTI_Tb01_Gestione

                '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_FORM assegnando _
                   l 'oggetto di classe alla variabile e creo creo l'oggetto di classe
                    'Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione

                '//PATH O FILE = Inserire o una path oppure una path+file con percoso esteso.
                    'Str1 = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_SERVIZI\MDB"
                    'MyCls_CLASSE_ATTI_Tb01_Gestione.APRO_PATH_O_FILE_GENERICI_pSub (MyPath_s & MyFile_s)


                '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
                    'Set MyCls_CLASSE_ATTI_Tb01_Gestione = Nothing

            '//.....................................................................................................




            '//APRO SORGENTE FORM FILTRO 66 CAMPI
            '//.....................................................................................................
            '//Note           : Tramite una Select vengono individuati i valori da restiuire.

                MyFile_s = ""
                '//ATTENZIONE occorre inserire uno spazio finale altrimenti non riconosce il percorso. Se invece è _
                   un programma oppure un eseguibile non occorre lo spazio.
                    'ESEMPIO DI APERTURA DI UNA PATH GENERICA.
                    'modello path MyPath_s = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_ELENCO"

                MyFile_s = par_MyPath_s                     'recupero la path passata con il parametro


                '//Imposto i parametri
                ProceduraAttivaEseguita_s = ""
                ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s




                '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
                '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
                '// prima la sessione dos e poi il pdf..
                '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
                Set WshPDF = CreateObject("wscript.shell")

                '// Unisco il comando pdf start + path definitiva
                Stringa1 = "%comspec% /c start " & MyPath_s & MyFile_s                               '//Comando di apertura
                '//Attivo il comando
                WshPDF.Run Stringa1


            '//*** fine ***
            '//APRO SORGENTE FORM FILTRO 66 CAMPI
            '//.....................................................................................................

        'USCITA  E GESTIONE ERRORI
        '..............................................................................................................

        Exit_APRO_PATH_O_FILE_GENERICI_pSub:
            Exit Sub

        Err_APRO_PATH_O_FILE_GENERICI_pSub:
            MsgBox Err.Description
            Resume Exit_APRO_PATH_O_FILE_GENERICI_pSub

        End Sub

        '//*** fine ***
        '//APRO LA PATH DEGLI ATTI DI GESTIONE -
        '//###########################################################################################################################//

'//ITERO ALL.FORMS - OGGETTI FORM DEL DB CORRENTE PER LA RICERCA DELL'OGGETTO FORM PASSATO COME PARAMETRO - _
  EVENTUALMENTE DA UTILIZZARE PER PULIRE LA FORM
    '//###########################################################################################################################//
    '//NOTE: eseguo il ciclo nell'insieme forms e nell'insieme dei controlli della form filtrata con il nome e condizionata _
        dalla if di filtro. Ciclo all'interno delle proprieta delle form e resetto LE CASELLE DI TESTO E LE LABEL.
    '//CODICE----------->LoopFormProp_ITERA_OGGETTI_FORM_ATTI_Tb01_Gestione_PSub


     Public Sub LoopFormProp_ITERA_OGGETTI_FORM_ATTI_Tb01_Gestione_PSub(par_NameForm_s As String)
        On Error GoTo Err_Handler
        Dim frmLoop As Object
        Dim ctrl As Control
        Dim propLoop As Property
        Dim strFormName As String

        '//La variabile BOOLEAN per indicare se esiste
        Dim Bool1 As Boolean
        Dim NameForm_s As String
        Dim Msg_s  As String


        '//CHIAMO LA ROUTINE DI ITERAZIONE DEI CAMPI TEXT E LABEL DELLA FORM
        '//--------------------------------------------------------------------------------//-//
        '//CODICE----------->LoopFormProp_ITERA_OGGETTI_FORM_ATTI_Tb01_Gestione_PSub.puliscoTutto

            '//ATTIVO LA ROUTINE Call LoopFormProp_ITERA_OGGETTI_FORM_ATTI_Tb01_Gestione_PSub

        '//--------------------------------------------------------------------------------//-//

        For Each frmLoop In CurrentProject.AllForms

            '//imposto la variabile bool
            Bool1 = False
            '//la form da ricercare nel db
            'NameForm_s = "MOD30_Frm01_S01_ELENCO_SCH01_65txt_campi_1key"

            NameForm_s = par_NameForm_s

            '//TROVA QUELLA CORRENTE
            'If frmLoop.NAME = "MOD30_Frm01_S01_ELENCO_SCH01_65txt_campi_1key" Then
            If frmLoop.Name = par_NameForm_s Then

            '//LA FORM CORRENTE
            '//=====================================================================================================//

                    '//form trovata imposto il nome e  la variabile
                    strFormName = frmLoop.Name
                    '//imposto la variabile bool a true per individuare che è stata trovata
                    Bool1 = True


                    '//APRO LA FORM IN MODALITA DESIGN
                    DoCmd.OpenForm strFormName, acDesign, , , , acHidden

                    '//Qualifico l'oggetto form pe le proprieta
                    With Forms(strFormName)

                        For Each ctrl In .Controls

                            DoEvents

                            For Each propLoop In ctrl.Properties

                                 If ctrl.Name = "TXT_01" Then
                                    'Stop
                                 End If


                            '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                            '//==================================================================================================//
                            '//NOTE : pulisco le caselle di testo impostando la proprieta recordsource a null mentre le label _
                                    corrispondenti vengono impostate al nome oggetto label corrente.
                            '//CODICE----------->LoopFormProp_ITERA_OGGETTI_FORM_ATTI_Tb01_Gestione_PSub.puliscoTutto



                                '//PULISCO LA PROPRIETA RECORD SOURCE DELLE CASELLE DI TESTO
                                '//--------------------------------------------------------------------------------//------------//
                                '//NELLA CASELLE TEXT riconosciute con il valore ControlSource _
                                viene azzerata la proprieta control source con valore ""
                                    If propLoop.Name = "ControlSource" Then
                                        Debug.Print propLoop.Value
                                        'propLoop.value = ""
                                    End If

                                '//--------------------------------------------------------------------------------//------------//



                                '//PULISCO LA PROPRIETA CAPTION DELLE LABEL
                                '//--------------------------------------------------------------------------------//------------//
                                '//se è una label riconosciuta con il nome della proprieta _
                                Caption, viene sostituita il valore del contenuto della _
                                etichetta con il nome della label
                                     If propLoop.Name = "Caption" Then
                                                Debug.Print propLoop.Value
                                                '//inserisco il nome label?
                                                'propLoop.value = ctrl.NAME


                                    End If

                                '//--------------------------------------------------------------------------------//------------//

                            '//*** FINE ***
                            '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                            '//==================================================================================================//

                                '//CONTROLLO oggetto
                                Debug.Print "Form [" & .Name & "] - Control[" & ctrl.Name & "](" & propLoop.Name & "):" & propLoop.Value

                            Next '//For Each propLoop In ctrl.Properties


                        Next    '//For Each ctrl In .Controls

                                '//SALVATAGGIO FORM APERTA IN DESIGN
                                'Richiama la funzione che chiude la maschera aperta
                                DoCmd.SetWarnings False
                                DoCmd.Close acForm, strFormName, acSaveYes
                                DoCmd.SetWarnings True
                                DoCmd.Close acForm, strFormName

                    End With '//With Forms(strFormName)


            '//=====================================================================================================//



            End If
        Next

                    '//messaggio di avviso se Bool1=false e cioè non esiste la form

                    If Bool1 = False Then
                        '//imposto il testo del messaggio
                        Msg_s = "ATTENZIONE LA FORM " & NameForm_s & " NON E' STATA TROVATA IN QUESTO DATABASE!"
                        MsgBox Msg_s, vbExclamation
                    End If




        Exit Sub
            Err_Handler:
            Select Case Err.Number
                Case 2186
                    '//SE LA PROPRIETA NON ESISTE ATTIVO L'ERRORE E RIENTRO NEL CICLO FOR EACH
                    'Property not available in design view  SE NON CI SONO LE  PROPRIETA RIPROVA.
                    Resume Next
                Case Else
                    MsgBox "Err: " & Err.Number & " - " & Err.Description
            End Select
        End Sub

        '//*** fine ***
        '//ITERO ALL.FORMS - OGGETTI FORM DEL DB CORRENTE PER LA RICERCA DELL'OGGETTO FORM PASSATO COME PARAMETRO - _
          EVENTUALMENTE DA UTILIZZARE PER PULIRE LA FORM
        '//###########################################################################################################################//

'//ITERO ALL.MODULES - OGGETTI MODULES DEL DB CORRENTE PER LA RICERCA DELL'OGGETTO PASSATO COME PARAMETRO
 '//###########################################################################################################################//
    '//NOTE: eseguo il ciclo nell'insieme MODULES e nell'insieme DELLE CLASSI ALLA RICERCA DI QUELLA PASSATA COME PARAMETRO _
       mediante la if di filtro. Ciclo all'interno delle proprieta delle CLASSI E MODULI. - CONTROLLO ESISTENZA MODULI
    '//CODICE----------->LoopFormProp_ITERA_OGGETTI_MODULI_E_CLASSI_PSub

    '//TODO: DA FINIRE!!!
    Public Function LoopModuli_ITERA_OGGETTI_MODULI_E_CLASSI_PFunction(par_NameModulo_s As String) As String
        On Error GoTo Err_Handler
        Dim ModuliClassiLoop As Object
        Dim ObjModule As Modules
        Dim propLoop As Property
        Dim strModuloName As String

        '//La variabile BOOLEAN per indicare se esiste
        Dim Bool1 As Boolean
        Dim NameForm_s As String
        Dim Msg_s  As String




         '//CHIAMO LA ROUTINE DI ITERAZIONE DEI MODULI E DELLE CLASSI NEL DB CORRENTE _
           per controlloare se esiste il modulo di classe da attivare
        '//--------------------------------------------------------------------------------//-//
        '//CODICE----------->LoopFormProp_ITERA_OGGETTI_MODULI_E_CLASSI_PSub.puliscoTutto.IteraTutto

           '
           'Dim NWClasse_MODULI_E_CLASSI As CLASSE_ATTI_Tb01_Gestione


           '//ATTIVO LA ROUTINE
            '
            'Set NWClasse_MODULI_E_CLASSI = New CLASSE_ATTI_Tb01_Gestione

                '//imposto il nome della forma da ricercare
                '
             '   Str1 = "MyCls_CLASSE_ATTI_Tb01_Gestione"
                '//ITERO NEGLI OGGETTI DEL DB ALLA RICERCA DEL MODULO O DELLA CLASSE
                '
              '  vV1 = NWClasse_MODULI_E_CLASSI.LoopModuli_ITERA_OGGETTI_MODULI_E_CLASSI_PFunction(Str1)

        '//--------------------------------------------------------------------------------//-//



        For Each ModuliClassiLoop In CurrentProject.AllModules

            '//imposto la variabile bool
            Bool1 = False

            '//IL MODULO O CLASSE DA ricercare nel db
            strModuloName = par_NameModulo_s



            With ModuliClassiLoop
            '//CONTROLLO oggetto CORRENTE
            Debug.Print "MODULO O CLASSE ITERATO:  [" & ModuliClassiLoop.Name & "]"

            End With

            '//TROVA E CONTROLLA CON QUELLO CORRENTE
            If ModuliClassiLoop.Name = par_NameModulo_s Then

            '//IL MODULO O CLASSE CORRENTE IN QUESTO DB
            '//=====================================================================================================//

                    '//form trovata imposto il nome e  la variabile
                    strModuloName = ModuliClassiLoop.Name
                    '//imposto la variabile bool a true per individuare che è stata trovata
                    Bool1 = True


                    '//APRO LA FORM IN MODALITA DESIGN
                    'DoCmd.OpenForm strModuloName, acDesign, , , , acHidden

                    '//Qualifico l'oggetto form pe le proprieta
                    With Modules(strModuloName)

                        For Each ObjModule In .Parent


                            DoEvents

                            For Each propLoop In ObjModule.Properties

                                 If ObjModule.Name = "TXT_01" Then
                                    'Stop  TROVATO DA SVILUPPARE!!!!
                                 End If


                            '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                            '//==================================================================================================//
                            '//NOTE : pulisco le caselle di testo impostando la proprieta recordsource a null mentre le label _
                                    corrispondenti vengono impostate al nome oggetto label corrente.
                            '//CODICE----------->LoopModuli_ITERA_OGGETTI_MODULI_E_CLASSI_PFunction.puliscoTutto



                                '//PULISCO LA PROPRIETA RECORD SOURCE DELLE CASELLE DI TESTO
                                '//--------------------------------------------------------------------------------//------------//
                                '//NELLA CASELLE TEXT riconosciute con il valore ControlSource _
                                viene azzerata la proprieta control source con valore ""
                                    If propLoop.Name = "ControlSource" Then
                                        Debug.Print propLoop.Value
                                        'propLoop.value = ""
                                    End If

                                '//--------------------------------------------------------------------------------//------------//



                                '//PULISCO LA PROPRIETA CAPTION DELLE LABEL - MODIFICARE PER I MODULI E CLASSI
                                '//--------------------------------------------------------------------------------//------------//
                                '//se è una label riconosciuta con il nome della proprieta _
                                Caption, viene sostituita il valore del contenuto della _
                                etichetta con il nome della label
                                     If propLoop.Name = "Caption" Then
                                                Debug.Print propLoop.Value
                                                '//inserisco il nome label?
                                                'propLoop.value = ObjModule.NAME


                                    End If

                                '//--------------------------------------------------------------------------------//------------//

                            '//*** FINE ***
                            '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                            '//==================================================================================================//

                                '//CONTROLLO oggetto
                                Debug.Print "MODULO O CLASSE [" & .Name & "] - Control[" & ObjModule.Name & "](" & propLoop.Name & "):" & propLoop.Value

                            Next '//For Each propLoop In ObjModule.Properties


                        Next    '//For Each ObjModule In .Controls

                                '//SALVATAGGIO FORM APERTA IN DESIGN ?? PUO ESSERE UTILIZZATO PER IL MODULO O _
                                PER LA CLASSE !! CONTROLLARE???
                                'Richiama la funzione che chiude la maschera aperta
                                'DoCmd.SetWarnings False
                                'DoCmd.Close acForm, strModuloName, acSaveYes
                                'DoCmd.SetWarnings True
                                'DoCmd.Close acForm, strModuloName

                    End With '//With MODULES(strModuloName)

            '*** FINE ***
            '//IL MODULO O CLASSE CORRENTE IN QUESTO DB
            '//=====================================================================================================//



            End If
        Next

                    '//messaggio di avviso se Bool1=false e cioè non esiste la form

                    If Bool1 = False Then
                        '//imposto il testo del messaggio
                        Msg_s = "ATTENZIONE IL MODULO O LA CLASSE RICERCATA  " & strModuloName & " NON E' STATA TROVATA IN QUESTO DATABASE!"
                        LoopModuli_ITERA_OGGETTI_MODULI_E_CLASSI_PFunction = "OGGETTO NON TROVATO"
                        MsgBox Msg_s, vbExclamation
                    End If




        Exit Function
    Err_Handler:
        Select Case Err.Number
            Case 2186
                '//SE LA PROPRIETA NON ESISTE ATTIVO L'ERRORE E RIENTRO NEL CICLO FOR EACH
                'Property not available in design view  SE NON CI SONO LE  PROPRIETA RIPROVA.
                Resume Next
            Case Else
                MsgBox "Err: " & Err.Number & " - " & Err.Description
        End Select
    End Function

    '//*** fine ***
    '//ITERO ALL.MODULES - OGGETTI MODULES DEL DB CORRENTE PER LA RICERCA DELL'OGGETTO PASSATO COME PARAMETRO - _
      EVENTUALMENTE DA UTILIZZARE PER PULIRE I MODULI
    '//###########################################################################################################################//

'//ITERO IN TUTTI GLI OGGETTI - DEL DB CORRENTE PER LA RICERCA DELL'OGGETTO PASSATO COME PARAMETRO
    '//###########################################################################################################################//
    '//NOTE: eseguo il ciclo nell'insieme individuato dal paramtro oggetto ALLA RICERCA dell'OGGETTO _
       mediante la if di filtro.  - CONTROLLO ESISTENZA OGGETTO
    '//CODICE----------->LoopFormProp_ITERA_OGGETTI_DB_CORRENTE

    '****************************************************************************************************************************
    '           PROCEDURA DI ITERAZIONE DEGLI OGGETTI DI DATABASE
    ' i parametri ricevuti sono : par_Scelta db = path db
    '
    '
    '****************************************************************************************************************************

    Private Sub ITERA_OGGETTO_DB()




        '// FUNZIONE ITERAZIONE OGGETTI, Tabelle, Query, Form, Report, Macro _
            e Moduli
        '//====================================================================//
        '//MODELLO : di ricerca oggetti, prova la funzione si devono comunque _
                    impostare i parametri.

                    '//Imposto le variabilie e le resetto
                    Dim sx_Scelta_db As String
                    Dim sx_Name_Tab As String
                    Dim sx_Name_Query As String
                    Dim sx_Name_FORM As String
                    Dim sx_Name_Report As String
                    Dim sx_Name_Macro As String
                    Dim sx_Name_Modulo As String
                    Dim sx_Tipo_Oggetto As String
                    Dim ix_Operazione As Integer
                    Dim ix_Indice1 As Integer
                    Dim bx_Flag_b As Boolean


                    sx_Scelta_db = ""
                    sx_Name_Tab = ""
                    sx_Name_Query = ""
                    sx_Name_FORM = ""
                    sx_Name_Report = ""
                    sx_Name_Macro = ""
                    sx_Name_Modulo = ""
                    sx_Tipo_Oggetto = ""
                    ix_Operazione = 0
                    ix_Indice1 = 0
                    bx_Flag_b = False


                    '//chiamo la procedura di iterazione oggetti
                    str1 = SALVATAGGIO_OGGETTI_DB_PFunct(sx_Scelta_db, _
                                              sx_Name_Tab, _
                                              sx_Name_Query, _
                                              sx_Name_FORM, _
                                              sx_Name_Report, _
                                              sx_Name_Macro, _
                                              sx_Name_Modulo, _
                                              sx_Tipo_Oggetto, _
                                              ix_Operazione, _
                                              ix_Indice1, _
                                              bx_Flag_b)

        '//====================================================================//




        End Sub

'//SALVATAGGIO_OGGETTI_DB_PFunct  = PROCEDURE DI SALVATAGGIO '
    '//###########################################################################################################//
    '//####
    '//####                             Funzione SALVATAGGIO OGGETTI DATABASE MDB
    '//####
    '//###########################################################################################################//

        '//CODICE ---> 2020.02.11_SALVATAGGIO_OGGETTI_DB_PFunct_DB
        'PARAMETRI IN IMPUT:
        'par_sx_Scelta_db   =database interno od esterno (se la variabile path è valorizzata es. = c:\Gestioni\FILE.mdb _
                             allora viene attivata la II opzione del database esterno);
        'par_sx_Name_Tab    =nome tabella da ricercare;
        'par_sx_Name_Query  =nome query da ricercare
        'par_sTipo_Oggetto  =il tipo oggetto del db sono identificati con la seguente stringa: _
                        Cmd_Elenco_Query    = Tutte le query _
                        Cmd_Elenco_Tabelle  = Tutte le tabelle _
                        Cmd_Elenco_Forms    = Tutte le Form _
                        ATTENZIONE mancano ancora i parametri Select per LE FORM, I REPORT, LE MACRO _
                        LE PAGINE WEB che sono da integrare.
        'tipo di operazione =?  da definire ;
        'par_ix_Indice1 =?  da definire ;
        'par_bx_Flag_b = True = individuato cosa ??
        '
        '//OUTPUT : RESTITUISCE una stringa ANCORA DA DEFINRE E IMPLEMENTARE ???
        '//MODULO ---> UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB
        '//DB MDB ---> UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB

        '//MODELLO FUNZIONE ITER OGGETTI
        'Public Function SALVATAGGIO_OGGETTI_DB_PFunct(par_sx_Scelta_db As String, _
                                           par_sx_Name_Tab As String, _
                                           par_sx_Name_Query As String, _
                                           par_sx_Name_Form As String, _
                                           par_sTipo_Oggetto As String, _
                                           par_ix_Operazione As Integer, _
                                           par_ix_Indice1 As Integer, _
                                           par_bx_Flag_b As Boolean) As String
        '//***********************************************************************************************
        '//ATTIVITA ESEGUITE: _
        '02.ITERAZIONE NEL GRUPPO OGGETTI TABLEDEFS ..LE TABELLE DEL DB _
        '02.01 - OGGETTI QUERY = CONTROLLO OGGETTO NEL DB: Faccio il Controllo l'esistenza delle TABELLE
        'pulendo prima tutta la tabella  oggetti da quelli precedenti.
        '03.01 - APRO LA TABELLA DI DATABASE CHE CONTIENE GLI OGGETTI TABELLA
        '        Iterando in tutti gli oggetti dell'insieme Tables
        '03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB = Itero nel db tra gli oggetti
        '       table negli archivi del database e salvo il nome dell'oggetto tabella nel record corrente
        '       facendo un ciclo for nell'insieme TableDefs
        '03.03 .... definire
        '
        '//                                     SALVATAGGIO OGGETTI DATABASE
        '//{@=============================================================================================================@}
        '//CODICE --------------------> SALVATAGGIO_OGGETTI_DB_PFunct.Funzione
        '//
        Public Function SALVATAGGIO_OGGETTI_DB_PFunct( _
                                           par_sx_Scelta_db As String, _
                                           par_sx_Name_Tab As String, _
                                           par_sx_Name_Query As String, _
                                           par_sx_Name_FORM As String, _
                                           par_sx_Name_Report As String, _
                                           par_sx_Name_MACRO As String, _
                                           par_sx_Name_Modulo As String, _
                                           par_sTipo_Oggetto As String, _
                                           par_ix_Operazione As Integer, _
                                           par_ix_Indice1 As Integer, _
                                           par_bx_Flag_b As Boolean) As String


                On Error GoTo Err_SALVATAGGIO_OGGETTI_DB_PFunct

                '// FUNZIONE ITERAZIONE OGGETTI, Tabelle, Query, Form, Report, Macro _
                    e Moduli
                '//====================================================================//
                'Str1 = SALVATAGGIO_OGGETTI_DB_PFunct(sScelta_db, _
                                            sName_Tab, _
                                            sName_Query, _
                                            sName_Form, _
                                            sName_Report, _
                                            sName_Macro, _
                                            sName_Modulo, _
                                            sTipo_Oggetto, _
                                            iOperazione, _
                                            iIndice1, _
                                            Flag_b)
                '//====================================================================//


                '------------------------------------------------------------------------------------------
                    '1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB

                            'OPEN DB: Apre il database scelto per la creazione o cancellazione
                            '         con il parametro Option = false per indicare che l'apertura è condivisa /
                            '                          Option = True apertura esclusiva

                        If par_sx_Scelta_db = "" Then
                                'apro il db corrente se non viene indicato un database esterno
                                Set dbs = CurrentDb
                            Else
                                'apro un database esterno con la path >0 passata con i parametri
                                Set dbs = OpenDatabase(par_sx_Scelta_db, False)
                        End If

                    Select Case par_sTipo_Oggetto

                        '//- CONTROLLO TABELLE -
                        '//CHIAVE =====> modulo:controllo-tabelle
                        Case Is = "Cmd_Elenco_Tabelle"

                                '02.ITERAZIONE NEL GRUPPO OGGETTI TABLEDEFS ..LE TABELLE DEL DB
                                '______________________________________________________________________________
                                '02.01 - OGGETTI QUERY


                                    'pulisco la tabella degli OGGETTI TABELLA
                                    sSql = ""
                                    sSql = sSql & "DELETE MSsys_TABELLE_Qry01_01_TUTTE.*"
                                    sSql = sSql & "FROM MSsys_TABELLE_Qry01_01_TUTTE "
                                    sSql = sSql & "WITH OWNERACCESS OPTION;"

                                    CurrentDb.Execute (sSql)

                                    '03.01 - APRO LA TABELLA
                                    'apro la tabella per il salvataggio delle query + effettuo il _
                                    controllo sulla stringa sql
                                    Debug.Print "---------------------------------------------------------------"
                                    Debug.Print "CONTROLLO STRINGA SQL                                          "
                                    Debug.Print
                                    Set ApriRs = dbs.OpenRecordset("MSsys_TABELLE_Qry01_01_TUTTE")


                                    With dbs

                                        '//reset
                                        icount = 0

                                        '//stampa di controllo
                                        Debug.Print .TableDefs.Count & _
                                            " TableDefs in database  " & sx_Scelta_db


                                        For intCiclo = 0 To .TableDefs.Count - 1
                                                Debug.Print "  " & .TableDefs(intCiclo).Name

                                                '03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB
                                                '................................................................
                                                    ApriRs.AddNew
                                                    ApriRs.Fields("Name1_s") = .TableDefs(intCiclo).Name

                                                    '//NRO di oggetto
                                                    icount = icount + 1         '//aggiorno il contatore oggetti delle tabelle
                                                    ApriRs.Fields("NRO_OGGETTO_i") = icount
                                                    '//Tipo di oggetto caricato = FORM
                                                    ApriRs.Fields("TIPOGGETTO_s") = "TABLE"

                                                    '//SALVATAGGIO DEGLI ATTRIBUTI DELLA TABELLA - proprieta Attributes
                                                    ApriRs.Fields("Attribute_lng") = .TableDefs(intCiclo).Attributes

                                                        '//CONTROLLO OGGETTO TABELLA
                                                        '//-----------------------------------------------------------

                                                            '//controllo degli attributi
                                                            Debug.Print "LA TABELLA "
                                                            Debug.Print .TableDefs(intCiclo).Name

                                                            Debug.Print "controllo attributi della tabella"
                                                            Debug.Print .TableDefs(intCiclo).Attributes

                                                        '//-----------------------------------------------------------
                                                    ApriRs.Update
                                                '................................................................



                                            '03.03 .... definire
                                            If .TableDefs(intCiclo).Name = par_sx_Name_Tab Then
                                                Bln1 = True    ' bx_Flag_b True = trovato tabelle si può cancellare
                                                SALVATAGGIO_OGGETTI_DB_PFunct = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                                GoTo Exit_SALVATAGGIO_OGGETTI_DB_PFunct  'esci
                                            Else
                                                Bln1 = False   ' bx_Flag_b False = non trovato TABELLA
                                                SALVATAGGIO_OGGETTI_DB_PFunct = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                            End If

                                        Next intCiclo

                                    End With

                       ' - CONTROLLO QUERY -
                        Case Is = "Cmd_Elenco_Query"
                            '//CHIAVE =====> modulo:controllo-Query

                                '//ITERAZIONE QUERYDEFS <<LE QUERY>>
                                '//====================================================================================================//
                                'OGGETTI QUERY
                                'CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della QUERY


                                'pulisco la tabella delle query prima del salvataggio dei nuovi nomi
                                CurrentDb.Execute ("DELETE MSsys_QUERY_Qry01_01_TUTTE.* FROM MSsys_QUERY_Qry01_01_TUTTE;")

                                'apro la tabella per il salvataggio delle query
                                Set ApriRs = dbs.OpenRecordset("MSsys_QUERY_Qry01_01_TUTTE") '// perche non funziona??? da errore

                                'itero nel db
                                        With dbs

                                            '//RESET VARIABILI CICLO
                                            icount = 0      '//contatore numero oggetti query


                                            Debug.Print .QueryDefs.Count & _
                                                " QueryDefs in database  " & sx_Scelta_db


                                        '//CICLO OGGETTI QUERY
                                        '//------------------------------------------------------------------------------//
                                        For intCiclo = 0 To .QueryDefs.Count - 1

                                                'controlli
                                                '................................................................


                                                        Debug.Print "-----------------------------------------------------------"
                                                        Debug.Print
                                                                    '//stampo il nome della query + il contentuto sql
                                                                    Debug.Print " stampo il NOME DELLA QUERY"
                                                                    Debug.Print "  " & .QueryDefs(intCiclo).Name

                                                                    '//AGGIORNAMENTO_DEL_28_02_2020
                                                                    'Stampo il contenuto sql della query
                                                                    Debug.Print
                                                                    Debug.Print "  " & .QueryDefs(intCiclo).SQL

                                                                    '//AGGIORNAMENTO_DEL_28_02_2020
                                                                    '//stampo il tipo di query scelta direttamente
                                                                    Debug.Print "IL TIPO DELLA Query"
                                                                    Debug.Print "  " & .QueryDefs(intCiclo).Type

                                                        Debug.Print "-----------------------------------------------------------"
                                                        Debug.Print


                                                '................................................................

                                                    'SALVATAGGIO NOME QUERY
                                                    '................................................................
                                                        '//AGGIORNAMENTO_DEL_28_02_2020
                                                        '//Aggiorno il numero oggetto query salvati nel campo della tabella _
                                                        dei salvataggi del nro query.
                                                        icount = icount + 1         '//aggiorno il contatore oggetti delle query


                                                        ApriRs.AddNew
                                                        '//aggiugno alla tabella il nome della query + il contenuto sql
                                                        ApriRs.Fields("Name1_s") = .QueryDefs(intCiclo).Name
                                                        ApriRs.Fields("IdOrd_lng") = icount

                                                        '//Tipo di oggetto caricato = QUERY
                                                        ApriRs.Fields("TIPOGGETTO_s") = "QUERY"

                                                        '//AGGIORNAMENTO_DEL_28_02_2020
                                                        '//SQL + TIPO DI QUERY
                                                        ApriRs.Fields("Expression_m") = .QueryDefs(intCiclo).SQL


                                                        ApriRs.Update




                                                        '03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB
                                                '................................................................



                                                    '................................................................

                                                'controllo nome query passata con parametro
                                                If .QueryDefs(intCiclo).Name = par_sx_Name_Query Then

                                                    Bln1 = True    ' bx_Flag_b True = trovato QUERY

                                                    'stampo il contenuto della query scelta direttamente _
                                                    solo se esiste la query ricercata.
                                                    Debug.Print "  " & .QueryDefs(par_sx_Name_Query).SQL


                                                    'recupero il contenuto della query
                                                    SALVATAGGIO_OGGETTI_DB_PFunct = .QueryDefs(intCiclo).SQL


                                                    GoTo Exit_SALVATAGGIO_OGGETTI_DB_PFunct  'esci
                                                Else
                                                    Bln1 = False   ' bx_Flag_b False = non trovato QUERY
                                                    SALVATAGGIO_OGGETTI_DB_PFunct = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                                End If

                                            Next intCiclo

                                        '//CICLO OGGETTI QUERY  --> Next intCiclo *** fine ***
                                        '//------------------------------------------------------------------------------//

                                        End With
                                '//ITERAZIONE QUERYDEFS <<LE QUERY>>   *** FINE ***
                                '//====================================================================================================//



                       'CICLO NEGLI OGGETTI FORMS
                       Case Is = "Cmd_Elenco_Forms"
                       '//CHIAVE =====> modulo:controllo-forms




                                '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                                '//==============================================================================================================//
                                '//OGGETTI FORMS
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS


                                    'Dim obj As AccessObject, obj_Oggetto As Object
                                    Set obj_Oggetto = Application.CurrentProject

                                     '//RESET DELLE VARIABILI
                                    intCiclo = 0
                                    icount = 0

                                    'NOME form prorvvisoria ca controllare, poi cancellare
                                    par_sx_Name_FORM = "form ??? da definire"


                                    'apro la tabella per il salvataggio delle query
                                    Set ApriRs = dbs.OpenRecordset("MSsys_FORM_Qry01_01_TUTTE")


                                    '//RESET DELLE VARIABILI
                                    intCiclo = 0
                                    icount = 0

                                    'pulisco la tabella delle FORMS prima del salvataggio dei nuovi nomi
                                    CurrentDb.Execute ("DELETE MSsys_FORM_Qry01_01_TUTTE.* FROM MSsys_FORM_Qry01_01_TUTTE;")



                                '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                                '______________________________________________________________________________
                                '//OGGETTI FORMS
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS



                                'itero nel db
                                With dbs



                                '//ITERAZIONE Forms DEFS <<LE FORMS>> - CONTROLLO FORM -
                                '//______________________________________________________________________________
                                '//OGGETTI FORMS
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della Forms

                                        For Each obj In obj_Oggetto.AllForms


                                                'SALVATAGGIO NOME Forms
                                                '................................................................
                                                '//AGGIORNAMENTO_DEL_28_02_2020
                                                '//Aggiorno il numero oggetto Forms  salvati nel campo della tabella _
                                                dei salvataggi del nro Forms.

                                                  ApriRs.AddNew

                                                    'CONTROLLO DEI NOMI DELE FORMS
                                                    '................................................................

                                                            '//Search for open AccessObject objects in AllForms collection. _
                                                            con la proprieta IsLoad = esegui il controllo delle Forms Aperti _
                                                            mentre senza questa condizione, vengono iterati tutti gli oggetti _
                                                            Report del Db. Per le Forms APERTI O CARICATI  viene impostata la _
                                                            proprieta IsLoaded = True mentre per quelli chiusi è impostata a _
                                                            FALSE. _
                                                            È possibile utilizzare la proprietà IsLoaded per determinare se _
                                                            un oggetto AccessObject è attualmente caricato. Valore Boolean in sola lettura.


                                                                        Debug.Print "  " & obj.Name
                                                                        Debug.Print "  " & obj.Type

                                                                         '//Tipo di oggetto caricato = QUERY
                                                                         ApriRs.Fields("TIPOGGETTO_s") = "FORMS"



                                                            '//obj.Type = 2 LA FORMS
                                                            If obj.Type = 2 Then

                                                                    If obj.IsLoaded = True Then

                                                                           '//aggiugno alla tabella il nome della FORM solo se oggetto di tipo 2
                                                                            ApriRs.Fields("Name1_s") = obj.Name

                                                                            Debug.Print "FORM APERTA = " & obj.Name
                                                                            Debug.Print "  " & obj.Name
                                                                            Debug.Print "Nro : " & icount
                                                                            ApriRs.Fields("PROPRIETA_s") = "True"
                                                                            ApriRs.Fields("VALOREPROPRIETA_m") = "FORM APERTA"

                                                                             icount = icount + 1         '//aggiorno il contatore oggetti delle Forms
                                                                             Debug.Print "Nro : " & icount

                                                                             '//NRO di oggetto
                                                                             ApriRs.Fields("IdOrd_lng") = icount



                                                                    Else

                                                                            '//aggiugno alla tabella il nome della FORM solo se oggetto di tipo 2
                                                                            ApriRs.Fields("Name1_s") = obj.Name


                                                                            Debug.Print "FORM NON APERTO ---> " & obj.Name
                                                                            ApriRs.Fields("PROPRIETA_s") = "FALSE"
                                                                            ApriRs.Fields("VALOREPROPRIETA_m") = "FORM CHIUSA"

                                                                             icount = icount + 1         '//aggiorno il contatore oggetti delle Forms
                                                                             Debug.Print "Nro : " & icount

                                                                             '//NRO di oggetto
                                                                             ApriRs.Fields("IdOrd_lng") = icount

                                                                    End If
                                                            End If


                                                    '................................................................

                                                  ApriRs.Update
                                                '................................................................

                                        Next obj

                                    End With

                    '//ITERAZIONE FORMS DEFS <<LE FORMS>>    *** FINE ***
                    '//==============================================================================================================//

                                    '**************************************************************



                      'CICLO NEGLI OGGETTI REPORTS
                       Case Is = "Cmd_Elenco_Reports"
                        '//CHIAVE =====> modulo:controllo-Report
                        '//Aggiornamento 22 MARZO 2020

                                '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                                '//______________________________________________________________________________
                                '//OGGETTI FORMS
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS



                                    'Dim obj As AccessObject, obj_Oggetto As Object
                                    Set obj_Oggetto = Application.CurrentProject

                                    '//RESET DELLE VARIABILI
                                    intCiclo = 0
                                    icount = 0

                                    'pulisco la tabella delle FORMS prima del salvataggio dei nuovi nomi
                                    CurrentDb.Execute ("DELETE MSsys_REPORT_Qry01_01_TUTTI.* FROM MSsys_REPORT_Qry01_01_TUTTI;")

                                    'apro la tabella per il salvataggio dei REPORT
                                    Set ApriRs = dbs.OpenRecordset("MSsys_REPORT_Qry01_01_TUTTI")


                                '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                                '______________________________________________________________________________
                                '//OGGETTI FORMS
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS



                                'itero nel db
                                With dbs



                                '//ITERAZIONE REPORTS DEFS <<I REPORT>> - CONTROLLO REPORT -
                                '//______________________________________________________________________________
                                '//OGGETTI REPORTS
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della REPORTS

                                        For Each obj In obj_Oggetto.AllReports


                                                'SALVATAGGIO NOME REPORTS
                                                '................................................................
                                                '//AGGIORNAMENTO_DEL_28_02_2020
                                                '//Aggiorno il numero oggetto REPORTS  salvati nel campo della tabella _
                                                dei salvataggi del nro REPORTS.



                                                  ApriRs.AddNew
                                                    '//aggiugno alla tabella il nome della query + il contenuto sql
                                                    ApriRs.Fields("Name1_s") = obj.Name




                                                    'CONTROLLO DEI NOMI DEI REPORT
                                                    '................................................................

                                                            '//Search for open AccessObject objects in AllReports collection. _
                                                            con la proprieta IsLoad = esegui il controllo dei Reports Aperti _
                                                            mentre senza questa condizione, vengono iterati tutti gli oggetti _
                                                            Report del Db. Per i reports APERTI O CARICATI  viene impostata la _
                                                            proprieta IsLoaded = True mentre per quelli chiusi è impostata a _
                                                            FALSE. _
                                                            È possibile utilizzare la proprietà IsLoaded per determinare se _
                                                            un oggetto AccessObject è attualmente caricato. Valore Boolean in sola lettura.


                                                                        Debug.Print "  " & obj.Name
                                                                        Debug.Print "  " & obj.Type

                                                                        '//Tipo di oggetto caricato = REPORTS
                                                                        ApriRs.Fields("TIPOGGETTO_s") = "REPORTS"

                                                            '//obj.Type = 3 LA FORMS
                                                            If obj.Type = 3 Then

                                                                    If obj.IsLoaded = True Then


                                                                            Debug.Print "REPORT APERTO = " & obj.Name
                                                                            Debug.Print "  " & obj.Name
                                                                            Debug.Print "Nro : " & icount
                                                                            ApriRs.Fields("PROPRIETA_s") = "True"
                                                                            ApriRs.Fields("VALOREPROPRIETA_m") = "FORM APERTA"

                                                                             icount = icount + 1         '//aggiorno il contatore oggetti delle REPORTS
                                                                             Debug.Print "Nro : " & icount

                                                                             '//NRO di oggetto
                                                                             ApriRs.Fields("IdOrd_lng") = icount



                                                                    Else
                                                                            Debug.Print "REPORT NON APERTO ---> " & obj.Name
                                                                            ApriRs.Fields("PROPRIETA_s") = "FALSE"
                                                                            ApriRs.Fields("VALOREPROPRIETA_m") = "FORM CHIUSA"

                                                                             icount = icount + 1         '//aggiorno il contatore oggetti delle REPORTS
                                                                             Debug.Print "Nro : " & icount

                                                                             '//NRO di oggetto
                                                                             ApriRs.Fields("IdOrd_lng") = icount

                                                                    End If
                                                            End If

                                                    '................................................................





                                                  ApriRs.Update
                                                '................................................................




                                        Next obj

                                    End With



                                '//ITERAZIONE REPORTS DEFS <<I REPORT>>
                                '//______________________________________________________________________________


                       'CICLO NEGLI OGGETTI MACRO - CONTROLLO MACRO -
                       Case Is = "Cmd_Elenco_MACRO"
                        '//CHIAVE =====> modulo:controllo-Macro
                        '//Aggiornamento 22 MARZO 2020

                                '//ITERAZIONE MACRO DEFS <<LE MACRO>>
                                '//______________________________________________________________________________
                                '//OGGETTI MACRO
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO



                                    'Dim obj As AccessObject, obj_Oggetto As Object
                                    Set obj_Oggetto = Application.CurrentProject

                                    '//RESET DELLE VARIABILI
                                    intCiclo = 0
                                    icount = 0

                                    'pulisco la tabella delLE MACRO prima del salvataggio dei nuovi nomi
                                    CurrentDb.Execute ("DELETE MSsys_MACRO_Qry01_01_TUTTE.* FROM MSsys_MACRO_Qry01_01_TUTTE;")

                                    'apro la tabella per il salvataggio dei Macro
                                    Set ApriRs = dbs.OpenRecordset("MSsys_MACRO_Qry01_01_TUTTE")


                                '//ITERAZIONE MACRO DEFS <<LE MACRO>>
                                '______________________________________________________________________________
                                '//OGGETTI MACRO
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO



                                'itero nel db
                                With dbs



                                '//ITERAZIONE MACRO DEFS <<I Macro>>
                                '//______________________________________________________________________________
                                '//OGGETTI MACRO
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO

                                        For Each obj In obj_Oggetto.AllMacros


                                                'SALVATAGGIO NOME MACRO
                                                '................................................................
                                                '//AGGIORNAMENTO_DEL_28_02_2020
                                                '//Aggiorno il numero oggetto MACRO  salvati nel campo della tabella _
                                                dei salvataggi del nro MACRO.



                                                  ApriRs.AddNew
                                                    '//aggiugno alla tabella il nome della query + il contenuto sql
                                                    ApriRs.Fields("Name1_s") = obj.Name





                                                    'CONTROLLO DEI NOMI DEI Macro
                                                    '................................................................

                                                            '//Search for open AccessObject objects in AllMACRO collection. _
                                                            con la proprieta IsLoad = esegui il controllo dei MACRO Aperti _
                                                            mentre senza questa condizione, vengono iterati tutti gli oggetti _
                                                            Macro del Db. Per i MACRO APERTI O CARICATI  viene impostata la _
                                                            proprieta IsLoaded = True mentre per quelli chiusi è impostata a _
                                                            FALSE. _
                                                            È possibile utilizzare la proprietà IsLoaded per determinare se _
                                                            un oggetto AccessObject è attualmente caricato. Valore Boolean in sola lettura.


                                                                        Debug.Print "  " & obj.Name
                                                                        Debug.Print "  " & obj.Type

                                                                        '//Tipo di oggetto caricato = MACRO
                                                                        ApriRs.Fields("TIPOGGETTO_s") = "MACRO"


                                                            '//obj.Type = 4 LA MACRO
                                                            If obj.Type = 4 Then




                                                                            Debug.Print "Macro APERTO = " & obj.Name
                                                                            Debug.Print "  " & obj.Name
                                                                            Debug.Print "Nro : " & icount
                                                                            'ApriRs.Fields("PROPRIETA_s") = "True"
                                                                            'ApriRs.Fields("VALOREPROPRIETA_m") = "FORM APERTA"

                                                                             icount = icount + 1         '//aggiorno il contatore oggetti delle MACRO
                                                                             Debug.Print "Nro : " & icount

                                                                             '//NRO di oggetto
                                                                             ApriRs.Fields("IdOrd_lng") = icount




                                                            End If

                                                    '................................................................





                                                  ApriRs.Update
                                                '................................................................




                                        Next obj

                                    End With



                                '//ITERAZIONE MACRO DEFS <<I Macro>>
                                '//______________________________________________________________________________



                            Case Else
                                    MsgBox "Scelta comando non effettuata", vbExclamation
                                     GoTo Exit_SALVATAGGIO_OGGETTI_DB_PFunct
                        End Select
                                    'chiudo i rs aperti + Rilascio la memoria
                                    ApriRs.Close
                                    Set ApriRs = Nothing



            '------------------------------------------------------------------------------
            '                       FINE FUNCTIONE E GESTIONE ERRORI

                Exit_SALVATAGGIO_OGGETTI_DB_PFunct:
                    dbs.Close
                    Set dbs = Nothing
                    Exit Function


                Err_SALVATAGGIO_OGGETTI_DB_PFunct:
                    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "SALVATAGGIO_OGGETTI_DB_PFunct"
                    Resume Exit_SALVATAGGIO_OGGETTI_DB_PFunct

                End Function

            '//*** FINE ***
            '//                                     SALVATAGGIO OGGETTI DATABASE
            '//{@=============================================================================================================@}




            '//*** fine ***
            '//ITERO IN TUTTI GLI OGGETTI - DEL DB CORRENTE PER LA RICERCA DELL'OGGETTO PASSATO COME PARAMETRO
            '//###########################################################################################################################//




        '****************************************************************************************************************************
        '           PROCEDURA DI ITERAZIONE DEGLI OGGETTI DI DATABASE PER IL CONTROLLO ESISTENZA OGGETTI
        ' i parametri ricevuti sono : par_Scelta db = path db
        '
        '
        '****************************************************************************************************************************

'//SALVO_ATTIVITA_CMB_02_IN_TMP_Sub = SALVATAGGIO PARAMETRI NELLA TABELLA TMP
    '//###########################################################################################################################//
    '//NOTE: Nella tabella <<GEST_Tb01_N01_PARAM_OGGETTI_TMP>> _
    viene effettuato il salvataggio dei parametri. I parametri in entrata sono i seguenti :
    'par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
    'par_sxTIPOGGETTO       =
    'par_sxATTIVITAEVENTO   =
    'par_sxCODATTIVITA      = CODICE ATTIVITA OBBLIGATORIO
    Public Sub SALVO_ATTIVITA_CMB_02_IN_TMP_Sub(par_sxCODOGGETTO As String, _
                                             par_sxTIPOGGETTO As String, _
                                             par_sxATTIVITAEVENTO As String, _
                                             par_sxCODATTIVITA As String)





        '//CHIAMO LA ROUTINE SALVO_ATTIVITA_CMB_02_IN_TMP_Sub DELLA CLASSE OGGETTI FORM
        '//CODICE----------->SALVO_ATTIVITA_CMB_02_IN_TMP_Sub.Salvataggio_ATTIVITA.Attivazione
        '//--------------------------------------------------------------------------------//-//
        '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_FORM
        '//PARAMETRI            : sono 4 di seguito indicati
                'par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
                'par_sxTIPOGGETTO       =
                'par_sxATTIVITAEVENTO   =
                'par_sxCODATTIVITA      = CODICE ATTIVITA OBBLIGATORIO
        '//NOTA  :Salvo i parametri del codice oggetto e del codice attiva nella tabella tmp solo nel caso in cui _
                sia IL CODICE OGGETTO che il CODICE ATTIVITA sono VALORIZZATI altrimenti la classe ricevente emette _
                un msg di avviso.



                    'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_FORM
                    'Assegno e creo l'oggetto di classe
                    'Dim MyCls_CLASSE_GESTIONE_OGGETTI_FORM As Classe_GESTIONE_OGGETTI_FORM
                    'Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = New Classe_GESTIONE_OGGETTI_FORM


                    'PARAMETRI DI ESTRAZIONE
                    'Attivo il metodo della classe rappresentata dalla routine passando
                    'i 4 parametri necessari
                    'sxCODOGGETTO = s
                    'sxTIPOGGETTO = s
                    'sxATTIVITAEVENTO = s
                    'sxCODATTIVITA = s
                    'MyCls_CLASSE_GESTIONE_OGGETTI_FORM.SALVO_ATTIVITA_CMB_02_IN_TMP_Sub sxCODOGGETTO, sxTIPOGGETTO, sxATTIVITAEVENTO, sxCODATTIVITA

                    'Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = Nothing

            '//--------------------------------------------------------------------------------//-//





    On Error GoTo Err_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub



                    '//CONTROLLO PARAMETRI RICEVENTI
                    '//--------------------------------------------------------------------------------//-//
                    '//NOTA : se almeno i due parametri più importanti, ossia CODICE ATTIVITA E CODICE
                    '//     OGGETTO non sono valorizzati, viene emesso un messaggio di avviso e di uscita.
                        If sxCODOGGETTO > "" And sxCODATTIVITA > "" Then
                            MsgBox "ATTENZIONE I PARAMETRI CODICE OGGETTO E CODICE ATTIVITA NON SONO VALORIZZATI", vbExclamation, "CLASSE - CONTROLLO PARAMETRI"
                        End If
                    '//--------------------------------------------------------------------------------//-//




                        'SALVATAGGIO ATTIVITA NELLA TABELLA GEST_Tb01_N01_PARAM_OGGETTI_TMP
                        '___________________________________________________________________
                            'Stringa ssql
                            'di salvataggio nella tabella tmp della attivita da svolgere
                            sSql = ""
                            sSql = sSql & "UPDATE GEST_Tb01_N01_PARAM_OGGETTI_TMP "
                            sSql = sSql & "SET GEST_Tb01_N01_PARAM_OGGETTI_TMP.CODATTIVITA = '" & sxCODATTIVITA & "', "
                                                    'sSql = sSql & "SET GEST_Tb01_N01_PARAM_OGGETTI_TMP.CODATTIVITA = '" & sxCODATTIVITA & "'; "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI_TMP.CODOGGETTO = '" & par_sxCODOGGETTO & "'; "


                            'controlo
                            Debug.Print sSql

                            'comando eseguito
                            CurrentDb.Execute (sSql)
                        '___________________________________________________________________



    'USCITA  E GESTIONE ERRORI
    '..............................................................................................................

    Exit_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub:
        Exit Sub

    Err_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub:
        MsgBox Err.Description
        Resume Exit_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub

    End Sub

    '//*** FINE ***
    '//SALVATAGGIO PARAMETRI NELLA TABELLA TMP
    '//###########################################################################################################################//




        '//###########################################################################################################//
        '//####
        '//####                             PROCEDURE DI SALVATAGGIO *** FINE ***
        '//####
        '//###########################################################################################################//

'//PROCEDURE DI RICERCA
    '//###########################################################################################################//
    '//####   Funzione RICERCA OGGETTI DATABASE MDB
    '//####
    '//###########################################################################################################//

        '//                 Funzione RICERCA OGGETTI DATABASE MDB
        '//CODICE ---> 2020.02.11_RICERCA_OGGETTI_DB_PFunct_DB
        'PARAMETRI IN IMPUT:
        'par_sx_Scelta_db   =database interno od esterno (se la variabile path è valorizzata es. = c:\Gestioni\FILE.mdb _
                             allora viene attivata la II opzione del database esterno);
        'par_sx_Name_Tab    =nome tabella da ricercare;
        'par_sx_Name_Query  =nome query da ricercare
        'par_sTipo_Oggetto  =il tipo oggetto del db sono identificati con la seguente stringa: _
                        Cmd_Elenco_Query    = Tutte le query _
                        Cmd_Elenco_Tabelle  = Tutte le tabelle _
                        Cmd_Elenco_Forms    = Tutte le Form _
                        ATTENZIONE mancano ancora i parametri Select per LE FORM, I REPORT, LE MACRO _
                        LE PAGINE WEB che sono da integrare.
        'tipo di operazione =?  da definire ;
        'par_ix_Indice1 =?  da definire ;
        'par_bx_Flag_b = True = individuato cosa ??
        '
        '//OUTPUT : RESTITUISCE UN MESSAGGIO "OGGETTO TROVATO NEL DB" O "NON TROVATO NEL DB"
        '//MODULO ---> UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB
        '//DB MDB ---> UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB

        '//***********************************************************************************************
        '//ATTIVITA ESEGUITE: _
        '02.ITERAZIONE NEL GRUPPO OGGETTI TABLEDEFS ..LE TABELLE DEL DB _
        '02.01 - OGGETTI QUERY = CONTROLLO OGGETTO NEL DB: Faccio il Controllo l'esistenza delle TABELLE
        'pulendo prima tutta la tabella  oggetti da quelli precedenti.
        '03.01 - APRO LA TABELLA DI DATABASE CHE CONTIENE GLI OGGETTI TABELLA
        '        Iterando in tutti gli oggetti dell'insieme Tables
        '03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB = Itero nel db tra gli oggetti
        '       table negli archivi del database e salvo il nome dell'oggetto tabella nel record corrente
        '       facendo un ciclo for nell'insieme TableDefs
        '03.03 .... definire
        '
        '//                                     SALVATAGGIO OGGETTI DATABASE
        '//{@=============================================================================================================@}
        '//CODICE --------------------> RICERCA_OGGETTI_DB_PFunct.Funzione
        '//
        Public Function RICERCA_OGGETTI_DB_PFunct( _
                                           par_sx_Scelta_db As String, _
                                           par_sx_Name_Tab As String, _
                                           par_sx_Name_Query As String, _
                                           par_sx_Name_FORM As String, _
                                           par_sx_Name_Report As String, _
                                           par_sx_Name_MACRO As String, _
                                           par_sx_Name_Modulo As String, _
                                           par_sTipo_Oggetto As String, _
                                           par_ix_Operazione As Integer, _
                                           par_ix_Indice1 As Integer, _
                                           par_bx_Flag_b As Boolean) As String


                On Error GoTo Err_RICERCA_OGGETTI_DB_PFunct


                '------------------------------------------------------------------------------------------
                    '1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB

                            'OPEN DB: Apre il database scelto per la creazione o cancellazione
                            '         con il parametro Option = false per indicare che l'apertura è condivisa /
                            '                          Option = True apertura esclusiva

                        If par_sx_Scelta_db = "" Then
                                'apro il db corrente se non viene indicato un database esterno
                                Set dbs = CurrentDb
                            Else
                                'apro un database esterno con la path >0 passata con i parametri
                                Set dbs = OpenDatabase(par_sx_Scelta_db, False)
                        End If

                    Select Case par_sTipo_Oggetto

                        '//- CONTROLLO TABELLE -
                        '//CHIAVE =====> modulo:controllo-tabelle
                        Case Is = "Cmd_Elenco_Tabelle"

                                '02.ITERAZIONE NEL GRUPPO OGGETTI TABLEDEFS ..LE TABELLE DEL DB
                                '______________________________________________________________________________
                                '02.01 - OGGETTI QUERY


                                    'pulisco la tabella degli OGGETTI TABELLA
                                    sSql = ""
                                    sSql = sSql & "DELETE MSsys_TABELLE_Qry01_01_TUTTE.*"
                                    sSql = sSql & "FROM MSsys_TABELLE_Qry01_01_TUTTE "
                                    sSql = sSql & "WITH OWNERACCESS OPTION;"

                                    CurrentDb.Execute (sSql)

                                    '03.01 - APRO LA TABELLA
                                    'apro la tabella per il salvataggio delle query + effettuo il _
                                    controllo sulla stringa sql
                                    Debug.Print "---------------------------------------------------------------"
                                    Debug.Print "CONTROLLO STRINGA SQL                                          "
                                    Debug.Print
                                    Set ApriRs = dbs.OpenRecordset("MSsys_TABELLE_Qry01_01_TUTTE")


                                    With dbs

                                        '//reset
                                        icount = 0

                                        '//stampa di controllo
                                        Debug.Print .TableDefs.Count & _
                                            " TableDefs in database  " & sx_Scelta_db


                                        For intCiclo = 0 To .TableDefs.Count - 1
                                                Debug.Print "  " & .TableDefs(intCiclo).Name

                                                '03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB
                                                '................................................................
                                                    ApriRs.AddNew
                                                    ApriRs.Fields("Name1_s") = .TableDefs(intCiclo).Name

                                                    '//NRO di oggetto
                                                    icount = icount + 1         '//aggiorno il contatore oggetti delle tabelle
                                                    ApriRs.Fields("NRO_OGGETTO_i") = icount
                                                    '//Tipo di oggetto caricato = FORM
                                                    ApriRs.Fields("TIPOGGETTO_s") = "TABLE"

                                                    '//SALVATAGGIO DEGLI ATTRIBUTI DELLA TABELLA - proprieta Attributes
                                                    ApriRs.Fields("Attribute_lng") = .TableDefs(intCiclo).Attributes

                                                        '//CONTROLLO OGGETTO TABELLA
                                                        '//-----------------------------------------------------------

                                                            '//controllo degli attributi
                                                            Debug.Print "LA TABELLA "
                                                            Debug.Print .TableDefs(intCiclo).Name

                                                            Debug.Print "controllo attributi della tabella"
                                                            Debug.Print .TableDefs(intCiclo).Attributes

                                                        '//-----------------------------------------------------------
                                                    ApriRs.Update
                                                '................................................................



                                            '03.03 .... definire
                                            If .TableDefs(intCiclo).Name = par_sx_Name_Tab Then
                                                Bln1 = True    ' bx_Flag_b True = trovato tabelle si può cancellare
                                                RICERCA_OGGETTI_DB_PFunct = "QUERY_TROVATA_NEL_DB" 'riassegno alla funzione il valore del bx_Flag_b
                                                GoTo Exit_RICERCA_OGGETTI_DB_PFunct  'esci
                                            Else
                                                Bln1 = False   ' bx_Flag_b False = non trovato TABELLA
                                                RICERCA_OGGETTI_DB_PFunct = "QUERY_NON_TROVATA_NEL_DB" 'riassegno alla funzione il valore del bx_Flag_b
                                            End If

                                        Next intCiclo

                                    End With

                       ' - CONTROLLO QUERY -
                        Case Is = "Cmd_Elenco_Query"
                            '//CHIAVE =====> modulo:controllo-Query

                                '//ITERAZIONE QUERYDEFS <<LE QUERY>>
                                '//====================================================================================================//
                                'OGGETTI QUERY
                                'CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della QUERY


                                'pulisco la tabella delle query prima del salvataggio dei nuovi nomi
                                CurrentDb.Execute ("DELETE MSsys_QUERY_Qry01_01_TUTTE.* FROM MSsys_QUERY_Qry01_01_TUTTE;")

                                'apro la tabella per il salvataggio delle query
                                Set ApriRs = dbs.OpenRecordset("MSsys_QUERY_Qry01_01_TUTTE") '// perche non funziona??? da errore

                                'itero nel db
                                        With dbs

                                            '//RESET VARIABILI CICLO
                                            icount = 0      '//contatore numero oggetti query


                                            Debug.Print .QueryDefs.Count & _
                                                " QueryDefs in database  " & sx_Scelta_db


                                        '//CICLO OGGETTI QUERY
                                        '//------------------------------------------------------------------------------//
                                        For intCiclo = 0 To .QueryDefs.Count - 1

                                                'controlli
                                                '................................................................


                                                        Debug.Print "-----------------------------------------------------------"
                                                        Debug.Print
                                                                    '//stampo il nome della query + il contentuto sql
                                                                    Debug.Print " stampo il NOME DELLA QUERY"
                                                                    Debug.Print "  " & .QueryDefs(intCiclo).Name

                                                                    '//AGGIORNAMENTO_DEL_28_02_2020
                                                                    'Stampo il contenuto sql della query
                                                                    Debug.Print
                                                                    Debug.Print "  " & .QueryDefs(intCiclo).SQL

                                                                    '//AGGIORNAMENTO_DEL_28_02_2020
                                                                    '//stampo il tipo di query scelta direttamente
                                                                    Debug.Print "IL TIPO DELLA Query"
                                                                    Debug.Print "  " & .QueryDefs(intCiclo).Type

                                                        Debug.Print "-----------------------------------------------------------"
                                                        Debug.Print


                                                '................................................................

                                                    'SALVATAGGIO NOME QUERY
                                                    '................................................................
                                                        '//AGGIORNAMENTO_DEL_28_02_2020
                                                        '//Aggiorno il numero oggetto query salvati nel campo della tabella _
                                                        dei salvataggi del nro query.
                                                        icount = icount + 1         '//aggiorno il contatore oggetti delle query


                                                        ApriRs.AddNew
                                                        '//aggiugno alla tabella il nome della query + il contenuto sql
                                                        ApriRs.Fields("Name1_s") = .QueryDefs(intCiclo).Name
                                                        ApriRs.Fields("IdOrd_lng") = icount

                                                        '//Tipo di oggetto caricato = QUERY
                                                        ApriRs.Fields("TIPOGGETTO_s") = "QUERY"

                                                        '//AGGIORNAMENTO_DEL_28_02_2020
                                                        '//SQL + TIPO DI QUERY
                                                        ApriRs.Fields("Expression_m") = .QueryDefs(intCiclo).SQL


                                                        ApriRs.Update




                                                        '03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB
                                                '................................................................



                                                    '................................................................

                                                'controllo nome query passata con parametro
                                                If .QueryDefs(intCiclo).Name = par_sx_Name_Query Then

                                                    Bln1 = True    ' bx_Flag_b True = trovato QUERY

                                                    'stampo il contenuto della query scelta direttamente _
                                                    solo se esiste la query ricercata.
                                                    Debug.Print "  " & .QueryDefs(par_sx_Name_Query).SQL


                                                    'recupero il contenuto della query
                                                    RICERCA_OGGETTI_DB_PFunct = .QueryDefs(intCiclo).SQL


                                                    GoTo Exit_RICERCA_OGGETTI_DB_PFunct  'esci
                                                Else
                                                    Bln1 = False   ' bx_Flag_b False = non trovato QUERY
                                                    RICERCA_OGGETTI_DB_PFunct = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                                End If

                                            Next intCiclo

                                        '//CICLO OGGETTI QUERY  --> Next intCiclo *** fine ***
                                        '//------------------------------------------------------------------------------//

                                        End With
                                '//ITERAZIONE QUERYDEFS <<LE QUERY>>   *** FINE ***
                                '//====================================================================================================//



                       'CICLO NEGLI OGGETTI FORMS
                       Case Is = "Cmd_Elenco_Forms"
                       '//CHIAVE =====> modulo:controllo-forms




                                '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                                '//==============================================================================================================//
                                '//OGGETTI FORMS
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS


                                    'Dim obj As AccessObject, obj_Oggetto As Object
                                    Set obj_Oggetto = Application.CurrentProject

                                     '//RESET DELLE VARIABILI
                                    intCiclo = 0
                                    icount = 0

                                    'NOME form prorvvisoria ca controllare, poi cancellare
                                    par_sx_Name_FORM = "form ??? da definire"


                                    'apro la tabella per il salvataggio delle query
                                    Set ApriRs = dbs.OpenRecordset("MSsys_FORM_Qry01_01_TUTTE")


                                    '//RESET DELLE VARIABILI
                                    intCiclo = 0
                                    icount = 0

                                    'pulisco la tabella delle FORMS prima del salvataggio dei nuovi nomi
                                    CurrentDb.Execute ("DELETE MSsys_FORM_Qry01_01_TUTTE.* FROM MSsys_FORM_Qry01_01_TUTTE;")



                                '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                                '______________________________________________________________________________
                                '//OGGETTI FORMS
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS



                                'itero nel db
                                With dbs



                                '//ITERAZIONE Forms DEFS <<LE FORMS>> - CONTROLLO FORM -
                                '//______________________________________________________________________________
                                '//OGGETTI FORMS
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della Forms

                                        For Each obj In obj_Oggetto.AllForms


                                                'SALVATAGGIO NOME Forms
                                                '................................................................
                                                '//AGGIORNAMENTO_DEL_28_02_2020
                                                '//Aggiorno il numero oggetto Forms  salvati nel campo della tabella _
                                                dei salvataggi del nro Forms.

                                                  ApriRs.AddNew

                                                    'CONTROLLO DEI NOMI DELE FORMS
                                                    '................................................................

                                                            '//Search for open AccessObject objects in AllForms collection. _
                                                            con la proprieta IsLoad = esegui il controllo delle Forms Aperti _
                                                            mentre senza questa condizione, vengono iterati tutti gli oggetti _
                                                            Report del Db. Per le Forms APERTI O CARICATI  viene impostata la _
                                                            proprieta IsLoaded = True mentre per quelli chiusi è impostata a _
                                                            FALSE. _
                                                            È possibile utilizzare la proprietà IsLoaded per determinare se _
                                                            un oggetto AccessObject è attualmente caricato. Valore Boolean in sola lettura.


                                                                        Debug.Print "  " & obj.Name
                                                                        Debug.Print "  " & obj.Type

                                                                         '//Tipo di oggetto caricato = QUERY
                                                                         ApriRs.Fields("TIPOGGETTO_s") = "FORMS"



                                                            '//obj.Type = 2 LA FORMS
                                                            If obj.Type = 2 Then

                                                                    If obj.IsLoaded = True Then

                                                                           '//aggiugno alla tabella il nome della FORM solo se oggetto di tipo 2
                                                                            ApriRs.Fields("Name1_s") = obj.Name

                                                                            Debug.Print "FORM APERTA = " & obj.Name
                                                                            Debug.Print "  " & obj.Name
                                                                            Debug.Print "Nro : " & icount
                                                                            ApriRs.Fields("PROPRIETA_s") = "True"
                                                                            ApriRs.Fields("VALOREPROPRIETA_m") = "FORM APERTA"

                                                                             icount = icount + 1         '//aggiorno il contatore oggetti delle Forms
                                                                             Debug.Print "Nro : " & icount

                                                                             '//NRO di oggetto
                                                                             ApriRs.Fields("IdOrd_lng") = icount



                                                                    Else

                                                                            '//aggiugno alla tabella il nome della FORM solo se oggetto di tipo 2
                                                                            ApriRs.Fields("Name1_s") = obj.Name


                                                                            Debug.Print "FORM NON APERTO ---> " & obj.Name
                                                                            ApriRs.Fields("PROPRIETA_s") = "FALSE"
                                                                            ApriRs.Fields("VALOREPROPRIETA_m") = "FORM CHIUSA"

                                                                             icount = icount + 1         '//aggiorno il contatore oggetti delle Forms
                                                                             Debug.Print "Nro : " & icount

                                                                             '//NRO di oggetto
                                                                             ApriRs.Fields("IdOrd_lng") = icount

                                                                    End If
                                                            End If


                                                    '................................................................

                                                  ApriRs.Update
                                                '................................................................

                                        Next obj

                                    End With

                    '//ITERAZIONE FORMS DEFS <<LE FORMS>>    *** FINE ***
                    '//==============================================================================================================//

                                    '**************************************************************



                      'CICLO NEGLI OGGETTI REPORTS
                       Case Is = "Cmd_Elenco_Reports"
                        '//CHIAVE =====> modulo:controllo-Report
                        '//Aggiornamento 22 MARZO 2020

                                '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                                '//______________________________________________________________________________
                                '//OGGETTI FORMS
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS



                                    'Dim obj As AccessObject, obj_Oggetto As Object
                                    Set obj_Oggetto = Application.CurrentProject

                                    '//RESET DELLE VARIABILI
                                    intCiclo = 0
                                    icount = 0

                                    'pulisco la tabella delle FORMS prima del salvataggio dei nuovi nomi
                                    CurrentDb.Execute ("DELETE MSsys_REPORT_Qry01_01_TUTTI.* FROM MSsys_REPORT_Qry01_01_TUTTI;")

                                    'apro la tabella per il salvataggio dei REPORT
                                    Set ApriRs = dbs.OpenRecordset("MSsys_REPORT_Qry01_01_TUTTI")


                                '//ITERAZIONE FORMS DEFS <<LE FORMS>>
                                '______________________________________________________________________________
                                '//OGGETTI FORMS
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS



                                'itero nel db
                                With dbs



                                '//ITERAZIONE REPORTS DEFS <<I REPORT>> - CONTROLLO REPORT -
                                '//______________________________________________________________________________
                                '//OGGETTI REPORTS
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della REPORTS

                                        For Each obj In obj_Oggetto.AllReports


                                                'SALVATAGGIO NOME REPORTS
                                                '................................................................
                                                '//AGGIORNAMENTO_DEL_28_02_2020
                                                '//Aggiorno il numero oggetto REPORTS  salvati nel campo della tabella _
                                                dei salvataggi del nro REPORTS.



                                                  ApriRs.AddNew
                                                    '//aggiugno alla tabella il nome della query + il contenuto sql
                                                    ApriRs.Fields("Name1_s") = obj.Name




                                                    'CONTROLLO DEI NOMI DEI REPORT
                                                    '................................................................

                                                            '//Search for open AccessObject objects in AllReports collection. _
                                                            con la proprieta IsLoad = esegui il controllo dei Reports Aperti _
                                                            mentre senza questa condizione, vengono iterati tutti gli oggetti _
                                                            Report del Db. Per i reports APERTI O CARICATI  viene impostata la _
                                                            proprieta IsLoaded = True mentre per quelli chiusi è impostata a _
                                                            FALSE. _
                                                            È possibile utilizzare la proprietà IsLoaded per determinare se _
                                                            un oggetto AccessObject è attualmente caricato. Valore Boolean in sola lettura.


                                                                        Debug.Print "  " & obj.Name
                                                                        Debug.Print "  " & obj.Type

                                                                        '//Tipo di oggetto caricato = REPORTS
                                                                        ApriRs.Fields("TIPOGGETTO_s") = "REPORTS"

                                                            '//obj.Type = 3 LA FORMS
                                                            If obj.Type = 3 Then

                                                                    If obj.IsLoaded = True Then


                                                                            Debug.Print "REPORT APERTO = " & obj.Name
                                                                            Debug.Print "  " & obj.Name
                                                                            Debug.Print "Nro : " & icount
                                                                            ApriRs.Fields("PROPRIETA_s") = "True"
                                                                            ApriRs.Fields("VALOREPROPRIETA_m") = "FORM APERTA"

                                                                             icount = icount + 1         '//aggiorno il contatore oggetti delle REPORTS
                                                                             Debug.Print "Nro : " & icount

                                                                             '//NRO di oggetto
                                                                             ApriRs.Fields("IdOrd_lng") = icount



                                                                    Else
                                                                            Debug.Print "REPORT NON APERTO ---> " & obj.Name
                                                                            ApriRs.Fields("PROPRIETA_s") = "FALSE"
                                                                            ApriRs.Fields("VALOREPROPRIETA_m") = "FORM CHIUSA"

                                                                             icount = icount + 1         '//aggiorno il contatore oggetti delle REPORTS
                                                                             Debug.Print "Nro : " & icount

                                                                             '//NRO di oggetto
                                                                             ApriRs.Fields("IdOrd_lng") = icount

                                                                    End If
                                                            End If

                                                    '................................................................





                                                  ApriRs.Update
                                                '................................................................




                                        Next obj

                                    End With



                                '//ITERAZIONE REPORTS DEFS <<I REPORT>>
                                '//______________________________________________________________________________


                       'CICLO NEGLI OGGETTI MACRO - CONTROLLO MACRO -
                       Case Is = "Cmd_Elenco_MACRO"
                        '//CHIAVE =====> modulo:controllo-Macro
                        '//Aggiornamento 22 MARZO 2020

                                '//ITERAZIONE MACRO DEFS <<LE MACRO>>
                                '//______________________________________________________________________________
                                '//OGGETTI MACRO
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO



                                    'Dim obj As AccessObject, obj_Oggetto As Object
                                    Set obj_Oggetto = Application.CurrentProject

                                    '//RESET DELLE VARIABILI
                                    intCiclo = 0
                                    icount = 0

                                    'pulisco la tabella delLE MACRO prima del salvataggio dei nuovi nomi
                                    CurrentDb.Execute ("DELETE MSsys_MACRO_Qry01_01_TUTTE.* FROM MSsys_MACRO_Qry01_01_TUTTE;")

                                    'apro la tabella per il salvataggio dei Macro
                                    Set ApriRs = dbs.OpenRecordset("MSsys_MACRO_Qry01_01_TUTTE")


                                '//ITERAZIONE MACRO DEFS <<LE MACRO>>
                                '______________________________________________________________________________
                                '//OGGETTI MACRO
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO



                                'itero nel db
                                With dbs



                                '//ITERAZIONE MACRO DEFS <<I Macro>>
                                '//______________________________________________________________________________
                                '//OGGETTI MACRO
                                '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO

                                        For Each obj In obj_Oggetto.AllMacros


                                                'SALVATAGGIO NOME MACRO
                                                '................................................................
                                                '//AGGIORNAMENTO_DEL_28_02_2020
                                                '//Aggiorno il numero oggetto MACRO  salvati nel campo della tabella _
                                                dei salvataggi del nro MACRO.



                                                  ApriRs.AddNew
                                                    '//aggiugno alla tabella il nome della query + il contenuto sql
                                                    ApriRs.Fields("Name1_s") = obj.Name





                                                    'CONTROLLO DEI NOMI DEI Macro
                                                    '................................................................

                                                            '//Search for open AccessObject objects in AllMACRO collection. _
                                                            con la proprieta IsLoad = esegui il controllo dei MACRO Aperti _
                                                            mentre senza questa condizione, vengono iterati tutti gli oggetti _
                                                            Macro del Db. Per i MACRO APERTI O CARICATI  viene impostata la _
                                                            proprieta IsLoaded = True mentre per quelli chiusi è impostata a _
                                                            FALSE. _
                                                            È possibile utilizzare la proprietà IsLoaded per determinare se _
                                                            un oggetto AccessObject è attualmente caricato. Valore Boolean in sola lettura.


                                                                        Debug.Print "  " & obj.Name
                                                                        Debug.Print "  " & obj.Type

                                                                        '//Tipo di oggetto caricato = MACRO
                                                                        ApriRs.Fields("TIPOGGETTO_s") = "MACRO"


                                                            '//obj.Type = 4 LA MACRO
                                                            If obj.Type = 4 Then




                                                                            Debug.Print "Macro APERTO = " & obj.Name
                                                                            Debug.Print "  " & obj.Name
                                                                            Debug.Print "Nro : " & icount
                                                                            'ApriRs.Fields("PROPRIETA_s") = "True"
                                                                            'ApriRs.Fields("VALOREPROPRIETA_m") = "FORM APERTA"

                                                                             icount = icount + 1         '//aggiorno il contatore oggetti delle MACRO
                                                                             Debug.Print "Nro : " & icount

                                                                             '//NRO di oggetto
                                                                             ApriRs.Fields("IdOrd_lng") = icount




                                                            End If

                                                    '................................................................





                                                  ApriRs.Update
                                                '................................................................




                                        Next obj

                                    End With



                                '//ITERAZIONE MACRO DEFS <<I Macro>>
                                '//______________________________________________________________________________



                            Case Else
                                    MsgBox "Scelta comando non effettuata", vbExclamation
                                     GoTo Exit_RICERCA_OGGETTI_DB_PFunct
                        End Select
                                    'chiudo i rs aperti + Rilascio la memoria
                                    ApriRs.Close
                                    Set ApriRs = Nothing



        '------------------------------------------------------------------------------
        '                       FINE FUNCTIONE E GESTIONE ERRORI

        Exit_RICERCA_OGGETTI_DB_PFunct:
            dbs.Close
            Set dbs = Nothing
            Exit Function


        Err_RICERCA_OGGETTI_DB_PFunct:
            MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "RICERCA_OGGETTI_DB_PFunct"
            Resume Exit_RICERCA_OGGETTI_DB_PFunct

        End Function

        '//*** FINE ***
        '//                                     SALVATAGGIO OGGETTI DATABASE
        '//{@=============================================================================================================@}




        '//###########################################################################################################//
        '//####
        '//####                             PROCEDURE DI RICERCA   *** FINE ***
        '//####
        '//###########################################################################################################//



        '// *** FINE ***
        '//CODICE----------->SALVO_ATTIVITA_CMB_02_IN_TMP_Sub.Salvataggio_ATTIVITA
        '//SALVATAGGIO PARAMETRI NELLA TABELLA TMP
        '//###########################################################################################################################//

'//PROCEDURE DI CONTROLLO *** INIZIO ***
    '//###########################################################################################################//
    '//####
    '//####
    '//#### CONTROLLO/ CONTROLLA...
    '//###########################################################################################################//


            '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DELL'OGGETTO_FORM
            '//=======================================================================================================//
            '//CODICE-----> Controlla_SE_OGGETTO_FORM_ESISTE_pFunct.attiva
            '//note: è stata inserita nella classe CLASSE_GESTIONE_OGGETTI_FORM
            Public Function Controlla_SE_OGGETTO_FORM_ESISTE_pFunct(par_sxNameForm As String) As Boolean






                '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DELL'OGGETTO_FORM
                '//CODICE-----> Controlla_SE_OGGETTO_FORM_ESISTE_pFunct.attiva.LaClasse
                '//--------------------------------------------------------------------------------//-//
                '//NOTA  :Controlla se esiste la form negli oggetti del database e restituisce TRUE _
                          se il nome della form passata con parametro par_sxNameForm esiste; altrimenti FALSE



                            'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_FORM
                            'Assegno e creo l'oggetto di classe
                            'Dim MyCls_CLASSE_GESTIONE_OGGETTI_FORM As Classe_GESTIONE_OGGETTI_FORM
                            'Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = New Classe_GESTIONE_OGGETTI_FORM


                            'PARAMETRI DI ESTRAZIONE
                            'Attivo il metodo della classe rappresentata dalla routine passando
                            'sxNameForm = "Nome della form"
                            'MyCls_CLASSE_GESTIONE_OGGETTI_FORM.Controlla_SE_OGGETTO_FORM_ESISTE_pFunct sxNameForm

                            'Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = Nothing

                    '//--------------------------------------------------------------------------------//-//








                Dim formName_s As String
                Dim ctl As Control

                '//reset
                FormGestione_b = False


                'application.Form[GEST_FrmM00_GESTIONE].
                Dim obj As AccessObject, dbs As Object
                Set dbs = Application.CurrentProject
                ' Search for open AccessObject objects in AllForms collection.
                For Each obj In dbs.AllForms

                        ' Print name of obj.
                        Debug.Print obj.Name
                        formName_s = obj.Name

                        '//CONTROLLO SUL NOME
                        If formName_s = par_sxNameForm Then

                           '//La form esiste restituisco True
                           FormGestione_b = True
                           Controlla_SE_OGGETTO_FORM_ESISTE_pFunct = True

                        Else
                             FormGestione_b = False

                        End If


                Next obj


                        If Bool1 = False Then
                            MsgBox "ATTENZIONE LA FORM ---> " & formName_s & " NON ESISTE NEL DB CORRENTE", vbExclamation, "CONTROLLO FORM"
                        End If





                    Set obj = Nothing
                    Set dbs = Nothing
                    Set ctl = Nothing
                    formName_s = vbNull




            End Function
            '//*** FINE ***
            '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DELL'OGGETTO_FORM
            '//=======================================================================================================//




            '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DEL_SINGOLO_OGGETTO_NEL_DB_CORRENTE
            '//=======================================================================================================//
            '//CODICE----->ACONTROLLA_OGGETTO_pFunction.Funcition

            '//ATTIVO LA FUNZIONE TROVA L'OGGETTO NEL DB CORRENTE
            '//--------------------------------------------------------------------------------//---//
            '//NOTE : attivo la ricerca dell'oggetto nel DB corrente iterando in tutti gli oggetti _
                    se trova l'oggetto passato come parametro, tipo form, modulo, query od altro _
                    ritrona il parametro OGGETTO TROVATO con il quale viene effettuato un controllo _
                    con la if che da il permesso di attivare la procedura della classe. Se il parametro _
                    ritorna NON TROVATO l'attivà della classe viene saltata esplicitamente.
            '//
            '//CODICE----->ACONTROLLA_OGGETTO_pFunction.Funcition.ATTIVA
                'Dim ObjDaTrovare_s As String
                'ObjDaTrovare_s = ""
                'Str1 = CONTROLLA_OGGETTO_pFunction(ObjDaTrovare_s)
                    'If ObjDaTrovare_s = "OGGETTO TROVATO" Then
                        'eseguo l'attività della classe altrimenti non viene eseguita
                        '...
                    'End If

            '//Controlla_SE_OGGETTO_TABELLA_ESISTE_pFunct = CONTROLLO ESISTENZA OGGETTO TABELLA
            '//--------------------------------------------------------------------------------//---//
            Public Function Controlla_SE_OGGETTO_TABELLA_ESISTE_pFunct(par_sxNameTABELLA As String)


                Dim obj As AccessObject
                Dim dbs As Object

                Dim objAllTables As AllTables

                  '//Variabile database corrente + reset
                Set dbs = Application.CurrentData
                Controlla_SE_OGGETTO_TABELLA_ESISTE_pFunct = False            '//imposto a ricerca fallita


                'Search for open AccessObject objects in AllForms collection.
                For Each obj In dbs.AllTables

                        ' Print name of obj.
                        Debug.Print obj.Name
                        TableName_s = obj.Name

                        '//CONTROLLO SUL NOME
                        If TableName_s = par_sxNameTABELLA Then

                           '//Se esiste la TABELLA restituisco True

                           Controlla_SE_OGGETTO_TABELLA_ESISTE_pFunct = True

                                '//controllo TABELLE ESISTENTE
                                 Debug.Print "----------------------LA TABELLA XLS ESISTE --------------------------------------"
                                 Debug.Print
                                 Debug.Print "tabella xls passata con parametro par_sxNameTABELLA ESISTE nel db               ----> " & par_sxNameTABELLA


                           Exit For

                        End If

                 Next obj


            End Function

            '//*** FINE ***
            '//Controlla_SE_OGGETTO_TABELLA_ESISTE_pFunct = CONTROLLO ESISTENZA OGGETTO TABELLA
            '//--------------------------------------------------------------------------------//---//


            '//Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct = CONTROLLO ESISTENZA OGGETTO QUERY
            '//--------------------------------------------------------------------------------//---//
            Public Function Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(par_sxNameQuery As String)


                Dim obj As AccessObject
                Dim dbs As Object

                Dim objAllQueries As AllQueries

                  '//Variabile database corrente + reset
                Set dbs = Application.CurrentData
                Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct = False            '//imposto a ricerca fallita


                'Search for open AccessObject objects in AllForms collection.
                For Each obj In dbs.AllQueries

                        ' Print name of obj.
                        Debug.Print obj.Name
                        formName_s = obj.Name

                        '//CONTROLLO SUL NOME
                        If formName_s = par_sxNameQuery Then

                           '//Se esiste la query restituisco True

                           Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct = True

                           Exit For

                        End If

                 Next obj


            End Function

            '//*** FINE ***
            '//Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct = CONTROLLO ESISTENZA OGGETTO QUERY
            '//--------------------------------------------------------------------------------//---//

            Public Function CONTROLLA_OGGETTO_pFunction(ByVal NameObject_s As String) As String


                    Dim obj As Object
                    Dim frm As Form
                        Dim ctrl As Control

                        '//impsto il valore a non trovato
                        CONTROLLA_OGGETTO_pFunction = "NON TROVATO"

                        For Each obj In CurrentProject.AllForms
                            If InStr(obj.Name, "MODELLO_Frm01_M01_GE_FILTRO_M15_H15_L35_NoButton") Then Debug.Print "Forms|" & obj.Name

                                'Trovato l'oggetto nel db corrente
                                CONTROLLA_OGGETTO_pFunction = "TROVATO"

                            'Set frm = Forms(obj.Name) '- this had problems as well, so I commented out
                            For Each ctrl In Forms(obj.Name).Controls
                                If InStr(ctrl.Name, "LBL_FILTRO_01") Then Debug.Print "Form|" & obj.Name & "|" & ctrl.Name
                            Next ctrl
                        Next obj

                    End Function



            '//*** FINE ***
            '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DEL_SINGOLO_OGGETTO_NEL_DB_CORRENTE
            '//=======================================================================================================//



            '//CODICE-->Controlla_SE_OGGETTO_FORM_E_APERTA_pFunct.01
            '//FUNZIONE_DI_CONTROLLO_DELLA_FORM_GESTIONE_APERTA.01
            '//=======================================================================================================//
            '//SI TROVA NEL MODULO -> GEST_FrmM00_GESTIONE_Mdl01_ControlloFormAperta

            Public Function Controlla_SE_OGGETTO_FORM_E_APERTA_pFunct(NameForm_s As String) As Boolean


            On Error GoTo Controlla_SE_OGGETTO_FORM_E_APERTA_pFunct_Err



                    '//FUNZIONE DELLA CLASSE PER IL CONTROLLO SE LA FORM E' APERTA
                    '//############################################################################################################################//

                            '//ISTANZIO LA CLASSE PER IL CONTROLLO DELLA FORM APERTA
                            '//.....................................................................................................


                                '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe e le altre variabili
                                    '//--> Dim MyCls_CLASSE_ATTI_Tb01_Gestione As CLASSE_ATTI_Tb01_Gestione
                                    '//--> Dim FormAperta_b As Boolean

                                '// RESET DELLE VARIABILI
                                    '//--> FormAperta_b = false

                                '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_FORM assegnando _
                                   l 'oggetto di classe alla variabile e creo creo l'oggetto di classe
                                    '//--> Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione

                                '//PATH O FILE = Inserire o una path oppure una path+file con percoso esteso.
                                    'Str1 = "**** DEVI INSERIRE IL NOME DELLA FORM DA CONTROLLARE ***"

                                    '// LA FUNZIONE RESTITUISCE TRUE SE E' APERTA ALTRIMENTI FALSE
                                    '//--> FormAperta_b = MyCls_CLASSE_ATTI_Tb01_Gestione.Controlla_SE_OGGETTO_FORM_E_APERTA_pFunct (Str1)



                                '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
                                    '//--> Set MyCls_CLASSE_ATTI_Tb01_Gestione = Nothing


                            '//.....................................................................................................


                    '//*** FINE ***
                    '//FUNZIONE DELLA CLASSE PER IL CONTROLLO SE LA FORM E' APERTA
                    '//############################################################################################################################//




            '//FUNZIONE_DI_CONTROLLO_DELLA_FORM_GESTIONE_APERTA.01_Esegui
            '//__________________________________________________________________________________________//

                Dim formName_s As String
                Dim ctl As Control

                '//reset
                FormGestione_b = False


                'application.Form[GEST_FrmM00_GESTIONE].
                Dim obj As AccessObject, dbs As Object
                Set dbs = Application.CurrentProject
                ' Search for open AccessObject objects in AllForms collection.
                For Each obj In dbs.AllForms

                    If obj.IsLoaded = True Then
                        ' Print name of obj.
                        Debug.Print obj.Name
                        formName_s = obj.Name

                        '//CONTROLLO SUL NOME
                        If formName_s = "GEST_FrmM00_GESTIONE" Then

                           '//FormGestione_b = true per indicare che la form è aperta e restituisco il nome della form _
                           alla routine chiamante
                           FormGestione_b = True
                           Controlla_SE_OGGETTO_FORM_E_APERTA_pFunct = True

                        End If

                    End If
                Next obj

                    Set obj = Nothing
                    Set dbs = Nothing
                    Set ctl = Nothing
                    formName_s = vbNull

            '//FUNZIONE_DI_CONTROLLO_DELLA_FORM_GESTIONE_APERTA.01_Esegui *** fine ***
            '//__________________________________________________________________________________________//




            '// USCITA E CONTROLLO ERRORI
            '//__________________________________________________________________________________________//


    Controlla_SE_OGGETTO_FORM_E_APERTA_pFunct_Exit:
                Exit Function

    Controlla_SE_OGGETTO_FORM_E_APERTA_pFunct_Err:
                MsgBox Error$
                Resume Controlla_SE_OGGETTO_FORM_E_APERTA_pFunct_Exit


            End Function
            '//*** FINE ***
            '//FUNZIONE_DI_CONTROLLO_DELLA_FORM_GESTIONE_APERTA
            '//=======================================================================================================//




    '//###########################################################################################################//
    '//####
    '//####                             PROCEDURE DI CONTROLLO *** FINE ***
    '//#### CONTROLLO/ CONTROLLA...
    '//###########################################################################################################//

'//TUTTE LE ESPORTAZIONI'
    '//@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@//
    '//                                             ESPORTAZIONI DATI PER IL SETTORE LAVORI PUBBLICI
    '//
    '//NOTE :ESPORTA LE TABELE XLS ESISTENTI        14.08.2020 del settore lavori pubblici.

    '//      ATTIVA IN VISUALIZZAZIONE LE QUERY ESISTENTI
    '//     TODO: non gestisce le tabelle collegate, da errore se esistono. Per il momento le tabelle collegate non funzionanti _
                  sono state cancellate. _
                  GESTIRE LE TABELLE COLLEGATE!!!!
    '//@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@//




    '//AGGIORNA I DATI CON LE QUERY ED ESPORTA IN MDB + XLS
    '//****************************************************************************************************************//
        '------------------------------------------------------------
        ' CODICE----->AggiornaDati_ed_EsportaDati_Pfunct.ATTIVA
        ' AggiornaDati_ed_EsportaDati_Pfunct
        ' si trova in questo modulo                 ----> LLPP_ATTI_Mdl01_AGGIORNA_Mcr01_01_AggiornaDati+EsportaDati
        ' ma sistemato all'interno della classe     ----> CLASSE_ATTI_Tb01_Gestione _
        ' perchè all'interno della classe posso fare un controllo degli oggetti prima di _
        ' attivare il comando Cmd; se non esistono nel db corrente da errore; ma _
        ' con il controllo degli oggetti puoi attivare quelli che sono presenti.
        '------------------------------------------------------------
        Public Function AggiornaDati_ed_EsportaDati_Pfunct()

            '//Variabile per le tabelle ESPORTATE E NON ESPORTATE
            Dim MsgTabelle_ESPORTATE_s As String
            Dim iCount_TABELLE_ESPORTATE As Integer
            Dim MsgTabelle_NON_ESPORTATE_s As String
            Dim iCount_NON_TABELLE_ESPORTATE As Integer


            Dim OggettoDaRicercare_s   As String
            Dim icount As Integer

            On Error GoTo AggiornaDati_ed_EsportaDati_Pfunct_Err




                '//RESET
                MsgTabelle_NON_ESPORTATE_s = ""
                icount = 0

            ' AGGIORNAMENTI + ACCODAMENTI
            ' ***********************************************************************************************************************
            ' Note  : Eseguo l'aggiornamento dei vari argomenti di controllo.


                '//PRES3000 AGGIORNAMENTI TABELLE PRESENZE MEDIANTE LA MACRO
                '//=========================================================================================================



                    '//RIRCERCA OGGETTO NEL DB CORRENTE - Query da ricercare
                    '//.......................................................................................
                    '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                        '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                           di ricerca secondo il tipo di oggetto

                        Bool1 = False
                        OggettoDaRicercare_s = "PRES3000_Qry01_10_AGGIORNA_Campo_IDKEY_CALENDARIO"      '//Aggiorno ID nel campo idKey del calendario, per preservare la chiave univoca.
                        Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                            '//SE TRUE = oggetto esistente ed eseguo il comando
                            If Bool1 Then
                                '//!!!! ATTENZIONE NON GESTISCE LE TABELLE COLLEGATE SISTEMARE ????????

                                '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                                DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                    '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                    iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                    MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                            Else

                                    '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                    icount = iCount_NON_TABELLE_ESPORTATE + 1
                                    MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                            End If


                   '//.......................................................................................



                    '//*** AGGIORNO LE TIMBRATURE ***
                    '//RIRCERCA OGGETTO NEL DB CORRENTE - Query da ricercare
                    '//.......................................................................................
                    '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                        '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                           di ricerca secondo il tipo di oggetto

                        Bool1 = False
                        OggettoDaRicercare_s = "PRES3000_Qry01_11_AGGIORNA_Campo_IDKEY_GIORNALIERE"  '//Aggiorno ID nel campo timbrature giornaliere
                        Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                            '//SE TRUE = oggetto esistente ed eseguo il comando
                            If Bool1 Then

                                  '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                                   DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                    '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                    iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                    MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                            Else

                                    '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                    icount = iCount_NON_TABELLE_ESPORTATE + 1
                                    MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                            End If


                   '//.......................................................................................


                '//PRES3000 AGGIORNAMENTI TABELLE PRESENZE MEDIANTE LA MACRO*** FINE ***
                '//=========================================================================================================




                '//AGGIORNAMENTI
                '//=========================================================================================================


                    '//INTESTATARI Aggiorna tabella INTESTARI - campo ->Argomento
                    '//================================================================================================================================//

                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "AGGIORNA_INTESTATARI_QyN01-01_Campo_Argomento"    '//AGGIORNA = la tabella INTESTARI - campo ->Argomento
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.


                                    DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                    '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                    iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                    MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                            Else

                                    '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                    icount = iCount_NON_TABELLE_ESPORTATE + 1
                                    MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                            End If



                           '//.......................................................................................


                   '//================================================================================================================================//


                    '//TABELLA ATTI AGGIORNAMENTI
                    '//************************************************************************************************************************************//


                            '//01) Aggiorna tabella TABELLA ATTI - campo ->IDKeyArch+IdOrd e contestualmente aggiorno la descrizione opera - campo ->DescrCodOpera_s
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "LLPP_ATTI_Qy01-01_Gestione_AGGIORNA_IDKeyArch+IdOrd"    '//AGGIORNA = campi ->IDKeyArch+IdOrd+DescrCodOpera_s
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.


                                            DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If

                           '//.......................................................................................




                            '//02) Aggiorna tabella TABELLA ATTI - campo ->DenomFascicolo_s e la denominazione del fascicolo prendendo i dati dalla tabella [LPP_DF06_FASCICOLI].[DenomFascicolo_s]
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "LLPP_ATTI_Qry01-06_Gestione_AGGIORNA_DenomFascicolo"    '//AGGIORNA = campi ->DenomFascicolo_s CON IL CAMPO [LPP_DF06_FASCICOLI].[DenomFascicolo_s]
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                                             DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If
                           '//.......................................................................................



                            '//03) Aggiorna tabella TABELLA ATTI - campo ->ANNO l'anno di creazione del documento viene aggiornato solo se null
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "LLPP_ATTI_Qy01-11_AGGIORNA_ANNO"    '//AGGIORNA = campo ->ANNO se > null
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                                            DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If

                           '//.......................................................................................



                    '//ATTI *** fine ***
                    '//************************************************************************************************************************************//




                    '//FALDONI + FASCICOLI *** fine ***
                    '//************************************************************************************************************************************//




                            '//01) Aggiorna tabella TABELLA ATTI - campo ->->CodDenomFaldone_s + DenomFaldone_s + AnnoFaldone_i e la denominazione DEL FALDONE
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "LLPP_ATTI_Qry01-06_Gestione_AGGIORNA_DenomFALDONE"    '//AGGIORNA = campi ->CodDenomFaldone_s + DenomFaldone_s + AnnoFaldone_i
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                                            DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If

                           '//.......................................................................................





                            '//02) Aggiorna tabella TABELLA ATTI - campo ->CodDenomFaldone_s (la denominazione DEL FASCICOLO)
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "LLPP_ATTI_Qry01-06_Gestione_AGGIORNA_DenomFascicolo"    '//AGGIORNA = campo ->CodDenomFaldone_s
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                                            DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If

                            '//.......................................................................................








                    '//FALDONI + FASCICOLI *** fine ***
                    '//************************************************************************************************************************************//





                    '//ALLEGATI
                    '//'//************************************************************************************************************************************//



                            '//01) Aggiorna tabella TABELLA ATTI - campo ->CodOpera_s
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "LLPP_ATTI_Qy02_01_Allegati_AGGIORNA_CodOpera"    '//AGGIORNA = campo ->CodOpera_s
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
                                            DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If

                            '//.......................................................................................


                            '//02) Aggiorna tabella TABELLA ATTI - i campi ->IDGestione + AnnoAtto_i + CodPratica_s +NroPratica_i
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "LLPP_ATTI_Qry02_01_Allegati_AGGIORNA_IDGestione"    '//AGGIORNA = campi ->IDGestione + AnnoAtto_i + CodPratica_s +NroPratica_i
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                                            DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If

                            '//.......................................................................................


                            '//03) Aggiorna tabella TABELLA ATTI - i campi ->NroPratica (il campo viene aggiornato con 0 (zero) solo se nullo)
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "LLPP_ATTI_Qy01_01_Gestione_AGGIORNA_Campo_NroPratica_s=Null"    '//AGGIORNA = campi ->NroPratica
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                                            DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If

                            '//.......................................................................................



                   '//************************************************************************************************************************************//


                    '//AFFIDAMENTI
                    '//************************************************************************************************************************************//


                            '//01) Aggiorna tabella TABELLA ATTI - i campi -> I dati degli affidamenti
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "AGGIORNA_AFFIDAMENTI_QyN10-01_Campi_TipoAtto+NroAtto+DataAtto"    '//AGGIORNA = campi -> AFFIDAMENTI vari campi
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
                                            DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If

                            '//.......................................................................................

                    '//************************************************************************************************************************************//



                    '//ACCODAMENTI
                    '//************************************************************************************************************************************//



                            '//01) Aggiorna tabella TABELLA ATTI - i campi -> I dati degli affidamenti
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "Accoda_QyDF15-01_NwEVENTI_Ins"    '//ACCODO = campi -> Nuovi eventi
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                                            DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If

                            '//.......................................................................................




                            '//02) ACCODO I NUOVI CODICI - LLPP_QyDF01_CodiceOpera_ACCODA_CODICI_Ins
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "LLPP_QryDF01_CodiceOpera_ACCODA_CODICI_Ins"    '//ACCODO = nuovi codice opera dalla tabelal GESTIONE campi -> Codici Opera
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.
                                            DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If

                            '//.......................................................................................

                    '//*** FINE ***
                    '//ACCODAMENTI
                    '//************************************************************************************************************************************//




                '//TIMBRATURE - PRES3000
                '//************************************************************************************************************************************//
                '//NOTE : accodo timbrature ed orari nella tabella PRES3000


                            '//01) ACCODO NUOVE ENTRATE  nella tabella orari.
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "PRES3000_Qry05_N01_ACCODA_Nw_ENTRATE->Orari"    '//ACCODO = I NUOVI ORARI
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                                            DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If
                            '//.......................................................................................


                            '//02) ACCODO NUOVE ENTRATE  nella tabella orari.
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                OggettoDaRicercare_s = "PRES3000_Qry05_N02_ACCODA_Nw_USCITE->Orari"    '//ACCODO = LE NUOVE USCITE
                                Bool1 = Controlla_SE_OGGETTO_QUERY_ESISTE_pFunct(OggettoDaRicercare_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '//ESEGUO IL COMANDO Docmd con 3 parametri, il nome della query + 2 variabili oggetto : _
                                        acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
                                        acEdit              1  L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record.

                                            DOCMDOpenQuery_Sub OggettoDaRicercare_s, 0, 1

                                            '//CONTEGGIO E MSG DELLE TABELLE ESPORTATE
                                            iCount_TABELLE_ESPORTATE = iCount_TABELLE_ESPORTATE + 1
                                            MsgTabelle_ESPORTATE_s = iCount_TABELLE_ESPORTATE & ") " & MsgTabelle_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    Else

                                            '//CONTEGGIO E MSG DELLE TABELLE NON ESPORTATE
                                            icount = iCount_NON_TABELLE_ESPORTATE + 1
                                            MsgTabelle_NON_ESPORTATE_s = iCount_NON_TABELLE_ESPORTATE & ") " & MsgTabelle_NON_ESPORTATE_s & OggettoDaRicercare_s & Chr$(13)
                                    End If

                            '//.......................................................................................






                '//*** FINE ***
                '//TIMBRATURE
                '//************************************************************************************************************************************//

            ' ***********************************************************************************************************************


            'ESPORTAZIONI
            '***********************************************************************************************************************

                '//TABELLA GESTIONE
                '//============================================================================================================


                   '//ESPORTO ATTI DI GESTIONE
                   '//Esporto la tabella con la specifica
                   '//ESPORTA I DATI ---------------------------------------------------------------------------------------------------------------------------------------> False = non visualizza il file
                   '//SOSPESA
                   'DoCmd.TransferText acExportDelim, "LLPP_ATTI_Tb01_Gestione - specifica di esportazione", "LLPP_ATTI_Tb01_Gestione", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ATTI_Tb01_Gestione.txt"
                   'Esporto tabella in dos
                   'DoCmd.OutputTo acOutputTable, "LLPP_ATTI_Tb01_Gestione", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ATTI_Tb01_Gestione_DOS.txt", False, "", 0, acExportQualityPrint
                   '//Apro la tabella esportata
                   'DoCmd.OpenTable "LLPP_ATTI_Tb01_Gestione", acViewNormal, acEdit



                   '//TABELLA ALLEGATI
                   '//Esporto la tabella con la specifica - LLPP_ATTI_Tb02_Allegati - specifica di esportazione
                   '//SOSPESA
                   'DoCmd.TransferText acExportDelim, "LLPP_ATTI_Tb02_Allegati - specifica di esportazione", "LLPP_ATTI_Tb02_Allegati", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ATTI_Tb02_Allegati.txt"
                   '//Esporto tabella in dos
                   'DoCmd.OutputTo acOutputTable, "LLPP_ATTI_Tb02_Allegati", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ATTI_Tb02_Allegati_Dos.txt", False, "", 0, acExportQualityPrint
                   '//Apro la tabella esportata
                   'DoCmd.OpenTable "LLPP_ATTI_Tb02_Allegati", acViewNormal, acEdit


                   '//GLI INDIRIZZI
                   '//........................................................................................................
                        '//TABELLA INTESTATARI
                        '//Esporto la tabella con la specifica - INDIRIZZI_Tb01_INTESTATARI - specifica di esportazione
                        '//SOSPESA
                        'DoCmd.TransferText acExportDelim, "INDIRIZZI_Tb01_INTESTATARI - specifica di esportazione", "Indirizzi_Tb01_INTESTATARI", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\INDIRIZZI_Tb01_INTESTATARI.txt"
                        '//Esporto tabella in dos
                       ' DoCmd.OutputTo acOutputTable, "Indirizzi_Tb01_INTESTATARI", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\INDIRIZZI_Tb01_INTESTATARI_Dos.txt", False, "", 0, acExportQualityPrint
                        '//Apro la tabella esportata
                        'DoCmd.OpenTable "Indirizzi_Tb01_INTESTATARI", acViewNormal, acEdit




                        '//TABELLA INDIRIZZI ELENCO
                        '//Esporto la tabella con la specifica - INDIRIZZI_Tb02_ELENCO - specifica di esportazione
                        '//SOSPESA
                        'DoCmd.TransferText acExportDelim, "INDIRIZZI_Tb02_ELENCO - specifica di esportazione", "Indirizzi_Tb02_ELENCO", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\INDIRIZZI_INDIRIZZI_Tb02_ELENCO.txt"
                        '//Esporto tabella in dos
                        'DoCmd.OutputTo acOutputTable, "Indirizzi_Tb02_ELENCO", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\INDIRIZZI_INDIRIZZI_Tb02_ELENCO_Dos.txt", False, "", 0, acExportQualityPrint
                        '//Apro la tabella esportata
                        'DoCmd.OpenTable "Indirizzi_Tb02_ELENCO", acViewNormal, acEdit

                      '//........................................................................................................


                'Indirizzi_Tb02_ELENCO - specifica di esportazione

                '//TABELLA GESTIONE *** FINE ***
                '//============================================================================================================

                '//PRES3000
                '//============================================================================================================

                    '//IL CALENDARIO
                    'Esporto la tabella con la specifica
                    'ESPORTA I DATI ---------------------------------------------------------------------------------------------------------------------------------------> False = non visualizza il file
                    '//SOSPESA
                    'DoCmd.TransferText acExportDelim, "PRES3000_Tb01_Calendario - specifica di esportazione", "PRES3000_Tb01_Calendario", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_Tb01_Calendario.txt"
                    'Esporto tabella in dos
                    'DoCmd.OutputTo acOutputTable, "PRES3000_Tb01_Calendario", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_Tb01_Calendario_DOS.txt", False, "", 0, acExportQualityPrint
                    '//Apro la tabella esportata
                    'DoCmd.OpenTable "PRES3000_Tb01_Calendario", acViewNormal, acEdit


                    '//LE GIORNALIERE - PRES3000_Tb02_ElencoGiornaliere
                    'Esporto la tabella con la specifica
                    'ESPORTA I DATI ---------------------------------------------------------------------------------------------------------------------------------------> False = non visualizza il file
                    '//SOSPESA
                    'DoCmd.TransferText acExportDelim, "PRES3000_Tb02_ElencoGiornaliere - specifica di esportazione", "PRES3000_Tb02_ElencoGiornaliere", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_Tb02_ElencoGiornaliere.txt"
                    'Esporto tabella in dos
                    'DoCmd.OutputTo acOutputTable, "PRES3000_Tb02_ElencoGiornaliere", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_Tb02_ElencoGiornaliere_DOS.txt", False, "", 0, acExportQualityPrint
                    '//Apro la tabella esportata
                    'DoCmd.OpenTable "PRES3000_Tb02_ElencoGiornaliere", acViewNormal, acEdit



                    '//LA DEFINIZIONI

                    'LLPP_DF01_CodiceOpera

                    'ESPORTA I DATI ---------------------------------------------------------------------------------------------------------------------------------------> False = non visualizza il file
                    '//SOSPESO
                    'DoCmd.TransferText acExportDelim, "LLPP_DF01_CodiceOpera - specifica di esportazione", "LLPP_DF01_CodiceOpera", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_DF01_CodiceOpera.txt"
                    'Esporto tabella in dos
                   ' DoCmd.OutputTo acOutputTable, "LLPP_DF01_CodiceOpera", "MS-DOSText(*.txt)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_DF01_CodiceOpera_DOS.txt", False, "", 0, acExportQualityPrint
                    '//Apro la tabella esportata
                    'DoCmd.OpenTable "LLPP_DF01_CodiceOpera", acViewNormal, acEdit




                '//PRES3000
                '//============================================================================================================


                '//ESPORTA_I_FILE_IN_FORMATO_XLS
                '//============================================================================================================


                            '//ATTIVO LA PROCEDURA DI ESPORTAZIONE DEI FILE IN FORMATO XLS DEGLI ATTI DI IMPEGNO
                            '//.................................................................................................
                            '//NOTE: Dim + Reset ed impostazioni

                            '//Chiamo la procedura di esportazione indicando il nome delle tabelle da esportare in xls
                            'ESPORTA_FILE_XLS_ConMatriceDati_pFunct "", _
                            "LLPP_ATTI_Tb01_Gestione", _
                            "LLPP_ATTI_Tb02_Allegati", _
                            "Indirizzi_Tb01_INTESTATARI", _
                            "Indirizzi_Tb02_ELENCO", _
                            "PRES3000_Tb01_Calendario", _
                            "PRES3000_Tb02_ElencoGiornaliere", _
                            "LLPP_DF01_CodiceOpera", _
                            "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", _
                            "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA"

            '//ESPORTA I DATI IN FILE XLS
            '//........................................................



               '//Chiamo la procedura di esportazione indicando il nome delle tabelle da esportare in xls _
                  ATTENZIONE, occorre inserire un record vuoto "", affinche possono essere considerati tutti gli altri _
                  file dell'esportazione altrimenti di viene saltato il primo "LLPP_ATTI_Tb01_Gestione", _
                  PER ESPORTARE CHIAMARE QUESTA FUNZIONE : CODICE------> ESPORTA_FILE_XLS_ConMatriceDati_pFunct

                            ESPORTA_FILE_XLS_ConMatriceDati_pFunct "", _
                            "LLPP_ATTI_Tb01_Gestione", _
                            "LLPP_ATTI_Tb02_Allegati", _
                            "Indirizzi_Tb01_INTESTATARI", _
                            "Indirizzi_Tb02_ELENCO", _
                            "PRES3000_Tb01_Calendario", _
                            "PRES3000_Tb02_ElencoGiornaliere", _
                            "LLPP_DF01_CodiceOpera", _
                            "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", _
                            "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA"

            '//........................................................






                             '//APRO LE  TABELLE ESPORTATE  --- SOSPESO
                             'DoCmd.OpenTable "LLPP_ATTI_Tb01_Gestione", acViewNormal, acEdit
                             'DoCmd.OpenTable "LLPP_ATTI_Tb02_Allegati", acViewNormal, acEdit
                             'DoCmd.OpenTable "Indirizzi_Tb01_INTESTATARI", acViewNormal, acEdit
                             'DoCmd.OpenTable "Indirizzi_Tb02_ELENCO", acViewNormal, acEdit
                             'DoCmd.OpenTable "PRES3000_Tb01_Calendario", acViewNormal, acEdit
                             'DoCmd.OpenTable "PRES3000_Tb02_ElencoGiornaliere", acViewNormal, acEdit
                             'DoCmd.OpenTable "LLPP_DF01_CodiceOpera", acViewNormal, acEdit
                             'DoCmd.OpenTable "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", acViewNormal, acEdit
                             'DoCmd.OpenTable "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA", acViewNormal, acEdit


                            '//.................................................................................................

                '//ESPORTA_I_FILE_IN_FORMATO_XLS         *** FINE ***
                '//============================================================================================================


            '***********************************************************************************************************************


                        If MsgTabelle_ESPORTATE_s > "" Then

                            MsgBox "LE TABELLE ESPORTATE SONO LE SEGUENTI: --->  " & MsgTabelle_ESPORTATE_s, vbExclamation
                        End If


                        If MsgTabelle_NON_ESPORTATE_s > "" Then

                            MsgBox "ATTENZIONE QUESTI OGGETTI NON SONO STATI TROVATI --->  " & MsgTabelle_NON_ESPORTATE_s, vbExclamation
                        End If


                MsgBox "FINE ATTIVITA' DI Esportazione, aggiornamento ed accodamento. Operazioni effettuate!", vbExclamation


        AggiornaDati_ed_EsportaDati_Pfunct_Exit:
                Exit Function

        AggiornaDati_ed_EsportaDati_Pfunct_Err:
                MsgBox Error$
                Resume AggiornaDati_ed_EsportaDati_Pfunct_Exit

            End Function







            '// *** FINE ***
            '//             AGGIORNA I DAITI CON LE QUERY ED ESPORTA IN MDB + XLS
            '//****************************************************************************************************************//




            '// *** FINE ***
            '//             I COMANDI CMD :
            '//                             DoCmdOpenQuery_Sub
            '//
            '//****************************************************************************************************************//
            '//@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@//

            '//DOCMD_OPEN_QUERY = Eseguo il comando docmd di esecuzione della query la cui _
                                esistenza viene controllata prima di chiamare questa routine. _
                                I parametri sono due: _
            1 paramentro par_acViewNormal, corrisponde alla variabile AcView sono i seguenti e devono essere indicati in par_acViewNormal _
            Nome Valore         Descrizione _
            acViewDesign        1 Visualizzazione Struttura _
            acViewLayout        6 Visualizzazione Layout _
            acViewNormal        0 (Impostazione predefinita) Visualizzazione Normale _
            acViewPivotChart    4 Visualizzazione Grafico pivot _
            acViewPivotTable    3 Visualizzazione Tabella pivot _
            acViewPreview       2 Anteprima di stampa _
            acViewReport        5 Visualizzazione Report _

            'Il 2 parametro, par_acEdit Deve indicare le modalita di apertura dell'oggetto: _
            Costante di enumerazione AcOpenDataMode _
            Specifica la modalità di immissione dati all'apertura di una tabella, una query, _
            una stored procedure, una visualizzazione o una funzione definita dall'utente. _
            Nome Valore     Descrizione _
            acAdd               0       L'utente può aggiungere nuovi record ma non visualizzare né modificare i record esistenti. _
            acEdit              1       L'utente può visualizzare o modificare i record esistenti e aggiungere nuovi record. _
            acReadOnly          2       L'utente può solo visualizzare i record.



        Public Sub DOCMDOpenQuery_Sub(par_Query_s As String, par_acViewNormal As Integer, par_acEdit As Integer)

                                '//Eseguo la query passata come parametro solo se valorizzata _
                                   la variabile che contiene il nome della query. Il controllo se _
                                   la query esiste viene fatto a prima della routine.
                                If par_Query_s > "" Then
                                        DoCmd.OpenQuery par_Query_s
                                Else
                                    MsgBox "Comando Cmd.OpenQuery non eseguito in quanto il parametro oggetto 'par_Query_s' è nullo:  --> " & par_Query_s
                                End If

            End Sub




















                                                '//*************************************//
                                                '//     LE ESPORTAZIONI
                                                '//*************************************//
                                                '
            '//**************************************************************************************************************************//
            '//         FUNZIONI DI ESPORTAZIONI LAVORI PUBBLICI IN FORMATO XLS
            '//**************************************************************************************************************************//
            '//

            '//01)_ESPORTA FILE XLS CON MATRICE = Funzione di esportazione dei dati dei lavori pubblici _
                    con esportazioni in tabelle excel. Viene utilizzata una matrice.
            '//     CODICE------> ESPORTA_FILE_XLS_ConMatriceDati_pFunct



        '//ESPORTO SU FILE ESTERNO _ per ora un xls
        Public Sub OutputTo_Sub(par_Tabella_s As String, par_Path_s As String, par_NuovaDenominazioneTabella_01_s As String, par_Bool As Boolean)

                    '//MODELLO DI ESPORTAZIONE
                    '//Esporto le tabelle in formato xls
                    '                       par_Tabella_s                          par_PATH_S                               par_Bool
                    'DoCmd.OutputTo acTable, Tabella_01_s, "MicrosoftExcel(*.xls)", PATH_s & NuovaDenominazioneTabella_01_s, False, ""
                    'MODELLO FUNZIONANTE ---> DoCmd.OutputTo acTable, "LLPP_ATTI_Tb01_Gestione", "MicrosoftExcel(*.xls)", "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LPP_ARCHIVI_SALVATAGGI\LLPP_ATTI_Tb01_Gestione.xls", False, ""
                    '//Controllo parametri:
                        Debug.Print "par_Tabella_s                              ----> " & par_Tabella_s
                        Debug.Print "par_PATH_S                                 ----> " & par_Path_s
                        Debug.Print "par_NuovaDenominazioneTabella_01_s         ----> " & par_NuovaDenominazioneTabella_01_s

                    DoCmd.OutputTo acTable, par_Tabella_s, "MicrosoftExcel(*.xls)", par_Path_s & par_NuovaDenominazioneTabella_01_s, par_Bool, ""

            End Sub






            '//**************************************************************************************************************************//

            '//FUNZIONE ESPORTA FILE XLS
            '//Chiamo matrice con Stringhe che rappresentano delle tabelle _
            Impegni , Presenze
            Private Sub CHIAMA_ESPORTA_FILE_XLS_ConMatriceDati_pFunct()

                '//Chiamo la procedura di esportazione indicando il nome delle tabelle da esportare in xls
                ESPORTA_FILE_XLS_ConMatriceDati_pFunct "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA", _
                "PRES3000_Tb01_Calendario", "PRES3000_Tb02_ElencoGiornaliere", "PRES3000_Tb01_Calendario", "Indirizzi_Tb01_INTESTATARI", _
                "Indirizzi_Tb02_ELENCO"

            End Sub


        '//01)_ESPORTA FILE XLS CON MATRICE = Funzione di esportazione dei dati dei lavori pubblici _
                    con esportazioni in tabelle excel. Viene utilizzata una matrice.
            '//     CODICE------> ESPORTA_FILE_XLS_ConMatriceDati_pFunct
            '//=================================================================================================//
            '//NOTE ------------->  : attenzione esporto la collezione di file ma il primo deve essere nullo ("") per _
                                    evitare che venga saltato.

            Public Function ESPORTA_FILE_XLS_ConMatriceDati_pFunct(strName As String, ParamArray intScores() As Variant)
                Dim PATH_s As String
                Dim Tabella_01_s As String
                Dim NuovaDenominazioneTabella_01_s As String
                Dim Tabella_02_s As String
                Dim NuovaDenominazioneTabella_02_s As String
                Dim intI As Integer
                Dim icount As Integer                   '//il contatore

                On Error GoTo ESPORTA_FILE_XLS_ConMatriceDati_pFunct_Err


                '//ESPORTA I DATI IN FILE XLS
                '//........................................................



                   '//Chiamo la procedura di esportazione indicando il nome delle tabelle da esportare in xls _
                      ATTENZIONE, occorre inserire un record vuoto "", affinche possono essere considerati tutti gli altri _
                      file dell'esportazione altrimenti di viene saltato il primo "LLPP_ATTI_Tb01_Gestione", _
                                Call ESPORTA_FILE_XLS_ConMatriceDati_pFunct "", _
                                "LLPP_ATTI_Tb01_Gestione", _
                                "LLPP_ATTI_Tb02_Allegati", _
                                "Indirizzi_Tb01_INTESTATARI", _
                                "Indirizzi_Tb02_ELENCO", _
                                "PRES3000_Tb01_Calendario", _
                                "PRES3000_Tb02_ElencoGiornaliere", _
                                "LLPP_DF01_CodiceOpera", _
                                "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO", _
                                "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA"

                '//........................................................


                '//Reset ed impostazioni
                '//........................................................
                    PATH_s = vbNull
                    PATH_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LPP_ARCHIVI_SALVATAGGI\"

                    Tabella_01_s = vbNull
                    Tabella_01_s = "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO"

                '//........................................................

                    Debug.Print
                    Debug.Print strName; " Punteggi"
                    ' Utilizza la funzione UBound per definire il
                    ' limite superiore della matrice.

                    '//Dimensione la matrice di file

                     Dim MatriceFile_s As String

                    icount = 0



                        '//ATTIVAZIONE CLASSE
                        '//ROUTINE_DI_ATTIVAZIONE_APERTURA_FILE
                        '//CODICE-----> Controlla_SE_OGGETTO_FORM_ESISTE_pFunct.attiva.LaClasse
                        '//--------------------------------------------------------------------------------//-//
                        '//NOTA  :Controlla se esiste la form negli oggetti del database


                                'If Str1 <> "OGGETTO NON TROVATO" Then
                                     '
                                        'Dim MyCls_CLASSE_ATTI_Tb01_Gestione As New CLASSE_ATTI_Tb01_Gestione
                                        'Dim OggettoDaRicercare_s As String

                                        'Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione



                                'End If

                            '//--------------------------------------------------------------------------------//-//







                For intI = 0 To UBound(intScores())
                    Debug.Print "          "; intScores(intI)
                    Tabella_01_s = vbNull
                    Tabella_01_s = intScores(intI)
                    icount = icount + 1
                    Bool1 = False
                    NuovaDenominazioneTabella_01_s = intScores(intI) & ".xls"


                        '//Controllo Parametri e tabella da esportare
                        Debug.Print
                        Debug.Print "----------------------CONTROLLO VARIABILI TABELLE XLS --------------------------------------"
                        Debug.Print
                        Debug.Print "tabella da ricercare - variabile           :Tabella_01_s                                   ----> " & Tabella_01_s
                        Debug.Print "la path e la variabile                     :par_PATH_S                                     ----> " & PATH_s
                        Debug.Print "nuova denominazione xls con la variabile   :NuovaDenominazioneTabella_01_s                 ----> " & NuovaDenominazioneTabella_01_s
                        Debug.Print
                        Debug.Print "--------------------------------------------------------------------------------------------"


                            '//ESPORTO LA TABELLA
                            '//.......................................................................................
                            '//NOTE : se non trova l'oggetto nel db dove si trova questa classe non esegue il comando. _
                                    la TABELLA viene presa nella matrice.

                                '//Imposto la variabile dell'ogggetto da ricercare e chiamo la funzione _
                                   di ricerca secondo il tipo di oggetto

                                Bool1 = False
                                'Tabella da ricercare = viene indicata nella matrice intScores(intI) senza .xls
                                Bool1 = MyCls_CLASSE_ATTI_Tb01_Gestione.Controlla_SE_OGGETTO_TABELLA_ESISTE_pFunct(Tabella_01_s)

                                    '//SE TRUE = oggetto esistente ed eseguo il comando
                                    If Bool1 Then

                                        '???? NON FUNZIONA
                                        'DoCmd.OutputTo acTable, Tabella_01_s, "MicrosoftExcel(*.xls)", PATH_s, NuovaDenominazioneTabella_01_s, False, ""


                                        '//ESEGUO IL COMANDO OutpuTo CON con 3 parametri, il nome della TABELLA + LA PATH dove esportare + TRUE ?
                                        'MODELLO ---> OutputTo_Sub(par_Tabella_s As String, par_PATH_S As String, par_NuovaDenominazioneTabella_01_s As String, par_Bool As Boolean)
                                        MyCls_CLASSE_ATTI_Tb01_Gestione.OutputTo_Sub Tabella_01_s, PATH_s, NuovaDenominazioneTabella_01_s, False

                                        '//APRO TABELLA ESPORTATA
                                        DoCmd.OpenTable Tabella_01_s, acViewNormal, acEdit



                                  End If

                           '//.......................................................................................



                                MatriceFile_s = MatriceFile_s & icount & ") " & NuovaDenominazioneTabella_01_s & " " & vbCrLf

                    '//Esporto le tabelle in formato xls
                    '                       par_Tabella_s                          par_PATH_S                               par_Bool
                    'sospesa DoCmd.OutputTo acTable, Tabella_01_s, "MicrosoftExcel(*.xls)", PATH_s, NuovaDenominazioneTabella_01_s, False, ""


                Next intI

                '//RILASCIO L'OGGETTO CLASSE
                Set CLASSE_ATTI_Tb01_Gestione = Nothing


                '//Messaggio di fine attivita'
                MsgBox "ESPORTAZIONE DATI IN FILE XLS EFFETTUATI in --> " & PATH_s & vbCrLf & MatriceFile_s & vbCrLf


        ESPORTA_FILE_XLS_ConMatriceDati_pFunct_Exit:
                        Exit Function

        ESPORTA_FILE_XLS_ConMatriceDati_pFunct_Err:
                        MsgBox Error$
                        Resume ESPORTA_FILE_XLS_ConMatriceDati_pFunct_Exit

                End Function
                '//01)_ESPORTA FILE XLS CON MATRICE = Funzione di esportazione dei dati dei lavori pubblici   ***FINE***
                '//=================================================================================================//

                '//ESPORTA_FILE_XLS_ConMatriceDati_pFunct
                '//*************************************************************************************************//
                '//NOTE : ESPORTA i file in formato xls (csv)   *** FINE ***
                '//*************************************************************************************************//




                                                            '//*** FINE ***
                                                    '//*************************************//
                                                    '//     LE ESPORTAZIONI
                                                    '//*************************************//

                '//**************************************************************************************************************************//
                '//         FUNZIONI DI ESPORTAZIONI LAVORI PUBBLICI IN FORMATO XLS
                '//**************************************************************************************************************************//



                                                    '//***INIZIO***
                                                    '//*************************************//
                                                    '//     LE FUNZIONI GENERALI
                                                    '//*************************************//
                '//********************************************************************************************************************************************************//
                '//                                 LE FUNZIONI DI GESTIONE ATTI Tb01_Gestione
                '//********************************************************************************************************************************************************//
                '//
                '//NOTA: In questa classe vengon raggruppate tutte le fuzioni che attengono alla gestione della TABELLA ATTI _
                        per la gestione dei lavori pubblici. _
                        IL seguente elenco specifica le funzioni implementate:


        '//01.AGGIORNA_FILE_ATTO_pfunct = Aggiorna file atto crea il NOME DI FILE da salvare composto da _
                                           fileAtto+NroAtto+Annoatto, solo se nro atto e data sono > 0 _
                                           CODICE------>AGGIORNA_FILE_ATTO_pfunct.attiva

            '//********************************************************************************************************************************************************//
            '//********************************************************************************************************************************************************//



            '//01.AGGIORNA_FILE_ATTO_pfunct = Aggiorna file atto crea il NOME DI FILE da salvare composto da _
                                            fileAtto+NroAtto+Annoatto, solo se nro atto e data sono > 0 _
                                            CODICE------>AGGIORNA_FILE_ATTO_pfunct.attiva
            '//===========================================================:===================================:================================
            '//AGGIORNA_ATTI_Mdl_N001_001_Function.001.01__________:Funzione 00 CALCOLA                                :Calcola
            '//...........................................:.....................................................................................
            '//NOTA PROCEDURA: La Funzione Aggiona il campo FILE_ATTO ricostruendo il nome del file con i parametri passati.

            Public Function AGGIORNA_FILE_ATTO_pfunct(par_TipoAtto_s As String, _
                                                      par_NroAtto_lng As Long, _
                                                      par_DataAtto_d As Date, _
                                                      par_AnnoAtto_i As Integer) As String

                '//.......................................................
                '// DIM VARIABILI
                'Dim sScelta_db As String




                    '//MODELLO-> AGGIORNA_FILE_ATTO
                    '//--------------------------------------------------------------------------//
                    '//NOTA   : Aggiorna il file atto costruendo il nome del file _
                               con i parametri passati ottengo il nome dai seguenti dati _
                               FILE_ATTO=TipoAtto+NroAtto+Annoatto a condizione che _
                               il NroAtto >0 and DataAtto >"", altrimenti la funzione restituisce _
                               il valore nullo. Vengono controllati anche le variabili Nro atto, _
                               data atto ed anno anno che devono essere sempre valorizzate.

                               '//RESETTO IL FILE ATTO _
                               in questo modo posso riattivare l'evento anche nel caso di errori nella _
                               digitazione del numero file atto, potendo ricostruire in automatico il _
                               nome completo del File_atto
                               'Me.FileAtto_s_Txt = ""

                                '//CONTROLLO VARIABILE E CHIAMATA PROCEDURA
                                '//.........................................................//
                                    'sSTR1 = ""
                                    'TipoAtto_s = Me.TipoAtto_s
                                    'If Me.NroAtto_lng > 0 Then
                                    '    NroAtto_lng = Me.NroAtto_lng
                                    'Else
                                    '    Me.NroAtto_lng = 0
                                    'End If

                                    '//CONTROLLO DATA = se la data nulla viene recuperato quella di calendario.
                                    'If IsNull(DataAtto_d) = TRUE Then
                                    '    DataAtto_d = Date
                                    'Else
                                        '//Altrimenti viene recuperata quella della casella di testo valorizzata.
                                    '    DataAtto_d = Me.DataAtto_d
                                    'End If

                                    '//Controllo variabile Anno deve essere valorizzata
                                    'If AnnoAtto_i = 0 Or IsNull(AnnoAtto_i) = True Then
                                        '//Se NULL o Zero, viena applicato l'anno corrente
                                    '    AnnoAtto_i = Year(Date)
                                    'Else
                                        '//Altrimenti viene recuperato l'anno del campo testo
                                    '    AnnoAtto_i = Me.AnnoAtto_i
                                    'End If



                                    '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
                                    'Dim MyCls_CLASSE_ATTI_Tb01_Gestione As CLASSE_ATTI_Tb01_Gestione

                                    '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_FORM assegnando _
                                    l 'oggetto di classe alla variabile e creo creo l'oggetto di classe
                                    'Set MyCls_CLASSE_ATTI_Tb01_Gestione = New CLASSE_ATTI_Tb01_Gestione

                                     '//MODELLO FUNZIONANTE
                                     'sSTR1 = AGGIORNA_FILE_ATTO_pfunct(TipoAtto_s, NroAtto_lng, DataAtto_d, AnnoAtto_i)


                                    'sSTR1 = MyCls_CLASSE_ATTI_Tb01_Gestione.AGGIORNA_FILE_ATTO_pfunct(TipoAtto_s, NroAtto_lng, DataAtto_d, AnnoAtto_i)

                                    '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
                                    'Set MyCls_CLASSE_ATTI_Tb01_Gestione = Nothing


                                    'Me.DATAAGG_TXT = Date
                                    'Me.TIMEOPER_TXT = Time()

                                '//.........................................................//

                    '//MODELLO-> AGGIORNA_FILE_ATTO *** FINE ***
                    '//--------------------------------------------------------------------------//




                        On Error GoTo Err_AGGIORNA_FILE_ATTO_pfunct


                        '//AGGIORNA_FILE_ATTO
                        '//-------------------------------------------:-----------------------------------------------
                        '//AGGIORNA_ATTI_Mdl_N001_001_Function.001.02__________:
                        '//NOTA   : Aggiorna il file atto costruendo il nome del file con i parametri passati.


                                        '//AGGIORNAMENTO CAMPO NRO ATTO -
                                        '//..............................................................
                                        '//Aggiorna il nro att and data > 0 allora _
                                        creo il nome del file con fileAtto+NroAtto+Annoatto, solo se _
                                        nro atto e data sono > 0

                                            If par_NroAtto_lng > 0 And par_DataAtto_d > 0 Then
                                                str1 = ""
                                                str1 = par_TipoAtto_s & "_" & par_NroAtto_lng & "_" & par_AnnoAtto_i
                                                '//controllo e salvataggio
                                                Debug.Print "-------------------------------------------------------"
                                                Debug.Print "file creato"
                                                Debug.Print str1
                                                Debug.Print "-------------------------------------------------------"
                                                AGGIORNA_FILE_ATTO_pfunct = str1


                                            End If

                                        '//..............................................................

                        '//-------------------------------------------:-----------------------------------------------







                '//--------------------------------------------------------------------------------//------------------
                '//                       FINE FUNCTION E GESTIONE ERRORI

            Exit_AGGIORNA_FILE_ATTO_pfunct:
                            Exit Function

            Err_AGGIORNA_FILE_ATTO_pfunct:
                                MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "AGGIORNA_FILE_ATTO_pfunct"
                                Resume Exit_AGGIORNA_FILE_ATTO_pfunct

                    End Function
                    '//AGGIORNA_FILE_ATTO_pfunct                                              *** FINE ***
                    '//===========================================================:===================================:================================



             '//TODO: APRI FINESTRA SHOWOPENFILE DIALOG

             '//01_API_FINESTRA_DI_SELEZIONE_FILE = APRO LA FINESTRA DI SELEZIONE DEI FILE CON PARAMETRI *** INIZIO ***
                '//==================================================================================================================//
                ' ----------------------------------------------------------------------------------
                ' Funzione pubblica che mostra la finestra di selezione file
                ' Parametri:
                '   allowMultiSelect     True = permette selezione multipla
                '   filtroDescrizione    Descrizione del tipo di file (es. "Immagini (*.jpg)")
                '   filtroEstensione     Estensione filtro (es. "*.jpg;*.png")
                '   titoloFinestra       Titolo della finestra (es. "Seleziona file")
                '   cartellaIniziale     Cartella iniziale di apertura
                ' Restituisce:
                '   - Stringa (singolo file)
                '   - Array di stringhe (più file)
                '   - Null se l'utente annulla
                ' ----------------------------------------------------------------------------------

                Public Function ShowOpenFileDialog(Optional ByVal allowMultiSelect As Boolean = False, _
                                                   Optional ByVal filtroDescrizione As String = "Tutti i file (*.*)", _
                                                   Optional ByVal filtroEstensione As String = "*.*", _
                                                   Optional ByVal titoloFinestra As String = "Seleziona file", _
                                                   Optional ByVal cartellaIniziale As String = vbNullString) As Variant
                    On Error GoTo ErrHandler ' Attiva la gestione degli errori

                    ' -- Dichiarazioni variabili
                    Dim OFN As OPENFILENAME           ' Struttura per l'API GetOpenFileName
                    Dim sFileName As String           ' Buffer per contenere i nomi dei file
                    Dim filterString As String        ' Filtro per i tipi di file
                    Dim ret As Long                   ' Risultato della chiamata API
                    Dim fileList() As String          ' Array per la selezione multipla
                    Dim i As Long                     ' Contatore
                    Dim s As String                   ' File singolo selezionato
                    Dim result() As String            ' Array con percorsi completi (se multi)

                    ' -- Crea il filtro nel formato richiesto dall'API (terminato da doppio Null)
                    filterString = filtroDescrizione & vbNullChar & filtroEstensione & vbNullChar

                    ' -- Crea una stringa di 4096 caratteri nulli come buffer per i file
                    sFileName = String$(4096, vbNullChar)

                    ' -- Compila la struttura OPENFILENAME con i parametri forniti
                    With OFN
                        .lStructSize = LenB(OFN)                              ' Dimensione della struttura
                        .hwndOwner = 0                                        ' Nessun proprietario (oppure Application.hWndAccessApp)
                        .lpstrFilter = filterString                           ' Filtro dei tipi di file
                        .nFilterIndex = 1                                     ' Indice filtro predefinito
                        .lpstrFile = sFileName                                ' Buffer per i nomi
                        .nMaxFile = Len(sFileName)                            ' Lunghezza massima del buffer
                        .lpstrFileTitle = String$(260, vbNullChar)            ' Buffer per il nome del file senza percorso
                        .nMaxFileTitle = 260                                  ' Lunghezza del nome file
                        .lpstrInitialDir = cartellaIniziale                   ' Cartella iniziale di apertura
                        .lpstrTitle = titoloFinestra                          ' Titolo della finestra di dialogo
                        .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_EXPLORER ' Opzioni standard
                        If allowMultiSelect Then
                            .flags = .flags Or OFN_ALLOWMULTISELECT           ' Aggiungi selezione multipla se richiesto
                        End If
                    End With

                    ' -- Mostra la finestra di dialogo: ritorna 1 se OK, 0 se Annulla
                    ret = GetOpenFileNameA(OFN)

                    ' -- Se è stato selezionato almeno un file
                    If ret <> 0 Then
                        ' -- Estrai il primo file (o il path base in caso di selezione multipla)
                        s = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, vbNullChar) - 1)

                        If allowMultiSelect Then
                            ' -- Divide i nomi dei file multipli usando il separatore Null
                            fileList = Split(Left$(OFN.lpstrFile, InStr(OFN.lpstrFile & vbNullChar, vbNullChar & vbNullChar) - 1), vbNullChar)

                            If UBound(fileList) = 0 Then
                                ' -- Se è stato selezionato un solo file
                                ShowOpenFileDialog = fileList(0)
                            Else
                                ' -- Se sono stati selezionati più file
                                ReDim result(UBound(fileList) - 1)
                                For i = 1 To UBound(fileList)
                                    ' -- Costruisce il percorso completo per ciascun file
                                    result(i - 1) = fileList(0) & "\" & fileList(i)
                                Next i
                                ShowOpenFileDialog = result ' Ritorna array con tutti i percorsi
                            End If
                        Else
                            ' -- Selezione singola: ritorna direttamente il percorso del file
                            ShowOpenFileDialog = s
                        End If
                    Else
                        ' -- L’utente ha annullato la finestra: ritorna Null
                        ShowOpenFileDialog = Null
                    End If

                    Exit Function ' Esce normalmente dalla funzione

                ' -- Blocco di gestione degli errori
            ErrHandler:
                        ' Log dell'errore tramite sub interna della classe - CHIAMO LA ROUTINE PRIVATA NRO 11
                        ScriviLogErrore "ShowOpenFileDialog", Err.Number, Err.Description, Erl
                        ' (opzionale)
                        MsgBox "Errore durante l'apertura della finestra di dialogo: " & Err.Description, vbCritical, "Errore"
                        ShowOpenFileDialog = Null ' Ritorna Null in caso di errore
                    End Function

                     '//01_API_FINESTRA_DI_SELEZIONE_FILE = APRO LA FINESTRA DI SELEZIONE DEI FILE CON PARAMETRI *** FINE ***
                    '//==================================================================================================================//



            '// TODO: FUNZIONE DA UTILIZZARE CON 4 PARAMETRI


                '//09_FUNZIONE_DA_UTILIZZABILE_API = LA FUNZIONE E UN BASE DI UTILIZZO  *** INIZIO ***
                '//==================================================================================================================//

                Public Function DA_UTILIZZARE_FUNCTION_Pfunct(filePath As String, Optional ByRef txtBox As Control = Nothing) As Boolean
                        On Error GoTo GestioneErrore



                        '// CREA ISTANZA DELLA CLASSE E APRI FILE
                        '//--------------------------------------------------------------------------------------//
                        '//Note: creo l'istanza della classe e chiamo la funzione apri file con le Api _
                                windows utilizzando la CLASSE per la gestonione dei file classe Msys_GEFILE_CLS01_GESTIONI _
                                Come parametro devo passare una path completa di file come ad esempio: _
                                filePath = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_BASE\MDB\MSYS_BASE.mdb" _
                                controllare se tale istanza della funzione è valida per tutti i tipi di file.


                            'Dim clsGestione As New Msys_GEFILE_CLS01_GESTIONI

                            'If clsGestione.DA_UTILIZZARE_FUNCTION_Pfunct(filePath) = False Then
                            '    MsgBox "Apertura file fallita.", vbExclamation
                            'End If



                        '//--------------------------------------------------------------------------------------//



                        #If VBA7 Then
                            Dim result As LongPtr
                        #Else
                            Dim result As Long
                        #End If

                        ' --- Verifica che il file esista prima di tentare di aprirlo ---
                        If Len(Dir(filePath)) = 0 Then
                            MsgBox "Il file selezionato non esiste o non è valido.", vbExclamation
                            DA_UTILIZZARE_FUNCTION_Pfunct = False
                            Exit Function
                        End If

                        ' --- Se è stata passata una TextBox, aggiorna il suo contenuto col percorso file ---
                        If Not txtBox Is Nothing Then
                            On Error Resume Next ' Evita errori se controllo non esiste o non è modificabile
                            If filePath <> "" Then
                                txtBox.SetFocus
                                txtBox.Value = filePath
                            End If
                            On Error GoTo GestioneErrore
                        End If

                        ' --- Chiama ShellExecute per aprire il file con l'applicazione predefinita ---
                        result = ShellExecute(0, "open", filePath, vbNullString, vbNullString, vbNormalFocus)

                        ' --- Controlla risultato ShellExecute, valori <= 32 indicano errore ---
                        If result <= 32 Then
                            MsgBox "Impossibile aprire il file: " & filePath, vbCritical
                            DA_UTILIZZARE_FUNCTION_Pfunct = False
                            Exit Function
                        End If

                        ' --- Se tutto è andato a buon fine ---
                        DA_UTILIZZARE_FUNCTION_Pfunct = True
                        Exit Function

                    ' --- Gestione errori generica ---
            GestioneErrore:
                    ' Registra il dettaglio errore nel log (funzione da implementare)
                    ScriviLogErrore "DA_UTILIZZARE_FUNCTION_Pfunct", Err.Number, Err.Description, Erl
                    MsgBox "Errore durante l'apertura del file: " & Err.Description, vbCritical, "Errore #" & Err.Number
                    DA_UTILIZZARE_FUNCTION_Pfunct = False
                End Function


                '//09_FUNZIONE_DA_UTILIZZABILE_API = LA FUNZIONE E UN BASE DI UTILIZZO  *** FINE ***
                '//==================================================================================================================//



            '//TODO : GESTIONE ERRORI:



                '//MostraProprietaFile  - 10_CLASSE_MOSTRA_PROPRIETA_FILE = LA FUNZIONE MOSTRA LE PROPRIETA DEL FILE PASSATO COME PARAMETRO
                    '//==================================================================================================================//
                    ' Funzione per mostrare le proprietà del file con controllo errori e log
                    ' Mostra le proprietà di un file con gestione robusta degli errori
                    ' e log automatico tramite ScriviLogErrore in caso di eccezioni

                    Public Sub MostraProprietaFile(filePath As String)
                        On Error GoTo ErrHandler ' Gestione errori centralizzata

                        ' Verifica che il percorso non sia vuoto
                        If Trim(filePath) = "" Then
                            MsgBox "Il percorso del file è vuoto.", vbExclamation, "Percorso mancante"
                            Exit Sub
                        End If

                        ' Verifica che il file esista
                        If Dir(filePath) = "" Then
                            MsgBox "Il file non esiste o non è stato trovato.", vbExclamation, "File non trovato"
                            Exit Sub
                        End If

                        ' Crea il FileSystemObject
                        Dim fso As Object
                        Set fso = CreateObject("Scripting.FileSystemObject")

                        Dim FILE As Object
                        Set FILE = fso.GetFile(filePath)

                        ' Mostra le proprietà del file
                        MsgBox "Proprietà del file:" & vbCrLf & _
                               "Percorso: " & FILE.Path & vbCrLf & _
                               "Nome: " & FILE.Name & vbCrLf & _
                               "Tipo: " & FILE.Type & vbCrLf & _
                               "Dimensione: " & Format(FILE.Size, "#,##0") & " bytes" & vbCrLf & _
                               "Data di creazione: " & FILE.DateCreated & vbCrLf & _
                               "Data di modifica: " & FILE.DateLastModified, _
                               vbInformation, "Informazioni sul file"

                        Exit Sub ' Uscita normale

                    '--- Gestione Errori ---
            ErrHandler:
                        ' Chiamata alla funzione di log errori (classe Msys_GEFILE_CLS01_GESTIONI)
                        ScriviLogErrore "MostraProprietaFile", Err.Number, Err.Description, Erl

                        MsgBox "Si è verificato un errore durante il recupero delle proprietà del file:" & vbCrLf & _
                               Err.Description, vbCritical, "Errore"
                    End Sub



                    '//==================================================================================================================//





                '//*** INIZIO ***
                '//11_CLASSE_GESTIONI_GESTIONE_ERRORI = LA GESTIONE DEGLI ERRORI CHE SI TROVA NELLA CLASSE UTILIZZABILE DALLE FUNZIONI
                    '//==================================================================================================================//
                    ' Note - Spiegazione: _
                        Quando succede un errore nella funzione ShowOpenFileDialog o altra, il flusso passa al label ErrHandler. _
                        A questo punto si richiama ScriviLogErrore, passando: _
                        Il nome della funzione dove è avvenuto l’errore (stringa "ShowOpenFileDialog"), _
                        Il numero dell’errore (Err.Number), _
                        La descrizione dell’errore (Err.Description), _
                        Il numero di riga dove si è verificato l’errore (Erl). _
                        Quindi, ogni volta che un errore si verifica in ShowOpenFileDialog, verrà registrato nel file di log tramite ScriviLogErrore.
                    '//miglioramenti: _
                        Formattazione più leggibile della data/ora _
                        Creazione automatica della cartella Logs _
                        Percorso di log leggermente più strutturato (\Logs\LogErrori.txt) _
                        Commenti chiari per manutenzione futura

                    Public Sub ScriviLogErrore(nomeFunzione As String, NumeroErrore As Long, DescrizioneErrore As String, RigaErrore As Long)
                        On Error GoTo GestioneErrore

                        Dim percorsoLog As String
                        Dim fileNum As Integer
                        Dim testo As String
                        Dim dataOra As String
                        Dim cartellaLogs As String

                        ' Format data/ora per log
                        dataOra = Format(Now, "yyyy-mm-dd HH:nn:ss")

                        ' Imposta la cartella Logs nel percorso del database
                        cartellaLogs = CurrentProject.Path & "\Logs"

                        ' Crea la cartella Logs se non esiste
                        If Dir(cartellaLogs, vbDirectory) = "" Then
                            MkDir cartellaLogs
                        End If

                        ' Costruisci il percorso completo del file di log
                        percorsoLog = cartellaLogs & "\LogErrori.txt"

                        ' Componi il testo da scrivere
                        testo = dataOra & " | Classe: Msys_GEFILE_CLS01_GESTIONI | Funzione: " & nomeFunzione & _
                                " | Errore #: " & NumeroErrore & " | Descrizione: " & DescrizioneErrore & _
                                " | Riga: " & RigaErrore

                        ' Scrive nel file in modalità Append
                        fileNum = FreeFile
                        Open percorsoLog For Append As #fileNum
                            Print #fileNum, testo
                        Close #fileNum

                        Exit Sub

            GestioneErrore:
                        ' Gestione errori interni alla funzione di log
                        MsgBox "Errore durante la scrittura del log." & vbCrLf & _
                               "Descrizione: " & Err.Description & vbCrLf & _
                               "Errore #: " & Err.Number, vbCritical, "Errore ScriviLogErrore"

                        ' Scrittura di emergenza in un file alternativo
                        On Error Resume Next
                        percorsoLog = CurrentProject.Path & "\LogErrore_di_Sistema.txt"
                        fileNum = FreeFile
                        Open percorsoLog For Append As #fileNum
                            Print #fileNum, Format(Now, "yyyy-mm-dd HH:nn:ss") & _
                                  " | Scrittura log fallita nella funzione: ScriviLogErrore" & _
                                  " | Errore interno #: " & Err.Number & " - " & Err.Description
                        Close #fileNum
                    End Sub

                    '//==================================================================================================================//

'//TODO: RISISTEMARE PER IL CONTROLL TABLE

'//CONTROLLO OGGETTI TABLE + FIELD TABLE +'
    '//##########################################################################################################################################//
    '//
    '//                     PROCEDURE DI CONTROLLO OGGETTI TABLE NEL DATABASE CORRENTE O ESTERNO
    '//
    '//##########################################################################################################################################//
    '//
    '//NOTE                     : Procedura a funzione che controlla l'esistenza dell'oggetto tabella nel database esterno o interno.
    '//MODULO DI CREAZIONE  --> UTIL_MDL40_N05_CONTROLLO_OGGETTI_DB_TABLES
    '//DATABASE             --> GE_OGGETTI_DB_MDB.mdb
    '//DOVE SI TROVA        --> c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_GESTIONE_OGGETTI_DB_MDB\GESTIONE_OGGETTI\GESTIONE_OGGETTI_MDB\
    '//
    '//LE FUNZIONI E ROUTINE :
    '//                         01)_FUNZIONE_CONTROLLO_OGGETTI_TABELLA = permette il controllo dell'esistenza della tabella oggetto all'interno del db corrente _
                                                                        oppure in un database esterno. _
                                                                        CODICE----------> pfCONTROLLO_OGGETTI_Table.Funzione


    '//                         02)...
    '//
    '//
    '//##########################################################################################################################################//

'//pfCONTROLLO_OGGETTI_Table - CONTROLLO_ESISTENZA_OGGETTI_TABELLA
    '//****************************************************************************************************************************
    '//01)_FUNZIONE_CONTROLLO_OGGETTI_TABELLA = permette il controllo dell'esistenza della tabella oggetto all'interno del db corrente _
                                                            oppure in un database esterno.
    '//CODICE----------> pfCONTROLLO_OGGETTI_Table.Funzione
    '//PARAMETRI: _
            01) sx_Scelta_db    = parametro di SCELTA DATABASE DATABASE da aprire e rappresenta il database esterno o interno; _
                           è esterno a condizione che ci sia una path estesa ed il mome di un db.mdb altrimenti in caso di parametro null _
                           viene inteso come parametro interno. Es.: _
                                    sx_Scelta_db = null il database è quello corrente; _
                                    sx_Scelta_db = C:\MyDb.mdb il database è esterno.
    '//         02) par_sx_Name_Table =?
    '//         03) par_sx_Proprieta_Table =?
    '//         04) par_ix_Operazione =?
    '//         05) par_ix_Indice1 =?
    '//         06) par_bx_Flag_b =?


    Public Function pfCONTROLLO_OGGETTI_Table( _
                                            par_sx_Scelta_db As String, _
                                            par_sx_Name_Table As String, _
                                            par_sx_Proprieta_Table As String, _
                                            par_ix_Operazione As Integer, _
                                            par_ix_Indice1 As Integer, _
                                            par_bx_Flag_b As Boolean) As Boolean



        On Error GoTo Err_pfCONTROLLO_OGGETTI_Table



                    '//CHIAMA IL CONTROLL TABELLE
                    '//------------------------------------------------------------------------------------------//
                    '//CODICE----> pfCONTROLLO_OGGETTI_Table.CALL

                            '//LE VARIABILE PER IL CONTROLLO OGGETTI
                            'Dim sx_Scelta_db As String
                            'Dim sx_Name_Table As String
                            'Dim sx_Proprieta_Table As String
                            'Dim ix_Operazione As Integer
                            'Dim ix_Indice1 As Integer
                            'Dim bx_Flag_b As Boolean

                            '//IMPOSTO LE VARIABILI
                            'sx_Scelta_db = ""
                            'sx_Name_Table = "GEST_MENU_TB01_VOCI"       '//RICERCO QUESTA TABELLA
                            'sx_Proprieta_Table = "Proprieta"
                            'ix_Operazione = 0
                            'ix_Indice1 = 0
                            'bx_Flag_b = False



                              '//ATTIVO LA FUNZIONE CON I PARAMETRI
                              'Bool1 = pfCONTROLLO_OGGETTI_Table(sx_Scelta_db, _
                                                                 sx_Name_Table, _
                                                                 sx_Proprieta_Table, _
                                                                 ix_Operazione, _
                                                                 ix_Indice1, _
                                                                 bx_Flag_b)

                    '//CHIAMA IL CONTROLL TABELLE *** FINE ***
                    '//------------------------------------------------------------------------------------------//




                    '//------------------------------------------------------------------------------------------//
                    '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB


                            'SET VARIABILI:
                            sx_Name_Tab = par_sx_Name_Table

                            ix_Operazione = par_ix_Operazione
                                'ix_Indice1= par_Index
                            bx_Flag_b = par_bx_Flag_b
                            sx_Name_Query = ""

                            'OPEN DB: Apre il database scelto per la creazione o cancellazione a condizione _
                            che esista una path ed il db altrimenti il databse da controllare è quello corrente.
                            If par_sx_Scelta_db > "" Then
                                Set dbs = OpenDatabase(par_sx_Scelta_db)
                            Else
                                Set dbs = CurrentDb
                            End If



                            '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della tabella Nell'insieme Tabledefs
                            '//CONTROLLO PARAMETRI
                            '//-----------------------------------------------------------------------------//

                                Debug.Print "........................................................"
                                Debug.Print "controllo parametri passati alla routine SALVATAGGIO_OGGETTI_DB_PFunct"
                                Debug.Print

                                      Debug.Print "par_sx_Scelta_db ----> " & par_sx_Scelta_db
                                      Debug.Print "par_sx_Name_Table ---> " & par_sx_Name_Table
                                      Debug.Print "par_sx_Name_Query ---> " & par_sx_Proprieta_Table
                                      Debug.Print "par_sx_Name_FORM  ---> " & par_ix_Operazione
                                      Debug.Print "par_sx_Name_Report --> " & par_sx_Name_Report
                                      Debug.Print "par_sx_Name_MACRO ---> " & par_ix_Indice1
                                      Debug.Print "par_sx_Name_Modulo --> " & par_bx_Flag_b
                                      Debug.Print

                                Debug.Print "........................................................"

                            '//-----------------------------------------------------------------------------//


                            '//DBS CICLO
                            '//-----------------------------------------------------------------------------//

                                With dbs

                                    Debug.Print .TableDefs.Count & _
                                        " TableDefs in database  " & sx_Scelta_db
                                    For intCiclo = 0 To .TableDefs.Count - 1

                                        '//BLOCCO ITERAZIONE _
                                        Ricerca di una tabella di sistema o di una tabella _
                                        inesistente con attivazione del blocco. Il nome della tabella _
                                        deve essere attivato manualmente.

                                            If .TableDefs(intCiclo).Name = "MSysIMEXColumns" Then
                                                MsgBox "BLOCCO ITERERAZIONE sulla tabella " & .TableDefs(intCiclo).Name, vbExclamation, "BLOCCO ITERAZIONE"
                                                Stop
                                                End If

                                        Debug.Print
                                        Debug.Print "  " & .TableDefs(intCiclo).Name


                                        '//TABELLA TROVATA
                                        If .TableDefs(intCiclo).Name = par_sx_Name_Table Then

                                             'COMANDO RICERCA TABELLA
                                            DoCmd.SelectObject acTable, par_sx_Name_Table, True

                                            Bln1 = True    ' bx_Flag_b True = trovato tabelle si può cancellare
                                            pfCONTROLLO_OGGETTI_Table = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                            GoTo Exit_pfCONTROLLO_OGGETTI_Table  'esci
                                        Else
                                            '//TABELLA NON TROVATA
                                            Bln1 = False   ' bx_Flag_b False = non trovato TABELLA
                                            pfCONTROLLO_OGGETTI_Table = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                        End If

                                    Next intCiclo

                                End With
                            '//-----------------------------------------------------------------------------//



                '------------------------------------------------------------------------------
                '                       FINE FUNCTIONE E GESTIONE ERRORI

    '//ERRORI ED USCITA
    '//--------------------------------------------------------------------------------------------------------------------------//
        Exit_pfCONTROLLO_OGGETTI_Table:
                    dbs.Close
                    Set dbs = Nothing

                    Exit Function

        Err_pfCONTROLLO_OGGETTI_Table:
                    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfCONTROLLO_OGGETTI_Table"
                    Resume Exit_pfCONTROLLO_OGGETTI_Table


    End Function

    '//*** FINE ***
    '//CONTROLLO_ESISTENZA_OGGETTI_TABELLA
    '//****************************************************************************************************************************


    '//##########################################################################################################################################//
    '//
    '//                     PROCEDURE DI CONTROLLO OGGETTI TABLE NEL DATABASE CORRENTE O ESTERNO *** fine ***
    '//
    '//##########################################################################################################################################//

'//Function CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct - CONTROLLO OGGETTI TABLE
    '//==============================================================================================================================//
    Public Function CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct(par_NameTable_s As String, par_NameField_s As String) As String

        '//MODULO:                  La funzione si trova nel modulo UTIL_MDL40_N05_CONTROLLO_OGGETTI_DB_TABLES del database GE_OGGETTI_DB_MDO.mdb in _
                                    questa path -> c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_GESTIONE_OGGETTI_DB_MDB\GESTIONE_OGGETTI\GESTIONE_OGGETTI_MDB\
        '//CHIAVI DI RICERCA        : Trova campi tabella e tipo di campo, ITERA NELLA TABELLA, Campo Tabella, Field type _
                                      itera campi tabella, trova campo, trova field, itera oggetti field table, _
                                      trova type field, ciclo field table, ciclo campi tabella access, _
                                      esempio itera campi tabella, esempio ciclo campi tabella. _
                                                        ******* _
                                                    IN QUESTA CLASSE ulteriori chiavi di ricerca : _
                                                            FUNZIONE DELLA CLASSE RICERCA TIPO DI CAMPI TABELLA, FUNZIONE TIPO FIELD


        On Error GoTo CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct_Err





        Dim MSG As String ' For error handling.
        Dim ws As Workspace, db As Database
        Dim tblDef As TableDef, fldDef As Field
        Dim i As Integer, handle As Integer
        Dim fldName As String, fldDataInfo As String
        Dim Size_s As String
        Dim TotField_i As Integer

                 ' -----------------------------------------------
                    Set db = CurrentDb()
                    str1 = par_NameTable_s
                 ' -----------------------------------------------


        '//imposto l'oggetto tabella con il nome della tabella passata come parametro
        Set tblDef = db.TableDefs(str1)
                 With tblDef
                        '//reset variabili
                            TotField_i = .Fields.Count - 1              '//calcolo il totale dei field della tabella
                    For i = 0 To .Fields.Count - 1
                       Set fldDef = .Fields(i)

                        '//qualifico l'oggetto
                        With fldDef
                                fldName = .Name

                                '//CONTROLLO FIELD
                                Debug.Print i & ") NOME CAMPO -> " & fldName

                                '//effettuo il controllo tra il campo della tabella e quello della ricerca _
                                  se trovato il campo di ricerca trovo il tipo e restituisco il timpo di _
                                  campo con l'enventuale grandezza
                                If fldName = par_NameField_s Then

                                          Select Case .Type
                                             Case dbBoolean
                                                fldDataInfo = "Bit/Boolean"
                                             Case dbByte
                                                fldDataInfo = "Byte"
                                             Case dbInteger
                                                fldDataInfo = "Short"
                                             Case dbLong
                                                fldDataInfo = "Integer"
                                             Case dbCurrency
                                                fldDataInfo = "Currency"
                                             Case dbSingle
                                                fldDataInfo = "Single"
                                             Case dbDouble
                                                fldDataInfo = "Double"
                                             Case dbDate
                                                fldDataInfo = "Date"
                                             Case dbText
                                                fldDataInfo = "Char"
                                                    '//FUTURI UTILIZZI _
                                                    Al momento non è necessariore restituire la grandezza del campo perchè crea problemi _
                                                    con il filtro impostato nella routine chiamante. Nel futuro puoi costruire una
                                                    Debug.Print "per futuri utilizzi puoi visualizzare"
                                                    Debug.Print "il char e la grandezza del campo"
                                                    Debug.Print "Char Width " & Format(Size > 0)


                                             Case dbLongBinary
                                                fldDataInfo = "OLE"
                                             Case dbMemo
                                                fldDataInfo = "LongChar"
                                             Case dbGUID
                                                fldDataInfo = "Char Width 16"
                                          End Select


                                            '//se trova il campo di ricerca lo restituisco al _
                                            chiamante e esco dal ciclo impostando il contatore dei campi _
                                            FIELD al massimo con l'operazione i = TotField_i in modo da uscire _
                                            in anticipo dal ciclo.
                                                    CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct = fldDataInfo
                                                        i = TotField_i

                                Else
                                        '//se non trova il campo di ricerca restituisco null
                                        CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct = ""

                                End If
                                            '//controllo tipo di campo iterato
                                            Debug.Print "         tipo campo ---> " & fldDataInfo
                                            Debug.Print

                       End With
                    Next i
                 End With

                            '//reset variabili ed oggetti
                            i = vbNull
                            Set tblDef = Nothing


        'EXIT E GESTIONE ERRORI
        '-----------------------------------------------------------------------------------------------

        CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct_Exit:
            Exit Function

        CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct_Err:
            MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfCONTROLLO_OGGETTI_Table"

        Resume CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct_Exit


    End Function

    '//*** fine **
    '//FUNZIONE CONTROLLA CAMPI TABELLA = attivo il controllo sulla tabella e sul campo
    '//==============================================================================================================================//





    '//********************************************************************************************************************************************************//
    '//                                 LE FUNZIONI DI GESTIONE ATTI Tb01_Gestione
    '//********************************************************************************************************************************************************//

                                        '//***FINE***
                                        '//*************************************//
                                        '//     LE FUNZIONI GENERALI
                                        '//*************************************//

'//LoopFormProps_FormCorrente_PULISCI_OGGETTI_TEXT_LABEL_PSub- PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE -
    '//###########################################################################################################################//
    '//NOTE: eseguo il ciclo nell'insieme forms e nell'insieme dei controlli della form filtrata con il nome e condizionata _
        dalla if di filtro. Cicolo all'interno delle proprieta delle form e resetto LE CASELLE DI TESTO E LE LABEL.
    '//CODICE----------->LoopFormProps_FormCorrente_PULISCI_OGGETTI_TEXT_LABEL_PSub
    '//@FUNZ_01=LoopFormProps_FormCorrente_PULISCI_OGGETTI_TEXT_LABEL_PSub


      Public Sub LoopFormProps_FormCorrente_PULISCI_OGGETTI_TEXT_LABEL_PSub(par_NameForm_s As String)
        On Error GoTo Err_Handler
        Dim frmLoop As Object
        Dim ctrl As Control
        Dim propLoop As Property
        Dim strFormName As String

        '//La variabile BOOLEAN per indicare se esiste
        Dim Bool1 As Boolean
        Dim NameForm_s As String
        Dim Msg_s  As String


        '//CHIAMO LA ROUTINE DI PULITURA DEI CAMPI TEXT E LABEL
        '//--------------------------------------------------------------------------------//-//
        '//CODICE----------->LoopFormProps_FormCorrente_PULISCI_OGGETTI_TEXT_LABEL_PSub.puliscoTutto

            '//ATTIVO LA ROUTINE Call LoopFormProps_FormCorrente_PULISCI_OGGETTI_TEXT_LABEL_PSub

        '//--------------------------------------------------------------------------------//-//

        For Each frmLoop In CurrentProject.AllForms

            '//imposto la variabile bool
            Bool1 = False
            '//la form da ricercare nel db
            'NameForm_s = "MOD30_Frm01_S01_ELENCO_SCH01_65txt_campi_1key"

            NameForm_s = par_NameForm_s

            '//TROVA QUELLA CORRENTE
            'If frmLoop.NAME = "MOD30_Frm01_S01_ELENCO_SCH01_65txt_campi_1key" Then
            If frmLoop.Name = par_NameForm_s Then

            '//LA FORM CORRENTE
            '//=====================================================================================================//

                    '//form trovata imposto il nome e  la variabile
                    strFormName = frmLoop.Name
                    '//imposto la variabile bool a true per individuare che è stata trovata
                    Bool1 = True


                    '//APRO LA FORM IN MODALITA DESIGN
                    DoCmd.OpenForm strFormName, acDesign, , , , acHidden

                    '//Qualifico l'oggetto form pe le proprieta
                    With Forms(strFormName)

                        For Each ctrl In .Controls

                            DoEvents

                            For Each propLoop In ctrl.Properties

                                 If ctrl.Name = "TXT_01" Then
                                    'Stop
                                 End If


                            '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                            '//==================================================================================================//
                            '//NOTE : pulisco le caselle di testo impostando la proprieta recordsource a null mentre le label _
                                    corrispondenti vengono impostate al nome oggetto label corrente.
                            '//CODICE----------->LoopFormProps_FormCorrente_PULISCI_OGGETTI_TEXT_LABEL_PSub.puliscoTutto



                                '//PULISCO LA PROPRIETA RECORD SOURCE DELLE CASELLE DI TESTO
                                '//--------------------------------------------------------------------------------//------------//
                                '//NELLA CASELLE TEXT riconosciute con il valore ControlSource _
                                viene azzerata la proprieta control source con valore ""
                                    If propLoop.Name = "ControlSource" Then
                                        Debug.Print propLoop.Value
                                        propLoop.Value = ""
                                    End If

                                '//--------------------------------------------------------------------------------//------------//



                                '//PULISCO LA PROPRIETA CAPTION DELLE LABEL
                                '//--------------------------------------------------------------------------------//------------//
                                '//se è una label riconosciuta con il nome della proprieta _
                                Caption, viene sostituita il valore del contenuto della _
                                etichetta con il nome della label
                                     If propLoop.Name = "Caption" Then
                                                Debug.Print propLoop.Value
                                                '//inserisco il nome label?
                                                propLoop.Value = ctrl.Name


                                    End If

                                '//--------------------------------------------------------------------------------//------------//

                            '//*** FINE ***
                            '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE - CICOLO FOR EACH -
                            '//==================================================================================================//

                                '//CONTROLLO oggetto
                                Debug.Print "Form [" & .Name & "] - Control[" & ctrl.Name & "](" & propLoop.Name & "):" & propLoop.Value

                            Next '//For Each propLoop In ctrl.Properties


                        Next    '//For Each ctrl In .Controls

                                '//SALVATAGGIO FORM APERTA IN DESIGN
                                'Richiama la funzione che chiude la maschera aperta
                                DoCmd.SetWarnings False
                                DoCmd.Close acForm, strFormName, acSaveYes
                                DoCmd.SetWarnings True
                                DoCmd.Close acForm, strFormName

                    End With '//With Forms(strFormName)


            '//=====================================================================================================//



            End If
        Next

                    '//messaggio di avviso se Bool1=false e cioè non esiste la form

                    If Bool1 = False Then
                        '//imposto il testo del messaggio
                        Msg_s = "ATTENZIONE LA FORM " & NameForm_s & " NON E' STATA TROVATA IN QUESTO DATABASE!"
                        MsgBox Msg_s, vbExclamation
                    End If




        Exit Sub
            Err_Handler:
                    Select Case Err.Number
                        Case 2186
                            '//SE LA PROPRIETA NON ESISTE ATTIVO L'ERRORE E RIENTRO NEL CICLO FOR EACH
                            'Property not available in design view  SE NON CI SONO LE  PROPRIETA RIPROVA.
                            Resume Next
                        Case Else
                            MsgBox "Err: " & Err.Number & " - " & Err.Description
                    End Select
                End Sub

        '//*** fine ***
        '//PULISCO I CAMPI TEXT E LABEL DELLA FORM CORRENTE -
        '//###########################################################################################################################//

'//SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO - SALVATAGGIO PARAMETRO DEL CODOGGETTO - SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO
    '//--------------------------------------------------------------------------------//------------------------------------------------//
    'NOTE: funzione Pubblica di salvataggio del codice oggetto nella tabella tmp parametri oggetti TMP
    '//MODULO DI PARTENZA = '3) UTIL_Nro61_N01_SALVATAGGIO_PARAMETRO_CODICE_OGGETTO = SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO
    '//CODICE ----------->SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO.AttivoLaFunzione

    Public Function SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO(par_sxCODOGGETTO) As String



    On Error GoTo Err_SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO

            'Controllo parametro passato
            If par_sxCODOGGETTO = "" Then

                MsgBox ("Funzione : SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO - Parametro Codice Oggetto = null ? - Uscita dalla funzione")

                Exit Function

            End If



    '//MODELLO_ATTIVAZIONE_CLASSE
    '//CHIAMO LA ROUTINE SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTOsxCODOGGETTO DELLA CLASSE OGGETTI FORM
    '//--------------------------------------------------------------------------------//------------------------------------------------//


       '//CODICE----------->SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO.AttivoLaFunzione
       '//--------------------------------------------------------------------------------//-//
       '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_FORM
       '//PARAMETRI            : sono 1 di seguito indicati
               'par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO

       '//NOTA  : AGGIORNO IL CODICE OGGETTO nella tabella oggetti tmp GEST_Tb01_N01_PARAM_OGGETTI_TMP



                   'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_FORM
                   'Assegno e creo l'oggetto di classe
                   'Creo la variabile oggetto                              ---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_FORM As Classe_GESTIONE_OGGETTI_FORM
                   'Creo l'oggetto classe                                  ---> Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = New CLASSE_GESTIONE_OGGETTI_DATABASE


                   'PARAMETRI DI ESTRAZIONE
                   'Attivo il metodo della classe rappresentata dalla routine passando
                   'i1 parametro sxCODOGGETTO

                   'Creo l'oggetto e attivo la funzione e ricevo la stringa ---> Str1 = MyCls_CLASSE_GESTIONE_OGGETTI_FORM.SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO(sxCODOGGETTO)

                   'Libero la memoria                                       ---> Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = Nothing

    '//--------------------------------------------------------------------------------//------------------------------------------------//








            'SALVO IL PARAMETRO NELLA TABELLA OGGETTI TMP
            'solo se esiste un codice oggetto.
            sSql = ""
            sSql = "UPDATE GEST_Tb01_N01_PARAM_OGGETTI_TMP SET GEST_Tb01_N01_PARAM_OGGETTI_TMP.CODOGGETTO = '" & par_sxCODOGGETTO & "';"

            'controllo
            Debug.Print sSql
            CurrentDb.Execute (sSql)


                    SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO = "Salvataggio Parametro Eseguito"




        'USCITA  E GESTIONE ERRORI
        '..............................................................................................................

            Exit_SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO:
                        Exit Function

            Err_SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO:
                        MsgBox Err.Description
                        Resume Exit_SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pbFunctsxCODOGGETTO


        End Function




    '//SALVATAGGIO PARAMETRO DEL CODOGGETTO                                     *** FINE ***
    '//--------------------------------------------------------------------------------//------------------------------------------------//
    '

'//OGGETTI_ITERA_TUTTE_LE_FORM_Sub  - FUNZIONE_ITERAZIONE_TUTTE_LE_FORM
    '//=======================================================================================================//

        Sub OGGETTI_ITERA_TUTTE_LE_FORM_Sub()

        Dim obj As Object
        Dim frm As Form
            Dim ctrl As Control

            For Each obj In CurrentProject.AllForms
                If InStr(obj.Name, "MODELLO_Frm01_M01_GE_FILTRO_M15_H15_L35_NoButton") Then Debug.Print "Forms|" & obj.Name
                'Set frm = Forms(obj.Name) '- this had problems as well, so I commented out
                For Each ctrl In Forms(obj.Name).Controls
                    If InStr(ctrl.Name, "LBL_FILTRO_01") Then Debug.Print "Form|" & obj.Name & "|" & ctrl.Name
                Next ctrl
            Next obj

        End Sub

        '//=======================================================================================================//

'//AllTables - ITERAZIONE NELL'INSIEME TABELLE   @TABELLE.itera.oggetti
    Sub AllTables()
        Dim obj As AccessObject
        Dim dbs As Object
        Set dbs = Application.CurrentData
            icount = 0
        ' Ricerca oggetti AccessObject aperti in insieme AllTables.
        For Each obj In dbs.AllTables


             DoEvents
                icount = icount + 1
                Debug.Print "NRO    -> " & icount
                Debug.Print "Tabelle -> " & obj.Name


            If obj.IsLoaded = True Then
                ' Stampa nome di obj.
                Debug.Print obj.Name
            End If


        Next obj


    End Sub

'//AllModules ITERAZIONE NELL'INSIEM MODULI   @moduli.itera.oggetti
    Sub AllModules()
        Set DbsObject = Application.CurrentProject
        icount = 0
        ' Ricerca oggetti AccessObject aperti in insieme AllModules.
        For Each obj In DbsObject.AllModules

            DoEvents
            icount = icount + 1
            Debug.Print "NRO    -> " & icount
            Debug.Print "modulo -> " & obj.Name

            If obj.IsLoaded = True Then
                ' Stampa nome di obj.
                Debug.Print "modulo APERTO -> " & obj.Name
            End If
        Next obj
        End Sub

'//pfCONTROLLO_ESISTENZA_OGGETTI_QUERY - CONTROLLA SE ESISTE L'OGGETTO QUERY NEL DB  - @itera.oggetti.QUERY, @itera.QUERYs, @itera.tutte.le.QUERY _

    '//ATTIVA_ESISTE_QUERY - CONTROLLO_ESISTENZA_OGGETTI_QUERY
        '//****************************************************************************************************************************

        Private Sub ATTIVA_ESISTE_QUERY()


            '//RICERCA ESISTENZA OGETTO QUERY  @esiste.QUERY,  @esiste.la.QUERY
            '//------------------------------------------------------------------------//
            '//NOTA: per vedere se esiste la QUERY invio come parametro il nome dell'oggetto _
                    QUERY DA RICERCARE e chiamo la funzione la quale restituisce un TRUE _
                    se esiste altrimenti un FALSE = oggetto non esiste nel database
                str1 = "MSys_ELENCO_ObjProgetti_OLD_01"
                Bool1 = pfCONTROLLO_ESISTENZA_OGGETTI_QUERY(str1)

            '//------------------------------------------------------------------------//


            If Bool1 = True Then

                MsgBox "Oggetto QUERY TROVATO : " & str1, vbInQUERYation, " ATTIVA FUNZIONE ESISTE QUERY"
                 'COMANDO RICERCA QUERY
                DoCmd.SelectObject acQuery, str1, True


            Else
                MsgBox "ATTENZIONE - CONTROLLO DELLA CLASSE DENOMINATA GESTIONE OGGETTI DATABASE:  --->   UN QUESTA QUERY NON E' STATA TROVATA : " & str1, vbInQUERYation, " ATTIVA FUNZIONE ESISTE QUERY"
            End If

        End Sub




            '//@CONTROLLO.OGGETTO.QUERY, @CONTROLLO.OGGETTO.QUERY, @QUERY.CONTROLLO.ESISTENZA _
                @QUERYS.CONTROLLO.ESISTENZA
            Public Function pfCONTROLLO_ESISTENZA_OGGETTI_QUERY(par_sx_Name_OGGETTO As String) As Boolean

                On Error GoTo Err_pfCONTROLLO_ESISTENZA_OGGETTI_QUERY





                            '//ITERAZIONE NEGLI OGGETTI QUERY
                            '//==============================================================================================================//
                            '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB





                                    '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della QUERY Nell'insieme QUERYdefs
                                    '//CONTROLLO PARAMETRI
                                    '//-----------------------------------------------------------------------------//
                                        '//RESET

                                         pfCONTROLLO_ESISTENZA_OGGETTI_QUERY = False

                                        Debug.Print "........................................................"
                                        Debug.Print "controllo parametri passati alla routine SALVATAGGIO_OGGETTI_DB_PFunct"
                                        Debug.Print

                                              Debug.Print "par_sx_Scelta_db ----> " & par_sx_Scelta_db
                                               Debug.Print

                                        Debug.Print "........................................................"

                                    '//-----------------------------------------------------------------------------//


                                            Set DbsObject = Application.CurrentProject
                                            icount = 0

                                            ' Ricerca oggetti AccessObject aperti in insieme AllModules.
                                            For Each obj In DbsObject.AllQueries

                                                DoEvents
                                                icount = icount + 1
                                                Debug.Print "NRO             -> " & icount
                                                Debug.Print "OGGETTO QUERY   -> " & obj.Name

                                                '//CONTROLLO SE OGGETTO TROVATO
                                                If obj.Name = par_sx_Name_OGGETTO Then

                                                        pfCONTROLLO_ESISTENZA_OGGETTI_QUERY = True

                                                End If

                                                '@QUERY.APERTA.? @?.QUERY.APERTA
                                                If obj.IsLoaded = True Then
                                                ' Stampa nome di obj.
                                                Debug.Print "QUERY.APERTA -> " & obj.Name
                                                End If
                                            Next obj


                                        '*******************************************







                                    '//-----------------------------------------------------------------------------//

                        '//ITERAZIONE NEGLI OGGETTI QUERY  *** FINE ****
                        '//==============================================================================================================//



            '//ERRORI ED USCITA FINE FUNCTIONE E GESTIONE ERRORI
            '//--------------------------------------------------------------------------------//------------------------------------------//
                        Exit_pfCONTROLLO_ESISTENZA_OGGETTI_QUERY:

                                Exit Function

                    Err_pfCONTROLLO_ESISTENZA_OGGETTI_QUERY:
                                MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfCONTROLLO_ESISTENZA_OGGETTI_QUERY"
                                Resume Exit_pfCONTROLLO_ESISTENZA_OGGETTI_QUERY


            End Function

        '//****************************************************************************************************************************

'//pfCONTROLLO_ESISTENZA_OGGETTI_TABLE - CONTROLLO_ESISTENZA_OGGETTI_TABLE
    '//****************************************************************************************************************************

    Private Sub ATTIVA_ESISTE_TABLE()


        '//RICERCA ESISTENZA OGETTO TABLE  @esiste.TABLE,  @esiste.la.TABLE
        '//------------------------------------------------------------------------//
        '//NOTA: per vedere se esiste la TABLE invio come parametro il nome dell'oggetto _
                TABLE DA RICERCARE e chiamo la funzione la quale restituisce un TRUE _
                se esiste altrimenti un FALSE = oggetto non esiste nel database
            str1 = "MSys_DF01_COMANDI"
            Bool1 = pfCONTROLLO_ESISTENZA_OGGETTI_TABLE(str1)

        '//------------------------------------------------------------------------//


        If Bool1 = True Then

            MsgBox "Oggetto TABLE TROVATO : " & str1, vbInTABLEation, " ATTIVA FUNZIONE ESISTE TABLE"
             'COMANDO RICERCA tabella
            DoCmd.SelectObject acTable, str1, True


        Else
            MsgBox "ATTENZIONE - CONTROLLO DELLA CLASSE DENOMINATA GESTIONE OGGETTI DATABASE:  --->   UN QUESTA TABLE NON E' STATA TROVATA : " & str1, vbInTABLEation, " ATTIVA FUNZIONE ESISTE TABLE"
        End If

    End Sub

'//pfCONTROLLO_ESISTENZA_OGGETTI_TABLE - CONTROLLA SE ESISTE L'OGGETTO TABLE NEL DB  - @itera.oggetti.TABLE, @itera.TABLEs, @itera.tutte.le.TABLE _
            @CONTROLLO.OGGETTO.TABLE, @CONTROLLO.OGGETTO.TABELLA, @TABELLE.CONTROLLO.ESISTENZA _
            @TABLES.CONTROLLO.ESISTENZA     @codice----> @RICERCA@ESISTENZA@TABELLE_(controllo se la tabella esiste nel db corrente)
            '//2024.08.07_MIGLIORATA_RICERCA_TABELLE

    Public Function pfCONTROLLO_ESISTENZA_OGGETTI_TABLE(par_sx_Name_OGGETTO As String) As Boolean
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim isTableOpen As Boolean

        On Error GoTo Err_pfCONTROLLO_ESISTENZA_OGGETTI_TABLE

        ' Inizializzazione
        pfCONTROLLO_ESISTENZA_OGGETTI_TABLE = False
        isTableOpen = False

        ' Ottieni il riferimento al database corrente
        Set db = CurrentDb

        ' Verifica se la tabella esiste come tabella fisica o collegata
        ' Modifica della Query SQL: La query ora include Type In (1, 6), dove: _
                1 rappresenta le tabelle fisiche. _
                6 rappresenta le tabelle collegate.
        Set rs = db.OpenRecordset("SELECT * FROM MSysObjects WHERE Type In (1, 6) AND Name='" & par_sx_Name_OGGETTO & "'")


        ' Verifica se il recordset è popolato
        If Not rs.EOF Then
            ' Tabella trovata
            pfCONTROLLO_ESISTENZA_OGGETTI_TABLE = True

            ' Controlla se la tabella è aperta _
            Controllo degli Oggtti: La verifica per SysCmd(acSysCmdGetObjectState, acTable, par_sx_Name_OGGETTO) rimane invariata perché è ancora valida per entrambi i tipi di tabelle _
            collegate e fisiche.
            If SysCmd(acSysCmdGetObjectState, acTable, par_sx_Name_OGGETTO) <> 0 Then
                ' La tabella è aperta
                isTableOpen = True
            End If

            ' Se la tabella è aperta, esci con messaggio di avviso E RESTITUISCE FALSE
            If isTableOpen Then
                MsgBox "ATTENZIONE - CONTROLLO DELLA CLASSE DENOMINATA GESTIONE OGGETTI DATABASE:  --->   TABELLA APERTA - uscita dalla funzione", vbExclamation, "pfCONTROLLO_ESISTENZA_OGGETTI_TABLE"
                pfCONTROLLO_ESISTENZA_OGGETTI_TABLE = False
                Exit Function
            End If
        End If

        ' Rilascia il riferimento al recordset e al database
        rs.Close
        Set rs = Nothing
        Set db = Nothing

                Exit_pfCONTROLLO_ESISTENZA_OGGETTI_TABLE:
                                Exit Function

                Err_pfCONTROLLO_ESISTENZA_OGGETTI_TABLE:
                                MsgBox "ERRORE FUNCTION PUBLIC " & Err.Number & " - " & Err.Description, vbCritical, "pfCONTROLLO_ESISTENZA_OGGETTI_TABLE"
                                Resume Exit_pfCONTROLLO_ESISTENZA_OGGETTI_TABLE
                            End Function



    '//*** FINE ***
    '//CONTROLLO_ESISTENZA_OGGETTI_TABLE
    '//****************************************************************************************************************************

'//pfCONTROLLO_ESISTENZA_OGGETTI_Form -CONTROLLO_ESISTENZA_OGGETTI_FORM
    '//****************************************************************************************************************************

    Private Sub ATTIVA_ESISTE_FORM()


        '//RICERCA ESISTENZA OGETTO FORM  @esiste.form,  @esiste.la.form
        '//------------------------------------------------------------------------//
        '//NOTA: per vedere se esiste la form invio come parametro il nome dell'oggetto _
                FORM DA RICERCARE e chiamo la funzione la quale restituisce un TRUE _
                se esiste altrimenti un FALSE = oggetto non esiste nel database
            str1 = "GEST_MENU_Tb02Frm01_M01_GE_FUNZIONI"
            Bool1 = pfCONTROLLO_ESISTENZA_OGGETTI_Form(str1)

        '//------------------------------------------------------------------------//


        If Bool1 = True Then

            MsgBox "Oggetto form TROVATO : " & str1, vbInformation, " ATTIVA FUNZIONE ESISTE FORM"
             'COMANDO RICERCA QUERY
            DoCmd.SelectObject acForm, str1, True


        Else
            MsgBox "ATTENZIONE - CONTROLLO DELLA CLASSE DENOMINATA GESTIONE OGGETTI DATABASE:  --->   UN QUESTA FORM NON E' STATA TROVATA : " & str1, vbInformation, " ATTIVA FUNZIONE ESISTE FORM"
        End If

    End Sub

'//pfCONTROLLO_ESISTENZA_OGGETTI_Form -@CONTROLLO@ESISTENZA@FORM_(Funzione che rileva se esiste nel db corrente la form restituendo True o False)
        '//CONTROLLA SE ESISTE L'OGGETTO FORM NEL DB  - @itera.oggetti.form, @itera.forms, @itera.tutte.le.form
        Public Function pfCONTROLLO_ESISTENZA_OGGETTI_Form(par_sx_Name_OGGETTO As String) As Boolean

            On Error GoTo Err_pfCONTROLLO_ESISTENZA_OGGETTI_Form





                    '//ITERAZIONE NEGLI OGGETTI FORMS
                    '//==============================================================================================================//
                    '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB





                            '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORM Nell'insieme Tabledefs
                            '//CONTROLLO PARAMETRI
                            '//-----------------------------------------------------------------------------//
                                '//RESET

                                 pfCONTROLLO_ESISTENZA_OGGETTI_Form = False

                                Debug.Print "........................................................"
                                Debug.Print "controllo parametri passati alla routine SALVATAGGIO_OGGETTI_DB_PFunct"
                                Debug.Print

                                      Debug.Print "par_sx_Scelta_db ----> " & par_sx_Scelta_db
                                       Debug.Print

                                Debug.Print "........................................................"

                            '//-----------------------------------------------------------------------------//


                                    Set DbsObject = Application.CurrentProject
                                    icount = 0
                                    ' Ricerca oggetti AccessObject aperti in insieme AllModules.
                                    For Each obj In DbsObject.AllForms

                                        DoEvents
                                        icount = icount + 1
                                        Debug.Print "NRO            -> " & icount
                                        Debug.Print "OGGETTO FORM   -> " & obj.Name

                                        '//CONTROLLO SE OGGETTO TROVATO
                                        If obj.Name = par_sx_Name_OGGETTO Then

                                                pfCONTROLLO_ESISTENZA_OGGETTI_Form = True

                                        End If


                                        If obj.IsLoaded = True Then
                                        ' Stampa nome di obj.
                                        Debug.Print "FORM APERTA APERTO -> " & obj.Name
                                        End If
                                    Next obj


                                '*******************************************







                            '//-----------------------------------------------------------------------------//

                '//ITERAZIONE NEGLI OGGETTI FORMS  *** FINE ****
                '//==============================================================================================================//



        '//ERRORI ED USCITA FINE FUNCTIONE E GESTIONE ERRORI
        '//--------------------------------------------------------------------------------//------------------------------------------//
                Exit_pfCONTROLLO_ESISTENZA_OGGETTI_Form:

                                            Exit Function

                Err_pfCONTROLLO_ESISTENZA_OGGETTI_Form:
                                            MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfCONTROLLO_ESISTENZA_OGGETTI_Form"
                                            Resume Exit_pfCONTROLLO_ESISTENZA_OGGETTI_Form


        End Function

    '//*** FINE ***
    '//CONTROLLO_ESISTENZA_OGGETTI_FORM
    '//****************************************************************************************************************************

'//pfCONTROLLO_ESISTENZA_OGGETTI_REPORTS - CONTROLLO_ESISTENZA_OGGETTI_REPORTS
    '//****************************************************************************************************************************
    '//@CONTROLLO@ESISTENZA@REPORT_(Funzione che rileva se esiste nel db corrente la REPORT restituendo True o False)
    Private Sub ATTIVA_ESISTE_REPORTS()


        '//RICERCA ESISTENZA OGETTO REPORTS  @esiste.REPORTS,  @esiste.la.REPORTS
        '//------------------------------------------------------------------------//
        '//NOTA: per vedere se esiste la REPORTS invio come parametro il nome dell'oggetto _
                REPORTS DA RICERCARE e chiamo la funzione la quale restituisce un TRUE _
                se esiste altrimenti un FALSE = oggetto non esiste nel database
            str1 = "REPORT_n01_OGGETTI_ITERA_TUTTE_LE_FOR_SubM"
            Bool1 = pfCONTROLLO_ESISTENZA_OGGETTI_REPORTS(str1)

        '//------------------------------------------------------------------------//


        If Bool1 = True Then

            MsgBox "Oggetto REPORTS TROVATO : " & str1, vbInREPORTSation, " ATTIVA FUNZIONE ESISTE REPORTS"
             'COMANDO RICERCA REPORTS
            DoCmd.SelectObject acReport, str1, True


        Else
            MsgBox "ATTENZIONE - CONTROLLO DELLA CLASSE DENOMINATA GESTIONE OGGETTI DATABASE:  --->   UN QUESTA REPORTS NON E' STATA TROVATA : " & str1, vbInREPORTSation, " ATTIVA FUNZIONE ESISTE REPORTS"
        End If

    End Sub




 '//pfCONTROLLO_ESISTENZA_OGGETTI_MODULI - CONTROLLA SE ESISTE L'OGGETTO REPORTS NEL DB  - @itera.oggetti.REPORTS, @itera.REPORTSs, @itera.tutte.le.REPORTS
    Public Function pfCONTROLLO_ESISTENZA_OGGETTI_REPORTS(par_sx_Name_OGGETTO As String) As Boolean

        On Error GoTo Err_pfCONTROLLO_ESISTENZA_OGGETTI_REPORTS





                    '//ITERAZIONE NEGLI OGGETTI REPORTSS
                    '//==============================================================================================================//
                    '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB



                            'RESET

                             pfCONTROLLO_ESISTENZA_OGGETTI_REPORTS = False


                            '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della REPORTS Nell'insieme Tabledefs
                            '//CONTROLLO PARAMETRI
                            '//-----------------------------------------------------------------------------//

                                Debug.Print "........................................................"
                                Debug.Print "controllo parametri passati alla routine SALVATAGGIO_OGGETTI_DB_PFunct"
                                Debug.Print

                                      Debug.Print "par_sx_Scelta_db ----> " & par_sx_Scelta_db
                                       Debug.Print

                                Debug.Print "........................................................"

                            '//-----------------------------------------------------------------------------//


                                    Set DbsObject = Application.CurrentProject
                                    icount = 0
                                    ' Ricerca oggetti AccessObject aperti in insieme AllModules.
                                    For Each obj In DbsObject.AllReports
                                        DoEvents
                                        icount = icount + 1
                                        Debug.Print "NRO            -> " & icount
                                        Debug.Print "OGGETTO REPORTS   -> " & obj.Name

                                        '//CONTROLLO SE OGGETTO TROVATO
                                        If obj.Name = par_sx_Name_OGGETTO Then

                                                pfCONTROLLO_ESISTENZA_OGGETTI_REPORTS = True

                                        End If


                                        If obj.IsLoaded = True Then
                                        ' Stampa nome di obj.
                                        Debug.Print "REPORTS APERTO -> " & obj.Name
                                        End If
                                    Next obj


                                '*******************************************







                            '//-----------------------------------------------------------------------------//

                '//ITERAZIONE NEGLI OGGETTI REPORTSS  *** FINE ****
                '//==============================================================================================================//



    '//ERRORI ED USCITA FINE FUNCTIONE E GESTIONE ERRORI
    '//--------------------------------------------------------------------------------//------------------------------------------//
            Exit_pfCONTROLLO_ESISTENZA_OGGETTI_REPORTS:

                                    Exit Function

            Err_pfCONTROLLO_ESISTENZA_OGGETTI_REPORTS:
                                    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfCONTROLLO_ESISTENZA_OGGETTI_REPORTS"
                                    Resume Exit_pfCONTROLLO_ESISTENZA_OGGETTI_REPORTS


            End Function

    '//*** FINE ***
    '//CONTROLLO_ESISTENZA_OGGETTI_REPORTS
    '//****************************************************************************************************************************

'//pfCONTROLLO_ESISTENZA_OGGETTI_MODULI - CONTROLLO_ESISTENZA_OGGETTI_MODULI
    '//****************************************************************************************************************************

    Private Sub ATTIVA_ESISTE_MODULI()


        '//RICERCA ESISTENZA OGETTO MODULI  @esiste.MODULI,  @esiste.la.MODULI
        '//------------------------------------------------------------------------//
        '//NOTA: per vedere se esiste la MODULI invio come parametro il nome dell'oggetto _
                MODULI DA RICERCARE e chiamo la funzione la quale restituisce un TRUE _
                se esiste altrimenti un FALSE = oggetto non esiste nel database
            str1 = "UTIL_MDL40_N09_CONTROLLO_ESISTENZA_OGGETTO_DB"
            Bool1 = pfCONTROLLO_ESISTENZA_OGGETTI_MODULI(str1)

        '//------------------------------------------------------------------------//


        If Bool1 = True Then

            MsgBox "Oggetto MODULI TROVATO : " & str1, vbInMODULIation, " ATTIVA FUNZIONE ESISTE MODULI"
             'COMANDO RICERCA MODULI
            DoCmd.SelectObject acModule, str1, True


        Else
            MsgBox "ATTENZIONE - CONTROLLO DELLA CLASSE DENOMINATA GESTIONE OGGETTI DATABASE:  --->   UN QUESTA MODULI NON E' STATA TROVATA : " & str1, vbInMODULIation, " ATTIVA FUNZIONE ESISTE MODULI"
        End If

    End Sub

'//pfCONTROLLO_ESISTENZA_OGGETTI_MODULI - CONTROLLA SE ESISTE L'OGGETTO MODULI NEL DB  - @itera.oggetti.MODULI, @itera.MODULIs, @itera.tutte.le.MODULI
    Public Function pfCONTROLLO_ESISTENZA_OGGETTI_MODULI(par_sx_Name_OGGETTO As String) As Boolean

        On Error GoTo Err_pfCONTROLLO_ESISTENZA_OGGETTI_MODULI





                    '//ITERAZIONE NEGLI OGGETTI MODULIS
                    '//==============================================================================================================//
                    '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB





                            '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MODULI Nell'insieme Tabledefs
                            '//CONTROLLO PARAMETRI
                            '//-----------------------------------------------------------------------------//

                                'RESET

                                pfCONTROLLO_ESISTENZA_OGGETTI_MODULI = False


                                Debug.Print "........................................................"
                                Debug.Print "controllo parametri passati alla routine SALVATAGGIO_OGGETTI_DB_PFunct"
                                Debug.Print

                                      Debug.Print "par_sx_Scelta_db ----> " & par_sx_Scelta_db
                                       Debug.Print

                                Debug.Print "........................................................"

                            '//-----------------------------------------------------------------------------//


                                    Set DbsObject = Application.CurrentProject
                                    icount = 0
                                    ' Ricerca oggetti AccessObject aperti in insieme AllModules.
                                    For Each obj In DbsObject.AllModules
                                        DoEvents
                                        icount = icount + 1
                                        Debug.Print "NRO            -> " & icount
                                        Debug.Print "OGGETTO MODULI   -> " & obj.Name

                                        '//CONTROLLO SE OGGETTO TROVATO
                                        If obj.Name = par_sx_Name_OGGETTO Then

                                                pfCONTROLLO_ESISTENZA_OGGETTI_MODULI = True


                                        End If


                                        If obj.IsLoaded = True Then
                                        ' Stampa nome di obj.
                                        Debug.Print "MODULI APERTA APERTO -> " & obj.Name
                                        End If
                                    Next obj


                                '*******************************************







                            '//-----------------------------------------------------------------------------//

                '//ITERAZIONE NEGLI OGGETTI MODULIS  *** FINE ****
                '//==============================================================================================================//



    '//ERRORI ED USCITA FINE FUNCTIONE E GESTIONE ERRORI
    '//--------------------------------------------------------------------------------//------------------------------------------//
            Exit_pfCONTROLLO_ESISTENZA_OGGETTI_MODULI:

                        Exit Function

            Err_pfCONTROLLO_ESISTENZA_OGGETTI_MODULI:
                                    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfCONTROLLO_ESISTENZA_OGGETTI_MODULI"
                                    Resume Exit_pfCONTROLLO_ESISTENZA_OGGETTI_MODULI


            End Function

    '//*** FINE ***
    '//CONTROLLO_ESISTENZA_OGGETTI_MODULI
    '//****************************************************************************************************************************

'//pfCONTROLLO_ESISTENZA_OGGETTI_MACRO -CONTROLLO_ESISTENZA_OGGETTI_MACRO
 '//****************************************************************************************************************************

    Private Sub ATTIVA_ESISTE_MACRO()


        '//RICERCA ESISTENZA OGETTO MACRO  @esiste.MACRO,  @esiste.la.MACRO
        '//------------------------------------------------------------------------//
        '//NOTA: per vedere se esiste la MACRO invio come parametro il nome dell'oggetto _
                MACRO DA RICERCARE e chiamo la funzione la quale restituisce un TRUE _
                se esiste altrimenti un FALSE = oggetto non esiste nel database
            str1 = "REPORT_n01_OGGETTI_ITERA_TUTTE_LE_FOR_SubM"
            Bool1 = pfCONTROLLO_ESISTENZA_OGGETTI_MACRO(str1)

        '//------------------------------------------------------------------------//


        If Bool1 = True Then

            MsgBox "Oggetto MACRO TROVATO : " & str1, vbInMACROation, " ATTIVA FUNZIONE ESISTE MACRO"
             'COMANDO RICERCA MACRO
            DoCmd.SelectObject acReport, str1, True


        Else
            MsgBox "ATTENZIONE - CONTROLLO DELLA CLASSE DENOMINATA GESTIONE OGGETTI DATABASE:  --->   UN QUESTA MACRO NON E' STATA TROVATA : " & str1, vbInMACROation, " ATTIVA FUNZIONE ESISTE MACRO"
        End If

    End Sub

    '//CONTROLLA SE ESISTE L'OGGETTO MACRO NEL DB  - @itera.oggetti.MACRO, @itera.MACRO, @itera.tutte.le.MACRO
    '//@CONTROLLO@ESISTENZA@MACRO_(Funzione che rileva se esiste nel db corrente dei MACRO restituendo True o False)
    Public Function pfCONTROLLO_ESISTENZA_OGGETTI_MACRO(par_sx_Name_OGGETTO As String) As Boolean

        On Error GoTo Err_pfCONTROLLO_ESISTENZA_OGGETTI_MACRO





                    '//ITERAZIONE NEGLI OGGETTI MACRO
                    '//==============================================================================================================//
                    '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB



                            'RESET

                             pfCONTROLLO_ESISTENZA_OGGETTI_MACRO = False


                            '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO Nell'insieme Tabledefs
                            '//CONTROLLO PARAMETRI
                            '//-----------------------------------------------------------------------------//

                                Debug.Print "........................................................"
                                Debug.Print "controllo parametri passati alla routine SALVATAGGIO_OGGETTI_DB_PFunct"
                                Debug.Print

                                      Debug.Print "par_sx_Scelta_db ----> " & par_sx_Scelta_db
                                       Debug.Print

                                Debug.Print "........................................................"

                            '//-----------------------------------------------------------------------------//


                                    Set DbsObject = Application.CurrentProject
                                    icount = 0
                                    ' Ricerca oggetti AccessObject aperti in insieme AllModules.
                                    For Each obj In DbsObject.AllMacros
                                        DoEvents
                                        icount = icount + 1
                                        Debug.Print "NRO            -> " & icount
                                        Debug.Print "OGGETTO MACRO   -> " & obj.Name

                                        '//CONTROLLO SE OGGETTO TROVATO
                                        If obj.Name = par_sx_Name_OGGETTO Then

                                                pfCONTROLLO_ESISTENZA_OGGETTI_MACRO = True

                                        End If


                                        If obj.IsLoaded = True Then
                                        ' Stampa nome di obj.
                                        Debug.Print "MACRO APERTA -> " & obj.Name
                                        End If
                                    Next obj


                                '*******************************************







                            '//-----------------------------------------------------------------------------//

                '//ITERAZIONE NEGLI OGGETTI MACRO  *** FINE ****
                '//==============================================================================================================//



    '//ERRORI ED USCITA FINE FUNCTIONE E GESTIONE ERRORI
    '//--------------------------------------------------------------------------------//------------------------------------------//
                Exit_pfCONTROLLO_ESISTENZA_OGGETTI_MACRO:

                        Exit Function

                Err_pfCONTROLLO_ESISTENZA_OGGETTI_MACRO:
                        MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfCONTROLLO_ESISTENZA_OGGETTI_MACRO"
                        Resume Exit_pfCONTROLLO_ESISTENZA_OGGETTI_MACRO


        End Function

    '//*** FINE ***
    '//CONTROLLO_ESISTENZA_OGGETTI_MACRO
    '//****************************************************************************************************************************

'//FORM_STATO_OGGETTO_funct - FORM REPORT ECC'
    '//****************************************************************************************************************************
    '//
    '//             CONTROLLO_STATO_OGGETTI: FORM, REPORT , TABELLE QUERY ECC. con SysCmd  **** INIZIO ****
    '//
    '//****************************************************************************************************************************


    '//todo: SI DEVONO CREARE ALTRE ROUTINE per vedere lo stato delle Tabelle, query report ecc. per ora sola form.
    '//@stato@oggetto_(comando con SysCmd che controlla l'ambiente windows per determinare lo stato dell'oggetto access)
    Private Sub FORM_STATO_OGGETTO_funct()


    '// In Microsoft Access, la variabile acSysCmdGetObjectState è utilizzata in combinazione con la funzione SysCmd per _
        ottenere lo stato di un oggetto di Access, come una tabella, una query, un modulo, un report, ecc.

    '//La funzione SysCmd permette di eseguire diverse operazioni di sistema e di interagire con l'ambiente di Access. _
        Quando si usa acSysCmdGetObjectState, puoi determinare se un oggetto specifico è aperto o chiuso. _
        Questo è particolarmente utile se vuoi verificare se un oggetto è attualmente in uso prima di eseguire un'altra operazione su di esso.

    '//sintassi SysCmd(acSysCmdGetObjectState, ObjectType, ObjectName) _
        acSysCmdGetObjectState: Specifica che vuoi ottenere lo stato dell'oggetto. _
        ObjectType: Il tipo di oggetto di Access (ad esempio, acTable, acQuery, acForm, acReport). _
        ObjectName: Il nome dell'oggetto di Access di cui vuoi controllare lo stato.

    '// Valori di ritorno _
        La funzione SysCmd restituirà uno dei seguenti valori: _
        0: L 'oggetto non è aperto. _
        1: L 'oggetto è aperto. _

    '//Esempio di utilizzo _
    Supponiamo di voler controllare se un modulo chiamato "GEST_MENU_Tb01Frm01_M01_GE_VOCI" è aperto. Puoi utilizzare il seguente codice VBA:

    '//FAQ: come controllare se un oggetto è aperto. _
            controllo se la form è aperta oppure no.

            Dim objState As Integer
            Dim myNameForm As String

            '//@Utilizzo@SysCmd_(vba variabile di sistema SysCmd per controllare l'oggetto se è aperto oppure no)
            myNameForm = "GEST_MENU_Tb01Frm01_M01_GE_VOCI"
            objState = SysCmd(acSysCmdGetObjectState, acForm, myNameForm)

            If objState = 1 Then
                MsgBox "la form '" & myNameForm & "' è APERTA."
            Else
                MsgBox "la form '" & MyForm & "' non è aperta."
            End If


    End Sub




        ''// ..... MANCANO GLI ALTRI OGGETTI DA CONTROLLARE : TABELLE, QUERY ECC..

        '//****************************************************************************************************************************
        '//
        '//             CONTROLLO_STATO_OGGETTI: FORM, REPORT , TABELLE QUERY ECC. con SysCmd  **** FINE ****
        '//
        '//****************************************************************************************************************************

'//Function Iterazione_oggetti - Funzione ITERAZIONE OGGETTI DEL DATABASE MDB CORRENTE - MODULO DA UTILIZZARE ----> UTIL_Nro40_N01_CONTROLLO_OGGETTI_DB
        '//PARAMETRI IN IMPUT:
        '//par_sx_Scelta_db =database ;                         = NON ATTIVATO          ----> _
           par_sx_Name_Tab                                      = ATTIVATO              ---> nome tabella da ricercare; _
           par_sx_Name_Query                                    = ATTIVATO              ---> nome query da ricercare _
           par_sTipo_Oggetto                                    = ATTIVATO              ---> tipo oggetto Query, tabella ecc..; _
           par_CodCmd_s                                         = ATTIVATO              ---> comando : cancella Query, salva tabella ecc..; _
           par_sx_Name_COD_PROGETTO_s                                   = ATTIVATO              ---> CODICE PROGETTO inserito il 22.05.2022 _
           par_ix_Operazione                                    = NON ATTIVATO          ----> Tipo di operazione =?  ; _
           par_ix_Indice1                                       = NON ATTIVATO          ----> ? ; _
           par_bx_Flag                                          = NON ATTIVATO          ----> True individuato ??
        '//OUTPUT : RESTITUISCE una stringa
        '//faq: @controllo.oggetti, @chiama.controllo oggetti, @attivo.il.controllo.oggetti, @iterazione.oggetti
        Public Function Iterazione_oggetti(par_sx_Scelta_db As String, _
                                           par_sx_Name_Tab As String, _
                                           par_sx_Name_Query As String, _
                                           par_sx_Name_FORM As String, _
                                           par_sx_Name_Report As String, _
                                           par_sx_Name_MODULI As String, _
                                           par_sx_Name_MACRO As String, _
                                           par_sx_Name_COD_PROGETTO_s As String, _
                                           par_sTipo_Oggetto As String, _
                                           par_CodCmd_s As String, _
                                           par_ix_Operazione As Integer, _
                                           par_ix_Indice1 As Integer, _
                                           par_bx_Flag As Boolean) As String


                On Error GoTo Err_Iterazione_oggetti

                '//reset
                icount = 0

                '------------------------------------------------------------------------------------------
                '1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB

                            'OPEN DB: Apre il database scelto per la creazione o cancellazione
                            '         con il parametro Option = false per indicare che l'apertura è condivisa /
                            '                          Option = True apertura esclusiva

                 If par_sx_Scelta_db = "" Then
                         'apro il db corrente se non viene indicato un database esterno
                         Set dbs = CurrentDb
                     Else
                         'apro un database esterno con la path >0 passata con i parametri
                         Set dbs = OpenDatabase(par_sx_Scelta_db, False)

                 End If




                '//ACCODA O CANCELLA ELENCO OGGETTO DELLA TABELLA @DF12
                '//--------------------------------------------------------------------------------//----------//
                '// NOTE: QUERY DI ACCODAMENTO E CANCELLAZIONE DEGLI ELEMENTI DELLA TABELLA
                '// @UTIL_Nro40_N01_CONTROLLO_OGGETTI_DB.OGGETTI.DF12
                '// @cancella.elenco.progetti @cmd.02_CANCELLA_MSys_ELENCO_ObjProgetti _
                    @comando.cmd.02_CANCELLA_MSys_ELENCO_ObjProgetti _
                    @cancella.DF12. @DF12.accoda @DF12.insert, @query.accoda.Msys_DF12


                            '//@CANCELLO.ELENCO.OGGETTI
                            If par_CodCmd_s = "Cmd_Cancella_MSys_ELENCO_ObjProgetti" Then
                                     '//se eseguo la query
                                     DoCmd.OpenQuery "MSys_ELENCO_ObjProgetti_Qry01_81_DELETE_TUTTI"


                                     par_CodCmd_s = "" '// annullo anche il codice comando eseguito per evitare errori
                                     '//esco dalla routine
                                     GoTo Exit_Iterazione_oggetti


                            End If


                            '

                             '//@CANCELLO.TUTTO @CANCELLO.TABELLE+QUERY+FORM+REPORT+MODULI, @Cmd_CANCELLA_TUTTO_TO_Msys_DF12, @oggetti.cancello.tutto _
                                @CANCELLAZIONE.COLLETTIVA.OGGETTI, @msys.oggetti.cancella.tutto. @cancellazione.multipla.oggetti
                            If par_CodCmd_s = "Cmd_CANCELLA_TUTTO_TO_Msys_DF12" Then

                                     '//se eseguo le 5 query di cancellazione : FORM + MODULI + QUERY _
                                                    + TABELLE + REPORT tutti gli oggetti msys
                                     DoCmd.OpenQuery "MSys_FORMS_Qry01_81_DELETE_TUTTE"
                                     DoCmd.OpenQuery "MSys_MODULI_Qry01_81_DELETE_TUTTE"
                                     DoCmd.OpenQuery "MSys_QUERY_Qry01_81_DELETE_TUTTE"
                                     DoCmd.OpenQuery "MSys_TABELLE_Qry01_81_DELETE_TUTTE"
                                     DoCmd.OpenQuery "MSys_REPORT_Qry01_81_DELETE_TUTTE"
                                     DoCmd.OpenQuery "MSys_MACRO_Qry01_81_DELETE_TUTTE"


                                     par_CodCmd_s = "" '// annullo anche il codice comando eseguito per evitare errori
                                     '//esco dalla routine
                                     GoTo Exit_Iterazione_oggetti


                            End If





                            '//ACCODO LE FORMS IN DF12, _
                            @accoda.elenco.progetti  @cmd.04_ACCODA_FORMS_TO_Msys_DF12 @comando.cmd.04_ACCODA_FORMS_TO_Msys_DF12 _
                            cmd.04_ACCODA_FORMS_TO_Msys_DF12 @accoda.DF12.FORMS, @accoda.form.to.DF12
                            '//.........................................................................................//
                                If par_CodCmd_s = "Cmd_Accoda_Forms_TO_Msys_DF12" Then

                                        DoCmd.OpenQuery "MSys_FORMS_Qry01_71_INSERT_toMsys_DF12"  '//@accoda.FORMS.to.DF12
                                        par_CodCmd_s = "" '// annullo anche il codice comando eseguito per evitare errori
                                        '//esco dalla routine
                                        GoTo Exit_Iterazione_oggetti


                                End If


                           '//ACCODO LE TABELLE IN DF12, _
                            @accoda.elenco.progetti  @cmd.04_ACCODA_TABELLE_TO_Msys_DF12 @comando.cmd.04_ACCODA_TABELLE_TO_Msys_DF12 _
                            cmd.04_ACCODA_TABELLE_TO_Msys_DF12 @accoda.DF12.TABELLE, @accoda.TABELLE.to.DF12
                            '//.........................................................................................//
                                If par_CodCmd_s = "Cmd_Accoda_Tabelle_TO_Msys_DF12" Then

                                        DoCmd.OpenQuery "MSys_TABELLE_Qry01_71_INSERT_toMsys_DF12"  '//@accoda.TABELLE.to.DF12
                                        par_CodCmd_s = "" '// annullo anche il codice comando eseguito per evitare errori
                                        '//esco dalla routine
                                        GoTo Exit_Iterazione_oggetti

                                End If
                            '//.........................................................................................//


                             '//ACCODO LE QUERY IN DF12, _
                               @accoda.elenco.progetti  @cmd.04_ACCODA_QUERY_TO_Msys_DF12 @comando.cmd.04_ACCODA_QUERY_TO_Msys_DF12 _
                               @accoda.DF12, @DF12.cancella @DF12.delete _
                               @accoda.query.to.DF12,

                            If par_CodCmd_s = "Cmd_Accoda_Query_TO_Msys_DF12" Then

                                     '//se eseguo la query
                                     DoCmd.OpenQuery "MSys_QUERY_Qry01_71_INSERT_toMsys_DF12"

                                     par_CodCmd_s = "" '// annullo anche il codice comando eseguito per evitare errori
                                     '//esco dalla routine
                                     GoTo Exit_Iterazione_oggetti
                            End If


                            '//ACCODO I MODULI IN DF12, _
                               @accoda.elenco.progetti  @cmd.04_ACCODA_MODULI_TO_Msys_DF12 @comando.cmd.04_ACCODA_MODULI_TO_Msys_DF12 _
                               @accoda.DF12, @DF12.cancella @DF12.delete _
                               @accoda.MODULI.to.DF12

                            If par_CodCmd_s = "Cmd_Accoda_MODULI_TO_Msys_DF12" Then

                                     '//se eseguo la QUERY SUI MODULI
                                     DoCmd.OpenQuery "MSys_MODULI_Qry01_71_INSERT_toMsys_DF12"

                                     par_CodCmd_s = "" '// annullo anche il codice comando eseguito per evitare errori
                                     '//esco dalla routine
                                     GoTo Exit_Iterazione_oggetti
                            End If


                             '//ACCODO I REPORT IN DF12, _
                               @accoda.elenco.progetti  @cmd.04_ACCODA_REPORT_TO_Msys_DF12 @comando.cmd.04_ACCODA_REPORT_TO_Msys_DF12 _
                               @accoda.DF12, @DF12.cancella @DF12.delete _
                               @accoda.REPORT.to.DF12
                            If par_CodCmd_s = "Cmd_Accoda_REPORT_TO_Msys_DF12" Then

                                     '//se eseguo la QUERY SUI REPORT
                                     DoCmd.OpenQuery "MSys_REPORT_Qry01_71_INSERT_toMsys_DF12"

                                     par_CodCmd_s = "" '// annullo anche il codice comando eseguito per evitare errori
                                     '//esco dalla routine
                                     GoTo Exit_Iterazione_oggetti
                            End If

                           '//ACCODO LE MACRO IN DF12, _
                            @accoda.elenco.progetti  @cmd.04_ACCODA_MACRO_TO_Msys_DF12 @comando.cmd.04_ACCODA_MACRO_TO_Msys_DF12 _
                            @accoda.DF12, @DF12.cancella @DF12.delete _
                            @accoda.MACRO.to.DF12
                            If par_CodCmd_s = "Cmd_Accoda_Macro_TO_Msys_DF12" Then

                                 '//se eseguo la QUERY SULLE MACRO
                                 DoCmd.OpenQuery "MSys_MACRO_Qry01_71_INSERT_toMsys_DF12"

                                 par_CodCmd_s = "" '// annullo anche il codice comando eseguito per evitare errori
                                 '//esco dalla routine
                                 GoTo Exit_Iterazione_oggetti
                            End If



                    '//ACCODA O CANCELLA ELENCO OGGETTO DELLA TABELLA @DF12  *** FINE ***
                    '//--------------------------------------------------------------------------------//----------//



                    '//@MSys.OGGETTI.TUTTI, @OGGETTI.MSys.TUTTI
                    Select Case par_sTipo_Oggetto




                        '//ESEGUO PER GLI OGGETTI DEL DB, @03_IMPORTA_TABELLE_IN_DB  @importa.tabelle _
                        @03_IMPORTA_TABELLE_IN_DB @funzione.importa.tabelle, @importa.tabelle.cmd, @cmd.importa.tabelle

                        Case Is = "TABLE"

                                'ITERAZIONE TABLEDEFS <<LE TABELLE>> @table.iterazione, @oggetti.table, @tabelle.iterazione
                                '______________________________________________________________________________
                                'OGGETTI TABLE
                                'CONTROLLO OGGETTO NEL DB: Controllo l'esistenza delle TABELLE



                                    '//@cancellazione.Tabelle.Msys, @cancella.Tabelle.tutte.Msys, @delete.table.tutte.Msys
                                    'pulisco la tabella delle Tabelle prima del salvataggio dei nuovi nomi,
                                    CurrentDb.Execute ("DELETE MSys_TABELLE_Qry01_01_SELECT_TUTTE.* FROM MSys_TABELLE_Qry01_01_SELECT_TUTTE;")

                                    '//SE SOLO CANCELLAZIONE ESCO DALLA ROUTINE
                                    If par_CodCmd_s = "Cmd_Cancella_Table" Then

                                        GoTo Exit_Iterazione_oggetti

                                    End If



                                    'apro la tabella per il salvataggio delle tabelle @importa.tabelle, @salva.tabelle
                                    Set ApriRs = dbs.OpenRecordset("MSys_TABELLE_Qry01_01_SELECT_TUTTE")


                                    With dbs



                                        Debug.Print .TableDefs.Count & _
                                            " TableDefs in database  " & sx_Scelta_db

                                        For intCiclo = 0 To .TableDefs.Count - 1



                                             '//CONTROLLO OGGETTO TABELLA
                                            '//-----------------------------------------------------------
                                                '//2022.04.19.ESCLUSIONE.TABELLE = è stata fatta una modifica _
                                                    per evitare di salvare nel rs anche le tabelle di sistema _
                                                    che non possono essere esportate. Il controllo viene fatto da _
                                                    una if multipla che verifica tutte le tabelle di sistema da escludere _
                                                    nel salvataggio

                                                    '//FILTRO_ESCLUSIONE @esclusione.tabelle, @tabelle.escluse.salvataggio

                                                        If .TableDefs(intCiclo).Name <> "MSysAccessObjects" _
                                                        And .TableDefs(intCiclo).Name <> "MSysAccessXML" _
                                                        And .TableDefs(intCiclo).Name <> "MSysACEs" _
                                                        And .TableDefs(intCiclo).Name <> "MSysNavPaneGroupCategories" _
                                                        And .TableDefs(intCiclo).Name <> "MSysNavPaneGroups" _
                                                        And .TableDefs(intCiclo).Name <> "MSysNavPaneGroupToObjects" _
                                                        And .TableDefs(intCiclo).Name <> "MSysNavPaneObjectIDs" _
                                                        And .TableDefs(intCiclo).Name <> "MSysObjects" _
                                                        And .TableDefs(intCiclo).Name <> "MSysQueries" _
                                                        And .TableDefs(intCiclo).Name <> "MSysRelationships" _
                                                        Then


                                                             'SALVATAGGIO NOME TABELLA, @salvataggio.tabelle
                                                             '................................................................

                                                                      Debug.Print "  " & .TableDefs(intCiclo).Name
                                                                         ApriRs.AddNew
                                                                         ApriRs.Fields("Name1_s") = .TableDefs(intCiclo).Name
                                                                         '//@imposto.tipo.oggetto
                                                                         ApriRs.Fields("TIPOGGETTO_s") = "TABELLE"
                                                                         '//il contatore record
                                                                         icount = icount + 1
                                                                         ApriRs.Fields("NRO_OGGETTO_i") = icount



                                                                        '//@errore.corretto.22.05.2022 = durante il salvataggio delle tabelle, _
                                                                        macro, form query, report e moduli è stato corretto l'errore _
                                                                        in quanto occorre anche il salvataggio del codice progetto _
                                                                        che viene passato come parametro: par_sx_Name_COD_PROGETTO_s


                                                                           ApriRs.Fields("COD_PROGETTO_s") = par_sx_Name_COD_PROGETTO_s


                                                                         ApriRs.Update


                                                                         '//CONTROLLO WINDOWS
                                                                         DoEvents

                                                             '................................................................
                                                        Else
                                                                Debug.Print "tabella di sistema non salvata  " & .TableDefs(intCiclo).Name
                                                        End If



                                            If .TableDefs(intCiclo).Name = par_sx_Name_Tab Then
                                                Bln1 = True    ' bx_Flag True = trovato tabelle si può cancellare
                                                Iterazione_oggetti = Bln1 'riassegno alla funzione il valore del bx_Flag
                                                GoTo Exit_Iterazione_oggetti  'esci
                                            Else
                                                Bln1 = False   ' bx_Flag False = non trovato TABELLA
                                                Iterazione_oggetti = Bln1 'riassegno alla funzione il valore del bx_Flag
                                            End If

                                        Next intCiclo

                                    End With

                        Case Is = "QUERY"

                                'ITERAZIONE QUERYDEFS <<LE QUERY>> @Query.iterazione, @oggetti.Query, @Query.iterazione _
                                @03_IMPORTA_QUERY_IN_DB, @IMPORTA_QUERY_IN_DB
                                '______________________________________________________________________________
                                'OGGETTI QUERY
                                'CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della QUERY

                                    '//todo: IMPOSTARE IL FILTRO DEI COMANDI _
                                        La azioni per i comandi sono: _
                                        Cmd_Cancella_Query = Esegue la cancellazione con la query ma esce dalla _
                                                            routine. _
                                        Cmd_Salva_Query     = Esegue la cancellazione e l'esecuzione della routine, _
                                                            la quale cicla fra gli oggetti del db e salva solo le query _
                                                            nella tabella QUERY.


                                    '//@cancellazione.query.Msys, @cancella.query.tutte.Msys, 02_CANCELLA_QUERY
                                    'pulisco la tabella delle query prima del salvataggio dei nuovi nomi,
                                    CurrentDb.Execute ("DELETE MSys_QUERY_Qry01_01_SELECT_TUTTE.* FROM MSys_QUERY_Qry01_01_SELECT_TUTTE;")

                                    '//SE SOLO CANCELLAZIONE ESCO DALLA ROUTINE
                                    If par_CodCmd_s = "Cmd_Cancella_Query" Then

                                            GoTo Exit_Iterazione_oggetti

                                    End If




                                    'apro la tabella per il salvataggio delle query - @importa.query.tutte
                                    Set ApriRs = dbs.OpenRecordset("MSys_QUERY_Qry01_01_SELECT_TUTTE")

                                                '//@salvataggio.query.Msys
                                                'itero nel db
                                                With dbs

                                                        Debug.Print .QueryDefs.Count & _
                                                            " QueryDefs in database  " & sx_Scelta_db

                                                        '// CICLO FOR OGGETTI QUERY: @ciclo.for.QueryDefs
                                                        For intCiclo = 0 To .QueryDefs.Count - 1



                                                            '// QUERY ESCLUSE CHE INIZIANO CON = ~sq_  , @QUERY.ESCLUSE UTIL_Nro40_N01_CONTROLLO_OGGETTI_DB.@query.salvate
                                                            '//=====================================================================================//
                                                            '//nota: salvo solo le query create non di sistema escludendo quelle cancellate con _
                                                                    segnatura iniziale che inizia con  ~sq_ ....

                                                                If Left(.QueryDefs(intCiclo).Name, 4) <> "~sq_" Then

                                                                        'controlli
                                                            '................................................................
                                                                '//@query.salvate
                                                                Debug.Print "query salvata:  " & .QueryDefs(intCiclo).Name

                                                                'Stampo il contenuto sql della query
                                                                Debug.Print
                                                                Debug.Print "  " & .QueryDefs(intCiclo).SQL

                                                            '................................................................

                                                                '//SALVATAGGIO NOME QUERY    @salvataggio.nome.query
                                                                '................................................................
                                                                    ApriRs.AddNew

                                                                    '//il contatore record
                                                                    icount = icount + 1
                                                                    ApriRs.Fields("NRO_OGGETTO_i") = icount

                                                                    '//@imposto.tipo.oggetto
                                                                    ApriRs.Fields("TIPOGGETTO_s") = "QUERY"


                                                                        '//@errore.corretto.22.05.2022 = durante il salvataggio delle tabelle, _
                                                                        macro, form query, report e moduli è stato corretto l'errore _
                                                                        in quanto occorre anche il salvataggio del codice progetto _
                                                                        che viene passato come parametro: par_sx_Name_COD_PROGETTO_s


                                                                    ApriRs.Fields("COD_PROGETTO_s") = par_sx_Name_COD_PROGETTO_s

                                                                    ApriRs.Fields("Name1_s") = .QueryDefs(intCiclo).Name


                                                                    ApriRs.Update
                                                                '................................................................


                                                                        '//RICERCA QUERY CON PARAMETRO _
                                                                           imposto il parametro con il nome della query da ricercare e lo _
                                                                           confronto con il nome oggetto _
                                                                           @query.ricerca.parametro _
                                                                           @controllo.nome.query.passata.con.parametro
                                                                        If .QueryDefs(intCiclo).Name = par_sx_Name_Query Then

                                                                            Bln1 = True    ' bx_Flag True = trovato QUERY

                                                                            'stampo il contenuto della query scelta direttamente _
                                                                            solo se esiste la query ricercata.
                                                                            Debug.Print "  " & .QueryDefs(par_sx_Name_Query).SQL


                                                                            'recupero il contenuto della query
                                                                            Iterazione_oggetti = .QueryDefs(intCiclo).SQL


                                                                            GoTo Exit_Iterazione_oggetti  'esci
                                                                        Else
                                                                            Bln1 = False   ' bx_Flag False = non trovato QUERY
                                                                            Iterazione_oggetti = Bln1 'riassegno alla funzione il valore del bx_Flag
                                                                        End If

                                                                Else
                                                                        Debug.Print "query di sistema non salvata:   " & .QueryDefs(intCiclo).Name

                                                                End If  '//If Left(.QueryDefs(intCiclo).Name, 4)
                                                            '// QUERY ESCLUSE CHE INIZIANO CON = ~sq_       @QUERY.ESCLUSE   *** fine ***
                                                            '//=====================================================================================//


                                                        Next intCiclo

                                                    End With


                                'ITERAZIONE QUERYDEFS <<LE QUERY>> *** fine
                                '______________________________________________________________________________

                        'CICLO NEGLI OGGETTI MODULI  - @moduli.itera.oggetti  TODO : costruisco i moduli, @importa.moduli _
                        @03_IMPORTA_MODULI_IN_DB, @IMPORTA_MODULI_IN_DB
                       Case Is = "MODULI"
                                '//TODO DA FINIRE come iteri nei moduli


                                'ITERAZIONE MODULIDEFS <<LE MODULI>> @MODULI.iterazione, @oggetti.MODULI, @MODULI.iterazione
                                '______________________________________________________________________________
                                'OGGETTI MODULI
                                'CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MODULI

                                    '//todo: IMPOSTARE IL FILTRO DEI COMANDI _
                                        La azioni per i comandi sono: _
                                        Cmd_Cancella_MODULI = Esegue la cancellazione con la MODULI ma esce dalla _
                                                            routine. _
                                        Cmd_Salva_MODULI     = Esegue la cancellazione e l'esecuzione della routine, _
                                                            la quale cicla fra gli oggetti del db e salva solo le MODULI _
                                                            nella tabella MODULI.


                                    '//@cancellazione.MODULI.Msys, @cancella.MODULI.tutte.Msys
                                    'pulisco la tabella delle MODULI prima del salvataggio dei nuovi nomi,
                                    CurrentDb.Execute ("DELETE MSys_MODULI_Qry01_01_SELECT_TUTTI.* FROM MSys_MODULI_Qry01_01_SELECT_TUTTI;")

                                    '//SE SOLO CANCELLAZIONE ESCO DALLA ROUTINE
                                    If par_CodCmd_s = "Cmd_Cancella_MODULI" Then

                                            GoTo Exit_Iterazione_oggetti

                                    End If




                                    'apro la tabella per il salvataggio delle MODULI - @importa.MODULI.tutte
                                    Set ApriRs = dbs.OpenRecordset("MSys_MODULI_Qry01_01_SELECT_TUTTI")


                                '//QUALIFICO DBS - @salvataggio.MODULI.Msys itero nel db - qualifico la variabile oggetto
                                '//i moduli : https://docs.microsoft.com/en-us/office/vba/api/access.module
                                '//=========================================================================================//
                                    With dbs

                                                        '//ITERAZIONE NELL'INSIEMO MODULI   @moduli.itera.oggetti
                                                        Set dbs = Application.CurrentProject
                                                         Debug.Print "tutti i moduli del db: " & dbs.AllModules.Count & _
                                                                                " MODULIDefs in database  " & sx_Scelta_db


                                                        icount = 0
                                                        ' Ricerca oggetti AccessObject aperti in insieme AllModules.
                                                        For Each obj In dbs.AllModules


                                                         'controlli DEI MODULI NAME
                                                        '................................................................



                                                            '//contatore dei moduli
                                                                icount = icount + 1
                                                            Debug.Print "NRO    -> " & icount
                                                            Debug.Print "modulo -> " & obj.Name
                                                            Debug.Print "modulo TYPE -> " & obj.Type
                                                            Debug.Print " FULL NAME =  " & obj.FullName






                                                        '................................................................


                                                        'SALVATAGGIO NOME MODULI
                                                        '................................................................
                                                            ApriRs.AddNew

                                                            '//il contatore record
                                                            icount = icount + 1
                                                            ApriRs.Fields("NRO_OGGETTO_i") = icount

                                                            '//@imposto.tipo.oggetto
                                                            ApriRs.Fields("TIPOGGETTO_s") = "MODULES"



                                                            '//@errore.corretto.22.05.2022 = durante il salvataggio delle tabelle, _
                                                            macro, form query, report e moduli è stato corretto l'errore _
                                                            in quanto occorre anche il salvataggio del codice progetto _
                                                            che viene passato come parametro: par_sx_Name_COD_PROGETTO_s


                                                            ApriRs.Fields("COD_PROGETTO_s") = par_sx_Name_COD_PROGETTO_s


                                                            ApriRs.Fields("Name1_s") = obj.Name


                                                            ApriRs.Update
                                                        '................................................................


                                                            'controllo nome MODULI passata con parametro @SOSTITUIRE.CON.par_sx_Name_MODULI
                                                            '................................................................
                                                            If obj.Name = par_sx_Name_MODULI Then

                                                            Bln1 = True    ' bx_Flag True = trovato MODULI

                                                            'stampo il contenuto della MODULI scelta direttamente _
                                                            solo se esiste la MODULI ricercata.
                                                            Debug.Print "  " & obj.Name


                                                            'recupero il contenuto della MODULI @MUDULO.SQL
                                                            Iterazione_oggetti = obj.SQL


                                                            GoTo Exit_Iterazione_oggetti  'esci
                                                            Else
                                                            Bln1 = False   ' bx_Flag False = non trovato MODULI
                                                            Iterazione_oggetti = Bln1 'riassegno alla funzione il valore del bx_Flag
                                                            End If
                                                            '................................................................



                                                            '//controllo se il modulo è aperto - @MODULO.APERTO
                                                            If obj.IsLoaded = True Then
                                                                ' Stampa nome di obj.
                                                                Debug.Print "modulo APERTO -> " & obj.Name
                                                            End If


                                                        Next obj


                                    End With
                                '//*** FINE ***
                                '//QUALIFICO DBS - @salvataggio.MODULI.Msys itero nel db - qualifico la variabile oggetto
                                '//=========================================================================================//

                                'ITERAZIONE MODULIDEFS <<LE MODULI>> *** fine
                                '______________________________________________________________________________





                       'CICLO NEGLI OGGETTI FORMS  @importa.forms @03_IMPORTA_FORMS_IN_DB, @IMPORTA_FORMS_IN_DB
                       Case Is = "FORMS"

                                'ITERAZIONE FORMS DEFS <<LE FORMS>>
                                '______________________________________________________________________________
                                'OGGETTI FORMS
                                'CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORMS


                                    'NOME form prorvvisoria ca controllare, poi cancellare
                                    par_sx_Name_FORM = "AMB_Frm01_M01_ARCHIVIO_AMBI"






                                    '//@cancellazione.forms.Msys, @cancella.form.tutte.Msys, @cancella.forms.Msys
                                    'pulisco la tabella delle FORM prima del salvataggio dei nuovi nomi
                                    CurrentDb.Execute ("DELETE MSys_FORMS_Qry01_01_SELECT_TUTTE.* FROM MSys_FORMS_Qry01_01_SELECT_TUTTE;")


                                    '//SE SOLO CANCELLAZIONE ESCO DALLA ROUTINE,
                                    If par_CodCmd_s = "Cmd_Cancella_Forms" Then

                                            GoTo Exit_Iterazione_oggetti

                                    End If


                                    'apro la tabella per il salvataggio delle forms,  @importa.forms.Msys, @importa.form.tutte.Msys
                                    Set ApriRs = dbs.OpenRecordset("MSys_FORMS_Qry01_01_SELECT_TUTTE")

                                     '//ITERAZIONE NELL'INSIEMO MODULI   @moduli.itera.oggetti
                                    Set dbs = Application.CurrentProject


                                    intCiclo = 0





                                'itero nel db
                                With dbs



                                        For Each obj In dbs.AllForms

                                            ' Debug.Print obj_Oggetto.Count _
                                           '& " FormsDefs in database  " & sx_Scelta_db


                                            'controlli
                                            '................................................................



                                                Debug.Print "  " & obj.Name

                                                'Stampo il contenuto sql della query
                                                Debug.Print
                                                'Debug.Print "  " & CurrentProject(intCiclo).sQL

                                                'stampo il contenuto della query scelta direttamente
                                                'Debug.Print "  " & CurrentProject("MSsys_VS01_MSys_QUERY_Qry01_01_SELECT_TUTTE").sQL
                                            '................................................................

                                                'SALVATAGGIO NOME FORMS
                                                '................................................................
                                                    ApriRs.AddNew
                                                    ApriRs.Fields("Name1_s") = obj.Name


                                                    '//il contatore record
                                                    icount = icount + 1
                                                    ApriRs.Fields("NRO_OGGETTO_i") = icount

                                                    '//@imposto.tipo.oggetto
                                                    ApriRs.Fields("TIPOGGETTO_s") = "FORMS"


                                                   '@errore.corretto.22.05.2022 = durante il salvataggio dei FORMS _
                                                    in quanto occorre anche il salvataggio del codice progetto _
                                                    che viene passato come parametro: par_sx_Name_COD_PROGETTO_s
                                                    ApriRs.Fields("COD_PROGETTO_s") = par_sx_Name_COD_PROGETTO_s




                                                    ApriRs.Update
                                                '................................................................

                                            'controllo nome query passata con parametro
                                            If obj.Name = par_sx_Name_FORM Then


                                                '??? erroe nella variabile di controllo
                                                'Trovo il numero dei controlli della form
                                                Vv1 = objx_Form.Controls.Count
                                                ix_Count_Control = 0

                                                'itero nei controlli della form
                                                For Each ctr In objx_Form.Controls

                                                        'Stampo il nome del controllo
                                                        Debug.Print " Nome del controllo  " & ctr(ix_Count_Control).Name

                                                    ix_Count_Control = ix_Count_Control + 1

                                                Next ctr


                                                Bln1 = True    ' bx_Flag True = trovato QUERY


                                                'recupero il contenuto della query
                                                'Iterazione_oggetti = CurrentProject(intCiclo).sQL


                                                GoTo Exit_Iterazione_oggetti  'esci
                                            Else
                                                Bln1 = False   ' bx_Flag False = non trovato FORMS
                                                Iterazione_oggetti = Bln1 'riassegno alla funzione il valore del bx_Flag
                                            End If


                                                intCiclo = intCiclo + 1

                                        Next obj

                                    End With

                                    '//CHIUSURA OGGETTO
                                    Set obj_Oggetto = Nothing


                      '@importa.report
                      Case Is = "REPORTS"


                                'ITERAZIONE REPORTDEFS <<LE REPORT>> @REPORT.iterazione, @oggetti.REPORT, @REPORT.iterazione _
                                @03_IMPORTA_REPORT_IN_DB, @IMPORTA_REPORT_IN_DB
                                '______________________________________________________________________________
                                'OGGETTI REPORT
                                'CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della REPORT

                                    '//todo: IMPOSTARE IL FILTRO DEI COMANDI _
                                        La azioni per i comandi sono: _
                                        Cmd_Cancella_Reports = Esegue la cancellazione con la REPORT ma esce dalla _
                                                            routine. _
                                        Cmd_Salva_REPORT     = Esegue la cancellazione e l'esecuzione della routine, _
                                                            la quale cicla fra gli oggetti del db e salva solo le REPORT _
                                                            nella tabella REPORT.


                                    '//@cancellazione.REPORT.Msys, @cancella.REPORT.tutte.Msys
                                    'pulisco la tabella delle REPORT prima del salvataggio dei nuovi nomi,
                                    CurrentDb.Execute ("DELETE MSys_REPORT_Qry01_01_SELECT_TUTTE.* FROM MSys_REPORT_Qry01_01_SELECT_TUTTE;")

                                    '//SE SOLO CANCELLAZIONE ESCO DALLA ROUTINE
                                    If par_CodCmd_s = "Cmd_Cancella_Reports" Then

                                            GoTo Exit_Iterazione_oggetti

                                    End If





                                    'apro la tabella per il salvataggio delle REPORT - @importa.REPORT.tutte
                                    Set ApriRs = dbs.OpenRecordset("MSys_REPORT_Qry01_01_SELECT_TUTTE")

                                                '//@salvataggio.REPORT.Msys
                                                'itero nel db
                                                'Set obj_Oggetto = Application.CurrentProject
                                             '//ITERAZIONE NELL'INSIEMO MODULI   @moduli.itera.oggetti
                                            Set dbs = Application.CurrentProject
                                                        icount = 0
                                                        ' Ricerca oggetti AccessObject aperti in insieme AllModules.


                                                   'itero nel db
                                With dbs



                                        For Each obj In dbs.AllReports

                                            ' Debug.Print obj_Oggetto.Count _
                                           '& " FormsDefs in database  " & sx_Scelta_db


                                            'controlli
                                            '................................................................



                                                Debug.Print "  " & obj.Name

                                                'Stampo il contenuto sql della query
                                                Debug.Print
                                                'Debug.Print "  " & CurrentProject(intCiclo).sQL

                                                'stampo il contenuto della query scelta direttamente
                                                'Debug.Print "  " & CurrentProject("MSsys_VS01_MSys_QUERY_Qry01_01_SELECT_TUTTE").sQL

                                            '................................................................

                                                'SALVATAGGIO NOME REPORT
                                                '................................................................
                                                    ApriRs.AddNew
                                                    ApriRs.Fields("Name1_s") = obj.Name


                                                    '//il contatore record
                                                    icount = icount + 1
                                                    ApriRs.Fields("NRO_OGGETTO_i") = icount



                                                    '@errore.corretto.22.05.2022 = durante il salvataggio dei REPORT _
                                                    in quanto occorre anche il salvataggio del codice progetto _
                                                    che viene passato come parametro: par_sx_Name_COD_PROGETTO_s
                                                    ApriRs.Fields("COD_PROGETTO_s") = par_sx_Name_COD_PROGETTO_s

                                                   '//@imposto.tipo.oggetto
                                                    ApriRs.Fields("TIPOGGETTO_s") = "REPORTS"






                                                    ApriRs.Update
                                                '................................................................

                                            'controllo nome query passata con parametro
                                            If obj.Name = par_sx_Name_Report Then


                                                '??? erroe nella variabile di controllo
                                                'Trovo il numero dei controlli della form
                                                Vv1 = objx_Form.Controls.Count
                                                ix_Count_Control = 0

                                                'itero nei controlli della form
                                                For Each ctr In objx_Form.Controls

                                                        'Stampo il nome del controllo
                                                        Debug.Print " Nome del controllo  " & ctr(ix_Count_Control).Name

                                                    ix_Count_Control = ix_Count_Control + 1

                                                Next ctr


                                                Bln1 = True    ' bx_Flag True = trovato REPORT


                                                'recupero il contenuto della query
                                                'Iterazione_oggetti = CurrentProject(intCiclo).sQL


                                                GoTo Exit_Iterazione_oggetti  'esci
                                            Else
                                                Bln1 = False   ' bx_Flag False = non trovato FORMS
                                                Iterazione_oggetti = Bln1 'riassegno alla funzione il valore del bx_Flag
                                            End If


                                                intCiclo = intCiclo + 1

                                        Next obj

                                    End With


                                    '//CHIUSURA OGGETTO, @chiusura.oggetto
                                    Set obj_Oggetto = Nothing


                                'ITERAZIONE REPORT DEFS <<LE REPORT>> *** fine
                                '______________________________________________________________________________

                      '//@importa.MACRO, @03_IMPORTA_MACRO_IN_DB, @IMPORTA_MACRO_IN_DB
                      Case Is = "MACRO"


                                'ITERAZIONE MACRODEFS <<LE MACRO>> @MACRO.iterazione, @oggetti.MACRO, @MACRO.iterazione, @04_ACCODA_MACRO_TO_Msys_DF12
                                '______________________________________________________________________________
                                'OGGETTI MACRO
                                'CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO

                                    '//todo: IMPOSTARE IL FILTRO DEI COMANDI _
                                        La azioni per i comandi sono: _
                                        Cmd_Cancella_MACROS = Esegue la cancellazione con la MACRO ma esce dalla _
                                                            routine. _
                                        Cmd_Salva_MACRO     = Esegue la cancellazione e l'esecuzione della routine, _
                                                            la quale cicla fra gli oggetti del db e salva solo le MACRO _
                                                            nella tabella MACRO.


                                    '//@cancellazione.MACRO.Msys, @cancella.MACRO.tutte.Msys
                                    'pulisco la tabella delle MACRO prima del salvataggio dei nuovi nomi,
                                    CurrentDb.Execute ("DELETE MSys_MACRO_Qry01_01_SELECT_TUTTI.* FROM MSys_MACRO_Qry01_01_SELECT_TUTTI;")

                                    '//SE SOLO CANCELLAZIONE ESCO DALLA ROUTINE
                                    If par_CodCmd_s = "Cmd_Cancella_Macro" Then

                                            GoTo Exit_Iterazione_oggetti

                                    End If





                                    'apro la tabella per il salvataggio delle MACRO - @importa.MACRO.tutte
                                    Set ApriRs = dbs.OpenRecordset("MSys_MACRO_Qry01_01_SELECT_TUTTI")

                                                '//@salvataggio.MACRO.Msys
                                                'itero nel db
                                                'Set obj_Oggetto = Application.CurrentProject
                                             '//ITERAZIONE NELL'INSIEMO MODULI   @moduli.itera.oggetti
                                            Set dbs = Application.CurrentProject
                                                        icount = 0
                                                        ' Ricerca oggetti AccessObject aperti in insieme AllMacros.


                                                   'itero nel db
                                With dbs



                                        For Each obj In dbs.AllMacros

                                            ' Debug.Print obj_Oggetto.Count _
                                           '& " FormsDefs in database  " & sx_Scelta_db


                                            'controlli
                                            '................................................................



                                                Debug.Print "  " & obj.Name

                                                'Stampo il contenuto sql della query
                                                Debug.Print
                                                'Debug.Print "  " & CurrentProject(intCiclo).sQL

                                                'stampo il contenuto della query scelta direttamente
                                                'Debug.Print "  " & CurrentProject("MSsys_VS01_MSys_QUERY_Qry01_01_SELECT_TUTTE").sQL

                                            '................................................................

                                                'SALVATAGGIO NOME MACRO
                                                '................................................................
                                                    ApriRs.AddNew
                                                    ApriRs.Fields("Name1_s") = obj.Name


                                                    '//il contatore record
                                                    icount = icount + 1
                                                    ApriRs.Fields("NRO_OGGETTO_i") = icount


                                                   '//@imposto.tipo.oggetto
                                                    ApriRs.Fields("TIPOGGETTO_s") = "MACRO"


                                                    '@errore.corretto.22.05.2022 = durante il salvataggio dei REPORT _
                                                    in quanto occorre anche il salvataggio del codice progetto _
                                                    che viene passato come parametro: par_sx_Name_COD_PROGETTO_s
                                                    ApriRs.Fields("COD_PROGETTO_s") = par_sx_Name_COD_PROGETTO_s







                                                    ApriRs.Update
                                                '................................................................

                                            'controllo nome query passata con parametro
                                            If obj.Name = par_sx_Name_MACRO Then


                                                '??? erroe nella variabile di controllo
                                                'Trovo il numero dei controlli della form
                                                Vv1 = objx_Form.Controls.Count
                                                ix_Count_Control = 0

                                                'itero nei controlli della form
                                                For Each ctr In objx_Form.Controls

                                                        'Stampo il nome del controllo
                                                        Debug.Print " Nome del controllo  " & ctr(ix_Count_Control).Name

                                                    ix_Count_Control = ix_Count_Control + 1

                                                Next ctr


                                                Bln1 = True    ' bx_Flag True = trovato MACRO


                                                'recupero il contenuto della query
                                                'Iterazione_oggetti = CurrentProject(intCiclo).sQL


                                                GoTo Exit_Iterazione_oggetti  'esci
                                            Else
                                                Bln1 = False   ' bx_Flag False = non trovato FORMS
                                                Iterazione_oggetti = Bln1 'riassegno alla funzione il valore del bx_Flag
                                            End If


                                                intCiclo = intCiclo + 1

                                        Next obj

                                    End With


                                    '//CHIUSURA OGGETTO, @chiusura.oggetto
                                    Set obj_Oggetto = Nothing


                                'ITERAZIONE INSIEME MACRO DEFS <<LE MACRO>> *** fine
                                '______________________________________________________________________________




                            Case Else

                                    MsgBox "Scelta comando non effettuata o comando non trovato --> " & par_CodCmd_s, vbExclamation
                                     GoTo Exit_Iterazione_oggetti
                        End Select
                                    'chiudo i rs aperti + Rilascio la memoria
                                    ApriRs.Close
                                    Set ApriRs = Nothing

        '------------------------------------------------------------------------------
        '                       FINE FUNCTIONE E GESTIONE ERRORI

                Exit_Iterazione_oggetti:

                                    Exit Function


                Err_Iterazione_oggetti:
                            dbs.Close
                            Set dbs = Nothing
                            MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "Iterazione_oggetti"
                            Resume Exit_Iterazione_oggetti

        End Function
        '//=======================================================================================================================================================//

'//RECUPERO_PARAMETRI_TABELLA_OGGETTI_pFc- LOTTO RECUPERO PARAMETRI TABELLA OGGETTI




    '1)I CAMPI GENERALI DEL RS
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    'Nota: oggetti generali della classe, prelevati dal rs.
                'VALORI CAMPI DEL RS <<GEST_Tb01_N01_PARAM_OGGETTI>>
                '============================================================================================
                '1) Parametro                   - CODOGGETTO
                '2) Parametro                   - TIPOGGETTO
                '3) Parametro                   - PROPRIETA
                '4) Parametro                   - VALOREPROPRIETA
                '5) Parametro                   - EVENTO
                '6) Parametro                   - ATTIVITAEVENTO
                '7) Parametro                   - METODO
                '8) Parametro                   - CODATTIVITA
                '9) Parametro                   - ID





                    '1) CODOGGETTO
                    'valore del campo
                    '........................................................................


                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let CODOGGETTO_let(par_sxCODOGGETTO)

                                            m_sxCODOGGETTO = par_sxCODOGGETTO

                        End Property


                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get CODOGGETTO_Get() As String

                                            CODOGGETTO_Get = m_sxCODOGGETTO

                        End Property
                    '........................................................................

                    '2)TIPOGGETTO
                    'valore del campo
                    '........................................................................


                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let TIPOGGETTO_let(par_sxTIPOGGETTO)

                                            m_sxTIPOGGETTO = par_sxTIPOGGETTO

                        End Property


                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get TIPOGGETTO_Get() As String

                                            TIPOGGETTO_Get = m_sxTIPOGGETTO

                        End Property
                    '........................................................................


                    '3)PROPRIETA
                    'valore del campo
                    '........................................................................


                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let PROPRIETA_let(par_sxPROPRIETA)

                                            m_sxPROPRIETA = par_sxPROPRIETA

                        End Property


                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get PROPRIETA_Get() As String

                                            PROPRIETA_Get = m_sxPROPRIETA

                        End Property
                    '........................................................................


                    '4)VALOREPROPRIETA
                    'valore del campo
                    '........................................................................


                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let VALOREPROPRIETA_let(par_vxVALOREPROPRIETA)

                                            m_vxVALOREPROPRIETA = par_vxVALOREPROPRIETA

                        End Property


                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get VALOREPROPRIETA_Get() As String

                                            VALOREPROPRIETA_Get = m_vxVALOREPROPRIETA

                        End Property
                    '........................................................................


                    '5)EVENTO
                    'valore del campo
                    '........................................................................


                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let EVENTO_let(par_sxEVENTO)

                                            m_sxEVENTO = par_sxEVENTO

                        End Property


                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get EVENTO_Get() As String

                                            EVENTO_Get = m_sxEVENTO

                        End Property
                    '........................................................................



                    '6)ATTIVITAEVENTO
                    'valore del campo
                    '........................................................................


                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let ATTIVITAEVENTO_let(par_sxATTIVITAEVENTO)

                                            m_sxATTIVITAEVENTO = par_sxATTIVITAEVENTO

                        End Property


                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get ATTIVITAEVENTO_Get() As String

                                            ATTIVITAEVENTO_Get = m_sxATTIVITAEVENTO

                        End Property
                    '........................................................................



                    '7) METODO
                    'valore del campo
                    '........................................................................


                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                      Public Property Let METODO_let(par_sxMETODO)

                                            m_sxMETODO = par_sxMETODO

                        End Property


                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get METODO_Get() As String

                                            METODO_Get = m_sxMETODO

                        End Property
                    '........................................................................


                    '8) CODATTIVITA
                    'valore del campo
                    '........................................................................


                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                      Public Property Let CODATTIVITA_let(par_sxCODATTIVITA)

                                            m_sxCODATTIVITA = par_sxCODATTIVITA

                        End Property


                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get CODATTIVITA_Get() As String

                                            CODATTIVITA_Get = m_sxCODATTIVITA

                        End Property
                    '........................................................................




                    '9) IDOGGETTO
                    'valore del campo
                    '........................................................................


                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                      Public Property Let IDOGGETTOlet(par_lngXID)

                                            m_lngxIDOGGETTO = par_lngXID

                        End Property


                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get IDOGGETTOGet() As String

                                            IDOGGETTOGet = m_lngxIDOGGETTO

                        End Property
                    '........................................................................


                'ALTRI PARAMETI DA ESTRARRE DAL RS <<GEST_Tb01_N01_PARAM_OGGETTI>>
                '============================================================================================
                '1) Parametro                   - TABLE
                '2) Parametro                   - QUERY





                    '1) TABLE
                    'valore del campo
                    '........................................................................


                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let TABLE_let(par_TABLE)

                                            m_sxTABLE = par_TABLE

                        End Property


                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get TABLE_Get() As String

                                            TABLE_Get = m_sxTABLE

                        End Property
                    '........................................................................




                    '2) RECORD_SOURCE
                    'valore del campo
                    '........................................................................
                    'Nota: Recupero il valore della proprieta RecordSource della form, che
                    'può essere una query o una tabella da inserire nella proprieta della
                    'form stessa.

                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let RECORD_SOURCE_sxLet(par_RECORD_SOURCE_FORM)

                                            m_FRM_DATI_01_sxOrigineRecord_RecordSource = par_RECORD_SOURCE_FORM

                        End Property

                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get RECORD_SOURCE_FORM_sxGet() As String

                                            RECORD_SOURCE_FORM_sxGet = m_FRM_DATI_01_sxOrigineRecord_RecordSource

                        End Property
                    '........................................................................




                    '3) FIELD_NUMERATI
                    'valore del campo
                    '........................................................................
                    'Nota: Recupero il valore della proprieta VISIBLE, che può essere
                    'impostata a True , oppure a False, ed in questo caso i campi vengono
                    'numerati o meno.



                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let Field_Numerati_bxLet(par_Field_Numerati)

                                            m_bxField_Numerati = par_Field_Numerati

                        End Property

                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get Field_Numerati_bxGet() As String

                                            Field_Numerati_bxGet = m_bxField_Numerati

                        End Property
                    '........................................................................



                    '4) QUERY PARAMETRI DI AGGIORNAMENTO E NUMERO PARAMETRI
                    'QuerySalvataggio_Parametri_Tmp
                    'valore del campo
                    '........................................................................
                    'Nota: Recupero il valore del campo VALOREPROPRIETA, nel quale viene recuperata
                    'una stringa ssql da utilizzare per l'aggiornamento della tabella parametri all'attivazione
                    'dell'evento ACTIVATE relativa al campo 1 della form.
                    'La stringa sql è una query parametrica, che può avere un numero di parametro diversi (da 1 a 10)
                    'e con diversi tipi (es. integer, long, String ecc..). La successiva procedura, la numero
                    '5),6), 7), 8),9) e 10), recupera i parametri necessari per il funzionamento della query di
                    'aggiornamento.

                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let QuerySalvataggio_Parametri_Tmp_sxLET(par_QuerySalvataggio_Parametri_Tmp)

                                            m_sxQuerySalvataggio_Parametri_Tmp = par_QuerySalvataggio_Parametri_Tmp

                        End Property

                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get QuerySalvataggio_Parametri_Tmp_sxGET() As String

                                            QuerySalvataggio_Parametri_Tmp_sxGET = m_sxQuerySalvataggio_Parametri_Tmp

                        End Property


                                'NUMERO PARAMETRI PER LA QUERY DI AGGIORNAMENTO
                                'Note   : Il numero parametri necessari per la query di aggiornamento
                                'rappresentano i campi Field_01, Field_02 Field_0N... che sono necessari
                                'per eseguire la query parametrica (es. Query con 1 parametro, campo
                                '1 e Field_01 ecc..).

                                'LET
                                'Assegno il valore alla variabile privata dal chiamante
                                Public Property Let Nro_Parametri_ixLET(par_Nro_Parametri)

                                                    m_ixNro_Parametri = par_Nro_Parametri

                                End Property

                                'GET
                                'Restituisco la variabile con la forma da aprire
                                Public Property Get Nro_Parametri_ixGET() As String

                                                    Nro_Parametri_ixGET = m_ixNro_Parametri

                                End Property

                    '........................................................................



                    '5) TIPO PARAMETRI 01-05 DELLA QUERY DI AGGIORNAMENTO
                    'Parametro per il funzionamento della Query di aggiornamento
                    '-
                    '........................................................................
                    'Nota: Recupero il valore del campo VALOREPROPRIETA, dal quale viene recuperata
                    'un dato che indica il TIPO DI parametro da utilizzare nella query di aggiornamento
                    'al punto 4) (es. recupera il dato = Integer del campo Field_01, 1 PARAMETRO da utilizzare
                    'nella query di aggiornamento parametrica).


                        'TIPO PARAMETRO_01
                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let vxTIPO_PARAMETRO_01_Tmp_sxLET(par_vxTIPO_PARAMETRO_01_Tmp)

                                            m_vxTIPO_PARAMETRO_01_Tmp = par_vxTIPO_PARAMETRO_01_Tmp

                        End Property

                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get vxTIPO_PARAMETRO_01_Tmp_sxGET() As String

                                            vxTIPO_PARAMETRO_01_Tmp_sxGET = m_vxTIPO_PARAMETRO_01_Tmp

                        End Property

                        'TIPO PARAMETRO_02
                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let vxTIPO_PARAMETRO_02_Tmp_sxLET(par_vxTIPO_PARAMETRO_02_Tmp)

                                            m_vxTIPO_PARAMETRO_02_Tmp = par_vxTIPO_PARAMETRO_02_Tmp

                        End Property

                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get vxTIPO_PARAMETRO_02_Tmp_sxGET() As String

                                            vxTIPO_PARAMETRO_02_Tmp_sxGET = m_vxTIPO_PARAMETRO_02_Tmp

                        End Property

                        'TIPO PARAMETRO_03
                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let vxTIPO_PARAMETRO_03_Tmp_sxLET(par_vxTIPO_PARAMETRO_03_Tmp)

                                            m_vxTIPO_PARAMETRO_03_Tmp = par_vxTIPO_PARAMETRO_03_Tmp

                        End Property

                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get vxTIPO_PARAMETRO_03_Tmp_sxGET() As String

                                            vxTIPO_PARAMETRO_03_Tmp_sxGET = m_vxTIPO_PARAMETRO_03_Tmp

                        End Property

                        'TIPO PARAMETRO_04
                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let vxTIPO_PARAMETRO_04_Tmp_sxLET(par_vxTIPO_PARAMETRO_04_Tmp)

                                            m_vxTIPO_PARAMETRO_04_Tmp = par_vxTIPO_PARAMETRO_04_Tmp

                        End Property

                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get vxTIPO_PARAMETRO_04_Tmp_sxGET() As String

                                            vxTIPO_PARAMETRO_04_Tmp_sxGET = m_vxTIPO_PARAMETRO_04_Tmp

                        End Property



                        'TIPO PARAMETRO_05
                        'LET
                        'Assegno il valore alla variabile privata dal chiamante
                        Public Property Let vxTIPO_PARAMETRO_05_Tmp_sxLET(par_vxTIPO_PARAMETRO_05_Tmp)

                                            m_vxTIPO_PARAMETRO_05_Tmp = par_vxTIPO_PARAMETRO_05_Tmp

                        End Property

                        'GET
                        'Restituisco la variabile con la forma da aprire
                        Public Property Get vxTIPO_PARAMETRO_05_Tmp_sxGET() As String

                                            vxTIPO_PARAMETRO_05_Tmp_sxGET = m_vxTIPO_PARAMETRO_05_Tmp

                        End Property


                    '........................................................................




        '*** fine ***
        '1)I CAMPI GENERALI DEL RS
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<











            'INIZIO - PROCEDURA E FUNZIONI
            ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            'Nota : Funzione di estrazione dei comandi Form, vengono estratti i nomi
            'dei comandi, e quando viene indicato il metodo Open viene estratta dal
            'campo valore della proprieta, il nome della form che deve essere aperta
            'con il comando.
            'Function pubblica
            'PARAMETRI IN IMPUT  :  - par_sxCODOGGETTO
            '                       - par_sxTIPOGGETTO
            '                       - par_sxPROPRIETA
            '                       - par_vxVALOREPROPRIETA
            '                       - par_sxEVENTO
            '                       - par_sxATTIVITAEVENTO
            '                       - par_sxMETODO
            '                       - par_sxCODATTIVITA
            '                       - par_ixNROGGETTO
            '                       - par_lngxID

            '
            '
            Public Function RECUPERO_PARAMETRI_TABELLA_OGGETTI_pFc(par_sxCODOGGETTO As String, _
                                                                par_sxTIPOGGETTO As String, _
                                                                par_sxPROPRIETA As String, _
                                                                par_vxVALOREPROPRIETA As Variant, _
                                                                par_sxEVENTO As String, _
                                                                par_sxATTIVITAEVENTO As String, _
                                                                par_sxMETODO As String, _
                                                                par_sxCODATTIVITA As String, _
                                                                par_ixNROGGETTO As Integer, _
                                                                par_lngXID As Long)






            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            'Variabili di database

            'DAO
            Dim DaoDB As DAO.Database          '//occorre il riferimento alla libreria  Microsoft DAO 3.6 Object Libray
            Dim DaoWks As DAO.Workspace
            Dim DaoRs As DAO.Recordset

            'ADO
            Dim ADODB As Database
            Dim AdodaoRs As Recordset

            Dim sSql As String                                          '//Stringa di estrazione SQL

            'Variabili generali
            Dim str1 As String
            Dim Int1 As Integer
            Dim Lng1 As Long
            Dim Dbl1 As Double
            Dim Bln1 As Boolean
            Dim Vv1 As Variant

            'dati terreni
            Dim lngx_IMREDAGR As Long                                   'importo reddito agrario
            Dim lngx_IMPAGR As Long                                     'importo reddito agrario totale

            Dim lngx_IMREDDOM As Long                                   'importo reddito domenicale
            Dim lngx_IMPDOM As Long                                     'importo reddito domenicale totale
            Dim icount As Integer

            Dim lngx_RTPOSSESSO As Long
            Dim ix_GGPOSSESSO As Integer
            Dim sx_FLCASIPART As String                                 'CASI PARTICOLARI (col.7)
            Dim sx_FLTITOLO As String                                   'titolo
            Dim lngx_IMCANOAFF As Long                                   'canone di affitto



            'NOME QUERY
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++----

            'Carica Nome Query

            Dim m_sxNomeQuery_01 As String
            Dim m_sxNomeQuery_02 As String
            Dim m_sxNomeQuery_03 As String
            Dim m_sxNomeQuery_04 As String
            Dim m_sxNomeQuery_05 As String
            Dim m_sxNomeQuery_06 As String
            Dim m_sxNomeQuery_07 As String
            Dim m_sxNomeQuery_08 As String
            Dim m_sxNomeQuery_09 As String
            Dim m_sxNomeQuery_10 As String
            '
            '



            'LE VARIABILI PRIVATE  ESTERNE DELLA CLASSE
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>:>>>>>>>>>>>>>>>>>>>>>>>>>>>>

            'GLI OGGETTI COMANDI FORM
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++----
            'Oggetti generali della form
            'Variabili interne della classe

            Dim m_sxCODOGGETTO As String
            Dim m_sxNOMEOGGETTO As String
            Dim m_sxTIPOGGETTO As String
            Dim m_sxPROPRIETA As String
            Dim m_vxVALOREPROPRIETA As Variant
            Dim m_sxEVENTO    As String
            Dim m_sxATTIVITAEVENTO    As String
            Dim m_sxMETODO    As String
            Dim m_sxCODATTIVITA As String
            Dim m_ixNROGGETTO   As Integer
            Dim m_lngxIDOGGETTO   As Long

            'Altri parametri
            Dim m_sxTABLE   As String
            Dim m_sxQUERY   As String
            Dim m_sxCARICO_DATI_FORM As String
            Dim m_bxField_Numerati As Boolean
            Dim m_sxQuerySalvataggio_Parametri_Tmp As String

            Dim m_ixNro_Parametri As Integer                            'Numero parametri per la query di aggiornamento
            Dim m_vxTIPO_PARAMETRO_01_Tmp As Variant                    'Tipo parametro 01 da utilizzare
            Dim m_vxTIPO_PARAMETRO_02_Tmp As Variant                    'Tipo parametro 02 da utilizzare
            Dim m_vxTIPO_PARAMETRO_03_Tmp As Variant                    'Tipo parametro 03 da utilizzare
            Dim m_vxTIPO_PARAMETRO_04_Tmp As Variant                    'Tipo parametro 04 da utilizzare
            Dim m_vxTIPO_PARAMETRO_05_Tmp As Variant                    'Tipo parametro 05 da utilizzare


            '
            '


            'PROPRIETA DELLA FORM
            '--------------------------------------------------------------------

            'A)FORMATO - FORM
            'Proprieta FORM
            Dim m_FRM_FORMATO_01_sxEtichetta_Caption As String
            Dim m_FRM_FORMATO_02_bxVisualizzazionePredefinita_DefaultView As Boolean
            Dim m_FRM_FORMATO_03_sxVisualizzazioneConsentite_ViewsAllowed As String
            Dim m_FRM_FORMATO_04_sxBarreScorrimento_ScrollBars As String
            Dim m_FRM_FORMATO_05_bxSelettoriRecord_RecordSelectors As Boolean
            Dim m_FRM_FORMATO_06_bxPulsantiSpostamento_NavigationButtons As Boolean
            Dim m_FRM_FORMATO_07_bxLineeDiDivisione_DividingLines As Boolean
            Dim m_FRM_FORMATO_08_bxDimensioniAutomatiche_AutoResize As Boolean
            Dim m_FRM_FORMATO_09_bxCentraturaAutomatica_AutoCenter As Boolean
            Dim m_FRM_FORMATO_10_bxDimensionabile_BorderStile As Boolean
            Dim m_FRM_FORMATO_11_bxCasellaMenuControllo_ControlBox As Boolean
            Dim m_FRM_FORMATO_12_ixPulsantiIngrandRiduz_MinMaxButtons As Integer
            Dim m_FRM_FORMATO_13_bxPulsanteChiusura_CloseButton As Boolean
            Dim m_FRM_FORMATO_14_bxPulsanteGuidaRapida_WhatsThisButton As Boolean
            Dim m_FRM_FORMATO_15_dxLarghezza_Width As Double
            Dim m_lxCmTwip As Long                                                            'Unita di misura 1 cm = 567 twip
            Dim m_FRM_FORMATO_16_sxImmagine_Picture As String


            'B)DATI - FORM
            'Proprieta FORM
            Dim m_FRM_DATI_01_sxOrigineRecord_RecordSource As String
            Dim m_FRM_DATI_02_sxFiltro_Filter As String
            Dim m_FRM_DATI_03_sxOrdinaPer_OrderBy As String
            Dim m_FRM_DATI_04_bxConsentiFiltri_AllowFilters As Boolean
            Dim m_FRM_DATI_05_bxConsentiModifiche_AllowEdits  As Boolean
            Dim m_FRM_DATI_06_bxConsentiEliminazioni_AllowDeletetion  As Boolean
            Dim m_FRM_DATI_07_bxConsentiAggiunte_AllowAddition  As Boolean
            Dim m_FRM_DATI_08_bxImmissioneDati_DataEntry  As Boolean
            Dim m_FRM_DATI_09_bxTipoRecordset_RecordsetType  As Boolean
            Dim m_FRM_DATI_10_bxTipoBloccoRecord_RecordLocks  As Boolean








            On Error GoTo Err_RECUPERO_PARAMETRI_TABELLA_OGGETTI_pFc


                            'APRO IL RS
                            sSql = ""

                            sSql = sSql & "SELECT "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.CODOGGETTO_s, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.NOMEOGGETTO,"
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.NOTA_OGGETTO, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.TIPOGGETTO, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.PROPRIETA, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.VALOREPROPRIETA, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.EVENTO, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.ATTIVITAEVENTO, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.METODO, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.MODFORM, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.CODATTIVITA, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.NROGGETTO, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.DATAINS_d, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.DATAAGG_d, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.TIMEOPER_d, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.IDOGGETTO_lng "

                            sSql = sSql & "FROM GEST_Tb01_N01_PARAM_OGGETTI "
                            sSql = sSql & "WHERE (((GEST_Tb01_N01_PARAM_OGGETTI.CODOGGETTO_s)='" & par_sxCODOGGETTO & "')"
                            sSql = sSql & "OR ((GEST_Tb01_N01_PARAM_OGGETTI.TIPOGGETTO)='" & par_sxTIPOGGETTO & "')"
                            sSql = sSql & "OR ((GEST_Tb01_N01_PARAM_OGGETTI.ATTIVITAEVENTO)='" & par_sxPROPRIETA & "'))"
                            'sSql = sSql & "ORDER BY GEST_Tb01_N01_PARAM_OGGETTI.NROGGETTO"
                            sSql = sSql & "ORDER BY "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.NROGGETTO, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.TIPOGGETTO, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.NROGGETTO, "
                            sSql = sSql & "GEST_Tb01_N01_PARAM_OGGETTI.NOMEOGGETTO;"

                            'OLD
                            'sSql = sSql & "FROM GEST_Tb01_N01_PARAM_OGGETTI "
                            'sSql = sSql & "WHERE (((GEST_Tb01_N01_PARAM_OGGETTI.CODOGGETTO_s)='" & par_sxCODOGGETTO & "')"
                            'sSql = sSql & "or ((GEST_Tb01_N01_PARAM_OGGETTI.TIPOGGETTO)='" & par_sxTIPOGGETTO & "')"
                            'sSql = sSql & "or ((GEST_Tb01_N01_PARAM_OGGETTI.METODO)='" & par_sxMETODO & "')"
                            'sSql = sSql & "or ((GEST_Tb01_N01_PARAM_OGGETTI.PROPRIETA)='" & par_sxPROPRIETA & "'))"
                            'sSql = sSql & "ORDER BY GEST_Tb01_N01_PARAM_OGGETTI.NROGGETTO"

                               Debug.Print sSql



                            'Apro rs estrazione Sql
                            Set DaoRs = CurrentDb.OpenRecordset(sSql)
                                'DaoRs.MoveFirst

                                With DaoRs

                                    'CICLO RS
                                    '..................................................................................................
                                        While Not DaoRs.EOF

                                            'reset
                                            m_sxPROPRIETA = ""
                                            m_sxCODATTIVITA = ""


                                                'CAMPI DEL RS
                                                '_________________________________________________________________
                                                'Note: Vengono recuperati i valori dei campi del rs.

                                                        '1)?? CODICE
                                                            'm_sx???
                                                        '2).....???
                                                            'm_sx....???
                                                        '5)CODATTIVITA
                                                            'm_sxCODATTIVITA

                                                    'VALORE RECUPERATO
                                                    '..........................................................

                                                            'Controllo
                                                            Debug.Print "0) : " & .Fields("IDOGGETTO").Name & " Valore : " & .Fields("IDOGGETTO")


                                                        '1)CODOGGETTO
                                                            'recupero il valore del campo
                                                            If .Fields("CODOGGETTO_s") > 0 Then
                                                                m_sxCODOGGETTO = .Fields("CODOGGETTO_s")

                                                                    'Controllo
                                                                    Debug.Print "1) : " & .Fields("CODOGGETTO_s").Name & " Valore : " & .Fields("CODOGGETTO_s")


                                                            End If

                                                        '2)TIPOGGETTO
                                                            'recupero il valore del campo
                                                            If .Fields("TIPOGGETTO") > 0 Then
                                                                m_sxTIPOGGETTO = .Fields("TIPOGGETTO")

                                                                    'Controllo
                                                                    Debug.Print "2) : " & .Fields("TIPOGGETTO").Name & " Valore : " & .Fields("TIPOGGETTO")


                                                            End If

                                                        '3)PROPRIETA
                                                            'recupero il valore del campo

                                                            If .Fields("PROPRIETA") > 0 Then

                                                                m_sxPROPRIETA = .Fields("PROPRIETA")

                                                                    'Controllo
                                                                    Debug.Print "3) : " & .Fields("PROPRIETA").Name & " Valore : " & .Fields("PROPRIETA")


                                                            End If

                                                        '4) VALOREPROPRIETA
                                                            'recupero il valore del campo
                                                            If .Fields("VALOREPROPRIETA") > 0 Then

                                                                m_vxVALOREPROPRIETA = .Fields("VALOREPROPRIETA")

                                                                    'Controllo
                                                                    Debug.Print "4) : " & .Fields("VALOREPROPRIETA").Name & " Valore : " & .Fields("VALOREPROPRIETA")



                                                            End If

                                                        '5)EVENTO
                                                            'recupero il valore del campo
                                                            If .Fields("EVENTO") > 0 Then

                                                                m_sxEVENTO = .Fields("EVENTO")

                                                                        'Controllo
                                                                        Debug.Print "5) : " & .Fields("EVENTO").Name & " Valore : " & .Fields("EVENTO")



                                                            End If

                                                        '8)CODATTIVITA
                                                            'recupero il valore del campo
                                                            If .Fields("CODATTIVITA") > 0 Then

                                                                m_sxCODATTIVITA = .Fields("CODATTIVITA")

                                                                        'Controllo
                                                                        Debug.Print "8) : " & .Fields("CODATTIVITA").Name & " Valore : " & .Fields("CODATTIVITA")



                                                            End If

                                                    '..........................................................


                                                    'TABELLA DI SALVATAGGIO DEI PARAMETRI
                                                    '..........................................................
                                                    '
                                                            '"QuerySalvataggio_Parametri_Tmp"

                                                            '1)RECUPERO LA TABELLA DI SALVATAGGIO

                                                            'TABLE
                                                            'PER LA TABELLA DI SALVATAGGIO INDIVIDUO
                                                            'Invidividuo il tipo oggetto <<Table>>
                                                            'ed il codice del campo ATTIVITA EVENTO = <<QuerySalvataggio_Parametri_Tmp>> con EVENTO = Null
                                                            If .Fields("TIPOGGETTO") = "Table" _
                                                                And .Fields("ATTIVITAEVENTO") = "QuerySalvataggio_Parametri_Tmp" _
                                                                And IsNull(.Fields("EVENTO")) = True Then
                                                                    'Valore proprieta
                                                                    m_sxTABLE = .Fields("NOMEOGGETTO")

                                                                        'Controllo
                                                                        Debug.Print "8) : " & .Fields("CODOGGETTO_s").Name & " Valore : " & .Fields("CODOGGETTO_s")



                                                            End If



                                                            '1) VISUALIZZA I DATI NELLA FORM
                                                            '...............................................................................
                                                                'PARAMETRO PER LA VISUALIZZAZIONE DEI DATI
                                                                'VALORE PROPRIETA DELLA QUERY
                                                                'ricerco il tipo oggetto <<Query>>
                                                                'ed il codice del campo ATTIVITA EVENTO = <<Carico_Dati_Form>> con EVENTO = Null
                                                                If .Fields("TIPOGGETTO") = "Query" _
                                                                    And .Fields("ATTIVITAEVENTO") = "Carico_Dati_Form" _
                                                                    And IsNull(.Fields("EVENTO")) = True Then

                                                                        'Nota : viene prelevato il nome oggetto, in quanto la proprieta recordsource
                                                                        'della form, può contenere al massimo una stringa di 255 caratteri. Quindi si
                                                                        'preferisce utilizzare il nome della query anziche il suo contegno in sql, inquanto
                                                                        'potrebbe superare tale soglia, causando errori di caricamento.

                                                                        'Nome oggetto
                                                                        m_FRM_DATI_01_sxOrigineRecord_RecordSource = .Fields("NOMEOGGETTO")

                                                                End If


                                                                'PARAMETRO PER LA VISUALIZZAZIONE DELLA NUMERAZIONE DEI CAMPI
                                                                'ricerco il tipo oggetto <<Field_Numerati>>
                                                                'ed il codice del campo TIPO OGGETTO = <<Field_Numerati>> con EVENTO = Null
                                                                If .Fields("TIPOGGETTO") = "Field_Numerati" _
                                                                    And .Fields("PROPRIETA") = "Visible" _
                                                                    And IsNull(.Fields("EVENTO")) = True Then
                                                                        'Valore proprieta
                                                                        m_bxField_Numerati = .Fields("VALOREPROPRIETA")

                                                                End If

                                                            '...............................................................................



                                                            '2)RECUPERO DELLA QUERY DA UTILIZZARE PER IL SALVATAGGIO
                                                            '...............................................................................

                                                            'QUERY
                                                            'PER LA TABELLA DI SALVATAGGIO INDIVIDUO
                                                            'Invidividuo il tipo oggetto <<Query>>
                                                            'ed il codice del campo ATTIVITA EVENTO = <<QuerySalvataggio_Parametri_Tmp>> con EVENTO = Null
                                                            If .Fields("TIPOGGETTO") = "Query" _
                                                                And .Fields("ATTIVITAEVENTO") = "QuerySalvataggio_Parametri_Tmp" _
                                                                And IsNull(.Fields("EVENTO")) = True Then

                                                                    'Valore proprieta
                                                                    m_sxQUERY = .Fields("NOMEOGGETTO")



                                                            End If

                                                            'PARAMETRO DI SALVATAGGIO
                                                            'VALORE PROPRIETA DELLA QUERY
                                                            'PER LA TABELLA DI SALVATAGGIO
                                                            'ricerco il tipo oggetto <<Query>>
                                                            'ed il codice del campo ATTIVITA EVENTO = <<QuerySalvataggio_Parametri_Tmp>> con EVENTO = Null
                                                            If .Fields("TIPOGGETTO") = "Query" _
                                                                And .Fields("ATTIVITAEVENTO") = "QuerySalvataggio_Parametri_Tmp" _
                                                                And IsNull(.Fields("EVENTO")) = True Then

                                                                        'Controllo
                                                                        '............................................................................................
                                                                            Debug.Print "<<QuerySalvataggio_Parametri_Tmp>> : " & .Fields("VALOREPROPRIETA").Name & " Valore : " & .Fields("VALOREPROPRIETA")
                                                                        '............................................................................................


                                                                    'Valore proprieta
                                                                    m_sxQuerySalvataggio_Parametri_Tmp = .Fields("VALOREPROPRIETA")

                                                            End If

                                                                'TIPO E NUMERO DI PARAMETRI DELLA QUERY PARAMETRICA
                                                                '-----------------------------------------------------------------------------------
                                                                'Nota   : Tipo di parametri recuperati da 1 a 5




                                                                    'NUMERO DI PARAMETRI
                                                                    'ricerco il tipo oggetto <<Nro_Field>>, per individuare il parametro
                                                                    'il valore che indica il numero di parametri utilizzati dalla query parametrica
                                                                    'Il codice del campo EVENTO = Null
                                                                    If .Fields("TIPOGGETTO") = "Nro_Field" _
                                                                         And .Fields("PROPRIETA") = "Nro_Parametri" _
                                                                         And IsNull(.Fields("EVENTO")) = True Then

                                                                                    'Controllo
                                                                                    '............................................................................................
                                                                                        Debug.Print "<<NUMERO PARAMETRI>> " & .Fields("VALOREPROPRIETA").Name & " Valore : " & .Fields("VALOREPROPRIETA")
                                                                                    '............................................................................................


                                                                             'Valore proprieta
                                                                             m_ixNro_Parametri = .Fields("VALOREPROPRIETA")

                                                                     End If



                                                                       'TIPO PARAMETRO_01
                                                                       'ricerco il tipo oggetto <<Field_01>>, per individuare il parametro
                                                                       'e cotenstualmente il tipo di parametro utilizzato (es. parametro 1 = integer, oppure
                                                                       'parametro 1 = String ecc.)
                                                                       'Il codice del campo EVENTO = Null
                                                                       If .Fields("TIPOGGETTO") = "Field_01" _
                                                                           And .Fields("PROPRIETA") = "Tipo" _
                                                                           And IsNull(.Fields("EVENTO")) = True Then


                                                                                    'Controllo
                                                                                    '............................................................................................
                                                                                        Debug.Print "TIPO PARAMETRO 1) : " & .Fields("VALOREPROPRIETA").Name & " Valore : " & .Fields("VALOREPROPRIETA")
                                                                                    '............................................................................................


                                                                               'Valore proprieta
                                                                               m_vxTIPO_PARAMETRO_01_Tmp = .Fields("VALOREPROPRIETA")

                                                                       End If


                                                                        'TIPO PARAMETRO_02
                                                                        'ricerco il tipo oggetto <<Field_02>>, per individuare il parametro
                                                                        'e cotenstualmente il tipo di parametro utilizzato (es. parametro 1 = integer, oppure
                                                                        'parametro 1 = String ecc.)
                                                                        'Il codice del campo EVENTO = Null
                                                                        If .Fields("TIPOGGETTO") = "Field_02" _
                                                                            And .Fields("PROPRIETA") = "Tipo" _
                                                                            And IsNull(.Fields("EVENTO")) = True Then


                                                                                    'Controllo
                                                                                    '............................................................................................
                                                                                        Debug.Print "TIPO PARAMETRO 2) : " & .Fields("VALOREPROPRIETA").Name & " Valore : " & .Fields("VALOREPROPRIETA")
                                                                                    '............................................................................................

                                                                               'Valore proprieta
                                                                               m_vxTIPO_PARAMETRO_02_Tmp = .Fields("VALOREPROPRIETA")

                                                                        End If


                                                                        'TIPO PARAMETRO_03
                                                                        'ricerco il tipo oggetto <<Field_03>>, per individuare il parametro
                                                                        'e cotenstualmente il tipo di parametro utilizzato (es. parametro 1 = integer, oppure
                                                                        'parametro 1 = String ecc.)
                                                                        'Il codice del campo EVENTO = Null
                                                                        If .Fields("TIPOGGETTO") = "Field_03" _
                                                                            And .Fields("PROPRIETA") = "Tipo" _
                                                                            And IsNull(.Fields("EVENTO")) = True Then



                                                                                    'Controllo
                                                                                    '............................................................................................
                                                                                        Debug.Print "TIPO PARAMETRO 3) : " & .Fields("VALOREPROPRIETA").Name & " Valore : " & .Fields("VALOREPROPRIETA")
                                                                                    '............................................................................................

                                                                               'Valore proprieta
                                                                               m_vxTIPO_PARAMETRO_03_Tmp = .Fields("VALOREPROPRIETA")

                                                                        End If


                                                                        'TIPO PARAMETRO_03
                                                                        'ricerco il tipo oggetto <<Field_03>>, per individuare il parametro
                                                                        'e cotenstualmente il tipo di parametro utilizzato (es. parametro 1 = integer, oppure
                                                                        'parametro 1 = String ecc.)
                                                                        'Il codice del campo EVENTO = Null
                                                                        If .Fields("TIPOGGETTO") = "Field_03" _
                                                                            And .Fields("PROPRIETA") = "Tipo" _
                                                                            And IsNull(.Fields("EVENTO")) = True Then



                                                                                    'Controllo
                                                                                    '............................................................................................
                                                                                        Debug.Print "TIPO PARAMETRO 3) : " & .Fields("VALOREPROPRIETA").Name & " Valore : " & .Fields("VALOREPROPRIETA")
                                                                                    '............................................................................................

                                                                               'Valore proprieta
                                                                               m_vxTIPO_PARAMETRO_03_Tmp = .Fields("VALOREPROPRIETA")

                                                                        End If



                                                                        'TIPO PARAMETRO_04
                                                                        'ricerco il tipo oggetto <<Field_04>>, per individuare il parametro
                                                                        'e cotenstualmente il tipo di parametro utilizzato (es. parametro 1 = integer, oppure
                                                                        'parametro 1 = String ecc.)
                                                                        'Il codice del campo EVENTO = Null
                                                                        If .Fields("TIPOGGETTO") = "Field_04" _
                                                                            And .Fields("PROPRIETA") = "Tipo" _
                                                                            And IsNull(.Fields("EVENTO")) = True Then


                                                                                    'Controllo
                                                                                    '............................................................................................
                                                                                        Debug.Print "TIPO PARAMETRO 4) : " & .Fields("VALOREPROPRIETA").Name & " Valore : " & .Fields("VALOREPROPRIETA")
                                                                                    '............................................................................................

                                                                               'Valore proprieta
                                                                               m_vxTIPO_PARAMETRO_04_Tmp = .Fields("VALOREPROPRIETA")

                                                                        End If

                                                                        'TIPO PARAMETRO_05
                                                                        'ricerco il tipo oggetto <<Field_05>>, per individuare il parametro
                                                                        'e cotenstualmente il tipo di parametro utilizzato (es. parametro 1 = integer, oppure
                                                                        'parametro 1 = String ecc.)
                                                                        'Il codice del campo EVENTO = Null
                                                                        If .Fields("TIPOGGETTO") = "Field_05" _
                                                                            And .Fields("PROPRIETA") = "Tipo" _
                                                                            And IsNull(.Fields("EVENTO")) = True Then

                                                                                    'Controllo
                                                                                    '............................................................................................
                                                                                        Debug.Print "TIPO PARAMETRO 5) : " & .Fields("VALOREPROPRIETA").Name & " Valore : " & .Fields("VALOREPROPRIETA")
                                                                                    '............................................................................................

                                                                               'Valore proprieta
                                                                               m_vxTIPO_PARAMETRO_05_Tmp = .Fields("VALOREPROPRIETA")

                                                                        End If


                                                                '*** FINE ****
                                                                'TIPO DI PARAMETRI DELLA QUERY PARAMETRICA
                                                                '-----------------------------------------------------------------------------------


                                                            '*** FINE ***
                                                            '2)RECUPERO DELLA QUERY DA UTILIZZARE PER IL SALVATAGGIO
                                                            '.......................................................................



                                                '*** FINE ***
                                                'CAMPI DEL RS
                                                '_________________________________________________________________





                                            .MoveNext
                                        Wend

                                    'FINE CICLO RS
                                    '..................................................................................................



                                End With

                            'Chiudo rs estrazione Sql
                            DaoRs.Close
                            Set DaoRs = Nothing




            'USCITA ED ERRORI
            '..............................................................
            Exit_RECUPERO_PARAMETRI_TABELLA_OGGETTI_pFc:
                            Exit Function

            Err_RECUPERO_PARAMETRI_TABELLA_OGGETTI_pFc:
                            MsgBox Err.Description
                            Resume Exit_RECUPERO_PARAMETRI_TABELLA_OGGETTI_pFc



            End Function
        '//###########################################################################################################################//

'//OGGETTI_ITERA_TUTTE_LE_FORM_COLLEGA_MSys_TABELLE_LINK_PFunct                     COLLEGA LE TABELLE   **** INIZIO ****
            '//=======================================================================================================================================================//
            '//FUNZIONAMENTO DEL MODULO COLLEGA LE TABELLE: _
                OGGETTI SYS _
                    Per il funzionamento di questo MODULO che permette il collegamento delle tabelle occorre prelevare dal database _
                    classe form di gestione del salario acess _
                        Occorrono le tabelle di sistema     :        MSys{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}Tab _
                                                                     MSysDF00}----------------------------------------------------@ _
                                                                     MSysDF01COMANDI _
                                                                     MSysELENCO}-------------------------------------------------@ _
                                                                     MSysELENCOOGGETTITABELLE _
                        Le seguenti query di sistema        :        MSsysQry0180}-------------------------------@DELETE _
                                                                     MSsys_Qry01_81_MSys_DELETE_ELENCO_OGGETTI_TABELLE_Tutte _
                FUNZIONAMENTO _
                    Per il funzionamento di questo modulo occorre valorizzare la tabella degli oggetti con il nome della tabelle _
                    da collegare, indicando anche i disco, la path ed il database da collegare. IL parametro Cancella_s _
                    deve essere impostato  TRUE se si vuole cancellare LE TABELLE COLLEGATE.




            '//COLLEGO LA TABELLA DIPENDENTI @collega.dipendenti _
                Note: la tabella dipendenti viene collegata chiamando questa funzione impostata _
                cha ha bisogno di 5 parametri, il : _
                                disco                           = il disco Dove si trova il db da collegare, c:, y: f: ecc _
                                la path                         = la path relativa del db da collegare _
                                il database                     = il nome del db da collegare _
                                la tabella                      = la tabella da collegare _
                                il flag cancella                = il flag TRUE per cancellare la tabella gia collegata evitando i duplicato, False = se la tabella non esiste e non deve essere cancellata _
                I primi 4 sono parametri generali mentre il 5 la tabella deve essere specifico per le chiamate singole.


            '//COLLEGA LE TABELLE _
                il collegamento deve essere costruito con l'iterazione del database corrente, ispezionando la _
                tabella MSys_TABELLE_LINK ed attivando i collegamento con i dati dei record esistenti.


            '//TODO: QUESTO MODULO deve essere incluso in unica classe gestion_Subi



            Private Sub OGGETTI_ITERA_TUTTE_LE_FORM_COLLEGA_MSys_TABELLE_LINK_PFunct()
                Call COLLEGA_MSys_TABELLE_LINK_PFunct
            End Sub

            Public Function COLLEGA_MSys_TABELLE_LINK_PFunct()



                '//todo aprire un rs.

                '//PARAMETRI GENERALI
                DATABASE_s = ""
                Disco_s = ""
                PATH_s = ""
                Table_s = ""
                Cancella_b = True

                    '//ITERO NEL RS
                    '//--------------------------------------------------------------------------------//-------//

                         '//Inserire tabella o stringa ssql
                         sSql = "MSys_TABELLE_LINK"
                         'RS ADO
                         'Solo l'anno indicato nella variabile
                         Set AdodaoRs = CurrentDb.OpenRecordset(sSql)

                        '//WITH AdodaoRs
                        '//............................................................................//
                         With AdodaoRs

                             'Controllo Rs
                             If AdodaoRs.EOF = False And AdodaoRs.BOF = False Then

                                 'RESET
                                 Int1 = 0
                                 icount = 0
                                 iTOTcount = 0
                                 ixNRO_ESTR = 0

                                 .MoveFirst

                                 'ITERAZIONE RS DI LETTURA
                                 '----------------------------------------------------------------------
                                     While Not AdodaoRs.EOF

                                     'Controllo windows
                                     DoEvents
                                     sSql = ""


                                        '//PARAMETRI COLLETTIVI PER TUTTE LE CHIAMATE

                                            Disco_s = .Fields("Disco_s")
                                            PATH_s = .Fields("Path_s")
                                            DATABASE_s = .Fields("Database_s")
                                            Table_s = .Fields("Table_s")
                                            Cancella_b = .Fields("Cancella_b")

                                        'Vado avanti finche la base è uguale
                                         '------------------------------------------------
                                             'Imposto i codici Base
                                             lngxID_TERN = .Fields("ID_lng")

                                             Debug.Print
                                             Debug.Print "--------------------------------------------"
                                             Debug.Print "CONTROLLO id tabella" & lngxID_TERN
                                             Debug.Print "collegato il seguente oggetto : " & Disco_s & PATH_s & DATABASE_s & Table_s & " cancella : " & Cancella_b

                                             Debug.Print "--------------------------------------------"
                                         '............................................

                                            ' Costruisce il percorso completo del database
                                            Dim fullPath As String
                                            fullPath = Disco_s & PATH_s & DATABASE_s

                                            ' Controlla se il percorso esiste
                                            If Dir(PATH_s, vbDirectory) = "" Or Dir(fullPath) = "" Then
                                                MsgBox "Database e path non esistenti   :  " & fullPath, vbCritical
                                                Exit Function
                                            End If


                                            '// Se il percorso esiste, chiama la funzione per collegare la tabella
                                            Vv1 = COLLEGA_TABELLA_pfunct(Disco_s, PATH_s, DATABASE_s, Table_s, Cancella_b)

                                         AdodaoRs.MoveNext

                                     Wend
                                 'ITERAZIONE RS DI LETTURA       *** fine ***
                                 '----------------------------------------------------------------------

                             End If '//If AdodaoRs.EOF = False And AdodaoRs.BOF = False Then

                                '// SELEZIONA L'ULTIMA TABELLA
                                '//======================================================================================================//




                                      '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DELL'OGGETTO_FORM
                                      '//CODICE-----> Controlla_SE_OGGETTO_FORM_ESISTE_pFunct.attiva.LaClasse
                                      '//---------------------------------------------------------------------------------//
                                      '//NOTA  :Controlla se esiste la form negli oggetti del database e restituisce TRUE _
                                                se il nome della form passata con parametro par_sxNameForm esiste; altrimenti FALSE



                                                  'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_FORM
                                                  'Assegno e creo l'oggetto di classe
                                                  '
                                                  Dim MyCls_CLASSE_GESTIONE_OGGETTI_FORM As CLASSE_GESTIONE_OGGETTI_DATABASE
                                                  Dim sxNameForm As String
                                                  '
                                                  Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = New CLASSE_GESTIONE_OGGETTI_DATABASE



                                                    '//NOTA: per vedere se esiste la TABLE invio come parametro il nome dell'oggetto _
                                                  TABLE DA RICERCARE e chiamo la funzione la quale restituisce un TRUE _
                                                  se esiste altrimenti un FALSE = oggetto non esiste nel database
                                                  str1 = par_Table_s


                                                  '
                                                  Boolean1 = MyCls_CLASSE_GESTIONE_OGGETTI_FORM.pfCONTROLLO_ESISTENZA_OGGETTI_TABLE(str1)

                                                  '
                                                  sxNameForm = vbNull
                                                  Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = Nothing

                                          '//---------------------------------------------------------------------------------//



                                    If Bool1 = True Then

                                        '//vado all'ultima tabella collegata e la seleziono
                                        DoCmd.SelectObject acTable, Table_s, True
                                    End If

                                '//======================================================================================================//



                         End With
                        '//WITH AdodaoRs     *** FINE ***
                        '//............................................................................//


            '//ITERO NEL RS     *** fine ***
            '//--------------------------------------------------------------------------------//-------//


            End Function



            '//                                         COLLEGA LE TABELLE   **** FINE ****
            '//=======================================================================================================================================================//

'//COLLEGA_TABELLA_pfunct   FUNZIONE_DI_COLLEGAMENTO_CON_PARAMETRI: *** INIZIO***
        '//======================================================================================================================//
        '//PARAMETRI: TODO: DA CREARE COME FUNZIONE GENERALE, per ora solo 2 parametri ma deve essere allargata anche al Flag = True per la cancellazione _
           alla tabella da collegare + e al database per renderlo piu generico possibile.
        Function COLLEGA_TABELLA_pfunct(par_Disco_s As String, _
                                                             par_Path_s As String, _
                                                             par_Database_s As String, _
                                                             par_Table_s As String, _
                                                             par_Cancella_b As Boolean)


        On Error GoTo COLLEGA_TABELLA_pfunct_Err




                  '//2024.08.07 = RICOSTRUITO IL CONTROLLO DI ESISTENZA DELLE TABELLE



                    '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DELL'OGGETTO_FORM
                    '//CODICE-----> Controlla_SE_OGGETTO_FORM_ESISTE_pFunct.attiva.LaClasse
                    '//---------------------------------------------------------------------------------//
                    '//NOTA  :Controlla se esiste la form negli oggetti del database e restituisce TRUE _
                              se il nome della form passata con parametro par_sxNameForm esiste; altrimenti FALSE



                                'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_FORM
                                'Assegno e creo l'oggetto di classe
                                '
                                Dim MyCls_CLASSE_GESTIONE_OGGETTI_FORM As CLASSE_GESTIONE_OGGETTI_DATABASE
                                Dim sxNameForm As String
                                '
                                Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = New CLASSE_GESTIONE_OGGETTI_DATABASE


                                str1 = par_Table_s

                                '
                                Boolean1 = MyCls_CLASSE_GESTIONE_OGGETTI_FORM.pfCONTROLLO_ESISTENZA_OGGETTI_TABLE(str1)

                                '
                                sxNameForm = vbNull
                                Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = Nothing

                        '//---------------------------------------------------------------------------------//





                    '//I° TIPO DI CANCELLAZIONE DELLA TABELLA SOLO SE ESISTE COME LINK _
                        prima di cancelleare controlla se la tabella esiste = TRUE
                    If Bool1 = True Then
                        '//se la tabella è esistente cancello
                        DoCmd.DeleteObject acTable, par_Table_s

                    End If


                    '//TODO: come impostare la cancellazione da record ?? da costruire ???
                    '//II° TIPO DI CANCELLAZIONE DELLA TABELLA SPECIFICATA SE SUL RECORD E' TRUE = CANCELLA _
                    questa cancellazione è ordinata dal Rs perchè è valorizzato a True il campo cancella.
                    If par_Cancella_b = True Then

                        'cancello solo se la ricerca ha trovato la tabella collegata
                        If Bool1 = True Then
                            DoCmd.DeleteObject acTable, par_Table_s
                        End If

                    End If

            ' COLLEGO
            ' collego la tabella es. DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N25_PIANTA_ORGANICA.mdb", acTable, "PRES3000_TB25_DIPENDENTI", "PRES3000_TB25_DIPENDENTI", False
            DoCmd.TransferDatabase acLink, "Microsoft Access", par_Disco_s & par_Path_s & par_Database_s, acTable, par_Table_s, par_Table_s, False

                    ' Azione macro dalla versione più recente di Access: SuErrore
                    ' prima elimina la tabella collegata gia esistente
                    'OLD DoCmd.DeleteObject acTable, "PRES3000_Tb25_}------------------------------------------------@"
                    ' collego la tabella PRES3000_Tb25_}------------------------------------------------@
                    'OLD DoCmd.TransferDatabase acLink, "Microsoft Access", "c:\GESTIONI\GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N25_PIANTA_ORGANICA.mdb", acTable, "PRES3000_Tb25_}------------------------------------------------@", "PRES3000_Tb25_}------------------------------------------------@", False

        COLLEGA_TABELLA_pfunct_Exit:
                                Exit Function

        COLLEGA_TABELLA_pfunct_Err:
                                MsgBox Error$
                                Resume COLLEGA_TABELLA_pfunct_Exit

    End Function

'//ATTIVA_CANCELLA_MSys_TABELLE_LINK_PFunct - FUNZIONE_DI_CANCELLAMENTO_LINK_TABELLE: *** INIZIO***
    '//======================================================================================================================//
    '//PARAMETRI: CANCELLO LE TABELLE SENGNALATE NEL RECORDSET CON IL Flag = True per la cancellazione _
    '//2024.08.07_@CANCELLO@TABELLE@LINK_(la funzione cancella le tabelle in precedenza collegate)


    Private Sub ATTIVA_CANCELLA_MSys_TABELLE_LINK_PFunct()
        Call CANCELLA_MSys_TABELLE_LINK_PFunct
    End Sub

    Public Function CANCELLA_MSys_TABELLE_LINK_PFunct()

        '//todo aprire un rs.

        '//PARAMETRI GENERALI
        DATABASE_s = ""
        Disco_s = ""
        PATH_s = ""
        Table_s = ""
        Cancella_b = True

            '//ITERO NEL RS
            '//--------------------------------------------------------------------------------//-------//

                 '//Inserire tabella o stringa ssql
                 sSql = "MSys_TABELLE_LINK"
                 'RS ADO
                 'Solo l'anno indicato nella variabile
                 Set AdodaoRs = CurrentDb.OpenRecordset(sSql)

                '//WITH AdodaoRs
                '//............................................................................//
                 With AdodaoRs

                     'Controllo Rs
                     If AdodaoRs.EOF = False And AdodaoRs.BOF = False Then

                         'RESET
                         Int1 = 0
                         icount = 0
                         iTOTcount = 0
                         ixNRO_ESTR = 0

                         .MoveFirst

                         'ITERAZIONE RS DI LETTURA
                         '----------------------------------------------------------------------
                             While Not AdodaoRs.EOF

                             'Controllo windows
                             DoEvents
                             sSql = ""


                                '//PARAMETRI COLLETTIVI PER TUTTE LE CHIAMATE
                                    Disco_s = .Fields("Disco_s")
                                    PATH_s = .Fields("Path_s")
                                    DATABASE_s = .Fields("Database_s")
                                    Table_s = .Fields("Table_s")
                                    Cancella_b = .Fields("Cancella_b")

                                'Vado avanti finche la base è uguale
                                 '------------------------------------------------
                                     'Imposto i codici Base
                                     lngxID_TERN = .Fields("ID_lng")

                                     Debug.Print
                                     Debug.Print "--------------------------------------------"
                                     Debug.Print "CONTROLLO id tabella" & lngxID_TERN
                                     Debug.Print "collegato il seguente oggetto : " & Disco_s & PATH_s & DATABASE_s & Table_s & " cancella : " & Cancella_b

                                     Debug.Print "--------------------------------------------"
                                 '............................................

                                    '//TRUE = la tabella esiste ed è stata cancellata - FALSE = tabella non esistente non è stata cancellata
                                    Vv1 = CANCELLA_TABELLE_LINK(Disco_s, PATH_s, DATABASE_s, Table_s, Cancella_b)

                                 AdodaoRs.MoveNext

                             Wend
                         'ITERAZIONE RS DI LETTURA       *** fine ***
                         '----------------------------------------------------------------------

                     End If '//If AdodaoRs.EOF = False And AdodaoRs.BOF = False Then



                 End With
                '//WITH AdodaoRs     *** FINE ***
                '//............................................................................//


        '//ITERO NEL RS     *** fine ***
        '//--------------------------------------------------------------------------------//-------//


    End Function





    '//CANCELLO LE TABELLE SE IL FLAG = TRUE SUL RECORDSET
    Function CANCELLA_TABELLE_LINK(par_Disco_s As String, _
                                    par_Path_s As String, _
                                    par_Database_s As String, _
                                    par_Table_s As String, _
                                    par_Cancella_b As Boolean)

    On Error GoTo CANCELLA_TABELLE_LINK_Err







                '//FUNZIONE_DI_CONTROLLO_DELLA_ESISTENZA_DELL'OGGETTO_FORM
                '//CODICE-----> Controlla_SE_OGGETTO_FORM_ESISTE_pFunct.attiva.LaClasse
                '//---------------------------------------------------------------------------------//
                '//NOTA  :Controlla se esiste la form negli oggetti del database e restituisce TRUE _
                          se il nome della form passata con parametro par_sxNameForm esiste; altrimenti FALSE



                            'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_FORM
                            'Assegno e creo l'oggetto di classe
                            '
                            Dim MyCls_CLASSE_GESTIONE_OGGETTI_FORM As CLASSE_GESTIONE_OGGETTI_DATABASE
                            Dim sxNameForm As String
                            '
                            Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = New CLASSE_GESTIONE_OGGETTI_DATABASE


                            str1 = par_Table_s

                            '
                            Boolean1 = MyCls_CLASSE_GESTIONE_OGGETTI_FORM.pfCONTROLLO_ESISTENZA_OGGETTI_TABLE(str1)

                            '
                            sxNameForm = vbNull
                            Set MyCls_CLASSE_GESTIONE_OGGETTI_FORM = Nothing

                    '//---------------------------------------------------------------------------------//


                '//ATTIVO LA CANCELLAZIONE DELLA TABELLA SPECIFICATA SE SUL RECORD E' TRUE = CANCELLA _
                            e se la tabella esiste.
                '//TRUE = sul record è impostato il flag per la cancellazione
                If par_Cancella_b = True Then

                    '//TRUE = ho fatto la ricerca e la tabella esiste deve essere cancellata
                    If Bool1 = True Then
                        DoCmd.DeleteObject acTable, par_Table_s
                    End If

                End If


            CANCELLA_TABELLE_LINK_Exit:
                Exit Function

            CANCELLA_TABELLE_LINK_Err:
                MsgBox Error$
                Resume CANCELLA_TABELLE_LINK_Exit

            End Function

    '//                             FUNZIONE_DI_CANCELLAMENTO_LINK_TABELLE: *** FINE ***
    '//======================================================================================================================//

'//ATTIVA_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct - SEZIONE.05.CODICE---------->UtilizzoDegliOggettiDao
        '//FUNZIONE DI CONTROLLO DEI CAMPI (FIELDS) DELLA TABELLA = controlla tutti gli elementi della tabella e ne restituisce il tipo.
        '//CODICE----------> ATTIVA_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.attiva
        '//****************************************************************************************************************************
        '//NOTE:    Per attivare la funzione occorre inviare il nome della tabella da controllare e il nome del campo e restituisce _
                    una stringa completa che dichiara il nome del campo ed eventualmente la grandezza.

        '//ATTIVO: Attivo la ricerca del campo
        '//-----------------------------------------------------------------------------//
        Private Sub ATTIVA_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct()
        Dim NameTable_s As String
        Dim NameField_s As String

            NameTable_s = "MSys_DF02_TIPO_OGGETTO"          '// nome della tabella da ricerca
            NameField_s = "Scel_b2"                          '// il campo da controllare. es. Scel_b, TIPOGGETTO_s
            '//Restituisco una stringa se trovato oppure null
            Vv1 = CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct(NameTable_s, NameField_s)

                '//CONTROLLO MSGBOX POSITIVO O NEGATIVO
                '//..........................................................................//
                    If Vv1 > "" Then
                       MsgBox NameTable_s & " QUESTA TABELLA CONTIENE questo tipo di campo : " & Vv1, _
                                            vbExclamation, "MSG CONTROLLO TIPO DI CAMPO"
                    Else
                       MsgBox "ATTENZIONE - CONTROLLO DELLA CLASSE DENOMINATA GESTIONE OGGETTI DATABASE:  --->  ! QUESTA TABELLA NON ESISTE -> " & Chr$(13) & Chr$(13) & NameTable_s & _
                                Chr$(13) & Chr$(13) & _
                               "OPPURE NON ESISTE IL TIPO DI CAMPO RICERCATO ->" & _
                                Chr$(13) & Chr$(13) & _
                                NameField_s, vbCritical, "ERRORE RICERCA TIPO DI CAMPO"
                    End If

                '//..........................................................................//

        End Sub

'//DbNortWind_CREA_TABLE_3CAMPI_1CALCOLATO- CICLO NELLE TABELLA DEL DB = Faccio il ciclo sulle tabelle del database corrente. _
                                        TODO:Apro un recordset su una tabella corrente, ma COME FARE IL CICLO SU CAMPI DEL RECORD corrente.??
        '//CODICE----------> DbNortWind_CREA_TABLE_3CAMPI.attiva
        '//****************************************************************************************************************************
        '//NOTE: Nell 'esempio seguente viene illustrato come creare un campo calcolato. _
                Il metodo CreateField crea un campo denominato FullName. La proprietà Expression viene quindi _
                impostata sull 'espressione che calcola il valore del campo.
        Sub DbNortWind_CREA_TABLE_3CAMPI_1CALCOLATO()


        Dim rs As New ADODB.Recordset
        Dim fld As ADODB.Field
        Dim dbs As DAO.Database



                                '02.ITERAZIONE NEL GRUPPO OGGETTI TABLEDEFS ..LE TABELLE DEL DB
                                '______________________________________________________________________________
                                '02.01 - OGGETTI QUERY


                                    'pulisco la tabella degli OGGETTI TABELLA
                                    sSql = ""
                                    sSql = sSql & "DELETE MSsys_TABELLE_Qry01_01_TUTTE.*"
                                    sSql = sSql & "FROM MSsys_TABELLE_Qry01_01_TUTTE "
                                    sSql = sSql & "WITH OWNERACCESS OPTION;"

                                    'CurrentDb.Execute (sSql)

                                    '03.01 - APRO LA TABELLA
                                    'apro la tabella per il salvataggio delle query + effettuo il _
                                    controllo sulla stringa sql
                                    Debug.Print "---------------------------------------------------------------"
                                    Debug.Print "CONTROLLO STRINGA SQL                                          "
                                    Debug.Print

                                    '//creo gli oggetti Dabatase, quello corrente + il recorset
                                    Set dbs = CurrentDb
                                    Set ApriRs = dbs.OpenRecordset("GEST_PARAM_OGGETTI")


                                    With dbs

                                        '//reset
                                        icount = 0

                                        '//stampa di controllo
                                        Debug.Print .TableDefs.Count & _
                                            " TableDefs in database  " & sx_Scelta_db

                                        '//TODO: ciclo sulle tabelle del db corrente
                                        For intCiclo = 0 To .TableDefs.Count - 1
                                                Debug.Print "  " & .TableDefs(intCiclo).Name

                                                '03.02 - SALVATAGGIO NOME OGGETTO TABELLA NELLA TABELLA DEL DB
                                                '................................................................
                                                    'ApriRs.AddNew
                                                    'ApriRs.Fields("Name1_s") = .TableDefs(intCiclo).Name

                                                    '//NRO di oggetto
                                                    icount = icount + 1         '//aggiorno il contatore oggetti delle tabelle
                                                    'ApriRs.Fields("NRO_OGGETTO_i") = icount
                                                    '//Tipo di oggetto caricato = FORM
                                                    'ApriRs.Fields("TIPOGGETTO_s") = "TABLE"

                                                    '//SALVATAGGIO DEGLI ATTRIBUTI DELLA TABELLA - proprieta Attributes
                                                    'ApriRs.Fields("Attribute_lng") = .TableDefs(intCiclo).Attributes

                                                        '//CONTROLLO OGGETTO TABELLA
                                                        '//-----------------------------------------------------------

                                                            '//controllo degli attributi
                                                            Debug.Print "LA TABELLA "
                                                            Debug.Print .TableDefs(intCiclo).Name

                                                            Debug.Print "controllo attributi della tabella"
                                                            Debug.Print .TableDefs(intCiclo).Attributes

                                                        '//-----------------------------------------------------------
                                                    'ApriRs.Update
                                                '................................................................




                                        Next intCiclo

                                    End With



        End Sub


        '//****************************************************************************************************************************

'//pfRICERCA_NEL_DB_OGGETTI_MACRO - CONTROLLO_ESISTENZA_OGGETTI_MACRO @esistenza.MACRO                                   **** INIZIO ****   A.G02
    '//****************************************************************************************************************************
    '//01)_FUNZIONE_CONTROLLO_OGGETTI_MACRO = permette il controllo dell'esistenza della MACRO oggetto all'interno del db corrente _
                                                            oppure in un database esterno.
    '//CODICE----------> pfRICERCA_NEL_DB_OGGETTI_MACRO.Funzione
    '//PARAMETRI: _
            01) sx_Scelta_db    = parametro di SCELTA DATABASE DATABASE da aprire e rappresenta il database esterno o interno; _
                           è esterno a condizione che ci sia una path estesa ed il mome di un db.mdb altrimenti in caso di parametro null _
                           viene inteso come parametro interno. Es.: _
                                    sx_Scelta_db = null il database è quello corrente; _
                                    sx_Scelta_db = C:\MyDb.mdb il database è esterno.
    '//         02) par_sx_Name_MACRO =?
    '//         03) par_sx_Proprieta_MACRO =?
    '//         04) par_ix_Operazione =?
    '//         05) par_ix_Indice1 =?
    '//         06) par_bx_Flag_b =?


    Public Function pfRICERCA_NEL_DB_OGGETTI_MACRO( _
                                            par_sx_Scelta_db As String, _
                                            par_sx_Name_MACRO As String, _
                                            par_sx_Proprieta_MACRO As String, _
                                            par_ix_Operazione As Integer, _
                                            par_ix_Indice1 As Integer, _
                                            par_bx_Flag_b As Boolean) As Boolean



        On Error GoTo Err_pfRICERCA_NEL_DB_OGGETTI_MACRO



                    '//CHIAMA IL CONTROLLO MACRO
                    '//--------------------------------------------------------------------------------//----------//
                    '//CODICE----> pfRICERCA_NEL_DB_OGGETTI_MACRO.CALL

                            '//LE VARIABILE PER IL CONTROLLO OGGETTI
                            'Dim sx_Scelta_db As String
                            'Dim sx_Name_MACRO As String
                            'Dim sx_Proprieta_MACRO As String
                            'Dim ix_Operazione As Integer
                            'Dim ix_Indice1 As Integer
                            'Dim bx_Flag_b As Boolean

                            '//IMPOSTO LE VARIABILI
                            'sx_Scelta_db = ""
                            'sx_Name_MACRO = "GEST_MENU_TB01_VOCI"       '//RICERCO QUESTA MACRO
                            'sx_Proprieta_MACRO = "Proprieta"
                            'ix_Operazione = 0
                            'ix_Indice1 = 0
                            'bx_Flag_b = False



                              '//ATTIVO LA FUNZIONE CON I PARAMETRI
                              'Bool1 = pfRICERCA_NEL_DB_OGGETTI_MACRO(sx_Scelta_db, _
                                                                 sx_Name_MACRO, _
                                                                 sx_Proprieta_MACRO, _
                                                                 ix_Operazione, _
                                                                 ix_Indice1, _
                                                                 bx_Flag_b)

                    '//CHIAMA IL CONTROLLO MACRO *** FINE ***
                    '//--------------------------------------------------------------------------------//----------//



                    '//ITERAZIONE NEGLI OGGETTI MACRO
                    '//==============================================================================================================//
                    '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB


                            'SET VARIABILI:
                            sx_Name_Tab = par_sx_Name_MACRO

                            ix_Operazione = par_ix_Operazione
                                'ix_Indice1= par_Index
                            bx_Flag_b = par_bx_Flag_b
                            sx_Name_Macro = ""

                            'OPEN DB: Apre il database scelto per la creazione o cancellazione a condizione _
                            che esista una path ed il db altrimenti il databse da controllare è quello corrente.
                            If par_sx_Scelta_db > "" Then
                                Set dbs = OpenDatabase(par_sx_Scelta_db)
                            Else
                                Set dbs = CurrentDb
                            End If



                            '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MACRO Nell'insieme Tabledefs
                            '//CONTROLLO PARAMETRI
                            '//-----------------------------------------------------------------------------//

                                '// reset
                                    pfRICERCA_NEL_DB_OGGETTI_MACRO = False

                                Debug.Print "........................................................"
                                Debug.Print "controllo parametri passati alla routine SALVATAGGIO_OGGETTI_DB_PFunct"
                                Debug.Print

                                      Debug.Print "par_sx_Scelta_db ----> " & par_sx_Scelta_db
                                      Debug.Print "par_sx_Name_MACRO ---> " & par_sx_Name_MACRO
                                      Debug.Print "par_sx_Name_MACRO ---> " & par_sx_Proprieta_MACRO
                                      Debug.Print "par_sx_Name_FORM  ---> " & par_ix_Operazione
                                      Debug.Print "par_sx_Name_MACRO --> " & par_sx_Name_MACRO
                                      Debug.Print "par_sx_Name_MACRO ---> " & par_ix_Indice1
                                      Debug.Print "par_sx_Name_Modulo --> " & par_bx_Flag_b
                                      Debug.Print

                                Debug.Print "........................................................"

                            '//-----------------------------------------------------------------------------//


                            '//DBS CICLO
                            '//-----------------------------------------------------------------------------//
                                '//ITERAZIONE NELL'INSIEMO MACRO   @MACRO.itera.oggetti
                                         Set obj_Oggetto = Application.CurrentProject
                                                    icount = 0
                                With dbs
                                '*****************************************


                                    For Each obj In obj_Oggetto.AllMacros

                                        ' Debug.Print obj_Oggetto.Count _
                                       '& " FormsDefs in database  " & sx_Scelta_db


                                        'controlli
                                        '................................................................

                                        '//BLOCCO ITERAZIONE _
                                        Ricerca di una MACRO di sistema o di una MACRO _
                                        inesistente con attivazione del blocco. Il nome della MACRO _
                                        deve essere attivato manualmente.

                                            If obj.Name = "MSysIMEXColumns" Then
                                                MsgBox "BLOCCO ITERERAZIONE sulla MACRO " & obj.Name, vbExclamation, "BLOCCO ITERAZIONE"
                                                Stop
                                            End If

                                            Debug.Print
                                            Debug.Print "  " & obj.Name


                                            '................................................................

                                              '//MACRO TROVATA
                                              If obj.Name = par_sx_Name_MACRO Then

                                                   'COMANDO RICERCA MODULO
                                                  DoCmd.SelectObject acMacro, par_sx_Name_MACRO, True



                                                  Bln1 = True    ' bx_Flag_b True = trovato tabelle si può cancellare
                                                    pfRICERCA_NEL_DB_OGGETTI_MACRO = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                                  GoTo Exit_pfRICERCA_NEL_DB_OGGETTI_MACRO 'esci

                                              End If



                                            '................................................................


                                            intCiclo = intCiclo + 1

                                    Next obj

                                End With


                                '//CHIUSURA OGGETTO, @chiusura.oggetto
                                Set obj_Oggetto = Nothing




                                '*******************************************







                            '//-----------------------------------------------------------------------------//

                '//ITERAZIONE NEGLI OGGETTI MACRO  *** FINE ****
                '//==============================================================================================================//



    '//ERRORI ED USCITA FINE FUNCTIONE E GESTIONE ERRORI
    '//--------------------------------------------------------------------------------//------------------------------------------//
            Exit_pfRICERCA_NEL_DB_OGGETTI_MACRO:
                        dbs.Close
                        Set dbs = Nothing

                        Exit Function

            Err_pfRICERCA_NEL_DB_OGGETTI_MACRO:
                        MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfRICERCA_NEL_DB_OGGETTI_MACRO"
                        Resume Exit_pfRICERCA_NEL_DB_OGGETTI_MACRO


        End Function

    '//*** FINE ***
    '//CONTROLLO_ESISTENZA_OGGETTI_MACRO
    '//****************************************************************************************************************************

'//pfRICERCA_NEL_DB_OGGETTI_Moduli CONTROLLO_ESISTENZA_OGGETTI_MODULI @esistenza.MODULI                                                A.G03
    '//****************************************************************************************************************************
    '//01)_FUNZIONE_CONTROLLO_OGGETTI_MODULI = permette il controllo dell'esistenza della MODULI oggetto all'interno del db corrente _
                                                            oppure in un database esterno.
    '//CODICE----------> pfRICERCA_NEL_DB_OGGETTI_Moduli.Funzione
    '//PARAMETRI: _
            01) sx_Scelta_db    = parametro di SCELTA DATABASE DATABASE da aprire e rappresenta il database esterno o interno; _
                           è esterno a condizione che ci sia una path estesa ed il mome di un db.mdb altrimenti in caso di parametro null _
                           viene inteso come parametro interno. Es.: _
                                    sx_Scelta_db = null il database è quello corrente; _
                                    sx_Scelta_db = C:\MyDb.mdb il database è esterno.
    '//         02) par_sx_Name_MODULI =?
    '//         03) par_sx_Proprieta_MODULI =?
    '//         04) par_ix_Operazione =?
    '//         05) par_ix_Indice1 =?
    '//         06) par_bx_Flag_b =?


    Public Function pfRICERCA_NEL_DB_OGGETTI_Moduli( _
                                            par_sx_Scelta_db As String, _
                                            par_sx_Name_MODULI As String, _
                                            par_sx_Proprieta_MODULI As String, _
                                            par_ix_Operazione As Integer, _
                                            par_ix_Indice1 As Integer, _
                                            par_bx_Flag_b As Boolean) As Boolean



        On Error GoTo Err_pfRICERCA_NEL_DB_OGGETTI_Moduli



                    '//CHIAMA IL CONTROLLO MODULI
                    '//--------------------------------------------------------------------------------//----------//
                    '//CODICE----> pfRICERCA_NEL_DB_OGGETTI_Moduli.CALL

                            '//LE VARIABILE PER IL CONTROLLO OGGETTI
                            'Dim sx_Scelta_db As String
                            'Dim sx_Name_MODULI As String
                            'Dim sx_Proprieta_MODULI As String
                            'Dim ix_Operazione As Integer
                            'Dim ix_Indice1 As Integer
                            'Dim bx_Flag_b As Boolean

                            '//IMPOSTO LE VARIABILI
                            'sx_Scelta_db = ""
                            'sx_Name_MODULI = "GEST_MENU_TB01_VOCI"       '//RICERCO QUESTA MODULI
                            'sx_Proprieta_MODULI = "Proprieta"
                            'ix_Operazione = 0
                            'ix_Indice1 = 0
                            'bx_Flag_b = False



                              '//ATTIVO LA FUNZIONE CON I PARAMETRI
                              'Bool1 = pfRICERCA_NEL_DB_OGGETTI_Moduli(sx_Scelta_db, _
                                                                 sx_Name_MODULI, _
                                                                 sx_Proprieta_MODULI, _
                                                                 ix_Operazione, _
                                                                 ix_Indice1, _
                                                                 bx_Flag_b)

                    '//CHIAMA IL CONTROLLO MODULI *** FINE ***
                    '//--------------------------------------------------------------------------------//----------//



                    '//ITERAZIONE NEGLI OGGETTI MODULI
                    '//==============================================================================================================//
                    '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB


                            'SET VARIABILI:
                            sx_Name_Tab = par_sx_Name_MODULI

                            ix_Operazione = par_ix_Operazione
                                'ix_Indice1= par_Index
                            bx_Flag_b = par_bx_Flag_b
                            sx_Name_MODULI = ""

                            'OPEN DB: Apre il database scelto per la creazione o cancellazione a condizione _
                            che esista una path ed il db altrimenti il databse da controllare è quello corrente.
                            If par_sx_Scelta_db > "" Then
                                Set dbs = OpenDatabase(par_sx_Scelta_db)
                            Else
                                Set dbs = CurrentDb
                            End If



                            '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della MODULI Nell'insieme Tabledefs
                            '//CONTROLLO PARAMETRI
                            '//-----------------------------------------------------------------------------//

                                Debug.Print "........................................................"
                                Debug.Print "controllo parametri passati alla routine SALVATAGGIO_OGGETTI_DB_PFunct"
                                Debug.Print

                                      Debug.Print "par_sx_Scelta_db ----> " & par_sx_Scelta_db
                                      Debug.Print "par_sx_Name_MODULI ---> " & par_sx_Name_MODULI
                                      Debug.Print "par_sx_Name_MODULI ---> " & par_sx_Proprieta_MODULI
                                      Debug.Print "par_sx_Name_FORM  ---> " & par_ix_Operazione
                                      Debug.Print "par_sx_Name_MODULI --> " & par_sx_Name_MODULI
                                      Debug.Print "par_sx_Name_MACRO ---> " & par_ix_Indice1
                                      Debug.Print "par_sx_Name_Modulo --> " & par_bx_Flag_b
                                      Debug.Print

                                Debug.Print "........................................................"

                            '//-----------------------------------------------------------------------------//


                            '//DBS CICLO
                            '//-----------------------------------------------------------------------------//
                                '//ITERAZIONE NELL'INSIEMO MODULI   @moduli.itera.oggetti
                                         Set obj_Oggetto = Application.CurrentProject
                                                    icount = 0
                                With dbs
                                '*****************************************

                                    For Each obj In obj_Oggetto.AllModules

                                        ' Debug.Print obj_Oggetto.Count _
                                       '& " FormsDefs in database  " & sx_Scelta_db


                                        'controlli
                                        '................................................................

                                        '//BLOCCO ITERAZIONE _
                                        Ricerca di una MODULI di sistema o di una MODULI _
                                        inesistente con attivazione del blocco. Il nome della MODULI _
                                        deve essere attivato manualmente.

                                            If obj.Name = "MSysIMEXColumns" Then
                                                MsgBox "BLOCCO ITERERAZIONE sulla MODULI " & obj.Name, vbExclamation, "BLOCCO ITERAZIONE"
                                                Stop
                                            End If

                                            Debug.Print
                                            Debug.Print "  " & obj.Name


                                            '................................................................

                                              '//MODULI TROVATA
                                              If obj.Name = par_sx_Name_MODULI Then

                                                   'COMANDO RICERCA MODULO
                                                  DoCmd.SelectObject acModule, par_sx_Name_MODULI, True



                                                  Bln1 = True    ' bx_Flag_b True = trovato tabelle si può cancellare
                                                  pfRICERCA_NEL_DB_OGGETTI_Moduli = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                                  GoTo Exit_pfRICERCA_NEL_DB_OGGETTI_Moduli 'esci
                                              Else
                                                  '//MODULO NON TROVATO
                                                  Bln1 = False   ' bx_Flag_b False = non trovato QUERY
                                                  pfRICERCA_NEL_DB_OGGETTI_Moduli = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                              End If



                                            '................................................................


                                            intCiclo = intCiclo + 1

                                    Next obj

                                End With


                                '//CHIUSURA OGGETTO, @chiusura.oggetto
                                Set obj_Oggetto = Nothing




                                '*******************************************







                            '//-----------------------------------------------------------------------------//

                '//ITERAZIONE NEGLI OGGETTI MODULI  *** FINE ****
                '//==============================================================================================================//



    '//ERRORI ED USCITA FINE FUNCTIONE E GESTIONE ERRORI
    '//--------------------------------------------------------------------------------//------------------------------------------//
            Exit_pfRICERCA_NEL_DB_OGGETTI_Moduli:
                        dbs.Close
                        Set dbs = Nothing

                        Exit Function

            Err_pfRICERCA_NEL_DB_OGGETTI_Moduli:
                        MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfRICERCA_NEL_DB_OGGETTI_Moduli"
                        Resume Exit_pfRICERCA_NEL_DB_OGGETTI_Moduli


    End Function

    '//*** FINE ***
    '//CONTROLLO_ESISTENZA_OGGETTI_MODULI                                       **** FINE ***
    '//****************************************************************************************************************************

'//pfRICERCA_NEL_DB_OGGETTI_Form -CONTROLLO_ESISTENZA_OGGETTI_FORM @esistenza.FORM                     A.G04
    '//****************************************************************************************************************************
    '//01)_FUNZIONE_CONTROLLO_OGGETTI_FORM = permette il controllo dell'esistenza della FORM oggetto all'interno del db corrente _
                                                            oppure in un database esterno.
    '//CODICE----------> pfRICERCA_NEL_DB_OGGETTI_Form.Funzione
    '//PARAMETRI: _
            01) sx_Scelta_db    = parametro di SCELTA DATABASE DATABASE da aprire e rappresenta il database esterno o interno; _
                           è esterno a condizione che ci sia una path estesa ed il mome di un db.mdb altrimenti in caso di parametro null _
                           viene inteso come parametro interno. Es.: _
                                    sx_Scelta_db = null il database è quello corrente; _
                                    sx_Scelta_db = C:\MyDb.mdb il database è esterno.
    '//         02) par_sx_Name_FORM =?
    '//         03) par_sx_Proprieta_FORM =?
    '//         04) par_ix_Operazione =?
    '//         05) par_ix_Indice1 =?
    '//         06) par_bx_Flag_b =?


    Public Function pfRICERCA_NEL_DB_OGGETTI_Form( _
                                            par_sx_Scelta_db As String, _
                                            par_sx_Name_FORM As String, _
                                            par_sx_Proprieta_FORM As String, _
                                            par_ix_Operazione As Integer, _
                                            par_ix_Indice1 As Integer, _
                                            par_bx_Flag_b As Boolean) As Boolean



        On Error GoTo Err_pfRICERCA_NEL_DB_OGGETTI_Form



                    '//CHIAMA IL CONTROLLO FORM
                    '//--------------------------------------------------------------------------------//----------//
                    '//CODICE----> pfRICERCA_NEL_DB_OGGETTI_Form.CALL

                            '//LE VARIABILE PER IL CONTROLLO OGGETTI
                            'Dim sx_Scelta_db As String
                            'Dim sx_Name_FORM As String
                            'Dim sx_Proprieta_FORM As String
                            'Dim ix_Operazione As Integer
                            'Dim ix_Indice1 As Integer
                            'Dim bx_Flag_b As Boolean

                            '//IMPOSTO LE VARIABILI
                            'sx_Scelta_db = ""
                            'sx_Name_FORM = "GEST_MENU_TB01_VOCI"       '//RICERCO QUESTA FORM
                            'sx_Proprieta_FORM = "Proprieta"
                            'ix_Operazione = 0
                            'ix_Indice1 = 0
                            'bx_Flag_b = False



                              '//ATTIVO LA FUNZIONE CON I PARAMETRI
                              'Bool1 = pfRICERCA_NEL_DB_OGGETTI_Form(sx_Scelta_db, _
                                                                 sx_Name_FORM, _
                                                                 sx_Proprieta_FORM, _
                                                                 ix_Operazione, _
                                                                 ix_Indice1, _
                                                                 bx_Flag_b)

                    '//CHIAMA IL CONTROLLO FORM *** FINE ***
                    '//--------------------------------------------------------------------------------//----------//



                    '//ITERAZIONE NEGLI OGGETTI FORMS
                    '//==============================================================================================================//
                    '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB


                            'SET VARIABILI:
                            sx_Name_Tab = par_sx_Name_FORM

                            ix_Operazione = par_ix_Operazione
                                'ix_Indice1= par_Index
                            bx_Flag_b = par_bx_Flag_b
                            sx_Name_FORM = ""

                            'OPEN DB: Apre il database scelto per la creazione o cancellazione a condizione _
                            che esista una path ed il db altrimenti il databse da controllare è quello corrente.
                            If par_sx_Scelta_db > "" Then
                                Set dbs = OpenDatabase(par_sx_Scelta_db)
                            Else
                                Set dbs = CurrentDb
                            End If



                            '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della FORM Nell'insieme Tabledefs
                            '//CONTROLLO PARAMETRI
                            '//-----------------------------------------------------------------------------//

                                Debug.Print "........................................................"
                                Debug.Print "controllo parametri passati alla routine SALVATAGGIO_OGGETTI_DB_PFunct"
                                Debug.Print

                                      Debug.Print "par_sx_Scelta_db ----> " & par_sx_Scelta_db
                                      Debug.Print "par_sx_Name_FORM ---> " & par_sx_Name_FORM
                                      Debug.Print "par_sx_Name_FORM ---> " & par_sx_Proprieta_FORM
                                      Debug.Print "par_sx_Name_FORM  ---> " & par_ix_Operazione
                                      Debug.Print "par_sx_Name_FORM --> " & par_sx_Name_FORM
                                      Debug.Print "par_sx_Name_MACRO ---> " & par_ix_Indice1
                                      Debug.Print "par_sx_Name_Modulo --> " & par_bx_Flag_b
                                      Debug.Print

                                Debug.Print "........................................................"

                            '//-----------------------------------------------------------------------------//


                            '//DBS CICLO
                            '//-----------------------------------------------------------------------------//
                                '//ITERAZIONE NELL'INSIEMO MODULI   @moduli.itera.oggetti
                                         Set obj_Oggetto = Application.CurrentProject
                                                    icount = 0
                                With dbs
                                '*****************************************

                                    For Each obj In obj_Oggetto.AllForms

                                        ' Debug.Print obj_Oggetto.Count _
                                       '& " FormsDefs in database  " & sx_Scelta_db


                                        'controlli
                                        '................................................................

                                        '//BLOCCO ITERAZIONE _
                                        Ricerca di una FORM di sistema o di una FORM _
                                        inesistente con attivazione del blocco. Il nome della FORM _
                                        deve essere attivato manualmente.

                                            If obj.Name = "MSysIMEXColumns" Then
                                                MsgBox "BLOCCO ITERERAZIONE sulla FORM " & obj.Name, vbExclamation, "BLOCCO ITERAZIONE"
                                                Stop
                                            End If

                                            Debug.Print
                                            Debug.Print "  " & obj.Name


                                            '................................................................

                                              '//FORM TROVATA
                                              If obj.Name = par_sx_Name_FORM Then

                                                   'COMANDO RICERCA FORM
                                                  'DoCmd.SelectObject acForm, par_sx_Name_FORM, True



                                                  Bln1 = True    ' bx_Flag_b True = trovato tabelle si può cancellare
                                                  pfRICERCA_NEL_DB_OGGETTI_Form = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                                  GoTo Exit_pfRICERCA_NEL_DB_OGGETTI_Form 'esci
                                              Else
                                                  '//QUERY NON TROVATA
                                                  Bln1 = False   ' bx_Flag_b False = non trovato form
                                                  pfRICERCA_NEL_DB_OGGETTI_Form = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                              End If



                                            '................................................................


                                            intCiclo = intCiclo + 1

                                    Next obj

                                End With


                                '//CHIUSURA OGGETTO, @chiusura.oggetto
                                Set obj_Oggetto = Nothing




                                '*******************************************







                            '//-----------------------------------------------------------------------------//

                '//ITERAZIONE NEGLI OGGETTI FORMS  *** FINE ****
                '//==============================================================================================================//



    '//ERRORI ED USCITA FINE FUNCTIONE E GESTIONE ERRORI
    '//--------------------------------------------------------------------------------//------------------------------------------//
            Exit_pfRICERCA_NEL_DB_OGGETTI_Form:
                        dbs.Close
                        Set dbs = Nothing

                        Exit Function

            Err_pfRICERCA_NEL_DB_OGGETTI_Form:
                        MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfRICERCA_NEL_DB_OGGETTI_Form"
                        Resume Exit_pfRICERCA_NEL_DB_OGGETTI_Form


        End Function

    '//*** FINE ***
    '//CONTROLLO_ESISTENZA_OGGETTI_FORM
    '//****************************************************************************************************************************

'//pfRICERCA_NEL_DB_OGGETTI_Report CONTROLLO_ESISTENZA_OGGETTI_REPORT @esistenza.report            A.G05
    '//****************************************************************************************************************************
    '//01)_FUNZIONE_CONTROLLO_OGGETTI_REPORT = permette il controllo dell'esistenza della REPORT oggetto all'interno del db corrente _
                                                            oppure in un database esterno.
    '//CODICE----------> pfRICERCA_NEL_DB_OGGETTI_Report.Funzione
    '//PARAMETRI: _
            01) sx_Scelta_db    = parametro di SCELTA DATABASE DATABASE da aprire e rappresenta il database esterno o interno; _
                           è esterno a condizione che ci sia una path estesa ed il mome di un db.mdb altrimenti in caso di parametro null _
                           viene inteso come parametro interno. Es.: _
                                    sx_Scelta_db = null il database è quello corrente; _
                                    sx_Scelta_db = C:\MyDb.mdb il database è esterno.
    '//         02) par_sx_Name_REPORT =?
    '//         03) par_sx_Proprieta_REPORT =?
    '//         04) par_ix_Operazione =?
    '//         05) par_ix_Indice1 =?
    '//         06) par_bx_Flag_b =?


    Public Function pfRICERCA_NEL_DB_OGGETTI_Report( _
                                            par_sx_Scelta_db As String, _
                                            par_sx_Name_Report As String, _
                                            par_sx_Proprieta_REPORT As String, _
                                            par_ix_Operazione As Integer, _
                                            par_ix_Indice1 As Integer, _
                                            par_bx_Flag_b As Boolean) As Boolean



        On Error GoTo Err_pfRICERCA_NEL_DB_OGGETTI_Report



                    '//CHIAMA IL CONTROLLO REPORT
                    '//--------------------------------------------------------------------------------//----------//
                    '//CODICE----> pfRICERCA_NEL_DB_OGGETTI_Report.CALL

                            '//LE VARIABILE PER IL CONTROLLO OGGETTI
                            'Dim sx_Scelta_db As String
                            'Dim sx_Name_REPORT As String
                            'Dim sx_Proprieta_REPORT As String
                            'Dim ix_Operazione As Integer
                            'Dim ix_Indice1 As Integer
                            'Dim bx_Flag_b As Boolean

                            '//IMPOSTO LE VARIABILI
                            'sx_Scelta_db = ""
                            'sx_Name_REPORT = "GEST_MENU_TB01_VOCI"       '//RICERCO QUESTA REPORT
                            'sx_Proprieta_REPORT = "Proprieta"
                            'ix_Operazione = 0
                            'ix_Indice1 = 0
                            'bx_Flag_b = False



                              '//ATTIVO LA FUNZIONE CON I PARAMETRI
                              'Bool1 = pfRICERCA_NEL_DB_OGGETTI_Report(sx_Scelta_db, _
                                                                 sx_Name_REPORT, _
                                                                 sx_Proprieta_REPORT, _
                                                                 ix_Operazione, _
                                                                 ix_Indice1, _
                                                                 bx_Flag_b)

                    '//CHIAMA IL CONTROLLO REPORT *** FINE ***
                    '//--------------------------------------------------------------------------------//----------//



                    '//ITERAZIONE NEGLI OGGETTI TABLES
                    '//==============================================================================================================//
                    '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB


                            'SET VARIABILI:
                            sx_Name_Tab = par_sx_Name_Report

                            ix_Operazione = par_ix_Operazione
                                'ix_Indice1= par_Index
                            bx_Flag_b = par_bx_Flag_b
                            sx_Name_Report = ""

                            'OPEN DB: Apre il database scelto per la creazione o cancellazione a condizione _
                            che esista una path ed il db altrimenti il databse da controllare è quello corrente.
                            If par_sx_Scelta_db > "" Then
                                Set dbs = OpenDatabase(par_sx_Scelta_db)
                            Else
                                Set dbs = CurrentDb
                            End If



                            '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della REPORT Nell'insieme Tabledefs
                            '//CONTROLLO PARAMETRI
                            '//-----------------------------------------------------------------------------//

                                Debug.Print "........................................................"
                                Debug.Print "controllo parametri passati alla routine SALVATAGGIO_OGGETTI_DB_PFunct"
                                Debug.Print

                                      Debug.Print "par_sx_Scelta_db ----> " & par_sx_Scelta_db
                                      Debug.Print "par_sx_Name_REPORT ---> " & par_sx_Name_Report
                                      Debug.Print "par_sx_Name_REPORT ---> " & par_sx_Proprieta_REPORT
                                      Debug.Print "par_sx_Name_FORM  ---> " & par_ix_Operazione
                                      Debug.Print "par_sx_Name_Report --> " & par_sx_Name_Report
                                      Debug.Print "par_sx_Name_MACRO ---> " & par_ix_Indice1
                                      Debug.Print "par_sx_Name_Modulo --> " & par_bx_Flag_b
                                      Debug.Print

                                Debug.Print "........................................................"

                            '//-----------------------------------------------------------------------------//


                            '//DBS CICLO
                            '//-----------------------------------------------------------------------------//
                                '//ITERAZIONE NELL'INSIEMO MODULI   @moduli.itera.oggetti
                                         Set obj_Oggetto = Application.CurrentProject
                                                    icount = 0
                                With dbs
                                '*****************************************

                                    For Each obj In obj_Oggetto.AllReports

                                        ' Debug.Print obj_Oggetto.Count _
                                       '& " FormsDefs in database  " & sx_Scelta_db


                                        'controlli
                                        '................................................................

                                        '//BLOCCO ITERAZIONE _
                                        Ricerca di una REPORT di sistema o di una REPORT _
                                        inesistente con attivazione del blocco. Il nome della REPORT _
                                        deve essere attivato manualmente.

                                            If obj.Name = "MSysIMEXColumns" Then
                                                MsgBox "BLOCCO ITERERAZIONE sulla REPORT " & obj.Name, vbExclamation, "BLOCCO ITERAZIONE"
                                                Stop
                                            End If

                                            Debug.Print
                                            Debug.Print "  " & obj.Name


                                            '................................................................

                                              '//REPORT TROVATA
                                              If obj.Name = par_sx_Name_Report Then

                                                   'COMANDO RICERCA QUERY
                                                  DoCmd.SelectObject acReport, par_sx_Name_Report, True



                                                  Bln1 = True    ' bx_Flag_b True = trovato tabelle si può cancellare
                                                  pfRICERCA_NEL_DB_OGGETTI_Report = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                                  GoTo Exit_pfRICERCA_NEL_DB_OGGETTI_Report 'esci
                                              Else
                                                  '//QUERY NON TROVATA
                                                  Bln1 = False   ' bx_Flag_b False = non trovato QUERY
                                                  pfRICERCA_NEL_DB_OGGETTI_Report = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                              End If



                                            '................................................................


                                            intCiclo = intCiclo + 1

                                    Next obj

                                End With


                                '//CHIUSURA OGGETTO, @chiusura.oggetto
                                Set obj_Oggetto = Nothing




                                '*******************************************







                            '//-----------------------------------------------------------------------------//

                '//ITERAZIONE NEGLI OGGETTI TABLES  *** FINE ****
                '//==============================================================================================================//



    '//ERRORI ED USCITA FINE FUNCTIONE E GESTIONE ERRORI
    '//--------------------------------------------------------------------------------//------------------------------------------//
                Exit_pfRICERCA_NEL_DB_OGGETTI_Report:
                        dbs.Close
                        Set dbs = Nothing

                        Exit Function

                Err_pfRICERCA_NEL_DB_OGGETTI_Report:
                        MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfRICERCA_NEL_DB_OGGETTI_Report"
                        Resume Exit_pfRICERCA_NEL_DB_OGGETTI_Report


            End Function

        '//*** FINE ***
        '//CONTROLLO_ESISTENZA_OGGETTI_REPORT
        '//****************************************************************************************************************************

'//pfRICERCA_NEL_DB_OGGETTI_Query - CONTROLLO_ESISTENZA_OGGETTI_QUERY      A.G06
    '//****************************************************************************************************************************
    '//01)_FUNZIONE_CONTROLLO_OGGETTI_QUERY = permette il controllo dell'esistenza della QUERY oggetto all'interno del db corrente _
                                                            oppure in un database esterno.
    '//CODICE----------> pfRICERCA_NEL_DB_OGGETTI_Query.Funzione
    'MODULO: UTIL_MDL40_N08_RICERCA_TUTTI_GLI_OGGETTI_NEL_DB _
            @ricerca@oggetti_(usare il modulo per la ricerca di tutti gli oggetti)

    '//PARAMETRI: _
            01) sx_Scelta_db    = parametro di SCELTA DATABASE DATABASE da aprire e rappresenta il database esterno o interno; _
                           è esterno a condizione che ci sia una path estesa ed il mome di un db.mdb altrimenti in caso di parametro null _
                           viene inteso come parametro interno. Es.: _
                                    sx_Scelta_db = null il database è quello corrente; _
                                    sx_Scelta_db = C:\MyDb.mdb il database è esterno.
    '//         02) par_sx_Name_Query =?
    '//         03) par_sx_Proprieta_Query =?
    '//         04) par_ix_Operazione =?
    '//         05) par_ix_Indice1 =?
    '//         06) par_bx_Flag_b =?


    Public Function pfRICERCA_NEL_DB_OGGETTI_Query( _
                                            par_sx_Scelta_db As String, _
                                            par_sx_Name_Query As String, _
                                            par_sx_Proprieta_Query As String, _
                                            par_ix_Operazione As Integer, _
                                            par_ix_Indice1 As Integer, _
                                            par_bx_Flag_b As Boolean) As Boolean



        On Error GoTo Err_pfRICERCA_NEL_DB_OGGETTI_Query



                    '//CHIAMA IL CONTROLLO QUERY
                    '//--------------------------------------------------------------------------------//----------//
                    '//CODICE----> pfRICERCA_NEL_DB_OGGETTI_Query.CALL

                            '//LE VARIABILE PER IL CONTROLLO OGGETTI
                            'Dim sx_Scelta_db As String
                            'Dim sx_Name_Query As String
                            'Dim sx_Proprieta_Query As String
                            'Dim ix_Operazione As Integer
                            'Dim ix_Indice1 As Integer
                            'Dim bx_Flag_b As Boolean

                            '//IMPOSTO LE VARIABILI
                            'sx_Scelta_db = ""
                            'sx_Name_Query = "GEST_MENU_TB01_VOCI"       '//RICERCO QUESTA QUERY
                            'sx_Proprieta_Query = "Proprieta"
                            'ix_Operazione = 0
                            'ix_Indice1 = 0
                            'bx_Flag_b = False



                              '//ATTIVO LA FUNZIONE CON I PARAMETRI
                              'Bool1 = pfRICERCA_NEL_DB_OGGETTI_Query(sx_Scelta_db, _
                                                                 sx_Name_Query, _
                                                                 sx_Proprieta_Query, _
                                                                 ix_Operazione, _
                                                                 ix_Indice1, _
                                                                 bx_Flag_b)

                    '//CHIAMA IL CONTROLLO QUERY *** FINE ***
                    '//--------------------------------------------------------------------------------//----------//



                    '//ITERAZIONE NEGLI OGGETTI TABLES
                    '//==============================================================================================================//
                    '//1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB


                            'SET VARIABILI:
                            sx_Name_Tab = par_sx_Name_Query

                            ix_Operazione = par_ix_Operazione
                                'ix_Indice1= par_Index
                            bx_Flag_b = par_bx_Flag_b
                            sx_Name_Query = ""

                            'OPEN DB: Apre il database scelto per la creazione o cancellazione a condizione _
                            che esista una path ed il db altrimenti il databse da controllare è quello corrente.
                            If par_sx_Scelta_db > "" Then
                                Set dbs = OpenDatabase(par_sx_Scelta_db)
                            Else
                                Set dbs = CurrentDb
                            End If



                            '//CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della QUERY Nell'insieme Tabledefs
                            '//CONTROLLO PARAMETRI
                            '//-----------------------------------------------------------------------------//

                                Debug.Print "........................................................"
                                Debug.Print "controllo parametri passati alla routine SALVATAGGIO_OGGETTI_DB_PFunct"
                                Debug.Print

                                      Debug.Print "par_sx_Scelta_db ----> " & par_sx_Scelta_db
                                      Debug.Print "par_sx_Name_Query ---> " & par_sx_Name_Query
                                      Debug.Print "par_sx_Name_Query ---> " & par_sx_Proprieta_Query
                                      Debug.Print "par_sx_Name_FORM  ---> " & par_ix_Operazione
                                      Debug.Print "par_sx_Name_Report --> " & par_sx_Name_Report
                                      Debug.Print "par_sx_Name_MACRO ---> " & par_ix_Indice1
                                      Debug.Print "par_sx_Name_Modulo --> " & par_bx_Flag_b
                                      Debug.Print

                                Debug.Print "........................................................"

                            '//-----------------------------------------------------------------------------//


                            '//DBS CICLO
                            '//-----------------------------------------------------------------------------//

                                With dbs


                                    Debug.Print .QueryDefs.Count & _
                                        " TableDefs in database  " & sx_Scelta_db
                                    For intCiclo = 0 To .QueryDefs.Count - 1

                                        '//BLOCCO ITERAZIONE _
                                        Ricerca di una QUERY di sistema o di una QUERY _
                                        inesistente con attivazione del blocco. Il nome della QUERY _
                                        deve essere attivato manualmente.

                                            If .QueryDefs(intCiclo).Name = "MSysIMEXColumns" Then
                                                MsgBox "BLOCCO ITERERAZIONE sulla QUERY " & .QueryDefs(intCiclo).Name, vbExclamation, "BLOCCO ITERAZIONE"
                                                Stop
                                                End If

                                        Debug.Print
                                        Debug.Print "  " & .QueryDefs(intCiclo).Name


                                        '//QUERY TROVATA
                                        If .QueryDefs(intCiclo).Name = par_sx_Name_Query Then

                                             'COMANDO RICERCA QUERY
                                            DoCmd.SelectObject acQuery, par_sx_Name_Query, True



                                            Bln1 = True    ' bx_Flag_b True = trovato tabelle si può cancellare
                                            pfRICERCA_NEL_DB_OGGETTI_Query = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                            GoTo Exit_pfRICERCA_NEL_DB_OGGETTI_Query  'esci
                                        Else
                                            '//QUERY NON TROVATA
                                            Bln1 = False   ' bx_Flag_b False = non trovato QUERY
                                            pfRICERCA_NEL_DB_OGGETTI_Query = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                        End If

                                    Next intCiclo

                                End With
                            '//-----------------------------------------------------------------------------//

                '//ITERAZIONE NEGLI OGGETTI TABLES  *** FINE ****
                '//==============================================================================================================//



        '//ERRORI ED USCITA FINE FUNCTIONE E GESTIONE ERRORI
        '//--------------------------------------------------------------------------------//------------------------------------------//

            Exit_pfRICERCA_NEL_DB_OGGETTI_Query:
                    dbs.Close
                    Set dbs = Nothing

                    Exit Function

            Err_pfRICERCA_NEL_DB_OGGETTI_Query:
                    MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfRICERCA_NEL_DB_OGGETTI_Query"
                    Resume Exit_pfRICERCA_NEL_DB_OGGETTI_Query


        End Function

    '//*** FINE ***
    '//CONTROLLO_ESISTENZA_OGGETTI_QUERY
    '//****************************************************************************************************************************

'/Public Function pfCONTROLLO_OGGETTI_DB - CONTROLLO OGGETTI DATABASE
    '****************************************************************************************************************************
    '           PROCEDURA CONTROLLO OGGETTI DATABASE                            *** inizio ****       A.G06
    '****************************************************************************************************************************
    ' i parametri ricevuti sono : par_Scelta db = path db
    '                           : par_TableDefNuovo = Tabella da Accodare
    '                           : par_sx_Name_Tab = Nome della Tabella da Cancellare
    '                           : par_ix_Operazione = Tipo di operazione da eseguire (1=inserire tabella; 2=cancellare tabella)
    '                           : par_bx_Flag = Segnala se l'operazione è stata effettuata
    '
    '****************************************************************************************************************************

    Public Function pfCONTROLLO_OGGETTI_DB(par_sx_Scelta_db As String, _
                                       par_sx_Name_Tab As String, par_sx_Name_Query As String, _
                                       par_ix_Operazione As Integer, par_ix_Indice1 As Integer, par_bx_Flag As Boolean) As Boolean




            On Error GoTo Err_pfCONTROLLO_OGGETTI_DB




            '------------------------------------------------------------------------------------------
            '1.1    SET VARIABILI, OPEN DB, CONTROLLO OGGETTO NEL DB


                    'SET VARIABILI:
                    sx_Scelta_db = par_sx_Scelta_db
                    sx_Name_Tab = par_sx_Name_Tab
                    ix_Operazione = par_ix_Operazione
                        'ix_Indice1= par_Index
                    bx_FlAG = par_bx_Flag
                    sx_Name_Query = ""

                    'OPEN DB: Apre il database scelto per la creazione o cancellazione
                    Set dbs = OpenDatabase(sx_Scelta_db)



                    'CONTROLLO OGGETTO NEL DB: Controllo l'esistenza della tabella Nell'insieme Tabledefs

                        With dbs

                            Debug.Print .TableDefs.Count & _
                                " TableDefs in database  " & sx_Scelta_db
                            For intCiclo = 0 To .TableDefs.Count - 1
                                Debug.Print "  " & .TableDefs(intCiclo).Name

                                If .TableDefs(intCiclo).Name = par_sx_Name_Tab Then
                                    Bln1 = True    ' bx_Flag True = trovato tabelle si può cancellare
                                    pfCONTROLLO_OGGETTI_DB = Bln1 'riassegno alla funzione il valore del bx_Flag
                                    GoTo Exit_pfCONTROLLO_OGGETTI_DB  'esci
                                Else
                                    Bln1 = False   ' bx_Flag False = non trovato TABELLA
                                    pfCONTROLLO_OGGETTI_DB = Bln1 'riassegno alla funzione il valore del bx_Flag
                                End If

                            Next intCiclo

                        End With



    '------------------------------------------------------------------------------
    '                       FINE FUNCTIONE E GESTIONE ERRORI

            Exit_pfCONTROLLO_OGGETTI_DB:
                dbs.Close
                Set dbs = Nothing

            Exit Function

            Err_pfCONTROLLO_OGGETTI_DB:
                MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfCONTROLLO_OGGETTI_DB"
                Resume Exit_pfCONTROLLO_OGGETTI_DB



        End Function




    '****************************************************************************************************************************
    '           PROCEDURA CONTROLLO OGGETTI DATABASE                            *** fine ****
    '****************************************************************************************************************************



    '********************************************************************************************************************************************************
    '                   RICERCA OGGETTI NEL DATABASE                                        *** fine ***
    '********************************************************************************************************************************************************

'//ApriFile_3Parametri_PFunct - APRI_IL_FILE_PDF - FUNZIONE INTERNA DI APERTURA DEI FILE CON 3 PARAMETRI-   A.G07
    '//==================================================================================================================//
    '//METODO DI APERTURA DI UN PROGRAMMA ESTERNO O DI UN COMANDO DOS MEDIANTE "WScript.Shell"
    '//PARAMETRI            -> passo 3 PARAMETRI 2 stringhe + 1 LONG
    '//VALORE_DI_RITORNO    -> Nulla
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFile_3Parametri_PFunct.01.00
    '//TODO: manca il 4 parametro che dovrebbe essere l'estensione che per ora non riesco a _
            gestire perche il comando dos non lo permette. TODO: occorre anche implementarenella classe LE API WINDOWS.

    Public Function ApriFile_3Parametri_PFunct(par_Path_s As String, _
                                                 par_NameFile_s As String, _
                                                 par_IDGestione_lng As Long) As String

    Dim Path_programma_s As String
    Dim NameFile_s As String
    Dim Stringa1 As String

    On Error GoTo ApriFile_3Parametri_PFunct_Err


    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFile_3Parametri_PFunct.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi

      '//IMPOSTO LE VARIABILI PATH E FILE
      '//Dim MyFile_s As String
      '//Dim MyPath_s  As String

            '//imposto i parametri - LA PATH
            'MyPath_s = "c:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            'MyFile_s = "Project_PROGETTI_MDB.sublime-project "


             '//chiamo la sub con i parametri =   'CALL (ApriFile_3Parametri_PFunct Path_programma_s, NameFile_s, IDGestione_lng)
                'Call ApriFile_3Parametri_PFunct(MyPath_s, MyFile_s, 0)


    '//--------------------------------------------------------------------------------//--------//


                            '//APRO FILE PDF SPECIFICO
                            '//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>//
                            '//Note       : Apro il file PDF speficico indicando come paramentro la path completa
                            '//           incluso il file PDF.

                              '//La path del file = Unisco la path + il file+estensione sopra individuata
                              par_Path_s = par_Path_s & par_NameFile_s & MyFileEstensione_s

                              '//controllo
                              Debug.Print par_Path_s

                              '// CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
                              '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
                              '// prima la sessione dos e poi il pdf..
                              '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
                              Set WshPDF = CreateObject("wscript.shell")

                              '// CREO UNA STRINGA UNICA, Unisco il comando pdf start + path definitiva altrimenti non funziona.
                              Stringa1 = "%comspec% /c start " & par_Path_s                               '//Comando di apertura
                              '//Attivo il comando
                              WshPDF.Run Stringa1

                              '// restituzione di ok file trovato. TODO: da sistemare
                              ApriFile_3Parametri_PFunct = "ok"

                              '//libero la memoria dalle variabili e dagli oggetti creati
                              WshPDF = Null
                              Stringa1 = vbNull
                              MyFileEstensione_s = vbNull
                              MyPath_s = vbNull



                                '//sistemare ??? invio tasti ???
                                'Lng = 0
                                'For Lng = 1 To 90000000 Step 1

                                'Next Lng

                              'SendKeys "{ESC}"
                              'SendKeys "{ESC}"
                              'SendKeys "{ESC}"


                            '//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>//

           '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
                '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
                '// prima la sessione dos e poi il pdf..
                '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
                'Set WshPDF = CreateObject("wscript.shell")

                '// Unisco il comando pdf start + path definitiva
                'Stringa1 = "%comspec% /c start " & MyPath_s & MyFile_s                               '//Comando di apertura
                    '//Attivo il comando ed aggiungo @exit per @chiudere@il@terminale occorre lo spazio per evitare l'errore & " ^exit"
                'WshPDF.Run Stringa1 & " ^exit"

     ApriFile_3Parametri_PFunct_Exit:
        Exit Function

     ApriFile_3Parametri_PFunct_Err:
        MsgBox Error$
        Resume ApriFile_3Parametri_PFunct_Exit

    End Function

    '//APRI_IL_FILE_PDF - FUNZIONE INTERNA DI APERTURA DEI FILE CON 3 PARAMETRI-         *** FINE ***
    '//==================================================================================================================//

'//SALVA_MENU_pFunct  Public Function e ROUTINE PER I MENU                           **** INIZIO ****
    '//*********************************************************************************************************************************************//
    '//
    '//
    '//*********************************************************************************************************************************************//


    '//SALVA_MENU_pFunct                         A.G08
    '//===========================================================:===================================:================================
    '//SEZX_Mdl_n000_000_Function.000.01__________:Funzione 00 CALCOLA                                :Calcola
    '//...........................................:.....................................................................................
    '//NOTA PROCEDURA: La Funzione attiva un rs per i conteggi dei record e per i calcoli vari.


    Public Function SALVA_MENU_pFunct(par_CodMenu_s As String, par_COD_PROGETTO_s As String, par_Menu_s As String)


    '//ATTIVA LA FUNZIONE SALVA I MENU
    '//.............................................................................//
    '//CODICE -> @SALVA_MENU_pFunct
    '//Note: attenzione occorrono 3 parametri che non devono essere nulli.
        'If Me.CodMenu_s_txt > "" And Me.COD_PROGETTO_s_TXT > "" And Me.Menu_s_TXT > "" Then
        '    Call SALVA_MENU_pFunct(Me.CodMenu_s_txt, Me.COD_PROGETTO_s_TXT, Me.Menu_s_TXT)
        'Else
        '    MsgBox "ATTENZIONE CAMPI MENU NULLI"
        'End If
    '//.............................................................................//



    '//.......................................................
    '// DIM SCELTA DATABASE DA APRIRE
    Dim sScelta_db As String



            On Error GoTo Err_SALVA_MENU_pFunct


            '//SALVA I MENU
            '//-------------------------------------------:-----------------------------------------------


                '//@SALVA@menu@Tmp_(il doppio click per salvare il menu selezionato nella tmp) _
                        01) svuoto la tabella voci tmp e delle funzioni tmp _
                        02) SE IL RS FUNZIONI HA IL MENU CONTROLLO RS SE IL RS E' POPOLATO = _
                            salva in tmp i MENU SOLO SE il MENU esiste nella tabella _
                            tb02 LE FUNZIONI, per cui viene fatto un controllo di popolamento _
                            del rs aperto sulla tabella FUNZIONI. Se il menu esiste la query _
                            estrae un rs e la tabella è popolata, altrimenti se il rs = nullo _
                            non viene fatto l'accodamento ma viene emesso um msg di avviso che _
                            il menu non esiste nelle FUNZIONI. _
                        03) 03 = IL RS E' POPOLATO ESISTE IL MENU E QUINDI INSERISCO LE FUNZIONI IN TMP _
                            dai menu recupera le voci delle funzioni e lo salva nella _
                            tabella tmp delle FUNZIONI_TMP


                    '//01)
                    DoCmd.OpenQuery "GEST_MENU_Tb01Qry01_01_VOCI_TMP_Delete"
                    DoCmd.OpenQuery "GEST_MENU_TB02Qry01_01_FUNZIONI_TMP_Delete"



                                '//02 = SE IL RS FUNZIONI HA IL MENU CONTROLLO RS SE IL RS E' POPOLATO
                                sSql = ""
                                sSql = sSql & "SELECT GEST_MENU_TB02_FUNZIONI.*, GEST_MENU_TB02_FUNZIONI.CodMenu_s AS CONTROLLO_MENU, * "
                                sSql = sSql & "FROM GEST_MENU_TB02_FUNZIONI "
                                sSql = sSql & "WHERE (((GEST_MENU_TB02_FUNZIONI.CodMenu_s) = '" & par_CodMenu_s & "'))"
                                sSql = sSql & "WITH OWNERACCESS OPTION;"


                                '//ITERO NELLA TABELLA
                                '//.....................................................................................................
                                '//NOTE------------>:Tramite una Select vengono individuati i valori da restiuire.
                                '//Codice---------->:DAO_RECORDSET_pSub.01.itero

                                     '//Apro il Database
                                     Set DaoDB = DBEngine.Workspaces(0).Databases(0)
                                     '//Apro un Recordset dal parametro ssql
                                     Set DaoRs = DaoDB.OpenRecordset(sSql)

                                    '//SE IL RS E' POPOLATO
                                    If DaoRs.EOF = False And DaoRs.BOF = False Then
                                        '//Posizione Primo record
                                        DaoRs.MoveFirst

                                                    '//Uscita Rs e chiusura oggetti
                                                    DaoRs.Close
                                                    Set DaoRs = Nothing


                                                    sSql = ""
                                                    sSql = sSql & "INSERT INTO "
                                                    sSql = sSql & "GEST_MENU_Tb01_VOCI_TMP "
                                                    sSql = sSql & "( CodMenu_s, COD_PROGETTO_s, Menu_s  ) "
                                                    sSql = sSql & "SELECT '" & par_CodMenu_s & "' As [codice_menu] , "
                                                    sSql = sSql & "'" & par_COD_PROGETTO_s & "' AS [COD_PROGETTO] ,"
                                                    sSql = sSql & "'" & par_Menu_s & "' AS [MENU] "
                                                    sSql = sSql & "WITH OWNERACCESS OPTION;"

                                                    '//controllo ed esecuzione
                                                    Debug.Print sSql
                                                    CurrentDb.Execute sSql


                                                    '//03 = IL RS E' POPOLATO ESISTE IL MENU E QUINDI INSERISCO LE FUNZIONI IN TMP
                                                    DoCmd.OpenQuery "GEST_MENU_TB02Qry01_02_INSERT_FUNZIONI_To_TMP"

                                        Else
                                            MsgBox "ATTENZIONE NON ESISTE IL MENU ---> " & par_CodMenu_s & " NELLE FUNZIONI, INSERIRE LE FUNZIONI!"
                                        End If  '//If DAORs.EOF = False And DAORs.BOF = False Then


                        '//*** fine ***
                        '//ITERO NELLA TABELLA
                        '//.....................................................................................................








    '//--------------------------------------------------------------------------------------------------
    '//                       FINE FUNCTION E GESTIONE ERRORI

        Exit_SALVA_MENU_pFunct:
            Exit Function

        Err_SALVA_MENU_pFunct:
                MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "SALVA_MENU_pFunct"
                Resume Exit_SALVA_MENU_pFunct

            End Function
    '//SALVA_MENU_pFunct                                              *** FINE ***
    '//===========================================================:===================================:================================




    '//**********************************************************************************************************//
    '//                     MENU:    GeCasaMENU       **** INIZIO ****   A.G09
    '//**********************************************************************************************************//
    Public Sub ATTIVA_ESEGUI_MENU_GeCasaMENU_TMP_pFunct()

        '//ATTIVO IL MENU CON I PARAMETRI :
        '//------------------------------------------------------------------------//
        '//NOTE: attivo la funzione corrispondente al menu CASA scelto
        '//PARAMETRI: _
            par_CodMenu_s           = codice menu , _
            par_COD_PROGETTO_s      = codice progetto _
            par_Menu_s              = il menu _
            par_Funzione            = la funzione attivata

            Call ESEGUI_MENU_GeCasaMENU_TMP_pFunct("", "", "", "GeSpeseCasa")

        '//------------------------------------------------------------------------//

    End Sub




    '//ESEGUI_MENU_GeCasaMENU_TMP_pFunct                                              *** INIZIO ***  A.G10
    '//===========================================================:===================================:================================
    '//NOTE: la funzione riceve 4 parametri: _

    '//01 = SELECT CASE = invididua quale FUNZIONE è attivo

    Public Function ESEGUI_MENU_GeCasaMENU_TMP_pFunct(par_CodMenu_s As String, par_COD_PROGETTO_s As String, par_Menu_s As String, par_Funzione_s As String)


    '//MessaggiDiErrore
    Dim ProceduraMessaggioErrore_s As String
    Dim ProceduraAttivaEseguita_s As String


            On Error GoTo Err_ESEGUI_MENU_GeCasaMENU_TMP_pFunct


            '//RESET
            ProceduraMessaggioErrore_s = ""
            ProceduraAttivaEseguita_s = ""

            ProceduraAttivaEseguita_s = "ERRORE NELLA FUNZIONE --> ESEGUI_MENU_GeCasaMENU_TMP_pFunct"

            '//SELCT QUALE FUNZIONE E' ATTIVATA
            '//............................................................................//

                '//controllo variabile isnull

                If par_Funzione_s = "" Then
                    ProceduraAttivaEseguita_s = "ATTENZIONE FUNZIONE NULLA!! - USCITA DALLA ROUTINE "

                    Debug.Print ProceduraAttivaEseguita_s

                    MsgBox ProceduraAttivaEseguita_s, vbCritical
                    '//USCITA DALLA FUNZIONE
                    GoTo Exit_ESEGUI_MENU_GeCasaMENU_TMP_pFunct

                End If

                    '>>CodMenu   Menu
                    Select Case par_Funzione_s
                        'GeCasaMENU  GESTIONE CASA
                        Case "GeSpeseCasa"
                            MsgBox "MENU CASA NON ATTIVO", vbCritical
                        Case ""
                        Case Else
                            MsgBox "NESSUNA ATTIVITA E' ASSOCIATA PER QUESTA FUNZIONE: " & par_Funzione_s
                        End Select



            '//............................................................................//



    '//--------------------------------------------------------------------------------------------------
    '//                       FINE FUNCTION E GESTIONE ERRORI

        Exit_ESEGUI_MENU_GeCasaMENU_TMP_pFunct:
            Exit Function

        Err_ESEGUI_MENU_GeCasaMENU_TMP_pFunct:
            MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "ESEGUI_MENU_GeCasaMENU_TMP_pFunct " & ProceduraAttivaEseguita_s
            Resume Exit_ESEGUI_MENU_GeCasaMENU_TMP_pFunct

        End Function
        '//ESEGUI_MENU_GeCasaMENU_TMP_pFunct                                              *** FINE ***
        '//===========================================================:===================================:================================


    '//**********************************************************************************************************//
    '//                     MENU:    GeCasaMENU       **** FINE ****
    '//**********************************************************************************************************//




    '//*********************************************************************************************************************************************//
    '//   Public Function e ROUTINE PER I MENU                           **** FINE ****
    '//
    '//
    '//*********************************************************************************************************************************************//

'//DbNortWind - CREA TABELLA CON 1 CAMPO + TABELLA CON 3 CAMPI
    '//****************************************************************************************************************************//
    '//                         UTILIZZO DEGLI OGGETTI DAO - CODICE---------->UtilizzoDegliOggettiDao.Titolo        *** INIZIO ***
            '//  @DB@ESTERNO@DbNortWind
            '//****************************************************************************************************************************//
            '//
            '//                     DAO oggetti table - CREAZIONE DI UNA TABELLA
            '//        ' Esempio IN UN DATABASE ESTERNO CHIAMATO Northwind.mdb _
                        viene creato un nuovo oggetto TableDef che viene aggiunto all'insieme TableDefs _
                         Viene quindi enumerato l'insieme TableDefs e l'insieme Properties del nuovo TableDef. Poi viene cancellato _
                        ma al momento è bloccato . ESEMPI IN -> https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/tabledefs-collection-dao

                    '//NORTWIND CREA UNA TABELLA = in un database esterno con gli oggetti DAO viene creata una tabella con 1 CAMPO.
                    '//CODICE----------> DbNortWind_CREA_TABLE.attiva

                    '//NORTWIND CREA UNA TABELLA = in un database esterno con gli oggetti DAO viene creata una tabella con 3 CAMPI.
                    '//CODICE----------> DbNortWind_CREA_TABLE_3CAMPI.attiva

                    '//FUNZIONE DI CONTROLLO DEI CAMPI (FIELDS) DELLA TABELLA = controlla tutti gli elementi della tabella e ne restituisce il tipo.
                    '//CODICE---------->CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.attiva




            '//NORTWIND CREA UNA TABELLA = in un database esterno con gli oggetti DAO viene creata una tabella con 3 CAMPI.
            '//CODICE----------> DbNortWind_CREA_TABLE_3CAMPI.attiva  @CREAZIONE@3@CAMPI   @CREA@TRE@CAMPI@DATABASE@ESTERNO
            '//****************************************************************************************************************************


            '//SEZIONE.05.CODICE---------->UtilizzoDegliOggettiDao   @CREA@UN@CAMPO
            '//NORTWIND CREA UNA TABELLA = in un database esterno con gli oggetti DAO viene creata una tabella con 1 CAMPO.@DATABASE@ESTERNO
            '//****************************************************************************************************************************
            '//CODICE----------> DbNortWind_CREA_TABLE.attiva
            '|A.G11 | DbNortWind_CREA_TABLE_1CAMPO.....................| crea tabella Nortwind db esterno con 1 campo ..........................'|

              Sub DbNortWind_CREA_TABLE_1CAMPO()

                   Dim dbsNorthwind As Database
                   Dim tdfNew As TableDef
                   Dim tdfLoop As TableDef
                   Dim prpLoop As Property


                   Set dbsNorthwind = OpenDatabase("c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_GESTIONE_OGGETTI_DB_MDB\GESTIONE_OGGETTI\GESTIONE_OGGETTI_MDB\Northwind.mdb")         '//DATABASE ESTERNO METTI LA PATH

                   ' Create new TableDef object, append Field objects
                   ' to its Fields collection, and append TableDef
                   ' object to the TableDefs collection of the
                   ' Database object.
                   Set tdfNew = dbsNorthwind.CreateTableDef("NewTableDef")
                   tdfNew.Fields.Append tdfNew.CreateField("Date", dbDate)
                   dbsNorthwind.TableDefs.Append tdfNew

                   With dbsNorthwind
                      Debug.Print .TableDefs.Count & _
                         " TableDefs in " & .Name

                      ' Enumerate TableDefs collection.
                      For Each tdfLoop In .TableDefs
                         Debug.Print "  " & tdfLoop.Name
                      Next tdfLoop


                      With tdfNew
                         DoEvents
                         Debug.Print "Properties of " & .Name

                        icount = 0

                         ' Enumerate Properties collection of new
                         ' TableDef object, only printing properties
                         ' with non-empty values. ENUMERA LE PROPRIETA DELLA TABELLA
                         For Each prpLoop In .Properties
                            icount = icount + 1
                            Debug.Print "Nro  " & icount & " " & prpLoop.Name & " - " & _
                               IIf(prpLoop = "", "[empty]", prpLoop)
                         Next prpLoop

                      End With

                      ' Delete new TableDef since this is a
                      ' demonstration.
                      '---> bloccato .TableDefs.Delete tdfNew.Name
                      .Close
                   End With

                End Sub

            '//****************************************************************************************************************************






        '|A.G12 | DbNortWind_CREA_TABLE_3CAMPI.....................| crea tabella Nortwind db esterno con 3 campi- ..........................'|

        Public Sub DbNortWind_CREA_TABLE_3CAMPI()

           Dim dbsNorthwind As Database
           Dim tdfNew As TableDef
           Dim prpLoop As Property


           Set dbsNorthwind = OpenDatabase("c:\CASA\LINGUAGGI\ACCESS\ACCESS_PROGETTI_MDB\ACCESS_GESTIONE_OGGETTI_DB_MDB\GESTIONE_OGGETTI\GESTIONE_OGGETTI_MDB\Northwind.mdb")         '//DATABASE ESTERNO METTI LA PATH

           ' Create a new TableDef object.
           Set tdfNew = dbsNorthwind.CreateTableDef("Contacts")

           With tdfNew
              ' ATTENZIONE = SE NON CREI APPEND TABELLA rimane provvisoria im memoria e non vien salvata nel db.
              ' Create fields and append them to the new TableDef
              ' object. This must be done before appending the
              ' TableDef object to the TableDefs collection of the
              ' Northwind database. CREO 4 CAMPI
              .Fields.Append .CreateField("FirstName", dbText)
              .Fields.Append .CreateField("LastName", dbText)
              .Fields.Append .CreateField("Phone", dbText)
              .Fields.Append .CreateField("Notes", dbMemo)
             dbsNorthwind.TableDefs.Append tdfNew           '//LA TABELLE VIENE SALVATA NEL DB negli oggetti DAO ma _
                                                                se disabiliti .Append rimane solo in memoria e non viene salvata.

              Debug.Print "Properties of new TableDef object " & _
                 "before appending to collection:"
                             icount = 0
              ' Enumerate Properties collection of new TableDef
              ' object. ciclo di tutte le tabelle compresa quella nuova.
              For Each prpLoop In .Properties
                 On Error Resume Next
                    icount = icount + 1
                 If prpLoop <> "" Then Debug.Print "  " & _
                          "Nro  " & icount & " " & prpLoop.Name & " = " & prpLoop
                 On Error GoTo 0
              Next prpLoop



           End With

           ' Delete new TableDef object since this is a
           ' demonstration.  CANCELLA LA TABELLA CONTATTI DISABILITA
           ' dbsNorthwind.TableDefs.Delete ("Contacts")

           dbsNorthwind.Close

        End Sub
    '//****************************************************************************************************************************

'//pfRICERCA_NEL_DB_OGGETTI_Table_ControllaCampi - FUNZIONE DI CONTROLLO DELLO OGGETTO TABELLA - CAMPI = controllo il tipo di campo della tabella.  *** INIZIO****
    '//CODICE----------> pfRICERCA_NEL_DB_OGGETTI_Table_ControllaCampi.Funzione
    '//****************************************************************************************************************************
    '//PARAMETRI: _
            01) sx_Scelta_db    = parametro di SCELTA DATABASE DATABASE da aprire e rappresenta il database esterno o interno; _
                           è esterno a condizione che ci sia una path estesa ed il mome di un db.mdb altrimenti in caso di parametro null _
                           viene inteso come parametro interno. Es.: _
                                    sx_Scelta_db = null il database è quello corrente; _
                                    sx_Scelta_db = C:\MyDb.mdb il database è esterno.
    '//         02) par_sx_Name_Table =?
    '//         03) par_sx_Proprieta_Table =?
    '//         04) par_ix_Operazione =?
    '//         05) par_ix_Indice1 =?
    '//         06) par_bx_Flag_b =?


    Private Sub ATTIVA_CONTROLLO()
    Dim sx_Scelta_db As String
    Dim sx_Name_Table As String
    Dim sx_Proprieta_Table As String
    Dim ix_Operazione As Integer
    Dim ix_Indice1 As Integer
    Dim bx_Flag_b As Boolean


    sx_Scelta_db = ""
    sx_Name_Table = ""
    sx_Proprieta_Table = "MSys_DF02_TIPO_OGGETTO"
    ix_Operazione = 0
    ix_Indice1 = 0
    bx_Flag_b = False

    Bool1 = pfRICERCA_NEL_DB_OGGETTI_Table_ControllaCampi(sx_Scelta_db, sx_Name_Table, sx_Proprieta_Table, ix_Operazione, ix_Indice1, bx_Flag_b)


    End Sub


  '// pfRICERCA_NEL_DB_OGGETTI_Table_ControllaCampi
                Public Function pfRICERCA_NEL_DB_OGGETTI_Table_ControllaCampi( _
                                                        par_sx_Scelta_db As String, _
                                                        par_sx_Name_Table As String, _
                                                        par_sx_Proprieta_Table As String, _
                                                        par_ix_Operazione As Integer, _
                                                        par_ix_Indice1 As Integer, _
                                                        par_bx_Flag_b As Boolean) As Boolean



                    On Error GoTo Err_pfRICERCA_NEL_DB_OGGETTI_Table





                                        '//DBS CICLO
                                        '//-----------------------------------------------------------------------------//

                                            With dbs

                                                Debug.Print .TableDefs.Count & _
                                                    " TableDefs in database  " & sx_Scelta_db
                                                For intCiclo = 0 To .TableDefs.Count - 1

                                                    '//BLOCCO ITERAZIONE _
                                                    Ricerca di una tabella di sistema o di una tabella _
                                                    inesistente con attivazione del blocco. Il nome della tabella _
                                                    deve essere attivato manualmente.

                                                        If .TableDefs(intCiclo).Name = "MSysIMEXColumns" Then
                                                            MsgBox "BLOCCO ITERERAZIONE sulla tabella " & .TableDefs(intCiclo).Name, vbExclamation, "BLOCCO ITERAZIONE"
                                                            Stop
                                                            End If

                                                    Debug.Print
                                                    Debug.Print "  " & .TableDefs(intCiclo).Name


                                                    '//TABELLA TROVATA
                                                    If .TableDefs(intCiclo).Name = par_sx_Name_Table Then

                                                         'COMANDO RICERCA TABELLA
                                                        DoCmd.SelectObject acTable, par_sx_Name_Table, True

                                                        Bln1 = True    ' bx_Flag_b True = trovato tabelle si può cancellare
                                                        pfRICERCA_NEL_DB_OGGETTI_Table_ControllaCampi = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                                        GoTo Exit_pfRICERCA_NEL_DB_OGGETTI_Table  'esci
                                                    Else
                                                        '//TABELLA NON TROVATA
                                                        Bln1 = False   ' bx_Flag_b False = non trovato TABELLA
                                                        pfRICERCA_NEL_DB_OGGETTI_Table_ControllaCampi = Bln1 'riassegno alla funzione il valore del bx_Flag_b
                                                    End If

                                                Next intCiclo

                                            End With
                                        '//-----------------------------------------------------------------------------//



                            '------------------------------------------------------------------------------
                            '                       FINE FUNCTIONE E GESTIONE ERRORI

                    '//ERRORI ED USCITA
                    '//--------------------------------------------------------------------------------//------------------------------------------//
        Exit_pfRICERCA_NEL_DB_OGGETTI_Table:
                                        'Dbs.Close
                                        'Set Dbs = Nothing

                                        Exit Function

        Err_pfRICERCA_NEL_DB_OGGETTI_Table:
                                        MsgBox "ERRORE FUNCTION PUBLIC    " & Err.Number & " - " & Err.Description, vbCritical, "pfRICERCA_NEL_DB_OGGETTI_Table"
                                        Resume Exit_pfRICERCA_NEL_DB_OGGETTI_Table


                    End Function

                    '//*** fine ***
                    '//FUNZIONE DI CONTROLLO DELLO OGGETTO TABELLA - CAMPI = controllo il tipo di campo della tabella.   *** FINE ****
                    '//****************************************************************************************************************************




                '//GESTIONE OGGETTI DAO '
                '//****************************************************************************************************************************//
                '//                 UTILIZZO DEGLI OGGETTI DAO - CODICE---------->UtilizzoDegliOggettiDao.Titolo        *** FINE ***
                '//
                '//****************************************************************************************************************************//



    '//**********************************************************************************************************************************************************
    '//
    '//                                     @ULTIME@FUNZIONI@AGGIUNTE                                    **** FINE *****
    '//
    '//**********************************************************************************************************************************************************

'//ZZFINE_MODELLI_ISTANZA_OGGETTI_CLASSE                          *** INIZIO ***
  '//****************************************************************************************************//'

    '//#######################################################################//
    '//
    '//                     MODELLI DI ISTANZIAMENTO CLASSE
    '//
    '//#######################################################################//

    Private Sub ZZFINE_____________________________________________()

            End Sub

    Private Sub ZZFINE_RECUPERO_CODOGGETTO_da_TMP_pbFunct()

            '//CHIAMO LA ROUTINE CONTROLLO_OGGETTO_FORM_LaFormAPERTA DELLA CLASSE OGGETTI FORM
            '//CODICE----------->CONTROLLO_OGGETTO_FORM_LaFormAPERTA.Salvataggio_ATTIVITA.Attivazione
            '//--------------------------------------------------------------------------------//-//
            '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
            '//PARAMETRI            : sono 4 di seguito indicati
            'par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
            'par_sxTIPOGGETTO       =
            'par_sxATTIVITAEVENTO   =
            'par_sxCODATTIVITA      = CODICE ATTIVITA OBBLIGATORIO
            '//NOTA  :Salvo i parametri del codice oggetto e del codice attiva nella tabella tmp solo nel caso in cui _
            sia IL CODICE OGGETTO che il CODICE ATTIVITA sono VALORIZZATI altrimenti la classe ricevente emette _
            un msg di avviso.



                    '1) ISTANZIO LA CLASSE
                    '........................................................

                    'Recupero la stringa da applicare alla
                    'successiva classe
                    '
                    'sxCODOGGETTO = ""
                    '//trovo il codice oggetto mediante il codice attivita
                    'sxCODOGGETTO = TROVA_CLASSE_DI_GESTIONE_pFunct(par_sxCODATTIVITA)

                    'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
                    'Assegno e creo l'oggetto di classe
                    '--->                     Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
                    '--->                     Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE



                        'PARAMETRI DI ESTRAZIONE
                        'Recupero la stringa da applicare alla
                        'successiva classe
                         '---> sxCODOGGETTO = ""
                         '---> sxCODOGGETTO = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.RECUPERO_CODOGGETTO_da_TMP_pbFunct


                    '--->                     Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing

            '//--------------------------------------------------------------------------------//-//
        End Sub

    Private Sub ZZFINE_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct()
        '//CODICE----------->ZZFINE_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.ATTIVA_LA_CLASSE


        '...........................................................................................................................................
        '                                       **** COPIARE DA QUI *****
        '   Devi copiare siala Routine che la funzione ed inserirla in un modulo, classe, form o report per utilizzare le funzioni comuni.
        '...........................................................................................................................................



        '//ROUTINE ATTIVA LA CLASSE PER I CONTROLLO OGGETTI TABLE =  Questa Routine multiuso _
                            puo' essere attivata per diverse chiamate sul tipo di campo, da utilizzare nel modulo _
                            di classe, di una form, report o modulo del progetto ospite. Questa Routine e la Funzione _
                            associata per l'attivazione della classe si trova nella classe identificata da codice _
                            di seguito indicato:
        '//CODICE----------->ZZFINE_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.ATTIVA_LA_CLASSE
        '//^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^//
        '---> Private Sub ATTIVA_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct()
        '---> Dim NameTable_s As String
        '---> Dim NameField_s As String
        '---> Dim AttivaMessaggio_b As Boolean

            '//PER ATTIVARE IL MESSAGGIO -> TRUE
            '---> AttivaMessaggio_b = True                        '//FALSE = MESSAGGIO NON ATTIVATO, TRUE MESSAGGIO DI CONTROLLO ATTIVATO

            '---> NameTable_s = "MSys_DF02_TIPO_OGGETTO"          '// nome della tabella da ricerca
            '---> NameField_s = "Scel_b"                          '// il campo da controllare. es. Scel_b, TIPOGGETTO_s
            '//Restituisco una stringa se trovato oppure null
            '---> vV1 = Cls_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct(NameTable_s, NameField_s)

                '//CONTROLLO MSGBOX POSITIVO O NEGATIVO
                '//..........................................................................//
                    '---> If AttivaMessaggio_b = True Then '//SE TRUE IL MESSAGGIO VIENE ATTIVATO
                        '---> If vV1 > "" Then
                           '---> MsgBox NameTable_s & " QUESTA TABELLA CONTIENE questo tipo di campo : " & vV1, _
                                                vbExclamation, "MSG CONTROLLO TIPO DI CAMPO"
                        '---> Else
                           '---> MsgBox "ATTENZIONE - CONTROLLO DELLA CLASSE DENOMINATA GESTIONE OGGETTI DATABASE:  --->  ! QUESTA TABELLA NON ESISTE -> " & Chr$(13) & Chr$(13) & NameTable_s & _
                                    Chr$(13) & Chr$(13) & _
                                   "OPPURE NON ESISTE IL TIPO DI CAMPO RICERCATO ->" & _
                                    Chr$(13) & Chr$(13) & _
                                    NameField_s, vbCritical, "ERRORE RICERCA TIPO DI CAMPO"
                        '---> End If
                    '---> End If
                '//..........................................................................//

        '--->End Sub




            '//FUNZIONE CHE ATTIVA LA CLASSE OGGETTI = la quale permette la ricerca dei campi della tabella _
                                passata come parametro. Questa funzione è riutilizzabile dalla _
                                routine di seguito indicata. Restituisce una stringa che rappresenta _
                                il campo e le sue proprieta. Puo essere utilizzata per verificare ogni _
                                singolo campo della tabella.
            '---> Private Function Cls_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct(par_NameTable_s As String, par_NameField_s As String) As String

            '//CODICE---------->CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.attiva.classe
            '//--------------------------------------------------------------------------------//-//
            '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
            '//PARAMETRI            : stringa che rappresenta il nome della della tabella + _
                                      il nome del campo da ricerca.
            '//RESTITUISCE          : Stringa ossia il tipo di campo esaminato
            '//NOTE         : L'attivazione della classe viene inserita in una Funzione denominata _
                            ATTIVA_LA_FUNZIONE_MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE _
            '//MODULO_DOVE_SI_TROVA:CODICE---------->CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.attiva
            '//CALSSE_CHE_LA_IMPLEMENTA:CODICE---------->CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct.attiva


                '1) ISTANZIO LA CLASSE
                '........................................................

                'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
                'Assegno e creo l'oggetto di classe
                '---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
                '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE



                    'CONTROLLO SE IL TIPO DI CAMPO
                    'Passo come parametri il nome della tabella in cui ricercare il campo
                     '--->                            NameField_s = "nome tabella?"
                     '--->                            NameField_s = "nome field?"
                     'Restituisce il nome del campo ed il titolo altrimenti se non esiste la tabella _
                        o non esiste il field viene emesso un msg di avviso dalla funzione.
                    '---> Cls_CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.CONTROLLO_OGGETTI_TABLE_Tipo_Field_pFunct(par_NameTable_s, par_NameField_s)
                    '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing
                '//--------------------------------------------------------------------------------//-//

                '---> End Function


                '...........................................................................................................................................
                '                                       **** FINO QUI *****
                '...........................................................................................................................................


        End Sub

    Private Sub ZZFINE_Controlla_FormGESTIONE_APERTA()





            '//CODICE----------->CONTROLLO_OGGETTO_FORM_LaFormAPERTA.Salvataggio_ATTIVITA.Attivazione
            '//--------------------------------------------------------------------------------//-//
            '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
            '//PARAMETRI            : stringa che rappresenta il nome della form.
            '//RESTITUISCE          : bool True se la form è aperta

                    '//LE VARIABILI
                    '--->                    Dim par_sxNameForm   As String                      '//La variabile della form passato come parametro

                    '1) ISTANZIO LA CLASSE
                    '........................................................

                    'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
                    'Assegno e creo l'oggetto di classe
                    '--->                    Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
                    '--->                    Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE



                        'CONTROLLO SE LA FORM E' APERTA
                        'Passo come parametro la stringa nome form e ritora True se è aperta
                         '--->                            par_sxNameForm = Str1
                         '--->                            Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.Controlla_FormGESTIONE_APERTA(par_sxNameForm)


                        '--->                            Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing

            '//--------------------------------------------------------------------------------//-//



                '//IN FONDO PUOI INSERIRE LA ROUTINE nel caso di parecchie operazioni di controllo della form se _
                        aperta oppure no. La routine contiene un solo richiamo della classe.
                '//ROUTINE PER IL CONTROLLO DELLA FORM APERTA
                '//CODICE -----> Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA
                '//==========================================================================================================================================//
                'Private Function Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA(Str1) As Boolean


                'On Error GoTo Err_Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA

                        '//CHIAMO LA FUNZIONE DI CONTROLLO SE LA FORM E' APERTA
                        '//--------------------------------------------------------------------------------//-//
                        '//NOTE: passare come stringa il nome della form da controllare. La funzione attiva _
                                la classe di gestione oggetti con la quale viene effettuato il controllo tra _
                                gli oggetti del db corrente, ossia le form e per quella specificata se la _
                                form passata come parametro è aperta viene restituito TRUE.

                                '----> Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA (Str1)

                        '//--------------------------------------------------------------------------------//-//


                        '// CONTROLLA SE LA FORM GESTIONE E' APERTA
                        '//============================================================================================//
                        '//NOTE : chiamo la funzione per il controllo se la form principale  è aperta _
                        viene passato come parametro il nome della form MASTER e restituisce un Boolean _
                        TRUE se la form è aperta.

                                '//1) ISTANZIO LA CLASSE
                                '//--------------------------------------------------------------------------------//-//


                                    'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
                                    'Assegno e creo l'oggetto di classe
                                    '--->
                                    'Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
                                    '--->
                                    'Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE


                                        '//
                                        'Bolean1 = False

                                        'Bolean1 = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.Controlla_FormGESTIONE_APERTA(Str1)


                                    '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing

                            '//--------------------------------------------------------------------------------//-//

                        '//============================================================================================//





            '//USCITA E CONTROLLO ERRORI
            '//==========================================================================================================================================//

            'Exit_Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA:
            '    Exit Function

            'Err_Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA:
            '    MsgBox Err.Description
            '    Resume Exit_Routine_CONTROLLO_FORM_DI_GESTIONE_APERTA
            'End Function

            'End Function
            '//ROUTINE PER IL CONTROLLO DELLA FORM APERTA *** FINE ***
            '//==========================================================================================================================================//

        End Sub

    Private Sub ZZFINE_Istanzia_TROVA_CLASSE_DI_GESTIONE_pFunct()






                    '//CHIAMO LA ROUTINE SALVO_ATTIVITA_CMB_02_IN_TMP_Sub DELLA CLASSE OGGETTI FORM
                    '//CODICE----------->SALVO_ATTIVITA_CMB_02_IN_TMP_Sub.Salvataggio_ATTIVITA.Attivazione
                    '//--------------------------------------------------------------------------------//-//
                    '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
                    '//PARAMETRI            : sono 4 di seguito indicati
                    'par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
                    'par_sxTIPOGGETTO       =
                    'par_sxATTIVITAEVENTO   =
                    'par_sxCODATTIVITA      = CODICE ATTIVITA OBBLIGATORIO
                    '//NOTA  :Salvo i parametri del codice oggetto e del codice attiva nella tabella tmp solo nel caso in cui _
                    sia IL CODICE OGGETTO che il CODICE ATTIVITA sono VALORIZZATI altrimenti la classe ricevente emette _
                    un msg di avviso.



                                '1) ISTANZIO LA CLASSE
                                '........................................................
                                'Uscita dalla routine se non trova la classe tramite il codice _
                                oggetto allora viene emesso un msg di errore.
                                'Recupero la stringa da applicare alla
                                'successiva classe tramite il codice attivita ritrovo il _
                                codice oggetto.
                                '---> sxCODOGGETTO = ""
                                '//trovo il codice oggetto mediante il codice attivita
                                '---> sxCODOGGETTO = TROVA_CLASSE_DI_GESTIONE_pFunct(par_sxCODATTIVITA)

                                'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
                                'Assegno e creo l'oggetto di classe
                                '---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
                                '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE

                                    'PARAMETRI DI ESTRAZIONE
                                    'Recupero la stringa da applicare alla
                                    'successiva classe
                                    '---> sxCODOGGETTO = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.TROVA_CLASSE_DI_GESTIONE_pFunct(par_sxCODATTIVITA)
                                    '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing


                                    'SE NON ESISTE LA FUNZIONE USCITA DALLA ROUTINE
                                    '........................................................
                                    'Uscita dalla routine se non trova la classe tramite il codice _
                                    oggetto allora viene emesso un msg di errore.

                                        '---> If sxCODOGGETTO = "" Then

                                                '---> MsgBox "Funzione non Abilitata!!", vbExclamation, "MSG GESTIONI"

                                                '---> GoTo Exit_RECUPERO_I_DATI_DALLA_CLASSE_Sub

                                        '---> End If
                                    '........................................................



                        '//--------------------------------------------------------------------------------//-//

        End Sub

    Private Sub ZZFINE_Istanzia_SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP()


        '//CODICE----------->SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct.AttivoLaClasse
        '//--------------------------------------------------------------------------------//-//
        '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
        '//PARAMETRI            : sono il par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO

        '//NOTA  : AGGIORNO IL CODICE OGGETTO nella tabella oggetti tmp GEST_Tb01_N01_PARAM_OGGETTI_TMP



                       'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
                       'Assegno e creo l'oggetto di classe
                       'Creo la variabile oggetto                              ---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
                       'Creo l'oggetto classe                                  ---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE

                       'RECUPERO PARAMETRO DA TABELLA OGGETTI
                        'Utilizzo la funzione pubblica di recupero del codice oggetto dal
                        'codice attivita. Restituisco la chiamante il CODICE OGGETTO
                        'Creo l'oggetto e attivo la funzione e ricevo la stringa ---> Str1 = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.SALVATAGGIO_PARAMETRO_CODOGGETTO_in_TMP_pFunct(par_sxCODOGGETTO)

                       'Libero la memoria                                       ---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing

        '//--------------------------------------------------------------------------------//------------------------------------------------//
        End Sub

    Private Sub ZZFINE_Istanzia_SALVO_ATTIVITA_CMB_02_IN_TMP_Sub()

            '//CHIAMO LA ROUTINE SALVO_ATTIVITA_CMB_02_IN_TMP_Sub DELLA CLASSE OGGETTI FORM
            '//CODICE----------->SALVO_ATTIVITA_CMB_02_IN_TMP_Sub.Salvataggio_ATTIVITA.Attivazione
            '//--------------------------------------------------------------------------------//-//
            '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
            '//PARAMETRI            : sono 4 di seguito indicati
            'par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO
            'par_sxTIPOGGETTO       =
            'par_sxATTIVITAEVENTO   =
            'par_sxCODATTIVITA      = CODICE ATTIVITA OBBLIGATORIO
            '//NOTA  :Salvo i parametri del codice oggetto e del codice attiva nella tabella tmp solo nel caso in cui _
            sia IL CODICE OGGETTO che il CODICE ATTIVITA sono VALORIZZATI altrimenti la classe ricevente emette _
            un msg di avviso.


                'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
                'Assegno e creo l'oggetto di classe
                '---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
                '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE



                'PARAMETRI DI ESTRAZIONE
                'Attivo il metodo della classe rappresentata dalla routine passando
                'i 4 parametri necessari
                'sxCODOGGETTO = s
                'sxTIPOGGETTO = s
                'sxATTIVITAEVENTO = s
                'sxCODATTIVITA = s
                '---> MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.SALVO_ATTIVITA_CMB_02_IN_TMP_Sub sxCODOGGETTO, sxTIPOGGETTO, sxATTIVITAEVENTO, sxCODATTIVITA


                '---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing


            '//--------------------------------------------------------------------------------//-//


        End Sub

    Private Sub ZZFINE_Istanzia_RECUPERO_CODOGGETTO()

        '//CODICE----------->RECUPERO_CODOGGETTO_pbFunct.AttivoLaClasse
        '//--------------------------------------------------------------------------------//-//
        '//CLASSE_OSPITE        : CLASSE_GESTIONE_OGGETTI_DATABASE
        '//PARAMETRI            : sono il par_sxCODOGGETTO       = CODICE OGGETTO OBBLIGATORIO

        '//NOTA  : AGGIORNO IL CODICE OGGETTO nella tabella oggetti tmp GEST_Tb01_N01_PARAM_OGGETTI_TMP



                       'ISTANZIO LA CLASSE CLASSE_GESTIONE_OGGETTI_DATABASE
                       'Assegno e creo l'oggetto di classe
                       'Creo la variabile oggetto                              ---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE
                       'Creo l'oggetto classe                                  ---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE

                       'RECUPERO PARAMETRO DA TABELLA OGGETTI
                        'Utilizzo la funzione pubblica di recupero del codice oggetto dal
                        'codice attivita. Restituisco la chiamante il CODICE OGGETTO
                        'Creo l'oggetto e attivo la funzione e ricevo la stringa ---> TROVA_CODOGGETTO_Funct = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.RECUPERO_CODOGGETTO_pbFunct(sxCODATTIVITA)

                       'Libero la memoria                                       ---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing

            '//--------------------------------------------------------------------------------//------------------------------------------------//


        End Sub

    Private Sub ZZFINE_Istanzia_CERCA_OGGETTI()


            '//01) CLASSE_GESTIONE_OGGETTI_DATABASE
            '//ISTANZIA LA CLASSE
            '//06.30
            '//FUNZIONE CERCA OGGETTI = Con due parametri Tipo oggetto e Nome oggetto lo ricerca nella db.
            '//CODICE-------->CERCA_OGGETTI.01.Attiva
            '//.....................................................................................................
                '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
                    '//---> Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE

                '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE assegnando _
                     l'oggetto di classe alla variabile e creo creo l'oggetto di classe
                    '//---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE

                '//PARAMETRI= TIPO OGGETTO E NOME OGGETTO.
                    '//---> Str1 = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.CERCA_OGGETTI(sxTIPOGGETTO, sxNOMEOGGETTO)

                '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
                    '//---> Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing

            '//.....................................................................................................



        End Sub

    Private Sub ZZFINE_ESEMPI_ISTANZIAMENTO_CLASSE() '//ESEMPI_ISTANZIAMENTO_CLASSE - @ESEMPI@ISTANZIAMENTO@CLASSE

        '//01) CLASSE CON LA PATH ED IL FILE DA APRIRE
        '//
        '//MODELLO ISTANZIAMENTO CLASSE ATTI DI GESTIONE
        '//.....................................................................................................
            '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
                'Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE

            '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE assegnando _
               l 'oggetto di classe alla variabile e creo creo l'oggetto di classe
                'Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE

            '//PATH O FILE = Inserire o una path oppure una path+file con percoso esteso.
                'Str1 = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_SERVIZI\MDB"
                'MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.APRO_PATH_O_FILE_GENERICI_pSub (Str1)

            '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
                'Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing

        '//.....................................................................................................


        '//02) CLASSE CON PER L'ATTIVAZIONE DI UNA FUNZIONE O ROUTINE E CON RETURN DI UN BOOLEAN
        '//
        '//MODELLO ISTANZIAMENTO CLASSE ATTI DI GESTIONE
        '//.....................................................................................................
            '//IMPOSTO LA VARIABILE OGGETTO = per l'utilizzo della classe
                'Dim MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE As CLASSE_GESTIONE_OGGETTI_DATABASE

            '//ISTANZIO LA CLASSE = MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE assegnando _
               l 'oggetto di classe alla variabile e creo creo l'oggetto di classe
                'Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = New CLASSE_GESTIONE_OGGETTI_DATABASE

            '//CHIAMO LA FUNZIONE CON I PARAMETRI E CON RESTITUZIONE DI UN BOOLEAN
                'Str1 = "C:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHVI_MDB\LLPP_GESTIONE\LPP_ANALISI\GESTIONE_SERVIZI\MDB"
                'Bool1 =  MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE.pfRICERCA_NEL_DB_OGGETTI_Table(sx_Scelta_db, _
                                                                                          sx_Name_Table, _
                                                                                          sx_Proprieta_Table, _
                                                                                          ix_Operazione, _
                                                                                          ix_Indice1, _
                                                                                          bx_Flag_b)




            '//RILASCIO LA MEMORIA = rilascio l'oggetto ed attivo il distruttore della classe
                'Set MyCls_CLASSE_GESTIONE_OGGETTI_DATABASE = Nothing

        '//.....................................................................................................


        End Sub





        '//                                     @CLASSE@OGGETTI@DATABASE                                    **** FINE *****
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@




        '//**********************************************************************************************************************************************************
        '//
        '//                                     @ULTIME@FUNZIONI@AGGIUNTE                                    **** INIZIO *****
        '//
        '//         @FUNZ_01=LoopFormProps_FormCorrente_PULISCI_OGGETTI_TEXT_LABEL_PSub
        '//         @FUNZ_02=SALVO_ATTIVITA_CMB_02_IN_TMP_Sub
        '//
        '//
        '//**********************************************************************************************************************************************************

    '//ZZFINE_MODELLI_ISTANZA_OGGETTI_CLASSE                          *** FINE ***
    '//****************************************************************************************************//'

