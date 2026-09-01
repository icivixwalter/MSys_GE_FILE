VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Msys_GEFILE_Frm01_GESTIONI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit


' -- API per mostrare la finestra standard di selezione file (GetOpenFileName)
#If VBA7 Then
    Private Declare PtrSafe Function GetOpenFileNameA Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
#Else
    Private Declare Function GetOpenFileNameA Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
#End If

' -- Costanti per il comportamento della finestra di dialogo
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXPLORER = &H80000

' -- Struttura necessaria per l'API GetOpenFileNameA
#If VBA7 Then
    Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As LongPtr
        hInstance As LongPtr
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
        lCustData As LongPtr
        lpfnHook As LongPtr
        lpTemplateName As String
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



'//@APRI@FILE_(DA @PATH E FILE DA SELEZIONARE)
'//-------------------------------------------------------------------------------------------//
'//Print Cosa; è; stato; aggiunto: _
On Error GoTo GestioneErrore per la gestione errori. _
Exit Sub per evitare di finire nella gestione se tutto va bene. _
Messaggio di errore dettagliato in caso di problemi (es. file non trovabile o accesso negato).

Private Sub Cmd_ApriFilePDF_Click()
    On Error GoTo GestioneErrore ' Attiva la gestione degli errori

    ' Variabile per il percorso del file selezionato
    Dim filePath As String

    ' Crea una nuova istanza di FileDialog (per aprire un file)
    Dim fileDialog As Object
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)

    ' Imposta il filtro per i file MDB
    fileDialog.Filters.Clear
    'fileDialog.Filters.Add "Microsoft Access Database", "*.MDB"

    ' Imposta la directory di partenza (ad esempio "C:\")
    fileDialog.InitialFileName = Me.PATH_BASE_TXT

    ' Mostra il dialogo e verifica se l'utente ha selezionato un file
    If fileDialog.Show = -1 Then
        ' Recupera il percorso del file selezionato
        filePath = fileDialog.SelectedItems(1)

        ' Mostra il percorso del file selezionato (ad esempio in una MsgBox)
        MsgBox "Hai selezionato il file: " & filePath

        ' ?? Apre il file selezionato
        FollowHyperlink filePath
    Else
        MsgBox "Nessun file selezionato."
    End If

    Exit Sub ' Esce prima della gestione errori

GestioneErrore:
    MsgBox "Errore durante l'apertura del file: " & Err.Description, vbCritical, "Errore"
End Sub

'//-------------------------------------------------------------------------------------------//

Private Sub Cmd_FILTER_Click()
    
    
    '//@IMPOSTA@PERCORSO_(nel campo della @casella.di@testo)
    '//----------------------------------------------------------------------------------------------//
        '//dim percorso per la casella di testo
        Dim percorso As Variant
    
        ' Mostra finestra di selezione file
        percorso = ShowOpenFileDialog(False, "Tutti i file (*.*)", "*.*", "Seleziona un file", "C:\")
        ' controllo percorso
        Debug.Print "percorso = " & percorso
        ' Se l'utente ha selezionato un file, mostra il percorso in una casella di testo
        If Not IsNull(percorso) Then
            '//STAMPO IL  PERCORSO ED IL FILE NELLA CASELLA DI TESTO
            Me!TxtPercorso_TXT = percorso  ' Cambia "TxtPercorso" con il nome del tuo controllo
        End If
        
    '//----------------------------------------------------------------------------------------------//

    
    
    ' Variabile per il percorso del file selezionato
    Dim filePath As String
    ' Crea una nuova istanza di FileDialog (per aprire un file)
    Dim fileDialog As Object
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    
    ' Imposta il filtro per i file MDB
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "Microsoft Access Database", "*.MDB"
    
    ' Imposta la directory di partenza (ad esempio "C:\")
    fileDialog.InitialFileName = "C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\GESTIONE_OGGETTI\MDB\"
    
    ' Mostra il dialogo e verifica se l'utente ha selezionato un file
    If fileDialog.Show = -1 Then
        ' Recupera il percorso del file selezionato
        filePath = fileDialog.SelectedItems(1)
        
        ' Mostra il percorso del file selezionato (ad esempio in una MsgBox)
        MsgBox "Hai selezionato il file: " & filePath
    Else
        MsgBox "Nessun file selezionato."
    End If
End Sub


'//APRI FILE CON FILTRO PREIMPOSTATO
'//-------------------------------------------------------------------------------------------//
Private Sub Cmd_FILTER_FILE_PREIMPOSTATO_FILE_PREIMPOSTATO_Click()

    ' Variabile per il percorso completo del file
    Dim filePath As String
    Dim officePath As String
    Dim command As String
    
    ' Crea una nuova istanza di FileDialog (per aprire un file)
    Dim fileDialog As Object
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    
    ' Imposta la directory di partenza
    fileDialog.InitialFileName = "C:\Casa\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSys_DF02.mdb"
    
    ' Imposta il filtro per i file MDB
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "Microsoft Access Database", "*.MDB"
    
    '//@IMPOSTA@PERCORSO_(nel campo della @casella.di@testo)

    ' Mostra il dialogo e verifica se l'utente ha selezionato un file
    If fileDialog.Show = -1 Then
        ' Recupera il percorso completo del file selezionato
        filePath = fileDialog.SelectedItems(1)
        
        
       '//----------------------------------------------------------------------------------------------//
        '//dim percorso per la casella di testo
        Dim percorso As Variant
    
        ' Mostra finestra di selezione file
        percorso = filePath
        ' controllo percorso
        Debug.Print "percorso = " & percorso
        ' Se l'utente ha selezionato un file, mostra il percorso in una casella di testo
        If Not IsNull(percorso) Then
            '//STAMPO IL  PERCORSO ED IL FILE NELLA CASELLA DI TESTO
            Me!TxtPercorso_TXT = percorso  ' Cambia "TxtPercorso" con il nome del tuo controllo
        End If
        
    '//----------------------------------------------------------------------------------------------//
    
        
        ' Mostra il percorso del file selezionato (ad esempio in una MsgBox)
        MsgBox "Hai selezionato il file: " & filePath
        
        ' Verifica se il file esiste
        If Len(Dir(filePath)) > 0 Then
            ' Ottieni il percorso di MSACCESS.EXE dalla variabile di ambiente
            officePath = Environ("ACCESS_2011")
            
            ' Verifica se la variabile di ambiente è correttamente impostata
            If officePath <> "" Then
                ' Concatenazione della variabile filePath con il percorso di Access
                command = """" & officePath & """" & " """ & filePath & """"
                
                '//controllo prima della shell:
                Debug.Print " file selezionato"
                Debug.Print " variabile valore = " & command
                
                ' Esegui il comando con Shell
                Shell command, vbNormalFocus
            Else
                MsgBox "La variabile di ambiente ACCESS_2011 non è impostata correttamente."
            End If
        Else
            MsgBox "Il file selezionato non esiste o non è valido."
        End If
    Else
        MsgBox "Nessun file selezionato."
    End If

End Sub

'//-------------------------------------------------------------------------------------------//


'//proprieta file
Private Sub Cmd_RecuperaFILE_Click()
    On Error GoTo Err_Handler
    
    ' Apri la finestra di selezione del file
    Dim result As Variant
    result = ShowOpenFileDialog(False, "Tutti i file (*.*)", "*.*", "Seleziona un file", "C:\")
    
    
   
    '//@IMPOSTA@PERCORSO_(nel campo della @casella.di@testo)
    '//----------------------------------------------------------------------------------------------//
        '//dim percorso per la casella di testo
        Dim percorso As Variant
    
        ' Mostra finestra di selezione file
        percorso = ShowOpenFileDialog(False, "Tutti i file (*.*)", "*.*", "Seleziona un file", "C:\")
        ' controllo percorso
        Debug.Print "percorso = " & percorso
        ' Se l'utente ha selezionato un file, mostra il percorso in una casella di testo
        If Not IsNull(percorso) Then
            '//STAMPO IL  PERCORSO ED IL FILE NELLA CASELLA DI TESTO
            Me!TxtPercorso_TXT = percorso  ' Cambia "TxtPercorso" con il nome del tuo controllo
        End If
        
    '//----------------------------------------------------------------------------------------------//

 
    
    ' Verifica se è stato selezionato un file
    If Not IsNull(result) Then
        ' Ottieni il percorso completo del file
        Dim filePath As String
        filePath = result
        
        ' Visualizza il file selezionato
        MsgBox "File selezionato: " & filePath
 
        ' Chiama una funzione per ottenere e mostrare le proprietà del file
        MostraProprietaFile filePath
        
        ' Se vuoi aprire il file (ad esempio PDF, Word, etc.)
        ' ApriFileConShellExecute filePath
    Else
        MsgBox "Nessun file selezionato."
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    MsgBox "Errore: " & Err.Description
    Resume Exit_Handler
End Sub

' Funzione per mostrare le proprietà del file
Private Sub MostraProprietaFile(filePath As String)
    ' Verifica che il file esista
    If Dir(filePath) <> "" Then
        ' Usa il file system object per ottenere le proprietà
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        Dim FILE As Object
        Set FILE = fso.GetFile(filePath)
        
        ' Mostra le proprietà del file
        MsgBox "Proprietà del file:" & vbCrLf & _
               "Percorso: " & FILE.PATH & vbCrLf & _
               "Nome: " & FILE.Name & vbCrLf & _
               "Tipo: " & FILE.Type & vbCrLf & _
               "Dimensione: " & FILE.Size & " bytes" & vbCrLf & _
               "Data di creazione: " & FILE.DateCreated & vbCrLf & _
               "Data di modifica: " & FILE.DateLastModified
    Else
        MsgBox "Il file non esiste o non è stato trovato."
    End If
End Sub
Private Sub Cmd_RecuperaPath_Click()
    On Error GoTo GestioneErrore ' Attiva la gestione degli errori
 
    '//@IMPOSTA@PERCORSO_(nel campo della @casella.di@testo)
    '//----------------------------------------------------------------------------------------------//
        '//dim percorso per la casella di testo
        Dim percorso As Variant
    
        ' Mostra finestra di selezione file
        percorso = ShowOpenFileDialog(False, "Tutti i file (*.*)", "*.*", "Seleziona un file", "C:\")
        ' controllo percorso
        Debug.Print "percorso = " & percorso
        ' Se l'utente ha selezionato un file, mostra il percorso in una casella di testo
        If Not IsNull(percorso) Then
            '//STAMPO IL  PERCORSO ED IL FILE NELLA CASELLA DI TESTO
            Me!TxtPercorso_TXT = percorso  ' Cambia "TxtPercorso" con il nome del tuo controllo
        End If
        
    '//----------------------------------------------------------------------------------------------//


    Exit Sub ' Esce prima della gestione errori

GestioneErrore:
    MsgBox "Errore durante la selezione del file: " & Err.Description, vbCritical, "Errore"
End Sub



'//@APRI@FILE_(l'evento DOPPIO CLICK della casella di testo APRO IL FILE ATTIVANDO IL COMANDO BUTTON )
'//-------------------------------------------------------------------------------------------//

Private Sub FILE_BASE_TXT_DblClick(Cancel As Integer)
    '//chiamo il comando di apertura DI APERTURA DEL FILE
        Cmd_ApriFilePDF_Click
    
    
End Sub
'//-------------------------------------------------------------------------------------------//

'//@APRI@PATH_(l'evento click della casella di testo che permette l'apertura della path)
'//-------------------------------------------------------------------------------------------//
Private Sub PATH_BASE_TXT_Click()
    On Error GoTo GestioneErrore ' Attiva la gestione degli errori

    ' Variabile per il percorso della cartella
    Dim folderPath As String

    ' Recupera il percorso della cartella dalla casella di testo
    folderPath = Me.PATH_BASE_TXT

    ' Controllo che la cartella esista
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "La cartella specificata non esiste: " & folderPath, vbExclamation
        Exit Sub
    End If

    ' Mostra il percorso selezionato (ad esempio in una MsgBox)
    MsgBox "Apro la cartella: " & folderPath

    ' ?? Apre la cartella nel file explorer
    Shell "explorer.exe """ & folderPath & """", vbNormalFocus

    Exit Sub ' Esce prima della gestione errori

GestioneErrore:
    MsgBox "Errore durante l'apertura della cartella: " & Err.Description, vbCritical, "Errore"
End Sub
'//-------------------------------------------------------------------------------------------//


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
    Dim OFN As OPENFILENAME
    Dim sFileName As String
    Dim filterString As String
    Dim ret As Long

    ' -- Crea il filtro nel formato richiesto (doppio null terminatore)
    filterString = filtroDescrizione & vbNullChar & filtroEstensione & vbNullChar

    ' -- Prepara un buffer per i nomi dei file (sufficiente per multi-selezione)
    sFileName = String$(4096, vbNullChar)

    ' -- Compila la struttura OPENFILENAME
    With OFN
        .lStructSize = LenB(OFN)
        .hwndOwner = 0  ' Compatibile con Access 2000 (non usa Application.hWndAccessApp)
        .lpstrFilter = filterString
        .nFilterIndex = 1
        .lpstrFile = sFileName
        .nMaxFile = Len(sFileName)
        .lpstrFileTitle = String$(260, vbNullChar)
        .nMaxFileTitle = 260
        .lpstrInitialDir = cartellaIniziale
        .lpstrTitle = titoloFinestra
        .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_EXPLORER
        If allowMultiSelect Then .flags = .flags Or OFN_ALLOWMULTISELECT
    End With

    ' -- Mostra la finestra di dialogo
    ret = GetOpenFileNameA(OFN)

    ' -- Se l'utente ha selezionato uno o più file
    If ret <> 0 Then
        Dim s As String
        s = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, vbNullChar) - 1)

        ' -- Se è abilitata la selezione multipla
        If allowMultiSelect Then
            Dim fileList() As String
            fileList = Split(Left$(OFN.lpstrFile, InStr(OFN.lpstrFile & vbNullChar, vbNullChar & vbNullChar) - 1), vbNullChar)

            If UBound(fileList) = 0 Then
                ' Solo un file selezionato
                ShowOpenFileDialog = fileList(0)
            Else
                ' Più file: combinare il percorso con ogni nome file
                Dim i As Long, result() As String
                ReDim result(UBound(fileList) - 1)
                For i = 1 To UBound(fileList)
                    result(i - 1) = fileList(0) & "\" & fileList(i)
                Next i
                ShowOpenFileDialog = result
            End If
        Else
            ' Solo un file selezionato
            ShowOpenFileDialog = s
        End If
    Else
        ' Utente ha annullato
        ShowOpenFileDialog = Null
    End If
End Function



'//@APRO IL FILE
Private Sub TxtPercorso_TXT_DblClick(Cancel As Integer)

  ' Variabile per il percorso della cartella
    Dim folderPath As String

    '//IF SE VALORIZZATO
    If Me.TxtPercorso_TXT > "" Then
            ' Recupera il percorso della cartella dalla casella di testo
            folderPath = Me.TxtPercorso_TXT
        
            ' Controllo che la cartella esista
            If Dir(folderPath, vbDirectory) = "" Then
                MsgBox "La cartella specificata non esiste: " & folderPath, vbExclamation
                Exit Sub
            End If
        
            ' Mostra il percorso selezionato (ad esempio in una MsgBox)
            ' MsgBox "Apro il file presalvato : " & folderPath
        
            ' ?? Apre la cartella nel file explorer
            Shell "explorer.exe """ & folderPath & """", vbNormalFocus
    End If
End Sub
