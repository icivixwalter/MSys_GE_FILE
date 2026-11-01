VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLASSE_clsFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' CLASSE: CLASSE_clsFileSystem
' RESPONSABILITA': Gestione di file e cartelle (apertura, selezione, proprieta)
' SUPPORTO: Percorsi lunghi fino a 32.767 caratteri (API W + prefisso \\->\)
' NOTE: tutte le API sono in versione Unicode (suffisso W) per gestire
'       correttamente caratteri non-ASCII e percorsi lunghi.
' ============================================================================
'
' INDICE DEI PARAGRAFI (ricercabili con Ctrl+F, es. "§3.2" o "3.2")
' ----------------------------------------------------------------------------
'  §1   COSTANTI
'   §1.1  Costanti ShellExecute
'   §1.2  Costanti GetOpenFileName / GetSaveFileName
'   §1.3  Costanti percorsi lunghi
'   §1.4  Costanti attributi file (WIN32)
'   §1.5  Costanti flag CopyFile / MoveFileEx
'   §1.6  Costanti handle API
'
'  §2   STRUTTURE
'   §2.1  WIN32_FILE_ATTRIBUTE_DATA
'   §2.2  WIN32_FIND_DATAW
'   §2.3  OPENFILENAME (VBA7 / pre-VBA7)
'
'  §3   DICHIARAZIONI API (tutte Unicode)
'   §3.1  Dichiarazioni VBA7 (PtrSafe)
'   §3.2  Dichiarazioni pre-VBA7
'
'  §4   METODI PUBBLICI
'   §4.1   SelezionaFile
'   §4.2   SalvaFile
'   §4.3   ApriFile
'   §4.4   FileEsiste
'   §4.5   CartellaEsiste
'   §4.6   MostraProprietaFile
'   §4.7   CopiaFile
'   §4.8   SpostaFile
'   §4.9   EliminaFile
'   §4.10  CreaCartella
'   §4.11  EliminaCartella
'   §4.12  GetFileName
'   §4.13  GetFolderPath
'   §4.14  GetFileExtension
'   §4.15  GetFileSizeFormatted
'   §4.16  IsLongPath
'   §4.17  NormalizzaPercorso
'   §4.18  ApriPathConControllo  (MODELLO - vedi §7)
'
'  §5   FUNZIONI PRIVATE
'   §5.1   TrimNull
'   §5.2   GetAttributiFile
'   §5.3   IsPercorsoAssoluto
'   §5.4   RisolviPercorsoAssoluto
'   §5.5   CreaCartellaRicorsiva
'   §5.6   EliminaCartellaRicorsiva
'   §5.7   GetFileAttributesText
'
'  §6   NOTE DI MIGRAZIONE (rispetto alla versione originale)
'
'  §7   MODELLI DI UTILIZZO DELLA CLASSE SYSTEM
'   §7.1  ApriPathConControllo (metodo pubblico)
'   §7.2  Esempio di chiamata dal form
' ============================================================================


Option Compare Database
Option Explicit

' ============================================================================
' §1   COSTANTI
' ============================================================================

' --- §1.1  Costanti ShellExecute -------------------------------------------
Private Const SW_SHOWNORMAL As Long = 1
Private Const SE_ERR_SUCCESS As Long = 32

' --- §1.2  Costanti GetOpenFileName / GetSaveFileName -----------------------
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_LONGNAMES As Long = &H200000

' --- §1.3  Costanti percorsi lunghi -----------------------------------------
Private Const MAX_PATH As Long = 260
Private Const LONG_PATH_PREFIX As String = "\\?\"
Private Const LONG_UNC_PREFIX As String = "\\?\UNC\"
Private Const BUFFER_LONG_PATH As Long = 32767

' --- §1.4  Costanti attributi file (WIN32) ----------------------------------
Private Const FILE_ATTRIBUTE_READONLY   As Long = &H1
Private Const FILE_ATTRIBUTE_HIDDEN     As Long = &H2
Private Const FILE_ATTRIBUTE_SYSTEM     As Long = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE    As Long = &H20
Private Const FILE_ATTRIBUTE_ALIAS      As Long = &H40
Private Const FILE_ATTRIBUTE_COMPRESSED As Long = &H800
Private Const INVALID_FILE_ATTRIBUTES   As Long = -1

' --- §1.5  Costanti flag CopyFile / MoveFileEx -------------------------------
Private Const COPY_FILE_FAIL_IF_EXISTS As Long = &H1
Private Const MOVEFILE_REPLACE_EXISTING As Long = &H1
Private Const MOVEFILE_COPY_ALLOWED     As Long = &H2


' --- §1.6  Costanti handle API ---------------------------------------------
' Valore restituito da FindFirstFileW in caso di errore (= (HANDLE)-1).
' Su VBA7 (LongPtr a 64 bit) va dichiarato come LongPtr; su VBA a 32 bit
' va dichiarato come Long. Per compatibilita' lo dichiariamo con #If.
#If VBA7 Then
    Private Const INVALID_HANDLE_VALUE As LongPtr = -1
#Else
    Private Const INVALID_HANDLE_VALUE As Long = -1
#End If

' ============================================================================
' §2   STRUTTURE
' ============================================================================

' --- §2.1  WIN32_FILE_ATTRIBUTE_DATA (36 byte) ------------------------------
Private Type WIN32_FILE_ATTRIBUTE_DATA
    dwFileAttributes As Long
    ftCreationTime   As Currency
    ftLastAccessTime As Currency
    ftLastWriteTime  As Currency
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
End Type

' --- §2.2  WIN32_FIND_DATAW (per enumerazione ricorsiva cartelle) -----------
Private Type WIN32_FIND_DATAW
    dwFileAttributes As Long
    ftCreationTime   As Currency
    ftLastAccessTime As Currency
    ftLastWriteTime  As Currency
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * 260
    cAlternateFileName As String * 14
End Type

' --- §2.3  OPENFILENAME -----------------------------------------------------
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
        lCustData As Long
        lpfnHook As LongPtr
        lpTemplateName As String
        pvReserved As LongPtr
        dwReserved As Long
        FlagsEx As Long
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
        lpfnHook As LongPtr
        lpTemplateName As String
    End Type
#End If

' ============================================================================
' §3   DICHIARAZIONI API (tutte Unicode, supportano \\->\)
' ============================================================================

' --- §3.1  Dichiarazioni VBA7 (PtrSafe) -------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecuteW Lib "shell32.dll" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As LongPtr, _
        ByVal lpFile As LongPtr, _
        ByVal lpParameters As LongPtr, _
        ByVal lpDirectory As LongPtr, _
        ByVal nShowCmd As Long) As LongPtr

    Private Declare PtrSafe Function GetOpenFileNameW Lib "comdlg32.dll" ( _
        pOpenfilename As OPENFILENAME) As Long

    Private Declare PtrSafe Function GetSaveFileNameW Lib "comdlg32.dll" ( _
        pOpenfilename As OPENFILENAME) As Long

    Private Declare PtrSafe Function GetFileAttributesExW Lib "kernel32" ( _
        ByVal lpFileName As LongPtr, _
        ByVal fInfoLevelId As Long, _
        ByRef lpFileInformation As WIN32_FILE_ATTRIBUTE_DATA) As Long

    Private Declare PtrSafe Function GetFullPathNameW Lib "kernel32" ( _
        ByVal lpFileName As LongPtr, _
        ByVal nBufferLength As Long, _
        ByVal lpBuffer As LongPtr, _
        ByRef lpFilePart As LongPtr) As Long

    Private Declare PtrSafe Function CopyFileW Lib "kernel32" ( _
        ByVal lpExistingFileName As LongPtr, _
        ByVal lpNewFileName As LongPtr, _
        ByVal bFailIfExists As Long) As Long

    Private Declare PtrSafe Function MoveFileExW Lib "kernel32" ( _
        ByVal lpExistingFileName As LongPtr, _
        ByVal lpNewFileName As LongPtr, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function DeleteFileW Lib "kernel32" ( _
        ByVal lpFileName As LongPtr) As Long

    Private Declare PtrSafe Function CreateDirectoryW Lib "kernel32" ( _
        ByVal lpPathName As LongPtr, _
        ByVal lpSecurityAttributes As LongPtr) As Long

    Private Declare PtrSafe Function RemoveDirectoryW Lib "kernel32" ( _
        ByVal lpPathName As LongPtr) As Long

    Private Declare PtrSafe Function FindFirstFileW Lib "kernel32" ( _
        ByVal lpFileName As LongPtr, _
        ByRef lpFindFileData As WIN32_FIND_DATAW) As LongPtr

    Private Declare PtrSafe Function FindNextFileW Lib "kernel32" ( _
        ByVal hFindFile As LongPtr, _
        ByRef lpFindFileData As WIN32_FIND_DATAW) As Long

    Private Declare PtrSafe Function FindClose Lib "kernel32" ( _
        ByVal hFindFile As LongPtr) As Long

    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

    Private Declare PtrSafe Function SysAllocString Lib "oleaut32.dll" ( _
        ByVal psz As LongPtr) As LongPtr

    Private Declare PtrSafe Function SysFreeString Lib "oleaut32.dll" ( _
        ByVal bstrString As LongPtr)
#Else

' --- §3.2  Dichiarazioni pre-VBA7 -------------------------------------------
    Private Declare Function ShellExecuteW Lib "shell32.dll" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As Long, _
        ByVal lpFile As Long, _
        ByVal lpParameters As Long, _
        ByVal lpDirectory As Long, _
        ByVal nShowCmd As Long) As Long

    Private Declare Function GetOpenFileNameW Lib "comdlg32.dll" ( _
        pOpenfilename As OPENFILENAME) As Long

    Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" ( _
        pOpenfilename As OPENFILENAME) As Long

    Private Declare Function GetFileAttributesExW Lib "kernel32" ( _
        ByVal lpFileName As Long, _
        ByVal fInfoLevelId As Long, _
        ByRef lpFileInformation As WIN32_FILE_ATTRIBUTE_DATA) As Long

    Private Declare Function GetFullPathNameW Lib "kernel32" ( _
        ByVal lpFileName As Long, _
        ByVal nBufferLength As Long, _
        ByVal lpBuffer As Long, _
        ByRef lpFilePart As Long) As Long

    Private Declare Function CopyFileW Lib "kernel32" ( _
        ByVal lpExistingFileName As Long, _
        ByVal lpNewFileName As Long, _
        ByVal bFailIfExists As Long) As Long

    Private Declare Function MoveFileExW Lib "kernel32" ( _
        ByVal lpExistingFileName As Long, _
        ByVal lpNewFileName As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function DeleteFileW Lib "kernel32" ( _
        ByVal lpFileName As Long) As Long

    Private Declare Function CreateDirectoryW Lib "kernel32" ( _
        ByVal lpPathName As Long, _
        ByVal lpSecurityAttributes As Long) As Long

    Private Declare Function RemoveDirectoryW Lib "kernel32" ( _
        ByVal lpPathName As Long) As Long

    Private Declare Function FindFirstFileW Lib "kernel32" ( _
        ByVal lpFileName As Long, _
        ByRef lpFindFileData As WIN32_FIND_DATAW) As Long

    Private Declare Function FindNextFileW Lib "kernel32" ( _
        ByVal hFindFile As Long, _
        ByRef lpFindFileData As WIN32_FIND_DATAW) As Long

    Private Declare Function FindClose Lib "kernel32" ( _
        ByVal hFindFile As Long) As Long

    Private Declare Function GetLastError Lib "kernel32" () As Long

    Private Declare Function SysAllocString Lib "oleaut32.dll" ( _
        ByVal psz As Long) As Long

    Private Declare Function SysFreeString Lib "oleaut32.dll" ( _
        ByVal bstrString As Long)
#End If

' ============================================================================
' §4   METODI PUBBLICI
' ============================================================================

' ----------------------------------------------------------------------------
' §4.1  SelezionaFile
' ----------------------------------------------------------------------------
' Mostra la finestra di dialogo "Apri file" (Unicode, supporta path lunghi).
Public Function SelezionaFile(Optional filtroDescrizione As String = "Tutti i file (*.*)", _
                               Optional filtroEstensione As String = "*.*", _
                               Optional cartellaIniziale As String = vbNullString, _
                               Optional titolo As String = "Seleziona un file") As String

    On Error GoTo ErrHandler

    Dim OFN As OPENFILENAME
    Dim strFileName As String
    Dim strFilter As String
    Dim ok As Long

    strFilter = filtroDescrizione & Chr$(0) & filtroEstensione & Chr$(0) & Chr$(0)

    ' Buffer Unicode ampio: GetOpenFileNameW accetta path lunghi
    strFileName = Space$(BUFFER_LONG_PATH)

    With OFN
        .lStructSize = Len(OFN)
        #If VBA7 Then
            .hwndOwner = Application.hWndAccessApp
        #Else
            .hwndOwner = Application.hWndAccessApp
        #End If
        .lpstrFilter = strFilter
        .nFilterIndex = 1
        .lpstrFile = strFileName
        .nMaxFile = Len(strFileName)
        .lpstrFileTitle = Space$(1024)
        .nMaxFileTitle = 1024
        .lpstrInitialDir = cartellaIniziale
        .lpstrTitle = titolo
        .flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_LONGNAMES Or OFN_EXPLORER
    End With

    ok = GetOpenFileNameW(OFN)
    If ok <> 0 Then
        SelezionaFile = TrimNull(OFN.lpstrFile)
    Else
        SelezionaFile = vbNullString
    End If

    Exit Function
ErrHandler:
    SelezionaFile = vbNullString
End Function

' ----------------------------------------------------------------------------
' §4.2  SalvaFile
' ----------------------------------------------------------------------------
' Mostra la finestra di dialogo "Salva file" (Unicode, supporta path lunghi).
Public Function SalvaFile(Optional filtroDescrizione As String = "Tutti i file (*.*)", _
                           Optional filtroEstensione As String = "*.*", _
                           Optional cartellaIniziale As String = vbNullString, _
                           Optional nomeFileSuggerito As String = "", _
                           Optional titolo As String = "Salva file") As String

    On Error GoTo ErrHandler

    Dim OFN As OPENFILENAME
    Dim strFileName As String
    Dim strFilter As String
    Dim ok As Long

    strFilter = filtroDescrizione & Chr$(0) & filtroEstensione & Chr$(0) & Chr$(0)

    If Len(nomeFileSuggerito) > 0 Then
        If Len(nomeFileSuggerito) >= BUFFER_LONG_PATH Then
            strFileName = Left$(nomeFileSuggerito, BUFFER_LONG_PATH - 1) & vbNullChar
        Else
            strFileName = nomeFileSuggerito & Space$(BUFFER_LONG_PATH - Len(nomeFileSuggerito))
        End If
    Else
        strFileName = Space$(BUFFER_LONG_PATH)
    End If

    With OFN
        .lStructSize = Len(OFN)
        #If VBA7 Then
            .hwndOwner = Application.hWndAccessApp
        #Else
            .hwndOwner = Application.hWndAccessApp
        #End If
        .lpstrFilter = strFilter
        .nFilterIndex = 1
        .lpstrFile = strFileName
        .nMaxFile = Len(strFileName)
        .lpstrFileTitle = Space$(1024)
        .nMaxFileTitle = 1024
        .lpstrInitialDir = cartellaIniziale
        .lpstrTitle = titolo
        .flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_LONGNAMES Or OFN_EXPLORER
    End With

    ok = GetSaveFileNameW(OFN)
    If ok <> 0 Then
        SalvaFile = TrimNull(OFN.lpstrFile)
    Else
        SalvaFile = vbNullString
    End If

    Exit Function
ErrHandler:
    SalvaFile = vbNullString
End Function

' ----------------------------------------------------------------------------
' §4.3  ApriFile
' ----------------------------------------------------------------------------
' Apre un file o una cartella con l'applicazione predefinita.
' Per le cartelle lunghe usa explorer.exe + parametro (piu' affidabile di
' "explore" con \\->\). Per le cartelle corte usa "explore" classico.
Public Function ApriFile(percorsoCompleto As String, Optional showCmd As Long = SW_SHOWNORMAL) As Boolean
    On Error GoTo ErrHandler

    If Trim$(percorsoCompleto) = "" Then
        ApriFile = False
        Exit Function
    End If

    Dim percorsoApi As String
    percorsoApi = NormalizzaPercorso(percorsoCompleto)

    Dim eCartella As Boolean
    eCartella = CartellaEsiste(percorsoCompleto)

    Dim risultato As LongPtr

    If eCartella Then
        If IsLongPath(percorsoApi) Then
            ' Cartella lunga: explorer.exe + parametro
            Dim sOp As String:    sOp = "open"
            Dim sFile As String:  sFile = "explorer.exe"
            Dim sParam As String: sParam = """" & percorsoApi & """"
            risultato = ShellExecuteW(0, StrPtr(sOp), StrPtr(sFile), StrPtr(sParam), 0, showCmd)
        Else
            ' Cartella corta: "explore" classico
            Dim sOp2 As String:  sOp2 = "explore"
            Dim sPath2 As String: sPath2 = percorsoApi
            risultato = ShellExecuteW(0, StrPtr(sOp2), StrPtr(sPath2), 0, 0, showCmd)
        End If
    Else
        ' File: "open"
        Dim sOp3 As String:  sOp3 = "open"
        Dim sFile3 As String: sFile3 = percorsoApi
        risultato = ShellExecuteW(0, StrPtr(sOp3), StrPtr(sFile3), 0, 0, showCmd)
    End If

    ApriFile = (risultato > SE_ERR_SUCCESS)

    Exit Function
ErrHandler:
    ApriFile = False
End Function



' ----------------------------------------------------------------------------
' §4.4  FileEsiste
' ----------------------------------------------------------------------------
' Restituisce True se il percorso esiste ed e' un file (non una directory).
Public Function FileEsiste(filePath As String) As Boolean
    On Error GoTo ErrHandler

    If Trim$(filePath) = "" Then
        FileEsiste = False
        Exit Function
    End If

    Dim attribs As Long
    attribs = GetAttributiFile(filePath)

    FileEsiste = (attribs <> INVALID_FILE_ATTRIBUTES And (attribs And FILE_ATTRIBUTE_DIRECTORY) = 0)

    Exit Function
ErrHandler:
    FileEsiste = False
End Function

' ----------------------------------------------------------------------------
' §4.5  CartellaEsiste
' ----------------------------------------------------------------------------
' Restituisce True se il percorso esiste ed e' una directory.
Public Function CartellaEsiste(cartellaPath As String) As Boolean
    On Error GoTo ErrHandler

    If Trim$(cartellaPath) = "" Then
        CartellaEsiste = False
        Exit Function
    End If

    Dim attribs As Long
    attribs = GetAttributiFile(cartellaPath)

    CartellaEsiste = (attribs <> INVALID_FILE_ATTRIBUTES And (attribs And FILE_ATTRIBUTE_DIRECTORY) <> 0)

    Exit Function
ErrHandler:
    CartellaEsiste = False
End Function

' ----------------------------------------------------------------------------
' §4.6  MostraProprietaFile
' ----------------------------------------------------------------------------
' Mostra un MsgBox con le principali proprieta del file (dimensione,
' attributi, flag percorso lungo). Usa GetFileAttributesExW, quindi
' funziona anche su path > 260.
Public Sub MostraProprietaFile(filePath As String)
    On Error GoTo ErrHandler

    If Trim$(filePath) = "" Then
        MsgBox "Percorso file vuoto.", vbExclamation, "Errore"
        Exit Sub
    End If

    If Not FileEsiste(filePath) Then
        MsgBox "File non trovato: " & filePath, vbExclamation, "Errore"
        Exit Sub
    End If

    Dim fad As WIN32_FILE_ATTRIBUTE_DATA
    Dim percorso As String
    percorso = NormalizzaPercorso(filePath)

    If GetFileAttributesExW(StrPtr(percorso), 0, fad) = 0 Then
        MsgBox "Impossibile leggere gli attributi del file.", vbExclamation, "Errore"
        Exit Sub
    End If

    Dim sizeDouble As Double
    sizeDouble = CDbl(fad.nFileSizeHigh) * 4294967296# + CDbl(fad.nFileSizeLow)

    Dim infoMsg As String
    infoMsg = "Proprieta del file:" & vbCrLf & vbCrLf
    infoMsg = infoMsg & "Percorso: " & filePath & vbCrLf
    infoMsg = infoMsg & "Nome: " & GetFileName(filePath) & vbCrLf
    infoMsg = infoMsg & "Dimensione: " & Format$(sizeDouble, "#,##0") & " bytes" & vbCrLf
    infoMsg = infoMsg & "Dimensione: " & Format$(sizeDouble / 1024, "#,##0.0") & " KB" & vbCrLf
    infoMsg = infoMsg & "Dimensione: " & Format$(sizeDouble / 1048576, "#,##0.0") & " MB" & vbCrLf
    infoMsg = infoMsg & "Attributi: " & GetFileAttributesText(filePath) & vbCrLf
    infoMsg = infoMsg & "Percorso lungo: " & IIf(Len(percorso) > MAX_PATH, "Si", "No")

    MsgBox infoMsg, vbInformation, "Informazioni sul file"

    Exit Sub
ErrHandler:
    MsgBox "Errore: " & Err.Description, vbCritical, "Errore"
End Sub

' ----------------------------------------------------------------------------
' §4.7  CopiaFile
' ----------------------------------------------------------------------------
' Copia un file con CopyFileW (supporta path lunghi e Unicode).
Public Function CopiaFile(sourcePath As String, destinationPath As String, Optional overwrite As Boolean = True) As Boolean
    On Error GoTo ErrHandler

    If Not FileEsiste(sourcePath) Then
        CopiaFile = False
        Exit Function
    End If

    Dim src As String, dst As String
    src = NormalizzaPercorso(sourcePath)
    dst = NormalizzaPercorso(destinationPath)

    Dim failIfExists As Long
    If overwrite Then failIfExists = 0 Else failIfExists = COPY_FILE_FAIL_IF_EXISTS

    CopiaFile = (CopyFileW(StrPtr(src), StrPtr(dst), failIfExists) <> 0)

    Exit Function
ErrHandler:
    CopiaFile = False
End Function

' ----------------------------------------------------------------------------
' §4.8  SpostaFile
' ----------------------------------------------------------------------------
' Sposta/rinomina un file con MoveFileExW (supporta path lunghi e Unicode).
Public Function SpostaFile(sourcePath As String, destinationPath As String) As Boolean
    On Error GoTo ErrHandler

    If Not FileEsiste(sourcePath) Then
        SpostaFile = False
        Exit Function
    End If

    Dim src As String, dst As String
    src = NormalizzaPercorso(sourcePath)
    dst = NormalizzaPercorso(destinationPath)

    SpostaFile = (MoveFileExW(StrPtr(src), StrPtr(dst), _
                              MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED) <> 0)

    Exit Function
ErrHandler:
    SpostaFile = False
End Function

' ----------------------------------------------------------------------------
' §4.9  EliminaFile
' ----------------------------------------------------------------------------
' Elimina un file con DeleteFileW (supporta path lunghi e Unicode).
Public Function EliminaFile(filePath As String) As Boolean
    On Error GoTo ErrHandler

    If Not FileEsiste(filePath) Then
        EliminaFile = False
        Exit Function
    End If

    Dim percorso As String
    percorso = NormalizzaPercorso(filePath)
    EliminaFile = (DeleteFileW(StrPtr(percorso)) <> 0)

    Exit Function
ErrHandler:
    EliminaFile = False
End Function

' ----------------------------------------------------------------------------
' §4.10  CreaCartella
' ----------------------------------------------------------------------------
' Crea una cartella (e tutta la catena di padri mancanti) con
' CreateDirectoryW. Supporta path lunghi e Unicode.
Public Function CreaCartella(cartellaPath As String) As Boolean
    On Error GoTo ErrHandler

    If Trim$(cartellaPath) = "" Then
        CreaCartella = False
        Exit Function
    End If

    If CartellaEsiste(cartellaPath) Then
        CreaCartella = True
        Exit Function
    End If

    Dim percorso As String
    percorso = NormalizzaPercorso(cartellaPath)

    CreaCartella = CreaCartellaRicorsiva(percorso)

    Exit Function
ErrHandler:
    CreaCartella = False
End Function

' ----------------------------------------------------------------------------
' §4.11  EliminaCartella
' ----------------------------------------------------------------------------
' Elimina una cartella. Se forceDelete = True, svuota ricorsivamente
' il contenuto prima di rimuoverla.
Public Function EliminaCartella(cartellaPath As String, Optional forceDelete As Boolean = False) As Boolean
    On Error GoTo ErrHandler

    If Not CartellaEsiste(cartellaPath) Then
        EliminaCartella = False
        Exit Function
    End If

    Dim percorso As String
    percorso = NormalizzaPercorso(cartellaPath)

    If forceDelete Then
        EliminaCartella = EliminaCartellaRicorsiva(percorso)
    Else
        EliminaCartella = (RemoveDirectoryW(StrPtr(percorso)) <> 0)
    End If

    Exit Function
ErrHandler:
    EliminaCartella = False
End Function

' ----------------------------------------------------------------------------
' §4.12  GetFileName
' ----------------------------------------------------------------------------
' Restituisce il nome del file (ultimo segmento) da un percorso completo.
' Pura manipolazione di stringhe: funziona anche su path > 260.
Public Function GetFileName(percorsoCompleto As String) As String
    On Error GoTo ErrHandler

    If Trim$(percorsoCompleto) = "" Then
        GetFileName = ""
        Exit Function
    End If

    Dim p As String
    p = percorsoCompleto
    Do While Len(p) > 1 And Right$(p, 1) = "\"
        p = Left$(p, Len(p) - 1)
    Loop

    Dim pos As Long
    pos = InStrRev(p, "\")
    If pos > 0 Then
        GetFileName = Mid$(p, pos + 1)
    Else
        GetFileName = p
    End If

    Exit Function
ErrHandler:
    GetFileName = ""
End Function

' ----------------------------------------------------------------------------
' §4.13  GetFolderPath
' ----------------------------------------------------------------------------
' Restituisce il percorso della cartella contenitrice.
Public Function GetFolderPath(percorsoCompleto As String) As String
    On Error GoTo ErrHandler

    If Trim$(percorsoCompleto) = "" Then
        GetFolderPath = ""
        Exit Function
    End If

    Dim p As String
    p = percorsoCompleto
    Do While Len(p) > 1 And Right$(p, 1) = "\"
        p = Left$(p, Len(p) - 1)
    Loop

    Dim pos As Long
    pos = InStrRev(p, "\")
    If pos > 1 Then
        GetFolderPath = Left$(p, pos - 1)
    Else
        GetFolderPath = ""
    End If

    Exit Function
ErrHandler:
    GetFolderPath = ""
End Function

' ----------------------------------------------------------------------------
' §4.14  GetFileExtension
' ----------------------------------------------------------------------------
' Restituisce l'estensione del file (senza punto).
Public Function GetFileExtension(percorsoCompleto As String) As String
    On Error GoTo ErrHandler

    Dim nomeFile As String
    nomeFile = GetFileName(percorsoCompleto)

    Dim pos As Long
    pos = InStrRev(nomeFile, ".")
    If pos > 0 And pos < Len(nomeFile) Then
        GetFileExtension = Mid$(nomeFile, pos + 1)
    Else
        GetFileExtension = ""
    End If

    Exit Function
ErrHandler:
    GetFileExtension = ""
End Function

' ----------------------------------------------------------------------------
' §4.15  GetFileSizeFormatted
' ----------------------------------------------------------------------------
' Restituisce la dimensione del file in formato leggibile (bytes/KB/MB/GB).
Public Function GetFileSizeFormatted(filePath As String) As String
    On Error GoTo ErrHandler

    If Not FileEsiste(filePath) Then
        GetFileSizeFormatted = ""
        Exit Function
    End If

    Dim fad As WIN32_FILE_ATTRIBUTE_DATA
    Dim percorso As String
    percorso = NormalizzaPercorso(filePath)

    If GetFileAttributesExW(StrPtr(percorso), 0, fad) = 0 Then
        GetFileSizeFormatted = ""
        Exit Function
    End If

    Dim size As Double
    size = CDbl(fad.nFileSizeHigh) * 4294967296# + CDbl(fad.nFileSizeLow)

    Select Case size
        Case Is < 1024
            GetFileSizeFormatted = Format$(size, "0") & " bytes"
        Case Is < 1048576
            GetFileSizeFormatted = Format$(size / 1024, "0.0") & " KB"
        Case Is < 1073741824
            GetFileSizeFormatted = Format$(size / 1048576, "0.0") & " MB"
        Case Else
            GetFileSizeFormatted = Format$(size / 1073741824, "0.0") & " GB"
    End Select

    Exit Function
ErrHandler:
    GetFileSizeFormatted = ""
End Function

' ----------------------------------------------------------------------------
' §4.16  IsLongPath
' ----------------------------------------------------------------------------
' True se il percorso supera MAX_PATH (260) caratteri.
Public Function IsLongPath(percorso As String) As Boolean
    IsLongPath = (Len(percorso) > MAX_PATH)
End Function

' ----------------------------------------------------------------------------
' §4.17  NormalizzaPercorso
' ----------------------------------------------------------------------------
' Prepara un percorso per le API:
'  - risolve i path relativi in assoluti (GetFullPathNameW);
'  - aggiunge \\->\ per i path locali lunghi;
'  - converte \\server\share in \\->\UNC\server\share per gli UNC lunghi;
'  - se il prefisso \\->\ e' gia' presente, non lo duplica.
Public Function NormalizzaPercorso(percorso As String) As String
    On Error GoTo ErrHandler

    If Trim$(percorso) = "" Then
        NormalizzaPercorso = ""
        Exit Function
    End If

    Dim risultato As String
    risultato = Trim$(percorso)

    If Left$(risultato, 4) = LONG_PATH_PREFIX Then
        NormalizzaPercorso = risultato
        Exit Function
    End If

    If Not IsPercorsoAssoluto(risultato) Then
        risultato = RisolviPercorsoAssoluto(risultato)
    End If

    If Len(risultato) > MAX_PATH Then
        If Left$(risultato, 2) = "\\" Then
            risultato = LONG_UNC_PREFIX & Mid$(risultato, 3)
        Else
            risultato = LONG_PATH_PREFIX & risultato
        End If
    End If

    NormalizzaPercorso = risultato

    Exit Function
ErrHandler:
    NormalizzaPercorso = percorso
End Function




' ============================================================
' 04.18 - APRI PERCORSO CON CONTROLLO
' ============================================================
Public Function ApriPathConControllo( _
    ByVal percorso As String, _
    Optional ByVal avvisaPercorsoLungo As Boolean = True, _
    Optional ByVal annullaSeLungo As Boolean = False) As String

    On Error GoTo ErrHandler

    Dim percorsoControllo As String
    Dim percorsoLungo As Boolean
    Dim eCartella As Boolean
    Dim esitoApertura As Boolean
    Dim risposta As VbMsgBoxResult
    Dim tipoOggetto As String

    ' 01 - CONTROLLO DEL PERCORSO
    percorsoControllo = Trim$(percorso)

    If Len(percorsoControllo) = 0 Then
        ApriPathConControllo = "FILE NON TROVATO"
        Exit Function
    End If

    ' 02 - CONTROLLO LUNGHEZZA DEL PERCORSO
    percorsoLungo = (Len(percorsoControllo) > 260)

    If percorsoLungo Then

        If annullaSeLungo Then

            If avvisaPercorsoLungo Then
                MsgBox "Il percorso supera 260 caratteri." & vbCrLf & _
                       "L'apertura è stata annullata.", _
                       vbExclamation, "Percorso troppo lungo"
            End If

            ApriPathConControllo = _
                "PATH > 260: APERTURA ANNULLATA"
            Exit Function

        ElseIf avvisaPercorsoLungo Then

            risposta = MsgBox( _
                "Il percorso supera 260 caratteri." & vbCrLf & vbCrLf & _
                "Vuoi tentare comunque l'apertura?", _
                vbQuestion + vbYesNo + vbDefaultButton2, _
                "Percorso lungo")

            If risposta = vbNo Then
                ApriPathConControllo = _
                    "PATH > 260: APERTURA ANNULLATA"
                Exit Function
            End If

        End If

    End If

    ' 03 - VERIFICA SE SI TRATTA DI UNA CARTELLA
    eCartella = CartellaEsiste(percorsoControllo)

    If eCartella Then
        tipoOggetto = "PATH"
    Else
        ' 04 - VERIFICA SE SI TRATTA DI UN FILE
        If FileEsiste(percorsoControllo) Then
            tipoOggetto = "FILE"
        Else
            ApriPathConControllo = "FILE NON TROVATO"
            Exit Function
        End If
    End If

    ' 05 - APERTURA DEL FILE O DELLA CARTELLA
    esitoApertura = ApriFile(percorsoControllo)

    If esitoApertura Then

        If eCartella Then
            ApriPathConControllo = "PATH TROVATA"
        Else
            ApriPathConControllo = "FILE TROVATO"
        End If

    Else

        If eCartella Then
            If percorsoLungo Then
                ApriPathConControllo = _
                    "ERRORE: impossibile aprire la cartella (path > 260)"
            Else
                ApriPathConControllo = _
                    "ERRORE: impossibile aprire la cartella"
            End If
        Else
            If percorsoLungo Then
                ApriPathConControllo = _
                    "ERRORE: impossibile aprire il file (path > 260)"
            Else
                ApriPathConControllo = _
                    "ERRORE: impossibile aprire il file"
            End If
        End If

    End If

    Exit Function

ErrHandler:
    ApriPathConControllo = _
        "ERRORE: impossibile aprire il file"

End Function



' ============================================================================
' §5   FUNZIONI PRIVATE
' ============================================================================

' ----------------------------------------------------------------------------
' §5.1  TrimNull
' ----------------------------------------------------------------------------
' Tronca la stringa al primo Chr(0). Usa Long per non andare in overflow
' con path molto lunghi.
Private Function TrimNull(str As String) As String
    Dim pos As Long
    pos = InStr(str, Chr$(0))
    If pos > 0 Then
        TrimNull = Left$(str, pos - 1)
    Else: TrimNull = str
    End If
End Function

' ----------------------------------------------------------------------------
' §5.2  GetAttributiFile
' ----------------------------------------------------------------------------
' Restituisce dwFileAttributes di un file/cartella tramite
' GetFileAttributesExW. INVALID_FILE_ATTRIBUTES se non esiste.
Private Function GetAttributiFile(filePath As String) As Long
    On Error GoTo ErrHandler

    Dim percorso As String
    percorso = NormalizzaPercorso(filePath)

    Dim fad As WIN32_FILE_ATTRIBUTE_DATA
    If GetFileAttributesExW(StrPtr(percorso), 0, fad) <> 0 Then
        GetAttributiFile = fad.dwFileAttributes
    Else
        GetAttributiFile = INVALID_FILE_ATTRIBUTES
    End If

    Exit Function
ErrHandler:
    GetAttributiFile = INVALID_FILE_ATTRIBUTES
End Function

' ----------------------------------------------------------------------------
' §5.3  IsPercorsoAssoluto
' ----------------------------------------------------------------------------
' True se il percorso e' assoluto (UNC \\server\share o locale C:\...).
Private Function IsPercorsoAssoluto(percorso As String) As Boolean
    If Len(percorso) < 2 Then
        IsPercorsoAssoluto = False
        Exit Function
    End If

    If Left$(percorso, 2) = "\\" Then
        IsPercorsoAssoluto = True
        Exit Function
    End If

    If Mid$(percorso, 2, 1) = ":" Then
        IsPercorsoAssoluto = True
        Exit Function
    End If

    IsPercorsoAssoluto = False
End Function

' ----------------------------------------------------------------------------
' §5.4  RisolviPercorsoAssoluto
' ----------------------------------------------------------------------------
' Converte un percorso relativo in assoluto usando GetFullPathNameW.
Private Function RisolviPercorsoAssoluto(percorso As String) As String
    On Error GoTo ErrHandler

    Dim buf As String
    Dim n As Long

    buf = Space$(BUFFER_LONG_PATH)
    n = GetFullPathNameW(StrPtr(percorso), Len(buf), StrPtr(buf), 0)

    If n > 0 And n < Len(buf) Then
        RisolviPercorsoAssoluto = Left$(buf, n)
    Else
        RisolviPercorsoAssoluto = percorso
    End If

    Exit Function
ErrHandler:
    RisolviPercorsoAssoluto = percorso
End Function

' ----------------------------------------------------------------------------
' §5.5  CreaCartellaRicorsiva
' ----------------------------------------------------------------------------
' Crea la cartella e tutti i padri mancanti tramite CreateDirectoryW.
Private Function CreaCartellaRicorsiva(percorso As String) As Boolean
    On Error GoTo ErrHandler

    If Len(percorso) = 0 Then
        CreaCartellaRicorsiva = False
        Exit Function
    End If

    Dim fad As WIN32_FILE_ATTRIBUTE_DATA
    If GetFileAttributesExW(StrPtr(percorso), 0, fad) <> 0 Then
        CreaCartellaRicorsiva = ((fad.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0)
        Exit Function
    End If

    Dim p As String
    p = percorso
    Do While Len(p) > 1 And Right$(p, 1) = "\"
        p = Left$(p, Len(p) - 1)
    Loop

    Dim pos As Long
    pos = InStrRev(p, "\")
    If pos > 3 Then
        Dim padre As String
        padre = Left$(p, pos - 1)
        If Right$(padre, 2) <> ":\" And Right$(padre, 4) <> "UNC\" Then
            If Not CreaCartellaRicorsiva(padre) Then
                CreaCartellaRicorsiva = False
                Exit Function
            End If
        End If
    End If

    If CreateDirectoryW(StrPtr(p), 0) <> 0 Then
        CreaCartellaRicorsiva = True
    Else
        If GetFileAttributesExW(StrPtr(p), 0, fad) <> 0 Then
            CreaCartellaRicorsiva = ((fad.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0)
        Else
            CreaCartellaRicorsiva = False
        End If
    End If

    Exit Function
ErrHandler:
    CreaCartellaRicorsiva = False
End Function

' ----------------------------------------------------------------------------
' §5.6  EliminaCartellaRicorsiva
' ----------------------------------------------------------------------------
' Svuota ricorsivamente una cartella (FindFirstFileW / FindNextFileW) e poi
' la rimuove con RemoveDirectoryW.
Private Function EliminaCartellaRicorsiva(percorso As String) As Boolean
    On Error GoTo ErrHandler

    If Len(percorso) = 0 Then
        EliminaCartellaRicorsiva = False
        Exit Function
    End If

    Dim fad As WIN32_FILE_ATTRIBUTE_DATA
    If GetFileAttributesExW(StrPtr(percorso), 0, fad) = 0 Then
        EliminaCartellaRicorsiva = False
        Exit Function
    End If
    If (fad.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
        EliminaCartellaRicorsiva = False
        Exit Function
    End If

    Dim pattern As String
    pattern = percorso
    If Right$(pattern, 1) <> "\" Then pattern = pattern & "\"
    pattern = pattern & "*"

    Dim wfd As WIN32_FIND_DATAW
    Dim hFind As LongPtr
    Dim nome As String
    Dim sottoPercorso As String

    #If VBA7 Then
        hFind = FindFirstFileW(StrPtr(pattern), wfd)
    #Else
        hFind = FindFirstFileW(StrPtr(pattern), wfd)
    #End If

    If hFind <> INVALID_HANDLE_VALUE Then
        Do
            nome = Left$(wfd.cFileName, InStr(wfd.cFileName, Chr$(0)) - 1)
            If nome <> "." And nome <> ".." Then
                sottoPercorso = percorso
                If Right$(sottoPercorso, 1) <> "\" Then sottoPercorso = sottoPercorso & "\"
                sottoPercorso = sottoPercorso & nome

                If (wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0 Then
                    EliminaCartellaRicorsiva sottoPercorso
                Else
                    DeleteFileW StrPtr(sottoPercorso)
                End If
            End If
        Loop While FindNextFileW(hFind, wfd) <> 0
        FindClose hFind
    End If

    EliminaCartellaRicorsiva = (RemoveDirectoryW(StrPtr(percorso)) <> 0)

    Exit Function
ErrHandler:
    EliminaCartellaRicorsiva = False
End Function

' ----------------------------------------------------------------------------
' §5.7  GetFileAttributesText
' ----------------------------------------------------------------------------
' Restituisce gli attributi del file in formato testuale.
Private Function GetFileAttributesText(filePath As String) As String
    On Error GoTo ErrHandler

    Dim attribs As Long
    attribs = GetAttributiFile(filePath)

    If attribs = INVALID_FILE_ATTRIBUTES Then
        GetFileAttributesText = "Non disponibile"
        Exit Function
    End If

    Dim s As String
    s = ""
    If attribs And FILE_ATTRIBUTE_READONLY Then s = s & "Sola lettura, "
    If attribs And FILE_ATTRIBUTE_HIDDEN Then s = s & "Nascosto, "
    If attribs And FILE_ATTRIBUTE_SYSTEM Then s = s & "Sistema, "
    If attribs And FILE_ATTRIBUTE_DIRECTORY Then s = s & "Directory, "
    If attribs And FILE_ATTRIBUTE_ARCHIVE Then s = s & "Archivio, "
    If attribs And FILE_ATTRIBUTE_ALIAS Then s = s & "Alias, "
    If attribs And FILE_ATTRIBUTE_COMPRESSED Then s = s & "Compresso, "

    If s = "" Then
        GetFileAttributesText = "Normale"
    Else
        GetFileAttributesText = Left$(s, Len(s) - 2)
    End If

    Exit Function
ErrHandler:
    GetFileAttributesText = "Non disponibile"
End Function

' ============================================================================
' §6   NOTE DI MIGRAZIONE (rispetto alla versione originale)
' ============================================================================
'  - Tutte le API sono passate alla versione W (Unicode).
'  - GetFileAttributesEx ora riceve WIN32_FILE_ATTRIBUTE_DATA (non un Long).
'  - FileEsiste / CartellaEsiste usano sempre GetFileAttributesExW.
'  - ApriFile usa ShellExecuteW + explorer.exe per le cartelle lunghe.
'  - SelezionaFile / SalvaFile usano GetOpenFileNameW / GetSaveFileNameW.
'  - CopiaFile / SpostaFile / EliminaFile usano CopyFileW / MoveFileExW /
'    DeleteFileW (niente piu' FSO, che non supporta \\->\).
'  - CreaCartella / EliminaCartella usano CreateDirectoryW / RemoveDirectoryW
'    con ricorsione reale.
'  - GetFileName / GetFolderPath / GetFileExtension: solo stringhe.
'  - GetFileSizeFormatted / MostraProprietaFile / GetFileAttributesText:
'    GetFileAttributesExW invece di FSO.
'  - NormalizzaPercorso risolve i relativi e gestisce UNC.
'  - TrimNull usa Long (no overflow).
'  - Rimossa NormalizzaPercorsoFSO: nessuna dipendenza da FSO.
' ============================================================================


' ============================================================================
' §7   MODELLI DI UTILIZZO DELLA CLASSE SYSTEM
' ============================================================================
'
' In questa sezione sono raccolti i MODELLI di utilizzo della classe.
' I metodi qui presenti NON interferiscono con il resto della classe:
'  - non usano Me. e non conoscono il form chiamante;
'  - non scrivono su controlli;
'  - non mostrano MsgBox obbligatori (solo se esplicitamente richiesto);
'  - restituiscono una stringa di esito che il chiamante interpreta.
'
' INDICE DEI MODELLI:
'   §7.1  ApriPathConControllo (metodo pubblico)
'   §7.2  Esempio di chiamata dal form
' ============================================================================


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
' Private Sub PathFileAtto_s_Txt_DblClick(Cancel As Integer)
'     On Error GoTo ErrHandler
'
'     '// Dichiarazione variabili
'     Dim MyPath_s As String
'     Dim IDGestione_lng As Long
'     Dim Messaggio_s As String
'     Dim rst As DAO.Recordset
'     Dim clsFileSystem As CLASSE_clsFileSystem
'
'     '// Uscita immediata se il record è vuoto
'     If IsNull(Me.IDGestione_TXT) Or Me.IDGestione_TXT = "" Then Exit Sub
'
'     '// Lettura path dal TextBox
'     If Me.PathFileAtto_s_Txt.Value > "" Then
'         MyPath_s = Me.PathFileAtto_s_Txt.Value
'         Debug.Print "Path da aprire: " & MyPath_s
'     End If
'
'     '// Lettura IDGestione (per il bookmark finale)
'     If Me!IDGestione_TXT.Value > 0 Then
'         IDGestione_lng = Me!IDGestione_TXT.Value
'     Else
'         MsgBox "ERRORE_CODICE ID NON TROVATO", vbCritical
'         Exit Sub
'     End If
'
'     '// Istanza della classe
'     Set clsFileSystem = New CLASSE_clsFileSystem
'
'     '// Chiamata al modello di apertura con controllo > 260
'     '//   Parametri: (path, avvisaPercorsoLungo, annullaSeLungo)
'     '//   Avvisa se > 260 e, se l'utente conferma, tenta comunque l'apertura.
'     Messaggio_s = clsFileSystem.ApriPathConControllo(MyPath_s, True, False)
'
'     '// Rilascio dell'istanza
'     Set clsFileSystem = Nothing
'
'     '// Scrittura esito nei campi del form
'     Select Case Messaggio_s
'         Case "FILE NON TROVATO"
'             Me.RICERCA_FileAtto_s_Txt = "FILE NON TROVATO"
'             Me.DATAAGG_TXT = Date
'             Me.TIMEOPER_Txt = Now
'
'         Case "PATH TROVATA"
'             Me.RICERCA_FileAtto_s_Txt = "PATH TROVATA"
'             Me.DATAAGG_TXT = Date
'             Me.TIMEOPER_Txt = Now
'
'         Case "FILE TROVATO"
'             Me.RICERCA_FileAtto_s_Txt = "FILE TROVATO"
'             Me.DATAAGG_TXT = Date
'             Me.TIMEOPER_Txt = Now
'
'         Case "PATH > 260: APERTURA ANNULLATA"
'             Me.RICERCA_FileAtto_s_Txt = "PATH > 260: APERTURA ANNULLATA"
'             Me.DATAAGG_TXT = Date
'             Me.TIMEOPER_Txt = Now
'
'         Case "ERRORE: impossibile aprire la cartella (path > 260)"
'             Me.RICERCA_FileAtto_s_Txt = "ERRORE APERTURA (> 260)"
'             Me.DATAAGG_TXT = Date
'             Me.TIMEOPER_Txt = Now
'
'         Case "ERRORE: impossibile aprire il file (path > 260)"
'             Me.RICERCA_FileAtto_s_Txt = "ERRORE APERTURA (> 260)"
'             Me.DATAAGG_TXT = Date
'             Me.TIMEOPER_Txt = Now
'
'         Case "ERRORE: impossibile aprire la cartella"
'             Me.RICERCA_FileAtto_s_Txt = "ERRORE APERTURA"
'             Me.DATAAGG_TXT = Date
'             Me.TIMEOPER_Txt = Now
'
'         Case "ERRORE: impossibile aprire il file"
'             Me.RICERCA_FileAtto_s_Txt = "ERRORE APERTURA"
'             Me.DATAAGG_TXT = Date
'             Me.TIMEOPER_Txt = Now
'
'         Case Else
'             MsgBox "Si è verificato un errore imprevisto: " & Messaggio_s, vbCritical
'     End Select
'
'     '// Refresh del campo esito
'     Me.RICERCA_FileAtto_s_Txt.Requery
'
'     '// Refresh form + riposizionamento sul record corrente
'     Me.Refresh
'     Set rst = Me.RecordsetClone
'     rst.FindFirst "IDGestione = " & IDGestione_lng
'     If Not rst.NoMatch Then
'         Me.Bookmark = rst.Bookmark
'     End If
'     rst.Close
'     Set rst = Nothing
'
'     '// Pulizia variabili
'     MyPath_s = vbNullString
'     IDGestione_lng = 0
'     Exit Sub
'
' ErrHandler:
'     MsgBox "Errore durante l'apertura della cartella o file: " & Err.Description, vbCritical
'     Set clsFileSystem = Nothing
' End Sub
'
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




