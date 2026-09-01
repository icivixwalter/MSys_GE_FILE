VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_UTILITA_Msys_Frm21_IMPORTA_OGGETTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'# CLASSE_Form_UTILITA_Msys_Frm21_IMPORTA_OGGETTI.md
'//@VERSIONE_DEL_2025_02_12=COMPLETA E FUNZIONANTE



'*****************************************************************************************************//
'*  CLASSE DELLA FORM: Form_UTILITA_Msys_Frm21_IMPORTA_OGGETTI    @classe@form@importa@oggetti_(CODICE Della classe form importa oggetti)
'*
'*  ATTIVITA        :questa classe della @form@IMPORTA@PROGETTO viene utilizzata per importare
'*                  tutti gli oggetti (tabelle, query, form, report ecc..) che in questo caso appartengono
'*                  al progetto MSYS_ATTIVA_GEST_OGGETTI.mdb ma puo essere adattata per importare
'*                  altri oggetti che appartengono ad al progetti access.
'*                  VANTAGGIO: è dato dal fatto che gli oggetti tabelle query ecc... NON SONO SALVATI
'*                  in un tabella anchessa da importare ma sono salvati all'interno della classe in
'*                  in diverse collezioni precaricate. In questo puo spostare tutta form master
'*                  in qualsiasi progetto mdb senza preoccuparti di portare dietro tabelle oggetti, query
'*                  ed altre form di gestione. L'importazione degli oggetti RISULTA SEMPLICATA.
'*
'*  FUNZIONAMENTO   : Il funzionamento si basa su queste collezione di oggetti precaricati che  possono
'*                  essere sostituite:
'*
'*                  @PARAMETRO_01 = LA PATH ED IL FILE.MDB DOVE SI TROVANO GLI OGGETTI
'*
'*                  @PARAMETRO_02= COLLECTION DELLE FORMS DA IMPORTARE
'*                              I° COLLEZIONE : OGGETTI FORM PRECARICATI PRONTI PER L'IMPORTAZIONE
'*                              VAI AL CODICE: '//CODICE ----> COLLECTION_FORM_PRECARICATE
'*
'*  APRI ANALISI    : puoi aprire l'analisi di questa classe con questo comando
'*                  START "APRI ANALISI GESTIONE DB ESTERNO " call "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\Project_MSYSDbEst_GE_OBJECT_ANALISI.sublime-project"
'*
'*  LIBRERIE
'*      Note
'*              Per il funzionamento occorrono le seguenti librerie:
'*              Dim db As DAO.Database  =
'*                   Microsoft DAO 3.6 Object Library                           =
'*                          C:\Program Files (x86)\Common Files\microsoft shared\DAO\dao360.dll
'*                   Microsoft Visual Basic for Application Extensibility 5.3   =
'*                          C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.olb
'*                   Microsoft Activex Data Object 2.1 Library                  =
'*                          C:\Program Files (x86)\Common Files\System\ado\msado21.tlb
    
'*  ANALISI ACCESS  : L'analisi generale di access si apre con questo comando
'*                      START "APRI ANALISI GENERALE DI ACCESS" call "c:\Casa\LINGUAGGI\ACCESS\Project_ACCESS.sublime-project"
'*  FAQ             :
'*                  COME IMPORTARE GLI OGGETTI TABLE, QUERY, FORM IN UN PROGETTO?
'*                  CLASSE IMPORTA PROGETTI IN UN DB ESTERNO
'*                  @FORM@IMPORTA@OGGETTI_(Questa form permette l'importazione degli oggetti table, query ecc. in modo autonomo)
'*
'*                           Verifica se il modulo esiste nel database esterno
'*  FILTRO IMPORTA MODULI   = Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = -32761 AND Name = '" & myNameModule_s & "'", dbOpenSnapshot)

'*                           '//FILTRO CONTROLLO ESISTENZA FORM
'*  FILTRO IMPORTA FORM     = Set rs = db.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=-32768 AND Name='" & objName & "'")
                            '//FILTRO CONTROLLO ESISTENZA MACRO
'*  FILTRO IMPORTA MACRO    = Set rs = db.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=-32766 AND Name='" & objName & "'")
                            '//FILTRO CONTROLLO ESISTENZA QUERY
'*  FILTRO IMPORTA QUERY    =  Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = 5 AND Name = '" & qryName & "'", dbOpenSnapshot)
                            '//FILTRO CONTROLLO ESISTENZA REPORT
'*  FILTRO IMPORTA REPORT   =  Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = -32764 AND Name = '" & varItem & "'", dbOpenSnapshot)

'*  AGGIORNAMENTO DEL 12.03.2025 fino qui mancano le tabelle..


                            '//FILTRO CONTROLLO ESISTENZA TABELLE   **** FA FINIRE
'*  FILTRO IMPORTA TABELLE   =  Set rs = dbCurrent.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = 6 AND Name = '" & varItem & "'", dbOpenSnapshot)


'//LE ROUTINE E LE FUNZIONI CREATE:
'//
'//@30@ROUTINE@APRI@FILE                            = ROUTINE DI APERTURA DEL FILE
'*
'*****************************************************************************************************//
Option Compare Database
    
'//DIM LE VARIABILI
'//-----------------------------------------------------------------------------------//
    '//APRI FILE
   
' Dim Tables_Collection As Collection
   
  
   
    '//DATABASE PATH SORGENTE, CONTATORI
    '// SE da errore caricare c:\Casa\LINGUAGGI\ACCESS\DLL_LIBRERIE\dao360.dll _
                              c:\Casa\LINGUAGGI\ACCESS\DLL_LIBRERIE\VBE6EXT.olbVBE6EXT _
                              c:\Casa\LINGUAGGI\ACCESS\DLL_LIBRERIE\VBE6EXT.msado21.tlb
    Dim db As DAO.Database                                  ' database DAO
    Dim ws As DAO.Workspace
    Dim sourceDB As DAO.Database
    'DAO
    Dim DaoDB As DAO.Database
    Dim dbCurrent As DAO.Database

    Dim DaoWks As DAO.Workspace
    Dim DaoRs As DAO.Recordset
    Dim DaoRs_Lett As DAO.Recordset
    Dim tdf As DAO.TableDef

    
    'ADO
    Dim ADODB As Database
    Dim AdodaoRs As Recordset
    
    'Contatori
    Dim iCount As Integer
    Dim iTOTcount As Integer
    Dim dbl_count As Double
    

    
    '//DIM input box
    Dim Message As Variant
    Dim Title As Variant
    Dim Default As Variant
    Dim MyValue As Variant
    Dim MyMsg_V As Variant              '//risultato se 0 exit 1 continua procedura se annulla stringa = ""
    
    
    'Variabili generali
    Dim i As Integer
    Dim Str1 As String
    Dim Int1 As Integer
    Dim Int2 As Integer
    Dim Int3 As Integer
    Dim Lng1 As Long
    Dim Dbl1 As Double
    Dim Bln1 As Boolean
    Dim Vv1 As Variant
    Dim obj1 As Object

    
    'Gestione parametri comandi
    Dim sxParamCmd_3 As String
    
    
    'Larghezza e numero di colonna
    Dim sLarg_Col As String
    Dim iNum_Col As Integer
    
    
    '//ERRORI PROCEDURA_FUNCTION O ROUTINE
    Dim ProceduraMessaggioErrore_s As String
    Dim ProceduraAttivaEseguita_s As String
    
    
    
    '//RECUPERO PARAMETRO DA TABELLA OGGETTI
    '//File da ricercare e la path
    Dim MyFile_s As String
    Dim MyPath_s As String
            
        
    ' // DIM variabili procedura importa e collega DLL  @DIM@DLL
    Dim PathDb_s As String
    
    ' // DIM MATRICI DINAMINCHE: senza valore - Array globali per i nomi degli oggetti, SOSTITUITO DA MATRICI STATICHE CON VALORE _
        A MATRICI DIMINAMICHE senza valore per essere ridimensionate : QueryArray() ecc..
    Dim queryArray() As String
    Dim TableArray() As String
    Dim ModuleArray() As String
    Dim MacroArray() As String

    
    '//DIM LE VARIABILI DEL PROGETTO ORIGINALE CURRENT PROJECT DEL DB ESTERNO DA LEGGERE
    '//........................................................................//
         Dim myCurrentProject As String
         Dim myOriginaleProject As String
          '//PATH E FILE DOVE IMPORTARE
         Dim sourceDBPath_s As String                       '//path del db esterno
         
         Dim PROGETTO_s As String                           '//nome del progetto attivato
         
    '//........................................................................//
    
    
    
   '//@DIM@COLLECTION@TABLE COLLEZIONI TABELLE
   '//...................................
      
    '//DIM IMPORTA TABELLE
    Dim tableName As String
    Dim tblCiclo As TableDef
    Dim mytableName_s As String
    Dim Tables_Collection As Collection              '//todo: doppia collezione di tabelle eliminare Tables_Collection
    Dim tablesCollection As Collection
    
    Dim varItem As Variant
    '//i contatori delle tabelle
    Dim tableExistsInCollectionTables_Funct_b As Boolean
    Dim importedTablesCount_i As Integer
    Dim deletedTablesCount_i As Integer
    Dim deletedTables_s As String
       
   

    
   
    '//...................................

   '//@DIM@COLLECTION@FORM COLLEZIONI FORM
   '//...................................
     
        '//DIM IMPORTA FORM
        Dim myforms_Name_s As String                      '//nome form da importare
        Dim forms_ToImport As Variant                    '//calcolo form da importare
        '//DIM della collezione form da importare
        Dim formsCollection As Collection                '//collezione delle form prestabilita
        
        '// i contatori delle form
        Dim deleted_Forms_s As String
        Dim deletet_FormsCount_i As Integer
        Dim imported_Forms_s As String
        Dim imported_FormsCount_i As Integer
    
        
   '//...................................


   '//@DIM@COLLECTION@QUERY COLLEZIONI QUERY
   '//...................................
     
        '//DIM IMPORTA QUERY
        Dim myName_Queryes_s As String                         '//nome query da importare
        Dim queryes_ToImport As Variant                      '//calcolo query da importare
        '//DIM della collezione form da importare
        Dim queryes_Collection As Collection                   '//collezione delle query prestabilita
        
        '// i contatori delle form
        Dim deleteName_Queryes_s As String
        Dim delete_QueryesCount_i As Integer
        Dim importedName_Queryes_s As String
        Dim imported_QueryesCount_i As Integer
    
        
   '//...................................




   '//@DIM@COLLECTION@REPORT COLLEZIONI REPORTS
   '//...................................
     
        '//DIM IMPORTA QUERY
        Dim myReports_Name_s As String                         '//nome query da importare
        Dim reports_ToImport As Variant                      '//calcolo REPORT da importare
        '//DIM della collezione form da importare
        Dim reports_Collection As Collection                   '//collezione delle REPORT prestabilita
        
        '// i contatori delle form
        Dim deleted_Reports_s As String
        Dim deletet_ReportesCount_i As Integer
        Dim imported_Reports_s As String
       
    
        
   '//...................................







    '//@DIM@COLLECTION@MACRO COLLEZIONI MACRO
    '//...................................
      
         '//DIM IMPORTA MACRO
         Dim myMACROS_Name_s As String                        '//nome MACRO da importare
         Dim macros_ToImport As Variant                      '//calcolo MACRO da importare
         '//DIM della collezione MACRO da importare
         Dim Macros_Collection As Collection                 '//collezione delle MACRO prestabilita
         
         '// i contatori delle MACRO
         Dim deleted_Macro_s As String
         Dim deletet_MacroCount_i As Integer
         Dim imported_Macro_s As String
         Dim imported_MacroCount_i As Integer
            
        '//DIM oggetti Access del db interno
        Dim obj As AccessObject
        Dim imported_macros_Count_i As Integer
        
        '//LA MATRICE DELLE MACRO INTERNE NEL DB
        Dim MATRICE_MACRO_Names_s() As String             'matrice di MACRO CON LE () SIGNIFICA MATRICE
    
    
         
    '//...................................

    


    '//@DIM@COLLECTION@MODULI COLLEZIONI MODULES
    '//...................................
      
         '//DIM IMPORTA MODULI
         Dim myModules_Name_s As String                         '//nome MODULI da importare
         Dim modules_Tot_Import As Variant                      '//calcolo MODULI da importare
         '//DIM della collezione form da importare
         Dim Modules_Collection As Collection                   '//collezione delle MODULI prestabilita
         
         '// i contatori delle form
         Dim deleted_Modules_s As String
         Dim deletet_Modules_Count_i As Integer
         Dim imported_Modules_s As String
         Dim imported_Modules_Count_i As Integer
     
         
    '//...................................

        
        
        
        'Le variabili generiche
        Dim sSql As String                                          '//Stringa di estrazione SQL
        Dim sSql_Lett As String                                     '//Stringa di estrazione SQL



'//DIM LE VARIABILI  *** FINE ***
'//-----------------------------------------------------------------------------------//




'//DIM LE COLLECTION DI OGGETTI PER LE UTILITA
'//-----------------------------------------------------------------------------------//
'//@UTILITA@DIM@COLLETION_(LE COLLEZIONI DI OGGETTI DA CARICARE PER LA GESTIONE DELLE UTILITA)
  Public coll_Tables_UTILITA As Collection    '//....vedi  @UTILITA@IMPORTA@COLLETION ...
  Public coll_Queryes_UTILITA As Collection
  Public coll_Modules_UTILITA As Collection

'//-----------------------------------------------------------------------------------//



'//DIM PER LA BARRA DI PROGRESSIONE @DIM@PRORESS@BAR
'//-----------------------------------------------------------------------------------//
    
    '//@PROGRESS@BAR@OPEN@DIM                       = DIM variabilie oggetto form
    Dim frmProgress As Form_UTILITA_Msys_Frm21_frmProgressBar          ' Dichiarazione globale per il riferimento alla form
    Dim TipoImport_s As String                      '//oggetto tipo di importazione (Stringa che indica oggetti FORM, REPORT,....)
    Dim Twips_dbl As Double                         '//@DIM@TWIPS = il numero di twips o cm base per la progressione in %
    Dim PercIncr_dbl As Double                      '//@DIM@%@INCREMENTO@PROGRESS@BAR = LA VARIABILE che rappersenta la % di incremento DELLA BARRA DI PROGRESSIONE

    Dim ProgressPercentTwips_dbl As Double
    Dim frm_obj As Form_UTILITA_Msys_Frm21_frmProgressBar ' <-- Cast esplicito al tipo corretto
    Dim NroCollection_i As Integer
      
    Dim TotPerc_dbl As Double                   '// E' la percentuale di incremnento
    
'//-----------------------------------------------------------------------------------//





'//PARAMETRO PATH E FILE DOVE SI TROVANO GLI OGGETTI
'//-----------------------------------------------------------------------------------//

Public Function PathFile_s_pFunct() As String

On Error GoTo CollectionForm_PFunct_Err

        
        '//recupero la path solo se valorizzata altrimenti è null
        If Me.sourceDBPath_s_Txt > "" Or IsNull(Me.sourceDBPath_s_Txt) = False Then
                 
            PathFile_s_pFunct = Me.sourceDBPath_s_Txt
        
        Else
            '//se non esiste la path inserisco null
            PathFile_s_pFunct = ""
        End If
        
        


'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionForm_PFunct:
    Exit Function

CollectionForm_PFunct_Err:
    MsgBox Error$
    Resume CollectionForm_PFunct

End Function
'//-----------------------------------------------------------------------------------//



'//************************************************************************************************************************//
'//***                              I PARAMETRI DELLE COLLECTION,
'//***                              COLLECTION OBJECT   *** INIZIO ***
'//************************************************************************************************************************//



'//LA COLLEZIONE DELLE FORM DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//
'//NOTE: in questa funzione precarico la collezione di oggetti form _
    @PARAMETRO_02= COLLECTION DELLE FORMS DA IMPORTARE
    
'//PRECARICO LA COLLECTION FORM DA IMPORTARE
Public Function CollectionForm_PFunct() As Collection       '//restituisco una collection
    

On Error GoTo CollectionForm_PFunct_Err
        


    Dim frm As Object             '//OGGETTO FORM DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_FORM_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI FORM   @importa@collection@oggetti@form_(@form la collezione delle form)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI FORM  CHE APPARTENGONO al progetto _
            MSYS_ATTIVA_GEST_OGGETTI.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le form del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della form di _
            importazione, in quanto le form non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alla form di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe FORM con i dati relativi in questo caso agli oggetti FORM. _


    '//@OGGETTI@FORM@PRECARICATI
    '//popola la collection delle FORM di DA IMPORTARE
    ' Aggiungi i nomi delle FORM o ELIMINA O COMMENTA LE FORM DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti form DI PROGETTI DIVERSI inserisci i nome delle form _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        Set objectFORMS = New Collection
        '//Aggiungo gli elementi della collezione
        
            
    '//@carico@collection@form_(carico la @collezione del @progetto scelto)_(@collezione@form)
     '//2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
            Le collection viene scelta in base combinata 2 che individua il progetto _
            al posto della combinata 1 e la sostituzione riguarda tutte le collection.
                    
     Select Case Me.Cmb_02_txt
        
        '//@COLLEZIONE@FORM@OGGETTI@IMPEGNI
        '//@IMPEGNI@FORM@UTILITA_(LA @collection DELLE @FORM per la gestione degli impegni)
        Case "IMPORTA_OGGETTI_IMPEGNI"
                 
            objectFORMS.Add "LLPP_IMPEGNI_Frm_S99_01_{@====================================@}"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_M01_OGGETTI_GEST"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S01_OGGETTI_FORMS"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S02_OGGETTI_QUERY"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S02_OGGETTI_QUERY_SOTTOSTANTE"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S03_OGGETTI_MODULI"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S04_OGGETTI_REPORT"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S05_OGGETTI_TABLE"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S06_ELENCO_TUTTI"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S07_{@=================================@}"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm17_}---------------------------------------@"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm17-00_M01_GE_CUP"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm17-01_S01_GROUPxANNO_CUP"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmM01_}-------------------------------------@"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmM01_01_GESTIONE_IMPEGNI_MASTER"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS01_12_GRUPPO_IMPEGNI_ANNO"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS01_13_GRUPPO_IMPEGNI_CLASSIFICAZIONE"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS01_13_GRUPPO_IMPEGNI_CLASSIFICAZIONE_OLD"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS01_14_00_GE_ELENCO_IMPEGNI"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS01_14_ELENCO_IMPEGNI"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS01_15_ELENCO_DI_SPESA"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS01_16_ELENCO_DI_SPESA_Totali"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS17-02_S01_CUP"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS99_}--------------------------------------@"
            objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS99_UTILITA_IMPORTA_OGGETTI"
            objectFORMS.Add "PROGETTI_Msys_{@==============GESTITI_DA_MSYS============@}_Tb00"
            objectFORMS.Add "PROGETTI_Msys_Tb00FrmM01_GE_PROGETTI"
            objectFORMS.Add "PROGETTI_Msys_TB01Frm00_}--------------------------------------@"
            objectFORMS.Add "PROGETTI_Msys_TB01Frm01_01_PROJECT"
            objectFORMS.Add "PROGETTI_Msys_TB01Frm01_01_PROJECT_TMP"
            objectFORMS.Add "PROGETTI_Msys_TB02Frm00_}--------------------------------------@"
            objectFORMS.Add "PROGETTI_Msys_TB02Frm01_01_TIPO_OGGETTI"
            objectFORMS.Add "PROGETTI_Msys_TB03Frm01_01_ELENCO_ObjProgetti"
            objectFORMS.Add "PROGETTI_Msys_TB03Frm01_02_ObjProgetti_FORM=COD_PROGETTO_s"
            objectFORMS.Add "PROGETTI_Msys_TB03Frm01_02_ObjProgetti_MACRO=COD_PROGETTO_s"
            objectFORMS.Add "PROGETTI_Msys_TB03Frm01_02_ObjProgetti_QUERY=COD_PROGETTO_s"
            objectFORMS.Add "PROGETTI_Msys_TB03Frm01_02_ObjProgetti_REPORT=COD_PROGETTO_s"
            objectFORMS.Add "PROGETTI_Msys_TB03Frm01_02_ObjProgetti_TABLE=COD_PROGETTO_s"
            
       
       
        '//@COLLEZIONE@FORM@OBJECT
        '//@collection@FORM@MSYS@OBJECT
        Case "MSYS_ATTIVA_GEST_OGGETTI.mdb"
            '//@popolo@collection@form@system_(delle form di sistema, inserisco nella @collezione)
            objectFORMS.Add "MsysDbEstTb01Frm01_Mts01_}-------------------------------------@"
            objectFORMS.Add "MsysDbEstTb01Frm01_Mts01_GE_OBJECT"
            objectFORMS.Add "MsysDbEstTb01Frm01_Stf01_OBJECT"
            objectFORMS.Add "MsysDbEstTb01Frm01_Stf01_OBJECT_TMP"
            objectFORMS.Add "MsysDbEstTb05Frm01_Mt00_}--------------------------------------@"
            objectFORMS.Add "MsysDbEstTb05Frm01_Stf01_DB_ESTERNO"
            objectFORMS.Add "UTILITA_MsysDbEst_Frm12_}--------------------------------------@"
            objectFORMS.Add "UTILITA_MsysDbEst_Frm12_SERVIZI_TUTTI"
        
        '//@COLLEZIONE@FORM@TREEVIEW
        '//@collection@FORM@TREEVIEW
        Case "IMPORTA_OGGETTI_TREEVIEW"
            
            objectFORMS.Add "Tree_Frm_M00_00_{@============================================@}"
            objectFORMS.Add "TREE_VIEW_FRM01"
            objectFORMS.Add "TreeView_FIL_}-------------------------------------------@"
            '//2024.08.06_CORREZIONE_VOCE_(corretto errori della collection DA objectFORMS.Add "TreView_FILE" A objectFORMS.Add)
            objectFORMS.Add "TreeView_FILE"
            objectFORMS.Add "TreeView_FrmS01_GE_TREEVIEW"
            objectFORMS.Add "TreeView_FrmS01_TREE"
            objectFORMS.Add "TreeView_FrmS02_StampaDiControllo"
            objectFORMS.Add "TreeView_FrmS03_StampaRisultati"
            objectFORMS.Add "TreeView_Tb01Frm01_}-------------------------------------------@"
            objectFORMS.Add "TreeView_Tb01Frm01_M01_GE_ALBERO"
            objectFORMS.Add "TreeView_Tb01Frm01_M01_RADICE"
            objectFORMS.Add "TreeView_Tb01Frm02_S01_TRONCO"
            objectFORMS.Add "TreeView_Tb01Frm03_S01_RAMI"
            objectFORMS.Add "TreeView_Tb01Frm04_S01_FOGLIE"
            objectFORMS.Add "TreeView_UTILITA_Frm12_}------------------------------@"
            objectFORMS.Add "TreeView_UTILITA_Frm12_SERVIZI_TUTTI"
            
            
        '//@COLLEZIONE@FORM@MSYS_ATTIVITA_GEST_OGGETTI
        '//@COLLECTION@FORM@MSys = form del progetto msys
        Case "MSYS_ATTIVA_GEST_OGGETTI.mdb"
            
            objectFORMS.Add "AA_MSys__}-----------------------------------------------@MASTER"
            objectFORMS.Add "AA_MSys_FORM_M01_GESTIONE_OGGETTI_MASTER"
            objectFORMS.Add "MSys_TABELLE_Frm01_S01_TUTTE"
            objectFORMS.Add "MSys_QUERY_Frm01_S01_TUTTE"
            objectFORMS.Add "MSys_FORMS_Frm01_S01_TUTTE"
            objectFORMS.Add "MSys_REPORT_Frm01_S01_TUTTI"
            objectFORMS.Add "MSys_MODULI_Frm01_S01_TUTTI"
            objectFORMS.Add "MSys_MACRO_Frm01_S01_TUTTE"
            objectFORMS.Add "MSys_DF10_DATABASE_Frm01_M0_ESPORTA"
            objectFORMS.Add "MSys_TABELLE_LINK_Frm01_S01_TUTTE"
            objectFORMS.Add "MSys_TUTTE_LE_PROPRIETA_Frm01_S01"
            objectFORMS.Add "Msys_DF11_PROGETTI_Frm00_M01_GESTIONE"
            objectFORMS.Add "Msys_DF11_PROGETTI_Frm01_M01_Obj"
            objectFORMS.Add "MSys_SERVIZI_Frm01_S01_TUTTI"
            objectFORMS.Add "MSys_DF10_DATABASE_Frm01_S01_ELENCO"
            objectFORMS.Add "MSys_TUTTI_PROG_FrmM01_01_PROGETTI"
            objectFORMS.Add "MSys_TUTTI_PROG_VIS_PAGINE_FrmS01_01_Carica"           '//AGGIUNTO
            objectFORMS.Add "Msys_DF11_PROGETTI_Frm02_S01_TIPO_OGGETTI"             '//AGGIUNTO
            objectFORMS.Add "MSys_ELENCO_ObjProgetti_Frm01_S01_Tutti"               '//AGGIUNTO

        '//@COLLEZIONE@FORM@ESPROPRI
        '@IMPORTA@FORM@EPROPRI
        '//@COLLECTION@FORM@ESPROPRI = form del progetto ESPROPRI
        Case "IMPORTA_OGGETTI_ESPROPRI"
            '//@COLLECTION@@FORM@ESPROPRI  - 2024.08.06_COLLECTION@FORM@ESPROPRI_(inserito la collezione form per gli espropri)
            objectFORMS.Add "LLPP_TB09Frm01_}-----------------------------------------------@"
            objectFORMS.Add "LLPP_TB09Frm01_M01_GE_ESPROPRI"
            objectFORMS.Add "LLPP_TB09Frm01_S01_ESPROPRI"
            
            objectFORMS.Add "UTILITA_EPROPRI_Frm19_}----------------------------------------@"
            objectFORMS.Add "UTILITA_ESPROPRI_Frm20_SERVIZI_TUTTI"
            
        '//@COLLEZIONE@FORM@DLL
        '//COLLECTION@IMPORTA@FORM@DLL@COLECTION   = @IMPORTA@FORM@DLL_(definisco gli oggetti collection FORM per le @liberie@DLL)
        Case "IMPORTA_OGGETTI_MSYS_DLL"
            '//2024.08.05_MODIFICHE_IMPORTA_DLL_(per l'importazione delle librerie DLL del progetto mysys inserito il blocco)
            '//@COLLECTION@DLL@FORM
            objectFORMS.Add ""
            objectFORMS.Add ""
            objectFORMS.Add ""
            objectFORMS.Add ""
        
        '//@COLLEZIONE@FORM@GE@ESPROPRI
        '//COLLECTION@IMPORTA@FORM@ESPROPRI   = @IMPORTA@FORM@DLL_(definisco gli oggetti collection FORM per GLI @ESPROPRI)
        Case "GE_ESPROPRI.mdb"
            '//@COLLECTION@ESPROPRI@FORM
            objectFORMS.Add "LLPP_TB09Frm01_}-----------------------------------------------@"
            objectFORMS.Add "LLPP_TB09Frm01_S01_ESPROPRI"
            objectFORMS.Add "UTILITA_EPROPRI_Frm19_}----------------------------------------@"
            objectFORMS.Add "UTILITA_ESPROPRI_Frm20_SERVIZI_TUTTI"
    
    '//@COLLEZIONE@FORM@CDM@GIUR
    '//@COLLECTION@FORM@CDM_(le collezioni del condominio)
    '//-------------------------------------------------------------------------------------//
        Case "IMPORTA_OGGETTI_CDM_GIUR"
                '//@CONDOMINIO@TABELLE@GIURIDICA
                objectFORMS.Add "GIUR_Frm00_{@==================GIURIDICA======================@}"
                objectFORMS.Add "GIUR_FrmTb01_}-------------------------------------------------@"
                objectFORMS.Add "GIUR_Tb01Frm01_M01_GEST_GIURIDICA"
                objectFORMS.Add "GIUR_Tb01Frm01_M01_PARAMETRI_TUTTI"
                objectFORMS.Add "GIUR_Tb01Frm01_S01_PARAMETRI_CONDOMIN"
                objectFORMS.Add "GIUR_Tb02Frm01_S01_PARAMETRI_GENERALI"
                objectFORMS.Add "GIUR_Tb02Frm01_S01_PARAMETRI_GENERALI_TMP"
                            
                '//@COLLEZIONE@FORM@CDM@CONSUNTIVO
                '//@CONDOMINIO@TABELLE@CONSUNTIVO
        Case "IMPORTA_OGGETTI_CDM_CONS"
                objectFORMS.Add "CONS_FrmTB08_}-------------------------------------------------@"
                objectFORMS.Add "CONS_FrmTB08_M01_GE_CONSUNTIVO"
                objectFORMS.Add "CONS_FrmTB08_S15_01_TOT_CONSUNT"
                objectFORMS.Add "CONS_FrmTB08_S15_02_PARZ_X_CODICE"
                
                '//@COLLEZIONE@FORM@CDM@ANAGRAFICA
                '//@CONDOMINIO@TABELLE@ANAGRAFICA
        Case "IMPORTA_OGGETTI_CDM_ANAGR"
                objectFORMS.Add "ANAGR_FrmTb01_}------------------------------------------------@"
                objectFORMS.Add "ANAGR_FrmTb01_M01_01_GEST_ANAGRAFICA"
                objectFORMS.Add "ANAGR_FrmTb01_S01_CONDOMINI"
                objectFORMS.Add "ANAGR_FrmTb01_S02_MILL_ASS"
                objectFORMS.Add "ANAGR_FrmTb01_S03_MILL_PRES"
                objectFORMS.Add "ANAGR_FrmTb01_S04_VOTAZIONI=ASTENUTI"
                objectFORMS.Add "ANAGR_FrmTb01_S04_VOTAZIONI=NO"
                objectFORMS.Add "ANAGR_FrmTb01_S04_VOTAZIONI=SI"
                
                '//@COLLEZIONE@FORM@CDM@GE@CONDOMINIO
                '//@CONDOMINIO@FORM@GECOND
        Case "IMPORTA_OGGETTI_CDM_GECOND"
                    objectFORMS.Add ""
                    objectFORMS.Add ""
                    
                '//@COLLEZIONE@FORM@CDM@PREVENTIVO
                '//@CONDOMINIO@FORM@PREVENTIVO
        Case "IMPORTA_OGGETTI_CDM_GE_PREV"
                    objectFORMS.Add "PREV_FrmTB09_}-------------------------------------------------@"
                    objectFORMS.Add "PREV_FrmTB09_M15_CONTROLLO_PREVENTIVO"
                    objectFORMS.Add "PREV_FrmTB09_S01_TOTALE_PREVENTIVO"
                    objectFORMS.Add "PREV_FrmTB09_S02_PREVENTIVO"
                    objectFORMS.Add "PREV_FrmTB09_S15_01_TOTALI_SOTTOTITOLO"
                    objectFORMS.Add "PREV_FrmTB09_S15_02_TOTALI_PER_TITOLO_DI_SPESA"
                    '//TODO: da finire l'inserimento delle form del preventivo condominiale
                    
                '//@COLLEZIONE@FORM@CDM@MILLESIMI
                '//@CONDOMINIO@FORM@FORM@MILLESIMALI
        Case "IMPORTA_OGGETTI_CDM_GE_TABELLE"
                objectFORMS.Add "TB01_MILL_Frm_}------------------------------------------------@"
                objectFORMS.Add "TB01_MILL_Frm01_S01_PROPR_GENERALI"
                
                '//@COLLEZIONE@FORM@CDM@UTILITA
                '//@CONDOMINIO@FORM@FORM@UTILITA
        Case "IMPORTA_OGGETTI_CDM_GE_UTIL"
                objectFORMS.Add "UTIL_FrmTB09_}-------------------------------------------------@"
                objectFORMS.Add "UTIL_FrmTB09_ATTI_M01_GEST_ATTI"
                objectFORMS.Add "UTIL_FrmTB09_ATTI_S01_ATTI"
                objectFORMS.Add "UTIL_FrmTB09_S15_GROUP_ANNO"
                    
                '//@COLLEZIONE@FORM@UTILITA@IMPEGNI
                '//@IMPEGNI@FORM@FORM@UTILITA
        Case "IMPORTA_OGGETTI_IMPEGNI"
                objectFORMS.Add "LLPP_IMPEGNI_Frm_S99_01_{@====================================@}"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_M01_OGGETTI_GEST"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S01_OGGETTI_FORMS"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S02_OGGETTI_QUERY"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S02_OGGETTI_QUERY_SOTTOSTANTE"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S03_OGGETTI_MODULI"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S04_OGGETTI_REPORT"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S05_OGGETTI_TABLE"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S06_ELENCO_TUTTI"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S07_{@=================================@}"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmM01_01_GESTIONE_IMPEGNI"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01Frm01_S01_OGGETTI_FORMS"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS01_12_GRUPPO_IMPEGNI_ANNO"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS01_13_GRUPPO_IMPEGNI_CLASSIFICAZIONE"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS01_14_ELENCO_IMPEGNI"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS01_15_ELENCO_DI_SPESA"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS01_16_ELENCO_DI_SPESA_Totali"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS99_}--------------------------------------@"
                objectFORMS.Add "LLPP_IMPEGNI_Tb01FrmS99_UTILITA_IMPORTA_OGGETTI"

            '//TODO: continuare con le altre form del condominio.....

    '//-------------------------------------------------------------------------------------//
    
        Case Else
            MsgBox "ATTENZIONE! NON CI SONO FORM PER QUESTO PROGETTO: " & Me.Cmb_01_txt, vbCritical, "AVVISO)"
                '//2024.08.10 =COLLECTION=0- prima di uscire imposto la collezione per evitare di restiuire null e poi esco dalla routine.
                Set CollectionForm_PFunct = objectFORMS
                '//SE NON CI SONO OGGETTI USCITA DALLA ROUTINE
                GoTo CollectionForm_PFunct_Exit
                


    End Select
        
    
        ' Inizializza le collezioni per FORM COLLEGATE
        Set connectedFORMS = New Collection
      
    '// *** FINE ****
    '//I° COLLEZIONE DI FORM DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
    
    
        '//CONTROLLO DELLE FORM PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
        '//NOTE: queste sono le form da carica nel nuovo progetto.
                        
                        '//RESET CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.TXT_OBJECT_ESTERNI.Requery
        
                                    
                  '//IMPOSTO LA PATH + FILE.MDB + TESTATA + OGGETTI VISUALIZZAZIONE DELLA CASELLA TXT
                  '//--------------------------------------------------------------------------------------------//
                  '//Note: @imposto@path_(imposto la path degli oggetti da recupeare, _
                           preparo la testata per la casela di testo, preparo gli oggetti e visualizzo tutto _
                           nella casella di testo.
                    
                    
                    '//RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI
                        Str1 = ""
                        Str1 = Str1 & "CONTROLLO PATH E DATABASE DA CUI RECUPERARE GLI OGGETTI" & vbCrLf
                        Str1 = Str1 & "-----------------------------------------------------------------" & vbCrLf
                        Str1 = Str1 & Me.sourceDBPath_s_Txt & vbCrLf
                        
                        '//INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO
                        Me.TXT_OBJECT_ESTERNI = Str1
                        Me.TXT_OBJECT_ESTERNI.Requery

                            '//PREPARO LA TESTATA DELLE COLLECTION PER IL II° INSERIMENTO
                            Str1 = Str1 & vbCrLf & "                          COLLECTION FORM DA IMPORTARE : " & vbCrLf
                            Str1 = Str1 & "//=========================================================================//" & vbCrLf
                            ' Stampa
                            Debug.Print
                            Debug.Print
                            Debug.Print Str1
                            
                            '//iterazione nella collezione per il controllo
                            For Each objfrm In objectFORMS
                                Debug.Print "LA FORM da importare nel progetto: " & objfrm
                                Str1 = Str1 & objfrm & vbCrLf
                            Next objfrm
                            
                            
                  '//--------------------------------------------------------------------------------------------//
                        
        '// *** FINE ***
        '//CONTROLLO DELLE FORM PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
                        
                        
                        ' Restituisci la collezione
                        Set CollectionForm_PFunct = objectFORMS
                            
                    '//CHIAMO LA PROCEDURA UNICA DI STAMPA E CARICAMENTO NELLA CASELLA DI TESTO _
                            parametro_01= la collezione, parametro_02 = il tipo
                      Vv1 = Stampa_MSYS_N01_OGGETTI_funct(CollectionForm_PFunct, "COLLECTION_FORMS")



'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionForm_PFunct_Exit:
    Exit Function

CollectionForm_PFunct_Err:
    MsgBox Error$
    Resume CollectionForm_PFunct_Exit

End Function



'//LA COLLEZIONE DELLE FORM DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//




'//LA COLLEZIONE DELLE QUERY DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//
'//NOTE: in questa funzione precarico la collezione di oggetti QUERY _
    @PARAMETRO_02= COLLECTION DELLE QUERYS DA IMPORTARE


'//@QUERY@INSERT.02=PRECARICO LA COLLECTION QUERY DAIMPORTARE
Public Function CollectionQuery_PFunct() As Collection       '//restituisco una collection
    

On Error GoTo CollectionQuery_PFunct_Err
        


    Dim frm As Variant             '//OGGETTO QUERY DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_QUERY_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI QUERY   @importa@collection@oggetti@query_(la collezione delle @query)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI QUERY  CHE APPARTENGONO al progetto _
            MSYS_ATTIVA_GEST_OGGETTI.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le QUERY del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della QUERY di _
            importazione, in quanto le QUERY non sono impostate una tabella esterna che dovrebbe essere _
            spostata insieme alla QUERY di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe QUERY con i dati relativi in questo caso agli oggetti QUERY. _

    '//@QUERY@INSERT.03=CREA LA COLLECTION QUERY
    '//@OGGETTI@QUERY@PRECARICATI
    '//popola la collection delle QUERY di DA IMPORTARE
    ' Aggiungi i nomi delle QUERY o ELIMINA O COMMENTA LE QUERY DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti QUERY DI PROGETTI DIVERSI inserisci i nome delle QUERY _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        Set objectQUERY = New Collection
        '//Aggiugno gli elementi della collezione
        
        
        
        '//@QUERY@INSERT.04=SELECT TRA VARI OGGETTI QUERY
        '//2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
                    Le collection viene scelta in base combinata 2 che individua il progetto _
                    al posto della combinata 1 e la sostituzione riguarda tutte le collection.
                    
        Select Case Me.Cmb_02_txt
            
            '//@COLLEZIONE@QUERY@IMPEGNI
            '//@IMPEGNI@QUERY@UTILITA_(LA @collection delle query relative agli impegni)
            Case "IMPORTA_OGGETTI_IMPEGNI"
                
                '//@collection@query@impegni
                
                  objectQUERY.Add "LLPP_ATTI_Qry01_01_Gestione"
                  objectQUERY.Add "LLPP_ATTI_Qry01_12_Gestione_Elenco"
                  objectQUERY.Add "LLPP_ATTI_Qry02_01_ALLEGATI"
                  objectQUERY.Add "LLPP_IMPEGNI_Qry01_13_ELENCO_IMPEGNI"
                  objectQUERY.Add "LLPP_IMPEGNI_Qry01_13_GRUPPO_ELENCO_IMPEGNI_ANNO"
                  objectQUERY.Add "LLPP_IMPEGNI_Qry01_13_GRUPPO_ELENCO_IMPEGNI_ANNO_CLASSIFICAZIONE"
                  objectQUERY.Add "LLPP_IMPEGNI_Qry01_19_ELENCO_IMPEGNI_COD_PRAT"
                  objectQUERY.Add "LLPP_IMPEGNI_Qry01_20_01_GRUPPO_IMPEGNI_XANNO"
                  objectQUERY.Add "LLPP_IMPEGNI_Qry01_20_02_GRUPPO_IMPEGNI_XANNO_XCLASSIFICAZ"
                  objectQUERY.Add "LLPP_IMPEGNI_Qry02_01_ELENCO_DI_SPESA"
                  objectQUERY.Add "LLPP_IMPEGNI_Qry02_02_ELENCO_DI_SPESA_Totali"
                  objectQUERY.Add "LLPP_IMPEGNI_QryDF02_01_VOCE_ATTI_IMPEGNO"
                  objectQUERY.Add "LLPP_IMPEGNI_QryDF02_02_GRUPPO_CODICE_VOCE_ATTI_IMPEGNO"
                  objectQUERY.Add "LLPP_IMPEGNI_Tb01Qry_01_OGGETTI"
                  objectQUERY.Add "LLPP_IMPEGNI_Tb01Qry01_12_GRUPPO_ELENCO_IMPEGNI_ANNO"
                  objectQUERY.Add "LLPP_IMPEGNI_Tb01Qry01_13_GRUPPO_ELENCO_IMPEGNI_ANNO_CLASSIF"
                  objectQUERY.Add "LLPP_IMPEGNI_Tb01Qry01_13_GRUPPO_ELENCO_IMPEGNI_ANNO_CLASSIFICAZ"
                  objectQUERY.Add "LLPP_IMPEGNI_Tb01Qry01_14_ELENCO_IMPEGNI_COD_PRAT"
                  objectQUERY.Add "LLPP_IMPEGNI_Tb01Qry02_15_ELENCO_DI_SPESA"
                  objectQUERY.Add "LLPP_IMPEGNI_Tb01Qry02_16_ELENCO_DI_SPESA_Totali"
                  objectQUERY.Add "LLPP_IMPEGNI_Tb01Qry02_17-01_CUP_GROUPxANNO"
                  objectQUERY.Add "LLPP_IMPEGNI_Tb01Qry02_17-02_CUP"
                  objectQUERY.Add "Msys_DF02_Qry01_}----------------------------------------------@"
                  objectQUERY.Add "Msys_DF02_Qry01_01_TIPO_OGGETTI"
                  objectQUERY.Add "MSys_QUERY_Qry00_{@===========================================@}"
                  objectQUERY.Add "MSys_QUERY_Qry01_}---------------------------------------------@"
                  objectQUERY.Add "MSys_QUERY_Qry01_01_SELECT_TUTTE"
                  objectQUERY.Add "MSys_QUERY_Qry01_01_SELECT_TUTTE_TMP"
                  objectQUERY.Add "MSys_QUERY_Qry01_71_INSERT_toMsys_DF12"
                  objectQUERY.Add "MSys_QUERY_Qry01_80_}------------------------------------@DELETE"
                  objectQUERY.Add "MSys_QUERY_Qry01_81_DELETE_TUTTE"
                  objectQUERY.Add "PROGETTI_Msys_Tb00_Qry00_{@===================================@}"
                  objectQUERY.Add "PROGETTI_Msys_TB01Qry01_01_PROJECT"
                  objectQUERY.Add "PROGETTI_Msys_TB01Qry01_01_PROJECT_TMP"
                  objectQUERY.Add "PROGETTI_Msys_TB01Qry01_01_PROJECT_TMP_Delete"
                  objectQUERY.Add "PROGETTI_Msys_TB02Qry01_01_TIPO_OGGETTI"
                  objectQUERY.Add "PROGETTI_Msys_TB03Qry01_01_ELENCO_ObjProgetti"
                  objectQUERY.Add "PROGETTI_Msys_TB03Qry01_02_ObjProgetti_FORM=COD_PROGETTO_s"
                  objectQUERY.Add "PROGETTI_Msys_TB03Qry01_02_ObjProgetti_MACRO=COD_PROGETTO_s"
                  objectQUERY.Add "PROGETTI_Msys_TB03Qry01_02_ObjProgetti_QUERY=COD_PROGETTO_s"
                  objectQUERY.Add "PROGETTI_Msys_TB03Qry01_02_ObjProgetti_REPORT=COD_PROGETTO_s"
                  objectQUERY.Add "PROGETTI_Msys_TB03Qry01_02_ObjProgetti_TABLE=COD_PROGETTO_s"
 

            '//@COLLEZIONE@QUERY@ATTIVA@GEST@OGGETTI
            Case "MSYS_ATTIVA_GEST_OGGETTI.mdb"
                '//@popolo@collection@QUERY@system_(delle QUERY di sistema, inserisco nella @collezione)_@carico@collection@query
                objectQUERY.Add "MSSys_DF02_Qry01_01_SELECT_TIPO_OGGETTO"
                objectQUERY.Add "MsysDbEstTb01Qry01_Dlt00_}-------------------------------------@"
                objectQUERY.Add "MsysDbEstTb01Qry01_Dlt01_OBJECT"
                objectQUERY.Add "MsysDbEstTb01Qry01_Dlt01_OBJECT_TMP"
                objectQUERY.Add "MsysDbEstTb01Qry01_Ins00_}-------------------------------------@"
                objectQUERY.Add "MsysDbEstTb01Qry01_Ins01_OBJECT_Tmp_To_Archivio"
                objectQUERY.Add "MsysDbEstTb01Qry01_Slt00_}-------------------------------------@"
                objectQUERY.Add "MsysDbEstTb01Qry01_Slt01_OBJECT"
                objectQUERY.Add "MsysDbEstTb01Qry01_Slt01_OBJECT_TMP"
                objectQUERY.Add "MsysDbEstTb05Qry01_Slt00_}-------------------------------------@"
                objectQUERY.Add "MsysDbEstTb05Qry01_Slt01_DB_ESTERNO"
                
              
        '//@COLLEZIONE@QUERY@OGGETTI@TREEVIEW
        '//@QUERY@TREEVIEW
        Case "IMPORTA_OGGETTI_TREEVIEW"
            
                objectQUERY.Add "Tree_Qry00_{@=============================================@}_Qry"
                objectQUERY.Add "Tree_Qry01_01_}------------------------------------------@SELECT"
                objectQUERY.Add "Tree_Qry01_CLASSI_STAMPA_DI_CONTROLLO_CANCELLA"
                objectQUERY.Add "Tree_Tb01_Qry01_N01_ELENCO"
                objectQUERY.Add "Tree_Tb01_Qry01_N01_ELENCO_TMP"
                objectQUERY.Add "Tree_Tb01_Qry01_n01_PADRI+figli"
                objectQUERY.Add "Tree_Tb01_Qry01_n01_PADRI+figli_Accoda_Tmp"
                objectQUERY.Add "TreeView_Tb01Qry01_01_RADICE"
                objectQUERY.Add "TreeView_Tb01Qry02_01_TRONCO"
                objectQUERY.Add "TreeView_Tb01Qry03_01_RAMI"
                objectQUERY.Add "TreeView_Tb01Qry04_01_FOGLIE"
        
            '//@COLLEZIONE@QUERY@GEST@OGGETTI
            '//@QUERY@MSys
            Case "MSYS_ATTIVA_GEST_OGGETTI.mdb"
                
                objectQUERY.Add "MSSys_DF00_Qry00_}---------------------------------------------@"
                objectQUERY.Add "MSSys_DF02_Qry01_01_SELECT_TIPO_OGGETTO"
                objectQUERY.Add "MSsys_DF05_Qry01_ATTRIBUTI"
                objectQUERY.Add "MSys_{@-----------------------OGGETTI_DB----------------------@}"
                objectQUERY.Add "MSys_DF01_COMANDI_Qry_}----------------------------------------@"
                objectQUERY.Add "MSys_DF01_COMANDI_Qry01_01_SELECT_TUTTI"
                objectQUERY.Add "MSys_DF01_COMANDI_Qry01_02_GROUP_TUTTI"
                objectQUERY.Add "MSys_DF10_DATABASE_Qry01_01_SELECT_TUTTI"
                objectQUERY.Add "MSys_DF10_DATABASE_Qry01_01_SELECT_TUTTI_TMP"
                objectQUERY.Add "MSys_DF10_DATABASE_Qry01_81_DELETE_TUTTI_TMP"
                objectQUERY.Add "Msys_DF11_PROGETTI_Qry00_}-------------------------------------@"
                objectQUERY.Add "Msys_DF11_PROGETTI_Qry01_10_GROUP_TUTTI"
                objectQUERY.Add "Msys_DF11_Qry01_01_PROGETTI"
                objectQUERY.Add "Msys_DF12_Qry01_}----------------------------------------------@"
                objectQUERY.Add "Msys_DF12_Qry01_01_TIPO_OGGETTI"
                objectQUERY.Add "MSys_ELENCO_ObjProgetti_Qry00-}--------------------------------@"
                objectQUERY.Add "MSys_ELENCO_ObjProgetti_Qry01_01_SELECT_TUTTI"
                objectQUERY.Add "MSys_ELENCO_ObjProgetti_Qry01_81_DELETE_TUTTI"
                objectQUERY.Add "MSys_FORMS_Qry_}-----------------------------------------------@"
                objectQUERY.Add "MSys_FORMS_Qry01_01_SELECT_TUTTE"
                objectQUERY.Add "MSys_FORMS_Qry01_70_}------------------------------------@INSERT"
                objectQUERY.Add "MSys_FORMS_Qry01_71_INSERT_toMsys_DF12"
                objectQUERY.Add "MSys_FORMS_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_MACRO_Qry01_00_}------------------------------------@SELECT"
                objectQUERY.Add "MSys_MACRO_Qry01_01_SELECT_TUTTI"
                objectQUERY.Add "MSys_MACRO_Qry01_70_}-----------------------------------@INSERT"
                objectQUERY.Add "MSys_MACRO_Qry01_71_INSERT_toMsys_DF12"
                objectQUERY.Add "MSys_MACRO_Qry01_80_}------------------------------------@DELETE"
                objectQUERY.Add "MSys_MACRO_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_MODULI_Qry01_00_}-----------------------------------@SELECT"
                objectQUERY.Add "MSys_MODULI_Qry01_01_SELECT_TUTTI"
                objectQUERY.Add "MSys_MODULI_Qry01_70_}-----------------------------------@INSERT"
                objectQUERY.Add "MSys_MODULI_Qry01_71_INSERT_toMsys_DF12"
                objectQUERY.Add "MSys_MODULI_Qry01_80_}-----------------------------------@DELETE"
                objectQUERY.Add "MSys_MODULI_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_QUERY_Qry_}----------------------------------------------@"
                objectQUERY.Add "MSys_QUERY_Qry01_01_SELECT_TUTTE"
                objectQUERY.Add "MSys_QUERY_Qry01_70_}------------------------------------@INSERT"
                objectQUERY.Add "MSys_QUERY_Qry01_71_INSERT_toMsys_DF12"
                objectQUERY.Add "MSys_QUERY_Qry01_80_}------------------------------------@DELETE"
                objectQUERY.Add "MSys_QUERY_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_REPORT_Qry_}----------------------------------------------@"
                objectQUERY.Add "MSys_REPORT_Qry01_01_SELECT_TUTTE"
                objectQUERY.Add "MSys_REPORT_Qry01_70_}-----------------------------------@INSERT"
                objectQUERY.Add "MSys_REPORT_Qry01_71_INSERT_toMsys_DF12"
                objectQUERY.Add "MSys_REPORT_Qry01_80_}-----------------------------------@DELETE"
                objectQUERY.Add "MSys_REPORT_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_TABELLE_LINK_Qry_}----------------------------------------@"
                objectQUERY.Add "MSys_TABELLE_LINK_Qry01_01_SELECT_TUTTE"
                objectQUERY.Add "MSys_TABELLE_Qry_}---------------------------------------------@"
                objectQUERY.Add "MSys_TABELLE_Qry01_01_SELECT_Name1_s"
                objectQUERY.Add "MSys_TABELLE_Qry01_01_SELECT_TUTTE"
                objectQUERY.Add "MSys_TABELLE_Qry01_70_}----------------------------------@INSERT"
                objectQUERY.Add "MSys_TABELLE_Qry01_71_INSERT_toMsys_DF12"
                objectQUERY.Add "MSys_TABELLE_Qry01_80_}----------------------------------@DELETE"
                objectQUERY.Add "MSys_TABELLE_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_TABELLE_TMP_Qry00_}---------------------------------------@"
                objectQUERY.Add "MSys_TABELLE_TMP_Qry01_01_SELECT_TUTTE"
                objectQUERY.Add "MSys_TABELLE_TMP_Qry01_80_}------------------------------@DELETE"
                objectQUERY.Add "MSys_TABELLE_TMP_Qry01_81_DELETE_TUTTE"
                objectQUERY.Add "MSys_TUTTE_LE_PROPRIETA_Qry01_00_}-----------------------@SELECT"
                objectQUERY.Add "MSys_TUTTE_LE_PROPRIETA_Qry01_01_SELECT"
                objectQUERY.Add "MSys_TUTTI_PROG_}--------------------------------@Qry"
                objectQUERY.Add "MSys_TUTTI_PROG_INSERT_Qry01_01_PROGETTI_in_TMP"
                objectQUERY.Add "MSys_TUTTI_PROG_SELECT_Qry01_01_PROGETTI"
                objectQUERY.Add "MSys_TUTTI_PROG_VIS_PAGINE_SELECT_Qry01_01_Carica"
                objectQUERY.Add "MSys_TUTTI_Qry99_{@=======================================@}_Qry"
    

        
        
        '//@COLLEZIONE@QUERY@ESPROPRI
        '//2024.08.06_COLLECTION@QUERY@ESPROPRI_(inserito la collezione QUERY per gli espropri)
        '//COLLECTION@IMPORTA@QUERY@ESPROPRI   = @IMPORTA@QUERY@ESPROPRI_(definisco gli oggetti collection QUERY per GLI @ESPROPRI)
        Case "IMPORTA_OGGETTI_ESPROPRI"
        '//@COLLECTION@ESPROPRI@QUERY
            
            objectQUERY.Add "LLPP_TB09Qry00_{@===============ESPROPRI======================@}"
            objectQUERY.Add "LLPP_TB09Qry01_00_}--------------------------------------------@"
            objectQUERY.Add "LLPP_TB09Qry01_01_ESPROPRI"
            objectQUERY.Add "LLPP_TB09Qry01_01_ESPROPRI_UPDATE_TuttiCampi"
            objectQUERY.Add "LLPP_TB09Qry01_61_UPDATE_}-----------------------------------@"
            objectQUERY.Add "LLPP_TB09Qry01_62_UPDATE_SINGOLO_RECORD"
            objectQUERY.Add "LLPP_TB09Qry01_62_UPDATE_TUTTI_RECORD"
            
        '//@COLLEZIONE@QUERY@DLL
         '//@IMPORTA@QUERY@DLL@COLECTION   = @IMPORTA@QUERY@DLL_(definisco gli oggetti collection QUERY per le @liberie@DLL)
        Case "IMPORTA_OGGETTI_MSYS_DLL"
                '//@COLLECTION@DLL@QUERY
                objectQUERY.Add "Msys_DF13Qry01_00_}-----------------------------------------@DLL"
                objectQUERY.Add "Msys_DF13Qry01_01_DLL_PROGETTO"
                objectQUERY.Add "Msys_DF14Qry01_01_DELETE_LIBRERIE"
                objectQUERY.Add "Msys_DF14Qry01_02_DLL_LIBRERIE"
                objectQUERY.Add "Msys_DF14Qry01_03_UPDATE_CAMPI"
                objectQUERY.Add "Msys_DLLQry01_00_}---------------------------------------------@"
                objectQUERY.Add ""
                objectQUERY.Add ""
                objectQUERY.Add ""


        Case Else
            MsgBox "ATTENZIONE! NON CI SONO QUERY PER QUESTO PROGETTO: " & Me.Cmb_01_txt, vbCritical, "AVVISO)"
                '//2024.08.10 =COLLECTION=0- prima di uscire imposto la collezione per evitare di restiuire null e poi esco dalla routine.
                Set CollectionQuery_PFunct = objectQUERY
                '//SE NON CI SONO OGGETTI USCITA DALLA ROUTINE
                GoTo CollectionQuery_PFunct_Exit

        End Select
   
    
    
    '// *** FINE ****
    '//I° COLLEZIONE DI QUERY DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
    
    
    
    
        '//CONTROLLO DELLE QUERY PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
        '//NOTE: queste sono le QUERY da carica nel nuovo progetto.
                                                    
                        '//@QUERY@INSERT.05=RESET CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.TXT_OBJECT_ESTERNI.Requery
        
                                    
                  '//IMPOSTO LA PATH + FILE.MDB + TESTATA + OGGETTI VISUALIZZAZIONE DELLA CASELLA TXT
                  '//--------------------------------------------------------------------------------------------//
                  '//Note: @imposto@path_(imposto la path degli oggetti da recupeare, _
                           preparo la testata per la casela di testo, preparo gli oggetti e visualizzo tutto _
                           nella casella di testo.
                    
                    
                    '//@QUERY@INSERT.06=STAMPO TESTATA DI RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI
                        Str1 = ""
                        Str1 = Str1 & "CONTROLLO PATH E DATABASE DA CUI RECUPERARE GLI OGGETTI" & vbCrLf
                        Str1 = Str1 & "-----------------------------------------------------------------" & vbCrLf
                        Str1 = Str1 & Me.sourceDBPath_s_Txt & vbCrLf
                        
                        
                        '//@QUERY@INSERT.07=INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO
                        Me.TXT_OBJECT_ESTERNI = Str1
                        Me.TXT_OBJECT_ESTERNI.Requery

                            '//PREPARO LA TESTATA DELLE COLLECTION PER IL II° INSERIMENTO
                            Str1 = Str1 & vbCrLf & "                          COLLECTION QUERY DA IMPORTARE : " & vbCrLf
                            Str1 = Str1 & "//=========================================================================//" & vbCrLf
                            ' Stampa
                            Debug.Print
                            Debug.Print
                            Debug.Print Str1
                                
                            '//@QUERY@INSERT.08=ITERO NEGLI OGGETTI QUERY CARICATI E FACCIO UN STAMPA DI CONTROLLO E CARICO LA STRINGA DA STAMPARE
                            '//iterazione nella collezione per il controllo
                            For Each objqry In objectQUERY

                                Debug.Print "la MODULI da importare nel progetto: " & objqry
                                Str1 = Str1 & objqry & vbCrLf
                                
                            Next objqry
                            
                            '//@QUERY@INSERT.09=STAMPO GLI OGGETTI NELLA CASELLA DI TESTO E LA AGGIORNO
                            '//II° INSERIMENTO = ACCODO TUTTI GLI OGGETTI DA CARICARE NELLA CASELLA DI TESTO
                            Me.TXT_OBJECT_ESTERNI = vbCrLf & vbCrLf & Str1
                            Me.Recalc
                  
                  '//--------------------------------------------------------------------------------------------//
                        
                        

        '// *** FINE ***
        '//CONTROLLO DELLE QUERY PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
                                            
                        '//@QUERY@INSERT.10=RESTITUISCO LA COLLEZIONE AL CHIAMANTE
                        ' Restituisci la collezione
                        Set CollectionQuery_PFunct = objectQUERY
               



        
        '//...............................................................................//
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionQuery_PFunct_Exit:
    Exit Function

CollectionQuery_PFunct_Err:
    MsgBox Error$
    Resume CollectionQuery_PFunct_Exit

End Function



'//LA COLLEZIONE DELLE TABLE DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//


'//PRECARICO LA COLLECTION TABELLE DA IMPORTARE
Public Function CollectionTable_PFunct() As Collection          '//restituisco una collection
    

On Error GoTo CollectionTable_PFunct_Err
        
    Dim tbl As Variant             '//OGGETTO TABELLE DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_TABELLE_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI TABELLE    @importa@collection@oggetti@table_(la collezione dei @table)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI TABELLE  CHE APPARTENGONO al progetto _
            MSYS_ATTIVA_GEST_OGGETTI.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le TABELLE del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della TABELLE di _
            importazione, in quanto le TABELLE non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alla TABELLE di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe TABELLE con i dati relativi in questo caso agli oggetti TABELLE. _


    '//@OGGETTI@TABELLE@PRECARICATI
    '//popola la collection delle TABELLE di DA IMPORTARE
    ' Aggiungi i nomi delle TABELLE o ELIMINA O COMMENTA LE TABELLE DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti TABELLE DI PROGETTI DIVERSI inserisci i nome delle TABELLE _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        
        Set Tables_Collection = New Collection
        '//Aggiugno gli elementi della collezione
        
        '//SELECT DELLE COLLECTION IN BASE AL VALORE DELLA  COMBINATA_02
        '//------------------------------------------------------------------------------------//
        
            '//2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
                    Le collection viene scelta in base combinata 2 che individua il progetto _
                    al posto della combinata 1 e la sostituzione riguarda tutte le collection.
            '//@ATTENZIONE NON INSERIRE IL DOPPIO NOME DELL'OGGETTO ALTRIMENTI L'OGGETTO VIENE CANCELLATO E NON IMPORTATO _
                    @STEP@03_IMPORTA@OGGETTI_(Imposto la collection tables)           = Oggetti Tabella da importare
        Select Case Me.Cmb_02_txt
            
            '//@COLLEZIONE@TABELLE@OGGETTI@PRES3000
            '//@PRES3000@TABELLE@UTILITA_(LA @collection DELLE @TABELLE relative agli PRES3000)
            Case "IMPORTA_OGGETTI_PRES3000"
                    '//per importare le tabelle cambiero diverse path impostanto a seconda _
                       del tipo delle tabelle la casella di testo della form.
                
                        '//TABELLE DEFINIZIONI ORARI _
                              Cambio la path per le tabelle calendario
                                Me.sourceDBPath_s_Txt = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N05_DEFINIZIONI_ORAR_GIORN_TIMB\PRES3000_N05_DEFINIZIONI_ORAR_GIORN_TIMB.mdb"
                                Me.sourceDBPath_s_Txt.Requery
                                

                                 Tables_Collection.Add "PRES3000_DF00_{@==========================================@}_Tab"
                                 Tables_Collection.Add "PRES3000_DF01_}------------------------------------------------@"
                                 Tables_Collection.Add "PRES3000_DF01_CalendarioFestivita"
                                 Tables_Collection.Add "PRES3000_DF05_}------------------------------------------------@"
                                 Tables_Collection.Add "PRES3000_DF05_ORARI"
                                 Tables_Collection.Add "PRES3000_DF10_GIORNALIERE"
                                 Tables_Collection.Add "PRES3000_DF11_}------------------------------------------------@"
                                 Tables_Collection.Add "PRES3000_DF11_TIMBRATURE"
                                 Tables_Collection.Add "PRES3000_DF12_}------------------------------------------------@"
                                 Tables_Collection.Add "PRES3000_DF12_ORARI_TIMBRATURE"
                                 Tables_Collection.Add "PRES3000_DF19_}------------------------------------------------@"
                                 Tables_Collection.Add "PRES3000_DF19_GiorniCalendario"
                                 Tables_Collection.Add "PRES3000_DF20_}------------------------------------------------@"
                                 Tables_Collection.Add "PRES3000_DF20_Help_Calendario"
                           
                
                
                            '//TABELLE CALENDARIO _
                                Cambio la path per le tabelle calendario
                                Me.sourceDBPath_s_Txt = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N01_CALENDARIO\MDB\PRES3000_N01_CALENDARIO.mdb"
                                Me.sourceDBPath_s_Txt.Requery
                                Tables_Collection.Add "PRES3000_DF01_CalendarioFestivita"
                                Tables_Collection.Add "PRES3000_DF19_GiorniCalendario"
                                Tables_Collection.Add "PRES3000_DF20_Help_Calendario"
                                Tables_Collection.Add "PRES3000_Tb01_Calendario"
                                Tables_Collection.Add "PRES3000_Tb02_CalendarioPianoFerie"

                
             '//@COLLEZIONE@TABELLE@OGGETTI@IMPEGNI
             '//@IMPEGNI@TABELLE@UTILITA_(LA @collection DELLE @TABELLE relative agli impegni)
            Case "IMPORTA_OGGETTI_IMPEGNI"
                
                
                Tables_Collection.Add "LLPP_ATTI_Tb01_{@=============================================@}"
                Tables_Collection.Add "LLPP_ATTI_Tb01_Gestione"
                Tables_Collection.Add "LLPP_ATTI_Tb02_{@====================================@}_ALLEGATI"
                Tables_Collection.Add "LLPP_ATTI_Tb02_Allegati"
                
                Tables_Collection.Add "LLPP_IMPEGNI_DF02_{@===============================@}_VOCE_ATTI"
                Tables_Collection.Add "LLPP_IMPEGNI_DF02_{@==========================================@}"
                Tables_Collection.Add "LLPP_IMPEGNI_DF02_}--------------------------------------------@"
                Tables_Collection.Add "LLPP_IMPEGNI_DF02_VOCE_ATTI_IMPEGNO"
                Tables_Collection.Add "LLPP_IMPEGNI_DF02_VOCE_ATTI_IMPEGNO_TMP"
                
                Tables_Collection.Add "LLPP_IMPEGNI_TB01_{@==========================================@}"
                Tables_Collection.Add "LLPP_IMPEGNI_TB01_}--------------------------------------------@"
                Tables_Collection.Add "LLPP_IMPEGNI_Tb01_OGGETT_}-------------------------------------@"
                Tables_Collection.Add "LLPP_IMPEGNI_Tb01_OGGETTI"
                Tables_Collection.Add "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO"
                Tables_Collection.Add "LLPP_IMPEGNI_Tb01_ATTI_DI_IMPEGNO_TMP"
                
                Tables_Collection.Add "LLPP_IMPEGNI_Tb02_{@==========================================@}"
                Tables_Collection.Add "LLPP_IMPEGNI_Tb02_}--------------------------------------------@"
                Tables_Collection.Add "LLPP_IMPEGNI_Tb02_ELENCO_DI_SPESA"
                
                Tables_Collection.Add "MSys_DF02_}----------------------------------------------------@"
                Tables_Collection.Add "MSys_DF02_TIPO_OGGETTO"
                Tables_Collection.Add "MSys_DF11_}----------------------------------------------------@"
                Tables_Collection.Add "Msys_DF11_PROGETTI"
                
                Tables_Collection.Add "MSys_T_}-------------------------------------------------------@"
                Tables_Collection.Add "MSys_TABELLE"
                
                Tables_Collection.Add "PROGETTI_MSYS_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Tab"
                Tables_Collection.Add "PROGETTI_Msys_TB01_}-------------------------------------------@"
                Tables_Collection.Add "PROGETTI_Msys_TB01_PROJECT"
                Tables_Collection.Add "PROGETTI_Msys_TB01_PROJECT_TMP"
                Tables_Collection.Add "PROGETTI_Msys_TB02_}-------------------------------------------@"
                Tables_Collection.Add "PROGETTI_Msys_TB02_TIPO_OGGETTI"
                Tables_Collection.Add "PROGETTI_Msys_TB03_}-------------------------------------------@"
                Tables_Collection.Add "PROGETTI_Msys_TB03_ELENCO_ObjProgetti"
                
                

                
                
                
   
                
            '//@COLLEZIONE@TABELLE@MYSY@GEST@OGGETTI
            '//@TABELLE@OGGETTI@MSYS_(LA @COLLECTION DI TABELLE DEL PROGETTO MSYS)
            Case "MSYS_ATTIVA_GEST_OGGETTI.mdb"
            
                '//@popolo@collection@TABELLE@system_(delle TABELLE di sistema, inserisco nella @collezione)_@CARICO@COLLECTION@TABLE
                Tables_Collection.Add "COLL_MSysTb06_IMPORTA_PROGETTO"
                Tables_Collection.Add "MSys_{@===================================================@}_Tab"
                Tables_Collection.Add "MSys_DF00_}----------------------------------------------------@"
                Tables_Collection.Add "MSys_DF10_DATABASE"
                Tables_Collection.Add "Msys_DF11_PROGETTI"
                Tables_Collection.Add "Msys_DF12_GE_ObjProgetti"
                Tables_Collection.Add "Msys_DF12_TIPO_OGGETTI"
                Tables_Collection.Add "MSys_ELENCO_ObjProgetti"
                Tables_Collection.Add "MSys_F_}-------------------------------------------------------@"
                Tables_Collection.Add "MsysDbEstTb00_}------------------------------------------------@"
                Tables_Collection.Add "MsysDbEstTb01_OBJECT"
                Tables_Collection.Add "MsysDbEstTb01_OBJECT_TMP"
                Tables_Collection.Add "MSysTb05_DB_EST"
                Tables_Collection.Add "MSysTb06_IMPORTA_PROGETTO"
                                    
            '//@COLLEZIONE@TABELLE@OGGETTI@TREVIEW
            '//@TABELLE@TREEVIEW
            Case "IMPORTA_OGGETTI_TREEVIEW"
                
                Tables_Collection.Add "Tre99_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Tab"
                Tables_Collection.Add "Tree"

                Tables_Collection.Add "Tree_CLASSI_STAMPA_DI_CONTROLLO"
                Tables_Collection.Add "Tree_DF01_COMANDI"
                Tables_Collection.Add "TreeTemporanea"
                Tables_Collection.Add "TreTab_}-------------------------------------------------------@"
            
            '//@COLLEZIONE@TABELLE@ATTIVA@GEST@OGGETTI
            '//@TABELLE@MSys
            Case "MSYS_ATTIVA_GEST_OGGETTI.mdb"
                
                Tables_Collection.Add "MSys_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Tab"
                Tables_Collection.Add "MSys_{@===================================================@}_Tab"
                Tables_Collection.Add "MSys_DF00_}----------------------------------------------------@"
                Tables_Collection.Add "MSys_DF01_COMANDI"
                Tables_Collection.Add "MSys_DF01_COMANDI_TMP"
                Tables_Collection.Add "MSys_DF02_TIPO_OGGETTO"
                Tables_Collection.Add "MSys_DF05_ATTRIBUTI_TABELLA"
                Tables_Collection.Add "MSys_DF10_DATABASE"
                Tables_Collection.Add "MSys_DF10_DATABASE_TMP"
                Tables_Collection.Add "Msys_DF11_PROGETTI"
                Tables_Collection.Add "Msys_DF12_GE_ObjProgetti"
                Tables_Collection.Add "Msys_DF12_TIPO_OGGETTI"
                Tables_Collection.Add "MSys_ELENCO_}-------------------------------------------------@"
                Tables_Collection.Add "MSys_ELENCO_ObjProgetti"
                Tables_Collection.Add "MSys_ELENCO_ObjProgetti_TMP"
                Tables_Collection.Add "MSys_F_}-------------------------------------------------------@"
                Tables_Collection.Add "MSys_FORMS"
                Tables_Collection.Add "MSys_MACRO"
                Tables_Collection.Add "MSys_MODULI"
                Tables_Collection.Add "MSys_PAGINE"
                Tables_Collection.Add "MSys_QUERY"
                Tables_Collection.Add "MSys_REPORT"
                Tables_Collection.Add "MSys_TABELLE"
                Tables_Collection.Add "MSys_TABELLE_LINK"
                Tables_Collection.Add "MSys_TABELLE_Tmp"
                Tables_Collection.Add "MSys_TUTTE_}---------------------------------------------------@"
                Tables_Collection.Add "MSys_TUTTE_BARRE_DEI_COMANDI"
                Tables_Collection.Add "MSys_TUTTE_LE_PAGINE"
                Tables_Collection.Add "MSys_TUTTE_LE_PROPRIETA"
                Tables_Collection.Add "MSys_TUTTI_PRO_{@=========================================@}_Tab"
                Tables_Collection.Add "MSys_TUTTI_PRO_}-----------------------------------------------@"
                Tables_Collection.Add "MSys_TUTTI_PROG"
                Tables_Collection.Add "MSys_TUTTI_PROG_TMP"
                Tables_Collection.Add "MSys_TUTTI_PROG_VIS_PAGINE"

            
            
        '//@COLLEZIONE@TABELLE@ESPROPRI
        '//@TABELLE@ESPROPRI = LE TABELLE del progetto ESPROPRI
        Case "GE_ESPROPRI.mdb"
     
            '//@COLLECTION@ESPROPRI@TABELLE
            Tables_Collection.Add "LLPP_ATTI_Tb01_{@=============================================@}"
            Tables_Collection.Add "LLPP_ATTI_Tb01_Gestione"
            Tables_Collection.Add "LLPP_ATTI_Tb01_Gestione_TMP"
            Tables_Collection.Add "LLPP_ATTI_Tb02_{@====================================@}_ALLEGATI"
            Tables_Collection.Add "LLPP_ATTI_Tb02_Allegati"
            Tables_Collection.Add "LLPP_DF01_{===============================================}_DF01"
            Tables_Collection.Add "LLPP_DF01_CodiceOpera"
            Tables_Collection.Add "LLPP_DF02_AttivitaOperaPubblica"
            Tables_Collection.Add "LLPP_DF12_{@========================================@}_TIPO_ATTO"
            Tables_Collection.Add "LLPP_DF12_TIPO_ATTO"
            Tables_Collection.Add "LLPP_TB09_ESPR_}-----------------------------------------------@"
            Tables_Collection.Add "LLPP_TB09_ESPROPRI"
            Tables_Collection.Add "LLPP_TB09_ESPROPRI_TMP"
            
            
            '//@COLLEZIONE@TABELLE@MSYS
             '//@IMPORTA@TABELLE@DLL@COLECTION   = @IMPORTA@TABELLE@DLL_(definisco gli oggetti collection tabella per le @liberie@DLL) _


            Case "IMPORTA_OGGETTI_MSYS_DLL"
                
                '//@COLLECTION@DLL@TABELLE
                Tables_Collection.Add "MSys_DF13_}----------------------------------------------------@"
                Tables_Collection.Add "Msys_DF13_DLL_PROGETTO"
                Tables_Collection.Add "Msys_DF14_DLL_LIBRERIE"
                Tables_Collection.Add "Msys_DLL_LIBRERIE"
            
            '//@COLLEZIONE@TABELLE@CDM@GIUR
            '//@COLLECTION@TABELLE@CDM_(le collezioni DI TABELLE DEL CONDOMINIO)
            '//-------------------------------------------------------------------------------------//
                '//@CONDOMINIO@TABELLE@GIURIDICA
                 Case "IMPORTA_OGGETTI_CDM_GIUR"
                        Tables_Collection.Add "GECOND_DF01_}--------------------------------------------------@"
                        Tables_Collection.Add "GECOND_DF01_TIPI"
                    
                '//@COLLEZIONE@TABELLE@CDM@CONSUNTIVO
                '//@CONDOMINIO@TABELLE@CONSUNTIVO
                Case "IMPORTA_OGGETTI_CDM_CONS"
                        Tables_Collection.Add "CONS_TB09_}----------------------------------------------------@"
                        Tables_Collection.Add "CONS_TB09_MASTRO"
                        
                '//@COLLEZIONE@TABELLE@CDM@ANAGRAFICA
                '//@CONDOMINIO@TABELLE@ANAGRAFICA
                Case "IMPORTA_OGGETTI_CDM_ANAGR"
                        Tables_Collection.Add "ANAGR_Tb01_}---------------------------------------------------@"
                        Tables_Collection.Add "ANAGR_TB01_PROPRIETARI"
                        Tables_Collection.Add "ANAGR_Tb02_}---------------------------------------------------@"
                        Tables_Collection.Add "ANAGR_Tb02_PROPR_CONDOM"
                        Tables_Collection.Add "ANAGR_Tb03_LOCAZIONI"
                        Tables_Collection.Add "ANAGR_TB04_PEC_EMAIL"
                
                '//@COLLEZIONE@TABELLE@CDM@GE@COND
                '//@CONDOMINIO@TABELLE@GECOND
                Case "IMPORTA_OGGETTI_CDM_GECOND"
                            Tables_Collection.Add "GECOND_DF01_}--------------------------------------------------@"
                            Tables_Collection.Add "GECOND_DF01_TIPI"
                            
                            
                '//@COLLEZIONE@TABELLE@CDM@PREVENTIVO
                '//@CONDOMINIO@TABELLE@PREVENTIVO
                Case "IMPORTA_OGGETTI_CDM_GE_PREV"
                            Tables_Collection.Add "PREV_TB09_}----------------------------------------------------@"
                            Tables_Collection.Add "PREV_TB09_DEF_SPESA"
                            Tables_Collection.Add "PREV_TB09_DEF_VOCI"
                            Tables_Collection.Add "PREV_TB09_PREVENTIVO"
                            Tables_Collection.Add "PREV_TB25_}----------------------------------------------------@"
                            Tables_Collection.Add "PREV_TB25_RIPARTO_PREV_MATTEI_2023"
                
                '//@COLLEZIONE@TABELLE@CDM@TABELLE@MILLESIMI
                '//@CONDOMINIO@TABELLE@MILLESIMALI
                Case "IMPORTA_OGGETTI_CDM_GE_TABELLE"
                        Tables_Collection.Add "TB01_MIL_}-----------------------------------------------------@"
                        Tables_Collection.Add "TB01_MILL"
                        Tables_Collection.Add "TB01_PROPR_}---------------------------------------------------@"
                        Tables_Collection.Add "TB01_PROPR_GEN"
                
                '//@COLLEZIONE@TABELLE@CDM@UTILITA
                '//@CONDOMINIO@TABELLE@UTILITA
                Case "IMPORTA_OGGETTI_CDM_GE_UTIL"
                        Tables_Collection.Add "UTIL_TB20_}----------------------------------------------------@"
                        Tables_Collection.Add "UTIL_Tb20_N01_GESTIONE_PROGETTI"
                        Tables_Collection.Add "UTIL_Tb20_N01_REPLICHE_OGGETTO"
        
        
            '//-------------------------------------------------------------------------------------//

            
        
            Case Else
            
                MsgBox "ATTENZIONE! NON CI SONO TABELLE PER QUESTO PROGETTO: " & Me.Cmb_01_txt, vbCritical, "AVVISO)"
                
                '//2024.08.10 =COLLECTION=0- prima di uscire imposto la collezione per evitare di restiuire null e poi esco dalla routine.
                Set CollectionTable_PFunct = Tables_Collection
                '//SE NON CI SONO OGGETTI USCITA DALLA ROUTINE
                GoTo CollectionTable_PFunct_Exit

            
        End Select
      
      '//SELECT DELLE COLLECTION IN BASE AL VALORE DELLA  COMBINATA_02  *** fine ***
      '//------------------------------------------------------------------------------------//
      
            
                    
    
        ' Inizializza le collezioni per TABELLE COLLEGATE
        Set connectedTABELLES = New Collection
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION TABLE per ottenere la collezione di TABELLE
                
                '//Set tablesCollection = CollectionTable_PFunct()
                
              '//------------------------------------------------------------------------//
        
      
    '// *** FINE ****
    '//I° COLLEZIONE DI TABELLE DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    

        '//CONTROLLO DELLE_TABELLE_PRECARICATE  NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
        '//NOTE: queste sono DELLE_TABELLE_PRECARICATE  carica nel nuovo progetto.
                        
                        '//RESET CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.TXT_OBJECT_ESTERNI.Requery
        
                                    
                  '//IMPOSTO LA PATH + FILE.MDB + TESTATA + OGGETTI VISUALIZZAZIONE DELLA CASELLA TXT
                  '//--------------------------------------------------------------------------------------------//
                  '//Note: @imposto@path_(imposto la path degli oggetti da recupeare, _
                           preparo la testata per la casela di testo, preparo gli oggetti e visualizzo tutto _
                           nella casella di testo.
                    
                    
                    '//RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI
                        Str1 = ""
                        Str1 = Str1 & "CONTROLLO PATH E DATABASE DA CUI RECUPERARE GLI OGGETTI" & vbCrLf
                        Str1 = Str1 & "-----------------------------------------------------------------" & vbCrLf
                        Str1 = Str1 & Me.sourceDBPath_s_Txt & vbCrLf
                        
                        '//INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO
                        Me.TXT_OBJECT_ESTERNI = Str1
                        Me.TXT_OBJECT_ESTERNI.Requery

                            '//PREPARO LA TESTATA DELLE COLLECTION PER IL II° INSERIMENTO
                            Str1 = Str1 & vbCrLf & "                COLLECTION DELLE_TABELLE_PRECARICATE DA IMPORTARE : " & vbCrLf
                            Str1 = Str1 & "//=========================================================================//" & vbCrLf
                            ' Stampa
                            Debug.Print
                            Debug.Print
                            Debug.Print Str1
                            
                            '//iterazione nella collezione per il controllo
                            For Each objtbl In Tables_Collection

                                Debug.Print "DELLE_TABELLE_PRECARICATE  importare nel progetto: " & objtbl
                                Str1 = Str1 & objtbl & vbCrLf
                                
                            Next objtbl
                            
                            
                            '//II° INSERIMENTO = ACCODO TUTTI GLI OGGETTI DA CARICARE NELLA CASELLA DI TESTO
                            Me.TXT_OBJECT_ESTERNI = vbCrLf & vbCrLf & Str1
                            Me.Recalc
                  '//--------------------------------------------------------------------------------------------//
                        
        '// *** FINE ***
        '//CONTROLLO DELLE_TABELLE_PRECARICATE  NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
                                             
                                                     
                            
                            
                     ' Restituisci la collezione
                        Set CollectionTable_PFunct = Tables_Collection
                        



        
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionTable_PFunct_Exit:
    Exit Function

CollectionTable_PFunct_Err:
    MsgBox Error$
    Resume CollectionTable_PFunct_Exit

End Function


'//LA COLLEZIONE DELLE TABLE DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//






'//LA COLLEZIONE DELLE MACRO DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//

'//@carico@collection@macro_(la collezione di oggetti macro viene @precaricata)
'//PRECARICO LA COLLECTION MACRO DA IMPORTARE
Public Function CollectionMacro_PFunct() As Collection          '//restituisco una collection
    

On Error GoTo CollectionMacro_PFunct_Err
        
    Dim frm As Variant             '//OGGETTO MACRO DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_MACRO_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI MACRO  @importa@collection@oggetti@macro_(la collezione delle @macro)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI MACRO  CHE APPARTENGONO al progetto _
            MSYS_ATTIVA_GEST_OGGETTI.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le MACRO del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della MACRO di _
            importazione, in quanto le MACRO non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alla MACRO di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe MACRO con i dati relativi in questo caso agli oggetti MACRO. _


    '//@OGGETTI@MACRO@PRECARICATI
    '//popola la collection dei MACRO di DA IMPORTARE
    ' Aggiungi i nomi dei MACRO o ELIMINA O COMMENTA I MACRO DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti MACRO DI PROGETTI DIVERSI inserisci i nome dei MACRO _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        
        Set objectMACRO = New Collection
        '//Aggiugno gli elementi della collezione
        

           
        
        '//2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
                    Le collection viene scelta in base combinata 2 che individua il progetto _
                    al posto della combinata 1 e la sostituzione riguarda tutte le collection.
        Select Case Me.Cmb_02_txt
            
            '//@COLLEZIONE@MACRO@IMPEGNI
            '//@collection@MACRO@IMPEGNI@OBJECT
             Case "IMPORTA_OGGETTI_IMPEGNI"
                  '//@popolo@collection@MACRO@IMPEGNI_(MACRO PER GLI OGGETTI @IMPEGNI, inserisco nella @collezione)
                  objectMACRO.Add "COLLEGA_Mcr_{@================Ricostruisci====================@}"
                  objectMACRO.Add "COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione"
                  objectMACRO.Add "COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione_TMP"
                  objectMACRO.Add "LLPP_IMPEGNI_Tb00Mcr00_}---------------------------------------@"
                  objectMACRO.Add "LLPP_IMPEGNI_Tb00Mcr01_IMPORTA_MODULI"
                  objectMACRO.Add "LLPP_IMPEGNI_Tb02Mcr10_ESPORTA_TUTTI_GLI_OGGETTI"
                  objectMACRO.Add "LLPP_IMPEGNI_Tb02Mcr10_}---------------------------------------@"
                  objectMACRO.Add "LLPP_IMPEGNI_Tb02Mcr09_RICOSTRUISCI_TUTTI_GLI_IMPEGNI"
                  objectMACRO.Add "LLPP_IMPEGNI_Tb02Mcr09_}---------------------------------------@"
                  objectMACRO.Add "ESPORTA_Mcr_{@================Esporta=========================@}"
                  objectMACRO.Add "ESPORTA_Mcr02_01_IMPEGNI_Tb02_TuttiGliOggetti"
                  objectMACRO.Add "ESPORTA_MACRO_DETERMINE->LLPP_ATTI_GESTIONE_mdb"
                  objectMACRO.Add "ESPORTA_Mcr_IMPEGNI_TUTTO**"

            
            '//@COLLEZIONE@MACRO@ATTIVA@GEST@IMPEGNI
            Case "MSYS_ATTIVA_GEST_OGGETTI.mdb"
                '//@popolo@collection@MACRO@system_(dei MACRO di sistema, inserisco nella @collezione)
                objectMACRO.Add "MsysDbEstTb00Mcr00_}--------------------------------------@Macro"
                objectMACRO.Add "MsysDbEstTb00Mcr01_N01_ESPORTA_XLS"
                objectMACRO.Add "MsysDbEstTb00Mcr01_N02_COLLEGA_XLS"

        '//@COLLEZIONE@MACRO@TREEVIEW
        '//@MACRO@TREEVIEW
        Case "IMPORTA_OGGETTI_TREEVIEW"
            
            objectMACRO.Add "CERCA_Mcr_{@==================================================@}}"
            objectMACRO.Add "CERCA_Mcr_FORM"
            objectMACRO.Add "CERCA_Mcr_MODULO"
            objectMACRO.Add "CERCA_Mcr_QUERY"
            objectMACRO.Add "CERCA_Mcr_TABELLA"
            objectMACRO.Add "MACRO_{@======================================================@}"
            objectMACRO.Add "MSys_Mcr00_}---------------------------------------------------@"
            objectMACRO.Add "MSys_Mcr01_ATTIVA_MACRO_ESPORTA_OGGETTI_MSSys"
            objectMACRO.Add "TREEVIEW_Mcr00_{@============================================@}"
            objectMACRO.Add "TREEVIEW_Mcr01_ATTIVA_LIBRERIE"
            objectMACRO.Add "TREEVIEW_Mcr01_STAMPA_LIBRERIE_ATTIVE"
            
        '//@COLLEZIONE@MACRO@MSYS
        '//@MACRO@MSys
        Case "MSYS_ATTIVA_GEST_OGGETTI.mdb"
            '//@COLLECTION@ESPROPRI@MACRO
            objectMACRO.Add "CERCA_Mcr_{@==================================================@}"
            objectMACRO.Add "CERCA_Mcr_FORM"
            objectMACRO.Add "CERCA_Mcr_MODULO"
            objectMACRO.Add "CERCA_Mcr_QUERY"
            objectMACRO.Add "CERCA_Mcr_TABELLA"
            objectMACRO.Add "MSys_Mcr00_{@=================================================@}"
            objectMACRO.Add "MSys_Mcr00_}---------------------------------------------------@"
            objectMACRO.Add "MSys_Mcr01_ATTIVA_MACRO_ESPORTA_OGGETTI_MSSys"
            
        
          
        
        
        '//@COLLEZIONE@MACRO@ESPROPRI
        '//@MACRO@ESPROPRI = LE MACRO del progetto ESPROPRI
        '//2024.08.06_COLLECTION@FORM@ESPROPRI_(inserito la collezione MACRO per gli espropri)
        '//COLLECTION@IMPORTA@MACRO@ESPROPRI   = @IMPORTA@MACRO@DLL_(definisco gli oggetti collection MACRO per GLI @ESPROPRI)
        Case "IMPORTA_OGGETTI_ESPROPRI"
            '//@COLLECTION@ESPROPRI@MACRO
            
                                   
                    objectMACRO.Add "COLLEGA_Mcr_}--------------------------------------------------@"
                    objectMACRO.Add "COLLEGA_Mcr_GEST_MENU_TB01_VOCI_Espropri"
                    objectMACRO.Add "COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione_Espropri"
                    objectMACRO.Add "COLLEGA_Mcr_LLPP_ATTI_Tb01_Gestione_TMP_Espropri"
                    objectMACRO.Add "COLLEGA_Mcr_LLPP_DF01_CodiceOpera_Espropri"
                    objectMACRO.Add "COLLEGA_Mcr_LLPP_DF02_AttivitaOperaPubblica_Espropri"
                    objectMACRO.Add "COLLEGA_Mcr_LLPP_DF12_TIPO_ATTO_Espropri"
                    objectMACRO.Add "COLLEGA_Mcr_LLPP_TUTTI_GLI_ESPROPRI"
                    objectMACRO.Add "COLLEGA_TABELLA_Mcr_ATTI_Tb01_Gestione_TUTTO_Espropri"
                    objectMACRO.Add "ESPR_Mcr00_}---------------------------------------------------@"
                    objectMACRO.Add "ESPR_Mcr01_CollegaTUTTO_(Atti+Allegati+CodOpera+TipoAtto+OOPP)"
                    objectMACRO.Add "LLPP_TB09Mcr01_}-----------------------------------------------@"
                    objectMACRO.Add "LLPP_TB09Mcr01_01_ESPROPRI_COLLEGA"
        
         '//@COLLEZIONE@MACRO@DLL
         '//@IMPORTA@MACRO@DLL@COLECTION   = @IMPORTA@MACRO@DLL_(definisco gli oggetti collection MACRO per le @liberie@DLL) _
            @STEP@04_IMPORTA@OGGETTI_(Imposto la collection Macros)           = Oggetti Macro da importare
        Case "IMPORTA_OGGETTI_MSYS_DLL"
                '//@COLLECTION@DLL@MACRO _

                objectMACRO.Add "UTILITA_MsysDF13_Mcr00_}------------------------------------@DLL"
                objectMACRO.Add "UTILITA_MsysDF13_Mcr01_STAMPA_LIBRERIE_ATTIVE_DLL"
                objectMACRO.Add "UTILITA_MsysDF13_Mcr02_MPORTA_OGGETTI_ATTIVA_LIBRERIE_DLL"
                objectMACRO.Add "UTILITA_MsysDF13_Mcr03_IMPORTA_LINK_TABELLE_Msys_DLL_PROGETTO"
                
        Case ""

        Case Else
            MsgBox "ATTENZIONE! NON CI SONO COLLECTION MACRO DA CARICARE IN QUESTO PER QUESTO PROGETTO: " & Me.Cmb_01_txt, vbCritical, "AVVISO)"
                '//2024.08.10 =COLLECTION=0- prima di uscire imposto la collezione per evitare di restiuire null e poi esco dalla routine.
                Set CollectionMacro_PFunct = objectMACRO
                '//SE NON CI SONO OGGETTI USCITA DALLA ROUTINE
                GoTo CollectionMacro_PFunct_Exit

        End Select
   


    
        ' Inizializza le collezioni per MACRO COLLEGATE
        Set connectedMACRO = New Collection
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION MACRO per ottenere la collezione di MACRO
                
                '//Set macroCollection = CollectionMacro_PFunct()
                
              '//------------------------------------------------------------------------//
        
      
    '// *** FINE ****
    '//I° COLLEZIONE DI MACRO DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
        
        '//CONTROLLO DELLE MACRO PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
        '//NOTE: queste sono le MACRO da carica nel nuovo progetto.
                        
                        '//RESET CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.TXT_OBJECT_ESTERNI.Requery
        
                                    
                  '//IMPOSTO LA PATH + FILE.MDB + TESTATA + OGGETTI VISUALIZZAZIONE DELLA CASELLA TXT
                  '//--------------------------------------------------------------------------------------------//
                  '//Note: @imposto@path_(imposto la path degli oggetti da recupeare, _
                           preparo la testata per la casela di testo, preparo gli oggetti e visualizzo tutto _
                           nella casella ditesto.
                    
                    
                    '//RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI
                        Str1 = ""
                        Str1 = Str1 & "CONTROLLO PATH E DATABASE DA CUI RECUPERARE GLI OGGETTI" & vbCrLf
                        Str1 = Str1 & "-----------------------------------------------------------------" & vbCrLf
                        Str1 = Str1 & Me.sourceDBPath_s_Txt & vbCrLf
                                                                        
                        '//INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO
                        Me.TXT_OBJECT_ESTERNI = Str1
                        Me.TXT_OBJECT_ESTERNI.Requery

                            '//PREPARO LA TESTATA DELLE COLLECTION PER IL II° INSERIMENTO
                            Str1 = Str1 & vbCrLf & "                          COLLECTION MACRO DA IMPORTARE : " & vbCrLf
                            Str1 = Str1 & "//=========================================================================//" & vbCrLf
                                                       
                            ' Stampa le MACRO fisiche
                            Debug.Print
                            Debug.Print
                            Debug.Print "MACRO DA IMPORTARE:"
                            '//iterazione nella collezione per il controllo
                            For Each mcr In objectMACRO
                                Debug.Print "la MACRO da importare nel progetto: " & mcr & vbCrLf
                                Str1 = Str1 & mcr & vbCrLf
                            Next mcr
         
                  '//--------------------------------------------------------------------------------------------//
                        
        '// *** FINE ***
        '//CONTROLLO DELLE MACRO PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
                            
                    
                    '//STAMPA NELLA CASELLA DI TESTO LA COLLEZIONE
                    '//...............................................................................//
                    '//NOTE: per la procedura di stampa utilizzando la funzione vengono eseguite _
                            due attivita: 01) carico gli oggetti; 02) attivo la funzione di stampa.
                            
                        '//01 CARICO GLI OGGETTI
                        'Restituisco la collezione macro
                        Set CollectionMacro_PFunct = objectMACRO
                        
                        '//02 ATTIVO LA PROCEDUR DI STAMPA UTILIZZANTO TRA I PARAMETRI LA COLLECTION DI OGGETTI
                        '//@PROCEDURA.DI@STAMPA_(@Stampa oggetti utilizzando la funzione)
                        '//CHIAMO LA PROCEDURA UNICA DI STAMPA E CARICAMENTO NELLA CASELLA DI TESTO _
                                parametro_01= la collezione, parametro_02 = il tipo
                      Vv1 = Stampa_MSYS_N01_OGGETTI_funct(CollectionMacro_PFunct, "COLLECTION_MACROS")
                    
                    '//...............................................................................//
                    
                    
                     
         

                  


        
        
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionMacro_PFunct_Exit:
    Exit Function

CollectionMacro_PFunct_Err:
    MsgBox Error$
    Resume CollectionMacro_PFunct_Exit

End Function


'//LA COLLEZIONE DELLE MACRO DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//










'//LA COLLEZIONE DELLE MODULI DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//


'//PRECARICO LA COLLECTION MODULI DA IMPORTARE
Public Function CollectionModuli_PFunct() As Collection          '//restituisco una collection
    

On Error GoTo CollectionModuli_PFunct_Err
        
     Dim moduleName As Variant                       '//OGGETTO MODULI DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_MODULI_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI MODULI   @importa@collection@oggetti@moduli_(la collezione dei @moduli)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI MODULI  CHE APPARTENGONO al progetto _
            MSYS_ATTIVA_GEST_OGGETTI.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le MODULI del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo della MODULI di _
            importazione, in quanto le MODULI non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alla MODULI di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe MODULI con i dati relativi in questo caso agli oggetti MODULI. _


    '//@OGGETTI@MODULI@PRECARICATI
    '//popola la collection dei MODULI di DA IMPORTARE
    ' Aggiungi i nomi dei MODULI o ELIMINA O COMMENTA I MODULI DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti MODULI DI PROGETTI DIVERSI inserisci i nome dei MODULI _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        
        Set objectModules = New Collection
        '//Aggiugno gli elementi della collezione
        
        
        
        
        
           
        
        
        '//2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
                    Le collection viene scelta in base combinata 2 che individua il progetto _
                    al posto della combinata 1 e la sostituzione riguarda tutte le collection.
        Select Case Me.Cmb_02_txt
                
           '//@COLLEZIONE@MODULI@IMPEGNI
           '//@IMPEGNI@MODULI@UTILITA_(LA @collection DEI @MODULI relativi agli impegni)
            Case "IMPORTA_OGGETTI_IMPEGNI"
            
                '//@COLLECTION@ESPROPRI@MODULI
                objectModules.Add "LLPP_IMPEGNI_TB01Mdl02_ESPORTA_TUTTI_GLI_OGGETTI"
                objectModules.Add "LLPP_IMPEGNI_TB01Mdl01_OGGETTI_ATTI_DI_IMPEGNO_Semplice"
                objectModules.Add "LLPP_IMPEGNI_TB01Mdl00_{@=============IMPEGNI=================@}"
                objectModules.Add "GEST_FrmM00_GESTIONE_Mdl01_ControlloFormAperta"
                objectModules.Add "BARRA_ATTI_Tb01_12_Gestione_Elenco"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_}--------------------------------------@"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_DATABASE"
                objectModules.Add "GEST_FrmM00_GESTIONE_Mdl05_DF05_FALDONI"
                objectModules.Add "FILE_Mdl00_{@=================================================@}"
                objectModules.Add "FILE_Mdl10_APRI_FILE_E_PATH"
                objectModules.Add "GEST_FrmM00_GESTIONE_Mdl05_}-----------------------------@_FALD"
                objectModules.Add "LLPP_DF06_{@==============================================@}_Mdl"
                objectModules.Add "LLPP_DF06_APRI_N06_FASCICOLI"
                objectModules.Add "BARRA_ATTI_Tb01_12_}-----------------------------------------@"
                objectModules.Add "LLPP_IMPEGNI_Mdl01_02------------------------------------------@"
                objectModules.Add "LLPP_IMPEGNI_Mdl01_02_LLPP_ATTI"
                objectModules.Add "CLASSE_ATTI_Tb01_Gestione"
                objectModules.Add "UTIL_Nrx1001_{@=============CLASSI_OGGETTI====================@}"
                objectModules.Add "UTIL_Nrx1001_n01_CLASSE_OGGETTI_DELLA_FORM"
                objectModules.Add "UTIL_Nrx1003_n01_RECUPERO_PARAMETRI_TABELLA_OGGETTI"
                

                  
                  '// la macro che attiva il modulo per l'esportazione della form UTILITA
                  objectModules.Add "ESPORTA_Mdl_}--------------------------------------------------@"
                  objectModules.Add "ESPORTA_Mdl01_FORM_UTILITA_Msys_Frm21_IMPORTA_OGGETTI"
              

         
            '//@COLLEZIONE@MODULI@GEST@OGGETTI
            '//@popolo@collection@MODULI@system_(dei MODULI di sistema, inserisco nella @collezione; )_@carico@collection@macro
            '//@MODULI@MSys
            Case "MSYS_ATTIVA_GEST_OGGETTI.mdb"
                '//@COLLECTION@ESPROPRI@MODULI
                objectModules.Add "UTIL_MDL40_{@=================================================@}}"
                objectModules.Add "UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB_SALVATAGGIO_TUTTI"
                objectModules.Add "UTIL_MDL40_N08_RICERCA_TUTTI_GLI_OGGETTI_NEL_DB"
                objectModules.Add "UTIL_MDL40_N09_CONTROLLO_ESISTENZA_OGGETTO_DB"
                objectModules.Add "UTIL_MDL40_N10_CONTROLLO_OGGETTI_APRI_ACCESS"
                objectModules.Add "UTIL_MDL40_N11_MACRO_ESPORTA_OGGETTI_MSSys"
                objectModules.Add "UTIL_MDL90_01_(VARIABILI_GLOBAL_FORM)"
                objectModules.Add "UTIL_Nrx1001_{@=============CLASSI_OGGETTI====================@}"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_}--------------------------------------@"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_DATABASE"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_FORM"
                objectModules.Add "UTIL_Nrx1003_n01_RECUPERO_PARAMETRI_TABELLA_OGGETTI"
                objectModules.Add "COLLEGA_Mdl_PRES3000_TB25_DIPENDENTI"
                objectModules.Add "CLASSE_{@=====================================================@}"
                objectModules.Add "CLASSE_ATTI_Tb01_Gestione"
       
            
        '//@MODULI@TREEVIEW
        Case "IMPORTA_OGGETTI_TREEVIEW"
            
                objectModules.Add "ACCESS_{@=================================================@}_Mdl"
                objectModules.Add "ACCESS_Classe_LIBRERIE_RIFERIMENTI"
                objectModules.Add "ACCESS_Mdl_REFERENZIA_LE_LIBRERIE"
                objectModules.Add "API_WINDOWS_Main"
                objectModules.Add "API_WINDOWS_Test_Excel"
                objectModules.Add "APRI_WINDOWS_File"
                objectModules.Add "BARRA_COMANDI_{@==================MDL=========================@}"
                objectModules.Add "BARRA_COMANDI_MyContest"
                objectModules.Add "BARRA_COMANDI_MyTEST"
                objectModules.Add "BARRA_DEI_COMANDI_ClasseEventi"
                objectModules.Add "BARRA_DEI_COMANDI_CopiaDueBarre"
                objectModules.Add "BARRA_DEI_MENU+BARRA_DEI_COMANDI_MSys_OGGETTI"
                objectModules.Add "CLASSE_{@=====================================================@}"
                objectModules.Add "CLASSE_ALBERO"
                objectModules.Add "CLASSE_ALBERO_Test"
                objectModules.Add "CLASSE_ATTI_Tb01_Gestione"
                objectModules.Add "CLASSE_ClsStack_TEST"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_}--------------------------------------@"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_DATABASE"
                objectModules.Add "CLASSE_GESTIONE_OGGETTI_FORM"
                objectModules.Add "CLASSE_NOD_}---------------------------------------------------@"
                objectModules.Add "CLASSE_NODO"
                objectModules.Add "CLASSE_NODO_Test"
                objectModules.Add "CLASSE_STACK"
                objectModules.Add "CLASSE_TREE"
                objectModules.Add "clsFile"
                objectModules.Add "ComboBoxHandler"
                objectModules.Add "CreaAlberoDaTabel_}--------------------------------------------@"
                objectModules.Add "CreaAlberoDaTabella"
                objectModules.Add "CreaAlberoDaTabella_TEST"
                objectModules.Add "ESPORTA_Mdl_}--------------------------------------------------@"
                objectModules.Add "GESTIONE_Mdl10_FILE_APRI_PDF"
                objectModules.Add "LLPP_ATTI_Tb01_Gestione_Classe"
                objectModules.Add "LLPP_CLASSE_MENU_LAVORI_PUBBLICI"
                objectModules.Add "PROGETTI_GESTIONE_TUTTI_Classe"
                objectModules.Add "UTIL_{=====================================================}_Mdl"
                objectModules.Add "UTIL_MDL30_CANCELLA_I_MODULI_DbEsterno"
                objectModules.Add "UTIL_MDL40_{@=================================================@}"
                objectModules.Add "UTIL_MDL40_N01_CONTROLLO_OGGETTI_DB_SALVATAGGIO_TUTTI"
                objectModules.Add "UTIL_MDL40_N08_RICERCA_TUTTI_GLI_OGGETTI_NEL_DB"
                objectModules.Add "UTIL_MDL40_N09_CONTROLLO_ESISTENZA_OGGETTO_DB"
                objectModules.Add "UTIL_MDL40_N10_CONTROLLO_OGGETTI_APRI_ACCESS"
                objectModules.Add "UTIL_MDL40_N11_MACRO_ESPORTA_OGGETTI_MSSys"
                objectModules.Add "UTIL_MDL90_01_(VARIABILI_GLOBAL_FORM)"
                objectModules.Add "UTIL_Nrx1001_{@=============CLASSI_OGGETTI====================@}"
                objectModules.Add "UTIL_Nrx1003_n01_RECUPERO_PARAMETRI_TABELLA_OGGETTI"
                 
                
                
            
            '//@COLLEZIONE@MODULI@OGGETTI@ESPROPRI
            '//2024.08.06_COLLECTION@MODULI@ESPROPRI_(inserito la collezione MODULI per gli espropri)
            '//COLLECTION@IMPORTA@MODULI@ESPROPRI   = @IMPORTA@MODULI@DLL_(definisco gli oggetti collection MODULI per GLI @ESPROPRI)
            Case "IMPORTA_OGGETTI_ESPROPRI"
            '//@COLLECTION@ESPROPRI@MODULI
            
            objectModules.Add "ESPORTA_FILE_Mdl00=============================================="
            objectModules.Add "FILE_Mdl01_TROVA_PATH_E_FILE"
            


        '//@COLLEZIONE@MODULI@DLL
        '//@IMPORTA@MODULI@DLL@COLECTION   = @IMPORTA@MODULI@DLL_(definisco gli oggetti collection MODULI per le @liberie@DLL) _
            @STEP@05_IMPORTA@OGGETTI_(Imposto la collection Modules)          = Oggetti Moduli da importare
            
        Case "IMPORTA_OGGETTI_MSYS_DLL"
                '//@COLLECTION@DLL@MODULI _

                objectModules.Add "UTILITA_MsysDF13Mdl01_DLL_REFERENZIA_LE_LIBRERIE"
                
        
        Case Else
            MsgBox "ATTENZIONE! NON CI SONO MODULI PER QUESTO PROGETTO: " & Me.Cmb_01_txt, vbCritical, "AVVISO)"
                '//2024.08.10 =COLLECTION=0- prima di uscire imposto la collezione per evitare di restiuire null e poi esco dalla routine.
                Set CollectionModuli_PFunct = objectModules
                '//SE NON CI SONO OGGETTI USCITA DALLA ROUTINE
                GoTo CollectionModuli_PFunct_Exit
                
        End Select
   
    
        ' Inizializza le collezioni per MODULI COLLEGATE
        Set connectedMODULIS = New Collection
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION MODULI per ottenere la collezione di MODULI
                
                '//Set tablesCollection = CollectionModuli_PFunct()
                
              '//------------------------------------------------------------------------//
        
      
    '// *** FINE ****
    '//I° COLLEZIONE DI MODULI DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
    
        '//CONTROLLO DELLE MODULI PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
        '//NOTE: queste sono le MODULI da carica nel nuovo progetto.
                        
                        '//RESET CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.TXT_OBJECT_ESTERNI.Requery
        
                                    
                  '//IMPOSTO LA PATH + FILE.MDB + TESTATA + OGGETTI VISUALIZZAZIONE DELLA CASELLA TXT
                  '//--------------------------------------------------------------------------------------------//
                  '//Note: @imposto@path_(imposto la path degli oggetti da recupeare, _
                           preparo la testata per la casela di testo, preparo gli oggetti e visualizzo tutto _
                           nella casella di testo.
                    
                    
                    '//RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI
                        Str1 = ""
                        Str1 = Str1 & "CONTROLLO PATH E DATABASE DA CUI RECUPERARE GLI OGGETTI" & vbCrLf
                        Str1 = Str1 & "-----------------------------------------------------------------" & vbCrLf
                        Str1 = Str1 & Me.sourceDBPath_s_Txt & vbCrLf
                        
                        '//INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO
                        Me.TXT_OBJECT_ESTERNI = Str1
                        Me.TXT_OBJECT_ESTERNI.Requery

                            '//PREPARO LA TESTATA DELLE COLLECTION PER IL II° INSERIMENTO
                            Str1 = Str1 & vbCrLf & "                          COLLECTION MODULI DA IMPORTARE : " & vbCrLf
                            Str1 = Str1 & "//=========================================================================//" & vbCrLf
                            ' Stampa
                            Debug.Print
                            Debug.Print
                            Debug.Print Str1
                            
                            '//iterazione nella collezione per il controllo
                            For Each objMdl In objectModules
                                Debug.Print "la MODULI da importare nel progetto: " & objMdl
                                Str1 = Str1 & objMdl & vbCrLf
                            Next objMdl
                            
                            
                  '//--------------------------------------------------------------------------------------------//
                        
        '// *** FINE ***
        '//CONTROLLO DELLE MODULI PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
                        
                        
                      
                      
                    '//STAMPA NELLA CASELLA DI TESTO LA COLLEZIONE
                    '//...............................................................................//
                    '//NOTE: per la procedura di stampa utilizzando la funzione vengono eseguite _
                            due attivita: 01) carico gli oggetti; 02) attivo la funzione di stampa.
                            
                        '//01 CARICO GLI OGGETTI
                        'Restituisco la collezione MODULI
                        Set CollectionModuli_PFunct = objectModules
                        
                        '//02 ATTIVO LA PROCEDUR DI STAMPA UTILIZZANTO TRA I PARAMETRI LA COLLECTION DI OGGETTI
                        '//@PROCEDURA.DI@STAMPA_(@Stampa oggetti utilizzando la funzione)
                        '//CHIAMO LA PROCEDURA UNICA DI STAMPA E CARICAMENTO NELLA CASELLA DI TESTO _
                                parametro_01= la collezione, parametro_02 = il tipo
                      Vv1 = Stampa_MSYS_N01_OGGETTI_funct(CollectionModuli_PFunct, "COLLECTION_MODULES")
                    
                    '//...............................................................................//
                    




        
        '//...............................................................................//
        

'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionModuli_PFunct_Exit:
    Exit Function

CollectionModuli_PFunct_Err:
    MsgBox Error$
    Resume CollectionModuli_PFunct_Exit

End Function


'//LA COLLEZIONE DELLE MODULI DA IMPORTARE  *** FINE ***
'//-----------------------------------------------------------------------------------//











'//LA COLLEZIONE DEI REPORT DA IMPORTARE  *** INIZIO ***
'//-----------------------------------------------------------------------------------//


'//PRECARICO LA COLLECTION REPORT DA IMPORTARE
Public Function CollectionReport_PFunct() As Collection          '//restituisco una collection
    

On Error GoTo CollectionReport_PFunct_Err
        
    Dim frm As Variant             '//OGGETTO REPORT DELLA COLLEZIONE PRECARICATA


    '//CODICE ----> COLLECTION_REPORT_PRECARICATE
    '//CREO LA COLLEZIONE E LA POPOLO
    '//I° COLLEZIONE : OGGETTI REPORT   @importa@collection@oggetti@Report_(la collezione dei @report)
    '//------------------------------------------------------------------------------//
    '//NOTE: questa collection, rappresenta gli OGGETTI REPORT  CHE APPARTENGONO al progetto _
            MSYS_ATTIVA_GEST_OGGETTI.mdb e che vengono precaricate in un collezione. Quindi tali _
            oggetti devono esere eliminati o aggiunti a seconda che le REPORT del progetto variano. _
            VANTAGGIO: il vantaggio di utilizzare la collezione facilita l'utilizzo delI REPORT di _
            importazione, in quanto le REPORT non sono impostate una tabella esterna che dovrebbe essere _
            spostata insime alI REPORT di importazione. MA ESSENDO PRECARICATE IN UNA COLLECTION si ha _
            sicurezza di spostare tutta la classe REPORT con i dati relativi in questo caso agli oggetti REPORT. _


    '//@OGGETTI@REPORT@PRECARICATI
    '//popola la collection delle REPORT di DA IMPORTARE
    ' Aggiungi i nomi delle REPORT o ELIMINA O COMMENTA LE REPORT DA ESCLUDERE. _
      SE VUOI UTILIZZARE LA COLLEZIONE anche per oggetti REPORT DI PROGETTI DIVERSI inserisci i nome delle REPORT _
      del progetto da importare. N.B. Rircordati di cambiare anche il parametro path+ file per rendere flessibile _
      questa collezione.
        
        Set objectreports = New Collection
        '//Aggiugno gli elementi della collezione
        
        

        '//2024.08.03_Select_Case_Me.Cmb_02_txt = SOSTITUITO IL VALORE CON LA CMB_02 al posto della CMB_01 _
                    Le collection viene scelta in base combinata 2 che individua il progetto _
                    al posto della combinata 1 e la sostituzione riguarda tutte le collection. _
                        @STEP@06_IMPORTA@OGGETTI_(Imposto la collection Reports)          = Oggetti Reports da importare
        Select Case Me.Cmb_02_txt
            Case "MSYS_ATTIVA_GEST_OGGETTI.mdb"
            
                '//@popolo@collection@REPORT@system_(delle REPORT di sistema, inserisco nella @collezione)_@carico@collection@REPORT
                objectreports.Add "MsysDbEstTb01Rpt00_}-------------------------------------------@"
                objectreports.Add "MsysDbEstTb01Rpt01_S01_OBJECT"
                objectreports.Add "MsysDbEstTb01Rpt01_S01_OBJECT_TMP"

            '//@COLLEZIONE@REPORT@OGGETTI@TREEVIEW
            '//@REPORT@TREEVIEW
            Case "IMPORTA_OGGETTI_TREEVIEW"
                
                objectreports.Add "Tree_Rpt00_}---------------------------------------------------@"
                objectreports.Add "Tree_Rpt01_01_PROVA"
                objectreports.Add "Tree_Rpt90_}---------------------------------------------------@"
                
            '//@COLLEZIONE@REPORT@OGGETTI@GE@OGGETTI
            '//@REPORTS@MSys
            Case "MSYS_ATTIVA_GEST_OGGETTI.mdb"
                
                objectreports.Add "MSys_Rpt00_{@------------------REPORT_DB----------------------@}"
                objectreports.Add "MSys_Rpt01_DF01_01_COMANDI_Qry01_01_SELECT_TUTTI"
                objectreports.Add "Msys_Rpt11_DF11_PROGETTI_Qry01_10_GROUP_TUTTI"
                objectreports.Add "Msys_Rpt12_DF12_Qry01_01_TIPO_OGGETTI}"
                
            
        '//@COLLEZIONE@REPORT@OGGETTI@DLL
        '//@IMPORTA@REPORT@DLL@COLECTION   = @IMPORTA@REPORT@DLL_(definisco gli oggetti collection REPORT per le @liberie@DLL) _

        Case "IMPORTA_OGGETTI_MSYS_DLL"
                '//2024.08.05_MODIFICHE_IMPORTA_DLL_(per l'importazione delle librerie DLL del progetto mysys inserito il blocco)
                '//@COLLECTION@DLL@FORM
                    objectreports.Add ""
        
        '//@COLLEZIONE@REPORT@OGGETTI@ESPROPRI
        '//@REPORT@ESPROPRI = I REPORT del progetto ESPROPRI
        '//@IMPORTA@REPORT@ESPROPRI@COLECTION   = 2024.08.06_COLLECTION@REPORT@ESPROPRI_(inserito la collezione REPORT per gli espropri)
        Case "IMPORTA_OGGETTI_ESPROPRI"
                
            '//@COLLECTION@ESPROPRI@REPORT
            objectreports.Add "LLPP_TB09Rpt00_}-----------------------------------------------@"
            objectreports.Add "LLPP_TB09Rpt00_01_ESPROPRI"

      
            
        Case Else
            
            MsgBox "ATTENZIONE! NON CI SONO REPORT PER QUESTO PROGETTO: " & Me.Cmb_01_txt, vbCritical, "AVVISO)"
                '//2024.08.10 =COLLECTION=0- prima di uscire imposto la collezione per evitare di restiuire null e poi esco dalla routine.
                Set CollectionReport_PFunct = objectreports
                '//SE NON CI SONO OGGETTI USCITA DALLA ROUTINE
                GoTo CollectionReport_PFunct_Exit
            
            
        End Select
   
        ' Inizializza le collezioni per REPORT COLLEGATE
        Set connectedREPORTS = New Collection
        
      
    '// *** FINE ****
    '//I° COLLEZIONE DI REPORT DI SISTEMA E LA POPOLO
    '//------------------------------------------------------------------------------//
    
        '//CONTROLLO DELLE REPORT PRECARICATE NELLA COLLEZIONE
        '//...............................................................................//
        '//NOTE: queste sono le REPORT da carica nel nuovo progetto.
                
                    
                    '//RIEPILOGO COLLEZIONI
                    '//---------------------------------------------------------//
                           
                            ' Stampa le REPORT fisiche
                            Debug.Print
                            Debug.Print
                            Debug.Print "REPORT DA IMPORTARE:"
                            '//iterazione nella collezione per il controllo
                            For Each frm In objectreports
                                Debug.Print "I REPORT da importare nel progetto: " & frm
                            Next frm
        
                    '//---------------------------------------------------------//
                    
                     
                     

    
        '//CONTROLLO I_REPORT PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
        '//NOTE: queste sono I_REPORT da carica nel nuovo progetto.
                        
                        '//RESET CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.TXT_OBJECT_ESTERNI.Requery
        
                                    
                  '//IMPOSTO LA PATH + FILE.MDB + TESTATA + OGGETTI VISUALIZZAZIONE DELLA CASELLA TXT
                  '//--------------------------------------------------------------------------------------------//
                  '//Note: @imposto@path_(imposto la path degli oggetti da recupeare, _
                           preparo la testata per la casela di testo, preparo gli oggetti e visualizzo tutto _
                           nella casella di testo.
                    
                    
                    '//RIEPILOGO COLLEZIONI E VISUALIZZAZIONE PATH ED OGGETTI PRECARICATI
                        Str1 = ""
                        Str1 = Str1 & "CONTROLLO PATH E DATABASE DA CUI RECUPERARE GLI OGGETTI" & vbCrLf
                        Str1 = Str1 & "-----------------------------------------------------------------" & vbCrLf
                        Str1 = Str1 & Me.sourceDBPath_s_Txt & vbCrLf
                        
                        '//INSERISCO LA PATH NELLA CASELLA DI TESTO ED AGGIORNO
                        Me.TXT_OBJECT_ESTERNI = Str1
                        Me.TXT_OBJECT_ESTERNI.Requery

                            '//PREPARO LA TESTATA DELLE COLLECTION PER IL II° INSERIMENTO
                            Str1 = Str1 & vbCrLf & "                          COLLECTION I_REPORT DA IMPORTARE : " & vbCrLf
                            Str1 = Str1 & "//=========================================================================//" & vbCrLf
                            ' Stampa
                            Debug.Print
                            Debug.Print
                            Debug.Print Str1
                            
                            '//iterazione nella collezione per il controllo
                            For Each objRpt In objectreports

                                Debug.Print "I_REPORT da importare nel progetto: " & objRpt
                                Str1 = Str1 & objRpt & vbCrLf
                                
                            Next objRpt
                            
                            
                            
                  '//--------------------------------------------------------------------------------------------//
                        
        '// *** FINE ***
        '//CONTROLLO I_REPORT PRECARICATE NELLA COLLEZIONE CARICO PATH + DB E COLLECTION NELLA CASELLA DI TESTO
        '//=====================================================================================================================//
                                             
                     
                     
                     
                     ' Restituisci la collezione
                        Set CollectionReport_PFunct = objectreports
        

                    '//CHIAMO LA PROCEDURA UNICA DI STAMPA E CARICAMENTO NELLA CASELLA DI TESTO _
                            parametro_01= la collezione, parametro_02 = il tipo
                      Vv1 = Stampa_MSYS_N01_OGGETTI_funct(CollectionModuli_PFunct, "COLLECTION_REPORTS")
       



        
'//USCITA ED ERRORI
'//--------------------------------------------------------------------------------------//

CollectionReport_PFunct_Exit:
    Exit Function

CollectionReport_PFunct_Err:
    MsgBox Error$
    Resume CollectionReport_PFunct_Exit

End Function













'//=====================================================================================================================//
'//                     CASELLE DI TESTO E COMBINATE DEL PROGETTO IMPORTA  *** INIZIO ***
'//                     + CurrentProject_txt + Cmb_01_txt + Cmb_02_txt + sourceDBPath_s_Txt  = @CMB_01, CMB_02 @CMB @PROGETTO
'//=====================================================================================================================//
'//NOTE: @caselle@combinate@progetto_(le caselle di testo e le caselle combinate che gestiscono il progetto + @cmb@help)





'//=====================================================================================================================//
'//                                             @CMB_01      *** INIZIO ***
'//
'//=====================================================================================================================//

'//CASELLA COMBINATA DI SCELTA DEL PROGETTO @CMB_01@CHANGE
'//-------------------------------------------------------------------------------------------//
'//@PATH@PROGETTI_(la path dei progetti da caricare)'
'//@CHANGE@CACELLA@COMBINATA_(se cambia la @casella@combinata abilito o disabilito i button)
'//NOTE
Private Sub Cmb_01_txt_AfterUpdate()

         
   
'//DIM della collezione form da importare
Dim formsCollection As Collection
Dim formName_v As Variant

'//DIM la collezione macro
Dim macroCollection As Collection
Dim MACROS_Name_v As Variant



'//DIM la collezione DEI MODULI
Dim Modules_Collection As Collection
Dim moduloName_v As Variant


'//DIM la collezione DELLE QUERY
Dim queryCollection As Collection
Dim queryName_v As Variant




'//DIM la collezione DELLE REPORT
Dim reportCollection As Collection
Dim reportName_v As Variant



'//DIM la collezione DELLE TABELLE
Dim tableCollection As Collection
Dim tableName_v As Variant
Dim CodiceProgetto_s As String



            
    '//RESET I BUTTON
    '//........................................//
    '//note : RESET VARIABILI
        PROGETTO_s = ""
        
        '//RESET
        Me.TXT_OBJECT_ESTERNI = ""
       
       
    '//........................................//
    
    
    '//@CMB_01_@DOPO@AGGIORNAMENTO
    '//ABILITO O DISABILITO I COMAND BUTTON + IMPOSTO PATH E DB PER LA CASELLA DI TESTO - LA COMBINATA 01 - @CMB_01
    '//===============================================================================================//
    '//NOTE: @scelta@progetto_(casella combinata di scelta del progetto)
    '//      @imposta@path_(per tutti i progetti vengon impostati nella combinata 01) _
    @STEP@02_IMPORTA@OGGETTI_(Cmb_01 valori casella)                  = Evento della Cmb_01 imposto i valori della casella
    
    Select Case Me.Cmb_01_txt.Value
    
        
        
            '//@ELENCO@PROGETTI
            '//2024.08.03_Cmb_01_txt_AfterUpdate = inserito il case ELENCO PROGETTI
            '//@MENU@PROGETTI
            Case "ELENCO PROGETTI"
                        
                        
                        
                        'se non esiste nessun progetto disabilito i button ed imposto la path a null
                        sourceDBPath_s = ""
                        '//La variabile path aggiornata la salvo nella casella di testo:
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.sourceDBPath_s_Txt = ""
                        Me.sourceDBPath_s_Txt.Requery
                        
                        '//2024.08.03_Disabilito_sourceDBPath_Sub = creato la Disabilitazione della path
                        Call Disabilito_sourceDBPath_Sub      '//DISABILITO LA PATH
                        Call DisabilitaButton_Sub
                        
                        Me.Cmb_02_txt = ""
                        Me.Cmb_02_txt.Requery
                        
            '//@MENU@OGGETTI
            Case "MSYS_ATTIVA_GEST_OGGETTI.mdb"
                        
                        
                        '//TODO: nel caso di scelta DI MSYS occorre anticipare il controllo progetto, perche deve essere _
                                ed inibita l'importazione o la cancellazione degli oggetti.
                        
                        
                        
                        '//IMPOSTO IL TITOLO DEL PROGETTO
                        PROGETTO_s = "PROGETTO MSYS_N01_OGGETTI"
                        
                        '//codice progetto
                        CodiceProgetto_s = "MSys"
                        Me.CODICE_PROGETTO_TXT = CodiceProgetto_s
                        
                        sourceDBPath_s = ""
                        Call DisabilitaButton_Sub
                          
                        '//MSYS_ATTIVA_GEST_OGGETTI.mdb = PER QUESTO DATABASE ABILITO LA PATH
                        ' Percorso del database di origine - todo: eliminato
                        sourceDBPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_ATTIVA_GEST_OGGETTI.mdb"
                        Me.sourceDBPath_s_Txt = sourceDBPath_s
                        
                        
                                                
                        '//IMPOSTO LA @LABEL@TITOLO@PRINCIPALE_(@LABEL@MASTER VIENE REIMPOSTATA)
                        Me.LBL_TIT01_TXT.ForeColor = RGB(255, 0, 0)   '//rosso
                        Me.LBL_TIT01_TXT.FontSize = 12
                        Me.LBL_TIT01_TXT.Caption = PROGETTO_s & " | codice progetto ---> " & CodiceProgetto_s & " | la path: " & sourceDBPath_s

                        
                        
                        '//Abilito la cmb_02 + la path
                        AbilitaCmd_02_Sub
                        
                        '//2024.08.03_Abilito_sourceDBPath_Sub = creato l'abilitazione della path
                        Call Abilito_sourceDBPath_Sub      '//ABILITO LA PATH
                        Call AbilitaButton_CMB_02_Sub      '//ABILITO LA CMB_02
    
        
                         '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
                         '//@CMB_02@IMPOSTO@_(I VALORI DELLA COMBINATA 02)
                         '//-------------------------------------------------------------------------//
                         '//@CMB_02@imposto_(reimposto agli oggetti TreeView della @casella@combinata)
                             Me.Cmb_02_txt.Value = ""
                             Str1 = ""
                             Str1 = Str1 & "OGGETTI MSYS;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_MSYS_DLL;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_MSYS_02??;"
                             Str1 = Str1 & ""
                         
                         
                             Me.Cmb_02_txt.RowSourceType = "elenco valori"
                             Me.Cmb_02_txt.RowSource = Str1
                             Me.Cmb_02_txt.Value = "OGGETTI MSYS"
                         '//-------------------------------------------------------------------------//
                            
                             '//RESET
                            Me.TXT_OBJECT_ESTERNI = ""
                            Me.Recalc
                            
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE  @CONTROLLO@PROGETTO
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                       
            '//@MENU@ESPROPRI
            Case "GE_ESPROPRI.mdb"
                                        
                        '//RESET
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.Recalc
                            
                    
                        '//IMPOSTO IL TITOLO DEL PROGETTO
                        PROGETTO_s = "PROGETTO ESPROPRI"
                        
                        '//codice progetto
                        CodiceProgetto_s = "ESPROPRI??"
                        Me.CODICE_PROGETTO_TXT = CodiceProgetto_s
                        
                        
                        '//Abilito la cmb_02 + la path + imposto la path
                        Call DisabilitaButton_Sub
                        Call Abilito_sourceDBPath_Sub      '//ABILITO LA PATH
                        Call AbilitaButton_CMB_02_Sub      '//ABILITO LA CMB_02
                        
                        sourceDBPath_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\ESPROPRI\MDB\GE_ESPROPRI.mdb"       '//percorso ESPROPRI
                        Me.sourceDBPath_s_Txt = sourceDBPath_s
                    
        
                         '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
                         '//@CMB_02@IMPOSTO@_(I VALORI DELLA COMBINATA 02)
                         '//-------------------------------------------------------------------------//
                         '//@CMB_02@imposto_(reimposto agli oggetti TreeView della @casella@combinata)
                             Me.Cmb_02_txt.Value = ""
                             Str1 = ""
                             Str1 = Str1 & "OGGETTI ESPROPRI;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_ESPORPRI_01??;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_ESPORPRI_02??;"
                             Str1 = Str1 & ""
                         
                         
                             Me.Cmb_02_txt.RowSourceType = "elenco valori"
                             Me.Cmb_02_txt.RowSource = Str1
                             Me.Cmb_02_txt.Value = "OGGETTI ESPROPRI"
                         '//-------------------------------------------------------------------------//
                            
                             '//RESET
                            Me.TXT_OBJECT_ESTERNI = ""
                            Me.Recalc
                       
            '//@MENU@TREEVIEW
            Case "TREEVIEW.mdb"
                    
                    '//RESET
                    Me.TXT_OBJECT_ESTERNI = ""
                        
                    
                        '//IMPOSTO IL TITOLO DEL PROGETTO
                        PROGETTO_s = "PROGETTO TREEVIEW"
                        
                        
                        '//codice progetto
                        CodiceProgetto_s = "TREE_VIEW"
                        Me.CODICE_PROGETTO_TXT = CodiceProgetto_s
                        
                        
                        '//Abilito la cmb_02 + la path + imposto la path
                        Call DisabilitaButton_Sub
                        sourceDBPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\TREVIEW\MDB\TREEVIEW.mdb"       '//percorso Treeview
                        
                        Call Abilito_sourceDBPath_Sub      '//ABILITO LA PATH
                        Call AbilitaButton_CMB_02_Sub      '//ABILITO LA CMB_02
                        
                        Me.sourceDBPath_s_Txt = sourceDBPath_s
                        Me.sourceDBPath_s_Txt.Requery
                        
        
                         '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
                         '//@CMB_02@IMPOSTO@_(I VALORI DELLA COMBINATA 02)
                         '//-------------------------------------------------------------------------//
                         '//@CMB_02@imposto_(reimposto agli oggetti TreeView della @casella@combinata)
                             Me.Cmb_02_txt.Value = ""
                             Str1 = ""
                             Str1 = Str1 & "OGGETTI TREVIEW;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_TREEVIEW;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_??;"
                             Str1 = Str1 & ""
                         
                             Me.Cmb_02_txt.RowSourceType = "elenco valori"
                             Me.Cmb_02_txt.RowSource = Str1
                             Me.Cmb_02_txt.Value = "OGGETTI TREVIEW"
                         '//-------------------------------------------------------------------------//
                        
                            
                             '//RESET
                            Me.TXT_OBJECT_ESTERNI = ""
                            Me.Recalc
                       
                    
            '//@MENU@ESPROPRI
            Case "GE_ESPROPRI.mdb"
                    
                      '//GE_ESPROPRI_MSYS_ATTIVA_GEST_OGGETTI.mdb = PER QUESTO DATABASE ABILITO LA PATH
                      ' Percorso del database di origine - todo: eliminato
                      sourceDBPath_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\ESPROPRI\MDB\GE_ESPROPRI.mdb"
                      Call DisabilitaButton_Sub
                        
                        Call Abilito_sourceDBPath_Sub      '//ABILITO LA PATH
                        Call AbilitaButton_CMB_02_Sub      '//ABILITO LA CMB_02
                        
                        
                        '//IMPOSTO IL TITOLO DEL PROGETTO
                        PROGETTO_s = "PROGETTO ESPROPRI"
                           
                        '//codice progetto
                        CodiceProgetto_s = "ESPROPRI?"
                        Me.CODICE_PROGETTO_TXT = CodiceProgetto_s
                        
                           
                           
                        '//VISUALIZZO NELLA CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = Stampa_N02_OGGETTI_ESPROPRI_funct
                        
                        
                         '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
                         '//@CMB_02@IMPOSTO@_(I VALORI DELLA COMBINATA 02)
                         '//-------------------------------------------------------------------------//
                         '//@CMB_02@imposto_(reimposto agli oggetti TreeView della @casella@combinata)
                             Me.Cmb_02_txt.Value = ""
                             Str1 = ""
                             Str1 = Str1 & "OGGETTI ESPROPRI;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_EPROPRI_01??;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_EPROPRI_02??;"
                             Str1 = Str1 & ""
                         
                         
                             Me.Cmb_02_txt.RowSourceType = "elenco valori"
                             Me.Cmb_02_txt.RowSource = Str1
                             Me.Cmb_02_txt.Value = "OGGETTI ESPROPRI"
                         '//-------------------------------------------------------------------------//
                        

                             '//RESET
                            Me.TXT_OBJECT_ESTERNI = ""
                            Me.Recalc
                        
                '//@MENU@CONDOMINIO
                Case "CDM"
                    
                    '//GE_ESPROPRI_MSYS_ATTIVA_GEST_OGGETTI.mdb = PER QUESTO DATABASE ABILITO LA PATH
                      ' Percorso del database di origine - todo: eliminato
                      sourceDBPath_s = "c:\CASA\CDM\GE_COND\"
                        Call DisabilitaButton_Sub
                        Call Abilito_sourceDBPath_Sub      '//ABILITO LA PATH
                        Call AbilitaButton_CMB_02_Sub      '//ABILITO LA CMB_02
                          
                        '//IMPOSTO IL TITOLO DEL PROGETTO
                        PROGETTO_s = "PROGETTO CONDOMINIO"
                           
                        '//codice progetto
                        CodiceProgetto_s = "CDM_ANGRAFE?"
                        Me.CODICE_PROGETTO_TXT = CodiceProgetto_s
                        
                           
                        '//VISUALIZZO NELLA CASELLA DI TESTO
                        Me.TXT_OBJECT_ESTERNI = Stampa_N03_OGGETTI_CONDOMINIO_funct
                        
                        
                         '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
                         '//@CMB_02@IMPOSTO@_(I VALORI DELLA COMBINATA 02)
                         '//-------------------------------------------------------------------------//
                         '//@CMB_02@imposto_(reimposto agli oggetti TreeView della @casella@combinata)
                             Me.Cmb_02_txt.Value = ""
                                Str1 = ""
                                Str1 = Str1 & "OGGETTI CONDOMINIO;"
                                Str1 = Str1 & "IMPORTA_OGGETTI_CDM_GECOND;"
                                Str1 = Str1 & "IMPORTA_OGGETTI_CDM_GIUR;"
                                Str1 = Str1 & "IMPORTA_OGGETTI_CDM_CONS;"
                                Str1 = Str1 & "IMPORTA_OGGETTI_CDM_ANAGR;"
                                Str1 = Str1 & "IMPORTA_OGGETTI_CDM_GE_PREV;"
                                Str1 = Str1 & "IMPORTA_OGGETTI_CDM_GE_TABELLE;"
                                Str1 = Str1 & "IMPORTA_OGGETTI_CDM_GE_UTIL;"
                                Str1 = Str1 & ""
                         
                         
                             Me.Cmb_02_txt.RowSourceType = "elenco valori"
                             Me.Cmb_02_txt.RowSource = Str1
                             Me.Cmb_02_txt.Value = "OGGETTI CONDOMINIO"
                         '//-------------------------------------------------------------------------//
                        

                             '//RESET
                            Me.TXT_OBJECT_ESTERNI = ""
                            Me.Recalc
                            
            '//2025.02.03_CMB_02_=INSERITO IL MENU IMPORTA IMPEGNI
            '//@MENU@IMPEGNI
            Case "IMPEGNI"
                    
                    '//RESET
                    Me.TXT_OBJECT_ESTERNI = ""
                        
                    
                        '//IMPOSTO IL TITOLO DEL PROGETTO
                        PROGETTO_s = "PROGETTO IMPEGNI"
                        
                        
                        '//codice progetto
                        CodiceProgetto_s = "IMPEGNI"
                        Me.CODICE_PROGETTO_TXT = CodiceProgetto_s
                        
                        
                        '//Abilito la cmb_02 + la path + imposto la path
                        Call DisabilitaButton_Sub
                        sourceDBPath_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\IMPEGNI\MDB\LLPP_IMPEGNI_ATTIVA_GESTIONE.mdb"       '//percorso Treeview
                        
                        Call Abilito_sourceDBPath_Sub      '//ABILITO LA PATH
                        Call AbilitaButton_CMB_02_Sub      '//ABILITO LA CMB_02
                        
                        Me.sourceDBPath_s_Txt = sourceDBPath_s
                        Me.sourceDBPath_s_Txt.Requery
                        
        
                         '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
                         '//@CMB_02@IMPOSTO@_(I VALORI DELLA COMBINATA 02)
                         '//-------------------------------------------------------------------------//
                         '//@CMB_02@imposto_(reimposto agli oggetti TreeView della @casella@combinata)
                             Me.Cmb_02_txt.Value = ""
                             Str1 = ""
                             Str1 = Str1 & "OGGETTI IMPEGNI;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_IMPEGNI;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_??;"
                             Str1 = Str1 & ""
                         
                             Me.Cmb_02_txt.RowSourceType = "elenco valori"
                             Me.Cmb_02_txt.RowSource = Str1
                             Me.Cmb_02_txt.Value = "OGGETTI IMPEGNI"
                         '//-------------------------------------------------------------------------//
                        
                            
                             '//RESET
                            Me.TXT_OBJECT_ESTERNI = ""
                            Me.Recalc
                            
                            
                        '//@MENU@PRES3000
            Case "PRES3000"
                    
                    '//RESET
                    Me.TXT_OBJECT_ESTERNI = ""
                        
                    
                        '//IMPOSTO IL TITOLO DEL PROGETTO
                        PROGETTO_s = "PROGETTO PRES3000"
                        
                        
                        '//codice progetto
                        CodiceProgetto_s = "PRES3000"
                        Me.CODICE_PROGETTO_TXT = CodiceProgetto_s
                        
                        
                        '//Abilito la cmb_02 + la path + imposto la path
                        Call DisabilitaButton_Sub
                        sourceDBPath_s = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\PRES3000_GESTIONE\ARCHIVI_PRES3000\PRES3000_N00_GESTIONE\GESTIONE_MDB\PRES3000_N00_GESTIONE.mdb"       '//percorso PRES3000
                        
                        Call Abilito_sourceDBPath_Sub      '//ABILITO LA PATH
                        Call AbilitaButton_CMB_02_Sub      '//ABILITO LA CMB_02
                        
                        Me.sourceDBPath_s_Txt = sourceDBPath_s
                        Me.sourceDBPath_s_Txt.Requery
                        
        
                         '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
                         '//@CMB_02@IMPOSTO@_(I VALORI DELLA COMBINATA 02)
                         '//-------------------------------------------------------------------------//
                         '//@CMB_02@imposto_(reimposto agli oggetti PRES3000 della @casella@combinata)
                             Me.Cmb_02_txt.Value = ""
                             Str1 = ""
                             Str1 = Str1 & "OGGETTI PRES3000;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_PRES3000;"
                             Str1 = Str1 & "IMPORTA_OGGETTI_??;"
                             Str1 = Str1 & ""
                         
                             Me.Cmb_02_txt.RowSourceType = "elenco valori"
                             Me.Cmb_02_txt.RowSource = Str1
                             Me.Cmb_02_txt.Value = "OGGETTI PRES3000"
                         '//-------------------------------------------------------------------------//
                        
                            
                             '//RESET
                            Me.TXT_OBJECT_ESTERNI = ""
                            Me.Recalc
                         
        
            Case Else
                        'se non esiste nessun progetto disabilito i button ed imposto la path a null
                        sourceDBPath_s = ""
                        '//La variabile path aggiornata la salvo nella casella di testo:
                        Me.TXT_OBJECT_ESTERNI = ""
                        Me.sourceDBPath_s_Txt = ""
                        Me.sourceDBPath_s_Txt.Requery
                         
                        
                        Me.Cmb_02_txt = ""
                        Me.Cmb_02_txt.Requery
                        
                         
                        
                    
                    
                        
                        '//DISABILITO LA CMB_02
                        Call Disabilito_sourceDBPath_Sub      '//DISABILITO LA PATH
                        Call DisabilitaButton_CMB_02_Sub      '//DISABILITO LA CMB_02
                                       
                        
            End Select
            
                    
        
            '//ABILITO O DISABILITO I COMAND BUTTON + IMPOSTO PATH E DB PER LA CASELLA DI TESTO
            '//===============================================================================================//
                          

End Sub


'//HELP@CMB_01
Private Sub Cmb_01_txt_HELP_Click()
'//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = ""
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\TUTORIAL\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "APRI_GIORNALE_OPERAZIONI_PROGETTO_MSYS_{Word2000}.bat"
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub


'//=====================================================================================================================//
'//                                             @CMB_01      *** FINE ***
'//
'//=====================================================================================================================//





'//=====================================================================================================================//
'//                                             @CMB_02      *** INIZIO ***
'//
'//=====================================================================================================================//





'//todo: finre AL POSTO DI CREARE QUESTO EVENTO OCCORRE alla scelta IMPORTA OGGETT MSYS DLL, ossia importa _
         le tabelle, query, form, report, moduli e macro utili per la lavorazione delle librerie ad oggetti _
            1) abilitare il comando di importazione delle tabelle e di cancellazione.
            
'//CMB_02= CASELLA COMBINATA DI SCELTA DEL PROGETTO @CMB_02@CHANGE
'//-------------------------------------------------------------------------------------------//
'//Note: la combinata 02 deve solo filtrare gli oggetti che appartengono al progetto principale senza _
        commistioni es. SE LA CMB_01 imposta il progetto MSYS la CMB_02 deve filtrare solo OGGETTI MSYS e non _
        altro come ad esempio oggetti TREEVIEW che sarebbe errato in quanto la path impostata DALLA COMBINATA 01 è per gli _
        oggetti MSYS  e non per i Treeview.



Private Sub Cmb_02_txt_DblClick(Cancel As Integer)
    Cmb_02_txt_AfterUpdate
End Sub

'//@STEP@07_IMPORTA@OGGETTI_(Abilito i Comandi)                      = Abilito i button Importa e esporta
Private Sub Cmb_02_txt_AfterUpdate()

    Select Case Me.Cmb_02_txt.Value
        
        '//2024.08.08_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
        '//IMPORTA OGGETTI LIBRERIE DLL  @CMB_02@IMPORTA@MSYS
        '//........................................................//
        Case "IMPORTA_OGGETTI_MSYS_DLL"
            
            '01_ABLITAZIONE_BUTTON= solo i comandi IMPORT e DELETE
            AbilitaButton_TABLE_Sub
            AbilitaButton_MODULI_Sub
            AbilitaButton_QUERY_Sub
            AbilitaButton_MACRO_Sub
            AbilitaButton_IMPORTA_TUTTO_Sub
            
            
            AbilitaButton_CMB_02_Sub
    
            
        '//........................................................//
                
         '//2025.07.10_CMB_02_SELEZIONE_OGGETTI_IMPOSTA_PATH_MODIFICA
        '//IMPORTA OGGETTI LIBRERIE DLL  @CMB_02@IMPORTA@MSYS
        '//........................................................//
        Case "IMPORTA_OGGETTI_PRES3000"
            
            '01_ABLITAZIONE_BUTTON= solo i comandi IMPORT e DELETE
            AbilitaButton_TABLE_Sub
            AbilitaButton_MODULI_Sub
            AbilitaButton_QUERY_Sub
            AbilitaButton_MACRO_Sub
            AbilitaButton_IMPORTA_TUTTO_Sub
            
            
            AbilitaButton_CMB_02_Sub
    
            
        '//........................................................//
              
                
        '//IMPORTA OGGETTI IMPEGNI
        '//........................................................//
        Case "IMPORTA_OGGETTI_IMPEGNI"
            
            '01_ABLITAZIONE_BUTTON= solo i comandi IMPORT e DELETE
            AbilitaButton_FORM_Sub
            AbilitaButton_TABLE_Sub
            AbilitaButton_MODULI_Sub
            AbilitaButton_QUERY_Sub
            AbilitaButton_MACRO_Sub
            AbilitaButton_IMPORTA_TUTTO_Sub
            
            AbilitaButton_CMB_02_Sub
    
        '//........................................................//
        
                
        '//IMPORTA OGGETTI TREEVIEW
        '//........................................................//
        Case "IMPORTA_OGGETTI_TREEVIEW"
            
            '01_ABLITAZIONE_BUTTON= solo i comandi IMPORT e DELETE
            AbilitaButton_FORM_Sub
            AbilitaButton_TABLE_Sub
            AbilitaButton_MODULI_Sub
            AbilitaButton_QUERY_Sub
            AbilitaButton_MACRO_Sub
            AbilitaButton_IMPORTA_TUTTO_Sub
            
            AbilitaButton_CMB_02_Sub
    
        '//........................................................//
        
        
        '//IMPORTA OGGETTI CONDOMINIO  @CMB_@02@CDM
        '//........................................................//
        Case "IMPORTA_OGGETTI_CDM_GECOND", _
              "IMPORTA_OGGETTI_CDM_GIUR", _
              "IMPORTA_OGGETTI_CDM_CONS", _
              "IMPORTA_OGGETTI_CDM_ANAGR", _
              "IMPORTA_OGGETTI_CDM_BIL", _
              "IMPORTA_OGGETTI_CDM_GE_PREV", _
              "IMPORTA_OGGETTI_CDM_GE_TABELLE", _
              "IMPORTA_OGGETTI_CDM_GE_UTIL", _
              ""
              


            
            '01_ABLITAZIONE_BUTTON= solo i comandi IMPORT e DELETE
            AbilitaButton_FORM_Sub
            AbilitaButton_TABLE_Sub
            AbilitaButton_MODULI_Sub
            AbilitaButton_QUERY_Sub
            AbilitaButton_MACRO_Sub
            AbilitaButton_IMPORTA_TUTTO_Sub
            
            AbilitaButton_CMB_02_Sub
            
 
            
            If Me.Cmb_02_txt.Value = "IMPORTA_OGGETTI_CDM_GIUR" Then
               Me.sourceDBPath_s_Txt = "c:\CASA\CDM\GE_COND\GIUR\MDB\GE_GIUR.mdb"
            End If
            
            If Me.Cmb_02_txt.Value = "IMPORTA_OGGETTI_CDM_CONS" Then
                   Me.sourceDBPath_s_Txt = "c:\CASA\CDM\GE_COND\CONS\MDB\GE_CONS.mdb"
            End If
            
            
            If Me.Cmb_02_txt.Value = "IMPORTA_OGGETTI_CDM_ANAGR" Then
               Me.sourceDBPath_s_Txt = "c:\CASA\CDM\GE_COND\ANAGR\MDB\GE_ANAGR.mdb"
            End If
            
            If Me.Cmb_02_txt.Value = "IMPORTA_OGGETTI_CDM_BIL" Then
               Me.sourceDBPath_s_Txt = "c:\CASA\CDM\GE_COND\BIL\MDB\GE_BIL.mdb"
            End If
            
            
            If Me.Cmb_02_txt.Value = "IMPORTA_OGGETTI_CDM_GECOND" Then
               Me.sourceDBPath_s_Txt = "c:\CASA\CDM\GE_COND\GECOND\MDB\GECOND.mdb"
            End If
            
            
            If Me.Cmb_02_txt.Value = "IMPORTA_OGGETTI_CDM_GE_PREV" Then
               Me.sourceDBPath_s_Txt = "c:\CASA\CDM\GE_COND\PREV\MDB\GE_PREV.mdb"
            End If
            
            
            
             If Me.Cmb_02_txt.Value = "IMPORTA_OGGETTI_CDM_GE_TABELLE" Then
                Me.sourceDBPath_s_Txt = "c:\CASA\CDM\GE_COND\TABELLE\MDB\GE_TABELLE_MILL.mdb"
             End If
             
             
             If Me.Cmb_02_txt.Value = "IMPORTA_OGGETTI_CDM_GE_UTIL" Then
                Me.sourceDBPath_s_Txt = "c:\CASA\CDM\GE_COND\UTILIT\MDB\GE_UTIL.mdb"
             End If
             
             If Me.Cmb_02_txt.Value = "IMPORTA_OGGETTI_CDM_GECOND" Then
                Me.sourceDBPath_s_Txt = "c:\CASA\CDM\GE_COND\GECOND\MDB\GECOND.mdb"
             End If
             
             
             If Me.Cmb_02_txt.Value = "IMPORTA_OGGETTI_IMPEGNI" Then
                Me.sourceDBPath_s_Txt = "c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\IMPEGNI\MDB\LLPP_IMPEGNI_ATTIVA_GESTIONE.mdb"
             End If
             
             
           
            
        '//........................................................//
        
    Case Else
        
        '//se non ci sono oggetti da importare disabilito tutti i button
                DisabilitaButton_FORM_Sub
                DisabilitaButton_TABLE_Sub
                DisabilitaButton_MODULI_Sub
                DisabilitaButton_QUERY_Sub
                DisabilitaButton_MACRO_Sub
                DisabilitaButton_IMPORTA_TUTTO_Sub
                DisabilitaButton_DELETE_TUTTO_Sub
                
                 '//RITORNO ALLA COMBINATA 01
                Me.Cmb_01_txt.SetFocus
                      
                DisabilitaButton_CMB_02_Sub
                Disabilito_sourceDBPath_Sub
                
    
    End Select
    
    
End Sub




'//CMB_02 = CASELLA COMBINATA DI SCELTA DEL PROGETTO   *** FINE ***

'//-------------------------------------------------------------------------------------------//



'//=====================================================================================================================//
'//                                             @CMB_02      *** FINE ***
'//
'//=====================================================================================================================//











'//@CANCELLA@FILE@LOG_(cancello tutti i file denominati LOG_IMPORTA *.*)
'//NOTE: cancello tutti i file .txt LOG costruiti per l'importazione degli oggetti query, form _
         report ecc. che si trovano nella directory corrente recuperara del database _
         corrente aperto con il comando : CurrentProject.PATH & "\"

Private Sub Cmd_Delete_FILE_LOG_Click()
    
    Dim strFolderPath As String
    Dim strFile As String
    Dim strElenco As String
    
    ' Ottiene la directory del database attualmente aperto
    strFolderPath = CurrentProject.PATH & "\"
    
    ' Cerca i file che iniziano con "LOG_IMPORTA"
    strFile = Dir(strFolderPath & "LOG_IMPORTA*.*")
    
    ' Controlla se ci sono file da eliminare
    If strFile = "" Then
        MsgBox "Nessun file trovato da eliminare nella cartella: " & vbCrLf & strFolderPath, vbExclamation, "Operazione conclusa"
        Exit Sub
    End If
    
    ' Cancella i file e crea l'elenco
    Do While strFile <> ""
        Kill strFolderPath & strFile
        strElenco = strElenco & strFile & vbCrLf  ' Aggiunge il file all'elenco
        strFile = Dir  ' Trova il prossimo file
    Loop
    
    ' Mostra l'elenco dei file cancellati
    MsgBox "I seguenti file sono stati eliminati dalla cartella:" & vbCrLf & strFolderPath & vbCrLf & vbCrLf & strElenco, vbInformation, "Eliminazione completata"
End Sub



'//CANCELLO I FILE LOG .TXT
Private Sub Cmd_Delete_FILE_LOG_HELP_Click()

    Dim strFolderPath As String
    Dim strFile As String
    Dim strElenco As String
    
    ' Ottiene la directory del database attualmente aperto
    strFolderPath = CurrentProject.PATH & "\"
    
    ' Cerca i file che iniziano con "LOG_IMPORTA"
    strFile = Dir(strFolderPath & "LOG_IMPORTA*.*")
    
    ' Controlla se ci sono file da eliminare
    If strFile = "" Then
        MsgBox "Nessun file trovato da eliminare nella cartella: " & vbCrLf & strFolderPath, vbExclamation, "Operazione conclusa"
        Exit Sub
    End If
    
    ' Cancella i file e crea l'elenco
    Do While strFile <> ""
        Kill strFolderPath & strFile
        strElenco = strElenco & strFile & vbCrLf  ' Aggiunge il file all'elenco
        strFile = Dir  ' Trova il prossimo file
    Loop
    
    ' Mostra l'elenco dei file cancellati
    MsgBox "I seguenti file sono stati eliminati dalla cartella:" & vbCrLf & strFolderPath & vbCrLf & vbCrLf & strElenco, vbInformation, "Eliminazione completata"
End Sub



'//***********************************************************************************************//
'//*        -----------------------------------------------------------------------------------
'//*                     ROUTINE PER AUMENTARE E RIDURRE LE DIMENSIONI      *** INIZIO ***
'//*
'//*        -----------------------------------------------------------------------------------
'//*
'//***********************************************************************************************//
'//*  CODICE ----> @AUMENTA@DIMENSIONI,  @RIDUCI@DIMENSIONI@FORM
'//*
'//*  FAQ: _
         come aumentare le dimensione della form _
         come RIDURRE le dimensione della form

'//*
'//* ROUTINE COSTRUITE:
'//* @01_@AUMENTA@DIMENSIONI              = AUMENTO LE DIMENSIONE DELLE FORM MASTER E DEGLI OGGETTI INCORPORATI
'//* @02_@RIDUCI@DIMENSIONI               = RIDUCO LE DIMENSIONE DELLE FORM MASTER E DEGLI OGGETTI INCORPORATI
'//* @03_@CONTROLLO@ESISTENZA@FORM        = AUMENTO O RIDUCO LA DIMENSIONE DELLE FORM SOLO SE ESISTONO NEL DB CORRENTE
'//* @04_@CMB@02@FORM@LARGHEZZA@ALTEZZA   = @Cmb_02_FORM_L_H_TXT @LARGHEZZA E @ALTEZZA DELLA @FORM @CMB@02_(combinata @misure @L@H della form)




Private Sub Cmb_01_FORM_TXT_GotFocus()
    On Error GoTo ErrHandler ' Attiva la gestione degli errori

    ' Verifico se la casella combinata esiste
    If Not Me.Cmb_01_FORM_TXT Is Nothing Then
        ' Pulisco il valore corrente
        Me.Cmb_01_FORM_TXT.Value = ""
        
        '//@CONTROLLO@FORM@DIMENSIONI_(individuo la form da @RIDURRE oppure da @AUMENTARE nelle dimensioni)
        ' Definisco la stringa con i dati per la combo _
            @IMPORTA@OGGETTI_01_(Nome oggetti)
        Dim Str1 As String
        Str1 = "OGGETTI MSYS;OGGETTO;" & _
               "IMPORTA_OGGETTI_MSYS_DLL;FORM 1;" & _
               "AA_PROVA;FORM 2;" & _
               "LLPP_IMPEGNI_Tb01FrmM01_01_GESTIONE_IMPEGNI_MASTER;FORM 3;" & _
               "4;FORM 4;" & _
               "-;-"

        ' Imposto la casella combinata
        Me.Cmb_01_FORM_TXT.RowSourceType = "Value List"
        Me.Cmb_01_FORM_TXT.RowSource = Str1
        Me.Cmb_01_FORM_TXT.ColumnCount = 2
        Me.Cmb_01_FORM_TXT.BoundColumn = 1

        ' Verifica se il valore predefinito è valido prima di assegnarlo
        If InStr(1, Str1, "*SCELTA DELLA FORM*", vbTextCompare) > 0 Then
            Me.Cmb_01_FORM_TXT.Value = "*SCELTA DELLA FORM*"
        Else
            Me.Cmb_01_FORM_TXT.Value = ""
        End If

        ' Aggiorno i dati della casella combinata
        Me.Cmb_01_FORM_TXT.Requery
    End If

ExitSub:
    Exit Sub

ErrHandler:
    MsgBox "Si è verificato un errore: " & Err.Description, vbExclamation, "Errore " & Err.Number
    
      '//@FILE@LOG@ERRORI
      '//.................................................................//
          Dim NomeRoutine_s As String
          Dim NumeroErrore_lng As Long
          Dim MessaggioErrore_s  As String
          
          NomeRoutine_s = "ROUTINE: Cmb_01_FORM_TXT_GotFocus"
          NumeroErrore_lng = Err.Number
          MessaggioErrore_s = "ERRORE DI ESECUZIONE ROUTINE : " & Err.Description
          
          ScriviLogErrore NomeRoutine_s, _
                          NumeroErrore_lng, _
                          MessaggioErrore_s
      
      '//.................................................................//
      
    Resume ExitSub
End Sub



'//SU MODIFICA DELLA COMBINATA
Private Sub Cmb_01_FORM_TXT_Change()
  

  Select Case Me.Cmb_01_FORM_TXT.Value
      Case "FORM 3"
          Stop
      Case Else
          'MsgBox "SCELTA EFFTTUATA NON ATTIVA  : " & Me.Cmb_01_FORM_TXT.Value, vbInformation, "TITOLO SCELTA FORM"
    End Select

End Sub



'//* @01_@AUMENTA@DIMENSIONI           = AUMENTO LE DIMENSIONE DELLE FORM MASTER E DEGLI OGGETTI INCORPORATI
'//-------------------------------------------------------------------------------------------------------------------//
'//Note: _
      Recupera il nome della form dalla casella combinata Cmb_01_FORM_TXT. _
      Scansiona tutte le form nel database con CurrentProject.AllForms. _
      Se il nome esiste, chiama AumentaDimensioni O RIDUCEDimensioni, altrimenti mostra un messaggio di errore. _
      Questo codice ti assicura che AumentaDimensioni venga chiamata solo se la form esiste nel database.
      




Private Sub Cmd_AUMENTA_DIMENSIONI_FORM_Click()
    On Error GoTo ErroreHandler ' Attiva la gestione degli errori
    
    Dim NomeForm_s As String
    
    ' Recupero il nome della form dalla casella combinata
    NomeForm_s = Nz(Me.Cmb_01_FORM_TXT.Value, "") ' Evita errori se la combo è vuota
    
    ' Controllo se la combo è vuota
    If NomeForm_s = "" Then
        MsgBox "Selezionare una form prima di procedere.", vbExclamation, "Attenzione"
        Exit Sub
    End If
    
    ' Verifica se la form esiste usando la funzione DI CONTROLLO (con trim per togliere gli spazi)
    If FormEsisteNelDatabase(Trim(NomeForm_s)) Then
        Call AumentaDimensioni(NomeForm_s)
        'Apro la form modificata
        DoCmd.OpenForm NomeForm_s
    Else
        MsgBox "La form '" & NomeForm_s & "' non esiste nel database.", vbExclamation, "Errore"
    End If

    Exit Sub ' Evita l'esecuzione del codice di gestione errori se tutto va bene

ErroreHandler:
    MsgBox "Si è verificato un errore in Cmd_AUMENTA_DIMENSIONI_FORM_Click: " & Err.Description, vbCritical, "Errore " & Err.Number
End Sub




'//* @02_@RIDUCI@DIMENSIONI            = RIDUCO LE DIMENSIONE DELLE FORM MASTER E DEGLI OGGETTI INCORPORATI
Private Sub Cmd_RIDUCI_DIMENSIONI_FORM_Click()
    On Error GoTo ErroreHandler ' Attiva la gestione degli errori
    
    Dim NomeForm_s As String
    
    ' Recupero il nome della form dalla casella combinata
    NomeForm_s = Nz(Me.Cmb_01_FORM_TXT.Value, "") ' Evita errori se la combo è vuota
    
    ' Controllo se la combo è vuota
    If NomeForm_s = "" Then
        MsgBox "Selezionare una form prima di procedere.", vbExclamation, "Attenzione"
        Exit Sub
    End If
    
    ' Verifica se la form esiste usando la funzione usando la funzione DI CONTROLLO (con trim per togliere gli spazi)
    If FormEsisteNelDatabase(Trim(NomeForm_s)) Then
        Call RiduciDimensioni(NomeForm_s)
        
        'Apro la form modificata
        DoCmd.OpenForm NomeForm_s
    Else
        MsgBox "La form '" & NomeForm_s & "' non esiste nel database.", vbExclamation, "Errore"
    End If

    Exit Sub ' Evita l'esecuzione del codice di gestione errori se tutto va bene

ErroreHandler:
    MsgBox "Si è verificato un errore in Cmd_RIDUCI_DIMENSIONI_FORM_Click: " & Err.Description, vbCritical, "Errore " & Err.Number
End Sub

'//-------------------------------------------------------------------------------------------------------------------//

'//* @03_@CONTROLLO@ESISTENZA@FORM      =  AUMENTO O RIDUCO LA DIMENSIONE DELLE FORM SOLO SE ESISTONO NEL DB CORRENTE
'//-------------------------------------------------------------------------------------------------------------------//

Public Function FormEsisteNelDatabase(NomeForm_s As String) As Boolean
    Dim obj As Object
    FormEsisteNelDatabase = False ' Default: la form non esiste

    ' Scansiona tutte le form del database
    For Each obj In CurrentProject.AllForms
        If obj.Name = NomeForm_s Then
            FormEsisteNelDatabase = True ' La form esiste
            Exit Function
        End If
    Next obj
End Function
'//-------------------------------------------------------------------------------------------------------------------//



'//@01_@AUMENTA@DIMENSIONI           = AUMENTO LE DIMENSIONE DELLE FORM MASTER E DEGLI OGGETTI INCORPORATI
Private Sub AumentaDimensioni(par_NameForm_s As String)
    Dim frm As Form
    Dim errMsg As String
    
    If par_NameForm_s > "" Then
    
        ' Chiudi la form se è aperta
        'If SysCmd(acSysCmdGetObjectState, acForm, "AA_PROVA") <> 0 Then
        '    DoCmd.Close acForm, "AA_PROVA", acSaveYes
        'End If
        
        If SysCmd(acSysCmdGetObjectState, acForm, par_NameForm_s) <> 0 Then
            DoCmd.Close acForm, par_NameForm_s, acSaveYes
        End If
    
        ' Apri la form in modalità DESIGN
        'DoCmd.OpenForm "AA_PROVA", acDesign
        'Set frm = forms("AA_PROVA")
        
        DoCmd.OpenForm par_NameForm_s, acDesign
        Set frm = forms(par_NameForm_s)
        
        
        
        Dim Larghezza_d As Double
        Dim Altezza_d As Double
       
        '// FORM PRINCIPALE
        '//..........................L  , H
        ' **Ordine corretto per AUMENTARE: FORM ? CORPO ? TAB CONTROL ? SOTTOFORM**
        Larghezza_d = 27
        Call ModificaFormPrincipale(frm, Larghezza_d)      ' Form principale a 36 cm : larghezza
        
        
       
       '// PAGINE DEL TAB CONTROL -                            ATTENZIONE modifiche le pagine trascinano a se gli oggetti sottoform
        '//..........................L  , H
        'Call ModificaPagineTab(frm, 19.551, 22.555)        ' Pagine del Tab Control a 32x23 cm
        Larghezza_d = 26.101
        Altezza_d = 16.383
      
        Call ModificaPagineTab(frm, Larghezza_d, Altezza_d)        ' Pagine del Tab Control a 32x23 cm
        
      
        '// SOTTOFORM
        '//..........................L , H
        Larghezza_d = 25.701
        Altezza_d = 15.376
        Call ModificaSottoform(frm, Larghezza_d, Altezza_d)               ' Sottoform a 15x31 cm    : larghezza, altezza
        
      
        '// TAB CONTROL
        '//..........................L  , H       ' TAB CONTROLL
        Larghezza_d = 25.999
        Altezza_d = 17.333
        Call ModificaTabControl(frm, Larghezza_d, Altezza_d)      ' Tab Control a 32x23 cm  : larghezza, altezza
    
        '// CORPO MASCHERA
        '//..........................L  , H
        Altezza_d = 17.555
        Call ModificaCorpoMaschera(frm, Altezza_d)       ' Corpo maschera a 30 cm  : altezza
      
        
      
        ' Salva e chiudi la form
        'DoCmd.Close acForm, "AA_PROVA", acSaveYes
        DoCmd.Close acForm, par_NameForm_s, acSaveYes
        
        
        Set frm = Nothing
        MsgBox "? Dimensioni aumentate con successo!", vbInformation, "Operazione Completata"
        
      Else
          MsgBox "? NON ESISTE LA FORM DA MODIFICARE!", vbInformation, "Operazione NON ESEGUITA"
          
      End If '//If par_NameForm_s > "" Then
End Sub

' --------------------------------
' ROUTINE PER RIDURRE LE DIMENSIONI
' --------------------------------
'// @02_@RIDUCI@DIMENSIONI            = RIDUCO LE DIMENSIONE DELLE FORM MASTER E DEGLI OGGETTI INCORPORATI
Public Sub RiduciDimensioni(par_NameForm_s As String)
    Dim frm As Form
    Dim errMsg As String
      
      If par_NameForm_s > "" Then
          ' Chiudi la form se è aperta
          'If SysCmd(acSysCmdGetObjectState, acForm, "AA_PROVA") <> 0 Then
          '    DoCmd.Close acForm, "AA_PROVA", acSaveYes
          'End If
          
          If SysCmd(acSysCmdGetObjectState, acForm, par_NameForm_s) <> 0 Then
              DoCmd.Close acForm, par_NameForm_s, acSaveYes
          End If
      
      
          ' Apri la form in modalità DESIGN
          'DoCmd.OpenForm "AA_PROVA", acDesign
          'Set frm = forms("AA_PROVA")
          
          DoCmd.OpenForm par_NameForm_s, acDesign
          Set frm = forms(par_NameForm_s)
      
          ' **Ordine corretto per RIDURRE: SOTTOFORM ? TAB CONTROL ? CORPO ? FORM**
          Call ModificaSottoform(frm, 20.301, 15.899)         ' Sottoform a 13x15 cm
          Call ModificaPagineTab(frm, 20.702, 16.109)         ' Pagine del Tab Control a 17x17 cm
          Call ModificaTabControl(frm, 20.999, 16.943)     ' Tab Control a 17x17 cm
          Call ModificaCorpoMaschera(frm, 17.243)   ' Corpo maschera a 17.101 cm
          Call ModificaFormPrincipale(frm, 18)      ' Form principale a 27 cm
      
          ' Salva e chiudi la form
          DoCmd.Close acForm, "AA_PROVA", acSaveYes
          Set frm = Nothing
          MsgBox "? Dimensioni ridotte con successo!", vbInformation, "Operazione Completata"
          
      Else
          MsgBox "? NON ESISTE LA FORM DA MODIFICARE!", vbInformation, "Operazione NON ESEGUITA"
          
      End If '//If par_NameForm_s > "" Then
    
End Sub

' -------------------------------
' **FUNZIONI DI SUPPORTO**
' -------------------------------

' Modifica la larghezza della Form principale
Private Sub ModificaFormPrincipale(frm As Form, Larghezza As Double)
    frm.Width = CmToTwips(Larghezza)
    'DoCmd.Save
    DoEvents
End Sub

' Modifica l'altezza del Corpo Maschera
Private Sub ModificaCorpoMaschera(frm As Form, Altezza As Double)
    frm.Section(acDetail).Height = CmToTwips(Altezza)
    DoCmd.Save
    DoEvents
End Sub

' Modifica le dimensioni del Tab Control
Private Sub ModificaTabControl(frm As Form, Larghezza As Double, Altezza As Double)
    Dim tabCtl As Control
    On Error Resume Next
    Set tabCtl = frm("TabCtl1")
    On Error GoTo 0
    If Not tabCtl Is Nothing Then
        tabCtl.Width = CmToTwips(Larghezza)
        tabCtl.Height = CmToTwips(Altezza)
        DoCmd.Save
        DoEvents
        DoCmd.Save
    End If
      
    Set tabCtl = Nothing
End Sub

' Modifica le dimensioni di tutte le pagine del Tab Control
Private Sub ModificaPagineTab(frm As Form, Larghezza As Double, Altezza As Double)
    Dim tabCtl As Control
    Dim i As Integer
    
    ' Ottiene il controllo TabCtl1
    On Error Resume Next
    Set tabCtl = frm("TabCtl1")
    On Error GoTo 0
    
    ' Se il Tab Control esiste, modifica le pagine
    If Not tabCtl Is Nothing Then
        For i = 0 To tabCtl.Pages.count - 1
            tabCtl.Pages(i).Width = CmToTwips(Larghezza)
            tabCtl.Pages(i).Height = CmToTwips(Altezza)
            DoEvents
            DoCmd.Save
        Next i
    End If
    
    ' Pulizia memoria
    Set tabCtl = Nothing
End Sub


' Modifica le dimensioni di tutti i sottoform e imposta il bordo con limiti MIN e MAX
Private Sub ModificaSottoform(frm As Form, Larghezza As Double, Altezza As Double)
    Const MIN_LARGHEZZA As Double = 5    ' Larghezza minima in cm
    Const MAX_LARGHEZZA As Double = 30   ' Larghezza massima in cm
    Const MIN_ALTEZZA As Double = 5      ' Altezza minima in cm
    Const MAX_ALTEZZA As Double = 25     ' Altezza massima in cm
    
    Dim i As Integer, ctl As Control
    Dim larghezzaEffettiva As Double, altezzaEffettiva As Double
    
    ' Verifica e limita le dimensioni nei range definiti
    larghezzaEffettiva = IIf(Larghezza < MIN_LARGHEZZA, MIN_LARGHEZZA, IIf(Larghezza > MAX_LARGHEZZA, MAX_LARGHEZZA, Larghezza))
    altezzaEffettiva = IIf(Altezza < MIN_ALTEZZA, MIN_ALTEZZA, IIf(Altezza > MAX_ALTEZZA, MAX_ALTEZZA, Altezza))
    
    For i = 1 To 12
        On Error Resume Next
        Set ctl = frm("sottoform_" & Format(i, "00"))
        On Error GoTo 0
        If Not ctl Is Nothing Then
            ' Modifica dimensioni con i valori limitati
            ctl.Width = CmToTwips(larghezzaEffettiva)
            ctl.Height = CmToTwips(altezzaEffettiva)
            Debug.Print "Modificato: "; ctl.Name; " - Larghezza: "; larghezzaEffettiva; " cm, Altezza: "; altezzaEffettiva; " cm"
            
            ' Imposta l'aspetto piatto, il colore e lo spessore del bordo
            ctl.BorderStyle = 1 ' Stile continuo (fmBorderStyleSingle)
            ctl.SpecialEffect = 0 ' Aspetto piatto (fmSpecialEffectFlat)
            ctl.BorderColor = RGB(&HBA, &H14, &H19) ' Colore bordo: #BA1419
            ctl.BorderWidth = 2 ' Spessore bordo: 2 pt
            
            DoCmd.Save
            DoEvents
        End If
        Set ctl = Nothing
    Next i
End Sub










'//CALCOLO TWIPS CON IL CONTROLLO DEI LIMITI MAX E MIN
Function CmToTwips(cm As Variant) As Long
    Const TWIPS_PER_CM As Double = 567 ' 1 cm = 567 twips
    Const MAX_TWIPS As Long = 32767 ' Limite massimo per larghezza/altezza degli oggetti
    Const MIN_TWIPS As Long = 1 ' Limite minimo per evitare valori negativi o nulli
    
    Dim result As Long
    
    On Error GoTo ErrHandler ' Abilita la gestione degli errori

    ' Controllo se il valore è numerico
    If IsNull(cm) Or Not IsNumeric(cm) Then
        MsgBox "Errore: Il valore inserito non è numerico o è nullo.", vbCritical, "Errore di conversione"
        Exit Function
    End If

    ' Controllo valori negativi
    If cm < 0 Then
        MsgBox "Errore: Il valore in cm non può essere negativo.", vbCritical, "Errore di conversione"
        Exit Function
    End If

    ' Conversione
    result = cm * TWIPS_PER_CM

    ' Applicare i limiti con messaggi di avviso
    If result > MAX_TWIPS Then
        MsgBox "Attenzione: Il valore convertito supera il limite massimo e verrà impostato a " & MAX_TWIPS & " twips.", vbInformation, "Limite superato"
        result = MAX_TWIPS
    ElseIf result < MIN_TWIPS Then
        MsgBox "Attenzione: Il valore convertito è troppo basso e verrà impostato a " & MIN_TWIPS & " twips.", vbInformation, "Limite minimo"
        result = MIN_TWIPS
    End If
    
    CmToTwips = result
    
ExitFunction:
    Exit Function

ErrHandler:
    MsgBox "Si è verificato un errore: " & Err.Description, vbExclamation, "Errore " & Err.Number

    ' //@FILE@LOG@ERRORI
    ' //.................................................................//
    Dim NomeRoutine_s As String
    Dim NumeroErrore_lng As Long
    Dim MessaggioErrore_s As String

    NomeRoutine_s = "ROUTINE: CmToTwips"
    NumeroErrore_lng = Err.Number
    MessaggioErrore_s = "ERRORE DI ESECUZIONE ROUTINE : " & Err.Description

    ScriviLogErrore NomeRoutine_s, _
                    NumeroErrore_lng, _
                    MessaggioErrore_s

    ' //.................................................................//

    Resume ExitFunction
End Function




'//***********************************************************************************//
'//**           COMBINATA 02 @CMB_02 = ALTEZZA LARGHEZZA FORM   *** INIZIO ***
'//**
'//***********************************************************************************//
'//* @04_@CMB@02@FORM@LARGHEZZA@ALTEZZA   = @Cmb_02_FORM_L_H_TXT @LARGHEZZA E @ALTEZZA DELLA @FORM @CMB@02_(combinata @misure @L@H della form)

'//EVENTO SU ATTIVATO IMPOSTO IL VALORE H+L DELLA FORM
Private Sub Cmb_02_FORM_L_H_TXT_GotFocus()
    On Error GoTo ErrHandler ' Attiva la gestione degli errori

    ' Verifico se la casella combinata esiste
    If Not Me.Cmb_02_FORM_L_H_TXT Is Nothing Then
        ' Creazione della lista di valori incrementali
        Dim Str1 As String
        Dim i As Integer
        Dim Larghezza As Integer
        Dim Altezza As Integer

        ' Aggiunge il valore predefinito come prima voce della lista
        Str1 = "LxH Larghezza form;;;" ' Testo visibile + colonne vuote

        ' Imposta i valori iniziali
        Larghezza = 10
        Altezza = 15

        ' Genera 10 valori incrementando di 50 twips
        For i = 1 To 10
            Str1 = Str1 & "L" & Larghezza & " H" & Altezza & ";" & Larghezza & ";" & Altezza & ";"
            Larghezza = Larghezza + 5
            Altezza = Altezza + 5
        Next i

        ' Imposto la casella combinata
        Me.Cmb_02_FORM_L_H_TXT.RowSourceType = "Value List"
        Me.Cmb_02_FORM_L_H_TXT.RowSource = Str1
        Me.Cmb_02_FORM_L_H_TXT.ColumnCount = 3 ' Tre colonne: Testo visibile, Larghezza, Altezza
        Me.Cmb_02_FORM_L_H_TXT.ColumnWidths = "3cm;0cm;0cm" ' Nasconde le colonne 2 e 3
        Me.Cmb_02_FORM_L_H_TXT.BoundColumn = 1 ' La colonna legata è la prima

        ' **Forza il valore predefinito dopo aver impostato la RowSource**
        Me.Cmb_02_FORM_L_H_TXT.Value = "LxH Larghezza form"

        ' Aggiorno i dati della casella combinata
        Me.Cmb_02_FORM_L_H_TXT.Requery
    End If
    
        

ExitSub:
    Exit Sub

ErrHandler:
    MsgBox "Si è verificato un errore: " & Err.Description, vbExclamation, "Errore " & Err.Number
    
    ' //@FILE@LOG@ERRORI
    ' //.................................................................//
    Dim NomeRoutine_s As String
    Dim NumeroErrore_lng As Long
    Dim MessaggioErrore_s As String

    NomeRoutine_s = "ROUTINE: Cmb_02_FORM_L_H_TXT_GotFocus"
    NumeroErrore_lng = Err.Number
    MessaggioErrore_s = "ERRORE DI ESECUZIONE ROUTINE : " & Err.Description

    ScriviLogErrore NomeRoutine_s, _
                    NumeroErrore_lng, _
                    MessaggioErrore_s

    ' //.................................................................//
    
    Resume ExitSub
End Sub

'==========================================
' Evento AfterUpdate LARGHEZZA + ALTEZZA DELLA FORM
'==========================================
'//Note: Dopo aver recuperato i valori di H + L dall'evento su ATTIVATO con questo _
        evento su modifica controllo se la form esiste (cmb_01) e poi attivo le _
        routine di attivazione larghezza form ed altezza pagina form.
        
Private Sub Cmb_02_FORM_L_H_TXT_AfterUpdate()
    On Error GoTo ErrHandler ' Attiva la gestione degli errori

    Dim ValoreSelezionato As String
    Dim Parti() As String
    Dim var1_L_dbl As Double
    Dim var2_H_dbl As Double
    Dim frm As Form                         '//oggetto della form fa modificare

    ' Verifica se la casella combinata contiene un valore valido
    If IsNull(Me.Cmb_02_FORM_L_H_TXT.Value) Or Me.Cmb_02_FORM_L_H_TXT.Value = "" Then
        Exit Sub
    End If

    ' Ottiene il valore selezionato dalla casella combinata (es. "L150 H145")
    ValoreSelezionato = Me.Cmb_02_FORM_L_H_TXT.Value

    ' Divide il valore in due parti utilizzando lo spazio come separatore
    Parti = Split(ValoreSelezionato, " ") ' Ora Parti(0) = "L150", Parti(1) = "H145"

    ' Verifica che ci siano almeno due parti valide
    If UBound(Parti) < 1 Then
        Err.Raise 1001, , "Formato valore non valido: " & ValoreSelezionato
    End If

    ' Rimuove la lettera "L" e converte in numero
    var1_L_dbl = CDbl(CInt(CVar(Mid(Parti(0), 2))))  ' Prende tutto tranne la prima lettera "L"

    ' Rimuove la lettera "H" e converte in numero
    var2_H_dbl = CDbl(CInt(CVar(Mid(Parti(1), 2)))) ' Prende tutto tranne la prima lettera "H"

    ' Ora var1_L_dbl  contiene la larghezza (es. 150) e var2_H_dbl l'altezza (es. 145)
    MsgBox "Larghezza: " & var1_L_dbl & " - Altezza: " & var2_H_dbl, vbInformation, "Valori Selezionati"

    ' Qui puoi chiamare altre routine passando var1_L_dbl  e var2_H_dbl
    ' Call ImpostaDimensioni(var1_L_dbl , var2_H_dbl)
    '//controllo prima se esiste la form da modificare
        If Cmb_01_FORM_TXT > "" Then
                
                Dim NomeForm_s As String
                ' Recupero il nome della form dalla casella combinata
                NomeForm_s = Nz(Me.Cmb_01_FORM_TXT.Value, "") ' Evita errori se la combo è vuota
                
        
            '//chiamo la procedura di controllo esistenza form se esiste TRUE
            Bool1 = False
            Bool1 = FormEsisteNelDatabase(NomeForm_s)
            
            If (Bool1) = True Then
            
                    ' Chiudi la form se è aperta
                    If SysCmd(acSysCmdGetObjectState, acForm, NomeForm_s) <> 0 Then
                        DoCmd.Close acForm, Trim(Me.Cmb_01_FORM_TXT), acSaveYes
                    End If
                
                    ' Apri la form in modalità DESIGN
                    DoCmd.OpenForm NomeForm_s, acDesign
                    Set frm = forms(NomeForm_s)
                    
               
            
                '// FORM PRINCIPALE L = LARGHEZZA H = CORPO MASCHERA
                '// Note: attivo le 2 routine che modificano la larghezza della form _
                          e l 'altezza della pagina della form
                '//..........................L  , H
                Call ModificaFormPrincipale(frm, var1_L_dbl)        ' Form principale : larghezza
                
                
                  '// CORPO MASCHERA
                  '//..........................L  , H
                  Call ModificaCorpoMaschera(frm, var2_H_dbl)       ' Corpo maschera : altezza
                
                        '//PULISCO LE RISORSE
                        Set frm = Nothing
                       
            
            End If
            
            
        
        End If

ExitSub:
    Exit Sub

ErrHandler:
    MsgBox "Si è verificato un errore: " & Err.Description, vbExclamation, "Errore " & Err.Number
    
    ' //@FILE@LOG@ERRORI
    ' //.................................................................//
    Dim NomeRoutine_s As String
    Dim NumeroErrore_lng As Long
    Dim MessaggioErrore_s  As String

    NomeRoutine_s = "ROUTINE: Cmb_02_FORM_L_H_TXT_AfterUpdate"
    NumeroErrore_lng = Err.Number
    MessaggioErrore_s = "ERRORE DI ESECUZIONE ROUTINE : " & Err.Description

    ScriviLogErrore NomeRoutine_s, _
                    NumeroErrore_lng, _
                    MessaggioErrore_s

    ' //.................................................................//
    
    Resume ExitSub
End Sub


'//***********************************************************************************//
'//**           COMBINATA 02 @CMB_02 = ALTEZZA LARGHEZZA FORM   *** FINE ***
'//**
'//***********************************************************************************//



'//***********************************************************************************//
'//**           COMBINATA 03 @CMB_03 = ALTEZZA LARGHEZZA @PAGINE@TAB@CONTROL   *** INIZIO ***
'//**
'//***********************************************************************************//
'//* @04_@CMB@02@FORM@LARGHEZZA@ALTEZZA   = @CMB_03_PAGINE_TAB_CONTROL_L_H_TXT @LARGHEZZA E @ALTEZZA DELLE @PAGINE@TAB@CONTROL_(combinata @misure @L@H PAGINE DEL TAB CONTROL)

'//EVENTO SU ATTIVATO IMPOSTO IL VALORE H+L DELLE PAGINE DEL TAB CONTROL
Private Sub CMB_03_PAGINE_TAB_CONTROL_L_H_TXT_GotFocus()
      On Error GoTo ErrHandler ' Attiva la gestione degli errori

    ' Verifico se la casella combinata esiste
    If Not Me.CMB_03_PAGINE_TAB_CONTROL_L_H_TXT Is Nothing Then
        ' Creazione della lista di valori incrementali
        Dim Str1 As String
        Dim i As Integer
        Dim Larghezza_dbl As Double
        Dim Altezza_dbl As Double

        ' Aggiunge il valore predefinito come prima voce della lista
        Str1 = "LxH PAGINE TAB CONTROL;;;" ' Testo visibile + colonne vuote
        
        
        ' Imposta i valori iniziali @dimensione@pagine@tab@control_(qui impostiamo la larghezza e l'altezza @delle@pagine del tab controll)
        Larghezza_dbl = 24.401
        Altezza_dbl = 16.501

        ' Genera 10 valori incrementando di 50 twips
        For i = 1 To 10
            Str1 = Str1 & "L" & Larghezza_dbl & " H" & Altezza_dbl & ";" & Larghezza_dbl & ";" & Altezza_dbl & ";"
            
            Larghezza_dbl = Larghezza_dbl + 1
            Altezza_dbl = Altezza_dbl + 1
        Next i

        ' Imposto la casella combinata
        Me.CMB_03_PAGINE_TAB_CONTROL_L_H_TXT.RowSourceType = "Value List"
        Me.CMB_03_PAGINE_TAB_CONTROL_L_H_TXT.RowSource = Str1
        Me.CMB_03_PAGINE_TAB_CONTROL_L_H_TXT.ColumnCount = 3 ' Tre colonne: Testo visibile, Larghezza, Altezza
        Me.CMB_03_PAGINE_TAB_CONTROL_L_H_TXT.ColumnWidths = "3cm;0cm;0cm" ' Nasconde le colonne 2 e 3
        Me.CMB_03_PAGINE_TAB_CONTROL_L_H_TXT.BoundColumn = 1 ' La colonna legata è la prima

        ' **Forza il valore predefinito dopo aver impostato la RowSource**
        Me.CMB_03_PAGINE_TAB_CONTROL_L_H_TXT.Value = "LxH Larghezza+Altezza PAGINE TAB CONTROLL"

        ' Aggiorno i dati della casella combinata
        Me.CMB_03_PAGINE_TAB_CONTROL_L_H_TXT.Requery
    End If
    
        

ExitSub:
    Exit Sub

ErrHandler:
    MsgBox "Si è verificato un errore: " & Err.Description, vbExclamation, "Errore " & Err.Number
    
    ' //@FILE@LOG@ERRORI
    ' //.................................................................//
    Dim NomeRoutine_s As String
    Dim NumeroErrore_lng As Long
    Dim MessaggioErrore_s As String

    NomeRoutine_s = "ROUTINE: CMB_03_PAGINE_TAB_CONTROL_L_H_TXT_GotFocus"
    NumeroErrore_lng = Err.Number
    MessaggioErrore_s = "ERRORE DI ESECUZIONE ROUTINE : " & Err.Description

    ScriviLogErrore NomeRoutine_s, _
                    NumeroErrore_lng, _
                    MessaggioErrore_s

    ' //.................................................................//
    
    Resume ExitSub
End Sub

'==========================================
' Evento AfterUpdate LARGHEZZA + ALTEZZA DELLE PAGINE DEL TAB CONTROL
'==========================================
'//Note: Dopo aver recuperato i valori di H + L dall'evento su ATTIVATO con questo _
        evento su modifica controllo se la form esiste (cmb_01) e poi attivo le _
        routine di attivazione larghezza DELLE PAGINE DEL TAB CONTROLL  della  form principale
        
Private Sub CMB_03_PAGINE_TAB_CONTROL_L_H_TXT_AfterUpdate()
    On Error GoTo ErrHandler ' Attiva la gestione degli errori

    Dim ValoreSelezionato As String
    Dim Parti() As String
    Dim var1_L_dbl As Double
    Dim var2_H_dbl As Double
    Dim frm As Form                         '//oggetto della form fa modificare

    ' Verifica se la casella combinata contiene un valore valido
    If IsNull(Me.CMB_03_PAGINE_TAB_CONTROL_L_H_TXT.Value) Or Me.CMB_03_PAGINE_TAB_CONTROL_L_H_TXT.Value = "" Then
        Exit Sub
    End If

    ' Ottiene il valore selezionato dalla casella combinata (es. "L150 H145")
    ValoreSelezionato = Me.CMB_03_PAGINE_TAB_CONTROL_L_H_TXT.Value

    ' Divide il valore in due parti utilizzando lo spazio come separatore
    Parti = Split(ValoreSelezionato, " ") ' Ora Parti(0) = "L150", Parti(1) = "H145"

    ' Verifica che ci siano almeno due parti valide
    If UBound(Parti) < 1 Then
        Err.Raise 1001, , "Formato valore non valido: " & ValoreSelezionato
    End If

    ' Rimuove la lettera "L" e converte in numero
    var1_L_dbl = CDbl(CVar(Mid(Parti(0), 2)))  ' Prende tutto tranne la prima lettera "L"

    ' Rimuove la lettera "H" e converte in numero
    var2_H_dbl = CDbl(CVar(Mid(Parti(1), 2))) ' Prende tutto tranne la prima lettera "H"

    ' Ora var1_L_dbl  contiene la larghezza (es. 150) e var2_H_dbl l'altezza (es. 145)
    MsgBox "DIMENSIONI DELLE PAGINE DEL TAB CONTROLL: Larghezza: " & var1_L_dbl & " - Altezza: " & var2_H_dbl, vbInformation, "Valori Selezionati"

    ' Qui puoi chiamare altre routine passando var1_L_dbl  e var2_H_dbl
    ' Call ImpostaDimensioni(var1_L_dbl , var2_H_dbl)
    '//controllo prima se esiste la form da modificare
        If Cmb_01_FORM_TXT > "" Then
                
                Dim NomeForm_s As String
                ' Recupero il nome della form dalla casella combinata
                NomeForm_s = Nz(Me.Cmb_01_FORM_TXT.Value, "") ' Evita errori se la combo è vuota
                
        
            '//chiamo la procedura di controllo esistenza form se esiste TRUE
            Bool1 = False
            Bool1 = FormEsisteNelDatabase(NomeForm_s)
            
            If (Bool1) = True Then
            
                    ' Chiudi la form se è aperta
                    If SysCmd(acSysCmdGetObjectState, acForm, NomeForm_s) <> 0 Then
                        DoCmd.Close acForm, Trim(Me.Cmb_01_FORM_TXT), acSaveYes
                    End If
                
                    ' Apri la form in modalità DESIGN
                    DoCmd.OpenForm NomeForm_s, acDesign
                    Set frm = forms(NomeForm_s)
                    
               
            
                '// TAB CONTROLL DELLA FORM PRINCIPALE L = LARGHEZZA H = ALTEZZA
                '// Note: attivo la routine che modificano la larghezza e l'altezza  delle pagine del tab controll
                        
                       
            '// PAGINE DEL TAB CONTROL -                            ATTENZIONE modifiche le pagine trascinano a se gli oggetti sottoform
            '//..........................L  , H
            'Call ModificaPagineTab(frm, 19.551, 22.555)           Pagine del Tab Control a 32x23 cm
            
            Call ModificaPagineTab(frm, var1_L_dbl, var2_H_dbl)      ' Pagine del Tab Control a 32x23 cm
            
            End If
            
            
        
        End If

ExitSub:
    Exit Sub

ErrHandler:
    MsgBox "Si è verificato un errore: " & Err.Description, vbExclamation, "Errore " & Err.Number
    
    ' //@FILE@LOG@ERRORI
    ' //.................................................................//
    Dim NomeRoutine_s As String
    Dim NumeroErrore_lng As Long
    Dim MessaggioErrore_s  As String

    NomeRoutine_s = "ROUTINE: CMB_03_PAGINE_TAB_CONTROL_L_H_TXT_AfterUpdate"
    NumeroErrore_lng = Err.Number
    MessaggioErrore_s = "ERRORE DI ESECUZIONE ROUTINE : " & Err.Description

    ScriviLogErrore NomeRoutine_s, _
                    NumeroErrore_lng, _
                    MessaggioErrore_s

    ' //.................................................................//
    
    Resume ExitSub
End Sub


'//***********************************************************************************//
'//**           COMBINATA 03 @CMB_03 = ALTEZZA LARGHEZZA PAGINE TAB CONTROL   *** FINE ***
'//**
'//***********************************************************************************//






'//***********************************************************************************//
'//**           COMBINATA 04 @CMB_04 = ALTEZZA LARGHEZZA DEL@TAB@CONTROL   *** INIZIO ***
'//**
'//***********************************************************************************//
'//* @04_@CMB@02@TAB@CONTROL@LARGHEZZA@ALTEZZA   = @CMB_04_TAB_CONTROL_L_H_TXT @LARGHEZZA E @ALTEZZA DEL @TAB@CONTROL_(combinata @misure @L@H DEL TAB CONTROL)

'//EVENTO SU ATTIVATO IMPOSTO IL VALORE H+L DEL TAB CONTROL
Private Sub CMB_04_TAB_CONTROL_L_H_TXT_GotFocus()
    On Error GoTo ErrHandler ' Attiva la gestione degli errori

    ' Verifico se la casella combinata esiste
    If Not Me.CMB_04_TAB_CONTROL_L_H_TXT Is Nothing Then
        ' Creazione della lista di valori incrementali
        Dim Str1 As String
        Dim i As Integer
        Dim Larghezza_dbl As Double
        Dim Altezza_dbl As Double

        ' Aggiunge il valore predefinito come prima voce della lista
        Str1 = "LxH TAB CONTROL;;;" ' Testo visibile + colonne vuote
        
        
        ' Imposta i valori iniziali @dimensione@pagine@tab@control_(qui impostiamo la larghezza e l'altezza @delle@pagine del tab controll)
        Larghezza_dbl = 24.4783
        Altezza_dbl = 176.333

        ' Genera 10 valori incrementando di 50 twips
        For i = 1 To 10
            Str1 = Str1 & "L" & Larghezza_dbl & " H" & Altezza_dbl & ";" & Larghezza_dbl & ";" & Altezza_dbl & ";"
            
            Larghezza_dbl = Larghezza_dbl + 0.75
            Altezza_dbl = Altezza_dbl + 0.75
        Next i

        ' Imposto la casella combinata
        Me.CMB_04_TAB_CONTROL_L_H_TXT.RowSourceType = "Value List"
        Me.CMB_04_TAB_CONTROL_L_H_TXT.RowSource = Str1
        Me.CMB_04_TAB_CONTROL_L_H_TXT.ColumnCount = 3 ' Tre colonne: Testo visibile, Larghezza, Altezza
        Me.CMB_04_TAB_CONTROL_L_H_TXT.ColumnWidths = "3cm;0cm;0cm" ' Nasconde le colonne 2 e 3
        Me.CMB_04_TAB_CONTROL_L_H_TXT.BoundColumn = 1 ' La colonna legata è la prima

        ' **Forza il valore predefinito dopo aver impostato la RowSource**
        Me.CMB_04_TAB_CONTROL_L_H_TXT.Value = "LxH Larghezza+Altezza TAB CONTROL"

        ' Aggiorno i dati della casella combinata
        Me.CMB_04_TAB_CONTROL_L_H_TXT.Requery
    End If
    
        

ExitSub:
    Exit Sub

ErrHandler:
    MsgBox "Si è verificato un errore: " & Err.Description, vbExclamation, "Errore " & Err.Number
    
    ' //@FILE@LOG@ERRORI
    ' //.................................................................//
    Dim NomeRoutine_s As String
    Dim NumeroErrore_lng As Long
    Dim MessaggioErrore_s As String

    NomeRoutine_s = "ROUTINE: CMB_04_TAB_CONTROL_L_H_TXT_GotFocus"
    NumeroErrore_lng = Err.Number
    MessaggioErrore_s = "ERRORE DI ESECUZIONE ROUTINE : " & Err.Description

    ScriviLogErrore NomeRoutine_s, _
                    NumeroErrore_lng, _
                    MessaggioErrore_s

    ' //.................................................................//
    
    Resume ExitSub
End Sub

'==========================================
' Evento AfterUpdate LARGHEZZA + ALTEZZA DEL TAB CONTROL
'==========================================
'//Note: Dopo aver recuperato i valori di H + L dall'evento su ATTIVATO con questo _
        evento su modifica controllo se la form esiste (cmb_01) e poi attivo le _
        routine di attivazione larghezza DEL TAB CONTROL  della  form principale
        
        
Private Sub CMB_04_TAB_CONTROL_L_H_TXT_AfterUpdate()
    On Error GoTo ErrHandler ' Attiva la gestione degli errori

    Dim ValoreSelezionato As String
    Dim Parti() As String
    Dim var1_L_dbl As Double
    Dim var2_H_dbl As Double
    Dim frm As Form                         '//oggetto della form fa modificare

    ' Verifica se la casella combinata contiene un valore valido
    If IsNull(Me.CMB_04_TAB_CONTROL_L_H_TXT.Value) Or Me.CMB_04_TAB_CONTROL_L_H_TXT.Value = "" Then
        Exit Sub
    End If

    ' Ottiene il valore selezionato dalla casella combinata (es. "L150 H145")
    ValoreSelezionato = Me.CMB_04_TAB_CONTROL_L_H_TXT.Value

    ' Divide il valore in due parti utilizzando lo spazio come separatore
    Parti = Split(ValoreSelezionato, " ") ' Ora Parti(0) = "L150", Parti(1) = "H145"

    ' Verifica che ci siano almeno due parti valide
    If UBound(Parti) < 1 Then
        Err.Raise 1001, , "Formato valore non valido: " & ValoreSelezionato
    End If

    ' Rimuove la lettera "L" e converte in numero
    var1_L_dbl = CDbl(CVar(Mid(Parti(0), 2)))  ' Prende tutto tranne la prima lettera "L"

    ' Rimuove la lettera "H" e converte in numero
    var2_H_dbl = CDbl(CVar(Mid(Parti(1), 2))) ' Prende tutto tranne la prima lettera "H"

    ' Ora var1_L_dbl  contiene la larghezza (es. 150) e var2_H_dbl l'altezza (es. 145)
    MsgBox "DIMENSIONI DEL TAB CONTROL: Larghezza: " & var1_L_dbl & " - Altezza: " & var2_H_dbl, vbInformation, "Valori Selezionati"

    ' Qui puoi chiamare altre routine passando var1_L_dbl  e var2_H_dbl
    ' Call ImpostaDimensioni(var1_L_dbl , var2_H_dbl)
    '//controllo prima se esiste la form da modificare
        If Cmb_01_FORM_TXT > "" Then
                
                Dim NomeForm_s As String
                ' Recupero il nome della form dalla casella combinata
                NomeForm_s = Nz(Me.Cmb_01_FORM_TXT.Value, "") ' Evita errori se la combo è vuota
                
        
            '//chiamo la procedura di controllo esistenza form se esiste TRUE
            Bool1 = False
            Bool1 = FormEsisteNelDatabase(NomeForm_s)
            
            If (Bool1) = True Then
            
                    ' Chiudi la form se è aperta
                    If SysCmd(acSysCmdGetObjectState, acForm, NomeForm_s) <> 0 Then
                        DoCmd.Close acForm, Trim(Me.Cmb_01_FORM_TXT), acSaveYes
                    End If
                
                    ' Apri la form in modalità DESIGN
                    DoCmd.OpenForm NomeForm_s, acDesign
                    Set frm = forms(NomeForm_s)
                    
               
            
                '// TAB CONTROL DELLA FORM PRINCIPALE L = LARGHEZZA H = ALTEZZA
                '// Note: attivo la routine che modificano la larghezza e l'altezza  delle pagine del tab controll
                        
                       
            '// DEL TAB CONTROL -                            ATTENZIONE modifiche le pagine trascinano a se gli oggetti sottoform
            '//..........................L  , H
              Call ModificaTabControl(frm, var1_L_dbl, var2_H_dbl)      ' Tab Control DIMENZIONI
            
            End If
            
            
        
        End If

ExitSub:
    Exit Sub

ErrHandler:
    MsgBox "Si è verificato un errore: " & Err.Description, vbExclamation, "Errore " & Err.Number
    
    ' //@FILE@LOG@ERRORI
    ' //.................................................................//
    Dim NomeRoutine_s As String
    Dim NumeroErrore_lng As Long
    Dim MessaggioErrore_s  As String

    NomeRoutine_s = "ROUTINE: CMB_04_TAB_CONTROL_L_H_TXT_AfterUpdate"
    NumeroErrore_lng = Err.Number
    MessaggioErrore_s = "ERRORE DI ESECUZIONE ROUTINE : " & Err.Description

    ScriviLogErrore NomeRoutine_s, _
                    NumeroErrore_lng, _
                    MessaggioErrore_s

    ' //.................................................................//
    
    Resume ExitSub
End Sub


'//***********************************************************************************//
'//**           COMBINATA 04 @CMB_04 = ALTEZZA E LARGHEZZA DEL TAB CONTROL   *** FINE ***
'//**
'//***********************************************************************************//









'//***********************************************************************************//
'//**           COMBINATA 05 @CMB_05 = ALTEZZA E LARGHEZZA DELLE SOTTOFORM   *** FINE ***
'//**
'//***********************************************************************************//


'//EVENTO SU ATTIVATO IMPOSTO IL VALORE H+L DELLE SOTTOFORM
Private Sub CMB_05_SottoForm_L_H_TXT_GotFocus()
    On Error GoTo ErrHandler ' Attiva la gestione degli errori

    ' Verifico se la casella combinata esiste
    If Not Me.CMB_05_SottoForm_L_H_TXT Is Nothing Then
        ' Creazione della lista di valori incrementali
        Dim Str1 As String
        Dim i As Integer
        Dim Larghezza_dbl As Double
        Dim Altezza_dbl As Double

        ' Aggiunge il valore predefinito come prima voce della lista
        Str1 = "LxH SottoForm;;;" ' Testo visibile + colonne vuote
        
        
        ' Imposta i valori iniziali @dimensione@pagine@tab@control_(qui impostiamo la larghezza e l'altezza @delle@SottoForm)
        Larghezza_dbl = 24.4783
        Altezza_dbl = 15.333

        ' Genera 10 valori incrementando di 50 twips
        For i = 1 To 10
            Str1 = Str1 & "L" & Larghezza_dbl & " H" & Altezza_dbl & ";" & Larghezza_dbl & ";" & Altezza_dbl & ";"
            
            Larghezza_dbl = Larghezza_dbl + 0.75
            Altezza_dbl = Altezza_dbl + 0.75
        Next i

        ' Imposto la casella combinata
        Me.CMB_05_SottoForm_L_H_TXT.RowSourceType = "Value List"
        Me.CMB_05_SottoForm_L_H_TXT.RowSource = Str1
        Me.CMB_05_SottoForm_L_H_TXT.ColumnCount = 3 ' Tre colonne: Testo visibile, Larghezza, Altezza
        Me.CMB_05_SottoForm_L_H_TXT.ColumnWidths = "3cm;0cm;0cm" ' Nasconde le colonne 2 e 3
        Me.CMB_05_SottoForm_L_H_TXT.BoundColumn = 1 ' La colonna legata è la prima

        ' **Forza il valore predefinito dopo aver impostato la RowSource**
        Me.CMB_05_SottoForm_L_H_TXT.Value = "LxH Larghezza+Altezza SottoForm"

        ' Aggiorno i dati della casella combinata
        Me.CMB_05_SottoForm_L_H_TXT.Requery
    End If
    
        

ExitSub:
    Exit Sub

ErrHandler:
    MsgBox "Si è verificato un errore: " & Err.Description, vbExclamation, "Errore " & Err.Number
    
    ' //@FILE@LOG@ERRORI
    ' //.................................................................//
    Dim NomeRoutine_s As String
    Dim NumeroErrore_lng As Long
    Dim MessaggioErrore_s As String

    NomeRoutine_s = "ROUTINE: CMB_05_SottoForm_L_H_TXT_GotFocus"
    NumeroErrore_lng = Err.Number
    MessaggioErrore_s = "ERRORE DI ESECUZIONE ROUTINE : " & Err.Description

    ScriviLogErrore NomeRoutine_s, _
                    NumeroErrore_lng, _
                    MessaggioErrore_s

    ' //.................................................................//
    
    Resume ExitSub
End Sub

'==========================================
' Evento AfterUpdate LARGHEZZA + ALTEZZA DELLE SOTTOFORM
'==========================================
'//Note: Dopo aver recuperato i valori di H + L dall'evento su ATTIVATO con questo _
        evento su modifica controllo se la form esiste (cmb_01) e poi attivo le _
        routine di attivazione larghezza DELLE SOTTOFORM  della  form principale
        
        
Private Sub CMB_05_SottoForm_L_H_TXT_AfterUpdate()
    On Error GoTo ErrHandler ' Attiva la gestione degli errori

    Dim ValoreSelezionato As String
    Dim Parti() As String
    Dim var1_L_dbl As Double
    Dim var2_H_dbl As Double
    Dim frm As Form                         '//oggetto della form fa modificare

    ' Verifica se la casella combinata contiene un valore valido
    If IsNull(Me.CMB_05_SottoForm_L_H_TXT.Value) Or Me.CMB_05_SottoForm_L_H_TXT.Value = "" Then
        Exit Sub
    End If

    ' Ottiene il valore selezionato dalla casella combinata (es. "L150 H145")
    ValoreSelezionato = Me.CMB_05_SottoForm_L_H_TXT.Value

    ' Divide il valore in due parti utilizzando lo spazio come separatore
    Parti = Split(ValoreSelezionato, " ") ' Ora Parti(0) = "L150", Parti(1) = "H145"

    ' Verifica che ci siano almeno due parti valide
    If UBound(Parti) < 1 Then
        Err.Raise 1001, , "Formato valore non valido: " & ValoreSelezionato
    End If

    ' Rimuove la lettera "L" e converte in numero
    var1_L_dbl = CDbl(CVar(Mid(Parti(0), 2)))  ' Prende tutto tranne la prima lettera "L"

    ' Rimuove la lettera "H" e converte in numero
    var2_H_dbl = CDbl(CVar(Mid(Parti(1), 2))) ' Prende tutto tranne la prima lettera "H"

    ' Ora var1_L_dbl  contiene la larghezza (es. 150) e var2_H_dbl l'altezza (es. 145)
    MsgBox "DIMENSIONI DELLE SOTTOFORM: Larghezza: " & var1_L_dbl & " - Altezza: " & var2_H_dbl, vbInformation, "Valori Selezionati"

    ' Qui puoi chiamare altre routine passando var1_L_dbl  e var2_H_dbl
    ' Call ImpostaDimensioni(var1_L_dbl , var2_H_dbl)
    '//controllo prima se esiste la form da modificare
        If Cmb_01_FORM_TXT > "" Then
                
                Dim NomeForm_s As String
                ' Recupero il nome della form dalla casella combinata
                NomeForm_s = Nz(Me.Cmb_01_FORM_TXT.Value, "") ' Evita errori se la combo è vuota
                
        
            '//chiamo la procedura di controllo esistenza form se esiste TRUE
            Bool1 = False
            Bool1 = FormEsisteNelDatabase(NomeForm_s)
            
            If (Bool1) = True Then
            
                    ' Chiudi la form se è aperta
                    If SysCmd(acSysCmdGetObjectState, acForm, NomeForm_s) <> 0 Then
                        DoCmd.Close acForm, Trim(Me.Cmb_01_FORM_TXT), acSaveYes
                    End If
                
                    ' Apri la form in modalità DESIGN
                    DoCmd.OpenForm NomeForm_s, acDesign
                    Set frm = forms(NomeForm_s)
                    
               
            
                '// SottoForm DELLA FORM PRINCIPALE L = LARGHEZZA H = ALTEZZA
                '// Note: attivo la routine che modificano la larghezza e l'altezza  delle SottoForm
                        
                       
            '// DELLE SOTTOFORM - ATTENZIONE modifiche sugli oggetti SottoForm
            '//..........................L  , H
              Call ModificaSottoform(frm, var1_L_dbl, var2_H_dbl)      ' SottoForm DIMENZIONI
            
            End If
            
            
        
        End If

ExitSub:
    Exit Sub

ErrHandler:
    MsgBox "Si è verificato un errore: " & Err.Description, vbExclamation, "Errore " & Err.Number
    
    ' //@FILE@LOG@ERRORI
    ' //.................................................................//
    Dim NomeRoutine_s As String
    Dim NumeroErrore_lng As Long
    Dim MessaggioErrore_s  As String

    NomeRoutine_s = "ROUTINE: CMB_05_SottoForm_L_H_TXT_AfterUpdate"
    NumeroErrore_lng = Err.Number
    MessaggioErrore_s = "ERRORE DI ESECUZIONE ROUTINE : " & Err.Description

    ScriviLogErrore NomeRoutine_s, _
                    NumeroErrore_lng, _
                    MessaggioErrore_s

    ' //.................................................................//
    
    Resume ExitSub
End Sub


'//***********************************************************************************//
'//**           COMBINATA 05 @CMB_05 = ALTEZZA E LARGHEZZA DELLE SOTTOFORM   *** FINE ***
'//**
'//***********************************************************************************//







'//***********************************************************************************************//
'//*        -----------------------------------------------------------------------------------
'//*                     ROUTINE PER AUMENTARE E RIDURRE LE DIMENSIONI      *** FINE ***
'//*
'//*        -----------------------------------------------------------------------------------
'//*
'//***********************************************************************************************//





'//=====================================================================================================================//
'//                     CASELLE DI TESTO E COMBINATE DEL PROGETTO IMPORTA  *** FINE ***
'//                     + CurrentProject_txt + Cmb_01_txt + Cmb_02_txt + sourceDBPath_s_Txt  = @CMB_01, CMB_02 @CMB @PROGETTO
'//=====================================================================================================================//






'//@attiva@controllo@oggetti@esterni
'//CHIAMO LA FUNZIONE DI CONTROLLO OGGETTI DI SISTEMA ESTERNI
'//Note: controllo oggetti esterni.@CONTROLLO@OGGETTI_(attivo la procedura per il controllo degli oggetti esterni)
Private Sub Cmd_CONTROLLO_OGGETTI_ESTERNI_Click()
    ListExternalDBObjects
End Sub









'//=====================================================================================================================//
'//                     DELETE OGGETTI QUERY DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE QUERY HELP
'//=====================================================================================================================//


'//CANCELLA GLI OGGETTI QUERY *** INIZIO *** '//DELETE QUERIES
'//==========================================================================================================//
'//@GESTIONE@DELLE@QUERY                    @GESTONE@QUERY

Private Sub Cmd_Delete_QUERY_Click()

    Dim dbCurrent As DAO.Database
    Dim qry As DAO.QueryDef
    Dim Tables_Collection As Collection
    Dim varItem As Variant
    Dim MATRICE_queryNames_s() As String             'matrice di Query
    Dim tempCount As Integer
    Dim queryCount As Integer
    Dim qdf  As QueryDef
    

    
'//TODO: controllare se tale routine funziona???
On Error GoTo Err_Cmd_Delete_QUERY_Click
    
    
    '//RESET
    iCount = 0
    delete_QueryesCount_i = 0
    deletedTablesCount_i = 0
    deletedTables_s = ""
   
        
        
        
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_ATTIVA_GEST_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE  @CONTROLLO@PROGETTO
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//

              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION QUERY per ottenere la collezione di QUERY
                
                
                
                '//VADO ALLA COLLEZIONE DI QUERY CHE VIENE CARICATA E RESTITUITA ALLA VARIABILE
                Set queryes_Collection = CollectionQuery_PFunct()
                
              '//------------------------------------------------------------------------//
                    
                    ' Loop attraverso tutte I REPORT nel database corrente
                     Set dbCurrent = CurrentDb
                     
                     '//CALCOLO INDICE QUERY
                     iCount = dbCurrent.QueryDefs.count
                                       
                    '//APRO IL CORRENTE DB
                    Set dbCurrent = CurrentDb

                        
                    '//se ci sono query da cancellare
                    If iCount > 0 Then
                    
                        '//reimposto la matrice sulle query esistenti nel db corrente
                        ReDim MATRICE_queryNames_s(iCount - 1)
                        
                        ' Costruisci un array con i nomi dei Query
                        tempCount = 0
                        
                        '//CARICO LA MATRICE: ciclo for per caricare MATRICE_queryNames_s(tempCount) CON LE QUERY SALVATE
                        For Each qdf In dbCurrent.QueryDefs
                            MATRICE_queryNames_s(tempCount) = qdf.Name
                            tempCount = tempCount + 1
                        Next qdf
                        
                        ' Cancella i Query usando l'array dei nomi
                        For i = 0 To tempCount - 1
                            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
                            Debug.Print "Cancellazione Query: " & MATRICE_queryNames_s(i)
                            
                            '//CONTROLLO MATRICE = COLLECTION
                            Bool1 = InCollectionQueryes_Funct_b(queryes_Collection, MATRICE_queryNames_s(i))
                            
                            '//cancello solo se la matrice (i) esiste nella collection
                            If Bool1 = True Then
                                    Debug.Print
                                    Debug.Print "cancello la query trovata sia nella collezione che nella matrice : " & MATRICE_queryNames_s(i)
                                    DoCmd.DeleteObject acQuery, MATRICE_queryNames_s(i)
                                    
                                    If Err.Number <> 0 Then
                                        Debug.Print "Errore durante la cancellazione del Query: " & MATRICE_queryNames_s(i) & " - " & Err.Description
                                        Err.Clear  ' Pulisci l'errore per la prossima iterazione
                                    Else
                                        delete_QueryesCount_i = delete_QueryesCount_i + 1
                                    End If
                                    
                            End If '//If Bool1 = True Then
                            
                            On Error GoTo 0  ' Reimposta gestione errori
                        Next i
                    End If
    
                      '//APRI RIQUADRO QUERY INDIPENDENTEMENTE DALLA LORO PRESENZA
                      '//-------------------------------------------------------------------------------------------------//
                      ' Apri la sezione query indipendentemente dalla presenza di oggetti _
                        Dobbiamo controllare solo le query utente, escludendo quelle nascoste o di sistema. _
                        Ecco il codice corretto: _
                        Perché questo codice è più affidabile? _
                          - Conta solo le query visibili, ignorando quelle di sistema (~xxx). _
                          - Evita errori dovuti a conteggi errati su query cancellate. _
                          - Apre comunque il riquadro Query, anche se è vuoto. _

                        
                              
                            Dim db As DAO.Database
                            
                            Dim queryVisibili As Integer
                            
                            Set db = CurrentDb
                            queryVisibili = 0 ' Inizializza il contatore
                            
                            ' Conta solo le query realmente visibili
                            For Each qdf In db.QueryDefs
                                If Left(qdf.Name, 1) <> "~" Then ' Esclude le query di sistema nascoste
                                    queryVisibili = queryVisibili + 1
                                End If
                            Next qdf
                            
                            ' Se ci sono query visibili, apri la sezione Query
                            If queryVisibili > 0 Then
                                DoCmd.SelectObject acQuery, , True
                            Else
                                ' Se non ci sono query, prova ad aprire il riquadro selezionando un altro oggetto
                                On Error Resume Next
                                DoCmd.SelectObject acTable, , True ' Seleziona una tabella (se esiste)
                                DoCmd.SelectObject acQuery, , True ' Ora prova a selezionare la sezione Query
                                On Error GoTo 0
                            End If
                            
                            ' Pulizia memoria
                            Set qdf = Nothing
                            Set db = Nothing
    
                                          
                    '//-------------------------------------------------------------------------------------------------//
                    
            ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                MsgBox "LE QUERY CANCELLATA DA QUESTO DATABASE SONO :  " & delete_QueryesCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
            End If





'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_QUERY_Click:
    Exit Sub

Err_Cmd_Delete_QUERY_Click:
        
    Set dbCurrent = Nothing
    
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_QUERY_Click

End Sub

'//*** FINE ***
'//SENZA_PARAMETRI
'//========================================================================================================================================//



' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionQueryes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionQueryes_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionQueryes_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionQueryes_Funct_b = False
    On Error GoTo 0
End Function




'//DELETE QUERIES        ***FINE ***
'//==========================================================================================================//


'//DELETE QUERY HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_Delete_QUERY_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "Help_delete_QUERY.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub


'//DELETE QUERY HELP *** FINE ***
'//-----------------------------------------------------------------------------------//


'//=====================================================================================================================//
'//                     DELETE OGGETTI QUERY DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE QUERY HELP
'//=====================================================================================================================//







'//************************************************************************************************************************//
'//***
'//***                              COLLECTION OBJECT   *** FINE ***
'//************************************************************************************************************************//



'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI FORM DEL DB ESTERNO *** INIZIO ***
'//                     @IMPORTA@OGGETTI@FORM@MSYS
'//=====================================================================================================================//



'//IMPORTA TUTTI GLI OGGETTI FORM DEL DB MSYSDBEST - LE ATTIVITA
'//----------------------------------------------------------------------------------------//
'//@01_@IMPORTA@FORM@DIM                  -------------------> DIM VARIABILI
'//@02_@IMPORTA@FORM@REST                 -------------------> RESET DELLE VARIABILI
'//@03_@IMPORTA@FORM@LOG                  -------------------> LOG FILE INIZIALE
'//@04_@IMPORTA@FORM@COLL                 -------------------> RECUPERO COLLECTION DA IMPORTARE
'//@05_@IMPORTA@@FORMCONTROLLO 1          -------------------> CONTROLLO 1 ESISTONO OGGETTI COLLECTION
'//@06_@IMPORTA@FORM@CONTROLLO 2          -------------------> CONTROLLO 2 ESISTE PATH VALIDA
'//@07_@IMPORTA@FORM@CONTROLLO 3          -------------------> CONTROLLO 3 PATH NULL EXIT
'//@08_@IMPORTA@FORM@CONTROLLO 4          -------------------> CONTROLLO 4 SE DB NULL EXIT
'//@09_@IMPORTA@FORM@AVVISI+MSG           -------------------> AVVISI E IMPOSTAZIONE MESSAGGIO FINALE
'//@10_@IMPORTA@FORM@PROGRSESSBAR_01      -------------------> PROGRESS BAR STEP 01 RECUPERO PARAMETRI
'//@11_@IMPORTA@FORM@CONTROLLO_PROG       -------------------> CONTROLLO SE STIAMO NEL PROGETTO ESTERNO
'//@12_@IMPORTA@FORM@PROG@CICLO           -------------------> NEL PROGETTO ESTERNO CICLO NEGLI OGGETTI
'//@13_@IMPORTA@FORM@PROG@CONTR_01        -------------------> CONTROLLO 1 SE ESTE OGGETTO DA IMPORTARE
'//@14_@IMPORTA@FORM@PROG@IMPORT          -------------------> IMPORTAZIONE DELL'OGGETTO DAL DB ESTERNO
'//@15_@IMPORTA@FORM@PROGRSESSBAR_02      -------------------> PROGRESS BAR STEP 02 CALCOLO ED INVIO PARAMETRI
'//@15.BIS_@IMPORTA@FORM@PROGRESSBAR_CALL -------------------> VIENE CHIAMATA LA PROGRESS BAR CON I PARAMETRI
'//@16_@IMPORTA@FORM@LOG@FILE             -------------------> SE IL FILE NON VIENE TROVATO SCRITTURAZIONE LOG
'//@17_@IMPORTA@FORM@CLOSE                -------------------> CHIUSURA CON RITARDO DELLA PROGRESS BAR
'//@18_@IMPORTA@FORM@NOTHING              -------------------> IMPOSTO A NULL TUTTI GLI OGGETTI


Private Sub Cmd_Importa_FORM_Click()
    
    '//@01_@IMPORTA@FORM@DIM                  -------------------> DIM VARIABILI
    ' Dichiarazione delle variabili
    '//oggetti form e variabile nome form
    Dim formsCollection As Collection                       '//oggetti collezione
    Dim formName_v As Variant
    Dim deleted_Forms_s As String
    Dim imported_Forms_s As String
    Dim myforms_Name_s As String
    
    '//il database + path
    Dim sourceDB As DAO.Database
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim sourceDBPath As String
    
    '//contatori oggetti: TOTALE OGGETTI PREVISTI; da importare e cancellare
    Dim TotCollectionForms_i As Integer                          '//totale degli oggetti form della collezione da importare
    Dim importedCount_i As Integer                              '//oggetti importati
    Dim deletedCount_i As Integer                               '//oggetti cancellati
    
    
    '//bool e log file
    Dim Bool1 As Boolean
    Dim MyValue As Integer
    Dim logFile As Integer
    
    On Error GoTo ErrorHandler
    
    '//@02_@IMPORTA@FORM@REST                 -------------------> RESET DELLE VARIABILI
    '//RESET = Inizializza i contatori e le stringhe di riepilogo
    TotCollectionForms_i = 0
    importedCount_i = 0
    deletedCount_i = 0
    deleted_Forms_s = ""
    imported_Forms_s = ""
    TotPerc_dbl = 0
    ProgressPercentTwips_dbl = 0
    
    '//@03_@IMPORTA@FORM@LOG                  -------------------> LOG FILE INIZIALE
    '@FILE@LOG@FORM: Apre il file di log per scrivere i dettagli dell'importazione apre nella path corrente.
    logFile = FreeFile()
    Open CurrentProject.PATH & "\LOG_IMPORTA_FORM.TXT" For Output As #logFile
    Print #logFile, "Log Importazione Form - " & Now
    Print #logFile, "------------------------------------------------"
    Print #logFile,
    
    '//@04_@IMPORTA@FORM@COLL                 -------------------> RECUPERO COLLECTION DA IMPORTARE
    'Recupero la collezione di form da importare
      Set formsCollection = CollectionForm_PFunct()
      
    '//@05_@IMPORTA@@FORMCONTROLLO 1          -------------------> CONTROLLO 1 ESISTONO OGGETTI COLLECTION
    '//se non esistono oggetti collection EXIT
    If formsCollection.count = 0 Then
        Print #logFile, "Nessuna form da importare."
        Close #logFile
        '//Prima di uscire restto l'oggetto collection
        Set formsCollection = Nothing
        Exit Sub
    End If
    
    '//@06_@IMPORTA@FORM@CONTROLLO 2          -------------------> CONTROLLO 2 ESISTE PATH VALIDA
    '//RECUPERO LA PATH =  Controllo se il percorso del database di origine è valido
    sourceDBPath = PathFile_s_pFunct
    
    '//@07_@IMPORTA@FORM@CONTROLLO 3          -------------------> CONTROLLO 3 PATH NULL EXIT
    If sourceDBPath = "" Then
        MsgBox "Errore: path e file nulli.", vbCritical
        Print #logFile, "Errore: path e file nulli."
        Close #logFile
        Exit Sub
    End If
    
    '//@08_@IMPORTA@FORM@CONTROLLO 4          -------------------> CONTROLLO 4 SE DB NULL EXIT
    '//SE NON ESISTE IL FILE DB = Controllo se il file del database di origine esiste
    If Dir(sourceDBPath) = "" Then
        MsgBox "Errore: il file del database di origine non esiste.", vbCritical
        Print #logFile, "Errore: il file del database di origine non esiste."
        Close #logFile
        Exit Sub
    End If
    
    '//@09_@IMPORTA@FORM@AVVISI+MSG           -------------------> AVVISI E IMPOSTAZIONE MESSAGGIO FINALE
    '//.......................................................................................//
      '//AVVISI MESSAGGIO IMPOSTATO = TRUE - >  Controllo del valore di avviso
      MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
      If MyValue = 0 Then Exit Sub
      
      '//CONTROLLO NOME PROGETTO =  Controllo del nome del progetto
      Bool1 = CheckProjectName_b_PFunct()
      If Bool1 = True Then Exit Sub
      
    '//.......................................................................................//
    
    
        '//@10_@IMPORTA@FORM@PROGRSESSBAR_01      -------------------> PROGRESS BAR STEP 01 RECUPERO PARAMETRI
        '//@PROGRESBAR_STEP_01=ATTIVO LA PROGRESS BAR CHIAMANDO ANCHE LE 4 FUNZIONI PER RECUPERARE I PARAMETRI  @progress@bar
        '//********************************************************************************************************************//
        '//@PROGRES@BAR@FASE_01_(imposto le variabili e recupero i valori dalle funzioni)
        
                
                '//NOTE: devi creare 4 funzioni per avere i 4 parametri per chiamare la progress bar: _
                     @01)@IMPOSTO@VARIABILE@NUMERO_COLLECTION          ----> recupero il NroCollection_i _
                     @02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl _
                     @03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
                     @04)@FUNZIONE@TIPO_IMPORT                         ----> RESTITUISCE TipoImport_s _
                     @05)@ROUTINE_APRI_LA_PROGRES_BAR                  ----> RESTITUISCE NULL
                     
                      
                      DoEvents                      '//PER CONSENTIRE IL BLOCCO DEBUG
                  
                      '//@01)@IMPOSTO@VARIABILE@NUMERO_COLLECTION          ----> recupero il NroCollection_i _
                             Imposto il totale oggetti form da importare se esistono
                      TotCollectionForms_i = formsCollection.count
             
                    
                    '//@02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl
                       Twips_dbl = Calcola_Twips_pFunct(TotCollectionForms_i)
                    
                    '//@03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
                     '//chiama la funzione percentuale incremento e ne restituisce la _
                         percentuale come double
                     PercIncr_dbl = PercentualeIncremento_pFunct(TotCollectionForms_i)             '//calcolo il numero di collection form da controllare
         
                     '//@04)@FUNZIONE@TIPO_IMPORT                         ----> RESTITUISCE TipoImport_s _
                     '//chiama la funzione TIPO IMPORT
                         TipoImport_s = TipoImport_pFunct("\LOG_IMPORTA_FORM.TXT")
         
                        '//CONTROLLO_TIPO_IMPORTAZIONE: se la funzione restituisce NON ESISTE esci dalla IMPORTAZIONE
                         If TipoImport_s <> "IMPORTA_FORM" Then
                            
                            '//chiudo il file log prima di uscire
                            Close #logFile
                            GoTo Exit_Cmd_Importa_FORM_Click
                         End If
                         
                           
                     '//@05)@ROUTINE_APRI_LA_PROGRES_BAR                  ----> RESTITUISCE NULL
                         '//Apri la form della progressBar
                         '//
                             Call ApriProgressBar
                             ' Aspetta che il form sia effettivamente caricato
                              DoEvents
                         
        '//@PROGRESBAR_STEP_01=ATTIVO LA PROGRESS BAR CHIAMANDO ANCHE LE 4 FUNZIONI PER RECUPERARE I PARAMETRI *** FINE ***
        '//********************************************************************************************************************//
    
    
    '//===========================================================================================================================//
    '//             ATTIVAZIONE CICLO DI IMPORTAZIONE DAL PROGETTO ESTERNO        *** INIZIO ****
    '//===========================================================================================================================//
            
          '//@11_@IMPORTA@FORM@CONTROLLO_PROG       -------------------> CONTROLLO SE STIAMO NEL PROGETTO ESTERNO
          ' Se siamo su un progetto esterno (Bool1 = False), procedo con l'importazione
          If Bool1 = False Then
                ' Apertura del database di origine e connessione al database corrente
                Set ws = DBEngine.Workspaces(0)
                Set sourceDB = ws.OpenDatabase(sourceDBPath)
                Set db = CurrentDb
                
                
                '//reset
                ProgressPercentTwips_dbl = 0
                
                
                '//@12_@IMPORTA@FORM@PROG@CICLO           -------------------> NEL PROGETTO ESTERNO CICLO NEGLI OGGETTI
                ' Ciclo attraverso la collezione di form da importare
                For Each formName_v In formsCollection
                    myforms_Name_s = formName_v
                    
                    '//@13_@IMPORTA@FORM@PROG@CONTR_01        -------------------> CONTROLLO 1 SE ESTE OGGETTO DA IMPORTARE
                    ' Controllo se la form esiste nel database di origine
                    If ObjectExists_FORM(sourceDB, acForm, myforms_Name_s) Then
                        ' Se la form esiste nel database corrente, la elimino prima di importarla
                        If ObjectExists_FORM(db, acForm, myforms_Name_s) Then
                            DoCmd.DeleteObject acForm, myforms_Name_s
                            deletedCount_i = deletedCount_i + 1
                            deleted_Forms_s = deleted_Forms_s & myforms_Name_s & vbCrLf
                            Print #logFile, "Form eliminata: " & myforms_Name_s
                             
                              
                        End If
                        
                          '//@14_@IMPORTA@FORM@PROG@IMPORT          -------------------> IMPORTAZIONE DELL'OGGETTO DAL DB ESTERNO
                          ' Importazione della form dal database di origine
                          DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acForm, myforms_Name_s, myforms_Name_s
                          importedCount_i = importedCount_i + 1
                          imported_Forms_s = imported_Forms_s & myforms_Name_s & vbCrLf
                          
                          '//SCRIVO SUL @FILE@LOG
                          Print #logFile, "NRO: " & importedCount_i & " Form importata: " & myforms_Name_s
                          
                              
                        '//@15_@IMPORTA@FORM@PROGRSESSBAR_02      -------------------> PROGRESS BAR STEP 02 CALCOLO ED INVIO PARAMETRI
                        '//@PROGRESBAR_STEP_02= CALCOLO PARAMETRI ED INVIO DATI ALLA FUNZIONE DI AGGIORNAMENTO   @progress@bar
                        '//********************************************************************************************************************//
                        '//NOTE: impostazioni e recupero valori dalla funzioni base per la progress bar.
                            
                              '// RECUPERO il messaggio tipo di importazione delle collection e numero di oggetti da importare .
                                TipoImport_s = "FORM IMPORTATE NRO : " & importedCount_i & " su TOTALE OGGETTI: " & TotCollectionForms_i       '// STRINGA MESSAGGIO PER IL TIPO DI IMPORTAZIONE E NRO OGGETTI
                                
                                '//LA PERCENTUALE INCREMENTATA = totale di incremento % per ogni oggetto importato
                                TotPerc_dbl = TotPerc_dbl + PercIncr_dbl
                                
                                '//TOTALE % DI INCREMENTO IN TWIPS, con la ProgressPercentTwips_dbl è la variabile che indica la %di incremento per ogni oggetto + i _
                                   twips necessari per ogni elemento (su 10 totale della barra)
                                ProgressPercentTwips_dbl = ProgressPercentTwips_dbl + (PercIncr_dbl + Twips_dbl)          '//la % di accrescimento barra
                                
                                '//TODO: DA CANCELLARE...
                                'Dbl1 = Dbl1 + Twips_dbl                                         '//twips di aumento barra
                                
                                
                                          '//@15.BIS_@IMPORTA@FORM@PROGRESSBAR_CALL -------------------> VIENE CHIAMATA LA PROGRESS BAR CON I PARAMETRI
                                          '//CHIAMO LA FUNZIONE @AGGIORNA@PROGRESS@BAR
                                          '//.................................................................//
                                          '//note: chiamo la routine di aggiornamento della progress bar che _
                                              puo essere utilizzata per ulteriori integrazioni di calcolo. _
                                              Devo inviare ma devo inviare 5 parametri : _
                                                  TotCollectionForms_i       = il TOTALE POTENZIALE degli oggetti da importare _
                                                  NroCollection_i           = il numero della collezione per il calcolo; _
                                                  Twips                     = i twip di incremento _
                                                  ProgressPercentTwips_dbl  =  La percentuale di incremento ( ProgressPercentTwips_dbl = ProgressPercentTwips_dbl + PercIncr_dbl) _
                                                  TipoImport_s              = il tipo di importazione
                                                  
                                              AggiornaProgressBar _
                                                          TotCollectionForms_i, _
                                                          ProgressPercentTwips_dbl, _
                                                          Twips_dbl, _
                                                          TotPerc_dbl, _
                                                          TipoImport_s
                                          '//.................................................................//
        
                               
                                T = Timer
                                While Timer < T + 0.001
                                    DoEvents  ' Continua a gestire gli eventi durante l'attesa
                                Wend
                          
                                  
                        '//@PROGRESBAR_STEP_02= CALCOLO PARAMETRI ED INVIO DATI ALLA FUNZIONE DI AGGIORNAMENTO   *** FINE ***
                        '//********************************************************************************************************************//
        
            
                        
                    Else
                        
                        '//@16_@IMPORTA@FORM@LOG@FILE             -------------------> SE IL FILE NON VIENE TROVATO SCRITTURAZIONE LOG
                        ' Se la form non esiste nel database di origine, lo scrivo nel log
                        Print #logFile, "Form non trovata nel database di origine: " & myforms_Name_s
                    End If
                
                Next formName_v
        
        
                                        
                        '//@17_@IMPORTA@FORM@CLOSE                -------------------> CHIUSURA CON RITARDO DELLA PROGRESS BAR
                        '//CHIUSURA CON RITARDO DELLA BARRA DI PROGRESSIONE @progress@bar
                        '//---------------------------------------------------------------------------//
                    
                                        T = Timer
                                        
                                        '// RITARDO CHIUSURA DI 5 SECONDI
                                        While Timer < T + 5
                                            DoEvents  ' Continua a gestire gli eventi durante l'attesa
                                        Wend
                                        
                                        '//CHIUDI_LA_FORM_PROGRESSBAR
                                        DoCmd.Close acForm, "UTILITA_Msys_Frm21_frmProgressBar", acSaveNo
                                        
                        '//---------------------------------------------------------------------------//
                        
                        
                        
                '//@18_@IMPORTA@FORM@NOTHING              -------------------> IMPOSTO A NULL TUTTI GLI OGGETTI
                '//---------------------------------------------------------------------------//
                ' Chiudo il database di origine
                        Set frm_obj = Nothing                 ' Rilascia il riferimento al form PROGRES BARR
                      
                      '//chiudo il db
                      sourceDB.Close
                      Set sourceDB = Nothing
                      Set ws = Nothing
                  
                      ' Messaggio finale con riepilogo delle operazioni
                      If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                          MsgBox "Importazione completata." & vbCrLf & _
                              "Forms importate: " & importedCount_i & vbCrLf & imported_Forms_s & _
                              "Forms esistenti cancellate: " & deletedCount_i & vbCrLf & deleted_Forms_s, vbInformation
                      End If
                  
                      ' Scrivo nel file di log il riepilogo finale
                      Print #logFile, "Importazione completata con " & importedCount_i & " forms importate."
                      Close #logFile
                  
                      ' Pulizia delle variabili del db e della collection forms
                      Set db = Nothing
                      Set formsCollection = Nothing
                      
                '//---------------------------------------------------------------------------//
                
          End If '//If Bool1 = False Then = Progetto esterno
      
    '//===========================================================================================================================//
    '//             ATTIVAZIONE CICLO DI IMPORTAZIONE DAL PROGETTO ESTERNO        *** FINE ****
    '//===========================================================================================================================//
    
    
'// USCITA E GESTIONE ERRORI
'//---------------------------------------------------------------------------------------------------------//

Exit_Cmd_Importa_FORM_Click:

Exit Sub
    
ErrorHandler:
    ' Gestione degli errori con messaggio a video e scrittura nel log
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Print #logFile, "Errore durante l'importazione: " & Err.Description
    If Not sourceDB Is Nothing Then sourceDB.Close
    Close #logFile
    Set sourceDB = Nothing
    Set ws = Nothing
End Sub


'//CONTROLLO_ESISTEMA_FORM(efficientata la procedura con l'interrogazione sulla tabella MSysObjects per il controllo delle form esterne)
'//FUNZIONE DI CONTROLLO SE ESISTE L'OGGETTO FORM
Function ObjectExists_FORM(db As DAO.Database, objType As Integer, objName As String) As Boolean
    Dim rs As DAO.Recordset
    Dim formExists As Boolean
    
    '//CONTROLLO OGGETTO
    Debug.Print " Function ObjectExists_FORM - oggetto della ricerca : objName "
    Debug.Print objName
    
    
    formExists = False
    On Error Resume Next
    '//FILTRO CONTROLLO ESISTENZA FORM
    Set rs = db.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=-32768 AND Name='" & objName & "'")
    If Not rs.EOF Then
        formExists = True
    End If
    rs.Close
    Set rs = Nothing
    ObjectExists_FORM = formExists
    On Error GoTo 0
End Function
'//IMPORTA TUTTI GLI OGGETTI FORM DEL DB MSYSDBEST *** FINE ***
'//----------------------------------------------------------------------------------------//



'//         IMPORTA PROGETTO OGGETTI FORM *** FINE ***
'//=========================================================================================//



'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI FORM DEL DB ESTERNO *** FINE ***
'//                     @IMPORTA@OGGETTI@FORM@MSYS
'//=====================================================================================================================//




'//=====================================================================================================================//
'//                     DELETE OGGETTI FORM DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE TABELLA HELP
'//=====================================================================================================================//



'//CANCELLA GLI OGGETTI FORM *** INIZIO *** DELETE_FORM
'//======================================================================================================//
'//@GESTIONE@DELLE@FORM    (@MODELLO@DELETE@OGGETTI)

'//NOTA: visto che la cancellazione riguarda il database corrente in cui è stata importata _
        la classe di IMPORTA PROGETTO, la routine cancella solo se le FORM esistono _
        nel database corrente mediante il cofronto tra : _
        LA COLLECTION FORM = GLI OGGETTI FORM  che sono attivi nel database ESTERNO. _
        Prima di cancellare viene effettuato il confronto di esistenza _
        @cancellazione@form@caricate_(cancellazione delle FORM gia caricate nel db esterno) _
        @cancella@TABLE@esterne_(attenzione puoi cancellare le FORM Nel @db@esterno)
            
'//FAQ : _
        DOVE SI CANCELLANO LE FORM PRECARICATE _
        COME CANCELLARE LE FORM ESTERNE _
        LE FORM ESTERNE DA CANCELLARE


'// PARAMETRI DA CAMBIARE : _
    -----> FORM _
    -----> DELLE@FORM _
    -----> DELETE_FORM _
    -----> CollectionForm_PFunct _
    -----> formsCollection _
    -----> objectFomr _
    -----> MATRICE_FORM_Names_s _
    -----> formsCount_i _
    -----> FormDef _
    -----> frm _


Private Sub Cmd_Delete_FORM_Click()

    Dim dbCurrent As DAO.Database
    Dim frm As Form
    Dim formsCollection As Collection
    Dim varItem As Variant
    Dim MATRICE_FORM_Names_s() As String             'matrice di FORM
    Dim tempCount As Integer
    Dim formsCount_i As Integer
    Dim fdf As AccessObject
    Dim iCount As Integer
    Dim i As Integer
    Dim Bool1 As Boolean
    
    On Error GoTo Err_Cmd_Delete_FORM_Click

    '//RESET
    iCount = 0
    deletet_FormsCount_i = 0
    deleted_Forms_s = ""
    
    
        
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_ATTIVA_GEST_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE  @CONTROLLO@PROGETTO
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//

                      
                      
                  
                      '//ATTIVO LA COLLEZIONE
                      '//------------------------------------------------------------------------//
                      '//Note: CHIAMO LA FUNZIONE COLLECTION FORM per ottenere la collezione di FORM
                  
                      Set formsCollection = CollectionForm_PFunct()
                      '//------------------------------------------------------------------------//
                  
                      ' Loop attraverso tutte le FORM nel database corrente
                      Set dbCurrent = CurrentDb
                  
                      '//CALCOLO INDICE FORM
                      iCount = Application.CurrentProject.AllForms.count
                  
                      '//se ci sono FORM da cancellare
                      If iCount > 0 Then
                  
                          '//reimposto la matrice sulle FORM esistenti nel db corrente
                          ReDim MATRICE_FORM_Names_s(iCount - 1)
                  
                          ' Costruisci un array con i nomi dei FORM
                          tempCount = 0
                  
                          '//CARICO LA MATRICE: ciclo for per caricare MATRICE_FORM_Names_s(tempCount)
                          For i = 0 To iCount - 1
                              Set fdf = Application.CurrentProject.AllForms(i)
                              MATRICE_FORM_Names_s(tempCount) = fdf.Name
                              tempCount = tempCount + 1
                          Next i
                  
                          ' Cancella i FORM usando l'array dei nomi
                          For i = 0 To tempCount - 1
                              On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
                              Debug.Print "Cancellazione FORM: " & MATRICE_FORM_Names_s(i)
                  
                              '//CONTROLLO MATRICE = COLLECTION
                              Bool1 = InCollectionFORMes_Funct_b(formsCollection, MATRICE_FORM_Names_s(i))
                  
                              '//cancello solo se la matrice (i) esiste nella collection
                              If Bool1 = True Then
                                  Debug.Print
                                  Debug.Print "cancello la FORM trovata sia nella collezione che nella matrice : " & MATRICE_FORM_Names_s(i)
                                  DoCmd.DeleteObject acForm, MATRICE_FORM_Names_s(i)
                  
                                  If Err.Number <> 0 Then
                                      Debug.Print "Errore durante la cancellazione del FORM: " & MATRICE_FORM_Names_s(i) & " - " & Err.Description
                                      Err.Clear  ' Pulisci l'errore per la prossima iterazione
                                  Else
                                      deletet_FormsCount_i = deletet_FormsCount_i + 1
                                  End If
                  
                              End If '//If Bool1 = True Then
                  
                              On Error GoTo 0  ' Reimposta gestione errori
                          Next i
                      End If
                  
                  
                  
                  
                              '//APRI RIQUADRO FORMS INDIPENDENTEMENTE DALLA LORO PRESENZA
                              '//-------------------------------------------------------------------------------------------------//
                              '//NOTE= Apri la sezione macro indipendentemente dalla presenza di oggetti _
                                  Conta solo i moduli visibili, escludendo quelli di sistema. _
                                  Se ci sono moduli, apre direttamente la sezione FORMS con DoCmd.SelectObject acForm, , True. _
                                  Print Se; non; ci; sono; FORMS, prova; a; selezionare; una; tabella; e; poi; la; sezione; FORMS, per; forzare; l 'apertura del riquadro. _
                                  Evita errori con On Error Resume Next in caso di assenza di oggetti.
                
                                      Dim formVisibili As Integer
                                      Dim tableVisibili As Integer
                                      
                                      ' Inizializza il database corrente
                                      Set db = CurrentDb
                                      tableVisibili = 0 ' Inizializza il contatore

                                      
                                      ' Inizializza i contatori per i form e le tabelle
                                      formVisibili = 0
                                      tableVisibili = 0
                                      
                                      ' Conta le maschere visibili nel database
                                      For i = 0 To Application.CurrentProject.AllForms.count - 1
                                          formVisibili = formVisibili + 1
                                      Next i
                                      
                                                                            
                                      
                                      ' Conta solo le tabelle visibili, escludendo quelle di sistema
                                      For Each TableDef In db.TableDefs
                                          If Left(TableDef.Name, 4) <> "MSys" Then ' Esclude le tabelle di sistema (iniziano con "MSys")
                                              tableVisibili = tableVisibili + 1
                                          End If
                                      Next TableDef
                                      
                                       ' Se ci sono maschere visibili, apri la sezione Maschere
                                      If formVisibili > 0 Then
                                          DoCmd.SelectObject acForm, , True
                                      ElseIf tableVisibili > 0 Then
                                          ' Se non ci sono form, ma ci sono tabelle, seleziona una tabella per evitare errori
                                          On Error Resume Next
                                          DoCmd.SelectObject acTable, , True
                                          On Error GoTo 0
                                      End If
                                      
                                      ' Pulizia memoria: Anche se VBA rilascia automaticamente le variabili locali,
                                      ' è buona pratica resettare eventuali riferimenti se necessario.
                                      formVisibili = 0
                                      tableVisibili = 0
                                      i = 0
                                                                            ' Inizializza il database corrente
                                      Set db = CurrentDb
                                      tableVisibili = 0 ' Inizializza il contatore

                                          
                                    
                              '//APRI RIQUADRO FORMS INDIPENDENTEMENTE DALLA LORO PRESENZA *** FINE ***
                              '//-------------------------------------------------------------------------------------------------//
                            

                  
                  
                  
                  
                      ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
                      If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                          MsgBox "LE FORM CANCELLATA DA QUESTO DATABASE SONO :  " & deletet_FormsCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
                      End If

'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_FORM_Click:
    Exit Sub

Err_Cmd_Delete_FORM_Click:
    Set dbCurrent = Nothing
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_FORM_Click

End Sub

' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionFORMes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionFORMes_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionFORMes_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionFORMes_Funct_b = False
    On Error GoTo 0
End Function




'//CANCELLA GLI OGGETTI FORM *** FINE *** DELETE_FORM
'//=========================================================================================================








'//Cmd_DELETE_FORM_HELP_Click  *** INIZIO ***
'//-----------------------------------------------------------------------------------//

'//HELP DELETE FORM
Private Sub Cmd_Delete_FORM_HELP_Click()
                 
 '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
'//--------------------------------------------------------------------------------//--------//
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf.01.01
'//PARAMETRI            -> par_Path_s         = PATH _
                        -> par_NameFile_s               = NOME FILE _
                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
  '//IMPOSTO LE VARIABILI PATH E FILE
  '//
  Dim MyFile_s As String
  '//
  Dim MyPath_s  As String
  
        '//imposto i parametri - LA PATH
        '
        MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
        
        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
        MyFile_s = "HELP_Cmd_Delete_FORM.pdf "
            
            
                     
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)

             
        
'//--------------------------------------------------------------------------------//--------//

End Sub


'//Cmd_DELETE_FORM_HELP_Click  *** FINE ***
'//-----------------------------------------------------------------------------------//






'//=====================================================================================================================//
'//                     DELETE OGGETTI FORM DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE DELETE HELP
'//=====================================================================================================================//








'//=========================================================================================//
'//
'//                 GESTIONE DELLE TABELLE ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@TABELLE                    @GESTONE@TABELLE
'//=========================================================================================//
'//FAQ: _
        @HELP@TABELLE    @HELP@IMPORTA@TABELLE  @HELP@CANCELLA@TABELLE _
             @LA@GESTIONE@DELLE@TABELLE _
        @GESTIONE@TABELLE  @IMPORTA@TABELLE @CANCELLA@TABELLE


'//Cmd_Importa_FORM_HELP_Click()   ***INIZIO***
'//------------------------------------------------------------------------------------------------------//
'//Tipo           : Routine pubblica.
'//Attività       : Esegue codice help - ISTRUZIONI COMANDO IMPORTA FORM -
'//Note           : apre il file pdf HELP
'//Parametro      : NO PARAMETRI
'//Codice         : Cmd_Importa_FORM_HELP_Click.01
'//

Public Sub Cmd_Importa_FORM_HELP_Click()

'//MessaggiDiErrore
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String
 
On Error GoTo Err_Cmd_Importa_FORM_HELP_Click


        
        '//Imposto i parametri
        ProceduraAttivaEseguita_s = "Cmd_Importa_FORM_HELP_Click "
        ProceduraMessaggioErrore_s = "Errore nella procedura HELP IMPORTA FORM ??"
        
        '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
        '//--------------------------------------------------------------------------------//--------//
        '//NOTE                 -> Apro il file di tipo doc, zip o pdf
        '//CODICE               -> Function ApriFilePdf.01.01
        '//PARAMETRI            -> par_Path_s         = PATH _
                               -> par_NameFile_s               = NOME FILE _
                               -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                               
         '//IMPOSTO LE VARIABILI PATH E FILE
         '//
         Dim MyFile_s As String
         '//
         Dim MyPath_s  As String
         
               
               '//TODO: help dei vari button importa o cancella si trova nella path : c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\ _
                        del progetto MSYS_OBJECT_DB_EST che dovra essere integrato completamente nel progetto MSYS che si trova nell'omonima cartella. _
                        Quando tutti gli oggetti saranno spostati nel progetto MSYS_ATTIVA_GEST_OGGETTI.mdb allora tutto l'Help dovra essere spostato nella _
                        cartella c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\OBJECT\HELP\HELP_FORM\. Per ora rimane qui.  ??? DA FINIRE ???
               '//imposto i parametri - LA PATH
               MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
               
               '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
               MyFile_s = "HELP_Cmd_Importa_FORM.pdf "
                   
                   
                '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
                   Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
                   anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
                   totale.
                Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
               
                    
               
        '//--------------------------------------------------------------------------------//--------//








'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................


Exit_Cmd_Importa_FORM_HELP_Click:
    Exit Sub

Err_Cmd_Importa_FORM_HELP_Click:
 '//-------------------------------------------------------------------------------
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Importa_FORM_HELP_Click
'//-------------------------------------------------------------------------------
        
End Sub

'//*** FINE ***
'//Cmd_Importa_FORM_HELP_Click()   ***INIZIO***
'//------------------------------------------------------------------------------------------------------//




'//**************************************************************************************************************************************//
'//*
'//*
'//*                                    GESTIONE DELLE MACRO   *** INIZIO ***
'//*
'//* FAQ: _
        @GESTIONE@DELLE@MACRO                    @GESTONE@MACRO _
        @LA@GESTIONE@DELLE@MACRO _
        @GESTIONE@MACRO  @IMPORTA@MACRO @CANCELLA@MACRO

'//**************************************************************************************************************************************//




'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI MACRO DEL DB ESTERNO *** INIZIO ***
'//                     @IMPORTA@OGGETTI@MACRO@MSYS
'//=====================================================================================================================//



'//IMPORTA TUTTI GLI OGGETTI MACRO DEL DB MSYSDBEST - LE ATTIVITA
'//----------------------------------------------------------------------------------------//
'//@01_@IMPORTA@MACRO@DIM                  -------------------> DIM VARIABILI
'//@02_@IMPORTA@MACRO@REST                 -------------------> RESET DELLE VARIABILI
'//@03_@IMPORTA@MACRO@LOG                  -------------------> LOG FILE INIZIALE
'//@04_@IMPORTA@MACRO@COLL                 -------------------> RECUPERO COLLECTION DA IMPORTARE
'//@05_@IMPORTA@@MACROCONTROLLO 1          -------------------> CONTROLLO 1 ESISTONO OGGETTI COLLECTION
'//@06_@IMPORTA@MACRO@CONTROLLO 2          -------------------> CONTROLLO 2 ESISTE PATH VALIDA
'//@07_@IMPORTA@MACRO@CONTROLLO 3          -------------------> CONTROLLO 3 PATH NULL EXIT
'//@08_@IMPORTA@MACRO@CONTROLLO 4          -------------------> CONTROLLO 4 SE DB NULL EXIT
'//@09_@IMPORTA@MACRO@AVVISI+MSG           -------------------> AVVISI E IMPOSTAZIONE MESSAGGIO FINALE
'//@10_@IMPORTA@MACRO@PROGRSESSBAR_01      -------------------> PROGRESS BAR STEP 01 RECUPERO PARAMETRI
'//@11_@IMPORTA@MACRO@CONTROLLO_PROG       -------------------> CONTROLLO SE STIAMO NEL PROGETTO ESTERNO
'//@12_@IMPORTA@MACRO@PROG@CICLO           -------------------> NEL PROGETTO ESTERNO CICLO NEGLI OGGETTI
'//@13_@IMPORTA@MACRO@PROG@CONTR_01        -------------------> CONTROLLO 1 SE ESTE OGGETTO DA IMPORTARE
'//@14_@IMPORTA@MACRO@PROG@IMPORT          -------------------> IMPORTAZIONE DELL'OGGETTO DAL DB ESTERNO
'//@15_@IMPORTA@MACRO@PROGRSESSBAR_02      -------------------> PROGRESS BAR STEP 02 CALCOLO ED INVIO PARAMETRI
'//@15.BIS_@IMPORTA@MACRO@PROGRESSBAR_CALL -------------------> VIENE CHIAMATA LA PROGRESS BAR CON I PARAMETRI
'//@16_@IMPORTA@MACRO@LOG@FILE             -------------------> SE IL FILE NON VIENE TROVATO SCRITTURAZIONE LOG
'//@17_@IMPORTA@MACRO@CLOSE                -------------------> CHIUSURA CON RITARDO DELLA PROGRESS BAR
'//@18_@IMPORTA@MACRO@NOTHING              -------------------> IMPOSTO A NULL TUTTI GLI OGGETTI


Private Sub Cmd_Importa_MACRO_Click()
    
    '//@01_@IMPORTA@MACRO@DIM                  -------------------> DIM VARIABILI
    ' Dichiarazione delle variabili
    '//oggetti MACRO e variabile nome MACRO
    Dim Macros_Collection As Collection                       '//oggetti collezione
    Dim MACROS_Name_v As Variant
    Dim deleted_Macros_s As String
    Dim imported_Macros_s As String
    Dim myMACROS_Name_s As String
    
    '//il database + path
    Dim sourceDB As DAO.Database
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim sourceDBPath As String
    
    '//contatori oggetti: TOTALE OGGETTI PREVISTI; da importare e cancellare
    Dim TotCollection_MACROS_i As Integer                          '//totale degli oggetti MACRO della collezione da importare
    Dim importedCount_i As Integer                              '//oggetti importati
    Dim deletedCount_i As Integer                               '//oggetti cancellati
    
    
    '//bool e log file
    Dim Bool1 As Boolean
    Dim MyValue As Integer
    Dim logFile As Integer
    
    On Error GoTo ErrorHandler
    
    '//@02_@IMPORTA@MACRO@REST                 -------------------> RESET DELLE VARIABILI
    '//RESET = Inizializza i contatori e le stringhe di riepilogo
    TotCollection_MACROS_i = 0
    importedCount_i = 0
    deletedCount_i = 0
    deleted_Macros_s = ""
    imported_Macros_s = ""
    TotPerc_dbl = 0
    ProgressPercentTwips_dbl = 0
    
    '//@03_@IMPORTA@MACRO@LOG                  -------------------> LOG FILE INIZIALE
    '@FILE@LOG@MACRO: Apre il file di log per scrivere i dettagli dell'importazione apre nella path corrente.
    logFile = FreeFile()
    Open CurrentProject.PATH & "\LOG_IMPORTA_MACRO.TXT" For Output As #logFile
    Print #logFile, "Log Importazione MACRO - " & Now
    Print #logFile, "------------------------------------------------"
    Print #logFile,
    
    '//@04_@IMPORTA@MACRO@COLL                 -------------------> RECUPERO COLLECTION DA IMPORTARE
    'Recupero la collezione di MACRO da importare
      Set Macros_Collection = CollectionMacro_PFunct()
      
    '//@05_@IMPORTA@@MACROCONTROLLO 1          -------------------> CONTROLLO 1 ESISTONO OGGETTI COLLECTION
    '//se non esistono oggetti collection EXIT
    If Macros_Collection.count = 0 Then
        Print #logFile, "Nessuna MACRO da importare."
        Close #logFile
        '//Prima di uscire restto l'oggetto collection
        Set Macros_Collection = Nothing
        Exit Sub
    End If
    
    '//@06_@IMPORTA@MACRO@CONTROLLO 2          -------------------> CONTROLLO 2 ESISTE PATH VALIDA
    '//RECUPERO LA PATH =  Controllo se il percorso del database di origine è valido
    sourceDBPath = PathFile_s_pFunct
    
    '//@07_@IMPORTA@MACRO@CONTROLLO 3          -------------------> CONTROLLO 3 PATH NULL EXIT
    If sourceDBPath = "" Then
        MsgBox "Errore: path e file nulli.", vbCritical
        Print #logFile, "Errore: path e file nulli."
        Close #logFile
        Exit Sub
    End If
    
    '//@08_@IMPORTA@MACRO@CONTROLLO 4          -------------------> CONTROLLO 4 SE DB NULL EXIT
    '//SE NON ESISTE IL FILE DB = Controllo se il file del database di origine esiste
    If Dir(sourceDBPath) = "" Then
        MsgBox "Errore: il file del database di origine non esiste.", vbCritical
        Print #logFile, "Errore: il file del database di origine non esiste."
        Close #logFile
        Exit Sub
    End If
    
    '//@09_@IMPORTA@MACRO@AVVISI+MSG           -------------------> AVVISI E IMPOSTAZIONE MESSAGGIO FINALE
    '//.......................................................................................//
      '//AVVISI MESSAGGIO IMPOSTATO = TRUE - >  Controllo del valore di avviso
      MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
      If MyValue = 0 Then Exit Sub
      
      '//CONTROLLO NOME PROGETTO =  Controllo del nome del progetto
      Bool1 = CheckProjectName_b_PFunct()
      If Bool1 = True Then Exit Sub
      
    '//.......................................................................................//
    
    
        '//@10_@IMPORTA@MACRO@PROGRSESSBAR_01      -------------------> PROGRESS BAR STEP 01 RECUPERO PARAMETRI
        '//@PROGRESBAR_STEP_01=ATTIVO LA PROGRESS BAR CHIAMANDO ANCHE LE 4 FUNZIONI PER RECUPERARE I PARAMETRI  @progress@bar
        '//********************************************************************************************************************//
        '//@PROGRES@BAR@FASE_01_(imposto le variabili e recupero i valori dalle funzioni)
        
                
                '//NOTE: devi creare 4 funzioni per avere i 4 parametri per chiamare la progress bar: _
                     @01)@IMPOSTO@VARIABILE@NUMERO_COLLECTION          ----> recupero il NroCollection_i _
                     @02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl _
                     @03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
                     @04)@FUNZIONE@TIPO_IMPORT                         ----> RESTITUISCE TipoImport_s _
                     @05)@ROUTINE_APRI_LA_PROGRES_BAR                  ----> RESTITUISCE NULL
                     
             
                    DoEvents                      '//PER CONSENTIRE IL BLOCCO DEBUG
                  
                      '//@01)@IMPOSTO@VARIABILE@NUMERO_COLLECTION          ----> recupero il NroCollection_i _
                             Imposto il totale oggetti MACRO da importare se esistono
                      TotCollection_MACROS_i = Macros_Collection.count
             
                    
                    '//@02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl
                       Twips_dbl = Calcola_Twips_pFunct(TotCollection_MACROS_i)
                    
                    '//@03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
                     '//chiama la funzione percentuale incremento e ne restituisce la _
                         percentuale come double
                     PercIncr_dbl = PercentualeIncremento_pFunct(TotCollection_MACROS_i)             '//calcolo il numero di collection MACRO da controllare
         
                     '//@04)@FUNZIONE@TIPO_IMPORT                         ----> RESTITUISCE TipoImport_s _
                     '//chiama la funzione TIPO IMPORT
                         TipoImport_s = TipoImport_pFunct("\LOG_IMPORTA_MACRO.TXT")
         
         
                        '//CONTROLLO_TIPO_IMPORTAZIONE: se la funzione restituisce NON ESISTE esci dalla IMPORTAZIONE
                         If TipoImport_s <> "IMPORTA_MACRO" Then
                             '//chiudo il file log prima di uscire
                             Close #logFile
                          
                            GoTo Exit_Cmd_Importa_MACRO_Click
                         End If
                     
                     
                     
                     '//@05)@ROUTINE_APRI_LA_PROGRES_BAR                  ----> RESTITUISCE NULL
                         '//Apri la MACRO della progressBar
                         '//
                             Call ApriProgressBar
                             ' Aspetta che il MACRO sia effettivamente caricato
                              DoEvents
                         
        '//@PROGRESBAR_STEP_01=ATTIVO LA PROGRESS BAR CHIAMANDO ANCHE LE 4 FUNZIONI PER RECUPERARE I PARAMETRI *** FINE ***
        '//********************************************************************************************************************//
    
    
    '//===========================================================================================================================//
    '//             ATTIVAZIONE CICLO DI IMPORTAZIONE DAL PROGETTO ESTERNO        *** INIZIO ****
    '//===========================================================================================================================//
            
          '//@11_@IMPORTA@MACRO@CONTROLLO_PROG       -------------------> CONTROLLO SE STIAMO NEL PROGETTO ESTERNO
          ' Se siamo su un progetto esterno (Bool1 = False), procedo con l'importazione
          If Bool1 = False Then
                ' Apertura del database di origine e connessione al database corrente
                Set ws = DBEngine.Workspaces(0)
                Set sourceDB = ws.OpenDatabase(sourceDBPath)
                Set db = CurrentDb
                
                
                '//reset
                ProgressPercentTwips_dbl = 0
                
                
                '//@12_@IMPORTA@MACRO@PROG@CICLO           -------------------> NEL PROGETTO ESTERNO CICLO NEGLI OGGETTI
                ' Ciclo attraverso la collezione di MACRO da importare
                For Each MACROS_Name_v In Macros_Collection
                    myMACROS_Name_s = MACROS_Name_v
                    
                    '//@13_@IMPORTA@MACRO@PROG@CONTR_01        -------------------> CONTROLLO 1 SE ESTE OGGETTO DA IMPORTARE
                    ' Controllo se la MACRO esiste nel database di origine
                    If ObjectExists_MACRO(sourceDB, acMacro, myMACROS_Name_s) Then
                        ' Se la MACRO esiste nel database corrente, la elimino prima di importarla
                        If ObjectExists_MACRO(db, acMacro, myMACROS_Name_s) Then
                            DoCmd.DeleteObject acMacro, myMACROS_Name_s
                            deletedCount_i = deletedCount_i + 1
                            deleted_Macros_s = deleted_Macros_s & myMACROS_Name_s & vbCrLf
                            Print #logFile, "MACRO eliminata: " & myMACROS_Name_s
                             
                              
                        End If
                        
                          '//@14_@IMPORTA@MACRO@PROG@IMPORT          -------------------> IMPORTAZIONE DELL'OGGETTO DAL DB ESTERNO
                          ' Importazione della MACRO dal database di origine
                          DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acMacro, myMACROS_Name_s, myMACROS_Name_s
                          importedCount_i = importedCount_i + 1
                          imported_Macros_s = imported_Macros_s & myMACROS_Name_s & vbCrLf
                          
                          '//SCRIVO SUL @FILE@LOG
                          Print #logFile, "NRO: " & importedCount_i & " MACRO importata: " & myMACROS_Name_s
                          
                              
                        '//@15_@IMPORTA@MACRO@PROGRSESSBAR_02      -------------------> PROGRESS BAR STEP 02 CALCOLO ED INVIO PARAMETRI
                        '//@PROGRESBAR_STEP_02= CALCOLO PARAMETRI ED INVIO DATI ALLA FUNZIONE DI AGGIORNAMENTO   @progress@bar
                        '//********************************************************************************************************************//
                        '//NOTE: impostazioni e recupero valori dalla funzioni base per la progress bar.
                            
                              '// RECUPERO il messaggio tipo di importazione delle collection e numero di oggetti da importare .
                                TipoImport_s = "MACRO IMPORTATE NRO : " & importedCount_i & " su TOTALE OGGETTI: " & TotCollection_MACROS_i       '// STRINGA MESSAGGIO PER IL TIPO DI IMPORTAZIONE E NRO OGGETTI
                                
                                '//LA PERCENTUALE INCREMENTATA = totale di incremento % per ogni oggetto importato
                                TotPerc_dbl = TotPerc_dbl + PercIncr_dbl
                                
                                '//TOTALE % DI INCREMENTO IN TWIPS, con la ProgressPercentTwips_dbl è la variabile che indica la %di incremento per ogni oggetto + i _
                                   twips necessari per ogni elemento (su 10 totale della barra)
                                ProgressPercentTwips_dbl = ProgressPercentTwips_dbl + (PercIncr_dbl + Twips_dbl)          '//la % di accrescimento barra
                                
                                '//TODO: DA CANCELLARE...
                                'Dbl1 = Dbl1 + Twips_dbl                                         '//twips di aumento barra
                                
                                
                                          '//@15.BIS_@IMPORTA@MACRO@PROGRESSBAR_CALL -------------------> VIENE CHIAMATA LA PROGRESS BAR CON I PARAMETRI
                                          '//CHIAMO LA FUNZIONE @AGGIORNA@PROGRESS@BAR
                                          '//.................................................................//
                                          '//note: chiamo la routine di aggiornamento della progress bar che _
                                              puo essere utilizzata per ulteriori integrazioni di calcolo. _
                                              Devo inviare ma devo inviare 5 parametri : _
                                                  TotCollection_MACROS_i       = il TOTALE POTENZIALE degli oggetti da importare _
                                                  NroCollection_i           = il numero della collezione per il calcolo; _
                                                  Twips                     = i twip di incremento _
                                                  ProgressPercentTwips_dbl  =  La percentuale di incremento ( ProgressPercentTwips_dbl = ProgressPercentTwips_dbl + PercIncr_dbl) _
                                                  TipoImport_s              = il tipo di importazione
                                                  
                                              AggiornaProgressBar _
                                                          TotCollection_MACROS_i, _
                                                          ProgressPercentTwips_dbl, _
                                                          Twips_dbl, _
                                                          TotPerc_dbl, _
                                                          TipoImport_s
                                          '//.................................................................//
        
                               
                                T = Timer
                                While Timer < T + 0.001
                                    DoEvents  ' Continua a gestire gli eventi durante l'attesa
                                Wend
                          
                                  
                        '//@PROGRESBAR_STEP_02= CALCOLO PARAMETRI ED INVIO DATI ALLA FUNZIONE DI AGGIORNAMENTO   *** FINE ***
                        '//********************************************************************************************************************//
        
            
                        
                    Else
                        
                        '//@16_@IMPORTA@MACRO@LOG@FILE             -------------------> SE IL FILE NON VIENE TROVATO SCRITTURAZIONE LOG
                        ' Se la MACRO non esiste nel database di origine, lo scrivo nel log
                        Print #logFile, "MACRO non trovata nel database di origine: " & myMACROS_Name_s
                    End If
                
                Next MACROS_Name_v
        
        
                                        
                        '//@17_@IMPORTA@MACRO@CLOSE                -------------------> CHIUSURA CON RITARDO DELLA PROGRESS BAR
                        '//CHIUSURA CON RITARDO DELLA BARRA DI PROGRESSIONE @progress@bar
                        '//---------------------------------------------------------------------------//
                    
                                        T = Timer
                                        
                                        '// RITARDO CHIUSURA DI 5 SECONDI
                                        While Timer < T + 5
                                            DoEvents  ' Continua a gestire gli eventi durante l'attesa
                                        Wend
                                        
                                        '//CHIUDI_LA_FORM_PROGRESSBAR
                                        DoCmd.Close acForm, "UTILITA_Msys_Frm21_frmProgressBar", acSaveNo
                                        
                        '//---------------------------------------------------------------------------//
                        
                        
                        
                '//@18_@IMPORTA@MACRO@NOTHING              -------------------> IMPOSTO A NULL TUTTI GLI OGGETTI
                '//---------------------------------------------------------------------------//
                ' Chiudo il database di origine
                        Set frm_obj = Nothing                 ' Rilascia il riferimento al MACRO PROGRES BARR
                      
                      '//chiudo il db
                      sourceDB.Close
                      Set sourceDB = Nothing
                      Set ws = Nothing
                  
                      ' Messaggio finale con riepilogo delle operazioni
                      If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                          MsgBox "Importazione completata." & vbCrLf & _
                              "MACROs importate: " & importedCount_i & vbCrLf & imported_Macros_s & _
                              "MACROs esistenti cancellate: " & deletedCount_i & vbCrLf & deleted_Macros_s, vbInMACROation
                      End If
                  
                      ' Scrivo nel file di log il riepilogo finale
                      Print #logFile, "Importazione completata con " & importedCount_i & " MACROs importate."
                      Close #logFile
                  
                      ' Pulizia delle variabili del db e della collection MACROs
                      Set db = Nothing
                      Set Macros_Collection = Nothing
                      
                '//---------------------------------------------------------------------------//
                
          End If '//If Bool1 = False Then = Progetto esterno
      
    '//===========================================================================================================================//
    '//             ATTIVAZIONE CICLO DI IMPORTAZIONE DAL PROGETTO ESTERNO        *** FINE ****
    '//===========================================================================================================================//
    
    
'// USCITA E GESTIONE ERRORI
'//---------------------------------------------------------------------------------------------------------//

Exit_Cmd_Importa_MACRO_Click:

Exit Sub
    
ErrorHandler:
    ' Gestione degli errori con messaggio a video e scrittura nel log
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Print #logFile, "Errore durante l'importazione: " & Err.Description
    If Not sourceDB Is Nothing Then sourceDB.Close
    Close #logFile
    Set sourceDB = Nothing
    Set ws = Nothing
End Sub


'//CONTROLLO_ESISTEMA_MACRO(efficientata la procedura con l'interrogazione sulla tabella MSysObjects per il controllo delle MACRO esterne)
'//FUNZIONE DI CONTROLLO SE ESISTE L'OGGETTO MACRO
Function ObjectExists_MACRO(db As DAO.Database, objType As Integer, objName As String) As Boolean
    Dim rs As DAO.Recordset
    Dim MACROExists As Boolean
    
    
    MACROExists = False
    On Error Resume Next
    '//FILTRO CONTROLLO ESISTENZA MACRO
    Set rs = db.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=-32766 AND Name='" & objName & "'")
    If Not rs.EOF Then
        MACROExists = True
    End If
    rs.Close
    Set rs = Nothing
    ObjectExists_MACRO = MACROExists
    On Error GoTo 0
End Function
'//IMPORTA TUTTI GLI OGGETTI MACRO DEL DB MSYSDBEST *** FINE ***
'//----------------------------------------------------------------------------------------//






'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI MACRO DEL DB ESTERNO *** FINE ***
'//                     @IMPORTA@OGGETTI@MACRO@MSYS
'//=====================================================================================================================//








'//=====================================================================================================================//
'//                     DELETE OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE MACRO HELP
'//=====================================================================================================================//
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti MACRO controllando prima se _
            nel db corrente di destinazione già esiste la MACRO per effettuare la cancellazione _
            prima dell 'importazione.
'//NOTA: visto che la cancellazione riguarda il database corrente in cui è stata importata _
        la classe di IMPORTA PROGETTO, la routine cancella solo se le MACRO esistono _
        nel database corrente mediante il confronto tra : _
        LA COLLECTION MACRO = GLI OGGETTI MACRO che sono attivi nel database ESTERNO. _
        Prima di cancellare viene effettuato il confronto di esistenza _
        @cancellazione@macro@caricati_(cancellazione delle MACRO già caricate nel db esterno) _
        @cancella@MACRO@esterne_(attenzione puoi cancellare le MACRO nel @db@esterno)

'//FAQ : _
        DOVE SI CANCELLANO LE MACRO PRECARICATE _
        COME CANCELLARE LE MACRO ESTERNE _
        LE MACRO ESTERNE DA CANCELLARE
 
'// PARAMETRI DA CAMBIARE : _
    -----> MACRO _
    -----> DELLE@MACRO _
    -----> DELETE_MACRO _
    -----> CollectionMacro_PFunct _
    -----> Macros_Collection _
    -----> objectMacro _
    -----> MATRICE_MACRO_Names_s _
    -----> macros_Count_i _
    -----> MacroDef _
    -----> mdf _


Private Sub Cmd_Delete_MACRO_Click()

    Dim dbCurrent As DAO.Database
    Dim Dbs As Object
    Dim obj As AccessObject
    
    Dim Macros_Collection As Collection
    Dim varItem As Variant
    Dim MATRICE_MACRO_Names_s() As String             'matrice di MACRO
    Dim tempCount As Integer
    Dim macros_Count_i As Integer
    Dim od As AccessObject
    Dim iCount As Integer                              'I CONTATORI
    Dim i As Integer
    Dim Bool1 As Boolean
    
    On Error GoTo Err_Cmd_Delete_MACRO_Click

    '//RESET
    iCount = 0
    deletetmacros_Count_i = 0
    deleted_Macros_s = ""

    
    
    
    
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_ATTIVA_GEST_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE  @CONTROLLO@PROGETTO
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//
    
    
    
    '//CANCELLO LE MACRO DEL DB CORRENTE SE UGUALI ALLA MATRICE MACRO DA IMPORTARE **** INIZIO ****
    '//=====================================================================================================================//
        
        '//@01MACRO_CARICO_LA_COLLEZIONE_OGGETTI=imposto la collection dei nomi delle _
                macro da caricare nell'oggetto colletion
        '//------------------------------------------------------------------------//
        
            Set Macros_Collection = CollectionMacro_PFunct()
            
        '//------------------------------------------------------------------------//
        
        '//@02MACRO_IMPOSTO_OGGETTI_DB_CORRENTE=imposto le variabili oggetto del _
            DB CORRENTE e del CONTATORE MACRO
        '//------------------------------------------------------------------------//
            '// Loop attraverso tutte le MACRO nel database corrente _
                Dim dbCurrent As DAO.Database
            Set dbCurrent = CurrentDb
        
            '//CALCOLO INDICE MACRO
            iCount = CurrentProject.AllMacros.count
        '//------------------------------------------------------------------------//
            
        
        
            '//@03MACRO_IF_DI_CONTROLLO_MACRO=IF>0 CI SONO MACRO DA CANCELLARE NE DB CORRENTE *** INIZIO ***
            '//--------------------------------------------------------------------------------------------------------//
                  If iCount > 0 Then
                                 
                    '//@03MACRO_IF.01=REIMPOSTO LA MATRICE DELLE MACRO+RESET CONTATORE
                    '//............................................................................................//
                        '//reimposto la matrice sulle MACRO esistenti nel db corrente _
                        Dim MATRICE_MACRO_Names_s() As String             'matrice di MACRO
                        ReDim MATRICE_MACRO_Names_s(iCount - 1)
                        ' e reset contatore temporaneo
                        tempCount = 0
                        
                    '//............................................................................................//
                    
                            
                    '//@03MACRO_IF.02=CARICO LA MATRICE_MACRO_Names_s(tempCount)
                    '//............................................................................................//
                                     
                                Set Dbs = Application.CurrentProject
                                ' Ricerca oggetti AccessObject aperti in insieme AllMacros. _
                                    Dim obj As AccessObject
                                For Each obj In Dbs.AllMacros
                                        ' Stampa nome di obj.
                                        Debug.Print obj.Name
                                        MATRICE_MACRO_Names_s(tempCount) = obj.Name
                                        tempCount = tempCount + 1
                                Next obj
                        
                    '//............................................................................................//
                        
                            
                            '//@03MACRO_IF.03=CANCELLO LE MACRO DEL DB CORRENTE SE = ARRAY MATRICE_MACRO_Names_s(tempCount)
                            '//............................................................................................//
                            
                                    ' Cancella le MACRO usando l'array dei nomi
                                    For i = 0 To tempCount - 1
                                        On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
                                        Debug.Print "Cancellazione MACRO: " & MATRICE_MACRO_Names_s(i)
                            
                                        '//CONTROLLO MATRICE = COLLECTION
                                        Bool1 = InCollectionMacrosDELETE_Funct_b(Macros_Collection, MATRICE_MACRO_Names_s(i))
                            
                                        '//cancello solo se la matrice (i) esiste nella collection
                                        If Bool1 = True Then
                                            Debug.Print
                                            Debug.Print "cancello la MACRO trovata sia nella collezione che nella matrice : " & MATRICE_MACRO_Names_s(i)
                                            DoCmd.DeleteObject acMacro, MATRICE_MACRO_Names_s(i)
                            
                                            If Err.Number <> 0 Then
                                                Debug.Print "Errore durante la cancellazione della MACRO: " & MATRICE_MACRO_Names_s(i) & " - " & Err.Description
                                                Err.Clear  ' Pulisci l'errore per la prossima iterazione
                                            Else
                                                deletetmacros_Count_i = deletetmacros_Count_i + 1
                                            End If
                            
                                        End If '//If Bool1 = True Then
                            
                                        On Error GoTo 0  ' Reimposta gestione errori
                                    Next i
                                    
                            '//............................................................................................//
                            
                    End If '// If icount > 0 Then
        
            '//@03MACRO_IF_DI_CONTROLLO_MACRO=IF>0 CI SONO MACRO DA CANCELLARE NE DB CORRENTE *** INIZIO ***
            '//--------------------------------------------------------------------------------------------------------//

    '//CANCELLO LE MACRO DEL DB CORRENTE SE UGUALI ALLA MATRICE MACRO DA IMPORTARE **** FINE ****
    '//=====================================================================================================================//
                
                
                
                
                      '//APRI RIQUADRO MACRO INDIPENDENTEMENTE DALLA LORO PRESENZA
                      '//-------------------------------------------------------------------------------------------------//
                      '//NOTE= Apri la sezione macro indipendentemente dalla presenza di oggetti _

                            
                            Dim db As DAO.Database
                            Dim doc As Document
                            Dim macroVisibili As Integer
                            
                            Set db = CurrentDb
                            macroVisibili = 0 ' Inizializza il contatore
                            
                            ' Conta solo le macro visibili nel database
                            For Each doc In db.Containers("Scripts").Documents
                                If Left(doc.Name, 1) <> "~" Then ' Esclude eventuali macro di sistema nascoste
                                    macroVisibili = macroVisibili + 1
                                End If
                            Next doc
                            
                            ' Se ci sono macro visibili, apri la sezione Macro
                            If macroVisibili > 0 Then
                                DoCmd.SelectObject acMacro, , True
                            Else
                                ' Se non ci sono macro, prova a selezionare un altro oggetto per forzare l'apertura del riquadro
                                On Error Resume Next
                                DoCmd.SelectObject acTable, , True ' Seleziona una tabella (se esiste)
                                DoCmd.SelectObject acMacro, , True ' Ora prova a selezionare la sezione Macro
                                On Error GoTo 0
                            End If
                            
                            ' Pulizia memoria
                            Set doc = Nothing
                            Set db = Nothing
                            
                            
                            
                      '//APRI RIQUADRO MACRO INDIPENDENTEMENTE DALLA LORO PRESENZA *** FINE ***
                      '//-------------------------------------------------------------------------------------------------//
    



                    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
                    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                        MsgBox "Le MACRO CANCELLATE DA QUESTO DATABASE SONO :  " & deletetmacros_Count_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
                    End If



'//USCITA E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_MACRO_Click:
    Exit Sub

Err_Cmd_Delete_MACRO_Click:
    Set dbCurrent = Nothing
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_MACRO_Click

End Sub

' Funzione per verificare se un elemento esiste nella collezione di MACRO
Function InCollectionMacrosDELETE_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionMacrosDELETE_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionMacrosDELETE_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionMacrosDELETE_Funct_b = False
    On Error GoTo 0
End Function



Private Sub Cmd_Delete_MACRO_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_Delete_MACRO.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa MODULI _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//


End Sub


'//=====================================================================================================================//
'//                     DELETE OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE MACRO HELP
'//=====================================================================================================================//




'//**************************************************************************************************************************************//
'//*
'//*
'//*                                    GESTIONE DELLE MACRO   *** INIZIO ***
'//*
'//*
'//**************************************************************************************************************************************//







'//IMPORTA MODULI HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_IMPORTA_MODULI_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_Importa_MODULI.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa MODULI _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub


'//IMPORTA MODULI HELP *** FINE ***
'//-----------------------------------------------------------------------------------//


'//***************************************************************************************************************************************//
'//*
'//*                              IMPORTA OGGETTI FORM DELLE UTILITA    *** INIZIO ***
'//*                              @FORM@UTILITA_(LA GESTIONE DEGLI OGGETTI DELLA FORM UTILITA @IMPORTA@OGGETTI)
'//*
'//***************************************************************************************************************************************//

'//@IMPORTA@OGGETTI@UTILITA_(vengono importati tutti gli oggetti al servizio della form utilita)
'//IMPORTA OGGETTI UTILITA *** INIZIO ***
'//==================================================================================================================//
'//@IMPORTA@OGGETTI@UTILITA_(vengono importati tutti gli oggetti al servizio della form utilita)
'//IMPORTA OGGETTI UTILITA *** INIZIO ***
'//@UTILITA@IMPORTA@COLLETION_(IMPORTA LE COLLEZIONI DI OGGETTI DA CARICARE PER LA GESTIONE DELLE UTILITA)
'//NOTE: Viene inibito il comando importa oggetti se ci troviamo nel progetto mdb originale, denominato _
         MSYS_ATTIVA_GEST_OGGETTI.mdb
Private Sub Cmd_IMPORTA_OGGETTI_UTILITA_Click()
    On Error GoTo ErrHandler ' Abilita la gestione degli errori

    ' --- INIZIO DEL CODICE DELLA PROCEDURA ---
    Dim dbCurrent As DAO.Database
    Dim dbSource As DAO.Database
    Dim tbl As String
    Dim tdf As DAO.TableDef
    Dim qdf As DAO.QueryDef
    Dim frm As AccessObject
    Dim myName As String
    Dim objectExistsInCollection As Boolean
    Dim countTabelle As Integer
    Dim countQuery As Integer
    Dim countModuli As Integer
    Dim countForms As Integer
    Dim countReports As Integer
    Dim countMacros As Integer
    Dim deleteTablesCount_i As Integer
    Dim importedTableCount_i As Integer
    Dim delete_QueryesCount_i As Integer
    Dim imported_QueryesCount_i As Integer
    Dim deleteNameTables_s As String
    Dim importedNameTables_s As String
    Dim deleteName_Queryes_s As String
    Dim importedName_Queryes_s As String
    Dim deleteNameForms_s As String
    Dim importedNameForms_s As String
    Dim deleteNameReports_s As String
    Dim importedNameReports_s As String
    Dim deleteNameMacros_s As String
    Dim importedNameMacros_s As String
    Dim varItem As Variant

    ' Percorso del database sorgente
    Dim sourceDBPath_s As String
    sourceDBPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_ATTIVA_GEST_OGGETTI.mdb"

    ' Crea le collezioni per tabelle, query, moduli, form, report e macro
    Dim coll_Tables_UTILITA As Collection
    Dim coll_Queryes_UTILITA As Collection
    Dim coll_Modules_UTILITA As Collection
    Dim colForms_UTILITA As Collection
    Dim colReports_UTILITA As Collection
    Dim colMacros_UTILITA As Collection

    ' Carica gli oggetti nelle collection
    Call CaricaOggetti_UTILITA(coll_Tables_UTILITA, coll_Queryes_UTILITA, coll_Modules_UTILITA, colForms_UTILITA, colReports_UTILITA, colMacros_UTILITA)

    ' Apri il database corrente
    Set dbCurrent = CurrentDb

    ' Apri il database di origine
    Set dbSource = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath_s)




    ' --- INIZIALIZZA I CONTATORI ---
    importedTableCount_i = 0
    imported_QueryesCount_i = 0
    importedModuliCount_i = 0
    imported_FormsCount_i = 0
    importedReportsCount_i = 0
    imported_macros_Count_i = 0

    deleteTablesCount_i = 0
    delete_QueryesCount_i = 0
    deleteModulesCount_i = 0
    deleteFormsCount_i = 0
    deleteReportsCount_i = 0
    deletemacros_Count_i = 0

    deleteNameTables_s = ""
    importedNameTables_s = ""
    deleteName_Queryes_s = ""
    importedName_Queryes_s = ""
    deleteNameForms_s = ""
    importedNameForms_s = ""
    deleteNameReports_s = ""
    importedNameReports_s = ""
    deleteNameMacros_s = ""
    importedNameMacros_s = ""



    ' --- Collega le tabelle ---
    Dim tblName As String
    Dim rs As DAO.Recordset
    Dim TableExists As Boolean
    
    
    '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_ATTIVA_GEST_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE  @CONTROLLO@PROGETTO
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//

    
    
        

    For Each varItem In coll_Tables_UTILITA
        tblName = CStr(varItem) ' Assicura che sia una stringa
        Debug.Print "CONTROLLO TABELLA DA COLLEGARE --> : "; tblName
        Debug.Print

        ' Verifica se la tabella esiste nel database esterno
        TableExists = False
        Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = 1 AND Name = '" & tblName & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            TableExists = True
        End If
        rs.Close
        Set rs = Nothing

        If TableExists Then
            ' Elimina la tabella se esiste già nel database corrente
            On Error Resume Next
            dbCurrent.TableDefs.Delete tblName
            On Error GoTo 0

            ' Conteggio delle tabelle cancellate
            deleteTablesCount_i = deleteTablesCount_i + 1
            deleteNameTables_s = deleteNameTables_s & tblName & vbCrLf

            ' Collega la tabella al database corrente
            Set tdf = dbCurrent.CreateTableDef(tblName)
            tdf.Connect = ";DATABASE=" & sourceDBPath_s
            tdf.SourceTableName = tblName
            dbCurrent.TableDefs.Append tdf

            ' Conteggio delle tabelle collegate
            importedTableCount_i = importedTableCount_i + 1
            importedNameTables_s = importedNameTables_s & tblName & vbCrLf
        Else
            Debug.Print "Tabella NON trovata nel database esterno: "; tblName
        End If
    Next varItem

    ' -------------- Importazione delle query ---
    Dim qryName As String
    Dim QUERYExists As Boolean

    For Each varItem In coll_Queryes_UTILITA
        qryName = CStr(varItem) ' Assicura che sia una stringa
        Debug.Print "CONTROLLO QUERY DA IMPORTARE --> : "; qryName
        Debug.Print

        ' Verifica se la query esiste nel database esterno
        QUERYExists = False
        Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = 5 AND Name = '" & qryName & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            QUERYExists = True
        End If
        rs.Close
        Set rs = Nothing

        If QUERYExists Then
            ' Elimina la query se esiste già nel database corrente
            On Error Resume Next
            DoCmd.DeleteObject acQuery, qryName
            On Error GoTo 0

            ' Conteggio delle query cancellate
            delete_QueryesCount_i = delete_QueryesCount_i + 1
            deleteName_Queryes_s = deleteName_Queryes_s & qryName & vbCrLf

            ' Importa la query dal database di origine
            DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acQuery, qryName, qryName

            ' Conteggio delle query importate
            imported_QueryesCount_i = imported_QueryesCount_i + 1
            importedName_Queryes_s = importedName_Queryes_s & qryName & vbCrLf
        Else
            Debug.Print "Query NON trovata nel database esterno: "; qryName
        End If
    Next varItem

    ' -------------- Importazione dei moduli ---
    Dim myNameModule_s As String
    Dim moduleExists As Boolean

    For Each MDL In coll_Modules_UTILITA
        myNameModule_s = MDL
        Debug.Print "CONTROLLO MODULO DA IMPORTARE --> : "; myNameModule_s
        Debug.Print

        ' Verifica se il modulo esiste nel database esterno
        moduleExists = False
        Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = -32761 AND Name = '" & myNameModule_s & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            moduleExists = True
        End If
        rs.Close
        Set rs = Nothing

        If moduleExists Then
            ' Elimina il modulo se esiste già nel database corrente
            On Error Resume Next
            Application.VBE.VBProjects(1).VBComponents.Remove Application.VBE.VBProjects(1).VBComponents(myNameModule_s)
            On Error GoTo 0

            ' Conteggio dei moduli cancellati
            deleteModulesCount_i = deleteModulesCount_i + 1
            deleteNameModules_s = deleteNameModules_s & myNameModule_s & vbCrLf

            ' Importa il modulo dal database di origine
            DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acModule, myNameModule_s, myNameModule_s

            ' Conteggio dei moduli importati
            importedModuliCount_i = importedModuliCount_i + 1
            importedNameModules_s = importedNameModules_s & myNameModule_s & vbCrLf
        End If
    Next MDL

    ' -------------- Importazione dei form ---
    Dim formExists As Boolean

    For Each varItem In colForms_UTILITA
        Debug.Print "CONTROLLO FORM DA IMPORTARE --> : "; varItem
        Debug.Print

        ' Verifica se la form esiste nel database di origine
        formExists = False
        Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = 6 AND Name = '" & varItem & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            formExists = True
        End If
        rs.Close
        Set rs = Nothing
        
        '@ciclo@forms_(questo è il ciclo negli oggetti @form)
        ' Se la form esiste nel database di origine, esegui la cancellazione e l'importazione
        If formExists Then
            ' Verifica se la form esiste già nel database corrente
            On Error Resume Next
            ' Controlla nella raccolta AllForms se esiste la form corrente
            If CurrentProject.AllForms(varItem).IsLoaded Then
                ' Se la form è caricata, eliminiamola
                DoCmd.DeleteObject acForm, varItem
                ' Conteggio delle form cancellate
                deleteFormsCount_i = deleteFormsCount_i + 1
                deleteNameForms_s = deleteNameForms_s & varItem & vbCrLf
            End If
            On Error GoTo 0

            ' Importa la form dal database di origine
            DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acForm, varItem, varItem

            ' Conteggio delle form importate
            imported_FormsCount_i = imported_FormsCount_i + 1
            importedNameForms_s = importedNameForms_s & varItem & vbCrLf
        Else
            Debug.Print "Form NON trovata nel database di origine: "; varItem
        End If
    Next varItem

    ' -------------- Importazione dei report ---
    Dim reportExists As Boolean
    Dim reportExistsInDestination As Boolean

    For Each varItem In colReports_UTILITA
        Debug.Print "CONTROLLO REPORT DA IMPORTARE --> : "; varItem
        Debug.Print

        ' Verifica se il report esiste nel database di origine
        reportExists = False
        reportExistsInDestination = False
        Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = -32764 AND Name = '" & varItem & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            reportExists = True
        End If
        rs.Close
        Set rs = Nothing

        ' Verifica se il report esiste nel database di destinazione
        On Error Resume Next
        Set rs = dbCurrent.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = -32764 AND Name = '" & varItem & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            reportExistsInDestination = True
        End If
        rs.Close
        Set rs = Nothing
        On Error GoTo 0

        ' Se il report esiste nel database di origine, esegui l'importazione
        If reportExists Then
            ' Se il report esiste già nel database di destinazione, cancellalo prima di importare
            If reportExistsInDestination Then
                DoCmd.DeleteObject acReport, varItem
                Debug.Print "Report esistente cancellato: "; varItem
            End If

            ' Importa il report dal database di origine
            DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acReport, varItem, varItem

            ' Conteggio dei report importati
            importedReportsCount_i = importedReportsCount_i + 1
            importedNameReports_s = importedNameReports_s & varItem & vbCrLf
        Else
            Debug.Print "Report NON trovato nel database di origine: "; varItem
        End If
    Next varItem

    ' -------------- Importazione delle macro ---
    Dim MACROExists As Boolean
    Dim macroExistsInDestination As Boolean

    For Each varItem In colMacros_UTILITA
        Debug.Print "CONTROLLO MACRO DA IMPORTARE --> : "; varItem
        Debug.Print

        ' Verifica se la macro esiste nel database di origine
        MACROExists = False
        macroExistsInDestination = False
        Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = -32766 AND Name = '" & varItem & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            MACROExists = True
        End If
        rs.Close
        Set rs = Nothing

        ' Verifica se la macro esiste nel database di destinazione
        On Error Resume Next
        Set rs = dbCurrent.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = -32766 AND Name = '" & varItem & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            macroExistsInDestination = True
        End If
        rs.Close
        Set rs = Nothing
        On Error GoTo 0

        ' Se la macro esiste nel database di origine, esegui l'importazione
        If MACROExists Then
            ' Se la macro esiste già nel database di destinazione, cancellala prima di importare
            If macroExistsInDestination Then
                DoCmd.DeleteObject acMacro, varItem
                Debug.Print "Macro esistente cancellata: "; varItem
            End If

            ' Importa la macro dal database di origine
            DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath_s, acMacro, varItem, varItem

            ' Conteggio delle macro importate
            imported_macros_Count_i = imported_macros_Count_i + 1
            importedNameMacros_s = importedNameMacros_s & varItem & vbCrLf
        Else
            Debug.Print "Macro NON trovata nel database di origine: "; varItem
        End If
    Next varItem

    ' Messaggio finale con il numero di oggetti importati
    MsgBox "Importazione completata con successo!" & vbCrLf & vbCrLf & _
           "Tabelle collegate: " & importedTableCount_i & vbCrLf & _
           "Query importate: " & imported_QueryesCount_i & vbCrLf & _
           "Moduli importati: " & importedModuliCount_i & vbCrLf & _
           "Forms importati: " & imported_FormsCount_i & vbCrLf & _
           "Report importati: " & importedReportsCount_i & vbCrLf & _
           "Macro importate: " & imported_macros_Count_i, _
           vbInformation, "Operazione completata"

    ' Rilascia le risorse
    dbSource.Close
    Set dbSource = Nothing
    Set dbCurrent = Nothing

    Exit Sub

ErrHandler:
    ' Log dell'errore nel Debug Window
    Debug.Print "Errore in Cmd_IMPORTA_OGGETTI_UTILITA_Click"
    Debug.Print "Numero errore: " & Err.Number
    Debug.Print "Descrizione errore: " & Err.Description
    Debug.Print "Origine errore: " & Err.Source

    ' Messaggio di errore all'utente
    MsgBox "Si è verificato un errore durante l'importazione." & vbCrLf & _
           "Errore " & Err.Number & ": " & Err.Description, _
           vbCritical, "Errore"

    ' Rilascia le risorse in caso di errore
    If Not dbSource Is Nothing Then
        dbSource.Close
        Set dbSource = Nothing
    End If
    If Not dbCurrent Is Nothing Then
        Set dbCurrent = Nothing
    End If
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
End Sub

' Funzione per verificare se un elemento esiste nella collezione di MACRO
Function InCollectionMacrosDbIntern_DELETE_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionMacrosDbIntern_DELETE_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionMacrosDbIntern_DELETE_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionMacrosDbIntern_DELETE_Funct_b = False
    On Error GoTo 0
End Function





' Funzione per caricare gli oggetti nelle collection
Public Sub CaricaOggetti_UTILITA(ByRef coll_Tables_UTILITA As Collection, _
                          ByRef coll_Queryes_UTILITA As Collection, _
                          ByRef coll_Modules_UTILITA As Collection, _
                          ByRef colForms_UTILITA As Collection, _
                          ByRef colReports_UTILITA As Collection, _
                          ByRef colMacros_UTILITA As Collection)
    ' Svuota le collezioni prima di aggiungere nuovi elementi
    Set coll_Tables_UTILITA = New Collection
    Set coll_Queryes_UTILITA = New Collection
    Set coll_Modules_UTILITA = New Collection
    Set colForms_UTILITA = New Collection
    Set colReports_UTILITA = New Collection
    Set colMacros_UTILITA = New Collection
    
    '//@LE.TABELLE@UTILITA
    '//-------------------------------------------------------------------------------------------------------//
    '//@UTILITA@COLLECTION@TABLES_(LE @TABELLE  per la @GESTIONE DELLE @UTILITA@TABELLE)
    '//@UTILITA@COLLETION_(Aggiungi @abelle, @query, @moduli, @form, @report e @macro alle rispettive collezioni)
        
        coll_Tables_UTILITA.Add "MSys_DF13_}----------------------------------------------------@"
        coll_Tables_UTILITA.Add "Msys_DF13_DLL_PROGETTO"
        coll_Tables_UTILITA.Add "Msys_DF14_DLL_LIBRERIE"
        '//TABELLE PROGETTI
        coll_Tables_UTILITA.Add "PROGETTI_MSYS_{@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@}_Tab"
        coll_Tables_UTILITA.Add "PROGETTI_Msys_TB01_}-------------------------------------------@"
        coll_Tables_UTILITA.Add "PROGETTI_Msys_TB01_PROJECT"
        coll_Tables_UTILITA.Add "PROGETTI_Msys_TB01_PROJECT_TMP"
        coll_Tables_UTILITA.Add "PROGETTI_Msys_TB02_}-------------------------------------------@"
        coll_Tables_UTILITA.Add "PROGETTI_Msys_TB02_TIPO_OGGETTI"
        coll_Tables_UTILITA.Add "PROGETTI_Msys_TB03_}-------------------------------------------@"
        coll_Tables_UTILITA.Add "PROGETTI_Msys_TB03_ELENCO_ObjProgetti"

    '//-------------------------------------------------------------------------------------------------------//
    
    
    '//@LE.QUERY@UTILITA
    '//-------------------------------------------------------------------------------------------------------//
    '//@UTILITA@COLLECTION@QUERYES_(LE @QUERY  per la @GESTIONE DELLE @UTILITA@QUERY)
        coll_Queryes_UTILITA.Add "Msys_DF13Qry01_00_}-----------------------------------------@DLL"
        coll_Queryes_UTILITA.Add "Msys_DF13Qry01_01_DLL_PROGETTO"
        coll_Queryes_UTILITA.Add "Msys_DF14Qry01_02_DLL_LIBRERIE"
        coll_Queryes_UTILITA.Add "Msys_DF14Qry01_01_DELETE_LIBRERIE"
        coll_Queryes_UTILITA.Add "Msys_DF14Qry01_03_UPDATE_CAMPI"
    
        '//QUERY PROGETTI
        coll_Queryes_UTILITA.Add "MSysObjects_Qry00_}--------------------------------------------@"
        coll_Queryes_UTILITA.Add "MSysObjects_Qry01_01_GROUP_OGGETTI"
        coll_Queryes_UTILITA.Add "MSysObjects_Qry01_02_OGGETTI_TUTTI"
        coll_Queryes_UTILITA.Add "MSysObjects_Qry01_03_CAMPO_DATA"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB01_Qry00_{@===================================@}"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB01_Qry00_}-------------------------------------@"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB01Qry01_01_PROJECT"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB01Qry01_01_PROJECT_TMP"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB01Qry01_02_PROJECT_TMP_Delete"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB01Qry01_03_PROJECT_TMP_UPDATE"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB02Qry01_01_TIPO_OGGETTI"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB03Qry01_01_ELENCO_ObjProgetti"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB03Qry01_02_ObjProgetti_FORM=COD_PROGETTO_s"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB03Qry01_02_ObjProgetti_MACRO=COD_PROGETTO_s"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB03Qry01_02_ObjProgetti_QUERY=COD_PROGETTO_s"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB03Qry01_02_ObjProgetti_REPORT=COD_PROGETTO_s"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB03Qry01_02_ObjProgetti_TABLE=COD_PROGETTO_s"
        coll_Queryes_UTILITA.Add "PROGETTI_Msys_TB03Qry01_03_GROUP_QUOTE_s"

    
    
    '//-------------------------------------------------------------------------------------------------------//
    
    
    '//@I.MODULI@UTILITA
    '//-------------------------------------------------------------------------------------------------------//
    '//@UTILITA@COLLECTION@MODULES_(I MODULI  per la @GESTIONE DELLE @UTILITA@MODULI)
        coll_Modules_UTILITA.Add "UTILITA_MsysDF13_Mdl00_}-----------------------------------@"
        coll_Modules_UTILITA.Add "UTILITA_MsysDF13Mdl01_DLL_REFERENZIA_LE_LIBRERIE"
        coll_Modules_UTILITA.Add "UTILITA_MsysDF13Mdl01_DLL_REFERENZIA_LE_LIBRERIE"
        
        '//.... per la gestione delle @API@ WINDOWS
        coll_Modules_UTILITA.Add "GEFILE_Mdl01_GESTIONI_FILE_API_VBA"
        
        '//.... per la gestione della libreria di access agganciata alle utilita _
            qui viene descritto il tutorial di access e le tabelle di oggetti access da studiare ...
        coll_Modules_UTILITA.Add "TUTORIAL_Mdl00_}-----------------------------------------------@"
        coll_Modules_UTILITA.Add "TUTORIAL_Mdl01_UTILITA_ACCESS"
        
        '//MODULI PROGETTI
        coll_Modules_UTILITA.Add "CLASSE_ATTI_Tb01_Gestione"
        
        
    '//-------------------------------------------------------------------------------------------------------//
    
    
    
    '//@LE.FORM@UTILITA
    '//-------------------------------------------------------------------------------------------------------//
    '//@UTILITA@COLLECTION@FORMS_(LE @FORM  per la @GESTIONE DELLE @UTILITA@FORMS)
        
        colForms_UTILITA.Add "UTILITA_Msys_Frm21_}-------------------------------------------@"
    
      '// ... questa è la form per la gestione dei file windows a cui si unisce ache il modulo PER LE @API@WINDOWS
        colForms_UTILITA.Add "GEFILE_Frm01_{@===============================================@}"
        colForms_UTILITA.Add "GEFILE_Frm01_GESTIONI"
        '//la form progress bar per le importazione degli oggetti
        colForms_UTILITA.Add "frmProgressBa_}------------------------------------------------@"
        colForms_UTILITA.Add "UTILITA_Msys_Frm21_frmProgressBar"
        
        '//FORM PROGETTI
        colForms_UTILITA.Add "PROGETTI_Msys_{@==============GESTITI_DA_MSYS============@}_Tb00"
        colForms_UTILITA.Add "PROGETTI_Msys_Tb00FrmM01_GE_PROGETTI"
        colForms_UTILITA.Add "PROGETTI_Msys_TB01Frm00_}--------------------------------------@"
        colForms_UTILITA.Add "PROGETTI_Msys_TB01Frm01_01_PROJECT"
        colForms_UTILITA.Add "PROGETTI_Msys_TB01Frm01_01_PROJECT_TMP"
        colForms_UTILITA.Add "PROGETTI_Msys_TB02Frm00_}--------------------------------------@"
        colForms_UTILITA.Add "PROGETTI_Msys_TB02Frm01_01_TIPO_OGGETTI"
        colForms_UTILITA.Add "PROGETTI_Msys_TB03Frm01_01_ELENCO_ObjProgetti"
        colForms_UTILITA.Add "PROGETTI_Msys_TB03Frm01_02_ObjProgetti_FORM=COD_PROGETTO_s"
        colForms_UTILITA.Add "PROGETTI_Msys_TB03Frm01_02_ObjProgetti_MACRO=COD_PROGETTO_s"
        colForms_UTILITA.Add "PROGETTI_Msys_TB03Frm01_02_ObjProgetti_QUERY=COD_PROGETTO_s"
        colForms_UTILITA.Add "PROGETTI_Msys_TB03Frm01_02_ObjProgetti_REPORT=COD_PROGETTO_s"
        colForms_UTILITA.Add "PROGETTI_Msys_TB03Frm01_02_ObjProgetti_TABLE=COD_PROGETTO_s"
        colForms_UTILITA.Add "UTILITA_Msys_Frm21_}-------------------------------------------@"
        colForms_UTILITA.Add "UTILITA_Msys_Frm21_frmProgressBar"
        colForms_UTILITA.Add "UTILITA_Msys_Frm21_IMPORTA_OGGETTI"
        colForms_UTILITA.Add "UTILITA_Msys_Frm21_UTILITA_ALTRI_PROGETTI"

        
    '//-------------------------------------------------------------------------------------------------------//
    
       
    
    '//@UTILITA@COLLECTION@REPORTS_(I @REPORT  per la @GESTIONE DELLE @UTILITA@REPORT)
    '//-------------------------------------------------------------------------------------------------------//
        colReports_UTILITA.Add "Msys_Rpt01_01_ESPORTA_DB_}-------------------------------------@"
        colReports_UTILITA.Add "Msys_Rpt01_01_ESPORTA_DB_EST"
        
        colReports_UTILITA.Add "Msys_Rpt01_01_ESPORTA_OGGETT_}---------------------------------@"
        colReports_UTILITA.Add "Msys_Rpt01_01_ESPORTA_OGGETTI"
    '//-------------------------------------------------------------------------------------------------------//
    
    
    '//@UTILITA@COLLECTION@MACROS_(le @MACRO per la @GESTIONE DELLE @UTILITA@MACRO)
    '//-------------------------------------------------------------------------------------------------------//
        colMacros_UTILITA.Add "CERCA_Mcr_{@==================================================@}"
        colMacros_UTILITA.Add "CERCA_Mcr_FORM"
        colMacros_UTILITA.Add "CERCA_Mcr_MODULO"
        colMacros_UTILITA.Add "CERCA_Mcr_QUERY"
        colMacros_UTILITA.Add "CERCA_Mcr_TABELLA"
        colMacros_UTILITA.Add "UTILITA_MSys_Mcr00_}----------------------------------------@DLL"
        colMacros_UTILITA.Add "UTILITA_Msys_Mcr01_STAMPA_LIBRERIE_ATTIVE_DLL"
    '//-------------------------------------------------------------------------------------------------------//
    
    
End Sub

'//CANCELLA OGGETTI UTILITA *** FINE ***
'//==================================================================================================================//
 



'//CANCELLA OGGETTI UTILITA *** INIZIO***
'//==================================================================================================================//
 '//@UTILITA@DELETE@OGGETTI_(elimino tutti gli oggetti caricati)
 
Private Sub Cmd_DELETE_OGGETTI_UTILITA_Click()

    On Error GoTo ErrHandler ' Abilita la gestione degli errori

    ' --- INIZIO DEL CODICE DELLA PROCEDURA ---
    Dim dbCurrent As DAO.Database
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim tdf As DAO.TableDef
    Dim varItem As Variant
    Dim myName As String
    Dim myNameModule_s As String
    Dim myNameReport_s As String
    Dim myNameForm_s As String
    Dim myNameMACRO_s As String
    Dim moduleExists As Boolean
    Dim reportExists As Boolean
    Dim formExists As Boolean
    Dim MACROExists As Boolean

    ' Dichiarazione delle collezioni
    Dim colForms_UTILITA As Collection
    Dim colMacros_UTILITA As Collection
    Dim coll_Modules_UTILITA As Collection
    Dim coll_Queryes_UTILITA As Collection
    Dim colReports_UTILITA As Collection
    Dim coll_Tables_UTILITA As Collection

    ' Contatori
    Dim deleteTablesCount_i As Integer: deleteTablesCount_i = 0
    Dim delete_QueryesCount_i As Integer: delete_QueryesCount_i = 0
    Dim deleteModulesCount_i As Integer: deleteModulesCount_i = 0
    Dim deleteReportsCount_i As Integer: deleteReportsCount_i = 0
    Dim deletemacros_Count_i As Integer: deletemacros_Count_i = 0
    Dim deleteFormsCount_i As Integer: deleteFormsCount_i = 0
    
    
    
    
    
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_ATTIVA_GEST_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE  @CONTROLLO@PROGETTO
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//

    
    
        
    
    
    
    
    
    
    

    ' Apri il database corrente
    Set dbCurrent = CurrentDb

    ' Carica gli oggetti nelle collezioni
    Call CaricaOggetti_UTILITA(coll_Tables_UTILITA, coll_Queryes_UTILITA, coll_Modules_UTILITA, colForms_UTILITA, colReports_UTILITA, colMacros_UTILITA)

    ' --- Cancellazione delle Tabelle ---
    For Each varItem In coll_Tables_UTILITA
        On Error Resume Next
        Set rs = dbCurrent.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type = 6 AND Name = '" & varItem & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            dbCurrent.TableDefs.Delete varItem
            deleteTablesCount_i = deleteTablesCount_i + 1
        End If
        On Error GoTo 0
        If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    Next varItem

    ' --- Cancellazione delle Query ---
    For Each varItem In coll_Queryes_UTILITA
        On Error Resume Next
        Set qdf = dbCurrent.QueryDefs(varItem)
        If Not qdf Is Nothing Then
            dbCurrent.QueryDefs.Delete varItem
            delete_QueryesCount_i = delete_QueryesCount_i + 1
        End If
        On Error GoTo 0
        Set qdf = Nothing
    Next varItem

    ' --- Cancellazione dei Moduli ---
    For Each varItem In coll_Modules_UTILITA
        myNameModule_s = varItem
        moduleExists = Not IsNull(DLookup("Name", "MSysObjects", "Name = '" & myNameModule_s & "' AND Type = -32761"))
        If moduleExists Then
            On Error Resume Next
            Application.VBE.VBProjects(1).VBComponents.Remove Application.VBE.VBProjects(1).VBComponents(myNameModule_s)
            deleteModulesCount_i = deleteModulesCount_i + 1
            On Error GoTo 0
        End If
    Next varItem

    ' --- Cancellazione dei Report ---
    For Each varItem In colReports_UTILITA
        myNameReport_s = varItem
        reportExists = Not IsNull(DLookup("Name", "MSysObjects", "Name = '" & myNameReport_s & "' AND Type = -32764"))
        If reportExists Then
            On Error Resume Next
            DoCmd.DeleteObject acReport, myNameReport_s
            deleteReportsCount_i = deleteReportsCount_i + 1
            On Error GoTo 0
        End If
    Next varItem

    ' --- Cancellazione delle Form ---
    For Each varItem In colForms_UTILITA
        myNameForm_s = varItem
        formExists = Not IsNull(DLookup("Name", "MSysObjects", "Name = '" & myNameForm_s & "' AND Type = -32768"))
        If formExists Then
            On Error Resume Next
            DoCmd.DeleteObject acForm, myNameForm_s
            deleteFormsCount_i = deleteFormsCount_i + 1
            On Error GoTo 0
        End If
    Next varItem

    ' --- Cancellazione delle Macro ---
    For Each varItem In colMacros_UTILITA
        myNameMACRO_s = varItem
        MACROExists = Not IsNull(DLookup("Name", "MSysObjects", "Name = '" & myNameMACRO_s & "' AND Type = -32766"))
        If MACROExists Then
            On Error Resume Next
            DoCmd.DeleteObject acMacro, myNameMACRO_s
            deletemacros_Count_i = deletemacros_Count_i + 1
            On Error GoTo 0
        End If
    Next varItem

Cleanup:
    ' Rilascio della memoria
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    Set dbCurrent = Nothing

    ' Messaggio finale
    MsgBox "Cancellazione completata con successo!" & vbCrLf & vbCrLf & _
           "Tabelle cancellate: " & deleteTablesCount_i & vbCrLf & _
           "Query cancellate: " & delete_QueryesCount_i & vbCrLf & _
           "Moduli cancellati: " & deleteModulesCount_i & vbCrLf & _
           "Report cancellati: " & deleteReportsCount_i & vbCrLf & _
           "Macro cancellate: " & deletemacros_Count_i & vbCrLf & _
           "Form cancellate: " & deleteFormsCount_i, vbInformation, "Operazione completata"

    Exit Sub

ErrHandler:
    MsgBox "Errore " & Err.Number & ": " & Err.Description, vbCritical, "Errore"
    Resume Cleanup

End Sub

'//CANCELLA OGGETTI UTILITA *** INIZIO***
'//==================================================================================================================//





'//***************************************************************************************************************************************//
'//*
'//*                              IMPORTA OGGETTI FORM DELLE UTILITA    *** FINE ***
'//*                              @FORM@UTILITA_(LA GESTIONE DEGLI OGGETTI DELLA FORM UTILITA @IMPORTA@OGGETTI)
'//*
'//***************************************************************************************************************************************//















'//=========================================================================================//
'//
'//                 GESTIONE DELLE REPORT ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@REPORT                   @GESTONE@REPORT
'//=========================================================================================//

 

 
 
'//==========================================================================================================//
'// IMPORTA GLI OGGETTI REPORT *** INIZIO *** '@IMPORT@REPORT
'//==========================================================================================================//

Private Sub Cmd_Importa_REPORT_Click()
    Dim obj As AccessObject
    Dim Dbs As Object
    Dim ProceduraMessaggioErrore_s As String
    Dim reports_Collection As Collection
    Dim reportExists As Boolean
    Dim sourceDBPath As String
    Dim externalDB As DAO.Database
    Dim rs As DAO.Recordset

    On Error GoTo Err_Cmd_Import_REPORT_Click

    ' Reset counts and strings
    deletet_ReportesCount_i = 0
    importedReportesCount_i = 0
    imported_Reports_s = ""

    '//=====================================================================================================================//
    '//                             STEP_02 = @CONTROLLO@PROGETTO
    '// CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
    '//=====================================================================================================================//
    '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
    '// non si possono caricare gli oggetti nel PROGETTO MSYS_ATTIVA_GEST_OGGETTI.mdb perche è originale _
    '// e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
    '// per caricare gli oggetti table, query , form ecc. da importare.
    '// 01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE
    '// 02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE E QUINDI ESCLUSIONE.

    ' 01) PRIMO CONTROLLO _ MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
    '//...........................................................................//
    ' Attenzione: L'AVVISO produce il messaggio solo se il parametro è TRUE, altrimenti il messaggio viene escluso.
    MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)

    If MyValue = 0 Then Exit Sub

    ' 02) SECONDO CONTROLLO
    ' CHIAMO IL CONTROLLO DEL PROGETTO
    '//...........................................................................//
    ' Note: True = siamo nel progetto originale non possiamo importare nessun oggetto perché distruggiamo il db corrente.
    ' False = siamo in un progetto esterno possiamo caricare tutto.
    Bool1 = CheckProjectName_b_PFunct

    If Bool1 = True Then Exit Sub

    '//=====================================================================================================================//
    '//                             STEP_02 = @CONTROLLO@PROGETTO
    '// CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
    '//=====================================================================================================================//

    ' ATTIVA LA COLLEZIONE DEI REPORT
    '------------------------------------------------------------------------//

    ' Recupero la collezione di report da importare
    Set reports_Collection = CollectionReport_PFunct()

    ' 2024.08.03_CollectionReport_PFunct = SE LA COLLECTION E' VUOTA ESCI DALLA ROUTINE
    ' Faccio un controllo sul numero degli oggetti della collection. Se = 0, allora non ci sono oggetti da caricare.
    If reports_Collection.count = 0 Then GoTo Exit_Cmd_Import_REPORT_Click

    '------------------------------------------------------------------------//

    ' Recupero il percorso e il nome del file del database esterno
    sourceDBPath = PathFile_s_pFunct

    If sourceDBPath = "" Then
        MsgBox "ATTENZIONE ERRORE: il percorso del file del database esterno non è stato specificato. Uscita dalla routine.", vbCritical, "Errore: Percorso del File Non Specificato"
        Exit Sub
    End If

    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "IMPORTO I REPORT DA QUESTO DATABASE; Percorso del file: " & sourceDBPath, vbExclamation, "Informazione: Chiamata Funzione PathFile"
    End If

    ' Apro il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath)

    ' Ciclo nella collezione per il controllo dei report precaricati
    For Each reportName In reports_Collection
        Debug.Print "Report nella collezione: " & myReports_Name_s

        ' Verifica se il report esiste nel database corrente
        reportExists = False
        Set Dbs = Application.CurrentProject
        For Each obj In Dbs.AllReports
            If obj.Name = reportName Then
                reportExists = True
                Exit For
            End If
        Next obj

        If reportExists Then
            ' Se il report esiste nel database corrente, cancellalo
            DoCmd.DeleteObject acReport, reportName
            deletet_ReportesCount_i = deletet_ReportesCount_i + 1
            Debug.Print "Report eliminato: " & reportName
        End If

        ' CONTROLLO SULLA TABELLA DI SISTEMA DEL DB ESTERNO MSysObjects
        '----------------------------------------------------------------------------------------------------------------------//
        ' Verifica se il report esiste nel database esterno utilizzando una query SQL
        ' La riga SELECT Name FROM MSysObjects WHERE Type=-32764 è una query SQL utilizzata per selezionare i
        ' report dal database esterno.

        reportExists = False
        Set rs = externalDB.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=-32764 AND Name='" & reportName & "'")
        If Not rs.EOF Then
            reportExists = True
        End If
        rs.Close

        If reportExists Then
            ' Se il report esiste nel database esterno, importalo
            DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acReport, reportName, reportName
            importedReportesCount_i = importedReportesCount_i + 1
            imported_Reports_s = imported_Reports_s & reportName & vbCrLf
            Debug.Print "Report importato: " & reportName
        Else
            ' Report non trovato nel database esterno
            Debug.Print "Report non trovato nel database esterno: " & reportName
        End If
    Next reportName

    ' Chiudo il database esterno
    externalDB.Close

    '----------------------------------------------------------------------------------------------------------------------//

    ' Messaggio di riepilogo finale
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO REPORT PRECARICATI DA IMPORTARE:" & vbCrLf & _
               "Report importati: " & importedReportesCount_i & vbCrLf & imported_Reports_s & _
               "Report esistenti cancellati: " & deletet_ReportesCount_i, vbInformation
    End If

    ' Ciclo nella collezione per il controllo delle QUERY precaricate per SELEZIONARE LA PRIMA
    '--------------------------------------------------------------------------------------//
    For Each obj In Dbs.AllReports
        ' Vado all'ultima query
        DoCmd.SelectObject acReport, obj.Name, True
    Next obj
    '--------------------------------------------------------------------------------------//

Exit_Cmd_Import_REPORT_Click:
    Exit Sub

Err_Cmd_Import_REPORT_Click:
    MsgBox Err.Description & " - Errore: " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Stop
    Resume Exit_Cmd_Import_REPORT_Click

End Sub


'//FUNZIONE DI CONTROLLO DEGLI OGGETTI REPORT DEL DB MSYSDBEST *** INIZIO ***
'//----------------------------------------------------------------------------------------//

'//FUNZIONE DI CONTROLLO SE ESISTE L'OGGETTO = todo: modificata
Function ObjectReportExists(db As DAO.Database, objType As Integer, objName As String) As Boolean
    'Dim obj As DAO.Document
    'On Error Resume Next
    'Set obj = db.Containers("Reports").Documents(objName)
    'ObjectReportExists = (Err.Number = 0)
    'Set obj = Nothing
    'On Error GoTo 0
    
    Dim obj As AccessObject
    On Error Resume Next
    For Each obj In CurrentProject.AllReports
        Debug.Print "report corrente controllato :" & obj.Name
        If obj.Name = objName Then
            ReportExistsInCurrentDb = True
            Exit Function
        End If
    Next obj
    ReportExistsInCurrentDb = False
    On Error GoTo 0
    
    
    
End Function


'//FUNZIONE DI CONTROLLO DEGLI OGGETTI REPORT DEL DB MSYSDBEST *** FINE ***
'//----------------------------------------------------------------------------------------//

'//==========================================================================================================//
'// IMPORTA GLI OGGETTI REPORT *** FINE *** '@IMPORT@REPORT
'//==========================================================================================================//




'//=====================================================================================================================//
'//                     DELETE OGGETTI REPROT DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE REPORT HELP
'//=====================================================================================================================//



'//CANCELLA GLI OGGETTI REPORT *** INIZIO *** DELETE_REPORT
'//======================================================================================================//
'//@GESTIONE@DEI@REPORT    (@MODELLO@DELETE@OGGETTI)

'//NOTA: visto che la cancellazione riguarda il database corrente in cui è stata importata _
        la classe di IMPORTA PROGETTO, la routine cancella solo se i REPORT esistono _
        nel database corrente mediante il confronto tra : _
        LA COLLECTION REPORT = GLI OGGETTI REPORT che sono attivi nel database ESTERNO. _
        Prima di cancellare viene effettuato il confronto di esistenza _
        @cancellazione@report@caricati_(cancellazione dei REPORT già caricati nel db esterno) _
        @cancella@TABLE@esterne_(attenzione puoi cancellare i REPORT nel @db@esterno)

'//FAQ : _
        DOVE SI CANCELLANO I REPORT PRECARICATI _
        COME CANCELLARE I REPORT ESTERNI _
        I REPORT ESTERNI DA CANCELLARE

'// PARAMETRI DA CAMBIARE : _
    -----> REPORT _
    -----> DEI@REPORT _
    -----> DELETE_REPORT _
    -----> CollectionReport_PFunct _
    -----> reports_Collection _
    -----> objectReport _
    -----> MATRICE_REPORT_Names_s _
    -----> reportsCount_i _
    -----> ReportDef _
    -----> rpt _

Private Sub Cmd_Delete_REPORT_Click()

    Dim dbCurrent As DAO.Database
    Dim rpt As Report
    Dim reports_Collection As Collection
    Dim varItem As Variant
    Dim MATRICE_REPORT_Names_s() As String             'matrice di REPORT
    Dim tempCount As Integer
    Dim reportsCount_i As Integer
    Dim rdf As AccessObject
    Dim iCount As Integer
    Dim i As Integer
    Dim Bool1 As Boolean
    
    On Error GoTo Err_Cmd_Delete_REPORT_Click

    '//RESET
    iCount = 0
    deletetReportsCount_i = 0
    deleted_Reports_s = ""
    
    
    
    
    
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_ATTIVA_GEST_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE  @CONTROLLO@PROGETTO
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//
     

    '//ATTIVO LA COLLEZIONE
    '//------------------------------------------------------------------------//
    '//Note: CHIAMO LA FUNZIONE COLLECTION REPORT per ottenere la collezione di REPORT

    Set reports_Collection = CollectionReport_PFunct()
    '//------------------------------------------------------------------------//

    ' Loop attraverso tutti i REPORT nel database corrente
    Set dbCurrent = CurrentDb

    '//CALCOLO INDICE REPORT
    iCount = Application.CurrentProject.AllReports.count

    '//se ci sono REPORT da cancellare
    If iCount > 0 Then

        '//reimposto la matrice sui REPORT esistenti nel db corrente
        ReDim MATRICE_REPORT_Names_s(iCount - 1)

        ' Costruisci un array con i nomi dei REPORT
        tempCount = 0

        '//CARICO LA MATRICE: ciclo for per caricare MATRICE_REPORT_Names_s(tempCount)
        For i = 0 To iCount - 1
            Set rdf = Application.CurrentProject.AllReports(i)
            MATRICE_REPORT_Names_s(tempCount) = rdf.Name
            tempCount = tempCount + 1
        Next i

        ' Cancella i REPORT usando l'array dei nomi
        For i = 0 To tempCount - 1
            On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
            Debug.Print "Cancellazione REPORT: " & MATRICE_REPORT_Names_s(i)

            '//CONTROLLO MATRICE = COLLECTION
            Bool1 = InCollectionREPORTes_Funct_b(reports_Collection, MATRICE_REPORT_Names_s(i))

            '//cancello solo se la matrice (i) esiste nella collection
            If Bool1 = True Then
                Debug.Print
                Debug.Print "cancello il REPORT trovato sia nella collezione che nella matrice : " & MATRICE_REPORT_Names_s(i)
                DoCmd.DeleteObject acReport, MATRICE_REPORT_Names_s(i)

                If Err.Number <> 0 Then
                    Debug.Print "Errore durante la cancellazione del REPORT: " & MATRICE_REPORT_Names_s(i) & " - " & Err.Description
                    Err.Clear  ' Pulisci l'errore per la prossima iterazione
                Else
                    deletetReportsCount_i = deletetReportsCount_i + 1
                End If

            End If '//If Bool1 = True Then

            On Error GoTo 0  ' Reimposta gestione errori
        Next i
    End If
              
              
              
              '//APRI RIQUADRO MODULI INDIPENDENTEMENTE DALLA LORO PRESENZA
              '//-------------------------------------------------------------------------------------------------//
              '//NOTE= Apri la sezione macro indipendentemente dalla presenza di oggetti _
                  Conta solo i moduli visibili, escludendo quelli di sistema. _
                  Se ci sono moduli, apre direttamente la sezione Moduli con DoCmd.SelectObject acModule, , True. _
                  Print Se; non; ci; sono; Moduli, prova; a; selezionare; una; tabella; e; poi; la; sezione; Moduli, per; forzare; l 'apertura del riquadro. _
                  Evita errori con On Error Resume Next in caso di assenza di oggetti.

                          Dim db As DAO.Database
                          Dim doc As Document
                          Dim reportVisibili As Integer
          
                          Set db = CurrentDb
                          reportVisibili = 0 ' Inizializza il contatore
          
                          ' Conta solo i report visibili nel database
                          For Each doc In db.Containers("Reports").Documents
                            If Left(doc.Name, 1) <> "~" Then ' Esclude eventuali report di sistema nascosti
                              reportVisibili = reportVisibili + 1
                            End If
          
                          Next doc
          
                          ' Se ci sono report visibili, apri la sezione Report
                          If reportVisibili > 0 Then
                            DoCmd.SelectObject acReport, , True
                          Else
                            ' Se non ci sono report, prova a selezionare un altro oggetto per forzare l'apertura del riquadro
                            On Error Resume Next
                            DoCmd.SelectObject acTable, , True ' Seleziona una tabella (se esiste)
                            DoCmd.SelectObject acReport, , True ' Ora prova a selezionare la sezione Report
                            On Error GoTo 0
          
                          End If
          
                          ' Pulizia memoria
                          Set doc = Nothing
                          Set db = Nothing
               
                                    
                    
              '//APRI RIQUADRO MODULI INDIPENDENTEMENTE DALLA LORO PRESENZA *** FINE ***
              '//-------------------------------------------------------------------------------------------------//
            

                
    ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "I REPORT CANCELLATI DA QUESTO DATABASE SONO :  " & deletetReportsCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
    End If

'//USCITA E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_REPORT_Click:
    Exit Sub

Err_Cmd_Delete_REPORT_Click:
    Set dbCurrent = Nothing
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_REPORT_Click

End Sub

' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionREPORTes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionREPORTes_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionREPORTes_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionREPORTes_Funct_b = False
    On Error GoTo 0
End Function


'//CANCELLA GLI OGGETTI REPORT *** FINE *** DELETE_REPORT
'//=========================================================================================================

'//DELETE REPORT HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_Delete_REPORT_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "Help_Cmd_DELETE_REPORT.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa REPORT _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub







Private Sub Comando98_Click()

End Sub

'//DELETE REPORT HELP *** FINE ***
'//-----------------------------------------------------------------------------------//




'//=====================================================================================================================//
'//                     DELETE OGGETTI REPROT DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE REPORT HELP
'//=====================================================================================================================//









'//=========================================================================================//
'//
'//                 GESTIONE DELLE REPORT ESTERNE O INTERNE   *** FINE ***
'//@GESTIONE@DELLE@REPORT                   @GESTONE@REPORT
'//=========================================================================================//









'//**************************************************************************************//
'//                     EVENTI FORM  @GLI@EVENTI@FORM@LOAD
'//**************************************************************************************//


'//funzione di ausilio alla form Load PER GARANTIRE L'IMPORTAZIONE DEGLI OGGETTI DI SERVIZIO.
Private Function OggettoAccess(nome As String) As AcObjectType
    If InStr(1, nome, "Qry", vbTextCompare) > 0 Then
        OggettoAccess = acQuery
    ElseIf InStr(1, nome, "Frm", vbTextCompare) > 0 Then
        OggettoAccess = acForm
    ElseIf InStr(1, nome, "Mdl", vbTextCompare) > 0 Then
        OggettoAccess = acModule
    Else
        OggettoAccess = acQuery ' fallback generico se non riconosciuto
    End If
End Function

Private Sub Form_Load()
    On Error GoTo GestioneErrore

    '//disabilito i butto
    Call DisabilitaButton_Sub
    Call Disabilito_sourceDBPath_Sub
    
    
    Me.LBL_TIT01_TXT.Caption = "CLASSE AUTONOMA PER IMPORTARE OGGETTI DAL  PROGETTO MSYSDbEst_GE_OBJECT.mdb"
    

    '//@avviso
    '//imposto il flag avviso ad abilitato per visualizzare il messaggio di proseguimento
    Me.ctr_AVVISO_ABILITATO_TXT = True
    
    '//@progetto@corrente @modificato = inserisco subito nella casella di testo ed evito _
        che il progetto corrente vada a finire nella casella combinata.
    Me.CurrentProject_txt = Me.Application.CurrentProject.Name
    
     
     '//@CMB_01 - _
        @STEP@01_IMPORTA@OGGETTI_(reset la Cmb_01 NOME PROGETTO)          = Reset di default la Cmb_01 con il NOME DEL PROGETTO
     '//-------------------------------------------------------------------------//
     '//@cmb@imposto_(i valore della @casella@combinata)
     '//IMPOSTO LA CASELLA COMBINATA IN PARTENZA @cmb@abilitazioni_(vari @codici dei @progetti)
        Dim Str1 As String
        Str1 = ""
        Str1 = Str1 & ";"
        Str1 = Str1 & "ELENCO PROGETTI;"
        Str1 = Str1 & "MSYS_ATTIVA_GEST_OGGETTI.mdb;"
        Str1 = Str1 & "GE_ESPROPRI.mdb;"
        Str1 = Str1 & "TREEVIEW.mdb;"
        Str1 = Str1 & "CDM;"
        Str1 = Str1 & "IMPEGNI;"
        Str1 = Str1 & "PRES3000;"
        Str1 = Str1 & ""
        
        Me.Cmb_01_txt.Value = "ELENCO PROGETTI"
        Me.Cmb_01_txt.RowSourceType = "elenco valori"
        Me.Cmb_01_txt.RowSource = Str1
     '//-------------------------------------------------------------------------//
    
    
    '//@CMB_02 _
        @STEP@01_IMPORTA@OGGETTI_(reset la Cmb_02)                        = Reset di default la Cmb_02
    '//-------------------------------------------------------------------------//
    '//@cmb@imposto_(i valore della @casella@combinata)
    '//IMPOSTO LA CASELLA COMBINATA IN PARTENZA @cmb@abilitazioni_(vari @codici dei @progetti) _
        @STEP@02_IMPORTA@OGGETTI_(Imposto la Cmb_02)                       = Imposto la voce del progetto
    
        Str1 = ""
        Me.Cmb_02_txt.Value = ""
        Str1 = Str1 & "ELENCO OGGETTI MSYS;"
        Str1 = Str1 & "IMPORTA_OGGETTI_MSYS_DLL;"                       '//@IMPORTA@OGGETTI@DLL
        Str1 = Str1 & "IMPORTA_OGGETTI_TREEVIEW;"                       '//@IMPORTA@OGGETTI@TREEVIEW
        Str1 = Str1 & "IMPORTA_IMPEGNI;"                                '//@IMPORTA@OGGETTI@IMPEGNI
        Str1 = Str1 & "IMPORTA_PRES3000;"                               '//@IMPORTA@OGGETTI@PRES3000
        Str1 = Str1 & ""
    
        Me.Cmb_02_txt.RowSourceType = "elenco valori"
        Me.Cmb_02_txt.RowSource = Str1
    '//-------------------------------------------------------------------------//
    
         
    '//@RESET@COLORI@TXT_(IMPOSTO I COLORI BASE DELLE CASELLE DI TESTO)
    '//-------------------------------------------------------------------------//
    '//TODO: da sistemare per tutte le caselle di testo.
        '//@IMPOSTA@COLORE.02=CHIAMA LE IMPOSTAZIONI
        Imposta_Colore_Sub
                
    '//-------------------------------------------------------------------------//
    
    '//             IMPORTAZIONE DEGLI OGGETTI DI FUNZIONAMENTO
    '//*************************************************************************//
        '//IMPORTAZIONE OGGETTI DAL DB ESTERNO SE NON CI TROVIAMO NEL DB ORIGINE
            Dim dbCorrente As String
            dbCorrente = Application.CurrentProject.Name
        
            If LCase(dbCorrente) <> "msys_attiva_gest_oggetti.mdb" Then
                Dim percorsoOrigine As String
                
                '//ASSEGNO IL PERCORSO ED EFFETTUO IL CONTROLLO: che il file MSYS_ATTIVA_GEST_OGGETTI.mdb esista prima di aprirlo
                '//...........................................................................//
                    percorsoOrigine = "C:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\MSYS_ATTIVA_GEST_OGGETTI.mdb"
                    
                    If Dir(percorsoOrigine) = "" Then
                            MsgBox "Il database di origine non è stato trovato: " & percorsoOrigine, vbCritical
                        Exit Sub
                    End If
                    
                '//...........................................................................//
                
        
                Dim colOggetti As New Collection
                Dim obj As Variant
        
                '//AGGIUNGO GLI OGGETTI DA IMPORTARE NELLA COLLECTION
        
                ' Moduli
                colOggetti.Add "GEFILE_Mdl00_}"
                colOggetti.Add "GEFILE_Mdl01_GESTIONI_FILE_API_VBA"
                
                ' Form
                colOggetti.Add "UTILITA_Msys_Frm21_}"
                colOggetti.Add "UTILITA_Msys_Frm21_frmProgressBar"
                colOggetti.Add "UTILITA_Msys_Frm21_IMPORTA_OGGETTI"
                colOggetti.Add "UTILITA_Msys_Frm21_UTILITA_ALTRI_PROGETTI"
                
                ' Query
                colOggetti.Add "UTILITA_Msys_Qry21_00_{@======================================@}"
                colOggetti.Add "UTILITA_Msys_Qry21_01_GRUPPO_OGGETTI_CORRENTI"
                colOggetti.Add "UTILITA_Msys_Qry21_02_Oggetti_Tutti"
                colOggetti.Add "UTILITA_Msys_Qry21_03_Oggetti_Form"
                colOggetti.Add "UTILITA_Msys_Qry21_04_Oggetti_Macro"
                colOggetti.Add "UTILITA_Msys_Qry21_05_Oggetti_Report"
                colOggetti.Add "UTILITA_Msys_Qry21_06_Oggetti_Moduli"
                colOggetti.Add "UTILITA_Msys_Qry21_07_Oggetti_TABELLE"
                colOggetti.Add "UTILITA_Msys_Qry21_08_Oggetti_Query"
        
                '//CONTROLLO ESISTENZA NEL DB ORIGINE PRIMA DI IMPORTARE
                Dim dbOrigine As DAO.Database
                Set dbOrigine = DBEngine.Workspaces(0).OpenDatabase(percorsoOrigine)
        
                Dim rs As DAO.Recordset
                Dim nomeOggetto As String
                Dim tipoOggetto As Integer
                Dim sqlCheck As String
        
                For Each obj In colOggetti
                    nomeOggetto = obj
        
                    sqlCheck = "SELECT Name FROM MSysObjects WHERE Name='" & Replace(nomeOggetto, "'", "''") & "'"
                    Set rs = dbOrigine.OpenRecordset(sqlCheck, dbOpenSnapshot)
        
                    If Not rs.EOF Then
                        '//SE L'OGGETTO ESISTE NEL DB ORIGINE, IMPORTO (eventualmente sovrascrivo)
                        DoCmd.TransferDatabase acImport, "Microsoft Access", percorsoOrigine, _
                            OggettoAccess(nomeOggetto), nomeOggetto, nomeOggetto, True
                    Else
                        Debug.Print "Oggetto non trovato nel DB origine: " & nomeOggetto
                    End If
        
                    rs.Close
                    Set rs = Nothing
                Next obj
        
                dbOrigine.Close
                Set dbOrigine = Nothing
            Else
                Debug.Print "Importazione non eseguita: siamo nel DB di origine"
            End If

    '//************************************************************************//
    
    
    
                    
    Exit Sub

GestioneErrore:
    MsgBox "Si è verificato un errore durante il caricamento del form: " & vbCrLf & _
           "Descrizione: " & Err.Description & vbCrLf & _
           "Numero: " & Err.Number, vbCritical, "Errore Form_Load"
End Sub





'//**************************************************************************************//
'//                     EVENTI FORM  *** FINE ***
'//**************************************************************************************//




'//=========================================================================================//
'//
'//                 GESTIONE DELLE MACRO ESTERNE O INTERNE   *** INIZIO ***
'//@GESTIONE@DELLE@MACRO                    @GESTONE@MACRO
'//=========================================================================================//


'//IMPORTA GLI OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//----------------------------------------------------------------------------------------//
'//NOTE:   dal database di origine importa solo gli oggetti MACRO controllando prima se _
            nel db corrente di destinazione già esiste la MACRO per effettuare la cancellazione _
            prima dell 'importazione.

'//IMPORTA GLI OGGETTI MACRO DEL DB ESTERNO IN TABELLA *** FINE ***
'//----------------------------------------------------------------------------------------//

'//=========================================================================================//
'//
'//                 GESTIONE DELLE MACRO ESTERNE O INTERNE   *** FINE ***
'//@GESTIONE@DELLE@MACRO                    @GESTONE@MACRO
'//=========================================================================================//









'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI QUERY DEL DB ESTERNO *** INIZIO ***
'//                     @IMPORTA@OGGETTI@QUERY@MSYS
'//=====================================================================================================================//



'//IMPORTA TUTTI GLI OGGETTI QUERY DEL DB MSYSDBEST - LE ATTIVITA
'//----------------------------------------------------------------------------------------//
'//@01_@IMPORTA@QUERY@DIM                  -------------------> DIM VARIABILI
'//@02_@IMPORTA@QUERY@REST                 -------------------> RESET DELLE VARIABILI
'//@03_@IMPORTA@QUERY@LOG                  -------------------> LOG FILE INIZIALE
'//@04_@IMPORTA@QUERY@COLL                 -------------------> RECUPERO COLLECTION DA IMPORTARE
'//@05_@IMPORTA@@QUERYCONTROLLO 1          -------------------> CONTROLLO 1 ESISTONO OGGETTI COLLECTION
'//@06_@IMPORTA@QUERY@CONTROLLO 2          -------------------> CONTROLLO 2 ESISTE PATH VALIDA
'//@07_@IMPORTA@QUERY@CONTROLLO 3          -------------------> CONTROLLO 3 PATH NULL EXIT
'//@08_@IMPORTA@QUERY@CONTROLLO 4          -------------------> CONTROLLO 4 SE DB NULL EXIT
'//@09_@IMPORTA@QUERY@AVVISI+MSG           -------------------> AVVISI E IMPOSTAZIONE MESSAGGIO FINALE
'//@10_@IMPORTA@QUERY@PROGRSESSBAR_01      -------------------> PROGRESS BAR STEP 01 RECUPERO PARAMETRI
'//@11_@IMPORTA@QUERY@CONTROLLO_PROG       -------------------> CONTROLLO SE STIAMO NEL PROGETTO ESTERNO
'//@12_@IMPORTA@QUERY@PROG@CICLO           -------------------> NEL PROGETTO ESTERNO CICLO NEGLI OGGETTI
'//@13_@IMPORTA@QUERY@PROG@CONTR_01        -------------------> CONTROLLO 1 SE ESTE OGGETTO DA IMPORTARE
'//@14_@IMPORTA@QUERY@PROG@IMPORT          -------------------> IMPORTAZIONE DELL'OGGETTO DAL DB ESTERNO
'//@15_@IMPORTA@QUERY@PROGRSESSBAR_02      -------------------> PROGRESS BAR STEP 02 CALCOLO ED INVIO PARAMETRI
'//@15.BIS_@IMPORTA@QUERY@PROGRESSBAR_CALL -------------------> VIENE CHIAMATA LA PROGRESS BAR CON I PARAMETRI
'//@16_@IMPORTA@QUERY@LOG@FILE             -------------------> SE IL FILE NON VIENE TROVATO SCRITTURAZIONE LOG
'//@17_@IMPORTA@QUERY@CLOSE                -------------------> CHIUSURA CON RITARDO DELLA PROGRESS BAR
'//@18_@IMPORTA@QUERY@NOTHING              -------------------> IMPOSTO A NULL TUTTI GLI OGGETTI


Private Sub Cmd_Importa_QUERY_Click()
    
    '//@01_@IMPORTA@QUERY@DIM                  -------------------> DIM VARIABILI
    ' Dichiarazione delle variabili
    '//oggetti QUERY e variabile nome QUERY
    Dim Queries_Collection As Collection                       '//oggetti collezione
    Dim Queries_Name_v As Variant
    Dim deleted_Queries_s As String
    Dim imported_Queries_s As String
    Dim myQueries_Name_s As String
    
    '//il database + path
    Dim sourceDB As DAO.Database
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim sourceDBPath As String
    
    '//contatori oggetti: TOTALE OGGETTI PREVISTI; da importare e cancellare
    Dim TotCollection_Queries_i As Integer                          '//totale degli oggetti QUERY della collezione da importare
    Dim importedCount_i As Integer                              '//oggetti importati
    Dim deletedCount_i As Integer                               '//oggetti cancellati
    
    
    '//bool e log file
    Dim Bool1 As Boolean
    Dim MyValue As Integer
    Dim logFile As Integer
    
    On Error GoTo ErrorHandler
    
    '//@02_@IMPORTA@QUERY@REST                 -------------------> RESET DELLE VARIABILI
    '//RESET = Inizializza i contatori e le stringhe di riepilogo
    TotCollection_Queries_i = 0
    importedCount_i = 0
    deletedCount_i = 0
    deleted_Queries_s = ""
    imported_Queries_s = ""
    TotPerc_dbl = 0
    ProgressPercentTwips_dbl = 0
    
    '//@03_@IMPORTA@QUERY@LOG                  -------------------> LOG FILE INIZIALE
    '@FILE@LOG@QUERY: Apre il file di log per scrivere i dettagli dell'importazione apre nella path corrente.
    logFile = FreeFile()
    Open CurrentProject.PATH & "\LOG_IMPORTA_QUERY.TXT" For Output As #logFile
    Print #logFile, "Log Importazione QUERY - " & Now
    Print #logFile, "------------------------------------------------"
    Print #logFile,
    
    '//@04_@IMPORTA@QUERY@COLL                 -------------------> RECUPERO COLLECTION DA IMPORTARE
    'Recupero la collezione di QUERY da importare
      Set Queries_Collection = CollectionQuery_PFunct()
      
    '//@05_@IMPORTA@@QUERYCONTROLLO 1          -------------------> CONTROLLO 1 ESISTONO OGGETTI COLLECTION
    '//se non esistono oggetti collection EXIT
    If Queries_Collection.count = 0 Then
        Print #logFile, "Nessuna QUERY da importare."
        Close #logFile
        '//Prima di uscire restto l'oggetto collection
        Set Queries_Collection = Nothing
        Exit Sub
    End If
    
    '//@06_@IMPORTA@QUERY@CONTROLLO 2          -------------------> CONTROLLO 2 ESISTE PATH VALIDA
    '//RECUPERO LA PATH =  Controllo se il percorso del database di origine è valido
    sourceDBPath = PathFile_s_pFunct
    
    '//@07_@IMPORTA@QUERY@CONTROLLO 3          -------------------> CONTROLLO 3 PATH NULL EXIT
    If sourceDBPath = "" Then
        MsgBox "Errore: path e file nulli.", vbCritical
        Print #logFile, "Errore: path e file nulli."
        Close #logFile
        Exit Sub
    End If
    
    '//@08_@IMPORTA@QUERY@CONTROLLO 4          -------------------> CONTROLLO 4 SE DB NULL EXIT
    '//SE NON ESISTE IL FILE DB = Controllo se il file del database di origine esiste
    If Dir(sourceDBPath) = "" Then
        MsgBox "Errore: il file del database di origine non esiste.", vbCritical
        Print #logFile, "Errore: il file del database di origine non esiste."
        Close #logFile
        Exit Sub
    End If
    
    '//@09_@IMPORTA@QUERY@AVVISI+MSG           -------------------> AVVISI E IMPOSTAZIONE MESSAGGIO FINALE
    '//.......................................................................................//
      '//AVVISI MESSAGGIO IMPOSTATO = TRUE - >  Controllo del valore di avviso
      MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
      If MyValue = 0 Then Exit Sub
      
      '//CONTROLLO NOME PROGETTO =  Controllo del nome del progetto
      Bool1 = CheckProjectName_b_PFunct()
      If Bool1 = True Then Exit Sub
      
    '//.......................................................................................//
    
    
        '//@10_@IMPORTA@QUERY@PROGRSESSBAR_01      -------------------> PROGRESS BAR STEP 01 RECUPERO PARAMETRI
        '//@PROGRESBAR_STEP_01=ATTIVO LA PROGRESS BAR CHIAMANDO ANCHE LE 4 FUNZIONI PER RECUPERARE I PARAMETRI  @progress@bar
        '//********************************************************************************************************************//
        '//@PROGRES@BAR@FASE_01_(imposto le variabili e recupero i valori dalle funzioni)
        
                
                '//NOTE: devi creare 4 funzioni per avere i 4 parametri per chiamare la progress bar: _
                     @01)@IMPOSTO@VARIABILE@NUMERO_COLLECTION          ----> recupero il NroCollection_i _
                     @02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl _
                     @03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
                     @04)@FUNZIONE@TIPO_IMPORT                         ----> RESTITUISCE TipoImport_s _
                     @05)@ROUTINE_APRI_LA_PROGRES_BAR                  ----> RESTITUISCE NULL
                     
                    
                      DoEvents                      '//PER CONSENTIRE IL BLOCCO DEBUG
                  
                      '//@01)@IMPOSTO@VARIABILE@NUMERO_COLLECTION          ----> recupero il NroCollection_i _
                             Imposto il totale oggetti QUERY da importare se esistono
                      TotCollection_Queries_i = Queries_Collection.count
             
                    
                    '//@02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl
                       Twips_dbl = Calcola_Twips_pFunct(TotCollection_Queries_i)
                    
                    '//@03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
                     '//chiama la funzione percentuale incremento e ne restituisce la _
                         percentuale come double
                     PercIncr_dbl = PercentualeIncremento_pFunct(TotCollection_Queries_i)             '//calcolo il numero di collection QUERY da controllare
         
                     '//@04)@FUNZIONE@TIPO_IMPORT                         ----> RESTITUISCE TipoImport_s _
                     '//chiama la funzione TIPO IMPORT
                         TipoImport_s = TipoImport_pFunct("\LOG_IMPORTA_QUERY.TXT")
                          
                         '//CONTROLLO_TIPO_IMPORTAZIONE: se la funzione restituisce NON ESISTE esci dalla IMPORTAZIONE
                         If TipoImport_s <> "IMPORTA_QUERY" Then
                            '//chiudo il file log prima di uscire
                            Close #logFile
                            GoTo Exit_Cmd_Importa_QUERY_Click
                         End If
                      
                     '//@05)@ROUTINE_APRI_LA_PROGRES_BAR                  ----> RESTITUISCE NULL
                         '//Apri la QUERY della progressBar
                         '//
                             Call ApriProgressBar
                             ' Aspetta che il QUERY sia effettivamente caricato
                              DoEvents
                         
        '//@PROGRESBAR_STEP_01=ATTIVO LA PROGRESS BAR CHIAMANDO ANCHE LE 4 FUNZIONI PER RECUPERARE I PARAMETRI *** FINE ***
        '//********************************************************************************************************************//
    
    
    '//===========================================================================================================================//
    '//             ATTIVAZIONE CICLO DI IMPORTAZIONE DAL PROGETTO ESTERNO        *** INIZIO ****
    '//===========================================================================================================================//
            
          '//@11_@IMPORTA@QUERY@CONTROLLO_PROG       -------------------> CONTROLLO SE STIAMO NEL PROGETTO ESTERNO
          ' Se siamo su un progetto esterno (Bool1 = False), procedo con l'importazione
          If Bool1 = False Then
                ' Apertura del database di origine e connessione al database corrente
                Set ws = DBEngine.Workspaces(0)
                Set sourceDB = ws.OpenDatabase(sourceDBPath)
                Set db = CurrentDb
                
                
                '//reset
                ProgressPercentTwips_dbl = 0
                
                
                '//@12_@IMPORTA@QUERY@PROG@CICLO           -------------------> NEL PROGETTO ESTERNO CICLO NEGLI OGGETTI
                ' Ciclo attraverso la collezione di QUERY da importare
                For Each Queries_Name_v In Queries_Collection
                    myQueries_Name_s = Queries_Name_v
                    
                    '//@13_@IMPORTA@QUERY@PROG@CONTR_01        -------------------> CONTROLLO 1 SE ESTE OGGETTO DA IMPORTARE
                    ' Controllo se la QUERY esiste nel database di origine
                    If ObjectExists_QUERY(sourceDB, acQuery, myQueries_Name_s) Then
                        ' Se la QUERY esiste nel database corrente, la elimino prima di importarla
                        If ObjectExists_QUERY(db, acQuery, myQueries_Name_s) Then
                            DoCmd.DeleteObject acQuery, myQueries_Name_s
                            deletedCount_i = deletedCount_i + 1
                            deleted_Queries_s = deleted_Queries_s & myQueries_Name_s & vbCrLf
                            Print #logFile, "QUERY eliminata: " & myQueries_Name_s
                             
                              
                        End If
                        
                          '//@14_@IMPORTA@QUERY@PROG@IMPORT          -------------------> IMPORTAZIONE DELL'OGGETTO DAL DB ESTERNO
                          ' Importazione della QUERY dal database di origine
                          DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acQuery, myQueries_Name_s, myQueries_Name_s
                          importedCount_i = importedCount_i + 1
                          imported_Queries_s = imported_Queries_s & myQueries_Name_s & vbCrLf
                          
                          '//SCRIVO SUL @FILE@LOG
                          Print #logFile, "NRO: " & importedCount_i & " QUERY importata: " & myQueries_Name_s
                          
                              
                        '//@15_@IMPORTA@QUERY@PROGRSESSBAR_02      -------------------> PROGRESS BAR STEP 02 CALCOLO ED INVIO PARAMETRI
                        '//@PROGRESBAR_STEP_02= CALCOLO PARAMETRI ED INVIO DATI ALLA FUNZIONE DI AGGIORNAMENTO   @progress@bar
                        '//********************************************************************************************************************//
                        '//NOTE: impostazioni e recupero valori dalla funzioni base per la progress bar.
                            
                              '// RECUPERO il messaggio tipo di importazione delle collection e numero di oggetti da importare .
                                TipoImport_s = "QUERY IMPORTATI NRO : " & importedCount_i & " su TOTALE OGGETTI: " & TotCollection_Queries_i       '// STRINGA MESSAGGIO PER IL TIPO DI IMPORTAZIONE E NRO OGGETTI
                                
                                '//LA PERCENTUALE INCREMENTATA = totale di incremento % per ogni oggetto importato
                                TotPerc_dbl = TotPerc_dbl + PercIncr_dbl
                                
                                '//TOTALE % DI INCREMENTO IN TWIPS, con la ProgressPercentTwips_dbl è la variabile che indica la %di incremento per ogni oggetto + i _
                                   twips necessari per ogni elemento (su 10 totale della barra)
                                ProgressPercentTwips_dbl = ProgressPercentTwips_dbl + (PercIncr_dbl + Twips_dbl)          '//la % di accrescimento barra
                                
                                '//TODO: DA CANCELLARE...
                                'Dbl1 = Dbl1 + Twips_dbl                                         '//twips di aumento barra
                                
                                
                                          '//@15.BIS_@IMPORTA@QUERY@PROGRESSBAR_CALL -------------------> VIENE CHIAMATA LA PROGRESS BAR CON I PARAMETRI
                                          '//CHIAMO LA FUNZIONE @AGGIORNA@PROGRESS@BAR
                                          '//.................................................................//
                                          '//note: chiamo la routine di aggiornamento della progress bar che _
                                              puo essere utilizzata per ulteriori integrazioni di calcolo. _
                                              Devo inviare ma devo inviare 5 parametri : _
                                                  TotCollection_Queries_i       = il TOTALE POTENZIALE degli oggetti da importare _
                                                  NroCollection_i           = il numero della collezione per il calcolo; _
                                                  Twips                     = i twip di incremento _
                                                  ProgressPercentTwips_dbl  =  La percentuale di incremento ( ProgressPercentTwips_dbl  = ProgressPercentTwips_dbl + PercIncr_dbl) _
                                                  TipoImport_s              = il tipo di importazione
                                                  
                                              AggiornaProgressBar _
                                                          TotCollection_Queries_i, _
                                                          ProgressPercentTwips_dbl, _
                                                          Twips_dbl, _
                                                          TotPerc_dbl, _
                                                          TipoImport_s
                                          '//.................................................................//
        
                               
                                T = Timer
                                While Timer < T + 0.001
                                    DoEvents  ' Continua a gestire gli eventi durante l'attesa
                                Wend
                          
                                  
                        '//@PROGRESBAR_STEP_02= CALCOLO PARAMETRI ED INVIO DATI ALLA FUNZIONE DI AGGIORNAMENTO   *** FINE ***
                        '//********************************************************************************************************************//
        
            
                        
                    Else
                        
                        '//@16_@IMPORTA@QUERY@LOG@FILE             -------------------> SE IL FILE NON VIENE TROVATO SCRITTURAZIONE LOG
                        ' Se la QUERY non esiste nel database di origine, lo scrivo nel log
                        Print #logFile, "QUERY non trovata nel database di origine: " & myQueries_Name_s
                    End If
                
                Next Queries_Name_v
        
        
                                        
                        '//@17_@IMPORTA@QUERY@CLOSE                -------------------> CHIUSURA CON RITARDO DELLA PROGRESS BAR
                        '//CHIUSURA CON RITARDO DELLA BARRA DI PROGRESSIONE @progress@bar
                        '//---------------------------------------------------------------------------//
                    
                                        T = Timer
                                        
                                        '// RITARDO CHIUSURA DI 5 SECONDI
                                        While Timer < T + 5
                                            DoEvents  ' Continua a gestire gli eventi durante l'attesa
                                        Wend
                                        
                                        ' Chiudi la QUERY UTILITA_Msys_Frm21_frmProgressBar
                                        DoCmd.Close acForm, "UTILITA_Msys_Frm21_frmProgressBar", acSaveNo
                                        
                        '//---------------------------------------------------------------------------//
                        
                        
                        
                '//@18_@IMPORTA@QUERY@NOTHING              -------------------> IMPOSTO A NULL TUTTI GLI OGGETTI
                '//---------------------------------------------------------------------------//
                ' Chiudo il database di origine
                        Set frm_obj = Nothing                 ' Rilascia il riferimento al QUERY PROGRES BARR
                      
                      '//chiudo il db
                      sourceDB.Close
                      Set sourceDB = Nothing
                      Set ws = Nothing
                  
                      ' Messaggio finale con riepilogo delle operazioni
                      If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                          MsgBox "Importazione completata." & vbCrLf & _
                              "Queries importate: " & importedCount_i & vbCrLf & imported_Queries_s & _
                              "Queries esistenti cancellate: " & deletedCount_i & vbCrLf & deleted_Queries_s, vbInQUERYation
                      End If
                  
                      ' Scrivo nel file di log il riepilogo finale
                      Print #logFile, "Importazione completata con " & importedCount_i & " Queries importate."
                      Close #logFile
                  
                      ' Pulizia delle variabili del db e della collection Queries
                      Set db = Nothing
                      Set Queries_Collection = Nothing
                      
                '//---------------------------------------------------------------------------//
                
          End If '//If Bool1 = False Then = Progetto esterno
      
    '//===========================================================================================================================//
    '//             ATTIVAZIONE CICLO DI IMPORTAZIONE DAL PROGETTO ESTERNO        *** FINE ****
    '//===========================================================================================================================//
    
    
'// USCITA E GESTIONE ERRORI
'//---------------------------------------------------------------------------------------------------------//

Exit_Cmd_Importa_QUERY_Click:

Exit Sub
    
ErrorHandler:
    ' Gestione degli errori con messaggio a video e scrittura nel log
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Print #logFile, "Errore durante l'importazione: " & Err.Description
    If Not sourceDB Is Nothing Then sourceDB.Close
    Close #logFile
    Set sourceDB = Nothing
    Set ws = Nothing
End Sub


'//CONTROLLO_ESISTEMA_QUERY(efficientata la procedura con l'interrogazione sulla tabella MSysObjects per il controllo delle QUERY esterne)
'//FUNZIONE DI CONTROLLO SE ESISTE L'OGGETTO QUERY
Function ObjectExists_QUERY(db As DAO.Database, objType As Integer, objName As String) As Boolean
    Dim rs As DAO.Recordset
    Dim QUERYExists As Boolean
    
    
    '//CONTROLLO OGGETTO
    Debug.Print " Function ObjectExists_QUERY - oggetto della ricerca : objName "
    Debug.Print objName
     
    QUERYExists = False
    On Error Resume Next
    '//FILTRO CONTROLLO ESISTENZA QUERY - CODICE TYPE 5
    Set rs = db.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=5 AND Name='" & objName & "'")
    If Not rs.EOF Then
        QUERYExists = True
    End If
    rs.Close
    Set rs = Nothing
    ObjectExists_QUERY = QUERYExists
    On Error GoTo 0
End Function
'//IMPORTA TUTTI GLI OGGETTI QUERY DEL DB MSYSDBEST *** FINE ***
'//----------------------------------------------------------------------------------------//






'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI QUERY DEL DB ESTERNO *** FINE ***
'//                     @IMPORTA@OGGETTI@QUERY@MSYS
'//=====================================================================================================================//






'//IMPORTA QUERY HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//
'//Cmd_Importa_QUERY_HELP_Click    = HELP GUIDA PER IL COMANDO DI IMPORTAZIONE QUERY
Private Sub Cmd_Importa_QUERY_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_IMPORTA_QUERY.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub

'//IMPORTA QUERY HELP *** FINE ***
'//-----------------------------------------------------------------------------------//


'//=========================================================================================//
'//
'//                 GESTIONE DELLE QUERY ESTERNE O INTERNE   *** FINE ***
'//@GESTIONE@DELLE@QUERY
'//=========================================================================================//





    












'//IMPORTA REPORT HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_Importa_REPORT_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_IMPORTA_REPORT.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa REPORT _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub

'//IMPORTA REPORT HELP *** FINE ***
'//-----------------------------------------------------------------------------------//









'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI TABELLE DEL DB ESTERNO *** INIZIO ***
'//                     @IMPORTA@OGGETTI@TABELLE@MSYS
'//=====================================================================================================================//



'//IMPORTA TUTTI GLI OGGETTI TABELLE DEL DB MSYSDBEST - LE ATTIVITA
'//----------------------------------------------------------------------------------------//
'//@01_@IMPORTA@TABELLE@DIM                  -------------------> DIM VARIABILI
'//@02_@IMPORTA@TABELLE@REST                 -------------------> RESET DELLE VARIABILI
'//@03_@IMPORTA@TABELLE@LOG                  -------------------> LOG FILE INIZIALE
'//@04_@IMPORTA@TABELLE@COLL                 -------------------> RECUPERO COLLECTION DA IMPORTARE
'//@05_@IMPORTA@@TABELLECONTROLLO 1          -------------------> CONTROLLO 1 ESISTONO OGGETTI COLLECTION
'//@06_@IMPORTA@TABELLE@CONTROLLO 2          -------------------> CONTROLLO 2 ESISTE PATH VALIDA
'//@07_@IMPORTA@TABELLE@CONTROLLO 3          -------------------> CONTROLLO 3 PATH NULL EXIT
'//@08_@IMPORTA@TABELLE@CONTROLLO 4          -------------------> CONTROLLO 4 SE DB NULL EXIT
'//@09_@IMPORTA@TABELLE@AVVISI+MSG           -------------------> AVVISI E IMPOSTAZIONE MESSAGGIO FINALE
'//@10_@IMPORTA@TABELLE@PROGRSESSBAR_01      -------------------> PROGRESS BAR STEP 01 RECUPERO PARAMETRI
'//@11_@IMPORTA@TABELLE@CONTROLLO_PROG       -------------------> CONTROLLO SE STIAMO NEL PROGETTO ESTERNO
'//@12_@IMPORTA@TABELLE@PROG@CICLO           -------------------> NEL PROGETTO ESTERNO CICLO NEGLI OGGETTI
'//@13_@IMPORTA@TABELLE@PROG@CONTR_01        -------------------> CONTROLLO 1 SE ESTE OGGETTO DA IMPORTARE
'//@14_@IMPORTA@TABELLE@PROG@IMPORT          -------------------> IMPORTAZIONE DELL'OGGETTO DAL DB ESTERNO
'//@15_@IMPORTA@TABELLE@PROGRSESSBAR_02      -------------------> PROGRESS BAR STEP 02 CALCOLO ED INVIO PARAMETRI
'//@15.BIS_@IMPORTA@TABELLE@PROGRESSBAR_CALL -------------------> VIENE CHIAMATA LA PROGRESS BAR CON I PARAMETRI
'//@16_@IMPORTA@TABELLE@LOG@FILE             -------------------> SE IL FILE NON VIENE TROVATO SCRITTURAZIONE LOG
'//@17_@IMPORTA@TABELLE@CLOSE                -------------------> CHIUSURA CON RITARDO DELLA PROGRESS BAR
'//@18_@IMPORTA@TABELLE@NOTHING              -------------------> IMPOSTO A NULL TUTTI GLI OGGETTI


Private Sub Cmd_Importa_TABELLE_Click()
    
    '//@01_@IMPORTA@TABELLE@DIM                  -------------------> DIM VARIABILI
    ' Dichiarazione delle variabili
    '//oggetti TABELLE e variabile nome TABELLE
    Dim Tables_Collection As Collection                       '//oggetti collezione
    Dim Tables_Name_v As Variant
    Dim deleted_Tables_s As String
    Dim imported_Tables_s As String
    Dim myTables_Name_s As String
    
    '//il database + path
    Dim sourceDB As DAO.Database
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim sourceDBPath As String
    
    '//contatori oggetti: TOTALE OGGETTI PREVISTI; da importare e cancellare
    Dim TotCollection_Tables_i As Integer                          '//totale degli oggetti TABELLE della collezione da importare
    Dim importedCount_i As Integer                              '//oggetti importati
    Dim deletedCount_i As Integer                               '//oggetti cancellati
    
    
    '//bool e log file
    Dim Bool1 As Boolean
    Dim MyValue As Integer
    Dim logFile As Integer
    
    On Error GoTo ErrorHandler
    
    '//@02_@IMPORTA@TABELLE@REST                 -------------------> RESET DELLE VARIABILI
    '//RESET = Inizializza i contatori e le stringhe di riepilogo
    TotCollection_Tables_i = 0
    importedCount_i = 0
    deletedCount_i = 0
    deleted_Tables_s = ""
    imported_Tables_s = ""
    TotPerc_dbl = 0
    ProgressPercentTwips_dbl = 0
    
    '//@03_@IMPORTA@TABELLE@LOG                  -------------------> LOG FILE INIZIALE
    '@FILE@LOG@TABELLE: Apre il file di log per scrivere i dettagli dell'importazione apre nella path corrente.
    logFile = FreeFile()
    Open CurrentProject.PATH & "\LOG_IMPORTA_TABELLE.TXT" For Output As #logFile
    Print #logFile, "Log Importazione TABELLE - " & Now
    Print #logFile, "------------------------------------------------"
    Print #logFile,
    
    '//@04_@IMPORTA@TABELLE@COLL                 -------------------> RECUPERO COLLECTION DA IMPORTARE
    'Recupero la collezione di TABELLE da importare
      Set Tables_Collection = CollectionTable_PFunct()
      
    '//@05_@IMPORTA@@TABELLECONTROLLO 1          -------------------> CONTROLLO 1 ESISTONO OGGETTI COLLECTION
    '//se non esistono oggetti collection EXIT
    If Tables_Collection.count = 0 Then
        Print #logFile, "Nessuna TABELLE da importare."
        Close #logFile
        '//Prima di uscire restto l'oggetto collection
        Set Tables_Collection = Nothing
        Exit Sub
    End If
    
    '//@06_@IMPORTA@TABELLE@CONTROLLO 2          -------------------> CONTROLLO 2 ESISTE PATH VALIDA
    '//RECUPERO LA PATH =  Controllo se il percorso del database di origine è valido
    sourceDBPath = PathFile_s_pFunct
    
    '//@07_@IMPORTA@TABELLE@CONTROLLO 3          -------------------> CONTROLLO 3 PATH NULL EXIT
    If sourceDBPath = "" Then
        MsgBox "Errore: path e file nulli.", vbCritical
        Print #logFile, "Errore: path e file nulli."
        Close #logFile
        Exit Sub
    End If
    
    '//@08_@IMPORTA@TABELLE@CONTROLLO 4          -------------------> CONTROLLO 4 SE DB NULL EXIT
    '//SE NON ESISTE IL FILE DB = Controllo se il file del database di origine esiste
    If Dir(sourceDBPath) = "" Then
        MsgBox "Errore: il file del database di origine non esiste.", vbCritical
        Print #logFile, "Errore: il file del database di origine non esiste."
        Close #logFile
        Exit Sub
    End If
    
    '//@09_@IMPORTA@TABELLE@AVVISI+MSG           -------------------> AVVISI E IMPOSTAZIONE MESSAGGIO FINALE
    '//.......................................................................................//
      '//AVVISI MESSAGGIO IMPOSTATO = TRUE - >  Controllo del valore di avviso
      MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
      If MyValue = 0 Then Exit Sub
      
      '//CONTROLLO NOME PROGETTO =  Controllo del nome del progetto
      Bool1 = CheckProjectName_b_PFunct()
      If Bool1 = True Then Exit Sub
      
    '//.......................................................................................//
    
    
        '//@10_@IMPORTA@TABELLE@PROGRSESSBAR_01      -------------------> PROGRESS BAR STEP 01 RECUPERO PARAMETRI
        '//@PROGRESBAR_STEP_01=ATTIVO LA PROGRESS BAR CHIAMANDO ANCHE LE 4 FUNZIONI PER RECUPERARE I PARAMETRI  @progress@bar
        '//********************************************************************************************************************//
        '//@PROGRES@BAR@FASE_01_(imposto le variabili e recupero i valori dalle funzioni)
        
                
                '//NOTE: devi creare 4 funzioni per avere i 4 parametri per chiamare la progress bar: _
                     @01)@IMPOSTO@VARIABILE@NUMERO_COLLECTION          ----> recupero il NroCollection_i _
                     @02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl _
                     @03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
                     @04)@FUNZIONE@TIPO_IMPORT                         ----> RESTITUISCE TipoImport_s _
                     @05)@ROUTINE_APRI_LA_PROGRES_BAR                  ----> RESTITUISCE NULL
                     
                    
                      DoEvents                      '//PER CONSENTIRE IL BLOCCO DEBUG
                  
                      '//@01)@IMPOSTO@VARIABILE@NUMERO_COLLECTION          ----> recupero il NroCollection_i _
                             Imposto il totale oggetti TABELLE da importare se esistono
                      TotCollection_Tables_i = Tables_Collection.count
             
                    
                    '//@02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl
                       Twips_dbl = Calcola_Twips_pFunct(TotCollection_Tables_i)
                    
                    '//@03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
                     '//chiama la funzione percentuale incremento e ne restituisce la _
                         percentuale come double
                     PercIncr_dbl = PercentualeIncremento_pFunct(TotCollection_Tables_i)             '//calcolo il numero di collection TABELLE da controllare
         
                     '//@04)@FUNZIONE@TIPO_IMPORT                         ----> RESTITUISCE TipoImport_s _
                     '//chiama la funzione TIPO IMPORT
                         TipoImport_s = TipoImport_pFunct("\LOG_IMPORTA_TABELLE.TXT")
                          
                         '//CONTROLLO_TIPO_IMPORTAZIONE: se la funzione restituisce NON ESISTE esci dalla IMPORTAZIONE
                         If TipoImport_s <> "IMPORTA_TABELLE" Then
                            '//chiudo il file log prima di uscire
                            Close #logFile
                            GoTo Exit_Cmd_Importa_TABELLE_Click
                         End If
                      
                     '//@05)@ROUTINE_APRI_LA_PROGRES_BAR                  ----> RESTITUISCE NULL
                         '//Apri la TABELLE della progressBar
                         '//
                             Call ApriProgressBar
                             ' Aspetta che il TABELLE sia effettivamente caricato
                              DoEvents
                         
        '//@PROGRESBAR_STEP_01=ATTIVO LA PROGRESS BAR CHIAMANDO ANCHE LE 4 FUNZIONI PER RECUPERARE I PARAMETRI *** FINE ***
        '//********************************************************************************************************************//
    
    
    '//===========================================================================================================================//
    '//             ATTIVAZIONE CICLO DI IMPORTAZIONE DAL PROGETTO ESTERNO        *** INIZIO ****
    '//===========================================================================================================================//
            
          '//@11_@IMPORTA@TABELLE@CONTROLLO_PROG       -------------------> CONTROLLO SE STIAMO NEL PROGETTO ESTERNO
          ' Se siamo su un progetto esterno (Bool1 = False), procedo con l'importazione
          If Bool1 = False Then
                ' Apertura del database di origine e connessione al database corrente
                Set ws = DBEngine.Workspaces(0)
                Set sourceDB = ws.OpenDatabase(sourceDBPath)
                Set db = CurrentDb
                
                
                '//reset
                ProgressPercentTwips_dbl = 0
                
                
                '//@12_@IMPORTA@TABELLE@PROG@CICLO           -------------------> NEL PROGETTO ESTERNO CICLO NEGLI OGGETTI
                ' Ciclo attraverso la collezione di TABELLE da importare
                For Each Tables_Name_v In Tables_Collection
                    myTables_Name_s = Tables_Name_v
                    
                    '//@13_@IMPORTA@TABELLE@PROG@CONTR_01        -------------------> CONTROLLO 1 SE ESTE OGGETTO DA IMPORTARE
                    ' Controllo se la TABELLE esiste nel database di origine
                    If ObjectExists_TABELLE(sourceDB, acTable, myTables_Name_s) Then
                        ' Se la TABELLE esiste nel database corrente, la elimino prima di importarla
                        If ObjectExists_TABELLE(db, acTable, myTables_Name_s) Then
                            DoCmd.DeleteObject acTable, myTables_Name_s
                            deletedCount_i = deletedCount_i + 1
                            deleted_Tables_s = deleted_Tables_s & myTables_Name_s & vbCrLf
                            Print #logFile, "TABELLE eliminata: " & myTables_Name_s
                             
                              
                        End If
                            '//todo: ATTENZIONE a volte importa le tabelle non le collega??? correggere NON DEVE IMPORTARE MA COLLEGARE!!!
                        
                          '//@14_@IMPORTA@TABELLE@PROG@IMPORT          -------------------> IMPORTAZIONE DELL'OGGETTO DAL DB ESTERNO
                          ' Importazione della TABELLE dal database di origine
                          DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acTable, myTables_Name_s, myTables_Name_s
                          importedCount_i = importedCount_i + 1
                          imported_Tables_s = imported_Tables_s & myTables_Name_s & vbCrLf
                          
                          '//SCRIVO SUL @FILE@LOG
                          Print #logFile, "NRO: " & importedCount_i & " TABELLE importata: " & myTables_Name_s
                          
                              
                        '//@15_@IMPORTA@TABELLE@PROGRSESSBAR_02      -------------------> PROGRESS BAR STEP 02 CALCOLO ED INVIO PARAMETRI
                        '//@PROGRESBAR_STEP_02= CALCOLO PARAMETRI ED INVIO DATI ALLA FUNZIONE DI AGGIORNAMENTO   @progress@bar
                        '//********************************************************************************************************************//
                        '//NOTE: impostazioni e recupero valori dalla funzioni base per la progress bar.
                            
                              '// RECUPERO il messaggio tipo di importazione delle collection e numero di oggetti da importare .
                                TipoImport_s = "TABELLE IMPORTATI NRO : " & importedCount_i & " su TOTALE OGGETTI: " & TotCollection_Tables_i       '// STRINGA MESSAGGIO PER IL TIPO DI IMPORTAZIONE E NRO OGGETTI
                                
                                '//LA PERCENTUALE INCREMENTATA = totale di incremento % per ogni oggetto importato
                                TotPerc_dbl = TotPerc_dbl + PercIncr_dbl
                                
                                '//TOTALE % DI INCREMENTO IN TWIPS, con la ProgressPercentTwips_dbl è la variabile che indica la %di incremento per ogni oggetto + i _
                                   twips necessari per ogni elemento (su 10 totale della barra)
                                ProgressPercentTwips_dbl = ProgressPercentTwips_dbl + (PercIncr_dbl + Twips_dbl)          '//la % di accrescimento barra
                                
                                '//TODO: DA CANCELLARE...
                                'Dbl1 = Dbl1 + Twips_dbl                                         '//twips di aumento barra
                                
                                
                                          '//@15.BIS_@IMPORTA@TABELLE@PROGRESSBAR_CALL -------------------> VIENE CHIAMATA LA PROGRESS BAR CON I PARAMETRI
                                          '//CHIAMO LA FUNZIONE @AGGIORNA@PROGRESS@BAR
                                          '//.................................................................//
                                          '//note: chiamo la routine di aggiornamento della progress bar che _
                                              puo essere utilizzata per ulteriori integrazioni di calcolo. _
                                              Devo inviare ma devo inviare 5 parametri : _
                                                  TotCollection_Tables_i       = il TOTALE POTENZIALE degli oggetti da importare _
                                                  NroCollection_i           = il numero della collezione per il calcolo; _
                                                  Twips                     = i twip di incremento _
                                                  ProgressPercentTwips_dbl  =  La percentuale di incremento ( ProgressPercentTwips_dbl  = ProgressPercentTwips_dbl + PercIncr_dbl) _
                                                  TipoImport_s              = il tipo di importazione
                                                  
                                              AggiornaProgressBar _
                                                          TotCollection_Tables_i, _
                                                          ProgressPercentTwips_dbl, _
                                                          Twips_dbl, _
                                                          TotPerc_dbl, _
                                                          TipoImport_s
                                          '//.................................................................//
        
                               
                                T = Timer
                                While Timer < T + 0.001
                                    DoEvents  ' Continua a gestire gli eventi durante l'attesa
                                Wend
                          
                                  
                        '//@PROGRESBAR_STEP_02= CALCOLO PARAMETRI ED INVIO DATI ALLA FUNZIONE DI AGGIORNAMENTO   *** FINE ***
                        '//********************************************************************************************************************//
        
            
                        
                    Else
                        
                        '//@16_@IMPORTA@TABELLE@LOG@FILE             -------------------> SE IL FILE NON VIENE TROVATO SCRITTURAZIONE LOG
                        ' Se la TABELLE non esiste nel database di origine, lo scrivo nel log
                        Print #logFile, "TABELLE non trovata nel database di origine: " & myTables_Name_s
                    End If
                
                Next Tables_Name_v
        
        
                                        
                        '//@17_@IMPORTA@TABELLE@CLOSE                -------------------> CHIUSURA CON RITARDO DELLA PROGRESS BAR
                        '//CHIUSURA CON RITARDO DELLA BARRA DI PROGRESSIONE @progress@bar
                        '//---------------------------------------------------------------------------//
                    
                                        T = Timer
                                        
                                        '// RITARDO CHIUSURA DI 5 SECONDI
                                        While Timer < T + 5
                                            DoEvents  ' Continua a gestire gli eventi durante l'attesa
                                        Wend
                                        
                                        ' Chiudi la TABELLE UTILITA_Msys_Frm21_frmProgressBar
                                        DoCmd.Close acForm, "UTILITA_Msys_Frm21_frmProgressBar", acSaveNo
                                        
                        '//---------------------------------------------------------------------------//
                        
                        
                        
                '//@18_@IMPORTA@TABELLE@NOTHING              -------------------> IMPOSTO A NULL TUTTI GLI OGGETTI
                '//---------------------------------------------------------------------------//
                ' Chiudo il database di origine
                        Set frm_obj = Nothing                 ' Rilascia il riferimento al TABELLE PROGRES BARR
                      
                      '//chiudo il db
                      sourceDB.Close
                      Set sourceDB = Nothing
                      Set ws = Nothing
                  
                      ' Messaggio finale con riepilogo delle operazioni
                      If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                          MsgBox "Importazione completata." & vbCrLf & _
                              "Tables importate: " & importedCount_i & vbCrLf & imported_Tables_s & _
                              "Tables esistenti cancellate: " & deletedCount_i & vbCrLf & deleted_Tables_s, vbInTABELLEation
                      End If
                  
                      ' Scrivo nel file di log il riepilogo finale
                      Print #logFile, "Importazione completata con " & importedCount_i & " Tables importate."
                      Close #logFile
                  
                      ' Pulizia delle variabili del db e della collection Tables
                      Set db = Nothing
                      Set Tables_Collection = Nothing
                      
                '//---------------------------------------------------------------------------//
                
          End If '//If Bool1 = False Then = Progetto esterno
      
    '//===========================================================================================================================//
    '//             ATTIVAZIONE CICLO DI IMPORTAZIONE DAL PROGETTO ESTERNO        *** FINE ****
    '//===========================================================================================================================//
    
    
'// USCITA E GESTIONE ERRORI
'//---------------------------------------------------------------------------------------------------------//

Exit_Cmd_Importa_TABELLE_Click:

Exit Sub
    
ErrorHandler:
    ' Gestione degli errori con messaggio a video e scrittura nel log
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Print #logFile, "Errore durante l'importazione: " & Err.Description
    If Not sourceDB Is Nothing Then sourceDB.Close
    Close #logFile
    Set sourceDB = Nothing
    Set ws = Nothing
End Sub




'//CONTROLLO TIPO DI TABELLA DA IMPORTARE NEL DB ESTERNO                            **** INIZIO ****
'//=================================================================================================================//
'//CONTROLLO_ESISTEMA_TABELLE(efficientata la procedura con l'interrogazione sulla tabella MSysObjects per il controllo delle TABELLE esterne)
'//FUNZIONE DI CONTROLLO SE ESISTE L'OGGETTO TABELLE
'    '//@CONTROLLO@TABELLE@DI@SISTEMA_(tipo di tabelle, collegata di sistema oppure 'originaria)
'    '//CONTROLLO OGGETTO TABELLA DISTINGUENDO TABELLA COLLEGATA TABELLA DI SISTEMA 'TABELLA ORGINARIA _
'            TABELLA COLLEGATA   campo Database  >"" (non Nullo) _
'            TABELLA COLLEGATA   campo Flags     = 2097152 _
'            TABELLA DI SISTEMA  campo Type      = 6 _

'            TABELLA DI SISTEMA  campo Database  ="" _
'            TABELLA DI SISTEMA  campo Flags     = _
'            TABELLA DI SISTEMA  campo Type      = 5 _

'            TABELLA ORIGINARIA  campo Database  = "" (Nullo) _
'            TABELLA ORIGINARIA  campo Type      = 0 _
'            TABELLA ORIGINARIA  campo Type      = 1 _

Function ObjectExists_TABELLE(db As DAO.Database, objType As Integer, objName As String) As Boolean
    Dim rs As DAO.Recordset
    Dim TABELLEExists As Boolean
    Dim sqlQuery As String
    Dim i As Integer
    Dim queryArray(2) As String

    Debug.Print "?? Controllo tabella: " & objName

    TABELLEExists = False
    On Error Resume Next
    
    '//FILTRO NELLA TABELLA DI SISTEMA MSysObjects = Il tipo è diverso a seconda del tipo di tabella ed _
       il filtro viene attivato a seconda dei diversi parametri necessari per distingure le tabella di sistema.
    ' 01= Tabelle ORIGINARIE (Type = 1, Flags = 0)
    queryArray(0) = "SELECT Name FROM MSysObjects " & _
                    "WHERE Type = 1 AND Flags = 0 " & _
                    "AND Name='" & objName & "'"

    ' 02= Tabelle COLLEGATE (Type = 6, Flags = 2097152, Database NOT NULL)
    queryArray(1) = "SELECT Name FROM MSysObjects " & _
                    "WHERE Type = 6 AND Flags = 2097152 " & _
                    "AND Database IS NOT NULL AND Database <> '' " & _
                    "AND Name='" & objName & "'"

    ' 03= Tabelle DI SISTEMA (Type = 5)
    queryArray(2) = "SELECT Name FROM MSysObjects " & _
                    "WHERE Type = 5 " & _
                    "AND Name='" & objName & "'"

    ' ?? Testa ogni query e restituisce True se trova la tabella
    For i = 0 To 2
        Debug.Print "?? Test Query: " & queryArray(i)
        Set rs = db.OpenRecordset(queryArray(i))

        If Not rs.EOF Then
            TABELLEExists = True
            rs.Close
            Set rs = Nothing
            Exit For   ' Esce dal loop appena trova un match
        End If

        rs.Close
        Set rs = Nothing
    Next i

    ObjectExists_TABELLE = TABELLEExists
    On Error GoTo 0
End Function


'//CONTROLLO TIPO DI TABELLA DA IMPORTARE NEL DB ESTERNO                            **** FINE ****
'//=================================================================================================================//

'//=================================================================================================================//




'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI TABELLE DEL DB ESTERNO *** FINE ***
'//                     @IMPORTA@OGGETTI@TABELLE@MSYS
'//=====================================================================================================================//





'//=====================================================================================================================//
'//                     DELETE OGGETTI TABLE DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE TABLE HELP
'//=====================================================================================================================//


'//CANCELLA GLI OGGETTI TABLE *** INIZIO *** 'DELETE TABLES
'//==========================================================================================================//

Private Sub Cmd_Delete_TABLE_Click()

    Dim dbCurrent As DAO.Database
    Dim tdf As DAO.TableDef
    Dim Tables_Collection As Collection
    Dim varItem As Variant
    
'//TODO: controllare se tale routine funziona???
On Error GoTo Err_Cmd_Delete_TABLE_Click
    
    
    '//RESET
    deletedTablesCount_i = 0
    deletedTables_s = ""
   
   
   
                   '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_ATTIVA_GEST_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE  @CONTROLLO@PROGETTO
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//

   
   
        
              '//ATTIVO LA COLLEZIONE
              '//------------------------------------------------------------------------//
              '//Note: CHIAMO LA FUNZIONE COLLECTION TABLE per ottenere la collezione di TABELLE
                
                '//
                Set tablesCollection = CollectionTable_PFunct()
                
              '//------------------------------------------------------------------------//
        
    
    ' Apri il database corrente
    Set dbCurrent = CurrentDb
    
    ' Loop attraverso tutte le tabelle nel database corrente
    For Each tdf In dbCurrent.TableDefs
        ' Controlla se la tabella è nella collezione e passo tutta la collezione Tables_Collection + il nome della tabella da controllare
        If InCollectionTables_Funct_b(tablesCollection, tdf.Name) Then
            ' Elimina la tabella
            On Error Resume Next
                deletedTablesCount_i = deletedTablesCount_i + 1
                DoCmd.DeleteObject acTable, tdf.Name
            On Error GoTo 0
        End If
    Next tdf
    
    
              '//APRI RIQUADRO TABELLE INDIPENDENTEMENTE DALLA LORO PRESENZA
              '//-------------------------------------------------------------------------------------------------//
              '//NOTE= Apri la sezione macro indipendentemente dalla presenza di oggetti _
                  Conta solo i moduli visibili, escludendo quelli di sistema. _
                  Se ci sono moduli, apre direttamente la sezione Moduli con DoCmd.SelectObject acModule, , True. _
                  Print Se; non; ci; sono; Moduli, prova; a; selezionare; una; tabella; e; poi; la; sezione; Moduli, per; forzare; l 'apertura del riquadro. _
                  Evita errori con On Error Resume Next in caso di assenza di oggetti.

                            
                            
                            
                          Dim db As DAO.Database
                          Dim tbl As TableDef
                          Dim tabelleVisibili As Integer
                          
                          Set db = CurrentDb
                          tabelleVisibili = 0 ' Inizializza il contatore
                          
                          ' Conta solo le tabelle visibili nel database
                          For Each tbl In db.TableDefs
                              If Left(tbl.Name, 4) <> "MSys" Then ' Esclude le tabelle di sistema
                                  tabelleVisibili = tabelleVisibili + 1
                              End If
                          Next tbl
                          
                          ' Se ci sono tabelle visibili, apri la sezione Tabelle
                          If tabelleVisibili > 0 Then
                              DoCmd.SelectObject acTable, , True
                          Else
                              ' Se non ci sono tabelle, prova a selezionare un altro oggetto per forzare l'apertura del riquadro
                              On Error Resume Next
                              DoCmd.SelectObject acQuery, , True ' Seleziona una query (se esiste)
                              DoCmd.SelectObject acTable, , True ' Ora prova a selezionare la sezione Tabelle
                              On Error GoTo 0
                          End If
                          
                          ' Pulizia memoria
                          Set tbl = Nothing
                          Set db = Nothing
                
                
                          
                    
              '//APRI RIQUADRO TABELLE INDIPENDENTEMENTE DALLA LORO PRESENZA *** FINE ***
              '//-------------------------------------------------------------------------------------------------//
            

    
            ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
            If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                MsgBox "LE TABELLE CANCELLATA DA QUESTO DATABASE SONO :  " & deletedTablesCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
            End If


                '




'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_TABLE_Click:
    Exit Sub

Err_Cmd_Delete_TABLE_Click:
        
    Set dbCurrent = Nothing
    
    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_TABLE_Click

End Sub


' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionTables_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col
        Debug.Print
        Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
        Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
        Debug.Print
        If varItem = key Then
            InCollectionTables_Funct_b = True
            Debug.Print "trovato restituisco : " & InCollectionTables_Funct_b
            Exit Function
        End If
    Next varItem
    InCollectionTables_Funct_b = False
    On Error GoTo 0
End Function



'//*** FINE ***
'//CANCELLA GLI OGGETTI TABLE *** INIZIO *** 'DELETE TABLES
'//==========================================================================================================//


'//DELETE TABLE HELP *** INIZIO ***
'//-----------------------------------------------------------------------------------//



Private Sub Cmd_Delete_TABLE_HELP_Click()
    
    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_Delete_TABLE.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto DELETE OGGETTI incluso in questa MODULI _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'DELETEzione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub




'//DELETE TABLE HELP *** FINE ***
'//-----------------------------------------------------------------------------------//


'//=====================================================================================================================//
'//                     DELETE OGGETTI TABLE DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE TABLE HELP
'//=====================================================================================================================//








'//@HELP@IMPORTA@TABELLE
Private Sub Importa_TABLE_HELP_Click()



    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_Importa_TABELLE.pdf "
                
                
            
            
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa form _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
             
            
    '//--------------------------------------------------------------------------------//--------//

      
End Sub



'//=====================================================================================================================//
'//                     DELETE OGGETTI QUERY DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE QUERY HELP
'//=====================================================================================================================//





'//CANCELLA GLI OGGETTI TABELLE *** INIZIO *** DELETE TABLE
'//==========================================================================================================//
'//@GESTIONE@DELLE@TABELLE    (@MODELLO@DELETE@OGGETTI)

'// PARAMETRI DA CAMBIARE : _
    -----> TABELLE _
    -----> DELLE@TABELLE _
    -----> DELETE TABLE _
    -----> CollectionTable_PFunct _
    -----> objectTable _
    -----> MATRICE_TABLE_Names_s _
    -----> tableiCount _
    -----> TableDef _
    -----> tdf _


Private Sub Cmd_Delete_TABELLE_Click()

    Dim dbCurrent As DAO.Database
    Dim qry As DAO.TableDef
    Dim objectQueries_Collection As Collection
    Dim varItem As Variant
    Dim MATRICE_TABLE_Names_s() As String             'matrice di TABELLE
    Dim tempCount As Integer
    Dim tableiCount As Integer
    Dim tdf As TableDef



'//TODO: controllare se tale routine funziona???
On Error GoTo Err_Cmd_Delete_TABELLE_Click


    '//RESET
    iCount = 0
    deleteTABELLEesCount_i = 0
    deletedTablesCount_i = 0
    deletedTables_s = ""




                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_ATTIVA_GEST_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE  @CONTROLLO@PROGETTO
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//


          '//ATTIVO LA COLLEZIONE
         '//------------------------------------------------------------------------//
          '//Note: CHIAMO LA FUNZIONE COLLECTION TABELLE per ottenere la collezione di TABELLE



        '//
        Set tablesCollection = CollectionTable_PFunct()

          '//------------------------------------------------------------------------//

            ' Loop attraverso tutte I REPORT nel database corrente
             Set dbCurrent = CurrentDb

             '//CALCOLO INDICE TABELLE
             iCount = dbCurrent.TableDefs.count

            Set dbCurrent = CurrentDb


            '//se ci sono TABELLE da cancellare
            If iCount > 0 Then

            '//reimposto la matrice sulle TABELLE esistenti nel db corrente
            ReDim MATRICE_TABLE_Names_s(iCount - 1)

            ' Costruisci un array con i nomi dei TABELLE
            tempCount = 0

            '//CARICO LA MATRICE: ciclo for per caricare MATRICE_TABLE_Names_s(tempCount)
            For Each tdf In dbCurrent.TableDefs
                MATRICE_TABLE_Names_s(tempCount) = tdf.Name
                tempCount = tempCount + 1
            Next tdf

            ' Cancella i TABELLE usando l'array dei nomi
            For i = 0 To tempCount - 1
                On Error Resume Next  ' Aggiungi gestione errori specifica per ogni iterazione
                Debug.Print "Cancellazione TABELLE: " & MATRICE_TABLE_Names_s(i)

                '//CONTROLLO MATRICE = COLLECTION
                Bool1 = InCollectionTABELLEes_Funct_b(tablesCollection, MATRICE_TABLE_Names_s(i))

                '//cancello solo se la matrice (i) esiste nella collection
                If Bool1 = True Then
                    Debug.Print
                    Debug.Print "cancello la TABELLE trovata sia nella collezione che nella matrice : " & MATRICE_TABLE_Names_s(i)
                    DoCmd.DeleteObject acTABELLE, MATRICE_TABLE_Names_s(i)

                    If Err.Number <> 0 Then
                    Debug.Print "Errore durante la cancellazione del TABELLE: " & MATRICE_TABLE_Names_s(i) & " - " & Err.Description
                    Err.Clear  ' Pulisci l'errore per la prossima iterazione
                    Else
                    deleteTABELLEesCount_i = deleteTABELLEesCount_i + 1
                    End If

                End If '//If Bool1 = True Then

                On Error GoTo 0  ' Reimposta gestione errori
            Next i
            End If

        ' Messaggio finale di riepilogo SE ABILITATO CON TRUE - FALSE DISABILITATO
        If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "LE TABELLE CANCELLATA DA QUESTO DATABASE SONO :  " & deleteTABELLEesCount_i, vbExclamation, "MSG CHIAMATO LA FUNZIONE PATHFILE"
        End If





'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................

Exit_Cmd_Delete_TABELLE_Click:
    Exit Sub

Err_Cmd_Delete_TABELLE_Click:

    Set dbCurrent = Nothing

    MsgBox Err.Description & " - Errore Messaggio -> : " & ProceduraMessaggioErrore_s & " Procedura -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Cmd_Delete_TABELLE_Click

End Sub

'//*** FINE ***
'//SENZA_PARAMETRI
'//========================================================================================================================================//



' Funzione per verificare se un elemento esiste nella collezione
Function InCollectionTABELLEes_Funct_b(col As Collection, key As String) As Boolean
    Dim varItem As Variant
    On Error Resume Next
    '//ITERNO NELLA COLLECTION PASSATA COME PARAMETRO
    For Each varItem In col

    Debug.Print
    Debug.Print "CONTROLLO COLLEZIONE = MATRICE (I)"
    Debug.Print "collezione: " & varItem & " = " & "matrice key: " & key
    Debug.Print
    If varItem = key Then
        InCollectionTABELLEes_Funct_b = True
        Debug.Print "trovato restituisco : " & InCollectionTABELLEes_Funct_b
        Exit Function
    End If
    Next varItem
    InCollectionTABELLEes_Funct_b = False
    On Error GoTo 0
End Function



'//CANCELLA GLI OGGETTI TABELLE *** FINE *** DELETE TABLE
'//=========================================================================================================






'//=========================================================================================//
'//
'//                 GESTIONE DELLE TABELLE ESTERNE O INTERNE   *** FINE ***
'//
'//=========================================================================================//





'//=========================================================================================//
'//
'//                 GESTIONE DELI MODULI ESTERNE O INTERNE   *** inizio ***
'//@GESTIONE@DELLE@MODULI                    @GESTONE@MODULI
'//=========================================================================================//




'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI MODULI DEL DB ESTERNO *** INIZIO ***
'//                     @IMPORTA@OGGETTI@MODULI@MSYS
'//=====================================================================================================================//



'//IMPORTA TUTTI GLI OGGETTI MODULI DEL DB MSYSDBEST - LE ATTIVITA
'//----------------------------------------------------------------------------------------//
'//@01_@IMPORTA@MODULI@DIM                  -------------------> DIM VARIABILI
'//@02_@IMPORTA@MODULI@REST                 -------------------> RESET DELLE VARIABILI
'//@03_@IMPORTA@MODULI@LOG                  -------------------> LOG FILE INIZIALE
'//@04_@IMPORTA@MODULI@COLL                 -------------------> RECUPERO COLLECTION DA IMPORTARE
'//@05_@IMPORTA@@MODULICONTROLLO 1          -------------------> CONTROLLO 1 ESISTONO OGGETTI COLLECTION
'//@06_@IMPORTA@MODULI@CONTROLLO 2          -------------------> CONTROLLO 2 ESISTE PATH VALIDA
'//@07_@IMPORTA@MODULI@CONTROLLO 3          -------------------> CONTROLLO 3 PATH NULL EXIT
'//@08_@IMPORTA@MODULI@CONTROLLO 4          -------------------> CONTROLLO 4 SE DB NULL EXIT
'//@09_@IMPORTA@MODULI@AVVISI+MSG           -------------------> AVVISI E IMPOSTAZIONE MESSAGGIO FINALE
'//@10_@IMPORTA@MODULI@PROGRSESSBAR_01      -------------------> PROGRESS BAR STEP 01 RECUPERO PARAMETRI
'//@11_@IMPORTA@MODULI@CONTROLLO_PROG       -------------------> CONTROLLO SE STIAMO NEL PROGETTO ESTERNO
'//@12_@IMPORTA@MODULI@PROG@CICLO           -------------------> NEL PROGETTO ESTERNO CICLO NEGLI OGGETTI
'//@13_@IMPORTA@MODULI@PROG@CONTR_01        -------------------> CONTROLLO 1 SE ESTE OGGETTO DA IMPORTARE
'//@14_@IMPORTA@MODULI@PROG@IMPORT          -------------------> IMPORTAZIONE DELL'OGGETTO DAL DB ESTERNO
'//@15_@IMPORTA@MODULI@PROGRSESSBAR_02      -------------------> PROGRESS BAR STEP 02 CALCOLO ED INVIO PARAMETRI
'//@15.BIS_@IMPORTA@MODULI@PROGRESSBAR_CALL -------------------> VIENE CHIAMATA LA PROGRESS BAR CON I PARAMETRI
'//@16_@IMPORTA@MODULI@LOG@FILE             -------------------> SE IL FILE NON VIENE TROVATO SCRITTURAZIONE LOG
'//@17_@IMPORTA@MODULI@CLOSE                -------------------> CHIUSURA CON RITARDO DELLA PROGRESS BAR
'//@18_@IMPORTA@MODULI@NOTHING              -------------------> IMPOSTO A NULL TUTTI GLI OGGETTI


Private Sub Cmd_Importa_MODULI_Click()
    
    '//@01_@IMPORTA@MODULI@DIM                  -------------------> DIM VARIABILI
    ' Dichiarazione delle variabili
    '//oggetti MODULI e variabile nome MODULI
    Dim Modules_Collection As Collection                       '//oggetti collezione
    Dim Modules_Name_v As Variant
    Dim deleted_Modules_s As String
    Dim imported_Modules_s As String
    Dim myModules_Name_s As String
    
    '//il database + path
    Dim sourceDB As DAO.Database
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim sourceDBPath As String
    
    '//contatori oggetti: TOTALE OGGETTI PREVISTI; da importare e cancellare
    Dim TotCollection_Modules_i As Integer                          '//totale degli oggetti MODULI della collezione da importare
    Dim importedCount_i As Integer                              '//oggetti importati
    Dim deletedCount_i As Integer                               '//oggetti cancellati
    
    
    '//bool e log file
    Dim Bool1 As Boolean
    Dim MyValue As Integer
    Dim logFile As Integer
    
    On Error GoTo ErrorHandler
    
    '//@02_@IMPORTA@MODULI@REST                 -------------------> RESET DELLE VARIABILI
    '//RESET = Inizializza i contatori e le stringhe di riepilogo
    TotCollection_Modules_i = 0
    importedCount_i = 0
    deletedCount_i = 0
    deleted_Modules_s = ""
    imported_Modules_s = ""
    TotPerc_dbl = 0
    ProgressPercentTwips_dbl = 0
    
    '//@03_@IMPORTA@MODULI@LOG                  -------------------> LOG FILE INIZIALE
    '@FILE@LOG@MODULI: Apre il file di log per scrivere i dettagli dell'importazione apre nella path corrente.
    logFile = FreeFile()
    Open CurrentProject.PATH & "\LOG_IMPORTA_MODULI.TXT" For Output As #logFile
    Print #logFile, "Log Importazione MODULI - " & Now
    Print #logFile, "------------------------------------------------"
    Print #logFile,
    
    '//@04_@IMPORTA@MODULI@COLL                 -------------------> RECUPERO COLLECTION DA IMPORTARE
    'Recupero la collezione di MODULI da importare
      Set Modules_Collection = CollectionModuli_PFunct()
      
    '//@05_@IMPORTA@@MODULICONTROLLO 1          -------------------> CONTROLLO 1 ESISTONO OGGETTI COLLECTION
    '//se non esistono oggetti collection EXIT
    If Modules_Collection.count = 0 Then
        Print #logFile, "Nessuna MODULI da importare."
        Close #logFile
        '//Prima di uscire restto l'oggetto collection
        Set Modules_Collection = Nothing
        Exit Sub
    End If
    
    '//@06_@IMPORTA@MODULI@CONTROLLO 2          -------------------> CONTROLLO 2 ESISTE PATH VALIDA
    '//RECUPERO LA PATH =  Controllo se il percorso del database di origine è valido
    sourceDBPath = PathFile_s_pFunct
    
    '//@07_@IMPORTA@MODULI@CONTROLLO 3          -------------------> CONTROLLO 3 PATH NULL EXIT
    If sourceDBPath = "" Then
        MsgBox "Errore: path e file nulli.", vbCritical
        Print #logFile, "Errore: path e file nulli."
        Close #logFile
        Exit Sub
    End If
    
    '//@08_@IMPORTA@MODULI@CONTROLLO 4          -------------------> CONTROLLO 4 SE DB NULL EXIT
    '//SE NON ESISTE IL FILE DB = Controllo se il file del database di origine esiste
    If Dir(sourceDBPath) = "" Then
        MsgBox "Errore: il file del database di origine non esiste.", vbCritical
        Print #logFile, "Errore: il file del database di origine non esiste."
        Close #logFile
        Exit Sub
    End If
    
    '//@09_@IMPORTA@MODULI@AVVISI+MSG           -------------------> AVVISI E IMPOSTAZIONE MESSAGGIO FINALE
    '//.......................................................................................//
      '//AVVISI MESSAGGIO IMPOSTATO = TRUE - >  Controllo del valore di avviso
      MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
      If MyValue = 0 Then Exit Sub
      
      '//CONTROLLO NOME PROGETTO =  Controllo del nome del progetto
      Bool1 = CheckProjectName_b_PFunct()
      If Bool1 = True Then Exit Sub
      
    '//.......................................................................................//
    
    
        '//@10_@IMPORTA@MODULI@PROGRSESSBAR_01      -------------------> PROGRESS BAR STEP 01 RECUPERO PARAMETRI
        '//@PROGRESBAR_STEP_01=ATTIVO LA PROGRESS BAR CHIAMANDO ANCHE LE 4 FUNZIONI PER RECUPERARE I PARAMETRI  @progress@bar
        '//********************************************************************************************************************//
        '//@PROGRES@BAR@FASE_01_(imposto le variabili e recupero i valori dalle funzioni)
        
                
                '//NOTE: devi creare 4 funzioni per avere i 4 parametri per chiamare la progress bar: _
                     @01)@IMPOSTO@VARIABILE@NUMERO_COLLECTION          ----> recupero il NroCollection_i _
                     @02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl _
                     @03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
                     @04)@FUNZIONE@TIPO_IMPORT                         ----> RESTITUISCE TipoImport_s _
                     @05)@ROUTINE_APRI_LA_PROGRES_BAR                  ----> RESTITUISCE NULL
                     
                    
                      DoEvents                      '//PER CONSENTIRE IL BLOCCO DEBUG
                  
                      '//@01)@IMPOSTO@VARIABILE@NUMERO_COLLECTION          ----> recupero il NroCollection_i _
                             Imposto il totale oggetti MODULI da importare se esistono
                      TotCollection_Modules_i = Modules_Collection.count
             
                    
                    '//@02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl
                       Twips_dbl = Calcola_Twips_pFunct(TotCollection_Modules_i)
                    
                    '//@03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
                     '//chiama la funzione percentuale incremento e ne restituisce la _
                         percentuale come double
                     PercIncr_dbl = PercentualeIncremento_pFunct(TotCollection_Modules_i)             '//calcolo il numero di collection MODULI da controllare
         
                     '//@04)@FUNZIONE@TIPO_IMPORT                         ----> RESTITUISCE TipoImport_s _
                     '//chiama la funzione TIPO IMPORT
                         TipoImport_s = TipoImport_pFunct("\LOG_IMPORTA_MODULI.TXT")
                          
                         '//CONTROLLO_TIPO_IMPORTAZIONE: se la funzione restituisce NON ESISTE esci dalla IMPORTAZIONE
                         If TipoImport_s <> "IMPORTA_MODULI" Then
                            '//chiudo il file log prima di uscire
                            Close #logFile
                            GoTo Exit_Cmd_IMPORTA_MODULI_Click
                         End If
                      
                     '//@05)@ROUTINE_APRI_LA_PROGRES_BAR                  ----> RESTITUISCE NULL
                         '//Apri la MODULI della progressBar
                         '//
                             Call ApriProgressBar
                             ' Aspetta che il MODULI sia effettivamente caricato
                              DoEvents
                         
        '//@PROGRESBAR_STEP_01=ATTIVO LA PROGRESS BAR CHIAMANDO ANCHE LE 4 FUNZIONI PER RECUPERARE I PARAMETRI *** FINE ***
        '//********************************************************************************************************************//
    
    
    '//===========================================================================================================================//
    '//             ATTIVAZIONE CICLO DI IMPORTAZIONE DAL PROGETTO ESTERNO        *** INIZIO ****
    '//===========================================================================================================================//
            
          '//@11_@IMPORTA@MODULI@CONTROLLO_PROG       -------------------> CONTROLLO SE STIAMO NEL PROGETTO ESTERNO
          ' Se siamo su un progetto esterno (Bool1 = False), procedo con l'importazione
          If Bool1 = False Then
                ' Apertura del database di origine e connessione al database corrente
                Set ws = DBEngine.Workspaces(0)
                Set sourceDB = ws.OpenDatabase(sourceDBPath)
                Set db = CurrentDb
                
                
                '//reset
                ProgressPercentTwips_dbl = 0
                
                
                '//@12_@IMPORTA@MODULI@PROG@CICLO           -------------------> NEL PROGETTO ESTERNO CICLO NEGLI OGGETTI
                ' Ciclo attraverso la collezione di MODULI da importare
                For Each Modules_Name_v In Modules_Collection
                    myModules_Name_s = Modules_Name_v
                    
                    '//@13_@IMPORTA@MODULI@PROG@CONTR_01        -------------------> CONTROLLO 1 SE ESTE OGGETTO DA IMPORTARE
                    ' Controllo se la MODULI esiste nel database di origine
                    If ObjectExists_MODULI(sourceDB, acModule, myModules_Name_s) Then
                        ' Se la MODULI esiste nel database corrente, la elimino prima di importarla
                        If ObjectExists_MODULI(db, acModule, myModules_Name_s) Then
                            DoCmd.DeleteObject acmoduli, myModules_Name_s
                            deletedCount_i = deletedCount_i + 1
                            deleted_Modules_s = deleted_Modules_s & myModules_Name_s & vbCrLf
                            Print #logFile, "MODULI eliminata: " & myModules_Name_s
                             
                              
                        End If
                        
                          '//@14_@IMPORTA@MODULI@PROG@IMPORT          -------------------> IMPORTAZIONE DELL'OGGETTO DAL DB ESTERNO
                          ' Importazione della MODULI dal database di origine
                          DoCmd.TransferDatabase acImport, "Microsoft Access", sourceDBPath, acModule, myModules_Name_s, myModules_Name_s
                          importedCount_i = importedCount_i + 1
                          imported_Modules_s = imported_Modules_s & myModules_Name_s & vbCrLf
                          
                          '//SCRIVO SUL @FILE@LOG
                          Print #logFile, "NRO: " & importedCount_i & " MODULI importata: " & myModules_Name_s
                          
                              
                        '//@15_@IMPORTA@MODULI@PROGRSESSBAR_02      -------------------> PROGRESS BAR STEP 02 CALCOLO ED INVIO PARAMETRI
                        '//@PROGRESBAR_STEP_02= CALCOLO PARAMETRI ED INVIO DATI ALLA FUNZIONE DI AGGIORNAMENTO   @progress@bar
                        '//********************************************************************************************************************//
                        '//NOTE: impostazioni e recupero valori dalla funzioni base per la progress bar.
                            
                              '// RECUPERO il messaggio tipo di importazione delle collection e numero di oggetti da importare .
                                TipoImport_s = "MODULI IMPORTATI NRO : " & importedCount_i & " su TOTALE OGGETTI: " & TotCollection_Modules_i       '// STRINGA MESSAGGIO PER IL TIPO DI IMPORTAZIONE E NRO OGGETTI
                                
                                '//LA PERCENTUALE INCREMENTATA = totale di incremento % per ogni oggetto importato
                                TotPerc_dbl = TotPerc_dbl + PercIncr_dbl
                                
                                '//TOTALE % DI INCREMENTO IN TWIPS, con la ProgressPercentTwips_dbl è la variabile che indica la %di incremento per ogni oggetto + i _
                                   twips necessari per ogni elemento (su 10 totale della barra)
                                ProgressPercentTwips_dbl = ProgressPercentTwips_dbl + (PercIncr_dbl + Twips_dbl)          '//la % di accrescimento barra
                                
                                '//TODO: DA CANCELLARE...
                                'Dbl1 = Dbl1 + Twips_dbl                                         '//twips di aumento barra
                                
                                
                                          '//@15.BIS_@IMPORTA@MODULI@PROGRESSBAR_CALL -------------------> VIENE CHIAMATA LA PROGRESS BAR CON I PARAMETRI
                                          '//CHIAMO LA FUNZIONE @AGGIORNA@PROGRESS@BAR
                                          '//.................................................................//
                                          '//note: chiamo la routine di aggiornamento della progress bar che _
                                              puo essere utilizzata per ulteriori integrazioni di calcolo. _
                                              Devo inviare ma devo inviare 5 parametri : _
                                                  TotCollection_Modules_i       = il TOTALE POTENZIALE degli oggetti da importare _
                                                  NroCollection_i           = il numero della collezione per il calcolo; _
                                                  Twips                     = i twip di incremento _
                                                  ProgressPercentTwips_dbl  =  La percentuale di incremento ( ProgressPercentTwips_dbl  = ProgressPercentTwips_dbl + PercIncr_dbl) _
                                                  TipoImport_s              = il tipo di importazione
                                                  
                                              AggiornaProgressBar _
                                                          TotCollection_Modules_i, _
                                                          ProgressPercentTwips_dbl, _
                                                          Twips_dbl, _
                                                          TotPerc_dbl, _
                                                          TipoImport_s
                                          '//.................................................................//
        
                               
                                T = Timer
                                While Timer < T + 0.001
                                    DoEvents  ' Continua a gestire gli eventi durante l'attesa
                                Wend
                          
                                  
                        '//@PROGRESBAR_STEP_02= CALCOLO PARAMETRI ED INVIO DATI ALLA FUNZIONE DI AGGIORNAMENTO   *** FINE ***
                        '//********************************************************************************************************************//
        
            
                        
                    Else
                        
                        '//@16_@IMPORTA@MODULI@LOG@FILE             -------------------> SE IL FILE NON VIENE TROVATO SCRITTURAZIONE LOG
                        ' Se la MODULI non esiste nel database di origine, lo scrivo nel log
                        Print #logFile, "MODULI non trovata nel database di origine: " & myModules_Name_s
                    End If
                
                Next Modules_Name_v
        
        
                                        
                        '//@17_@IMPORTA@MODULI@CLOSE                -------------------> CHIUSURA CON RITARDO DELLA PROGRESS BAR
                        '//CHIUSURA CON RITARDO DELLA BARRA DI PROGRESSIONE @progress@bar
                        '//---------------------------------------------------------------------------//
                    
                                        T = Timer
                                        
                                        '// RITARDO CHIUSURA DI 5 SECONDI
                                        While Timer < T + 5
                                            DoEvents  ' Continua a gestire gli eventi durante l'attesa
                                        Wend
                                        
                                        ' Chiudi la MODULI UTILITA_Msys_Frm21_frmProgressBar
                                        DoCmd.Close acForm, "UTILITA_Msys_Frm21_frmProgressBar", acSaveNo
                                        
                        '//---------------------------------------------------------------------------//
                        
                        
                        
                '//@18_@IMPORTA@MODULI@NOTHING              -------------------> IMPOSTO A NULL TUTTI GLI OGGETTI
                '//---------------------------------------------------------------------------//
                ' Chiudo il database di origine
                        Set frm_obj = Nothing                 ' Rilascia il riferimento al MODULI PROGRES BARR
                      
                      '//chiudo il db
                      sourceDB.Close
                      Set sourceDB = Nothing
                      Set ws = Nothing
                  
                      ' Messaggio finale con riepilogo delle operazioni
                      If Me.ctr_AVVISO_ABILITATO_TXT = True Then
                          MsgBox "Importazione completata." & vbCrLf & _
                              "Modules importate: " & importedCount_i & vbCrLf & imported_Modules_s & _
                              "Modules esistenti cancellate: " & deletedCount_i & vbCrLf & deleted_Modules_s, vbInMODULIation
                      End If
                  
                      ' Scrivo nel file di log il riepilogo finale
                      Print #logFile, "Importazione completata con " & importedCount_i & " Modules importate."
                      Close #logFile
                  
                      ' Pulizia delle variabili del db e della collection Modules
                      Set db = Nothing
                      Set Modules_Collection = Nothing
                      
                '//---------------------------------------------------------------------------//
                
          End If '//If Bool1 = False Then = Progetto esterno
      
    '//===========================================================================================================================//
    '//             ATTIVAZIONE CICLO DI IMPORTAZIONE DAL PROGETTO ESTERNO        *** FINE ****
    '//===========================================================================================================================//
    
    
'// USCITA E GESTIONE ERRORI
'//---------------------------------------------------------------------------------------------------------//

Exit_Cmd_IMPORTA_MODULI_Click:

Exit Sub
    
ErrorHandler:
    ' Gestione degli errori con messaggio a video e scrittura nel log
    MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical
    Print #logFile, "Errore durante l'importazione: " & Err.Description
    If Not sourceDB Is Nothing Then sourceDB.Close
    Close #logFile
    Set sourceDB = Nothing
    Set ws = Nothing
End Sub


'//CONTROLLO_ESISTEMA_MODULI(efficientata la procedura con l'interrogazione sulla tabella MSysObjects per il controllo delle MODULI esterne)
'//FUNZIONE DI CONTROLLO SE ESISTE L'OGGETTO MODULI
Function ObjectExists_MODULI(db As DAO.Database, objType As Integer, objName As String) As Boolean
    Dim rs As DAO.Recordset
    Dim MODULIExists As Boolean
    
    
    '//CONTROLLO OGGETTO
    Debug.Print " Function ObjectExists_MODULI - oggetto della ricerca : objName "
    Debug.Print objName
     
    MODULIExists = False
    On Error Resume Next
    '//FILTRO CONTROLLO ESISTENZA MODULI - CODICE -32761
    Set rs = db.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=-32761 AND Name='" & objName & "'")
    If Not rs.EOF Then
        MODULIExists = True
    End If
    rs.Close
    Set rs = Nothing
    ObjectExists_MODULI = MODULIExists
    On Error GoTo 0
End Function
'//IMPORTA TUTTI GLI OGGETTI MODULI DEL DB MSYSDBEST *** FINE ***
'//----------------------------------------------------------------------------------------//






'//=====================================================================================================================//
'//                     CMD_IMPORTA OGGETTI MODULI DEL DB ESTERNO *** FINE ***
'//                     @IMPORTA@OGGETTI@MODULI@MSYS
'//=====================================================================================================================//





'//=====================================================================================================================//
'//                     DELETE OGGETTI MODULI DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE MODULI HELP
'//=====================================================================================================================//

Private Sub Cmd_Delete_MODULI_Click()
    Dim obj As AccessObject
    Dim Dbs As Object
    Dim ProceduraMessaggioErrore_s As String
    Dim Modules_Collection As Collection
    'Dim Modules_Collection_v As Variant
    Dim moduleExists As Boolean
    'Dim deletedModulesCount_i As Integer
    'Dim deletet_Modules_Count_i As Integer
    'Dim deleted_Modules_s As String
    'Dim imported_Modules_s As String
    
    On Error GoTo Err_Cmd_Delete_MODULI_Click
   
    ' Reset count
    deletedModulesCount_i = 0
    deletet_Modules_Count_i = 0
    deleted_Modules_s = ""
    importedCount = 0
    imported_Modules_s = ""
     
     
     
     
     

                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_ATTIVA_GEST_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE  @CONTROLLO@PROGETTO
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//
         
     
    ' Recupero la collezione moduli da importare
    Set Modules_Collection = CollectionModuli_PFunct()
                        
    ' Ciclo nella collezione per il controllo delle form precaricate
    For Each Modules_Collection_v In Modules_Collection
        DoEvents
        
        Debug.Print "I MODULI nella collezione: " & Modules_Collection_v
        
        
        ' Verifica se il modulo esiste nel database
        moduleExists = False
        Set Dbs = Application.CurrentProject
        
        '//itero nell'insime ALLMODULES
        For Each obj In Dbs.AllModules
        
            '//CONFRONTA SE ESISTE IMPOSTA TRUE
            If obj.Name = Modules_Collection_v Then
                moduleExists = True
                Exit For
            End If
        Next obj
        
        '//TRUE= ESISTE E CANCELLO IL MODULO
        If moduleExists Then
            ' Se il modulo esiste, cancellalo
            DoCmd.DeleteObject acModule, Modules_Collection_v
            deletet_Modules_Count_i = deletet_Modules_Count_i + 1
            deleted_Modules_s = deleted_Modules_s & Modules_Collection_v & vbCrLf
            Debug.Print "Modulo eliminato: " & Modules_Collection_v
        Else
            ' Se il modulo non esiste, considera per l'importazione
            imported_Modules_Count_i = imported_Modules_Count_i + 1
            imported_Modules_s = imported_Modules_s & Modules_Collection_v & vbCrLf
        End If
    Next Modules_Collection_v
                        
    ' Messaggio finale di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
        MsgBox "CONTROLLO MODULI PRECARICATE DA IMPORTARE:" & vbCrLf & _
               "Modules importate: " & imported_Modules_Count_i & vbCrLf & imported_Modules_s & _
               "Modules esistenti cancellate: " & deletet_Modules_Count_i & vbCrLf & deleted_Modules_s, vbInformation
    End If
            
                
                
                
              '//APRI RIQUADRO MODULI INDIPENDENTEMENTE DALLA LORO PRESENZA
              '//-------------------------------------------------------------------------------------------------//
              '//NOTE= Apri la sezione macro indipendentemente dalla presenza di oggetti _
                  Conta solo i moduli visibili, escludendo quelli di sistema. _
                  Se ci sono moduli, apre direttamente la sezione Moduli con DoCmd.SelectObject acModule, , True. _
                  Print Se; non; ci; sono; Moduli, prova; a; selezionare; una; tabella; e; poi; la; sezione; Moduli, per; forzare; l 'apertura del riquadro. _
                  Evita errori con On Error Resume Next in caso di assenza di oggetti.

                  Dim db As DAO.Database
                  Dim doc As Document
                  Dim moduliVisibili As Integer
      
                  Set db = CurrentDb
                  moduliVisibili = 0 ' Inizializza il contatore
      
                  ' Conta solo i moduli visibili nel database
                  For Each doc In db.Containers("Modules").Documents
                    If Left(doc.Name, 1) <> "~" Then ' Esclude eventuali moduli di sistema nascosti
                      moduliVisibili = moduliVisibili + 1
                    End If
                  Next doc
      
                  ' Se ci sono moduli visibili, apri la sezione Moduli
                  If moduliVisibili > 0 Then
                    DoCmd.SelectObject acModule, , True
                  Else
                    ' Se non ci sono moduli, prova a selezionare un altro oggetto per forzare l'apertura del riquadro
                    On Error Resume Next
                    DoCmd.SelectObject acTable, , True ' Seleziona una tabella (se esiste)
                    DoCmd.SelectObject acModule, , True ' Ora prova a selezionare la sezione Moduli
                    On Error GoTo 0
                  End If
      
                  ' Pulizia memoria
                  Set doc = Nothing
                  Set db = Nothing
                          
                          
                    
              '//APRI RIQUADRO MODULI INDIPENDENTEMENTE DALLA LORO PRESENZA *** FINE ***
              '//-------------------------------------------------------------------------------------------------//
            


            
     ' Messaggio finale di riepilogo
    If Me.ctr_AVVISO_ABILITATO_TXT = True Then
   
        ' Optional: display a message indicating that the modules have been deleted
        MsgBox "Tutti i moduli standard sono stati cancellati. Totale moduli cancellati: " & deletet_Modules_Count_i, vbInformation, "Moduli Cancellati"
    
    End If

Exit_Cmd_Delete_MODULI_Click:
    Exit Sub

Err_Cmd_Delete_MODULI_Click:
    MsgBox Err.Description & " - Error Message -> : " & ProceduraMessaggioErrore_s
    Debug.Print ProceduraMessaggioErrore_s
    Stop
    Resume Exit_Cmd_Delete_MODULI_Click

End Sub

'//DELETE MODULI        ***FINE ***
'//==========================================================================================================//



'//Cmd_DELETE_MODULI_HELP_Click  *** INIZIO ***
'//-----------------------------------------------------------------------------------//

Private Sub Cmd_DELETE_MODULI_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                           
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
     
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_Delete_MODULI.pdf "
               
               
             '//CHIAMO LA FUNZIONE INTERNA PER L'APERTURA DEI FILE
                Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
                 
           
    '//--------------------------------------------------------------------------------//--------//




End Sub

'//Cmd_DELETE_MODULI_HELP_Click  *** FINE ***
'//-----------------------------------------------------------------------------------//


'//=====================================================================================================================//
'//                     DELETE OGGETTI MODULI DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE MODULI HELP
'//=====================================================================================================================//






'//=========================================================================================//
'//
'//                 GESTIONE DEI MODULI ESTERNI O INTERNI   *** FINE ***
'//
'//=========================================================================================//
















'//*********************************************************************************************//
'//***
'//***                                            VARIE FUNZIONI
'//*********************************************************************************************//

'//CONTROLLO DEL PROGETTO CORRENTE *** INIZIO *** Function CheckProjectName_b_PFunct() As Boolean
'//----------------------------------------------------------------------------------------//
'//NOTE: QUESTA ROUTINE  controlla se il progetto in cui si trova la form carica oggetti, _
    è quello di origine ossia il database MSYS_ATTIVA_GEST_OGGETTI.mdb. Interroganto la _
    proprita name se ottengo il nome del DB DI ORIGINE non eseguo l'importazione perchè altrimenti _
    distrutto il db di origine. QUESTE FUNZIONI SI POSSONO UTILIZZARE SOLO IN DB DIVERSI DA _
    QUELLO DI ORIGINE.  @PROGETTO@SORGENTE_(il nome del progetto sorgente che impedesce le attivia MSYS_ATTIVA_GEST_OGGETTI.mdb)
'//CODICE: @CHECK@PROJECT@PROGETTO@SORGENTE

'// CHIAVI_SOSTITUIBILI .......................................................***INIZIO *** _
        PROGETTO ---------> MSYS_ATTIVA_GEST_OGGETTI.mdb

'// CHIAVI_SOSTITUIBILI .......................................................***FINE ***


Public Function CheckProjectName_b_PFunct() As Boolean
    
    '//CHIAMO IL CONTROLLO DEL PROGETTO
    '//..........................................................//
    '//Note  : True = siamo nel progetto originale non possiamo _
               importare nessun oggetto perche distruggiamo il db corrente. _
               Fale = siamo in un progetto esterno possiamo caricare tutto.
                   
        '//solo se FALSE attivo la procedura
        'Bool1 = CheckProjectName_b_PFunct
    '//..........................................................//
    
        
    ' Variabile per memorizzare il nome del progetto corrente
    Dim projectName As String
    
    ' Ottieni il nome del progetto corrente
    projectName = Application.CurrentProject.Name
    
    
    '//2024.08.10 =CHECKPROJECT_MODIFICA_CONTROLLO_CON_RAFFORZAMENTO
    ' Controlla se il nome del DATABASE QUELLO ORIGINARIO OSSIA "MSYS_ATTIVA_GEST_OGGETTI.mdb"
    If projectName = "MSYS_ATTIVA_GEST_OGGETTI.mdb" Then
        MsgBox "USCITA DALLA FUNZIONE: ATTENZIONE CI TROVIAMO NEL PROGETTO ORIGINALE E NON POSSIAMO CARICARE NE CANCELLARE " & _
               "NESSUN OGGETTO SUL PROGETTO DI PARTENZA DENOMINATO : " & projectName & vbCrLf & _
               "La funzione puo essere usata solo in un progetto Esterno.", vbCritical
               
            Debug.Print "CONTROLLO PROGETTTO: ATTENZIONE IL PROGETTO E' QUELLO ORIGINALE  : " & projectName & " NON POSSIAMO CARICARE NESSUNO OGGETTO!!"
            '//TRUE = progetto originale  NON SI PUO CARICARE NESSU OGGETTO
            CheckProjectName_b_PFunct = True
    Else
            '//FALSE = progetto diverso dall'originale  possiamo importare gli oggetti."
            CheckProjectName_b_PFunct = False
            Debug.Print "Il progetto si chiama diversamente: " & projectName & " POSSIAMO CARICARE GLI OGGETTI !!"
    End If
    
    
End Function

'//CONTROLLO DEL PROGETTO CORRENTE *** FINE *** Function CheckProjectName_b_PFunct() As Boolean
'//----------------------------------------------------------------------------------------//







'//AVVISO DI CONTINUAZIONE *** INIZIO ***
'//----------------------------------------------------------------------------------------//

Function CheckAvviso_i_PFunct(par_Bool As Boolean) As Integer
    
    '//il messaggio viene attivato solo se il parametro è TRUE
    If par_Bool = True Then
      '//AVVISO  - @MSGBOX
      Message = "ATTENZIONE QUESTO COMANDO DI IMPORTAZIONE VALE SOLO PER II PROGETTI DIVERSI  " & Chr$(13) & _
                "DAL PROGETTO MADRE MSYS_ATTIVA_GEST_OGGETTI.mdb; PER ANNULLARE INSERIRE 0"               ' Imposta il messaggio.
                Title = "AVVISO DI CONTINUAZIONE"           ' Imposta il titolo.
                Default = "1"                               ' Imposta il valore predefinito.
            
                ' Visualizza la finestra di dialogo nella posizione 100,100
                
                      MyMsg_V = InputBox(Message, Title, Default, 100, 100)
            
                
                
                
                '//CONTROLLO SE LA SCELTA = 1 O SCELTA = 0
                If MyMsg_V <> "" Or MyMsg_V <> Null Then
                    '//se scelta =0 oppure = 1 restisce un valore
                    CheckAvviso_i_PFunct = MyMsg_V
                Else
                    '//se annulla restituisce stringa = null e quindi restituisce 0 ANNULLATA
                    CheckAvviso_i_PFunct = 0
                    MyMsg_V = 0             ' 0 = SCELTA ANNULLATA ANCHE SE HAI FATTO ANNULLA
                End If
                
                
                '//AVVISO DI USCITA PERCHE SCELTO 0
                If MyMsg_V = 0 Then
                
                        '//AVVISO  DI USCITA- @MSGBOX
                        
                        MsgBox "USCITA DALLA ROUTINE", vbInformation, "AVVISO DI USCITA"
                        
   
                End If
                       
                
                
    Else
            '//CON IL MESSAGGIO DISABILITATO SI INTENDE SEMPRE 1 = ATTIVARE LA PROCEURA
             MyMsg_i = 1
            CheckAvviso_i_PFunct = MyMsg_i
        
    End If
            
End Function

'//AVVISO DI CONTINUAZIONE *** FINE ***
'//----------------------------------------------------------------------------------------//



'//APRI_IL_FILE_PDF - FUNZIONE INTERNA DI APERTURA DEI FILE -
'//==================================================================================================================//
'//METODO DI APERTURA DI UN PROGRAMMA ESTERNO O DI UN COMANDO DOS MEDIANTE "WScript.Shell"
'//PARAMETRI            -> passo 2 stringhe per parametro, la path e il nome del file.pdf
'//VALORE_DI_RITORNO    -> Nulla
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.00

Public Function ApriFilePdf_Interna_s_PFunct(par_Path_s As String, par_NameFile_s As String, par_IDGestione_lng As Long) As String

Dim Path_programma_s As String
Dim NameFile_s As String
Dim Stringa1 As String

On Error GoTo ApriFilePdf_Interna_s_PFunct_Err
        
    
'//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
'//--------------------------------------------------------------------------------//--------//
'//NOTE                 -> Apro il file di tipo doc, zip o pdf
'//CODICE               -> Function ApriFilePdf_Interna_s_PFunct.01.01
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
            
            
         '//chiamo la sub con i parametri =   'CALL (ApriFilePdf_Interna_s_PFunct Path_programma_s, NameFile_s, IDGestione_lng)
            'Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
             
        
'//--------------------------------------------------------------------------------//--------//
                        
         
                        '//APRO FILE PDF SPECIFICO
                        '//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>//
                        '//Note       : Apro il file PDF speficico indicando come paramentro la path completa
                        '//           incluso il file PDF.
                          
                          '//La path del file = Unisco la path + il file+estensione sopra individuata
                          par_Path_s = par_Path_s & par_NameFile_s & MyFileEstensione_s
                                 
                          '//controllo
                          Debug.Print par_Path_s
                                 
                          '//CREO LA SHELL come oggetto, attivo il comando "%comspec% /c start " e gli assegno la path per estesa con il nome del file
                          '// perchè Shell lancia un EXE e non il PDF, quindi è necessario costruire la stringa in modo da far lanciare
                          '// prima la sessione dos e poi il pdf..
                          '// La stringa di comando ("%comspec% /c start ") deve essere unica : comando shell + path + file
                          Set WshPDF = CreateObject("wscript.shell")
                          
                          
                          '//CONTROLLO
                          Debug.Print " ---------------------CONTROLLO PATH -----------------------------------"
                          Debug.Print par_Path_s
                          Debug.Print
                          
                          '// Unisco il comando pdf start + path definitiva
                          Stringa1 = "%comspec% /c start " & par_Path_s                               '//Comando di apertura
                          '//Attivo il comando
                          WshPDF.Run Stringa1
                                  
                                  
                          '//libero la memoria dalle variabili e dagli oggetti creati
                          WshPDF = Null
                          Stringa1 = vbNull
                          MyFileEstensione_s = vbNull
                          MyPath_s = vbNull
                            
                            '//sistemare ??? invio tasti ???
                            Lng = 0
                            For Lng = 1 To 90000000 Step 1
                            
                            Next Lng
                            
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
      
      
ApriFilePdf_Interna_s_PFunct_Exit:
    Exit Function

ApriFilePdf_Interna_s_PFunct_Err:
    MsgBox Error$
    Resume ApriFilePdf_Interna_s_PFunct_Exit

End Function

'//APRI_IL_FILE_PDF - FUNZIONE INTERNA DI APERTURA DEI FILE -         *** FINE ***
'//==================================================================================================================//





'//*** INIZIO ***
'//APRI FILE @APRO@ANALISI@SUBLIME.TEXT  @CMD@03@CMD.03@CMD.3    @HELP@SUBLIME@TEXT@ANLISI
'//========================================================================================================================================//

Private Sub Cmd_HELP_Click()

On Error GoTo Cmd_HELP_Err
    
      
      
      
                     
        
        '//APRO SORGENTE FORM FILTRO 66 CAMPI
        '//.....................................................................................................
        '//Note           : Tramite una Select vengono individuati i valori da restiuire.
                   
            '//reset
            '//Imposto i parametri
            ProceduraAttivaEseguita_s = ""
            ProceduraMessaggioErrore_s = ""
            
            MyFile_s = ""
            MyPath_s = ""
            
            '//Imposto i parametri
            ProceduraAttivaEseguita_s = ""
            ProceduraMessaggioErrore_s = "Errore nella procedura - apri FILE GESTIONE """ & MyPath_s & MyFile_s
            
            
        
                    
                '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF
                '//--------------------------------------------------------------------------------//--------//
                '//NOTE                 -> Apro il file di tipo doc, zip o pdf
                '//CODICE               -> Function ApriFilePdf.01.01
                '//PARAMETRI            -> par_Path_s         = PATH _
                                        -> par_NameFile_s               = NOME FILE _
                                        -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                        
                        '//imposto i parametri - LA PATH
                        MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\MSYS\MDB\"
                        '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
                        MyFile_s = "Project_MSYSDbEst_GE_OBJECT_ANALISI.sublime-project "
                
                            '//CHIAMO LA FUNZIONE INTERNA PER L'APERTURA DEI FILE
                            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
             
                        
                '//--------------------------------------------------------------------------------//--------//
                        
                 
          
 

           
        '//*** fine ***
        '//APRO IL FILE PASSANDO COME PARAMETRO PATH E NOME
        '//.....................................................................................................
 
                    







        

'//*** fine ***
'//APRI FILE @APRO@ANALISI@SUBLIME.TEXT
'//###########################################################################################################################//


        
    '//........................................................//


'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Cmd_HELP_Exit:
    Exit Sub

Cmd_HELP_Err:
    MsgBox Error$
    Resume Cmd_HELP_Exit

End Sub


'//*** FINE ***
'//APRI PATH
'//========================================================================================================================================//







'//ESEGUI QUERY SULLA TABELLA DI SISTEMA MSysObjects PER INDIVIDUARE GLI OGGETTI ATTIVI (TABELLE, QUERY FORM REPORT)
'//========================================================================================================================================//
'//Recupero del percorso del database esterno: La funzione PathFile_s_pFunct viene chiamata per ottenere il percorso del database esterno. _
    Se il percorso è vuoto, la funzione mostra un messaggio di errore e esce. Apertura del database esterno: Il database esterno viene aperto _
    utilizzando DBEngine.Workspaces(0).OpenDatabase. Esecuzione della query su MSysObjects: Viene eseguita una query su MSysObjects per _
    ottenere tutti gli oggetti che non sono nascosti (Flags=0) e che sono di tipo tabella, query, form, modulo, report o macro. _
    Ciclo attraverso i risultati: La funzione cicla attraverso i risultati della query e stampa il nome e il tipo di ciascun _
    oggetto nella finestra di debug. La variabile objectTypeDescription viene utilizzata per fornire una descrizione leggibile del tipo di oggetto. _
    Chiusura del recordset e del database: Il recordset e il database esterno vengono chiusi al termine della funzione. _
    Messaggio informativo: Viene visualizzato un messaggio informativo per indicare che la lista degli oggetti è stata stampata nella finestra di debug.


Private Sub ListExternalDBObjects()
    Dim sourceDBPath As String
    Dim externalDB As DAO.Database
    Dim rs As DAO.Recordset
    Dim objectType As Integer
    Dim objectName As String
    Dim objectTypeDescription As String
    
    ' Liste per ogni tipo di oggetto
    Dim tables As String
    Dim queries As String
    Dim forms As String
    Dim modules As String
    Dim reports As String
    Dim macros As String
    Dim others As String
    
    ' Recupero la path e il file dal database esterno
    sourceDBPath = PathFile_s_pFunct
    
    If sourceDBPath = "" Or IsNull(sourceDBPath) = True Then
        MsgBox "ATTENZIONE ERRORE: path e file sono null, uscita dalla routine", vbCritical, "MSG CONTROLLO PATH E FILE VALORIZZATI"
        Exit Sub
    End If
    
    ' RESET
    Me.TXT_OBJECT_ESTERNI.Value = ""
    Me.TXT_OBJECT_ESTERNI.Requery
    
    ' Apro il database esterno
    Set externalDB = DBEngine.Workspaces(0).OpenDatabase(sourceDBPath)
    
    ' Eseguo la query su MSysObjects per ottenere tutti gli oggetti
    Set rs = externalDB.OpenRecordset("SELECT Name, Type FROM MSysObjects WHERE Flags=0 AND Type IN (1, 4, 5, -32761, -32764, -32766)")
    
    ' Ciclo attraverso i risultati e raccolgo gli oggetti nelle liste appropriate
    Do While Not rs.EOF
    
        objectName = rs!Name
        objectType = rs!Type
        
        
        '// Chr$(13) =  è una funzione in VBA che restituisce il carattere corrispondente al codice ASCII specificato. _
            In questo caso, 13 è il codice ASCII per il carattere di ritorno a capo (Carriage Return, CR), che viene spesso _
            utilizzato in combinazione con Chr$(10) (line feed, LF) per creare una nuova linea su molte piattaforme. Tuttavia, per creare una _
            nuova linea in modo più robusto e compatibile, è meglio usare vbCrLf, che combina entrambi i caratteri di controllo (Chr$(13) & Chr$(10)). _
            vbCrLf è una costante predefinita in VBA che rappresenta una nuova linea e garantisce che il testo appaia correttamente formattato _
            su tutte le piattaforme. _
            Chr$(10) = Il Line Feed (LF) è un carattere di controllo utilizzato per indicare la fine di una linea di testo e l'inizio di una nuova. _
            Nel sistema ASCII, il codice per il Line Feed è 10. Questo carattere è comunemente utilizzato nei sistemi Unix e Linux _
            per terminare le linee di testo. Quando un Line Feed viene incontrato in un testo, il cursore o il puntatore si _
            sposta all 'inizio della riga successiva. Tuttavia, il Line Feed da solo non riporta il cursore all'inizio della riga corrente, _
            per questo motivo spesso viene combinato con il Carriage Return (CR, codice ASCII 13), che invece riporta il _
            cursore all 'inizio della riga corrente.0
            
        ' Descrivo il tipo di oggetto
        Select Case objectType
            Case 1
                objectTypeDescription = "Tabella"
                tables = tables & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case 4
                objectTypeDescription = "Query"
                queries = queries & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case 5
                objectTypeDescription = "Form"
                forms = forms & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case -32761
                objectTypeDescription = "Modulo"
                modules = modules & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case -32764
                objectTypeDescription = "Report"
                reports = reports & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case -32766
                objectTypeDescription = "Macro"
                macros = macros & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
            Case Else
                objectTypeDescription = "Altro"
                others = others & "Nome: " & objectName & ", Tipo: " & objectTypeDescription & " - codice : " & objectType & vbCrLf
        End Select
        
        rs.MoveNext
    Loop
    
    ' Stampo le liste in ordine
    Debug.Print "Oggetti nel database esterno (" & sourceDBPath & "):"
    Debug.Print "-----------------------------------------------"
    Debug.Print
    
    If tables <> "" Then Debug.Print tables
    Debug.Print
    If queries <> "" Then Debug.Print queries
    Debug.Print
    If forms <> "" Then Debug.Print forms
    Debug.Print
    If modules <> "" Then Debug.Print modules
    Debug.Print
    If reports <> "" Then Debug.Print reports
    Debug.Print
    If macros <> "" Then Debug.Print macros
    Debug.Print
    If others <> "" Then Debug.Print others
    Debug.Print
    
    ' Imposto il valore del controllo di testo
    Me.TXT_OBJECT_ESTERNI.Value = tables & vbCrLf & queries & vbCrLf & forms & vbCrLf & modules & vbCrLf & reports & vbCrLf & macros & vbCrLf & others
    
    ' Chiudo il recordset e il database
    rs.Close
    externalDB.Close
    
    MsgBox "Lista degli oggetti nel database esterno stampata nella finestra di debug.", vbInformation, "Completato"
End Sub


'//*** FINE ***
'//ESEGUI QUERY SULLA TABELLA DI SISTEMA MSysObjects PER INDIVIDUARE GLI OGGETTI ATTIVI (TABELLE, QUERY FORM REPORT)
'//========================================================================================================================================//








'//=====================================================================================================================//
'//                     DELETE TUTTI GLI OGGETTI DEL DB ESTERNO IN TABELLA *** INIZIO ***
'//                     + DELETE TUTTO HELP
'//=====================================================================================================================//




'//CANCELLA TUTTO
Private Sub Cmd_DELETE_TUTTO_Click()
        
        
    
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_ATTIVA_GEST_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE  @CONTROLLO@PROGETTO
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '//=====================================================================================================================//
                    
                    '//2024.08.03_MODIFICA_CMB_DELETE_TUTTO = la modifica riguarda la routine di attivazione
                    '//ATTIVO LA CANCELLAZIONE DI TUTTI GLI OGGETTI SE OGNI  BUTTON E' ABILITATO CON TRUE '
                                            
                        If Me.Cmd_Delete_FORM.Enabled = True Then
                            Cmd_Delete_FORM_Click
                        End If
                        If Me.Cmd_Delete_MODULI.Enabled = True Then
                            Cmd_Delete_MODULI_Click
                        End If
                        If Me.Cmd_Delete_QUERY.Enabled = True Then
                            Cmd_Delete_QUERY_Click
                        End If
                        If Me.Cmd_Delete_REPORT.Enabled = True Then
                            Cmd_Delete_REPORT_Click
                        End If
                        If Me.Cmd_Delete_TABLE.Enabled = True Then
                            Cmd_Delete_TABLE_Click
                        End If
                        
                        If Me.Cmd_Delete_MACRO.Enabled = True Then
                            Cmd_Delete_MACRO_Click
                        End If

                        '//FINE CON MESSAGGIO IN CASELLA DI TESTO'
                        Me.TXT_OBJECT_ESTERNI = vbCrLf & vbCrLf & "  //********************************************************************// " & vbCrLf & " ......  FINE CANCELLAZIONE DI TUTTI GLI OGGETTI "
                        Me.TXT_OBJECT_ESTERNI.Requery
                        

End Sub


Private Sub Cmd_Delete_Tutto_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_DELETE_TUTTO.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa REPORT _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub






'//=====================================================================================================================//
'//                     DELETE TUTTI GLI OGGETTI DEL DB ESTERNO IN TABELLA *** FINE ***
'//                     + DELETE TUTTO HELP
'//=====================================================================================================================//






Private Sub Cmd_IMPORTA_TUTTO_Click()



                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** INIZIO ***
                '//=====================================================================================================================//
                '//NOTE: chiamo la funzione che controlla se il progetto corrente è il progetto originale, infatti _
                         non si possono caricare gli oggetti nel PROGETTO MSYS_ATTIVA_GEST_OGGETTI.mdb perhè è originale _
                         e vi è una distruzione degli stessi. LA FUNZIONE CONTROLLO SE E' UN PROGETTO ESTERNO AMMISSIBILE _
                         per caricare gli oggetti table, query , form ecc. da importare. _
                         01) CONTROLLO = MSG DI AVVISO CON POSSIBILITA DI ANNULLARE _
                         02) CONTROLLO = CONTROLLO SE STIAMO NEL PROGETTO ORIGINALE  E QUINDI ESCLUSIONE.
                                 
                           '//01) PRIMO CONTROLLO _
                             MESSAGGIO DI IMPORTAZIONE 0 = ANNULLA  1 = CONTINUA
                           '//..........................................................//
                             '//Attenzione L'AVVISO produce il messaggio solo se il _
                                parametro è TRUE altrimenti il messaggio viene escluso.
                             MyValue = CheckAvviso_i_PFunct(Me.ctr_AVVISO_ABILITATO_TXT)
                             
                             If MyValue = 0 Then Exit Sub
                           '//..........................................................//
                                 
                           '//02) SECONDO CONTROLLO
                           '//CHIAMO IL CONTROLLO DEL PROGETTO
                           '//..........................................................//
                           '//Note  : True = siamo nel progetto originale non possiamo _
                                      importare nessun oggetto perche distruggiamo il db corrente. _
                                      False = siamo in un progetto esterno possiamo caricare tutto.
                                          
                               '//SE VIENE RESTITUITO TRUE è FUORI PROGETTO PUO CARICARE, SE FALSE EXIT PERCHE PROGETTO ORIGINALE  @CONTROLLO@PROGETTO
                               Bool1 = CheckProjectName_b_PFunct
                               
                               If Bool1 = True Then Exit Sub
                           '//..........................................................//
                
                '//=====================================================================================================================//
                '//                             STEP_02 = @CONTROLLO@PROGETTO
                '//CHIAMO IL CONTROLLO DEL PROGETTO  - 01 E 02) CONTROLLO - *** FINE ***
                '

                '//CHIAMO LE PROCEDURE MA LE TABELLE DEVONO ESSERE CHIAMATE PER PRIMA _
                   e le query per ultimo, altrimenti le query non si definiscono benei i campi.'
                
                '//2024.08.03_MODIFICA_CMB_IMPORTA_TUTTO = la modifica riguarda la routine di attivazione
                '//ATTIVO LA IMPORTAZIONE DI TUTTI GLI OGGETTI SE OGNI  BUTTON E' ABILITATO CON TRUE '
                    
                    If Me.Cmd_Importa_TABELLE.Enabled = True Then
                        Cmd_Importa_TABELLE_Click
                        
                    End If
                    
                    If Me.Cmd_Importa_FORM.Enabled = True Then
                        Cmd_Importa_FORM_Click
                    End If
                    
                    If Me.Cmd_Importa_MODULI.Enabled = True Then
                        Cmd_Importa_MODULI_Click
                    End If
                    
                    If Me.Cmd_Importa_REPORT.Enabled = True Then
                        Cmd_Importa_REPORT_Click
                    End If
                    
                    If Me.Cmd_Importa_MACRO.Enabled = True Then
                        Cmd_Importa_MACRO_Click
                    End If
                    
                    If Me.Cmd_Importa_QUERY.Enabled = True Then
                        Cmd_Importa_QUERY_Click
                    End If


                     '//FINE CON MESSAGGIO IN CASELLA DI TESTO'
                        Me.TXT_OBJECT_ESTERNI = vbCrLf & vbCrLf & "  //********************************************************************// " & vbCrLf & " ......  FINE IMPORTAZIONE DI TUTTI GLI OGGETTI "
                        Me.TXT_OBJECT_ESTERNI.Requery
                       

End Sub




Private Sub Cmd_IMPORTA_Tutto_HELP_Click()

    '//APRO FILE PDF SPECIFICO CON IL COMANDO OGGETTO PDF - @APRI.FILE
    '//--------------------------------------------------------------------------------//--------//
    '//NOTE                 -> Apro il file di tipo doc, zip o pdf
    '//CODICE               -> Function ApriFilePdf.01.01
    '//PARAMETRI            -> par_Path_s         = PATH _
                            -> par_NameFile_s               = NOME FILE _
                            -> par_IDGestione_lng           = ID FILE DA RICERCARE per futuri utilizzi
                            
      '//IMPOSTO LE VARIABILI PATH E FILE
      '//
      Dim MyFile_s As String
      '//
      Dim MyPath_s  As String
      
            '//imposto i parametri - LA PATH
            '
            MyPath_s = "c:\CASA\LINGUAGGI\ACCESS\PROGETTI_MDB\MSYS_OGGETTI\OBJECT\HELP\"
            '//IL FILE = attenzione al file ho lasciato un spazio perche a volte non funziona senza
            '
            MyFile_s = "HELP_Cmd_IMPORTA_TUTTO.pdf "
                
                
                       
         '//chiamo la funzione INTERNA DI APERTURA DEL FILE con i parametri =   'CALL (ApriFilePdf Path_programma_s, NameFile_s, IDGestione_lng) _
            Il vantaggio della procedura interna è quello di evitare di portare dietro al progetto IMPORTA OGGETTI incluso in questa REPORT _
            anche i moduli e classi che potrebbero essere oggetto di duplicazione con l'importazione e cancellazione in caso di resettamento _
            totale.
            Call ApriFilePdf_Interna_s_PFunct(MyPath_s, MyFile_s, 0)
      
            
    '//--------------------------------------------------------------------------------//--------//

End Sub



'//*********************************************************************************************//
'//***
'//***                                            VARIE FUNZIONI *** FINE ***
'//*********************************************************************************************//












'//***********************************************************************************************************************//
'//                             @ABILITAZIONE E @DISABILLITAZIONI        *** INIZIO *****
'//
'//***********************************************************************************************************************//




'//DISABILITO LA PATH DEL DB
Private Sub Disabilito_sourceDBPath_Sub()
    
    '//DISABILITO LA PATH ESTERNA
    '//........................................//
       'Call DisabilitosourceDBPath_s_Txt_Sub
    '//........................................//
    
    '//disabilito la path
    Me.sourceDBPath_s_Txt.Enabled = False
    
    
End Sub

'//ABILITO LA PATH DEL DB
Private Sub Abilito_sourceDBPath_Sub()
    
    '//ABILITO LA PATH ESTERNA
    '//........................................//
       'Call Abilito_sourceDBPath_Sub
    '//........................................//
    
    '//Abilito la path
    Me.sourceDBPath_s_Txt.Enabled = True
    
    
End Sub




Private Sub DisabilitaButton_Sub()
    
    
    '//DISABILITO I BUTTON
    '//........................................//
       'Call DisAbilitaButton_Sub
    '//........................................//
    
    
    '//@DISABILITO@BUTTON_(disabilito tutti i comandi button se ci troviano nel progetto originale MSYS_ATTIVA_GEST_OGGETTI.mdb)
    Me.Cmd_Delete_FORM.Enabled = False
    Me.Cmd_Delete_MACRO.Enabled = False
    Me.Cmd_Delete_MODULI.Enabled = False
    Me.Cmd_Delete_QUERY.Enabled = False
    Me.Cmd_Delete_REPORT.Enabled = False
    Me.Cmd_Delete_TABLE.Enabled = False
    Me.Cmd_DELETE_TUTTO.Enabled = False
    Me.Cmd_Importa_FORM.Enabled = False
    Me.Cmd_Importa_MACRO.Enabled = False
    Me.Cmd_Importa_MODULI.Enabled = False
    Me.Cmd_Importa_QUERY.Enabled = False
    Me.Cmd_Importa_REPORT.Enabled = False
    Me.Cmd_Importa_TABELLE.Enabled = False
    Me.Cmd_IMPORTA_TUTTO.Enabled = False
    
    
End Sub


'//ABILITAZIONE@BUTTON
'//NOTE : se corrisponde a diverse voci della combinata abilita il button
Private Sub AbilitaButton_Sub()
    
    '//ABILITO I BUTTON
    '//........................................//
    '//note : chiamo la funzione di abilitazione _
        dei button se corrispone ad una voce diversa _
        da ELENCO PROGETTI
       'Call AbilitaButton_Sub
    '//........................................//
 
            '//@ABILITO@BUTTON_(ABILITO tutti i comandi button se ci troviano nel progetto originale MSYS_ATTIVA_GEST_OGGETTI.mdb)
            Me.Cmd_Delete_FORM.Enabled = True
            Me.Cmd_Delete_MACRO.Enabled = True
            Me.Cmd_Delete_MODULI.Enabled = True
            Me.Cmd_Delete_QUERY.Enabled = True
            Me.Cmd_Delete_REPORT.Enabled = True
            Me.Cmd_Delete_TABLE.Enabled = True
            Me.Cmd_DELETE_TUTTO.Enabled = True
            Me.Cmd_Importa_FORM.Enabled = True
            Me.Cmd_Importa_MACRO.Enabled = True
            Me.Cmd_Importa_MODULI.Enabled = True
            Me.Cmd_Importa_QUERY.Enabled = True
            Me.Cmd_Importa_REPORT.Enabled = True
            Me.Cmd_Importa_TABELLE.Enabled = True
            Me.Cmd_IMPORTA_TUTTO.Enabled = True
            
            
                
            '//ABILITO ANCHE LA COMBINATA 2 = CMB_02
            Me.Cmb_02_txt.Enabled = False
            Me.Cmb_02_txt.Value = "scelta"

    
End Sub





'//ABILITAZIONE E DIABLITAIONE DEI BUTTON , TABLE, FORM, QUERY MACRO REPORT *** INIZIO ***
'//-------------------------------------------------------------------------------------------------------------//
'2024.08.03_Abilito_Table+Report+Macro+Query+Moduli = creato routine per
'l 'abilitazione delle query, macro, report, table ecc.
'2024.08.03_DisabilitoAbilito_Table+Report+Macro+Query+Moduli = creato routine per
'la DISABILITAZIONE  delle query, macro, report, table ecc.

'//ABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_FORM_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA E CANCELLA FORM
            '//@ABILITO@BUTTON_(abilito il @comando@form @cmd@form@delete.abilitato + cmd@form@importa.abilitato)
            Me.Cmd_Delete_FORM.Enabled = True
            Me.Cmd_Importa_FORM.Enabled = True
          
End Sub



'//DISABILITAZIONE@BUTTON@CMB_02
'//NOTE : DISABILITA LA COMBINATA 02
Private Sub DisabilitaButton_CMB_02_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA TUTTO
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@CMB_02)
            Me.Cmb_02_txt.Enabled = False
          
End Sub


'//BILITAZIONE@BUTTON@CMB_02
'//NOTE : ABILITA LA COMBINATA 02
Private Sub AbilitaButton_CMB_02_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA TUTTO
            '//@DISABILITO@BUTTON_(ABILITO il @comando@CMB_02)
            Me.Cmb_02_txt.Enabled = True
          
End Sub




'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub DisabilitaButton_FORM_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA E CANCELLA FORM
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@form @cmd@form@delete.disabilitato + cmd@form@importa.disabilitato)
            Me.Cmd_Delete_FORM.Enabled = False
            Me.Cmd_Importa_FORM.Enabled = False
          
End Sub


'//ABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_TABLE_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA E CANCELLA TABLE
            '//@ABILITO@BUTTON_(abilito il @comando@TABLE @cmd@TABLE@delete.abilitato + cmd@TABLE@importa.abilitato)
            Me.Cmd_Delete_TABLE.Enabled = True
            Me.Cmd_Importa_TABELLE.Enabled = True
          
End Sub



'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub DisabilitaButton_TABLE_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA E CANCELLA TABLE
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@TABLE @cmd@TABLE@delete.disabilitato + cmd@TABLE@importa.disabilitato)
            Me.Cmd_Delete_TABLE.Enabled = False
            Me.Cmd_Importa_TABELLE.Enabled = False
          
End Sub




'//ABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_QUERY_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA E CANCELLA QUERY
            '//@ABILITO@BUTTON_(abilito il @comando@QUERY @cmd@QUERY@delete.abilitato + cmd@QUERY@importa.abilitato)
            Me.Cmd_Delete_QUERY.Enabled = True
            Me.Cmd_Importa_QUERY.Enabled = True
          
End Sub



'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub DisabilitaButton_QUERY_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA E CANCELLA QUERY
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@QUERY @cmd@QUERY@delete.disabilitato + cmd@QUERY@importa.disabilitato)
            Me.Cmd_Delete_QUERY.Enabled = False
            Me.Cmd_Importa_QUERY.Enabled = False
          
End Sub


'//ABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_MODULI_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA E CANCELLA MODULI
            '//@ABILITO@BUTTON_(abilito il @comando@MODULI @cmd@MODULI@delete.abilitato + cmd@MODULI@importa.abilitato)
            Me.Cmd_Delete_MODULI.Enabled = True
            Me.Cmd_Importa_MODULI.Enabled = True
          
End Sub



'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub DisabilitaButton_MODULI_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA E CANCELLA MODULI
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@MODULI @cmd@MODULI@delete.disabilitato + cmd@MODULI@importa.disabilitato)
            Me.Cmd_Delete_MODULI.Enabled = False
            Me.Cmd_Importa_MODULI.Enabled = False
          
End Sub



'//ABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_REPORT_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA E CANCELLA REPORT
            '//@ABILITO@BUTTON_(abilito il @comando@REPORT @cmd@REPORT@delete.abilitato + cmd@REPORT@importa.abilitato)
            Me.Cmd_Delete_REPORT.Enabled = True
            Me.Cmd_Importa_REPORT.Enabled = True
          
End Sub


'//ABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_MACRO_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA E CANCELLA MACRO
            '//@ABILITO@BUTTON_(abilito il @comando@MACRO @cmd@MACRO@delete.abilitato + cmd@MACRO@importa.abilitato)
            Me.Cmd_Delete_MACRO.Enabled = True
            Me.Cmd_Importa_MACRO.Enabled = True
          
End Sub







'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub DisabilitaButton_MACRO_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA E CANCELLA MACRO
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@MACRO @cmd@MACRO@delete.disabilitato + cmd@MACRO@importa.disabilitato)
            Me.Cmd_Delete_MACRO.Enabled = False
            Me.Cmd_Importa_MACRO.Enabled = False
          
End Sub




'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub DisabilitaButton_REPORT_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA E CANCELLA REPORT
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@REPORT @cmd@REPORT@delete.disabilitato + cmd@REPORT@importa.disabilitato)
            Me.Cmd_Delete_REPORT.Enabled = False
            Me.Cmd_Importa_REPORT.Enabled = False
          
End Sub




'//DISABILITAZIONE@BUTTON
'//NOTE : ABILITA SOLO 1 COMANDO IMPORTA ED 1 CANCELLA
Private Sub AbilitaButton_IMPORTA_TUTTO_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA TUTTO
            '//@DISABILITO@BUTTON_(ABILITO il @comando@IMPORTA@TUTTO)
            Me.Cmd_IMPORTA_TUTTO.Enabled = True
            Me.Cmd_DELETE_TUTTO.Enabled = True
            
End Sub



'//DISABILITAZIONE@BUTTON
'//NOTE : DISABILITA SOLO 1 COMANDO IMPORTA TUTTO
Private Sub DisabilitaButton_IMPORTA_TUTTO_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA TUTTO
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@IMPORTA@TUTTO)
            Me.Cmd_IMPORTA_TUTTO.Enabled = False
            
End Sub


'//DISABILITAZIONE@BUTTON@DELETE@TUTTO
'//NOTE : DISABILITA SOLO 1 COMANDO DELETE TUTTO
Private Sub DisabilitaButton_DELETE_TUTTO_Sub()
    
            '//DISABILITO SOLO I COMANDI, IMPORTA TUTTO
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@IMPORTA@TUTTO)
            Me.Cmd_DELETE_TUTTO.Enabled = False
          
End Sub



'//ABILITAZIONE@BUTTON@DELETE@TUTTO
'//NOTE : ABILITA SOLO 1 COMANDO DELETE TUTTO
Private Sub AbilitaButton_DELETE_TUTTO_Sub()
    
            '//ABILITO SOLO I COMANDI, IMPORTA TUTTO
            '//@DISABILITO@BUTTON_(DISABILITO il @comando@IMPORTA@TUTTO)
            Me.Cmd_DELETE_TUTTO.Enabled = True
          
End Sub


'//ABILITAZIONE E DIABLITAIONE DEI BUTTON , TABLE, FORM, QUERY MACRO REPORT *** INIZIO ***
'//-------------------------------------------------------------------------------------------------------------//




'//@ABILITA@COMBINATA_01=@CMB_02
Private Sub AbilitaCmd_02_Sub()
    'Me.Cmb_01_txt.Selected
    Me.Cmb_02_txt.Enabled = True
    Me.Cmb_02_txt.Value = "ELENCO OGGETTI MSYS"

End Sub






'//***********************************************************************************************************************//
'//                             @ABILITAZIONE E @DISABILLITAZIONI        *** FINE *****
'//
'//***********************************************************************************************************************//




'//***********************************************************************************************************************//
'//                             @STAMPE@OGGETTI@CASELLA@DI@TESTO        *** INIZIO *****
'//
'//***********************************************************************************************************************//


'//@STAMPA@OGGETTI@MSYS
'//================================================================================================================//
'// STAMPA OGGETTI COLLETTIVA *** INIZIO ***
'//================================================================================================================//

Private Function Stampa_MSYS_N01_OGGETTI_funct(par_LaCollection As Collection, par_Tipo_s As String) As String


'//DIM della collezione form da importare
Dim formsCollection As Collection
Dim formName_v As Variant

'//DIM la collezione macro
Dim macroCollection As Collection
Dim MACROS_Name_v As Variant



'//DIM la collezione DEI MODULI
Dim Modules_Collection As Collection
Dim moduloName_v As Variant


'//DIM la collezione DELLE QUERY
Dim queryCollection As Collection
Dim queryName_v As Variant




'//DIM la collezione DELLE REPORT
Dim reportCollection As Collection
Dim reportName_v As Variant



'//DIM la collezione DELLE TABELLE
Dim tableCollection As Collection
Dim tableName_v As Variant


'//Le variabili di controllo della sub
Dim ProceduraMessaggioErrore_s As String
Dim ProceduraAttivaEseguita_s As String
   



    '//....
On Error GoTo Err_Stampa_MSYS_N01_OGGETTI_funct


'//RESET VARIABILI
ProceduraMessaggioErrore_s = "PROCEDURA_ESEGUITA_Stampa_MSYS_N01_OGGETTI_funct"
ProceduraAttivaEseguita_s = "ATTIVITA_DI_CONTROLLO_controllo_record"


'//LA_SUB_Stampa_MSYS_N01_OGGETTI_funct_01.01
'//______________________________________________________________________________
            

            
                 
                '//CREO IL MESSAGGIO DA VISUALIZZARE NELLA CASELLA DI TESTO UNITAMENTE AGLI OGGETTI DELLA COLLEZIONE
                '//---------------------------------------------------------------------------------------------------------//
                '//NOTE: 01_TESTATA_MESSAGGIO = cre la testata del messaggio da riportare nella casella di testo
                ProceduraAttivaEseguita_s = ""
                ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ************************* PROGETTO DA CARICARE : " & PROGETTO_s & " ********************" & vbCrLf & vbCrLf
                ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " PATH DEL PROGETTO: " & sourceDBPath_s & vbCrLf
                
                
                
                ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf
                ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI FORM DA CARICARE  " & vbCrLf
                ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                
                
                Select Case par_Tipo_s
                
                    Case "COLLECTION_FORMS"
                
                   ' ATTIVA LA COLLEZIONE DELLE FORMS
                    '//------------------------------------------------------------------------//
                
                        
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI FORMS DA CARICARE  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        
                        '//ciclo nella collezione PASSATA COME PARAMETRO per il CARICAMENTO DELLE FORMS NEL STRINGA MESSAGGIO
                        For Each FORMSName_v In par_LaCollection
                            Debug.Print "FORMS nella collezione: " & FORMSName_v
                            ' Aggiungi l'oggetto della collezione FORMS nella stringa di messaggio
                            ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & FORMSName_v
                            
                        Next FORMSName_v
                        
                        
                           
                        '//VISUALIZZO NELLA CASELLA DI TESTO IL TITOLO E LE FORMS DA CARICARE
                        Me.TXT_OBJECT_ESTERNI = ProceduraAttivaEseguita_s & vbCrLf

                        
                        
                    '//------------------------------------------------------------------------//
                        
                        
                Case "COLLECTION_MACROS"
                        
                    ' ATTIVA LA COLLEZIONE DELLE MACRO
                    '//------------------------------------------------------------------------//
                
                        
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI MACRO DA CARICARE  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        
                        '//ciclo nella collezione PASSATA COME PARAMETRO per il CARICAMENTO DELLE MACRO NEL STRINGA MESSAGGIO
                        For Each MACROS_Name_v In par_LaCollection
                            Debug.Print "MACRO nella collezione: " & MACROS_Name_v
                            ' Aggiungi l'oggetto della collezione MACRO nella stringa di messaggio
                            ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & MACROS_Name_v
                            
                        Next MACROS_Name_v
                        
                        
                           
                        '//VISUALIZZO NELLA CASELLA DI TESTO IL TITOLO E LE MACRO DA CARICARE
                        Me.TXT_OBJECT_ESTERNI = ProceduraAttivaEseguita_s & vbCrLf

                        
                        
                    '//------------------------------------------------------------------------//

                Case "COLLECTION_MODULES"
                
                    ' ATTIVA LA COLLEZIONE DEI MODULI
                    '//------------------------------------------------------------------------//
                
                        
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI DEI MODULI DA CARICARE  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        
                        
                        
                        '//ciclo nella collezione PASSATA COME PARAMETRO per il CARICAMENTO DELLE MACRO NEL STRINGA MESSAGGIO
                        For Each moduloName_v In par_LaCollection
                            Debug.Print "MACRO nella collezione: " & MACROS_Name_v
                            ' Aggiungi l'oggetto della collezione MACRO nella stringa di messaggio
                            ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & moduloName_v
                            
                    
                        Next moduloName_v
                        
                                
                        '//VISUALIZZO NELLA CASELLA DI TESTO IL TITOLO I MODULI DA CARICARE
                        Me.TXT_OBJECT_ESTERNI = ProceduraAttivaEseguita_s & vbCrLf

                        
                    '//------------------------------------------------------------------------//
                    
                    
                    
                Case "COLLECTION_QUERYES"
                    
                    '// ATTIVA LA COLLEZIONE DELLE QUERY
                    '//------------------------------------------------------------------------//
                
                        
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI DELLE QUERY DA CARICARE  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " Attenzione per le query viene prima esaminata tutti gli oggetti che sono   " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " salvati nel database esterno, sul quale viene asperta una interrogazione   " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " sql che filtra dal database esterno e dalla tabella di sistema denominata  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " msysobject nel database MSYS_ATTIVA_GEST_OGGETTI.mdb. La sql di estrazione applica " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " due filtri;                                                                " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " 01 filtro = esclude l'estrazione dei nomi che iniziano                     " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "  ~sq = solo le tabelle cancellate.                                         " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " 02 filtro = escludo dalla collection tali query e quindi vengono estratte  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " solo le tabelle che sono in lavorazione.                                   " & vbCrLf
                        
                        
                       
                        '//ciclo nella collezione PASSATA COME PARAMETRO per il CARICAMENTO DELLE QUERY NEL STRINGA MESSAGGIO
                        For Each queryName_v In par_LaCollection
                            
                            Debug.Print "QUERY nella collezione: " & queryName_v
                            
                            'Aggiungi l'oggetto della collezione  QUERY nella stringa di messaggio
                            ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & queryName_v
                            
                        Next queryName_v
                                                            
                                    
                        '//VISUALIZZO NELLA CASELLA DI TESTO IL TITOLO LE QUERY DA CARICARE
                        Me.TXT_OBJECT_ESTERNI = ProceduraAttivaEseguita_s & vbCrLf
                        
                    '//------------------------------------------------------------------------//
                    
                    
                    
                Case "COLLECTION_REPORTS"
                
                    '// ATTIVA LA COLLEZIONE DELLE REPORT
                    '//------------------------------------------------------------------------//
                
                        
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI DELLE REPORT DA CARICARE  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf
                           
                           
                         '//ciclo nella collezione PASSATA COME PARAMETRO per il CARICAMENTO DELLE FORM NEL STRINGA MESSAGGIO
                        For Each reportName_v In par_LaCollection
                            
                            Debug.Print "REPORT nella collezione: " & reportName_v
                            
                            'Aggiungi l'oggetto della collezione  REPORT nella stringa di messaggio
                            ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & reportName_v
                    
                        Next reportName_v
                        
                            
                        '//VISUALIZZO NELLA CASELLA DI TESTO IL TITOLO LE REPORT DA CARICARE
                        Me.TXT_OBJECT_ESTERNI = ProceduraAttivaEseguita_s & vbCrLf
                    
                    
                    '//------------------------------------------------------------------------//

            
            Case "COLLECTION_TABLES"
                    '// ATTIVA LA COLLEZIONE DELLE TABELLE
                    '//------------------------------------------------------------------------//
                
                        
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & "         LE COLLEZIONI DELLE TABELLE DA CARICARE  " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & " ..................................................................... " & vbCrLf
                        ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf
                           
                         '//ciclo nella collezione PASSATA COME PARAMETRO per il CARICAMENTO DELLE TABELLE NEL STRINGA MESSAGGIO
                        For Each tableName_v In par_LaCollection
                            
                            Debug.Print "TABELLE nella collezione: " & tableName_v
                            
                            'Aggiungi l'oggetto della collezione  TABELLE nella stringa di messaggio
                            ProceduraAttivaEseguita_s = ProceduraAttivaEseguita_s & vbCrLf & tableName_v
                    
                        Next tableName_v
                        
                        
                        '//VISUALIZZO NELLA CASELLA DI TESTO IL TITOLO LE TABELLE DA CARICARE
                        Me.TXT_OBJECT_ESTERNI = ProceduraAttivaEseguita_s & vbCrLf
                    '//------------------------------------------------------------------------//
                    
                Case Else
                        MsgBox "NESSUNA COLLECTION E' STATA STAMPATA PERCHE NON ESISTE", vbCritical, "CONTROLLO DI STAMPA"
                    
                End Select
                              
                   
          
'//______________________________________________________________________________




'//USCITA  E GESTIONE ERRORI
'//..............................................................................................................


Exit_Stampa_MSYS_N01_OGGETTI_funct:
    Exit Function

Err_Stampa_MSYS_N01_OGGETTI_funct:
    MsgBox Err.Description
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Stop
    Resume Exit_Stampa_MSYS_N01_OGGETTI_funct

End Function


'//================================================================================================================//
'// STAMPA OGGETTI COLLETTIVA *** FINE ***
'//================================================================================================================//







'***********************************************************************************************************************
'                                     IL TAB CONTROLL *** INIZIO ***
'
'***********************************************************************************************************************

'//====================================================================================================================//
'                         @TAB@CONTROLL@EVENTI_(@EVENTI DEL TAB CONTROLL)
'
'  01) TabCtl_01_Change()
'
'//====================================================================================================================//
'//@MODIFICHE _
    2025.01.29= CREATO GLI EVENTI DEL TAB CONTROLL PER SCRIVERE SUL LABEL PRINCIPALE IL NOME DELLE PAGINE


'//DENOMINAZIONE---------> TabCtl_01_Change
'//=================================================================================================================//
'//TIPO---------------------> Routine privata.
'//ATTIVITA--------------> GESTIONE DELLE PAGINE TAB CONTROLL
'//NOTE---------------------> ....
'//PARAMETRI-------------> Nessuno
'//VALORE_DI_RITORNO-----> Nulla
'//CODICE-------------------> TabCtl_01_Change.01.00   @TAB@CONTROL_(oggetto tab controll della form @99@UTILITA@IMPEGNI)
'//
'//02.@TAB@CONTROL= LA GESTIONE DELLE PAGINE DEL TAB CONTROL
'//
'//=================================================================================================================//
'//
'//ATTIVITA ED EVENTI DEL TAB CONTROL: _
  01.TABCONTROL= RESET LE VARIABILI _
  02.TABCONTROL=RECUPERO IL NUMERO DELLA PAGINA CORRENTE _
  03.TABCONTROL=IF DI CONTROLLO DEL VALORE PAGINA E ATTIVITA DA ESEGUIRE _
  04.TABCONTROL=ASSEGNO IL VALORE PAGINE E NOME DELLA PAGINA ALLA LABEL TITOLO PRINCIPALE






'//@EVENTO@CHANGE_(@cambio@pagina del @tab@control)'
Private Sub TabCtl_01_Change()

'//DIM--------------------------
Dim ixPage As Integer


'//-----------------------------

    On Error GoTo Err_TabCtl_01_Change
        
       '//Riaggiorno la form principale
       'Me.Recalc
       'Me.Requery
       
    '//RESET
    ProceduraMessaggioErrore_s = ""
    ProceduraAttivaEseguita_s = ""


    '//IMPOSTO LE VARIABILI
    ProceduraMessaggioErrore_s = "GESTIONE DEI SERVIZI DEGLI IMPEGNI"
    ProceduraAttivaEseguita_s = "Modello_FUNZIONE_N01_PFucnt"

       
    'PAGINE
    '____________________________________________________________________________________________________________
        
        
        '02.TABCONTROL=RECUPERO IL NUMERO DELLA PAGINA CORRENTE
        '............................................................................
            'reset variabili
            'sxFORM_CHIAMANTE_GENERALE_CORRENTE = Me.Form.Name
            
            ixPage = Me.TabCtl_01
            
            
            'chiamo la funzione
            'sxPagCorrente_Caption = pfunc_Recupero_NOME_PAGINA_CORRENTE(Dim sxFORM_CHIAMANTE_GENERALE_CORRENTE, 1)
        '............................................................................
        
        '03.TABCONTROL=IF DI CONTROLLO DEL VALORE PAGINA E ATTIVITA DA ESEGUIRE
        If Me.TabCtl_01.Value = 0 Then
               
                    'PG_01
                    '............................................................................
                    'Imposto il valore della casella combinata
                    'valore = 0 : pannello lotto; valore = 1 : ambo ecc.
                    'Me.Cmb_01.Value = 1
                        
                            'Imposta i parametri del comando per l'ambo
                            'ixCommand = 1
                            'vxArgument = 412
                            'la scheda corrente attiva
                            'sxPagCorrente_Caption = ""
                            
                          
                         'Me.SottForm_Pg_02.Requery
                    
                        'Aggiorno Sottoform_01
                        'Me.SottForm_Pg_01.Requery
                '............................................................................
                 
    'Pag 2
    '____________________________________________________________________________________________________________
        ElseIf Me.TabCtl_01.Value = 1 Then
    
                    
                    'PG_02
                    '............................................................................
                    'Imposto il valore della casella combinata
                    
                            'Imposta i parametri del comando per l'ambo
                            'ixCommand = 1
                            'vxArgument = 413
                            'la scheda corrente attiva
                            'sxPagCorrente_Caption = "TERNO"
                '............................................................................
                
                        '//REQUERY DELLA SOTTOFORM_2
                        '//Cambio proprieta record source della sottoform caricata e il requery è automatico
                        '//MascheraPrincipale------->Sottoform02.....>RecordoSource......="NomeQuery"
                        'Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_02].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco_ANNUALITA_TMP_Folium"
                        
                         'Me.SottForm_Pg_02.Requery
                
    'Pag 3
    '____________________________________________________________________________________________________________
        ElseIf Me.TabCtl_01.Value = 2 Then
    
                
                'PG_03 - TAB CONTROLL -> TabCtl_01.value=2
                '............................................................................
                    'Imposto il valore della casella combinata
                        'Imposta i parametri del comando per l'ambo
                            'ixCommand = 1
                            'vxArgument = 414
                            'la scheda corrente attiva
                            'sxPagCorrente_Caption = "QUADERNA"
                        
                        '............................................................................
                
                        '//REQUERY DELLA SOTTOFORM_2
                        '//Cambio proprieta record source della sottoform caricata e il requery è automatico
                        '//MascheraPrincipale------->Sottoform02.....>RecordoSource......="NomeQuery"
                        'Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_03].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco_ANNUALITA_TMP_Folium"
                        
                        
                         'Me.SottForm_Pg_03.Requery
                        
                '............................................................................
                    
                
    'Pag 4
    '____________________________________________________________________________________________________________
        ElseIf Me.TabCtl_01.Value = 3 Then
    
                
                'PG_04
                '............................................................................
                    'Imposto il valore della casella combinata
                    
                        
                            'Imposta i parametri del comando per l'ambo
                            'ixCommand = 1
                            'vxArgument = 415
                            'la scheda corrente attiva
                            'sxPagCorrente_Caption = "?"
                               
                        '............................................................................
                
                        '//REQUERY DELLA SOTTOFORM_4
                        '//Cambio proprieta record source della sottoform caricata e il requery è automatico
                        '//MascheraPrincipale------->Sottoform02.....>RecordoSource......="NomeQuery"
                        'Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_04].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_DET"
                        
                        'Me.SottForm_Pg_04.Requery
                        
                '............................................................................
            
    '____________________________________________________________________________________________________________

                 
    'Pag 5
    '____________________________________________________________________________________________________________
        ElseIf Me.TabCtl_01.Value = 4 Then
    
                'Aggiorno form
                'Me.Pg_05.Requery
                
                
                
                'PG_05 TAB CONTROLL -> TabCtl_01.value=4
                '............................................................................
                    'Imposto il valore della casella combinata
                   '............................................................................
                
                        '//REQUERY DELLA SOTTOFORM_3
                        '//Cambio proprieta record source della sottoform caricata e il requery è automatico
                        '//MascheraPrincipale------->Sottoform02.....>RecordoSource......="NomeQuery"
                        'Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_05].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco_ANNUALITA_TMP_DGC"
                        
                  
                    
                '............................................................................
                
                '//REQUERY DELLA SOTTOFORM_5
                'Me.SottForm_Pg_05.Requery
                         
    '____________________________________________________________________________________________________________

    'Pag 6
    '____________________________________________________________________________________________________________
        ElseIf Me.TabCtl_01.Value = 5 Then
    
                'Aggiorno form
                'Me.Pg_06.Requery
                
                
                
                'PG_06 TAB CONTROLL -> TabCtl_01.value=5
                '............................................................................
                        '//REQUERY DELLA SOTTOFORM_2
                        '//Cambio proprieta record source della sottoform caricata e il requery è automatico
                        '//MascheraPrincipale------->Sottoform02.....>RecordoSource......="NomeQuery"
                        'Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_06].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco_ANNUALITA_TMP_DCC"
                        
                         
                        
                '............................................................................
                
                '//REQUERY DELLA SOTTOFORM_6
                'Me.SottForm_Pg_06.Requery
                         
    '____________________________________________________________________________________________________________
        
            
            
    'Pag 7
    '____________________________________________________________________________________________________________
        ElseIf Me.TabCtl_01.Value = 6 Then
    
                'Aggiorno form
                'Me.Pg_07.Requery
                
                
                
                'PG_07 ???
                '............................................................................
                    'Imposto il valore della casella combinata
                      '//REQUERY DELLA SOTTOFORM_2
                        '//Cambio proprieta record source della sottoform caricata e il requery è automatico
                        '//MascheraPrincipale--------->Sottoform02..............................Sottoform03...........................>RecordoSource......="NomeQuery"
                         'Forms![GEST_FrmM00_GESTIONE]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco"
                         'Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_07]![LLPP_ATTI_FrmS01-21_GRUPPO_CODICE_OPERA]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco"
                          'Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_07]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01-12_Gestione_Elenco"
                          '[Forms]![GEST_FrmM00_GESTIONE]![SottForm_Pg_07].[Form]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].[Form].[RecordSource] = "LLPP_ATTI_Qry01-12_Gestione_Elenco"
                          'Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_07]![LLPP_ATTI_FrmS01-21_GRUPPO_CODICE_OPERA].[LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"
                          'Application.Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_07]![LLPP_ATTI_FrmS01-21_GRUPPO_CODICE_OPERA].[LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"
                          'Forms![GEST_FrmM00_GESTIONE]![LLPP_ATTI_FrmS01-21_GRUPPO_CODICE_OPERA].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"
                          
                          'Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_07]![LLPP_ATTI_FrmS01-21_GRUPPO_CODICE_OPERA]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"
                         'Application.Forms.Item(14).Form
                          'Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_07]![LLPP_ATTI_FrmS01-12_Gestione_Elenco].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"
                            
                            'Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_07].Form.RecordSource = "LLPP_ATTI_Qry01_12_Gestione_Elenco_ANNUALITA_TMP_ALTRO"

                  '............................................................................
                
                '//REQUERY DELLA SOTTOFORM_7
                'Me.SottForm_Pg_07.Requery
                         
    '____________________________________________________________________________________________________________
            
        

            
    'Pag 8
    '____________________________________________________________________________________________________________
        ElseIf Me.TabCtl_01.Value = 7 Then
    
                'Aggiorno form
                'Me.Pg_08.Requery
                
                
                
                'PG_08
                '............................................................................
                    'Imposto il valore della casella combinata
                    '//REQUERY DELLA SOTTOFORM_2
                     '//Cambio proprieta record source della sottoform caricata e il requery è automatico
                     '//MascheraPrincipale------->Sottoform02.....>RecordoSource......="NomeQuery"
                     'Forms![GEST_FrmM00_GESTIONE]![SottForm_Pg_07].Form.RecordSource = ""
                                                                                       
                    
                        
                '............................................................................
                
                '//REQUERY DELLA SOTTOFORM_8
                'Me.SottForm_Pg_08.Requery
                         
    '____________________________________________________________________________________________________________
            
                
                
                
            
    'Pag 9
    '____________________________________________________________________________________________________________
        ElseIf Me.TabCtl_01.Value = 8 Then
    
                'Aggiorno form
                'Me.Pg_09.Requery
                
                
                
                'PG_09
                '............................................................................
                    'Imposto il valore della casella combinata
                    
                        
                '............................................................................
                
                '//REQUERY DELLA SOTTOFORM_9
                'Me.SottForm_Pg_09.Requery
                
                         
    '____________________________________________________________________________________________________________
            


            
    'Pag 10
    '____________________________________________________________________________________________________________
        ElseIf Me.TabCtl_01.Value = 9 Then
    
                'Aggiorno form
                'Me.Pg_10.Requery
                
                
                
                'PG_10
                '............................................................................
                    'Imposto il valore della casella combinata
                    
                        
                '............................................................................
                
                '//REQUERY DELLA SOTTOFORM_10
                'Me.SottForm_Pg_10.Requery
                         
    '____________________________________________________________________________________________________________
            

            
    'Pag 11
    '____________________________________________________________________________________________________________
        ElseIf Me.TabCtl_01.Value = 10 Then
    
              
    '____________________________________________________________________________________________________________
            

    'Pag 12
    '____________________________________________________________________________________________________________
        ElseIf Me.TabCtl_01.Value = 11 Then
    
              
    '____________________________________________________________________________________________________________
            
   
                
        
    End If
    
    
    '//04.TABCONTROL=ASSEGNO IL VALORE PAGINE E NOME DELLA PAGINA ALLA LABEL TITOLO PRINCIPALE
    Me.LBL_TIT01_TXT.Caption = Me.TabCtl_01.Pages(ixPage).Caption & " TAB NRO : " & Me.TabCtl_01.Pages(ixPage).Name
    
    
    
    '//Aggiorno la form master
    'Me.Recalc
    'Me.Requery
                

'USCITA  E GESTIONE ERRORI
'..............................................................................................................

Exit_TabCtl_01_Change:
    Exit Sub

Err_TabCtl_01_Change:
   
    MsgBox Err.Description & " " & ProceduraMessaggioErrore_s & " - " & ProceduraAttivaEseguita_s & " --> errore Error$ -> : " & Error$
    Debug.Print ProceduraMessaggioErrore_s
    Debug.Print ProceduraAttivaEseguita_s
    Debug.Assert "BLOCCO PROCEDURA -> " & ProceduraAttivaEseguita_s

    
    Resume Exit_TabCtl_01_Change


End Sub

'//DENOMINAZIONE---------> TabCtl_01_Change *** FINE ***
'//=================================================================================================================//



'***********************************************************************************************************************
'                                     IL TAB CONTROLL *** FINE ***
'
'***********************************************************************************************************************



'***********************************************************************************************************************
'                                    CASELLA DI TESTO --> CurrentProject_txt  *** INIZIO ***
'
'***********************************************************************************************************************

'//@CASELLA.DI.TESTO_(@casella di @testo @progetto@corrente RESETTA TUTTE LE @LABEL DELLE CASELLE DI TESTO)
Private Sub CurrentProject_txt_Click()
    Call Form_Load
    Call Imposta_Colore_Sub             '//@IMPOSTA@COLORE.01=RESET

End Sub



'***********************************************************************************************************************
'                                    CASELLA DI TESTO --> CurrentProject_txt  *** FINE ***
'
'***********************************************************************************************************************





'***********************************************************************************************************************
'                                    STAMPA OGGETTI *** INIZIO ***
'
'***********************************************************************************************************************
'//01@STAMPA@TABELLE    = STAMPO LE TABELLE DEL DB CORRENTE;    FILE 01_LOG_TABELLE SALVATO NELLA PATH CORRENTE
'//02@STAMPA@QUERY      = STAMPO LE QUERY DEL DB CORRENTE;      FILE 02_LOG_QUERY SALVATO NELLA PATH CORRENTE
'//03@STAMPA@FORM       = STAMPO LE FORM DEL DB CORRENTE;       FILE 03_LOG_FORM SALVATO NELLA PATH CORRENTE
'//04@STAMPA@REPORT     = STAMPO LE REPORT DEL DB CORRENTE;     FILE 04_LOG_REPORT SALVATO NELLA PATH CORRENTE
'//05@STAMPA@MODULI     = STAMPO LE MODULI DEL DB CORRENTE;     FILE 05_LOG_MODULI SALVATO NELLA PATH CORRENTE
'//06@STAMPA@TUTTO      = STAMPO TUTTI GLI OGGETTI DEL DB CORRENTE.
'//07@STAMPA@LOG@ERRORI = @STAMPO IL @FILE@LOG DI @GESTIONE DEGLI @ERRORI DELLA SINGOLA FUNZIONE O ROUTINE


'//01@STAMPA@TABELLE    = STAMPO LE TABELLE DEL DB CORRENTE;    FILE 01_LOG_TABELLE SALVATO NELLA PATH CORRENTE _
Impostazione del database corrente: Set db = CurrentDb() imposta il database corrente come oggetto db. _
Ciclo attraverso le tabelle: Il ciclo For Each tdf In db.TableDefs scorre tutte le tabelle (TableDefs) nel database. _
Esclusione delle tabelle di sistema: Le tabelle di sistema (che iniziano con "MSys" o "~") vengono escluse dalla stampa. _
Stampa dei nomi delle tabelle: I nomi delle tabelle vengono stampati nella finestra di debug (Debug.Print). _
Rilascio delle risorse: Le risorse vengono rilasciate impostando tdf e db su Nothing. _
Messaggio di conferma: Un messaggio di conferma viene visualizzato per indicare che la stampa è stata completata. _

Private Sub Cmd_STAMPA_TABLE_DB_CORRENTE_Click()
    On Error GoTo ErrorHandler ' Attiva il controllo degli errori
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim tblName As String
    Dim i As Integer
    Dim logFilePath As String
    Dim fileNumber As Integer
    Dim dbPath As String
    Dim dbFolder As String
    
    ' Imposta il database corrente
    Set db = CurrentDb()
    
    ' Ottieni il percorso della cartella corrente del database
    dbPath = CurrentDb.Name
    dbFolder = Left(dbPath, InStrRev(dbPath, "\")) ' Estrae la cartella
    
    ' Definisci il percorso del file di log nella stessa cartella del database
    logFilePath = dbFolder & "01_LOG_TABELLE.txt" ' Nome del file di log
    
    ' Inizializza il contatore
    i = 1
    Me.TXT_GESTIONE_UTILITA.Value = ""
    Me.TXT_GESTIONE_UTILITA.Requery
    
    ' Apri il file di log in modalità scrittura
    fileNumber = FreeFile
    '//APRE CON APPEND che aggiugne ai dati precedenti altrimenti se Open logFilePath For Output As #fileNumber _
                       cancella i dati precedenti on Output
    Open logFilePath For Append As #fileNumber
    
    
    ' Cicla attraverso tutte le tabelle nel database
    For Each tdf In db.TableDefs
            tblName = tdf.Name
                
        ' Escludi le tabelle di sistema che iniziano con "MSysA ...ecc"
        If Left(tdf.Name, 5) <> "MSysA" And Left(tdf.Name, 5) <> "MSysN" And Left(tdf.Name, 5) <> "MSysO" _
                            And Left(tdf.Name, 5) <> "MSysQ" And Left(tdf.Name, 5) <> "MSysR" _
                            And Left(tdf.Name, 5) <> "MSysS" Then
                            
            ' Verifica se la tabella è collegata o locale
            If (tdf.Attributes And dbAttachedTable) = 0 Then
                tblName = tdf.Name
                
                ' Stampa il nome della tabella  COLLEGATA nella casella di testo
                Me.TXT_GESTIONE_UTILITA.Value = Me.TXT_GESTIONE_UTILITA.Value & "Tabella " & i & ": " & tblName & vbCrLf
                
                ' Scrivi il nome della tabella nel file di log
                Print #fileNumber, "Tabella " & i & ": " & tblName
                
                i = i + 1
            
            Else
                '//tabella LOCALE non collegata stampata e salvata nel log
                i = i + 1
                ' Stampa il nome della tabella nella casella di testo
                Me.TXT_GESTIONE_UTILITA.Value = Me.TXT_GESTIONE_UTILITA.Value & "Tabella " & i & ": " & tblName & vbCrLf
                 
                 ' Scrivi il nome della tabella nel file di log
                Print #fileNumber, "Tabella " & i & ": " & tblName
              
            End If
            
            
            
        End If
          
             
    Next tdf
    
    ' Chiudi il file di log
    Close #fileNumber
    
    ' Aggiorna la casella di testo
    Me.TXT_GESTIONE_UTILITA.Requery
    
    ' Rilascia le risorse
    Set tdf = Nothing
    Set db = Nothing
    
    ' Messaggio di conferma
    MsgBox "Stampa delle tabelle completata! File di log creato in: " & logFilePath, vbInformation
    
    Exit Sub ' Esci dalla routine per evitare di eseguire il gestore degli errori

ErrorHandler:
    ' Gestione degli errori
    MsgBox "Si è verificato un errore: " & Err.Description, vbCritical
    ' Rilascia le risorse in caso di errore
    If Not tdf Is Nothing Then Set tdf = Nothing
    If Not db Is Nothing Then Set db = Nothing
    ' Chiudi il file di log se aperto
    If fileNumber > 0 Then Close #fileNumber
End Sub




'//02@STAMPA@QUERY      = STAMPO LE QUERY DEL DB CORRENTE;   FILE 02_LOG_QUERY SALVATO NELLA PATH CORRENTE
'//@ESEMPIO@STAMPA@QUERY_(Esempio di @stampa@oggetti@QUERY nel @database corrente)
'//@faq@QUERY           =  @COME @STARE LE @QUERY DEL DATABASE CORRENTE _
                        =  @STAMPARE@LE@QUERY@DEL@DATABASE _
                        =  @STAMPA@DELLE@QUERY  @NEL @DATABASE @CORRENTE _
                        =  @FAQ@STAMPA@QUERY _
                        =  @ROUTINE@STAMPA@QUERY _
                        =  @COME@STAMPARE@LE@QUERY


Private Sub Cmd_STAMPA_QUERY_DB_CORRENTE_Click()
    On Error GoTo ErrorHandler ' Attiva il controllo degli errori
    
    Dim db As DAO.Database
    Dim qry As DAO.QueryDef
    Dim qryName As String
    Dim i As Integer
    Dim logFilePath As String
    Dim fileNumber As Integer
    Dim dbPath As String
    Dim dbFolder As String
    
    ' Imposta il database corrente
    Set db = CurrentDb()
    
    ' Ottieni il @percorso della cartella corrente del database  @estrazione@path@corrente_(vba @estra la @cartella @corrente)
    dbPath = CurrentDb.Name
    dbFolder = Left(dbPath, InStrRev(dbPath, "\")) ' Estrae la cartella @path@corrente_(la path della @cartella@corrente)
    
    ' Definisci il percorso del file di log nella stessa cartella del database
    logFilePath = dbFolder & "02_LOG_QUERY.txt" ' Nome del file di log
    
    ' Inizializza il contatore
    i = 1
    Me.TXT_GESTIONE_UTILITA.Value = ""
    Me.TXT_GESTIONE_UTILITA.Requery
    
    ' Apri il file di log in modalità scrittura
    fileNumber = FreeFile
    Open logFilePath For Output As #fileNumber
    
    ' Cicla attraverso tutte le QUERY nel database
    For Each qry In db.QueryDefs
            qryName = qry.Name
                
                ' Escludi le QUERY di sistema che iniziano con "MSys" o "~"
                If Left(qry.Name, 5) <> "~sq_c" Then
                    qryName = qry.Name
                
                
                ' Stampa il nome della Query nella casella di testo
                Me.TXT_GESTIONE_UTILITA.Value = Me.TXT_GESTIONE_UTILITA.Value & "Query " & i & ": " & qryName & vbCrLf
                
                ' Scrivi il nome della Query nel file di log
                Print #fileNumber, "Query " & i & ": " & qryName
                
                i = i + 1
           
                End If
                
                
    Next qry
    
    
    ' Chiudi il file di log
    Close #fileNumber
    
    ' Aggiorna la casella di testo
    Me.TXT_GESTIONE_UTILITA.Requery
    
    ' Rilascia le risorse
    Set qry = Nothing
    Set db = Nothing
    
    ' Messaggio di conferma
    MsgBox "Stampa delle QUERY completata! File di log creato in: " & logFilePath, vbInformation
    
    Exit Sub ' Esci dalla routine per evitare di eseguire il gestore degli errori

ErrorHandler:
    ' Gestione degli errori
    MsgBox "Si è verificato un errore: " & Err.Description, vbCritical
    ' Rilascia le risorse in caso di errore
    If Not qry Is Nothing Then Set qry = Nothing
    If Not db Is Nothing Then Set db = Nothing
    ' Chiudi il file di log se aperto
    If fileNumber > 0 Then Close #fileNumber
End Sub




'//03@STAMPA@FORM       = STAMPO LE FORM DEL DB CORRENTE;       FILE 03_LOG_FORM SALVATO NELLA PATH CORRENTE
'//@ESEMPIO@STAMPA@FORM_(Esempio di @stampa@oggetti@FORM nel @database corrente)
'//@faq@FORM            =  @COME @STARE LE @FORM DEL DATABASE CORRENTE _
                        =  @STAMPARE@LE@FORM@DEL@DATABASE _
                        =  @STAMPA@DELLE@FORM  @NEL @DATABASE @CORRENTE _
                        =  @FAQ@STAMPA@FORM _
                        =  @ROUTINE@STAMPA@FORM _
                        =  @COME@STAMPARE@LE@FORM
Private Sub Cmd_STAMPA_FORM_DB_CORRENTE_Click()
    On Error GoTo ErrorHandler ' Attiva il controllo degli errori
    
    Dim db As DAO.Database
    Dim frm As Object
    Dim frmName As String
    Dim i As Integer
    Dim logFilePath As String
    Dim fileNumber As Integer
    Dim dbPath As String
    Dim dbFolder As String
    
    ' Imposta il database corrente
    Set db = CurrentDb()
    
    ' Ottieni il percorso della cartella corrente del database
    dbPath = CurrentDb.Name
    dbFolder = Left(dbPath, InStrRev(dbPath, "\")) ' Estrae la cartella
    
    ' Definisci il percorso del file di log nella stessa cartella del database
    logFilePath = dbFolder & "03_LOG_FORM.txt" ' Nome del file di log
    
    ' Inizializza il contatore
    i = 1
    Me.TXT_GESTIONE_UTILITA.Value = ""
    Me.TXT_GESTIONE_UTILITA.Requery
    
    ' Apri il file di log in modalità scrittura
    fileNumber = FreeFile
    Open logFilePath For Output As #fileNumber
    
    ' Cicla attraverso tutte le FORM nel database
    For Each frm In db.Containers("Forms").Documents
        frmName = frm.Name
        
        ' Escludi le FORM che iniziano con "~sq_c" (se necessario)
        If Left(frmName, 5) <> "~sq_c" Then
            ' Stampa il nome della Form nella casella di testo
            Me.TXT_GESTIONE_UTILITA.Value = Me.TXT_GESTIONE_UTILITA.Value & "Form " & i & ": " & frmName & vbCrLf
            
            ' Scrivi il nome della Form nel file di log
            Print #fileNumber, "Form " & i & ": " & frmName
            
            ' Incrementa il contatore solo per le form valide
            i = i + 1
        End If
    Next frm
    
    ' Chiudi il file di log
    Close #fileNumber
    
    ' Aggiorna la casella di testo
    Me.TXT_GESTIONE_UTILITA.Requery
    
    ' Rilascia le risorse
    Set frm = Nothing
    Set db = Nothing
    
    ' Messaggio di conferma
    MsgBox "Stampa delle FORM completata! File di log creato in: " & logFilePath, vbInformation
    
    Exit Sub ' Esci dalla routine per evitare di eseguire il gestore degli errori

ErrorHandler:
    ' Gestione degli errori
    MsgBox "Si è verificato un errore nella riga " & Erl & ": " & Err.Description, vbCritical
    ' Rilascia le risorse in caso di errore
    If Not frm Is Nothing Then Set frm = Nothing
    If Not db Is Nothing Then Set db = Nothing
    ' Chiudi il file di log se aperto
    If fileNumber > 0 Then Close #fileNumber
End Sub










'//04@STAMPA@REPORT     = STAMPO LE REPORT DEL DB CORRENTE;       FILE 04_LOG_REPORT SALVATO NELLA PATH CORRENTE
'//@ESEMPIO@STAMPA@REPORT_(Esempio di @stampa@oggetti@REPORT nel @database corrente)
'//@faq@REPORT          =  @COME @STARE LE @REPORT DEL DATABASE CORRENTE _
                        =  @STAMPARE@LE@REPORT@DEL@DATABASE _
                        =  @STAMPA@DELLE@REPORT  @NEL @DATABASE @CORRENTE _
                        =  @FAQ@STAMPA@REPORT _
                        =  @ROUTINE@STAMPA@REPORT _
                        =  @COME@STAMPARE@LE@REPORT

'//Vantaggi delle correzioni _
01) Accesso corretto ai report usando CurrentProject.AllReports invece di db.Containers("REPORTs").Documents. _
02) Migliore gestione del file di log: ora aggiunge (Append) invece di sovrascrivere (Output). _
03) Gestione dei Null nella casella di testo per evitare errori. _
04) Gestione degli errori più chiara senza Erl, con Err.Number e Err.Description. _
05) Rilascio sicuro delle risorse per evitare file bloccati o oggetti non deallocati.

Private Sub Cmd_STAMPA_REPORT_DB_CORRENTE_Click()
    On Error GoTo ErrorHandler ' Attiva il controllo degli errori
    
    Dim db As DAO.Database
    Dim rpt As AccessObject
    Dim rptName As String
    Dim i As Integer
    Dim logFilePath As String
    Dim fileNumber As Integer
    Dim dbPath As String
    Dim dbFolder As String
    
    ' Imposta il database corrente
    Set db = CurrentDb()
    
    ' Ottieni il percorso della cartella corrente del database
    dbPath = CurrentDb.Name
    dbFolder = Left(dbPath, InStrRev(dbPath, "\")) ' Estrae la cartella
    
    ' Definisci il percorso del file di log nella stessa cartella del database
    logFilePath = dbFolder & "04_LOG_REPORT.txt"
    
    ' Inizializza il contatore
    i = 1
    If IsNull(Me.TXT_GESTIONE_UTILITA.Value) Then
        Me.TXT_GESTIONE_UTILITA.Value = ""
    Else
        Me.TXT_GESTIONE_UTILITA.Value = ""
    End If
    Me.TXT_GESTIONE_UTILITA.Requery
    
    ' Apri il file di log in modalità APPEND per non sovrascrivere i dati precedenti
    fileNumber = FreeFile
    Open logFilePath For Append As #fileNumber
    
    
    ' Cicla attraverso tutti i report nel database
    For Each rpt In CurrentProject.AllReports
        rptName = rpt.Name
        
        ' Escludi i report che iniziano con "~sq_c"
        If Left(rptName, 5) <> "~sq_c" Then
            ' Scrivi il nome del report nella casella di testo
            Me.TXT_GESTIONE_UTILITA.Value = Me.TXT_GESTIONE_UTILITA.Value & "REPORT " & i & ": " & rptName & vbCrLf
            
            ' Scrivi il nome del report nel file di log
            Print #fileNumber, "REPORT " & i & ": " & rptName
            
            ' Incrementa il contatore solo per i report validi
            i = i + 1
        End If
    Next rpt
    
    ' Chiudi il file di log
    Close #fileNumber
    
    ' Aggiorna la casella di testo
    Me.TXT_GESTIONE_UTILITA.Requery
    
    ' Rilascia le risorse
    Set rpt = Nothing
    Set db = Nothing
    
    ' Messaggio di conferma
    MsgBox "Stampa dei REPORT completata! File di log creato in: " & logFilePath, vbInformation
    
    Exit Sub ' Esci dalla routine per evitare di eseguire il gestore degli errori

ErrorHandler:
    ' Gestione degli errori
    MsgBox "Si è verificato un errore: " & Err.Number & " - " & Err.Description, vbCritical
    
    ' Rilascia le risorse in caso di errore
    If Not rpt Is Nothing Then Set rpt = Nothing
    If Not db Is Nothing Then Set db = Nothing
    
    ' Chiudi il file di log se aperto
    If fileNumber > 0 Then Close #fileNumber
End Sub





'//05@STAMPA@MODULI     = STAMPO LE MODULI DEL DB CORRENTE;     FILE 05_LOG_MODULI SALVATO NELLA PATH CORRENTE
'//@ESEMPIO@STAMPA@MODULI_(Esempio di @stampa@oggetti@MODULI nel @database corrente)
'//@faq@MODULI          =  @COME @STARE LE @MODULI DEL DATABASE CORRENTE _
                        =  @STAMPARE@LE@MODULI@DEL@DATABASE _
                        =  @STAMPA@DELLE@MODULI  @NEL @DATABASE @CORRENTE _
                        =  @FAQ@STAMPA@MODULI _
                        =  @ROUTINE@STAMPA@MODULI _
                        =  @COME@STAMPARE@LE@MODULI

'//Vantaggi delle correzioni _
01) Accesso corretto ai MODULI usando CurrentProject.AllMODULIs invece di db.Containers("MODULIs").Documents. _
02) Migliore gestione del file di log: ora aggiunge (Append) invece di sovrascrivere (Output). _
03) Gestione dei Null nella casella di testo per evitare errori. _
04) Gestione degli errori più chiara senza Erl, con Err.Number e Err.Description. _
05) Rilascio sicuro delle risorse per evitare file bloccati o oggetti non deallocati.


Private Sub Cmd_STAMPA_MODULI_DB_CORRENTE_Click()
    On Error GoTo ErrorHandler ' Attiva il controllo degli errori
    
    Dim db As DAO.Database
    Dim MDL As Object
    Dim mdlName As String
    Dim i As Integer
    Dim logFilePath As String
    Dim fileNumber As Integer
    Dim dbPath As String
    Dim dbFolder As String
    
    ' Imposta il database corrente
    Set db = CurrentDb()
    
    ' Ottieni il percorso della cartella corrente del database
    dbPath = CurrentDb.Name
    dbFolder = Left(dbPath, InStrRev(dbPath, "\")) ' Estrae la cartella
    
    ' Definisci il percorso del file di log nella stessa cartella del database
    logFilePath = dbFolder & "05_LOG_MODULI.txt"
    
    ' Inizializza il contatore
    i = 1
    If IsNull(Me.TXT_GESTIONE_UTILITA.Value) Then
        Me.TXT_GESTIONE_UTILITA.Value = ""
    Else
        Me.TXT_GESTIONE_UTILITA.Value = ""
    End If
    Me.TXT_GESTIONE_UTILITA.Requery
    
    ' Apri il file di log in modalità APPEND per non sovrascrivere i dati precedenti
    fileNumber = FreeFile
    Open logFilePath For Append As #fileNumber
    
    ' Cicla attraverso tutti i moduli del database
    For Each MDL In Application.VBE.VBProjects(1).VBComponents
        mdlName = MDL.Name
        
        ' Filtra solo i moduli standard (1) e di classe (2), escludendo form (3) e report (100)
        If MDL.Type = 1 Or MDL.Type = 2 Then
            ' Stampa il nome del modulo nella casella di testo
            Me.TXT_GESTIONE_UTILITA.Value = Me.TXT_GESTIONE_UTILITA.Value & "Modulo " & i & ": " & mdlName & vbCrLf
            
            ' Scrivi il nome del modulo nel file di log
            Print #fileNumber, "Modulo " & i & ": " & mdlName
            
            ' Incrementa il contatore solo per i moduli validi
            i = i + 1
        End If
    Next MDL
    
    ' Chiudi il file di log
    Close #fileNumber
    
    ' Aggiorna la casella di testo
    Me.TXT_GESTIONE_UTILITA.Requery
    
    ' Rilascia le risorse
    Set MDL = Nothing
    Set db = Nothing
    
    ' Messaggio di conferma
    MsgBox "Stampa dei moduli completata! File di log creato in: " & logFilePath, vbInformation
    
    Exit Sub ' Esci dalla routine per evitare di eseguire il gestore degli errori

ErrorHandler:
    ' Gestione degli errori
    MsgBox "Si è verificato un errore: " & Err.Number & " - " & Err.Description, vbCritical
    
    ' Rilascia le risorse in caso di errore
    If Not MDL Is Nothing Then Set MDL = Nothing
    If Not db Is Nothing Then Set db = Nothing
    
    ' Chiudi il file di log se aperto
    If fileNumber > 0 Then Close #fileNumber
End Sub


'//06@STAMPA@TUTTO      = STAMPO TUTTI GLI OGGETTI DEL DB CORRENTE.
Private Sub Cmd_STAMPA_TUTTI_GLI_OGGETTI_DB_CORRENTE_Click()
        '//STAMPO TUTTI GLI OGGETTI
        Call Cmd_STAMPA_FORM_DB_CORRENTE_Click
        Call Cmd_STAMPA_TABLE_DB_CORRENTE_Click
        Call Cmd_STAMPA_QUERY_DB_CORRENTE_Click
        Call Cmd_STAMPA_REPORT_DB_CORRENTE_Click
        Call Cmd_STAMPA_MODULI_DB_CORRENTE_Click

End Sub


'//07@STAMPA@LOG@ERRORI = @STAMPO IL @FILE@LOG DI @GESTIONE DEGLI @ERRORI DELLA SINGOLA FUNZIONE O ROUTINE

Public Sub ScriviLogErrore(NomeRoutine As String, NumeroErrore As Long, MessaggioErrore As String)
      
      '//@FILE@LOG@ERRORI@MODELLO_(per l'attivazione della routine errori)
      '//.................................................................//
          'Dim NomeRoutine_s As String
          'Dim NumeroErrore_lng As Long
          'Dim MessaggioErrore_s  As String
          
          'NomeRoutine_s = "ROUTINE: Cmb_01_FORM_TXT_GotFocus"
          'NumeroErrore_lng = Err.number
          'MessaggioErrore_s = "ERRORE DI ESECUZIONE ROUTINE : " & Err.Description
          
          'ScriviLogErrore NomeRoutine_s, _
                          NumeroErrore_lng, _
                          MessaggioErrore_s
      
      '//.................................................................//
      


    Dim FileNum As Integer
    Dim TestoLog_s As String
    Dim FileLog_s As String
    
    '//reset variabili
    FileLog_s = ""
    FileLog_s = "\LOG_ERRORI_ROUTINE_E_FUNZIONI.TXT"
    
    ' Percorso e nome del file di log (stesso percorso del database)
    FileLog = CurrentProject.PATH & FileLog_s
    
    ' Formato del testo da scrivere nel file di log
    TestoLog_s = "--------------------------------------------------" & vbCrLf
    TestoLog_s = TestoLog_s & "Data:        " & Now & vbCrLf
    TestoLog_s = TestoLog_s & "Routine:     " & NomeRoutine & vbCrLf
    TestoLog_s = TestoLog_s & "Errore #:    " & NumeroErrore & vbCrLf
    TestoLog_s = TestoLog_s & "Descrizione: " & MessaggioErrore & vbCrLf
    TestoLog_s = TestoLog_s & "--------------------------------------------------" & vbCrLf & vbCrLf

    ' Apre il file in modalità Append (aggiunge il nuovo errore senza sovrascrivere i precedenti)
    FileNum = FreeFile()
    Open FileLog For Append As #FileNum
    Print #FileNum, TestoLog_s
    Close #FileNum
    
    MsgBox "LOG ERRORI STAMPATO NELLA PATH: " & CurrentProject.PATH & " FILE_LOG_ERRORI: " & FileLog_s
    
End Sub




'***********************************************************************************************************************
'                                    STAMPA OGGETTI *** FINE ***
'
'***********************************************************************************************************************









'//ROUTINE PER IL COLORE
'//


Sub Imposta_Colore_Sub()

    Dim curImportoDovuto As Currency, lngNero As Long
    Dim lngRosso As Long, lngGiallo As Long, lngBianco As Long

    If Not IsNull(Me!CurrentProject_txt.Value) Then
        curImportoDovuto = 0 'Me!CurrentProject_txt.Value
    Else
        Exit Sub
    End If
    
    lngRosso = RGB(255, 0, 0)
    lngNero = RGB(0, 0, 0)
    lngGiallo = RGB(255, 255, 0)
    lngBianco = RGB(255, 255, 255)
    
    If curImportoDovuto > 100 Then
        Me!CurrentProject_txt.BorderColor = lngRosso
        Me!CurrentProject_txt.ForeColor = lngRosso
        Me!CurrentProject_txt.BackColor = lngGiallo
    Else
        Me!CurrentProject_txt.BorderColor = lngNero
        '//colore testo
        Me!CurrentProject_txt.ForeColor = lngNero
        '//sfondo
        Me!CurrentProject_txt.BackColor = lngBianco
    End If
End



End Sub


'// IMPOSTA IL COLORE       *** FINE ***
'//-------------------------------------------------------------------------------------------------------------//



'//======================================================================================================//'
'//                         LA @PROGRESS@BAR  EVENTI'                         *** INIZIO ***
'//======================================================================================================//'
'//'
'//EVENTI ED ATTIVITA _ '
'//@PROGRESS@BAR@OPEN@DIM                       = DIM variabilie oggetto form _
'//@PROGRESS@BAR@OPEN@FORM                      = APRI OGGETTO FORM _
'//@PROGRESS@BAR@OPEN@UPDATE                    = AGGIORNA DATI PROGRESS BAR   *** forse non serve *** _
'//@PROGRESS@BAR@OPEN@CLOSE                     = CHIUDI LA FORM PROGRESS BAR _


'//TODO: CREARE LE 4 FUNZIONI PER LA RESTITUZIONE DEI VALORI:

  '//TODO: devi creare 4 funzioni per avere i 4 parametri per chiamare la progress bar: _
        01) FUNZIONE  NUMERO_COLLECTION         ----> restituisce NroCollection_i _
        @02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl _
        @03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
        04) FUNZIONE  TIPO_IMPORT               ----> RESTITUISCE TipoImport_s
'//todo: da eliminare
'//01) FUNZIONE  NUMERO_COLLECTION         ----> restituisce NroCollection_i
Public Function Nro_Collection_Pfunct() As Integer
    
    Dim formsCollection As Form_UTILITA_Msys_Frm21_frmProgressBar ' <-- Cast esplicito al tipo corretto
    Dim NroCollection_i As Integer
    On Error GoTo ErrorHandler ' Abilita la gestione degli errori
    
     '//calcolo parametri
      NroCollection_i = formsCollection.count             '//calcolo il numero di collection form da controllare
            
    '//chiama la funzione nro collection
      'i = Nro_Collection_Pfunct
    
    ' Calcola il numero di form aperti
    Nro_Collection_Pfunct = frm_obj.count
    Exit Function ' Esce dalla funzione in caso di successo

ErrorHandler:
    MsgBox "Errore durante il calcolo del numero di form aperti: " & Err.Description, vbCritical, "Errore " & Err.Number
    Nro_Collection_Pfunct = -1 ' Restituisce -1 in caso di errore per segnalare il problema
End Function

'//@02)@FUNZIONE@NUMERO@TWIPS                        ----> restituisce Twips_dbl _
    Note: riceve un integer ossia il numero oggetti e restituisce un _
    dbl ossia il NUMERO TWIPS di incremento per ogni oggetto.
Public Function Calcola_Twips_pFunct(par_NRO_COLLECTION_i As Integer) As Double
    On Error GoTo ErrorHandler ' Abilita la gestione degli errori

    Dim NroCollection_i As Integer
    
      
      '//chiama la funzione calcola twips
      'Dbl1 = Calcola_Twips_pFunct

    
    NroCollection_i = par_NRO_COLLECTION_i  ' Nro_Collection_Pfunct() ' Ottiene il numero di form aperti

    ' Verifica che il numero di form non sia 0 per evitare divisione per zero
    If NroCollection_i > 0 Then
        Calcola_Twips_pFunct = 5670 / NroCollection_i ' Calcola i twips
    Else
        Calcola_Twips_pFunct = 0 ' Se non ci sono form, i twips sono 0
    End If

    Exit Function ' Esce dalla funzione in caso di successo

ErrorHandler:
    MsgBox "Errore nel calcolo dei Twips: " & Err.Description, vbCritical, "Errore " & Err.Number
    Calcola_Twips_pFunct = -1 ' Restituisce -1 in caso di errore
End Function

'//@03)@FUNZIONE@PERCENTUALE_INCREMENTO              ----> restituisce PercIncr_dbl _
       note: riceve il numero della collection di oggetti un integer e _
       calcola la @%@di@incremento in double = % DI INCREMENTO PER OGNI OGGETTO.
Public Function PercentualeIncremento_pFunct(par_NRO_COLLECTION_i As Integer) As Double
    On Error GoTo ErrorHandler ' Abilita la gestione degli errori

    Dim NroCollection_i As Integer
    
      '//chiama la funzione percentuale incremento
      'Doble1 =PercentualeIncremento_pFunct (par_NRO_COLLECTION_i As Integer) As Double

    
    'NroCollection_i = Nro_Collection_Pfunct() ' Ottiene il numero di form aperti

    ' Evita la divisione per zero
    If par_NRO_COLLECTION_i > 0 Then
        '//la casella + di 10 e la percentuale 100/nro *10
        PercentualeIncremento_pFunct = (100 / par_NRO_COLLECTION_i) ' Calcola la percentuale
    Else
        PercentualeIncremento_pFunct = 0 ' Se non ci sono form, la percentuale è 0
    End If

    Exit Function ' Esce in caso di successo

ErrorHandler:
    MsgBox "Errore nel calcolo della percentuale di incremento: " & Err.Description, vbCritical, "Errore " & Err.Number
    PercentualeIncremento_pFunct = -1 ' Restituisce -1 in caso di errore
End Function

'//04) FUNZIONE  TIPO_IMPORT               ----> RESTITUISCE TipoImport_s
Public Function TipoImport_pFunct(par_TipoImportazione_s As String) As String
    
    On Error GoTo ErrorHandler ' Abilita la gestione degli errori


      '//chiama la funzione TIPO IMPORT
      'Str1 = PercentualeIncremento_pFunct (<Stringa>)


             '//scelgo il TIPO DI OGGETTO DA IMPORTARE
              Select Case par_TipoImportazione_s
              '//TIPO DI OGGETTO DA IMPORTARE
              Case Is = "\LOG_IMPORTA_FORM.TXT"
                        TipoImport_pFunct = "IMPORTA_FORM"
              Case Is = "\LOG_IMPORTA_MACRO.TXT"
                        TipoImport_pFunct = "IMPORTA_MACRO"
                        
             Case Is = "\LOG_IMPORTA_QUERY.TXT"
                        TipoImport_pFunct = "IMPORTA_QUERY"
                        
                        
              Case Is = "\LOG_IMPORTA_MODULI.TXT"
                        TipoImport_pFunct = "IMPORTA_MODULI"
                
               Case Is = "\LOG_IMPORTA_TABELLE.TXT"
                        TipoImport_pFunct = "IMPORTA_TABELLE"
                
              Case Else
                        '//CONTROLLO OGGETTO NON ESISTENTE
                        MsgBox "ATTENZIONE TIPO DI IMPORTAZIONE ERRATO O NON ESISTENTE - RESTITUISCO AVVISO: USCITA_DALLA_IMPORTAZIONE "
                        
                        TipoImport_pFunct = "USCITA_DALLA_IMPORTAZIONE"
                    
              End Select
              


    Exit Function ' Esce in caso di successo

ErrorHandler:
    MsgBox "Errore nella determinazione del tipo di importazione: " & Err.Description, vbCritical, "Errore " & Err.Number
    TipoImport_pFunct = "ERRORE" ' Restituisce un valore di errore
End Function


'//@DA@ATTIVARE= PROCEDURA ANCORA DA ATTIVARE
'//@PROGRESS@BAR@OPEN@FORM                      = APRI OGGETTO FORM
'// APRE LA FORM DI PROGRESSO
Public Sub ApriProgressBar()

                        
              
      ' Apri il form @UTILITA_Msys_Frm21_frmProgressBar prima di passare il parametro  @APRI@progress@bar
      '//---------------------------------------------------------------------------//
           DoCmd.OpenForm "UTILITA_Msys_Frm21_frmProgressBar", acNormal
           
           ' Aspetta che il form sia effettivamente caricato
           DoEvents
       
           
      '//---------------------------------------------------------------------------//

    
End Sub

'//@DA@ATTIVARE= PROCEDURA ANCORA DA ATTIVARE
'//@PROGRESS@BAR@OPEN@UPDATE                    = AGGIORNA DATI PROGRESS BAR   *** forse non serve ***
'// AGGIORNA LA PROGRESS BAR
Public Sub AggiornaProgressBar(par_NroCollection_i As Integer, _
                               par_ProgressPercentTwips_dbl As Double, _
                               par_Twips_dbl As Double, _
                               par_TotPerc_dbl As Double, _
                               par_TipoImport_s As String)

    


'// CHIAMO LA FORM PROGRESS BAR
'//...................................................................................................//
'//Note: passo i 5 parametri per aggiornare la label e il rettangono:
'//par_NroCollection_i    = Nro di oggetti collection
'//par_ProgressPercentTwips_dbl  = La progressione di incremento
'//par_Twips_dbl          = i twips per ogni oggetto da importare
'//par_PercIncr_dbl       = la percetuale di incremento per ogni oggetto da importare
'//par_TipoImport_s       = il tipo di oggetto di importazione TOT IMPORTATI - TOT NON IMPORTATI
               
'//AggiornaProgressBar(par_NroCollection_i As Integer, _
                       par_ProgressPercentTwips_dbl As Double, _
                       par_Twips_dbl As Double, _
                       par_PercIncr_dbl As Double, _
                       par_TipoImport_s As String)

'//...................................................................................................//

    
    
        
        
            
      '//01) AGGIORNO LA LABEL DELLA BARRA DI PROGRESSIONE @progress@bar@AGGIORNA@LABEL
      '//---------------------------------------------------------------------------//
          ' Ottieni un riferimento al form aperto e CAST alla classe specifica
          '//todo la CHIAMATA ALLA FOR DI BARRA DI PROGRESSO  inviando solo i 4 _
             parametri e quindi chiamarla UNA SOLA VOLTA anzieche 2 _
                CREA UNA FUNZIONE PUBBLICA DENOMINATA : _
                frm_obj.AttivaBarraDiProgressione con 3 parametri _
                par_PercIncr_dbl, _
                par_Twips_dbl, _
                par_TipoImport_s
             
          Set frm_obj = forms("UTILITA_Msys_Frm21_frmProgressBar")
          
          '//Aggiorno LA LABEL della barra di progressione
             frm_obj.AttivaBarraDiProgressione par_NroCollection_i, _
                                par_ProgressPercentTwips_dbl, _
                                par_Twips_dbl, _
                                par_TotPerc_dbl, _
                                par_TipoImport_s
          

        

          '//chiudo oggetto  *** todo: chiudere oggetto oppure no ???
          Set frm_obj = Nothing
          
      '//---------------------------------------------------------------------------//
 
 

           
    
    
    
    
    
    
End Sub

'//@DA@ATTIVARE= PROCEDURA ANCORA DA ATTIVARE
'//@PROGRESS@BAR@OPEN@CLOSE                     = CHIUDI LA FORM PROGRESS BAR
'// CHIUDE LA FORM CON RITARDO
Public Sub ChiudiProgressBar()
    
    
      
    '//CHIUDI LA FORM PROGRESS BAR
    '//.................................................................//
    '//NOTE: attivo la procedura di chiusura della form
        
        '//Call ChiudiProgressBar
    
    '//.................................................................//
    
    
    
    Dim T As Double
    T = Timer
    While Timer < T + 5
        DoEvents
    Wend
    
    '//CHIUDI_LA_FORM_PROGRESSBAR
    DoCmd.Close acForm, "UTILITA_Msys_Frm21_frmProgressBar", acSaveNo
    
    
    
End Sub




'//======================================================================================================//'
'//                         LA @PROGRESS@BAR  EVENTI'                         *** FINE ***
'//======================================================================================================//'






'***********************************************************************************************************************
'                                    @ROUTINE@GENERICHE              **** FINE ****
'
'***********************************************************************************************************************





'//======================================================================================================//'
'//                         @APRI@FILE                        *** FINE ***
'//======================================================================================================//'
'//@30@ROUTINE@APRI@FILE                            = ROUTINE DI APERTURA DEL FILE
'//
'//CODICE -----> @APRI@FILE
'//EVENTI ED ATTIVITA _ '
'//@01@APRI@FILE@RESET                              = RESET DELLE VARIABILI _




'//APRI_FILE
'//==================================================================================================================//

Private Sub sourceDBPath_s_Txt_DblClick(Cancel As Integer)
    On Error GoTo Err_sourceDBPath_s_Txt_DblClick
    
    '//@01@APRI@FILE@RESET                              = RESET DELLE VARIABILI _
    ' Reset variabili di gestione errori
    ProceduraMessaggioErrore_s = "GESTIONE DEI FALDONI"
    ProceduraAttivaEseguita_s = "sourceDBPath_s_Txt_DblClick"
    
    ' Recupero il percorso dalla casella di testo
    Dim filePath As String
    filePath = Trim(sourceDBPath_s_Txt.Value)
    
    ' Verifica se il percorso è vuoto
    If filePath = "" Then
        MsgBox "Nessun percorso specificato.", vbExclamation, "Attenzione"
        Exit Sub
    End If
    
    ' Controllo se il file esiste
    If Dir$(filePath) = "" Then
        MsgBox "Il file specificato non esiste." & vbCrLf & filePath, vbCritical, "Errore"
        Exit Sub
    End If
    
    ' Tentativo di apertura con FollowHyperlink
    On Error Resume Next
    Application.FollowHyperlink filePath
    If Err.Number <> 0 Then
        Err.Clear
        
        ' Se FollowHyperlink fallisce, provo con Shell e Access
        Dim accessPath As String
        accessPath = Me.sourceDBPath_s_Txt
        
        ' Controllo se Access è installato nel percorso predefinito
        If Dir$(accessPath) = "" Then
            accessPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE"
        End If
        
        ' Se Access è trovato, uso Shell per aprire il file
        If Dir$(accessPath) <> "" Then
            Shell """" & accessPath & """ """ & filePath & """", vbNormalFocus
            DoEvents
        Else
            MsgBox "Microsoft Access non trovato. Impossibile aprire il file.", vbCritical, "Errore"
        End If
    End If
    On Error GoTo 0

    Exit Sub
    
Err_sourceDBPath_s_Txt_DblClick:
    MsgBox "Si è verificato un errore: " & Err.Description, vbCritical, "Errore"
    Resume Exit_sub

Exit_sub:
    Exit Sub
End Sub

'//==================================================================================================================//




'//======================================================================================================//'
'//                         @APRI@FILE                        *** FINE ***
'//======================================================================================================//'

















