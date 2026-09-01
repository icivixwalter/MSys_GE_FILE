@PROGETTO@GE@FILE
   Note
      progetto di gestione dei file e delle cartelle con le API WINDOWS ed i codici per 
      identificare i progetti sono :

      CODICE_ACCESS     -----> MSys_GE_FILE
      GROUP_QUOTE_s     -----> GROUP_QUOTA_MSys_GE_FILE
      NOTE_PROGETTO_s   -----> PROGETTO GESTIONE FILE
      
MSys_TAB502_Qry01_01_VISUAL_PAG_SEMPLICE

AGGIORNAMENTI
   2025_10_28_MODULO_LIBRERIE = creato un modulo per l'inserimento automatico delle librerie.

   2025_10_28_MODULO_LIBRERIE = AGGIORNATA LA CLASSE GESTIONE OGGETTI
      Migliorato l'esempio di attivazione e apertura file con le api windows.


PARAMETRI_VBA
   Note
      per gestire al meglio vba e Access 2011 ho creato una variabile di sistema che identifica
      Access 2011 denominata come VARIABILE DI SISTEMA = "ACCESS_2011"

      VARIABILE DI AMBIENTE ACCESS CASA
         Per casa la path PER L'ESEGUIBILE DI ACCESS E PER LA COSTRUZIONE DELLA VARIABILE DI AMBIENTE è la seguente :
            variabile denominata: "ACCESS_2011" = C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE

      VARIABILE DI AMBIENTE ACCESS UFFICIO
            variabile denominata: "ACCESS_2011" = C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE



MACRO ESPORTA OGGETTI to 
   creato la macro esporta oggetti in LLPP_ATTI_GESTIONE.mdb c:\GESTIONI\GESTIONE_LLPP\25_GESTIONE_LLPP\LLPP_ARCHIVI_MDB\LLPP_ATTI_GESTIONE.mdb

   denominata:
      ESPORTA_Mcr_01_GEFILE_to_LLPP_ATTI_GESTIONE_mdb

RUST DOCUMENTAZIONE HTML
   Note
      costruita procedura in Rust che crea una documentazione in html
      della classe form. Per creare la libreria in html occorre:
         1) estrarre la classe della form o altra classe i
         2) fornire la path completa ed il nome della classe alla procedura main come 
            in questo esempio: c:\tmp\classe.cls.
            Per assegnare il percorso completo segui questo codice ----> @path@completa+file   

         3) Attivare la procedura rust per la costruzione della documentazione in html.
            Vai in mani rust ----> 
