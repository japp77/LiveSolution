Attribute VB_Name = "ModMain"
Public CnDMT As DmtOleDbLib.adoConnection
Public gResource As Resource
'**********************VARIABILI GLOBALI AZIENDA**************************
    Public VarIDEsercizio As Long
    
'*************************************************************************

Public b_Loading As Boolean

Public Const NUMERICTYPE = 0
Public Const STRINGTYPE = 1
Public Const DATETYPE = 2

Public Var_DaDataRinnovo As String
Public Var_ADataRinnovo As String
Public Link_cliente_Ric As Long
Public Link_Tipo_Contratto_Ric As Long

Public NumeroRecord As Long

'VARIABILI DEL TIPO CONTRATTO
Public Descrizione_Tipo_Contratto As String

'VARIABILI DELLA DURATA CONTRATTO
Public Mesi_Durata_Contratto As Long
Public Giorni_Durata_Contratto As Long

'VARIABILI DEL TIPO RINNOVO
Public Mesi_Rinnovo_Contratto As Long
Public Giorni_Rinnovo_Contratto As Long 'Indica il giorno esatto del mese
Public AnnoPrecedente_Rinnovo_Contratto As Long 'Indica se la data di decorrenza deve concidere con quella dell'anno precedente

'VARIABILI DEL TIPO DURATA ASSISTENZA
Public Descrizione_Tipo_Assistenza As String
Public Mesi_Tipo_Assistenza As Long
Public Giorni_Tipo_Assistenza As Long 'Indica il giorno esatto del mese


'VARIABILI DEL TIPO DI RATEIZZAZIONE
Public Mesi_Rate As Long
Public Numero_Rate As Long
Public Pagamento_Anticipato_Periodo As Boolean
Public Rata_Iniziale As Long
Public Anno_Solare As Long


Public MesiRimanentiScadenzaContratto As Long

Public PercentualeIstat As Double
Public Link_Istat As Long

Public Password As String
Public Utente As String


Public LINK_SEZIONALE_RATE As Long



'API di uso comune.
Public Declare Function fnAnsi2Jet Lib "Diamante.dll" Alias "fnAnsi2jet" (ByVal sSQL As String) As String
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long)
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long


Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOP = 0
Public Const WM_SETREDRAW = &HB

'Costanti globali
Public Const TOTAL_CONTROLS_NUMBER = 10
Public Const SPLITLIMIT = 1000
Public Const SRCNEXT = 1
Public Const SRCPREVIOUS = 2
Public Const HELP_FINDER = &HB
Public Const HELP_CONTEXT = &H1
Public Const NATIVE_LANGUAGE = 1
Public Const URL_DIAMANTE = "http://www.diamante.it"


'Elenco errori
Public Const ERR_TABLE_STRUCT = vbObjectError + 10000
Public Const ERR_NO_DEFAULT_TABLEVIEW = vbObjectError + 10001
Public Const ERR_NO_PROCESSES = vbObjectError + 10002
Public Const ERR_NDELFILTER = vbObjectError + 2500



'La variabile globale TheApp mantiene un riferimento all'oggetto
'applicazione che viene utilizzato per eseguire le funzionalità
'ed i relativi processi del gestore.
Public TheApp As Application



'Oggetto Semaforo usato per gestire i conflitti di multiutenza.
Public gSemaphore As Semaforo.dmtSemaphore



'La variabile globale Application_Name è valorizzata nella Sub Main.
Public Application_Name As String


'La variabile globale Current_Process_ID è valorizzata nella Sub Main
'e rappresenta l'ID del processo in esecuzione.
Public Current_Process_ID As Long

'////////////////////////////////////////////////////////
'Impostare questa costante con il nome del
'processo previsto per questa manutenzione
'////////////////////////////////////////////////////////
Public Const PROCESS_NAME = "Manutenzione"


'**+
'Nome                   : Main
'Parametri              : Nessuno
'Valori di ritorno      : Nessuno
'Funzionalità           : In questa Sub vengono eseguite
'                           : le operazioni di Startup.
'**/
Sub Main()
    Dim Proc  As DMTRunAppLib.Process
    Dim ErrorMsg As String

    On Error GoTo ErrorHandler
    
    'L'Applicazione.
    Set TheApp = New Application
    
    'Il nome della applicazione.
   TheApp.Name = App.EXEName
    
    
    DmtRegistry2.EXEName = App.EXEName
    
    'L'oggetto che si occupa della lettura delle risorse.
    Set gResource = New Resource
    'FrmInizio.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    'Carica il form del Wizard senza mostrarlo.
    Load FrmInizio
    
    Set FrmInizio.Application = TheApp
    'Esegue l'applicazione
    TheApp.Run FrmInizio.hWnd
    
    'Inizializza l'oggetto Semaforo per la gestione dei conflitti di multiutenza.
    InitSemaphore
    
    'Viene individuato il nome della funzione.
    Application_Name = TheApp.FunctionName
    
    'Lettura file di help
    'App.HelpFile = TheApp.Path & "\Diamante.hlp"


    '----------------------------------------------------------
    'Ciclo sui processi della funzione
    '----------------------------------------------------------
    For Each Proc In TheApp.Processes
        
        'L'identificativo del Tipo Oggetto correntemente gestito.
        Current_Process_ID = Proc.IDocType.ID
        FrmInizio.InitControlli
        '..............................................................................................................................
        'Gestione della Semaforo
        '..............................................................................................................................
        
        ' Verifica se l'applicazione corrente è bloccata da altri gestori.
        ' (Il controllo avviene sul Tipo Oggetto correntemente trattato ovvero Current_Process_ID)
        If Not gSemaphore.IsActionAvailable(Current_Process_ID, SemAllObjects, SemAllActions) Then
            '-------------------------------------------------------------
            'Il programma è bloccato da un'altra manutenzione in esecuzione.
            'Sarà pertanto terminato.
            '-------------------------------------------------------------
    
            'Scarica il form
            Unload FrmInizio
    
            'Termina il programma
            End
        End If
        
        
        '----------------------------------------------------
        'Il programma non è bloccato e prosegue normalmente.
        '----------------------------------------------------
        
        'Ripulisce la tabella semaforo.
        'Se era avvenuto un crash di sistema questo garantisce il ripristino della situazione.
        SemaphoreUnlock
        
        'Imposta gli eventuali blocchi (semaforo) su altre manutenzioni.
        SemaphoreLock
        
        '..............................................................................................................................
        '..............................................................................................................................
        
        
        
        
        '-------------------------------------------------------------------------------------
        'In funzione del processo da gestire la manutenzione si deve comportare di conseguenza
        '-------------------------------------------------------------------------------------
        Select Case Proc.Name
        
            '*
            'Inserire qui il codice per la gestione del processo (o dei processi)
            '*
            
            Case "Manutenzione"  ' <---Tipicamente è questo l'unico processo gestito
            
            '   For Each Parameter In Proc.Parameters
            '       Select Case Parameter.Name
            '       *
            '       Inserire il codice per la gestione del parametro
            '       *
            '       Case ParameterName??????
            '       End Select
            '   next
                  
                '-------- Di solito --------
                
                'Inizializzazioni preliminari
                'FrmIniziotControlli
                FrmInizio.ConnessioneADO
                'Viene mostrato il form.
                FrmInizio.Show
    
                b_Loading = True
                
            Case Else
                ErrorMsg = "No processes to execute" & vbCrLf
                ErrorMsg = ErrorMsg & "This application is able to execute these processes:" & vbCrLf
                '*
                '/////////////////////////////////////////////////////
                'Inserire i processi che l'applicazione sa eseguire
                '/////////////////////////////////////////////////////
                '*
                'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE & vbCrLf
                'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE_EXTENDED_DATABASE & vbCrLf
                'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE_DA_SHELL & vbCrLf
                Err.Raise ERR_NO_PROCESSES, , ErrorMsg
        End Select
    Next


    
    Exit Sub
ErrorHandler:

    'Ripulisce la tabella Semaforo
    SemaphoreUnlock
    
    'Scarica il form
    Unload FrmInizio
    
    If Err.Number = 1 + vbObjectError Then
        'Questo programma può essere eseguito solo all'interno dell'applicativo Diamante.
        'Prima di TheApp.Run si ha TheApp.FunctionName = "" allora nella Caption del messaggio si avrà TheApp.Name.
        sbMsgInfo gResource.GetMessage(MESS_RUNOUTOFDIAMANTE), IIf(TheApp.FunctionName <> "", TheApp.FunctionName, TheApp.Name)
    Else
        Err.Raise Err.Number
    End If
    
End Sub






Public Sub PrelevaAzienda()

    Dim TmpFiliale As String
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    TmpFiliale = GetSetting("DMT Professional v3.1", "MenuSettings", "LASTBRANCH")
    
    sSQL = "SELECT Azienda.IDAzienda, Anagrafica.Anagrafica, AttivitaAzienda.IDAttivitaAzienda, AttivitaAzienda.AttivitaAzienda, Filiale.IDFiliale, Filiale.Filiale"
    sSQL = sSQL & " FROM (Anagrafica INNER JOIN Azienda ON Anagrafica.IDAnagrafica = Azienda.IDAnagrafica) INNER JOIN (Filiale INNER JOIN AttivitaAzienda ON Filiale.IDAttivitaAzienda = AttivitaAzienda.IDAttivitaAzienda) ON Azienda.IDAzienda = AttivitaAzienda.IDAzienda"
    sSQL = sSQL & " WHERE (((Filiale.IDFiliale)=" & TmpFiliale & "))"
    
    
    Set rs = CnDMT.OpenResultset(sSQL)
        VarIDAttivitaAzienda = rs!IDAttivitaAzienda
        
    rs.CloseResultset
    Set rs = Nothing
    
    
End Sub
Private Function fncEsercizio() As String
    fncEsercizio = ""
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "Select IDEsercizio, Esercizio"
    sSQL = sSQL & " FROM Esercizio"
    sSQL = sSQL & " WHERE (IDAzienda = " & TheApp.IDFirm & ")"
    sSQL = sSQL & " AND (IDTipoEsercizio = 1)"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        VarIDEsercizio = rs!IDEsercizio
        fncEsercizio = rs!Esercizio
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Public Sub ChiusuraConnessione()
    
    
End Sub
Public Sub TipoContratto(IDTipoContratto As Long)
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT DescrizioneAggiuntiva "
    sSQL = sSQL & "FROM RV_POTipoContratto "
    sSQL = sSQL & "WHERE IDRV_POTipoContratto=" & IDTipoContratto
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        Descrizione_Tipo_Contratto = ""
    Else
        Descrizione_Tipo_Contratto = fnNotNull(rs!DescrizioneAggiuntiva)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Sub
Public Sub DurataAssistenza(IDTipoDurataAssistenza As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT * FROM RV_POTipoDurataAssistenza "
    sSQL = sSQL & "WHERE IDRV_POTipoDurataAssistenza=" & IDTipoDurataAssistenza
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        Descrizione_Tipo_Assistenza = fnNotNull(rs!TipoDurataAssistenza)
        Mesi_Tipo_Assistenza = fnNotNullN(rs!Mesi)
        Giorni_Tipo_Assistenza = fnNotNullN(rs!Giorni)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Sub
Public Sub DurataContratto(IDDurataContratto As Long)
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT * FROM RV_PODurataContratto "
    sSQL = sSQL & "WHERE IDRV_PODurataContratto=" & IDDurataContratto
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        Mesi_Durata_Contratto = fnNotNullN(rs!Mesi)
        Giorni_Durata_Contratto = fnNotNullN(rs!Giorni)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Sub
Public Sub TipoRinnovo(IDTipoRinnovo As Long)
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    

    sSQL = "SELECT * FROM RV_POTipoRinnovo "
    sSQL = sSQL & "WHERE IDRV_POTipoRinnovo=" & IDTipoRinnovo
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        Mesi_Rinnovo_Contratto = fnNotNullN(rs!Mesi)
        Giorni_Rinnovo_Contratto = fnNotNullN(rs!Giorni)
        AnnoPrecedente_Rinnovo_Contratto = fnNotNullN(rs!AnnoPrecedente)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Sub
Public Sub TipoRateizzazione(IDRateizzazione As Long)
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT * FROM RV_PORateizzazione "
    sSQL = sSQL & "WHERE IDRV_PORateizzazione=" & IDRateizzazione
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        Mesi_Rate = fnNotNullN(rs!Mesi)
        Numero_Rate = fnNotNullN(rs!numerorate)
        Pagamento_Anticipato_Periodo = fnNotNullN(rs!PagamentoInizioPeriodo)
        Rata_Iniziale = fnNotNullN(rs!RataInizialeRataFinale)
        Anno_Solare = fnNotNullN(rs!AnnoSolare)
    Else
        Mesi_Rate = 0
        Numero_Rate = 0
        Pagamento_Anticipato_Periodo = False
        Rata_Iniziale = 0
        Anno_Solare = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing

End Sub
Public Function fnGetNewKey(Tabella As String, CampoKey As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
        
        sSQL = "SELECT " & CampoKey & " FROM " & Tabella & " ORDER BY " & CampoKey & " DESC"
        
        Set rs = CnDMT.OpenResultset(fnAnsi2Jet(sSQL))
    
        If rs.EOF = True Then
        
            fnGetNewKey = 1
    
        Else
            
            fnGetNewKey = fnNotNullN(rs.adoColumns(CampoKey).Value) + 1
    
        End If

        rs.CloseResultset
        Set rs = Nothing
    

    
End Function
 Private Sub InitSemaphore()
    
    Set gSemaphore = New Semaforo.dmtSemaphore
    Set gSemaphore.Database = TheApp.Database.Connection
    Set gSemaphore.objRes = gResource
    gSemaphore.IDUser = TheApp.IDUser
    gSemaphore.IDBranch = TheApp.Branch
    gSemaphore.IDFunction = TheApp.FunctionID
    
End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 12/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: SemaphoreLock
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'                 ////////////////////////////////////////////////////////////////////////
'                     Impostare qui gli eventuali blocchi sulle altre manutenzioni
'                 ////////////////////////////////////////////////////////////////////////
'**/
Public Sub SemaphoreLock()

    If Not gSemaphore Is Nothing Then
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        'Personalizzare, se necessario, le righe sottostanti
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        
'        gSemaphore.SetObjectAction TO_TIPO_OGGETTO_XXX, SemAllObjects, SemAllActions
'        gSemaphore.SetObjectAction TO_TIPO_OGGETTO_YYY, SemAllObjects, SemAllActions
'        gSemaphore.SetObjectAction TO_TIPO_OGGETTO_ZZZ, SemAllObjects, SemAllActions

        'Decommentare questa riga se si deve impedire ad un altro utente di entrare nella manutenzione corrente.
'        gSemaphore.SetObjectAction Current_Process_ID, SemAllObjects, SemAllActions

    End If
    
End Sub

'**+
'Autore: Diamante s.p.a
'Data creazione: 12/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: SemaphoreUnlock
'
'Parametri:
'
'Valori di ritorno:

'Funzionalità:
'                 //////////////////////////////////////////////////////////////////////////////////////////////////
'                     Sbloccare qui le altre manutenzioni (bloccate precedentemente in SemaphoreLock)
'                 //////////////////////////////////////////////////////////////////////////////////////////////////
'
'**/
Public Sub SemaphoreUnlock()

    If Not gSemaphore Is Nothing Then
    
        'Ripulisce la tabella semaforo per quanto riguarda il Tipo Oggetto e l'utente correnti
        gSemaphore.ClearObjectAction Current_Process_ID, SemAllObjects, SemAllActions
        
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        'Personalizzare, se necessario, le righe sottostanti
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        
        'Sblocca le manutenzioni bloccate precedentemente
'        gSemaphore.ClearObjectAction TO_TIPO_OGGETTO_XXX, SemAllObjects, SemAllActions
'        gSemaphore.ClearObjectAction TO_TIPO_OGGETTO_YYY, SemAllObjects, SemAllActions
'        gSemaphore.ClearObjectAction TO_TIPO_OGGETTO_ZZZ, SemAllObjects, SemAllActions
    
        'Decommentare questa riga se in SemaphoreLock è stato fatto altrettanto.
'        gSemaphore.ClearObjectAction Current_Process_ID, SemAllObjects, SemAllActions
    
    End If
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    
    If Not (CnDMT Is Nothing) Then
        CnDMT.CloseConnection
        Set CnDMT = Nothing
    End If
    
    'Sblocca gli eventuali gestori bloccati da questa manutenzione
    SemaphoreUnlock
    
    '--------------------------------
    'Distruzione degli oggetti allocati.
    '--------------------------------
    
    Set gSemaphore = Nothing
    
    
End Sub

