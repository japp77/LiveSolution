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
'applicazione che viene utilizzato per eseguire le funzionalit?
'ed i relativi processi del gestore.
Public TheApp As Application



'Oggetto Semaforo usato per gestire i conflitti di multiutenza.
Public gSemaphore As Semaforo.dmtSemaphore



'La variabile globale Application_Name ? valorizzata nella Sub Main.
Public Application_Name As String


'La variabile globale Current_Process_ID ? valorizzata nella Sub Main
'e rappresenta l'ID del processo in esecuzione.
Public Current_Process_ID As Long

'////////////////////////////////////////////////////////
'Impostare questa costante con il nome del
'processo previsto per questa manutenzione
'////////////////////////////////////////////////////////
Public Const PROCESS_NAME = "Manutenzione"

Public rsContrattiReg As ADODB.Recordset
Public NumeroRecordContratti As Long

Public STRINGA_SQL As String
Public Const REGISTRY_KEY_PERS As String = "Point Office Company srl"
Public Const SECTION_REGISTRY_KEY_PERS As String = "ContrattiPRO"
Public TipoAvvio As Long

Public LINK_AZIENDA As Long
Public LINK_ATTIVITA_AZIENDA As Long
Public LINK_FILIALE As Long
Public LINK_ANAGRAFICA_AZIENDA As Long
Public LINK_UTENTE As Long


'**+
'Nome                   : Main
'Parametri              : Nessuno
'Valori di ritorno      : Nessuno
'Funzionalit?           : In questa Sub vengono eseguite
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
    FrmInizio.Icon = gResource.GetIcon(IDI_DIAMANTE16)
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
        
        '..............................................................................................................................
        'Gestione della Semaforo
        '..............................................................................................................................
        
        ' Verifica se l'applicazione corrente ? bloccata da altri gestori.
        ' (Il controllo avviene sul Tipo Oggetto correntemente trattato ovvero Current_Process_ID)
        If Not gSemaphore.IsActionAvailable(Current_Process_ID, SemAllObjects, SemAllActions) Then
            '-------------------------------------------------------------
            'Il programma ? bloccato da un'altra manutenzione in esecuzione.
            'Sar? pertanto terminato.
            '-------------------------------------------------------------
    
            'Scarica il form
            Unload FrmInizio
    
            'Termina il programma
            End
        End If
        
        
        '----------------------------------------------------
        'Il programma non ? bloccato e prosegue normalmente.
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
            
            Case "Manutenzione"  ' <---Tipicamente ? questo l'unico processo gestito
            
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
        'Questo programma pu? essere eseguito solo all'interno dell'applicativo Diamante.
        'Prima di TheApp.Run si ha TheApp.FunctionName = "" allora nella Caption del messaggio si avr? TheApp.Name.
        sbMsgInfo gResource.GetMessage(MESS_RUNOUTOFDIAMANTE), IIf(TheApp.FunctionName <> "", TheApp.FunctionName, TheApp.Name)
    Else
        MsgBox "Errore indefinito", vbCritical, App.EXEName
    End If
    
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
Public Function fnGetNewKey(Tabella As String, CampoKey As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
        
        sSQL = "SELECT " & CampoKey & " FROM " & Tabella & " ORDER BY " & CampoKey & " DESC"
        
        Set rs = CnDMT.OpenResultset(fnAnsi2Jet(sSQL))
    
        If rs.EOF = True Then
        
            fnGetNewKey = 1
    
        Else
            
            fnGetNewKey = fnNotNullN(rs.adoColumns(CampoKey)) + 1
    
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
'Funzionalit?:
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

'Funzionalit?:
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
    
        'Decommentare questa riga se in SemaphoreLock ? stato fatto altrettanto.
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

Public Function ConnessioneADODBLib() As Boolean
On Error GoTo ERR_ConnessioneADODBLib
Dim StringaDiConnessione As String
Dim Motore As String
Dim Catalogo As String
Dim IDFiliale As Long

    
    TipoAvvio = fnNotNullN(GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Tipo Avvio"))
    
    Set gResource = New Resource
    
    If TipoAvvio = 0 Then
        REGISTRY_KEY = Trim(gResource.GetMessage(LBL_REGISTRY_KEY))
        
        Utente = GetSetting(REGISTRY_KEY, "MenuSettings", "LASTUSER")
        Password = fnCryptString(GetSetting(REGISTRY_KEY, "MenuSettings", "LASTUSERPWD"))
        
        IDFiliale = MenuOptions.LastBranch
        
        If Right(Trim(MenuOptions.ConnectionString), 1) = ";" Then
            StringaDiConnessione = MenuOptions.ConnectionString
        Else
            StringaDiConnessione = MenuOptions.ConnectionString & ";"
        End If
    Else
        IDFiliale = GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "IDFiliale")
        Motore = GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Motore")
        Catalogo = GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Catalogo")
        
        Utente = GetSetting(gResource.GetMessage(LBL_REGISTRY_KEY), App.EXEName, "Utente", "")
        Password = GetSetting(gResource.GetMessage(LBL_REGISTRY_KEY), App.EXEName, "Password", "")
        
        StringaDiConnessione = "Provider=SQLOLEDB.1;Initial Catalog=" & Catalogo & ";Data Source=" & Motore & ";"

    End If
    
    Set CnDMT = DmtOleDbLib.adoEnvironments(0).OpenConnection((StringaDiConnessione & "User Id=" & Utente & ";Password=" & Password))
    
    'Set CnDMT = DmtOleDbLib.adoEnvironments(0).OpenConnection((StringaDiConnessione & "User Id=" & Utente & ";Password="))


    PrelevaAzienda Utente, IDFiliale
    
    ConnessioneADODBLib = True
    
Exit Function

ERR_ConnessioneADODBLib:
    MsgBox Err.Description, vbCritical, "Connessione"
    ConnessioneADODBLib = False
End Function
Public Sub PrelevaAzienda(E_NomeUtente As String, IDFiliale As Long)
Dim TmpFiliale As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT Anagrafica.IDAnagrafica, Azienda.IDAzienda, Anagrafica.Anagrafica, AttivitaAzienda.IDAttivitaAzienda, "
    sSQL = sSQL & "AttivitaAzienda.AttivitaAzienda, Filiale.IDFiliale, Filiale.Filiale"
    sSQL = sSQL & " FROM (Anagrafica INNER JOIN Azienda ON Anagrafica.IDAnagrafica = Azienda.IDAnagrafica)"
    sSQL = sSQL & " INNER JOIN (Filiale INNER JOIN AttivitaAzienda ON Filiale.IDAttivitaAzienda = AttivitaAzienda.IDAttivitaAzienda)"
    sSQL = sSQL & " ON Azienda.IDAzienda = AttivitaAzienda.IDAzienda"
    sSQL = sSQL & " WHERE Filiale.IDFiliale=" & IDFiliale
    
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        LINK_AZIENDA = fnNotNullN(rs!IDAzienda)
        LINK_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
        LINK_FILIALE = fnNotNullN(rs!IDFiliale)
        LINK_UTENTE = GET_LINK_UTENTE(E_NomeUtente)
        LINK_ANAGRAFICA_AZIENDA = fnNotNullN(rs!IDAnagrafica)
    rs.CloseResultset
    Set rs = Nothing
    End If
    
End Sub

Private Function GET_LINK_UTENTE(NomeUtente As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUtente FROM Utente "
sSQL = sSQL & "WHERE Utente=" & fnNormString(NomeUtente)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_UTENTE = 0
Else
    GET_LINK_UTENTE = fnNotNullN(rs!IDUtente)
End If

rs.CloseResultset
Set rs = Nothing
End Function
