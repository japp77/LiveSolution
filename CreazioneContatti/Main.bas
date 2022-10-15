Attribute VB_Name = "ModMain"
'Variabile archivio Diamante
Public CnDMT As DmtOleDbLib.adoConnection
Public oRes As DmtOleDbLib.adoResultset


Public gResource As Resource
'Variabile che indica se è un'utente nuovo da configurare o
'è un utente da modificare i parametri di configurazione
'0 = Nuova configurazione
'1 = Configurazione da modificare
Public VarNewUser As Integer

'**********************VARIABILI GLOBALI AZIENDA**************************
Public VarIDAzienda As Long
Public VarIDAttivitaAzienda As Long
Public VarIDFiliale As Long
Public VarIDUtente As Long

'*******************VARIABILI DI PASSAGGIO TRA I FORM*********************
Public VarIDCliente As Long

'*******************COSTANTI DEL PROGRAMMA*********************************
Private Const NUMERICTYPE = 0
Private Const STRINGTYPE = 1
Private Const DATETYPE = 2

'Variabile per controllare i comandi tra una form e l'altra

Public DinamicNumber As Integer


'Variabile di opzione scelta per l'esportazione
Public optOperazione As Integer
Public optTipoCliente As Long
Public optTipoContratto As Long
Public optTipoContatto As Long

Public NumeroRecord As Long

Public Password As String

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

Public b_Loading As Boolean



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
    FrmMain.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    'Carica il form del Wizard senza mostrarlo.
    Load FrmMain
    
    Set FrmMain.Application = TheApp
    'Esegue l'applicazione
    TheApp.Run FrmMain.hWnd
    
    
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
        FrmMain.InitControlli
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
            Unload FrmMain
    
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
                'FrmMaintControlli
                FrmMain.ConnessioneADO
                'Viene mostrato il form.
                FrmMain.Show
    
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
    Unload FrmMain
    
    If Err.Number = 1 + vbObjectError Then
        'Questo programma può essere eseguito solo all'interno dell'applicativo Diamante.
        'Prima di TheApp.Run si ha TheApp.FunctionName = "" allora nella Caption del messaggio si avrà TheApp.Name.
        sbMsgInfo gResource.GetMessage(MESS_RUNOUTOFDIAMANTE), IIf(TheApp.FunctionName <> "", TheApp.FunctionName, TheApp.Name)
    Else
        Err.Raise Err.Number
    End If
    
End Sub







Public Sub PrelevaAzienda()

    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
   
    
    sSQL = "SELECT Azienda.IDAzienda, Anagrafica.Anagrafica, AttivitaAzienda.IDAttivitaAzienda, AttivitaAzienda.AttivitaAzienda, Filiale.IDFiliale, Filiale.Filiale"
    sSQL = sSQL & " FROM (Anagrafica INNER JOIN Azienda ON Anagrafica.IDAnagrafica = Azienda.IDAnagrafica) INNER JOIN (Filiale INNER JOIN AttivitaAzienda ON Filiale.IDAttivitaAzienda = AttivitaAzienda.IDAttivitaAzienda) ON Azienda.IDAzienda = AttivitaAzienda.IDAzienda"
    sSQL = sSQL & " WHERE (((Filiale.IDFiliale)=" & TheApp.Branch & "))"
    
    
    Set rs = CnDMT.OpenResultset(sSQL)
        
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


Public Sub CreazioneContatti()
'On Error Resume Next
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim sSQL_WHERE As String
    Dim UnitaProgresso As Double
    Dim TotaleRecord As Long
    
    CnDMT.Execute "DELETE FROM RV_POTMPContact"
    
    Select Case optOperazione
    
        Case 1
            sSQL = "SELECT RV_POContactDettaglio.IDRV_POContactDettaglio, RV_POContactDettaglio.IDRV_POContact, RV_POContactDettaglio.Email, "
            sSQL = sSQL & "RV_POContactDettaglio.IDTitolo, RV_POContactDettaglio.Nominativo, RV_POContactDettaglio.Mansione, dbo.RV_POContactDettaglio.Fax,"
            sSQL = sSQL & "RV_POContactDettaglio.Telefono1, RV_POContactDettaglio.Cellulare, RV_POContactDettaglio.Telefono2,"
            sSQL = sSQL & "RV_POTipoContattoDettaglio.IDRV_POTipoContatto, dbo.SitoPerAnagrafica.SitoPerAnagrafica, dbo.RV_POContact.IDRVTipoCliente,"
            sSQL = sSQL & "Comune.Comune, Provincia.Provincia, Comune_1.Comune AS ComuneFiliale, Provincia_1.Provincia AS ProvinciaFiliale,"
            sSQL = sSQL & "Anagrafica.Indirizzo, Anagrafica.Cap, SitoPerAnagrafica.Indirizzo AS IndirizzoFiliale, SitoPerAnagrafica.Cap AS CapFiliale,"
            sSQL = sSQL & "RV_POContact.IDAnagrafica , RV_POContact.IDSitoPerAnagrafica, RV_POContratto.IDTipoContratto, Anagrafica.Anagrafica "
            sSQL = sSQL & "FROM RV_POContact LEFT OUTER JOIN "
            sSQL = sSQL & "RV_POContratto ON RV_POContact.IDSitoPerAnagrafica = RV_POContratto.IDSitoPerAnagrafica AND "
            sSQL = sSQL & "RV_POContact.IDAnagrafica = RV_POContratto.IDAnagrafica LEFT OUTER JOIN "
            sSQL = sSQL & "SitoPerAnagrafica LEFT OUTER JOIN "
            sSQL = sSQL & "Provincia Provincia_1 RIGHT OUTER JOIN "
            sSQL = sSQL & "Comune Comune_1 ON Provincia_1.IDProvincia = Comune_1.IDProvincia ON SitoPerAnagrafica.IDComune = Comune_1.IDComune ON "
            sSQL = sSQL & "RV_POContact.IDSitoPerAnagrafica = SitoPerAnagrafica.IDSitoPerAnagrafica LEFT OUTER JOIN "
            sSQL = sSQL & "Provincia RIGHT OUTER JOIN "
            sSQL = sSQL & "Comune ON Provincia.IDProvincia = Comune.IDProvincia RIGHT OUTER JOIN "
            sSQL = sSQL & "Anagrafica ON Comune.IDComune = Anagrafica.IDComune ON "
            sSQL = sSQL & "RV_POContact.IDAnagrafica = Anagrafica.IDAnagrafica RIGHT OUTER JOIN "
            sSQL = sSQL & "RV_POContactDettaglio ON RV_POContact.IDRV_POContact = RV_POContactDettaglio.IDRV_POContact LEFT OUTER JOIN "
            sSQL = sSQL & "RV_POTipoContatto RIGHT OUTER JOIN "
            sSQL = sSQL & "RV_POTipoContattoDettaglio ON RV_POTipoContatto.IDRV_POTipoContatto = RV_POTipoContattoDettaglio.IDRV_POTipoContatto ON "
            sSQL = sSQL & "RV_POContactDettaglio.IDRV_POContactDettaglio = RV_POTipoContattoDettaglio.IDRV_POContactDettaglio "
            sSQL_WHERE = ""
            If optTipoCliente > 0 Then
                sSQL_WHERE = "WHERE RV_POContact.IDRVTipoCliente=" & optTipoCliente
            End If
            If optTipoContatto > 0 Then
                If sSQL_WHERE = "" Then
                    sSQL_WHERE = "WHERE RV_POTipoContattoDettaglio.IDRV_POTipoContatto=" & optTipoContatto
                Else
                    sSQL_WHERE = sSQL_WHERE & " AND RV_POTipoContattoDettaglio.IDRV_POTipoContatto=" & optTipoContatto
                End If
            End If
            If optTipoContratto > 0 Then
                If sSQL_WHERE = "" Then
                    sSQL_WHERE = "WHERE RV_POContratto.IDTipoContratto=" & optTipoContratto
                Else
                    sSQL_WHERE = sSQL_WHERE & " AND RV_POContratto.IDTipoContratto=" & optTipoContratto
                End If
            End If

        Case 0
            sSQL = "SELECT RV_POContactDettaglio.IDRV_POContactDettaglio, RV_POContactDettaglio.IDRV_POContact, RV_POContactDettaglio.Email, "
            sSQL = sSQL & "RV_POContactDettaglio.IDTitolo, RV_POContactDettaglio.Nominativo, RV_POContactDettaglio.Mansione, dbo.RV_POContactDettaglio.Fax,"
            sSQL = sSQL & "RV_POContactDettaglio.Telefono1, RV_POContactDettaglio.Cellulare, RV_POContactDettaglio.Telefono2,"
            sSQL = sSQL & "RV_POTipoContattoDettaglio.IDRV_POTipoContatto, dbo.SitoPerAnagrafica.SitoPerAnagrafica, dbo.RV_POContact.IDRVTipoCliente,"
            sSQL = sSQL & "Comune.Comune, Provincia.Provincia, Comune_1.Comune AS ComuneFiliale, Provincia_1.Provincia AS ProvinciaFiliale,"
            sSQL = sSQL & "Anagrafica.Indirizzo, Anagrafica.Cap, SitoPerAnagrafica.Indirizzo AS IndirizzoFiliale, SitoPerAnagrafica.Cap AS CapFiliale,"
            sSQL = sSQL & "RV_POContact.IDAnagrafica , RV_POContact.IDSitoPerAnagrafica, RV_POContratto.IDTipoContratto, Anagrafica.Anagrafica "
            sSQL = sSQL & "FROM RV_POContact LEFT OUTER JOIN "
            sSQL = sSQL & "RV_POContratto ON RV_POContact.IDSitoPerAnagrafica = RV_POContratto.IDSitoPerAnagrafica AND "
            sSQL = sSQL & "RV_POContact.IDAnagrafica = RV_POContratto.IDAnagrafica LEFT OUTER JOIN "
            sSQL = sSQL & "SitoPerAnagrafica LEFT OUTER JOIN "
            sSQL = sSQL & "Provincia Provincia_1 RIGHT OUTER JOIN "
            sSQL = sSQL & "Comune Comune_1 ON Provincia_1.IDProvincia = Comune_1.IDProvincia ON SitoPerAnagrafica.IDComune = Comune_1.IDComune ON "
            sSQL = sSQL & "RV_POContact.IDSitoPerAnagrafica = SitoPerAnagrafica.IDSitoPerAnagrafica LEFT OUTER JOIN "
            sSQL = sSQL & "Provincia RIGHT OUTER JOIN "
            sSQL = sSQL & "Comune ON Provincia.IDProvincia = Comune.IDProvincia RIGHT OUTER JOIN "
            sSQL = sSQL & "Anagrafica ON Comune.IDComune = Anagrafica.IDComune ON "
            sSQL = sSQL & "RV_POContact.IDAnagrafica = Anagrafica.IDAnagrafica RIGHT OUTER JOIN "
            sSQL = sSQL & "RV_POContactDettaglio ON RV_POContact.IDRV_POContact = RV_POContactDettaglio.IDRV_POContact LEFT OUTER JOIN "
            sSQL = sSQL & "RV_POTipoContatto RIGHT OUTER JOIN "
            sSQL = sSQL & "RV_POTipoContattoDettaglio ON RV_POTipoContatto.IDRV_POTipoContatto = RV_POTipoContattoDettaglio.IDRV_POTipoContatto ON "
            sSQL = sSQL & "RV_POContactDettaglio.IDRV_POContactDettaglio = RV_POTipoContattoDettaglio.IDRV_POContactDettaglio "
            sSQL_WHERE = ""
            If optTipoCliente > 0 Then
                sSQL_WHERE = "WHERE RV_POContact.IDRVTipoCliente=" & optTipoCliente
            End If
            If optTipoContatto > 0 Then
                If sSQL_WHERE = "" Then
                    sSQL_WHERE = "WHERE RV_POTipoContattoDettaglio.IDRV_POTipoContatto=" & optTipoContatto
                Else
                    sSQL_WHERE = sSQL_WHERE & " AND RV_POTipoContattoDettaglio.IDRV_POTipoContatto=" & optTipoContatto
                End If
            End If
            If optTipoContratto > 0 Then
                If sSQL_WHERE = "" Then
                    sSQL_WHERE = "WHERE RV_POContratto.IDTipoContratto=" & optTipoContratto
                Else
                    sSQL_WHERE = sSQL_WHERE & " AND RV_POContratto.IDTipoContratto=" & optTipoContratto
                End If
            End If
        Case 2
            
    End Select
    
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL & sSQL_WHERE, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    'Set rs = CnDMT.OpenResultset(sSQL & sSQL_WHERE)
    NumeroRecord = 0
    frmParametri.ProgressBar1.Value = 0
    frmParametri.ProgressBar1.Max = 100
    
    If rs.EOF Then
        UnitaProgresso = 0
    Else
        If rs.RecordCount = 0 Then
            UnitaProgresso = 0
        Else
            UnitaProgresso = FormatNumber((frmParametri.ProgressBar1.Max / rs.RecordCount), 2)
            TotaleRecord = rs.RecordCount
        End If
    End If
    If UnitaProgresso = 0 Then
        frmParametri.lblInfo.Caption = "NESSUNA INFORMAZIONE"
        frmParametri.ProgressBar1.Value = frmParametri.ProgressBar1.Max
        Exit Sub
    End If
    
    
    
        While Not rs.EOF
            sSQL = "INSERT INTO RV_POTMPContact ("
            sSQL = sSQL & "IDRV_POTMPContact, IDAnagrafica, Anagrafica, Indirizzo, Comune, Cap, Provincia, "
            sSQL = sSQL & "IDSitoPerAnagrafica, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Nominativo, Fax, Email) "
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & fnGetNewKey("RV_POTMPContact", "IDRV_POTMPContact") & ", "
            sSQL = sSQL & rs!IDAnagrafica & ", "
            sSQL = sSQL & fnNormString(rs!Anagrafica) & ", "
            sSQL = sSQL & fnNormString(rs!Indirizzo) & ", "
            sSQL = sSQL & fnNormString(rs!Comune) & ", "
            sSQL = sSQL & fnNormString(rs!Cap) & ", "
            sSQL = sSQL & fnNormString(rs!Provincia) & ", "
            sSQL = sSQL & rs!IDSitoPerAnagrafica & ", "
            sSQL = sSQL & fnNormString(rs!SitoPerAnagrafica) & ", "
            sSQL = sSQL & fnNormString(rs!IndirizzoFiliale) & ", "
            sSQL = sSQL & fnNormString(rs!ComuneFiliale) & ", "
            sSQL = sSQL & fnNormString(rs!CapFiliale) & ", "
            sSQL = sSQL & fnNormString(rs!ProvinciaFiliale) & ", "
            sSQL = sSQL & fnNormString(rs!Nominativo) & ", "
            If optOperazione = 0 Then
                sSQL = sSQL & fnNormString(rs!fax) & ", "
            Else
                sSQL = sSQL & fnNormString("") & ", "
            End If
            If optOperazione = 1 Then
                sSQL = sSQL & fnNormString(rs!email) & ")"
            Else
                sSQL = sSQL & fnNormString("") & ")"
            End If
            
            
            
            CnDMT.Execute sSQL
            
        If (frmParametri.ProgressBar1.Value + UnitaProgresso) >= frmParametri.ProgressBar1.Max Then
            frmParametri.ProgressBar1.Value = frmParametri.ProgressBar1.Max
        Else
            frmParametri.ProgressBar1.Value = frmParametri.ProgressBar1.Value + UnitaProgresso
        End If
            
        NumeroRecord = NumeroRecord + 1
        frmParametri.lblInfo.Caption = "Elaborazione di record " & NumeroRecord & " su " & TotaleRecord
        DoEvents
        rs.MoveNext
        Wend
  
    rs.Close
    Set rs = Nothing
End Sub
