Attribute VB_Name = "ModMain"
Option Explicit
Public ID_ContactDettaglio As Long

'**+
'Nome                   : Main
'Parametri              : Nessuno
'Valori di ritorno      : Nessuno
'Funzionalità           : In questa Sub vengono eseguite
'                       : le operazioni di Startup.
'                       : Questa procedura deve essere
'                       : sempre presente nel prog. gestore.
'**/
Sub Main()
    On Error GoTo ErrorHandler
    'L'Applicazione
    Set TheApp = New Application

    'Carica il form della applicazione senza mostrarlo
    Load frmMain
    
    'Abilita il form della applicazione alla ricezione
    'degli eventi da parte della applicazione Diamante
    Set frmMain.Application = TheApp

    'Il nome della applicazione
    TheApp.Name = App.EXEName
    
    
    'Viene inizializzata la componente DmtRegistry2 che si occupa della gestione
    'degli accessi al registry.
    DmtRegistry2.EXEName = App.EXEName
    
    
    Set gResource = New Resource
    REGISTRY_KEY = Trim(gResource.GetMessage(LBL_REGISTRY_KEY))
    
    'L'icona di Diamante
    frmMain.Icon = gResource.GetIcon(IDI_DIAMANTE16)
        
    'Esegue l'applicazione
    TheApp.Run frmMain.hwnd
    
    
    'La vista alla partenza deve essere quella del Form
    
    frmMain.BrwMain.Visible = False
    
    'alla fine del caricamento dell'applicazione
    'visualizza il layout del form.
    frmMain.Show
    
    
    '..............................................................................
    'NOTA: In Form_Activate() viene determinata la modalità iniziale con cui si   .
    'avvia il programma:                                                          .
    'Se il filtro predefinito restituisce dei record si va in modalità variazione .
    'sul primo record altrimenti si va in modalità inserimento.                   .
    '..............................................................................
    
    Exit Sub
ErrorHandler:
    Unload frmMain
    If Err.Number = 1 + vbObjectError Then
        'Questo programma può essere eseguito solo all'interno dell'applicativo Diamante.
        'Prima di TheApp.Run si ha TheApp.FunctionName = "" allora nella Caption del messaggio si avrà TheApp.Name.
        sbMsgInfo gResource.GetMessage(MESS_RUNOUTOFDIAMANTE), IIf(TheApp.FunctionName <> "", TheApp.FunctionName, TheApp.Name)
    Else
        Err.Raise Err.Number
    End If
End Sub
Public Function fnGetEsercizio(dData As Date) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "Select IDEsercizio FROM Esercizio WHERE "
    sSQL = sSQL & "((IDAzienda = " & TheApp.IDFirm & ")"
    sSQL = sSQL & " AND (DataInizio <= #" & dData & "#)"
    sSQL = sSQL & " AND (DataFine >= #" & dData & "#))"
   

    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetEsercizio = rsEse!IDEsercizio
    Else
        fnGetEsercizio = 0
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Public Function fnGetNewKey(Tabella As String, CampoKey As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
        
        sSQL = "SELECT " & CampoKey & " FROM " & Tabella & " ORDER BY " & CampoKey & " DESC"
        
        Set rs = Cn.OpenResultset(fnAnsi2Jet(sSQL))
    
        If rs.EOF = True Then
        
            fnGetNewKey = 1
    
        Else
            
            fnGetNewKey = fnNotNullN(rs.adoColumns(CampoKey)) + 1
    
        End If

        rs.CloseResultset
        Set rs = Nothing
    

    
End Function

