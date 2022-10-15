VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmConfiguraServizi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CONFIGURAZIONE SERVIZI"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRiporta 
      Caption         =   "RIPORTA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   3
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdDeselezionaTutto 
      Caption         =   "Deseleziona tutto"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdSelezionaTutto 
      Caption         =   "Seleziona tutto"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   1815
   End
   Begin DmtGridCtl.DmtGrid GrigliaServizi 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   8070
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableMove      =   0   'False
      ColumnsHeaderHeight=   20
   End
End
Attribute VB_Name = "frmConfiguraServizi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsNew As ADODB.Recordset

Private Sub cmdDeselezionaTutto_Click()
On Error GoTo ERR_cmdDeselezionaTutto_Click
If Not (rsNew.BOF And rsNew.EOF) Then

    rsNew.MoveFirst
    While Not rsNew.EOF
        'If rsNew!Riportato = 0 Then
            sbSelectSelectedRow Not CBool(rsNew.Fields("DaRiportare").Value), 2
        'End If
    rsNew.MoveNext
    Wend
End If
Exit Sub
ERR_cmdDeselezionaTutto_Click:
    MsgBox Err.Description, vbCritical, "cmdDeselezionaTutto_Click"
End Sub

Private Sub cmdRiporta_Click()
    ELABORA_SERVIZI_PER_CONTRATTO
    Unload Me
End Sub

Private Sub cmdSelezionaTutto_Click()
On Error GoTo ERR_cmdSelezionaTutto_Click
If Not (rsNew.BOF And rsNew.EOF) Then

    rsNew.MoveFirst
    While Not rsNew.EOF
        If rsNew!Riportato = 0 Then
            sbSelectSelectedRow Not CBool(rsNew.Fields("DaRiportare").Value), 2
        End If
    rsNew.MoveNext
    Wend
End If
Exit Sub
ERR_cmdSelezionaTutto_Click:
    MsgBox Err.Description, vbCritical, "cmdSelezionaTutto_Click"
End Sub

Private Sub Form_Load()
    CREA_TABELLA_TEMPORANEA
    INSERIMENTO_DATI
    SETTAGGIO_GRIGLIA
    RIPORTA_SERVIZI = False
End Sub
Private Sub CREA_TABELLA_TEMPORANEA()

Set rsNew = New ADODB.Recordset

rsNew.CursorLocation = adUseClient

rsNew.Fields.Append "DaRiportare", adSmallInt, , adFldIsNullable
rsNew.Fields.Append "Riportato", adSmallInt, , adFldIsNullable
rsNew.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsNew.Fields.Append "CodiceArticolo", adVarChar, 250, adFldIsNullable
rsNew.Fields.Append "Articolo", adVarChar, 250, adFldIsNullable
rsNew.Fields.Append "IDRV_POCriterioRicorrenza", adInteger, , adFldIsNullable
rsNew.Fields.Append "CriterioRicorrenza", adVarChar, 250, adFldIsNullable
rsNew.Fields.Append "OgniNumeroGiorni", adInteger, , adFldIsNullable
rsNew.Fields.Append "OgniNumeroMesi", adInteger, , adFldIsNullable
rsNew.Fields.Append "OgniNumeroSettimane", adInteger, , adFldIsNullable
rsNew.Fields.Append "IDRV_POTipoDataInizioRicorrenza", adInteger, , adFldIsNullable
rsNew.Fields.Append "TipoDataInizioRicorrenza", adVarChar, 250, adFldIsNullable
rsNew.Fields.Append "GiornoInizioRicorrenza", adInteger, , adFldIsNullable
rsNew.Fields.Append "MeseInizioRicorrenza", adInteger, , adFldIsNullable
rsNew.Fields.Append "IDRV_POTipoDataFineRicorrenza", adInteger, , adFldIsNullable
rsNew.Fields.Append "TipoDataFineRicorrenza", adVarChar, 250, adFldIsNullable
rsNew.Fields.Append "GiornoFineRicorrenza", adInteger, , adFldIsNullable
rsNew.Fields.Append "MeseFineRicorrenza", adInteger, , adFldIsNullable
rsNew.Fields.Append "NumeroRicorrenza", adInteger, , adFldIsNullable
rsNew.Fields.Append "IDRV_POTipoAnnoInizioRicorrenza", adInteger, , adFldIsNullable
rsNew.Fields.Append "IDRV_POTipoAnnoFineRicorrenza", adInteger, , adFldIsNullable
End Sub
Private Sub INSERIMENTO_DATI()
On Error GoTo ERR_INSERIMENTO_DATI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POTipoContrattoServizi.IDRV_POTipoContrattoServizi, RV_POTipoContrattoServizi.IDRV_POTipoContratto, RV_POTipoContrattoServizi.IDArticolo, "
sSQL = sSQL & "RV_POTipoContrattoServizi.IDRV_POCriterioRicorrenza, RV_POTipoContrattoServizi.OgniNumeroGiorni, RV_POTipoContrattoServizi.OgniNumeroMesi,"
sSQL = sSQL & "RV_POTipoContrattoServizi.OgniNumeroSettimane, RV_POTipoContrattoServizi.IDRV_POTipoDataInizioRicorrenza,"
sSQL = sSQL & "RV_POTipoContrattoServizi.GiornoInizioRicorrenza, RV_POTipoContrattoServizi.MeseInizioRicorrenza,"
sSQL = sSQL & "RV_POTipoContrattoServizi.IDRV_POTipoDataFineRicorrenza, RV_POTipoContrattoServizi.GiornoFineRicorrenza,"
sSQL = sSQL & "RV_POTipoContrattoServizi.MeseFineRicorrenza, RV_POTipoContrattoServizi.NumeroRicorrenze, Articolo.CodiceArticolo, Articolo.Articolo,"
sSQL = sSQL & "RV_POTipoDataFineRicorrenza.TipoDataFineRicorrenza, RV_POTipoDataInizioRicorrenza.TipoDataInizioRicorrenza, "
sSQL = sSQL & "RV_POTipoContrattoServizi.IDRV_POTipoAnnoInizioRicorrenza, RV_POTipoContrattoServizi.IDRV_POTipoAnnoFineRicorrenza, "
sSQL = sSQL & "RV_POCriterioRicorrenza.CriterioRicorrenza "
sSQL = sSQL & "FROM RV_POTipoContrattoServizi LEFT OUTER JOIN "
sSQL = sSQL & "Articolo ON RV_POTipoContrattoServizi.IDArticolo = Articolo.IDArticolo LEFT OUTER JOIN "
sSQL = sSQL & "RV_POTipoDataFineRicorrenza ON "
sSQL = sSQL & "RV_POTipoContrattoServizi.IDRV_POTipoDataFineRicorrenza = RV_POTipoDataFineRicorrenza.IDRV_POTipoDataFineRicorrenza LEFT OUTER JOIN "
sSQL = sSQL & "RV_POTipoDataInizioRicorrenza ON "
sSQL = sSQL & "RV_POTipoContrattoServizi.IDRV_POTipoDataInizioRicorrenza = RV_POTipoDataInizioRicorrenza.IDRV_POTipoDataInizioRicorrenza LEFT OUTER JOIN "
sSQL = sSQL & "RV_POCriterioRicorrenza ON RV_POTipoContrattoServizi.IDRV_POCriterioRicorrenza = RV_POCriterioRicorrenza.IDRV_POCriterioRicorrenza "
sSQL = sSQL & "WHERE IDRV_POTipoContratto=" & frmMain.cboTipoContratto.CurrentID
 
Set rs = Cn.OpenResultset(sSQL)
rsNew.Open , , adOpenKeyset, adLockBatchOptimistic
While Not rs.EOF
    rsNew.AddNew
        rsNew!DaRiportare = GET_ARTICOLO_IN_CONTRATTO(fnNotNullN(rs!IDArticolo))
        rsNew!Riportato = IIf((rsNew!DaRiportare = 1), 0, 1)
        rsNew!IDArticolo = rs!IDArticolo
        rsNew!CodiceArticolo = rs!CodiceArticolo
        rsNew!Articolo = rs!Articolo
        rsNew!IDRV_POCriterioRicorrenza = rs("IDRV_POCriterioRicorrenza").Value
        rsNew!CriterioRicorrenza = rs!CriterioRicorrenza
        rsNew!OgniNumeroGiorni = rs!OgniNumeroGiorni
        rsNew!OgniNumeroMesi = rs!OgniNumeroMesi
        rsNew!OgniNumeroSettimane = rs!OgniNumeroSettimane
        rsNew!IDRV_POTipoDataInizioRicorrenza = rs!IDRV_POTipoDataInizioRicorrenza
        rsNew!GiornoInizioRicorrenza = rs!GiornoInizioRicorrenza
        rsNew!MeseInizioRicorrenza = rs!MeseInizioRicorrenza
        rsNew!IDRV_POTipoDataFineRicorrenza = rs!IDRV_POTipoDataFineRicorrenza
        rsNew!GiornoFineRicorrenza = rs!GiornoFineRicorrenza
        rsNew!MeseFineRicorrenza = rs!MeseFineRicorrenza
        rsNew!NumeroRicorrenza = rs!NumeroRicorrenze
        rsNew!IDRV_POTipoAnnoInizioRicorrenza = rs!IDRV_POTipoAnnoInizioRicorrenza
        rsNew!IDRV_POTipoAnnoFineRicorrenza = rs!IDRV_POTipoAnnoFineRicorrenza
    rsNew.Update
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_INSERIMENTO_DATI:
    MsgBox Err.Description, vbCritical, "INSERIMENTO_DATI"
End Sub

Private Sub SETTAGGIO_GRIGLIA()
On Error GoTo ERR_SettaggioGrigliaRate
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    

        
    With Me.GrigliaServizi
        .EnableMove = True
        .UpdatePosition = False
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear

        .ColumnsHeader.Clear
            Set cl = .ColumnsHeader.Add("DaRiportare", "Riporta", dgBoolean, True, 1000, dgAligncenter)
                cl.Editable = True
            
            .ColumnsHeader.Add "Riportato", "Esistente", dgBoolean, True, 1000, dgAligncenter
            .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgVarChar, True, 1500, dgAlignleft
            .ColumnsHeader.Add "Articolo", "Articolo", dgVarChar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POCriterioRicorrenza", "IDRV_POCriterioRicorrenza", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "OgniNumeroGiorni", "Ogni n° giorni", dgInteger, True, 1500, dgAlignRight
            .ColumnsHeader.Add "OgniNumeroMesi", "Ogni n° mesi", dgInteger, True, 1500, dgAlignRight
            .ColumnsHeader.Add "OgniNumeroSettimane", "Ogni n° settimane", dgInteger, True, 1500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POTipoDataInizioRicorrenza", "IDRV_POTipoDataInizioRicorrenza", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "TipoDataInizioRicorrenza", "Tipo data inizio Ric.", dgVarChar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "GiornoInizioRicorrenza", "Giorno inizio ric.", dgInteger, True, 1500, dgAlignRight
            .ColumnsHeader.Add "MeseInizioRicorrenza", "Mese inizio ric.", dgInteger, True, 1500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POTipoDataFineRicorrenza", "IDRV_POTipoDataFineRicorrenza", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "TipoDataFineRicorrenza", "Tipo data fine Ric.", dgVarChar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "GiornoFineRicorrenza", "Giorno fine ric.", dgInteger, True, 1500, dgAlignRight
            .ColumnsHeader.Add "MesefineRicorrenza", "Mese fine ric.", dgInteger, True, 1500, dgAlignRight
            .ColumnsHeader.Add "NumeroRicorrenza", "NumeroRicorrenza", dgInteger, True, 1500, dgAlignRight
                
        Set .Recordset = rsNew
        .Refresh
    End With
    
    Cn.CursorLocation = OLDCursor

Exit Sub
ERR_SettaggioGrigliaRate:
    MsgBox Err.Description, vbCritical, "SettaggioGrigliaRate"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (rsNew Is Nothing) Then
        rsNew.Close
        Set rsNew = Nothing
    End If
End Sub

Private Sub GrigliaServizi_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GrigliaServizi.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsNew.Fields("DaRiportare").Value), 2
            
        End If
    End If
    

End Sub

Private Sub GrigliaServizi_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If GrigliaServizi.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < GrigliaServizi.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If GrigliaServizi.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsNew.Fields("DaRiportare").Value), 2
            End If
        End If
    End If

End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
Dim TestoMessaggio As String

TestoMessaggio = "ATTENZIONE!!!!" & vbCrLf
TestoMessaggio = TestoMessaggio & "Il servizio selezionato è già presente nel contratto" & vbCrLf
TestoMessaggio = TestoMessaggio & "Vuoi continuare con questo comando?"

        If Not rsNew.EOF And Not rsNew.BOF Then
            If rsNew.Fields("DaRiportare").Value = 0 Then
                If rsNew.Fields("Riportato").Value = 1 Then
                    If MsgBox(TestoMessaggio, vbQuestion + vbYesNo, "Importazione servizi") = vbNo Then Exit Sub
                End If
            End If
            
            rsNew.Fields("DaRiportare").Value = Abs(CLng(Selected))
            
            Me.GrigliaServizi.Refresh
            
        End If
End Sub
Private Function GET_ARTICOLO_IN_CONTRATTO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT IDArticolo FROM RV_POContrattoServizi "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDRV_POContratto=" & Link_Contratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ARTICOLO_IN_CONTRATTO = 1
Else
    GET_ARTICOLO_IN_CONTRATTO = 0
End If
    
rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ELABORA_SERVIZI_PER_CONTRATTO()
On Error GoTo ERR_ELABORA_SERVIZI_PER_CONTRATTO
Dim sSQL As String
Dim rsServ As ADODB.Recordset
Dim OLD_CAPTION_FORM
Dim X As Long
Dim ErroreCoda As String
Dim OLD_Cursor As Long

    SCRIVI_CODA Link_Contratto
    APERTURA_FORM_CODA = False
    OLD_CAPTION_FORM = Me.Caption
    
    ''''''''''''''''''''''''''''''CONTROLLA LA CODA DEI SALVATAGGI'''''''''''''''''''''''''''''
    X = 0
    ErroreCoda = False
    Do
        X = GET_NUMERO_DOCUMENTO(False)
        If X = -1 Then
            X = 1
            ErroreCoda = True
        End If
    Loop Until X = 1
    
    If ErroreCoda = True Then
        X = -1
    End If
    
    If X = -1 Then
        Me.Enabled = True
        Me.SetFocus
        Me.Caption = OLD_CAPTION_FORM
        Screen.MousePointer = 0
        ''''''''ELIMINAZIONE RIFERIMENTO CODA'''''''''''''''''''''''''''''''
        sSQL = "DELETE FROM RV_POTMP "
        sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
        sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto(App.EXEName)
        Cn.Execute sSQL
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Me.Enabled = True
    Me.SetFocus
    Me.Caption = OLD_CAPTION_FORM
    
    OLD_Cursor = Cn.CursorLocation
    Cn.CursorLocation = adUseClient
    
    'frmAttesa.Show
    Me.Enabled = False
    
    DoEvents
    
    Me.Caption = "SALVATAGGIO IN CORSO..................."
    DoEvents
    
    'frmAttesa.lblInfo = Me.Caption
    DoEvents
    
    Screen.MousePointer = 11
    
    rsNew.UpdateBatch


    rsNew.Filter = "DaRiportare=1"
    
    If Not rsNew.EOF Then
        rsNew.MoveFirst
        
        Cn.BeginTrans
        
        Set rsServ = New ADODB.Recordset
        rsServ.Open "SELECT * FROM RV_POContrattoServizi WHERE IDRV_POContrattoServizi=0", Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
        
        While Not rsNew.EOF
            rsServ.AddNew
                rsServ!IDRV_POContrattoServizi = fnGetNewKey("RV_POContrattoServizi", "IDRV_POContrattoServizi")
                rsServ!IDRV_POContratto = Link_Contratto
                rsServ!IDRV_POStoriaContratto = Link_StoriaContratto
                rsServ("IDArticolo").Value = rsNew("IDArticolo").Value
                rsServ("IDRV_POCriterioRicorrenza").Value = rsNew("IDRV_POCriterioRicorrenza").Value
                rsServ("OgniNumeroGiorni").Value = rsNew("OgniNumeroGiorni").Value
                rsServ("OgniNumeroMesi").Value = rsNew("OgniNumeroMesi").Value
                rsServ("OgniNumeroSettimane").Value = rsNew("OgniNumeroSettimane").Value
                rsServ("IDRV_POTipoDataInizioRicorrenza").Value = rsNew("IDRV_POTipoDataInizioRicorrenza").Value
                rsServ("GiornoInizioRicorrenza").Value = rsNew("GiornoInizioRicorrenza").Value
                rsServ("MeseInizioRicorrenza").Value = rsNew("MeseInizioRicorrenza").Value
                rsServ("IDRV_POTipoDataFineRicorrenza").Value = rsNew("IDRV_POTipoDataFineRicorrenza").Value
                rsServ("GiornoFineRicorrenza").Value = rsNew("GiornoFineRicorrenza").Value
                rsServ("MeseFineRicorrenza").Value = rsNew("MeseFineRicorrenza").Value
                rsServ("NumeroRicorrenze").Value = rsNew("NumeroRicorrenza").Value
                rsServ("IDRV_POTipoAnnoInizioRicorrenza").Value = rsNew("IDRV_POTipoAnnoInizioRicorrenza").Value
                rsServ("IDRV_POTipoAnnoFineRicorrenza").Value = rsNew("IDRV_POTipoAnnoFineRicorrenza").Value
                
                If fnNotNullN(rsServ("IDRV_POTipoAnnoInizioRicorrenza").Value) = 0 Then
                    rsServ("IDRV_POTipoAnnoInizioRicorrenza").Value = GET_TIPO_ANNO_INIZIO_SERVIZIO(fnNotNullN(rsServ("IDArticolo").Value))
                End If
                If fnNotNullN(rsServ("IDRV_POTipoAnnoFineRicorrenza").Value) = 0 Then
                    rsServ("IDRV_POTipoAnnoFineRicorrenza").Value = GET_TIPO_ANNO_FINE_SERVIZIO(fnNotNullN(rsServ("IDArticolo").Value))
                End If
                
            rsServ.Update
        rsNew.MoveNext
        Wend
    End If
    Cn.CommitTrans
    rsServ.Close
    Set rsServ = Nothing



'Unload frmAttesa
Me.Enabled = True
Me.SetFocus
Me.Caption = OLD_CAPTION_FORM
Cn.CursorLocation = OLD_Cursor

''''''''''''''''''''''''''''''''''''ELIMINAZIONE CODA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POTMP "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
'sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto(App.EXEName)
Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Screen.MousePointer = 0
Exit Sub
ERR_ELABORA_SERVIZI_PER_CONTRATTO:
    On Error Resume Next
    'Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    Cn.RollbackTrans
    MsgBox Err.Description, vbCritical, "Elaborazione dati"

    'Cn.RollbackTrans
    ''''''''''''''''''''ELIMINAZIONE RIGA DI CODA'''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
    'sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Cn.CursorLocation = OLD_Cursor
    
    Me.Caption = OLD_CAPTION_FORM
    
    Screen.MousePointer = 0

    
End Sub
Private Sub SCRIVI_CODA(IDOggetto As Long)
Dim rs As ADODB.Recordset
Dim sSQL As String

'''''''''''''''''ELIMINAZIONE DATI UTENTE PER IL TIPO OGGETTO'''''''''''''''''''

sSQL = "DELETE FROM RV_POTMP "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
'sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID

Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set rs = New ADODB.Recordset

rs.Open "RV_POTMP", Cn.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    'rs!IDSessione = fnGetNewKey("RV_POTMP", "IDSessione")
    rs!IDUtente = TheApp.IDUser
    rs!IDTipoOggetto = fnGetTipoOggetto(App.EXEName)
    rs!IDOggetto = IDOggetto
    rs!Utente = TheApp.User
rs.Update

rs.Close
Set rs = Nothing

End Sub
Private Function GET_NUMERO_DOCUMENTO(NuovoDocumento As Boolean) As Long
On Error GoTo ERR_GET_NUMERO_DOCUMENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim X_FRM As Form
Dim OLD_Cursor As Long

GET_NUMERO_DOCUMENTO = 0

sSQL = "SELECT * FROM RV_POTMP "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto(App.EXEName)
sSQL = sSQL & " ORDER BY IDSessione, IDUtente"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!IDUtente) = TheApp.IDUser Then
        Me.Caption = "SALVATAGGIO IN CORSO.........."
        
        DoEvents
       
        'If APERTURA_FORM_CODA = True Then
        '    Unload frmCoda
        '    APERTURA_FORM_CODA = False
        'End If
        
        GET_NUMERO_DOCUMENTO = 1
        
        rs.CloseResultset
        Set rs = Nothing
    Else
        rs.CloseResultset
        Set rs = Nothing
    
        'If APERTURA_FORM_CODA = False Then
        '    APERTURA_FORM_CODA = True
        '    Me.Enabled = False
        '    frmCoda.Show
        'End If
        
        Me.Caption = "ATTENDERE......."
        DoEvents
        'GET_NUMERO_DOCUMENTO NuovoDocumento
        
    End If
End If
Exit Function

ERR_GET_NUMERO_DOCUMENTO:
    MsgBox Err.Description, vbCritical, "Errore coda"
    GET_NUMERO_DOCUMENTO = -1
    Unload frmCoda
End Function


Private Function fnGetTipoOggetto(NomeGestore) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(NomeGestore)
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Function GET_TIPO_ANNO_INIZIO_SERVIZIO(IDArticolo As Long) As Long
On Error GoTo ERR_GET_TIPO_ANNO_SERVIZIO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_TIPO_ANNO_INIZIO_SERVIZIO = 0

sSQL = "SELECT IDRV_POTipoAnnoInizioRicorrenza "
sSQL = sSQL & " FROM RV_POConfigurazioneServizio "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_ANNO_INIZIO_SERVIZIO = 0
Else
    GET_TIPO_ANNO_INIZIO_SERVIZIO = fnNotNullN(rs!IDRV_POTipoAnnoInizioRicorrenza)
End If

rs.CloseResultset
Set rs = Nothing


Exit Function
ERR_GET_TIPO_ANNO_SERVIZIO:


End Function
Private Function GET_TIPO_ANNO_FINE_SERVIZIO(IDArticolo As Long) As Long
On Error GoTo ERR_GET_TIPO_ANNO_SERVIZIO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_TIPO_ANNO_FINE_SERVIZIO = 0

sSQL = "SELECT IDRV_POTipoAnnoFineRicorrenza "
sSQL = sSQL & " FROM RV_POConfigurazioneServizio "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_ANNO_FINE_SERVIZIO = 0
Else
    GET_TIPO_ANNO_FINE_SERVIZIO = fnNotNullN(rs!IDRV_POTipoAnnoFineRicorrenza)
End If

rs.CloseResultset
Set rs = Nothing


Exit Function
ERR_GET_TIPO_ANNO_SERVIZIO:
    
End Function
