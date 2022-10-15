VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmContatori 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Attivazione contatori"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18540
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
   ScaleHeight     =   7635
   ScaleWidth      =   18540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSoloProdAss 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Picture         =   "frmContatori.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Solo prodotti associati"
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton cmdRiporta 
      Caption         =   "CONFIGURA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16320
      TabIndex        =   3
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton cmdDeSelTutto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Picture         =   "frmContatori.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deseleziona tutto"
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton cmdSelTutto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Picture         =   "frmContatori.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Seleziona tutto"
      Top             =   7080
      Width           =   855
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18495
      _ExtentX        =   32623
      _ExtentY        =   12303
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   1920
      X2              =   1920
      Y1              =   7080
      Y2              =   7560
   End
End
Attribute VB_Name = "frmContatori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private soloContConf As Boolean

Private Sub CREATE_RECORDSET()
On Error GoTo ERR_CREATE_RECORDSET
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroRecord As Long
Dim UnitaProgresso As Double
Dim NumeroElaborazione As Long
Dim Avvia As Boolean

If Not (rsGriglia Is Nothing) Then
    If rsGriglia.State > 0 Then
        rsGriglia.Close
    End If
    Set rsGriglia = Nothing
End If

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

rsGriglia.Fields.Append "IDRV_POContatoreProdotto", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "Codice", adChar, 250, adFldIsNullable
rsGriglia.Fields.Append "Descrizione", adChar, 250, adFldIsNullable
rsGriglia.Fields.Append "IDRV_POUMContatore", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "UMContatore", adChar, 250, adFldIsNullable
rsGriglia.Fields.Append "NelProdotto", adBoolean, , adFldIsNullable
rsGriglia.Fields.Append "IDRV_POContrattoProdottiContatori", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "QuantitaMax", adDouble, , adFldIsNullable
rsGriglia.Fields.Append "IDRV_POUnitaDiMisuraPeriodo", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "UMPeriodo", adChar, 250, adFldIsNullable
rsGriglia.Fields.Append "QuantitaPeriodo", adDouble, , adFldIsNullable
rsGriglia.Fields.Append "QuantitaInizio", adDouble, , adFldIsNullable
rsGriglia.Fields.Append "ImportoUnitario", adDouble, , adFldIsNullable

rsGriglia.Fields.Append "Riporta", adBoolean, , adFldIsNullable

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic

sSQL = "SELECT * FROM RV_POIEContatoreProdotto "

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
        
    rsGriglia.AddNew

        rsGriglia!IDRV_POContatoreProdotto = fnNotNullN(rs!IDRV_POContatoreProdotto)
        rsGriglia!Codice = fnNotNull(rs!Codice)
        rsGriglia!Descrizione = fnNotNull(rs!Descrizione)
        rsGriglia!IDRV_POUMContatore = fnNotNullN(rs!IDRV_POUMContatore)
        rsGriglia!UMContatore = fnNotNull(rs!UMContatore)
        
        rsGriglia!NelProdotto = GET_CONTATORE_PRODOTTO(LINK_PRODOTTO_SEL, fnNotNullN(rs!IDRV_POContatoreProdotto))


        GET_DATI_CONTATORE_CONTRATTO rsGriglia, LINK_PRODOTTO_RIGA_SEL, LINK_PRODOTTO_SEL, fnNotNullN(rs!IDRV_POContatoreProdotto)
        rsGriglia!Riporta = False
        
        If fnNotNullN(rsGriglia!IDRV_POContrattoProdottiContatori) > 0 Then
            rsGriglia!Riporta = True
        End If
    rsGriglia.Update
        
    
    DoEvents
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

GET_GRIGLIA

Exit Sub
ERR_CREATE_RECORDSET:
    MsgBox Err.Description, vbCritical, "CREATE_RECORDSET"
End Sub

Private Function GET_CONTATORE_PRODOTTO(IDProdotto As Long, IDContatore As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTATORE_PRODOTTO = False

sSQL = "SELECT IDRV_POProdottoContatori FROM RV_POProdottoContatori "
sSQL = sSQL & "WHERE IDRV_POProdotto=" & IDProdotto
sSQL = sSQL & " AND IDRV_POContatoreProdotto=" & IDContatore

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTATORE_PRODOTTO = False
Else
    GET_CONTATORE_PRODOTTO = True
End If

rs.CloseResultset
Set rs = Nothing


End Function

Private Sub GET_DATI_CONTATORE_CONTRATTO(rstmp As ADODB.Recordset, IDRV_POContrattoProdotti As Long, IDProdotto As Long, IDContatore As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT * FROM RV_POContrattoProdottiContatori "
sSQL = sSQL & "WHERE IDRV_POProdotto=" & IDProdotto
sSQL = sSQL & " AND IDRV_POContatoreProdotto=" & IDContatore
sSQL = sSQL & " AND IDRV_POContrattoProdotti=" & IDRV_POContrattoProdotti
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    rstmp!IDRV_POContrattoProdottiContatori = 0
    rstmp!QuantitaMax = Null
    rstmp!IDRV_POUnitaDiMisuraPeriodo = Null
    rstmp!UMPeriodo = ""
    rstmp!QuantitaPeriodo = Null
    rstmp!QuantitaInizio = Null
    rstmp!ImportoUnitario = Null
    
Else
    rstmp!IDRV_POContrattoProdottiContatori = fnNotNullN(rs!IDRV_POContrattoProdottiContatori)
    rstmp!QuantitaMax = fnNotNullN(rs!QuantitaMax)
    rstmp!IDRV_POUnitaDiMisuraPeriodo = fnNotNullN(rs!IDRV_POUnitaDiMisuraPeriodo)
    rstmp!UMPeriodo = GET_DESCRIZIONE_PERIODO(fnNotNullN(rs!IDRV_POUnitaDiMisuraPeriodo))
    rstmp!QuantitaPeriodo = fnNotNullN(rs!QuantitaPeriodo)
    rstmp!QuantitaInizio = fnNotNullN(rs!QuantitaInizio)
    rstmp!ImportoUnitario = fnNotNullN(rs!ImportoUnitario)
    
End If

rs.CloseResultset
Set rs = Nothing


End Sub
Private Function GET_DESCRIZIONE_PERIODO(IDUMPeriodo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_DESCRIZIONE_PERIODO = ""

sSQL = "SELECT * FROM RV_POUnitaDiMisuraPeriodo "
sSQL = sSQL & "WHERE IDRV_POUnitaDiMisuraPeriodo=" & IDUMPeriodo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_PERIODO = ""
Else
    GET_DESCRIZIONE_PERIODO = fnNotNull(rs!Descrizione)
End If

rs.CloseResultset
Set rs = Nothing


End Function

Private Sub cmdRiporta_Click()
On Error GoTo ERR_cmdRiporta_Click
Dim sSQL As String
Dim rs As ADODB.Recordset

If ((rsGriglia.EOF) And (rsGriglia.BOF)) Then Exit Sub

rsGriglia.Filter = "Riporta=" & fnNormBoolean(1)

While Not rsGriglia.EOF

    sSQL = "SELECT * FROM RV_POContrattoProdottiContatori "
    sSQL = sSQL & "WHERE IDRV_POContrattoProdottiContatori=" & fnNotNullN(rsGriglia!IDRV_POContrattoProdottiContatori)
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rs.EOF Then
        rs.AddNew
        rs!IDRV_POContrattoProdottiContatori = fnGetNewKey("RV_POContrattoProdottiContatori", "IDRV_POContrattoProdottiContatori")
    End If
    
    
    rs!IDRV_POContrattoProdotti = LINK_PRODOTTO_RIGA_SEL
    rs!IDRV_POProdotto = LINK_PRODOTTO_SEL
    rs!IDRV_POContatoreProdotto = fnNotNullN(rsGriglia!IDRV_POContatoreProdotto)
    rs!QuantitaMax = rsGriglia!QuantitaMax
    rs!IDRV_POUnitaDiMisuraPeriodo = rsGriglia!IDRV_POUnitaDiMisuraPeriodo
    rs!QuantitaPeriodo = rsGriglia!QuantitaPeriodo
    rs!QuantitaInizio = rsGriglia!QuantitaInizio
    rs!ImportoUnitario = rsGriglia!ImportoUnitario
    
    
    
    rs.Update
    
    rs.Close
    Set rs = Nothing
rsGriglia.MoveNext
Wend


Unload Me

Exit Sub
ERR_cmdRiporta_Click:
    MsgBox Err.Description, vbCritical, "cmdRiporta_Click"
End Sub

Private Sub cmdSelTutto_Click()
On Error GoTo ERR_cmdSelTutto_Click
If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
    rsGriglia.MoveFirst
    
    While Not rsGriglia.EOF
        rsGriglia!Riporta = True
        rsGriglia.Update
    rsGriglia.MoveNext
    Wend
    
    Me.GrigliaCorpo.Refresh
End If

Exit Sub
ERR_cmdSelTutto_Click:
    MsgBox Err.Description, vbCritical, "cmdSelTutto_Click"
End Sub
Private Sub cmdDeSelTutto_Click()
On Error GoTo ERR_cmdDeSelTutto_Click
If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
    rsGriglia.MoveFirst
    
    While Not rsGriglia.EOF
        rsGriglia!Riporta = False
        rsGriglia.Update
    rsGriglia.MoveNext
    Wend
    
    Me.GrigliaCorpo.Refresh
End If
Exit Sub
ERR_cmdDeSelTutto_Click:
    MsgBox Err.Description, vbCritical, "cmdDeSelTutto_Click"
End Sub

Private Sub cmdSoloProdAss_Click()
    If soloContConf = True Then
        soloContConf = False
        cmdSoloProdAss.ToolTipText = "Solo prodotti associati"
    Else
        soloContConf = True
        cmdSoloProdAss.ToolTipText = "Tutti i prodotti"
    End If
    
    GET_GRIGLIA
    
End Sub

Private Sub Form_Load()
    soloContConf = False
    
    CREATE_RECORDSET
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    rsGriglia.Filter = vbNullString
    If soloContConf = True Then
        rsGriglia.Filter = "NelProdotto=" & fnNormBoolean(soloContConf)
    End If
    
    With Me.GrigliaCorpo
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
            Set cl = .ColumnsHeader.Add("Riporta", "Configura", dgBoolean, True, 1000, dgAligncenter)
                cl.Editable = True
            .ColumnsHeader.Add "IDRV_POContatoreProdotto", "IDRV_POContatoreProdotto", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Codice", "Codice", dgchar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "Descrizione", "Descrizione", dgchar, True, 2500, dgAlignleft
         
            .ColumnsHeader.Add "IDRV_POUMContatore", "IDRV_POUMContatore", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "UMContatore", "U.M.", dgchar, True, 1500, dgAlignleft
            
            .ColumnsHeader.Add "NelProdotto", "Ass. al prod.", dgBoolean, True, 1500, dgAligncenter
            
            .ColumnsHeader.Add "IDRV_POContrattoProdottiContatori", "IDRV_POContrattoProdottiContatori", dgInteger, False, 500, dgAlignleft
            Set cl = .ColumnsHeader.Add("QuantitaMax", "Q.tà max", dgDouble, True, 1500, dgAlignRight)
                cl.BackColor = vbYellow
                cl.Editable = True
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 1
                cl.FormatOptions.FormatNumericThousandSep = "."
                
            .ColumnsHeader.Add "IDRV_POUnitaDiMisuraPeriodo", "IDRV_POUnitaDiMisuraPeriodo", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "UMPeriodo", "Periodo", dgchar, True, 2500, dgAlignleft
            
            Set cl = .ColumnsHeader.Add("QuantitaPeriodo", "Q.tà periodo", dgDouble, True, 1500, dgAlignRight)
                cl.BackColor = vbYellow
                cl.Editable = True
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 1
                cl.FormatOptions.FormatNumericThousandSep = "."
                
            Set cl = .ColumnsHeader.Add("QuantitaInizio", "Q.tà inizio", dgDouble, True, 1500, dgAlignRight)
                cl.BackColor = vbYellow
                cl.Editable = True
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 1
                cl.FormatOptions.FormatNumericThousandSep = "."
             
             Set cl = .ColumnsHeader.Add("ImportoUnitario", "Importo unitario", dgDouble, True, 2000, dgAlignRight)
                cl.BackColor = vbYellow
                cl.Editable = True
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
        
        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With
    Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "Reperimento dati prodotti"

Exit Sub

End Sub

Private Sub Form_Resize()
    Me.GrigliaCorpo.Width = Me.Width - 180
    
    Me.cmdRiporta.Left = Me.Width - Me.cmdRiporta.Width - 180
    
    
End Sub

Private Sub GrigliaCorpo_DblClick()
On Error GoTo ERR_GrigliaCorpo_DblClick
If ((rsGriglia.EOF) And (rsGriglia.BOF)) Then Exit Sub

frmSelPeriodo.Show vbModal

If LINK_UM_PERIODO_SEL > 0 Then
    rsGriglia.Fields("IDRV_POUnitaDiMisuraPeriodo").Value = LINK_UM_PERIODO_SEL
    rsGriglia.Fields("UMPeriodo").Value = DESCRIZIONE_UM_PERIODO_SEL
    
    Me.GrigliaCorpo.Refresh

End If

Exit Sub
ERR_GrigliaCorpo_DblClick:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_DblClick"
End Sub

Private Sub GrigliaCorpo_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_GrigliaCorpo_KeyPress
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GrigliaCorpo.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsGriglia.Fields("Riporta").Value), 2
        End If
    End If

Exit Sub
ERR_GrigliaCorpo_KeyPress:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_KeyPress"
End Sub

Private Sub GrigliaCorpo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_GrigliaCorpo_MouseUp
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If GrigliaCorpo.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < GrigliaCorpo.ColumnsHeader("Riporta").Width Then
            'Se non siamo in modalità filtri
            If GrigliaCorpo.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsGriglia.Fields("Riporta").Value), 2
            End If
        End If
    End If
Exit Sub
ERR_GrigliaCorpo_MouseUp:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_MouseUp"
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
On Error GoTo ERR_sbSelectSelectedRow
        If Not rsGriglia.EOF And Not rsGriglia.BOF Then
        
            rsGriglia.Fields("Riporta").Value = Abs(CLng(Selected))
            
            Me.GrigliaCorpo.Refresh
            
        End If

Exit Sub
ERR_sbSelectSelectedRow:
    MsgBox Err.Description, vbCritical, "sbSelectSelectedRow"
End Sub
