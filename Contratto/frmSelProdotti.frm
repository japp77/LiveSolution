VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelProdotti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SELEZIONE PRODOTTI"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14460
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
   ScaleHeight     =   7605
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDeSelTutto 
      Height          =   375
      Left            =   1920
      Picture         =   "frmSelProdotti.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton cmdSelTutto 
      Height          =   375
      Left            =   120
      Picture         =   "frmSelProdotti.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   2880
      TabIndex        =   14
      Top             =   7080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
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
      Height          =   495
      Left            =   12240
      TabIndex        =   13
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox txtArticolo 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Frame fraRicerca 
      Caption         =   "FILTRI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6975
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2775
      Begin VB.TextBox txtMatricola 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   3000
         Width           =   2535
      End
      Begin DMTDataCmb.DMTCombo cboProdottoGenerico 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtCodiceArticolo 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtDescrizioneServizio 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblProdotto 
         Caption         =   "Matricola"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label lblProdotto 
         Caption         =   "Prodotto generico"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label lblProdotto 
         Caption         =   "Articolo"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblProdotto 
         Caption         =   "Codice articolo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblProdotto 
         Caption         =   "Descrizione servizio"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   6975
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
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
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   7200
      Width           =   9015
   End
   Begin VB.Label lblProdotto 
      Caption         =   "Descrizione servizio"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   2535
   End
End
Attribute VB_Name = "frmSelProdotti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub cboProdottoGenerico_Click()
    GET_FILTRO_SQL
End Sub

Private Sub cmdDeSelTutto_Click()
If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
    rsGriglia.MoveFirst
    
    While Not rsGriglia.EOF
        rsGriglia!Riporta = False
        rsGriglia.Update
    rsGriglia.MoveNext
    Wend
    
    Me.GrigliaCorpo.Refresh
End If
End Sub

Private Sub cmdRiporta_Click()
On Error GoTo ERR_cmdRiporta_Click
Dim rsNew As ADODB.Recordset
Dim sSQL As String


    rsGriglia.Update
    
    rsGriglia.Filter = "Riporta=" & fnNormBoolean(1)
        
    sSQL = "SELECT * FROM RV_POContrattoProdotti "
    sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=0"
    
    Set rsNew = New ADODB.Recordset
    rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
        
    While Not rsGriglia.EOF
        rsNew.AddNew
            rsNew!IDRV_POContrattoProdotti = fnGetNewKey("RV_POContrattoProdotti", "IDRV_POContrattoProdotti")
            rsNew!IDRV_POContratto = Link_Contratto
            rsNew!IDRV_POContrattoPadre = frmMain.txtIDContrattoPadre.Value
            rsNew!IDRV_POProdotto = rsGriglia!IDRV_POProdotto
            rsNew!Quantita = rsGriglia!Quantita
            rsNew!Dismesso = 0
        rsNew.Update
        
        AGGIORNA_INTERVENTI_PRODOTTO rsNew!IDRV_POProdotto, rsNew!IDRV_POContrattoProdotti
        
    rsGriglia.MoveNext
    Wend
    
    rsNew.Close
    Set rsNew = Nothing
    
    Unload Me

Exit Sub
ERR_cmdRiporta_Click:
    MsgBox Err.Description, vbCritical, "cmdRiporta_Click"
End Sub
Private Sub AGGIORNA_INTERVENTI_PRODOTTO(IDProdotto As Long, IDRigaProdottoContratto As Long)
On Error GoTo ERR_AGGIORNA_INTERVENTI_PRODOTTO
Dim sSQL As String


sSQL = "UPDATE RV_POIntervento SET "
sSQL = sSQL & "IDRV_POContratto=" & Link_Contratto & ", "
sSQL = sSQL & "IDAnagraficaCliente=" & frmMain.CDCliente.KeyFieldID & ", "
sSQL = sSQL & "IDRV_POContrattoPadre=" & frmMain.txtIDContrattoPadre.Value & ", "
sSQL = sSQL & "IDAnagraficaFatturazione=" & IDClienteFatturazione & ", "
sSQL = sSQL & "IDRV_POContrattoProdotti=" & IDRigaProdottoContratto

sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDRV_POProdotto=" & IDProdotto
sSQL = sSQL & " AND GeneratoDa=1"
sSQL = sSQL & " AND DataAppuntamento>" & fnNormDate(Date)


Cn.Execute sSQL
Exit Sub
ERR_AGGIORNA_INTERVENTI_PRODOTTO:
    MsgBox Err.Description, vbCritical, "AGGIORNA_INTERVENTI_PRODOTTO"
End Sub
Private Sub cmdSelTutto_Click()
If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
    rsGriglia.MoveFirst
    
    While Not rsGriglia.EOF
        rsGriglia!Riporta = True
        rsGriglia.Update
    rsGriglia.MoveNext
    Wend
    
    Me.GrigliaCorpo.Refresh
End If
End Sub

Private Sub Form_Activate()
    CREATE_RECORDSET
End Sub

Private Sub Form_Load()

    With Me.cboProdottoGenerico
        Set .Database = Cn
        .AddFieldKey "IDRV_POSINO"
        .DisplayField = "SINO"
        .SQL = "SELECT * FROM RV_POSINO ORDER BY SINO"
        .Fill
    End With

    If frmMain.cboTipoImpostazione.CurrentID <> 1 Then
        Me.cmdRiporta.Enabled = False
        Me.cmdDeSelTutto.Enabled = False
        Me.cmdSelTutto.Enabled = False
    End If

End Sub
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

rsGriglia.Fields.Append "IDRV_POProdotto", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "Prodotto", adChar, 255, adFldIsNullable
rsGriglia.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "CodiceArticolo", adChar, 255, adFldIsNullable
rsGriglia.Fields.Append "Articolo", adChar, 255, adFldIsNullable
rsGriglia.Fields.Append "ProdottoGenerico", adBoolean, , adFldIsNullable
rsGriglia.Fields.Append "Matricola", adChar, 255, adFldIsNullable
rsGriglia.Fields.Append "Quantita", adDouble, , adFldIsNullable
rsGriglia.Fields.Append "Riporta", adBoolean, , adFldIsNullable

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic

'CONTEGGIO RECORD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT COUNT (IDRV_POProdotto) as NumeroRecord "
sSQL = sSQL & "FROM RV_POProdotto"
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND Dismesso=" & fnNormBoolean(0)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If NumeroRecord = 0 Then
    GET_GRIGLIA
    Exit Sub
End If

UnitaProgresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 4)

sSQL = "SELECT * FROM RV_POIEProdotto "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND Dismesso=" & fnNormBoolean(0)

Set rs = Cn.OpenResultset(sSQL)
Avvia = False

While Not rs.EOF
    Avvia = False
    If (SEL_PROD_NON_CONTRATTO = 0) Then
        Avvia = True
    Else
        If (rs!ProdottoGenerico = False) Then
            If (GET_ESISTENZA_PRODOTTO_CONTRATTO(fnNotNullN(rs!IDRV_POProdotto)) = False) Then
                Avvia = True
            End If
        Else
            Avvia = True
        End If
    End If
    
    If Avvia = True Then
        rsGriglia.AddNew
            rsGriglia!Riporta = False
            rsGriglia!IDRV_POProdotto = fnNotNullN(rs!IDRV_POProdotto)
            rsGriglia!Prodotto = Mid(fnNotNull(rs!Descrizione), 1, 255)
            rsGriglia!IDArticolo = fnNotNullN(rs!IDArticolo)
            rsGriglia!CodiceArticolo = fnNotNull(rs!CodiceArticolo)
            rsGriglia!Articolo = Mid(fnNotNull(rs!Articolo), 1, 255)
            rsGriglia!ProdottoGenerico = rs!ProdottoGenerico
            rsGriglia!Matricola = fnNotNull(rs!Matricola)
            rsGriglia!Quantita = 1
        rsGriglia.Update
    End If
    
    If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
    End If
    NumeroElaborazione = NumeroElaborazione + 1
    lblInfo.Caption = "Elaborazione numero " & NumeroElaborazione & " di " & NumeroRecord
    DoEvents
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

lblInfo.Caption = ""
Me.ProgressBar1.Value = 0

GET_GRIGLIA

Exit Sub
ERR_CREATE_RECORDSET:
    MsgBox Err.Description, vbCritical, "CREATE_RECORDSET"
End Sub
Private Function GET_ESISTENZA_PRODOTTO_CONTRATTO(IDProdotto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POContrattoProdotti.IDRV_POContrattoProdotti, RV_POContrattoProdotti.IDRV_POContratto, RV_POContrattoProdotti.Dismesso, RV_POContratto.Disdetta, "
sSQL = sSQL & "RV_POContrattoProdotti.IDRV_POProdotto "
sSQL = sSQL & "FROM RV_POContrattoProdotti INNER JOIN "
sSQL = sSQL & "RV_POContratto ON RV_POContrattoProdotti.IDRV_POContratto = RV_POContratto.IDRV_POContratto"
sSQL = sSQL & " WHERE RV_POContrattoProdotti.IDRV_POProdotto=" & IDProdotto
sSQL = sSQL & " AND RV_POContrattoProdotti.Dismesso=" & fnNormBoolean(0)
sSQL = sSQL & " AND RV_POContratto.Disdetta=" & fnNormBoolean(0)
sSQL = sSQL & " AND RV_POContratto.ContrattoAttuale=" & fnNormBoolean(1)
sSQL = sSQL & " AND RV_POContratto.Chiuso=" & fnNormBoolean(0)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_PRODOTTO_CONTRATTO = False
Else
    If fnNotNullN(rs!Dismesso) = 1 Then
        GET_ESISTENZA_PRODOTTO_CONTRATTO = False
    Else
        GET_ESISTENZA_PRODOTTO_CONTRATTO = True
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3

    With Me.GrigliaCorpo
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
            If frmMain.cboTipoImpostazione.CurrentID = 1 Then
            Set cl = .ColumnsHeader.Add("Riporta", "Riporta", dgBoolean, True, 1000, dgAligncenter)
                cl.Editable = True
            End If
            .ColumnsHeader.Add "IDRV_POProdotto", "IDRV_PO08_Prodotto", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "Prodotto", "Prodotto", dgchar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "CodiceArticolo", "Codice articolo", dgchar, True, 1500, dgAlignleft
            .ColumnsHeader.Add "Articolo", "Articolo", dgchar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "ProdottoGenerico", "Generico", dgBoolean, True, 1000, dgAligncenter
            .ColumnsHeader.Add "Matricola", "Matricola", dgchar, True, 2000, dgAlignleft
            Set cl = .ColumnsHeader.Add("Quantita", "Quantità", dgDouble, True, 1500, dgAlignRight)
                cl.BackColor = vbYellow
                cl.Editable = True
        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With

    
    Cn.CursorLocation = OLDCursor

Exit Sub


ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub



Private Sub GrigliaCorpo_DblClick()
    If frmMain.cboTipoImpostazione.CurrentID = 1 Then Exit Sub
    
    frmMain.txtIDProdotto.Value = Me.GrigliaCorpo.AllColumns("IDRV_POProdotto").Value
    frmMain.txtQtaProdotto.Value = Me.GrigliaCorpo.AllColumns("Quantita").Value
    Unload Me
    
End Sub

Private Sub GrigliaCorpo_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If frmMain.cboTipoImpostazione.CurrentID <> 1 Then Exit Sub
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GrigliaCorpo.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsGriglia.Fields("Riporta").Value), 2
        End If
    End If
End Sub

Private Sub GrigliaCorpo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    If frmMain.cboTipoImpostazione.CurrentID <> 1 Then Exit Sub
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

End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)

        If Not rsGriglia.EOF And Not rsGriglia.BOF Then
        
            rsGriglia.Fields("Riporta").Value = Abs(CLng(Selected))
            
            Me.GrigliaCorpo.Refresh
            
            
            
        End If
End Sub

Private Sub GrigliaCorpo_ValidateFieldValue(ByVal Column As DmtGridCtl.dgColumnHeader, Value As Variant)
    If Column.FieldName = "Quantita" Then
        If rsGriglia!ProdottoGenerico = False Then
            Value = 1
            Me.GrigliaCorpo.Refresh
        End If
    End If
End Sub

Private Sub txtArticolo_Change()
    GET_FILTRO_SQL
End Sub

Private Sub txtCodiceArticolo_Change()
    GET_FILTRO_SQL
End Sub

Private Sub txtDescrizioneServizio_Change()

    GET_FILTRO_SQL
End Sub
Private Sub GET_FILTRO_SQL()
On Error GoTo ERR_GET_FILTRO_SQL
    rsGriglia.Filter = vbNullString
    
    If Len(Trim(Me.txtDescrizioneServizio.Text)) > 0 Then
        If rsGriglia.Filter = 0 Then
            rsGriglia.Filter = "Prodotto LIKE " & fnNormString("%" & txtDescrizioneServizio.Text & "%")
        Else
            rsGriglia.Filter = rsGriglia.Filter & " AND Prodotto LIKE " & fnNormString("%" & txtDescrizioneServizio.Text & "%")
        End If
    End If
    
    If Len(Trim(Me.txtCodiceArticolo.Text)) > 0 Then
        If rsGriglia.Filter = 0 Then
            rsGriglia.Filter = "CodiceArticolo LIKE " & fnNormString("%" & txtCodiceArticolo.Text & "%")
        Else
            rsGriglia.Filter = rsGriglia.Filter & " AND CodiceArticolo LIKE " & fnNormString("%" & txtCodiceArticolo.Text & "%")
        End If
    End If
    
    If Len(Trim(Me.txtArticolo.Text)) > 0 Then
        If rsGriglia.Filter = 0 Then
            rsGriglia.Filter = "Articolo LIKE " & fnNormString("%" & txtArticolo.Text & "%")
        Else
            rsGriglia.Filter = rsGriglia.Filter & " AND Articolo LIKE " & fnNormString("%" & txtArticolo.Text & "%")
        End If
    End If
    
    If Len(Trim(Me.txtMatricola.Text)) > 0 Then
        If rsGriglia.Filter = 0 Then
            rsGriglia.Filter = "Matricola LIKE " & fnNormString("%" & txtMatricola.Text & "%")
        Else
            rsGriglia.Filter = rsGriglia.Filter & " AND Matricola LIKE " & fnNormString("%" & txtMatricola.Text & "%")
        End If
    End If
    If Me.cboProdottoGenerico.CurrentID > 0 Then
        If Me.cboProdottoGenerico.CurrentID = 1 Then
            If rsGriglia.Filter = 0 Then
                rsGriglia.Filter = "ProdottoGenerico=" & fnNormBoolean(True)
            Else
                rsGriglia.Filter = rsGriglia.Filter & " AND ProdottoGenerico=" & fnNormBoolean(True)
            End If
        End If
        If Me.cboProdottoGenerico.CurrentID = 2 Then
            If rsGriglia.Filter = 0 Then
                rsGriglia.Filter = "ProdottoGenerico=" & fnNormBoolean(False)
            Else
                rsGriglia.Filter = rsGriglia.Filter & " AND ProdottoGenerico=" & fnNormBoolean(False)
            End If
        End If
    End If
    GET_GRIGLIA
Exit Sub
ERR_GET_FILTRO_SQL:
    MsgBox Err.Description, vbCritical, "GET_FILTRO_SQL"
End Sub

Private Sub txtMatricola_Change()
    GET_FILTRO_SQL
End Sub
