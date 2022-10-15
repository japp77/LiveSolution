VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelProdottiServizi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SELEZIONE PRODOTTI DEL CONTRATTO"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15780
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
   ScaleHeight     =   7650
   ScaleWidth      =   15780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   4
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton cmdEliminaFiltri 
         Height          =   315
         Left            =   2280
         TabIndex        =   17
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox txtArticolo 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtDescrizioneServizio 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtCodiceArticolo 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   2535
      End
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
         TabIndex        =   6
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
      Begin DMTDataCmb.DMTCombo cboAssociato 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   3600
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
      Begin VB.Label lblProdotto 
         Caption         =   "Prodotto associato al servizio"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label lblProdotto 
         Caption         =   "Descrizione servizio"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblProdotto 
         Caption         =   "Codice articolo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblProdotto 
         Caption         =   "Articolo"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblProdotto 
         Caption         =   "Prodotto generico"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label lblProdotto 
         Caption         =   "Matricola"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdRiporta 
      Caption         =   "ASSOCIA"
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
      Left            =   13440
      TabIndex        =   1
      Top             =   7080
      Width           =   2175
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   6975
      Left            =   2880
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      _ExtentX        =   22675
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   2880
      TabIndex        =   2
      Top             =   7080
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   7200
      Width           =   10335
   End
End
Attribute VB_Name = "frmSelProdottiServizi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

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

rsGriglia.Fields.Append "IDRV_POContrattoProdotti", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "IDRV_POContrattoServiziProdotti", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "IDRV_POProdotto", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "Prodotto", adChar, 250, adFldIsNullable
rsGriglia.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "CodiceArticolo", adChar, 250, adFldIsNullable
rsGriglia.Fields.Append "Articolo", adChar, 250, adFldIsNullable
rsGriglia.Fields.Append "ProdottoGenerico", adBoolean, , adFldIsNullable
rsGriglia.Fields.Append "Matricola", adChar, 250, adFldIsNullable
rsGriglia.Fields.Append "Quantita", adDouble, , adFldIsNullable
rsGriglia.Fields.Append "Riporta", adBoolean, , adFldIsNullable
rsGriglia.Fields.Append "Associato", adBoolean, , adFldIsNullable

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic

'CONTEGGIO RECORD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT COUNT (IDRV_POContrattoProdotti) as NumeroRecord "
sSQL = sSQL & "FROM RV_POContrattoProdotti"
sSQL = sSQL & " WHERE IDRV_POContratto=" & Link_Contratto
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

sSQL = "SELECT * FROM RV_POIEProdottiContratto "
sSQL = sSQL & " WHERE IDRV_POContratto=" & Link_Contratto
sSQL = sSQL & " AND Dismesso=" & fnNormBoolean(0)

Set rs = Cn.OpenResultset(sSQL)
Avvia = False

While Not rs.EOF
    
    rsGriglia.AddNew
        rsGriglia!IDRV_POContrattoProdotti = fnNotNullN(rs!IDRV_POContrattoProdotti)
        rsGriglia!IDRV_POContrattoServiziProdotti = GET_PRODOTTO_ASSOCIATO(fnNotNullN(rs!IDRV_POContrattoProdotti), Link_Contratto_Servizio)
        rsGriglia!IDRV_POContrattoProdotti = fnNotNullN(rs!IDRV_POContrattoProdotti)
        rsGriglia!IDRV_POProdotto = fnNotNullN(rs!IDRV_POProdotto)
        rsGriglia!Prodotto = fnNotNull(rs!Descrizione)
        rsGriglia!IDArticolo = fnNotNullN(rs!IDArticolo)
        rsGriglia!CodiceArticolo = fnNotNull(rs!CodiceArticolo)
        rsGriglia!Articolo = fnNotNull(rs!Articolo)
        rsGriglia!ProdottoGenerico = rs!ProdottoGenerico
        rsGriglia!Matricola = fnNotNull(rs!Matricola)
        rsGriglia!Quantita = fnNotNullN(rs!Quantita)
        rsGriglia!Riporta = False
        If (rsGriglia!IDRV_POContrattoServiziProdotti > 0) Then
            If (GET_PRODOTTO_ELIMINATO(fnNotNullN(rs!IDRV_POContrattoProdotti), Link_Contratto_Servizio) = True) Then
                rsGriglia!Riporta = False
            Else
                rsGriglia!Riporta = True
            End If
        End If
        rsGriglia!Associato = rsGriglia!Riporta
        
    rsGriglia.Update
    
    
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
Private Function GET_PRODOTTO_ASSOCIATO(IDProdottoContratto As Long, IDServizioContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POContrattoServiziProdotti FROM RV_POContrattoServiziProdotti "
sSQL = sSQL & "WHERE IDRV_POContrattoServizi=" & IDServizioContratto
sSQL = sSQL & " AND IDRV_POContrattoProdotti=" & IDProdottoContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PRODOTTO_ASSOCIATO = 0
Else
    GET_PRODOTTO_ASSOCIATO = fnNotNullN(rs!IDRV_POContrattoServiziProdotti)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_PRODOTTO_ELIMINATO(IDProdottoContratto As Long, IDServizioContratto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Eliminato FROM RV_POContrattoServiziProdotti "
sSQL = sSQL & "WHERE IDRV_POContrattoServizi=" & IDServizioContratto
sSQL = sSQL & " AND IDRV_POContrattoProdotti=" & IDProdottoContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PRODOTTO_ELIMINATO = False
Else
    If rs!Eliminato = True Then
        GET_PRODOTTO_ELIMINATO = True
    Else
        GET_PRODOTTO_ELIMINATO = False
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
            Set cl = .ColumnsHeader.Add("Riporta", "Riporta", dgBoolean, True, 1000, dgAligncenter)
                cl.Editable = True
            .ColumnsHeader.Add "IDRV_POContrattoProdotti", "IDRV_POContrattoProdotti", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POContrattoServiziProdotti", "IDRV_POContrattoServiziProdotti", dgInteger, False, 500, dgAlignleft
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
    MsgBox Err.Description, vbCritical, "Reperimento dati documentazione"


End Sub

Private Sub cboAssociato_Click()
    GET_FILTRO_SQL

End Sub

Private Sub cboProdottoGenerico_Click()
    GET_FILTRO_SQL

End Sub

Private Sub cmdEliminaFiltri_Click()
    Me.txtArticolo.Text = ""
    Me.txtDescrizioneServizio.Text = ""
    Me.txtCodiceArticolo.Text = ""
    Me.txtMatricola.Text = ""
    Me.cboAssociato.WriteOn 0
    Me.cboProdottoGenerico.WriteOn 0
    
End Sub

Private Sub cmdRiporta_Click()
On Error GoTo ERR_cmdRiporta_Click
Dim sSQL As String
Dim rs As ADODB.Recordset


rsGriglia.Update

rsGriglia.Filter = "Riporta=" & fnNormBoolean(True) & " OR Associato=" & fnNormBoolean(True)

Set rs = New ADODB.Recordset

If Not ((rsGriglia.EOF) And (rsGriglia.BOF)) Then
    While Not rsGriglia.EOF
    
        Set rs = Nothing
        
        sSQL = "SELECT * FROM RV_POContrattoServiziProdotti "
        sSQL = sSQL & "WHERE IDRV_POContrattoServiziProdotti=" & rsGriglia!IDRV_POContrattoServiziProdotti
        Set rs = New ADODB.Recordset
        
        rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
        If rs.EOF Then
            rs.AddNew
        End If
        
        rs!IDRV_POContrattoServizi = Link_Contratto_Servizio
        rs!IDRV_POContrattoProdotti = rsGriglia!IDRV_POContrattoProdotti
        rs!Eliminato = False
        If rsGriglia!IDRV_POContrattoServiziProdotti > 0 Then
            If ((rsGriglia!Associato = True) And (rsGriglia!Riporta = False)) Then
                rs!Eliminato = True
            End If
        End If
        rs.Update
    rsGriglia.MoveNext
    Wend
End If

Unload Me
Exit Sub
ERR_cmdRiporta_Click:
    MsgBox Err.Description, vbCritical, "cmdRiporta_Click"
End Sub

Private Sub Form_Load()
    CREATE_RECORDSET
    GET_GRIGLIA
    INIT_CONTROLLI
    
    Me.cboAssociato.WriteOn FILTRO_PRODOTTO_ASSOCIATO
    
End Sub
Private Sub GrigliaCorpo_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
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
Dim Testo As String

        If Not rsGriglia.EOF And Not rsGriglia.BOF Then
            
            If PermessoAggiornamento(Link_Contratto_Servizio, rsGriglia!IDRV_POContrattoProdotti) = False Then
                Testo = "ATTENZIONE!!!" & vbCrLf
                Testo = Testo & "Il prodotto per questo servizio risulta movimentato negli interventi che sono stati elaborati in precedenza" & vbCrLf
                Testo = Testo & "Impossibile eliminare il riferimento del prodotto"
                Exit Sub
            End If
            
            rsGriglia.Fields("Riporta").Value = Abs(CLng(Selected))
            
            Me.GrigliaCorpo.Refresh
            
        End If
End Sub
Private Sub INIT_CONTROLLI()
    With Me.cboProdottoGenerico
        Set .Database = Cn
        .AddFieldKey "IDRV_POSINO"
        .DisplayField = "SINO"
        .SQL = "SELECT * FROM RV_POSINO ORDER BY SINO"
        .Fill
    End With
    
    With Me.cboAssociato
        Set .Database = Cn
        .AddFieldKey "IDRV_POSINO"
        .DisplayField = "SINO"
        .SQL = "SELECT * FROM RV_POSINO ORDER BY SINO"
        .Fill
    End With
End Sub
Private Sub GET_FILTRO_SQL()
    rsGriglia.Filter = vbNullString
    
    If Len(Trim(Me.txtDescrizioneServizio.Text)) > 0 Then
        If rsGriglia.Filter = 0 Then
            rsGriglia.Filter = "Prodotto LIKE " & fnNormString(txtDescrizioneServizio.Text & "%")
        Else
            rsGriglia.Filter = rsGriglia.Filter & " AND Prodotto LIKE " & fnNormString(txtDescrizioneServizio.Text & "%")
        End If
    End If
    
    If Len(Trim(Me.txtCodiceArticolo.Text)) > 0 Then
        If rsGriglia.Filter = 0 Then
            rsGriglia.Filter = "CodiceArticolo LIKE " & fnNormString(txtCodiceArticolo.Text & "%")
        Else
            rsGriglia.Filter = rsGriglia.Filter & " AND CodiceArticolo LIKE " & fnNormString(txtCodiceArticolo.Text & "%")
        End If
    End If
    
    If Len(Trim(Me.txtArticolo.Text)) > 0 Then
        If rsGriglia.Filter = 0 Then
            rsGriglia.Filter = "Articolo LIKE " & fnNormString(txtArticolo.Text & "%")
        Else
            rsGriglia.Filter = rsGriglia.Filter & " AND Articolo LIKE " & fnNormString(txtArticolo.Text & "%")
        End If
    End If
    
    If Len(Trim(Me.txtMatricola.Text)) > 0 Then
        If rsGriglia.Filter = 0 Then
            rsGriglia.Filter = "Matricola LIKE " & fnNormString(txtMatricola.Text & "%")
        Else
            rsGriglia.Filter = rsGriglia.Filter & " AND Matricola LIKE " & fnNormString(txtMatricola.Text & "%")
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
    
    If Me.cboAssociato.CurrentID > 0 Then
        If Me.cboAssociato.CurrentID = 1 Then
            If rsGriglia.Filter = 0 Then
                rsGriglia.Filter = "Riporta=" & fnNormBoolean(True)
            Else
                rsGriglia.Filter = rsGriglia.Filter & " AND Riporta=" & fnNormBoolean(True)
            End If
        End If
        If Me.cboAssociato.CurrentID = 2 Then
            If rsGriglia.Filter = 0 Then
                rsGriglia.Filter = "Riporta=" & fnNormBoolean(False)
            Else
                rsGriglia.Filter = rsGriglia.Filter & " AND Riporta=" & fnNormBoolean(False)
            End If
        End If
    End If
    
    
    GET_GRIGLIA
    
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

Private Sub txtMatricola_Change()
    GET_FILTRO_SQL

End Sub
Private Function PermessoAggiornamento(IDRigaServizioContratto As Long, IDRigaProdottoContratto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT IDRV_POIntervento FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POContrattoServizi=" & IDRigaServizioContratto
sSQL = sSQL & " AND IDRV_POContrattoProdotti=" & IDRigaProdottoContratto
sSQL = sSQL & " AND Manuale=1"
sSQL = sSQL & " AND Elaborata=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    PermessoAggiornamento = True
Else
    PermessoAggiornamento = False
End If

rs.CloseResultset
Set rs = Nothing
End Function
