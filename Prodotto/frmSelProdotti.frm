VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmSelProdotti 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Seleziona contatori"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11595
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
   ScaleHeight     =   7755
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   120
      Picture         =   "frmSelProdotti.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   855
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
      Left            =   1200
      Picture         =   "frmSelProdotti.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   855
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
      Left            =   9360
      TabIndex        =   1
      Top             =   7200
      Width           =   2175
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
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
End
Attribute VB_Name = "frmSelProdotti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub cmdRiporta_Click()
On Error GoTo ERR_cmdRiporta_Click
Dim sSQL As String
Dim rs As ADODB.Recordset


sSQL = "DELETE FROM RV_POProdottoContatori "
sSQL = sSQL & "WHERE IDRV_POProdotto=" & LINK_PRODOTTO
cn.Execute sSQL

sSQL = "SELECT * FROM RV_POProdottoContatori "
sSQL = sSQL & "WHERE IDRV_POProdottoContatori=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, cn.InternalConnection, adOpenKeyset, adLockPessimistic

rsGriglia.Filter = "Riporta=1"

While Not rsGriglia.EOF

    rs.AddNew
        rs!IDRV_POProdottoContatori = fnGetNewKey("RV_POProdottoContatori", "IDRV_POProdottoContatori")
        rs!IDRV_POProdotto = LINK_PRODOTTO
        rs!IDRV_POContatoreProdotto = fnNotNullN(rsGriglia!IDRV_POContatoreProdotto)
    rs.Update

rsGriglia.MoveNext
Wend

rs.Close
Set rs = Nothing

Unload Me
Exit Sub
ERR_cmdRiporta_Click:
    MsgBox Err.Description, vbCritical, "cmdRiporta_Click"
End Sub

Private Sub Form_Load()
    
    CREATE_RECORDSET
    
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

rsGriglia.Fields.Append "IDRV_POContatoreProdotto", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "Codice", adChar, 250, adFldIsNullable
rsGriglia.Fields.Append "Descrizione", adChar, 250, adFldIsNullable
rsGriglia.Fields.Append "Riporta", adBoolean, , adFldIsNullable

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic

sSQL = "SELECT * FROM RV_POIEContatoreProdotto "

Set rs = cn.OpenResultset(sSQL)
Avvia = True

While Not rs.EOF
    If Avvia = True Then
        
        rsGriglia.AddNew
            
            rsGriglia!Riporta = GET_ESISTENZA_PRODOTTO(fnNotNullN(rs!IDRV_POContatoreProdotto), LINK_PRODOTTO)
            rsGriglia!IDRV_POContatoreProdotto = fnNotNullN(rs!IDRV_POContatoreProdotto)
            rsGriglia!Codice = fnNotNull(rs!Codice)
            rsGriglia!Descrizione = fnNotNull(rs!Descrizione)
            
            
        rsGriglia.Update
        
    End If
    
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
Private Function GET_ESISTENZA_PRODOTTO(IDContatore As Long, IDProdotto As Long) As Boolean
On Error GoTo ERR_GET_ESISTENZA_PRODOTTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_ESISTENZA_PRODOTTO = False
sSQL = "SELECT * FROM RV_POProdottoContatori "
sSQL = sSQL & " WHERE IDRV_POProdotto=" & IDProdotto
sSQL = sSQL & " AND IDRV_POContatoreProdotto=" & IDContatore

Set rs = cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_ESISTENZA_PRODOTTO = True
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_ESISTENZA_PRODOTTO:
    MsgBox Err.Description, vbCritical, "GET_ESISTENZA_PRODOTTO"
End Function
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    OLDCursor = cn.CursorLocation
    cn.CursorLocation = 3

    With Me.GrigliaCorpo
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
            Set cl = .ColumnsHeader.Add("Riporta", "Riporta", dgBoolean, True, 1000, dgAligncenter)
                cl.Editable = True
            .ColumnsHeader.Add "IDRV_POContatoreProdotto", "IDRV_POContatoreProdotto", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Codice", "Codice", dgchar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "Descrizione", "Descrizione", dgchar, True, 2500, dgAlignleft
        
        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With
    cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "Reperimento dati prodotti"

Exit Sub

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

Private Sub GrigliaCorpo_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error GoTo ERR_GrigliaCorpo_MouseUp
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If GrigliaCorpo.HitTest(X, y) > 0 Then
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
