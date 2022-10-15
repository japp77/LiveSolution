VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDistinta 
   Caption         =   "Distinta articolo"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDistinta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   12270
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   8160
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   8520
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Visualizza solamente gli articoli della distinta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   7680
      Width           =   4815
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   2
      Top             =   7320
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8760
      TabIndex        =   1
      Top             =   7560
      Width           =   3495
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   12726
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
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Ricerca per descrizione"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   8520
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Ricerca per codice"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   8160
      Width           =   1695
   End
End
Attribute VB_Name = "frmDistinta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private rsDistinta As ADODB.Recordset

Private Sub Check1_Click()
    GET_GRIGLIA
    
End Sub

Private Sub cmdConferma_Click()
    
    CONFERMA LINK_PRODOTTO

End Sub

Private Sub Form_Activate()
    CREA_RECORDSET
End Sub

Private Sub CREA_RECORDSET()
On Error GoTo ERR_CREA_RECORDSET
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim NumeroRecord As Long
Dim UnitaProgresso As Double

NumeroRecord = 0

GET_DISTINTA LINK_PRODOTTO

If Not (rsGriglia Is Nothing) Then
    If rsGriglia.State > 0 Then
        rsGriglia.Close
    End If
    Set rsGriglia = Nothing
End If

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient


rsGriglia.Fields.Append "Riporta", adBoolean, , adFldIsNullable
rsGriglia.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "CodiceArticolo", adVarChar, 50, adFldIsNullable
rsGriglia.Fields.Append "DescrizioneArticolo", adVarChar, 250, adFldIsNullable

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic

sSQL = "SELECT COUNT(IDArticolo) AS NumeroRecord "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm

Set rs = New ADODB.Recordset
rs.Open sSQL, cn.InternalConnection

If rs.EOF = False Then
    NumeroRecord = rs!NumeroRecord
End If

rs.Close
Set rs = Nothing

If NumeroRecord = 0 Then Exit Sub

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100

UnitaProgresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 2)



sSQL = "SELECT IDArticolo, CodiceArticolo, Articolo "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm

Set rs = New ADODB.Recordset
rs.Open sSQL, cn.InternalConnection

While Not rs.EOF
    rsGriglia.AddNew
        rsGriglia!Riporta = GET_ESISTENZA_PRODOTTO(fnNotNullN(rs!IDArticolo), LINK_PRODOTTO)
        rsGriglia!IDArticolo = rs!IDArticolo
        rsGriglia!CodiceArticolo = rs!CodiceArticolo
        rsGriglia!DescrizioneArticolo = rs!Articolo
    rsGriglia.Update
    
    If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
    End If
    DoEvents
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

rsDistinta.Close
Set rsDistinta = Nothing

GET_GRIGLIA

Me.ProgressBar1.Value = 0
Exit Sub
ERR_CREA_RECORDSET:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET"
End Sub
Private Sub GET_DISTINTA(IDProdotto As Long)
On Error GoTo ERR_GET_DISTINTA
Dim sSQL As String
Dim rs As ADODB.Recordset


Set rsDistinta = New ADODB.Recordset
rsDistinta.CursorLocation = adUseClient

rsDistinta.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsDistinta.Fields.Append "IDProdotto", adInteger, , adFldIsNullable
rsDistinta.Open , , adOpenKeyset, adLockBatchOptimistic


sSQL = "SELECT * FROM RV_POProdottoDistinta "
sSQL = sSQL & "WHERE IDRV_POProdotto=" & IDProdotto

Set rs = New ADODB.Recordset

rs.Open sSQL, cn.InternalConnection

While Not rs.EOF
    rsDistinta.AddNew
        rsDistinta!IDArticolo = rs!IDArticolo
        rsDistinta!IDProdotto = rs!IDRV_POProdotto
    rsDistinta.Update
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Exit Sub
ERR_GET_DISTINTA:
    MsgBox Err.Description, vbCritical, "GET_DISTINTA"
End Sub

Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    rsGriglia.Filter = vbNullString
    
    If Me.Check1.Value = vbChecked Then
        If rsGriglia.Filter = 0 Then
            rsGriglia.Filter = "Riporta=1"
        Else
            rsGriglia.Filter = rsGriglia.Filter & " AND Riporta=1"
        End If
    End If
    
    If Len(Me.Text1.Text) > 0 Then
        If rsGriglia.Filter = 0 Then
            rsGriglia.Filter = "CodiceArticolo LIKE " & fnNormString("%" & Text1.Text & "%")
        Else
            rsGriglia.Filter = rsGriglia.Filter & " AND CodiceArticolo LIKE " & fnNormString("%" & Text1.Text & "%")
        End If
    End If
    
    If Len(Me.Text2.Text) > 0 Then
        If rsGriglia.Filter = 0 Then
            rsGriglia.Filter = "DescrizioneArticolo LIKE " & fnNormString("%" & Text2.Text & "%")
        Else
            rsGriglia.Filter = rsGriglia.Filter & " AND DescrizioneArticolo LIKE " & fnNormString("%" & Text2.Text & "%")
        End If
    End If
    
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
            .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "CodiceArticolo", "Codice", dgchar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "DescrizioneArticolo", "Descrizione", dgchar, True, 4500, dgAlignleft
        
        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With
    cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "Reperimento dati articoli"

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
Private Function GET_ESISTENZA_PRODOTTO(IDArticolo As Long, IDProdotto As Long) As Boolean
On Error GoTo ERR_GET_ESISTENZA_PRODOTTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_ESISTENZA_PRODOTTO = False


rsDistinta.Filter = "IDProdotto=" & IDProdotto
rsDistinta.Filter = rsDistinta.Filter & " AND IDArticolo=" & IDArticolo


If rsDistinta.EOF = False Then
    GET_ESISTENZA_PRODOTTO = True
End If

rsDistinta.Filter = vbNullString


Exit Function
ERR_GET_ESISTENZA_PRODOTTO:
    MsgBox Err.Description, vbCritical, "GET_ESISTENZA_PRODOTTO"
End Function
Private Sub CONFERMA(IDProdotto As Long)
On Error GoTo ERR_CONFERMA
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "DELETE FROM RV_POProdottoDistinta "
sSQL = sSQL & "WHERE IDRV_POProdotto=" & IDProdotto
cn.Execute sSQL


sSQL = "SELECT * FROM RV_POProdottoDistinta "
sSQL = sSQL & "WHERE IDRV_POProdotto=" & IDProdotto

Set rs = New ADODB.Recordset

rs.Open sSQL, cn.InternalConnection, adOpenKeyset, adLockPessimistic


rsGriglia.Filter = "Riporta=1"


While Not rsGriglia.EOF
    rs.AddNew
        rs!IDRV_POProdotto = IDProdotto
        rs!IDArticolo = rsGriglia!IDArticolo
    rs.Update
rsGriglia.MoveNext
Wend

rs.Close
Set rs = Nothing

CREA_RECORDSET

Exit Sub
ERR_CONFERMA:
    MsgBox Err.Description, vbCritical, "CONFERMA"
End Sub

Private Sub Text1_Change()
    GET_GRIGLIA
    
End Sub

Private Sub Text2_Change()
    GET_GRIGLIA
    
End Sub
