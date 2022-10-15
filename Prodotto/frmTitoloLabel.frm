VERSION 5.00
Begin VB.Form frmTitoloLabel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Titolo etichette"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2730
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
   ScaleHeight     =   2505
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodice1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   345
      Width           =   2535
   End
   Begin VB.TextBox txtCodice2 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtCodice3 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONFERMA"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblCodice1 
      Caption         =   "Codice 1"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblCodice2 
      Caption         =   "Codice 2"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblCodice3 
      Caption         =   "Codice 3"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "frmTitoloLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Codice1_Pred As String = "Codice 1"
Const Codice2_Pred As String = "Codice 2"
Const Codice3_Pred As String = "Codice 3"

Const Codice1_Pred_Field As String = "Codice1"
Const Codice2_Pred_Field As String = "Codice2"
Const Codice3_Pred_Field As String = "Codice3"

Const tabella As String = "Prodotto"

Private Sub Command1_Click()
Dim testo As String
Dim sSQL As String
Dim rs As ADODB.Recordset

testo = "Sei sicuro di voler modificare le etichette?"

If MsgBox(testo, vbQuestion + vbYesNo, "Modifica") = vbNo Then Exit Sub

sSQL = "DELETE FROM RV_POEtichetteCampo"
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND NomeTabella=" & fnNormString(tabella)
cn.Execute sSQL

sSQL = "SELECT * FROM RV_POEtichetteCampo "
sSQL = sSQL & "WHERE IDRV_POEtichetteCampo=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, cn.InternalConnection, adOpenKeyset, adLockPessimistic

SALVA_CAMPO Codice1_Pred_Field, rs, txtCodice1.Text
SALVA_CAMPO Codice2_Pred_Field, rs, txtCodice2.Text
SALVA_CAMPO Codice3_Pred_Field, rs, txtCodice3.Text

rs.Close
Set rs = Nothing

Unload Me

End Sub
Private Sub SALVA_CAMPO(NomeCampo As String, rs As ADODB.Recordset, testo As String)

rs.AddNew
    rs!IDRV_POEtichetteCampo = fnGetNewKey("RV_POEtichetteCampo", "IDRV_POEtichetteCampo")
    rs!IDAzienda = TheApp.IDFirm
    rs!NomeTabella = tabella
    rs!NomeCampo = NomeCampo
    rs!testo = testo
rs.Update

End Sub

Private Sub Form_Load()
    GET_DATI
End Sub
Private Sub GET_DATI()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Codice1 As String
Dim Codice2 As String
Dim Codice3 As String

Codice1 = Codice1_Pred
Codice2 = Codice2_Pred
Codice3 = Codice3_Pred

sSQL = "SELECT * FROM RV_POEtichetteCampo "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND NomeTabella=" & fnNormString(tabella)

Set rs = cn.OpenResultset(sSQL)

While Not rs.EOF
    Select Case fnNotNull(rs!NomeCampo)
        Case "Codice1"
            Codice1 = fnNotNull(rs!testo)
        Case "Codice2"
            Codice2 = fnNotNull(rs!testo)
        Case "Codice3"
            Codice3 = fnNotNull(rs!testo)
    End Select
rs.MoveNext
Wend



rs.CloseResultset
Set rs = Nothing

txtCodice1.Text = Codice1
txtCodice2.Text = Codice2
txtCodice3.Text = Codice3
End Sub

