VERSION 5.00
Begin VB.Form frmGruppo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestione gruppi"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkIdentificaEmail 
      Caption         =   "Identifica una email"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox txtGruppo 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Gruppo"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmGruppo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GET_ESISTENZA_GRUPPO(Me.txtGruppo.Text, LINK_GRUPPO) Then
            MsgBox "Gruppo esistente", vbCritical, "Impossibile salvare"
            Exit Sub
        End If

        SALVA_GRUPPO LINK_GRUPPO
        Unload Me
    End If
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    If LINK_GRUPPO > 0 Then
        Me.txtGruppo.Text = frmProprieta.cboGruppo.Text
        Me.chkIdentificaEmail.Value = GET_IDENTIFICA_EMAIL(LINK_GRUPPO)
    Else
        Me.txtGruppo.Text = ""
        Me.chkIdentificaEmail.Value = 0
    End If
End Sub
Private Sub SALVA_GRUPPO(IDGruppo As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset




sSQL = "SELECT * FROM RV_POContactGruppo "
sSQL = sSQL & "WHERE IDRV_POContactGruppo=" & IDGruppo

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
    rs!IDRV_POContactGruppo = fnGetNewKey("RV_POContactGruppo", "IDRV_POContactGruppo")
   
End If

rs!ContactGruppo = Me.txtGruppo.Text
rs!Email = Me.chkIdentificaEmail.Value
rs.Update

LINK_GRUPPO = rs!IDRV_POContactGruppo

rs.Close
Set rs = Nothing
End Sub
Private Function GET_IDENTIFICA_EMAIL(IDGruppo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POContactGruppo "
sSQL = sSQL & "WHERE IDRV_POContactGruppo=" & IDGruppo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_IDENTIFICA_EMAIL = 0
Else
    GET_IDENTIFICA_EMAIL = Abs(fnNotNullN(rs!Email))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ESISTENZA_GRUPPO(Gruppo As String, IDGruppo As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POContactGruppo "
sSQL = sSQL & "WHERE ContactGruppo=" & fnNormString(Gruppo)
sSQL = sSQL & " AND IDRV_POContactGruppo<>" & IDGruppo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_GRUPPO = False
Else
    GET_ESISTENZA_GRUPPO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
