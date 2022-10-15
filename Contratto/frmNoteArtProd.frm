VERSION 5.00
Begin VB.Form frmNoteArtProd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Descrizione aggiuntiva articolo del prodotto"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   2085
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNoteProdotto 
      Height          =   2055
      Left            =   0
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmNoteArtProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Me.txtNoteProdotto.Text = frmMain.txtNoteProdotto.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.txtNoteProdotto.Text = Me.txtNoteProdotto.Text
End Sub

