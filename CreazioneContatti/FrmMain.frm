VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREAZIONE CONTATTI    (Passo 1 di 3)"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdIIndietro 
      Caption         =   "Indietro"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Avanti"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdFine 
      Caption         =   "Fine"
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   4200
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4515
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   0
      Width           =   2895
   End
   Begin VB.OptionButton OptScelta 
      Caption         =   "Creazione di etichette"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   2160
      Width           =   4095
   End
   Begin VB.OptionButton OptScelta 
      Caption         =   "Creazione contatti per indirizzi di posta elettronica"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   1320
      Width           =   4095
   End
   Begin VB.OptionButton OptScelta 
      Caption         =   "Creazione contatti con numeri di FAX"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents m_App As DMTRunAppLib.Application
Attribute m_App.VB_VarHelpID = -1


Public Sub ConnessioneADO()
    If Not (CnDMT Is Nothing) Then
        CnDMT.CloseConnection
        Set CnDMT = Nothing
    End If
  Set CnDMT = m_App.Database.Connection
  
  
  
  VarPassword = m_App.Password
  VarUtente = m_App.User

    PrelevaAzienda
  
  Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set m_App = NewValue
End Property

Public Property Get Application() As DMTRunAppLib.Application
    Set Application = m_App
End Property

Private Sub cmdAnnulla_Click()
    If MsgBox("Vuoi abbandonare la procedura?", vbInformation + vbYesNo, "Chiusura applicazioni") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdAvanti_Click()
    Unload Me
End Sub
Public Sub InitControlli()
    
End Sub
Private Sub Form_Load()
    Me.OptScelta(0).Value = True

    DinamicNumber = 1
                    
    
    


    ControlloComandi
    
    
End Sub


Private Sub ControlloComandi()
    Me.cmdIIndietro.Enabled = False
    Me.cmdFine.Enabled = False
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If Me.cmdAvanti.Value = True Then
        frmParametri.Show
        Exit Sub
    End If
    
End Sub

Private Sub OptScelta_Click(Index As Integer)
    
    optOperazione = Index

End Sub
