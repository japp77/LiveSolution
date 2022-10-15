VERSION 5.00
Begin VB.Form frmAttesa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
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
   ScaleHeight     =   3390
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2460
      Left            =   40
      Picture         =   "frmAttesa.frx":0000
      ScaleHeight     =   2460
      ScaleWidth      =   9015
      TabIndex        =   1
      Top             =   120
      Width           =   9015
      Begin VB.Label lblInfo2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ATTENDERE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2040
         Width           =   8775
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7920
      Top             =   1800
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   8895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   3360
      Left            =   15
      Top             =   10
      Width           =   9075
   End
End
Attribute VB_Name = "frmAttesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
    Me.Timer1.Enabled = True

End Sub


Public Function FORM_ATTESA_UNLOAD()
    Me.Timer1.Enabled = False

    Unload Me
End Function

Private Sub Timer1_Timer()
    DoEvents
End Sub
