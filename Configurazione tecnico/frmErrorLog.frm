VERSION 5.00
Object = "{FB1FAA0A-DAD5-11D1-850D-002018802E11}#3.0#0"; "DMTERROR.OCX"
Begin VB.Form frmErrorLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Errore"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6810
   Icon            =   "frmErrorLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DMTErrorLogCtrl.DmtError DMTErrorContol 
      Height          =   4380
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   7726
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogFile 
      Caption         =   "&Log file"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "frmErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdLogFile_Click()
    Me.DMTErrorContol.CreateLogFile
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub
