VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDettaglioIntervento 
   Caption         =   "Dettaglio intervento"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14865
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
   ScaleHeight     =   7185
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtRichesta 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   11880
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmDettaglioIntervento.frx":0000
   End
   Begin RichTextLib.RichTextBox txtLavoro 
      Height          =   6735
      Left            =   5040
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   11880
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmDettaglioIntervento.frx":007C
   End
   Begin RichTextLib.RichTextBox txtAnnotazioni 
      Height          =   6735
      Left            =   9960
      TabIndex        =   2
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   11880
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmDettaglioIntervento.frx":00F8
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Annotazioni"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   5
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Lavoro eseguito"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Richiesta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmDettaglioIntervento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
