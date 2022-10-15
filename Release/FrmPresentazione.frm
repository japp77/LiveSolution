VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmPresentazione 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FrmPresentazione.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   13665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "CONTROLLO COLLEGAMENTI DOCUMENTI"
      Height          =   615
      Left            =   11280
      TabIndex        =   29
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Frame frmLicenza 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Funzionalità abilitate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   5040
      TabIndex        =   27
      Top             =   2880
      Width           =   6015
      Begin DmtGridCtl.DmtGrid Griglia 
         Height          =   4815
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8493
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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   11280
      TabIndex        =   26
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdParametriFiliale 
      Caption         =   "CONFIGURAZIONE PARAMETRI FILIALE"
      Height          =   615
      Left            =   11280
      TabIndex        =   25
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HELP DESK IN ITALIANO"
      Height          =   615
      Left            =   11280
      TabIndex        =   24
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CommandButton cmdSbloccaUtente 
      Caption         =   "SBLOCCA UTENTE"
      Height          =   615
      Left            =   11280
      TabIndex        =   23
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdAggiornamento1 
      Caption         =   "Aggiornamento per la versione 3.02.00 (Facoltativo)"
      Height          =   855
      Left            =   11280
      TabIndex        =   22
      Tag             =   "Aggiornamento_3_02_00_4.exe"
      Top             =   7200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   13575
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data ultimo aggiornamento:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   18
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label lblDataInstallazione 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6120
         TabIndex        =   17
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label lblVersione 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   6120
         TabIndex        =   16
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Versione:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   15
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   2310
         Left            =   0
         Picture         =   "FrmPresentazione.frx":4781A
         Top             =   0
         Width           =   9000
      End
   End
   Begin VB.Frame FraHelpDesk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CONFIGURAZIONE HELP DESK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      Top             =   2880
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "Elimina avvio automatico"
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Top             =   4680
         Width           =   2295
      End
      Begin VB.CommandButton cmdAvvioAutomatico 
         Caption         =   "Avvio automatico"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Frame FraDatabase 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dati di avvio personalizzato"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3135
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   4695
         Begin VB.CommandButton cmdConferma 
            Caption         =   "CONFERMA"
            Height          =   255
            Left            =   1320
            TabIndex        =   19
            Top             =   2760
            Width           =   1815
         End
         Begin DMTDataCmb.DMTCombo cboAzienda 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   1680
            Width           =   4455
            _ExtentX        =   7858
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
         Begin VB.TextBox txtPassword 
            Height          =   285
            Left            =   2400
            TabIndex        =   10
            Top             =   1680
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtNomeUtente 
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   1680
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtNomeCatalogo 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   4455
         End
         Begin VB.TextBox txtNomeServer 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   4455
         End
         Begin DMTDataCmb.DMTCombo cboFiliale 
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   2280
            Width           =   4455
            _ExtentX        =   7858
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
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Filiale"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   2040
            Width           =   4095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Azienda"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome motore SQL SERVER"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome catalogo"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   4095
         End
      End
      Begin VB.OptionButton optScelta 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Avvio indipendente"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   4335
      End
      Begin VB.OptionButton optScelta 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Avvia da DMT PROFESSIONAL"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   44
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D3705
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D3A1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D3FB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D4553
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D46AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D4AFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D4E19
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D5133
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D5585
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D589F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D5CF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D6143
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D645D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D65B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D7409
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D785B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D9565
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D987F
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D9B99
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":D9FEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":DA43D
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":DA757
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":DA85D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":DACAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":DAFC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":DBB9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":DBFED
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":DC43F
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":DC891
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":DCCE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":DD135
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E33CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E36E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E3A03
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E3D1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E4037
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E4489
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E4A23
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E4FBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E5557
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E5AF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E608B
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E6625
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":E6BBF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   11160
      X2              =   11160
      Y1              =   3120
      Y2              =   8040
   End
End
Attribute VB_Name = "FrmPresentazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFF
Private Const WAIT_OBJECT_0 = 0
Private Const WAIT_TIMEOUT = &H102

Private Const ABS_AUTOHIDE = &H1
Private Const ABS_ONTOP = &H2
Private Const ABM_GETSTATE = &H4
Private Const ABM_GETTASKBARPOS = &H5

Private Type RECT
        left As Long
        top As Long
        right As Long
        bottom As Long
End Type
Private Type APPBARDATA
    cbSize As Long
    hwnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type


' *** Icon loading functions
Private Const LR_LOADFROMFILE = &H10 ' Not NT
Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2
Private Const IMAGE_ENHMETAFILE = 3


Private Type NOTIFYICONDATA
    cbSize As Long              ' Size of the NotifyIconData structure
    hwnd As Long                ' Window handle of the window processing the icon events
    uID As Long                 ' Icon ID (to allow multiple icons per application)
    uFlags As Long              ' NIF Flags
    uCallbackMessage As Long    ' The message received for the system tray icon if NIF_MESSAGE
                                ' specified. Can be in the range 0x0400 through 0x7FFF (1024 to 32767)
    hIcon As Long               ' The memory location of our icon if NIF_ICON is specifed
    szTip As String * 64        ' Tooltip if NIF_TIP is specified (64 characters max)
End Type

' Shell_NotifyIconA() messages
Private Const NIM_ADD = &H0      ' Add icon to the System Tray
Private Const NIM_MODIFY = &H1   ' Modify System Tray icon
Private Const NIM_DELETE = &H2   ' Delete icon from System Tray

' NotifyIconData Flags
Private Const NIF_MESSAGE = &H1  ' Send event messages to the parent window
Private Const NIF_ICON = &H2     ' Display the icon
Private Const NIF_TIP = &H4      ' Use a tooltip
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10


Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

'icone
Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_GUID = &H5
Private Const NIIF_ICON_MASK = &HF
Private Const NIIF_NOSOUND = &H10
Private Const WM_USER = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

' The events sent appear in lParam and are as follows:
Private Const MOUSE_MOVE = 512
Private Const MOUSE_LEFT_DOWN = 513
Private Const MOUSE_LEFT_UP = 514
Private Const MOUSE_LEFT_DBLCLICK = 515
Private Const MOUSE_RIGHT_DOWN = 516
Private Const MOUSE_RIGHT_UP = 517
Private Const MOUSE_RIGHT_DBLCLICK = 518
Private Const MOUSE_MIDDLE_DOWN = 519
Private Const MOUSE_MIDDLE_UP = 520
Private Const MOUSE_MIDDLE_DBLCLICK = 521

'''SetWindowsPos
Private Const HWND_TOP As Long = 0
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Private Const GWL_WNDPROC = -4

Private OldWindowProc As Long


Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const CSIDL_COMMON_APPDATA = &H1C
Private Const MAX_PATH = 260
Const SW_SHOWNORMAL As Long = 1
Const SW_HIDE As Long = 0
Private Sub CopiaFile(Origine As String, Destinazione As String)
  Dim h As Long
  Rem Copio il file autoexec.bat dal drive C:\ al drive D:\
  Rem Notare che grazie al parametro SW_HIDE la finestra del prompt resta nascosta!
  h = ShellExecute(Me.hwnd, "open", "XCopy.exe", Origine & " " & Destinazione, vbNullString, SW_HIDE)
  
    
End Sub
Public Function TrovaCartella(IDLCartella As Long) As String

    TrovaCartella = String$(MAX_PATH, 0)
    
    Call SHGetSpecialFolderPath(ByVal 0&, TrovaCartella, IDLCartella, ByVal 0&)
    
    TrovaCartella = left$(TrovaCartella, InStr(1, TrovaCartella, Chr$(0)) - 1)
    
    If Len(TrovaCartella) > 0 And right$(TrovaCartella, 1) <> "\" Then TrovaCartella = TrovaCartella & "\"
    
End Function
Private Sub cboAzienda_Click()
    'Filiale
    With Me.cboFiliale
        Set .Database = CnDMT
        .AddFieldKey "IDFiliale"
        .DisplayField = "Filiale"
        .Sql = "SELECT Filiale, IDFiliale, Azienda.IDAzienda "
        .Sql = .Sql & "FROM Filiale INNER JOIN "
        .Sql = .Sql & "AttivitaAzienda ON Filiale.IDAttivitaAzienda = AttivitaAzienda.IDAttivitaAzienda INNER JOIN "
        .Sql = .Sql & "Azienda ON AttivitaAzienda.IDAzienda = Azienda.IDAzienda "
        .Sql = .Sql & "WHERE Azienda.IDAzienda=" & Me.cboAzienda.CurrentID
        .Fill
    End With
End Sub
Private Sub cmdAggiornamento1_Click()
On Error GoTo ERR_cmdAggiornamento1_Click
    LanciaEAspetta App.Path & "\" & cmdAggiornamento1.Tag
Exit Sub
ERR_cmdAggiornamento1_Click:
    MsgBox Err.Description, vbCritical, "cmdAggiornamento1_Click"
End Sub

Private Sub cmdAvvioAutomatico_Click()
On Error GoTo ERR_cmdAvviaPoint_Click
    
    ShowAtStartup "RV_PO51_HD", "HELPDESK"
    
    
    'If CONTROLLO_PROCESSO_ATTIVO("RV_POConsoleIntervento.exe") = False Then
    '    Shell MenuOptions.ProgramsPath & "\RV_POConsoleIntervento.exe", vbNormalFocus
    'End If
    
    
    
    CREA_COLLEGAMENTO_HELP_DESK
    
    MsgBox "PROCEDURA AVVIATA CON SUCCESSO", vbInformation, "LIVE SOLUTION"
    
    
Exit Sub
ERR_cmdAvviaPoint_Click:
    MsgBox Err.Description, vbCritical, "cmdAvvioAutomatico_Click"
End Sub

Private Sub cmdConferma_Click()
    If Me.optScelta(0).Value = True Then
        SaveSetting REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Tipo Avvio", Me.optScelta(0).Index
    Else
        SaveSetting REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Tipo Avvio", Me.optScelta(1).Index
    End If
    
    SaveSetting REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Motore", Me.txtNomeServer.Text
    SaveSetting REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Catalogo", Me.txtNomeCatalogo.Text
    SaveSetting REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "NomeUtente", Me.txtNomeUtente.Text
    SaveSetting REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Password", fnCryptString(Me.txtPassword.Text)
    SaveSetting REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "IDAzienda", Me.cboAzienda.CurrentID
    SaveSetting REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "IDFiliale", Me.cboFiliale.CurrentID
    
    MsgBox "Configurazione avvenuta con successo", vbInformation, "Conferma configurazione"
    
    
End Sub

Private Sub cmdControlla_Click()

    
End Sub

Private Sub cmdSblocca_Click()

    
End Sub

Private Sub cmdTrovaAzienda_Click()
    frmAzienda.Show vbModal
End Sub

Private Sub cmdParametriFiliale_Click()
    frmConfigFiliale.Show
    
End Sub

Private Sub cmdSbloccaUtente_Click()
    frmCoda.Show vbModal
    
End Sub

Private Sub Command1_Click()
On Error GoTo ERR_Command1_Click
    
    DontShowAtStartup "HELPDESK"
    
    MsgBox "Elimazione dell'avvio automatico dell'HELP-DESK avvenuto con successo", vbInformation, "Avvio automatico"
    
Exit Sub
ERR_Command1_Click:
    MsgBox Err.Description, vbCritical, "Eliminazione avvio automatico"
End Sub





Private Sub Command2_Click()
On Error GoTo ERR_Command2_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroRecord As Long


sSQL = "SELECT COUNT(IDRV_POContrattoProdotti) AS NumeroRecordSel "
sSQL = sSQL & "FROM RV_POContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContratto=-1"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rs!NumeroRecordSel)
End If

rs.CloseResultset
Set rs = Nothing

If MsgBox("Record da eliminare: " & NumeroRecord, vbInformation + vbYesNo, "Controllo dati") = vbNo Then Exit Sub

If NumeroRecord = 0 Then Exit Sub

sSQL = "DELETE FROM RV_POContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContratto=-1"
CnDMT.Execute sSQL

MsgBox "OPERAZIONE AVVENUTA CON SUCCESSO", vbInformation, "Controllo dati"

Exit Sub
ERR_Command2_Click:
    MsgBox Err.Description, vbCritical, "Command2_Click"

End Sub

Private Sub Command3_Click()
On Error GoTo ERR_Command3_Click
Dim PercorsoOrigine As String
Dim PercorsoDestinazione As String
Dim PercorsoDestinazioneTmp As String

PercorsoOrigine = MenuOptions.ProgramsPath & "\DOCUMENTAZIONE\it" '& GET_RECUPERA_NOME_PROGRAMMA & "\it"
PercorsoDestinazione = MenuOptions.ProgramsPath & "\it"
PercorsoDestinazioneTmp = TrovaCartella(CSIDL_COMMON_APPDATA)

Set F = New FileSystemObject

If (F.FolderExists(PercorsoDestinazioneTmp & "LiveSolution")) = False Then
    F.CreateFolder (PercorsoDestinazioneTmp & "LiveSolution")
    F.CreateFolder (PercorsoDestinazioneTmp & "LiveSolution\it")
End If

If F.FolderExists(PercorsoOrigine) Then
    
    Set VFolder = F.GetFolder(PercorsoOrigine)
    
    Set VFile = VFolder.Files
    For Each CountFile In VFile
        FileCopy PercorsoOrigine & "\" & CountFile.Name, PercorsoDestinazioneTmp & "LiveSolution\it" & "\" & CountFile.Name
    Next
    
    Set VFolder = F.GetFolder(PercorsoDestinazioneTmp & "LiveSolution\it")
    
    Set VFile = VFolder.Files

    For Each CountFile In VFile
        FileCopy PercorsoDestinazioneTmp & "LiveSolution\it" & "\" & CountFile.Name, PercorsoDestinazione & "\" & CountFile.Name
    Next
    
End If

MsgBox "Operazione avvenuta con successo", vbInformation, "Copia file"

Exit Sub
ERR_Command3_Click:
    MsgBox Err.Description, vbCritical, "Command3_Click"
    
End Sub
Private Function GET_RECUPERA_NOME_PROGRAMMA() As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT * FROM RV_POProgramma "
sSQL = sSQL & "WHERE IDRV_POProgramma=" & 51

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_RECUPERA_NOME_PROGRAMMA = fnNotNull(rs!Programma)
Else
    GET_RECUPERA_NOME_PROGRAMMA = ""
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_PERCORSO_ORIGINE() As String
Dim F As FileSystemObject
Dim Percorso As String

'Lista Help
'Set f = New FileSystemObject
GET_PERCORSO_ORIGINE = MenuOptions.ProgramsPath & "\Help_"

'If f.FolderExists(Percorso) Then
'    Set VFolder = f.GetFolder(Percorso)
'
'    Set VFile = VFolder.Files
'
'    For Each CountFile In VFile
'        If right(CountFile.Name, 3) = "pdf" Then
' '           Me.LV.ListItems.Add , "H_" & CountFile.Name, CountFile.Name, 2, 2
'        End If
'    Next
'End If
End Function

Private Sub Command4_Click()
    frmControlloDoc.Show vbModal
    
End Sub

Private Sub Form_Load()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If ConnessioneADODBLib = True Then
    PrelevaAzienda
    
    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    Me.optScelta(TipoAvvio) = True


    'Azienda
    With Me.cboAzienda
        Set .Database = CnDMT
        .AddFieldKey "IDAzienda"
        .DisplayField = "AnagraficaAnagrafica"
        .Sql = "SELECT * FROM RepAzienda"
        .Fill
    End With
    
    'Filiale
    With Me.cboFiliale
        Set .Database = CnDMT
        .AddFieldKey "IDFiliale"
        .DisplayField = "Filiale"
        .Sql = "SELECT Filiale, IDFiliale, Azienda.IDAzienda "
        .Sql = .Sql & "FROM Filiale INNER JOIN "
        .Sql = .Sql & "AttivitaAzienda ON Filiale.IDAttivitaAzienda = AttivitaAzienda.IDAttivitaAzienda INNER JOIN "
        .Sql = .Sql & "Azienda ON AttivitaAzienda.IDAzienda = Azienda.IDAzienda "
        .Fill
    End With
    
    ''''''''''''''' DATI PERSONALIZZATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Me.txtNomeServer.Text = GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Motore")
    Me.txtNomeCatalogo.Text = GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Catalogo")
    Me.txtNomeUtente.Text = GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "NomeUtente")
    Me.txtPassword.Text = GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Password")
    Me.cboAzienda.WriteOn fnNotNullN(GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "IDAzienda"))
    Me.cboFiliale.WriteOn fnNotNullN(GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "IDFiliale"))

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'PrelevaAzienda
    
    sSQL = "SELECT * FROM RV_PORelease WHERE IDRV_POProgramma=" & IdentificativoProgramma
    sSQL = sSQL & " ORDER BY IDRelease DESC"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        Me.Caption = "Modulo " & NomeProgramma & " (Release " & Trim(rs!Release) & ")"
        Me.lblVersione.Caption = Trim(rs!Release)
        Me.lblDataInstallazione.Caption = fnNotNull(rs!DataInstallazione)
    Else
        Me.Caption = "Modulo " & NomeProgramma
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    GET_RECUPERA_DATI_LICENZA
    CONTROLLO_LICENZA
    GET_GRIGLIA 51
    
End If

End Sub

Private Sub GET_RECUPERA_DATI_LICENZA()

End Sub
Private Sub CONTROLLO_LICENZA()


End Sub
Private Function GET_CODICE_DIAMANTE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Descrizione FROM ComponenteSwAbilitata "
sSQL = sSQL & "WHERE NomeCompSW=" & fnNormString("*IDSW___")

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICE_DIAMANTE = ""
Else
    GET_CODICE_DIAMANTE = Trim(fnNotNull(rs!Descrizione))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub SALVATAGGIO_DATI()

End Sub

Private Sub optScelta_Click(Index As Integer)
    'If Index = 0 Then
    '    Me.FraDatabase.Enabled = False
    'Else
    '    Me.FraDatabase.Enabled = True
    'End If
End Sub
Private Sub CREA_COLLEGAMENTO_HELP_DESK()
On Error GoTo ERR_cmdCreaCollegamento_Click
Dim F As FileSystemObject
Dim strDesktop As String
Dim WshShell
Dim oMyShotcut
Dim PercorsoAssouluto As String
Dim X As Long

Set F = New FileSystemObject

If F.FileExists(MenuOptions.ProgramsPath & "\RV_POCreaCollegamento.exe.manifest") Then
    F.DeleteFile MenuOptions.ProgramsPath & "\RV_POCreaCollegamento.exe.manifest", True
End If

Set F = Nothing


X = Shell(MenuOptions.ProgramsPath & "\RV_POCreaCollegamento.exe", vbNormalFocus)

Exit Sub

ERR_cmdCreaCollegamento_Click:
    MsgBox Err.Description
    
    MsgBox "Collegamento al desktop non avvenuto", vbCritical, "Collegamento"
End Sub
Public Sub LanciaEAspetta(nomeProcesso As String)

Dim lPid As Long
Dim lHnd As Long
Dim lRet As Long

If Trim$(nomeProcesso) = "" Then Exit Sub

lPid = Shell(nomeProcesso, vbNormalFocus)
If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle lHnd
        End If
        ''MsgBox "Finito!.", vbInformation
End If

End Sub

Private Sub GET_GRIGLIA(IdentificativoProgramma As Long)
On Error GoTo ERR_fnGrigliaAssegnazione
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
    sSQL = "SELECT * FROM RV_POProgrammaModulo "
    sSQL = sSQL & "WHERE IdentificazioneProgramma=" & IdentificativoProgramma
    
    
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, CnDMT.InternalConnection
        
        With Me.Griglia
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
                .ColumnsHeader.Add "IDRV_POProgrammaModulo", "IDRV_POProgrammaModulo", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "DescrizioneModulo", "Modulo", dgchar, True, 3500, dgAlignleft
                .ColumnsHeader.Add "Attivato", "Attivato", dgBoolean, True, 1500, dgAligncenter
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaAssegnazione:
    MsgBox "Errore recupero licenza", vbCritical, "Funzionalità abilitate"
End Sub
