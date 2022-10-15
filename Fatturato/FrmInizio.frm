VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Begin VB.Form FrmInizio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fatturato (Passo 1 di 2)"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   Icon            =   "FrmInizio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin DMTDataCmb.DMTCombo cboAMese 
      Height          =   315
      Left            =   4920
      TabIndex        =   17
      Top             =   4080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DMTDataCmb.DMTCombo cboDaMese 
      Height          =   315
      Left            =   4920
      TabIndex        =   16
      Top             =   3600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   0
      Picture         =   "FrmInizio.frx":030A
      ScaleHeight     =   5025
      ScaleWidth      =   2385
      TabIndex        =   13
      Top             =   0
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "RIFERIMENTO AZIENDA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3375
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   4455
      Begin VB.Label Label1 
         Caption         =   "Azienda"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Attivita Azienda"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Filiale"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Esercizio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label LblAzienda 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label LblAttivitaAzienda 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label LblFiliale 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label LblEsercizio 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "Utente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   4095
      End
      Begin VB.Label LblUtente 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   3000
         Width           =   4215
      End
   End
   Begin VB.CommandButton CmdFine 
      Caption         =   "Annulla"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton CmdAvanti 
      Caption         =   "Avanti"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "A mese"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   15
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Da mese"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   14
      Top             =   3600
      Width           =   975
   End
End
Attribute VB_Name = "FrmInizio"
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

Private Sub cboAMese_Click()
    VAR_A_MESE = Me.cboAMese.CurrentID
End Sub

Private Sub cboDaMese_Click()
    VAR_DA_MESE = Me.cboDaMese.CurrentID
End Sub

Private Sub cmdAvanti_Click()
    VAR_DA_MESE = Me.cboDaMese.CurrentID
    VAR_A_MESE = Me.cboAMese.CurrentID
    Unload Me
End Sub

Private Sub CmdFine_Click()
Dim Risposta As Integer
    Risposta = MsgBox("Vuoi abbandonare il wizard per il passaggio degli interventi in contabilità?", vbInformation + vbYesNo, "Abbandono")
        
    If Risposta = vbYes Then
        Unload Me
    End If

End Sub


Public Function InitControlli()


    fncMese

    Me.cboAMese.WriteOn DatePart("m", Date)
    Me.cboDaMese.WriteOn DatePart("m", Date)

End Function
Private Sub Form_Load()
On Error GoTo ERR_Form_Load
   
If b_Loading = True Then
    fncMese

    Me.cboAMese.WriteOn DatePart("m", Date)
    Me.cboDaMese.WriteOn DatePart("m", Date)
    PrelevaAzienda
End If

    
Exit Sub

ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERR_Form_Unload
Dim Risposta As Integer

    If Me.CmdAvanti.Value = True Then
        FrmMovimenti.Show
       
    End If

Exit Sub

ERR_Form_Unload:
    MsgBox Err.Description, vbCritical, "Form_Unload"
End Sub
Private Sub fncMese()
    With Me.cboAMese
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POMese"
        .DisplayField = "Mese"
        .Sql = "SELECT * FROM RV_POMese ORDER BY IDRV_POMese"
        .Fill
    End With
    
    With Me.cboDaMese
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POMese"
        .DisplayField = "Mese"
        .Sql = "SELECT * FROM RV_POMese ORDER BY IDRV_POMese"
        .Fill
    End With
    
End Sub

