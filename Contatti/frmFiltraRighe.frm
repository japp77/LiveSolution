VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Begin VB.Form frmFiltraRighe 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "FILTRA CONTATTI"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5025
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
   ScaleHeight     =   2025
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CheckBox chkDestNonAssociata 
         Alignment       =   1  'Right Justify
         Caption         =   "Non associata"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdConferma 
         Caption         =   "Conferma"
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtNominativo 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4575
      End
      Begin DMTDataCmb.DMTCombo cboFiliale 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   980
         Width           =   4575
         _ExtentX        =   8070
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
      Begin VB.Label Label1 
         Caption         =   "Altra sede"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Nominativo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmFiltraRighe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboFiliale_Click()
    If Me.cboFiliale.CurrentID > 0 Then
        Me.chkDestNonAssociata.Enabled = False
        Me.chkDestNonAssociata.Value = 0
    Else
        Me.chkDestNonAssociata.Enabled = True
    End If
End Sub

Private Sub chkDestNonAssociata_Click()
    If Me.chkDestNonAssociata.Value = vbChecked Then
        Me.cboFiliale.WriteOn 0
        Me.cboFiliale.Enabled = False
    Else
        Me.cboFiliale.Enabled = True
    End If
End Sub

Private Sub cmdConferma_Click()
    LINK_DESTINAZIONE_TROVA = Me.cboFiliale.CurrentID
    NOMINATIVO_TROVA = Me.txtNominativo.Text
    DESTINAZIONE_NON_ASSOCIATA_TROVA = Me.chkDestNonAssociata.Value
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    If KeyCode = vbKeyReturn Then
        cmdConferma_Click
    End If
End Sub

Private Sub Form_Load()
    With Me.cboFiliale
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT IDSitoPerAnagrafica, SitoPerAnagrafica FROM SitoPerAnagrafica "
        .SQL = .SQL & "WHERE IDAnagrafica=" & frmMain.CDCliente.KeyFieldID
        .Fill
    End With
    
    Me.cboFiliale.WriteOn LINK_DESTINAZIONE_TROVA
    Me.txtNominativo.Text = NOMINATIVO_TROVA
    Me.chkDestNonAssociata.Value = Abs(DESTINAZIONE_NON_ASSOCIATA_TROVA)
End Sub
