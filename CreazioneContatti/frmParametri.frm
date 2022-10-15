VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParametri 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CREAZIONE CONTATTI   (Passo 2 di 3)"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
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
   ScaleHeight     =   4530
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstProprieta 
      Height          =   1635
      Left            =   3120
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   1080
      Width           =   3135
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   3120
      TabIndex        =   7
      Top             =   3840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin DMTDataCmb.DMTCombo cboTipoContratto 
      Height          =   315
      Left            =   3120
      TabIndex        =   6
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
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
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdIIndietro 
      Caption         =   "Indietro"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Avanti"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdFine 
      Caption         =   "Fine"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Proprietà che indicano una Email"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label lblInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   3480
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo contratto"
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmParametri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboTipoCliente_Click()
    optTipoCliente = Me.cboTipoCliente.CurrentID
    
End Sub
Private Sub cboTipoContatto_Click()
    optTipoContatto = Me.cboTipoContatto.CurrentID
End Sub
Private Sub cboTipoContratto_Click()
    optTipoContratto = Me.cboTipoContratto.CurrentID
End Sub

Private Sub cmdAnnulla_Click()
    If MsgBox("Vuoi abbandonare la procedura?", vbInformation + vbYesNo, "Chiusura applicazioni") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdAvanti_Click()

    CreazioneContatti
    Unload Me
End Sub

Private Sub cmdIIndietro_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    DinamicNumber = 2
    ControlloComandi
    
    fncTipoContratto
    GET_PROPRIETA_EMAIL
End Sub
Private Sub ControlloComandi()
    
    Me.cmdFine.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.cmdAvanti.Value = True Then
        frmVisualizzazione.Show
        Exit Sub
    End If
    If Me.cmdIIndietro.Value = True Then
        FrmMain.Show
        Exit Sub
    End If
    
End Sub

Private Sub fncTipoContratto()
    With Me.cboTipoContratto
        Set .Database = CnDMT
        .DisplayField = "TipoContratto"
        .AddFieldKey "IDRV_POTipoContratto"
        .Sql = "Select * From RV_POTipoContratto"
        .Refresh
    End With
End Sub
Private Sub GET_PROPRIETA_EMAIL()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim I As Integer
Me.lstProprieta.Clear

sSQL = "SELECT RV_POContactProprieta.IDRV_POContactProprieta, RV_POContactProprieta.IDRV_POContactGruppo, RV_POContactProprieta.Proprieta, "
sSQL = sSQL & "RV_POContactGruppo.Email "
sSQL = sSQL & "FROM RV_POContactProprieta INNER JOIN "
sSQL = sSQL & "RV_POContactGruppo ON RV_POContactProprieta.IDRV_POContactGruppo = RV_POContactGruppo.IDRV_POContactGruppo "
sSQL = sSQL & "WHERE RV_POContactGruppo.Email=1"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    Me.lstProprieta.AddItem fnNotNull(rs!Proprieta)
    I = Me.lstProprieta.NewIndex
    Me.lstProprieta.ItemData(I) = fnNotNullN(rs!IDRV_POContactProprieta)
    Me.lstProprieta.Selected(I) = True
 
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
