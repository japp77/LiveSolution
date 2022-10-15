VERSION 5.00
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmInserimento 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Inserimento dati"
   ClientHeight    =   1140
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6600
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
   ScaleHeight     =   1140
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "A capo"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   750
      Width           =   3015
   End
   Begin VB.CommandButton cmdSalva 
      Height          =   375
      Left            =   6000
      Picture         =   "frmInserimento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Carattere di spazio"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin DMTEDITNUMLib.dmtNumber txtPosizione 
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Top             =   360
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   503
      _StockProps     =   253
      BackColor       =   16777215
      Appearance      =   1
   End
   Begin VB.TextBox txtDescrizione 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Posizione"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblDescrizione 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmInserimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LINK_STRINGA As Long

Private Sub cmdSalva_Click()
    If SALVATAGGIO = True Then Unload Me
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If SALVATAGGIO = True Then Unload Me
    End If
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    If TIPO_INSERIMENTO = 1 Then
        Me.txtDescrizione.Locked = True
        Me.txtDescrizione.Text = FrmInizio.lstCampi.Text
    Else
        Me.txtDescrizione.Locked = False
        Me.txtDescrizione.Text = ""
    End If
    
    Me.txtPosizione.Value = GET_POSIZIONE
    
End Sub
Private Function GET_POSIZIONE()
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT MAX(Posizione) as MaxPos "
sSQL = sSQL & "FROM RV_POStringaEmail "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDRV_POTipoSezioneEmail=" & FrmInizio.cboTipoSezioneEmail.CurrentID


Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    GET_POSIZIONE = 1
Else
    GET_POSIZIONE = fnNotNullN(rs!MaxPos) + 1
End If

rs.Close
Set rs = Nothing
End Function
Private Sub RICALCOLO_POSIZIONI(IDRigaStringa As Long, Posizione As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsUpd As ADODB.Recordset

sSQL = "SELECT * FROM RV_POStringaEmail "
sSQL = sSQL & "WHERE IDRV_POStringaEmail<>" & IDRigaStringa
sSQL = sSQL & " AND Posizione=" & Posizione
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDRV_POTipoSezioneEmail=" & FrmInizio.cboTipoSezioneEmail.CurrentID
sSQL = sSQL & " ORDER BY Posizione"

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    rs.CloseResultset
    Set rs = Nothing
    
    sSQL = "SELECT * FROM RV_POStringaEmail "
    sSQL = sSQL & "WHERE IDRV_POStringaEmail<>" & IDRigaStringa
    sSQL = sSQL & " AND Posizione>=" & Posizione
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDRV_POTipoSezioneEmail=" & FrmInizio.cboTipoSezioneEmail.CurrentID
    sSQL = sSQL & " ORDER BY Posizione"
    
    Set rsUpd = New ADODB.Recordset
    
    rsUpd.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    While Not rsUpd.EOF
        rsUpd!Posizione = rsUpd!Posizione + 1
        rsUpd.Update
    rsUpd.MoveNext
    Wend
    
    rsUpd.Close
    Set rsUpd = Nothing
    
Else

    rs.CloseResultset
    Set rs = Nothing

End If
End Sub

Private Function SALVATAGGIO() As Boolean
'On Error GoTo ERR_SALVATAGGIO
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POStringaEmail "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDRV_POTipoSezioneEmail=" & FrmInizio.cboTipoSezioneEmail.CurrentID


Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    rs!IDRV_POStringaEmail = fnGetNewKey("RV_POStringaEmail", "IDRV_POStringaEmail")
    rs!IDAzienda = TheApp.IDFirm
    rs!IDRV_POTipoSezioneEmail = FrmInizio.cboTipoSezioneEmail.CurrentID
    If TIPO_INSERIMENTO = 1 Then
        rs!NomeCampo = Me.txtDescrizione.Text
    Else
        rs!ValoreCampo = Me.txtDescrizione.Text
    End If
    rs!Posizione = Me.txtPosizione.Value
    rs!CarattereSpazio = Me.Check1.Value
    rs!CarattereACapo = Me.Check2.Value
    
    
rs.Update

LINK_STRINGA = rs!IDRV_POStringaEmail


rs.Close
Set rs = Nothing

RICALCOLO_POSIZIONI LINK_STRINGA, Me.txtPosizione.Value
SALVATAGGIO = True
Exit Function
ERR_SALVATAGGIO:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    SALVATAGGIO = False

End Function
