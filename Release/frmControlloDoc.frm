VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Begin VB.Form frmControlloDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ELIMINAZIONE FLUSSO DOCUMENTI"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmControlloDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdElimina 
      Caption         =   "ELIMINA FLUSSO"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3960
      Width           =   2775
   End
   Begin DMTDATETIMELib.dmtDate txtDataDoc 
      Height          =   315
      Left            =   7560
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   253
      BackColor       =   16777215
      Appearance      =   1
   End
   Begin DMTEDITNUMLib.dmtNumber txtNumeroDoc 
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   253
      Text            =   "0"
      BackColor       =   16777215
      Appearance      =   1
      AllowEmpty      =   0   'False
   End
   Begin DMTDataCmb.DMTCombo cboTipoOggetto 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   3255
      _ExtentX        =   5741
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
   Begin DMTDataCmb.DMTCombo cboSezionale 
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.Label Label2 
      Caption         =   "Data doc."
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "N° doc."
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2295
      Left            =   0
      Picture         =   "frmControlloDoc.frx":4781A
      Top             =   0
      Width           =   9000
   End
   Begin VB.Label Label2 
      Caption         =   "Sezionale"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   6
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo documento"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   3255
   End
End
Attribute VB_Name = "frmControlloDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboTipoOggetto_Click()
On Error GoTo ERR_cboTipoOggetto_Click
    With Me.cboSezionale
        Set .Database = CnDMT
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .Sql = "SELECT Sezionale.IDSezionale, Sezionale.Sezionale "
        .Sql = .Sql & "FROM RegistroIvaPerTipoOggetto INNER JOIN "
        .Sql = .Sql & "Sezionale ON RegistroIvaPerTipoOggetto.IDRegistroIva = Sezionale.IDRegistroIva AND "
        .Sql = .Sql & "RegistroIvaPerTipoOggetto.IDFiliale = Sezionale.IDFiliale LEFT OUTER JOIN "
        .Sql = .Sql & "TipoOggetto ON RegistroIvaPerTipoOggetto.IDTipoOggetto = TipoOggetto.IDTipoOggetto "
        .Sql = .Sql & "WHERE RegistroIvaPerTipoOggetto.IDTipoOggetto = " & Me.cboTipoOggetto.CurrentID
        .Sql = .Sql & " AND RegistroIvaPerTipoOggetto.IDFiliale = " & VarIDFiliale
        .Fill
    End With
Exit Sub
ERR_cboTipoOggetto_Click:
    MsgBox Err.Description, vbCritical, "cboTipoOggetto_Click"

End Sub

Private Sub cmdElimina_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDTipoOggettoRataContratto As Long
Dim IDTipoOggettoRilContatori As Long
Dim IDOggettoDocumento As Long

IDOggettoDocumento = 0

If Me.cboTipoOggetto.CurrentID = 0 Then Exit Sub
If Me.cboSezionale.CurrentID = 0 Then Exit Sub
If txtNumeroDoc.Value = 0 Then Exit Sub
If txtDataDoc.Value = 0 Then Exit Sub



sSQL = "SELECT IDOggetto FROM Oggetto "
sSQL = sSQL & "WHERE IDTipoOggetto=" & Me.cboTipoOggetto.CurrentID
sSQL = sSQL & " AND IDSezionale=" & Me.cboSezionale.CurrentID
sSQL = sSQL & " AND Numero=" & fnNormString(Me.txtNumeroDoc.Text)
sSQL = sSQL & " AND DataEmissione=" & fnNormString(Me.txtDataDoc.Text)

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    IDOggettoDocumento = fnNotNullN(rs!IDOggetto)
End If

If IDOggettoDocumento = 0 Then
    MsgBox "Non è stato possibile recuperare il documento", vbInformation, "Eliminazione flusso documentale"
    Exit Sub
End If

IDTipoOggettoRataContratto = fnGetTipoOggetto("RV_PORateContratto")
IDTipoOggettoRilContatori = fnGetTipoOggetto("RV_POContatoreRilevamenti")

sSQL = "DELETE FROM FlussoOggettiCollegati "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoDocumento
sSQL = sSQL & " AND IDTipoOggetto=" & Me.cboTipoOggetto.CurrentID
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoRataContratto
CnDMT.Execute sSQL

sSQL = "DELETE FROM FlussoOggettiCollegati "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoDocumento
sSQL = sSQL & " AND IDTipoOggetto=" & Me.cboTipoOggetto.CurrentID
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoRilContatori
CnDMT.Execute sSQL


sSQL = "DELETE FROM FlussoFunzioneCollegato "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoDocumento
sSQL = sSQL & " AND IDTipoOggetto=" & Me.cboTipoOggetto.CurrentID
sSQL = sSQL & " AND IDFlussoFunzione >= 10000"
CnDMT.Execute sSQL



MsgBox "OPERAZIONE COMPLETATA!", vbInformation, "Eliminazione flusso documentale"

End Sub

Private Sub Form_Load()
    initControlli
End Sub
Private Sub initControlli()
On Error GoTo ERR_initControlli
    Dim sSQL As String
    
    sSQL = "SELECT IDTipoOggetto, Oggetto"
    sSQL = sSQL & " FROM TipoOggetto"
    sSQL = sSQL & " WHERE IDGestore=15"
    sSQL = sSQL & " ORDER BY Oggetto"
    With Me.cboTipoOggetto
        Set .Database = CnDMT
        .DisplayField = "Oggetto"
        .AddFieldKey "IDTipoOggetto"
        .Sql = sSQL
        .Refresh
    End With
    
Exit Sub
ERR_initControlli:
    MsgBox Err.Description, vbCritical, "initControlli"
End Sub

Private Function fnGetTipoOggetto(Optional Gestore As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    If Gestore = "" Then
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(App.EXEName)
    Else
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(Gestore)
    End If
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If Not rs.EOF Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

