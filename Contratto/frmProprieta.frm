VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmProprieta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestione delle proprietà del contatto"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7980
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
   ScaleHeight     =   5805
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DMTEDITNUMLib.dmtNumber txtIDProprieta 
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      Top             =   960
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   253
      Text            =   "0"
      BackColor       =   16777215
      Appearance      =   1
      AllowEmpty      =   0   'False
   End
   Begin VB.CommandButton cmdNuovoGruppo 
      Height          =   315
      Left            =   7440
      Picture         =   "frmProprieta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Nuovo gruppo"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdModificaGruppo 
      Height          =   315
      Left            =   6960
      Picture         =   "frmProprieta.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Modifica gruppo"
      Top             =   360
      Width           =   495
   End
   Begin VB.CheckBox chkVisualizza 
      Caption         =   "Visualizza"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CheckBox chkGestisci 
      Caption         =   "Gestisci"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   3735
   End
   Begin DMTDataCmb.DMTCombo cboGruppo 
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
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
   Begin VB.TextBox txtProprieta 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalva 
      Caption         =   "Salva"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdNuovo 
      Caption         =   "Nuovo"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7011
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
   Begin VB.Label Label1 
      Caption         =   "Proprietà"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblGruppo 
      Caption         =   "Gruppo"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmProprieta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Nuovo As Boolean
Private rsGriglia As ADODB.Recordset


Private Sub cmdElimina_Click()
Dim sSQL As String
Dim Testo As String

Testo = "Sei sicuro di voler eliminare questa proprietà?"


'CONTROLLO SULL'ELIMINAZIONE
If Nuovo = True Then Exit Sub

If ((Me.GrigliaCorpo.Recordset.EOF) And (Me.GrigliaCorpo.Recordset.BOF)) Then Exit Sub

If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione proprietà") = vbNo Then Exit Sub

If GET_CONTROLLO_UTILIZZO(Me.txtIDProprieta.Value) = True Then
    MsgBox "La proprietà è già stata utilizzata, quindi è impossibile eliminare", vbCritical, "Eliminazione proprietà"
    Exit Sub
End If
sSQL = "DELETE FROM RV_POContactProprieta "
sSQL = sSQL & "WHERE IDRV_POContactProprieta=" & Me.txtIDProprieta.Value
Cn.Execute sSQL

GET_GRIGLIA

End Sub

Private Sub cmdModificaGruppo_Click()
Dim LINK_GRUPPO_LOCAL As Long
Dim NumeroRecord As Long
    LINK_GRUPPO = Me.cboGruppo.CurrentID
    
    frmGruppo.Show vbModal
    
    If LINK_GRUPPO > 0 Then
        LINK_GRUPPO_LOCAL = Me.cboGruppo.CurrentID
        
        Me.cboGruppo.Refresh
        If LINK_GRUPPO_LOCAL = 0 Then
            Me.cboGruppo.WriteOn LINK_GRUPPO
        Else
            Me.cboGruppo.WriteOn LINK_GRUPPO_LOCAL
        End If
        
    End If
    
    If Me.txtIDProprieta.Value > 0 Then
        NumeroRecord = Me.GrigliaCorpo.ListIndex - 1
        GET_GRIGLIA
        Me.GrigliaCorpo.Recordset.Move NumeroRecord
    End If
    
End Sub

Private Sub cmdNuovo_Click()
    Me.txtIDProprieta.Value = 0
    Me.txtProprieta.Text = ""
    Me.cboGruppo.WriteOn 0
    Me.chkGestisci.Value = vbChecked
    Me.chkVisualizza.Value = vbChecked
    Me.txtProprieta.SetFocus
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA_TOUR
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
    sSQL = "SELECT * FROM RV_POIEContactProprieta "
    sSQL = sSQL & "ORDER BY ContactGruppo,Proprieta"
    
    Set rsGriglia = New ADODB.Recordset
        rsGriglia.Open sSQL, Cn.InternalConnection
        
        With Me.GrigliaCorpo
            .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDRV_POContactProprieta", "IDRV_POContactProprieta", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POContactGruppo", "IDRV_POContactGruppo", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "ContactGruppo", "Gruppo", dgchar, True, 1500, dgAlignleft
            .ColumnsHeader.Add "Proprieta", "Proprieta", dgchar, True, 2500, dgAlignleft
            .ColumnsHeader.Add "Gestisci", "Gestisci", dgBoolean, True, 1500, dgAligncenter
            .ColumnsHeader.Add "Visualizza", "Visualizza", dgBoolean, True, 1500, dgAligncenter
            
            Set .Recordset = rsGriglia
            .Refresh
            .LoadUserSettings
        End With
    
    Cn.CursorLocation = OLDCursor
    
Exit Sub
ERR_GET_GRIGLIA_TOUR:
    MsgBox Err.Description, vbCritical, "Reperimento dati tour planning"
End Sub

Private Sub cmdNuovoGruppo_Click()
Dim LINK_GRUPPO_LOCAL As Long
Dim NumeroRecord As Long

    LINK_GRUPPO = 0
    
    frmGruppo.Show vbModal
    
    If LINK_GRUPPO > 0 Then
        LINK_GRUPPO_LOCAL = Me.cboGruppo.CurrentID
        
        Me.cboGruppo.Refresh
        If LINK_GRUPPO_LOCAL = 0 Then
            Me.cboGruppo.WriteOn LINK_GRUPPO
        Else
            Me.cboGruppo.WriteOn LINK_GRUPPO_LOCAL
        End If
        
    End If
    
    If Me.txtIDProprieta.Value > 0 Then
        NumeroRecord = Me.GrigliaCorpo.ListIndex - 1
        GET_GRIGLIA
        Me.GrigliaCorpo.Recordset.Move NumeroRecord
    End If
End Sub

Private Sub cmdSalva_Click()
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim NumeroRecord As Long
Dim NomeCampoProp As String

NomeCampoProp = GET_NOME_CAMPO_PROPRIETA(Me.txtProprieta.Text)

If Len(Trim(Me.txtProprieta.Text)) = 0 Then
    MsgBox "Inserire la proprietà", vbCritical, "Gestione proprieta"
    Exit Sub
End If
If Me.cboGruppo.CurrentID = 0 Then
    MsgBox "Inserire il gruppo della proprietà", vbCritical, "Gestione proprieta"
    Exit Sub
End If
If GET_ESISTENZA_DESCRIZIONE(Me.txtProprieta.Text, Me.txtIDProprieta.Value) = True Then
    MsgBox "La proprietà inserita è già esistente", vbCritical, "Gestione proprieta"
    Exit Sub
End If
If GET_ESISTENZA_CAMPO_PROP(NomeCampoProp, Me.txtIDProprieta.Value) = True Then
    MsgBox "Il nome del campo è già stato inserito", vbCritical, "Gestione proprieta"
    Exit Sub
End If


sSQL = "SELECT * FROM RV_POContactProprieta "
sSQL = sSQL & "WHERE IDRV_POContactProprieta=" & Me.txtIDProprieta.Value

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
    rs!IDRV_POContactProprieta = fnGetNewKey("RV_POContactProprieta", "IDRV_POContactProprieta")
End If

rs!Proprieta = Me.txtProprieta.Text
rs!IDRV_POContactGruppo = Me.cboGruppo.CurrentID
rs!Gestisci = Abs(Me.chkGestisci.Value)
rs!Visualizza = Abs(Me.chkVisualizza.Value)
rs!nomeCampo = NomeCampoProp
rs.Update

rs.Close
Set rs = Nothing

If Me.txtIDProprieta.Value = 0 Then
    NumeroRecord = Me.GrigliaCorpo.ListCount
Else
    NumeroRecord = Me.GrigliaCorpo.ListIndex - 1
End If

GET_GRIGLIA

Me.GrigliaCorpo.Recordset.Move NumeroRecord

End Sub

Private Function GET_ESISTENZA_DESCRIZIONE(Descrizione As String, IDProprieta As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POContactProprieta "
sSQL = sSQL & "WHERE Proprieta=" & fnNormString(Descrizione)
If IDProprieta > 0 Then
    sSQL = sSQL & " AND IDRV_POContactProprieta<>" & IDProprieta
End If
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_DESCRIZIONE = False
Else
    GET_ESISTENZA_DESCRIZIONE = True
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_ESISTENZA_CAMPO_PROP(CampoProp As String, IDProprieta As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POContactProprieta "
sSQL = sSQL & "WHERE NomeCampo=" & fnNormString(CampoProp)
If IDProprieta > 0 Then
    sSQL = sSQL & " AND IDRV_POContactProprieta<>" & IDProprieta
End If

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_CAMPO_PROP = False
Else
    GET_ESISTENZA_CAMPO_PROP = True
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub Form_Load()
    With Me.cboGruppo
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POContactGruppo"
        .DisplayField = "ContactGruppo"
        .SQL = "SELECT * FROM RV_POContactGruppo ORDER BY ContactGruppo"
        .Fill
    End With
    
    
    
    GET_GRIGLIA
End Sub

Private Sub GrigliaCorpo_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    Me.txtIDProprieta.Value = fnNotNullN(AllColumns("IDRV_POContactProprieta").Value)
    Me.cboGruppo.WriteOn fnNotNullN(AllColumns("IDRV_POContactGruppo").Value)
    Me.txtProprieta.Text = fnNotNull(AllColumns("Proprieta").Value)
    Me.chkGestisci.Value = fnNotNullN(AllColumns("Gestisci").Value)
    Me.chkVisualizza.Value = fnNotNullN(AllColumns("Visualizza").Value)
End Sub
Private Function GET_CONTROLLO_UTILIZZO(IDProprieta As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POContactValoriProp "
sSQL = sSQL & "WHERE IDRV_POContactProprieta=" & IDProprieta
sSQL = sSQL & " AND LEN(Valore)>0 "

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_UTILIZZO = False
Else
    GET_CONTROLLO_UTILIZZO = True
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_NOME_CAMPO_PROPRIETA(Nominativo As String) As String
Dim I As Long
Dim ReturnString As String

ReturnString = ""
For I = 1 To Len(Nominativo)
    If Mid(Nominativo, I, 1) <> " " Then
        ReturnString = ReturnString & Mid(Nominativo, I, 1)
    End If
Next

ReturnString = Replace(ReturnString, "-", "")
ReturnString = Replace(ReturnString, "/", "")
ReturnString = Replace(ReturnString, "\", "")
ReturnString = Replace(ReturnString, "(", "")
ReturnString = Replace(ReturnString, ")", "")
ReturnString = Replace(ReturnString, "%", "")
ReturnString = Replace(ReturnString, "$", "")
ReturnString = Replace(ReturnString, "#", "")
ReturnString = Replace(ReturnString, "@", "")
ReturnString = Replace(ReturnString, "é", "e")
ReturnString = Replace(ReturnString, "ò", "o")
ReturnString = Replace(ReturnString, "à", "à")
ReturnString = Replace(ReturnString, "ù", "u")
ReturnString = Replace(ReturnString, "ì", "ì")
ReturnString = Replace(ReturnString, "è", "e")
ReturnString = Replace(ReturnString, ".", "")
ReturnString = Replace(ReturnString, ":", "")
ReturnString = Replace(ReturnString, "=", "")
ReturnString = Replace(ReturnString, "?", "")
ReturnString = Replace(ReturnString, "^", "")
ReturnString = Replace(ReturnString, "<", "")
ReturnString = Replace(ReturnString, ">", "")
ReturnString = Replace(ReturnString, "|", "")
ReturnString = Replace(ReturnString, "!", "")
ReturnString = Replace(ReturnString, Chr(34), "")
ReturnString = Replace(ReturnString, "£", "")
ReturnString = Replace(ReturnString, "€", "")
ReturnString = Replace(ReturnString, "ç", "")
ReturnString = Replace(ReturnString, "[", "")
ReturnString = Replace(ReturnString, "]", "")
ReturnString = Replace(ReturnString, "+", "")
ReturnString = Replace(ReturnString, "*", "")
ReturnString = Replace(ReturnString, "'", "")
ReturnString = Replace(ReturnString, ";", "")

GET_NOME_CAMPO_PROPRIETA = ReturnString

End Function


