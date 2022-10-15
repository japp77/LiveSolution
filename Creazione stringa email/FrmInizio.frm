VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Begin VB.Form FrmInizio 
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14250
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInizio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "FrmInizio.frx":030A
   ScaleHeight     =   8190
   ScaleWidth      =   14250
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   615
      Left            =   705
      TabIndex        =   2
      Top             =   120
      Width           =   270
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8145
      ScaleWidth      =   14145
      TabIndex        =   0
      Top             =   0
      Width           =   14175
      Begin DMTDataCmb.DMTCombo cboTipoSezioneEmail 
         Height          =   315
         Left            =   6840
         TabIndex        =   11
         Top             =   360
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
      Begin VB.CommandButton cmdGiu 
         Height          =   375
         Left            =   6360
         Picture         =   "FrmInizio.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Giù"
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton cmdSu 
         Height          =   375
         Left            =   6360
         Picture         =   "FrmInizio.frx":0B9E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Su"
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdElimina 
         Height          =   375
         Left            =   6360
         Picture         =   "FrmInizio.frx":1128
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Elimina"
         Top             =   3000
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   1815
         Left            =   6840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   960
         Width           =   7215
      End
      Begin DmtGridCtl.DmtGrid Griglia 
         Height          =   5175
         Left            =   6840
         TabIndex        =   5
         Top             =   2880
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   9128
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
      Begin VB.ListBox lstCampi 
         Height          =   7860
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label2 
         Caption         =   "Sezione del messaggio da inviare"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   6840
         TabIndex        =   12
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Descrizione"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   6840
         TabIndex        =   7
         Top             =   720
         Width           =   7215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Campi da utilizzare"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   6135
      End
   End
End
Attribute VB_Name = "FrmInizio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents TheApp As DMTRunAppLib.Application
Attribute TheApp.VB_VarHelpID = -1
Private rsGriglia As ADODB.Recordset
Private bLoading As Boolean

Public Sub ConnessioneADO()

    If Not (CnDMT Is Nothing) Then
        CnDMT.CloseConnection
        Set CnDMT = Nothing
    End If
    
    Set CnDMT = TheApp.Database.Connection
    GET_CAMPI_DISPONIBILI
    GET_GRIGLIA_ATTIVITA
    Me.Text1.Text = GET_STRINGA
    
    Me.Caption = TheApp.FunctionName
    INIT_CONTROLLI
End Sub

Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set TheApp = NewValue
End Property

Public Property Get Application() As DMTRunAppLib.Application
    Set Application = TheApp
End Property

Private Sub cboTipoSezioneEmail_Click()
    GET_GRIGLIA_ATTIVITA
    Me.Text1.Text = GET_STRINGA
End Sub

Private Sub cmdElimina_Click()
On Error GoTo ERR_cmdElimina_Click
Dim sSQL As String

If (Me.Griglia.Recordset.EOF) And (Me.Griglia.Recordset.BOF) Then Exit Sub

'''ELIMINAZIONE RIGA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POStringaEmail "
sSQL = sSQL & "WHERE IDRV_POStringaEmail=" & Me.Griglia.AllColumns("IDRV_POStringaEmail").Value
CnDMT.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

RICALCOLO_DOPO_ELIMINAZIONE_RIGA

GET_GRIGLIA_ATTIVITA

Me.Text1.Text = GET_STRINGA

Exit Sub

ERR_cmdElimina_Click:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
End Sub

Private Sub cmdGiu_Click()
On Error GoTo ERR_cmdGiu_Click
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim NumeroRecord As Long
Dim MAX_POSIZIONE As Long

If (Me.Griglia.Recordset.EOF) And (Me.Griglia.Recordset.BOF) Then Exit Sub

''''''''''''''''''''CALCOLO DELLA MASSIMA POSIZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT MAX(Posizione) as MaxPosizione "
sSQL = sSQL & "FROM RV_POStringaEmail "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDRV_POTipoSezioneEmail=" & Me.cboTipoSezioneEmail.CurrentID
Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    MAX_POSIZIONE = 1
Else
    MAX_POSIZIONE = fnNotNullN(rs!MaxPosizione)
End If

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Me.Griglia.AllColumns("Posizione").Value = MAX_POSIZIONE Then Exit Sub

GET_CAMBIO_POSIZIONE_GIU Me.Griglia.AllColumns("IDRV_POStringaEmail").Value, Me.Griglia.AllColumns("Posizione").Value

NumeroRecord = Me.Griglia.ListIndex - 1

GET_GRIGLIA_ATTIVITA

Me.Griglia.Recordset.Move NumeroRecord + 1

Me.Text1 = GET_STRINGA

Exit Sub
ERR_cmdGiu_Click:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
End Sub

Private Sub cmdSu_Click()
On Error GoTo ERR_cmdSu_Click
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim NumeroRecord As Long
Dim MAX_POSIZIONE As Long

If (Me.Griglia.Recordset.EOF) And (Me.Griglia.Recordset.BOF) <= 0 Then Exit Sub

If Me.Griglia.AllColumns("Posizione").Value = 1 Then Exit Sub

GET_CAMBIO_POSIZIONE_SU Me.Griglia.AllColumns("IDRV_POStringaEmail").Value, Me.Griglia.AllColumns("Posizione").Value

NumeroRecord = Me.Griglia.ListIndex - 1

GET_GRIGLIA_ATTIVITA


Me.Griglia.Recordset.Move NumeroRecord - 1

Me.Text1.Text = GET_STRINGA

Exit Sub
ERR_cmdSu_Click:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lstCampi_DblClick
    End If
    
End Sub

Private Sub Form_Load()
    With HScroll1
      .Max = (Pic1.ScaleWidth)
      .LargeChange = .Max \ 10
      .SmallChange = .Max \ 10
      
    End With

    With VScroll1
      .Max = (Pic1.ScaleHeight)
      .LargeChange = .Max \ 10
      .SmallChange = .Max \ 10
    End With
    
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
    

        If Me.ScaleWidth < Me.Pic1.ScaleWidth Then
            Me.HScroll1.Visible = True
            Me.HScroll1.Top = Me.ScaleHeight - Me.HScroll1.Height
            Me.HScroll1.Left = 0
            
        Else
            Me.HScroll1.Visible = False
        End If
        
        If Me.ScaleHeight < Me.Pic1.ScaleHeight Then
            Me.VScroll1.Visible = True
            Me.VScroll1.Top = 0
            Me.VScroll1.Left = Me.ScaleWidth - Me.VScroll1.Width
            
        Else
            Me.VScroll1.Visible = False
        End If
        
        If (VScroll1.Visible = True) And (HScroll1.Visible = False) Then
            Me.VScroll1.Height = Me.ScaleHeight '- Me.HScroll1.Height
        Else
            Me.VScroll1.Height = Me.ScaleHeight - Me.HScroll1.Height
        End If
        
        If (HScroll1.Visible = True) And (HScroll1.Visible = True) Then
            Me.HScroll1.Width = Me.ScaleWidth '- Me.VScroll1.Width
        Else
            Me.HScroll1.Width = Me.ScaleWidth - Me.VScroll1.Width
        End If
            
        With HScroll1
            .Max = (Pic1.ScaleWidth - Me.ScaleWidth + Me.VScroll1.Width)
            If .Max > 0 Then
                .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With

        With VScroll1
            .Max = (Pic1.ScaleHeight - Me.ScaleHeight + Me.HScroll1.Height)
            If .Max > 0 Then
                 .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With
        
        
    End If

End Sub


Private Sub lstCampi_DblClick()
    If Me.cboTipoSezioneEmail.CurrentID = 0 Then Exit Sub

    If Me.lstCampi.ListIndex > 0 Then
        TIPO_INSERIMENTO = 1
    Else
        TIPO_INSERIMENTO = 2
    End If
    
    frmInserimento.Show vbModal
    GET_GRIGLIA_ATTIVITA
    Me.Text1.Text = GET_STRINGA
    
End Sub

Private Sub VScroll1_Change()
   Me.Pic1.Top = -VScroll1.Value
End Sub
Private Sub VScroll1_Scroll()
   Me.Pic1.Top = -VScroll1.Value
End Sub
Private Sub HScroll1_Change()
   Me.Pic1.Left = -HScroll1.Value
End Sub
Private Sub HScroll1_Scroll()
   Me.Pic1.Left = -HScroll1.Value
End Sub

Private Function fncTrovaIDFunzione(Gestore As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Funzione.IDFunzione, Gestore.Gestore "
sSQL = sSQL & "FROM Gestore INNER JOIN "
sSQL = sSQL & "TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore INNER JOIN "
sSQL = sSQL & "Funzione ON TipoOggetto.IDTipoOggetto = Funzione.IDTipoOggetto "
sSQL = sSQL & "WHERE (Gestore.Gestore = " & fnNormString(Gestore) & ") "
sSQL = sSQL & "AND (Funzione.IDFunzione >= 10000)"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaIDFunzione = fnNotNullN(rs!IDFunzione)
Else
    fncTrovaIDFunzione = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_GRIGLIA_ATTIVITA()
'On Error GoTo ERR_fnGrigliaAssegnazione
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader



    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3

    sSQL = "SELECT * FROM RV_POStringaEmail "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDRV_POTipoSezioneEmail=" & Me.cboTipoSezioneEmail.CurrentID
    sSQL = sSQL & " ORDER BY Posizione"
    

    
    Set rsGriglia = New ADODB.Recordset
    rsGriglia.CursorLocation = adUseClient
    rsGriglia.Open sSQL, CnDMT.InternalConnection
    
        With Me.Griglia
            .EnableMove = True
            .UpdatePosition = True
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear


            With Me.Griglia.ColumnsHeader
                .Add "IDRV_POStringaEmail", "IDRV_POStringaEmail", dgInteger, False, 500, dgAlignleft
                .Add "IDAzienda", "IDAzienda", dgInteger, False, 500, dgAlignleft
                .Add "IDRV_POTipoSezioneEmail", "IDRV_POTipoSezioneEmail", dgInteger, False, 500, dgAlignleft
                .Add "Posizione", "Posizione", dgInteger, True, 1000, dgAlignRight
                .Add "NomeCampo", "Nome campo", dgchar, True, 2500, dgAlignleft
                .Add "ValoreCampo", "Valore", dgchar, True, 2000, dgAlignleft
                .Add "CarattereSpazio", "Spazio", dgBoolean, True, 1500, dgAligncenter
                .Add "CarattereACapo", "A capo", dgBoolean, True, 1500, dgAligncenter
            End With
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaAssegnazione:
    MsgBox Err.Description, vbCritical, "Reperimento dati assegnazione"
End Sub


Private Sub GET_CAMPI_DISPONIBILI()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Long

sSQL = "SELECT * FROM RV_POIEIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

Me.lstCampi.Clear

Me.lstCampi.AddItem " AGGIUNGI DESCRIZIONE PERSONALIZZATA"
For I = 0 To rs.Fields.Count - 1
    If Mid(rs.Fields(I).Name, 1, 2) <> "ID" Then
        Me.lstCampi.AddItem fnNotNull(rs.Fields(I).Name)
    End If
Next
rs.Close
Set rs = Nothing
End Sub
Private Function GET_STRINGA() As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POStringaEmail "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDRV_POTipoSezioneEmail=" & Me.cboTipoSezioneEmail.CurrentID
sSQL = sSQL & " ORDER BY Posizione"

Set rs = CnDMT.OpenResultset(sSQL)
GET_STRINGA = ""
While Not rs.EOF
    If Len(fnNotNull(rs!ValoreCampo)) = 0 Then
        If fnNotNullN(rs!CarattereSpazio) = 0 Then
            GET_STRINGA = GET_STRINGA & "[" & fnNotNull(rs!NomeCampo) & "]"
        Else
            GET_STRINGA = GET_STRINGA & " [" & fnNotNull(rs!NomeCampo) & "]"
        End If
    Else
        If fnNotNullN(rs!CarattereSpazio) = 0 Then
            GET_STRINGA = GET_STRINGA & fnNotNull(rs!ValoreCampo)
        Else
            GET_STRINGA = GET_STRINGA & " " & fnNotNull(rs!ValoreCampo)
        End If
    End If
    If fnNotNullN(rs!CarattereACapo) = 1 Then
        GET_STRINGA = GET_STRINGA & vbCrLf
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub RICALCOLO_DOPO_ELIMINAZIONE_RIGA()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsUpd As ADODB.Recordset
Dim NumeroPosizione As Long

sSQL = "SELECT * FROM RV_POStringaEmail "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDRV_POTipoSezioneEmail=" & Me.cboTipoSezioneEmail.CurrentID
sSQL = sSQL & " ORDER BY Posizione"

Set rsUpd = New ADODB.Recordset

rsUpd.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
NumeroPosizione = 1
While Not rsUpd.EOF
    rsUpd!Posizione = NumeroPosizione
    rsUpd.Update
    NumeroPosizione = NumeroPosizione + 1
rsUpd.MoveNext
Wend

rsUpd.Close
Set rsUpd = Nothing
End Sub
Private Function GET_CAMBIO_POSIZIONE_SU(IDRigaStringa As Long, PosizioneOriginale As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset

''''''AGGIORNAMENTO POSIZIONE ORDINE'''''''''''''''''''''''''''''''''''''''''''
sSQL = "UPDATE RV_POStringaEmail SET "
sSQL = sSQL & "Posizione=" & PosizioneOriginale - 1
sSQL = sSQL & " WHERE IDRV_POStringaEmail=" & IDRigaStringa
sSQL = sSQL & " AND IDRV_POTipoSezioneEmail=" & Me.cboTipoSezioneEmail.CurrentID

CnDMT.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_POStringaEmail "
sSQL = sSQL & "WHERE IDRV_POStringaEmail<>" & IDRigaStringa
sSQL = sSQL & " AND Posizione=" & (PosizioneOriginale - 1)
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDRV_POTipoSezioneEmail=" & Me.cboTipoSezioneEmail.CurrentID
sSQL = sSQL & " ORDER BY Posizione"

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rs.EOF
    rs!Posizione = rs!Posizione + 1
    rs.Update
rs.MoveNext
Wend
rs.Close
Set rs = Nothing
End Function

Private Function GET_CAMBIO_POSIZIONE_GIU(IDRigaStringa As Long, PosizioneOriginale As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset

''''''AGGIORNAMENTO POSIZIONE ORDINE'''''''''''''''''''''''''''''''''''''''''''
sSQL = "UPDATE RV_POStringaEmail SET "
sSQL = sSQL & "Posizione=" & PosizioneOriginale + 1
sSQL = sSQL & " WHERE IDRV_POStringaEmail=" & IDRigaStringa
sSQL = sSQL & " AND IDRV_POTipoSezioneEmail=" & Me.cboTipoSezioneEmail.CurrentID

CnDMT.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_POStringaEmail "
sSQL = sSQL & "WHERE IDRV_POStringaEmail<>" & IDRigaStringa
sSQL = sSQL & " AND Posizione=" & (PosizioneOriginale + 1)
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDRV_POTipoSezioneEmail=" & Me.cboTipoSezioneEmail.CurrentID
sSQL = sSQL & " ORDER BY Posizione"

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
While Not rs.EOF
    rs!Posizione = rs!Posizione - 1
    rs.Update
rs.MoveNext
Wend
rs.Close
Set rs = Nothing
End Function
Private Sub INIT_CONTROLLI()
    'Utente di inserimento
    With Me.cboTipoSezioneEmail
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoSezioneEmail"
        .DisplayField = "TipoSezioneEmail"
        .Sql = "SELECT * FROM RV_POTipoSezioneEmail ORDER BY IDRV_POTipoSezioneEmail"
        .Fill
    End With
    
    Me.cboTipoSezioneEmail.WriteOn 1
End Sub
