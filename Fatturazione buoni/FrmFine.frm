VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Begin VB.Form FrmParametri 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fatturazione buoni  (Passo 2 di 3)"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVisAnteprima 
      Caption         =   "Visualizza anteprima"
      Height          =   495
      Left            =   2400
      TabIndex        =   17
      Top             =   4680
      Width           =   4455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Crea fattura per ogni addebito"
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Top             =   3840
      Width           =   4455
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
      Height          =   5895
      Left            =   0
      Picture         =   "FrmFine.frx":4781A
      ScaleHeight     =   5835
      ScaleWidth      =   2235
      TabIndex        =   15
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Avanti"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton CmdFine 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton CmdIndietro 
      Caption         =   "Indietro"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo di documento"
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
      Height          =   3615
      Left            =   2400
      TabIndex        =   6
      Top             =   0
      Width           =   4455
      Begin DMTDataCmb.DMTCombo CboPagamento 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   3120
         Width           =   4215
         _ExtentX        =   7435
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
      Begin DMTDataCmb.DMTCombo CboMagazzino 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   4215
         _ExtentX        =   7435
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
      Begin DMTDATETIMELib.dmtDate DataDoc 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo CboValuta 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
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
      Begin DMTDataCmb.DMTCombo CboSezionale 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   4215
         _ExtentX        =   7435
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
      Begin DMTDataCmb.DMTCombo CboTipoOggetto 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
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
         Caption         =   "Pagamento di Default"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Magazzino"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Data documento"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Valuta"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Sezionale"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   3495
      End
   End
End
Attribute VB_Name = "FrmParametri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Indica se ci sono parametri di default per l'azienda
'per l'importazione dei dati
'0 = non ci sono dati
'1 = ci sono i dati
Public NewRecDefault As Integer
Public odbLib As DMTDataLayer.Database

Private Function ControlloInserimento() As String
    ControlloInserimento = ""
    
    If Me.CboTipoOggetto.Text = "" Then
        ControlloInserimento = ControlloInserimento & "Inserire il tipo di documento" & vbCrLf
    Else
        Link_TipoOggetto = Me.CboTipoOggetto.CurrentID
        Var_TipoOggetto = Me.CboTipoOggetto.Text
    End If
    If Me.CboSezionale.CurrentID = 0 Then
        ControlloInserimento = ControlloInserimento & "Inserire il tipo di sezionale del documento" & vbCrLf
    Else
        Link_Sezionale = Me.CboSezionale.CurrentID
    End If
    If Me.CboValuta.CurrentID = 0 Then
        ControlloInserimento = ControlloInserimento & "Inserire il tipo di valuta del documento" & vbCrLf
    Else
        Link_Valuta = Me.CboValuta.CurrentID
    End If
    If Me.CboMagazzino.CurrentID = 0 Then
        ControlloInserimento = ControlloInserimento & "Inserire il magazzino" & vbCrLf
    Else
        Link_Magazzino = Me.CboMagazzino.CurrentID
    End If
    If Me.CboPagamento.CurrentID = 0 Then
        ControlloInserimento = ControlloInserimento & "Inserire il tipo di pagamento di default" & vbCrLf
    Else
        Link_PagamentoDefault = Me.CboPagamento.CurrentID
    End If
    
    SingolaFattura = Me.Check1.Value
    
    
    Data_Documento = Me.DataDoc.Text
    
End Function

Private Sub CboTipoOggetto_Click()
    With Me.CboSezionale
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .Sql = "SELECT Sezionale.IDSezionale, Sezionale.Sezionale "
        .Sql = .Sql & "FROM RegistroIvaPerTipoOggetto INNER JOIN "
        .Sql = .Sql & "Sezionale ON RegistroIvaPerTipoOggetto.IDRegistroIva = Sezionale.IDRegistroIva AND "
        .Sql = .Sql & "RegistroIvaPerTipoOggetto.IDFiliale = Sezionale.IDFiliale LEFT OUTER JOIN "
        .Sql = .Sql & "TipoOggetto ON RegistroIvaPerTipoOggetto.IDTipoOggetto = TipoOggetto.IDTipoOggetto "
        .Sql = .Sql & "WHERE RegistroIvaPerTipoOggetto.IDTipoOggetto = " & Me.CboTipoOggetto.CurrentID
        .Sql = .Sql & " AND RegistroIvaPerTipoOggetto.IDFiliale = " & TheApp.Branch
        .Fill
    End With
End Sub

Private Sub cmdAvanti_Click()
    If ControlloInserimento = "" Then
        'fncPassaggioDocumenti
        SalvataggioParametriDefault
        Unload Me
    Else
        MsgBox ControlloInserimento, vbInformation, "Parametri mancanti"
    End If
End Sub

Private Sub CmdFine_Click()
    Dim Risposta As Integer
    Risposta = MsgBox("Vuoi abbandonare il wizard per il passaggio dei buoni in fatturazione?", vbInformation + vbYesNo, "Abbandono")
    If Risposta = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdIndietro_Click()
    Unload Me
End Sub
Private Sub SalvataggioParametriDefault()
On Error GoTo ERR_SalvataggioParametriDefault
    Dim sSQL As String
    
    If NewRecDefault = 0 Then
    'NUOVO RECORD DI DEFAULT
        sSQL = "INSERT INTO RV_POParametriDefault "
        sSQL = sSQL & "(IDAzienda, IDTipoOggetto, IDSezionale, IDValuta, IDMagazzino, IDPagamento, SingolaFattura, "
        sSQL = sSQL & "RaggrAltraDestinazioneFattAdd, RaggrCorpoAltraDestinazioneFattAdd, RaggrCorpoInterventoFattAdd, SovrascriviDescrizioneArticolo) "
        sSQL = sSQL & " VALUES ("
        sSQL = sSQL & TheApp.IDFirm & ", "
        sSQL = sSQL & Me.CboTipoOggetto.CurrentID & ", "
        sSQL = sSQL & Me.CboSezionale.CurrentID & ", "
        sSQL = sSQL & Me.CboValuta.CurrentID & ", "
        sSQL = sSQL & Me.CboMagazzino.CurrentID & ", "
        sSQL = sSQL & Me.CboPagamento.CurrentID & ", "
        sSQL = sSQL & Me.Check1.Value & ", "
        sSQL = sSQL & FLAG_RAGGR_ALTRA_DEST & ", "
        sSQL = sSQL & FLAG_RAGGR_CORPO_ALTRA_DEST & ", "
        sSQL = sSQL & FLAG_RAGGR_CORPO_INT & ", "
        sSQL = sSQL & FLAG_SOVRAS_DESCRIZIONE & ")"
    Else
    'AGGIORNA IL RECORD DI DEFAULT
        sSQL = "UPDATE RV_POParametriDefault SET "
        sSQL = sSQL & "IDTipoOggetto=" & Me.CboTipoOggetto.CurrentID & ", "
        sSQL = sSQL & "IDSezionale=" & Me.CboSezionale.CurrentID & ", "
        sSQL = sSQL & "IDValuta=" & Me.CboValuta.CurrentID & ", "
        sSQL = sSQL & "IDMagazzino=" & Me.CboMagazzino.CurrentID & ", "
        sSQL = sSQL & "IDPagamento=" & Me.CboPagamento.CurrentID & ", "
        sSQL = sSQL & "SingolaFattura=" & Me.Check1.Value & ", "
        sSQL = sSQL & "RaggrAltraDestinazioneFattAdd=" & FLAG_RAGGR_ALTRA_DEST & ", "
        sSQL = sSQL & "RaggrCorpoAltraDestinazioneFattAdd=" & FLAG_RAGGR_CORPO_ALTRA_DEST & ", "
        sSQL = sSQL & "RaggrCorpoInterventoFattAdd=" & FLAG_RAGGR_CORPO_INT & ", "
        sSQL = sSQL & "SovrascriviDescrizioneArticolo=" & FLAG_SOVRAS_DESCRIZIONE
        sSQL = sSQL & " WHERE IDAzienda= " & TheApp.IDFirm
        
    End If
    
    CnDMT.Execute sSQL
    
Exit Sub

ERR_SalvataggioParametriDefault:
    MsgBox Err.Description, vbCritical, "Salvataggio Parametri Default"
End Sub
Private Sub DefaultImportazione()
    Dim rs As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT * From RV_POParametriDefault Where IDAzienda=" & TheApp.IDFirm
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        
        Me.CboTipoOggetto.WriteOn IIf(IsNull(rs!IDTipoOggetto), 0, rs!IDTipoOggetto)
        Me.CboSezionale.WriteOn IIf(IsNull(rs!IDSezionale), 0, rs!IDSezionale)
        Me.CboValuta.WriteOn IIf(IsNull(rs!IDValuta), 0, rs!IDValuta)
        Me.CboMagazzino.WriteOn IIf(IsNull(rs!IDMagazzino), 0, rs!IDMagazzino)
        Me.CboPagamento.WriteOn IIf(IsNull(rs!IDPagamento), 0, rs!IDPagamento)
        Me.Check1.Value = Abs(fnNotNullN(rs!SingolaFattura))
        
        'Me.CDArticolo.Load fnNotNullN(rs!IDArticolo)
        NewRecDefault = 1
    Else
        Me.CboSezionale.WriteOn 0
        Me.CboTipoOggetto.WriteOn 0
        Me.CboValuta.WriteOn 0
        Me.CboMagazzino.WriteOn 0
        Me.CboPagamento.WriteOn 0
        Me.Check1.Value = vbUnchecked
        
        'Me.CDArticolo.Load 0
        NewRecDefault = 0
    End If
    
    
End Sub

Private Sub cmdVisAnteprima_Click()
    FrmMovimenti.Show vbModal
    
End Sub

Private Sub Form_Activate()
    Me.Caption = TheApp.FunctionName & " (Passo 2 di 3)"
    Me.SetFocus
    
    
    If ((FLAG_RAGGR_ALTRA_DEST = 1) Or (FLAG_RAGGR_CORPO_ALTRA_DEST = 1) Or (FLAG_RAGGR_CORPO_INT = 1)) Then
        
        Me.Check1.Enabled = False
        Me.Check1.Value = vbUnchecked

    End If
    
End Sub

Private Sub Form_Load()
'On Error GoTo ERR_Form_Load
    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    
    fncTipoOggetto
    'fncSezionale
    fncValuta
    fncMagazzino
    fncPagamento
    'fncArticolo
    Me.DataDoc.Text = Date
    
    
    DefaultImportazione
    'CREA_TABELLA_TEMPORANEA_ANAGRAFICA
    'CREA_TABELLA_TEMPORANEA_BUONI
    
Exit Sub

ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERR_Form_Unload
    If Me.cmdAvanti.Value = True Then
        
        frmCreazioneDocumenti.Show
        

        Exit Sub
    End If
    
    
    If Me.cmdIndietro.Value = True Then
        
        FrmInizio.Show
        Exit Sub
        
    End If


Exit Sub

ERR_Form_Unload:
    MsgBox Err.Description, vbCritical, "Form_Unload"
    
    
End Sub
Private Function SettaggioProgressBar() As Long

    
End Function
Private Sub fncTipoOggetto()
    Dim sSQL As String
    
    sSQL = "SELECT IDTipoOggetto, Oggetto"
    sSQL = sSQL & " FROM TipoOggetto"
    sSQL = sSQL & " WHERE IDGestore=15"
    sSQL = sSQL & " ORDER BY Oggetto"
    

    With Me.CboTipoOggetto
    Set .Database = CnDMT
    .DisplayField = "Oggetto"
    .AddFieldKey "IDTipoOggetto"
    .Sql = sSQL
    .Refresh
    End With
    
End Sub
    
    
Private Sub fncSezionale()
    Dim sSQL As String
    
    sSQL = "SELECT IDSezionale, Sezionale"
    sSQL = sSQL & " FROM Sezionale"
    sSQL = sSQL & " WHERE ((IDFiliale=" & TheApp.Branch & ") AND (IDRegistroIva = 1))"
    sSQL = sSQL & " ORDER BY Sezionale"

    With Me.CboSezionale
    Set .Database = CnDMT
    .DisplayField = "Sezionale"
    .AddFieldKey "IDSezionale"
    .Sql = sSQL
    .Refresh
    End With
    
End Sub
Private Sub fncValuta()
    Dim sSQL As String
    'Dim sSQLValuta As String
    'Dim rs As DmtOLedbLib.adoResultset
    
    
    sSQL = "SELECT IDValuta, Valuta"
    sSQL = sSQL & " FROM Valuta"
    'sSQL = sSQL & " WHERE ((IDFiliale=" & theapp.branch & ") AND (IDRegistroIva = 1))"
    sSQL = sSQL & " ORDER BY Valuta"

    With Me.CboValuta
    Set .Database = CnDMT
    .DisplayField = "Valuta"
    .AddFieldKey "IDValuta"
    .Sql = sSQL
    .Refresh
    End With
    
End Sub
Private Sub fncMagazzino()
    Dim sSQL As String
    'Dim sSQLValuta As String
    'Dim rs As DmtOLedbLib.adoResultset
    
    
    sSQL = "SELECT IDMagazzino, Magazzino"
    sSQL = sSQL & " FROM Magazzino"
    sSQL = sSQL & " WHERE IDAzienda = " & TheApp.IDFirm
    sSQL = sSQL & " ORDER BY Magazzino"

    With Me.CboMagazzino
    Set .Database = CnDMT
    .DisplayField = "Magazzino"
    .AddFieldKey "IDMagazzino"
    .Sql = sSQL
    .Refresh
    End With

End Sub
Private Sub fncPagamento()
    Dim sSQL As String

    
    
    sSQL = "SELECT IDPagamento, Pagamento"
    sSQL = sSQL & " FROM Pagamento"
    sSQL = sSQL & " ORDER BY Pagamento"

    With Me.CboPagamento
    Set .Database = CnDMT
    .DisplayField = "Pagamento"
    .AddFieldKey "IDPagamento"
    .Sql = sSQL
    .Refresh
    End With

End Sub
Private Sub CREA_TABELLA_TEMPORANEA_ANAGRAFICA()
Set rsAnaFatt = New ADODB.Recordset

rsAnaFatt.CursorLocation = adUseClient

rsAnaFatt.Fields.Append "IDAnagraficaCliente", adInteger, , adFldIsNullable
rsAnaFatt.Fields.Append "AnagraficaCliente", adVarChar, 250, adFldIsNullable
rsAnaFatt.Fields.Append "NomeCliente", adVarChar, 250, adFldIsNullable

rsAnaFatt.Open , , adOpenKeyset, adLockBatchOptimistic

If Not (rsAna.EOF And rsAnaFatt.BOF) Then
    rsAna.MoveFirst
    While Not rsAna.EOF
        rsAnaFatt.AddNew
            rsAnaFatt!IDAnagraficaCliente = rsAna!IDAnagraficaCliente
            rsAnaFatt!AnagraficaCliente = rsAna!AnagraficaCliente
            rsAnaFatt!NomeCliente = rsAna!NomeCliente
        rsAna.Update
    rsAna.MoveNext
    Wend
rsAna.MoveFirst
End If

End Sub
Private Sub CREA_TABELLA_TEMPORANEA_BUONI()

Set rsBuoniFatt = New ADODB.Recordset

rsBuoniFatt.CursorLocation = adUseClient

rsBuoniFatt.Fields.Append "IDRV_POInterventoRigheDett", adInteger, , adFldIsNullable
rsBuoniFatt.Fields.Append "IDRV_POInterventoRighe", adInteger, , adFldIsNullable
rsBuoniFatt.Fields.Append "IDRV_POIntervento", adInteger, , adFldIsNullable
rsBuoniFatt.Fields.Append "NumeroBuono", adInteger, , adFldIsNullable
rsBuoniFatt.Fields.Append "DataBuono", adVarChar, 250, adFldIsNullable
rsBuoniFatt.Fields.Append "IDAnagraficaTecnico", adInteger, , adFldIsNullable
rsBuoniFatt.Fields.Append "AnagraficaTecnico", adVarChar, 250, adFldIsNullable
rsBuoniFatt.Fields.Append "NomeTecnico", adVarChar, 250, adFldIsNullable
rsBuoniFatt.Fields.Append "IDAnagraficaCliente", adInteger, , adFldIsNullable
rsBuoniFatt.Fields.Append "AnagraficaCliente", adVarChar, 250, adFldIsNullable
rsBuoniFatt.Fields.Append "NomeCliente", adVarChar, 250, adFldIsNullable
rsBuoniFatt.Fields.Append "IDAnagraficaIntervento", adInteger, , adFldIsNullable
rsBuoniFatt.Fields.Append "AnagraficaIntervento", adVarChar, 250, adFldIsNullable
rsBuoniFatt.Fields.Append "NomeAnagraficaIntervento", adVarChar, 250, adFldIsNullable
rsBuoniFatt.Fields.Append "ImportoFatturazione", adDouble, , adFldIsNullable
rsBuoniFatt.Fields.Append "Quantita", adDouble, , adFldIsNullable
rsBuoniFatt.Fields.Append "ImportoUnitario", adDouble, , adFldIsNullable
rsBuoniFatt.Fields.Append "IDIva", adInteger, , adFldIsNullable
rsBuoniFatt.Fields.Append "AliquotaIva", adDouble, , adFldIsNullable
rsBuoniFatt.Fields.Append "Sconto1", adDouble, , adFldIsNullable
rsBuoniFatt.Fields.Append "Sconto2", adDouble, , adFldIsNullable
rsBuoniFatt.Fields.Append "Sconto3", adDouble, , adFldIsNullable
rsBuoniFatt.Fields.Append "TotaleRigaNettoIva", adDouble, , adFldIsNullable
rsBuoniFatt.Fields.Append "TotaleRigaIva", adDouble, , adFldIsNullable
rsBuoniFatt.Fields.Append "TotaleRigaLordoIva", adDouble, , adFldIsNullable
rsBuoniFatt.Fields.Append "RitenutaAcconto", adSmallInt, , adFldIsNullable
rsBuoniFatt.Fields.Append "IDSitoPerAnagrafica", adInteger, , adFldIsNullable
rsBuoniFatt.Fields.Append "IDAccordoCommerciale", adInteger, , adFldIsNullable
rsBuoniFatt.Fields.Append "IDRaggruppamentoFatturato", adInteger, , adFldIsNullable
rsBuoniFatt.Fields.Append "EntePubblico", adSmallInt, , adFldIsNullable
rsBuoniFatt.Fields.Append "IDMagazzino", adInteger, , adFldIsNullable



rsBuoniFatt.Open , , adOpenKeyset, adLockBatchOptimistic

rsBuoni.Filter = ""

If Not (rsBuoni.EOF And rsBuoni.BOF) Then
    rsBuoni.MoveFirst
    While Not rsBuoni.EOF
        rsBuoniFatt.AddNew
            rsBuoniFatt!IDRV_POInterventoRigheDett = rsBuoni("IDRV_POInterventoRigheDett").Value
            rsBuoniFatt!IDRV_POInterventoRighe = rsBuoni("IDRV_POInterventoRighe").Value
            rsBuoniFatt!IDRV_POintervento = rsBuoni("IDRV_POIntervento").Value
            rsBuoniFatt!NumeroBuono = rsBuoni("NumeroBuono").Value
            rsBuoniFatt!DataBuono = rsBuoni("DataBuono").Value
            rsBuoniFatt!IDAnagraficaTecnico = rsBuoni("IDAnagraficaTecnico").Value
            rsBuoniFatt!AnagraficaTecnico = rsBuoni("AnagraficaTecnico").Value
            rsBuoniFatt!NomeTecnico = rsBuoni("NomeTecnico").Value
            rsBuoniFatt!IDAnagraficaCliente = rsBuoni("IDAnagraficaCliente").Value
            rsBuoniFatt!AnagraficaCliente = rsBuoni("AnagraficaCliente").Value
            rsBuoniFatt!NomeCliente = rsBuoni("NomeCliente").Value
            rsBuoniFatt!IDAnagraficaIntervento = rsBuoni("IDAnagraficaIntervento").Value
            rsBuoniFatt!AnagraficaIntervento = rsBuoni("AnagraficaIntervento").Value
            rsBuoniFatt!NomeAnagraficaIntervento = rsBuoni("NomeAnagraficaIntervento").Value
            
            rsBuoniFatt!ImportoFatturazione = rsBuoni!ImportoFatturazione

            rsBuoniFatt!RitenutaAcconto = rsBuoni!RitenutaAcconto
            rsBuoniFatt!Quantita = rsBuoni!Quantita
            rsBuoniFatt!ImportoUnitario = rsBuoni!ImportoUnitario
            rsBuoniFatt!IDIva = rsBuoni!IDIva
            rsBuoniFatt!AliquotaIva = rsBuoni!AliquotaIva
            rsBuoniFatt!Sconto1 = rsBuoni!Sconto1
            rsBuoniFatt!Sconto2 = rsBuoni!Sconto2
            rsBuoniFatt!Sconto3 = rsBuoni!Sconto3
            rsBuoniFatt!TotaleRigaNettoIva = rsBuoni!TotaleRigaNettoIva
            rsBuoniFatt!TotaleRigaIva = rsBuoni!TotaleRigaIva
            rsBuoniFatt!TotaleRigaLordoIva = rsBuoni!TotaleRigaLordoIva
            
            rsBuoniFatt.Fields.Append "IDSitoPerAnagrafica", adInteger, , adFldIsNullable
            rsBuoniFatt.Fields.Append "IDAccordoCommerciale", adInteger, , adFldIsNullable
            rsBuoniFatt.Fields.Append "IDRaggruppamentoFatturato", adInteger, , adFldIsNullable
            rsBuoniFatt.Fields.Append "EntePubblico", adSmallInt, , adFldIsNullable
            rsBuoniFatt.Fields.Append "IDMagazzino", adInteger, , adFldIsNullable
            
            
        rsBuoni.Update
    
    rsBuoni.MoveNext
    Wend

End If
End Sub
