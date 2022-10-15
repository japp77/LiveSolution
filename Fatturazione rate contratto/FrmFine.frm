VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Begin VB.Form FrmParametri 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creazione documenti  (Passo 3 di 4)"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   Icon            =   "FrmFine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   7095
      Left            =   0
      Picture         =   "FrmFine.frx":4781A
      ScaleHeight     =   7035
      ScaleWidth      =   2475
      TabIndex        =   15
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton cmdAvanti 
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
      TabIndex        =   6
      Top             =   6600
      Width           =   1335
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
      Left            =   4800
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton CmdIndietro 
      Caption         =   "Indietro"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   6600
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
      Height          =   6375
      Left            =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CheckBox chkRipDescrContr 
         Caption         =   "Stampa riferimento contratto nel corpo"
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
         TabIndex        =   23
         ToolTipText     =   "Raggruppamento per cliente di fatturazione - per cliente del contratto - pagamento - ritenuta acconto"
         Top             =   5880
         Width           =   4095
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroDocumento 
         Height          =   315
         Left            =   2040
         TabIndex        =   21
         Top             =   2160
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
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
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin VB.CheckBox chkRaggrClienteContr 
         Caption         =   "Raggruppa per anagrafica del contratto"
         Enabled         =   0   'False
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
         TabIndex        =   18
         ToolTipText     =   "Raggruppamento per cliente di fatturazione - per cliente del contratto - pagamento - ritenuta acconto"
         Top             =   5400
         Width           =   4095
      End
      Begin VB.CheckBox chkRaggruppaCliente 
         Caption         =   "Raggruppa per anagrafica di fatturazione"
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
         TabIndex        =   17
         ToolTipText     =   "Raggruppamento per cliente di fatturazione - pagamento - ritenuta acconto"
         Top             =   4920
         Width           =   4095
      End
      Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   3960
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   1085
         PropCodice      =   $"FrmFine.frx":4E4D0
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"FrmFine.frx":4E51F
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"FrmFine.frx":4E576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin DMTDataCmb.DMTCombo CboPagamento 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   3600
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
         TabIndex        =   12
         Top             =   2880
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
         TabIndex        =   10
         Top             =   2160
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   253
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
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo CboValuta 
         Height          =   315
         Left            =   3120
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
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
         TabIndex        =   2
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
         TabIndex        =   1
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
      Begin DMTDataCmb.DMTCombo cboSezionalePA 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1560
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
         Caption         =   "N° Doc."
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
         Index           =   6
         Left            =   2040
         TabIndex        =   22
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Sezionale per fatturazione PA"
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
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Pagamento di Default"
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
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Magazzino"
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
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Data documento"
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
         TabIndex        =   8
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Valuta"
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
         Left            =   3120
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Sezionale"
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
         TabIndex        =   3
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
    
    Dim IDEsercizio As Long
    
    IDEsercizio = GET_LINK_ESERCIZIO(Me.DataDoc.Text)
    
    If IDEsercizio = 0 Then
        ControlloInserimento = ControlloInserimento & "L'esercizio non esiste" & vbCrLf
    End If
    
    
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
    If Me.CDArticolo.KeyFieldID = 0 Then
        ControlloInserimento = ControlloInserimento & "Inserire l'articolo di riferimento" & vbCrLf
    Else
        Link_Articolo = Me.CDArticolo.KeyFieldID
        Var_Codice_Articolo = fnNotNull(Me.CDArticolo.Code)
    End If
    
    If Me.chkRaggruppaCliente.Value = vbChecked Then
        RaggruppamentoAnagrafica = 1
    Else
        RaggruppamentoAnagrafica = 0
    End If
    
    If Me.chkRaggrClienteContr.Value = vbChecked Then
        RaggruppamentoAnaContratto = 1
    Else
        RaggruppamentoAnaContratto = 0
    End If
    
    If Me.chkRipDescrContr.Value = vbChecked Then
        StampaRifContratto = 1
    Else
        StampaRifContratto = 0
    End If
    
    Data_Documento = Me.DataDoc.Text
    Var_NumeroDocumento = Me.txtNumeroDocumento.Value
    
    Link_SezionalePA = Me.cboSezionalePA.CurrentID
    
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

Private Sub chkRaggruppaCliente_Click()
    If Me.chkRaggruppaCliente.Value = vbChecked Then
        Me.chkRaggrClienteContr.Enabled = True
    Else
        Me.chkRaggrClienteContr.Enabled = False
        Me.chkRaggrClienteContr.Value = vbUnchecked
    End If
End Sub

Private Sub cmdAvanti_Click()
    If ControlloInserimento = "" Then
        'fncPassaggioDocumenti
        
        If RaggruppamentoAnagrafica = 1 Then
            SCRIVI_RAGGRUPPAMENTO_ANAGRAFICA
        End If
        
        SalvataggioParametriDefault
        
        Unload Me
    Else
        MsgBox ControlloInserimento, vbInformation, "Parametri mancanti"
    End If
End Sub

Private Sub CmdFine_Click()
    Dim Risposta As Integer
    Risposta = MsgBox("Vuoi abbandonare il wizard per il passaggio degli interventi in contabilità?", vbInformation + vbYesNo, "Abbandono")
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
        sSQL = sSQL & "(IDAzienda, IDTipoOggetto, IDSezionale, IDSezionalePA, IDValuta, "
        sSQL = sSQL & "IDMagazzino, IDPagamento, IDArticolo, "
        sSQL = sSQL & "RaggrClienteFatturazione, RaggrClienteContratto, StampaRiferimentoContratto)"
        sSQL = sSQL & " VALUES ("
        sSQL = sSQL & TheApp.IDFirm & ", "
        sSQL = sSQL & Me.CboTipoOggetto.CurrentID & ", "
        sSQL = sSQL & Me.CboSezionale.CurrentID & ", "
        sSQL = sSQL & Me.cboSezionalePA.CurrentID & ", "
        sSQL = sSQL & Me.CboValuta.CurrentID & ", "
        sSQL = sSQL & Me.CboMagazzino.CurrentID & ", "
        sSQL = sSQL & Me.CboPagamento.CurrentID & ", "
        sSQL = sSQL & Me.CDArticolo.KeyFieldID & ", "
        sSQL = sSQL & Abs(Me.chkRaggruppaCliente.Value) & ", "
        sSQL = sSQL & Abs(Me.chkRaggrClienteContr.Value) & ", "
        sSQL = sSQL & Abs(Me.chkRipDescrContr.Value) & ")"
    Else
    'AGGIORNA IL RECORD DI DEFAULT
        sSQL = "UPDATE RV_POParametriDefault SET "
        sSQL = sSQL & "IDTipoOggetto=" & Me.CboTipoOggetto.CurrentID & ", "
        sSQL = sSQL & "IDSezionale=" & Me.CboSezionale.CurrentID & ", "
        sSQL = sSQL & "IDSezionalePA=" & Me.cboSezionalePA.CurrentID & ", "
        sSQL = sSQL & "IDValuta=" & Me.CboValuta.CurrentID & ", "
        sSQL = sSQL & "IDMagazzino=" & Me.CboMagazzino.CurrentID & ", "
        sSQL = sSQL & "IDArticolo=" & Me.CDArticolo.KeyFieldID & ", "
        sSQL = sSQL & "IDPagamento=" & Me.CboPagamento.CurrentID & ", "
        sSQL = sSQL & "RaggrClienteFatturazione=" & Abs(Me.chkRaggruppaCliente.Value) & ", "
        sSQL = sSQL & "RaggrClienteContratto=" & Abs(Me.chkRaggrClienteContr.Value) & ", "
        sSQL = sSQL & "StampaRiferimentoContratto=" & Abs(Me.chkRipDescrContr.Value)
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
    
    sSQL = "SELECT * From RV_POParametriDefault "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        Me.CboTipoOggetto.WriteOn IIf(IsNull(rs!IDTipoOggetto), 0, rs!IDTipoOggetto)
        Me.CboSezionale.WriteOn IIf(IsNull(rs!IDSezionale), 0, rs!IDSezionale)
        Me.cboSezionalePA.WriteOn IIf(IsNull(rs!IDSezionalePA), 0, rs!IDSezionalePA)
        Me.CboValuta.WriteOn IIf(IsNull(rs!IDValuta), 0, rs!IDValuta)
        Me.CboMagazzino.WriteOn IIf(IsNull(rs!IDMagazzino), 0, rs!IDMagazzino)
        Me.CboPagamento.WriteOn IIf(IsNull(rs!IDPagamento), 0, rs!IDPagamento)
        Me.CDArticolo.Load fnNotNullN(rs!IDArticolo)
        Me.chkRaggruppaCliente.Value = Abs(fnNotNullN(rs!RaggrClienteFatturazione))
        Me.chkRaggrClienteContr.Value = Abs(fnNotNullN(rs!RaggrClienteContratto))
        Me.chkRipDescrContr.Value = Abs(fnNotNullN(rs!StampaRiferimentoContratto))
        NewRecDefault = 1
    Else
        Me.CboSezionale.WriteOn 0
        Me.cboSezionalePA.WriteOn 0
        Me.CboTipoOggetto.WriteOn 0
        Me.CboValuta.WriteOn 0
        Me.CboMagazzino.WriteOn 0
        Me.CboPagamento.WriteOn 0
        Me.CDArticolo.Load 0
        Me.chkRaggruppaCliente.Value = 0
        Me.chkRaggrClienteContr.Value = 0
        Me.chkRipDescrContr.Value = 0
        NewRecDefault = 0
    End If
    
    chkRaggruppaCliente_Click
    
End Sub
Private Sub Form_Load()
'On Error GoTo ERR_Form_Load
    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    
    fncTipoOggetto
    fncSezionale
    fncSezionalePA
    fncValuta
    fncMagazzino
    fncPagamento
    fncArticolo
    Me.DataDoc.Text = Date
    
    
    DefaultImportazione
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
        FrmMovimenti.Show
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
Private Sub fncSezionalePA()
    Dim sSQL As String
    
    'SEZIONALE PA
    sSQL = "SELECT IDSezionale, Sezionale"
    sSQL = sSQL & " FROM Sezionale"
    sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch
    sSQL = sSQL & " AND FatturaElettronica=" & fnNormBoolean(1)
    sSQL = sSQL & " AND IDRegistroIva=1"
    sSQL = sSQL & " ORDER BY Sezionale"
    
    With Me.cboSezionalePA
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
Private Sub fncArticolo()

Set odbLib = New DMTDataLayer.Database

'odbLib.OpenConnection "sa", Password, "MSDASQL", dbUseServer

    'Articolo
    With Me.CDArticolo
        'Set .Application = m_App
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 6 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    

End Sub
Private Sub SCRIVI_RAGGRUPPAMENTO_ANAGRAFICA()
    
    If Not (rsAnagrafica Is Nothing) Then
        If rsAnagrafica.State > 0 Then
            rsAnagrafica.Close
        End If
        Set rsAnagrafica = Nothing
    End If
    
    Set rsAnagrafica = New ADODB.Recordset
    rsAnagrafica.CursorLocation = adUseClient
    
    rsAnagrafica.Fields.Append "IDAnagrafica", adInteger, , adFldIsNullable 'IDAnagrafica di fatturazione
    If RaggruppamentoAnaContratto = 1 Then
        rsAnagrafica.Fields.Append "IDAnagraficaContratto", adInteger, , adFldIsNullable 'IDAnagrafica del contratto
    End If
    rsAnagrafica.Fields.Append "IDPagamento", adInteger, , adFldIsNullable
    rsAnagrafica.Fields.Append "RitenutaAcconto", adSmallInt, , adFldIsNullable
    rsAnagrafica.Fields.Append "IDRaggruppamentoFatturato", adInteger, , adFldIsNullable
    rsAnagrafica.Fields.Append "IDAccordoCommerciale", adInteger, , adFldIsNullable
    rsAnagrafica.Fields.Append "IDContrattoBancario", adInteger, , adFldIsNullable
    rsAnagrafica.Fields.Append "EntePubblico", adInteger, , adFldIsNullable
    rsAnagrafica.Fields.Append "IDSitoPerAnagrafica", adInteger, , adFldIsNullable
    rsAnagrafica.Fields.Append "IDAnagraficaAgente", adInteger, , adFldIsNullable
    
    rsAnagrafica.Open , , adOpenKeyset, adLockBatchOptimistic
    
    rsRateDaFatt.MoveFirst
        
    While Not rsRateDaFatt.EOF
        rsAnagrafica.Filter = "IDAnagrafica=" & fnNotNullN(rsRateDaFatt!IDAnagraficaFatturazione)
        If RaggruppamentoAnaContratto = 1 Then
            rsAnagrafica.Filter = "IDAnagraficaContratto=" & fnNotNullN(rsRateDaFatt!IDAnagrafica)
        End If
        rsAnagrafica.Filter = rsAnagrafica.Filter & " AND IDPagamento=" & fnNotNullN(rsRateDaFatt!IDPagamentoRata)
        rsAnagrafica.Filter = rsAnagrafica.Filter & " AND RitenutaAcconto=" & fnNotNullN(rsRateDaFatt!RitenutaAcconto)
        rsAnagrafica.Filter = rsAnagrafica.Filter & " AND IDRaggruppamentoFatturato=" & fnNotNullN(rsRateDaFatt!IDRaggruppamentoFatturato)
        rsAnagrafica.Filter = rsAnagrafica.Filter & " AND IDAccordoCommerciale=" & fnNotNullN(rsRateDaFatt!IDAccordoCommerciale)
        rsAnagrafica.Filter = rsAnagrafica.Filter & " AND IDContrattoBancario=" & fnNotNullN(rsRateDaFatt!IDContrattoBancario)
        rsAnagrafica.Filter = rsAnagrafica.Filter & " AND EntePubblico=" & fnNotNullN(rsRateDaFatt!EntePubblico)
        rsAnagrafica.Filter = rsAnagrafica.Filter & " AND IDSitoPerAnagrafica=" & fnNotNullN(rsRateDaFatt!IDSitoPerAnagrafica)
        rsAnagrafica.Filter = rsAnagrafica.Filter & " AND IDAnagraficaAgente=" & fnNotNullN(rsRateDaFatt!IDAnagraficaAgente)
        
        If rsAnagrafica.EOF Then
            rsAnagrafica.Filter = ""
            rsAnagrafica.AddNew
                rsAnagrafica!IDAnagrafica = fnNotNullN(rsRateDaFatt!IDAnagraficaFatturazione)
                If RaggruppamentoAnaContratto = 1 Then
                    rsAnagrafica!IDAnagraficaContratto = fnNotNullN(rsRateDaFatt!IDAnagrafica)
                End If
                rsAnagrafica!IDPagamento = fnNotNullN(rsRateDaFatt!IDPagamentoRata)
                rsAnagrafica!RitenutaAcconto = fnNotNullN(rsRateDaFatt!RitenutaAcconto)
                rsAnagrafica!IDRaggruppamentoFatturato = fnNotNullN(rsRateDaFatt!IDRaggruppamentoFatturato)
                rsAnagrafica!IDAccordoCommerciale = fnNotNullN(rsRateDaFatt!IDAccordoCommerciale)
                rsAnagrafica!IDContrattoBancario = fnNotNullN(rsRateDaFatt!IDContrattoBancario)
                rsAnagrafica!EntePubblico = fnNotNullN(rsRateDaFatt!EntePubblico)
                rsAnagrafica!IDSitoPerAnagrafica = fnNotNullN(rsRateDaFatt!IDSitoPerAnagrafica)
                rsAnagrafica!IDAnagraficaAgente = fnNotNullN(rsRateDaFatt!IDAnagraficaAgente)
                
            rsAnagrafica.Update
            NumeroRecordTesta = NumeroRecordTesta + 1
        End If
        
    rsRateDaFatt.MoveNext
    Wend
End Sub

Private Function GET_SEZIONALE_PA(IDSezionale As Long) As Long
On Error GoTo ERR_GET_SEZIONALE_PA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDRegIva As Long

IDRegIva = 0

sSQL = "SELECT IDSezionale, IDRegistroIva FROM Sezionale "
sSQL = sSQL & " WHERE IDSezionale=" & IDSezionale

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    IDRegIva = fnNotNullN(rs!IDRegistroIva)
End If

rs.CloseResultset
Set rs = Nothing

If IDRegIva = 1 Then
    cboSezionalePA.Enabled = True
    cboSezionalePA.WriteOn GET_LINK_SEZIONALE_PA
Else
    cboSezionalePA.Enabled = False
    cboSezionalePA.WriteOn 0
End If
Exit Function

ERR_GET_SEZIONALE_PA:
    MsgBox Err.Description, vbCritical, "GET_SEZIONALE_PA"
End Function
Private Function GET_LINK_SEZIONALE_PA() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSezionale FROM Sezionale "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND FatturaElettronica=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDRegistroIva=1"
sSQL = sSQL & " ORDER BY Sezionale"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_SEZIONALE_PA = 0
Else
    GET_LINK_SEZIONALE_PA = fnNotNullN(rs!IDSezionale)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ANAGRAFICA_ENTE_PUBBLICO(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT EntePubblico FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ANAGRAFICA_ENTE_PUBBLICO = 0
Else
    GET_ANAGRAFICA_ENTE_PUBBLICO = fnNotNullN(rs!EntePubblico)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub CboSezionale_Click()
    GET_SEZIONALE_PA Me.CboSezionale.CurrentID
End Sub
