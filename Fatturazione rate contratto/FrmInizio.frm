VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.9#0"; "DmtCodDesc.ocx"
Object = "{A83BB158-4E50-11D2-B95E-002018813989}#8.3#0"; "DmtSearchAccount.OCX"
Begin VB.Form FrmInizio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creazione documenti (Passo 1 di 4)"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Filtri"
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
      Height          =   5055
      Left            =   2520
      TabIndex        =   18
      Top             =   3360
      Width           =   5175
      Begin VB.ComboBox cboTipoFatt 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   4080
         Width           =   4935
      End
      Begin DmtSearchAccount.DmtSearchACS ACS 
         Height          =   585
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   1032
         WidthDescription=   3500
         WidthSecondDescription=   1300
         VisibleCode     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HideLeaf        =   0   'False
         BeginProperty FontLabel {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionDescription=   "Anagrafica di fatturazione"
         CaptionCode     =   ""
         OnlyAccounts    =   -1  'True
      End
      Begin DMTDataCmb.DMTCombo cboTipoContratto 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   4935
         _ExtentX        =   8705
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
      Begin DMTDATETIMELib.dmtDate txtAData 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtDate txtDaData 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DmtCodDescCtl.DmtCodDesc CDAmministratore 
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   3240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1085
         PropCodice      =   $"FrmInizio.frx":4781A
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"FrmInizio.frx":47872
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"FrmInizio.frx":478C2
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
      Begin DMTDataCmb.DMTCombo cboRaggrFatturato 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   4935
         _ExtentX        =   8705
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
      Begin DMTDataCmb.DMTCombo cboCateAnaCliente 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   4935
         _ExtentX        =   8705
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
      Begin VB.Label Label4 
         Caption         =   "Categoria anagrafica"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   2640
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "Raggruppamento fatturato"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo da fatturare"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   3840
         Width           =   4815
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo contratto"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Da data"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "A data"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8895
      Left            =   0
      Picture         =   "FrmInizio.frx":4791C
      ScaleHeight     =   8865
      ScaleWidth      =   2385
      TabIndex        =   17
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
      Left            =   2520
      TabIndex        =   6
      Top             =   0
      Width           =   5175
      Begin VB.Label Label1 
         Caption         =   "Azienda"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Attivita Azienda"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Filiale"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Esercizio"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label LblAzienda 
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label LblAttivitaAzienda 
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   4935
      End
      Begin VB.Label LblFiliale 
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   4935
      End
      Begin VB.Label LblEsercizio 
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   4935
      End
      Begin VB.Label Label3 
         Caption         =   "Utente"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   4095
      End
      Begin VB.Label LblUtente 
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   4935
      End
   End
   Begin VB.CommandButton CmdFine 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton CmdAvanti 
      Caption         =   "Avanti"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   8520
      Width           =   1335
   End
End
Attribute VB_Name = "FrmInizio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents m_App As DMTRunAppLib.Application
Attribute m_App.VB_VarHelpID = -1
Private rsTipoAddebito As ADODB.Recordset

Public Sub ConnessioneADO()
    If Not (CnDMT Is Nothing) Then
        CnDMT.CloseConnection
        Set CnDMT = Nothing
    End If
    
    Set CnDMT = m_App.Database.Connection
    
    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
End Sub
Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set m_App = NewValue
End Property
Public Property Get Application() As DMTRunAppLib.Application
    Set Application = m_App
End Property

Public Sub InitControlli()
    
    fncTipoContratto
    getTipoAddebito
    
    With ACS
        'Imposta la connessione attiva al controllo
        Set .Connection = m_App.Database.Connection
        'Imposta il nome dell'applicazione
        '.ApplicationName = m_App.
        'Imposta il nome dell'eseguibile dell'applicazione
        .Client = App.EXEName
        'Imposta l'identificativo dell'azienda corrente
        .IDFirm = TheApp.Branch
        'Imposta l'identificativo dell'utente corrente
        .IDUser = TheApp.IDUser
        '.UserName = m_App.User
        'Impostare con la proprietà Hwnd del form che contiene
        'il controllo. Serve per l'esegui gestione
        .HwndContainer = Me.hwnd
        
    End With

    'Anagrafica del tecnico di riferimento
    With Me.CDAmministratore
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Anagrafica"
        .DescriptionField = "Nome"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIEAnagraficaPerTipo"
        .Filter = "IDAzienda = " & TheApp.IDFirm & " AND IDTipoAnagrafica=" & GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDTipoAnagraficaAmministratore")
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Cognome"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Nome"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Cognome"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Nome"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 29 'Anagrafica
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    With Me.cboRaggrFatturato
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRaggruppamentoFatturato"
        .DisplayField = "RaggruppamentoFatturato"
        .Sql = "SELECT * FROM RaggruppamentoFatturato ORDER BY RaggruppamentoFatturato"
        .Fill
    End With

    With Me.cboCateAnaCliente
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDCategoriaAnagrafica"
        .DisplayField = "CategoriaAnagrafica"
        .Sql = "SELECT * FROM CategoriaAnagrafica WHERE IDTipoAnagrafica=2 ORDER BY CategoriaAnagrafica"
        .Fill
    End With

    Me.cboTipoFatt.ListIndex = 0
End Sub
Private Sub cmdAvanti_Click()
    VAR_DA_DATA = Me.txtDaData.Text
    VAR_A_DATA = Me.txtAData.Text
    LINK_TIPO_CONTRATTO = Me.cboTipoContratto.CurrentID
    LINK_CLIENTE = Me.ACS.IDAnagrafica
    LINK_AMMINISTRATORE = Me.CDAmministratore.KeyFieldID
    TIPO_SELEZIONE = Me.cboTipoFatt.ListIndex
    LINK_RAGGR_FATT_CLIENTE = Me.cboRaggrFatturato.CurrentID
    LINK_CAT_ANA_CLIENTE = Me.cboCateAnaCliente.CurrentID
    
    Unload Me
End Sub

Private Sub CmdFine_Click()
Dim Risposta As Integer
    Risposta = MsgBox("Vuoi abbandonare il wizard per il passaggio degli interventi in fatturazione?", vbInformation + vbYesNo, "Abbandono")
        
    If Risposta = vbYes Then
        Unload Me
    End If

End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    LINK_CONTRATTO_SELEZIONATO = GetSetting(Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), App.EXEName, "IDContratto", 0)
    
    If LINK_CONTRATTO_SELEZIONATO > 0 Then
        cmdAvanti.Value = True
        'cmdAvanti_Click
    End If
    
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load

'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)

If b_Loading = True Then
    InitControlli
    PrelevaAzienda
End If


Exit Sub

ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"

End Sub
Private Sub fncTipoContratto()
    With Me.cboTipoContratto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoContratto"
        .DisplayField = "TipoContratto"
        .Sql = "SELECT * FROM RV_POTipoContratto"
        .Fill
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERR_Form_Unload
Dim Risposta As Integer

    If Me.cmdAvanti.Value = True Then
        FrmMovimenti.Show
    End If

Exit Sub

ERR_Form_Unload:
    MsgBox Err.Description, vbCritical, "Form_Unload"
End Sub


Private Function GET_PARAMETRO_AZIENDA_LONG(IDFiliale As Long, NomeCampo As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_POParametriAzienda "
sSQL = sSQL & " WHERE IDFiliale=" & IDFiliale

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PARAMETRO_AZIENDA_LONG = 0
Else
    GET_PARAMETRO_AZIENDA_LONG = fnNotNullN(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub getTipoAddebito()

cboTipoFatt.Clear

cboTipoFatt.AddItem "Fattura tutto"
cboTipoFatt.AddItem "Fattura solamente le rate del contratto"
cboTipoFatt.AddItem "Fattura solamente le eccedenze"






End Sub
