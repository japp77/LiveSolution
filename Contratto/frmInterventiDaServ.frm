VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmInterventiDaServ 
   Caption         =   "INTERVENTI DA CONTRATTO"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17070
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterventiDaServ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   17070
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   1200
      TabIndex        =   28
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393217
      TextRTF         =   $"frmInterventiDaServ.frx":4781A
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   5
      Top             =   8400
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdGenera 
      Caption         =   "GENERA"
      Height          =   495
      Left            =   14880
      TabIndex        =   1
      Top             =   8040
      Width           =   2175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   13996
      _Version        =   393216
      Tab             =   2
      TabHeight       =   706
      TabCaption(0)   =   "Interventi non modificabili"
      TabPicture(0)   =   "frmInterventiDaServ.frx":4789F
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Griglia1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Interventi da eliminare"
      TabPicture(1)   =   "frmInterventiDaServ.frx":478BB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Griglia2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Interventi da generare"
      TabPicture(2)   =   "frmInterventiDaServ.frx":478D7
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Griglia3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Proprietà intervento"
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
         Height          =   7215
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   3735
         Begin VB.CheckBox Check1 
            Caption         =   "Inserisci l'operatore se configurato nei prodotti del contratto"
            Height          =   495
            Left            =   120
            TabIndex        =   27
            Top             =   5520
            Width           =   3495
         End
         Begin VB.CheckBox chkEliminaSolamente 
            Caption         =   "Elimina solamente"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   5160
            Width           =   3495
         End
         Begin VB.TextBox txtTipoAddebito 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   3480
            Width           =   3495
         End
         Begin VB.TextBox txtTecnicOpe 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox txtCategoria 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   4680
            Width           =   3495
         End
         Begin VB.TextBox txtClasse 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   4080
            Width           =   3495
         End
         Begin VB.TextBox txtTipoIntervento 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   2280
            Width           =   3495
         End
         Begin VB.TextBox txtStatoIntervento 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   2880
            Width           =   3495
         End
         Begin VB.TextBox txtRiferimentoInt 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   1080
            Width           =   3495
         End
         Begin VB.TextBox txtCliente 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   480
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Categoria"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   22
            Top             =   4440
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Classe"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   20
            Top             =   3840
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo addebito"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   19
            Top             =   3240
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo intervento"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   17
            Top             =   2040
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Stato intervento"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   2640
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Tecnico operativo"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Riferimento interno"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   3495
         End
      End
      Begin DmtGridCtl.DmtGrid Griglia1 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   2
         Top             =   600
         Width           =   16815
         _ExtentX        =   29660
         _ExtentY        =   12515
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
      Begin DmtGridCtl.DmtGrid Griglia2 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   3
         Top             =   600
         Width           =   16815
         _ExtentX        =   29660
         _ExtentY        =   12515
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
      Begin DmtGridCtl.DmtGrid Griglia3 
         Height          =   7095
         Left            =   3960
         TabIndex        =   4
         Top             =   600
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   12515
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
   End
   Begin VB.Label Label1 
      Caption         =   "Stato intervento"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   8160
      Width           =   14535
   End
End
Attribute VB_Name = "frmInterventiDaServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia1 As ADODB.Recordset
Private rsGriglia2 As ADODB.Recordset
Private rsGriglia3 As ADODB.Recordset

Private rsGrigliaServizi As ADODB.Recordset

Private rsInterventoUpdate As ADODB.Recordset

Private NUMERO_INTERVENTI_DA_ELIMINARE
Private NUMERO_INTERVENTI_DA_NON_ELABORARE
Private NUMERO_INTERVENTI_DA_CREARE

Private LINK_CLIENTE_LOCAL As Long
Private LINK_RIFERIMENTO_INT_LOCAL As Long
Private LINK_TECNICO_OPERATIVO_LOCAL As Long
Private LINK_STATO_LOCAL As Long
Private LINK_TIPO_LOCAL As Long
Private LINK_TIPO_ADDEBITO_LOCAL As Long
Private LINK_CLASSE_LOCAL As Long
Private LINK_CATEGORIA_LOCAL As Long
Private LINK_TIPO_ANA_TEC_RIF As Long
Private LINK_TIPO_ANA_TEC_OPE As Long
Private Sub cmdGenera_Click()
On Error GoTo ERR_cmdGenera_Click
Dim sSQL As String
Dim UnitaProgresso As Double
Dim rsIntervento As ADODB.Recordset
Dim link_riga_servizio As Long
Dim NumeroIntervento As Long

ELIMINA_INTERVENTI

If chkEliminaSolamente.Value = vbChecked Then
    Unload Me
    Exit Sub
End If

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100
 
UnitaProgresso = FormatNumber((Me.ProgressBar1.Max / NUMERO_INTERVENTI_DA_CREARE), 4)

Me.lblInfo.Caption = "CREAZIONE INTERVENTI IN CORSO..."
DoEvents

'''''AGGIORNAMENTO INTERVENTO PADRE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Not (rsInterventoUpdate Is Nothing) Then
    If rsInterventoUpdate.State > 0 Then
        rsInterventoUpdate.Close
    End If
    
    Set rsInterventoUpdate = Nothing
End If

Set rsInterventoUpdate = New ADODB.Recordset
rsInterventoUpdate.Fields.Append "IDIntervento", adInteger, , adFldIsNullable

rsInterventoUpdate.Open , , adOpenKeyset, adLockBatchOptimistic
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


rsGriglia3.Filter = "Registra=" & fnNormBoolean(True)

If ((rsGriglia3.EOF) And (rsGriglia3.BOF)) Then Exit Sub

rsGriglia3.Sort = "IDRV_POContrattoServizi"

link_riga_servizio = 0

sSQL = "SELECT * FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=0"

Set rsIntervento = New ADODB.Recordset

rsIntervento.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

link_riga_servizio = 0
NumeroIntervento = 0

While Not rsGriglia3.EOF
    
'    If link_riga_servizio <> fnNotNullN(rsGriglia3!IDRV_POContrattoServizi) Then
'        'NumeroIntervento = GET_NUMERO_INTERVENTO(Year(Date))
'
'        link_riga_servizio = fnNotNullN(rsGriglia3!IDRV_POContrattoServizi)
'
'
'    End If
    
    If fnNotNullN(rsGriglia3!NumeroInterventoSub) = 1 Then
        NumeroIntervento = GET_NUMERO_INTERVENTO(Year(rsGriglia3!DataAppuntamento))
    End If
    
    CREA_INTERVENTO rsIntervento, NumeroIntervento
    
    If fnNotNullN(rsIntervento!IDRV_POIntervento) > 0 Then
        rsInterventoUpdate.AddNew
            rsInterventoUpdate!IDIntervento = fnNotNullN(rsIntervento!IDRV_POIntervento)
        rsInterventoUpdate.Update
    End If
    
    
    If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
    End If
    
    DoEvents
rsGriglia3.MoveNext
Wend

rsIntervento.Close
Set rsIntervento = Nothing

If Not ((rsInterventoUpdate.EOF) And (rsInterventoUpdate.BOF)) Then
    rsInterventoUpdate.MoveFirst
    
    While Not rsInterventoUpdate.EOF
        sSQL = "UPDATE RV_POIntervento SET "
        sSQL = sSQL & " IDRV_POInterventoPadre=" & rsInterventoUpdate!IDIntervento
        sSQL = sSQL & " WHERE IDRV_POIntervento=" & rsInterventoUpdate!IDIntervento
        Cn.Execute sSQL
    rsInterventoUpdate.MoveNext
    Wend
End If


Unload Me
Exit Sub
ERR_cmdGenera_Click:
    MsgBox Err.Description, vbCritical, "cmdGenera_Click"
End Sub
Private Sub Form_Activate()

    NUMERO_INTERVENTI_DA_NON_ELABORARE = 0
    NUMERO_INTERVENTI_DA_CREARE = 0
    NUMERO_INTERVENTI_DA_ELIMINARE = 0
    
    CREA_RECORDSET_SERVIZI Elabora_Tutti_Servizi_Contratto
    
    GET_DATI_RECORDSET_1
    GET_DATI_RECORDSET_2
    GET_DATI_RECORDSET_3
    
    GET_GRIGLIA_1
    GET_GRIGLIA_2
    GET_GRIGLIA_3
    
    Me.lblInfo.Caption = ""
    Me.ProgressBar1.Value = 0
    
    SSTab1.TabCaption(0) = SSTab1.TabCaption(0) & " (" & NUMERO_INTERVENTI_DA_NON_ELABORARE & ")"
    SSTab1.TabCaption(1) = SSTab1.TabCaption(1) & " (" & NUMERO_INTERVENTI_DA_ELIMINARE & ")"
    SSTab1.TabCaption(2) = SSTab1.TabCaption(2) & " (" & NUMERO_INTERVENTI_DA_CREARE & ")"
    
End Sub

Private Sub Form_Load()

    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    LINK_CLIENTE_LOCAL = frmMain.CDCliente.KeyFieldID
    LINK_RIFERIMENTO_INT_LOCAL = frmMain.CDTecnico.KeyFieldID
    LINK_TECNICO_OPERATIVO_LOCAL = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDAnaTecOpeDaProdotto")
    If LINK_TECNICO_OPERATIVO_LOCAL = 0 Then LINK_TECNICO_OPERATIVO_LOCAL = LINK_RIFERIMENTO_INT_LOCAL
    
    LINK_TIPO_ANA_TEC_RIF = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDTipoAnagraficaTecnicoIntRif")
    LINK_TIPO_ANA_TEC_OPE = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDTipoAnagraficaTecnicoFaseRif")
    
    LINK_STATO_LOCAL = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDRV_POStatoInterventoInserimento")
    LINK_TIPO_LOCAL = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDRV_POTipoFaseInterventoEla")
    LINK_TIPO_ADDEBITO_LOCAL = GET_PARAMETRI_TEC_OPE(LINK_TECNICO_OPERATIVO_LOCAL, "IDRV_POTipoAddebito")
    LINK_CLASSE_LOCAL = GET_PARAMETRI_TEC_OPE(LINK_TECNICO_OPERATIVO_LOCAL, "IDRV_POTipoClasseIntervento")
    LINK_CATEGORIA_LOCAL = GET_PARAMETRI_TEC_OPE(LINK_TECNICO_OPERATIVO_LOCAL, "IDRV_POCategoriaFase")


    Me.txtCliente.Text = GET_ANAGRAFICA(LINK_CLIENTE_LOCAL)
    Me.txtRiferimentoInt.Text = GET_ANAGRAFICA(LINK_RIFERIMENTO_INT_LOCAL)
    Me.txtTecnicOpe.Text = GET_ANAGRAFICA(LINK_TECNICO_OPERATIVO_LOCAL)
    
    Me.txtTipoIntervento.Text = GET_DESCRIZIONE_TABELLA("RV_POTipoFaseIntervento", "TipoFaseIntervento", "IDRV_POTipoFaseIntervento", LINK_TIPO_LOCAL)
    Me.txtStatoIntervento.Text = GET_DESCRIZIONE_TABELLA("RV_POStatoIntervento", "StatoIntervento", "IDRV_POStatoIntervento", LINK_STATO_LOCAL)
    Me.txtCategoria.Text = GET_DESCRIZIONE_TABELLA("RV_POCategoriaFase", "CategoriaFase", "IDRV_POCategoriaFase", LINK_CATEGORIA_LOCAL)
    Me.txtTipoAddebito.Text = GET_DESCRIZIONE_TABELLA("RV_POTipoAddebito", "TipoAddebito", "IDRV_POTipoAddebito", LINK_TIPO_ADDEBITO_LOCAL)
    Me.txtClasse.Text = GET_DESCRIZIONE_TABELLA("RV_POTipoClasseIntervento", "TipoClasseIntervento", "IDRV_POTipoClasseIntervento", LINK_CLASSE_LOCAL)
    
    

    CREA_RECORSET_1
    CREA_RECORSET_2
    CREA_RECORSET_3
    

End Sub

Private Sub CREA_RECORSET_1()
Dim sSQL As String
Dim rs As ADODB.Recordset

Set rsGriglia1 = Nothing

Set rsGriglia1 = New ADODB.Recordset

rsGriglia1.CursorLocation = adUseClient

'INTERVENTO
rsGriglia1.Fields.Append "IDRV_POIntervento", adInteger, , adFldIsNullable
rsGriglia1.Fields.Append "NumeroIntervento", adInteger, , adFldIsNullable
rsGriglia1.Fields.Append "AnnoIntervento", adInteger, , adFldIsNullable
rsGriglia1.Fields.Append "NumeroInterventoSub", adInteger, , adFldIsNullable
rsGriglia1.Fields.Append "FaseIntervento", adInteger, , adFldIsNullable
'CONTRATTO
rsGriglia1.Fields.Append "IDRV_POContratto", adInteger, , adFldIsNullable
rsGriglia1.Fields.Append "AnnoContratto", adInteger, , adFldIsNullable
rsGriglia1.Fields.Append "NumeroContratto", adInteger, , adFldIsNullable
rsGriglia1.Fields.Append "IDRV_POContrattoPadre", adInteger, , adFldIsNullable

'SERVIZIO
rsGriglia1.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsGriglia1.Fields.Append "CodiceServizio", adVarChar, 50, adFldIsNullable
rsGriglia1.Fields.Append "DescrizioneServizio", adVarChar, 250, adFldIsNullable
'PRODOTTO
rsGriglia1.Fields.Append "IDRV_POProdotto", adInteger, , adFldIsNullable
rsGriglia1.Fields.Append "DescrizioneProdotto", adVarChar, 250, adFldIsNullable
rsGriglia1.Fields.Append "Matricola", adVarChar, 250, adFldIsNullable

'RIFERIMENTI CLIENTE
rsGriglia1.Fields.Append "IDAnagraficaCliente", adInteger, , adFldIsNullable
rsGriglia1.Fields.Append "AnagraficaCliente", adVarChar, 250, adFldIsNullable

'RIFERIMENTI RIGHE DEL CONTRATTO
rsGriglia1.Fields.Append "IDRV_POContrattoServizi", adInteger, , adFldIsNullable
rsGriglia1.Fields.Append "IDRV_POContrattoProdotti", adInteger, , adFldIsNullable

'TECNICO DI RIFERIMENTO
rsGriglia1.Fields.Append "IDAnagraficaTecnicoRif", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia1.Fields.Append "AnagraficaTecnicoRif", adVarChar, 250, adFldIsNullable

'TECNICO OPERATIVO
rsGriglia1.Fields.Append "IDAnagraficaTecnicoOperativo", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia1.Fields.Append "AnagraficaTecnicoOperativo", adVarChar, 250, adFldIsNullable

'DATA APPUNTAMENTO
rsGriglia1.Fields.Append "DataAppuntamento", adDBDate, , adFldIsNullable 'Tecnico di riferimento
rsGriglia1.Fields.Append "OraAppuntamento", adVarChar, 250, adFldIsNullable

'ALTRI DATI UTILI DA VISUALIZZARE
rsGriglia1.Fields.Append "Manuale", adBoolean, , adFldIsNullable 'Tecnico di riferimento
rsGriglia1.Fields.Append "Elaborato", adBoolean, , adFldIsNullable
rsGriglia1.Fields.Append "InterventoChiuso", adSmallInt, , adFldIsNullable

'CATEGORIA INTERVENTO
rsGriglia1.Fields.Append "IDRV_POCategoriaIntervento", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia1.Fields.Append "CategoriaIntervento", adVarChar, 250, adFldIsNullable

'TIPO ADDEBITO INTERVENTO
rsGriglia1.Fields.Append "IDRV_POTipoAddebito", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia1.Fields.Append "TipoAddebitoIntervento", adVarChar, 250, adFldIsNullable

'CLASSE INTERVENTO
rsGriglia1.Fields.Append "IDRV_POTipoClasseIntervento", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia1.Fields.Append "ClasseIntervento", adVarChar, 250, adFldIsNullable

'STATO INTERVENTO
rsGriglia1.Fields.Append "IDRV_POStatoIntervento", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia1.Fields.Append "StatoIntervento", adVarChar, 250, adFldIsNullable

'TIPO FASE INTERVENTO
rsGriglia1.Fields.Append "IDRV_POTipoFaseIntervento", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia1.Fields.Append "TipoFaseIntervento", adVarChar, 250, adFldIsNullable


rsGriglia1.Open , , adOpenKeyset, adLockPessimistic

End Sub
Private Sub CREA_RECORSET_2()
Dim sSQL As String
Dim rs As ADODB.Recordset

Set rsGriglia2 = Nothing

Set rsGriglia2 = New ADODB.Recordset

rsGriglia2.CursorLocation = adUseClient

'INTERVENTO
rsGriglia2.Fields.Append "IDRV_POIntervento", adInteger, , adFldIsNullable
rsGriglia2.Fields.Append "NumeroIntervento", adInteger, , adFldIsNullable
rsGriglia2.Fields.Append "AnnoIntervento", adInteger, , adFldIsNullable
rsGriglia2.Fields.Append "NumeroInterventoSub", adInteger, , adFldIsNullable
rsGriglia2.Fields.Append "FaseIntervento", adInteger, , adFldIsNullable
'CONTRATTO
rsGriglia2.Fields.Append "IDRV_POContratto", adInteger, , adFldIsNullable
rsGriglia2.Fields.Append "AnnoContratto", adInteger, , adFldIsNullable
rsGriglia2.Fields.Append "NumeroContratto", adInteger, , adFldIsNullable
rsGriglia2.Fields.Append "IDRV_POContrattoPadre", adInteger, , adFldIsNullable

'SERVIZIO
rsGriglia2.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsGriglia2.Fields.Append "CodiceServizio", adVarChar, 50, adFldIsNullable
rsGriglia2.Fields.Append "DescrizioneServizio", adVarChar, 250, adFldIsNullable
'PRODOTTO
rsGriglia2.Fields.Append "IDRV_POProdotto", adInteger, , adFldIsNullable
rsGriglia2.Fields.Append "DescrizioneProdotto", adVarChar, 250, adFldIsNullable
rsGriglia2.Fields.Append "Matricola", adVarChar, 250, adFldIsNullable

'RIFERIMENTI CLIENTE
rsGriglia2.Fields.Append "IDAnagraficaCliente", adInteger, , adFldIsNullable
rsGriglia2.Fields.Append "AnagraficaCliente", adVarChar, 250, adFldIsNullable


'RIFERIMENTI RIGHE DEL CONTRATTO
rsGriglia2.Fields.Append "IDRV_POContrattoServizi", adInteger, , adFldIsNullable
rsGriglia2.Fields.Append "IDRV_POContrattoProdotti", adInteger, , adFldIsNullable

'TECNICO DI RIFERIMENTO
rsGriglia2.Fields.Append "IDAnagraficaTecnicoRif", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia2.Fields.Append "AnagraficaTecnicoRif", adVarChar, 250, adFldIsNullable

'TECNICO OPERATIVO
rsGriglia2.Fields.Append "IDAnagraficaTecnicoOperativo", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia2.Fields.Append "AnagraficaTecnicoOperativo", adVarChar, 250, adFldIsNullable

'DATA APPUNTAMENTO
rsGriglia2.Fields.Append "DataAppuntamento", adDBDate, , adFldIsNullable 'Tecnico di riferimento
rsGriglia2.Fields.Append "OraAppuntamento", adVarChar, 250, adFldIsNullable

'ALTRI DATI UTILI DA VISUALIZZARE
rsGriglia2.Fields.Append "Manuale", adBoolean, , adFldIsNullable 'Tecnico di riferimento
rsGriglia2.Fields.Append "Elaborato", adBoolean, , adFldIsNullable
rsGriglia2.Fields.Append "InterventoChiuso", adSmallInt, , adFldIsNullable

'CATEGORIA INTERVENTO
rsGriglia2.Fields.Append "IDRV_POCategoriaIntervento", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia2.Fields.Append "CategoriaIntervento", adVarChar, 250, adFldIsNullable

'TIPO ADDEBITO INTERVENTO
rsGriglia2.Fields.Append "IDRV_POTipoAddebito", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia2.Fields.Append "TipoAddebitoIntervento", adVarChar, 250, adFldIsNullable

'CLASSE INTERVENTO
rsGriglia2.Fields.Append "IDRV_POTipoClasseIntervento", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia2.Fields.Append "ClasseIntervento", adVarChar, 250, adFldIsNullable

'STATO INTERVENTO
rsGriglia2.Fields.Append "IDRV_POStatoIntervento", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia2.Fields.Append "StatoIntervento", adVarChar, 250, adFldIsNullable

'TIPO FASE INTERVENTO
rsGriglia2.Fields.Append "IDRV_POTipoFaseIntervento", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia2.Fields.Append "TipoFaseIntervento", adVarChar, 250, adFldIsNullable

rsGriglia2.Fields.Append "Elimina", adBoolean, , adFldIsNullable
rsGriglia2.Fields.Append "EliminaObbligatorio", adBoolean, , adFldIsNullable

rsGriglia2.Open , , adOpenKeyset, adLockPessimistic

End Sub

Private Sub CREA_RECORSET_3()
Dim sSQL As String
Dim rs As ADODB.Recordset

Set rsGriglia3 = Nothing

Set rsGriglia3 = New ADODB.Recordset

rsGriglia3.CursorLocation = adUseClient

'INTERVENTO
rsGriglia3.Fields.Append "IDRV_POIntervento", adInteger, , adFldIsNullable
rsGriglia3.Fields.Append "NumeroIntervento", adInteger, , adFldIsNullable
rsGriglia3.Fields.Append "AnnoIntervento", adInteger, , adFldIsNullable
rsGriglia3.Fields.Append "NumeroInterventoSub", adInteger, , adFldIsNullable
rsGriglia3.Fields.Append "FaseIntervento", adInteger, , adFldIsNullable

'CONTRATTO
rsGriglia3.Fields.Append "IDRV_POContratto", adInteger, , adFldIsNullable
rsGriglia3.Fields.Append "AnnoContratto", adInteger, , adFldIsNullable
rsGriglia3.Fields.Append "NumeroContratto", adInteger, , adFldIsNullable
rsGriglia3.Fields.Append "IDRV_POContrattoPadre", adInteger, , adFldIsNullable

'SERVIZIO
rsGriglia3.Fields.Append "IDArticolo", adInteger, , adFldIsNullable
rsGriglia3.Fields.Append "CodiceServizio", adVarChar, 50, adFldIsNullable
rsGriglia3.Fields.Append "DescrizioneServizio", adVarChar, 250, adFldIsNullable
'PRODOTTO
rsGriglia3.Fields.Append "IDRV_POProdotto", adInteger, , adFldIsNullable
rsGriglia3.Fields.Append "DescrizioneProdotto", adVarChar, 250, adFldIsNullable
rsGriglia3.Fields.Append "Matricola", adVarChar, 250, adFldIsNullable

'RIFERIMENTI CLIENTE
rsGriglia3.Fields.Append "IDAnagraficaCliente", adInteger, , adFldIsNullable
rsGriglia3.Fields.Append "AnagraficaCliente", adVarChar, 250, adFldIsNullable

'RIFERIMENTI RIGHE DEL CONTRATTO
rsGriglia3.Fields.Append "IDRV_POContrattoServizi", adInteger, , adFldIsNullable
rsGriglia3.Fields.Append "IDRV_POContrattoProdotti", adInteger, , adFldIsNullable

'TECNICO DI RIFERIMENTO
rsGriglia3.Fields.Append "IDAnagraficaTecnicoRif", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia3.Fields.Append "AnagraficaTecnicoRif", adVarChar, 250, adFldIsNullable

'TECNICO OPERATIVO
rsGriglia3.Fields.Append "IDAnagraficaTecnicoOperativo", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia3.Fields.Append "AnagraficaTecnicoOperativo", adVarChar, 250, adFldIsNullable

'DATA APPUNTAMENTO
rsGriglia3.Fields.Append "DataAppuntamento", adDBDate, , adFldIsNullable 'Tecnico di riferimento
rsGriglia3.Fields.Append "OraAppuntamento", adVarChar, 250, adFldIsNullable

'ALTRI DATI UTILI DA VISUALIZZARE
rsGriglia3.Fields.Append "Manuale", adBoolean, , adFldIsNullable 'Tecnico di riferimento
rsGriglia3.Fields.Append "Elaborato", adBoolean, , adFldIsNullable
rsGriglia3.Fields.Append "InterventoChiuso", adSmallInt, , adFldIsNullable

'CATEGORIA INTERVENTO
rsGriglia3.Fields.Append "IDRV_POCategoriaIntervento", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia3.Fields.Append "CategoriaIntervento", adVarChar, 250, adFldIsNullable

'TIPO ADDEBITO INTERVENTO
rsGriglia3.Fields.Append "IDRV_POTipoAddebito", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia3.Fields.Append "TipoAddebitoIntervento", adVarChar, 250, adFldIsNullable

'CLASSE INTERVENTO
rsGriglia3.Fields.Append "IDRV_POTipoClasseIntervento", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia3.Fields.Append "ClasseIntervento", adVarChar, 250, adFldIsNullable

'STATO INTERVENTO
rsGriglia3.Fields.Append "IDRV_POStatoIntervento", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia3.Fields.Append "StatoIntervento", adVarChar, 250, adFldIsNullable

'TIPO FASE INTERVENTO
rsGriglia3.Fields.Append "IDRV_POTipoFaseIntervento", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia3.Fields.Append "TipoFaseIntervento", adVarChar, 250, adFldIsNullable

rsGriglia3.Fields.Append "Elimina", adBoolean, , adFldIsNullable
rsGriglia3.Fields.Append "EliminaObbligatorio", adBoolean, , adFldIsNullable
rsGriglia3.Fields.Append "Registra", adBoolean, , adFldIsNullable

rsGriglia3.Open , , adOpenKeyset, adLockPessimistic

End Sub
Private Sub GET_DATI_RECORDSET_1()
On Error GoTo ERR_GET_DATI_RECORDSET_1
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim UnitaProgresso As Double
Dim NumeroRecord As Long

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100

NumeroRecord = 0

If ((rsGrigliaServizi.EOF) And (rsGrigliaServizi.BOF)) Then Exit Sub

rsGrigliaServizi.MoveFirst

While Not rsGrigliaServizi.EOF
    Link_Contratto_Servizio = rsGrigliaServizi!IDRV_POContrattoServizi
    
    sSQL = "SELECT COUNT(IDRV_POIntervento) AS NumeroRecord "
    sSQL = sSQL & " FROM RV_POIEIntervento "
    sSQL = sSQL & "WHERE IDRV_POContrattoServizi=" & Link_Contratto_Servizio
    sSQL = sSQL & " AND Elaborata=1"
    sSQL = sSQL & " AND Manuale=1"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        NumeroRecord = NumeroRecord + fnNotNullN(rs!NumeroRecord)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
rsGrigliaServizi.MoveNext
Wend


If NumeroRecord = 0 Then Exit Sub

NUMERO_INTERVENTI_DA_NON_ELABORARE = NumeroRecord


UnitaProgresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 4)
Me.lblInfo.Caption = "Interventi non modificabili..."
DoEvents


If ((rsGrigliaServizi.EOF) And (rsGrigliaServizi.BOF)) Then Exit Sub

rsGrigliaServizi.MoveFirst

While Not rsGrigliaServizi.EOF

    Link_Contratto_Servizio = rsGrigliaServizi!IDRV_POContrattoServizi

    sSQL = "SELECT * FROM RV_POIEIntervento "
    sSQL = sSQL & "WHERE IDRV_POContrattoServizi=" & Link_Contratto_Servizio
    sSQL = sSQL & " AND Elaborata=1"
    sSQL = sSQL & " AND Manuale=1"
    'sSQL = sSQL & " AND DataAppuntamento<" & fnNormDate(Date)
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        rsGriglia1.AddNew
            rsGriglia1!IDRV_POIntervento = fnNotNullN(rs!IDRV_POIntervento)
            rsGriglia1!NumeroIntervento = fnNotNullN(rs!NumeroIntervento)
            rsGriglia1!AnnoIntervento = fnNotNullN(rs!AnnoIntervento)
            rsGriglia1!NumeroInterventoSub = fnNotNullN(rs!NumeroInterventoSub)
            rsGriglia1!FaseIntervento = fnNotNullN(rs!NumeroFase)
            rsGriglia1!IDRV_POContratto = fnNotNullN(rs!IDRV_POContratto)
            rsGriglia1!AnnoContratto = fnNotNullN(rs!AnnoContratto)
            rsGriglia1!NumeroContratto = fnNotNullN(rs!NumeroContratto)
            rsGriglia1!IDRV_POContrattoPadre = fnNotNullN(rs!IDRV_POContrattoPadre)
            
            rsGriglia1!IDArticolo = fnNotNullN(rs!IDArticolo)
            rsGriglia1!CodiceServizio = fnNotNull(rs!CodiceArticolo)
            rsGriglia1!DescrizioneServizio = fnNotNull(rs!Articolo)
            rsGriglia1!IDRV_POProdotto = fnNotNullN(rs!IDRV_POProdotto)
            rsGriglia1!DescrizioneProdotto = fnNotNull(rs!DescrizioneProdotto)
            rsGriglia1!Matricola = fnNotNull(rs!MatricolaProdotto)
            rsGriglia1!IDAnagraficaCliente = fnNotNullN(rs!IDAnagraficaCliente)
            rsGriglia1!AnagraficaCliente = fnNotNull(rs!AnagraficaCliente) & " " & fnNotNull(rs!NomeCliente)
            rsGriglia1!IDRV_POContrattoServizi = fnNotNullN(rs!IDRV_POContrattoServizi)
            rsGriglia1!IDRV_POContrattoProdotti = fnNotNullN(rs!IDRV_POContrattoProdotti)
            rsGriglia1!IDAnagraficaTecnicoRif = fnNotNullN(rs!IDAnagraficaTecnicoRif)
            rsGriglia1!AnagraficaTecnicoRif = fnNotNull(rs!AnagraficaTecnico) & " " & fnNotNull(rs!NomeTecnico)
            rsGriglia1!IDAnagraficaTecnicoOperativo = fnNotNullN(rs!IDAnagraficaTecnicoOperativo)
            rsGriglia1!AnagraficaTecnicoOperativo = fnNotNull(rs!AnagraficaTecnicoOperativo) & " " & fnNotNull(rs!NomeTecnicoOperativo)
            rsGriglia1!DataAppuntamento = rs!DataAppuntamento
            rsGriglia1!OraAppuntamento = fnNotNull(OraAppuntamento)
            rsGriglia1!Manuale = rs!Manuale
            rsGriglia1!Elaborato = rs!Elaborata
            rsGriglia1!InterventoChiuso = fnNotNullN(rs!InterventoChiuso)
            rsGriglia1!IDRV_POCategoriaIntervento = fnNotNullN(rs!IDRV_POCategoriaIntervento)
            rsGriglia1!CategoriaIntervento = fnNotNull(rs!CategoriaIntervento)
            rsGriglia1!IDRV_POTipoAddebito = fnNotNullN(rs!IDRV_POTipoAddebito)
            rsGriglia1!TipoAddebitoIntervento = fnNotNull(rs!TipoAddebito)
            rsGriglia1!IDRV_POTipoClasseIntervento = fnNotNullN(rs!IDRV_POTipoClasseIntervento)
            rsGriglia1!ClasseIntervento = fnNotNull(rs!TipoClasseIntervento)
            rsGriglia1!IDRV_POStatoIntervento = fnNotNullN(rs!IDRV_POStatoIntervento)
            rsGriglia1!StatoIntervento = fnNotNull(rs!StatoIntervento)
            rsGriglia1!IDRV_POTipoFaseIntervento = fnNotNullN(rs!IDRV_POTipoFaseIntervento)
            rsGriglia1!TipoFaseIntervento = fnNotNull(rs!TipoFaseIntervento)
    
        rsGriglia1.Update
        
        If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
            Me.ProgressBar1.Value = Me.ProgressBar1.Max
        Else
            Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
        End If
        
        DoEvents
    rs.MoveNext
    Wend
    
rsGrigliaServizi.MoveNext
Wend
    
    
rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_GET_DATI_RECORDSET_1:
    MsgBox Err.Description, vbCritical, "GET_DATI_RECORDSET_1"
End Sub
Private Sub GET_DATI_RECORDSET_2()
On Error GoTo ERR_GET_DATI_RECORDSET_2
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim UnitaProgresso As Double
Dim NumeroRecord As Long

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100

NumeroRecord = 0

If ((rsGrigliaServizi.EOF) And (rsGrigliaServizi.BOF)) Then Exit Sub

rsGrigliaServizi.MoveFirst

While Not rsGrigliaServizi.EOF
    Link_Contratto_Servizio = rsGrigliaServizi!IDRV_POContrattoServizi
    
    sSQL = "SELECT COUNT(IDRV_POIntervento) AS NumeroRecord "
    sSQL = sSQL & " FROM RV_POIEIntervento "
    sSQL = sSQL & "WHERE IDRV_POContrattoServizi=" & Link_Contratto_Servizio
    sSQL = sSQL & " AND Elaborata=1"
    sSQL = sSQL & " AND Manuale=0"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        NumeroRecord = NumeroRecord + fnNotNullN(rs!NumeroRecord)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
rsGrigliaServizi.MoveNext
Wend


If NumeroRecord = 0 Then Exit Sub

UnitaProgresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 4)
Me.lblInfo.Caption = "Interventi da eliminare..."
DoEvents

NUMERO_INTERVENTI_DA_ELIMINARE = 0

If ((rsGrigliaServizi.EOF) And (rsGrigliaServizi.BOF)) Then Exit Sub

rsGrigliaServizi.MoveFirst

While Not rsGrigliaServizi.EOF

    Link_Contratto_Servizio = rsGrigliaServizi!IDRV_POContrattoServizi
    
    sSQL = "SELECT * FROM RV_POIEIntervento "
    sSQL = sSQL & "WHERE IDRV_POContrattoServizi=" & Link_Contratto_Servizio
    sSQL = sSQL & " AND Elaborata=1"
    sSQL = sSQL & " AND Manuale=0"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        rsGriglia2.AddNew
            rsGriglia2!IDRV_POIntervento = fnNotNullN(rs!IDRV_POIntervento)
            rsGriglia2!NumeroIntervento = fnNotNullN(rs!NumeroIntervento)
            rsGriglia2!AnnoIntervento = fnNotNullN(rs!AnnoIntervento)
            rsGriglia2!NumeroInterventoSub = fnNotNullN(rs!NumeroInterventoSub)
            rsGriglia2!FaseIntervento = fnNotNullN(rs!NumeroFase)
            rsGriglia2!IDRV_POContratto = fnNotNullN(rs!IDRV_POContratto)
            rsGriglia2!AnnoContratto = fnNotNullN(rs!AnnoContratto)
            rsGriglia2!NumeroContratto = fnNotNullN(rs!NumeroContratto)
            rsGriglia2!IDRV_POContrattoPadre = fnNotNullN(rs!IDRV_POContrattoPadre)
            
            rsGriglia2!IDArticolo = fnNotNullN(rs!IDArticolo)
            rsGriglia2!CodiceServizio = fnNotNull(rs!CodiceArticolo)
            rsGriglia2!DescrizioneServizio = fnNotNull(rs!Articolo)
            rsGriglia2!IDRV_POProdotto = fnNotNullN(rs!IDRV_POProdotto)
            rsGriglia2!DescrizioneProdotto = fnNotNull(rs!DescrizioneProdotto)
            rsGriglia2!Matricola = fnNotNull(rs!MatricolaProdotto)
            rsGriglia2!IDAnagraficaCliente = fnNotNullN(rs!IDAnagraficaCliente)
            rsGriglia2!AnagraficaCliente = fnNotNull(rs!AnagraficaCliente) & " " & fnNotNull(rs!NomeCliente)
            rsGriglia2!IDRV_POContrattoServizi = fnNotNullN(rs!IDRV_POContrattoServizi)
            rsGriglia2!IDRV_POContrattoProdotti = fnNotNullN(rs!IDRV_POContrattoProdotti)
            rsGriglia2!IDAnagraficaTecnicoRif = fnNotNullN(rs!IDAnagraficaTecnicoRif)
            rsGriglia2!AnagraficaTecnicoRif = fnNotNull(rs!AnagraficaTecnico) & " " & fnNotNull(rs!NomeTecnico)
            rsGriglia2!IDAnagraficaTecnicoOperativo = fnNotNullN(rs!IDAnagraficaTecnicoOperativo)
            rsGriglia2!AnagraficaTecnicoOperativo = fnNotNull(rs!AnagraficaTecnicoOperativo) & " " & fnNotNull(rs!NomeTecnicoOperativo)
            rsGriglia2!DataAppuntamento = rs!DataAppuntamento
            rsGriglia2!OraAppuntamento = fnNotNull(OraAppuntamento)
            rsGriglia2!Manuale = rs!Manuale
            rsGriglia2!Elaborato = rs!Elaborata
            rsGriglia2!InterventoChiuso = fnNotNullN(rs!InterventoChiuso)
            rsGriglia2!IDRV_POCategoriaIntervento = fnNotNullN(rs!IDRV_POCategoriaIntervento)
            rsGriglia2!CategoriaIntervento = fnNotNull(rs!CategoriaIntervento)
            rsGriglia2!IDRV_POTipoAddebito = fnNotNullN(rs!IDRV_POTipoAddebito)
            rsGriglia2!TipoAddebitoIntervento = fnNotNull(rs!TipoAddebito)
            rsGriglia2!IDRV_POTipoClasseIntervento = fnNotNullN(rs!IDRV_POTipoClasseIntervento)
            rsGriglia2!ClasseIntervento = fnNotNull(rs!TipoClasseIntervento)
            rsGriglia2!IDRV_POStatoIntervento = fnNotNullN(rs!IDRV_POStatoIntervento)
            rsGriglia2!StatoIntervento = fnNotNull(rs!StatoIntervento)
            rsGriglia2!IDRV_POTipoFaseIntervento = fnNotNullN(rs!IDRV_POTipoFaseIntervento)
            rsGriglia2!TipoFaseIntervento = fnNotNull(rs!TipoFaseIntervento)
            rsGriglia2!Elimina = True
            rsGriglia2!EliminaObbligatorio = True
'            If fnNotNullN(rs!DataAppuntamento) = 0 Then
'
'            Else
'                If DateDiff("d", rs!DataAppuntamento, Date) >= 0 Then
'                    rsGriglia2!Elimina = False
'                    rsGriglia2!EliminaObbligatorio = False
'                Else
'                    rsGriglia2!Elimina = True
'                    rsGriglia2!EliminaObbligatorio = True
'                End If
'            End If
        rsGriglia2.Update
        NUMERO_INTERVENTI_DA_ELIMINARE = NUMERO_INTERVENTI_DA_ELIMINARE + 1
        
        If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
         Me.ProgressBar1.Value = Me.ProgressBar1.Max
        Else
            Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
        End If
        
        DoEvents
    rs.MoveNext
    Wend
rsGrigliaServizi.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_GET_DATI_RECORDSET_2:
    MsgBox Err.Description, vbCritical, "GET_DATI_RECORDSET_2"
End Sub

Private Sub GET_DATI_RECORDSET_3()
On Error GoTo ERR_GET_DATI_RECORDSET_3
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsServizio As ADODB.Recordset
Dim rsProdotti As ADODB.Recordset
Dim rsFestivita As ADODB.Recordset

'''''''''''''''DICHIARAZIONE DELLE VARIABILI''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim DataInizioServizio As String
Dim DataInizioPersonalizzata As String
Dim DataFinePersonalizzata As String

Dim DataFineServizio As String

Dim X_Ricorrenza As Long
Dim I As Integer
Dim NumeroIntervento As Long
Dim LINK_INTERVENTO As Long
Dim NumeroInterventoSub As Long
Dim IProdotto As Long
Dim AvviaConProdotti As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim UnitaProgresso As Double
Dim NumeroRecord As Long

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100

NumeroRecord = 100

If NumeroRecord = 0 Then Exit Sub

UnitaProgresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 4)
Me.lblInfo.Caption = "Interventi da creare..."
DoEvents

NUMERO_INTERVENTI_DA_CREARE = 0

sSQL = "SELECT * FROM RV_POFestivita"

Set rsFestivita = New ADODB.Recordset
rsFestivita.Open sSQL, Cn.InternalConnection

If ((rsGrigliaServizi.EOF) And (rsGrigliaServizi.BOF)) Then Exit Sub

rsGrigliaServizi.MoveFirst

While Not rsGrigliaServizi.EOF

    Link_Contratto_Servizio = rsGrigliaServizi!IDRV_POContrattoServizi

    'CONNESSIONE AL SERVIZIO
    sSQL = "SELECT RV_POContrattoServizi.IDRV_POContrattoServizi, RV_POContrattoServizi.IDRV_POContratto, RV_POContrattoServizi.IDRV_POStoriaContratto, RV_POContrattoServizi.IDArticolo, "
    sSQL = sSQL & "RV_POContrattoServizi.IDRV_POCriterioRicorrenza, RV_POContrattoServizi.OgniNumeroGiorni, RV_POContrattoServizi.OgniNumeroMesi, RV_POContrattoServizi.OgniNumeroSettimane, "
    sSQL = sSQL & "RV_POContrattoServizi.IDRV_POTipoDataInizioRicorrenza, RV_POContrattoServizi.GiornoInizioRicorrenza, RV_POContrattoServizi.MeseInizioRicorrenza, "
    sSQL = sSQL & "RV_POContrattoServizi.IDRV_POTipoDataFineRicorrenza, RV_POContrattoServizi.GiornoFineRicorrenza, RV_POContrattoServizi.MeseFineRicorrenza, RV_POContrattoServizi.NumeroRicorrenze, "
    sSQL = sSQL & "RV_POContrattoServizi.IDRV_POContrattoPadre , Articolo.CodiceArticolo, Articolo.Articolo, RV_POContrattoServizi.IDRV_POTipoAnnoInizioRicorrenza, RV_POContrattoServizi.IDRV_POTipoAnnoFineRicorrenza "
    sSQL = sSQL & "FROM RV_POContrattoServizi INNER JOIN "
    sSQL = sSQL & "Articolo ON RV_POContrattoServizi.IDArticolo = Articolo.IDArticolo "
    sSQL = sSQL & "WHERE RV_POContrattoServizi.IDRV_POContrattoServizi=" & Link_Contratto_Servizio
    
    Set rsServizio = New ADODB.Recordset
    
    rsServizio.Open sSQL, Cn.InternalConnection
    
    If rsServizio.EOF Then
        rsServizio.Close
        Set rsServizio = Nothing
        Exit Sub
    End If
    
    'CONNESSIONE AI PRODOTTI COLLEGATI AL SERVIZIO
    sSQL = "SELECT RV_POContrattoServiziProdotti.IDRV_POContrattoServiziProdotti, RV_POContrattoServiziProdotti.IDRV_POContrattoServizi, RV_POContrattoServiziProdotti.IDRV_POContrattoProdotti, "
    sSQL = sSQL & "RV_POContrattoServiziProdotti.Eliminato , RV_POContrattoProdotti.IDRV_POProdotto, RV_POContrattoProdotti.Quantita, RV_POProdotto.Descrizione, RV_POProdotto.Matricola, RV_POContrattoProdotti.Dismesso, "
    sSQL = sSQL & "RV_POContrattoProdotti.DataInizioPeriodo, RV_POContrattoProdotti.DataFinePeriodo, RV_POContrattoProdotti.EscludiGiorniFestivi, RV_POContrattoProdotti.EscludiSabato, RV_POContrattoProdotti.Conducente, "
    sSQL = sSQL & "RV_POContrattoProdotti.IDAnagraficaOperatore "
    sSQL = sSQL & "FROM RV_POContrattoServiziProdotti INNER JOIN "
    sSQL = sSQL & "RV_POContrattoProdotti ON RV_POContrattoServiziProdotti.IDRV_POContrattoProdotti = RV_POContrattoProdotti.IDRV_POContrattoProdotti INNER JOIN "
    sSQL = sSQL & "RV_POProdotto ON RV_POContrattoProdotti.IDRV_POProdotto = RV_POProdotto.IDRV_POProdotto "
    sSQL = sSQL & "WHERE RV_POContrattoServiziProdotti.IDRV_POContrattoServizi=" & Link_Contratto_Servizio
    sSQL = sSQL & " AND RV_POContrattoServiziProdotti.Eliminato=" & fnNormBoolean(0)
    sSQL = sSQL & " AND RV_POContrattoProdotti.Dismesso =" & fnNormBoolean(0)
    
    Set rsProdotti = New ADODB.Recordset
    rsProdotti.Open sSQL, Cn.InternalConnection
    
    If rsProdotti.EOF Then
        AvviaConProdotti = False
    Else
        AvviaConProdotti = True
    End If
    
    'CALCOLO DELLA DATA INIZIO RICORRENZA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (frmMain.cboTipoImpostazione.CurrentID <> 3) Then
        DataInizioServizio = frmMain.txtDataDecorrenza.Text
    Else
        If (fnNotNullN(rsServizio!IDRV_POTipoDataInizioRicorrenza) = 0) And (fnNotNullN(rsServizio!IDRV_POTipoDataFineRicorrenza) > 0) Then
            DataInizioServizio = fnNotNull(rsProdotti!DataFinePeriodo)
        Else
            If AvviaConProdotti = True Then
                DataInizioServizio = fnNotNull(rsProdotti!DataInizioPeriodo)
            Else
                DataInizioServizio = frmMain.txtDataDecorrenza.Text
            End If
        End If
    End If
    
    
    If (fnNotNullN(rsServizio!OgniNumeroGiorni) > 0) Then
        DataInizioServizio = DateAdd("d", -1, DataInizioServizio)
    End If
    
    Select Case fnNotNullN(rsServizio!IDRV_POTipoDataInizioRicorrenza)
        Case 1
            'DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), frmMain.txtDataDecorrenza.Text)
            DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), DataInizioServizio)
            DataInizioServizio = DateAdd("d", IIf((fnNotNullN(rsServizio!OgniNumeroGiorni) = 0), 0, fnNotNullN(rsServizio!OgniNumeroGiorni)), DataInizioServizio)
            DataInizioServizio = DateAdd("ww", IIf((fnNotNullN(rsServizio!OgniNumeroSettimane) = 0), 0, fnNotNullN(rsServizio!OgniNumeroSettimane)), DataInizioServizio)
    '        If (frmMain.cboTipoImpostazione.CurrentID = 3) Then
    '            DataInizioServizio = DateAdd("d", -1, DataInizioServizio)
    '        End If
        Case 2
            'DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), frmMain.txtDataDecorrenza.Text)
            DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), DataInizioServizio)
            DataInizioServizio = DateAdd("d", IIf((fnNotNullN(rsServizio!OgniNumeroGiorni) = 0), 0, fnNotNullN(rsServizio!OgniNumeroGiorni)), DataInizioServizio)
            DataInizioServizio = DateAdd("ww", IIf((fnNotNullN(rsServizio!OgniNumeroSettimane) = 0), 0, fnNotNullN(rsServizio!OgniNumeroSettimane)), DataInizioServizio)
        Case 3
            DataInizioPersonalizzata = GET_COSTRUZIONE_DATA_PERS(fnNotNullN(rsServizio!GiornoInizioRicorrenza), fnNotNullN(rsServizio!MeseInizioRicorrenza)) '& Year(frmMain.txtDataDecorrenza.Text)
            If (fnNotNullN(rsServizio!IDRV_POTipoAnnoInizioRicorrenza)) = 1 Then DataInizioPersonalizzata = DataInizioPersonalizzata & Year(frmMain.txtDataDecorrenza.Text)
            If (fnNotNullN(rsServizio!IDRV_POTipoAnnoInizioRicorrenza)) = 2 Then DataInizioPersonalizzata = DataInizioPersonalizzata & Year(frmMain.txtDataScadenzaPerRinnovo.Text)
            
'            If (fnNotNullN(rsServizio!GiornoInizioRicorrenza) = fnNotNullN(rsServizio!GiornoFineRicorrenza)) And (fnNotNullN(rsServizio!MeseInizioRicorrenza) = fnNotNullN(rsServizio!MeseFineRicorrenza)) Then
'                DataInizioPersonalizzata = GET_COSTRUZIONE_DATA_PERS(fnNotNullN(rsServizio!GiornoInizioRicorrenza), fnNotNullN(rsServizio!MeseInizioRicorrenza)) & Year(frmMain.txtDataFineAssistenza.Text)
'            End If
            If (IsDate(DataInizioPersonalizzata)) Then DataInizioServizio = DataInizioPersonalizzata

            'DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), DataInizioPersonalizzata)
            'DataInizioServizio = DateAdd("d", IIf((fnNotNullN(rsServizio!OgniNumeroGiorni) = 0), 0, fnNotNullN(rsServizio!OgniNumeroGiorni)), DataInizioServizio)
            'DataInizioServizio = DateAdd("ww", IIf((fnNotNullN(rsServizio!OgniNumeroSettimane) = 0), 0, fnNotNullN(rsServizio!OgniNumeroSettimane)), DataInizioServizio)
        Case Else
            'DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), frmMain.txtDataDecorrenza.Text)
            DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), DataInizioServizio)
            DataInizioServizio = DateAdd("d", IIf((fnNotNullN(rsServizio!OgniNumeroGiorni) = 0), 0, fnNotNullN(rsServizio!OgniNumeroGiorni)), DataInizioServizio)
            DataInizioServizio = DateAdd("ww", IIf((fnNotNullN(rsServizio!OgniNumeroSettimane) = 0), 0, fnNotNullN(rsServizio!OgniNumeroSettimane)), DataInizioServizio)
            If (frmMain.cboTipoImpostazione.CurrentID <> 3) Then
                DataInizioServizio = DateAdd("d", -1, DataInizioServizio)
            End If
    End Select
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'CALCOLO DELLA DATA DI FINE RICORRENZA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case fnNotNullN(rsServizio!IDRV_POTipoDataFineRicorrenza)
        Case 1
            If (frmMain.cboTipoImpostazione.CurrentID <> 3) Then
                DataFineServizio = frmMain.txtDataScadenza.Text
            Else
                If AvviaConProdotti = True Then
                    DataFineServizio = fnNotNull(rsProdotti!DataFinePeriodo)
                Else
                    DataFineServizio = frmMain.txtDataScadenza.Text
                End If
            End If
        Case 2
            If (frmMain.cboTipoImpostazione.CurrentID <> 3) Then
                DataFineServizio = frmMain.txtDataScadenzaPerRinnovo.Text
            Else
                If AvviaConProdotti = True Then
                    DataFineServizio = fnNotNull(rsProdotti!DataFinePeriodo)
                Else
                    DataFineServizio = frmMain.txtDataScadenzaPerRinnovo.Text
                End If
            End If
        Case 3
            If (frmMain.cboTipoImpostazione.CurrentID <> 3) Then
                DataFineServizio = frmMain.txtDataFineAssistenza.Text
            Else
                If AvviaConProdotti = True Then
                    DataFineServizio = fnNotNull(rsProdotti!DataFinePeriodo)
                Else
                    DataFineServizio = frmMain.txtDataFineAssistenza.Text
                End If
            End If
        Case 4
            If (frmMain.cboTipoImpostazione.CurrentID <> 3) Then
                'DataFineServizio = GET_COSTRUZIONE_DATA_PERS(fnNotNullN(rsServizio!GiornoFineRicorrenza), fnNotNullN(rsServizio!MeseFineRicorrenza)) & Year(frmMain.txtDataFineAssistenza.Text)
                DataFinePersonalizzata = GET_COSTRUZIONE_DATA_PERS(fnNotNullN(rsServizio!GiornoFineRicorrenza), fnNotNullN(rsServizio!MeseFineRicorrenza)) '& Year(frmMain.txtDataFineAssistenza.Text)
                If (fnNotNullN(rsServizio!IDRV_POTipoAnnoFineRicorrenza)) = 1 Then DataFinePersonalizzata = DataFinePersonalizzata & Year(frmMain.txtDataDecorrenza.Text)
                If (fnNotNullN(rsServizio!IDRV_POTipoAnnoFineRicorrenza)) = 2 Then DataFinePersonalizzata = DataFinePersonalizzata & Year(frmMain.txtDataScadenzaPerRinnovo.Text)
                
                If (IsDate(DataFinePersonalizzata)) Then
                    DataFineServizio = DataFinePersonalizzata
                Else
                    DataFineServizio = DataInizioServizio
                End If
                
            Else
                If AvviaConProdotti = True Then
                    DataFineServizio = GET_COSTRUZIONE_DATA_PERS(fnNotNullN(rsServizio!GiornoFineRicorrenza), fnNotNullN(rsServizio!MeseFineRicorrenza)) & Year(fnNotNull(rsProdotti!DataFinePeriodo))
                Else
                    DataFinePersonalizzata = GET_COSTRUZIONE_DATA_PERS(fnNotNullN(rsServizio!GiornoFineRicorrenza), fnNotNullN(rsServizio!MeseFineRicorrenza)) '& Year(frmMain.txtDataFineAssistenza.Text)
                    If (fnNotNullN(rsServizio!IDRV_POTipoAnnoFineRicorrenza)) = 1 Then DataFinePersonalizzata = DataFinePersonalizzata & Year(frmMain.txtDataDecorrenza.Text)
                    If (fnNotNullN(rsServizio!IDRV_POTipoAnnoFineRicorrenza)) = 2 Then DataFinePersonalizzata = DataFinePersonalizzata & Year(frmMain.txtDataScadenzaPerRinnovo.Text)
                    
                    If (IsDate(DataFinePersonalizzata)) Then
                        DataFineServizio = DataFinePersonalizzata
                    Else
                        DataFineServizio = DataInizioServizio
                    End If
                End If
            End If
        Case Else
            DataFineServizio = DataInizioServizio
    End Select
    
    
    If (fnNotNullN(rsServizio!IDRV_POTipoDataInizioRicorrenza) = 0) Then DataInizioServizio = DataFineServizio
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'CALCOLO NUMERO RICORRENZE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        X_Ricorrenza = fnNotNullN(rsServizio!NumeroRicorrenze)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    NumeroInterventoSub = 1
    
    If X_Ricorrenza > 0 Then
        For I = 1 To X_Ricorrenza
            
            If DataFineServizio <> "" Then
                If DateDiff("d", DataFineServizio, DataInizioServizio) > 0 Then
                    Exit For
                End If
            End If
            'If DateDiff("d", Date, DataInizioServizio) >= 0 Then
                'CICLO DEI PRODOTTI
                If AvviaConProdotti = True Then
                    rsProdotti.MoveFirst
                    While Not rsProdotti.EOF
                    
                        For IProdotto = 1 To fnNotNullN(rsProdotti!Quantita)
                            
                            If GET_DATA_FESTIVO(DataInizioServizio, fnNotNull(rsProdotti!EscludiGiorniFestivi), fnNotNullN(rsProdotti!EscludiSabato), rsFestivita) = False Then
                                rsGriglia3.AddNew
                                    
                                    rsGriglia3!NumeroInterventoSub = NumeroInterventoSub
                                    rsGriglia3!FaseIntervento = 1
                                    rsGriglia3!IDRV_POContratto = Link_Contratto
                                    rsGriglia3!IDRV_POContrattoPadre = frmMain.txtIDContrattoPadre.Value
                                    
                                    rsGriglia3!IDArticolo = fnNotNullN(rsServizio!IDArticolo)
                                    rsGriglia3!CodiceServizio = fnNotNull(rsServizio!CodiceArticolo)
                                    rsGriglia3!DescrizioneServizio = fnNotNull(rsServizio!Articolo)
                                    
                                    rsGriglia3!IDRV_POProdotto = fnNotNullN(rsProdotti!IDRV_POProdotto)
                                    rsGriglia3!DescrizioneProdotto = fnNotNull(rsProdotti!Descrizione)
                                    rsGriglia3!Matricola = fnNotNull(rsProdotti!Matricola)
                                    
                                    rsGriglia3!IDRV_POContrattoServizi = Link_Contratto_Servizio
                                    rsGriglia3!IDRV_POContrattoProdotti = fnNotNullN(rsProdotti!IDRV_POContrattoProdotti)
                                    'rsGriglia3!IDRV_POProdottoServizi = fnNotNullN(rsProdotti!IDRV_POContrattoServiziProdotti)
        
                                    rsGriglia3!DataAppuntamento = DataInizioServizio
                                    rsGriglia3!OraAppuntamento = "09:00"
        
        
                                   rsGriglia3!Registra = True
                            
                                rsGriglia3.Update
                                NumeroInterventoSub = NumeroInterventoSub + 1
                                NUMERO_INTERVENTI_DA_CREARE = NUMERO_INTERVENTI_DA_CREARE + 1
                            End If
                            
                            If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
                             Me.ProgressBar1.Value = Me.ProgressBar1.Max
                            Else
                                Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
                            End If
                            
                            DoEvents
    
                        Next
                    
                    rsProdotti.MoveNext
                    
                    Wend
                Else
                    If (GET_CONTROLLO_ESISTENZA_INTERVENTO(Link_Contratto_Servizio, DataInizioServizio) = False) Then
                        rsGriglia3.AddNew
                            
                            rsGriglia3!NumeroInterventoSub = NumeroInterventoSub
                            rsGriglia3!FaseIntervento = 1
                            rsGriglia3!IDRV_POContratto = Link_Contratto
                            rsGriglia3!IDRV_POContrattoPadre = frmMain.txtIDContrattoPadre.Value
                            
                            rsGriglia3!IDArticolo = fnNotNullN(rsServizio!IDArticolo)
                            rsGriglia3!CodiceServizio = fnNotNull(rsServizio!CodiceArticolo)
                            rsGriglia3!DescrizioneServizio = fnNotNull(rsServizio!Articolo)
                            
                          
                            rsGriglia3!IDRV_POContrattoServizi = Link_Contratto_Servizio
                            
                            rsGriglia3!DataAppuntamento = DataInizioServizio
                            rsGriglia3!OraAppuntamento = "09:00"
                            
                            rsGriglia3!Registra = True
            
                            NUMERO_INTERVENTI_DA_CREARE = NUMERO_INTERVENTI_DA_CREARE + 1
                        rsGriglia3.Update
                        
                        If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
                         Me.ProgressBar1.Value = Me.ProgressBar1.Max
                        Else
                            Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
                        End If
                        
                        DoEvents
                    End If
                End If
            'End If
            
            
            DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), DataInizioServizio)
            DataInizioServizio = DateAdd("d", IIf((fnNotNullN(rsServizio!OgniNumeroGiorni) = 0), 0, fnNotNullN(rsServizio!OgniNumeroGiorni)), DataInizioServizio)
            DataInizioServizio = DateAdd("ww", IIf((fnNotNullN(rsServizio!OgniNumeroSettimane) = 0), 0, fnNotNullN(rsServizio!OgniNumeroSettimane)), DataInizioServizio)
            'DataInizioServizio = DateAdd("d", -1, DataInizioServizio)
            NumeroInterventoSub = 1
        Next
    Else
        If (X_Ricorrenza = 0) And (Len(DataFineServizio) > 0) Then
            While Not DateDiff("d", DataFineServizio, DataInizioServizio) > 0
                'If DateDiff("d", Date, DataInizioServizio) >= 0 Then
                    'CICLO DEI PRODOTTI
                    If AvviaConProdotti = True Then
                        rsProdotti.MoveFirst
                        While Not rsProdotti.EOF
                            IProdotto = 1
                            For IProdotto = 1 To fnNotNullN(rsProdotti!Quantita)
                                If GET_DATA_FESTIVO(DataInizioServizio, fnNotNull(rsProdotti!EscludiGiorniFestivi), fnNotNullN(rsProdotti!EscludiSabato), rsFestivita) = False Then
    
                                    rsGriglia3.AddNew
                                        rsGriglia3!NumeroInterventoSub = NumeroInterventoSub
                                        rsGriglia3!FaseIntervento = 1
                                        rsGriglia3!IDRV_POContratto = Link_Contratto
                                        rsGriglia3!IDRV_POContrattoPadre = frmMain.txtIDContrattoPadre.Value
                                        
                                        rsGriglia3!IDArticolo = fnNotNullN(rsServizio!IDArticolo)
                                        rsGriglia3!CodiceServizio = fnNotNull(rsServizio!CodiceArticolo)
                                        rsGriglia3!DescrizioneServizio = fnNotNull(rsServizio!Articolo)
                                        
                                        rsGriglia3!IDRV_POProdotto = fnNotNullN(rsProdotti!IDRV_POProdotto)
                                        rsGriglia3!DescrizioneProdotto = fnNotNull(rsProdotti!Descrizione)
                                        rsGriglia3!Matricola = fnNotNull(rsProdotti!Matricola)
                                        
                                        
                                        rsGriglia3!IDRV_POContrattoServizi = Link_Contratto_Servizio
                                        rsGriglia3!IDRV_POContrattoProdotti = fnNotNullN(rsProdotti!IDRV_POContrattoProdotti)
                                        
                                        rsGriglia3!DataAppuntamento = DataInizioServizio
                                        rsGriglia3!OraAppuntamento = "09:00"
                                                                        
                                        rsGriglia3!Registra = True
                                
                                    rsGriglia3.Update
                                    NumeroInterventoSub = NumeroInterventoSub + 1
                                    NUMERO_INTERVENTI_DA_CREARE = NUMERO_INTERVENTI_DA_CREARE + 1
                                End If
                                If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
                                 Me.ProgressBar1.Value = Me.ProgressBar1.Max
                                Else
                                    Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
                                End If
                                
                                DoEvents
                                
                            Next
                        
                        rsProdotti.MoveNext
    
    
                        Wend
                    Else
                        If (GET_CONTROLLO_ESISTENZA_INTERVENTO(Link_Contratto_Servizio, DataInizioServizio) = False) Then
                            rsGriglia3.AddNew
                                rsGriglia3!NumeroInterventoSub = NumeroInterventoSub
                                rsGriglia3!FaseIntervento = 1
                                rsGriglia3!IDRV_POContratto = Link_Contratto
                                rsGriglia3!IDRV_POContrattoPadre = frmMain.txtIDContrattoPadre.Value
                                
                                rsGriglia3!IDArticolo = fnNotNullN(rsServizio!IDArticolo)
                                rsGriglia3!CodiceServizio = fnNotNull(rsServizio!CodiceArticolo)
                                rsGriglia3!DescrizioneServizio = fnNotNull(rsServizio!Articolo)
                                
                               
                                rsGriglia3!IDAnagraficaCliente = frmMain.CDCliente.KeyFieldID
                                rsGriglia3!AnagraficaCliente = frmMain.CDCliente.Code & " " & frmMain.CDCliente.Description
                                
                                rsGriglia3!IDRV_POContrattoServizi = Link_Contratto_Servizio
                                
                                rsGriglia3!IDAnagraficaTecnicoRif = frmMain.CDTecnico.KeyFieldID
                                rsGriglia3!AnagraficaTecnicoRif = frmMain.CDTecnico.Code & " " & frmMain.CDTecnico.Description
                                
                                rsGriglia3!IDAnagraficaTecnicoOperativo = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDAnaTecRifDaProdotto")
                                rsGriglia3!AnagraficaTecnicoOperativo = GET_ANAGRAFICA(rsGriglia3!IDAnagraficaTecnicoOperativo)
                                
                                rsGriglia3!DataAppuntamento = DataInizioServizio
                                rsGriglia3!OraAppuntamento = "09:00"
        
                                                            
                                rsGriglia3!Registra = True
                
                            NUMERO_INTERVENTI_DA_CREARE = NUMERO_INTERVENTI_DA_CREARE + 1
                            rsGriglia3.Update
        
                            If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
                             Me.ProgressBar1.Value = Me.ProgressBar1.Max
                            Else
                                Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
                            End If
                            
                            DoEvents
                        End If
                    End If
                'End If
                
                DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), DataInizioServizio)
                DataInizioServizio = DateAdd("d", IIf((fnNotNullN(rsServizio!OgniNumeroGiorni) = 0), 0, fnNotNullN(rsServizio!OgniNumeroGiorni)), DataInizioServizio)
                DataInizioServizio = DateAdd("ww", IIf((fnNotNullN(rsServizio!OgniNumeroSettimane) = 0), 0, fnNotNullN(rsServizio!OgniNumeroSettimane)), DataInizioServizio)
                'DataInizioServizio = DateAdd("d", -1, DataInizioServizio)
                NumeroInterventoSub = 1
            Wend
        End If
    End If

rsGrigliaServizi.MoveNext
Wend

rsFestivita.Close
Set rsFestivita = Nothing

rsServizio.Close
Set rsServizio = Nothing

rsProdotti.Close
Set rsProdotti = Nothing
Exit Sub

ERR_GET_DATI_RECORDSET_3:
    MsgBox Err.Description, vbCritical, "GET_DATI_RECORDSET_3"
End Sub

Public Function GET_NOMECOMPUTER() As String
Dim dwLen As Long
Dim strString As String
Const MAX_COMPUTERNAME_LENGTH As Long = 31
    
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    'Show the computer name
    GET_NOMECOMPUTER = strString
End Function

Function GET_NOMEUTENTE() As String
    Dim strString As String
    Dim lunghezzaStringa As Long
    lunghezzaStringa = 32
    strString = String(lunghezzaStringa, " ")
    GetUserName strString, lunghezzaStringa
    strString = Left(strString, lunghezzaStringa)
    GET_NOMEUTENTE = strString
    GET_NOMEUTENTE = Mid(GET_NOMEUTENTE, 1, Len(GET_NOMEUTENTE) - 1)
End Function
Private Function GET_ORARIO(StringaData As String) As String
Dim Ora As String
Dim Minuti As String
Dim Secondi As String

If Len(DatePart("h", StringaData)) = 1 Then
    Ora = "0" & DatePart("h", StringaData)
Else
    Ora = DatePart("h", StringaData)
End If
If Len(DatePart("n", StringaData)) = 1 Then
    Minuti = "0" & DatePart("n", StringaData)
Else
    Minuti = DatePart("n", StringaData)
End If
If Len(DatePart("s", StringaData)) = 1 Then
    Secondi = "0" & DatePart("s", StringaData)
Else
    Secondi = DatePart("s", StringaData)
End If

GET_ORARIO = Ora & "." & Minuti & "." & Secondi


End Function

Private Function GET_NUMERO_INTERVENTO(Anno As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(NumeroIntervento) as Numero "
sSQL = sSQL & "FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND AnnoIntervento=" & Anno

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_INTERVENTO = 1
Else
    GET_NUMERO_INTERVENTO = fnNotNullN(rs!numero) + 1
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PARAMETRO_AZIENDA_LONG(IDFiliale As Long, nomeCampo As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & nomeCampo
sSQL = sSQL & " FROM RV_POParametriAzienda "
sSQL = sSQL & " WHERE IDFiliale=" & IDFiliale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PARAMETRO_AZIENDA_LONG = 0
Else
    GET_PARAMETRO_AZIENDA_LONG = fnNotNullN(rs.adoColumns(nomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_GRIGLIA_1()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3

    With Me.Griglia1
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear

        'Set cl = .ColumnsHeader.Add("Elimina", "Elimina", dgInteger, True, 1000, dgAligncenter)
        '    cl.Editable = True
        '.ColumnsHeader.Add "Elimina", "EliminaObbligatorio", dgInteger, True, 1000, dgAligncenter
        
        .ColumnsHeader.Add "IDRV_POIntervento", "IDRV_POIntervento", dgInteger, False, 500, dgAlignleft
        
        '''''''''''''''''''''''''''''''''''''''''DATI INTERVENTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .ColumnsHeader.Add "InterventoChiuso", "Chiuso", dgBoolean, True, 1000, dgAligncenter
        .ColumnsHeader.Add "AnnoIntervento", "Anno", dgInteger, False, 1200, dgAlignRight
        .ColumnsHeader.Add "NumeroIntervento", "N° Intervento", dgInteger, True, 1200, dgAlignRight
        .ColumnsHeader.Add "NumeroInterventoSub", "N° sub", dgInteger, True, 1200, dgAlignRight
        .ColumnsHeader.Add "FaseIntervento", "Fase", dgInteger, True, 1200, dgAlignRight
        
        .ColumnsHeader.Add "IDRV_POContratto", "IDRV_POContratto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDRV_POContrattoPadre", "IDRV_POContrattoPadre", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnnoContratto", "Anno contratto", dgInteger, False, 1200, dgAlignRight
        .ColumnsHeader.Add "NumeroContratto", "Numero contratto", dgInteger, True, 1200, dgAlignRight
        
        
        .ColumnsHeader.Add "IDRV_POContrattoServizi", "IDRV_POContrattoServizi", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "CodiceServizio", "Codice articolo int.", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "DescrizioneServizio", "Articolo int.", dgchar, True, 2500, dgAlignleft

        .ColumnsHeader.Add "IDRV_POContrattoProdotti", "IDRV_POContrattoProdotti", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDRV_POProdotto", "IDRV_POProdotto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "DescrizioneProdotto", "Prodotto", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "Matricola", "Matricola", dgchar, True, 2500, dgAlignleft
        
        .ColumnsHeader.Add "DataAppuntamento", "Data appuntamento", dgDate, True, 1500, dgAligncenter
        .ColumnsHeader.Add "OraAppuntamento", "Ora appuntamento", dgchar, True, 1500, dgAlignRight
        
        .ColumnsHeader.Add "IDAnagraficaCliente", "IDAnagraficaCliente", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaCliente", "Cliente", dgchar, False, 2500, dgAlignleft
        
        .ColumnsHeader.Add "IDAnagraficaTecnicoRif", "IDAnagraficaTecnicoRiferimento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaTecnicoRif", "Rif. interno", dgchar, True, 2500, dgAlignleft
        
        .ColumnsHeader.Add "IDAnagraficaTecnicoOperativo", "IDAnagraficaTecnicoOperativo", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaTecnicoOperativo", "Tecnico operativo", dgchar, False, 2500, dgAlignleft
       
        .ColumnsHeader.Add "IDRV_POTipoFaseIntervento", "IDRV_POTipoFaseIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "TipoFaseIntervento", "Tipo fase", dgchar, True, 1500, dgAlignleft
       
        .ColumnsHeader.Add "IDRV_POCategoriaIntervento", "IDRV_POCategoriaIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "CategoriaIntervento", "Categoria Int.", dgchar, True, 1500, dgAlignleft
        
        .ColumnsHeader.Add "IDRV_POStatoIntervento", "IDRV_POStatoIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "StatoIntervento", "Stato Int.", dgchar, True, 1500, dgAlignleft
        

        
        Set .Recordset = rsGriglia1
        .Refresh
        .LoadUserSettings
    End With

    
    Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "Reperimento dati 1"


End Sub
Private Sub GET_GRIGLIA_2()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3

    With Me.Griglia2
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear

        
        Set cl = .ColumnsHeader.Add("Elimina", "Elimina", dgBoolean, True, 1000, dgAligncenter)
            cl.Editable = True
        .ColumnsHeader.Add "EliminaObbligatorio", "Obbligatorio", dgBoolean, True, 1000, dgAligncenter
        .ColumnsHeader.Add "IDRV_POIntervento", "IDRV_POIntervento", dgInteger, False, 500, dgAlignleft
        
        '''''''''''''''''''''''''''''''''''''''''DATI INTERVENTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .ColumnsHeader.Add "InterventoChiuso", "Chiuso", dgBoolean, True, 1000, dgAligncenter
        .ColumnsHeader.Add "AnnoIntervento", "Anno", dgInteger, False, 1200, dgAlignRight
        .ColumnsHeader.Add "NumeroIntervento", "N° Intervento", dgInteger, True, 1200, dgAlignRight
        .ColumnsHeader.Add "NumeroInterventoSub", "N° sub", dgInteger, True, 1200, dgAlignRight
        .ColumnsHeader.Add "FaseIntervento", "Fase", dgInteger, True, 1200, dgAlignRight
        
        .ColumnsHeader.Add "IDRV_POContratto", "IDRV_POContratto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDRV_POContrattoPadre", "IDRV_POContrattoPadre", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnnoContratto", "Anno contratto", dgInteger, False, 1200, dgAlignRight
        .ColumnsHeader.Add "NumeroContratto", "Numero contratto", dgInteger, True, 1200, dgAlignRight
        
        
        .ColumnsHeader.Add "IDRV_POContrattoServizi", "IDRV_POContrattoServizi", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "CodiceServizio", "Codice articolo int.", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "DescrizioneServizio", "Articolo int.", dgchar, True, 2500, dgAlignleft
        
        .ColumnsHeader.Add "IDRV_POContrattoProdotti", "IDRV_POContrattoProdotti", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDRV_POProdotto", "IDRV_POProdotto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "DescrizioneProdotto", "Prodotto", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "Matricola", "Matricola", dgchar, True, 2500, dgAlignleft
        
        .ColumnsHeader.Add "DataAppuntamento", "Data appuntamento", dgDate, True, 1500, dgAligncenter
        .ColumnsHeader.Add "OraAppuntamento", "Ora appuntamento", dgchar, True, 1500, dgAlignRight
        
        .ColumnsHeader.Add "IDAnagraficaCliente", "IDAnagraficaCliente", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaCliente", "Cliente", dgchar, False, 2500, dgAlignleft
        
        .ColumnsHeader.Add "IDAnagraficaTecnicoRif", "IDAnagraficaTecnicoRiferimento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaTecnicoRif", "Rif. interno", dgchar, True, 2500, dgAlignleft
        
        .ColumnsHeader.Add "IDAnagraficaTecnicoOperativo", "IDAnagraficaTecnicoOperativo", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaTecnicoOperativo", "Tecnico operativo", dgchar, False, 2500, dgAlignleft

        .ColumnsHeader.Add "IDRV_POTipoFaseIntervento", "IDRV_POTipoFaseIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "TipoFaseIntervento", "Tipo fase", dgchar, True, 1500, dgAlignleft

        .ColumnsHeader.Add "IDRV_POCategoriaIntervento", "IDRV_POCategoriaIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "CategoriaIntervento", "Categoria Int.", dgchar, True, 1500, dgAlignleft
        
        .ColumnsHeader.Add "IDRV_POStatoIntervento", "IDRV_POStatoIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "StatoIntervento", "Stato Int.", dgchar, True, 1500, dgAlignleft
        



           
        Set .Recordset = rsGriglia2
        .Refresh
        .LoadUserSettings
    End With

    
    Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "Reperimento dati 2"


End Sub
Private Sub GET_GRIGLIA_3()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3

    With Me.Griglia3
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear

        Set cl = .ColumnsHeader.Add("Registra", "Registra", dgBoolean, True, 1000, dgAligncenter)
            'cl.Editable = True
        '.ColumnsHeader.Add "IDRV_POIntervento", "IDRV_POIntervento", dgInteger, False, 500, dgAlignleft
        
        '''''''''''''''''''''''''''''''''''''''''DATI INTERVENTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '.ColumnsHeader.Add "InterventoChiuso", "Chiuso", dgBoolean, True, 1000, dgAligncenter
        '.ColumnsHeader.Add "NumeroFase", "N°", dgInteger, True, 500, dgAlignRight
        
        '.ColumnsHeader.Add "AnnoIntervento", "Anno intervento", dgInteger, False, 1200, dgAlignRight
        '.ColumnsHeader.Add "NumeroIntervento", "Numero Intervento", dgInteger, True, 1200, dgAlignRight
        .ColumnsHeader.Add "NumeroInterventoSub", "Numero sub", dgInteger, True, 1200, dgAlignRight
        '.ColumnsHeader.Add "FaseIntervento", "Fase", dgInteger, True, 1200, dgAlignRight
        
        '.ColumnsHeader.Add "IDRV_POContratto", "IDRV_POContratto", dgInteger, False, 500, dgAlignleft
        '.ColumnsHeader.Add "IDRV_POContrattoPadre", "IDRV_POContrattoPadre", dgInteger, False, 500, dgAlignleft
        '.ColumnsHeader.Add "AnnoContratto", "Anno contratto", dgInteger, False, 1200, dgAlignRight
        '.ColumnsHeader.Add "NumeroContratto", "Numero contratto", dgInteger, True, 1200, dgAlignRight
        
        
        .ColumnsHeader.Add "IDRV_POContrattoServizi", "IDRV_POContrattoServizi", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "CodiceServizio", "Codice articolo int.", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "DescrizioneServizio", "Articolo int.", dgchar, True, 2500, dgAlignleft

        .ColumnsHeader.Add "IDRV_POContrattoProdotti", "IDRV_POContrattoProdotti", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDRV_POProdotto", "IDRV_POProdotto", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "DescrizioneProdotto", "Prodotto", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "Matricola", "Matricola", dgchar, True, 2500, dgAlignleft
        
        .ColumnsHeader.Add "DataAppuntamento", "Data appuntamento", dgDate, True, 1500, dgAligncenter
        .ColumnsHeader.Add "OraAppuntamento", "Ora appuntamento", dgchar, True, 1500, dgAlignRight
        
        .ColumnsHeader.Add "IDAnagraficaCliente", "IDAnagraficaCliente", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaCliente", "Cliente", dgchar, False, 2500, dgAlignleft
        
        .ColumnsHeader.Add "IDAnagraficaTecnicoRif", "IDAnagraficaTecnicoRiferimento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaTecnicoRif", "Rif. interno", dgchar, True, 2500, dgAlignleft
        
        .ColumnsHeader.Add "IDAnagraficaTecnicoOperativo", "IDAnagraficaTecnicoOperativo", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaTecnicoOperativo", "Tecnico operativo", dgchar, False, 2500, dgAlignleft

        .ColumnsHeader.Add "IDRV_POTipoFaseIntervento", "IDRV_POTipoFaseIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "TipoFaseIntervento", "Tipo fase", dgchar, True, 1500, dgAlignleft

        .ColumnsHeader.Add "IDRV_POCategoriaIntervento", "IDRV_POCategoriaIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "CategoriaIntervento", "Categoria Int.", dgchar, True, 1500, dgAlignleft
        
        .ColumnsHeader.Add "IDRV_POStatoIntervento", "IDRV_POStatoIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "StatoIntervento", "Stato Int.", dgchar, True, 1500, dgAlignleft
        

            
        
        Set .Recordset = rsGriglia3
        .Refresh
        .LoadUserSettings
    End With

    
    Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "Reperimento dati 3"


End Sub

Private Function GET_COSTRUZIONE_DATA_PERS(Giorno As String, mese As String) As String
Dim GiornoInizio As String
Dim MeseInizio As String

If Len(Giorno) = 1 Then
    GiornoInizio = "0" & Giorno
ElseIf Len(Giorno) = 0 Then
    GiornoInizio = "01" '& Giorno
Else
    GiornoInizio = Giorno
End If
    
If Len(mese) = 1 Then
    MeseInizio = "0" & mese
ElseIf Len(mese) = 0 Then
    MeseInizio = "01" '& Giorno
Else
    MeseInizio = mese
End If

GET_COSTRUZIONE_DATA_PERS = GiornoInizio & "/" & MeseInizio & "/"
End Function
Private Function GET_ANAGRAFICA(IDAnagrafica As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica, Anagrafica, Nome "
sSQL = sSQL & "FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ANAGRAFICA = ""
Else
    GET_ANAGRAFICA = fnNotNull(rs!Anagrafica) & " " & fnNotNull(rs!Nome)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_TABELLA(Tabella As String, NomeCampoReturn As String, NomeCampoWhere As String, ValoreIDWhere As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampoReturn
sSQL = sSQL & " FROM " & Tabella
sSQL = sSQL & " WHERE " & NomeCampoWhere & "=" & ValoreIDWhere

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_TABELLA = ""
Else
    GET_DESCRIZIONE_TABELLA = fnNotNull(rs.adoColumns(NomeCampoReturn).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PARAMETRI_TEC_OPE(IDAnagrafica As Long, Nome As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & Nome
sSQL = sSQL & " FROM RV_POConfigurazioneTecnicoOpe "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PARAMETRI_TEC_OPE = 0
Else
    GET_PARAMETRI_TEC_OPE = fnNotNullN(rs.adoColumns(Nome).Value)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_LINK_STAGIONE(VarData As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POStagione FROM RV_POStagione "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND DataInizio<=" & fnNormDate(VarData)
sSQL = sSQL & " AND DataFine>=" & fnNormDate(VarData)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_STAGIONE = 0
Else
    GET_LINK_STAGIONE = fnNotNullN(rs!IDRV_POStagione)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub ELIMINA_INTERVENTI()
Dim sSQL As String
Dim UnitaProgresso As Double

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100
If NUMERO_INTERVENTI_DA_ELIMINARE = 0 Then Exit Sub

UnitaProgresso = FormatNumber((Me.ProgressBar1.Max / NUMERO_INTERVENTI_DA_ELIMINARE), 4)

Me.lblInfo.Caption = "ELIMINAZIONE IN CORSO..."
DoEvents


rsGriglia2.Filter = "Elimina=" & fnNormBoolean(True)

If ((rsGriglia2.EOF) And (rsGriglia2.BOF)) Then Exit Sub

While Not rsGriglia2.EOF
    
    sSQL = "DELETE FROM RV_POIntervento"
     sSQL = sSQL & " WHERE IDRV_POIntervento=" & fnNotNullN(rsGriglia2!IDRV_POIntervento)
    Cn.Execute sSQL

    '''ELIMINAZIONE BUONI INTERVENTO
    sSQL = "DELETE FROM RV_POInterventoRigheDett "
    sSQL = sSQL & "WHERE IDRV_POIntervento=" & fnNotNullN(rsGriglia2!IDRV_POIntervento)
    Cn.Execute sSQL
    
    '''ELIMINAZIONE BUONI INTERVENTO
    sSQL = "DELETE FROM RV_PODocumentazione "
    sSQL = sSQL & "WHERE IDRV_POIntervento=" & fnNotNullN(rsGriglia2!IDRV_POIntervento)
    Cn.Execute sSQL

    '''ELIMINAZIONE BUONI INTERVENTO
    sSQL = "DELETE FROM RV_POInterventoEmail "
    sSQL = sSQL & "WHERE IDRV_POIntervento=" & fnNotNullN(rsGriglia2!IDRV_POIntervento)
    Cn.Execute sSQL
    
    sSQL = "DELETE FROM Appuntamento "
    sSQL = sSQL & "WHERE RV_POIDIntervento=" & fnNotNullN(rsGriglia2!IDRV_POIntervento)
    Cn.Execute sSQL
    
    If (Me.ProgressBar1.Value + UnitaProgresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
    End If
    
    DoEvents
    
rsGriglia2.MoveNext
Wend


End Sub

Private Sub CREA_INTERVENTO(rsNew As ADODB.Recordset, NumeroIntervento As Long)

rsNew.AddNew

    'rsnew!IDRV_POIntervento = fnGetNewKey("RV_POIntervento", "IDRV_POIntervento")
    
    rsNew!IDRV_POContrattoServizi = rsGriglia3!IDRV_POContrattoServizi
    rsNew!IDRV_POContratto = rsGriglia3!IDRV_POContratto
    rsNew!IDRV_POContrattoPadre = rsGriglia3!IDRV_POContrattoPadre
    
    rsNew!Elaborata = 1
    rsNew!Manuale = 0
    
    rsNew!IDAnagraficaCliente = LINK_CLIENTE_LOCAL
    rsNew!IDAnagraficaFatturazione = rsNew!IDAnagraficaCliente
    
    rsNew!IDAzienda = TheApp.IDFirm
    rsNew!IDFiliale = TheApp.Branch
    rsNew!AnnoIntervento = Year(rsGriglia3!DataAppuntamento)
    rsNew!NumeroIntervento = NumeroIntervento 'GET_NUMERO_INTERVENTO(Year(Date))
    rsNew!NumeroInterventoSub = rsGriglia3!NumeroInterventoSub
    rsNew!NumeroFase = 1
    rsNew!InterventoChiuso = 0
    
    rsNew!IDAnagraficaTecnicoRif = LINK_RIFERIMENTO_INT_LOCAL
    rsNew!IDTipoAnagraficaTecnicoRif = LINK_TIPO_ANA_TEC_RIF
    
    rsNew!IDAnagraficaTecnicoOperativo = LINK_TECNICO_OPERATIVO_LOCAL
    rsNew!IDTipoAnagraficaTecnicoOpe = LINK_TIPO_ANA_TEC_OPE
    
    Me.RichTextBox1.Text = fnNotNull(rsGriglia3!DescrizioneServizio)
    
    rsNew!Richiesta = Me.RichTextBox1.TextRTF
    
    rsNew!DataAppuntamento = rsGriglia3!DataAppuntamento
    rsNew!OraAppuntamento = rsGriglia3!OraAppuntamento
    rsNew!LavoroEseguito = ""
    rsNew!Annotazioni = ""
    
    rsNew!IDRV_POStagione = GET_LINK_STAGIONE(rsGriglia3!DataAppuntamento)
    
    rsNew!IDRV_POCategoriaIntervento = LINK_CATEGORIA_LOCAL
    
    rsNew!IDRV_POTipoAddebito = LINK_TIPO_ADDEBITO_LOCAL
    
    rsNew!IDRV_POTipoClasseIntervento = LINK_CLASSE_LOCAL
    
    rsNew!IDRV_POStatoIntervento = LINK_STATO_LOCAL
    
    rsNew!IDRV_POTipoFaseIntervento = LINK_TIPO_LOCAL

    rsNew!IDRV_POProdotto = fnNotNullN(rsGriglia3!IDRV_POProdotto)
    rsNew!IDArticolo = fnNotNullN(rsGriglia3!IDArticolo)
    rsNew!IDRV_POContrattoProdotti = fnNotNullN(rsGriglia3!IDRV_POContrattoProdotti)
    
    rsNew!DataInserimento = Now
    rsNew!OraInserimento = GET_ORARIO(Now)
    rsNew!IDUtenteInserimento = TheApp.IDUser
    rsNew!NomeComputerInserimento = GET_NOMECOMPUTER
    rsNew!UtenteComputerInserimento = GET_NOMEUTENTE
    
    rsNew!DataUltimaModifica = Now
    rsNew!OraUltimaModifica = GET_ORARIO(Now)
    rsNew!IDUtenteUltimaModifica = TheApp.IDUser
    rsNew!NomeComputerModifica = GET_NOMECOMPUTER
    rsNew!UtenteComputerModifica = GET_NOMEUTENTE
    
    rsNew!Verificato = 0
    rsNew!AppuntamentoConfermato = 0
    rsNew!FeedBack = 0
    rsNew!AppuntamentoPressoCliente = 0
    rsNew!VisualizzaInPlanning = 0
    
    rsNew!DataChiamata = rsGriglia3!DataAppuntamento
    rsNew!OraChiamata = rsGriglia3!OraAppuntamento
    
    
rsNew.Update
End Sub
Private Function GET_DATA_FESTIVO(DataInizioPeriodo, EscludiFestivi As Long, EscludiSabato As Long, rs As ADODB.Recordset) As Boolean
On Error GoTo ERR_GET_CALCOLO_QUANTITA_EFFETTIVA
Dim DataElaborata As String
Dim Giorni As Long
Dim I As Long
Dim NumeroGiorni As Long
Dim Incrementa As Boolean
'Dim rs As ADODB.Recordset
Dim sSQL As String
Dim DataPasqua As String
Dim DataPasquaFinePeriodo As String
Dim DataPasquetta As String
Dim DataPasquettaFinePeriodo As String
Dim MeseDataEla As Long
Dim GiornoDataEla As Long

GET_DATA_FESTIVO = False

DataPasqua = ""
DataPasquetta = ""
DataPasquaFinePeriodo = ""
DataPasquettaFinePeriodo = ""

DataPasqua = CalcolaPasqua(Year(DataInizioPeriodo))
DataPasquetta = DateAdd("d", 1, DataPasqua)


DataElaborata = DataInizioPeriodo
Incrementa = False

MeseDataEla = Month(DataElaborata)
GiornoDataEla = Day(DataElaborata)

If EscludiFestivi = 1 Then
    If DatePart("w", DataElaborata) = 1 Then
        Incrementa = True
    End If
    
    'FESTIVITA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rs.Filter = "Mese=" & MeseDataEla
    rs.Filter = rs.Filter & " AND Giorno=" & GiornoDataEla
    rs.Filter = rs.Filter & " AND FestivitaNazionale=1"
    
    If Not rs.EOF Then
        Incrementa = True
    End If
    
    rs.Filter = vbNullString
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'PASQUA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Len(DataPasqua) > 0 Then
        If DateDiff("d", DataElaborata, DataPasqua) = 0 Then
            Incrementa = True
        End If
    End If
    
     If Len(DataPasquetta) > 0 Then
        If DateDiff("d", DataElaborata, DataPasquetta) = 0 Then
            Incrementa = True
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If

If EscludiSabato = 1 Then
    If DatePart("w", DataElaborata) = 7 Then
        Incrementa = True
    End If
End If
    
GET_DATA_FESTIVO = Incrementa

Exit Function

ERR_GET_CALCOLO_QUANTITA_EFFETTIVA:
    MsgBox Err.Description, vbCritical, "GET_DATA_FESTIVO"
End Function

Private Function CalcolaPasqua(Anno As Integer) As String
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim m As Double
Dim n As Double
Dim Giorno As Double
Dim mese As Double
   
   If (Anno <= 2099) Then
      m = 24
      n = 5
   ElseIf (Anno <= 2199) Then
      m = 24
      n = 6
   ElseIf (Anno <= 2299) Then
      m = 25
      n = 0
   ElseIf (Anno <= 2399) Then
      m = 26
      n = 1
   ElseIf (Anno <= 2499) Then
      m = 25
      n = 1
   End If
 
   a = Anno Mod 19
   b = Anno Mod 4
   c = Anno Mod 7
   d = ((19 * a) + m) Mod 30
   e = ((2 * b) + (4 * c) + (6 * d) + n) Mod 7
 
   If ((d + e) < 10) Then
      Giorno = d + e + 22
      mese = 3
   Else
      Giorno = d + e - 9
      mese = 4
   End If
 
   If (Giorno = 26 And mese = 4) Then
      Giorno = 19
      mese = 4
   End If
 
   If (Giorno = 25 And mese = 4 And d = 28 And e = 6 And a > 10) Then
      Giorno = 18
      mese = 4
   End If
 
   CalcolaPasqua = DateSerial(Anno, mese, Giorno)
End Function
Private Sub CREA_RECORDSET_SERVIZI(tutto As Long)
Dim sSQL As String
Dim I As Long
Dim rs As ADODB.Recordset

sSQL = "SELECT * FROM RV_POIEContrattoServizi "
sSQL = sSQL & "WHERE IDRV_POContrattoServizi=0"

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection

Set rsGrigliaServizi = New ADODB.Recordset
rsGrigliaServizi.CursorLocation = adUseClient


For I = 0 To rs.Fields.Count - 1
    Select Case rs.Fields(I).Type
    
        Case adChar, adVarChar, adVarWChar, adWChar, 201
            rsGrigliaServizi.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, rs.Fields(I).DefinedSize, rs.Fields(I).Attributes
            
        Case adInteger
            rsGrigliaServizi.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
            
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsGrigliaServizi.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes
            
        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
            rsGrigliaServizi.Fields.Append rs.Fields(I).Name, adBoolean, , rs.Fields(I).Attributes
            
        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
            rsGrigliaServizi.Fields.Append rs.Fields(I).Name, adDouble, , rs.Fields(I).Attributes
    End Select
Next

rs.Close
Set rs = Nothing

rsGrigliaServizi.Open , , adOpenKeyset, adLockPessimistic

sSQL = "SELECT * FROM RV_POIEContrattoServizi "
sSQL = sSQL & "WHERE IDRV_POContratto=" & Link_Contratto
If (tutto = 0) Then
    sSQL = sSQL & " AND IDRV_POContrattoServizi=" & Link_Contratto_Servizio
End If

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

While Not rs.EOF
    rsGrigliaServizi.AddNew
        For I = 0 To rs.Fields.Count - 1
            rsGrigliaServizi.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
        Next
    rsGrigliaServizi.Update
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

End Sub
Private Function GET_CONTROLLO_ESISTENZA_INTERVENTO(IDContrattoServizi As Long, DataAppuntamento As String) As Boolean
On Error GoTo ERR_GET_CONTROLLO_ESISTENZA_INTERVENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_ESISTENZA_INTERVENTO = False

sSQL = "SELECT IDRV_POIntervento FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POContrattoServizi=" & IDContrattoServizi
sSQL = sSQL & " AND Elaborata=1"
sSQL = sSQL & " AND Manuale=1"
sSQL = sSQL & " AND DataAppuntamento=" & fnNormDate(DataAppuntamento)

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_ESISTENZA_INTERVENTO = True
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_CONTROLLO_ESISTENZA_INTERVENTO:
    
End Function


