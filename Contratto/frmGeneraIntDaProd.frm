VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmGeneraIntDaProd 
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGeneraIntDaProd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9840
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUbicazioneProdotto 
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   9240
      Width           =   8295
   End
   Begin VB.TextBox txtNoteIntervento 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   7920
      Width           =   8295
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8520
      TabIndex        =   30
      Top             =   7920
      Width           =   2895
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   13361
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
   Begin VB.Frame fraIntervento 
      Caption         =   "INTERVENTO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   7335
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CheckBox chkElimina 
         Caption         =   "Elimina"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   7080
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ELIMINA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   28
         Top             =   7680
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   5880
         Width           =   855
      End
      Begin DMTDATETIMELib.dmtDate txtDataAppuntamento 
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   5880
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   503
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Forza aggiornamento"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   6720
         Width           =   3495
      End
      Begin VB.CheckBox chkRegistra 
         Caption         =   "Registra"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   6360
         Width           =   3495
      End
      Begin VB.CommandButton cmdSelServ 
         Height          =   285
         Left            =   3240
         TabIndex        =   22
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtProdotto 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtRiferimentoInt 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtStatoIntervento 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3480
         Width           =   3495
      End
      Begin VB.TextBox txtTipoIntervento 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txtClasse 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4680
         Width           =   3495
      End
      Begin VB.TextBox txtCategoria 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   5280
         Width           =   3495
      End
      Begin VB.TextBox txtTecnicOpe 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtTipoAddebito 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   4080
         Width           =   3495
      End
      Begin VB.CommandButton cmdSalva 
         Caption         =   "SALVA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   7680
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuovo 
         Caption         =   "NUOVO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   7680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Data e ora appuntamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3600
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Label Label1 
         Caption         =   "Servizio"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Prodotto"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Riferimento interno"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Tecnico operativo"
         Height          =   255
         Index           =   3
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
         TabIndex        =   16
         Top             =   3240
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo intervento"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo addebito"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   3840
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Classe"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   13
         Top             =   4440
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Categoria"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   12
         Top             =   5040
         Width           =   3495
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393217
      TextRTF         =   $"frmGeneraIntDaProd.frx":4781A
   End
   Begin VB.CheckBox chkIncludiOggi 
      Caption         =   "Includi oggi"
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   6960
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Ubicazione"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   9000
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "Annotazioni intervento"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   7680
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   5160
      TabIndex        =   33
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "frmGeneraIntDaProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia1 As ADODB.Recordset

Private LINK_CLIENTE_LOCAL As Long
Private LINK_RIFERIMENTO_INT_LOCAL As Long
Private LINK_TECNICO_OPERATIVO_LOCAL As Long
Private LINK_STATO_LOCAL As Long
Private LINK_STATO_LOCAL_CHIUSO As Long

Private LINK_TIPO_LOCAL As Long
Private LINK_TIPO_ADDEBITO_LOCAL As Long
Private LINK_CLASSE_LOCAL As Long
Private LINK_CATEGORIA_LOCAL As Long
Private LINK_TIPO_ANA_TEC_RIF As Long
Private LINK_TIPO_ANA_TEC_OPE As Long

Private Sub cmdConferma_Click()
Dim sSQL As String
Dim rs As ADODB.Recordset

CHIUDI_INTERVENTI Link_Contratto_Prodotto, Date

Me.Griglia.UpdatePosition = False
DoEvents

If Not ((rsGriglia1.EOF) And (rsGriglia1.BOF)) Then
    rsGriglia1.MoveFirst
    
    While Not rsGriglia1.EOF
                
        If ESISTENZA_INTERVENTO(Link_Contratto_Prodotto, rsGriglia1!DataAppuntamento) = False Then
            If rsGriglia1!Selezionato = 1 Then
                CREA_INTERVENTO
            End If
        End If
        
    rsGriglia1.MoveNext
    
    Wend
    
End If


Note_Prodotto_Intervento = Me.txtNoteIntervento.Text

Unload Me
End Sub

Private Sub CREA_INTERVENTO()
On Error GoTo ERR_CREA_INTERVENTO
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim IDIntervento As Long

sSQL = "SELECT * FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=0"

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic


    rsNew.AddNew
    
        'rsnew!IDRV_POIntervento = fnGetNewKey("RV_POIntervento", "IDRV_POIntervento")
        
        rsNew!IDRV_POContrattoProdotti = Link_Contratto_Prodotto
        
        rsNew!IDRV_POContrattoServizi = 0
        rsNew!IDRV_POContratto = Link_Contratto
        rsNew!IDRV_POContrattoPadre = Link_Contratto_padre
        
        rsNew!Elaborata = 1
        rsNew!Manuale = 0
        
        rsNew!IDAnagraficaCliente = frmMain.CDCliente.KeyFieldID
        rsNew!IDAnagraficaFatturazione = IDClienteFatturazione
        
        rsNew!IDAzienda = TheApp.IDFirm
        rsNew!IDFiliale = TheApp.Branch
        rsNew!AnnoIntervento = Year(Date)
        rsNew!NumeroIntervento = GET_NUMERO_INTERVENTO(Year(Date))
        rsNew!NumeroInterventoSub = 1
        rsNew!NumeroFase = 1
        rsNew!InterventoChiuso = 0
        
        rsNew!IDAnagraficaTecnicoRif = LINK_RIFERIMENTO_INT_LOCAL
        
        rsNew!IDAnagraficaTecnicoOperativo = LINK_TECNICO_OPERATIVO_LOCAL
        
        Me.RichTextBox1.Text = frmMain.txtNoteProdotto.Text
        rsNew!Richiesta = Me.RichTextBox1.TextRTF
        
        rsNew!DataAppuntamento = rsGriglia1!DataAppuntamento
        rsNew!OraAppuntamento = rsGriglia1!OraAppuntamento
        rsNew!LavoroEseguito = ""
        
        Me.RichTextBox1.Text = Me.txtNoteIntervento.Text
        rsNew!Annotazioni = Me.RichTextBox1.TextRTF
        
        rsNew!IDRV_POStagione = GET_LINK_STAGIONE(rsGriglia1!DataAppuntamento)
        
        rsNew!IDRV_POCategoriaIntervento = LINK_CATEGORIA_LOCAL
        
        rsNew!IDRV_POTipoAddebito = LINK_TIPO_ADDEBITO_LOCAL
        
        rsNew!IDRV_POTipoClasseIntervento = LINK_CLASSE_LOCAL
        
        rsNew!IDRV_POStatoIntervento = LINK_STATO_LOCAL
        
        rsNew!IDRV_POTipoFaseIntervento = LINK_TIPO_LOCAL
    
        rsNew!IDRV_POProdotto = frmMain.txtIDProdotto.Value
        
        
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
        
        rsNew!DataChiamata = rsGriglia1!DataAppuntamento
        rsNew!OraChiamata = rsGriglia1!OraAppuntamento
        rsNew!UbicazioneProdotto = Me.txtUbicazioneProdotto.Text
        
    rsNew.Update
    
    IDIntervento = fnNotNullN(rsNew!IDRV_POIntervento)
    
rsNew.Close
Set rsNew = Nothing


If (IDIntervento > 0) Then
    sSQL = "UPDATE RV_POIntervento SET "
    sSQL = sSQL & "IDRV_POInterventoPadre=" & IDIntervento
    sSQL = sSQL & " WHERE IDRV_POIntervento=" & IDIntervento
    Cn.Execute sSQL
End If

Exit Sub
ERR_CREA_INTERVENTO:
    MsgBox Err.Description, vbCritical, "CREA_INTERVENTO"
End Sub
Private Function ESISTENZA_INTERVENTO(IDContrattoProdotti As Long, DataAppuntamento As String) As Boolean
Dim sSQL As String
Dim rs As ADODB.Recordset
    
    ESISTENZA_INTERVENTO = False
    
    Set rs = New ADODB.Recordset
    
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM RV_POIntervento "
    sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDContrattoProdotti
    sSQL = sSQL & " AND DataAppuntamento=" & fnNormDate(DataAppuntamento)
    
    rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If Not rs.EOF Then
        
        rs!InterventoChiuso = 0
        rs!IDRV_POStatoIntervento = LINK_STATO_LOCAL
        rs!Annotazioni = ""
        rs!IDAnagraficaTecnicoRif = LINK_RIFERIMENTO_INT_LOCAL
        rs!IDAnagraficaTecnicoOperativo = LINK_TECNICO_OPERATIVO_LOCAL
        
        Me.RichTextBox1.Text = frmMain.txtNoteProdotto.Text
        rs!Richiesta = Me.RichTextBox1.TextRTF
        
        Me.RichTextBox1.Text = Me.txtNoteIntervento.Text
        rs!Annotazioni = Me.RichTextBox1.TextRTF
        
        rs.Update
    
        ESISTENZA_INTERVENTO = True
        
    End If
    
rs.Close
Set rs = Nothing
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyEscape) Then Unload Me
    
End Sub

Private Sub Form_Load()

    LINK_CLIENTE_LOCAL = frmMain.CDCliente.KeyFieldID
    
    LINK_RIFERIMENTO_INT_LOCAL = frmMain.CDTecnico.KeyFieldID
    
    If LINK_RIFERIMENTO_INT_LOCAL = 0 Then LINK_RIFERIMENTO_INT_LOCAL = GET_PARAMETRO_UTENTE(TheApp.IDUser, TheApp.IDFirm, "IDAnagraficaTecRif") 'PRENDERLO DALL'UTENTE
    
    LINK_TECNICO_OPERATIVO_LOCAL = 0
    
    If frmMain.chkConducente.Value = vbChecked Then
        LINK_TECNICO_OPERATIVO_LOCAL = frmMain.cboAnaOperatoreProd.CurrentID
    End If
    
    If LINK_TECNICO_OPERATIVO_LOCAL = 0 Then LINK_TECNICO_OPERATIVO_LOCAL = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDAnaTecOpeDaProdotto")
    
    'If LINK_TECNICO_OPERATIVO_LOCAL = 0 Then LINK_TECNICO_OPERATIVO_LOCAL = LINK_RIFERIMENTO_INT_LOCAL
    
    'LINK_TIPO_ANA_TEC_RIF = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDTipoAnagraficaTecnicoIntRif")
    'LINK_TIPO_ANA_TEC_OPE = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDTipoAnagraficaTecnicoFaseRif")
    
    LINK_STATO_LOCAL = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDRV_POStatoInterventoInserimento")
    LINK_STATO_LOCAL_CHIUSO = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDRV_POStatoInterventoChiuso")

    LINK_TIPO_LOCAL = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDRV_POTipoFaseInterventoEla")
    LINK_TIPO_ADDEBITO_LOCAL = GET_PARAMETRI_TEC_OPE(LINK_TECNICO_OPERATIVO_LOCAL, "IDRV_POTipoAddebito")
    LINK_CLASSE_LOCAL = GET_PARAMETRI_TEC_OPE(LINK_TECNICO_OPERATIVO_LOCAL, "IDRV_POTipoClasseIntervento")
    LINK_CATEGORIA_LOCAL = GET_PARAMETRI_TEC_OPE(LINK_TECNICO_OPERATIVO_LOCAL, "IDRV_POCategoriaFase")
    
    'Me.txtProdotto.Text = frmMain.txtProdotto.Text & " (" & frmMain.txtValIdentProd.Text & ")"
    
    Me.Caption = "GENERAZIONE INTERVENTI PER IL PRODOTTO " & frmMain.txtProdotto.Text
    If Len(frmMain.txtValIdentProd.Text) > 0 Then
        Me.Caption = Me.Caption & " (" & frmMain.txtValIdentProd.Text & ")"
    End If
    
    
    
    'Me.txtRiferimentoInt.Text = GET_ANAGRAFICA(LINK_RIFERIMENTO_INT_LOCAL)
    'Me.txtTecnicOpe.Text = GET_ANAGRAFICA(LINK_TECNICO_OPERATIVO_LOCAL)
    
    'Me.txtTipoIntervento.Text = GET_DESCRIZIONE_TABELLA("RV_POTipoFaseIntervento", "TipoFaseIntervento", "IDRV_POTipoFaseIntervento", LINK_TIPO_LOCAL)
    'Me.txtStatoIntervento.Text = GET_DESCRIZIONE_TABELLA("RV_POStatoIntervento", "StatoIntervento", "IDRV_POStatoIntervento", LINK_STATO_LOCAL)
    'Me.txtCategoria.Text = GET_DESCRIZIONE_TABELLA("RV_POCategoriaFase", "CategoriaFase", "IDRV_POCategoriaFase", LINK_CATEGORIA_LOCAL)
    'Me.txtTipoAddebito.Text = GET_DESCRIZIONE_TABELLA("RV_POTipoAddebito", "TipoAddebito", "IDRV_POTipoAddebito", LINK_TIPO_ADDEBITO_LOCAL)
    'Me.txtClasse.Text = GET_DESCRIZIONE_TABELLA("RV_POTipoClasseIntervento", "TipoClasseIntervento", "IDRV_POTipoClasseIntervento", LINK_CLASSE_LOCAL)
    
    Me.txtNoteIntervento.Text = Note_Prodotto_Intervento
    Me.txtUbicazioneProdotto.Text = frmMain.txtDescrProdotto.Text
    
    CREA_RECORSET_1
    
End Sub
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

Private Sub GET_DATI_RECORDSET_1()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim UnitaProgresso As Double
Dim NumeroRecord As Long
Dim DataInizio As String
Dim DataFine As String
Dim NumeroGiorni As Long
Dim I As Long
Dim DataInizioElaborazione As String
Dim rsFestivita As ADODB.Recordset

'''''FESTIVITA''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POFestivita"

Set rsFestivita = New ADODB.Recordset
rsFestivita.Open sSQL, Cn.InternalConnection
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

DataInizio = frmMain.txtDataInizioProd.Text
DataFine = frmMain.txtDataFineProd.Text

NumeroGiorni = DateDiff("d", DataInizio, DataFine)

DataInizioElaborazione = DataInizio

If (frmMain.chkConducente.Value = vbChecked) Then
    For I = 0 To NumeroGiorni
        If DateDiff("d", Date, DataInizioElaborazione) >= 0 Then
           If GET_DATA_FESTIVO(DataInizioElaborazione, Abs(frmMain.chkEscludiGiorniFestivi.Value), Abs(frmMain.chkEscludiSabato.Value), rsFestivita) = False Then
                rsGriglia1.AddNew
                    rsGriglia1!DataAppuntamento = DataInizioElaborazione
                    rsGriglia1!OraAppuntamento = "08:00"
                    rsGriglia1!IDAnagraficaTecnicoOperativo = LINK_TECNICO_OPERATIVO_LOCAL
                    rsGriglia1!AnagraficaTecnicoOperativo = GET_ANAGRAFICA(LINK_TECNICO_OPERATIVO_LOCAL)
                    rsGriglia1!IDAnagraficaTecnicoRif = LINK_RIFERIMENTO_INT_LOCAL
                    rsGriglia1!AnagraficaTecnicoRif = GET_ANAGRAFICA(LINK_RIFERIMENTO_INT_LOCAL)
                    rsGriglia1!Selezionato = 1
                rsGriglia1.Update
            End If
        End If
           
        DataInizioElaborazione = DateAdd("d", 1, DataInizioElaborazione)
    Next
Else
    rsGriglia1.AddNew
        rsGriglia1!DataAppuntamento = DataInizioElaborazione
        rsGriglia1!OraAppuntamento = "08:00"
        rsGriglia1!IDAnagraficaTecnicoOperativo = LINK_TECNICO_OPERATIVO_LOCAL
        rsGriglia1!AnagraficaTecnicoOperativo = GET_ANAGRAFICA(LINK_TECNICO_OPERATIVO_LOCAL)
        rsGriglia1!IDAnagraficaTecnicoRif = LINK_RIFERIMENTO_INT_LOCAL
        rsGriglia1!AnagraficaTecnicoRif = GET_ANAGRAFICA(LINK_RIFERIMENTO_INT_LOCAL)
        rsGriglia1!Selezionato = 1
    rsGriglia1.Update
End If

End Sub

Private Sub CREA_RECORSET_1()
Dim sSQL As String
Dim rs As ADODB.Recordset

If Not (rsGriglia1 Is Nothing) Then
    
    Set rsGriglia1 = Nothing
End If

Set rsGriglia1 = New ADODB.Recordset

rsGriglia1.CursorLocation = adUseClient

rsGriglia1.Fields.Append "Selezionato", adSmallInt, , adFldIsNullable 'Interventi da creare

'TECNICO DI RIFERIMENTO
rsGriglia1.Fields.Append "IDAnagraficaTecnicoRif", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia1.Fields.Append "AnagraficaTecnicoRif", adVarChar, 250, adFldIsNullable

'TECNICO OPERATIVO
rsGriglia1.Fields.Append "IDAnagraficaTecnicoOperativo", adInteger, , adFldIsNullable 'Tecnico di riferimento
rsGriglia1.Fields.Append "AnagraficaTecnicoOperativo", adVarChar, 250, adFldIsNullable

'DATA APPUNTAMENTO
rsGriglia1.Fields.Append "DataAppuntamento", adDBDate, , adFldIsNullable 'Tecnico di riferimento
rsGriglia1.Fields.Append "OraAppuntamento", adVarChar, 250, adFldIsNullable

rsGriglia1.Open , , adOpenKeyset, adLockBatchOptimistic

GET_DATI_RECORDSET_1

GET_GRIGLIA_1

End Sub

Private Sub GET_GRIGLIA_1()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3

    With Me.Griglia
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
        
        
        Set cl = .ColumnsHeader.Add("Selezionato", "Sel.", dbBoolean, True, 1500, dgAligncenter)
            cl.Editable = True
        
        .ColumnsHeader.Add "DataAppuntamento", "Data appuntamento", dgDate, True, 2500, dgAligncenter
        .ColumnsHeader.Add "OraAppuntamento", "Ora appuntamento", dgchar, True, 1500, dgAlignRight

        .ColumnsHeader.Add "IDAnagraficaTecnicoRif", "IDAnagraficaTecnicoRiferimento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaTecnicoRif", "Rif. interno", dgchar, True, 3000, dgAlignleft
        
        .ColumnsHeader.Add "IDAnagraficaTecnicoOperativo", "IDAnagraficaTecnicoOperativo", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaTecnicoOperativo", "Tecnico operativo", dgchar, True, 3000, dgAlignleft
       
        Set .Recordset = rsGriglia1
        .Refresh
        .LoadUserSettings
    End With

    
    Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "Reperimento dati 1"


End Sub

Private Function GET_PARAMETRO_UTENTE(IDUtente As Long, IDAzienda As Long, nomeCampo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & nomeCampo & " FROM RV_POParametriUtente "
sSQL = sSQL & "WHERE IDUtente=" & IDUtente
sSQL = sSQL & " AND IDAzienda=" & IDAzienda

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PARAMETRO_UTENTE = 0
Else
    GET_PARAMETRO_UTENTE = fnNotNullN(rs.adoColumns(nomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub CHIUDI_INTERVENTI(IDContrattoProdotto As Long, DataInizio As String)
Dim sSQL As String

Me.RichTextBox1.Text = "INTERVENTO CHIUSO DA ELABORAZIONE DA CONTRATTO (GENERAZIONE INTERVENTI DA PRODOTTO)"

sSQL = "UPDATE RV_POIntervento SET "
sSQL = sSQL & "InterventoChiuso=1, "
sSQL = sSQL & "IDRV_POStatoIntervento=" & LINK_STATO_LOCAL_CHIUSO & ", "
sSQL = sSQL & "Annotazioni=" & fnNormString(Me.RichTextBox1.TextRTF)
sSQL = sSQL & " WHERE IDRV_POContrattoProdotti=" & IDContrattoProdotto
sSQL = sSQL & " AND DataAppuntamento>=" & fnNormDate(DataInizio)

Cn.Execute sSQL

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
    GET_NUMERO_INTERVENTO = fnNotNullN(rs!Numero) + 1
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

Private Sub Griglia_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_GrigliaCorpo_KeyPress
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.Griglia.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsGriglia1.Fields("Selezionato").Value), 2
        End If
    End If

Exit Sub
ERR_GrigliaCorpo_KeyPress:
    MsgBox Err.Description, vbCritical, "Griglia_KeyPress"
End Sub

Private Sub Griglia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_GrigliaCorpo_MouseUp
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If Griglia.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < Griglia.ColumnsHeader("Selezionato").Width Then
            'Se non siamo in modalità filtri
            If Griglia.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsGriglia1.Fields("Selezionato").Value), 2
            End If
        End If
    End If
Exit Sub
ERR_GrigliaCorpo_MouseUp:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_MouseUp"
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
On Error GoTo ERR_sbSelectSelectedRow

If Not rsGriglia1.EOF And Not rsGriglia1.BOF Then

    rsGriglia1.Fields("Selezionato").Value = Abs(CLng(Selected))
    
    Me.Griglia.Refresh
    
End If

Exit Sub
ERR_sbSelectSelectedRow:
    MsgBox Err.Description, vbCritical, "sbSelectSelectedRow"
End Sub

