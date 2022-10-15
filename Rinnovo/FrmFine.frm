VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmFine 
   Caption         =   "Rinnovo Contratti (Passo 3 di 3)"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   Icon            =   "FrmFine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   12180
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtRichiesta 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"FrmFine.frx":4781A
   End
   Begin VB.ListBox lstRinnovi 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4320
      Left            =   2880
      TabIndex        =   8
      Top             =   600
      Width           =   9255
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Avanti"
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
      Height          =   375
      Left            =   11040
      TabIndex        =   6
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdIndietro 
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
      Left            =   7440
      TabIndex        =   5
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnulla 
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
      Left            =   8640
      TabIndex        =   4
      Top             =   6120
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   0
      Picture         =   "FrmFine.frx":4789C
      ScaleHeight     =   4815
      ScaleWidth      =   2775
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.CommandButton Fine 
      Caption         =   "Fine"
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
      Left            =   9840
      TabIndex        =   3
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RINNOVO CONTRATTI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   12015
   End
End
Attribute VB_Name = "FrmFine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsEla As DmtOleDbLib.adoResultset
Private rsNew As ADODB.Recordset

Private ImportoContrattoAttuale As Double

''''''VARIABILI PARAMETRI AZIENDA'''''''''''''''''''''''
Public LINK_TIPO_ANA_TEC_INT As Long
Public LINK_STATO_INT_NUOVO As Long
Public LINK_STATO_INT_CHIUSO As Long
Public LINK_TIPO_ANA_TEC_FASE As Long
Public LINK_TIPO_FASE_DA_PRODOTTO As Long
Public LINK_STATO_FASE_NUOVA As Long
Public LINK_STATO_FASE_CHIUSA As Long
Public LINK_TIPO_FASE_ELA As Long
Public LINK_TIPO_FASE_MANUALE As Long
Public LINK_TIPO_TASCA As Long
Public LINK_TIPO_ANA_TEC_CONTRATTO As Long
Public FLAG_RIPORTA_TEC_CONTRATTO As Long
Public LINK_TIPO_GESTIONE_BUONI As Long
Public LINK_TIPO_NUMERAZIONE_BUONO As Long
Public VOISPEED As Long

Private rsGriglia3 As ADODB.Recordset
Private NUMERO_INTERVENTI_DA_CREARE As Long
Private rsGrigliaErrori As ADODB.Recordset

Private rsInterventoUpdate As ADODB.Recordset

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


Private Function GetNumeroRecord() As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT COUNT(IDRV_POContratto) AS NumeroRecord "
    sSQL = sSQL & "FROM RV_POTMPRinnovoContratto "
    sSQL = sSQL & "WHERE (DaRinnovare = 1) "
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        If IsNull(rs!NumeroRecord) Then
            GetNumeroRecord = 0
        Else
            GetNumeroRecord = rs!NumeroRecord
        End If
    Else
        GetNumeroRecord = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function


Private Sub cmdIndietro_Click()
    Unload Me
End Sub

Private Sub Fine_Click()
On Error GoTo ERR_Fine_Click
Dim sSQL As String
Dim Unita_Progresso As Double
Dim NumeroRecord As Long
Dim NumeroElaborazioni As Long

Me.Fine.Enabled = False


Me.lblInfo.Caption = ""
Me.ProgressBar1.Value = 0

NumeroRecord = GetNumeroRecord
Me.ProgressBar1.Max = 100


If NumeroRecord = 0 Then Exit Sub

lblInfo.Caption = "START IN CORSO........"
DoEvents

Unita_Progresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 4)
NumeroElaborazioni = 1

GET_PARAMETRI_AZIENDA TheApp.IDFirm
    
CREA_RECORSET_3
CREA_RECORDSET_ERRORI

NUMERO_INTERVENTI_DA_CREARE = 0
''''APERTURA RECORDSET CONTRATTI''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & 0

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Me.lstRinnovi.Clear
    
    sSQL = "SELECT * FROM RV_POTMPRinnovoContratto "
    sSQL = sSQL & "WHERE DaRinnovare=" & fnNormBoolean(1)
    Set rsEla = CnDMT.OpenResultset(sSQL)
    
    While Not rsEla.EOF
        
        Me.lstRinnovi.AddItem "Rinnovo del contratto " & GET_ANAGRAFICA_CONTRATTO(fnNotNullN(rsEla!IDRV_POContratto))
        lstRinnovi.ListIndex = lstRinnovi.ListCount - 1
        
        DoEvents
        
        TipoRinnovo fnNotNullN(rsEla!IDTipoRinnovo)
        TipoRateizzazione fnNotNullN(rsEla!IDRateizzazione)
        DurataContratto fnNotNullN(rsEla!IDDurataContratto)
        DescrizioneRateizzazione fnNotNullN(rsEla!IDRateizzazione)
        DescrizioneTipoContratto fnNotNullN(rsEla!IDRV_POContratto)
        DescrizioneTipoDurata fnNotNullN(rsEla!IDDurataContratto)
        DescrizioneTipoRinnovo fnNotNullN(rsEla!IDTipoRinnovo)
                
        LINK_TIPO_ANA_TEC_RIF = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDTipoAnagraficaTecnicoIntRif")
        LINK_TIPO_ANA_TEC_OPE = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDTipoAnagraficaTecnicoFaseRif")
        
        GET_LINK_NUOVO_CONTRATTO rsEla!IDRV_POContratto, rsEla, rsNew
        
        If (Me.ProgressBar1.Value + Unita_Progresso) >= Me.ProgressBar1.Max Then
            Me.ProgressBar1.Value = Me.ProgressBar1.Max
        Else
            Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_Progresso
        End If
        
        ''''''Aggiornamento del contratto rinnovato'''''''''''''''''''''''''''''''''''''''
        sSQL = "UPDATE RV_POContratto SET "
        sSQL = sSQL & "ContrattoAttuale=0, "
        sSQL = sSQL & "Chiuso=1 "
        sSQL = sSQL & "WHERE IDRV_POContratto=" & fnNotNullN(rsEla!IDRV_POContratto)
        CnDMT.Execute sSQL
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Me.lblInfo.Caption = NumeroElaborazioni & " di " & NumeroRecord
        
        NumeroElaborazioni = NumeroElaborazioni + 1
        
        DoEvents

    rsEla.MoveNext
    Wend
    
    
    rsEla.CloseResultset
    Set rsEla = Nothing
    
    
    rsNew.Close
    Set rsNew = Nothing

    
    CREA_INTERVENTI
    
    Me.lstRinnovi.AddItem "OPERAZIONE CONCLUSA"
    DoEvents
    
    lblInfo.Caption = "OPERAZIONE CONCLUSA"
    DoEvents
    
    
    
Exit Sub
ERR_Fine_Click:
    MsgBox Err.Description, vbCritical, "Fine_Click"
    Me.Fine.Enabled = True
End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

If cmdAnnulla.Value = True Then
    ChiusuraConnessione
    Exit Sub
    
End If
If Me.cmdIndietro.Value = True Then
    FrmContrattiPerRinnovo.Show
    Exit Sub
End If

ChiusuraConnessione
End Sub

Private Sub ElaborazioneRinnovoContratto()
Dim sSQL As String
        
        
    sSQL = "UPDATE RV_POContratto SET "
    sSQL = sSQL & "ContrattoAttuale=" & fnNormBoolean(0) & ", "
    sSQL = sSQL & "IDUtentePerModifica=" & TheApp.IDUser & ", "
    sSQL = sSQL & "DataModifica=" & fnNormDate(Date) & " "
    sSQL = sSQL & "WHERE IDRV_POContratto=" & fnNormNumber(rsEla!IDRV_POContratto)
    CnDMT.Execute sSQL
     
    ImportoContrattoAttuale = FormatNumber(rsEla!ImportoContratto, 2)
    
End Sub
    
Private Sub RINNOVO_SERVIZI(IDContratto As Long, IDContrattoOLD As Long, IDAzienda As Long, IDStoriaContrattoOLD As Long)
On Error GoTo ERR_RINNOVO_SERVIZI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsContratto As DmtOleDbLib.adoResultset
Dim rsServizi As ADODB.Recordset

'''''''''''''''''''''PRELEVO I DATI DAL CONTRATTO''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rsContratto = CnDMT.OpenResultset(sSQL)

If rsContratto.EOF Then
    rsContratto.CloseResultset
    Set rsContratto = Nothing
    Exit Sub
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''PROCEDURA DI INSERIMENTO AUTOMATICO DEI SERVIZI E LORO CALCOLO''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_POContrattoServizi "
sSQL = sSQL & "WHERE IDRV_POContratto=0"
'sSQL = sSQL & " AND IDRV_POStoriaContratto=0"

Set rsServizi = New ADODB.Recordset
rsServizi.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic


sSQL = "SELECT * FROM RV_POContrattoServizi "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContrattoOLD

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    
    rsServizi.AddNew
        rsServizi("IDRV_POContrattoServizi").Value = fnGetNewKey("RV_POContrattoServizi", "IDRV_POContrattoServizi")
        rsServizi("IDRV_POContratto").Value = IDContratto
        rsServizi("IDRV_POContrattoPadre").Value = IDStoriaContrattoOLD
        rsServizi("IDArticolo").Value = rs("IDArticolo").Value
        rsServizi("IDRV_POCriterioRicorrenza").Value = rs("IDRV_POCriterioRicorrenza").Value
        rsServizi("OgniNumeroGiorni").Value = rs("OgniNumeroGiorni").Value
        rsServizi("OgniNumeroMesi").Value = rs("OgniNumeroMesi").Value
        rsServizi("OgniNumeroSettimane").Value = rs("OgniNumeroSettimane").Value
        rsServizi("IDRV_POTipoDataInizioRicorrenza").Value = rs("IDRV_POTipoDataInizioRicorrenza").Value
        rsServizi("GiornoInizioRicorrenza").Value = rs("GiornoInizioRicorrenza").Value
        rsServizi("MeseInizioRicorrenza").Value = rs("MeseInizioRicorrenza").Value
        rsServizi("IDRV_POTipoDataFineRicorrenza").Value = rs("IDRV_POTipoDataFineRicorrenza").Value
        rsServizi("GiornoFineRicorrenza").Value = rs("GiornoFineRicorrenza").Value
        rsServizi("MeseFineRicorrenza").Value = rs("MeseFineRicorrenza").Value
        rsServizi("NumeroRicorrenze").Value = rs("NumeroRicorrenze").Value
        rsServizi("IDRV_POContrattoServiziRinnovo").Value = rs("IDRV_POContrattoServizi").Value
        rsServizi("IDRV_POTipoAnnoInizioRicorrenza").Value = rs("IDRV_POTipoAnnoInizioRicorrenza").Value
        rsServizi("IDRV_POTipoAnnoFineRicorrenza").Value = rs("IDRV_POTipoAnnoFineRicorrenza").Value
        If fnNotNullN(rsServizi("IDRV_POTipoAnnoInizioRicorrenza").Value) = 0 Then
            rsServizi("IDRV_POTipoAnnoInizioRicorrenza").Value = GET_TIPO_ANNO_INIZIO_SERVIZIO(fnNotNullN(rsServizi("IDArticolo").Value))
        End If
        If fnNotNullN(rsServizi("IDRV_POTipoAnnoFineRicorrenza").Value) = 0 Then
            rsServizi("IDRV_POTipoAnnoFineRicorrenza").Value = GET_TIPO_ANNO_FINE_SERVIZIO(fnNotNullN(rsServizi("IDArticolo").Value))
        End If
        
    rsServizi.Update
rs.MoveNext
Wend

rsServizi.Close
Set rsServizi = Nothing

rsContratto.CloseResultset
Set rsContratto = Nothing

rs.CloseResultset
Set rs = Nothing


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Sub
ERR_RINNOVO_SERVIZI:
    MsgBox Err.Description, vbCritical, "RINNOVO_SERVIZI"
End Sub
Private Sub ELABORAZIONE_INTERVENTI_PER_SERVIZIO(IDAnagraficaCliente As Long, IDAnagraficaTecndmtdmticoRif As Long, IDContratto As Long, IDStoriaContratto As Long, IDArticolo As Long, _
IDTipoRicorrenza As Integer, GiornoRic As String, MeseRic As String, SettimanaRic As String, _
IDTipoDataInizioRic As Integer, GiorniInizioRic As String, MeseInizioRic As String, _
IDTipoDataFineRic As String, GiornoFineRic As String, MeseFineRic As String, _
NumeroRicorrenze As String, IDTipoAnaTecRif As Long, IDTipoStatoInt As Long, IDTipoStatoFase As Long, _
IDTipoFase As Long, DescrizioneArticolo As String, DataDecorrenzaContratto As String, DataScadenzaContratto As String, _
DataRinnovoContratto As String, DataFineAssistenza As String, IDTipoAnaTecOpe As Long, IDCategoriaIntervento As Long)

'''''''''''''''DICHIARAZIONE DELLE VARIABILI'''''''''''''''''''''''''''''''''
Dim DataInizioServizio As String
Dim DataInizioPersonalizzata As String
Dim DataFineServizio As String
Dim DataFinePersonalizzata As String
Dim X_Ricorrenza As Long
Dim I As Integer
Dim oItem As MSComctlLib.ListItem
Dim NumeroIntervento As Long

Dim sSQL As String
Dim rsIntervento As ADODB.Recordset
Dim rsFase As ADODB.Recordset
Dim LINK_INTERVENTO As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'If IDContratto = 13 Then
'    MsgBox "STOP"
'End If
'CALCOLO DELLA DATA INIZIO RICORRENZA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Select Case IDTipoDataInizioRic
    Case 1
        DataInizioServizio = DateAdd("m", IIf((MeseRic = ""), 0, MeseRic), DataDecorrenzaContratto)
        DataInizioServizio = DateAdd("d", IIf((GiornoRic = ""), 0, GiornoRic), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((SettimanaRic = ""), 0, SettimanaRic), DataInizioServizio)
    Case 2
        DataInizioServizio = DateAdd("m", IIf((MeseRic = ""), 0, MeseRic), DataDecorrenzaContratto)
        DataInizioServizio = DateAdd("d", IIf((GiornoRic = ""), 0, GiornoRic), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((SettimanaRic = ""), 0, SettimanaRic), DataInizioServizio)
    Case 3
        
        DataInizioPersonalizzata = GET_COSTRUZIONE_DATA_PERS(GiorniInizioRic, MeseInizioRic) & Year(DataDecorrenzaContratto)
        If (GiorniInizioRic = GiornoFineRic) And (MeseInizioRic = MeseFineRic) Then
            DataInizioPersonalizzata = GET_COSTRUZIONE_DATA_PERS(GiorniInizioRic, MeseInizioRic) & Year(DataFineAssistenza)
        End If
        DataInizioServizio = DateAdd("m", IIf((MeseRic = ""), 0, MeseRic), DataInizioPersonalizzata)
        DataInizioServizio = DateAdd("d", IIf((GiornoRic = ""), 0, GiornoRic), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((SettimanaRic = ""), 0, SettimanaRic), DataInizioServizio)
    Case Else
        DataInizioServizio = DateAdd("m", IIf((MeseRic = ""), 0, MeseRic), DataDecorrenzaContratto)
        DataInizioServizio = DateAdd("d", IIf((GiornoRic = ""), 0, GiornoRic), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((SettimanaRic = ""), 0, SettimanaRic), DataInizioServizio)
End Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'CALCOLO DELLA DATA DI FINE RICORRENZA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Select Case IDTipoDataFineRic
    Case 1
        DataFineServizio = DataScadenzaContratto
    Case 2
        DataFineServizio = DataRinnovoContratto
        
    Case 3
        DataFineServizio = DataFineAssistenza
    Case 4
        DataFineServizio = GET_COSTRUZIONE_DATA_PERS(GiornoFineRic, MeseFineRic) & Year(DataFineAssistenza)
        
    Case Else
        DataFineServizio = ""
End Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CALCOLO NUMERO RICORRENZE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If (Len(NumeroRicorrenze) = 0) Or (NumeroRicorrenze = "0") Then
    X_Ricorrenza = 0
Else
    X_Ricorrenza = NumeroRicorrenze
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set rsIntervento = New ADODB.Recordset

rsIntervento.Open "SELECT * FROM RV_POIntervento WHERE IDRV_POIntervento=0", CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If X_Ricorrenza > 0 Then
    For I = 1 To X_Ricorrenza
        
        If DataFineServizio <> "" Then
            If DateDiff("d", DataFineServizio, DataInizioServizio) > 0 Then
                Exit For
            End If
        End If
            
        rsIntervento.AddNew
            LINK_INTERVENTO = fnGetNewKey("RV_POIntervento", "IDRV_POIntervento")
            rsIntervento!IDRV_POIntervento = LINK_INTERVENTO
            rsIntervento!IDRV_POInterventoPadre = LINK_INTERVENTO
            rsIntervento!IDRV_POContratto = IDContratto
            rsIntervento!IDRV_POStoriaContratto = IDStoriaContratto
            rsIntervento!IDAnagraficaCliente = IDAnagraficaCliente
            rsIntervento!IDAnagraficaFatturazione = IDAnagraficaCliente
            rsIntervento!IDAzienda = TheApp.IDFirm
            rsIntervento!IDFiliale = TheApp.Branch
            rsIntervento!NumeroIntervento = GET_NUMERO_INTERVENTO(Year(Date))
            rsIntervento!AnnoIntervento = Year(Date)
            rsIntervento!IDArticolo = IDArticolo
            rsIntervento!IDAnagraficaTecnicoRif = IDAnagraficaTecndmtdmticoRif
            rsIntervento!IDTipoAnagraficaTecnicoRif = IDTipoAnaTecRif
            rsIntervento!IDRV_POStatoIntervento = IDTipoStatoInt
            rsIntervento!InterventoChiuso = 0
            rsIntervento!DataInserimento = Date
            rsIntervento!OraInserimento = GET_ORARIO(Now)
            rsIntervento!IDUtenteInserimento = TheApp.IDUser
            rsIntervento!DataUltimaModifica = Date
            rsIntervento!OraUltimaModifica = GET_ORARIO(Now)
            rsIntervento!IDUtenteUltimaModifica = TheApp.IDUser
            rsIntervento!Elaborata = 1
            rsIntervento!Manuale = 0
            rsIntervento!Richiesta = DescrizioneArticolo
            rsIntervento!NomeComputerInserimento = GET_NOMECOMPUTER
            rsIntervento!UtenteComputerInserimento = GET_NOMEUTENTE
            rsIntervento!NomeComputerModifica = GET_NOMECOMPUTER
            rsIntervento!UtenteComputerModifica = GET_NOMEUTENTE
            rsIntervento!NumeroFase = 1
            rsIntervento!IDRV_POTipoFaseIntervento = IDTipoFase
            rsIntervento!IDAnagraficaTecnicoOperativo = IDAnagraficaTecndmtdmticoRif
            rsIntervento!IDTipoAnagraficaTecnicoOpe = IDTipoAnaTecOpe
            rsIntervento!DataAppuntamento = DataInizioServizio
            rsIntervento!OraAppuntamento = "09.00"
            rsIntervento!LavoroEseguito = DescrizioneArticolo
            rsIntervento!Annotazioni = ""
            rsIntervento!IDRV_POStagione = GET_LINK_STAGIONE(DataInizioServizio)
            rsIntervento!IDRV_POCategoriaIntervento = GET_PARAMETRI_TEC_OPE(fnNotNullN(rsIntervento!IDAnagraficaTecnicoOperativo), "IDRV_POCategoriaFase")
            rsIntervento!IDRV_POTipoAddebito = GET_PARAMETRI_TEC_OPE(fnNotNullN(rsIntervento!IDAnagraficaTecnicoOperativo), "IDRV_POTipoAddebito")
            rsIntervento!IDRV_POTipoClasseIntervento = GET_PARAMETRI_TEC_OPE(fnNotNullN(rsIntervento!IDAnagraficaTecnicoOperativo), "IDRV_POTipoClasseIntervento")
            If fnNotNullN(rsIntervento!IDRV_POCategoriaIntervento) = 0 Then
                rsIntervento!IDRV_POCategoriaIntervento = IDCategoriaIntervento
            End If
        rsIntervento.Update
        
        DataInizioServizio = DateAdd("m", IIf((MeseRic = ""), 0, MeseRic), DataInizioServizio)
        DataInizioServizio = DateAdd("d", IIf((GiornoRic = ""), 0, GiornoRic), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((SettimanaRic = ""), 0, SettimanaRic), DataInizioServizio)
        
    Next
End If

If (X_Ricorrenza = 0) And (Len(DataFineServizio) > 0) Then
    While Not DateDiff("d", DataFineServizio, DataInizioServizio) > 0

        
        'INSERIMENTO INTERVENTO TESTA
        rsIntervento.AddNew
            LINK_INTERVENTO = fnGetNewKey("RV_POIntervento", "IDRV_POIntervento")
            rsIntervento!IDRV_POIntervento = LINK_INTERVENTO
            rsIntervento!IDRV_POInterventoPadre = LINK_INTERVENTO
            rsIntervento!IDRV_POContratto = IDContratto
            rsIntervento!IDRV_POStoriaContratto = IDStoriaContratto
            rsIntervento!IDAnagraficaCliente = IDAnagraficaCliente
            rsIntervento!IDAnagraficaFatturazione = IDAnagraficaCliente
            rsIntervento!IDAzienda = TheApp.IDFirm
            rsIntervento!IDFiliale = TheApp.Branch
            rsIntervento!NumeroIntervento = GET_NUMERO_INTERVENTO(Year(Date))
            rsIntervento!AnnoIntervento = Year(Date)
            rsIntervento!IDArticolo = IDArticolo
            rsIntervento!IDAnagraficaTecnicoRif = IDAnagraficaTecndmtdmticoRif
            rsIntervento!IDTipoAnagraficaTecnicoRif = IDTipoAnaTecRif
            rsIntervento!IDRV_POStatoIntervento = IDTipoStatoInt
            rsIntervento!InterventoChiuso = 0
            rsIntervento!DataInserimento = Date
            rsIntervento!OraInserimento = GET_ORARIO(Now)
            rsIntervento!IDUtenteInserimento = TheApp.IDUser
            rsIntervento!DataUltimaModifica = Date
            rsIntervento!OraUltimaModifica = GET_ORARIO(Now)
            rsIntervento!IDUtenteUltimaModifica = TheApp.IDUser
            rsIntervento!Elaborata = 1
            rsIntervento!Manuale = 0
            rsIntervento!Richiesta = DescrizioneArticolo
            rsIntervento!NomeComputerInserimento = GET_NOMECOMPUTER
            rsIntervento!UtenteComputerInserimento = GET_NOMEUTENTE
            rsIntervento!NomeComputerModifica = GET_NOMECOMPUTER
            rsIntervento!UtenteComputerModifica = GET_NOMEUTENTE
            rsIntervento!NumeroFase = 1
            rsIntervento!IDRV_POTipoFaseIntervento = IDTipoFase
            rsIntervento!IDAnagraficaTecnicoOperativo = IDAnagraficaTecndmtdmticoRif
            rsIntervento!IDTipoAnagraficaTecnicoOpe = IDTipoAnaTecOpe
            rsIntervento!DataAppuntamento = DataInizioServizio
            rsIntervento!OraAppuntamento = "09.00"
            rsIntervento!LavoroEseguito = DescrizioneArticolo
            rsIntervento!Annotazioni = ""
            rsIntervento!IDRV_POStagione = GET_LINK_STAGIONE(DataInizioServizio)
            rsIntervento!IDRV_POCategoriaIntervento = GET_PARAMETRI_TEC_OPE(fnNotNullN(rsIntervento!IDAnagraficaTecnicoOperativo), "IDRV_POCategoriaFase")
            rsIntervento!IDRV_POTipoAddebito = GET_PARAMETRI_TEC_OPE(fnNotNullN(rsIntervento!IDAnagraficaTecnicoOperativo), "IDRV_POTipoAddebito")
            rsIntervento!IDRV_POTipoClasseIntervento = GET_PARAMETRI_TEC_OPE(fnNotNullN(rsIntervento!IDAnagraficaTecnicoOperativo), "IDRV_POTipoClasseIntervento")
            If fnNotNullN(rsIntervento!IDRV_POCategoriaIntervento) = 0 Then
                rsIntervento!IDRV_POCategoriaIntervento = IDCategoriaIntervento
            End If
        rsIntervento.Update
            
        
        DataInizioServizio = DateAdd("m", IIf((MeseRic = ""), 0, MeseRic), DataInizioServizio)
        DataInizioServizio = DateAdd("d", IIf((GiornoRic = ""), 0, GiornoRic), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((SettimanaRic = ""), 0, SettimanaRic), DataInizioServizio)
    Wend
End If
rsIntervento.Close
Set rsIntervento = Nothing

End Sub

Private Function GET_COSTRUZIONE_DATA_PERS(Giorno As String, Mese As String) As String
Dim GiornoInizio As String
Dim MeseInizio As String


If Len(Giorno) = 1 Then
    GiornoInizio = "0" & Giorno
ElseIf Len(Giorno) = 0 Then
    GiornoInizio = "01" '& Giorno
Else
    GiornoInizio = Giorno
End If
    
If Len(Mese) = 1 Then
    MeseInizio = "0" & Mese
ElseIf Len(Mese) = 0 Then
    MeseInizio = "01" '& Giorno
Else
    MeseInizio = Mese
End If

GET_COSTRUZIONE_DATA_PERS = GiornoInizio & "/" & MeseInizio & "/"
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

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_INTERVENTO = 1
Else
    GET_NUMERO_INTERVENTO = fnNotNullN(rs!Numero) + 1
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_PARAMETRI_AZIENDA(IDAzienda As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POParametriAzienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    LINK_TIPO_ANA_TEC_INT = 0
    LINK_STATO_INT_NUOVO = 0
    LINK_STATO_INT_CHIUSO = 0
    
    LINK_TIPO_ANA_TEC_FASE = 0
    LINK_STATO_FASE_NUOVA = 0
    LINK_STATO_FASE_CHIUSA = 0
    LINK_TIPO_FASE_ELA = 0
    LINK_TIPO_FASE_MANUALE = 0
    If LINK_TIPO_TASCA = 0 Then
        LINK_TIPO_TASCA = 0
    End If
    VOISPEED = 0
    FLAG_RIPORTA_TEC_CONTRATTO = 0
    LINK_TIPO_ANA_TEC_CONTRATTO = 0
    LINK_TIPO_FASE_DA_PRODOTTO = 0
Else
    LINK_TIPO_ANA_TEC_INT = fnNotNullN(rs!IDTipoAnagraficaTecnicoIntRif)
    LINK_STATO_INT_NUOVO = fnNotNullN(rs!IDRV_POStatoInterventoInserimento)
    LINK_STATO_INT_CHIUSO = fnNotNullN(rs!IDRV_POStatoInterventoChiuso)
    
    LINK_TIPO_ANA_TEC_FASE = fnNotNullN(rs!IDTipoAnagraficaTecnicoFaseRif)
    LINK_STATO_FASE_NUOVA = fnNotNullN(rs!IDRV_POStatoFaseInserimento)
    LINK_STATO_FASE_CHIUSA = fnNotNullN(rs!IDRV_POStatoFaseChiusa)
    LINK_TIPO_FASE_ELA = fnNotNullN(rs!IDRV_POTipoFaseInterventoEla)
    LINK_TIPO_FASE_MANUALE = fnNotNullN(rs!IDRV_POTipoFaseInterventoMan)
    If LINK_TIPO_TASCA = 0 Then
        LINK_TIPO_TASCA = fnNotNullN(rs!IDRV_POTipoTasca)
    End If
    VOISPEED = fnNotNullN(rs!UtilizzaVoispeed)
    FLAG_RIPORTA_TEC_CONTRATTO = fnNotNullN(rs!RiportaTecContratto)
    LINK_TIPO_ANA_TEC_CONTRATTO = fnNotNullN(rs!IDTipoAnagraficaContratto)
    LINK_TIPO_FASE_DA_PRODOTTO = fnNotNullN(rs!IDStatoInterventoDaProdotto)
    If LINK_TIPO_FASE_DA_PRODOTTO = 0 Then LINK_TIPO_FASE_DA_PRODOTTO = LINK_TIPO_FASE_ELA
End If

rs.CloseResultset
Set rs = Nothing
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

Private Function GET_DESCRIZIONE_ARTICOLO(IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Articolo FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_ARTICOLO = ""
Else
    GET_DESCRIZIONE_ARTICOLO = fnNotNull(rs!Articolo)
End If

rs.CloseResultset
Set rs = Nothing
End Function


Private Function GET_DATI_CONTRATTO(IDContratto As Long, sSQL As String) As String
Dim sSQLContratto As String
Dim rs As DmtOleDbLib.adoResultset

GET_DATI_CONTRATTO = sSQL

sSQLContratto = "SELECT * FROM RV_POContratto "
sSQLContratto = sSQLContratto & "WHERE IDRV_POContratto=" & IDContratto

Set rs = CnDMT.OpenResultset(sSQLContratto)

If Not rs.EOF Then

    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(rs!IDAnagraficaAgente) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNormString(rs!AnagraficaAgente) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNormString(rs!NomeAgente) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(rs!IDTipoAnagraficaCommesso) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(rs!IDAnagraficaCommesso) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNormString(rs!AnagraficaCommesso) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNormString(rs!NomeCommesso) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(rs!IDAnagraficaAmministratore) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(rs!IDTipoAnagraficaAmministratore) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNormString(rs!AnagraficaAmministratore) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNormString(rs!NomeAmministratore) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(rs!RitenutaAcconto) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(rs!IDContrattoBancario) & ")"
Else
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(0) & ", "
    GET_DATI_CONTRATTO = sSQL & fnNormString("") & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNormString("") & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(0) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(rsEla!IDAnagraficaCommesso) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNormString("") & ", " & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNormString("") & ", " & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(rsEla!IDAnagraficaAmministratore) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(rsEla!IDTipoAnagraficaAmministratore) & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNormString("") & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNormString("") & ", "
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(0) & ","
    GET_DATI_CONTRATTO = GET_DATI_CONTRATTO & fnNotNullN(0) & ")"
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

Set rs = CnDMT.OpenResultset(sSQL)

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

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_STAGIONE = 0
Else
    GET_LINK_STAGIONE = fnNotNullN(rs!IDRV_POStagione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_NUOVO_CONTRATTO(IDContratto As Long, rstmp As DmtOleDbLib.adoResultset, rsNew As ADODB.Recordset) As Long
On Error GoTo ERR_GET_LINK_NUOVO_CONTRATTO
Dim sSQL As String
Dim LINK_LOCAL_CONTRATTO As Long
Dim rsContratto As DmtOleDbLib.adoResultset

GET_LINK_NUOVO_CONTRATTO = 0

sSQL = "SELECT * FROM RV_POContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rsContratto = CnDMT.OpenResultset(sSQL)

If Not rsContratto.EOF Then
    rsNew.AddNew
        LINK_LOCAL_CONTRATTO = fnGetNewKey("RV_POContratto", "IDRV_POContratto")
        rsNew!IDRV_POContratto = LINK_LOCAL_CONTRATTO
        rsNew!IDRV_POContrattoPadre = fnNotNullN(rstmp!IDRV_POContrattoPadre)
        rsNew!NumeroRinnovo = fnNotNullN(rstmp!NumeroRinnovo)
        rsNew!IDAnagrafica = fnNotNullN(rsContratto!IDAnagrafica)
        rsNew!IDAnagraficaFatturazione = fnNotNullN(rsContratto!IDAnagraficaFatturazione)
        rsNew!IDSitoPerAnagrafica = fnNotNullN(rsContratto!IDSitoPerAnagrafica)
        rsNew!IDTipoContratto = fnNotNullN(rsContratto!IDTipoContratto)
        rsNew!DataStipula = rsContratto!DataStipula
        rsNew!ImportoContrattoStipula = fnNotNullN(rsContratto!ImportoContrattoStipula)
        rsNew!ImportoContrattoAttuale = fnNotNullN(rstmp!ImportoContratto)
        rsNew!DataDecorrenza = rstmp!DataDecorrenza
        rsNew!DescrizioneTipoContratto = fnNotNull(rsContratto!DescrizioneTipoContratto)
        rsNew!IDDurataContratto = fnNotNullN(rstmp!IDDurataContratto)
        rsNew!IDTipoRinnovo = fnNotNullN(rstmp!IDTipoRinnovo)
        rsNew!DataScadenza = rstmp!DataScadenzaContratto
        rsNew!DataScadenzaPerRinnovo = rstmp!DataScadenzaPerRinnovo
        rsNew!MeseScadenza = DatePart("m", rstmp!DataScadenzaPerRinnovo)
        rsNew!AnnoScadenza = DatePart("yyyy", rstmp!DataScadenzaPerRinnovo)
        rsNew!IDRateizzazione = fnNotNullN(rstmp!IDRateizzazione)
        rsNew!IDPagamentoRate = fnNotNullN(rstmp!IDPagamento)
        rsNew!Disdetta = 0
        rsNew!AttivoPassivo = 1
        rsNew!RinnovoAutomatico = fnNotNullN(rstmp!RinnovoAutomatico)
        rsNew!AdeguamentoIstat = fnNotNullN(rstmp!AdeguamentoIstat)
        rsNew!IDIstat = fnNotNullN(rstmp!IDIstat)
        rsNew!IDAzienda = fnNotNullN(rsContratto!IDAzienda)
        rsNew!IDFiliale = fnNotNullN(rsContratto!IDFiliale)
        rsNew!NonFatturare = fnNotNullN(rsContratto!NonFatturare)
        rsNew!IDCodiceConto = fnNotNullN(rsContratto!IDCodiceConto)
        rsNew!CodiceConto = fnNotNull(rsContratto!CodiceConto)
        rsNew!DescrizioneConto = fnNotNull(rsContratto!DescrizioneConto)
        rsNew!IDUtentePerInserimento = TheApp.IDUser
        rsNew!DataInserimento = Date
        rsNew!NumeroLicenze = fnNotNullN(rsContratto!NumeroLicenze)
        rsNew!IDAnagraficaAgente = fnNotNullN(rsContratto!IDAnagraficaAgente)
        rsNew!AnagraficaAgente = fnNotNull(rsContratto!AnagraficaAgente)
        rsNew!NomeAgente = fnNotNull(rsContratto!NomeAgente)
        rsNew!IDAnagraficaCommesso = fnNotNullN(rsContratto!IDAnagraficaCommesso)
        rsNew!AnagraficaCommesso = fnNotNull(rsContratto!AnagraficaCommesso)
        rsNew!NomeCommesso = fnNotNull(rsContratto!NomeCommesso)
        rsNew!NumeroProtocollo = fnNotNull(rsContratto!NumeroProtocollo)
        rsNew!IDContrattoBancario = fnNotNullN(rsContratto!IDContrattoBancario)
        rsNew!IDAnagraficaAmministratore = fnNotNullN(rsContratto!IDAnagraficaAmministratore)
        rsNew!IDTipoAnagraficaAmministratore = fnNotNullN(rsContratto!IDTipoAnagraficaAmministratore)
        rsNew!AnagraficaAmministratore = fnNotNull(rsContratto!AnagraficaAmministratore)
        rsNew!NomeAmministratore = fnNotNull(rsContratto!NomeAmministratore)
        rsNew!RitenutaAcconto = fnNotNullN(rsContratto!RitenutaAcconto)
        rsNew!IDTipoAnagraficaCommesso = fnNotNullN(rsContratto!IDTipoAnagraficaCommesso)
        rsNew!IDRV_POTipoDurataAssistenza = fnNotNullN(rstmp!IDRV_POTipoDurataAssistenza)
        rsNew!DataFineAssistenza = rstmp!DataFineAssistenza
        rsNew!Maggiorazione = fnNotNullN(rstmp!Maggiorazione)
        rsNew!IDAdeguamentoIstat = fnNotNullN(rstmp!IDIstat)
        rsNew!ContrattoAttuale = 1
        rsNew!IDRaggruppamentoFatturato = fnNotNullN(rstmp!IDRaggruppamentoFatturato)
        rsNew!IDRV_POTipoClassificazioneContratto = fnNotNullN(rstmp!IDRV_POTipoClassificazioneContratto)
        rsNew!IDArticoloContratto = fnNotNullN(rstmp!IDArticoloContratto)
        rsNew!NoteFattura = fnNotNull(rsContratto!NoteFattura)
        rsNew!Annotazioni = fnNotNull(rsContratto!Annotazioni)
        rsNew!AnnoContratto = Year(rsNew!DataDecorrenza)
        rsNew!NumeroContratto = GET_NUMERO_CONTRATTO(rsNew!AnnoContratto)
        rsNew!Chiuso = 0
        rsNew!NumeroGiorniPrimaRata = fnNotNullN(rsContratto!NumeroGiorniPrimaRata)
        rsNew!IDRV_POTipoImpostazioneContratto = fnNotNullN(rsContratto!IDRV_POTipoImpostazioneContratto)
        rsNew!GeneraRatePerProdotto = fnNotNullN(rsContratto!GeneraRatePerProdotto)
        rsNew!FatturazioneRicorrente = fnNotNullN(rsContratto!FatturazioneRicorrente)
        rsNew!TotaleContrattoDaProdotti = fnNotNullN(rsContratto!TotaleContrattoDaProdotti)
        rsNew!Offerta = 0
        rsNew!FineContratto = rstmp!FineContratto
        If rsNew!FineContratto = 0 Then
            rsNew!IDDurataContrattoProssimoRinnovo = rsContratto!IDDurataContrattoProssimoRinnovo
            rsNew!DataScadenzaSecondoContratto = rsContratto!DataScadenzaSecondoContratto
        End If
        rsNew!DataPrimaDecorrenza = rsContratto!DataPrimaDecorrenza
        
        
    'rsNew.Update
    
    INSERIMENTO_PRODOTTI IDContratto, rsNew!IDRV_POContratto, rsNew!IDRV_POContrattoPadre, rsNew!DataDecorrenza, rsNew!DataScadenzaPerRinnovo, GET_PERCENTUALE_ISTAT(rsNew!IDIstat)
    
    If rsNew!TotaleContrattoDaProdotti = 1 Then
        rsNew!ImportoContrattoAttuale = GET_TOTALE_CONTRATTO(rsNew!IDRV_POContratto)
        
    End If
    
    rsNew.Update
    
    ElaborazioneRate rsNew, fnNotNullN(rsNew!IDRV_POContratto), fnNotNull(rsNew!DataDecorrenza), fnNotNullN(rsNew!ImportoContrattoAttuale), fnNotNullN(rsNew!IDPagamentoRate)
    
    INSERIMENTO_ADEGUAMENTI rsNew, IDContratto, rsNew!IDRV_POContratto, rsNew!IDRV_POContrattoPadre, rsNew!NumeroRinnovo
    
    RINNOVO_SERVIZI fnNotNullN(rsNew!IDRV_POContratto), fnNotNullN(rstmp!IDRV_POContratto), TheApp.IDFirm, fnNotNullN(rstmp!IDRV_POContrattoPadre)
    
    RINNOVO_SERVIZI_PRODOTTI rsNew!IDRV_POContratto, IDContratto
    
    CREA_TMP_INTERVENTI rsNew
    
End If

rsContratto.CloseResultset
Set rsContratto = Nothing

Exit Function

ERR_GET_LINK_NUOVO_CONTRATTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_NUOVO_CONTRATTO"

End Function

Private Function GET_NUMERO_CONTRATTO(AnnoContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT COUNT(NumeroContratto) AS Numero "
sSQL = sSQL & "FROM RV_POContratto "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND AnnoContratto=" & AnnoContratto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_CONTRATTO = 1
Else
    GET_NUMERO_CONTRATTO = fnNotNullN(rs!Numero) + 1
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub INSERIMENTO_ADEGUAMENTI(rsNuovoContratto As ADODB.Recordset, IDContrattoOLD As Long, IDContrattoNew As Long, IDContrattoPadre As Long, NumeroRinnovo As Long)
On Error GoTo ERR_INSERIMENTO_ADEGUAMENTI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset

''''''''''''''APERTURA RECORDSET PER INSERIMENTO ADEGUAMENTO''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POContrattoAdeguamento "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContrattoNew

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''APERTURA RECORDSET PER ADEGUAMENTI VECCHI''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POContrattoAdeguamento "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContrattoOLD
sSQL = sSQL & " AND IDRV_POTipoAdeguamento=2"
Set rs = CnDMT.OpenResultset(sSQL)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

While Not rs.EOF
    If (fnNotNullN(rs!RiportaProssimoRinnovo) = 1) Then
        If fnNotNullN(rs!DataFineAdeguamento) = 0 Then
            rsNew.AddNew
                rsNew!IDRV_POContrattoAdeguamento = fnGetNewKey("RV_POContrattoAdeguamento", "IDRV_POContrattoAdeguamento")
                rsNew!IDRV_POContratto = IDContrattoNew
                rsNew!IDRV_POContrattoPadre = IDContrattoPadre
                rsNew!DataStipula = rs!DataStipula
                rsNew!DataDecorrenza = rsNuovoContratto!DataDecorrenza
                rsNew!AdeguamentoIstat = fnNotNullN(rs!AdeguamentoIstat)
                rsNew!RiportaProssimoRinnovo = Abs(fnNotNullN(rs!RiportaProssimoRinnovo))
                rsNew!AdeguaContrattoAttuale = Abs(fnNotNullN(rs!AdeguaContrattoAttuale))
                rsNew!Annotazioni = fnNotNull(rs!Annotazioni)
                rsNew!IDRV_POTipoAdeguamento = fnNotNullN(rs!IDRV_POTipoAdeguamento)
                rsNew!IDArticolo = fnNotNullN(rs!IDArticolo)
                rsNew!NumeroProtocollo = fnNotNull(rs!NumeroProtocollo)
                rsNew!NumeroAdeguamento = fnNotNullN(rs!NumeroAdeguamento)
                rsNew!DescrizioneAdeguamento = "Adeguamento numero " & rsNew!NumeroAdeguamento & " - " & NumeroRinnovo
                rsNew!DescrizionePerFatturazione = fnNotNull(rs!DescrizionePerFatturazione)
                rsNew!NoCalcPeriodoFatt = fnNotNullN(rs!NoCalcPeriodoFatt)
                GET_IMPORTO_ADEGUAMENTO rsNew, fnNotNullN(rs!Importo)
            
            rsNew.Update
            
            GENERA_RATE_ADEGUAMENTO rsNew!IDRV_POContrattoAdeguamento, rsNuovoContratto!IDDurataContratto, rsNew!IDRV_POContratto, rsNew!IDRV_POContrattoPadre, _
            DatePart("yyyy", fnNotNull(rsNuovoContratto!DataDecorrenza)), rsNew!DataDecorrenza, rsNuovoContratto!DataScadenzaPerRinnovo, rsNew!Importo, rsNuovoContratto!IDPagamentoRate, rsNew!NumeroAdeguamento, rsNew!NumeroProtocollo, rsNew!IDArticolo, fnNotNull(rsNew!DescrizionePerFatturazione), fnNotNullN(rsNew!NoCalcPeriodoFatt)
            
        Else
            rsNew.AddNew
                rsNew!IDRV_POContrattoAdeguamento = fnGetNewKey("RV_POContrattoAdeguamento", "IDRV_POContrattoAdeguamento")
                rsNew!IDRV_POContratto = IDContrattoNew
                rsNew!IDRV_POContrattoPadre = IDContrattoPadre
                rsNew!DataStipula = rs!DataStipula
                rsNew!AdeguamentoIstat = fnNotNullN(rs!AdeguamentoIstat)
                rsNew!RiportaProssimoRinnovo = Abs(fnNotNullN(rs!RiportaProssimoRinnovo))
                rsNew!AdeguaContrattoAttuale = Abs(fnNotNullN(rs!AdeguaContrattoAttuale))
                rsNew!Annotazioni = fnNotNull(rs!Annotazioni)
                rsNew!IDRV_POTipoAdeguamento = fnNotNullN(rs!IDRV_POTipoAdeguamento)
                rsNew!IDArticolo = fnNotNullN(rs!IDArticolo)
                rsNew!NumeroProtocollo = fnNotNull(rs!NumeroProtocollo)
                rsNew!NumeroAdeguamento = fnNotNullN(rs!NumeroAdeguamento)
                rsNew!DescrizioneAdeguamento = "Adeguamento numero " & rsNew!NumeroAdeguamento & " - " & NumeroRinnovo
                rsNew!DescrizionePerFatturazione = fnNotNull(rs!DescrizionePerFatturazione)
                rsNew!DataDecorrenza = DateAdd("yyyy", 1, rs!DataDecorrenza)
                rsNew!DataFineAdeguamento = DateAdd("yyyy", 1, rs!DataFineAdeguamento)
                rsNew!IDRateizzazione = rs!IDRateizzazione
                rsNew!NoCalcPeriodoFatt = fnNotNullN(rs!NoCalcPeriodoFatt)
                
                GET_IMPORTO_ADEGUAMENTO rsNew, fnNotNullN(rs!Importo)
                
                
            rsNew.Update
        
            ElaborazioneRate rsNuovoContratto, IDContrattoNew, rsNew!DataDecorrenza, rsNew!Importo, rsNuovoContratto!IDPagamentoRate, rsNew!IDRV_POContrattoAdeguamento, fnNotNullN(rsNew!NoCalcPeriodoFatt)
        
        End If
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

rsNew.Close
Set rsNew = Nothing
Exit Sub
ERR_INSERIMENTO_ADEGUAMENTI:
    MsgBox Err.Description, vbCritical, "INSERIMENTO_ADEGUAMENTI"
End Sub
Private Sub GET_IMPORTO_ADEGUAMENTO(rsAdeg As ADODB.Recordset, ImportoOLD As Double)

If rsAdeg!AdeguamentoIstat = 0 Then
    rsAdeg!Importo = ImportoOLD
    rsAdeg!IDIstat = 0
    rsAdeg!MaggiorazioneIstat = 0
Else
    rsAdeg!Importo = ImportoOLD + ((ImportoOLD / 100) * PercentualeIstat)
    rsAdeg!IDIstat = Link_Istat
    rsAdeg!MaggiorazioneIstat = ((ImportoOLD / 100) * PercentualeIstat)
End If

End Sub
Private Sub GENERA_RATE_ADEGUAMENTO(IDAdeguamentoContratto As Long, IDDurataContratto As Long, IDContratto As Long, IDContrattoPadre As Long, AnnoContratto As Long, DataDecorrenzaAdeg As String, DataScadenzaContratto As String, ImportoAdeg As Double, IDPagamentoRate As Long, NumeroAdeguamento As Long, ProtocolloAdeg As String, IDArticolo As Long, DescrizionePerFatturazione As String, Optional NoCalcolaPeriodo As Long = 0)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim NumeroGiorniContratto As Long
Dim NumeroGiorniRimanenti As Double
Dim ImportoGiornalieroAdeg As Double
Dim ImportoAdeguamentoTotale As Double
Dim NumeroRateDaPagare As Long
Dim ImportoAdegPerRata As Double
Dim NumeroRataElaborata As Long
Dim ImportoAdegProgressivo As Double
Dim ImportoDaRegistrare As Double
Dim ImportoAdegPerRataComp As Double
Dim ImportoAdegProgressivoComp As Double

Dim Periodo As String
Dim ArrayRate As String
Dim SplitArrayRate() As String
Dim I As Integer

ArrayRate = ""

'ELIMINAZIONE RATE DELL'ADEGUAMENTO DEL CONTRATTO'''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento=" & IDAdeguamentoContratto
CnDMT.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Year(DataScadenzaContratto) Mod 4 <> 0 Then
    DataScadenzaContratto = DateAdd("d", 1, DataScadenzaContratto)

    If Year(DataDecorrenzaAdeg) Mod 4 <> 0 Then
        DataDecorrenzaAdeg = DataDecorrenzaAdeg
    Else
        If DateDiff("d", DataDecorrenzaAdeg, "29/02/" & Year(DataDecorrenzaAdeg)) >= 0 Then
            DataDecorrenzaAdeg = DateAdd("d", 1, DataDecorrenzaAdeg)
        Else
            DataDecorrenzaAdeg = DataDecorrenzaAdeg
        End If
    End If

    NumeroGiorniContratto = 365
    NumeroGiorniRimanenti = DateDiff("d", DataDecorrenzaAdeg, DataScadenzaContratto)
    
    ImportoGiornalieroAdeg = ImportoAdeg / NumeroGiorniContratto
    
    If DatePart("yyyy", DataScadenzaContratto) Mod 4 = 0 Then
        ImportoAdeguamentoTotale = ImportoGiornalieroAdeg * (NumeroGiorniRimanenti)
    Else
        ImportoAdeguamentoTotale = ImportoGiornalieroAdeg * NumeroGiorniRimanenti
    End If
    
Else
    If DateDiff("d", DataScadenzaContratto, "29/02/" & Year(DataScadenzaContratto)) <= 0 Then
        DataScadenzaContratto = DataScadenzaContratto
        
        NumeroGiorniContratto = 365
        NumeroGiorniRimanenti = DateDiff("d", DataDecorrenzaAdeg, DataScadenzaContratto)
        
        ImportoGiornalieroAdeg = ImportoAdeg / NumeroGiorniContratto
        
        If DatePart("yyyy", DataScadenzaContratto) Mod 4 = 0 Then
            ImportoAdeguamentoTotale = ImportoGiornalieroAdeg * (NumeroGiorniRimanenti)
        Else
            ImportoAdeguamentoTotale = ImportoGiornalieroAdeg * NumeroGiorniRimanenti
        End If
    Else
        DataScadenzaContratto = DateAdd("d", 1, DataScadenzaContratto)
        If Year(DataDecorrenzaAdeg) Mod 4 <> 0 Then
            DataDecorrenzaAdeg = DataDecorrenzaAdeg
        Else
            If DateDiff("d", DataDecorrenzaAdeg, "29/02/" & Year(DataDecorrenzaAdeg)) >= 0 Then
                DataDecorrenzaAdeg = DateAdd("d", 1, DataDecorrenzaAdeg)
            Else
                DataDecorrenzaAdeg = DataDecorrenzaAdeg
            End If
        End If

        NumeroGiorniContratto = 365
        NumeroGiorniRimanenti = DateDiff("d", DataDecorrenzaAdeg, DataScadenzaContratto)
        
        ImportoGiornalieroAdeg = ImportoAdeg / NumeroGiorniContratto
        
        If DatePart("yyyy", DataScadenzaContratto) Mod 4 = 0 Then
            ImportoAdeguamentoTotale = ImportoGiornalieroAdeg * (NumeroGiorniRimanenti)
        Else
            ImportoAdeguamentoTotale = ImportoGiornalieroAdeg * NumeroGiorniRimanenti
        End If
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT COUNT(IDRV_PORateContratto) AS NumeroRecord "
sSQL = sSQL & "FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Fatturata=0"
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    NumeroRateDaPagare = 0
Else
    NumeroRateDaPagare = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=0"

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If NumeroRateDaPagare > 0 Then
    ImportoAdegPerRataComp = ImportoAdeguamentoTotale / NumeroRateDaPagare
    ImportoAdegPerRata = FormatNumber(ImportoAdegPerRataComp, 2)
    ImportoAdegProgressivo = ImportoAdegPerRata
    ImportoAdegProgressivoComp = ImportoAdegPerRataComp
    
    sSQL = "SELECT * FROM RV_PORateContratto "
    sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
    sSQL = sSQL & " AND Fatturata=0"
    sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"
    sSQL = sSQL & " ORDER BY DataRata "
    
    Set rs = CnDMT.OpenResultset(sSQL)

    While Not rs.EOF
        
        If ImportoAdegProgressivoComp >= ImportoAdeguamentoTotale Then
            ImportoDaRegistrare = FormatNumber((ImportoAdeguamentoTotale - (ImportoAdegProgressivo - ImportoAdegPerRataComp)), 2)
        Else
            ImportoDaRegistrare = ImportoAdegPerRata
        End If
        
        rsNew.AddNew
            rsNew!IDRV_PORateContratto = fnGetNewKey("RV_PORateContratto", "IDRV_PORateContratto")
            rsNew!IDRV_POContratto = IDContratto
            rsNew!NumeroRata = fnNotNullN(rs!NumeroRata)
            rsNew!DataRata = rs!DataRata
            rsNew!IDPagamentoRata = fnNotNullN(rs!IDPagamentoRata)
            rsNew!ImportoRata = ImportoDaRegistrare
            rsNew!Fatturata = 0
            rsNew!NonFatturare = 0
            rsNew!IDOggettoCollegato = 0
            rsNew!IDTipoOggettoCollegato = 0
            rsNew!Mese = DatePart("m", fnNotNull(rs!DataRata))
            rsNew!Anno = DatePart("yyyy", fnNotNull(rs!DataRata))
            rsNew!Periodo = Mid(DescrizionePerFatturazione, 1, 250)
            rsNew!Manuale = 0
            rsNew!ContrattoAttuale = 1
            rsNew!IDRV_POContrattoPadre = IDContrattoPadre
            rsNew!IDRV_POContrattoAdeguamento = IDAdeguamentoContratto
            rsNew!IDArticolo = IDArticolo
            rsNew!IDTipoOggetto = fnGetTipoOggetto("RV_PORateContratto")
            rsNew!IDOggetto = GET_LINK_OGGETTO(0, rsNew!IDTipoOggetto, rsNew!NumeroRata, rsNew!DataRata)
            rsNew!DataInizioPeriodo = rs!DataInizioPeriodo
            rsNew!DataFinePeriodo = rs!DataFinePeriodo
            
            If Len(ArrayRate) > 0 Then
                ArrayRate = ArrayRate & "|"
            End If
            ArrayRate = ArrayRate & fnNotNullN(rsNew!IDRV_PORateContratto)
            
        rsNew.Update
        ImportoAdegProgressivo = ImportoAdegProgressivo + ImportoAdegPerRata
        ImportoAdegProgressivoComp = ImportoAdegProgressivoComp + ImportoAdegPerRataComp
        
        
    rs.MoveNext
    Wend
    rs.CloseResultset
    Set rs = Nothing
Else
    rsNew.AddNew
        rsNew!IDRV_PORateContratto = fnGetNewKey("RV_PORateContratto", "IDRV_PORateContratto")
        rsNew!IDRV_POContratto = IDContratto
        rsNew!NumeroRata = GET_NUMERO_RATA_CONTRATTO(IDContratto)
        rsNew!DataRata = DataDecorrenzaAdeg
        rsNew!IDPagamentoRata = IDPagamentoRate
        rsNew!ImportoRata = ImportoAdeguamentoTotale
        rsNew!Fatturata = 0
        rsNew!IDOggettoCollegato = 0
        rsNew!IDTipoOggettoCollegato = 0
        rsNew!Mese = DatePart("m", DataDecorrenzaAdeg)
        rsNew!Anno = DatePart("yyyy", DataDecorrenzaAdeg)
        rsNew!Periodo = Mid(DescrizionePerFatturazione, 1, 255)
        rsNew!Manuale = 0
        rsNew!ContrattoAttuale = 1
        rsNew!IDRV_POContrattoPadre = IDContrattoPadre
        rsNew!IDRV_POContrattoAdeguamento = IDAdeguamentoContratto
        rsNew!IDArticolo = IDArticolo
        rsNew!IDTipoOggetto = fnGetTipoOggetto("RV_PORateContratto")
        rsNew!IDOggetto = GET_LINK_OGGETTO(0, rsNew!IDTipoOggetto, rsNew!NumeroRata, rsNew!DataRata)
        rsNew!DataInizioPeriodo = DataDecorrenzaAdeg
        rsNew!DataFinePeriodo = DataScadenzaContratto
        If Len(ArrayRate) > 0 Then
            ArrayRate = ArrayRate & "|"
        End If
        ArrayRate = ArrayRate & fnNotNullN(rsNew!IDRV_PORateContratto)
    rsNew.Update
End If

rsNew.Close
Set rsNew = Nothing

SplitArrayRate = Split(ArrayRate, "|")

For I = 0 To UBound(SplitArrayRate)
    
    If NoCalcolaPeriodo = 0 Then
        Periodo = GET_STRINGA_PERIODO_ADEG(2, TheApp.Branch, IDContratto, CLng(SplitArrayRate(I)), GET_LINK_ADEGUAMENTO_DA_RATA(CLng(SplitArrayRate(I))), 0)
    Else
        Periodo = GET_DESCRIZIONE_ARTICOLO(IDArticolo)
    End If
    
    sSQL = "UPDATE RV_PORateContratto SET "
    sSQL = sSQL & "Periodo=" & fnNormString(Mid(Periodo, 1, 250))
    sSQL = sSQL & "WHERE IDRV_PORateContratto=" & SplitArrayRate(I)
    
    CnDMT.Execute sSQL
    
Next


End Sub
Private Sub INSERIMENTO_PRODOTTI(IDContrattoOLD As Long, IDContrattoNew As Long, IDContrattoPadre As Long, DataInizioContratto As String, DataFineContratto As String, percMagg As Double)
On Error GoTo ERR_INSERIMENTO_PRODOTTI
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim I As Long

sSQL = "SELECT * FROM RV_POContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContrattoNew
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

sSQL = "SELECT * FROM RV_POContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContrattoOLD
sSQL = sSQL & " AND Dismesso=0"
sSQL = sSQL & " AND ((NonRinnovare=0) OR (NonRinnovare IS NULL))"

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection

While Not rs.EOF
    rsNew.AddNew
        
        rsNew!IDRV_POContrattoProdotti = fnGetNewKey("RV_POContrattoProdotti", "IDRV_POContrattoProdotti")
        rsNew!IDRV_POContratto = IDContrattoNew
        rsNew!IDRV_POContrattoPadre = IDContrattoPadre
        rsNew!IDArticolo = rs!IDArticolo
        rsNew!ValoreIndentificativo = rs!ValoreIndentificativo
        rsNew!DescrizioneAggiuntiva = rs!DescrizioneAggiuntiva
        rsNew!Annotazioni = rs!Annotazioni
        rsNew!Quantita = rs!Quantita
        rsNew!ImportoUnitario = rs!ImportoUnitario
        rsNew!ImportoUnitarioScontato = rs!ImportoUnitarioScontato
        rsNew!IDIva = rs!IDIva
        rsNew!Imponibile = rs!Imponibile
        rsNew!ImportoIva = rs!ImportoIva
        rsNew!TotaleRiga = rs!TotaleRiga
        rsNew!DataInizioGaranzia = rs!DataInizioGaranzia
        rsNew!DataFineGaranzia = rs!DataFineGaranzia
        rsNew!Dismesso = rs!Dismesso
        rsNew!Sconto1 = rs!Sconto1
        rsNew!Sconto2 = rs!Sconto2
        rsNew!DataDismesso = rs!DataDismesso
        rsNew!IDRV_POContrattoAdeguamento = rs!IDRV_POContrattoAdeguamento
        rsNew!IDArticoloServizio = rs!IDArticoloServizio
        rsNew!IDRV_POProdotto = rs!IDRV_POProdotto
        rsNew!IDRV_POContrattoProdottiRinnovo = fnNotNullN(rs!IDRV_POContrattoProdotti)
        rsNew!DataInizioPeriodo = DataInizioContratto
        rsNew!OraInizioPeriodo = rs!OraInizioPeriodo
        rsNew!DataFinePeriodo = DataFineContratto
        rsNew!OraFinePeriodo = rs!OraFinePeriodo
        rsNew!QuantitaPeriodo = rs!QuantitaPeriodo
        rsNew!IDRV_POUnitaDiMisuraPeriodo = rs!IDRV_POUnitaDiMisuraPeriodo
        rsNew!ImportoComplessivo = rs!ImportoComplessivo
        rsNew!IDUnitaDiMisuraArticolo = rs!IDUnitaDiMisuraArticolo
        rsNew!IDListino = rs!IDListino
        rsNew!QuantitaArticolo = rs!QuantitaArticolo
        rsNew!ScontoAImporto = rs!ScontoAImporto
        rsNew!IDRV_POTipoPeriodo = rs!IDRV_POTipoPeriodo
        rsNew!AliquotaIva = rs!AliquotaIva
        rsNew!QuantitaEffettiva = rs!QuantitaEffettiva
        rsNew!EscludiGiorniFestivi = rs!EscludiGiorniFestivi
        rsNew!EscludiSabato = rs!EscludiSabato
        rsNew!Conducente = rs!Conducente
        rsNew!ACorpo = rs!ACorpo
        rsNew!TestoStampa = rs!TestoStampa
        rsNew!IDAnagraficaOperatore = rs!IDAnagraficaOperatore
        rsNew!IDRV_POInterventoRigheDett = rs!IDRV_POInterventoRigheDett
        rsNew!IDRV_POContatoreRilevamenti = rs!IDRV_POContatoreRilevamenti
        rsNew!IDRV_POContrattoProdottiCollegato = rs!IDRV_POContrattoProdottiCollegato
        rsNew!NonRateizzare = rs!NonRateizzare
        rsNew!NonRinnovare = rs!NonRinnovare
        rsNew!AnnotazioniPerIntervento = rs!AnnotazioniPerIntervento
        rsNew!IDRateizzazione = rs!IDRateizzazione
        
        GET_TOTALE_RIGA rsNew, percMagg
        
    rsNew.Update
    
    INSERIMENTO_PRODOTTI_CONF_CONT fnNotNullN(rs!IDRV_POContrattoProdotti), fnNotNullN(rsNew!IDRV_POContrattoProdotti)
    
    
rs.MoveNext
Wend

rsNew.Close
Set rsrsnew = Nothing

rs.Close
Set rs = Nothing
Exit Sub
ERR_INSERIMENTO_PRODOTTI:
    MsgBox Err.Description, vbCritical, "INSERIMENTO_PRODOTTI"
End Sub
Private Function GET_NUMERO_RATA_CONTRATTO(IDContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(NumeroRata) as MaxNumeroRata "
sSQL = sSQL & "FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_RATA_CONTRATTO = 1
Else
    GET_NUMERO_RATA_CONTRATTO = fnNotNullN(rs!MaxNumeroRata) + 1
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_OGGETTO(IDOggetto As Long, IDTipoOggetto As Long, NumeroRata As Long, DataScadenzaRata As String) As Long
On Error GoTo ERR_GET_LINK_OGGETTO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim IDOggettoLocal As Long
Dim IDFunzione As Long

GET_LINK_OGGETTO = 0

IDFunzione = GET_LINK_FUNZIONE(IDTipoOggetto)

If IDFunzione = 0 Then Exit Function

sSQL = "SELECT * FROM Oggetto "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = True Then
    rs.AddNew
        rs!IDTipoOggetto = IDTipoOggetto
        rs!IDFunzione = IDFunzione
        rs!IDAzienda = TheApp.IDFirm
        rs!IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.Branch)
        rs!IDSezionale = 0
        rs!Oggetto = GET_DESCRIZIONE_FUNZIONE(IDFunzione)
        rs!DataEmissione = DataScadenzaRata
        rs!Numero = NumeroRata
        rs!DataUltimaVariazione = Date
        rs!IDUtenteUltimaVariazione = TheApp.IDUser
        rs!VirtualDelete = 0
        rs!IDOggetto = fnGetNewKey("Oggetto", "IDOggetto")
        GET_LINK_OGGETTO = rs!IDOggetto
    rs.Update
Else
    rs!DataEmissione = DataScadenzaRata
    rs!Numero = NumeroRata
    rs!DataUltimaVariazione = Date
    rs!IDUtenteUltimaVariazione = TheApp.IDUser
    rs!VirtualDelete = 0
    GET_LINK_OGGETTO = rs!IDOggetto
    rs.Update
End If

rs.Close
Set rs = Nothing

Exit Function

ERR_GET_LINK_OGGETTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_OGGETTO"
    GET_LINK_OGGETTO = 0
End Function
Private Function GET_LINK_FUNZIONE(IDTipoOggetto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDFunzione FROM Funzione "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_FUNZIONE = 0
Else
    GET_LINK_FUNZIONE = fnNotNullN(rs!IDFunzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_FUNZIONE(IDFunzione As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Funzione FROM Funzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_FUNZIONE = ""
Else
    GET_DESCRIZIONE_FUNZIONE = fnNotNull(rs!Funzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_ATTIVITA_AZIENDA(IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAttivitaAzienda FROM Filiale "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ATTIVITA_AZIENDA = 0
Else
    GET_LINK_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function fnGetTipoOggetto(NomeGestore) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(NomeGestore)
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Function GET_LINK_ADEGUAMENTO_DA_RATA(IDRataContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POContrattoAdeguamento FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_PORateContratto=" & IDRataContratto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ADEGUAMENTO_DA_RATA = 0
Else
    GET_LINK_ADEGUAMENTO_DA_RATA = fnNotNullN(rs!IDRV_POContrattoAdeguamento)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub RINNOVO_SERVIZI_PRODOTTI(IDContratto As Long, IDContrattoOLD As Long)
On Error GoTo ERR_RINNOVO_SERVIZI_PRODOTTI
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim rsOld As DmtOleDbLib.adoResultset

Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POContrattoServiziProdotti "
sSQL = sSQL & "WHERE IDRV_POContrattoServiziProdotti=0"
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic


sSQL = "SELECT * FROM RV_POIERinnovoServiziProdotti "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Eliminato=" & fnNormBoolean(0)

Set rsOld = CnDMT.OpenResultset(sSQL)


While Not rsOld.EOF
    rsNew.AddNew
        rsNew!IDRV_POContrattoServizi = fnNotNullN(rsOld!IDRV_POContrattoServiziNew)
        rsNew!IDRV_POContrattoProdotti = fnNotNullN(rsOld!IDRV_POContrattoProdottiNew)
        rsNew!Eliminato = False
    rsNew.Update
rsOld.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rsOld.CloseResultset
Set rsOld = Nothing
Exit Sub
ERR_RINNOVO_SERVIZI_PRODOTTI:
    MsgBox Err.Description, vbCritical, "RINNOVO_SERVIZI_PRODOTTI"
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

rsGriglia3.Fields.Append "IDAnagraficaClienteFatt", adInteger, , adFldIsNullable

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
Private Sub CREA_RECORDSET_ERRORI()
Dim sSQL As String
Dim rs As ADODB.Recordset

Set rsGrigliaErrori = Nothing

Set rsGrigliaErrori = New ADODB.Recordset

rsGrigliaErrori.CursorLocation = adUseClient


rsGrigliaErrori.Fields.Append "IDContratto", adInteger, , adFldIsNullable
rsGrigliaErrori.Fields.Append "NumeroContratto", adInteger, , adFldIsNullable
rsGrigliaErrori.Fields.Append "AnnoContratto", adInteger, , adFldIsNullable
rsGrigliaErrori.Fields.Append "NumeroRinnovo", adInteger, , adFldIsNullable
rsGrigliaErrori.Fields.Append "IDAnagraficaCliente", adInteger, , adFldIsNullable
rsGrigliaErrori.Fields.Append "DescrizioneAnagraficaCliente", adVarChar, 250, adFldIsNullable
rsGrigliaErrori.Fields.Append "FunzioneErrore", adVarChar, 250, adFldIsNullable
rsGrigliaErrori.Fields.Append "Errore", adVarChar, 2500, adFldIsNullable

rsGrigliaErrori.Open , , adOpenKeyset, adLockPessimistic
End Sub

Private Sub GET_DATI_RECORDSET_3(IDContrattoServizi As Long, IDContratto As Long, IDContrattoPadre As Long, IDCliente As Long, IDTecnicoRiferimento As Long, IDTecnicoOperativo As Long, rsContratto As ADODB.Recordset)
On Error GoTo ERR_GET_DATI_RECORDSET_3
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsServizio As ADODB.Recordset
Dim rsProdotti As ADODB.Recordset
Dim rsFestivita As ADODB.Recordset

'''''''''''''''DICHIARAZIONE DELLE VARIABILI''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim DataInizioServizio As String
Dim DataInizioPersonalizzata As String
Dim DataFineServizio As String
Dim DataFinePersonalizzata As String
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

'NUMERO_INTERVENTI_DA_CREARE = 0


LINK_STATO_LOCAL = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDRV_POStatoInterventoInserimento")
LINK_TIPO_LOCAL = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDRV_POTipoFaseInterventoEla")

LINK_RIFERIMENTO_INT_LOCAL = fnNotNullN(rsContratto!IDAnagraficaCommesso)
LINK_TECNICO_OPERATIVO_LOCAL = LINK_RIFERIMENTO_INT_LOCAL

LINK_TIPO_ADDEBITO_LOCAL = GET_PARAMETRI_TEC_OPE(LINK_TECNICO_OPERATIVO_LOCAL, "IDRV_POTipoAddebito")
LINK_CLASSE_LOCAL = GET_PARAMETRI_TEC_OPE(LINK_TECNICO_OPERATIVO_LOCAL, "IDRV_POTipoClasseIntervento")
LINK_CATEGORIA_LOCAL = GET_PARAMETRI_TEC_OPE(LINK_TECNICO_OPERATIVO_LOCAL, "IDRV_POCategoriaFase")


'CONNESSIONE AL SERVIZIO
sSQL = "SELECT RV_POContrattoServizi.IDRV_POContrattoServizi, RV_POContrattoServizi.IDRV_POContratto, RV_POContrattoServizi.IDRV_POStoriaContratto, RV_POContrattoServizi.IDArticolo, "
sSQL = sSQL & "RV_POContrattoServizi.IDRV_POCriterioRicorrenza, RV_POContrattoServizi.OgniNumeroGiorni, RV_POContrattoServizi.OgniNumeroMesi, RV_POContrattoServizi.OgniNumeroSettimane, "
sSQL = sSQL & "RV_POContrattoServizi.IDRV_POTipoDataInizioRicorrenza, RV_POContrattoServizi.GiornoInizioRicorrenza, RV_POContrattoServizi.MeseInizioRicorrenza, "
sSQL = sSQL & "RV_POContrattoServizi.IDRV_POTipoDataFineRicorrenza, RV_POContrattoServizi.GiornoFineRicorrenza, RV_POContrattoServizi.MeseFineRicorrenza, RV_POContrattoServizi.NumeroRicorrenze, "
sSQL = sSQL & "RV_POContrattoServizi.IDRV_POContrattoPadre , Articolo.CodiceArticolo, Articolo.Articolo, RV_POContrattoServizi.IDRV_POTipoAnnoInizioRicorrenza, RV_POContrattoServizi.IDRV_POTipoAnnoFineRicorrenza "
sSQL = sSQL & "FROM RV_POContrattoServizi INNER JOIN "
sSQL = sSQL & "Articolo ON RV_POContrattoServizi.IDArticolo = Articolo.IDArticolo "
sSQL = sSQL & "WHERE RV_POContrattoServizi.IDRV_POContrattoServizi=" & IDContrattoServizi

Set rsServizio = New ADODB.Recordset

rsServizio.Open sSQL, CnDMT.InternalConnection

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
sSQL = sSQL & "WHERE RV_POContrattoServiziProdotti.IDRV_POContrattoServizi=" & IDContrattoServizi
sSQL = sSQL & " AND RV_POContrattoServiziProdotti.Eliminato=" & fnNormBoolean(0)
sSQL = sSQL & " AND RV_POContrattoProdotti.Dismesso =" & fnNormBoolean(0)

Set rsProdotti = New ADODB.Recordset
rsProdotti.Open sSQL, CnDMT.InternalConnection

If rsProdotti.EOF Then
    AvviaConProdotti = False
Else
    AvviaConProdotti = True
End If

'CALCOLO DELLA DATA INIZIO RICORRENZA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (fnNotNullN(rsContratto!IDRV_POTipoImpostazioneContratto) <> 3) Then
    DataInizioServizio = rsContratto!DataDecorrenza
Else
    If (fnNotNullN(rsServizio!IDRV_POTipoDataInizioRicorrenza) = 0) And (fnNotNullN(rsServizio!IDRV_POTipoDataFineRicorrenza) > 0) Then
        If AvviaConProdotti = True Then
            DataInizioServizio = rsContratto!DataScadenzaPerRinnovo
        Else
            DataInizioServizio = fnNotNull(rsProdotti!DataFinePeriodo)
        End If

    Else
        If AvviaConProdotti = True Then
            DataInizioServizio = fnNotNull(rsProdotti!DataInizioPeriodo)
        Else
            DataInizioServizio = rsContratto!DataDecorrenza
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
        'If (frmMain.cboTipoImpostazione.CurrentID <> 3) Then
            'DataInizioServizio = DateAdd("d", -1, DataInizioServizio)
        'End If
    Case 2
        'DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), frmMain.txtDataDecorrenza.Text)
        DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), DataInizioServizio)
        DataInizioServizio = DateAdd("d", IIf((fnNotNullN(rsServizio!OgniNumeroGiorni) = 0), 0, fnNotNullN(rsServizio!OgniNumeroGiorni)), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((fnNotNullN(rsServizio!OgniNumeroSettimane) = 0), 0, fnNotNullN(rsServizio!OgniNumeroSettimane)), DataInizioServizio)
    Case 3
        DataInizioPersonalizzata = GET_COSTRUZIONE_DATA_PERS(fnNotNullN(rsServizio!GiornoInizioRicorrenza), fnNotNullN(rsServizio!MeseInizioRicorrenza)) '& Year(rsContratto!DataDecorrenza)
        If (fnNotNullN(rsServizio!IDRV_POTipoAnnoInizioRicorrenza)) = 1 Then DataInizioPersonalizzata = DataInizioPersonalizzata & Year(rsContratto!DataDecorrenza)
        If (fnNotNullN(rsServizio!IDRV_POTipoAnnoInizioRicorrenza)) = 2 Then DataInizioPersonalizzata = DataInizioPersonalizzata & Year(rsContratto!DataScadenzaPerRinnovo)
        If (fnNotNullN(rsServizio!IDRV_POTipoAnnoInizioRicorrenza)) = 0 Then DataInizioPersonalizzata = DataInizioPersonalizzata
        
        If (IsDate(DataInizioPersonalizzata)) Then DataInizioServizio = DataInizioPersonalizzata

'        If (fnNotNullN(rsServizio!GiornoInizioRicorrenza) = fnNotNullN(rsServizio!GiornoFineRicorrenza)) And (fnNotNullN(rsServizio!MeseInizioRicorrenza) = fnNotNullN(rsServizio!MeseFineRicorrenza)) Then
'            DataInizioPersonalizzata = GET_COSTRUZIONE_DATA_PERS(fnNotNullN(rsServizio!GiornoInizioRicorrenza), fnNotNullN(rsServizio!MeseInizioRicorrenza)) & Year(rsContratto!DataFineAssistenza)
'        End If

        'DataInizioServizio = DataInizioPersonalizzata
'        DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), DataInizioPersonalizzata)
'        DataInizioServizio = DateAdd("d", IIf((fnNotNullN(rsServizio!OgniNumeroGiorni) = 0), 0, fnNotNullN(rsServizio!OgniNumeroGiorni)), DataInizioServizio)
'        DataInizioServizio = DateAdd("ww", IIf((fnNotNullN(rsServizio!OgniNumeroSettimane) = 0), 0, fnNotNullN(rsServizio!OgniNumeroSettimane)), DataInizioServizio)
    Case Else
        'DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), frmMain.txtDataDecorrenza.Text)
        DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), DataInizioServizio)
        DataInizioServizio = DateAdd("d", IIf((fnNotNullN(rsServizio!OgniNumeroGiorni) = 0), 0, fnNotNullN(rsServizio!OgniNumeroGiorni)), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((fnNotNullN(rsServizio!OgniNumeroSettimane) = 0), 0, fnNotNullN(rsServizio!OgniNumeroSettimane)), DataInizioServizio)
        'If (frmMain.cboTipoImpostazione.CurrentID <> 3) Then
        '    DataInizioServizio = DateAdd("d", -1, DataInizioServizio)
        'End If
        DataInizioServizio = ""
End Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'CALCOLO DELLA DATA DI FINE RICORRENZA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Select Case fnNotNullN(rsServizio!IDRV_POTipoDataFineRicorrenza)
    Case 1
        If (fnNotNullN(rsContratto!IDRV_POTipoImpostazioneContratto) <> 3) Then
            DataFineServizio = rsContratto!DataScadenza
        Else
            If AvviaConProdotti = True Then
                DataFineServizio = fnNotNull(rsProdotti!DataFinePeriodo)
            Else
                DataFineServizio = rsContratto!DataScadenza
            End If
        End If
    Case 2
        If (fnNotNullN(rsContratto!IDRV_POTipoImpostazioneContratto) <> 3) Then
            DataFineServizio = rsContratto!DataScadenzaPerRinnovo
        Else
            If AvviaConProdotti = True Then
                DataFineServizio = fnNotNull(rsProdotti!DataFinePeriodo)
            Else
                DataFineServizio = rsContratto!DataScadenzaPerRinnovo
            End If
        End If
    Case 3
        If (fnNotNullN(rsContratto!IDRV_POTipoImpostazioneContratto) <> 3) Then
            DataFineServizio = rsContratto!DataFineAssistenza
        Else
            If AvviaConProdotti = True Then
                DataFineServizio = fnNotNull(rsProdotti!DataFinePeriodo)
            Else
                DataFineServizio = rsContratto!DataFineAssistenza
            End If
        End If
    Case 4
        If (fnNotNullN(rsContratto!IDRV_POTipoImpostazioneContratto) <> 3) Then
            DataFinePersonalizzata = GET_COSTRUZIONE_DATA_PERS(fnNotNullN(rsServizio!GiornoFineRicorrenza), fnNotNullN(rsServizio!MeseFineRicorrenza)) '& Year(frmMain.txtDataFineAssistenza.Text)
            If (fnNotNullN(rsServizio!IDRV_POTipoAnnoFineRicorrenza)) = 1 Then DataFinePersonalizzata = DataFinePersonalizzata & Year(rsContratto!DataDecorrenza)
            If (fnNotNullN(rsServizio!IDRV_POTipoAnnoFineRicorrenza)) = 2 Then DataFinePersonalizzata = DataFinePersonalizzata & Year(rsContratto!DataScadenzaPerRinnovo)
            
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
                If (fnNotNullN(rsServizio!IDRV_POTipoAnnoFineRicorrenza)) = 1 Then DataFinePersonalizzata = DataFinePersonalizzata & Year(rsContratto!DataDecorrenza)
                If (fnNotNullN(rsServizio!IDRV_POTipoAnnoFineRicorrenza)) = 2 Then DataFinePersonalizzata = DataFinePersonalizzata & Year(rsContratto!DataScadenzaPerRinnovo)
                
                If (IsDate(DataFinePersonalizzata)) Then
                    DataFineServizio = DataFinePersonalizzata
                Else
                    DataFineServizio = DataInizioServizio
                End If
            End If
        End If
    Case Else
        DataFineServizio = DataFineServizio
End Select

If (fnNotNullN(rsServizio!IDRV_POTipoDataInizioRicorrenza) = 0) Then DataInizioServizio = DataFineServizio

If (DataFineServizio = "") Then DataFineServizio = DataInizioServizio
If (DataInizioServizio = "") Then DataInizioServizio = DataFineServizio

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
                        
                        'If GET_DATA_FESTIVO(DataInizioServizio, fnNotNull(rsProdotti!EscludiGiorniFestivi), fnNotNullN(rsProdotti!EscludiSabato), rsFestivita) = False Then
                        rsGriglia3.AddNew
                            
                            rsGriglia3!IDAnagraficaCliente = rsContratto!IDAnagrafica
                            rsGriglia3!IDAnagraficaClienteFatt = rsContratto!IDAnagraficaFatturazione
                            
                            rsGriglia3!NumeroInterventoSub = NumeroInterventoSub
                            rsGriglia3!FaseIntervento = 1
                            rsGriglia3!IDRV_POContratto = rsContratto!IDRV_POContratto
                            rsGriglia3!IDRV_POContrattoPadre = rsContratto!IDRV_POContrattoPadre
                            
                            rsGriglia3!IDAnagraficaTecnicoRif = LINK_RIFERIMENTO_INT_LOCAL
                            rsGriglia3!IDAnagraficaTecnicoOperativo = LINK_TECNICO_OPERATIVO_LOCAL
                            rsGriglia3!IDRV_POCategoriaIntervento = LINK_CATEGORIA_LOCAL
                            rsGriglia3!IDRV_POTipoAddebito = LINK_TIPO_ADDEBITO_LOCAL
                            rsGriglia3!IDRV_POTipoClasseIntervento = LINK_CLASSE_LOCAL
                            rsGriglia3!IDRV_POStatoIntervento = LINK_STATO_LOCAL
                            rsGriglia3!IDRV_POTipoFaseIntervento = LINK_TIPO_LOCAL
                            
                            
                            rsGriglia3!IDArticolo = fnNotNullN(rsServizio!IDArticolo)
                            rsGriglia3!CodiceServizio = fnNotNull(rsServizio!CodiceArticolo)
                            rsGriglia3!DescrizioneServizio = fnNotNull(rsServizio!Articolo)
                            
                            rsGriglia3!IDRV_POProdotto = fnNotNullN(rsProdotti!IDRV_POProdotto)
                            rsGriglia3!DescrizioneProdotto = fnNotNull(rsProdotti!Descrizione)
                            rsGriglia3!Matricola = fnNotNull(rsProdotti!Matricola)
                            
                            rsGriglia3!IDRV_POContrattoServizi = IDContrattoServizi
                            rsGriglia3!IDRV_POContrattoProdotti = fnNotNullN(rsProdotti!IDRV_POContrattoProdotti)
                            'rsGriglia3!IDRV_POProdottoServizi = fnNotNullN(rsProdotti!IDRV_POContrattoServiziProdotti)

                            rsGriglia3!DataAppuntamento = DataInizioServizio
                            rsGriglia3!OraAppuntamento = "09:00"

                            
                            rsGriglia3!Registra = True
                    
                        rsGriglia3.Update
                        NumeroInterventoSub = NumeroInterventoSub + 1
                        NUMERO_INTERVENTI_DA_CREARE = NUMERO_INTERVENTI_DA_CREARE + 1
                        'End If
                        
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
                rsGriglia3.AddNew
                
                    rsGriglia3!IDAnagraficaCliente = rsContratto!IDAnagrafica
                    rsGriglia3!IDAnagraficaClienteFatt = rsContratto!IDAnagraficaFatturazione
                
                    rsGriglia3!IDAnagraficaTecnicoRif = LINK_RIFERIMENTO_INT_LOCAL
                    rsGriglia3!IDAnagraficaTecnicoOperativo = LINK_TECNICO_OPERATIVO_LOCAL
                    rsGriglia3!IDRV_POCategoriaIntervento = LINK_CATEGORIA_LOCAL
                    rsGriglia3!IDRV_POTipoAddebito = LINK_TIPO_ADDEBITO_LOCAL
                    rsGriglia3!IDRV_POTipoClasseIntervento = LINK_CLASSE_LOCAL
                    rsGriglia3!IDRV_POStatoIntervento = LINK_STATO_LOCAL
                    rsGriglia3!IDRV_POTipoFaseIntervento = LINK_TIPO_LOCAL
                
                    rsGriglia3!NumeroInterventoSub = NumeroInterventoSub
                    rsGriglia3!FaseIntervento = 1
                    rsGriglia3!IDRV_POContratto = rsContratto!IDRV_POContratto
                    rsGriglia3!IDRV_POContrattoPadre = rsContratto!IDRV_POContrattoPadre
                    
                    rsGriglia3!IDArticolo = fnNotNullN(rsServizio!IDArticolo)
                    rsGriglia3!CodiceServizio = fnNotNull(rsServizio!CodiceArticolo)
                    rsGriglia3!DescrizioneServizio = fnNotNull(rsServizio!Articolo)
                    
                  
                    rsGriglia3!IDRV_POContrattoServizi = IDContrattoServizi
                    
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
                            'If GET_DATA_FESTIVO(DataInizioServizio, fnNotNull(rsProdotti!EscludiGiorniFestivi), fnNotNullN(rsProdotti!EscludiSabato), rsFestivita) = False Then

                                rsGriglia3.AddNew
                                
                                    rsGriglia3!IDAnagraficaCliente = rsContratto!IDAnagrafica
                                    rsGriglia3!IDAnagraficaClienteFatt = rsContratto!IDAnagraficaFatturazione
                                
                                    rsGriglia3!IDAnagraficaTecnicoRif = LINK_RIFERIMENTO_INT_LOCAL
                                    rsGriglia3!IDAnagraficaTecnicoOperativo = LINK_TECNICO_OPERATIVO_LOCAL
                                    rsGriglia3!IDRV_POCategoriaIntervento = LINK_CATEGORIA_LOCAL
                                    rsGriglia3!IDRV_POTipoAddebito = LINK_TIPO_ADDEBITO_LOCAL
                                    rsGriglia3!IDRV_POTipoClasseIntervento = LINK_CLASSE_LOCAL
                                    rsGriglia3!IDRV_POStatoIntervento = LINK_STATO_LOCAL
                                    rsGriglia3!IDRV_POTipoFaseIntervento = LINK_TIPO_LOCAL
                                    
                                    rsGriglia3!NumeroInterventoSub = NumeroInterventoSub
                                    rsGriglia3!FaseIntervento = 1
                                    rsGriglia3!IDRV_POContratto = rsContratto!IDRV_POContratto
                                    rsGriglia3!IDRV_POContrattoPadre = rsContratto!IDRV_POContrattoPadre
                                    
                                    rsGriglia3!IDArticolo = fnNotNullN(rsServizio!IDArticolo)
                                    rsGriglia3!CodiceServizio = fnNotNull(rsServizio!CodiceArticolo)
                                    rsGriglia3!DescrizioneServizio = fnNotNull(rsServizio!Articolo)
                                    
                                    rsGriglia3!IDRV_POProdotto = fnNotNullN(rsProdotti!IDRV_POProdotto)
                                    rsGriglia3!DescrizioneProdotto = fnNotNull(rsProdotti!Descrizione)
                                    rsGriglia3!Matricola = fnNotNull(rsProdotti!Matricola)
                                    
                                    
                                    rsGriglia3!IDRV_POContrattoServizi = IDContrattoServizi
                                    rsGriglia3!IDRV_POContrattoProdotti = fnNotNullN(rsProdotti!IDRV_POContrattoProdotti)
                                    
                                    rsGriglia3!DataAppuntamento = DataInizioServizio
                                    rsGriglia3!OraAppuntamento = "09:00"
                                                                    
                                    rsGriglia3!Registra = True
                            
                                rsGriglia3.Update
                                NumeroInterventoSub = NumeroInterventoSub + 1
                                NUMERO_INTERVENTI_DA_CREARE = NUMERO_INTERVENTI_DA_CREARE + 1
                            'End If
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
                    rsGriglia3.AddNew
                        rsGriglia3!NumeroInterventoSub = NumeroInterventoSub
                        rsGriglia3!FaseIntervento = 1
                        rsGriglia3!IDRV_POContratto = rsContratto!IDRV_POContratto
                        rsGriglia3!IDRV_POContrattoPadre = rsContratto!IDRV_POContrattoPadre
                        
                        rsGriglia3!IDArticolo = fnNotNullN(rsServizio!IDArticolo)
                        rsGriglia3!CodiceServizio = fnNotNull(rsServizio!CodiceArticolo)
                        rsGriglia3!DescrizioneServizio = fnNotNull(rsServizio!Articolo)
                        
                       
                        rsGriglia3!IDAnagraficaCliente = rsContratto!IDAnagrafica
                        rsGriglia3!IDAnagraficaClienteFatt = rsContratto!IDAnagraficaFatturazione
                        
                        rsGriglia3!IDAnagraficaTecnicoRif = LINK_RIFERIMENTO_INT_LOCAL
                        rsGriglia3!IDAnagraficaTecnicoOperativo = LINK_TECNICO_OPERATIVO_LOCAL
                        rsGriglia3!IDRV_POCategoriaIntervento = LINK_CATEGORIA_LOCAL
                        rsGriglia3!IDRV_POTipoAddebito = LINK_TIPO_ADDEBITO_LOCAL
                        rsGriglia3!IDRV_POTipoClasseIntervento = LINK_CLASSE_LOCAL
                        rsGriglia3!IDRV_POStatoIntervento = LINK_STATO_LOCAL
                        rsGriglia3!IDRV_POTipoFaseIntervento = LINK_TIPO_LOCAL
                        'rsGriglia3!AnagraficaCliente = frmMain.CDCliente.Code & " " & frmMain.CDCliente.Description
                        
                        rsGriglia3!IDRV_POContrattoServizi = IDContrattoServizi
                        
                        rsGriglia3!IDAnagraficaTecnicoRif = IDTecnicoRiferimento
                        'rsGriglia3!AnagraficaTecnicoRif = frmMain.CDTecnico.Code & " " & frmMain.CDTecnico.Description
                        
                        rsGriglia3!IDAnagraficaTecnicoOperativo = IDTecnicoOperativo
                        'rsGriglia3!AnagraficaTecnicoOperativo = GET_ANAGRAFICA(rsGriglia3!IDAnagraficaTecnicoOperativo)
                        
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
            'End If
    
            DataInizioServizio = DateAdd("m", IIf((fnNotNullN(rsServizio!OgniNumeroMesi) = 0), 0, fnNotNullN(rsServizio!OgniNumeroMesi)), DataInizioServizio)
            DataInizioServizio = DateAdd("d", IIf((fnNotNullN(rsServizio!OgniNumeroGiorni) = 0), 0, fnNotNullN(rsServizio!OgniNumeroGiorni)), DataInizioServizio)
            DataInizioServizio = DateAdd("ww", IIf((fnNotNullN(rsServizio!OgniNumeroSettimane) = 0), 0, fnNotNullN(rsServizio!OgniNumeroSettimane)), DataInizioServizio)
            'DataInizioServizio = DateAdd("d", -1, DataInizioServizio)
            NumeroInterventoSub = 1
        Wend
    End If
End If



rsServizio.Close
Set rsServizio = Nothing

rsProdotti.Close
Set rsProdotti = Nothing

Exit Sub
ERR_GET_DATI_RECORDSET_3:
    MsgBox Err.Description, vbCritical, "GET_DATI_RECORDSET_3"
End Sub

Private Sub CREA_TMP_INTERVENTI(rsContratto As ADODB.Recordset)
On Error GoTo ERR_CREA_TMP_INTERVENTI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POContrattoServizi "
sSQL = sSQL & "WHERE IDRV_POContratto=" & fnNotNullN(rsContratto!IDRV_POContratto)

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    GET_DATI_RECORDSET_3 fnNotNullN(rs!IDRV_POContrattoServizi), fnNotNullN(rsContratto!IDRV_POContratto), fnNotNullN(rsContratto!IDRV_POContrattoPadre), fnNotNullN(rsContratto!IDAnagrafica), fnNotNullN(rsContratto!IDAnagraficaCommesso), fnNotNullN(rsContratto!IDAnagraficaCommesso), rsContratto
rs.MoveNext
Wend

rs.CloseResultset


Exit Sub
ERR_CREA_TMP_INTERVENTI:
    MsgBox Err.Description, vbCritical, "CREA_TMP_INTERVENTI"
End Sub
Private Sub CREA_INTERVENTI()
On Error GoTo ERR_CREA_INTERVENTI
Dim sSQL As String
Dim UnitaProgresso As Double
Dim rsIntervento As ADODB.Recordset
Dim link_riga_servizio As Long
Dim link_cliente As Long

Dim NumeroIntervento As Long
 
Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100

If NUMERO_INTERVENTI_DA_CREARE = 0 Then Exit Sub

UnitaProgresso = FormatNumber((Me.ProgressBar1.Max / NUMERO_INTERVENTI_DA_CREARE), 4)

Me.lblInfo.Caption = "CREAZIONE INTERVENTI IN CORSO..."
Screen.MousePointer = 11
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

rsGriglia3.Filter = "Registra=" & fnNormBoolean(True)

If ((rsGriglia3.EOF) And (rsGriglia3.BOF)) Then Exit Sub

rsGriglia3.Sort = "IDRV_POContrattoServizi"

sSQL = "SELECT * FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=0"

Set rsIntervento = New ADODB.Recordset

rsIntervento.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
link_cliente = 0

link_riga_servizio = 0
NumeroIntervento = 0

While Not rsGriglia3.EOF
    Screen.MousePointer = 11
    
    If (rsGriglia3!NumeroInterventoSub = 1) Then
        'NumeroIntervento = GET_NUMERO_INTERVENTO(Year(rsGriglia3!DataAppuntamento))
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
    Screen.MousePointer = 0
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
        CnDMT.Execute sSQL
    rsInterventoUpdate.MoveNext
    Wend
End If

Exit Sub
ERR_CREA_INTERVENTI:
    MsgBox Err.Description, vbCritical, "CREA_INTERVENTI"
    Screen.MousePointer = 0
End Sub

Private Sub CREA_INTERVENTO(rsNew As ADODB.Recordset, NumeroIntervento As Long)
On Error GoTo ERR_CREA_INTERVENTO

rsNew.AddNew

    'rsnew!IDRV_POIntervento = fnGetNewKey("RV_POIntervento", "IDRV_POIntervento")
    'rsNew!IDRV_POInterventoPadre = rsNew!IDRV_POIntervento
    rsNew!IDRV_POContrattoServizi = rsGriglia3!IDRV_POContrattoServizi
    rsNew!IDRV_POContratto = rsGriglia3!IDRV_POContratto
    rsNew!IDRV_POContrattoPadre = rsGriglia3!IDRV_POContrattoPadre
    
    rsNew!Elaborata = 1
    rsNew!Manuale = 0
    
    rsNew!IDAnagraficaCliente = rsGriglia3!IDAnagraficaCliente
    rsNew!IDAnagraficaFatturazione = rsGriglia3!IDAnagraficaClienteFatt
    
    rsNew!IDAzienda = TheApp.IDFirm
    rsNew!IDFiliale = TheApp.Branch
    rsNew!AnnoIntervento = Year(rsGriglia3!DataAppuntamento)
    rsNew!NumeroIntervento = NumeroIntervento 'GET_NUMERO_INTERVENTO(Year(Date))
    rsNew!NumeroInterventoSub = rsGriglia3!NumeroInterventoSub
    rsNew!NumeroFase = 1
    rsNew!InterventoChiuso = 0
    
    rsNew!IDAnagraficaTecnicoRif = rsGriglia3!IDAnagraficaTecnicoRif
    rsNew!IDTipoAnagraficaTecnicoRif = LINK_TIPO_ANA_TEC_RIF
    
    rsNew!IDAnagraficaTecnicoOperativo = rsGriglia3!IDAnagraficaTecnicoOperativo
    rsNew!IDTipoAnagraficaTecnicoOpe = LINK_TIPO_ANA_TEC_OPE
    
    Me.txtRichiesta.Text = fnNotNull(rsGriglia3!DescrizioneServizio)
    
    rsNew!Richiesta = Me.txtRichiesta.TextRTF
    
    'rsNew!Richiesta = fnNotNull(rsGriglia3!DescrizioneServizio)
    
    rsNew!DataAppuntamento = rsGriglia3!DataAppuntamento
    rsNew!OraAppuntamento = rsGriglia3!OraAppuntamento
    rsNew!LavoroEseguito = ""
    rsNew!Annotazioni = ""
    
    rsNew!IDRV_POStagione = GET_LINK_STAGIONE(rsGriglia3!DataAppuntamento)
    
    rsNew!IDRV_POCategoriaIntervento = rsGriglia3!IDRV_POCategoriaIntervento
    
    rsNew!IDRV_POTipoAddebito = rsGriglia3!IDRV_POTipoAddebito
    
    rsNew!IDRV_POTipoClasseIntervento = rsGriglia3!IDRV_POTipoClasseIntervento
    
    rsNew!IDRV_POStatoIntervento = rsGriglia3!IDRV_POStatoIntervento
    
    rsNew!IDRV_POTipoFaseIntervento = rsGriglia3!IDRV_POTipoFaseIntervento

    rsNew!IDRV_POProdotto = fnNotNullN(rsGriglia3!IDRV_POProdotto)
    rsNew!IDArticolo = fnNotNullN(rsGriglia3!IDArticolo)
    rsNew!IDRV_POContrattoProdotti = fnNotNullN(rsGriglia3!IDRV_POContrattoProdotti)
    
    rsNew!DataInserimento = Date
    rsNew!OraInserimento = GET_ORARIO(Now)
    rsNew!IDUtenteInserimento = TheApp.IDUser
    rsNew!DataUltimaModifica = Date
    rsNew!OraUltimaModifica = GET_ORARIO(Now)
    rsNew!IDUtenteUltimaModifica = TheApp.IDUser
    rsNew!NomeComputerInserimento = GET_NOMECOMPUTER
    rsNew!UtenteComputerInserimento = GET_NOMEUTENTE
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

Exit Sub

ERR_CREA_INTERVENTO:
    MsgBox Err.Description, vbCritical, "CREA_INTERVENTO"
    On Error Resume Next
    rsNew.CancelUpdate
End Sub


Private Sub GET_TOTALE_RIGA(rstmpProd As ADODB.Recordset, percMagg As Double)
Dim Imponibile As Double
Dim importoTotale As Double
Dim ImportoIva As Double

Imponibile = FormatNumber((fnNotNullN(rstmpProd!ImportoUnitario) * ((percMagg / 100) + 1)), 5)
Imponibile = Imponibile - ((Imponibile / 100) * fnNotNullN(rstmpProd!Sconto1))
Imponibile = Imponibile - ((Imponibile / 100) * fnNotNullN(rstmpProd!Sconto2))
Imponibile = Imponibile * fnNotNullN(rstmpProd!QuantitaArticolo)

Imponibile = Imponibile * fnNotNullN(rstmpProd!QuantitaPeriodo)
Imponibile = Imponibile - fnNotNullN(rstmpProd!ScontoAImporto)

Imponibile = FormatNumber(Imponibile, 2)

importoTotale = Imponibile * ((fnNotNullN(rstmpProd!AliquotaIva) / 100) + 1)

ImportoIva = importoTotale - Imponibile

rstmpProd!ImportoUnitario = FormatNumber((fnNotNullN(rstmpProd!ImportoUnitario) * ((percMagg / 100) + 1)), 5)
rstmpProd!Imponibile = Imponibile
rstmpProd!ImportoIva = ImportoIva
rstmpProd!TotaleRiga = importoTotale
rstmpProd!ImportoComplessivo = Imponibile

End Sub
Private Sub INSERIMENTO_PRODOTTI_CONF_CONT(IDContrattoProdottoOLD As Long, IDContrattoProdottoNew As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsNew As ADODB.Recordset

sSQL = "SELECT * FROM RV_POContrattoProdottiContatori "
sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDContrattoProdottoNew
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

sSQL = "SELECT * FROM RV_POContrattoProdottiContatori "
sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDContrattoProdottoOLD

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_POContrattoProdottiContatori = fnGetNewKey("RV_POContrattoProdottiContatori", "IDRV_POContrattoProdottiContatori")
        rsNew!IDRV_POContrattoProdotti = IDContrattoProdottoNew
        rsNew!IDRV_POProdotto = fnNotNullN(rs!IDRV_POProdotto)
        rsNew!IDRV_POContatoreProdotto = fnNotNullN(rs!IDRV_POContatoreProdotto)
        rsNew!QuantitaMax = fnNotNullN(rs!QuantitaMax)
        rsNew!IDRV_POUnitaDiMisuraPeriodo = fnNotNullN(rs!IDRV_POUnitaDiMisuraPeriodo)
        rsNew!QuantitaPeriodo = fnNotNullN(rs!QuantitaPeriodo)
        rsNew!QuantitaInizio = GET_QUANTITA_INIZIO_CONTATORE(fnNotNullN(rs!IDRV_POContrattoProdottiContatori), fnNotNullN(rs!QuantitaInizio))
        rsNew!ImportoUnitario = fnNotNullN(rs!ImportoUnitario)
    rsNew.Update
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rs.Close
Set rs = Nothing
End Sub
Private Function GET_QUANTITA_INIZIO_CONTATORE(IDContrattoContProd As Long, QuantitaInizio As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim QtaCont As Double

GET_QUANTITA_INIZIO_CONTATORE = QuantitaInizio
QtaCont = 0

sSQL = "SELECT MAX(Quantita) AS MaxQuantita "
sSQL = sSQL & "FROM RV_POContatoreRilevamenti "
sSQL = sSQL & "WHERE IDRV_POContrattoProdottiContatori=" & IDContrattoContProd

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    QtaCont = fnNotNullN(rs!MaxQuantita)
End If

rs.CloseResultset
Set rs = Nothing

If QtaCont > 0 Then
    GET_QUANTITA_INIZIO_CONTATORE = QtaCont
End If

End Function
Private Function GET_TOTALE_CONTRATTO(IDContratto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT SUM(ImportoComplessivo) AS TotaleRiga "
sSQL = sSQL & "FROM RV_POContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_CONTRATTO = 0
Else
    GET_TOTALE_CONTRATTO = fnNotNullN(rs!TotaleRiga)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_PERCENTUALE_ISTAT(IDIstat As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_PERCENTUALE_ISTAT = 0

sSQL = "SELECT IDRV_POIstat, Percentuale "
sSQL = sSQL & "FROM RV_POIstat "
sSQL = sSQL & "WHERE IDRV_POIstat=" & IDIstat

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PERCENTUALE_ISTAT = 0
Else
    GET_PERCENTUALE_ISTAT = fnNotNullN(rs!Percentuale)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ANAGRAFICA_CONTRATTO(IDContratto As Long) As String
On Error GoTo ERR_GET_ANAGRAFICA_CONTRATTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POContratto, Anagrafica, TipoContratto, AnnoContratto, NumeroContratto, NumeroRinnovo "
sSQL = sSQL & "FROM RV_POViewContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ANAGRAFICA_CONTRATTO = ""
Else
    GET_ANAGRAFICA_CONTRATTO = fnNotNull(rs!Anagrafica) & " - " & fnNotNull(rs!TipoContratto)
    GET_ANAGRAFICA_CONTRATTO = GET_ANAGRAFICA_CONTRATTO & " - numero " & fnNotNullN(rs!AnnoContratto) & "-" & fnNotNullN(rs!NumeroContratto) & "-" & fnNotNullN(rs!NumeroRinnovo)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_ANAGRAFICA_CONTRATTO:
    
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
Dim Mese As Double
   
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
      Mese = 3
   Else
      Giorno = d + e - 9
      Mese = 4
   End If
 
   If (Giorno = 26 And Mese = 4) Then
      Giorno = 19
      Mese = 4
   End If
 
   If (Giorno = 25 And Mese = 4 And d = 28 And e = 6 And a > 10) Then
      Giorno = 18
      Mese = 4
   End If
 
   CalcolaPasqua = DateSerial(Anno, Mese, Giorno)
End Function

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

Private Function GET_TIPO_ANNO_INIZIO_SERVIZIO(IDArticolo As Long) As Long
On Error GoTo ERR_GET_TIPO_ANNO_SERVIZIO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_TIPO_ANNO_INIZIO_SERVIZIO = 0

sSQL = "SELECT IDRV_POTipoAnnoInizioRicorrenza "
sSQL = sSQL & " FROM RV_POConfigurazioneServizio "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_ANNO_INIZIO_SERVIZIO = 0
Else
    GET_TIPO_ANNO_INIZIO_SERVIZIO = fnNotNullN(rs!IDRV_POTipoAnnoInizioRicorrenza)
End If

rs.CloseResultset
Set rs = Nothing


Exit Function
ERR_GET_TIPO_ANNO_SERVIZIO:


End Function
Private Function GET_TIPO_ANNO_FINE_SERVIZIO(IDArticolo As Long) As Long
On Error GoTo ERR_GET_TIPO_ANNO_SERVIZIO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_TIPO_ANNO_FINE_SERVIZIO = 0

sSQL = "SELECT IDRV_POTipoAnnoFineRicorrenza "
sSQL = sSQL & " FROM RV_POConfigurazioneServizio "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_ANNO_FINE_SERVIZIO = 0
Else
    GET_TIPO_ANNO_FINE_SERVIZIO = fnNotNullN(rs!IDRV_POTipoAnnoFineRicorrenza)
End If

rs.CloseResultset
Set rs = Nothing


Exit Function
ERR_GET_TIPO_ANNO_SERVIZIO:
    
End Function
