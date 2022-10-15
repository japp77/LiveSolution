Attribute VB_Name = "ModPassaggioInterventi"
Option Explicit

Public ObjDoc As DmtDocs.cDocument
Dim cDefault As Collection
Dim cPerfoming As New CPerforming
Public rsReg As DmtOleDbLib.adoResultset

'Variabile che contiene il numero documento creato
Private VarNumeroDoc As String

Private Const NOMETABELLAPIANA = "ValoriOggettoPerTipo"
Private Const NOMETABELLADETTAGLIO = "ValoriOggettoDettaglio"

Private NumeroOreLavorate As Double

Private ArrayCli(0, 12) As String
Private ArrayArt(0, 10) As String

Dim rsFattura As DmtOleDbLib.adoResultset

Private ScontoCliente As Double
Public Link_TipoOggetto As Long
Public Link_Sezionale As Long
Public Link_SezionalePA As Long
Public Data_Documento As String
Public Link_Valuta As Long
Public Link_Magazzino As Long
Public Link_PagamentoDefault As Long
Public Var_TipoOggetto As String
Public Link_Articolo As Long
Public Var_Codice_Articolo As String
Public Var_NumeroDocumento As Long

Public LINK_IVA_CLIENTE As Long


Private oReport As dmtReportLib.dmtReport
Private progMin As Double
Private Link_IDOggetto As Long
Private Link_IDRataRiferimento As Long
Private IDRata As Long

Private TotaleRecord As Integer
Private NomeAltraSede As String


Public ARTICOLO_PIU_PERIODO_FATT As Long


'VAriabile che contengono valori di errore di elaborazione
'Contiene il valore della funzione che va inserita nel titolo del messaggio
 Private VARErroreFunzione As String
'Contiene il numero di intervento
 Private VARErroreIDIntervento As String
'Contiene l'identificativo dell'articolo elaborato
 Private VARErroreIDArticolo As String
'Contiene l'identificatico dell'anagrafica
 Private VARErroreIDAnagrafica As String
'Contiene tutta la stringa errore
 Private VARErroreGenerico As String
 
Public Sub fncPassaggioDocumenti()
On Error GoTo ERR_fncPassaggioDocumenti
Dim sSQL As String
Dim lID As Long
Dim Identificativo As Long
Dim Unita_Progresso As Double
Dim f As FileSystemObject
Dim NomeCartella As String


    ''''RECUPERO DATI CARTELLA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    NomeCartella = TrovaCartella(CSIDL_COMMON_APPDATA)
    
    Set f = New FileSystemObject
    
    If (f.FolderExists(NomeCartella & "LiveSolution") = False) Then
        f.CreateFolder NomeCartella & "LiveSolution"
    End If
    
    If (f.FolderExists(NomeCartella & "LiveSolution\ExpPDF") = True) Then
        f.DeleteFolder NomeCartella & "LiveSolution\ExpPDF", True
    End If
    
    f.CreateFolder NomeCartella & "LiveSolution\ExpPDF"
    
    Set f = Nothing
    
    NomeCartella = NomeCartella & "LiveSolution\ExpPDF\"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Set ObjDoc = New DmtDocs.cDocument

    ''''''''''''''''''''''''''INIZIO PASSAGGIO''''''''''''''''''''''''''''''''''''''''
    If RaggruppamentoAnagrafica = 0 Then
        If NumeroRecordRighe = 0 Then Exit Sub
        frmCreazioneDocumenti.ProgressBar1.Value = 0
        frmCreazioneDocumenti.ProgressBar1.Max = 100
        Unita_Progresso = FormatNumber((frmCreazioneDocumenti.ProgressBar1.Max / NumeroRecordRighe), 4)
        
        rsRateDaFatt.Filter = vbNullString
        rsRateDaFatt.MoveFirst
        While Not rsRateDaFatt.EOF
            
            DoEvents
            
            Settaggio fnNotNullN(rsRateDaFatt!EntePubblico)
            
            fncTestata 0
            
            fncRighe 0
            
            If (rsRateDaFatt!IDTipo = 1) Then
                IDRata = rsRateDaFatt!IDRV_PORateContratto
            End If
            If (rsRateDaFatt!IDTipo = 2) Then
                IDRata = rsRateDaFatt!RV_POContatoreRilevamenti
            End If
            If InserimentoDMT = True Then
                fncIDOggettoCollegatoIntervento IDRata, rsRateDaFatt!IDTipo
                
                If frmCreazioneDocumenti.chkStampa.Value = 1 Then
                    DoEvents
                    ObjDoc.Prepare2Print TheApp.IDFirm, TheApp.IDUser, ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto
                    StampaDocumento
                    DoEvents
                End If
                
                If frmCreazioneDocumenti.chkSalvaInPDF = vbChecked Then
                    DoEvents
                    ObjDoc.Prepare2Print TheApp.IDFirm, TheApp.IDUser, ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto
                    SendDocument recPDF, NomeCartella
                    DoEvents
                End If
                
            End If
            If (frmCreazioneDocumenti.ProgressBar1.Value + Unita_Progresso) >= frmCreazioneDocumenti.ProgressBar1.Max Then
                frmCreazioneDocumenti.ProgressBar1.Value = frmCreazioneDocumenti.ProgressBar1.Max
            Else
                frmCreazioneDocumenti.ProgressBar1.Value = frmCreazioneDocumenti.ProgressBar1.Value + Unita_Progresso
            End If
        
        DoEvents
        rsRateDaFatt.MoveNext
        Wend
    Else
        If NumeroRecordTesta = 0 Then Exit Sub
        frmCreazioneDocumenti.ProgressBar1.Value = 0
        frmCreazioneDocumenti.ProgressBar1.Max = 100
        Unita_Progresso = FormatNumber((frmCreazioneDocumenti.ProgressBar1.Max / NumeroRecordTesta), 4)

        rsAnagrafica.Filter = ""
        rsAnagrafica.MoveFirst
    
        While Not rsAnagrafica.EOF
            
            DoEvents
            
            Settaggio fnNotNullN(rsAnagrafica!EntePubblico)
            
            fncTestata 1
            
            fncRigheRaggruppa CLng(RaggruppamentoAnaContratto)
            
            If InserimentoDMT = True Then
                
                AGGIORNA_RATE_CONTRATTI_PER_RAGGRUPPAMENTO
                
                If (frmCreazioneDocumenti.ProgressBar1.Value + Unita_Progresso) >= frmCreazioneDocumenti.ProgressBar1.Max Then
                    frmCreazioneDocumenti.ProgressBar1.Value = frmCreazioneDocumenti.ProgressBar1.Max
                Else
                    frmCreazioneDocumenti.ProgressBar1.Value = frmCreazioneDocumenti.ProgressBar1.Value + Unita_Progresso
                End If
                
                DoEvents
                
                If frmCreazioneDocumenti.chkStampa.Value = 1 Then
                    DoEvents
                    ObjDoc.Prepare2Print TheApp.IDFirm, TheApp.IDUser, ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto
                    StampaDocumento
                    DoEvents
                End If
                
                If frmCreazioneDocumenti.chkSalvaInPDF = vbChecked Then
                    DoEvents
                    ObjDoc.Prepare2Print TheApp.IDFirm, TheApp.IDUser, ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto
                    SendDocument recPDF, NomeCartella
                    DoEvents
                End If
                
            End If
        rsAnagrafica.MoveNext
        Wend


    End If
    
    If frmCreazioneDocumenti.chkSalvaInPDF = vbChecked Then
        Shell "explorer.exe /e, " & NomeCartella, vbNormalFocus
    End If
Exit Sub

ERR_fncPassaggioDocumenti:
    MsgBox Err.Description, vbCritical, "fncPassaggioDocumenti"
    frmCreazioneDocumenti.cmdFine.Enabled = True
End Sub

Private Function fncTestata(RaggruppaPerAnagrafica As Long) As Boolean
On Error GoTo ERR_fncTestata
Dim IDLetteraIntento As Long


 
With ObjDoc.Tables
'Imposta la riga attiva per la tabella di testata
    
    ObjDoc.Tables(NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)).SetActiveRetail 1
    'Dati generici del documento
    .Field "Link_Val_cambio", Null, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Doc_data", ObjDoc.DataEmissione, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Doc_numero", ObjDoc.Numero, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Link_Doc_magazzino", Link_Magazzino, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Link_Doc_sezionale", ObjDoc.IDSezionale, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Doc_prefisso", GET_PREFISSO_SEZ(ObjDoc.IDSezionale), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
    If RaggruppaPerAnagrafica = 0 Then
        'TrovaAnagrafica IDRata
        ObjDoc.ReadDataFromCliFo rsRateDaFatt!IDAnagraficaFatturazione, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        ObjDoc.ReadDataFromCliFoSite rsRateDaFatt!IDSitoPerAnagrafica, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        
        
        ObjDoc.ReadDataFromPayment rsRateDaFatt!IDPagamentoRata, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_perc_rit_acc", ObjDoc.DBDefaults.PercentualeRitenutaAcconto, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Nom_calcola_rit_acc", fnNotNullN(rsRateDaFatt!RitenutaAcconto), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        
        LINK_IVA_CLIENTE = fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))
        'If (LINK_IVA_CLIENTE > 0) Then
        '    LINK_IVA_CLIENTE = GET_LINK_IVA_CLIENTE_ESENTE(rsRateDaFatt!IDAnagraficaFatturazione, 0, ObjDoc.DataEmissione)
        'End If
        
        'IDLetteraIntento = 0
        'If GET_CONTROLLO_NUMERO_LETTERE_INTENTO(rsRateDaFatt!IDAnagraficaFatturazione, TheApp.IDFirm, Date) = 1 Then
        '    IDLetteraIntento = GET_LINK_LETTERA_INTENTO(rsRateDaFatt!IDAnagraficaFatturazione, TheApp.IDFirm, Year(Date))
        '    LINK_IVA_CLIENTE = GET_LINK_IVA_LETTERA_INTENTO(IDLetteraIntento, LINK_IVA_CLIENTE)
        'End If
        
        'ObjDoc.Field "Link_Nom_IVA", LINK_IVA_CLIENTE, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        'ObjDoc.Field "Link_Nom_lettera_intento", IDLetteraIntento, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        
        If fnNotNullN(rsRateDaFatt!IDRaggruppamentoFatturato) > 0 Then
            ObjDoc.Field "Link_Nom_raggrup_fatturato", fnNotNullN(rsRateDaFatt!IDRaggruppamentoFatturato), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        End If
        If fnNotNullN(rsRateDaFatt!IDAccordoCommerciale) > 0 Then
            ObjDoc.Field "Link_Nom_accordi_commerciali", fnNotNullN(rsRateDaFatt!IDAccordoCommerciale), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        End If
        If fnNotNullN(rsRateDaFatt!IDContrattoBancario) > 0 Then
            ObjDoc.Field "Link_Nom_contratto_bancario", fnNotNullN(rsRateDaFatt!IDContrattoBancario), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        End If
        
        If fnNotNullN(rsRateDaFatt!IDAnagraficaAgente) > 0 Then
            ObjDoc.ReadDataFromAgent fnNotNullN(rsRateDaFatt!IDAnagraficaAgente)
        End If
        
        
    Else
        ObjDoc.ReadDataFromCliFo rsAnagrafica!IDAnagrafica, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        ObjDoc.ReadDataFromCliFoSite rsAnagrafica!IDSitoPerAnagrafica, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        
        ObjDoc.ReadDataFromPayment rsAnagrafica!IDPagamento, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_perc_rit_acc", ObjDoc.DBDefaults.PercentualeRitenutaAcconto, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Nom_calcola_rit_acc", fnNotNullN(rsAnagrafica!RitenutaAcconto), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        
        LINK_IVA_CLIENTE = fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))
        
        'If (LINK_IVA_CLIENTE = 0) Then
        '    LINK_IVA_CLIENTE = GET_LINK_IVA_CLIENTE_ESENTE(rsAnagrafica!IDAnagrafica, 0, ObjDoc.DataEmissione)
        'End If
        
'        IDLetteraIntento = 0
'        If GET_CONTROLLO_NUMERO_LETTERE_INTENTO(rsAnagrafica!IDAnagrafica, TheApp.IDFirm, Date) = 1 Then
'            IDLetteraIntento = GET_LINK_LETTERA_INTENTO(rsAnagrafica!IDAnagrafica, TheApp.IDFirm, Year(Date))
'            LINK_IVA_CLIENTE = GET_LINK_IVA_LETTERA_INTENTO(IDLetteraIntento, LINK_IVA_CLIENTE)
'        End If
'
'        ObjDoc.Field "Link_Nom_IVA", LINK_IVA_CLIENTE, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
'        ObjDoc.Field "Link_Nom_lettera_intento", IDLetteraIntento, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        
        If fnNotNullN(rsAnagrafica!IDRaggruppamentoFatturato) > 0 Then
            ObjDoc.Field "Link_Nom_raggrup_fatturato", fnNotNullN(rsAnagrafica!IDRaggruppamentoFatturato), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        End If
        If fnNotNullN(rsAnagrafica!IDAccordoCommerciale) > 0 Then
            ObjDoc.Field "Link_Nom_accordi_commerciali", fnNotNullN(rsAnagrafica!IDAccordoCommerciale), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        End If
        If fnNotNullN(rsAnagrafica!IDContrattoBancario) > 0 Then
            ObjDoc.Field "Link_Nom_contratto_bancario", fnNotNullN(rsAnagrafica!IDContrattoBancario), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        End If
        If fnNotNullN(rsAnagrafica!IDAnagraficaAgente) > 0 Then
            ObjDoc.ReadDataFromAgent fnNotNullN(rsAnagrafica!IDAnagraficaAgente)
        End If
    End If

    If fnNotNullN(.Field("Link_Doc_pagamento", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) = 0 Then
        .Field "Link_Doc_pagamento", Link_PagamentoDefault, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    If fnNotNullN(.Field("Link_Val_valuta", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) = 0 Then
        .Field "Link_Val_valuta", Link_Valuta, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    
    ObjDoc.Field "Link_Spe_esenti_art_10_IVA", ObjDoc.DBDefaults.Link_Spe_esenti_art_10_IVA, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Link_Spe_bolli_eff_art_15_IVA", ObjDoc.DBDefaults.Link_Spe_bolli_eff_art_15_IVA, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
End With

fncTestata = True
     
Exit Function
ERR_fncTestata:
MsgBox Err.Description, vbCritical, "ERR_fncTestata"

End Function

Private Function fncRighe(RaggruppaPerCliente As Long) As Boolean
On Error GoTo ERR_fncRighe
VARErroreFunzione = "fncRighe"
Dim Prova As Boolean
Dim Ciao As Double
Dim I As Integer
Dim sSQL As String
Dim rsART As DmtOleDbLib.adoResultset
Dim rsAgente As DmtOleDbLib.adoResultset
'Serve per vedere se esiste uno sconto a livello di articolo
'1 = Esiste
'0 = Non Esiste
Dim VarRegSconto As Integer
Dim IDArticolo As Long
Dim DescrizioneArticolo As String
Dim IDIvaFatturazione As Long
Dim AliquotaIvaFatturazione As Double
Dim IDContoContabile As Long
Dim CodicePDC As String
Dim DescrizionePDC As String
Dim SplitPeriodo() As String
Dim J As Long
Dim AnnotazioniStringa As String

I = 1

If fnNotNullN(rsRateDaFatt!IDAnagraficaFatturazione) <> fnNotNullN(rsRateDaFatt!IDAnagrafica) Then
    ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
    With ObjDoc.Tables
        .Field "Art_descrizione", GET_RIFERIMENTO_ANA_CONTRATTO(fnNotNullN(rsRateDaFatt!IDAnagrafica), fnNotNullN(rsRateDaFatt!IDSitoPerAnagrafica)), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    End With
    I = I + 1
End If

If StampaRifContratto = 1 Then
    ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
     With ObjDoc.Tables
         .Field "Art_descrizione", "Riferimento contratto numero " & fnNotNullN(rsRateDaFatt!AnnoContratto) & "-" & fnNotNullN(rsRateDaFatt!NumeroContratto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
         .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
         .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
         .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
         .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
     End With
     I = I + 1
  
End If

ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
        
    
With ObjDoc.Tables
    'IDArticolo'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If fnNotNullN(rsRateDaFatt!IDRV_POTipoImpostazioneContratto) <= 2 Then
        IDArticolo = fnNotNullN(rsRateDaFatt!IDArticoloRataContratto)
        IDIvaFatturazione = fnNotNullN(rsRateDaFatt!IDIvaArticoloRata)
        AliquotaIvaFatturazione = fnNotNullN(rsRateDaFatt!AliquotaIvaArticoloRata)
        
        If IDArticolo = 0 Then
            IDArticolo = fnNotNullN(rsRateDaFatt!IDArticoloContratto)
            IDIvaFatturazione = fnNotNullN(rsRateDaFatt!IDIvaVenditaArticoloContratto)
            AliquotaIvaFatturazione = fnNotNullN(rsRateDaFatt!AliquotaIvaArticoloContratto)
            If IDArticolo = 0 Then
                IDArticolo = fnNotNullN(rsRateDaFatt!IDArticoloTipoContratto)
                IDIvaFatturazione = fnNotNullN(rsRateDaFatt!IDIvaVenditaTipoContratto)
                AliquotaIvaFatturazione = fnNotNullN(rsRateDaFatt!AliquotaIvaTipoContratto)
            
                If IDArticolo = 0 Then
                    IDArticolo = Link_Articolo
                    IDIvaFatturazione = GET_IVA_ARTICOLO(IDArticolo)
                    AliquotaIvaFatturazione = GET_ALIQUOTA_IVA_ARTICOLO(IDIvaFatturazione)
                End If
            
            End If
        End If
    Else
'        IDArticolo = GET_LINK_ARTICOLO_DA_PROD_CONTR(rsRateDaFatt!IDRV_POContrattoProdotti)
'        IDIvaFatturazione = GET_IVA_ARTICOLO(IDArticolo)
'        AliquotaIvaFatturazione = GET_ALIQUOTA_IVA_ARTICOLO(IDIvaFatturazione)
        IDArticolo = GET_LINK_ARTICOLO_DA_PROD_CONTR(rsRateDaFatt!IDRV_POContrattoProdotti, IDIvaFatturazione, AliquotaIvaFatturazione)
        If (IDIvaFatturazione = 0) Then
            IDIvaFatturazione = GET_IVA_ARTICOLO(IDArticolo)
            AliquotaIvaFatturazione = GET_ALIQUOTA_IVA_ARTICOLO(IDIvaFatturazione)
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'DESCRIZIONE ARTICOLO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If fnNotNullN(rsRateDaFatt!IDRV_POTipoImpostazioneContratto) <= 2 Then
        DescrizioneArticolo = Trim(fnNotNull(rsRateDaFatt!Periodo))
        If Len(DescrizioneArticolo) = 0 Then
            DescrizioneArticolo = GET_DESCRIZIONE_ARTICOLO(IDArticolo)
        End If
    Else
        DescrizioneArticolo = GET_DESCRIZIONE_ARTICOLO(IDArticolo)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Conto contabile'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    IDContoContabile = fnNotNullN(rsRateDaFatt!IDContoPDCContratto)
    CodicePDC = fnNotNull(rsRateDaFatt!CodicePDCContratto)
    DescrizionePDC = fnNotNull(rsRateDaFatt!DescrizionePDCContratto)
    If IDContoContabile = 0 Then
        IDContoContabile = fnNotNullN(rsRateDaFatt!IDContoPDCTipoContratto)
        CodicePDC = fnNotNull(rsRateDaFatt!CodicePDCTipoContratto)
        DescrizionePDC = fnNotNull(rsRateDaFatt!DescrizionePDCTipoContratto)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ObjDoc.ReadDataFromArticle IDArticolo
    .Field "Art_descrizione", DescrizioneArticolo, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    '.Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    If fnNotNullN(rsRateDaFatt!IDRV_POTipoImpostazioneContratto) <= 2 Then
        .Field "Art_quantita_totale", fnNotNullN(rsRateDaFatt!Quantita), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Art_prezzo_unitario_neutro", fnNotNullN(rsRateDaFatt!ImportoRataUnitaria), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        
        .Field "Art_sco_in_percentuale_1", fnNotNullN(rsRateDaFatt!Sconto1), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Art_sco_in_percentuale_2", fnNotNullN(rsRateDaFatt!Sconto2), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    Else
        If fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti) > 0 Then
            If fnNotNullN(rsRateDaFatt!IDRV_POProdotto) > 0 Then
                If (GET_CONTROLLO_UNA_RATA(fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti), fnNotNullN(rsRateDaFatt!IDRV_POContratto)) = True) Then
                    GET_DATI_DA_RIGA_CONTR fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti), True
                Else
                    .Field "Art_quantita_totale", fnNotNullN(rsRateDaFatt!Quantita), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_prezzo_unitario_neutro", fnNotNullN(rsRateDaFatt!ImportoRataUnitaria), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                    .Field "Art_sco_in_percentuale_1", fnNotNullN(rsRateDaFatt!Sconto1), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_sco_in_percentuale_2", fnNotNullN(rsRateDaFatt!Sconto2), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                    GET_DATI_DA_RIGA_CONTR_PIU_RATE fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti), fnNotNullN(rsRateDaFatt!IDRV_POContratto), fnNotNullN(rsRateDaFatt!ImportoRataUnitaria)
                    
                End If
            Else
                If fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti) > 0 Then
                    GET_DATI_DA_RIGA_CONTR fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti)
                Else
                    .Field "Art_quantita_totale", fnNotNullN(rsRateDaFatt!Quantita), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_prezzo_unitario_neutro", fnNotNullN(rsRateDaFatt!ImportoRataUnitaria), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                    .Field "Art_sco_in_percentuale_1", fnNotNullN(rsRateDaFatt!Sconto1), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_sco_in_percentuale_2", fnNotNullN(rsRateDaFatt!Sconto2), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End If
            End If
        Else
            .Field "Art_quantita_totale", fnNotNullN(rsRateDaFatt!Quantita), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_prezzo_unitario_neutro", fnNotNullN(rsRateDaFatt!ImportoRataUnitaria), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            
            .Field "Art_sco_in_percentuale_1", fnNotNullN(rsRateDaFatt!Sconto1), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_sco_in_percentuale_2", fnNotNullN(rsRateDaFatt!Sconto2), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        End If
    End If
    
    If LINK_IVA_CLIENTE > 0 Then
        .Field "Link_art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    Else
        .Field "Link_Art_IVA", IDIvaFatturazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Art_aliquota_IVA", AliquotaIvaFatturazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    End If
    .Field "Link_art_magazzino", Link_Magazzino, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        
    If IDContoContabile > 0 Then
        .Field "Link_Art_IDCContropartita", IDContoContabile, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    End If
    .Field "Art_data_inizio_competenza", fnNotNullN(rsRateDaFatt!DataInizioPeriodo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    .Field "Art_data_fine_competenza", fnNotNullN(rsRateDaFatt!DataFinePeriodo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    ''''''ANAGRAFICA AGENTE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome, Agente.Codice, RegolaProvvAgente.IDRegolaProvvAgente, "
    sSQL = sSQL & "RegolaProvvAgente.Predefinita , RegolaProvvAgente.IDRegolaProvv, RegolaProvv.RegolaProvv "
    sSQL = sSQL & "FROM Agente INNER JOIN "
    sSQL = sSQL & "Anagrafica ON Agente.IDAnagrafica = Anagrafica.IDAnagrafica INNER JOIN "
    sSQL = sSQL & "RegolaProvvAgente ON Anagrafica.IDAnagrafica = RegolaProvvAgente.IDAnagrafica INNER JOIN "
    sSQL = sSQL & "RegolaProvv ON RegolaProvvAgente.IDRegolaProvv = RegolaProvv.IDRegolaProvv "
    sSQL = sSQL & "WHERE (Anagrafica.IDAnagrafica =" & fnNotNullN(rsRateDaFatt!IDAnagraficaAgente) & ") "
    sSQL = sSQL & "AND (RegolaProvvAgente.Predefinita =" & fnNormBoolean(1) & ")"
    
    Set rsAgente = CnDMT.OpenResultset(sSQL)
    
    If rsAgente.EOF = False Then
        ObjDoc.Field "Art_age_ragione_sociale", fnNotNull(rsAgente!Anagrafica), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_age_codice", fnNotNull(rsAgente!Codice), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_age_nome", fnNotNull(rsAgente!Nome), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Link_Art_agente", fnNotNullN(rsAgente!IDAnagrafica), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Link_Art_age_regola_provv", fnNotNullN(rsAgente!IDRegolaProvv), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_age_regola_provv", fnNotNull(rsAgente!RegolaProvv), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    End If
    
    rsAgente.CloseResultset
    Set rsAgente = Nothing
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End With

If fnNotNullN(rsRateDaFatt!IDRV_POTipoImpostazioneContratto) <= 2 Then
    If rsRateDaFatt!IDTipo = 1 Then
        If fnNotNullN(rsRateDaFatt!IDRV_POContrattoAdeguamento) = 0 Then
            If rsRateDaFatt!NoteFattura <> "" Then
                I = I + 1
                ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                With ObjDoc.Tables
                    .Field "Art_descrizione", rsRateDaFatt!NoteFattura, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End With
            End If
        End If
    End If
End If

If Len(Trim(fnNotNull(rsRateDaFatt!Periodo))) > 0 Then
    I = I + 1
    If fnNotNullN(rsRateDaFatt!IDRV_POTipoImpostazioneContratto) = 3 Then
        If (ARTICOLO_PIU_PERIODO_FATT = 1) Then
            SplitPeriodo = Split(fnNotNull(rsRateDaFatt!Periodo), vbCrLf)
            For J = 0 To UBound(SplitPeriodo)
                ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                With ObjDoc.Tables
                    .Field "Art_descrizione", Mid(SplitPeriodo(J), 1, 255), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End With
            I = I + 1
            Next
        End If
    End If
End If


If fnNotNullN(rsRateDaFatt!IDRV_POTipoImpostazioneContratto) = 3 Then
    'SE CI SONO LE ANNOTAZIONI INSERIRLE NELLA FATTURA NELLO STESSO MODO DEL PERIODO
    AnnotazioniStringa = GET_ANNOTAZIONI_DA_PRODOTTO(fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti), fnNotNullN(rsRateDaFatt!IDRV_POContratto))
    If Len(AnnotazioniStringa) > 0 Then
        SplitPeriodo = Split(AnnotazioniStringa, vbCrLf)
        For J = 0 To UBound(SplitPeriodo)
            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
            With ObjDoc.Tables
                .Field "Art_descrizione", Mid(SplitPeriodo(J), 1, 255), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            End With
        I = I + 1
        Next
    End If
End If

RIPORTA_DESCRIZIONI_AUTOMATICHE I + 1, TheApp.Branch, ObjDoc.IDTipoOggetto

fncRighe = True
Exit Function
ERR_fncRighe:
    fncRighe = False


    MsgBox Err.Description, vbCritical, "ERR_fncRighe"
    
End Function

Private Function fnContrattoAltraFiliale() As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT RV_POContratto.IDSitoPerAnagrafica, SitoPerAnagrafica.SitoPerAnagrafica "
    sSQL = sSQL & "FROM RV_POContratto LEFT OUTER JOIN "
    sSQL = sSQL & "SitoPerAnagrafica ON dbo.RV_POContratto.IDSitoPerAnagrafica = SitoPerAnagrafica.IDSitoPerAnagrafica "
    sSQL = sSQL & "WHERE IDRV_POContratto=" & 0
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        fnContrattoAltraFiliale = False
    Else
        
        If IsNull(rs!IDSitoPerAnagrafica) Then
            fnContrattoAltraFiliale = False
        ElseIf (rs!IDSitoPerAnagrafica = 0) Then
            fnContrattoAltraFiliale = False
        Else
            NomeAltraSede = fnNotNull(rs!SitoPerAnagrafica)
            fnContrattoAltraFiliale = True
        End If
        
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub Settaggio(EntePubblico As Integer)
    With ObjDoc
        Set .Connection = CnDMT
        .IDAzienda = TheApp.IDFirm
        .IDAttivitaAzienda = GET_ATTIVITA_AZIENDA
        .IDFiliale = TheApp.Branch
        .SetTipoOggetto Link_TipoOggetto
        .IDFunzione = fncTrovaIDFunzione
        .UseAutomation = True
        .IDEsercizio = GET_LINK_ESERCIZIO(Data_Documento)
        If EntePubblico = 0 Then
            .IDSezionale = Link_Sezionale
        Else
            If Link_SezionalePA = 0 Then
                .IDSezionale = Link_Sezionale
            Else
                .IDSezionale = Link_SezionalePA
            End If
        End If
        .IDTipoAnagrafica = 2
        .IDUtente = TheApp.IDUser
        .Descrizione = Var_TipoOggetto
        .DataEmissione = Data_Documento
        .Numero = Var_NumeroDocumento
        If .Tables.Count = 0 Then
        'Se Tables.Count = 0 vuol dire che l'oggetto
        'DmtDocs non è mai stato inizializzato
            .Clear
            .SetTipoOggetto Link_TipoOggetto
        Else
            .ClearValues
        End If
    
    End With
End Sub
Private Function InserimentoDMT() As Boolean
On Error GoTo ERR_InserimentoDMT
VARErroreFunzione = "InserimentoDMT"


Screen.MousePointer = vbHourglass
    
    Set ObjDoc.Scadenze = Nothing
    ObjDoc.PerformDocument Nothing
    
    VarNumeroDoc = ObjDoc.Insert
    
    If VarNumeroDoc > 0 Then
        InserimentoDMT = True
    Else
        InserimentoDMT = False
    End If
    
    
    
    
Screen.MousePointer = vbDefault
    
Exit Function

ERR_InserimentoDMT:
    InserimentoDMT = False
    MsgBox Err.Description, vbCritical, "Creazione fattura"
End Function

Private Sub SetDefault()

            
            
    'Imposta i default per i campi relativi alla valuta corrente
    'al tipo di arrotondamento, alle spese e ai bolli
    'Questi default sono obbligatori per il calcolo del documento
    Set cDefault = New Collection
    'Valore di arrotondamento per la valuta corrente
    cDefault.Add 1, "Val_arrotondamento"
    'Tipo di arrotondomento per la valuta corrente
    cDefault.Add 1, "Link_Val_tipo_arrotondamento"
    'ID della valuta corrente
    cDefault.Add 0, "Link_Val_valuta"
    'Spesse incassi in percentuale
    'cDefault.Add 5.16, "Spe_incasso_netto_IVA"
    'cDefault.Add 2, "Link_Nom_spese_incasso"
    'Importo del bollo
    cDefault.Add 0, "Nom_bollo_esente"
    'Importo limite per il pagamento del bollo
    cDefault.Add 0, "Nom_bollo_esente_limite"
    'ID del contratto bancario azienda
    'cDefault.Add 11, "Link_Doc_contratto_bancario_az"
    'cDefault.Add 37261, "Link_Nom_contratto_bancario"
    'ID della natura delle scadenze
    cDefault.Add 0, "IDNaturaScadenza"
End Sub

Public Function fncTrovaIDFunzione() As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT IDFunzione FROM Funzione WHERE IDTipoOggetto = " & Link_TipoOggetto & " ORDER BY IDFunzione"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        fncTrovaIDFunzione = rs!IDFunzione
    End If
    
    
    rs.CloseResultset
    Set rs = Nothing

    
End Function
Private Function fncGetNumeroIntervento(IDIntervento As Long) As String

Dim sSQL As String
Dim rsInt As DmtOleDbLib.adoResultset

sSQL = "SELECT NumeroIntervento, AnnoIntervento FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_PoIntervento=" & IDIntervento

Set rsInt = CnDMT.OpenResultset(sSQL)

If rsInt.EOF = False Then
    fncGetNumeroIntervento = fnNotNull(rsInt!NumeroIntervento) & "-" & fnNotNull(rsInt!annoIntervento)
Else
    fncGetNumeroIntervento = "Impossibile recuperare il numero intervento (" & IDIntervento & ")"
End If

rsInt.CloseResultset
Set rsInt = Nothing

End Function
Public Sub StampaDocumento()
On Error GoTo ERR_StampaDocumento

Dim IDReport As Long

    Set oReport = New dmtReportLib.dmtReport
    Set oReport.Connection = CnDMT
    If MenuOptions.DBType = 1 Then
        'parametri di accesso al database ACCESS
        oReport.Password = "dmt192981046"
        oReport.User = "admin"
    Else
        'parametri di accesso al database SQL Server
        oReport.Password = TheApp.Password
        oReport.User = TheApp.User
    End If


    oReport.BranchID = TheApp.Branch 'IDFiliale

    oReport.DocTypeID = ObjDoc.IDTipoOggetto
    'oReport.Where = "IDOggetto = 873" '& Val(Me.Txt_Reg_IDRegistro)
    oReport.Where = "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto) & ".IDOggetto = " & ObjDoc.IDOggetto
    oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser

    oReport.Copies = frmCreazioneDocumenti.txtNumeroCopie.Text
    If (Len(frmCreazioneDocumenti.cboStampante.Text)) = 0 Then
        oReport.DoPrint Printer.DeviceName
    Else
        oReport.DoPrint frmCreazioneDocumenti.cboStampante.Text
    End If
        
    
            
Exit Sub
ERR_StampaDocumento:
    MsgBox Err.Description, vbCritical, "Stampa Documento"
End Sub

Private Function GET_NOME_FILE_DOCUMENTO() As String

Dim NomeFile As String

GET_NOME_FILE_DOCUMENTO = ""

GET_NOME_FILE_DOCUMENTO = ObjDoc.Descrizione & ""
GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " " & fnNotNull(ObjDoc.Field("Nom_ragione_sociale_o_cognome", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) & fnNotNull(ObjDoc.Field("Nom_nome", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))
'GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " (" & fnNotNull(ObjDoc.Field("Nom_codice", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) & ")"
GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " [" & GET_DATA_FORMATTATA(ObjDoc.DataEmissione) & "]"
    
'GET_NOME_FILE_DOCUMENTO = ObjDoc.Descrizione & ""
'GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " " & fnNotNull(rsGrigliaSelDocTMP!Nom_ragione_sociale_o_cognome) & fnNotNull(rsGrigliaSelDocTMP!Nom_nome)
'GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " (" & fnNotNull(rsGrigliaSelDocTMP!Nom_codice) & ")"
'GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " [" & GET_DATA_FORMATTATA(rsGrigliaSelDocTMP!Doc_data) & "]"


End Function
Private Function GET_DATA_FORMATTATA(DataF As String) As String
On Error GoTo ERR_GET_DATA_FORMATTATA
Dim Anno As String
Dim mese As String
Dim giorno As String

GET_DATA_FORMATTATA = ""

Anno = Year(DataF)
mese = Month(DataF)
giorno = Day(DataF)

If Len(mese) = 1 Then mese = "0" & mese
If Len(giorno) = 1 Then giorno = "0" & giorno

GET_DATA_FORMATTATA = Anno & "-" & mese & "-" & giorno

GET_DATA_FORMATTATA = GET_DATA_FORMATTATA & " n. "

If Len(Trim(fnNotNull(ObjDoc.Field("Doc_Prefisso", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))))) > 0 Then
    GET_DATA_FORMATTATA = GET_DATA_FORMATTATA & Trim(fnNotNull(ObjDoc.Field("Doc_Prefisso", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))) & "-"
End If

GET_DATA_FORMATTATA = GET_DATA_FORMATTATA & ObjDoc.Numero

Exit Function
ERR_GET_DATA_FORMATTATA:

End Function

Private Sub SendDocument(ByVal Appl As Long, percorso As String)
On Error GoTo errHandler
Dim OLDCursor As Integer
Dim SExt As String
Dim DataDocumento As String
Dim NomeCartella As String
Dim NomeFile As String
Dim InvioEmailPersonalizzata As Boolean
    
    Set oReport = New dmtReportLib.dmtReport
    Set oReport.Connection = CnDMT
    If MenuOptions.DBType = 1 Then
        'parametri di accesso al database ACCESS
        oReport.Password = "dmt192981046"
        oReport.User = "admin"
    Else
        'parametri di accesso al database SQL Server
        oReport.Password = TheApp.Password
        oReport.User = TheApp.User
    End If


    oReport.BranchID = TheApp.Branch 'IDFiliale

    oReport.DocTypeID = ObjDoc.IDTipoOggetto
    'oReport.Where = "IDOggetto = 873" '& Val(Me.Txt_Reg_IDRegistro)
    oReport.Where = "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto) & ".IDOggetto = " & ObjDoc.IDOggetto
    oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
    
    
    OLDCursor = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
   
    Select Case Appl
        Case 0
            SExt = ".xls"
        Case 1
            SExt = ".doc"
        Case 2
            SExt = ".html"
        Case 3
            SExt = ".pdf"
    End Select
 
    NomeFile = GET_NOME_FILE_DOCUMENTO '& SExt
    
    oReport.ExportFileName = percorso & NomeFile
    oReport.ShowExportFile = False
    oReport.Export Appl
    
    Screen.MousePointer = OLDCursor
    
    
    Exit Sub
errHandler:
    Screen.MousePointer = OLDCursor
    
    MsgBox Err.Description, vbCritical, "SendDocument"
    

End Sub


Private Function fncTrovaReport(NomeReport As String, IDTipoOggetto As Long) As Long
On Error GoTo ERR_fncTrovaReport
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDReportTipoOggetto FROM ReportTipoOggetto "
sSQL = sSQL & "WHERE ((ReportTipoOggetto=" & fnNormString(NomeReport) & ") AND (IDTipoOggetto=" & IDTipoOggetto & "))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaReport = rs!IDReportTipoOggetto
Else
    fncTrovaReport = 0
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_fncTrovaReport:
    MsgBox Err.Description, vbCritical, "Trova report per stampa"
    fncTrovaReport = 0

End Function
Private Function fncImpostaDefaultReport(ByVal IDReportDefault As Long)
On Error GoTo ERR_fncImpostaDefaultReport
    Dim sSQL As String
    
    
    sSQL = "UPDATE DefaultFilialePerTipoOggetto SET "
    sSQL = sSQL & "IDReportTipoOggetto=" & IDReportDefault
    sSQL = sSQL & " WHERE IDTipoOggetto = " & Link_TipoOggetto & " AND IDFiliale = " & TheApp.Branch
    
    CnDMT.Execute sSQL
    
Exit Function
ERR_fncImpostaDefaultReport:
    MsgBox Err.Description, vbCritical, "Settaggio report di default"
End Function
Private Function fncTrovaDocumento() As Long
On Error GoTo ERR_fncTrovaDocumento
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto From Oggetto WHERE ("
sSQL = sSQL & "(IDTipoOggetto=" & Link_TipoOggetto & ") "
sSQL = sSQL & "AND (Numero=" & fnNormString(VarNumeroDoc) & ") "
sSQL = sSQL & "AND (DataEmissione=" & fnNormDate(Data_Documento) & ")) "
sSQL = sSQL & "ORDER BY IDOggetto DESC"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaDocumento = rs!IDOggetto
Else
    fncTrovaDocumento = 0
End If

rs.CloseResultset
Set rs = Nothing



Exit Function
ERR_fncTrovaDocumento:
    MsgBox Err.Description, vbCritical, "Impossibile stampare"
    fncTrovaDocumento = 0
End Function
Private Sub fncIDOggettoCollegatoIntervento(IDRata As Long, IDTipo As Long)
Dim sSQL As String
Dim IDOggettoScadenza As Long
Dim IDOggettoRata As Long
Dim IDTipoOggettoRata As Long

If IDRata > 0 Then
    If IDTipo = 1 Then
        IDOggettoRata = fnNotNullN(rsRateDaFatt!IDOggetto)
        IDTipoOggettoRata = fnNotNullN(rsRateDaFatt!IDTipoOggetto)
        
        If IDTipoOggettoRata = 0 Then
            IDTipoOggettoRata = fnGetTipoOggetto("RV_PORateContratto")
        End If
        
        If IDOggettoRata = 0 Then
            IDOggettoRata = GET_LINK_OGGETTO(IDOggettoRata, IDTipoOggettoRata, rsRateDaFatt!NumeroRata, rsRateDaFatt!DataRata)
        End If
        
        sSQL = "UPDATE RV_PORateContratto SET"
        sSQL = sSQL & " IDOggettoCollegato=" & ObjDoc.IDOggetto & ", "
        sSQL = sSQL & " Fatturata=1, "
        sSQL = sSQL & " IDOggetto=" & IDOggettoRata & ", "
        sSQL = sSQL & " IDTipoOggetto=" & IDTipoOggettoRata
        sSQL = sSQL & " WHERE IDRV_PORateContratto=" & fnNotNullN(rsRateDaFatt!IDRV_PORateContratto)
        CnDMT.Execute sSQL
    End If

    If IDTipo = 2 Then
        IDOggettoRata = fnNotNullN(rsRateDaFatt!IDOggetto)
        IDTipoOggettoRata = fnNotNullN(rsRateDaFatt!IDTipoOggetto)
        

        sSQL = "UPDATE RV_POContatoreRilevamenti SET"
        sSQL = sSQL & " IDOggettoCollegato=" & ObjDoc.IDOggetto & ", "
        sSQL = sSQL & " IDTipoOggettoCollegato=" & ObjDoc.IDTipoOggetto & ", "
        sSQL = sSQL & " Fatturata=1 "
        sSQL = sSQL & " WHERE IDRV_POContatoreRilevamenti=" & fnNotNullN(rsRateDaFatt!IDRV_POContatoreRilevamenti)
        CnDMT.Execute sSQL
    End If
End If


If (IDOggettoRata > 0) And (IDTipoOggettoRata > 0) Then
    
    If rsRateDaFatt!IDTipo = 1 Then
        CREA_FLUSSO_DOCUMENTALE ObjDoc.IDTipoOggetto, ObjDoc.IDOggetto, IDOggettoRata, IDTipoOggettoRata, "Documento di vendita -> Rata contratto"
    End If
    
    If rsRateDaFatt!IDTipo = 1 Then
        CREA_FLUSSO_DOCUMENTALE ObjDoc.IDTipoOggetto, ObjDoc.IDOggetto, IDOggettoRata, IDTipoOggettoRata, "Documento di vendita -> Rilevamento"
    End If
    
    If (rsRateDaFatt!IDTipo = 1) Then
        ''''''ELIMINAZIONE SCADENZA DEL CONTRATTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(fnNotNullN(rsRateDaFatt!IDOggetto), fnNotNullN(rsRateDaFatt!IDTipoOggetto), 0)
        
        If IDOggettoScadenza > 0 Then
            ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, fnNotNullN(rsRateDaFatt!IDOggetto), fnNotNullN(rsRateDaFatt!IDTipoOggetto), "Rata contratto -> Scadenza"
            ELIMINA_SCADENZA IDOggettoScadenza
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
End If

End Sub
Private Function fncTrovaIDOggettoCollegato(IDTipoOggetto As Long, NumeroIntervento As String) As Long
'On Error GoTo ERR_fncTrovaIDOggettoCollegato
Dim sSQL As String
Dim rsOgg As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto FROM Oggetto "
sSQL = sSQL & "WHERE ("
sSQL = sSQL & "(Numero=" & fnNormString(NumeroIntervento) & ") AND "
sSQL = sSQL & "(IDTipoOggetto=" & IDTipoOggetto & ") AND "
sSQL = sSQL & "(IDAzienda=" & TheApp.IDFirm & ") AND "

sSQL = sSQL & "(IDAttivitaAzienda=" & GET_ATTIVITA_AZIENDA & ") AND "
sSQL = sSQL & "(IDSezionale=" & ObjDoc.IDSezionale & ") AND "
sSQL = sSQL & "(DataEmissione=" & fnNormDate(ObjDoc.DataEmissione) & ")) "

Set rsOgg = CnDMT.OpenResultset(sSQL)

If rsOgg.EOF = False Then
    fncTrovaIDOggettoCollegato = rsOgg!IDOggetto
Else
    fncTrovaIDOggettoCollegato = 0
End If

Exit Function
ERR_fncTrovaIDOggettoCollegato:
    fncTrovaIDOggettoCollegato = 0
End Function
Private Sub TotaleRateDaFatturare()
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT COUNT(IDTMP) AS Totale "
    sSQL = sSQL & "From RV_POTMPFatturazioneRate "
    sSQL = sSQL & "WHERE NonFatturare=0"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        If IsNull(rs!Totale) Then
            
            TotaleRecord = 0
        Else
            
            TotaleRecord = rs!Totale
        End If
    Else
        
        TotaleRecord = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
End Sub
Private Function GET_ATTIVITA_AZIENDA() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAttivitaAzienda FROM AttivitaAzienda "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ATTIVITA_AZIENDA = 0
Else
    GET_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub GET_CODICE_ARTICOLO_PER_TIPO_CONTRATTO(IDTipoContratto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDArticolo FROM RV_POTipoContratto "
sSQL = sSQL & "WHERE IDRV_POTipoContratto=" & IDTipoContratto

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!IDArticolo) > 0 Then
        Link_Articolo = fnNotNullN(rs!IDArticolo)
        Var_Codice_Articolo = GET_CODICE_ARTICOLO(fnNotNullN(rs!IDArticolo))

    End If
End If



rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_CODICE_ARTICOLO(IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CodiceArticolo FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CODICE_ARTICOLO = fnNotNull(rs!CodiceArticolo)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function fncRigheRaggruppa(TipoRaggruppamento As Long) As Boolean
'On Error GoTo ERR_fncRighe
VARErroreFunzione = "fncRighe"
Dim Prova As Boolean
Dim Ciao As Double
Dim I As Integer
Dim sSQL As String
Dim rsART As DmtOleDbLib.adoResultset
Dim rsAgente As DmtOleDbLib.adoResultset
'Serve per vedere se esiste uno sconto a livello di articolo
'1 = Esiste
'0 = Non Esiste
Dim VarRegSconto As Integer
Dim IDArticolo As Long
Dim DescrizioneArticolo As String
Dim IDIvaFatturazione As Long
Dim AliquotaIvaFatturazione As Double
Dim IDContoContabile As Long
Dim CodicePDC As String
Dim DescrizionePDC As String
Dim IDAnagraficaContratto As Long
Dim SplitPeriodo() As String
Dim J As Long
Dim IDContratto As Long
Dim AnnotazioniStringa As String

I = 1
    
    rsRateDaFatt.Filter = "IDAnagraficaFatturazione=" & fnNotNullN(rsAnagrafica!IDAnagrafica)
    rsRateDaFatt.Filter = rsRateDaFatt.Filter & " AND IDPagamentoRata=" & fnNotNullN(rsAnagrafica!IDPagamento)
    rsRateDaFatt.Filter = rsRateDaFatt.Filter & " AND RitenutaAcconto=" & fnNotNullN(rsAnagrafica!RitenutaAcconto)
    rsRateDaFatt.Filter = rsRateDaFatt.Filter & " AND IDRaggruppamentoFatturato=" & fnNotNullN(rsAnagrafica!IDRaggruppamentoFatturato)
    rsRateDaFatt.Filter = rsRateDaFatt.Filter & " AND IDAccordoCommerciale=" & fnNotNullN(rsAnagrafica!IDAccordoCommerciale)
    rsRateDaFatt.Filter = rsRateDaFatt.Filter & " AND IDContrattoBancario=" & fnNotNullN(rsAnagrafica!IDContrattoBancario)
    rsRateDaFatt.Filter = rsRateDaFatt.Filter & " AND EntePubblico=" & fnNotNullN(rsAnagrafica!EntePubblico)
    rsRateDaFatt.Filter = rsRateDaFatt.Filter & " AND IDSitoPerAnagrafica=" & fnNotNullN(rsAnagrafica!IDSitoPerAnagrafica)
    rsRateDaFatt.Filter = rsRateDaFatt.Filter & " AND IDAnagraficaAgente=" & fnNotNullN(rsAnagrafica!IDAnagraficaAgente)
    
    If TipoRaggruppamento = 1 Then
        rsRateDaFatt.Filter = rsRateDaFatt.Filter & " AND IDAnagrafica=" & fnNotNullN(rsAnagrafica!IDAnagraficaContratto)
    End If
    
    If TipoRaggruppamento = 0 Then
        If StampaRifContratto = 0 Then
            rsRateDaFatt.Sort = "IDAnagrafica, CodiceArticoloRata"
        Else
            rsRateDaFatt.Sort = "IDRV_PORateContratto, IDRV_POContratto, IDAnagrafica, CodiceArticoloRata"
        End If
    Else
        If StampaRifContratto = 0 Then
            rsRateDaFatt.Sort = "CodiceArticoloRata"
        Else
            rsRateDaFatt.Sort = "IDRV_PORateContratto, IDRV_POContratto, CodiceArticoloRata"
        End If
    End If
    
    IDAnagraficaContratto = 0
    IDContratto = 0
    
    While Not rsRateDaFatt.EOF
        If IDAnagraficaContratto <> fnNotNullN(rsRateDaFatt!IDAnagrafica) Then
            IDAnagraficaContratto = fnNotNullN(rsRateDaFatt!IDAnagrafica)
            If fnNotNullN(rsRateDaFatt!IDAnagrafica) <> fnNotNullN(rsAnagrafica!IDAnagrafica) Then
                ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                With ObjDoc.Tables
                    .Field "Art_descrizione", GET_RIFERIMENTO_ANA_CONTRATTO(fnNotNullN(rsRateDaFatt!IDAnagrafica), fnNotNullN(rsRateDaFatt!IDSitoPerAnagrafica)), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End With
                I = I + 1
            End If
        End If
        
        If StampaRifContratto = 1 Then
            If IDContratto <> fnNotNullN(rsRateDaFatt!IDRV_POContratto) Then
                IDContratto = fnNotNullN(rsRateDaFatt!IDRV_POContratto)
                ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                With ObjDoc.Tables
                    .Field "Art_descrizione", "Riferimento contratto numero " & fnNotNullN(rsRateDaFatt!AnnoContratto) & "-" & fnNotNullN(rsRateDaFatt!NumeroContratto), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End With
                I = I + 1
            End If
        End If
        
        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I

        'IDArticolo'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If fnNotNullN(rsRateDaFatt!IDRV_POTipoImpostazioneContratto) <= 2 Then
            IDArticolo = fnNotNullN(rsRateDaFatt!IDArticoloRataContratto)
            IDIvaFatturazione = fnNotNullN(rsRateDaFatt!IDIvaArticoloRata)
            AliquotaIvaFatturazione = fnNotNullN(rsRateDaFatt!AliquotaIvaArticoloRata)
            
            If IDArticolo = 0 Then
                IDArticolo = fnNotNullN(rsRateDaFatt!IDArticoloContratto)
                IDIvaFatturazione = fnNotNullN(rsRateDaFatt!IDIvaVenditaArticoloContratto)
                AliquotaIvaFatturazione = fnNotNullN(rsRateDaFatt!AliquotaIvaArticoloContratto)
                If IDArticolo = 0 Then
                    IDArticolo = fnNotNullN(rsRateDaFatt!IDArticoloTipoContratto)
                    IDIvaFatturazione = fnNotNullN(rsRateDaFatt!IDIvaVenditaTipoContratto)
                    AliquotaIvaFatturazione = fnNotNullN(rsRateDaFatt!AliquotaIvaTipoContratto)
                
                    If IDArticolo = 0 Then
                        IDArticolo = Link_Articolo
                        IDIvaFatturazione = GET_IVA_ARTICOLO(IDArticolo)
                        AliquotaIvaFatturazione = GET_ALIQUOTA_IVA_ARTICOLO(IDIvaFatturazione)
                    End If
                
                End If
            End If
        Else
            IDArticolo = GET_LINK_ARTICOLO_DA_PROD_CONTR(rsRateDaFatt!IDRV_POContrattoProdotti, IDIvaFatturazione, AliquotaIvaFatturazione)
            If (IDIvaFatturazione = 0) Then
                IDIvaFatturazione = GET_IVA_ARTICOLO(IDArticolo)
                AliquotaIvaFatturazione = GET_ALIQUOTA_IVA_ARTICOLO(IDIvaFatturazione)
            End If
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        'DESCRIZIONE ARTICOLO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If fnNotNullN(rsRateDaFatt!IDRV_POTipoImpostazioneContratto) <= 2 Then
            DescrizioneArticolo = Trim(fnNotNull(rsRateDaFatt!Periodo))
            If Len(DescrizioneArticolo) = 0 Then
                DescrizioneArticolo = GET_DESCRIZIONE_ARTICOLO(IDArticolo)
            End If
        Else
            DescrizioneArticolo = GET_DESCRIZIONE_ARTICOLO(IDArticolo)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        'Conto contabile'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        IDContoContabile = fnNotNullN(rsRateDaFatt!IDContoPDCContratto)
        CodicePDC = fnNotNull(rsRateDaFatt!CodicePDCContratto)
        DescrizionePDC = fnNotNull(rsRateDaFatt!DescrizionePDCContratto)
        If IDContoContabile = 0 Then
            IDContoContabile = fnNotNullN(rsRateDaFatt!IDContoPDCTipoContratto)
            CodicePDC = fnNotNull(rsRateDaFatt!CodicePDCTipoContratto)
            DescrizionePDC = fnNotNull(rsRateDaFatt!DescrizionePDCTipoContratto)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
        With ObjDoc.Tables
            'ObjDoc.ReadDataFromAgent rsReg!IDAgente, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.ReadDataFromArticle IDArticolo
            .Field "Art_descrizione", DescrizioneArticolo, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            If fnNotNullN(rsRateDaFatt!IDRV_POTipoImpostazioneContratto) <= 2 Then
                .Field "Art_quantita_totale", fnNotNullN(rsRateDaFatt!Quantita), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_prezzo_unitario_neutro", fnNotNullN(rsRateDaFatt!ImportoRataUnitaria), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                .Field "Art_sco_in_percentuale_1", fnNotNullN(rsRateDaFatt!Sconto1), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_sco_in_percentuale_2", fnNotNullN(rsRateDaFatt!Sconto2), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            Else
                If fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti) > 0 Then
                    If fnNotNullN(rsRateDaFatt!IDRV_POProdotto) > 0 Then
                        ''SE IL NOLEGGIO HA UNA SOLA RATA, ALLORA PRENDERE I RIFERIMENTI DALLA RIGA PRODOTTI ALTRIMENTI CONTINUARE NELLO STESSO MODO
                        If (GET_CONTROLLO_UNA_RATA(fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti), fnNotNullN(rsRateDaFatt!IDRV_POContratto)) = True) Then
                            GET_DATI_DA_RIGA_CONTR fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti), True
                        Else
                            .Field "Art_quantita_totale", fnNotNullN(rsRateDaFatt!Quantita), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            .Field "Art_prezzo_unitario_neutro", fnNotNullN(rsRateDaFatt!ImportoRataUnitaria), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            
                            .Field "Art_sco_in_percentuale_1", fnNotNullN(rsRateDaFatt!Sconto1), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            .Field "Art_sco_in_percentuale_2", fnNotNullN(rsRateDaFatt!Sconto2), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            
                            GET_DATI_DA_RIGA_CONTR_PIU_RATE fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti), fnNotNullN(rsRateDaFatt!IDRV_POContratto), fnNotNullN(rsRateDaFatt!ImportoRataUnitaria)
                        End If
                    Else
                        If fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti) > 0 Then
                            GET_DATI_DA_RIGA_CONTR fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti)
                        Else
                            .Field "Art_quantita_totale", fnNotNullN(rsRateDaFatt!Quantita), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            .Field "Art_prezzo_unitario_neutro", fnNotNullN(rsRateDaFatt!ImportoRataUnitaria), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            
                            .Field "Art_sco_in_percentuale_1", fnNotNullN(rsRateDaFatt!Sconto1), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            .Field "Art_sco_in_percentuale_2", fnNotNullN(rsRateDaFatt!Sconto2), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        End If
                    End If
                Else
                    .Field "Art_quantita_totale", fnNotNullN(rsRateDaFatt!Quantita), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_prezzo_unitario_neutro", fnNotNullN(rsRateDaFatt!ImportoRataUnitaria), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                    .Field "Art_sco_in_percentuale_1", fnNotNullN(rsRateDaFatt!Sconto1), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_sco_in_percentuale_2", fnNotNullN(rsRateDaFatt!Sconto2), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End If
            End If
            
            'If fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
            If LINK_IVA_CLIENTE > 0 Then
                .Field "Link_art_IVA", LINK_IVA_CLIENTE, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA(LINK_IVA_CLIENTE), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            Else
                .Field "Link_Art_IVA", IDIvaFatturazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_aliquota_IVA", AliquotaIvaFatturazione, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            End If
            .Field "Link_art_magazzino", Link_Magazzino, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
            If IDContoContabile > 0 Then
                .Field "Link_Art_IDCContropartita", IDContoContabile, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            End If
            .Field "Art_data_inizio_competenza", IIf(fnNotNullN(rsRateDaFatt!DataInizioPeriodo) = 0, Date, rsRateDaFatt!DataInizioPeriodo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_data_fine_competenza", IIf(fnNotNullN(rsRateDaFatt!DataFinePeriodo) = 0, Date, rsRateDaFatt!DataFinePeriodo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)

            ''''''ANAGRAFICA AGENTE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = "SELECT Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome, Agente.Codice, RegolaProvvAgente.IDRegolaProvvAgente, "
            sSQL = sSQL & "RegolaProvvAgente.Predefinita , RegolaProvvAgente.IDRegolaProvv, RegolaProvv.RegolaProvv "
            sSQL = sSQL & "FROM Agente INNER JOIN "
            sSQL = sSQL & "Anagrafica ON Agente.IDAnagrafica = Anagrafica.IDAnagrafica INNER JOIN "
            sSQL = sSQL & "RegolaProvvAgente ON Anagrafica.IDAnagrafica = RegolaProvvAgente.IDAnagrafica INNER JOIN "
            sSQL = sSQL & "RegolaProvv ON RegolaProvvAgente.IDRegolaProvv = RegolaProvv.IDRegolaProvv "
            sSQL = sSQL & "WHERE (Anagrafica.IDAnagrafica =" & fnNotNullN(rsRateDaFatt!IDAnagraficaAgente) & ") "
            sSQL = sSQL & "AND (RegolaProvvAgente.Predefinita =" & fnNormBoolean(1) & ")"
            
            Set rsAgente = CnDMT.OpenResultset(sSQL)
            
            If rsAgente.EOF = False Then
                ObjDoc.Field "Art_age_ragione_sociale", fnNotNull(rsAgente!Anagrafica), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_age_codice", fnNotNull(rsAgente!Codice), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_age_nome", fnNotNull(rsAgente!Nome), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Link_Art_agente", fnNotNullN(rsAgente!IDAnagrafica), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Link_Art_age_regola_provv", fnNotNullN(rsAgente!IDRegolaProvv), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_age_regola_provv", fnNotNull(rsAgente!RegolaProvv), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            End If
            
            rsAgente.CloseResultset
            Set rsAgente = Nothing
            
        End With
                
        If fnNotNullN(rsRateDaFatt!IDRV_POTipoImpostazioneContratto) <= 2 Then
            If rsRateDaFatt!IDTipo = 1 Then
                If fnNotNullN(rsRateDaFatt!IDRV_POContrattoAdeguamento) = 0 Then
                    If rsRateDaFatt!NoteFattura <> "" Then
                        I = I + 1
                        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                        With ObjDoc.Tables
                            .Field "Art_descrizione", fnNotNull(rsRateDaFatt!NoteFattura), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        End With
                    
                    End If
                End If
            End If
        End If
        
'        If fnContrattoAltraFiliale = True Then
'            I = I + 1
'            ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
'            With ObjDoc.Tables
'                .Field "Art_descrizione", "Filiale: " & NomeAltraSede, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
'                .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
'                .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
'                .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
'                .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
'            End With
'        End If
    I = I + 1
    
    If Len(Trim(fnNotNull(rsRateDaFatt!Periodo))) > 0 Then
        If fnNotNullN(rsRateDaFatt!IDRV_POTipoImpostazioneContratto) = 3 Then
            If (ARTICOLO_PIU_PERIODO_FATT = 1) Then
                SplitPeriodo = Split(fnNotNull(rsRateDaFatt!Periodo), vbCrLf)
                For J = 0 To UBound(SplitPeriodo)
                    ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                    With ObjDoc.Tables
                        .Field "Art_descrizione", Mid(SplitPeriodo(J), 1, 255), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    End With
                I = I + 1
                Next
                
            End If
        End If
    End If
    
    If fnNotNullN(rsRateDaFatt!IDRV_POTipoImpostazioneContratto) = 3 Then
        'SE CI SONO LE ANNOTAZIONI INSERIRLE NELLA FATTURA NELLO STESSO MODO DEL PERIODO
        AnnotazioniStringa = GET_ANNOTAZIONI_DA_PRODOTTO(fnNotNullN(rsRateDaFatt!IDRV_POContrattoProdotti), fnNotNullN(rsRateDaFatt!IDRV_POContratto))
        If Len(AnnotazioniStringa) > 0 Then
            SplitPeriodo = Split(AnnotazioniStringa, vbCrLf)
            For J = 0 To UBound(SplitPeriodo)
                ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                With ObjDoc.Tables
                    .Field "Art_descrizione", Mid(SplitPeriodo(J), 1, 255), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End With
            I = I + 1
            Next
        End If
    End If
    
    
    rsRateDaFatt.MoveNext
    Wend

    'RIPORTA_DESCRIZIONI_AUTOMATICHE I, TheApp.Branch, ObjDoc.IDTipoOggetto
    
'fncRighe = True
Exit Function
'ERR_fncRighe:
'    fncRighe = False
'    VARErroreIDIntervento = "GENERALITA':" & vbCrLf & "IDCliente : " & IDClienteOP
'    'VARErroreIDArticolo = vbCrLf & "Articolo: " & rsArt!Articolo
'    VARErroreGenerico = Err.Description & vbCrLf & VARErroreIDIntervento ' & VARErroreIDArticolo


End Function
Private Sub AGGIORNA_RATE_CONTRATTI_PER_RAGGRUPPAMENTO()
Dim sSQL As String
Dim VarIDOggettoCollegato As Long
Dim IDOggettoScadenza As Long
Dim IDOggettoRata As Long
Dim IDTipoOggettoRata As Long

rsRateDaFatt.MoveFirst

While Not rsRateDaFatt.EOF
    If rsRateDaFatt!IDTipo = 1 Then
        If fnNotNullN(rsRateDaFatt!IDRV_PORateContratto) > 0 Then
            IDOggettoRata = fnNotNullN(rsRateDaFatt!IDOggetto)
            IDTipoOggettoRata = fnNotNullN(rsRateDaFatt!IDTipoOggetto)
            
            If IDTipoOggettoRata = 0 Then
                IDTipoOggettoRata = fnGetTipoOggetto("RV_PORateContratto")
            End If
            
            If IDOggettoRata = 0 Then
                IDOggettoRata = GET_LINK_OGGETTO(IDOggettoRata, IDTipoOggettoRata, rsRateDaFatt!NumeroRata, rsRateDaFatt!DataRata)
            End If
            
            sSQL = "UPDATE RV_PORateContratto SET"
            sSQL = sSQL & " IDOggettoCollegato=" & ObjDoc.IDOggetto & ", "
            sSQL = sSQL & " Fatturata=1, "
            sSQL = sSQL & " IDOggetto=" & IDOggettoRata & ", "
            sSQL = sSQL & " IDTipoOggetto=" & IDTipoOggettoRata
            sSQL = sSQL & " WHERE IDRV_PORateContratto=" & fnNotNullN(rsRateDaFatt!IDRV_PORateContratto)
            CnDMT.Execute sSQL
        End If
    End If
    If rsRateDaFatt!IDTipo = 2 Then
        If fnNotNullN(rsRateDaFatt!IDRV_POContatoreRilevamenti) > 0 Then
            IDOggettoRata = fnNotNullN(rsRateDaFatt!IDOggetto)
            IDTipoOggettoRata = fnNotNullN(rsRateDaFatt!IDTipoOggetto)
            
            sSQL = "UPDATE RV_POContatoreRilevamenti SET"
            sSQL = sSQL & " IDOggettoCollegato=" & ObjDoc.IDOggetto & ", "
            sSQL = sSQL & " IDTipoOggettoCollegato=" & ObjDoc.IDTipoOggetto & ", "
            sSQL = sSQL & " Fatturata=1 "
            sSQL = sSQL & " WHERE IDRV_POContatoreRilevamenti=" & fnNotNullN(rsRateDaFatt!IDRV_POContatoreRilevamenti)
            CnDMT.Execute sSQL
        End If
    End If
    
    
    
    If ((IDOggettoRata > 0) And (IDTipoOggettoRata > 0)) Then
        If rsRateDaFatt!IDTipo = 1 Then
            CREA_FLUSSO_DOCUMENTALE ObjDoc.IDTipoOggetto, ObjDoc.IDOggetto, IDOggettoRata, IDTipoOggettoRata, "Documento di vendita -> Rata contratto"
        End If
        If rsRateDaFatt!IDTipo = 2 Then
            CREA_FLUSSO_DOCUMENTALE ObjDoc.IDTipoOggetto, ObjDoc.IDOggetto, IDOggettoRata, IDTipoOggettoRata, "Documento di vendita -> Rilevamento"
        End If
        
        ''''''ELIMINAZIONE SCADENZA DEL CONTRATTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If (rsRateDaFatt!IDTipo = 1) Then
            IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(fnNotNullN(rsRateDaFatt!IDOggetto), fnNotNullN(rsRateDaFatt!IDTipoOggetto), 0)
            
            If IDOggettoScadenza > 0 Then
                ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, fnNotNullN(rsRateDaFatt!IDOggetto), fnNotNullN(rsRateDaFatt!IDTipoOggetto), "Rata contratto -> Scadenza"
                ELIMINA_SCADENZA IDOggettoScadenza
            
            End If
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If


rsRateDaFatt.MoveNext
Wend

End Sub
Private Function GET_IVA_ARTICOLO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIvaVendita FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_IVA_ARTICOLO = 0
Else
    GET_IVA_ARTICOLO = fnNotNullN(rs!IDIvaVendita)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ALIQUOTA_IVA_ARTICOLO(IDIva As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AliquotaIva FROM Iva "
sSQL = sSQL & "WHERE IDIva=" & IDIva

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ALIQUOTA_IVA_ARTICOLO = 0
Else
    GET_ALIQUOTA_IVA_ARTICOLO = fnNotNullN(rs!aliquotaIva)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_ARTICOLO(IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Articolo FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_ARTICOLO = ""
Else
    GET_DESCRIZIONE_ARTICOLO = fnNotNull(rs!Articolo)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREA_FLUSSO_DOCUMENTALE(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long, DescrizioneFunzione As String)
On Error GoTo ERR_CREA_FLUSSO_DOCUMENTALE
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim IDFunzioneVend As Long
Dim IDFunzioneRata As Long
Dim IDFlussoGruppo As Long
Dim IDFlussoFunzione As Long

IDFunzioneVend = GET_LINK_FUNZIONE(IDTipoOggettoVend)
IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)

If IDFunzioneVend = 0 Then Exit Sub
If IDFunzioneRata = 0 Then Exit Sub
'''''''''''''''''''''''''''''''''GRUPPO FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoGruppo "
sSQL = sSQL & "WHERE Descrizione=" & fnNormString(DescrizioneFunzione)
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFlussoGruppo")
        rsNew!Descrizione = DescrizioneFunzione
    rsNew.Update
End If

IDFlussoGruppo = fnNotNullN(rsNew!IDFlussoGruppo)

rsNew.Close
Set rsNew = Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoFunzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzioneVend
sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneRata
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoFunzione = fnGetNewKeyTipoOggetto("FlussoFunzione", "IDFlussoFunzione")
        rsNew!IDFunzione = IDFunzioneVend
        rsNew!IDFunzioneSuccessiva = IDFunzioneRata
        rsNew!Cardinalita = 3
        rsNew!TipoAutomatismo = 1
        rsNew!Attributo = 14
        rsNew!TipoDipendenza = 1
        rsNew!IDFlussoGruppo = IDFlussoGruppo
    rsNew.Update
End If

IDFlussoFunzione = fnNotNullN(rsNew!IDFlussoFunzione)

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE COLLEGATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoFunzioneCollegato "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoVend
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoFunzione = IDFlussoFunzione
        rsNew!IDOggetto = IDOggettoVend
        rsNew!IDTipoOggetto = IDTipoOggettoVend
End If

rsNew!FlussoFunzioneCollegato = 2
rsNew.Update

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FLUSSO OGGETTI COLLEGATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggetto=" & IDOggettoVend
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoRata

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
    rsNew!IDFlussoFunzione = IDFlussoFunzione
    rsNew!IDOggetto = IDOggettoVend
    rsNew!IDTipoOggetto = IDTipoOggettoVend
    rsNew!IDTipoOggettoCollegato = IDTipoOggettoRata
    rsNew!IDOggettoCollegato = IDOggettoRata
    rsNew.Update
End If

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Sub
ERR_CREA_FLUSSO_DOCUMENTALE:
MsgBox Err.Description, vbCritical, "CREA_FLUSSO_DOCUMENTALE"
End Sub

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
Private Function GET_RIFERIMENTO_ANA_CONTRATTO(IDAnagraficaContratto As Long, IDSitoPerAnagrafica As Long) As String
On Error GoTo ERR_GET_RIFERIMENTO_ANA_CONTRATTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaContratto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_RIFERIMENTO_ANA_CONTRATTO = "Riferimento: "
Else
    GET_RIFERIMENTO_ANA_CONTRATTO = "Riferimento: " & fnNotNull(rs!Anagrafica) & " " & rs!Nome
End If

rs.CloseResultset
Set rs = Nothing

If IDSitoPerAnagrafica = 0 Then Exit Function

sSQL = "SELECT * FROM SitoPerAnagrafica "
sSQL = sSQL & "WHERE IDSitoPerAnagrafica=" & IDSitoPerAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_RIFERIMENTO_ANA_CONTRATTO = ""
Else
    GET_RIFERIMENTO_ANA_CONTRATTO = " (Filiale: " & fnNotNull(rs!SitoPerAnagrafica) & ")"
End If

rs.CloseResultset
Set rs = Nothing


Exit Function
ERR_GET_RIFERIMENTO_ANA_CONTRATTO:
    GET_RIFERIMENTO_ANA_CONTRATTO = "Riferimento: "
End Function
Public Function fnGetNewKeyTipoOggetto(Tabella As String, CampoKey As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

    
    sSQL = "SELECT MAX(" & CampoKey & ") AS NumeroRecord "
    sSQL = sSQL & " FROM " & Tabella
    sSQL = sSQL & " WHERE " & CampoKey & ">=10000"
    
    Set rs = CnDMT.OpenResultset(sSQL)

    If rs.EOF = True Then
    
        fnGetNewKeyTipoOggetto = 10000

    Else
        If fnNotNullN(rs.adoColumns("NumeroRecord").Value) = 0 Then
            fnGetNewKeyTipoOggetto = 10000
        Else
            fnGetNewKeyTipoOggetto = fnNotNullN(rs.adoColumns("NumeroRecord").Value) + 1
        End If
    End If

    rs.CloseResultset
    Set rs = Nothing
End Function

Private Function GET_LINK_OGGETTO_SCADENZA_COLLEGATA(IDOggettoRata As Long, IDTipoOggettoRata As Long, IDSezionale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDOggettoScadenza As Long
Dim IDFunzioneRata As Long
Dim IDFunzioneScadenza As Long
Dim IDFlussoFunzione As Long

IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)
IDFunzioneScadenza = GET_LINK_FUNZIONE(131)
IDFlussoFunzione = GET_LINK_FLUSSO_FUNZIONE(IDFunzioneRata, IDFunzioneScadenza)

If IDFlussoFunzione = 0 Then
    GET_LINK_OGGETTO_SCADENZA_COLLEGATA = 0
    Exit Function
End If

sSQL = "SELECT IDOggettoCollegato "
sSQL = sSQL & "FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=131"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_OGGETTO_SCADENZA_COLLEGATA = 0
Else
    GET_LINK_OGGETTO_SCADENZA_COLLEGATA = fnNotNullN(rs!IDOggettoCollegato)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_FLUSSO_FUNZIONE(IDFunzione As Long, IDFunzioneSucc As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDFlussoFunzione FROM FlussoFunzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzione
sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneSucc

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_FLUSSO_FUNZIONE = 0
Else
    GET_LINK_FLUSSO_FUNZIONE = fnNotNullN(rs!IDFlussoFunzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long, DescrizioneFunzione As String)
On Error GoTo ERR_ELIMINA_FLUSSO_DOCUMENTALE
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim IDFunzioneVend As Long
Dim IDFunzioneRata As Long
Dim IDFlussoGruppo As Long
Dim IDFlussoFunzione As Long

IDFunzioneVend = GET_LINK_FUNZIONE(IDTipoOggettoVend)
IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)

'''''''''''''''''''''''''''''''''GRUPPO FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoGruppo "
sSQL = sSQL & "WHERE Descrizione=" & fnNormString(DescrizioneFunzione)
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFLussoGruppo")
        rsNew!Descrizione = DescrizioneFunzione
    rsNew.Update
End If

IDFlussoGruppo = fnNotNullN(rsNew!IDFlussoGruppo)

rsNew.Close
Set rsNew = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IDFunzioneVend > 0 Then
    sSQL = "SELECT * FROM FlussoFunzione "
    sSQL = sSQL & "WHERE IDFunzione=" & IDFunzioneRata
    sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneVend
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDFlussoFunzione = fnGetNewKeyTipoOggetto("FlussoFunzione", "IDFlussoFunzione")
            rsNew!IDFunzione = IDFunzioneVend
            rsNew!IDFunzioneSuccessiva = IDFunzioneRata
            rsNew!Cardinalita = 3
            rsNew!TipoAutomatismo = 1
            rsNew!Attributo = 14
            rsNew!TipoDipendenza = 1
            rsNew!IDFlussoGruppo = IDFlussoGruppo
        rsNew.Update
    End If
    
    IDFlussoFunzione = fnNotNullN(rsNew!IDFlussoFunzione)
    
    rsNew.Close
    Set rsNew = Nothing
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE COLLEGATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggettoCollegato<>" & IDOggettoVend

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    sSQL = "DELETE FROM FlussoFunzioneCollegato "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoRata
    sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
    sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
    CnDMT.Execute sSQL
End If
rsNew.Close
Set rsNew = Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''FLUSSO OGGETTI COLLEGATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoVend
CnDMT.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Sub
ERR_ELIMINA_FLUSSO_DOCUMENTALE:
    MsgBox Err.Description, vbCritical, "ELIMINA_FLUSSO_DOCUMENTALE"
    
End Sub
Private Sub CREA_FLUSSO_DOCUMENTALE_SCADENZA(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long)
On Error GoTo ERR_CREA_FLUSSO_DOCUMENTALE_SCADENZA
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim IDFunzioneVend As Long
Dim IDFunzioneRata As Long
Dim IDFlussoGruppo As Long
Dim IDFlussoFunzione As Long

IDFunzioneVend = GET_LINK_FUNZIONE(IDTipoOggettoVend)
IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)

If IDFunzioneVend = 0 Then Exit Sub
If IDFunzioneRata = 0 Then Exit Sub
'''''''''''''''''''''''''''''''''GRUPPO FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoGruppo "
sSQL = sSQL & "WHERE Descrizione=" & fnNormString("Rata contratto -> Scadenza")
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFlussoGruppo")
        rsNew!Descrizione = "Rata contratto -> Scadenza"
    rsNew.Update
End If

IDFlussoGruppo = fnNotNullN(rsNew!IDFlussoGruppo)

rsNew.Close
Set rsNew = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoFunzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzioneRata
sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneVend
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoFunzione = fnGetNewKeyTipoOggetto("FlussoFunzione", "IDFlussoFunzione")
        rsNew!IDFunzione = IDFunzioneRata
        rsNew!IDFunzioneSuccessiva = IDFunzioneVend
        rsNew!Cardinalita = 3
        rsNew!TipoAutomatismo = 1
        rsNew!Attributo = 14
        rsNew!TipoDipendenza = 1
        rsNew!IDFlussoGruppo = IDFlussoGruppo
    rsNew.Update
End If

IDFlussoFunzione = fnNotNullN(rsNew!IDFlussoFunzione)

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE COLLEGATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoFunzioneCollegato "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoFunzione = IDFlussoFunzione
        rsNew!IDOggetto = IDOggettoRata
        rsNew!IDTipoOggetto = IDTipoOggettoRata
End If

rsNew!FlussoFunzioneCollegato = 2
rsNew.Update

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FLUSSO OGGETTI COLLEGATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoVend

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
    rsNew!IDFlussoFunzione = IDFlussoFunzione
    rsNew!IDOggetto = IDOggettoRata
    rsNew!IDTipoOggetto = IDTipoOggettoRata
    rsNew!IDTipoOggettoCollegato = IDTipoOggettoVend
    rsNew!IDOggettoCollegato = IDOggettoVend
    rsNew.Update
End If

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Sub
ERR_CREA_FLUSSO_DOCUMENTALE_SCADENZA:
MsgBox Err.Description, vbCritical, "CREA_FLUSSO_DOCUMENTALE_SCADENZA"
End Sub

Private Sub ELIMINA_SCADENZA(IDOggettoScadenza As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDTestataScadenza As Long


sSQL = "SELECT IDTestataScadenza FROM TestataScadenza "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoScadenza

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    IDTestataScadenza = 0
Else
    IDTestataScadenza = fnNotNullN(rs!IDTestataScadenza)
End If


rs.CloseResultset
Set rs = Nothing

If IDTestataScadenza > 0 Then
    sSQL = "DELETE FROM TestataScadenza "
    sSQL = sSQL & "WHERE IDTestataScadenza=" & IDTestataScadenza
    CnDMT.Execute sSQL

    sSQL = "DELETE FROM DettaglioScadenza "
    sSQL = sSQL & "WHERE IDTestataScadenza=" & IDTestataScadenza
    CnDMT.Execute sSQL

    sSQL = "DELETE FROM Oggetto "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoScadenza
    CnDMT.Execute sSQL
    
End If

End Sub

Private Function GET_LINK_SCADENZA(ImportoComplessivoScadenza As Double, IDAnagrafica As Long, NumeroDocumento As Long, DataDocumento As String, IDSezionale As Long, Periodo As String) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim Link_Oggetto As Long
Dim IDTestataScadenza As Long


Link_Oggetto = GET_LINK_OGGETTO_SCADENZA(DataDocumento, IDSezionale, NumeroDocumento)

If Link_Oggetto > 0 Then
        
    Set rs = New ADODB.Recordset
    sSQL = "SELECT * FROM TestataScadenza "
    sSQL = sSQL & "WHERE IDOggetto=" & Link_Oggetto
    
    rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    rs.AddNew
        rs!IDTestataScadenza = fnGetNewKey("TestataScadenza", "IDTestataScadenza")
        rs!IDOggetto = Link_Oggetto
        rs!IDTipoOggetto = 131
        rs!IDFiliale = TheApp.Branch
        rs!IDNaturaScadenza = 6
        rs!IDAnagrafica_CF = IDAnagrafica
        rs!IDTipoAnagrafica_CF = 2
        rs!IDAzienda_CF = TheApp.IDFirm
        rs!IDAzienda = TheApp.IDFirm
        rs!IDPagamento = 52
        rs!IDRegistroIva = GET_LINK_REGISTRO_IVA(IDSezionale)
        rs!IDTipoStatoScadenza = 2
        rs!IDSezionale = IDSezionale
        rs!NumeroDocumento = NumeroDocumento
        rs!DataDocumento = DataDocumento
        rs!DataInizioScadenza = DataDocumento
        rs!ImportoComplessivo = ImportoComplessivoScadenza
        rs!ImportoIva = 0
        rs!ScadenzaAttivaPassiva = 0
        rs!GiornoScadenzaFissa = 0
        rs!IvaSuScadenza = 2
        rs!NumeroRataIVA = 0
        rs!NonRipartireIva = False
        rs!DataUltimaVariazione = Date
        rs!IDUtenteUltimaVariazione = TheApp.IDUser
        rs!VirtualDelete = 0
        
        IDTestataScadenza = fnNotNullN(rs!IDTestataScadenza)
    rs.Update

    rs.Close
    Set rs = Nothing
    If IDTestataScadenza > 0 Then
        GENERA_DETTAGLIO_SCADENZA IDTestataScadenza, ImportoComplessivoScadenza, NumeroDocumento, DataDocumento, Periodo
    End If
    
    
    GET_LINK_SCADENZA = Link_Oggetto
End If
End Function

Private Function GET_LINK_OGGETTO_SCADENZA(DataDocumento As String, IDSezionale As Long, NumeroDocumento As Long) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

GET_LINK_OGGETTO_SCADENZA = 0

Set rs = New ADODB.Recordset

rs.Open "Oggetto", CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    rs!IDOggetto = fnGetNewKey("Oggetto", "IDOggetto")
    rs!IDTipoOggetto = 131
    rs!IDFunzione = 88
    rs!Oggetto = "Testata Scadenza del " & DataDocumento
    rs!IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.Branch)
    rs!IDAzienda = TheApp.IDFirm
    rs!IDSezionale = IDSezionale
    rs!DataEmissione = DataDocumento
    rs!Numero = NumeroDocumento
    rs!DataUltimaVariazione = Date
    rs!IDUtenteUltimaVariazione = TheApp.IDUser
    rs!VirtualDelete = 0
    
rs.Update

GET_LINK_OGGETTO_SCADENZA = rs!IDOggetto

rs.Close
Set rs = Nothing
End Function

Private Function GENERA_DETTAGLIO_SCADENZA(IDTestataScadenza As Long, ImportoComplessivo As Double, NumeroDocumento As Long, DataDocumento As String, Periodo As String) As Double
Dim rs As ADODB.Recordset
Dim rsNew As ADODB.Recordset
Dim sSQL As String
Dim IRata As Integer
Dim RIMANENZA As Double
Dim IDTipoStatoScadenza As Long


sSQL = "SELECT * FROM DettaglioScadenza "
sSQL = sSQL & "WHERE IDTestataScadenza=" & IDTestataScadenza

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

rsNew.AddNew
    rsNew!IDTestataScadenza = IDTestataScadenza
    rsNew!IDDettaglioScadenza = fnGetNewKey("DettaglioScadenza", "IDDettaglioScadenza")
    rsNew!DataScadenza = DataDocumento
    rsNew!ImportoScadenza = ImportoComplessivo
    rsNew!IDTipoStatoScadenza = 2
    rsNew!IDTipoPosizioneScadenza = 3
    rsNew!IDTipoOggetto = 0
    rsNew!IDTipoPagamento = 1
    rsNew!RaggruppamentoScadenza = False
    rsNew!RIBA = False
    rsNew!TrasferitoOutlook = False
    rsNew!Contabilizzata = False
    rsNew!Note = Periodo
    rsNew!NumeroRata = NumeroDocumento
    rsNew!DataUltimaVariazione = Date
    rsNew!IDUtenteUltimaVariazione = TheApp.IDUser
    rsNew!VirtualDelete = 0
rsNew.Update

End Function
Private Function GET_LINK_REGISTRO_IVA(IDSezionale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRegistroIva FROM Sezionale "
sSQL = sSQL & "WHERE IDSezionale=" & IDSezionale

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_REGISTRO_IVA = 1
Else
    GET_LINK_REGISTRO_IVA = fnNotNullN(rs!IDRegistroIva)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ELIMINA_FLUSSO_RATE_DA_DOCUMENTO(IDContratto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDOggettoScadenza As Long

sSQL = "SELECT * FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), 0)
    If IDOggettoScadenza > 0 Then
        ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), "Rata contratto -> Scadenza"
        ELIMINA_SCADENZA IDOggettoScadenza
    End If
rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing
End Sub

Private Function GET_CONTROLLO_NUMERO_LETTERE_INTENTO(IDAnagrafica As Long, IDAzienda As Long, Anno As Long) As Long
On Error GoTo ERR_GET_CONTROLLO_NUMERO_LETTERE_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Count(IDLetteraIntento) AS NumeroRecord "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDAzienda_CF=" & IDAzienda
sSQL = sSQL & " AND IDAnagrafica_CF=" & IDAnagrafica
sSQL = sSQL & " AND IDTipoAnagrafica_CF=2"
sSQL = sSQL & " AND ((Anno=" & Anno & ") OR (AnnoOperazione=" & Anno & "))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_NUMERO_LETTERE_INTENTO = 0
Else
    GET_CONTROLLO_NUMERO_LETTERE_INTENTO = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_CONTROLLO_NUMERO_LETTERE_INTENTO:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_NUMERO_LETTERE_INTENTO"
End Function
Private Function GET_LINK_LETTERA_INTENTO(IDAnagrafica As Long, IDAzienda As Long, Anno As Long) As Long
On Error GoTo ERR_GET_LINK_LETTERA_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDLetteraIntento "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDAzienda_CF=" & IDAzienda
sSQL = sSQL & " AND IDAnagrafica_CF=" & IDAnagrafica
sSQL = sSQL & " AND IDTipoAnagrafica_CF=2"
sSQL = sSQL & " AND ((Anno=" & Anno & ") OR (AnnoOperazione=" & Anno & "))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_LETTERA_INTENTO = 0
Else
    GET_LINK_LETTERA_INTENTO = fnNotNullN(rs!IDLetteraIntento)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_LINK_LETTERA_INTENTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_LETTERA_INTENTO"
End Function
Private Function GET_LINK_IVA_LETTERA_INTENTO(IDLetteraIntento As Long, IDIvaCliente As Long) As Long
On Error GoTo ERR_GET_LINK_IVA_LETTERA_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIva "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDLetteraIntento=" & IDLetteraIntento

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_LETTERA_INTENTO = IDIvaCliente
Else
    If fnNotNullN(rs!IDIva) > 0 Then
        GET_LINK_IVA_LETTERA_INTENTO = fnNotNullN(rs!IDIva)
    Else
        GET_LINK_IVA_LETTERA_INTENTO = IDIvaCliente
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_LINK_IVA_LETTERA_INTENTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_IVA_LETTERA_INTENTO"
End Function

Private Function GET_LINK_IVA_CLIENTE(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIva "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_CLIENTE = 0
Else

    GET_LINK_IVA_CLIENTE = fnNotNullN(rs!IDIva)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function

End Function
Private Function GET_ALIQUOTA_IVA(IDIva As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AliquotaIva FROM Iva "
sSQL = sSQL & "WHERE IDIva=" & IDIva

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ALIQUOTA_IVA = 0
Else
    GET_ALIQUOTA_IVA = fnNotNullN(rs!aliquotaIva)
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
Private Sub RIPORTA_DESCRIZIONI_AUTOMATICHE(IRiga As Long, IDFiliale As Long, IDTipoOggetto As Long)
On Error GoTo ERR_RIPORTA_DESCRIZIONI_AUTOMATICHE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim I As Long

sSQL = "SELECT DescrizioneAggiuntivaPerFiliale.IDDescrizioneAggiuntivaPerFiliale, DescrizioneAggiuntivaPerFiliale.IDFiliale, DescrizioneAggiuntivaPerFiliale.IDTestataDescrizioneAggiuntiva, "
sSQL = sSQL & "DescrizioneAggiuntivaPerFiliale.IDTipoOggetto, DescrizioneAggiuntivaPerFiliale.DataInizio, DescrizioneAggiuntivaPerFiliale.DataFine, DescrizioneAggiuntivaPerFiliale.NonRiportoFattDiff,"
sSQL = sSQL & "DescrizioneAggiuntivaPerFiliale.NonRiportoDocumenti, DescrizioneAggiuntivaPerFiliale.Sequenza, TestataDescrizioneAggiuntiva.CodiceDescrizione,"
sSQL = sSQL & "TestataDescrizioneAggiuntiva.DescrizioneRidotta "
sSQL = sSQL & "FROM DescrizioneAggiuntivaPerFiliale INNER JOIN "
sSQL = sSQL & "TestataDescrizioneAggiuntiva ON DescrizioneAggiuntivaPerFiliale.IDTestataDescrizioneAggiuntiva = TestataDescrizioneAggiuntiva.IDTestataDescrizioneAggiuntiva "
sSQL = sSQL & " WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND NonRiportoDocumenti=" & fnNormBoolean(False)
sSQL = sSQL & " AND DataInizio<=" & fnNormDate(ObjDoc.DataEmissione)
sSQL = sSQL & " AND DataFine>=" & fnNormDate(ObjDoc.DataEmissione)
sSQL = sSQL & " ORDER BY Sequenza "

Set rs = CnDMT.OpenResultset(sSQL)
I = IRiga

While Not rs.EOF
    ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
    With ObjDoc.Tables
        .Field "Art_descrizione", fnNotNull(rs!DescrizioneRidotta), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Link_Art_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        .Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    End With
    I = I + 1
    
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_RIPORTA_DESCRIZIONI_AUTOMATICHE:
    
End Sub
Private Function GET_LINK_IVA_CLIENTE_ESENTE(IDAnagrafica As Long, IDIva As Long, DataDocumento As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_IVA_CLIENTE_ESENTE = 0

sSQL = "SELECT IDIva, DataEsenzioneDa, DataEsenzioneA "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_CLIENTE_ESENTE = 0
Else
    If (fnNotNullN(rs!DataEsenzioneDa) = 0) Then
        GET_LINK_IVA_CLIENTE_ESENTE = fnNotNullN(rs!IDIva)
    Else
        If ((DateDiff("d", rs!DataEsenzioneDa, DataDocumento) >= 0) And (DateDiff("d", DataDocumento, rs!DataEsenzioneA) >= 0)) Then
            GET_LINK_IVA_CLIENTE_ESENTE = fnNotNullN(rs!IDIva)
        End If
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
End Function

Private Function GET_LINK_ARTICOLO_DA_PROD_CONTR(IDContrattoProdotti, idIvaFatt As Long, aliquotaIva As Double) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_ARTICOLO_DA_PROD_CONTR = 0

sSQL = "SELECT IDArticolo, IDIva "
sSQL = sSQL & "FROM RV_POContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDContrattoProdotti

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_ARTICOLO_DA_PROD_CONTR = fnNotNullN(rs!IDArticolo)
    idIvaFatt = fnNotNullN(rs!IDIva)
    aliquotaIva = GET_ALIQUOTA_IVA(idIvaFatt)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_DATI_DA_RIGA_CONTR(IDContrattoProdotti As Long, Optional DaProdotto As Boolean = False)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDContrattoProdotti

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    If (DaProdotto = False) Then
        ObjDoc.Field "Art_quantita_totale", fnNotNullN(rs!QuantitaArticolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    Else
        If (fnNotNullN(rs!ACorpo) = 0) Then
            ObjDoc.Field "Art_quantita_totale", fnNotNullN(rs!QuantitaEffettiva) * fnNotNullN(rs!QuantitaArticolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        Else
            ObjDoc.Field "Art_quantita_totale", fnNotNullN(rs!QuantitaArticolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        End If
        'ObjDoc.Field "Art_quantita_totale", fnNotNullN(rs!QuantitaPeriodo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    End If
    
    ObjDoc.Field "Art_prezzo_unitario_neutro", fnNotNullN(rs!ImportoUnitario), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    
    ObjDoc.Field "Art_sco_in_percentuale_1", fnNotNullN(rs!Sconto1), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Art_sco_in_percentuale_2", fnNotNullN(rs!Sconto2), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Link_Art_unita_di_misura", fnNotNullN(rs!IDUnitaDiMisuraArticolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Art_sigla_unita_di_misura", GET_DESCRIZIONE_UM(fnNotNullN(rs!IDUnitaDiMisuraArticolo)), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    
    
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_DATI_DA_RIGA_CONTR_PIU_RATE(IDContrattoProdotti As Long, IDContratto As Long, ImportoRata As Double)
On Error GoTo ERR_GET_DATI_DA_RIGA_CONTR_PIU_RATE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoUnitario As Double
Dim QuantitaDaFatturare As Double
Dim NumeroRate As Long
Dim IDUnitaDiMisura As Long

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDContrattoProdotti

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    NumeroRate = GET_NUMERO_RATE_PRODOTTO_CONTRATTO(IDContrattoProdotti, IDContratto)
    ImportoUnitario = ImportoRata / (fnNotNullN(rs!QuantitaEffettiva) * fnNotNullN(rs!QuantitaArticolo))
    QuantitaDaFatturare = (fnNotNullN(rs!QuantitaEffettiva) * fnNotNullN(rs!QuantitaArticolo)) / NumeroRate
    
    If ((QuantitaDaFatturare * fnNotNullN(rs!ImportoUnitario)) = ImportoRata) Then
        ObjDoc.Field "Art_quantita_totale", QuantitaDaFatturare, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_prezzo_unitario_neutro", fnNotNullN(rs!ImportoUnitario), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_sco_in_percentuale_1", fnNotNullN(rs!Sconto1), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_sco_in_percentuale_2", fnNotNullN(rs!Sconto2), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Link_Art_unita_di_misura", fnNotNullN(rs!IDUnitaDiMisuraArticolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_sigla_unita_di_misura", GET_DESCRIZIONE_UM(fnNotNullN(rs!IDUnitaDiMisuraArticolo)), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    Else
        IDUnitaDiMisura = CREA_UM_PERIODO
        If (IDUnitaDiMisura > 0) Then
            ObjDoc.Field "Link_Art_unita_di_misura", IDUnitaDiMisura, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_sigla_unita_di_misura", GET_DESCRIZIONE_UM(IDUnitaDiMisura), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        End If
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_GET_DATI_DA_RIGA_CONTR_PIU_RATE:

    MsgBox Err.Description, vbCritical, "GET_DATI_DA_RIGA_CONTR_PIU_RATE"

End Sub
Private Function GET_DESCRIZIONE_UM(IDUnitaDiMisura As Long) As String
On Error GoTo ERR_GET_DESCRIZIONE_UM
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoUnitario As Double

GET_DESCRIZIONE_UM = ""

sSQL = "SELECT * "
sSQL = sSQL & "FROM UnitaDiMisura "
sSQL = sSQL & "WHERE IDUnitaDiMisura=" & IDUnitaDiMisura

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_DESCRIZIONE_UM = fnNotNull(rs!DescrizioneFattura)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_DESCRIZIONE_UM:

End Function
Public Function TrovaCartella(IDLCartella As Long) As String

    TrovaCartella = String$(MAX_PATH, 0)
    
    Call SHGetSpecialFolderPath(ByVal 0&, TrovaCartella, IDLCartella, ByVal 0&)
    
    TrovaCartella = Left$(TrovaCartella, InStr(1, TrovaCartella, Chr$(0)) - 1)
    
    If Len(TrovaCartella) > 0 And Right$(TrovaCartella, 1) <> "\" Then TrovaCartella = TrovaCartella & "\"
End Function
Private Function GET_CONTROLLO_UNA_RATA(IDContrattoProdotti As Long, IDContratto As Long) As Boolean
On Error GoTo ERR_GET_CONTROLLO_UNA_RATA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_UNA_RATA = False

sSQL = "SELECT COUNT(IDRV_PORateContratto) AS NumeroRate "
sSQL = sSQL & "FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND IDRV_POContrattoProdotti=" & IDContrattoProdotti

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    If (fnNotNullN(rs!NumeroRate) <= 1) Then
        GET_CONTROLLO_UNA_RATA = True
    Else
        GET_CONTROLLO_UNA_RATA = False
    End If
Else
    GET_CONTROLLO_UNA_RATA = True
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_CONTROLLO_UNA_RATA:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_UNA_RATA"
End Function
Private Function GET_ANNOTAZIONI_DA_PRODOTTO(IDContrattoProdotti As Long, IDContratto As Long) As String
On Error GoTo ERR_GET_ANNOTAZIONI_DA_PRODOTTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_ANNOTAZIONI_DA_PRODOTTO = ""

sSQL = "SELECT IDRV_POContrattoProdotti, IDRV_POContratto, Annotazioni "
sSQL = sSQL & " FROM RV_POContrattoProdotti "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND IDRV_POContrattoProdotti=" & IDContrattoProdotti

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_ANNOTAZIONI_DA_PRODOTTO = Trim(fnNotNull(rs!Annotazioni))
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_ANNOTAZIONI_DA_PRODOTTO:
    MsgBox Err.Description, vbCritical, "GET_ANNOTAZIONI_DA_PRODOTTO"
End Function

Private Function GET_NUMERO_RATE_PRODOTTO_CONTRATTO(IDContrattoProdotti As Long, IDContratto As Long) As Long
On Error GoTo ERR_GET_CONTROLLO_UNA_RATA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_RATE_PRODOTTO_CONTRATTO = 1

sSQL = "SELECT COUNT(IDRV_PORateContratto) AS NumeroRate "
sSQL = sSQL & "FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND IDRV_POContrattoProdotti=" & IDContrattoProdotti

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!NumeroRate) > 0 Then
        GET_NUMERO_RATE_PRODOTTO_CONTRATTO = fnNotNullN(rs!NumeroRate)
    End If
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_CONTROLLO_UNA_RATA:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_UNA_RATA"
End Function
Private Function CREA_UM_PERIODO() As Long
On Error GoTo ERR_CREA_UM_PERIODO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim ID As Long

    CREA_UM_PERIODO = 0
    
    sSQL = "SELECT * FROM UnitaDiMisura "
    sSQL = sSQL & "WHERE DescrizioneFattura=" & fnNormString("Periodo")
    
    Set rs = New ADODB.Recordset
    rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rs.EOF Then
        rs.AddNew
        rs!IDUnitaDiMisura = fnGetNewKey("UnitaDiMisura", "IDUnitaDiMisura")
    End If
    ID = fnNotNullN(rs!IDUnitaDiMisura)
    rs!UnitaDiMisura = "Periodo"
    rs!DescrizioneFattura = "Rata"
    rs.Update
    
    CREA_UM_PERIODO = ID
    
rs.Close
Set rs = Nothing
Exit Function
ERR_CREA_UM_PERIODO:

End Function

Private Function GET_PREFISSO_SEZ(IDSezionale As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Prefisso FROM Sezionale "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDSezionale=" & IDSezionale

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREFISSO_SEZ = ""
Else
    GET_PREFISSO_SEZ = Trim(fnNotNull(rs!Prefisso))
End If

rs.CloseResultset
Set rs = Nothing
End Function
