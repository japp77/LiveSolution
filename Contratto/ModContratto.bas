Attribute VB_Name = "ModContratto"
Public Link_Contratto As Long
Public Link_StoriaContratto As Long
Public Link_Contratto_padre As Long

'VARIABILI DEL TIPO CONTRATTO
Public Descrizione_Tipo_Contratto As String

'VARIABILI DELLA DURATA CONTRATTO
Public Mesi_Durata_Contratto As Long
Public Giorni_Durata_Contratto As Long

Public Mesi_Durata_Contratto_prox As Long
Public Giorni_Durata_Contratto_prox As Long

'VARIABILI DEL TIPO DI RINNOVO
Public Mesi_Rinnovo_Contratto As Long 'Indica ogni quanti mesi il contratto deve essere rinnovato
Public Giorni_Rinnovo_Contratto As Long 'Indica il giorno esatto del mese
Public AnnoPrecedente_Rinnovo_Contratto As Long 'Indica se la data di decorrenza deve concidere con quella dell'anno precedente

Public Mesi_Rinnovo_Contratto_prox As Long 'Indica ogni quanti mesi il contratto deve essere rinnovato
Public Giorni_Rinnovo_Contratto_prox As Long 'Indica il giorno esatto del mese
Public AnnoPrecedente_Rinnovo_Contratto_prox As Long 'Indica se la data di decorrenza deve concidere con quella dell'anno precedente


'VARIABILI DEL TIPO DI RATEIZZAZIONE
Public Numero_Rate As Long 'Indica il numero di rate del contratto
Public Pagamento_Anticipato_Periodo As Boolean 'Indica se la rata deve essere pagata ad inizio periodo
Public Mesi_Rate As Long 'Indica ogni quanti mesi devono essere create le rate
Public Rata_Iniziale As Long
Public Anno_Solare As Long


'VARIABILI DELLA DURATA ASSISTENZA DEL CONTRATTO
Public Mesi_Durata_Assistenza As Long
Public Giorni_Durata_Assistenza As Long

Public Link_Contratto_Servizio As Long
Public Link_Contratto_Prodotto As Long
Public Elabora_Tutti_Servizi_Contratto As Long
Public Note_Prodotto_Intervento As String



Public rsRateEla As ADODB.Recordset
Public ImportoRataTotale As Double
Public Sub SviluppoRateContratto(IDContratto As Long, Optional IDAdeguamento As Long)
Dim sSQL As String
Dim rsDelete As DmtOleDbLib.adoResultset
Dim IDOggettoScadenza As Long
Dim rsAdeg As DmtOleDbLib.adoResultset

    If (IDAdeguamento > 0) Then
    
        'Dim PagamentoAnticipato As Boolean
        sSQL = "SELECT * FROM RV_PORateContratto "
        sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
        sSQL = sSQL & " AND Fatturata=0 "
        sSQL = sSQL & " AND IDRV_POContrattoAdeguamento=" & IDAdeguamento
        
        Set rsDelete = Cn.OpenResultset(sSQL)
        
        While Not rsDelete.EOF
            sSQL = "DELETE FROM RV_PORateContratto "
            sSQL = sSQL & " WHERE IDRV_PORateContratto=" & fnNotNullN(rsDelete!IDRV_PORateContratto)
'            sSQL = sSQL & " AND IDRV_POContrattoAdeguamento=" & IDAdeguamentoContratto
'                sSQL = sSQL & " AND Adeguamento=0"
'                sSQL = sSQL & " AND Manuale=0"
'                sSQL = sSQL & " AND ContrattoAttuale=1"
'                sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"
            Cn.Execute sSQL
            
            IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(fnNotNullN(rsDelete!IDOggetto), fnNotNullN(rsDelete!IDTipoOggetto), 0)
            
            If IDOggettoScadenza > 0 Then
                ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, fnNotNullN(rsDelete!IDOggetto), fnNotNullN(rsDelete!IDTipoOggetto)
                ELIMINA_SCADENZA IDOggettoScadenza
            End If
            
        rsDelete.MoveNext
        Wend
    
        sSQL = "SELECT * FROM RV_POIEAdeguamentiContratto "
        sSQL = sSQL & " WHERE IDRV_POContrattoAdeguamento=" & IDAdeguamento
        sSQL = sSQL & " AND IDRV_POContratto=" & IDContratto
        
        Set rsAdeg = Cn.OpenResultset(sSQL)
        
        If Not rsAdeg.EOF Then
            ElaborazioneRate 1, TheApp.IDFirm, IDContratto, fnNotNull(rsAdeg!DataDecorrenza), fnNotNullN(rsAdeg!Importo), IIf(fnNotNullN(rsAdeg!IDRateizzazione) = 0, Mesi_Rate, fnNotNullN(rsAdeg!Mesi)), IIf(fnNotNullN(rsAdeg!IDRateizzazione) = 0, Numero_Rate, fnNotNullN(rsAdeg!numerorate)), IIf(fnNotNullN(rsAdeg!IDRateizzazione) = 0, Pagamento_Anticipato_Periodo _
                            , fnNotNullN(rsAdeg!PagamentoInizioPeriodo)), frmMain.cboPagamentoRataContratto.CurrentID, frmMain.txtIDContrattoPadre.Value, frmMain.cboTipoContratto.CurrentID, frmMain.cboSitoPerAnagrafica.CurrentID, IIf(fnNotNullN(rsAdeg!IDRateizzazione) = 0, Rata_Iniziale, fnNotNullN(rsAdeg!RataInizialeRataFinale)), IIf(fnNotNullN(rsAdeg!IDRateizzazione) = 0, Anno_Solare, fnNotNullN(rsAdeg!AnnoSolare)), _
                            IIf((Len(Trim(fnNotNull(rsAdeg!DataFineAdeguamento))) = 0), frmMain.txtDataScadenzaPerRinnovo.Text, rsAdeg!DataFineAdeguamento), frmMain.txtNGGPrimaRata.Value, IIf(fnNotNullN(rsAdeg!IDRateizzazione) = 0, frmMain.cboTipoRateizzazione.CurrentID, fnNotNullN(rsAdeg!IDRateizzazione)), IDAdeguamento, fnNotNullN(rsAdeg!IDArticolo), fnNotNullN(rsAdeg!NoCalcPeriodoFatt), fnNotNull(rsAdeg!ArticoloAdeg)
        End If
        
        rsAdeg.CloseResultset
        Set rsAdeg = Nothing
    End If
        
        
    If ((frmMain.cboTipoImpostazione.CurrentID = 1) And (IDAdeguamento = 0)) Then
        'Dim PagamentoAnticipato As Boolean
        sSQL = "SELECT * FROM RV_PORateContratto "
        sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
        sSQL = sSQL & " AND ((Fatturata IS NULL) OR (Fatturata=0))"
        sSQL = sSQL & " AND Adeguamento=0"
        sSQL = sSQL & " AND Manuale=0"
        sSQL = sSQL & " AND ContrattoAttuale=1"
        sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"
        
        Set rsDelete = Cn.OpenResultset(sSQL)
        
        While Not rsDelete.EOF
            sSQL = "DELETE FROM RV_PORateContratto "
            sSQL = sSQL & " WHERE IDRV_PORateContratto=" & fnNotNullN(rsDelete!IDRV_PORateContratto)
            Cn.Execute sSQL
            
            IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(fnNotNullN(rsDelete!IDOggetto), fnNotNullN(rsDelete!IDTipoOggetto), 0)
        
            If IDOggettoScadenza > 0 Then
                ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, fnNotNullN(rsDelete!IDOggetto), fnNotNullN(rsDelete!IDTipoOggetto)
                ELIMINA_SCADENZA IDOggettoScadenza
            End If
            
        rsDelete.MoveNext
        Wend
        'PagamentoAnticipato = GET_PagamentoAnticipato(TheApp.Branch)
        ElaborazioneRate 1, TheApp.Branch, IDContratto, frmMain.txtDataDecorrenza.Text, frmMain.txtImportoAttuale.Value, Mesi_Rate, Numero_Rate, Pagamento_Anticipato_Periodo, frmMain.cboPagamentoRate.CurrentID, frmMain.txtIDContrattoPadre.Value, frmMain.cboTipoContratto.CurrentID, frmMain.cboSitoPerAnagrafica.CurrentID, Rata_Iniziale, Anno_Solare, frmMain.txtDataScadenzaPerRinnovo.Text, frmMain.txtNGGPrimaRata.Value
    End If
    
    If ((frmMain.cboTipoImpostazione.CurrentID = 2) And (IDAdeguamento = 0)) Then
        'Dim PagamentoAnticipato As Boolean
        sSQL = "SELECT * FROM RV_PORateContratto "
        sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
        
        Set rsDelete = Cn.OpenResultset(sSQL)
        
        While Not rsDelete.EOF
            sSQL = "DELETE FROM RV_PORateContratto "
            sSQL = sSQL & " WHERE IDRV_PORateContratto=" & fnNotNullN(rsDelete!IDRV_PORateContratto)
            sSQL = sSQL & " AND Adeguamento=0"
            sSQL = sSQL & " AND Manuale=0"
            sSQL = sSQL & " AND ContrattoAttuale=1"
            sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"
            sSQL = sSQL & " AND ((Fatturata IS NULL) OR (Fatturata=0))"
            Cn.Execute sSQL
            
            IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(fnNotNullN(rsDelete!IDOggetto), fnNotNullN(rsDelete!IDTipoOggetto), 0)
        
            If IDOggettoScadenza > 0 Then
                ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, fnNotNullN(rsDelete!IDOggetto), fnNotNullN(rsDelete!IDTipoOggetto)
                ELIMINA_SCADENZA IDOggettoScadenza
            End If
            
        rsDelete.MoveNext
        Wend
    
        If (frmMain.chkGeneraRateProd.Value = vbUnchecked) Then
            'ElaborazioneRateNoleggio TheApp.Branch, IDContratto, frmMain.txtDataDecorrenza.Text, frmMain.txtImportoAttuale.Value, Mesi_Rate, Numero_Rate, Pagamento_Anticipato_Periodo, frmMain.cboPagamentoRate.CurrentID, frmMain.txtIDContrattoPadre.Value, frmMain.cboTipoContratto.CurrentID, frmMain.cboSitoPerAnagrafica.CurrentID
            ElaborazioneRate 2, TheApp.Branch, IDContratto, frmMain.txtDataDecorrenza.Text, frmMain.txtImportoAttuale.Value, Mesi_Rate, Numero_Rate, Pagamento_Anticipato_Periodo, frmMain.cboPagamentoRate.CurrentID, frmMain.txtIDContrattoPadre.Value, frmMain.cboTipoContratto.CurrentID, frmMain.cboSitoPerAnagrafica.CurrentID, Rata_Iniziale, Anno_Solare, frmMain.txtDataScadenzaPerRinnovo.Text, frmMain.txtNGGPrimaRata.Value
        Else
            ElaborazioneRate 3, TheApp.Branch, IDContratto, frmMain.txtDataDecorrenza.Text, frmMain.txtImportoAttuale.Value, Mesi_Rate, Numero_Rate, Pagamento_Anticipato_Periodo, frmMain.cboPagamentoRate.CurrentID, frmMain.txtIDContrattoPadre.Value, frmMain.cboTipoContratto.CurrentID, frmMain.cboSitoPerAnagrafica.CurrentID, Rata_Iniziale, Anno_Solare, frmMain.txtDataScadenzaPerRinnovo.Text, frmMain.txtNGGPrimaRata.Value
        End If
        
    End If
    
End Sub
Private Sub ElaborazioneRate(Tipo As Long, IDFiliale As Long, IDContratto As Long, DataDecorrenza As String, ImportoContratto As Double, MesiRate As Long, numerorate As Long, PagamentoAnticipato As Boolean, IDPagamentoRata As Long, IDContrattoPadre As Long, IDTipoContratto As Long, IDSitoDestinazione As Long, RataIniziale As Long, AnnoSolare As Long, DataFineContratto As String, NGGPrimaRata As Long, Optional IDRateizzazione As Long = 0, Optional IDAdeguamento As Long = 0, Optional Articolo As Long = 0, Optional StampaPeriodo As Long = 1, Optional DescrizionePeriodo As String = "")
Dim DataRata As String
Dim IDOggettoRata As Long
Dim IDTipoOggettoRata As Long
Dim Periodo As String
Dim IDOggettoScadenza As Long
Dim IDRata As Long
Dim rsProd As DmtOleDbLib.adoResultset
Dim Avvia As Boolean
Dim Periodo_local As String

If Tipo = 3 Then
    
    CREA_RECORDSET_RATE
    
    sSQL = "SELECT * FROM RV_POContrattoProdotti "
    sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
    sSQL = sSQL & " AND Dismesso=0"
    
    Set rsProd = Cn.OpenResultset(sSQL)

    While Not rsProd.EOF
        
        Elaborazione IDContratto, DataDecorrenza, DataFineContratto, fnNotNullN(rsProd!ImportoComplessivo), numerorate, MesiRate, RataIniziale, AnnoSolare, rsProd!IDRV_POContrattoProdotti, rsProd!IDRV_POProdotto, NGGPrimaRata, fnNotNullN(rsProd!NonRateizzare)
        
    rsProd.MoveNext
    Wend
    
Else
    CREA_RECORDSET_RATE

    If (IDAdeguamento = 0) Then
        
        
        Elaborazione IDContratto, DataDecorrenza, DataFineContratto, ImportoContratto, numerorate, MesiRate, RataIniziale, AnnoSolare, 0, 0, NGGPrimaRata
    Else
        
        Elaborazione IDContratto, DataDecorrenza, DataFineContratto, ImportoContratto, numerorate, MesiRate, RataIniziale, AnnoSolare, 0, 0, NGGPrimaRata, 0, IDAdeguamento, Articolo
    
    End If
End If

If ((rsRateEla.BOF) And (rsRateEla.EOF)) Then Exit Sub

rsRateEla.MoveFirst

While Not rsRateEla.EOF
    
    If PagamentoAnticipato = True Then
        DataRata = rsRateEla!DataInizioPeriodo
    Else
        DataRata = rsRateEla!DataFinePeriodo
    End If
    
    IDRata = fnGetNewKey("RV_PORateContratto", "IDRV_PORateContratto")
    
    If IDAdeguamento = 0 Then
        Periodo_local = "Canone " & frmMain.cboTipoRateizzazione.Text & " " & frmMain.cboTipoContratto.Text & vbCrLf
        Periodo_local = Periodo_local & "Periodo di riferimento dal " & rsRateEla!DataInizioPeriodo & " al " & rsRateEla!DataFinePeriodo & vbCrLf
        Periodo_local = Periodo_local & "Periodo contratto dal " & frmMain.txtDataDecorrenza.Text & " al " & frmMain.txtDataScadenzaPerRinnovo.Text
    Else
        If StampaPeriodo = 1 Then
            Periodo_local = DescrizionePeriodo
        End If
    End If
    Periodo = ""
    
    sSQL = "INSERT INTO RV_PORateContratto ("
    sSQL = sSQL & "IDRV_PORateContratto, IDRV_POContratto, NumeroRata, DataRata, IDPagamentoRata, ImportoRata, "
    sSQL = sSQL & "Mese, Anno, Periodo, Adeguamento, Manuale, "
    sSQL = sSQL & "ContrattoAttuale, Fatturata, IDRV_POContrattoPadre, DataInizioPeriodo, DataFinePeriodo, "
    sSQL = sSQL & "IDTipoOggetto, IDOggetto, NonFatturare, AnnotazioniNonFatturare, IDRV_POProdotto, IDRV_POContrattoProdotti"
    If IDAdeguamento > 0 Then sSQL = sSQL & ", IDRV_POContrattoAdeguamento, IDArticolo"
    sSQL = sSQL & ") "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & IDRata & ", "
    sSQL = sSQL & IDContratto & ", "
    sSQL = sSQL & rsRateEla!numerorata & ", "
    sSQL = sSQL & fnNormDate(DataRata) & ", "
    sSQL = sSQL & IDPagamentoRata & ", "
    sSQL = sSQL & fnNormNumber(rsRateEla!ImportoRata) & ", "
    sSQL = sSQL & fnNormNumber(DatePart("m", DataRata)) & ", "
    sSQL = sSQL & fnNormNumber(DatePart("yyyy", DataRata)) & ", "
    sSQL = sSQL & fnNormString(Periodo) & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & 1 & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & IDContrattoPadre & ", "
    sSQL = sSQL & fnNormDate(rsRateEla!DataInizioPeriodo) & ", "
    sSQL = sSQL & fnNormDate(rsRateEla!DataFinePeriodo) & ", "
    
    IDTipoOggettoRata = fnGetTipoOggetto("RV_PORateContratto")
    IDOggettoRata = GET_LINK_OGGETTO(0, fnGetTipoOggetto("RV_PORateContratto"), rsRateEla!numerorata, DataRata)
    
    sSQL = sSQL & IDTipoOggettoRata & ", "
    sSQL = sSQL & IDOggettoRata & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & fnNormString("") & ", "
    sSQL = sSQL & rsRateEla!IDRV_POProdotto & ", "
    sSQL = sSQL & rsRateEla!IDRV_POContrattoProdotti
    If IDAdeguamento > 0 Then
        sSQL = sSQL & ", " & rsRateEla!IDRV_POContrattoAdeguamento & ", "
        sSQL = sSQL & rsRateEla!IDArticolo
    End If
    
    sSQL = sSQL & ")"
    Cn.Execute sSQL
    
    ''''''''''''''''FLUSSO DOCUMENTALE SCADENZARIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If LINK_SEZIONALE_RATE > 0 Then
        IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(IDOggettoRata, IDTipoOggettoRata, 0)
        
        IDOggettoScadenza = GET_LINK_SCADENZA(rsRateEla!ImportoRata, IDClienteFatturazione, rsRateEla!numerorata, DataRata, LINK_SEZIONALE_RATE, Periodo)
        
        CREA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, IDOggettoRata, IDTipoOggettoRata
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (IDAdeguamento = 0) Then
        Periodo = GET_STRINGA_PERIODO_ADEG(1, TheApp.Branch, IDContratto, IDRata, IDAdeguamento, fnNotNullN(rsRateEla!IDRV_POProdotto), fnNotNullN(rsRateEla!IDRV_POContrattoProdotti))
    Else
        Periodo = GET_STRINGA_PERIODO_ADEG(2, TheApp.Branch, IDContratto, IDRata, IDAdeguamento, fnNotNullN(rsRateEla!IDRV_POProdotto), fnNotNullN(rsRateEla!IDRV_POContrattoProdotti))
    End If
    
    If Len(Periodo) = 0 Then
        Periodo = Periodo_local
    End If
    
    If IDAdeguamento = 0 Then
        sSQL = "UPDATE RV_PORateContratto SET "
        sSQL = sSQL & "Periodo=" & fnNormString(Periodo)
        sSQL = sSQL & "WHERE IDRV_PORateContratto=" & IDRata
        Cn.Execute sSQL
    Else
        If StampaPeriodo = 0 Then
            sSQL = "UPDATE RV_PORateContratto SET "
            sSQL = sSQL & "Periodo=" & fnNormString(Periodo)
            sSQL = sSQL & "WHERE IDRV_PORateContratto=" & IDRata
            Cn.Execute sSQL
        Else
            sSQL = "UPDATE RV_PORateContratto SET "
            sSQL = sSQL & "Periodo=" & fnNormString(Periodo_local)
            sSQL = sSQL & "WHERE IDRV_PORateContratto=" & IDRata
            Cn.Execute sSQL
        End If
    End If
rsRateEla.MoveNext
Wend

rsRateEla.Close
Set rsRateEla = Nothing

End Sub
Public Function ControlloRatePagate(IDContratto As Long) As Boolean
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT * FROM RV_PORateContratto WHERE ("
    sSQL = sSQL & "(IDRV_POContratto=" & IDContratto & ") AND "
    sSQL = sSQL & "(IDOggettoCollegato>0))"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        ControlloRatePagate = False
    Else
        ControlloRatePagate = True
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub AccodaStoriaRata(IDStoriaContratto As Long, IDRiferimentoRata As Long, numerorata As Long, DataScadenza As String, ImportoRata As Double, IDPagamentoRata As Long, mese As Long, Anno As Long, Periodo As String, Adeguamento As Boolean, Manuale As Boolean, ContrattoAttuale As Boolean, Fatturata As Boolean)
    Dim sSQL As String
    
                 
        sSQL = "INSERT INTO RV_POStoriaRateContratto ("
        sSQL = sSQL & "IDRV_POStoriaRateContratto, IDRV_POStoriaContratto, IDRiferimentoRata, NumeroRata, DataRata, IDPagamentoRate, ImportoRata, Mese, Anno, Periodo, Adeguamento, Manuale, ContrattoAttuale, Fatturata) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnGetNewKey("RV_POStoriaRateContratto", "IDRV_POStoriaRateContratto") & ", "
        sSQL = sSQL & Link_StoriaContratto & ", "
        sSQL = sSQL & IDRiferimentoRata & ", "
        sSQL = sSQL & numerorata & ", "
        sSQL = sSQL & fnNormDate(DataScadenza) & ", "
        sSQL = sSQL & IDPagamentoRata & ", "
        sSQL = sSQL & fnNormNumber(ImportoRata) & ", "
        sSQL = sSQL & fnNormNumber(mese) & ", "
        sSQL = sSQL & fnNormNumber(Anno) & ", "
        sSQL = sSQL & fnNormString(Periodo) & ", "
        sSQL = sSQL & Adeguamento & ", "
        sSQL = sSQL & Manuale & ", "
        sSQL = sSQL & ContrattoAttuale & ", "
        sSQL = sSQL & Fatturata & ")"
        
        
        Cn.Execute sSQL
    
End Sub
Public Sub RiAccodaStoriaRata(IDStoriaContratto As Long, IDContratto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_PORateContratto WHERE IDRV_POContratto=" & IDContratto
Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
                 
        sSQL = "INSERT INTO RV_POStoriaRateContratto ("
        sSQL = sSQL & "IDRV_POStoriaRateContratto, IDRV_POStoriaContratto, IDRiferimentoRata, NumeroRata, DataRata, "
        sSQL = sSQL & "IDPagamentoRate, ImportoRata, Mese, Anno, Periodo, "
        sSQL = sSQL & "Adeguamento, Manuale, ContrattoAttuale, IDOggettoCollegato, Fatturata) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnGetNewKey("RV_POStoriaRateContratto", "IDRV_POStoriaRateContratto") & ", "
        sSQL = sSQL & IDStoriaContratto & ", "
        sSQL = sSQL & fnNotNullN(rs!IDRV_PORateContratto) & ", "
        sSQL = sSQL & fnNotNullN(rs!numerorata) & ", "
        sSQL = sSQL & fnNormDate(rs!DataRata) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDPagamentoRata) & ", "
        sSQL = sSQL & fnNormNumber(rs!ImportoRata) & ", "
        sSQL = sSQL & fnNormNumber(rs!mese) & ", "
        sSQL = sSQL & fnNormNumber(rs!Anno) & ", "
        sSQL = sSQL & fnNormString(rs!Periodo) & ", "
        sSQL = sSQL & fnNotNullN(rs!Adeguamento) & ", "
        sSQL = sSQL & fnNotNullN(rs!Manuale) & ", "
        sSQL = sSQL & fnNotNullN(rs!ContrattoAttuale) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDOggettoCollegato) & ", "
        If fnNotNullN(rs!IDOggettoCollegato) > 0 Then
            sSQL = sSQL & 1 & ")"
        Else
            sSQL = sSQL & 0 & ")"
        End If
        
        Cn.Execute sSQL
    rs.MoveNext
    Wend

rs.CloseResultset
Set rs = Nothing
End Sub
Public Function ContrattoAttualeStorico(IDContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POStoriaContratto FROM RV_POStoriaContratto WHERE ("
sSQL = sSQL & "(IDRV_POContratto=" & IDContratto & ") AND "
sSQL = sSQL & "(ContrattoAttuale=" & 1 & "))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    ContrattoAttualeStorico = fnNotNullN(rs!IDRV_POStoriaContratto)
Else
    ContrattoAttualeStorico = 0
End If
End Function
Private Sub AggiornaStoriaRata(IDRiferimentoRata As Long)
    Dim sSQL As String
    
                 
        sSQL = "UPDATE  RV_POStoriaRateContratto SET "
        sSQL = sSQL & "DataRata=" & fnNormDate(DataScadenza) & ", "
        sSQL = sSQL & "PagamentoRate=" & IDPagamentoRata & ", "
        sSQL = sSQL & "ImportoRata=" & fnNormNumber(ImportoRata) & ")"
        sSQL = sSQL & "Fatturata=" & frmMain.chkRataFatturata
        sSQL = sSQL & "IDRiferimentoRata=" & IDRiferimentoRata
        
        Cn.Execute sSQL
    
End Sub
Private Sub EliminaStoriaRata(IDRiferimentoRata As Long)
    Dim sSQL As String
    
                 
        sSQL = "DELETE FROM RV_POStoriaRateContratto WHERE IDRiferimentoRata=" & IDRiferimentoRata
        Cn.Execute sSQL
    
End Sub
Public Function numerorate(IDContratto As Long)
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT NumeroRata FROM RV_PORateContratto WHERE IDRV_POContratto=" & IDContratto
    sSQL = sSQL & " ORDER BY NumeroRata DESC"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        numerorate = 1
    Else
        numerorate = fnNotNullN(rs!numerorata) + 1
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Public Function GET_STRINGA_PERIODO(IDFiliale As Long, IDContratto As Long, numerorata As Long, DataInizioRata As String, DataFineRata As String) As String
Dim rsContratto As DmtOleDbLib.adoResultset
Dim rsAdeguamento As DmtOleDbLib.adoResultset
Dim rs As DmtOleDbLib.adoResultset
Dim AvviaStringa As Boolean


GET_STRINGA_PERIODO = ""
AvviaStringa = False

'RECORDSET CONTRATTO
sSQL = "SELECT * FROM RV_POViewContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rsContratto = Cn.OpenResultset(sSQL)

If Not rsContratto.EOF Then
    AvviaStringa = True
End If

'RECORDSET ADEGUAMENTO
'sSQL = "SELECT * FROM RV_POIEAdeguamentiContratto "
'sSQL = sSQL & "WHERE IDRV_POContrattoAdeguamento=" & IDContrattoAdeguamento
'
'Set rsAdeguamento = Cn.OpenResultset(sSQL)'
'
'If rsAdeguamento.EOF Then
'    AvviaStringa = False
'End If
If AvviaStringa = False Then
    rsContratto.CloseResultset
    Set rsContratto = Nothing
    Exit Function
End If
sSQL = "SELECT RV_POStringaPeriodoRighe.IDRV_POCampoPeriodo, RV_POStringaPeriodoRighe.Posizione, RV_POStringaPeriodoRighe.Tipo, "
sSQL = sSQL & "RV_POStringaPeriodoRighe.Testo "
sSQL = sSQL & "FROM RV_POStringaPeriodoTesta INNER JOIN "
sSQL = sSQL & "RV_POStringaPeriodoRighe ON RV_POStringaPeriodoTesta.IDRV_POStringaPeriodoTesta = RV_POStringaPeriodoRighe.IDRV_POStringaPeriodoTesta "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & " AND RV_POStringaPeriodoRighe.Tipo=1"
sSQL = sSQL & " ORDER BY Posizione "

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    Select Case rs!IDRV_POCampoPeriodo
        Case 1 'Stringa personalizzata
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rs!Testo)
        Case 2 'Tipo contratto
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!TipoContratto) 'TipoContratto
        Case 3 'Data decorrenza
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!DataDecorrenza) 'DataDecorrenza
        Case 4 'Data scadenza
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!DataScadenza) 'DataScadenza
        Case 5 'Data rinnovo
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!DataScadenzaPerRinnovo) 'DataRinnovo
        Case 6 'Tipo rateizzazione
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!Rateizzazione) 'TipoRateizzazione
        Case 7 'Durata contratto
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!DurataContratto) 'DurataContratto
        Case 8 'Tipo rinnovo
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!TipoRinnovo) 'TipoRinnovo
        Case 9 'Numero licenza
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!NumeroLicenze) 'NumeroLicenza
        Case 10 ' Numero rate
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & numerorata
        Case 11 'Data inizio rata
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & DataInizioRata
        Case 12 'Data fine rata
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & DataFineRata
        Case 13 'Carattere speciale spazio
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & " "
        Case 14 'Carattere speciale A Capo
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & vbCrLf
        Case 15 'Descrizione del tipo contratto
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!DescrizioneAggTipoContratto) 'DescrizioneTipoContratto
        Case 16 'ELENCO SERVIZI
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO
        Case 17 'Descrizione articolo tipo contratto
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!ArticoloTipoContratto) 'GET_ARTICOLO_TIPO_CONTRATTO(IDTipoContratto, "Articolo")
        Case 18 'Descrizione ridotta dell'articolo del tipo contratto
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!DescrizioneAggTipoContratto) 'GET_ARTICOLO_TIPO_CONTRATTO(IDTipoContratto, "DescrizioneArticoloRidotta")
        Case 19 'Sito destinazione
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!SitoPerAnagrafica) 'GET_DESTINAZIONE_DIVERSA(IDSitoDestinazione)
        Case 20 'Descrizione articolo contratto
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!ArticoloContratto)
        Case 21 'Descrizione articolo ridotta contratto
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!DescrizioneArtRidContratto)
        Case 22 'Numero rinnovo
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNullN(rsContratto!NumeroRinnovo)
        Case 23 'Numero protocollo contratto
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!NumeroProtocollo)
        Case 24 'Descrizione aggiuntiva del contratto
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rsContratto!DescrizioneTipoContratto)
        Case 25 'Descrizione articolo adeguamento
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO
        Case 26 'Descrizione articolo ridotta adeguamento
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO
        Case 27 'Numero protocollo adeguamento
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO
        
    End Select
     
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing





rsContratto.CloseResultset
Set rsContratto = Nothing


End Function
Public Function GET_PagamentoAnticipato(IDFiliale As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PagamentoAnticipato FROM RV_POStringaPeriodoTesta "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    PagamentoAnticipato = False
Else
    PagamentoAnticipato = fnNotNullN(rs!PagamentoAnticipato)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_IVA_TIPO_CONTRATTO(IDTipoContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIva FROM RV_POTipoContratto "
sSQL = sSQL & "WHERE IDRV_POTipoContratto=" & IDTipoContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_TIPO_CONTRATTO = 0
Else
    GET_LINK_IVA_TIPO_CONTRATTO = fnNotNullN(rs!IDIva)
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
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Function GET_LINK_OGGETTO(IDOggetto As Long, IDTipoOggetto As Long, numerorata As Long, DataScadenzaRata As String) As Long
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
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
        rs!IDTipoOggetto = IDTipoOggetto
        rs!IDFunzione = IDFunzione
        rs!IDAzienda = TheApp.IDFirm
        rs!IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.Branch)
        rs!IDSezionale = 0
        rs!Oggetto = GET_DESCRIZIONE_FUNZIONE(IDFunzione)
        rs!DataEmissione = DataScadenzaRata
        rs!numero = numerorata
        rs!DataUltimaVariazione = Date
        rs!IDUtenteUltimaVariazione = TheApp.IDUser
        rs!VirtualDelete = 0
        rs!IDOggetto = fnGetNewKey("Oggetto", "IDOggetto")
        GET_LINK_OGGETTO = rs!IDOggetto
    rs.Update
Else
    rs!DataEmissione = DataScadenzaRata
    rs!numero = numerorata
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

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ATTIVITA_AZIENDA = 0
Else
    GET_LINK_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
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

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_FLUSSO_FUNZIONE = 0
Else
    GET_LINK_FLUSSO_FUNZIONE = fnNotNullN(rs!IDFlussoFunzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long)
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
sSQL = sSQL & "WHERE Descrizione=" & fnNormString("Rata contratto -> Scadenza")
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFLussoGruppo")
        rsNew!Descrizione = "Rata contratto -> Scadenza"
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
    
    rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
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

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    sSQL = "DELETE FROM FlussoFunzioneCollegato "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoRata
    sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
    sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
    Cn.Execute sSQL
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
Cn.Execute sSQL
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

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

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

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

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

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

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

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

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

Set rs = Cn.OpenResultset(sSQL)

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
    Cn.Execute sSQL

    sSQL = "DELETE FROM DettaglioScadenza "
    sSQL = sSQL & "WHERE IDTestataScadenza=" & IDTestataScadenza
    Cn.Execute sSQL

    sSQL = "DELETE FROM Oggetto "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoScadenza
    Cn.Execute sSQL
    
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
    
    rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
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

sSQL = "SELECT * FROM Oggetto "
sSQL = sSQL & "WHERE IDOggetto=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    rs!IDOggetto = fnGetNewKey("Oggetto", "IDOggetto")
    rs!IDTipoOggetto = 131
    rs!IDFunzione = 88
    rs!Oggetto = "Testata Scadenza del " & DataDocumento
    rs!IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.Branch)
    rs!IDAzienda = TheApp.IDFirm
    rs!IDSezionale = IDSezionale
    rs!DataEmissione = DataDocumento
    rs!numero = NumeroDocumento
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

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

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
    rsNew!numerorata = NumeroDocumento
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

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_REGISTRO_IVA = 1
Else
    GET_LINK_REGISTRO_IVA = fnNotNullN(rs!IDRegistroIva)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ARTICOLO_TIPO_CONTRATTO(IDTipoContratto As Long, NomeCampoArticolo As String) As String
On Error GoTo ERR_GET_ARTICOLO_TIPO_CONTRATTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Articolo." & NomeCampoArticolo
sSQL = sSQL & " FROM RV_POTipoContratto INNER JOIN "
sSQL = sSQL & "Articolo ON RV_POTipoContratto.IDArticolo = Articolo.IDArticolo "
sSQL = sSQL & "WHERE IDRV_POTipoContratto=" & IDTipoContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ARTICOLO_TIPO_CONTRATTO = ""
Else
    GET_ARTICOLO_TIPO_CONTRATTO = fnNotNull(rs.adoColumns(NomeCampoArticolo).Value)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_ARTICOLO_TIPO_CONTRATTO:
    MsgBox Err.Description, vbCritical, "GET_ARTICOLO_TIPO_CONTRATTO"
End Function
Private Function GET_DESTINAZIONE_DIVERSA(IDSitoPerAnagrafica As Long) As String
On Error GoTo ERR_GET_DESTINAZIONE_DIVERSA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SitoPerAnagrafica FROM SitoPerAnagrafica "
sSQL = sSQL & "WHERE IDSitoPerAnagrafica=" & IDSitoPerAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESTINAZIONE_DIVERSA = ""
Else
    GET_DESTINAZIONE_DIVERSA = fnNotNull(rs!SitoPerAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_DESTINAZIONE_DIVERSA:
    MsgBox Err.Description, vbCritical, "GET_DESTINAZIONE_DIVERSA"
End Function
Public Function GET_STRINGA_PERIODO_ADEG(Tipo As Long, IDFiliale As Long, IDContratto As Long, IDContrattoRate As Long, IDContrattoAdeguamento As Long, IDProdotto As Long, Optional IDContrattoProdotti As Long = 0) As String
On Error GoTo ERR_GET_STRINGA_PERIODO_ADEG
Dim rsContratto As DmtOleDbLib.adoResultset
Dim rsAdeguamento As DmtOleDbLib.adoResultset
Dim rsRate As DmtOleDbLib.adoResultset
Dim rs As DmtOleDbLib.adoResultset
Dim AvviaStringa As Boolean
Dim sSQL As String
Dim rsProd As DmtOleDbLib.adoResultset
Dim rsProdContr As DmtOleDbLib.adoResultset


GET_STRINGA_PERIODO_ADEG = ""

AvviaStringa = False

'RECORDSET CONTRATTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POViewContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rsContratto = Cn.OpenResultset(sSQL)

If Not rsContratto.EOF Then
    AvviaStringa = True
End If

'RECORDSET ADEGUAMENTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIEAdeguamentiContratto "
sSQL = sSQL & "WHERE IDRV_POContrattoAdeguamento=" & IDContrattoAdeguamento

Set rsAdeguamento = Cn.OpenResultset(sSQL)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'RECORDSET RATE CONTRATTO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_PORateContratto=" & IDContrattoRate

Set rsRate = Cn.OpenResultset(sSQL)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'RECORDSET PRODOTTO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POProdotto "
sSQL = sSQL & "WHERE IDRV_POProdotto=" & IDProdotto

Set rsProd = Cn.OpenResultset(sSQL)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'RECORDSET PRODOTTO NEL CONTRATTO'''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDRV_POContrattoProdotti, CodiceArticolo, Articolo, DataInizioPeriodo, DataFinePeriodo, "
sSQL = sSQL & "DescrizioneAggiuntiva, Annotazioni, ValoreIndentificativo "
sSQL = sSQL & "FROM RV_POIEContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDContrattoProdotti

Set rsProdContr = Cn.OpenResultset(sSQL)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT RV_POStringaPeriodoRighe.IDRV_POCampoPeriodo, RV_POStringaPeriodoRighe.Posizione, RV_POStringaPeriodoRighe.Tipo, "
sSQL = sSQL & "RV_POStringaPeriodoRighe.Testo "
sSQL = sSQL & "FROM RV_POStringaPeriodoTesta INNER JOIN "
sSQL = sSQL & "RV_POStringaPeriodoRighe ON RV_POStringaPeriodoTesta.IDRV_POStringaPeriodoTesta = RV_POStringaPeriodoRighe.IDRV_POStringaPeriodoTesta "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & " AND RV_POStringaPeriodoRighe.Tipo=" & Tipo
sSQL = sSQL & " ORDER BY Posizione "

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    Select Case rs!IDRV_POCampoPeriodo
        Case 1 'Stringa personalizzata
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rs!Testo)
        Case 2 'Tipo contratto
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!TipoContratto) 'TipoContratto
        Case 3 'Data decorrenza
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!DataDecorrenza) 'DataDecorrenza
        Case 4 'Data scadenza
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!DataScadenza) 'DataScadenza
        Case 5 'Data rinnovo
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!DataScadenzaPerRinnovo) 'DataRinnovo
        Case 6 'Tipo rateizzazione
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!Rateizzazione) 'TipoRateizzazione
        Case 7 'Durata contratto
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!DurataContratto) 'DurataContratto
        Case 8 'Tipo rinnovo
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!TipoRinnovo) 'TipoRinnovo
        Case 9 'Numero licenza
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!NumeroLicenze) 'NumeroLicenza
        Case 10 ' Numero rate
            If Not rsRate.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsRate!numerorata)
            End If
        Case 11 'Data inizio rata
            If Not rsRate.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsRate!DataInizioPeriodo)
            End If
        Case 12 'Data fine rata
            If Not rsRate.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsRate!DataFinePeriodo)
            End If
        Case 13 'Carattere speciale spazio
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & " "
        Case 14 'Carattere speciale A Capo
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & vbCrLf
        Case 15 'Descrizione del tipo contratto
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!DescrizioneAggTipoContratto) 'DescrizioneTipoContratto
        Case 16 'ELENCO SERVIZI
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG
        Case 17 'Descrizione articolo tipo contratto
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!ArticoloTipoContratto) 'GET_ARTICOLO_TIPO_CONTRATTO(IDTipoContratto, "Articolo")
        Case 18 'Descrizione ridotta dell'articolo del tipo contratto
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!DescrizioneAggTipoContratto) 'GET_ARTICOLO_TIPO_CONTRATTO(IDTipoContratto, "DescrizioneArticoloRidotta")
        Case 19 'Sito destinazione
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!SitoPerAnagrafica) 'GET_DESTINAZIONE_DIVERSA(IDSitoDestinazione)
        Case 20 'Descrizione articolo contratto
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!ArticoloContratto)
        Case 21 'Descrizione articolo ridotta contratto
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!DescrizioneArtRidContratto)
        Case 22 'Numero rinnovo
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNullN(rsContratto!NumeroRinnovo)
        Case 23 'Numero protocollo contratto
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!NumeroProtocollo)
        Case 24 'Descrizione aggiuntiva del contratto
            GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsContratto!DescrizioneTipoContratto)
        Case 25 'Descrizione articolo adeguamento
            If Not rsAdeguamento.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsAdeguamento!ArticoloAdeg)
            End If
        Case 26 'Descrizione articolo ridotta adeguamento
            If Not rsAdeguamento.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsAdeguamento!DescrizioneArticoloRidottaAdeg)
            End If
        Case 27 'Numero protocollo adeguamento
            If Not rsAdeguamento.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsAdeguamento!NumeroProtocollo)
            End If
        Case 28
            If Not rsAdeguamento.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsAdeguamento!DescrizionePerFatturazione)
            End If
        Case 29 'Descrizione prodotto
            If Not rsProd.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsProd!Descrizione)
            End If
        Case 30 'Matricola prodotto
            If Not rsProdContr.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsProdContr!ValoreIndentificativo)
            End If
        Case 31 'Codice articolo del prodotto nel contratto
            If Not rsProdContr.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsProdContr!CodiceArticolo)
            End If
        Case 32 'Descrizione articolo del prodotto nel contratto
            If Not rsProdContr.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsProdContr!Articolo)
            End If
        Case 33 'Data inizio periodo prodotto nel contratto
            If Not rsProdContr.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsProdContr!DataInizioPeriodo)
            End If
        Case 34 'Data fine periodo prodotto nel contratto
            If Not rsProdContr.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsProdContr!DataFinePeriodo)
            End If
        Case 35 'Annotazioni prodotto nel contratto
            If Not rsProdContr.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsProdContr!Annotazioni)
            End If
        Case 36 'Ubicazione prodotto nel contratto
            If Not rsProdContr.EOF Then
                GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG & fnNotNull(rsProdContr!DescrizioneAggiuntiva)
            End If
    End Select
     
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

rsContratto.CloseResultset
Set rsContratto = Nothing

If Not rsAdeguamento.EOF Then
    rsAdeguamento.CloseResultset
    Set rsAdeguamento = Nothing
End If
If Not rsRate.EOF Then
    rsRate.CloseResultset
    Set rsRate = Nothing
End If
If Not rsProd.EOF Then
    rsProd.CloseResultset
    Set rsProd = Nothing
End If
If Not rsProdContr.EOF Then
    rsProdContr.CloseResultset
    Set rsProdContr = Nothing
End If

Exit Function
ERR_GET_STRINGA_PERIODO_ADEG:
    MsgBox Err.Description, vbCritical, "ERR_GET_STRINGA_PERIODO_ADEG"
    GET_STRINGA_PERIODO_ADEG = GET_STRINGA_PERIODO_ADEG
    
End Function
Public Sub SviluppoRateContrattoProdotto(IDContratto As Long, IDContrattoProdotti As Long)
On Error GoTo ERR_SviluppoRateContrattoProdotto
Dim sSQL As String
Dim rsDelete As DmtOleDbLib.adoResultset
Dim IDOggettoScadenza As Long
Dim rsContrattoProd As ADODB.Recordset

    If frmMain.cboTipoImpostazione.CurrentID = 3 Then
'        'Dim PagamentoAnticipato As Boolean
        sSQL = "SELECT * FROM RV_PORateContratto "
        sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
        sSQL = sSQL & " AND IDRV_POContrattoProdotti=" & IDContrattoProdotti
        sSQL = sSQL & " AND Fatturata=0"
        
        Set rsDelete = Cn.OpenResultset(sSQL)

        While Not rsDelete.EOF
            sSQL = "DELETE FROM RV_PORateContratto "
            sSQL = sSQL & " WHERE IDRV_PORateContratto=" & fnNotNullN(rsDelete!IDRV_PORateContratto)
            Cn.Execute sSQL

            IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(fnNotNullN(rsDelete!IDOggetto), fnNotNullN(rsDelete!IDTipoOggetto), 0)

            If IDOggettoScadenza > 0 Then
                ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, fnNotNullN(rsDelete!IDOggetto), fnNotNullN(rsDelete!IDTipoOggetto)
                ELIMINA_SCADENZA IDOggettoScadenza
            End If

        rsDelete.MoveNext
        Wend

        rsDelete.CloseResultset
        Set rsDelete = Nothing
        
        '''RECUPERO DATI DELLA RIGA PRODOTTO NEL CONTRATTO''''''''''''''''''''''''''''''''''
        sSQL = "SELECT * FROM RV_POContrattoProdotti "
        sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDContrattoProdotti
        
        Set rsContrattoProd = New ADODB.Recordset
        
        rsContrattoProd.Open sSQL, Cn.InternalConnection
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not rsContrattoProd.EOF Then
            ElaborazioneRateProdotto TheApp.Branch, IDContratto, rsContrattoProd!DataInizioPeriodo, fnNotNullN(rsContrattoProd!Imponibile), frmMain.cboPagamentoRate.CurrentID, frmMain.txtIDContrattoPadre.Value, frmMain.cboTipoContratto.CurrentID, frmMain.cboSitoPerAnagrafica.CurrentID, rsContrattoProd!DataFinePeriodo, rsContrattoProd!IDRateizzazione, rsContrattoProd!IDRV_POContrattoProdotti, rsContrattoProd!IDRV_POProdotto, fnNotNullN(rsContrattoProd!QuantitaPeriodo), fnNotNullN(rsContrattoProd!IDRV_POUnitaDiMisuraPeriodo)
        End If
        
        rsContrattoProd.Close
        Set contrattoprod = Nothing
    End If

Exit Sub
ERR_SviluppoRateContrattoProdotto:
    MsgBox Err.Description, vbCritical, "Sviluppo rate contratto prodotto"
End Sub
Private Sub ElaborazioneRateProdotto(IDFiliale As Long, IDContratto As Long, DataDecorrenza As String, ImportoContratto As Double, IDPagamentoRata As Long, IDContrattoPadre As Long, IDTipoContratto As Long, IDSitoDestinazione As Long, DataFineContratto As String, IDRateizzazione As Long, IDContrattoProdotti As Long, IDProdotto As Long, QuantitaPeriodo As Double, IDTipoPeriodo As Long)
Dim DataRata As String
Dim IDOggettoRata As Long
Dim IDTipoOggettoRata As Long
Dim Periodo As String
Dim IDOggettoScadenza As Long
Dim IDRata As Long
Dim rsProd As DmtOleDbLib.adoResultset
Dim Avvia As Boolean

Dim MesiRateLocal As Long
Dim NumeroRateLocal As Long
Dim PagamentoAnticipatoLocal As Long
Dim RataInizialeLocal As Long
Dim AnnoSolareLocal As Long
Dim ConsideraQuantitaPeriodoLocal As Long
Dim PercentualePrimaRataLocal As Double
Dim DueRateLocal As Long
Dim rsRateizzazione As DmtOleDbLib.adoResultset


''''RECUPERO DATI DELLA RATEIZZAZIONE''''''''''''''''''''''''''''
MesiRateLocal = 1
NumeroRateLocal = 1
PagamentoAnticipatoLocal = 1
RataInizialeLocal = 0
AnnoSolareLocal = 0
ConsideraQuantitaPeriodoLocal = 0
PercentualePrimaRataLocal = 0
DueRateLocal = 0

sSQL = "SELECT * FROM RV_PORateizzazione "
sSQL = sSQL & "WHERE IDRV_PORateizzazione=" & IDRateizzazione

Set rsRateizzazione = Cn.OpenResultset(sSQL)

If Not rsRateizzazione.EOF Then
    MesiRateLocal = fnNotNullN(rsRateizzazione!Mesi)
    NumeroRateLocal = fnNotNullN(rsRateizzazione!numerorate)
    PagamentoAnticipatoLocal = fnNotNullN(rsRateizzazione!PagamentoInizioPeriodo)
    RataInizialeLocal = fnNotNullN(rsRateizzazione!RataInizialeRataFinale)
    AnnoSolareLocal = fnNotNullN(rsRateizzazione!AnnoSolare)
    ConsideraQuantitaPeriodoLocal = fnNotNullN(rsRateizzazione!ConsideraQuantitaPeriodo)
    PercentualePrimaRataLocal = fnNotNullN(rsRateizzazione!PercentualePrimaRata)
    DueRateLocal = fnNotNullN(rsRateizzazione!duerate)
End If

rsRateizzazione.CloseResultset
Set rsRateizzazione = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

CREA_RECORDSET_RATE

ElaborazionePerProdotto IDContratto, DataDecorrenza, DataFineContratto, ImportoContratto, NumeroRateLocal, MesiRateLocal, RataInizialeLocal, AnnoSolareLocal, IDContrattoProdotti, IDProdotto, 0, PercentualePrimaRataLocal, DueRateLocal, ConsideraQuantitaPeriodoLocal, QuantitaPeriodo, IDTipoPeriodo
    
If ((rsRateEla.BOF) And (rsRateEla.EOF)) Then Exit Sub

rsRateEla.MoveFirst

While Not rsRateEla.EOF
    
    If PagamentoAnticipato = True Then
        DataRata = rsRateEla!DataInizioPeriodo
    Else
        DataRata = rsRateEla!DataFinePeriodo
    End If
    
    
    IDRata = fnGetNewKey("RV_PORateContratto", "IDRV_PORateContratto")
    
    Periodo_local = "Canone " & frmMain.cboTipoRateizzazione.Text & " " & frmMain.cboTipoContratto.Text & vbCrLf
    Periodo_local = Periodo_local & "Periodo di riferimento dal " & rsRateEla!DataInizioPeriodo & " al " & rsRateEla!DataFinePeriodo & vbCrLf
    Periodo_local = Periodo_local & "Periodo contratto dal " & frmMain.txtDataDecorrenza.Text & " al " & frmMain.txtDataScadenzaPerRinnovo.Text

    
    Periodo = ""
    
    sSQL = "INSERT INTO RV_PORateContratto ("
    sSQL = sSQL & "IDRV_PORateContratto, IDRV_POContratto, NumeroRata, DataRata, IDPagamentoRata, ImportoRata, "
    sSQL = sSQL & "Mese, Anno, Periodo, Adeguamento, Manuale, "
    sSQL = sSQL & "ContrattoAttuale, Fatturata, IDRV_POContrattoPadre, DataInizioPeriodo, DataFinePeriodo, "
    sSQL = sSQL & "IDTipoOggetto, IDOggetto, NonFatturare, AnnotazioniNonFatturare, IDRV_POProdotto, IDRV_POContrattoProdotti) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & IDRata & ", "
    sSQL = sSQL & IDContratto & ", "
    sSQL = sSQL & rsRateEla!numerorata & ", "
    sSQL = sSQL & fnNormDate(DataRata) & ", "
    sSQL = sSQL & IDPagamentoRata & ", "
    sSQL = sSQL & fnNormNumber(rsRateEla!ImportoRata) & ", "
    sSQL = sSQL & fnNormNumber(DatePart("m", DataRata)) & ", "
    sSQL = sSQL & fnNormNumber(DatePart("yyyy", DataRata)) & ", "
    sSQL = sSQL & fnNormString(Periodo) & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & 1 & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & IDContrattoPadre & ", "
    sSQL = sSQL & fnNormDate(rsRateEla!DataInizioPeriodo) & ", "
    sSQL = sSQL & fnNormDate(rsRateEla!DataFinePeriodo) & ", "
    
    IDTipoOggettoRata = fnGetTipoOggetto("RV_PORateContratto")
    IDOggettoRata = GET_LINK_OGGETTO(0, fnGetTipoOggetto("RV_PORateContratto"), rsRateEla!numerorata, DataRata)
    
    sSQL = sSQL & IDTipoOggettoRata & ", "
    sSQL = sSQL & IDOggettoRata & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & fnNormString("") & ", "
    sSQL = sSQL & rsRateEla!IDRV_POProdotto & ", "
    sSQL = sSQL & rsRateEla!IDRV_POContrattoProdotti & ")"
    Cn.Execute sSQL
    
    ''''''''''''''''FLUSSO DOCUMENTALE SCADENZARIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If LINK_SEZIONALE_RATE > 0 Then
        IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(IDOggettoRata, IDTipoOggettoRata, 0)
        
        IDOggettoScadenza = GET_LINK_SCADENZA(rsRateEla!ImportoRata, IDClienteFatturazione, rsRateEla!numerorata, DataRata, LINK_SEZIONALE_RATE, Periodo)
        
        CREA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, IDOggettoRata, IDTipoOggettoRata
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Periodo = GET_STRINGA_PERIODO_ADEG(1, TheApp.Branch, IDContratto, IDRata, 0, fnNotNullN(rsRateEla!IDRV_POProdotto), fnNotNullN(rsRateEla!IDRV_POContrattoProdotti))
    
    If fnNotNullN(rsRateEla!IDRV_POProdotto) = 0 Then
        If (NO_CALCOLO_PERIODO_FATT = 1) Then
            Periodo = ""
        Else
            If Len(Periodo) = 0 Then
                Periodo = Periodo_local
            End If
        End If
    Else
        If Len(Periodo) = 0 Then
            Periodo = Periodo_local
        End If
    End If
    
    sSQL = "UPDATE RV_PORateContratto SET "
    sSQL = sSQL & "Periodo=" & fnNormString(Periodo)
    sSQL = sSQL & "WHERE IDRV_PORateContratto=" & IDRata
    Cn.Execute sSQL
    
rsRateEla.MoveNext
Wend

rsRateEla.Close
Set rsRateEla = Nothing

End Sub
