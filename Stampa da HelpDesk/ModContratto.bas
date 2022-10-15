Attribute VB_Name = "ModContratto"
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Link_StoriaContratto As Long
Public Link_StoriaContrattoPrecedente As Long
Public Var_Rateizzazione As String 'Contiene la descrizione del tipo di rateizzazione
Public Var_TipoContratto As String 'Contiene la descrizione del tipo di contratto
Public Var_TipoDurata As String 'Contiene la descrizione della durata del contratto
Public Var_TipoRinnovo As String 'Contiene la descrizione del tipo rinnovo del contratto
Public Sub EliminaRateContratto(IDContratto As Long)
    Dim sSQL As String
    
    sSQL = "DELETE FROM RV_PORateContratto WHERE IDRV_POContratto=" & IDContratto
        
    CnDMT.Execute sSQL
    
End Sub
Public Sub ElaborazioneRate(rsNuovoContratto As ADODB.Recordset, IDContratto As Long, DataDecorrenza As String, ImportoContratto As Double, IDPagamentoRata As Long)
Dim ImportoRata As Double
Dim UltimaRata As Double
Dim ImportoRataProgressiva As Double
Dim DataRataProgressiva As String
Dim I As Long
Dim sSQL As String
Dim IDRata As Long
Dim DataFinePeriodo As String
Dim Periodo As String
Dim Var_NumeroRata As Long
Dim Var_RimanenzaMesiDiRate As Long
Dim rs As DmtOleDbLib.adoResultset
Dim MesiRate As Long
Dim NumeroRate As Long
Dim PagamentoAnticipato As Boolean
Dim DataRata As String
Dim IDTipoOggettoRata As Long
Dim IDOggettoRata As Long
Dim IDOggettoScadenza As Long

    GET_INFO_RATEIZZAZIONE fnNotNullN(rsNuovoContratto!IDRateizzazione), MesiRate, NumeroRate, PagamentoAnticipato
    
    ImportoRata = 0
    ImportoRataProgressiva = 0
    
    DataRataProgressiva = DataDecorrenza
    DataFinePeriodo = DateAdd("m", MesiRate, DataRataProgressiva) - 1
    
    If PagamentoAnticipato = True Then
        DataRata = DataRataProgressiva
    Else
        DataRata = DataFinePeriodo
    End If
    
    If ImportoContratto > 0 Then
        ImportoRata = FormatNumber((ImportoContratto / NumeroRate), 2)
    Else
        ImportoRata = ImportoContratto
    End If
    
    Var_NumeroRata = 0
    Var_RimanenzaMesiDiRate = 0
    
    'If NumeroRate > 1 Then
    '    MsgBox "STOP"
    'End If
    
    For I = 1 To NumeroRate 'Mesi_Rate To Mesi_Rinnovo_Contratto Step Mesi_Rate
        Var_NumeroRata = Var_NumeroRata + 1

        If Var_NumeroRata = NumeroRate Then
            ImportoRata = ImportoContratto - ImportoRataProgressiva
        End If
        'If Var_NumeroRata > 1 Then
        '    DataRataProgressiva = DateAdd("m", MesiRate, DataRataProgressiva)
        'End If

        Periodo = GET_STRINGA_PERIODO(TheApp.Branch, Var_TipoContratto, DataDecorrenza, fnNotNull(rsNuovoContratto!DataScadenza), fnNotNull(rsNuovoContratto!DataScadenzaPerRinnovo), Var_Rateizzazione, Var_TipoDurata, Var_TipoRinnovo, fnNotNullN(rsNuovoContratto!NumeroLicenze), Var_NumeroRata, DataRataProgressiva, DataFinePeriodo, Descrizione_Tipo_Contratto, PagamentoAnticipato)
        If Periodo = "" Then
            Periodo = "Canone " & Var_Rateizzazione & " " & Var_TipoContratto & vbCrLf
            Periodo = Periodo & "Periodo di riferimento dal " & DataRataProgressiva & " al " & DataFinePeriodo & vbCrLf
            Periodo = Periodo & "Periodo contratto dal " & DataDecorrenza & " al " & fnNotNull(rs!DataScadenzaPerRinnovo)
        End If

        IDRata = fnGetNewKey("RV_PORateContratto", "IDRV_PORateContratto")
        sSQL = "INSERT INTO RV_PORateContratto ("
        sSQL = sSQL & "IDRV_PORateContratto, IDRV_POContratto, NumeroRata, DataRata, IDPagamentoRata, "
        sSQL = sSQL & "ImportoRata, Mese, Anno, Periodo, Adeguamento, Manuale, ContrattoAttuale, "
        sSQL = sSQL & "IDRV_POContrattoPadre, IDArticolo, DataInizioPeriodo, DataFinePeriodo, "
        sSQL = sSQL & "IDTipoOggetto, IDOggetto, NonFatturare, AnnotazioniNonFatturare) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & IDRata & ", "
        sSQL = sSQL & IDContratto & ", "
        sSQL = sSQL & Var_NumeroRata & ", "
        sSQL = sSQL & fnNormDate(DataRata) & ", "
        sSQL = sSQL & IDPagamentoRata & ", "
        sSQL = sSQL & fnNormNumber(ImportoRata) & ", "
        sSQL = sSQL & fnNormNumber(DatePart("m", DataRata)) & ", "
        sSQL = sSQL & fnNormNumber(DatePart("yyyy", DataRata)) & ", "
        sSQL = sSQL & fnNormString(Periodo) & ", "
        sSQL = sSQL & 0 & ", "
        sSQL = sSQL & 0 & ", "
        sSQL = sSQL & 1 & ", "
        sSQL = sSQL & fnNotNullN(rsNuovoContratto!IDRV_POContrattoPadre) & ", "
        sSQL = sSQL & 0 & ", " 'IDArticolo
        sSQL = sSQL & fnNormDate(DataRataProgressiva) & ", "
        sSQL = sSQL & fnNormDate(DataFinePeriodo) & ", "
        IDTipoOggettoRata = fnGetTipoOggetto("RV_PORateContratto")
        IDOggettoRata = GET_LINK_OGGETTO(0, fnGetTipoOggetto("RV_PORateContratto"), Var_NumeroRata, DataRata)
        sSQL = sSQL & IDTipoOggettoRata & ", "
        sSQL = sSQL & IDOggettoRata & ", "
        sSQL = sSQL & 0 & ", "
        sSQL = sSQL & fnNormString("") & ")"
    
        CnDMT.Execute sSQL
        If LINK_SEZIONALE_RATE > 0 Then
             ''''''''''''''''FLUSSO DOCUMENTALE SCADENZARIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
             'IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(IDOggettoRata, IDTipoOggettoRata, 0)
            
             IDOggettoScadenza = GET_LINK_SCADENZA(ImportoRata, rsNuovoContratto!IDAnagraficaFatturazione, Var_NumeroRata, DataRata, LINK_SEZIONALE_RATE, Periodo)
             
             CREA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, IDOggettoRata, IDTipoOggettoRata
             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
        
        DataRataProgressiva = DateAdd("d", 1, DataFinePeriodo)
        DataFinePeriodo = DateAdd("m", MesiRate, DataRataProgressiva) - 1
        ImportoRataProgressiva = ImportoRataProgressiva + FormatNumber(ImportoRata, 2)

        If PagamentoAnticipato = True Then
            DataRata = DataRataProgressiva
        Else
            DataRata = DataFinePeriodo
        End If
    Next


End Sub


Public Function NumeroRate(IDContratto As Long)
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT NumeroRata FROM RV_PORateContratto WHERE IDRV_POContratto=" & IDContratto
    sSQL = sSQL & " ORDER BY NumeroRata DESC"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        NumeroRate = 1
    Else
        NumeroRate = rs!NumeroRata + 1
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Public Function DescrizioneRateizzazione(ID) As String
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT Rateizzazione FROM RV_PORateizzazione WHERE IDRV_PORateizzazione=" & ID
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        Var_Rateizzazione = rs!Rateizzazione
    Else
        Var_Rateizzazione = ""
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Public Function DescrizioneTipoRinnovo(ID) As String
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoRinnovo FROM RV_POTipoRinnovo WHERE IDRV_POTipoRinnovo=" & ID
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        Var_TipoRinnovo = fnNotNull(rs!TipoRinnovo)
    Else
        Var_TipoRinnovo = ""
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Public Function DescrizioneTipoDurata(ID) As String
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT DurataContratto FROM RV_PODurataContratto WHERE IDRV_PODurataContratto=" & ID
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        Var_TipoDurata = fnNotNull(rs!DurataContratto)
    Else
        Var_TipoDurata = ""
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Public Function DescrizioneTipoContratto(IDContratto) As String
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT RV_POTipoContratto.TipoContratto, RV_POTipoContratto.DescrizioneAggiuntiva "
    sSQL = sSQL & "FROM RV_POContratto LEFT OUTER JOIN "
    sSQL = sSQL & "RV_POTipoContratto ON RV_POContratto.IDTipoContratto = RV_POTipoContratto.IDRV_POTipoContratto "
    sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
    
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        Var_TipoContratto = fnNotNull(rs!TipoContratto)
        Descrizione_Tipo_Contratto = fnNotNull(rs!DescrizioneAggiuntiva)
    Else
        Var_TipoContratto = ""
        Descrizione_Tipo_Contratto = ""
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Public Sub AggiornamentoRatePrecedenti()
    Dim sSQL As String
    
    sSQL = "UPDATE RV_POStoriaRateContratto SET "
    sSQL = sSQL & "IDRiferimentoRata=0" & ", "
    sSQL = sSQL & "ContrattoAttuale=" & 1 & " "
    sSQL = sSQL & "WHERE IDRV_POStoriaContratto=" & Link_StoriaContrattoPrecedente
    
    CnDMT.OpenResultset sSQL
End Sub
Public Function GET_STRINGA_PERIODO(IDFiliale As Long, TipoContratto As String, DataDecorrenza As String, DataScadenza As String, DataRinnovo As String, TipoRateizzazione As String, DurataContratto As String, TipoRinnovo As String, NumeroLicenza As Long, NumeroRata As Long, DataInizioRata As String, DataFineRata As String, DescrizioneTipoContratto As String, PagamentoAnticipato As Boolean) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


GET_STRINGA_PERIODO = ""


sSQL = "SELECT RV_POStringaPeriodoRighe.IDRV_POCampoPeriodo, RV_POStringaPeriodoRighe.Posizione, "
sSQL = sSQL & "RV_POStringaPeriodoRighe.Testo "
sSQL = sSQL & "FROM RV_POStringaPeriodoTesta INNER JOIN "
sSQL = sSQL & "RV_POStringaPeriodoRighe ON RV_POStringaPeriodoTesta.IDRV_POStringaPeriodoTesta = RV_POStringaPeriodoRighe.IDRV_POStringaPeriodoTesta "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & " ORDER BY Posizione "

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    Select Case rs!IDRV_POCampoPeriodo
        Case 1 'Stringa personalizzata
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & fnNotNull(rs!Testo)
        Case 2 'Tipo contratto
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & TipoContratto
        Case 3 'Data decorrenza
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & DataDecorrenza
        Case 4 'Data scadenza
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & DataScadenza
        Case 5 'Data rinnovo
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & DataRinnovo
        Case 6 'Tipo rateizzazione
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & TipoRateizzazione
        Case 7 'Durata contratto
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & DurataContratto
        Case 8 'Tipo rinnovo
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & TipoRinnovo
        Case 9 'Numero licenza
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & NumeroLicenza
        Case 10 ' Numero rate
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & NumeroRata
        Case 11 'Data inizio rata
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & DataInizioRata
        Case 12 'Data fine rata
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & DataFineRata
        Case 13 'Carattere speciale spazio
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & " "
        Case 14 'Carattere speciale A Capo
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & vbCrLf
        Case 15
            GET_STRINGA_PERIODO = GET_STRINGA_PERIODO & DescrizioneTipoContratto
    End Select
     
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_INFO_RATEIZZAZIONE(IDTipoRateizzazione As Long, MesiRate As Long, NumeroRate As Long, PagamentoAnticipato As Boolean)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_PORateizzazione "
sSQL = sSQL & "WHERE IDRV_PORateizzazione=" & IDTipoRateizzazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    MesiRate = 0
    NumeroRate = 0
    PagamentoAnticipato = False
Else
    MesiRate = fnNotNullN(rs!Mesi)
    NumeroRate = fnNotNullN(rs!NumeroRate)
    PagamentoAnticipato = fnNotNullN(rs!PagamentoInizioPeriodo)
End If


rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_LINK_IVA_TIPO_CONTRATTO(IDTipoContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIva FROM RV_POTipoContratto "
sSQL = sSQL & "WHERE IDRV_POTipoContratto=" & IDTipoContratto

Set rs = CnDMT.OpenResultset(sSQL)

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

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF = False Then
    fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
Else
    fnGetTipoOggetto = 0
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

If rs.EOF Then
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

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

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


