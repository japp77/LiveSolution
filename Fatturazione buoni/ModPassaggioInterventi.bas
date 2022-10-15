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
Public Data_Documento As String
Public Link_Valuta As Long
Public Link_Magazzino As Long
Public Link_PagamentoDefault As Long
Public Var_TipoOggetto As String
Public Link_Articolo As Long
Public Var_Codice_Articolo As String
Public Link_Iva_Cliente As Long
Public SingolaFattura As Long


Private oReport As dmtReportLib.dmtReport
Private progMin As Double
Private Link_IDOggetto As Long
Private Link_IDRataRiferimento As Long
Private IDRata As Long

Private TotaleRecord As Integer
Private NomeAltraSede As String


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
 
 
Public LINK_TIPO_GESTIONE_BUONI As Long
Public LINK_TIPO_NUMERAZIONE_BUONO As Long


Public NUMERO_DOC_PER_CLIENTE As Long
Public NUMERO_DOCUMENTO As Long


Public Sub fncPassaggioDocumenti()
'On Error GoTo ERR_fncPassaggioDocumenti
Dim sSQL As String
Dim I As Integer
Dim Unita_Progresso As Double
Dim Numero As Long

    'Funzione che setta la variabile ObjDoc con i dati predefiniti
    
    If SingolaFattura = 0 Then
        If (NUMERO_ANAGRAFICHE = 0) Then Exit Sub
        
        frmCreazioneDocumenti.ProgressBar1.Value = 0
        
        frmCreazioneDocumenti.ProgressBar1.Max = 100
        
        Unita_Progresso = frmCreazioneDocumenti.ProgressBar1.Max / NUMERO_ANAGRAFICHE
        
        frmCreazioneDocumenti.lblInfo.Caption = "FATTURAZIONE IN CORSO..."
        DoEvents
        
        I = 0
    '''''''''''''''''''''''''''''INIZIO PASSAGGIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        rsAna.MoveFirst
        Numero = 0
        
        While Not rsAna.EOF
        
            Numero = Numero + 1
            
            frmCreazioneDocumenti.lblInfo.Caption = "FATTURAZIONE " & Numero & " di " & NUMERO_ANAGRAFICHE
            DoEvents
        
            If Not (ObjDoc Is Nothing) Then
                Set ObjDoc = Nothing
            End If
            
            Settaggio
            
            fncTestata rsAna!IDAnagraficaCliente, rsAna!RitenutaAcconto, rsAna!IDSitoPerAnagrafica
            
            fncRighe rsAna!IDAnagraficaCliente, rsAna!RitenutaAcconto, rsAna!IDSitoPerAnagrafica
            
            If (InserimentoDMT = True) Then
                AGGIORNAMENTO_COLLEGAMENTO_BUONO rsAna!IDAnagraficaCliente, ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto, rsAna!RitenutaAcconto, rsAna!IDSitoPerAnagrafica
                
                If frmCreazioneDocumenti.chkStampa.Value = 1 Then
                    
                    If Link_IDOggetto > 0 Then
                        DoEvents
                        ObjDoc.Prepare2Print TheApp.IDFirm, TheApp.IDUser, ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto
                        StampaDocumento
                        DoEvents
                    End If
                    
                End If
            Else
                MsgBox VARErroreGenerico, vbCritical, VARErroreFunzione
            End If
            
            
            If (frmCreazioneDocumenti.ProgressBar1.Value + Unita_Progresso) >= frmCreazioneDocumenti.ProgressBar1.Max Then
                frmCreazioneDocumenti.ProgressBar1.Value = frmCreazioneDocumenti.ProgressBar1.Max
            Else
                frmCreazioneDocumenti.ProgressBar1.Value = frmCreazioneDocumenti.ProgressBar1.Value + Unita_Progresso
            End If
            DoEvents
            
        rsAna.MoveNext
        Wend
    
    Else
        If (NUMERO_ADDEBITI_DA_FATTURARE = 0) Then Exit Sub
        
        frmCreazioneDocumenti.ProgressBar1.Value = 0
       
        frmCreazioneDocumenti.ProgressBar1.Max = 100
        
        Unita_Progresso = frmCreazioneDocumenti.ProgressBar1.Max / NUMERO_ADDEBITI_DA_FATTURARE
        
        frmCreazioneDocumenti.lblInfo.Caption = "FATTURAZIONE IN CORSO..."
        DoEvents
        
        I = 0
    '''''''''''''''''''''''''''''INIZIO PASSAGGIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        rsnew.MoveFirst
        Numero = 0
        
        While Not rsnew.EOF
        
            Numero = Numero + 1
            
            frmCreazioneDocumenti.lblInfo.Caption = "FATTURAZIONE " & Numero & " di " & NUMERO_ADDEBITI_DA_FATTURARE
            DoEvents
        
            If Not (ObjDoc Is Nothing) Then
                Set ObjDoc = Nothing
            End If
            
            Settaggio
            
            fncTestata rsnew!IDAnagraficaFatturazione, rsnew!RitenutaAcconto, rsnew!IDSitoPerAnagraficaIntervento
            
            fncRigheSingole rsnew!IDAnagraficaFatturazione, rsnew!RitenutaAcconto
            
            If (InserimentoDMT = True) Then
            
                AGGIORNAMENTO_COLLEGAMENTO_BUONO_SINGOLO rsnew!IDAnagraficaFatturazione, ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto, rsnew!RitenutaAcconto
                
                If frmCreazioneDocumenti.chkStampa.Value = 1 Then
                    
                    If Link_IDOggetto > 0 Then
                        DoEvents
                        ObjDoc.Prepare2Print TheApp.IDFirm, TheApp.IDUser, ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto
                        StampaDocumento
                        DoEvents
                    End If
                    
                End If
            Else
                MsgBox VARErroreGenerico, vbCritical, VARErroreFunzione
            End If
            
            
            If (frmCreazioneDocumenti.ProgressBar1.Value + Unita_Progresso) >= frmCreazioneDocumenti.ProgressBar1.Max Then
                frmCreazioneDocumenti.ProgressBar1.Value = frmCreazioneDocumenti.ProgressBar1.Max
            Else
                frmCreazioneDocumenti.ProgressBar1.Value = frmCreazioneDocumenti.ProgressBar1.Value + Unita_Progresso
            End If
            DoEvents
            
        rsnew.MoveNext
        Wend
    End If

Exit Sub
ERR_fncPassaggioDocumenti:
    MsgBox Err.Description, vbCritical, "fncPassaggioDocumenti"

End Sub

Private Function fncTestata(IDAnagrafica As Long, RitenutaAcconto As Long, IDSitoPerAnagrafica) As Boolean
'On Error GoTo ERR_fncTestata
Dim IDLetteraIntento As Long

VARErroreFunzione = "fncTestata"
   
         With ObjDoc.Tables
        
        'Imposta la riga attiva per la tabella di testata
            
            ObjDoc.Tables(NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)).SetActiveRetail 1
            
            'TrovaAnagrafica IDRata
            ObjDoc.ReadDataFromCliFo IDAnagrafica, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            If (FLAG_RAGGR_ALTRA_DEST = 1) Then
                ObjDoc.ReadDataFromCliFoSite IDSitoPerAnagrafica, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            End If
            .Field "Doc_perc_rit_acc", ObjDoc.DBDefaults.PercentualeRitenutaAcconto, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Nom_calcola_rit_acc", RitenutaAcconto, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "RV_PONonStampaDescrAgg", 1, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            
            'Dati generici del documento
            .Field "Link_Doc_magazzino", Link_Magazzino, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Link_Doc_sezionale", Link_Sezionale, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Doc_prefisso", GET_PREFISSO_SEZ(ObjDoc.IDSezionale), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            If .Field("Link_Doc_pagamento", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)) = 0 Then
                ObjDoc.ReadDataFromPayment Link_PagamentoDefault, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            End If
            
            .Field "Link_Val_valuta", Link_Valuta, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Link_Val_cambio", Null, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Doc_data", Data_Documento, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Doc_numero", 0, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            
             Link_Iva_Cliente = fnNotNullN(.Field("Link_nom_Iva", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))
'
'            Link_Iva_Cliente = GET_LINK_IVA_CLIENTE_ESENTE(IDAnagrafica, 0, Data_Documento)
'
'            IDLetteraIntento = 0
'
'            If GET_CONTROLLO_NUMERO_LETTERE_INTENTO(IDAnagrafica, TheApp.IDFirm, Date) = 1 Then
'                IDLetteraIntento = GET_LINK_LETTERA_INTENTO(IDAnagrafica, TheApp.IDFirm, Year(Date))
'                Link_Iva_Cliente = GET_LINK_IVA_LETTERA_INTENTO(IDLetteraIntento, Link_Iva_Cliente)
'            End If
'
'            ObjDoc.Field "Link_Nom_IVA", Link_Iva_Cliente, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
'            ObjDoc.Field "Link_Nom_lettera_intento", IDLetteraIntento, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
          
            
            
        End With
        
        fncTestata = True
     
Exit Function
'ERR_fncTestata:
'    fncTestata = False
'    VARErroreIDIntervento = "GENERALITA':" & vbCrLf & "IDCliente : " & IDClienteOP
'    VARErroreGenerico = Err.Description & vbCrLf & VARErroreIDIntervento
    
End Function

Private Function fncRigheSingole(IDAnagrafica As Long, RitenutaAcconto As Long) As Boolean
'On Error GoTo ERR_fncRighe
Dim I As Integer
Dim J As Integer
Dim X As Integer
Dim Y As Integer
Dim DescrizioneRiga As String

Dim sSQL As String
Dim ArrayDescrizioneFattura() As String
Dim ArrayDescrFattConACapo() As String
Dim Stringa_Buono As String
Dim IDAnagraficaIntervento As Long
Dim DescrizioneClienteIntervento As String


VARErroreFunzione = "fncRighe"
    
    
    
'sSQL = "SELECT * FROM RV_POTMPFatturazioneBuoni "
'sSQL = sSQL & "WHERE IDAnagraficaCliente=" & IDAnagrafica
'sSQL = sSQL & " AND RitenutaAcconto=" & RitenutaAcconto - 1
'sSQL = sSQL & " ORDER BY IDAnagraficaIntervento"
'Set rsnew = CnDMT.OpenResultset(sSQL)

'rsnew.Filter = "IDAnagraficaFatturazione=" & IDAnagrafica
'rsnew.Filter = rsnew.Filter & " AND RitenutaAcconto=" & RitenutaAcconto

IDAnagraficaIntervento = 0
I = 1
'While Not rsnew.EOF
    
    With ObjDoc.Tables
        'INTESTAZIONE CLIENTE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If IDAnagraficaIntervento <> fnNotNullN(rsnew!IDAnagraficaCliente) Then
            If fnNotNullN(rsnew!IDAnagraficaFatturazione) <> fnNotNullN(rsnew!IDAnagraficaCliente) Then
                DescrizioneClienteIntervento = GET_ANAGRAFICA_INTERVENTO(fnNotNullN(rsnew!IDAnagraficaFatturazione))
                If Len(DescrizioneClienteIntervento) > 0 Then
                    ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                    .Field "Art_descrizione", DescrizioneClienteIntervento, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    I = I + 1
                End If
            End If
            IDAnagraficaIntervento = fnNotNullN(rsnew!IDAnagraficaFatturazione)
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'INTESTAZIONE BUONO
        Stringa_Buono = GET_STRINGA_BUONO(fnNotNullN(rsnew!IDRV_POInterventoRigheDett))
        
        If LINK_TIPO_GESTIONE_BUONI = 2 Then
            If Len(Stringa_Buono) > 0 Then
                ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                .Field "Art_descrizione", Mid(Stringa_Buono, 1, 255), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                I = I + 1
            End If
        End If
        
        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
        ObjDoc.ReadDataFromArticle rsnew!IDArticolo
        
        If LINK_TIPO_GESTIONE_BUONI <= 1 Then
            If Len(Stringa_Buono) = 0 Then
                .Field "Art_descrizione", "Numero " & fnNotNullN(rsnew!NumeroDocumento) & " del " & fnNotNull(rsnew!DataDocumento), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            Else
                .Field "Art_descrizione", Mid(Stringa_Buono, 1, 255), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            End If
            .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_prezzo_unitario_neutro", fnNotNullN(rsnew!ImportoDiFatturazione), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            If Link_Iva_Cliente > 0 Then
                .Field "Link_art_IVA", Link_Iva_Cliente, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA(Link_Iva_Cliente), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            End If
        Else
            GET_DETTAGLIO_RIGA_BUONO rsnew
        End If
        
        I = I + 1
        
        If (FLAG_SOVRAS_DESCRIZIONE = 0) Then
            If GET_CONTROLLO_DESCRIZIONE_ART_RIGA_FATT(rsnew!IDRV_POInterventoRigheDett) = False Then
                'RIGHE DESCRITTIVE
                ArrayDescrizioneFattura = Split(GET_DESCRIZIONE_FATTURA(rsnew!IDRV_POInterventoRigheDett), "||")
                For J = 0 To UBound(ArrayDescrizioneFattura)
                    If Len(Trim(ArrayDescrizioneFattura(J))) > 0 Then
                        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                        'ArrayDescrFattConACapo = Split(Trim(ArrayDescrizioneFattura(J)), Chr(13))
                        'For X = 0 To UBound(ArrayDescrFattConACapo)
                            .Field "Art_descrizione", Trim(ArrayDescrizioneFattura(J)), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        'Next X
                    End If
                    I = I + 1
                Next J
            End If
        End If
    End With
'rsnew.MoveNext
'Wend


    
'fncRigheSingole = True

'rsnew.Filter = vbNullString


Exit Function

'ERR_fncRighe:
'    fncRighe = False
'    VARErroreIDIntervento = "GENERALITA':" & vbCrLf & "IDCliente : " & IDClienteOP
'    'VARErroreIDArticolo = vbCrLf & "Articolo: " & rsArt!Articolo
'    VARErroreGenerico = Err.Description & vbCrLf & VARErroreIDIntervento ' & VARErroreIDArticolo


End Function

Private Function fncRighe(IDAnagrafica As Long, RitenutaAcconto As Long, IDSitoPerAnagrafica As Long) As Boolean
'On Error GoTo ERR_fncRighe
Dim I As Integer
Dim J As Integer
Dim X As Integer
Dim Y As Integer
Dim DescrizioneRiga As String

Dim sSQL As String
Dim ArrayDescrizioneFattura() As String
Dim ArrayDescrFattConACapo() As String
Dim Stringa_Buono As String
Dim IDAnagraficaIntervento As Long
Dim IDSitoPerAnagraficaInt As Long
Dim IDIntervento As Long

Dim DescrizioneClienteIntervento As String


VARErroreFunzione = "fncRighe"
    
'sSQL = "SELECT * FROM RV_POTMPFatturazioneBuoni "
'sSQL = sSQL & "WHERE IDAnagraficaCliente=" & IDAnagrafica
'sSQL = sSQL & " AND RitenutaAcconto=" & RitenutaAcconto - 1
'sSQL = sSQL & " ORDER BY IDAnagraficaIntervento"
'Set rsnew = CnDMT.OpenResultset(sSQL)

rsnew.Filter = "IDAnagraficaFatturazione=" & IDAnagrafica
rsnew.Filter = rsnew.Filter & " AND RitenutaAcconto=" & RitenutaAcconto
If (FLAG_RAGGR_ALTRA_DEST = 1) Then
    rsnew.Filter = rsnew.Filter & " AND IDSitoPerAnagraficaIntervento=" & IDSitoPerAnagrafica
End If

rsnew.Sort = "IDSitoPerAnagraficaIntervento, AnnoIntervento, NumeroIntervento"

IDAnagraficaIntervento = 0
IDSitoPerAnagraficaInt = 0
IDIntervento = 0

I = 1
While Not rsnew.EOF
    
    With ObjDoc.Tables
        'INTESTAZIONE CLIENTE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If IDAnagraficaIntervento <> fnNotNullN(rsnew!IDAnagraficaCliente) Then
            If fnNotNullN(rsnew!IDAnagraficaFatturazione) <> fnNotNullN(rsnew!IDAnagraficaCliente) Then
                DescrizioneClienteIntervento = GET_ANAGRAFICA_INTERVENTO(fnNotNullN(rsnew!IDAnagraficaFatturazione))
                If Len(DescrizioneClienteIntervento) > 0 Then
                    ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                    .Field "Art_descrizione", DescrizioneClienteIntervento, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    I = I + 1
                End If
            End If
            IDAnagraficaIntervento = fnNotNullN(rsnew!IDAnagraficaFatturazione)
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'INTESTAZIONE BUONO
        Stringa_Buono = GET_STRINGA_BUONO(fnNotNullN(rsnew!IDRV_POInterventoRigheDett))
        
        If LINK_TIPO_GESTIONE_BUONI = 2 Then
            If (FLAG_RAGGR_CORPO_ALTRA_DEST = 1) Then
                If (IDSitoPerAnagraficaInt <> fnNotNullN(rsnew!IDSitoPerAnagraficaIntervento)) Then
                    If fnNotNullN(rsnew!IDSitoPerAnagraficaIntervento) > 0 Then
                        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                        .Field "Art_descrizione", "Filiale: " & GET_DESCRIZIONE_SITO(fnNotNullN(rsnew!IDSitoPerAnagraficaIntervento)), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        I = I + 1
                        IDSitoPerAnagraficaInt = fnNotNullN(rsnew!IDSitoPerAnagraficaIntervento)
                    End If
                End If
            End If
            If (FLAG_RAGGR_CORPO_INT = 1) Then
                If (IDIntervento <> fnNotNullN(rsnew!IDRV_POintervento)) Then
                    ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                    .Field "Art_descrizione", "Intervento numero " & fnNotNullN(rsnew!AnnoIntervento) & "-" & fnNotNullN(rsnew!NumeroIntervento) & " del " & GET_FORMATTA_DATA(fnNotNull(rsnew!DataChiamata)), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    I = I + 1
                    IDIntervento = fnNotNullN(rsnew!IDRV_POintervento)
                End If
            Else
                If Len(Stringa_Buono) > 0 Then
                    ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                    .Field "Art_descrizione", Mid(Stringa_Buono, 1, 255), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    I = I + 1
                End If
            End If
        End If
        
        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
        ObjDoc.ReadDataFromArticle rsnew!IDArticolo
        
        If LINK_TIPO_GESTIONE_BUONI <= 1 Then
            If Len(Stringa_Buono) = 0 Then
                .Field "Art_descrizione", "Numero " & fnNotNullN(rsnew!NumeroDocumento) & " del " & fnNotNull(rsnew!DataDocumento), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            Else
                .Field "Art_descrizione", Mid(Stringa_Buono, 1, 255), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            End If
            .Field "Art_quantita_totale", 1, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_prezzo_unitario_neutro", fnNotNullN(rsnew!ImportoDiFatturazione), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            If Link_Iva_Cliente > 0 Then
                .Field "Link_art_IVA", Link_Iva_Cliente, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA(Link_Iva_Cliente), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            End If
        Else
            GET_DETTAGLIO_RIGA_BUONO rsnew
        End If
        
        I = I + 1
        
        If (FLAG_SOVRAS_DESCRIZIONE = 0) Then
            If GET_CONTROLLO_DESCRIZIONE_ART_RIGA_FATT(rsnew!IDRV_POInterventoRigheDett) = False Then
                'RIGHE DESCRITTIVE
                ArrayDescrizioneFattura = Split(GET_DESCRIZIONE_FATTURA(rsnew!IDRV_POInterventoRigheDett), "||")
                For J = 0 To UBound(ArrayDescrizioneFattura)
                    If Len(Trim(ArrayDescrizioneFattura(J))) > 0 Then
                        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                        'ArrayDescrFattConACapo = Split(Trim(ArrayDescrizioneFattura(J)), Chr(13))
                        'For X = 0 To UBound(ArrayDescrFattConACapo)
                            .Field "Art_descrizione", Trim(ArrayDescrizioneFattura(J)), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                            .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                        'Next X
                    End If
                    I = I + 1
                Next J
            End If
        End If
    End With
rsnew.MoveNext
Wend


    
fncRighe = True

rsnew.Filter = vbNullString


Exit Function

'ERR_fncRighe:
'    fncRighe = False
'    VARErroreIDIntervento = "GENERALITA':" & vbCrLf & "IDCliente : " & IDClienteOP
'    'VARErroreIDArticolo = vbCrLf & "Articolo: " & rsArt!Articolo
'    VARErroreGenerico = Err.Description & vbCrLf & VARErroreIDIntervento ' & VARErroreIDArticolo


End Function

Private Function fnContrattoAltraFiliale() As Boolean
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT RV_POContratto.IDSitoPerAnagrafica, SitoPerAnagrafica.SitoPerAnagrafica "
    sSQL = sSQL & "FROM RV_POContratto LEFT OUTER JOIN "
    sSQL = sSQL & "SitoPerAnagrafica ON dbo.RV_POContratto.IDSitoPerAnagrafica = SitoPerAnagrafica.IDSitoPerAnagrafica "
    sSQL = sSQL & "WHERE IDRV_POContratto=" & rsReg!IDRV_POCOntratto
    
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




Private Sub Settaggio()
Set ObjDoc = New DmtDocs.cDocument
    With ObjDoc
        Set .Connection = CnDMT
        .IDAzienda = TheApp.IDFirm
        .IDAttivitaAzienda = GET_ATTIVITA_AZIENDA
        .IDFiliale = TheApp.Branch
        .SetTipoOggetto Link_TipoOggetto
        .IDFunzione = fncTrovaIDFunzione
        .UseAutomation = True
        .IDEsercizio = fnGetEsercizio(Data_Documento)
        .IDSezionale = Link_Sezionale
        .IDTipoAnagrafica = 2
        .IDUtente = TheApp.IDUser
        .Descrizione = Var_TipoOggetto
        .DataEmissione = Data_Documento
        .Numero = 0
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
'On Error GoTo ERR_InserimentoDMT
VARErroreFunzione = "InserimentoDMT"



Screen.MousePointer = vbHourglass
    
    
    'SetDefault
    
    Set ObjDoc.Scadenze = Nothing
    ObjDoc.PerformDocument Nothing
    
    VarNumeroDoc = ObjDoc.Insert
    'ObjDoc.Update
    If VarNumeroDoc > 0 Then
        InserimentoDMT = True
    Else
        InserimentoDMT = False
    End If
    
Screen.MousePointer = vbDefault
    
Exit Function

ERR_InserimentoDMT:
    InserimentoDMT = False
        
        VARErroreIDIntervento = "GENERALITA':" & vbCrLf & "Passaggio in contabilità della rata"
        VARErroreGenerico = Err.Description & vbCrLf & VARErroreIDIntervento
    Screen.MousePointer = vbDefault
    
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
    fncGetNumeroIntervento = fnNotNull(rsInt!NumeroIntervento) & "-" & fnNotNull(rsInt!AnnoIntervento)
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
        
        
        'Imposta l'idfiliale di appartenenza del documento da stampare
            oReport.BranchID = TheApp.Branch 'IDFiliale
        'Imposta l'identificativo del tipo di documento
            oReport.DocTypeID = Link_TipoOggetto
            'oReport.Where = "IDOggetto = 873" '& Val(Me.Txt_Reg_IDRegistro)
            oReport.Where = "ValoriOggettoPerTipo" & fnGetHex(Link_TipoOggetto) & ".IDOggetto = " & Link_IDOggetto
            oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
    
            oReport.Copies = frmCreazioneDocumenti.txtNumeroCopie.Text
            oReport.DoPrint frmCreazioneDocumenti.DMTDialog.PrinterName
   
Exit Sub
ERR_StampaDocumento:
    MsgBox Err.Description, vbCritical, "Stampa Documento"
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

sSQL = "SELECT IDOggetto From Oggetto "
sSQL = sSQL & "WHERE IDTipoOggetto=" & Link_TipoOggetto
sSQL = sSQL & " AND IDSezionale=" & Link_Sezionale
sSQL = sSQL & " AND Numero=" & fnNormString(VarNumeroDoc)
sSQL = sSQL & " AND DataEmissione=" & fnNormDate(Data_Documento)
sSQL = sSQL & " ORDER BY IDOggetto DESC"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaDocumento = fnNotNullN(rs!IDOggetto)
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
Private Sub fncIDOggettoCollegatoIntervento(IDRata As Long, NumeroIntervento As String, ID As Long)
Dim sSQL As String

Dim VarIDOggettoCollegato As Long

VarIDOggettoCollegato = fncTrovaIDOggettoCollegato(Link_TipoOggetto, NumeroIntervento)

    sSQL = "UPDATE RV_POStoriaRateContratto SET"
    sSQL = sSQL & " IDOggettoCollegato=" & VarIDOggettoCollegato & ", "
    sSQL = sSQL & " Fatturata = 1"
    sSQL = sSQL & " WHERE IDRV_POStoriaRateContratto=" & IDRata

    CnDMT.Execute sSQL


    If Link_IDRataRiferimento > 0 Then
    
        sSQL = "UPDATE RV_PORateContratto SET"
        sSQL = sSQL & " IDOggettoCollegato=" & VarIDOggettoCollegato & ", "
        sSQL = sSQL & " Fatturata = 1"
        sSQL = sSQL & " WHERE IDRV_PORateContratto=" & Link_IDRataRiferimento
    
        CnDMT.Execute sSQL
    
    End If

    sSQL = "UPDATE RV_POTMPFatturazioneRate SET "
    sSQL = sSQL & "IDOggetto=" & fnNotNullN(VarIDOggettoCollegato) & " "
    sSQL = sSQL & "WHERE IDTMP=" & ID

    CnDMT.Execute sSQL

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
Private Function GET_ARTICOLO_BUONO(IDBuono As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDArticolo FROM RV_POInterventoRigheDett"
sSQL = sSQL & " WHERE IDRV_POInterventoRigheDett=" & IDBuono

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ARTICOLO_BUONO = 0
Else
    GET_ARTICOLO_BUONO = fnNotNullN(rs!IDArticolo)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_FATTURA(IDBuono As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RigaDiFatturazione FROM RV_POInterventoRigheDett"
sSQL = sSQL & " WHERE IDRV_POInterventoRigheDett=" & IDBuono

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_FATTURA = ""
Else
    GET_DESCRIZIONE_FATTURA = fnNotNull(rs!RigaDiFatturazione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Public Function fnGetEsercizio(dData As String) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "Select IDEsercizio FROM Esercizio "
    sSQL = sSQL & " WHERE IDAzienda = " & TheApp.IDFirm
    sSQL = sSQL & " AND DataInizio <=" & fnNormDate(dData)
    sSQL = sSQL & " AND DataFine >= " & fnNormDate(dData)
   

    
    Set rsEse = CnDMT.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetEsercizio = fnNotNullN(rsEse!IDEsercizio)
    Else
        fnGetEsercizio = 0
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Private Sub AGGIORNAMENTO_COLLEGAMENTO_BUONO_SINGOLO(IDAnagrafica As Long, IDOggetto As Long, IDTipoOggetto As Long, RitenutaAcconto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDTipoOggettoBuono As Long
Dim IDOggettoBuono As Long
Dim NumeroDocumentoEsteso As String



'While Not rsnew.EOF

    sSQL = "UPDATE RV_POInterventoRigheDett SET "
    sSQL = sSQL & "IDOggetto=" & IDOggetto & ", "
    sSQL = sSQL & "IDTipoOggetto=" & IDTipoOggetto & ", "
    sSQL = sSQL & "Fatturata=1 "
    sSQL = sSQL & "WHERE IDRV_POInterventoRigheDett=" & fnNotNullN(rsnew!IDRV_POInterventoRigheDett)
    
    CnDMT.Execute sSQL
    
    GET_OGGETTO_BUONO fnNotNullN(rsnew!IDRV_POInterventoRigheDett), IDOggettoBuono, IDTipoOggettoBuono
    
    If IDOggettoBuono = 0 Then
        NumeroDocumentoEsteso = GET_NUMERO_INTERVENTO(fnNotNullN(rsnew!IDRV_POintervento)) & "/" & rsnew!NumeroDocumento
        IDTipoOggettoBuono = fnGetTipoOggetto("RV_POBuonoIntervento")
        IDOggettoBuono = GET_LINK_OGGETTO(0, IDTipoOggettoBuono, NumeroDocumentoEsteso, rsnew!DataDocumento)
    End If
    
    CREA_FLUSSO_DOCUMENTALE IDTipoOggetto, IDOggetto, IDOggettoBuono, IDTipoOggettoBuono
    
'rsnew.MoveNext
'Wend

'rsnew.Filter = vbNullString


End Sub


Private Sub AGGIORNAMENTO_COLLEGAMENTO_BUONO(IDAnagrafica As Long, IDOggetto As Long, IDTipoOggetto As Long, RitenutaAcconto As Long, IDSitoPerAnagrafica As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDTipoOggettoBuono As Long
Dim IDOggettoBuono As Long
Dim NumeroDocumentoEsteso As String


rsnew.Filter = "IDAnagraficaFatturazione=" & IDAnagrafica
rsnew.Filter = rsnew.Filter & " AND RitenutaAcconto=" & RitenutaAcconto
If (FLAG_RAGGR_ALTRA_DEST = 1) Then
   rsnew.Filter = rsnew.Filter & " AND IDSitoPerAnagraficaIntervento=" & IDSitoPerAnagrafica
End If
While Not rsnew.EOF

    sSQL = "UPDATE RV_POInterventoRigheDett SET "
    sSQL = sSQL & "IDOggetto=" & IDOggetto & ", "
    sSQL = sSQL & "IDTipoOggetto=" & IDTipoOggetto & ", "
    sSQL = sSQL & "Fatturata=1 "
    sSQL = sSQL & "WHERE IDRV_POInterventoRigheDett=" & fnNotNullN(rsnew!IDRV_POInterventoRigheDett)
    
    CnDMT.Execute sSQL
    
    GET_OGGETTO_BUONO fnNotNullN(rsnew!IDRV_POInterventoRigheDett), IDOggettoBuono, IDTipoOggettoBuono
    
    If IDOggettoBuono = 0 Then
        NumeroDocumentoEsteso = GET_NUMERO_INTERVENTO(fnNotNullN(rsnew!IDRV_POintervento)) & "/" & rsnew!NumeroDocumento
        IDTipoOggettoBuono = fnGetTipoOggetto("RV_POBuonoIntervento")
        IDOggettoBuono = GET_LINK_OGGETTO(0, IDTipoOggettoBuono, NumeroDocumentoEsteso, rsnew!DataDocumento)
    End If
    
    CREA_FLUSSO_DOCUMENTALE IDTipoOggetto, IDOggetto, IDOggettoBuono, IDTipoOggettoBuono
    
rsnew.MoveNext
Wend

rsnew.Filter = vbNullString


End Sub
Private Function GET_OGGETTO_BUONO(IDRigaIntervento As Long, IDOggettoBuono As Long, IDTipoOggettoBuono As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POInterventoRigheDett "
sSQL = sSQL & " WHERE IDRV_POInterventoRigheDett=" & IDRigaIntervento
Set rs = CnDMT.OpenResultset(sSQL)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    IDOggettoBuono = 0
    IDTipoOggettoBuono = 0
Else
    IDOggettoBuono = fnNotNullN(rs!IDOggettoBuono)
    IDTipoOggettoBuono = fnNotNullN(rs!IDTipoOggettoBuono)
End If


rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_CAPTION_FATTURAZIONE(IDAnagrafica As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT IDAnagrafica, Nome, Anagrafica "
sSQL = sSQL & "FROM IERepCliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CAPTION_FATTURAZIONE = ""
Else
    GET_CAPTION_FATTURAZIONE = "Creazione fattura per il cliente " & UCase(fnNotNull(rs!Anagrafica) & " " & fnNotNull(rs!Nome))
End If

rs.CloseResultset
Set rs = Nothing
End Function

Public Sub GET_PARAMETRI_FILIALE_BUONI()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POParametriAzienda "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    LINK_TIPO_GESTIONE_BUONI = 0
    LINK_TIPO_NUMERAZIONE_BUONO = 0
Else
    LINK_TIPO_GESTIONE_BUONI = fnNotNullN(rs!IDRV_POTipoGestioneBuono)
    LINK_TIPO_NUMERAZIONE_BUONO = fnNotNullN(rs!IDRV_POTipoNumerazioneBuono)
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub GET_DETTAGLIO_RIGA_BUONO(rs As ADODB.Recordset)
Dim sSQL As String
    
    If (FLAG_SOVRAS_DESCRIZIONE = 1) Then
        ObjDoc.Field "Art_descrizione", Mid(GET_DESCRIZIONE_FATTURA(rsnew!IDRV_POInterventoRigheDett), 1, 255), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    End If
        
    ObjDoc.Field "Art_quantita_totale", fnNotNullN(rs!Quantita), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Art_prezzo_unitario_neutro", fnNotNullN(rs!ImportoUnitario), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Art_sco_in_percentuale_1", fnNotNullN(rs!Sconto1), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Art_sco_in_percentuale_2", fnNotNullN(rs!Sconto2), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Art_sco_in_percentuale_3", fnNotNullN(rs!Sconto3), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    
    If fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
        ObjDoc.Field "Link_art_IVA", fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_aliquota_IVA", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    Else
        ObjDoc.Field "Link_art_IVA", fnNotNullN(rs!IDIva), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_aliquota_IVA", fnNotNullN(rs!AliquotaIva), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
    End If
    
End Sub
Private Function GET_NUMERO_DOCUMENTI_PER_CLIENTE(IDAnagraficaCliente As Long, NumeroDocumento As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_DOCUMENTI_PER_CLIENTE = 0

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POTMPFatturazioneBuoni "
sSQL = sSQL & " WHERE IDAnagraficaCliente=" & IDAnagraficaCliente
sSQL = sSQL & " AND RitenutaAcconto=" & NumeroDocumento - 1

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_DOCUMENTI_PER_CLIENTE = False
Else
    GET_NUMERO_DOCUMENTI_PER_CLIENTE = True
End If

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Function
Private Function GET_STRINGA_BUONO(IDRigaBuono As Long) As String
On Error GoTo ERR_GET_STRINGA_BUONO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_STRINGA_BUONO = ""

sSQL = "SELECT * FROM RV_POStringaBuono "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " ORDER BY Posizione"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    If Len(Trim(fnNotNull(rs!ValoreCampo))) > 0 Then
        If fnNotNullN(rs!CarattereSpazio) = 0 Then
            GET_STRINGA_BUONO = GET_STRINGA_BUONO & fnNotNull(rs!ValoreCampo)
        Else
            GET_STRINGA_BUONO = GET_STRINGA_BUONO & " " & fnNotNull(rs!ValoreCampo)
        End If
    Else
        If fnNotNullN(rs!CarattereSpazio) = 0 Then
            GET_STRINGA_BUONO = GET_STRINGA_BUONO & GET_VALORE_CAMPO_BUONO(fnNotNull(rs!NomeCampo), IDRigaBuono)
        Else
            GET_STRINGA_BUONO = GET_STRINGA_BUONO & " " & GET_VALORE_CAMPO_BUONO(fnNotNull(rs!NomeCampo), IDRigaBuono)
        End If
    End If
rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_STRINGA_BUONO:
    GET_STRINGA_BUONO = GET_STRINGA_BUONO

End Function
Private Function GET_VALORE_CAMPO_BUONO(NomeCampo As String, IDRigaBuono As Long) As String
On Error GoTo ERR_GET_VALORE_CAMPO_IMBARCAZIONE
Dim rsVal As ADODB.Recordset
Dim sSQL As String

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_POIEBuoniIntervervento "
sSQL = sSQL & "WHERE IDRV_POInterventoRigheDett=" & IDRigaBuono

Set rsVal = New ADODB.Recordset

rsVal.Open sSQL, CnDMT.InternalConnection

If rsVal.EOF Then
    GET_VALORE_CAMPO_BUONO = ""
Else
    If (rsVal.Fields(0).Type) = adDBTimeStamp Then
        GET_VALORE_CAMPO_BUONO = GET_FORMATTA_DATA(fnNotNull(rsVal.Fields(NomeCampo).Value))
    Else
        GET_VALORE_CAMPO_BUONO = fnNotNull(rsVal.Fields(NomeCampo).Value)
    End If
End If

rsVal.Close
Set rsVal = Nothing

Exit Function
ERR_GET_VALORE_CAMPO_IMBARCAZIONE:
    GET_VALORE_CAMPO_BUONO = ""
End Function
Private Function GET_CONTROLLO_DESCRIZIONE_ART_RIGA_FATT(IDRigaDettaglio As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POIEBuoniIntervervento "
sSQL = sSQL & "WHERE IDRV_POInterventoRigheDett=" & IDRigaDettaglio

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNull(rs!Articolo) = fnNotNull(rs!AnnotazioniPerFattura) Then
        GET_CONTROLLO_DESCRIZIONE_ART_RIGA_FATT = True
    Else
        GET_CONTROLLO_DESCRIZIONE_ART_RIGA_FATT = False
    End If
End If


rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_ANAGRAFICA_INTERVENTO(IDAnagrafica As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT * FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ANAGRAFICA_INTERVENTO = ""
Else
    GET_ANAGRAFICA_INTERVENTO = "Cliente intervento: " & fnNotNull(rs!Anagrafica) & " " & fnNotNull(rs!Nome)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREA_FLUSSO_DOCUMENTALE(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long)
'On Error GoTo ERR_CREA_FLUSSO_DOCUMENTALE
Dim sSQL As String
Dim rsnew As ADODB.Recordset
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
sSQL = sSQL & "WHERE Descrizione=" & fnNormString("Vendita -> Buoni intervento")
Set rsnew = New ADODB.Recordset

rsnew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsnew.EOF Then
    rsnew.AddNew
        rsnew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFlussoGruppo")
        rsnew!Descrizione = "Vendita -> Buoni intervento"
    rsnew.Update
End If

IDFlussoGruppo = fnNotNullN(rsnew!IDFlussoGruppo)

rsnew.Close
Set rsnew = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoFunzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzioneVend
sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneRata
Set rsnew = New ADODB.Recordset

rsnew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsnew.EOF Then
    rsnew.AddNew
        rsnew!IDFlussoFunzione = fnGetNewKeyTipoOggetto("FlussoFunzione", "IDFlussoFunzione")
        rsnew!IDFunzione = IDFunzioneVend
        rsnew!IDFunzioneSuccessiva = IDFunzioneRata
        rsnew!Cardinalita = 3
        rsnew!TipoAutomatismo = 1
        rsnew!Attributo = 14
        rsnew!TipoDipendenza = 1
        rsnew!IDFlussoGruppo = IDFlussoGruppo
    rsnew.Update
End If

IDFlussoFunzione = fnNotNullN(rsnew!IDFlussoFunzione)

rsnew.Close
Set rsnew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE COLLEGATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoFunzioneCollegato "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoVend
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
Set rsnew = New ADODB.Recordset

rsnew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsnew.EOF Then
    rsnew.AddNew
        rsnew!IDFlussoFunzione = IDFlussoFunzione
        rsnew!IDOggetto = IDOggettoVend
        rsnew!IDTipoOggetto = IDTipoOggettoVend
End If

rsnew!FlussoFunzioneCollegato = 2
rsnew.Update

rsnew.Close
Set rsnew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FLUSSO OGGETTI COLLEGATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggetto=" & IDOggettoVend
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoRata

Set rsnew = New ADODB.Recordset

rsnew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsnew.EOF Then
    rsnew.AddNew
    rsnew!IDFlussoFunzione = IDFlussoFunzione
    rsnew!IDOggetto = IDOggettoVend
    rsnew!IDTipoOggetto = IDTipoOggettoVend
    rsnew!IDTipoOggettoCollegato = IDTipoOggettoRata
    rsnew!IDOggettoCollegato = IDOggettoRata
    rsnew.Update
End If

rsnew.Close
Set rsnew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Sub
ERR_CREA_FLUSSO_DOCUMENTALE:
MsgBox Err.Description, vbCritical, "CREA_FLUSSO_DOCUMENTALE"
End Sub
Private Function GET_LINK_OGGETTO(IDOggetto As Long, IDTipoOggetto As Long, NumeroBuono As String, DataBuono As String) As Long
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
        rs!IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.IDFirm)
        rs!IDSezionale = 0
        rs!Oggetto = GET_DESCRIZIONE_FUNZIONE(IDFunzione)
        rs!DataEmissione = DataBuono
        rs!Numero = NumeroBuono
        rs!DataUltimaVariazione = Date
        rs!IDUtenteUltimaVariazione = TheApp.IDUser
        rs!VirtualDelete = 0
        rs!IDOggetto = fnGetNewKey("Oggetto", "IDOggetto")
        GET_LINK_OGGETTO = rs!IDOggetto
    rs.Update
Else
    rs!DataEmissione = DataBuono
    rs!Numero = NumeroBuono
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

Private Function GET_NUMERO_INTERVENTO(IDIntervento As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AnnoIntervento, NumeroIntervento, NumeroFase "
sSQL = sSQL & "FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=" & IDIntervento

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_INTERVENTO = ""
Else
    GET_NUMERO_INTERVENTO = fnNotNullN(rs!AnnoIntervento) & "-" & fnNotNullN(rs!NumeroIntervento) & "-" & fnNotNullN(rs!NumeroFase)
End If

rs.CloseResultset
Set rs = Nothing
End Function
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
Private Function GET_DESCRIZIONE_SITO(ID As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_DESCRIZIONE_SITO = ""
sSQL = "SELECT IDSitoPerAnagrafica, SitoPerAnagrafica "
sSQL = sSQL & "FROM SitoPerAnagrafica "
sSQL = sSQL & "WHERE IDSitoPerAnagrafica=" & ID

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_DESCRIZIONE_SITO = fnNotNull(rs!SitoPerAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_FORMATTA_DATA(Data As String) As String
On Error GoTo ERR_GET_FORMATTA_DATA
Dim Anno As String
Dim Mese As String
Dim Giorno As String


If Len(Trim(Data)) = 0 Then
    GET_FORMATTA_DATA = ""
    Exit Function
End If

Anno = Year(Data)
Mese = Month(Data)
Giorno = Day(Data)

If Len(Mese) = 1 Then Mese = "0" & Mese
If Len(Giorno) = 1 Then Giorno = "0" & Giorno


GET_FORMATTA_DATA = Giorno & "/" & Mese & "/" & Anno


Exit Function
ERR_GET_FORMATTA_DATA:
    GET_FORMATTA_DATA = Data
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
    GET_ALIQUOTA_IVA = fnNotNullN(rs!AliquotaIva)
End If

rs.CloseResultset
Set rs = Nothing
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

