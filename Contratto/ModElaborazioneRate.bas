Attribute VB_Name = "ModElaborazioneRate"

Public Sub CREA_RECORDSET_RATE()

If Not (rsRateEla Is Nothing) Then
    Set rsRateEla = Nothing
End If

Set rsRateEla = New ADODB.Recordset
rsRateEla.CursorLocation = adUseClient

rsRateEla.Fields.Append "NumeroRata", adInteger, , adFldIsNullable
rsRateEla.Fields.Append "DataInizioPeriodo", adDBDate, , adFldIsNullable
rsRateEla.Fields.Append "DataFinePeriodo", adDBDate, , adFldIsNullable
rsRateEla.Fields.Append "ImportoRata", adDouble, , adFldIsNullable
rsRateEla.Fields.Append "IDRV_POProdotto", adInteger, , adFldIsNullable
rsRateEla.Fields.Append "IDRV_POContrattoProdotti", adInteger, , adFldIsNullable
rsRateEla.Fields.Append "IDRV_POContrattoAdeguamento", adInteger, , adFldIsNullable
rsRateEla.Fields.Append "IDArticolo", adInteger, , adFldIsNullable

rsRateEla.Open , , adOpenKeyset, adLockBatchOptimistic

End Sub
Public Sub Elaborazione(IDContratto As Long, DataInizioContratto As String, DataFineContratto As String, ImportoContratto As Double, numerorate As Long, CadenzaRate As Long, CalcolaPrimaRata As Long, AnnoSolare As Long, IDRV_POContrattoProdotti As Long, IDRV_POProdotto As Long, NGGPrimaRata As Long, Optional GeneraUnaSolaRata As Long = 0, Optional IDAdeguamento As Long = 0, Optional IDArticolo As Long = 0)
Dim NumeroGiorniContratto As Long
Dim ImportoGiornalieroContratto As Double
Dim ImportoRata As Double
Dim DataInizioRata As String
Dim DataFineRata As String
Dim NRata As Long
Dim ImportoRataParziale As Double


Dim TotaleRatePagate As Double
Dim DataUltimaRataPagata As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoContrattoPerRate As Double
Dim NumeroRatePagate As Long

Dim NumeroProgRata As Long


Dim Avvia As Boolean
Avvia = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT SUM(ImportoRata) AS TotaleImportoPagato, "
sSQL = sSQL & "MAX(DataRata) as DataUltRataPag, "
sSQL = sSQL & "COUNT(IDRV_PORateContratto) AS NRatePag"
sSQL = sSQL & " FROM RV_PORateContratto "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Adeguamento=0"
sSQL = sSQL & " AND Manuale=0"
sSQL = sSQL & " AND ContrattoAttuale=1"
If IDAdeguamento = 0 Then
    sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"
Else
    sSQL = sSQL & " AND IDRV_POContrattoAdeguamento=" & IDAdeguamento
End If
sSQL = sSQL & " AND Fatturata=1 "
If IDRV_POContrattoProdotti > 0 Then
    sSQL = sSQL & " AND IDRV_POContrattoProdotti=" & IDRV_POContrattoProdotti
End If

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    TotaleRatePagate = 0
    DataUltimaRataPagata = ""
    NumeroRatePagate = 0
Else
    TotaleRatePagate = fnNotNullN(rs!TotaleImportoPagato)
    DataUltimaRataPagata = fnNotNull(rs!DataUltRataPag)
    NumeroRatePagate = fnNotNullN(rs!NRatePag)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ImportoRata = 0
ImportoRataProgressiva = 0
ImportoContrattoPerRate = ImportoContratto - TotaleRatePagate

NumeroGiorniContratto = DateDiff("d", DataInizioContratto, DataFineContratto) + 1
ImportoGiornalieroContratto = FormatNumber((ImportoContratto / NumeroGiorniContratto), 2)
ImportoRata = FormatNumber((ImportoContratto / numerorate), 2)

ImportoRataTotale = 0

DataInizioRata = DataInizioContratto

If GeneraUnaSolaRata = 1 Then
    'CONTROLLO SE E' STATA PAGATA
    If ImportoContrattoPerRate <> 0 Then
        rsRateEla.AddNew
            rsRateEla!numerorata = 1
            rsRateEla!DataInizioPeriodo = DataInizioRata
            rsRateEla!DataFinePeriodo = DataFineContratto
            rsRateEla!ImportoRata = ImportoContrattoPerRate
            rsRateEla!IDRV_POContrattoProdotti = IDRV_POContrattoProdotti
            rsRateEla!IDRV_POProdotto = IDRV_POProdotto
            If IDAdeguamento > 0 Then
                rsRateEla!IDRV_POContrattoAdeguamento = IDAdeguamento
                rsRateEla!IDArticolo = IDArticolo
            End If
        rsRateEla.Update
    End If
    Exit Sub
End If

If (NGGPrimaRata > 0) Then
    DataInizioRata = DateAdd("d", NGGPrimaRata, DataInizioContratto) - 1
End If

NRata = 1
NumeroProgRata = 1

If IDAdeguamento > 0 Then
    NumeroProgRata = GET_NUMERO_PROGRESSIVO_ADEG(IDAdeguamento, NumeroProgRata)
End If

If CalcolaPrimaRata = 1 Then
    If (Day(DataInizioRata) > 1) Then
        DataInizioRata = GET_PRIMA_RATA(numerorate, CadenzaRate, CalcolaPrimaRata, AnnoSolare, ImportoGiornalieroContratto, DataInizioRata, DataFineContratto, ImportoRata, DataUltimaRataPagata, IDRV_POContrattoProdotti, IDRV_POProdotto, IDAdeguamento, IDArticolo, NumeroProgRata)
        NumeroProgRata = NumeroProgRata + 1
        NRata = NRata + 1
    End If
End If

Do While (DateDiff("d", DataInizioRata, DataFineContratto) > 0)
    DataFineRata = DateAdd("m", CadenzaRate, DataInizioRata)
    DataFineRata = DateAdd("d", -1, DataFineRata)
    
    If ((AnnoSolare = 0) And (CalcolaPrimaRata = 0)) Then
        If (NRata = numerorate) Then
            DataFineRata = DataFineContratto
        End If
    End If
    
    If AnnoSolare = 1 Then
        DataFineRata = GET_DATA_FINE_SOLARE(DataInizioRata, numerorate, CadenzaRate)
    End If
    
    If DateDiff("d", DataFineRata, DataFineContratto) <= 0 Then
        
        DataFineRata = DataFineContratto
        ImportoRataParziale = ImportoContratto - TotaleRatePagate - ImportoRataTotale 'Somma delle rate precedenti
        
        If DataUltimaRataPagata = "" Then
            Avvia = True
        Else
            If (DateDiff("d", DataInizioRata, DataUltimaRataPagata) < 0) Then
                Avvia = True
            End If
        End If
        If Avvia = True Then
            rsRateEla.AddNew
                rsRateEla!numerorata = NumeroProgRata 'NRata
                rsRateEla!DataInizioPeriodo = DataInizioRata
                rsRateEla!DataFinePeriodo = DataFineRata
                rsRateEla!ImportoRata = ImportoRataParziale
                rsRateEla!IDRV_POContrattoProdotti = IDRV_POContrattoProdotti
                rsRateEla!IDRV_POProdotto = IDRV_POProdotto
                If IDAdeguamento > 0 Then
                    rsRateEla!IDRV_POContrattoAdeguamento = IDAdeguamento
                    rsRateEla!IDArticolo = IDArticolo
                End If
            rsRateEla.Update
        End If
    Else
        If AnnoSolare = 1 Then
            ImportoRataParziale = GET_IMPORTO_PER_ANNO_SOLARE(DataInizioRata, DataFineRata, ImportoRata, CadenzaRate, ImportoGiornalieroContratto)
        Else
            ImportoRataParziale = ImportoRata
        End If
        
        If DataUltimaRataPagata = "" Then
            Avvia = True
        Else
            If (DateDiff("d", DataInizioRata, DataUltimaRataPagata) < 0) Then
                Avvia = True
            End If
        End If
        
        If Avvia = True Then
            rsRateEla.AddNew
                rsRateEla!numerorata = NumeroProgRata 'NRata
                rsRateEla!DataInizioPeriodo = DataInizioRata
                rsRateEla!DataFinePeriodo = DataFineRata
                rsRateEla!ImportoRata = ImportoRataParziale
                rsRateEla!IDRV_POContrattoProdotti = IDRV_POContrattoProdotti
                rsRateEla!IDRV_POProdotto = IDRV_POProdotto
                If IDAdeguamento > 0 Then
                    rsRateEla!IDRV_POContrattoAdeguamento = IDAdeguamento
                    rsRateEla!IDArticolo = IDArticolo
                End If
            rsRateEla.Update
        End If
        
    End If
    
    If Avvia = True Then
        ImportoRataTotale = ImportoRataTotale + FormatNumber(ImportoRataParziale, 2)
    End If
    DataInizioRata = DateAdd("d", 1, DataFineRata)
    NRata = NRata + 1
    NumeroProgRata = NumeroProgRata + 1
Loop

End Sub
Private Function GET_PRIMA_RATA(numerorate As Long, CadenzaRate As Long, CalcolaPrimaRata As Long, AnnoSolare As Long, importogiornaliero As Double, DataInizioContratto As String, DataFineContratto As String, ImportoRata As Double, DataUltimaRataPagata As String, IDRV_POContrattoProdotti As Long, IDRV_POProdotto As Long, Optional IDAdeguamento As Long = 0, Optional IDArticolo As Long, Optional numero As Long = 1) As String
Dim NumeroGiorniPrimaRata As Long
Dim DataFineRata As String
Dim DataInizioRata As String
Dim ImportoRataParziale As Double
Dim Avvia As Boolean

DataInizioRata = DataInizioContratto

Avvia = False
If CalcolaPrimaRata = 1 Then
    NumeroGiorniPrimaRata = GET_NUMERO_GIORNI(DataInizioContratto, DataFineRata)
    ImportoRataParziale = FormatNumber((NumeroGiorniPrimaRata * importogiornaliero), 2)
    
    If DataUltimaRataPagata = "" Then
        Avvia = True
    Else
        If (DateDiff("d", DataInizioRata, DataUltimaRataPagata) < 0) Then
            Avvia = True
        End If
    End If
    
    If Avvia = True Then
        rsRateEla.AddNew
            rsRateEla!numerorata = numero
            rsRateEla!DataInizioPeriodo = DataInizioRata
            rsRateEla!DataFinePeriodo = DataFineRata
            rsRateEla!ImportoRata = ImportoRataParziale
            rsRateEla!IDRV_POContrattoProdotti = IDRV_POContrattoProdotti
            rsRateEla!IDRV_POProdotto = IDRV_POProdotto
            If IDAdeguamento > 0 Then
                rsRateEla!IDRV_POContrattoAdeguamento = IDAdeguamento
                rsRateEla!IDArticolo = IDArticolo
            End If
        rsRateEla.Update
    End If
    
    DataInizioRata = DateAdd("d", 1, DataFineRata)
    
    If Avvia = True Then
        ImportoRataTotale = ImportoRataTotale + ImportoRataParziale
    End If
    
End If

GET_PRIMA_RATA = DataInizioRata

End Function
Private Function GET_NUMERO_GIORNI(DataInizio As String, DataFine) As Long
Dim mese As Long
Dim Giorno As Long


Select Case Month(DataInizio)
    Case 1, 3, 5, 7, 8, 10, 12
        Giorno = 31
    Case 2
        If ((Year(DataInizio) Mod 4) = 0) Then
            Giorno = 29
        Else
            Giorno = 28
        End If
    Case 4, 6, 9, 11
        Giorno = 30
End Select

mese = Month(DataInizio)
If Len(mese) = 1 Then mese = "0" & mese
DataFine = Giorno & "/" & mese & "/" & Year(DataInizio)

GET_NUMERO_GIORNI = DateDiff("d", DataInizio, DataFine) + 1

End Function


'Private Function GET_NUMERO_GIORNI_ANNO_SOLARE(DataInizio As String, numerorate As Long) As Long
'Dim mese As Long
'Dim Giorno As Long
'Dim DataFine As String
'
'DataFine = GET_DATA_FINE_SOLARE(DataInizio, numerorate)
'
'GET_NUMERO_GIORNI_ANNO_SOLARE = DateDiff("d", DataInizio, DataFine) + 1
'
'
'End Function

Private Function GET_DATA_FINE_MESE(MeseSel As Long, Anno As Long) As String
Dim Giorno As Long
Dim DataFine As String
Dim mese As String

Select Case MeseSel
    Case 1, 3, 5, 7, 8, 10, 12
        Giorno = 31
    Case 2
        If (Anno Mod 4) = 0 Then
            Giorno = 29
        Else
            Giorno = 28
        End If
    Case 4, 6, 9, 11
    
        Giorno = 30
End Select

mese = MeseSel

If Len(mese) = 1 Then mese = "0" & mese

GET_DATA_FINE_MESE = Giorno & "/" & mese & "/" & Anno


End Function
Private Function GET_DATA_FINE_SOLARE(DatainizioPer As String, numerorate As Long, CadenzaRate As Long) As String
'Dim CadenzaRate As Long
Dim DataInizioSolare As String
'CadenzaRate = 12 / numerorate
Dim MeseInizio As Long
Dim DataFinePer As String
Dim Anno As Long
MeseInizio = Month(DatainizioPer)
Anno = Year(DatainizioPer)

Select Case CadenzaRate
    Case 1
        GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(Month(DatainizioPer), Anno)
    Case 2
        Select Case MeseInizio
            Case 1, 2
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(2, Anno)
            Case 3, 4
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(4, Anno)
            Case 5, 6
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(6, Anno)
            Case 7, 8
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(8, Anno)
            Case 9, 10
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(10, Anno)
            Case 11, 12
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(12, Anno)
        End Select
    Case 3
        Select Case MeseInizio
        
            Case 1, 2, 3
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(3, Anno)
            Case 4, 5, 6
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(6, Anno)
            Case 7, 8, 9
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(9, Anno)
            Case 10, 11, 12
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(12, Anno)
        End Select
    Case 4
        Select Case MeseInizio
        
            Case 1, 2, 3, 4
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(4, Anno)
            Case 5, 6, 7, 8
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(8, Anno)
            Case 9, 10, 11, 12
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(12, Anno)
                
        End Select
    Case 6
        Select Case MeseInizio
        
            Case 1, 2, 3, 4, 5, 6
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(6, Anno)
            Case 7, 8, 9, 10, 11, 12
                GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(12, Anno)
        End Select
    Case 12
        GET_DATA_FINE_SOLARE = GET_DATA_FINE_MESE(12, Anno)
End Select


End Function
Private Function GET_IMPORTO_PER_ANNO_SOLARE(DataInizioRata As String, DataFineRata As String, ImportoRata As Double, CadenzaRate As Long, ImpGiornoContratto As Double) As Double
Dim DataInizio As String
Dim DataFine As String
Dim Importo As Double
Dim mese As String
Dim MeseInizio As Long
Dim Anno As Long
Dim DiffGiorniUfficiali As Long
Dim DiffGiorniRata As Long
Dim ImportoRataPeriodo As Double
Dim GiornoInizio As Long
Dim NumeroGiorniParziali As Long


MeseInizio = Month(DataFineRata)
Anno = Year(DataFineRata)
mese = MeseInizio
GiornoInizio = Day(DataInizioRata)

If Len(mese) = 1 Then mese = "0" & mese


Select Case CadenzaRate
    Case 1
        DataInizio = "01/" & mese & "/" & Anno
        ImportoRataPeriodo = ImportoRata
    Case 2
        Select Case MeseInizio
            Case 1, 2
                DataInizio = "01/01/" & Anno
            Case 3, 4
                DataInizio = "01/03/" & Anno
            Case 5, 6
                DataInizio = "01/05/" & Anno
            Case 7, 8
                DataInizio = "01/07/" & Anno
            Case 9, 10
                DataInizio = "01/09/" & Anno
            Case 11, 12
                DataInizio = "01/11/" & Anno
        End Select
        ImportoRataPeriodo = ImportoRata / 2
    Case 3
        Select Case MeseInizio
        
            Case 1, 2, 3
                DataInizio = "01/01/" & Anno
            Case 4, 5, 6
                DataInizio = "01/04/" & Anno
            Case 7, 8, 9
                DataInizio = "01/07/" & Anno
            Case 10, 11, 12
                DataInizio = "01/10/" & Anno
        End Select
        
        ImportoRataPeriodo = ImportoRata / 3
    Case 4
        Select Case MeseInizio
        
            Case 1, 2, 3, 4
                DataInizio = "01/01/" & Anno
            Case 5, 6, 7, 8
                DataInizio = "01/05/" & Anno
            Case 9, 10, 11, 12
                DataInizio = "01/09/" & Anno
                
        End Select
        ImportoRataPeriodo = ImportoRata / 4
    Case 6
        Select Case MeseInizio
        
            Case 1, 2, 3, 4, 5, 6
                DataInizio = "01/01/" & Anno
            Case 7, 8, 9, 10, 11, 12
                DataInizio = "01/07/" & Anno
        End Select
        
        ImportoRataPeriodo = ImportoRata / 6
    Case 12
        DataInizio = "01/12/" & Anno
        ImportoRataPeriodo = ImportoRata / 12
End Select

DiffGiorniUfficiali = DateDiff("d", DataInizio, DataFineRata)
DiffGiorniRata = DateDiff("d", DataInizioRata, DataFineRata)


If DiffGiorniUfficiali = DiffGiorniRata Then
    GET_IMPORTO_PER_ANNO_SOLARE = ImportoRata
Else
    If (GiornoInizio = 1) Then
        
        GET_IMPORTO_PER_ANNO_SOLARE = GET_IMPORTO_PARZIALE_SOLARE(Month(DataInizioRata), CadenzaRate, ImportoRataPeriodo)
    Else
        NumeroGiorniParziali = DateDiff("d", DataInizioRata, DataFineRata) + 1
        GET_IMPORTO_PER_ANNO_SOLARE = FormatNumber((NumeroGiorniParziali * ImpGiornoContratto), 2)
    End If
End If


End Function
Private Function GET_IMPORTO_PARZIALE_SOLARE(MeseInizio As Long, CadenzaRate As Long, ImportoRataPeriodo As Double) As Double
Dim MeseNelPeriodo As Long
Dim Periodo As Long

Select Case MeseInizio
    Case 1
        MeseNelPeriodo = 1
    Case 2
        Select Case CadenzaRate
            Case 1
                MeseNelPeriodo = 1
            Case 2
                MeseNelPeriodo = 2
            Case 3
                MeseNelPeriodo = 2
            Case 4
                MeseNelPeriodo = 2
            Case 6
                MeseNelPeriodo = 2
            Case 12
                MeseNelPeriodo = 2
            
        End Select
            
    Case 3
        Select Case CadenzaRate
            Case 1
                MeseNelPeriodo = 1
            Case 2
                MeseNelPeriodo = 1
            Case 3
                MeseNelPeriodo = 3
            Case 4
                MeseNelPeriodo = 3
            Case 6
                MeseNelPeriodo = 3
            Case 12
                MeseNelPeriodo = 3
            
        End Select
    Case 4
        Select Case CadenzaRate
            Case 1
                MeseNelPeriodo = 1
            Case 2
                MeseNelPeriodo = 2
            Case 3
                MeseNelPeriodo = 1
            Case 4
                MeseNelPeriodo = 4
            Case 6
                MeseNelPeriodo = 4
            Case 12
                MeseNelPeriodo = 4
            
        End Select
    Case 5
        Select Case CadenzaRate
            Case 1
                MeseNelPeriodo = 1
            Case 2
                MeseNelPeriodo = 1
            Case 3
                MeseNelPeriodo = 2
            Case 4
                MeseNelPeriodo = 1
            Case 6
                MeseNelPeriodo = 5
            Case 12
                MeseNelPeriodo = 5
            
        End Select
    Case 6
        Select Case CadenzaRate
            Case 1
                MeseNelPeriodo = 1
            Case 2
                MeseNelPeriodo = 2
            Case 3
                MeseNelPeriodo = 3
            Case 4
                MeseNelPeriodo = 2
            Case 6
                MeseNelPeriodo = 6
            Case 12
                MeseNelPeriodo = 6
            
        End Select
    Case 7
        Select Case CadenzaRate
            Case 1
                MeseNelPeriodo = 1
            Case 2
                MeseNelPeriodo = 1
            Case 3
                MeseNelPeriodo = 1
            Case 4
                MeseNelPeriodo = 3
            Case 6
                MeseNelPeriodo = 1
            Case 12
                MeseNelPeriodo = 7
            
        End Select
    Case 8
        Select Case CadenzaRate
            Case 1
                MeseNelPeriodo = 1
            Case 2
                MeseNelPeriodo = 2
            Case 3
                MeseNelPeriodo = 2
            Case 4
                MeseNelPeriodo = 4
            Case 6
                MeseNelPeriodo = 2
            Case 12
                MeseNelPeriodo = 8
            
        End Select
    Case 9
        Select Case CadenzaRate
            Case 1
                MeseNelPeriodo = 1
            Case 2
                MeseNelPeriodo = 1
            Case 3
                MeseNelPeriodo = 3
            Case 4
                MeseNelPeriodo = 1
            Case 6
                MeseNelPeriodo = 3
            Case 12
                MeseNelPeriodo = 9
            
        End Select
    Case 10
        Select Case CadenzaRate
            Case 1
                MeseNelPeriodo = 1
            Case 2
                MeseNelPeriodo = 2
            Case 3
                MeseNelPeriodo = 1
            Case 4
                MeseNelPeriodo = 2
            Case 6
                MeseNelPeriodo = 4
            Case 12
                MeseNelPeriodo = 10
            
        End Select
    Case 11
        Select Case CadenzaRate
            Case 1
                MeseNelPeriodo = 1
            Case 2
                MeseNelPeriodo = 1
            Case 3
                MeseNelPeriodo = 2
            Case 4
                MeseNelPeriodo = 3
            Case 6
                MeseNelPeriodo = 5
            Case 12
                MeseNelPeriodo = 11
            
        End Select
    Case 12
        Select Case CadenzaRate
            Case 1
                MeseNelPeriodo = 1
            Case 2
                MeseNelPeriodo = 2
            Case 3
                MeseNelPeriodo = 3
            Case 4
                MeseNelPeriodo = 4
            Case 6
                MeseNelPeriodo = 6
            Case 12
                MeseNelPeriodo = 12
            
        End Select

End Select


Periodo = (CadenzaRate - MeseNelPeriodo) + 1


GET_IMPORTO_PARZIALE_SOLARE = FormatNumber((Periodo * ImportoRataPeriodo), 2)

End Function


Public Sub ElaborazionePerRigaProdotto(IDContratto As Long, DataInizioContratto As String, DataFineContratto As String, ImportoContratto As Double, numerorate As Long, CadenzaRate As Long, CalcolaPrimaRata As Long, AnnoSolare As Long, IDRV_POContrattoProdotti As Long, IDRV_POProdotto As Long, NGGPrimaRata As Long, IDRateizzazione As Long)
Dim NumeroGiorniContratto As Long
Dim ImportoGiornalieroContratto As Double
Dim ImportoRata As Double
Dim DataInizioRata As String
Dim DataFineRata As String
Dim NRata As Long
Dim ImportoRataParziale As Double


Dim TotaleRatePagate As Double
Dim DataUltimaRataPagata As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoContrattoPerRate As Double
Dim NumeroRatePagate As Long

Dim Avvia As Boolean
Avvia = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'sSQL = "SELECT SUM(ImportoRata) AS TotaleImportoPagato, "
'sSQL = sSQL & "MAX(DataRata) as DataUltRataPag, "
'sSQL = sSQL & "COUNT(IDRV_PORateContratto) AS NRatePag"
'sSQL = sSQL & " FROM RV_PORateContratto "
'sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
'sSQL = sSQL & " AND Adeguamento=0"
'sSQL = sSQL & " AND Manuale=0"
'sSQL = sSQL & " AND ContrattoAttuale=1"
'sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"
'sSQL = sSQL & " AND Fatturata=1 "
'If IDRV_POContrattoProdotti > 0 Then
'    sSQL = sSQL & " AND IDRV_POContrattoProdotti=" & IDRV_POContrattoProdotti
'End If
'
'Set rs = Cn.OpenResultset(sSQL)
'
'If rs.EOF Then
    TotaleRatePagate = 0
    DataUltimaRataPagata = ""
    NumeroRatePagate = 0
'Else
'    TotaleRatePagate = fnNotNullN(rs!TotaleImportoPagato)
'    DataUltimaRataPagata = fnNotNull(rs!DataUltRataPag)
'    NumeroRatePagate = fnNotNullN(rs!NRatePag)
'End If
'
'rs.CloseResultset
'Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ImportoRata = 0
ImportoRataProgressiva = 0
ImportoContrattoPerRate = ImportoContratto - TotaleRatePagate

NumeroGiorniContratto = DateDiff("d", DataInizioContratto, DataFineContratto) + 1
ImportoGiornalieroContratto = FormatNumber((ImportoContratto / NumeroGiorniContratto), 2)
ImportoRata = FormatNumber((ImportoContratto / numerorate), 2)

ImportoRataTotale = 0

DataInizioRata = DataInizioContratto

If (NGGPrimaRata > 0) Then
    DataInizioRata = DateAdd("d", NGGPrimaRata, DataInizioContratto) - 1
End If

NRata = 1

If CalcolaPrimaRata = 1 Then
    If (Day(DataInizioRata) > 1) Then
        DataInizioRata = GET_PRIMA_RATA(numerorate, CadenzaRate, CalcolaPrimaRata, AnnoSolare, ImportoGiornalieroContratto, DataInizioRata, DataFineContratto, ImportoRata, DataUltimaRataPagata, IDRV_POContrattoProdotti, IDRV_POProdotto)
        NRata = 2
    End If
End If

Do While (DateDiff("d", DataInizioRata, DataFineContratto) > 0)
    DataFineRata = DateAdd("m", CadenzaRate, DataInizioRata)
    DataFineRata = DateAdd("d", -1, DataFineRata)
    
    If ((AnnoSolare = 0) And (CalcolaPrimaRata = 0)) Then
        If (NRata = numerorate) Then
            DataFineRata = DataFineContratto
        End If
    End If
    
    If AnnoSolare = 1 Then
        DataFineRata = GET_DATA_FINE_SOLARE(DataInizioRata, numerorate, CadenzaRate)
    End If
    
    If DateDiff("d", DataFineRata, DataFineContratto) <= 0 Then
        
        DataFineRata = DataFineContratto
        ImportoRataParziale = ImportoContratto - TotaleRatePagate - ImportoRataTotale 'Somma delle rate precedenti
        
        If DataUltimaRataPagata = "" Then
            Avvia = True
        Else
            If (DateDiff("d", DataInizioRata, DataUltimaRataPagata) < 0) Then
                Avvia = True
            End If
        End If
        If Avvia = True Then
            rsRateEla.AddNew
                rsRateEla!numerorata = NRata
                rsRateEla!DataInizioPeriodo = DataInizioRata
                rsRateEla!DataFinePeriodo = DataFineRata
                rsRateEla!ImportoRata = ImportoRataParziale
                rsRateEla!IDRV_POContrattoProdotti = IDRV_POContrattoProdotti
                rsRateEla!IDRV_POProdotto = IDRV_POProdotto
            rsRateEla.Update
        End If
        
    Else
        If AnnoSolare = 1 Then
            ImportoRataParziale = GET_IMPORTO_PER_ANNO_SOLARE(DataInizioRata, DataFineRata, ImportoRata, CadenzaRate, ImportoGiornalieroContratto)
        Else
            ImportoRataParziale = ImportoRata
        End If
        
        If DataUltimaRataPagata = "" Then
            Avvia = True
        Else
            If (DateDiff("d", DataInizioRata, DataUltimaRataPagata) < 0) Then
                Avvia = True
            End If
        End If
        
        If Avvia = True Then
            rsRateEla.AddNew
                rsRateEla!numerorata = NRata
                rsRateEla!DataInizioPeriodo = DataInizioRata
                rsRateEla!DataFinePeriodo = DataFineRata
                rsRateEla!ImportoRata = ImportoRataParziale
                rsRateEla!IDRV_POContrattoProdotti = IDRV_POContrattoProdotti
                rsRateEla!IDRV_POProdotto = IDRV_POProdotto
            rsRateEla.Update
        End If
        
    End If
    
    If Avvia = True Then
        ImportoRataTotale = ImportoRataTotale + FormatNumber(ImportoRataParziale, 2)
    End If
    DataInizioRata = DateAdd("d", 1, DataFineRata)
    NRata = NRata + 1
    
Loop

End Sub

Public Sub ElaborazionePerProdotto(IDContratto As Long, DataInizioContratto As String, DataFineContratto As String, ImportoContratto As Double, numerorate As Long, CadenzaRate As Long, CalcolaPrimaRata As Long, AnnoSolare As Long, IDRV_POContrattoProdotti As Long, IDRV_POProdotto As Long, NGGPrimaRata As Long, PercPrimaRata As Double, DueRata As Long, ConsideraQuantitaPeriodo As Long, QuantitaPeriodo As Double, IDTipoPeriodo As Long)
Dim NumeroGiorniContratto As Long
Dim ImportoGiornalieroContratto As Double
Dim ImportoRata As Double
Dim DataInizioRata As String
Dim DataFineRata As String
Dim NRata As Long
Dim ImportoRataParziale As Double
Dim ImportoPrimaRata As Double
Dim CalcolaPrimaRataPerc As Long

Dim TotaleRatePagate As Double
Dim DataUltimaRataPagata As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoContrattoPerRate As Double
Dim NumeroRatePagate As Long

Dim Avvia As Boolean
Avvia = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT SUM(ImportoRata) AS TotaleImportoPagato, "
sSQL = sSQL & "MAX(DataRata) as DataUltRataPag, "
sSQL = sSQL & "COUNT(IDRV_PORateContratto) AS NRatePag"
sSQL = sSQL & " FROM RV_PORateContratto "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Adeguamento=0"
sSQL = sSQL & " AND Manuale=0"
sSQL = sSQL & " AND ContrattoAttuale=1"
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"
sSQL = sSQL & " AND Fatturata=1 "
If IDRV_POContrattoProdotti > 0 Then
    sSQL = sSQL & " AND IDRV_POContrattoProdotti=" & IDRV_POContrattoProdotti
End If

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    TotaleRatePagate = 0
    DataUltimaRataPagata = ""
    NumeroRatePagate = 0
Else
    TotaleRatePagate = fnNotNullN(rs!TotaleImportoPagato)
    DataUltimaRataPagata = fnNotNull(rs!DataUltRataPag)
    NumeroRatePagate = fnNotNullN(rs!NRatePag)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ImportoRata = 0
ImportoRataProgressiva = 0
ImportoContrattoPerRate = ImportoContratto - TotaleRatePagate

NumeroGiorniContratto = DateDiff("d", DataInizioContratto, DataFineContratto) + 1
ImportoGiornalieroContratto = FormatNumber((ImportoContratto / NumeroGiorniContratto), 2)

If PercPrimaRata > 0 Then
    ImportoPrimaRata = (ImportoContrattoPerRate / 100) * PercPrimaRata
    
    ImportoRata = ImportoContrattoPerRate - ImportoPrimaRata
    
    If ConsideraQuantitaPeriodo = 0 Then
        ImportoRata = FormatNumber((ImportoRata / (numerorate - 1)), 2)
    Else
        ImportoRata = FormatNumber((ImportoRata / (QuantitaPeriodo - 1)), 2)
    End If

Else
    If ConsideraQuantitaPeriodo = 0 Then
        ImportoRata = FormatNumber((ImportoContrattoPerRate / numerorate), 2)
    Else
        ImportoRata = FormatNumber((ImportoContrattoPerRate / QuantitaPeriodo), 2)
    End If
End If

ImportoRataTotale = 0
CalcolaPrimaRataPerc = 0

DataInizioRata = DataInizioContratto

NRata = 1

If DueRata = 0 Then
    If CalcolaPrimaRata = 1 Then
        If (Day(DataInizioRata) > 1) Then
            DataInizioRata = GET_PRIMA_RATA(numerorate, CadenzaRate, CalcolaPrimaRata, AnnoSolare, ImportoGiornalieroContratto, DataInizioRata, DataFineContratto, ImportoRata, DataUltimaRataPagata, IDRV_POContrattoProdotti, IDRV_POProdotto)
            NRata = 2
        End If
    End If
End If

If DueRata = 0 Then
    Do While (DateDiff("d", DataInizioRata, DataFineContratto) >= 0)
        If ConsideraQuantitaPeriodo = 0 Then
            DataFineRata = DateAdd("m", CadenzaRate, DataInizioRata)
            DataFineRata = DateAdd("d", -1, DataFineRata)
        Else
            Select Case IDTipoPeriodo
                Case 2 'Giorni
                    DataFineRata = DateAdd("d", 1, DataInizioRata)
                    DataFineRata = DateAdd("d", -1, DataFineRata)
                Case 3 'Settimane
                    DataFineRata = DateAdd("ww", 1, DataInizioRata)
                    DataFineRata = DateAdd("d", -1, DataInizioRata)
                Case 4 'Mese
                    DataFineRata = DateAdd("m", 1, DataInizioRata)
                    DataFineRata = DateAdd("d", -1, DataFineRata)
                Case Else
                    DataFineRata = DateAdd("m", 1, DataInizioRata)
                    DataFineRata = DateAdd("d", -1, DataFineRata)
                
            End Select
        End If
        
        If ((AnnoSolare = 0) And (CalcolaPrimaRata = 0)) Then
            If ConsideraQuantitaPeriodo = 0 Then
                If (NRata = numerorate) Then
                    DataFineRata = DataFineContratto
                End If
            End If
        End If
        
        If AnnoSolare = 1 Then
            DataFineRata = GET_DATA_FINE_SOLARE(DataInizioRata, numerorate, CadenzaRate)
        End If
        
        If DateDiff("d", DataFineRata, DataFineContratto) <= 0 Then
            
            DataFineRata = DataFineContratto
            ImportoRataParziale = ImportoContratto - TotaleRatePagate - ImportoRataTotale 'Somma delle rate precedenti
            
            If DataUltimaRataPagata = "" Then
                Avvia = True
            Else
                If (DateDiff("d", DataInizioRata, DataUltimaRataPagata) < 0) Then
                    Avvia = True
                End If
            End If
            If Avvia = True Then
                rsRateEla.AddNew
                    rsRateEla!numerorata = NRata
                    rsRateEla!DataInizioPeriodo = DataInizioRata
                    rsRateEla!DataFinePeriodo = DataFineRata
                    rsRateEla!ImportoRata = ImportoRataParziale
                    rsRateEla!IDRV_POContrattoProdotti = IDRV_POContrattoProdotti
                    rsRateEla!IDRV_POProdotto = IDRV_POProdotto
                rsRateEla.Update
            End If
        Else
            If (PercPrimaRata > 0) And (CalcolaPrimaRataPerc = 0) Then
                If AnnoSolare = 1 Then
                    ImportoRataParziale = GET_IMPORTO_PER_ANNO_SOLARE(DataInizioRata, DataFineRata, ImportoPrimaRata, CadenzaRate, ImportoGiornalieroContratto)
                Else
                    ImportoRataParziale = ImportoPrimaRata
                End If
                
                CalcolaPrimaRataPerc = 1
            Else
                If AnnoSolare = 1 Then
                    ImportoRataParziale = GET_IMPORTO_PER_ANNO_SOLARE(DataInizioRata, DataFineRata, ImportoRata, CadenzaRate, ImportoGiornalieroContratto)
                Else
                    ImportoRataParziale = ImportoRata
                End If
            End If
            If DataUltimaRataPagata = "" Then
                Avvia = True
            Else
                If (DateDiff("d", DataInizioRata, DataUltimaRataPagata) < 0) Then
                    Avvia = True
                End If
            End If
            
            If Avvia = True Then
                rsRateEla.AddNew
                    rsRateEla!numerorata = NRata
                    rsRateEla!DataInizioPeriodo = DataInizioRata
                    rsRateEla!DataFinePeriodo = DataFineRata
                    rsRateEla!ImportoRata = ImportoRataParziale
                    rsRateEla!IDRV_POContrattoProdotti = IDRV_POContrattoProdotti
                    rsRateEla!IDRV_POProdotto = IDRV_POProdotto
                rsRateEla.Update
            End If
            
        End If
        
        If Avvia = True Then
            ImportoRataTotale = ImportoRataTotale + FormatNumber(ImportoRataParziale, 2)
        End If
        DataInizioRata = DateAdd("d", 1, DataFineRata)
        NRata = NRata + 1
        
    Loop
End If

If DueRata = 1 Then
    
    If PercPrimaRata = 0 Then
        ImportoRataParziale = FormatNumber((ImportoContrattoPerRate / 2), 2)
    Else
        ImportoRataParziale = ImportoPrimaRata
    End If
    
    rsRateEla.AddNew
        rsRateEla!numerorata = 1
        rsRateEla!DataInizioPeriodo = DataInizioRata
        rsRateEla!DataFinePeriodo = DataInizioRata
        rsRateEla!ImportoRata = ImportoRataParziale
        rsRateEla!IDRV_POContrattoProdotti = IDRV_POContrattoProdotti
        rsRateEla!IDRV_POProdotto = IDRV_POProdotto
    rsRateEla.Update

    rsRateEla.AddNew
        rsRateEla!numerorata = 2
        rsRateEla!DataInizioPeriodo = DataFineContratto
        rsRateEla!DataFinePeriodo = DataFineContratto
        rsRateEla!ImportoRata = ImportoContrattoPerRate - ImportoRataParziale
        rsRateEla!IDRV_POContrattoProdotti = IDRV_POContrattoProdotti
        rsRateEla!IDRV_POProdotto = IDRV_POProdotto
    rsRateEla.Update
    
End If

End Sub

Private Function GET_NUMERO_PROGRESSIVO_ADEG(IDAdeguamento As Long, numerorata As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Ris As Long

Ris = numerorata

sSQL = "SELECT IDRV_POContrattoAdeguamento, NumeroPartenza "
sSQL = sSQL & "FROM RV_POContrattoAdeguamento "
sSQL = sSQL & " WHERE IDRV_POContrattoAdeguamento=" & IDAdeguamento

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!NumeroPartenza) > 0 Then
        Ris = fnNotNullN(rs!NumeroPartenza)
    End If
End If

rs.CloseResultset
Set rs = Nothing

GET_NUMERO_PROGRESSIVO_ADEG = Ris

End Function
