Attribute VB_Name = "ModAltriDati"
Public IDClassContratto As Long
Public DescrClassContratto As String

Public IDRaggrFattContratto As Long
Public DescrRaggrFattContratto As String

Public IDClienteFatturazione As Long
Public ClienteFatturazione As String

Public IDContrattoBancario As Long
Public DescrContrattoBancario As String

Public IDAccordoCommerciale As Long
Public DescrAccordoCommerciale As String

Public IDArticoloContratto As Long
Public ArticoloContratto As String

Public IDPDCContratto As Long
Public CodicePDCContratto As String
Public DescrPDCContratto As String

Public NumeroContratto As Long
Public AnnoContratto As Long
Public NumeroRinnovo As Long
Public IDIstatContratto As Long
Public DescrIstaContratto  As String

Public MaggIstatContratto As Double
Public NumeroLicenze As Long

Public RapprLegaleAzienda As String
Public RuoloRapprAzienda As String
Public RapprLegaleCliente As String
Public RuoloRapprCliente As String

Public IDUtenteInserimento As Long
Public DescrUtenteInserimento As String
Public DataInsContratto As String
Public IDUtenteUltMod As Long
Public DescrUtenteUltMod As String
Public DataUltMod As String

Public CONFERMA_MODIFICA As Long

Public Sub InitVariabiliAltriDati(m_Document As DmtDocManLib.DBFormDocument)

IDClassContratto = 0
DescrClassContratto = ""

IDRaggrFattContratto = 0
DescrRaggrFattContratto = ""

IDClienteFatturazione = 0
ClienteFatturazione = ""

IDContrattoBancario = 0
DescrContrattoBancario = ""

IDAccordoCommerciale = 0
DescrAccordoCommerciale = ""

IDArticoloContratto = 0
ArticoloContratto = ""

IDPDCContratto = 0
CodicePDCContratto = ""
DescrPDCContratto = ""

IDIstatContratto = 0
DescrIstaContratto = ""

MaggIstatContratto = 0
NumeroLicenze = 0

RapprLegaleAzienda = ""
RuoloRapprAzienda = ""
RapprLegaleCliente = ""
RuoloRapprCliente = ""

IDUtenteInserimento = TheApp.IDUser
DataInsContratto = Date

IDUtenteUltMod = TheApp.IDUser
DataUltMod = Date

If ((m_Document.EOF) And (m_Document.BOF)) Then

    IDUtenteInserimento = TheApp.IDUser
    DataInsContratto = Date
    DescrUtenteInserimento = GET_DESCRIZIONE_UTENTE(IDUtenteInserimento)
    
    IDUtenteUltMod = TheApp.IDUser
    DataUltMod = Date
    DescrUtenteUltMod = GET_DESCRIZIONE_UTENTE(IDUtenteInserimento)

Else
    
    IDClassContratto = fnNotNullN(m_Document("IDRV_POTipoClassificazioneContratto").Value)
    DescrClassContratto = fnNotNull(m_Document("TipoClassificazioneContratto").Value)
    
    IDRaggrFattContratto = fnNotNullN(m_Document("IDRaggruppamentoFatturato").Value)
    DescrRaggrFattContratto = fnNotNull(m_Document("RaggruppamentoFatturato").Value)
    
    IDClienteFatturazione = fnNotNullN(m_Document("IDAnagraficaFatturazione").Value)
    ClienteFatturazione = fnNotNull(m_Document("AnagraficaFatturazione").Value) & " " & fnNotNull(m_Document("NomeAnagraficaFatturazione").Value)
    
    IDContrattoBancario = fnNotNullN(m_Document("IDContrattoBancario").Value)
    DescrContrattoBancario = fnNotNull(m_Document("BancaPerAnagrafica").Value)
    
    IDAccordoCommerciale = fnNotNullN(m_Document("IDAccordoCommerciale").Value)
    DescrAccordoCommerciale = fnNotNull(m_Document("DescrizioneAccordoCommerciale").Value)
    
    IDArticoloContratto = fnNotNullN(m_Document("IDArticoloContratto").Value)
    ArticoloContratto = fnNotNull(m_Document("CodiceArticoloContratto").Value) & " - " & fnNotNull(m_Document("ArticoloContratto").Value)
    
    IDPDCContratto = fnNotNullN(m_Document("IDCodiceConto").Value)
    CodicePDCContratto = fnNotNull(m_Document("CodiceConto").Value)
    DescrPDCContratto = fnNotNull(m_Document("DescrizioneConto").Value)
    
    IDIstatContratto = fnNotNullN(m_Document("IDIstat").Value)
    DescrIstaContratto = fnNotNull(m_Document("Istat").Value)
    
    MaggIstatContratto = fnNotNullN(m_Document("Maggiorazione").Value)
    NumeroLicenze = fnNotNullN(m_Document("NumeroLicenze").Value)
    
    RapprLegaleAzienda = fnNotNull(m_Document("RiferimentoAzienda").Value)
    RuoloRapprAzienda = fnNotNull(m_Document("RuoloRifAzienda").Value)
    RapprLegaleCliente = fnNotNull(m_Document("RiferimentoCliente").Value)
    RuoloRapprCliente = fnNotNull(m_Document("RuoloRifCliente").Value)
    
    IDUtenteInserimento = fnNotNullN(m_Document("IDUtentePerInserimento").Value)
    DataInsContratto = fnNotNull(m_Document("DataInserimento").Value)
    DescrUtenteInserimento = fnNotNull(m_Document("UtentePerInserimento").Value)
    
    IDUtenteUltMod = fnNotNullN(m_Document("IDUtentePerModifica").Value)
    DataUltMod = fnNotNull(m_Document("DataModifica").Value)
    DescrUtenteUltMod = fnNotNull(m_Document("UtentePerModifica").Value)
End If
End Sub

Public Sub CONFERMA_MOD_ALTRI_DATI()

With frmAltriDati

    IDClassContratto = .cboClassificazioneContratto.CurrentID
    DescrClassContratto = .cboClassificazioneContratto.Text
    
    IDRaggrFattContratto = .cboRaggrFatturato.CurrentID
    DescrRaggrFattContratto = .cboRaggrFatturato.Text
    
    IDClienteFatturazione = .CDAnaFatt.KeyFieldID
    ClienteFatturazione = .CDAnaFatt.Code & " " & .CDAnaFatt.Description
    
    IDContrattoBancario = .cboContrattoBancario.CurrentID
    DescrContrattoBancario = .cboContrattoBancario.Text
    
    IDAccordoCommerciale = .cboAccordoCommerciale.CurrentID
    DescrAccordoCommerciale = .cboAccordoCommerciale.Text
    
    IDArticoloContratto = .CDArticoloContratto.KeyFieldID
    ArticoloContratto = .CDArticoloContratto.Code & " - " & .CDArticoloContratto.Description
    
    IDPDCContratto = Link_ContoPDC
    CodicePDCContratto = .txtCodiceConto.Text
    DescrPDCContratto = .txtDescrizioneConto.Text

    
    IDIstatContratto = .cboIstat.CurrentID
    DescrIstaContratto = .cboIstat.Text
    
    MaggIstatContratto = .txtMaggiorazioneIstat.Value
    NumeroLicenze = .txtNumeroLicenze.Value
    
    RapprLegaleAzienda = .cboRapprAzienda.Text
    RuoloRapprAzienda = .cboRuoloRapprAzienda.Text
    RapprLegaleCliente = .cboRapprCliente.Text
    RuoloRapprCliente = .cboRuoloRapprCliente.Text
    
    IDUtenteInserimento = .cboUtenteInserimento.CurrentID
    DataInsContratto = .txtDataInserimento.Text
    DescrUtenteInserimento = .cboUtenteInserimento.Text
    
    IDUtenteUltMod = .cboUtenteModifica.CurrentID
    DataUltMod = .txtDataModifica.Text
    DescrUtenteUltMod = .cboUtenteModifica.Text

End With
End Sub

Public Function GET_DESCRIZIONE_ALTRI_DATI() As String

GET_DESCRIZIONE_ALTRI_DATI = ""

GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Cliente di fatturazione: " & ClienteFatturazione & vbCrLf
GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Articolo contratto: " & ArticoloContratto & vbCrLf
GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Piano dei conti: " & CodicePDCContratto & " -  " & DescrPDCContratto & vbCrLf
GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Raggruppamento fatturato: " & DescrRaggrFattContratto & vbCrLf
GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Classificazione: " & DescrClassContratto & vbCrLf
GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Contratto bancario: " & DescrContrattoBancario & vbCrLf
GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Accordo commerciale: " & DescrAccordoCommerciale & vbCrLf
GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Numero licenze: " & NumeroLicenze & vbCrLf
GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Istat rinnovo: " & DescrIstaContratto & vbCrLf
GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Maggiorazione istat: " & FormatNumber(MaggIstatContratto, 2) & vbCrLf

GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Rappresentante azienda: " & RapprLegaleAzienda & vbCrLf
GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Rappresentante cliente: " & RapprLegaleCliente & vbCrLf

GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Utente inserimento: " & DescrUtenteInserimento & vbCrLf
GET_DESCRIZIONE_ALTRI_DATI = GET_DESCRIZIONE_ALTRI_DATI & "Utente ult. mod.: " & DescrUtenteUltMod

End Function
Public Function GET_DESCRIZIONE_UTENTE(IDUtente As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Utente FROM Utente "
sSQL = sSQL & "WHERE IDUtente=" & IDUtente

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_UTENTE = ""
Else
    GET_DESCRIZIONE_UTENTE = fnNotNull(rs!Utente)
End If

rs.CloseResultset
Set rs = Nothing
End Function



Public Sub InitVariabiliContrSalvato(IDContratto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


IDClassContratto = 0
DescrClassContratto = ""

IDRaggrFattContratto = 0
DescrRaggrFattContratto = ""

IDClienteFatturazione = 0
ClienteFatturazione = ""

IDContrattoBancario = 0
DescrContrattoBancario = ""

IDAccordoCommerciale = 0
DescrAccordoCommerciale = ""

IDArticoloContratto = 0
ArticoloContratto = ""

IDPDCContratto = 0
CodicePDCContratto = ""
DescrPDCContratto = ""

IDIstatContratto = 0
DescrIstaContratto = ""

MaggIstatContratto = 0
NumeroLicenze = 0

RapprLegaleAzienda = ""
RuoloRapprAzienda = ""
RapprLegaleCliente = ""
RuoloRapprCliente = ""

IDUtenteInserimento = TheApp.IDUser
DataInsContratto = Date

IDUtenteUltMod = TheApp.IDUser
DataUltMod = Date

IDContrattoPadre = 0

sSQL = "SELECT * FROM RV_POViewContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then

    IDClassContratto = fnNotNullN(rs("IDRV_POTipoClassificazioneContratto").Value)
    DescrClassContratto = fnNotNull(rs("TipoClassificazioneContratto").Value)
    
    IDRaggrFattContratto = fnNotNullN(rs("IDRaggruppamentoFatturato").Value)
    DescrRaggrFattContratto = fnNotNull(rs("RaggruppamentoFatturato").Value)
    
    IDClienteFatturazione = fnNotNullN(rs("IDAnagraficaFatturazione").Value)
    ClienteFatturazione = fnNotNull(rs("AnagraficaFatturazione").Value) & " " & fnNotNull(rs("NomeAnagraficaFatturazione").Value)
    
    IDContrattoBancario = fnNotNullN(rs("IDContrattoBancario").Value)
    DescrContrattoBancario = fnNotNull(rs("BancaPerAnagrafica").Value)
    
    IDAccordoCommerciale = fnNotNullN(rs("IDAccordoCommerciale").Value)
    DescrAccordoCommerciale = fnNotNull(rs("DescrizioneAccordoCommerciale").Value)
    
    IDArticoloContratto = fnNotNullN(rs("IDArticoloContratto").Value)
    ArticoloContratto = fnNotNull(rs("CodiceArticoloContratto").Value) & " - " & fnNotNull(rs("ArticoloContratto").Value)
    
    IDPDCContratto = fnNotNullN(rs("IDCodiceConto").Value)
    CodicePDCContratto = fnNotNull(rs("CodiceConto").Value)
    DescrPDCContratto = fnNotNull(rs("DescrizioneConto").Value)
    
    IDIstatContratto = fnNotNullN(rs("IDIstat").Value)
    DescrIstaContratto = fnNotNull(rs("Istat").Value)
    
    MaggIstatContratto = fnNotNullN(rs("Maggiorazione").Value)
    NumeroLicenze = fnNotNullN(rs("NumeroLicenze").Value)
    
    RapprLegaleAzienda = fnNotNull(rs("RiferimentoAzienda").Value)
    RuoloRapprAzienda = fnNotNull(rs("RuoloRifAzienda").Value)
    RapprLegaleCliente = fnNotNull(rs("RiferimentoCliente").Value)
    RuoloRapprCliente = fnNotNull(rs("RuoloRifCliente").Value)
    
    IDUtenteInserimento = fnNotNullN(rs("IDUtentePerInserimento").Value)
    DataInsContratto = fnNotNull(rs("DataInserimento").Value)
    DescrUtenteInserimento = fnNotNull(rs("UtentePerInserimento").Value)
    
    IDUtenteUltMod = fnNotNullN(rs("IDUtentePerModifica").Value)
    DataUltMod = fnNotNull(rs("DataModifica").Value)
    DescrUtenteUltMod = fnNotNull(rs("UtentePerModifica").Value)
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Public Sub GET_CARATTERISTICHE_RISORSA(grid As MSFlexGrid)
On Error Resume Next
Dim NumeroColonne As Long
Dim NumeroRighe As Long
Dim N_Col As Long
Dim N_Row As Long




N_Col = 0
N_Row = 0


NumeroColonne = 2
NumeroRighe = 13

grid.Clear
grid.Cols = 1
grid.Rows = 1
DoEvents

With grid
    .Cols = NumeroColonne
    .Rows = NumeroRighe '+ 1
End With

''''INTESTAZIONI'''''''''''''''''''''''''''''''''''''''''''''''
'grid.TextMatrix(N_Row, 0) = "Descrizione"
'grid.TextMatrix(N_Row, 1) = "Valore"
'grid.Col = 0
'grid.Row = N_Row
'grid.CellFontBold = True
'grid.Col = 1
'Mgrid.Row = N_Row
'grid.CellFontBold = True

grid.ColWidth(0) = 2500
grid.ColWidth(1) = 2600

N_Row = 1

    grid.TextMatrix(0, 0) = "Cliente di fatturazione"
    grid.TextMatrix(0, 1) = ClienteFatturazione
    
    grid.TextMatrix(1, 0) = "Articolo contratto"
    grid.TextMatrix(1, 1) = ArticoloContratto
    
    grid.TextMatrix(2, 0) = "Piano dei conti"
    grid.TextMatrix(2, 1) = CodicePDCContratto & " -  " & DescrPDCContratto
    
    grid.TextMatrix(3, 0) = "Raggruppamento fatturato"
    grid.TextMatrix(3, 1) = DescrRaggrFattContratto
    
    grid.TextMatrix(4, 0) = "Classificazione"
    grid.TextMatrix(4, 1) = DescrClassContratto
    
    grid.TextMatrix(5, 0) = "Accordo commerciale"
    grid.TextMatrix(5, 1) = DescrAccordoCommerciale
    
    grid.TextMatrix(6, 0) = "Numero licenze"
    grid.TextMatrix(6, 1) = NumeroLicenze
    
    grid.TextMatrix(7, 0) = "Istat rinnovo"
    grid.TextMatrix(7, 1) = DescrIstaContratto
    
    grid.TextMatrix(8, 0) = "Maggiorazione istat"
    grid.TextMatrix(8, 1) = FormatNumber(MaggIstatContratto, 2)
    
    grid.TextMatrix(9, 0) = "Rappresentante azienda"
    grid.TextMatrix(9, 1) = RapprLegaleAzienda
    
    grid.TextMatrix(10, 0) = "Rappresentante cliente"
    grid.TextMatrix(10, 1) = RapprLegaleCliente

    grid.TextMatrix(11, 0) = "Utente inserimento"
    grid.TextMatrix(11, 1) = DescrUtenteInserimento

    grid.TextMatrix(12, 0) = "Utente ult. mod."
    grid.TextMatrix(12, 1) = DescrUtenteUltMod
    
    
    
'    Me.MSFlexGrid1.Col = 1
'    Me.MSFlexGrid1.Row = N_Row
'    Me.MSFlexGrid1.CellFontBold = True
        


End Sub


