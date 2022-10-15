VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmFine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Passaggio attività (passo 2 di 2)"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdIndietro 
      Caption         =   "Indietro"
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdFine 
      Caption         =   "Fine"
      Height          =   375
      Left            =   9480
      TabIndex        =   0
      Top             =   6360
      Width           =   1575
   End
   Begin VB.ListBox LstRisultato 
      Height          =   4740
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   11055
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5880
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tutti i dati sono stati reperiti, quindi cliccare su fine per eseguire l'operazione"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   11055
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5520
      Width           =   11055
   End
End
Attribute VB_Name = "FrmFine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LINK_ESERCIZIO As Long
Private LINK_SEZIONALE As Long
Private LINK_PERIODO_IVA As Long
Private Mov As DmtMovim.cMovimentazione

Private Sub cmdAnnulla_Click()

End Sub

Private Sub cmdFine_Click()
On Error GoTo ERR_cmdFine_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsCount As DmtOleDbLib.adoResultset
Dim UnitaProgresso As Double
Dim NumeroRecord As Long
Dim LINK_INTERVENTO As Long
Dim PrezzoLordoIvaCommessa As Long
Me.cmdFine.Enabled = False
'''''CONTEGGIO RECORD
sSQL = "SELECT COUNT(Registra) AS NumeroRecord "
sSQL = sSQL & "FROM RV_PO08_TMPPassaggioAttivita "
sSQL = sSQL & "WHERE Registra=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser


Set rsCount = CnDMT.OpenResultset(sSQL)

If rsCount.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rsCount!NumeroRecord)
End If

rsCount.CloseResultset
Set rsCount = Nothing


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If NumeroRecord = 0 Then
    Me.cmdFine.Enabled = False
    Exit Sub
End If

sSQL = "SELECT * FROM RV_PO08_TMPPassaggioAttivita "
sSQL = sSQL & "WHERE Registra=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " ORDER BY TipoAssociazione, IDRV_PO08_ParcoNatanti, IDAnagrafica"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    Screen.MousePointer = 11
    DoEvents
    Me.LstRisultato.AddItem "***********************************************************************************************************************************"
    Me.LstRisultato.AddItem "Passaggio dell'attività " & fnNotNull(rs!CodiceArticolo) & "-" & fnNotNull(rs!Articolo) & " per il natante " & fnNotNull(rs!NomeNatante)
    Me.LstRisultato.ListIndex = Me.LstRisultato.ListCount - 1
    
    Select Case fnNotNull(rs!TipoAssociazione)
        Case 0
            Me.LstRisultato.AddItem "Tipo associazione: Commessa specifica"
            Me.LstRisultato.AddItem "Associazione alla commessa N° " & fnNotNullN(rs!NumeroDocumentoAssociato) & " del " & fnNotNull(rs!DataDocumentoAssociato)
            Me.LstRisultato.ListIndex = Me.LstRisultato.ListCount - 1
            LINK_INTERVENTO = fnNotNullN(rs!IDRV_PO08_InterventoAssociato)
        Case 1
            Me.LstRisultato.AddItem "Tipo associazione: Crea una nuova commessa se non esiste, altrimenti associa alla prima disponibile"
            LINK_INTERVENTO = GET_LINK_ASSOCIAZIONE_INTERVENTO(rs, fnNotNullN(rs!TipoAssociazione))
            Me.LstRisultato.AddItem "Associazione commessa N° " & GET_DESCRIZIONE_INTERVENTO(LINK_INTERVENTO)
            Me.LstRisultato.ListIndex = Me.LstRisultato.ListCount - 1
            
        Case 2
            Me.LstRisultato.AddItem "Tipo associazione: Crea una nuova commessa"
            LINK_INTERVENTO = GET_LINK_ASSOCIAZIONE_INTERVENTO(rs, fnNotNullN(rs!TipoAssociazione))
            Me.LstRisultato.AddItem "Associazione alla commessa N° " & GET_DESCRIZIONE_INTERVENTO(LINK_INTERVENTO)
            Me.LstRisultato.ListIndex = Me.LstRisultato.ListCount - 1
    End Select
    
    DoEvents
    Me.LstRisultato.AddItem "Inserimento e ricalcolo della commessa in corso....... "
    Me.LstRisultato.ListIndex = Me.LstRisultato.ListCount - 1
    DoEvents
    
    INSERISCI_MACRO_ATTIVITA LINK_INTERVENTO, fnNotNullN(rs!IDRV_PO08_InterventoLavorazione), 0

    PrezzoLordoIvaCommessa = GET_PREZZO_LORDO_IVA_COMMESSA(LINK_INTERVENTO)
    
    GET_CAMBIO_PREZZO_LORDO_IVA LINK_INTERVENTO, PrezzoLordoIvaCommessa
    
    GET_TOTALI_DOCUMENTO LINK_INTERVENTO

    MovimentazioneDocumento fnGetTipoOggetto("RV_PO08_Intervento"), LINK_INTERVENTO
    
    CHIUDI_LAVORAZIONE fnNotNullN(rs!IDRV_PO08_InterventoLavorazione)
    
    
    
    Me.LstRisultato.AddItem "Passaggio attività eseguito con successo"
    Me.LstRisultato.AddItem "***********************************************************************************************************************************"
    Me.LstRisultato.ListIndex = Me.LstRisultato.ListCount - 1
    DoEvents
    
    If (UnitaProgresso + Me.ProgressBar1.Value) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
    End If
    DoEvents
    Screen.MousePointer = 0
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Me.lblInfo.Caption = "OPERAZIONE COMPLETATA"
Unload Me

Exit Sub
ERR_cmdFine_Click:
    MsgBox Err.Description, vbCritical, "cmdFine_Click"
    Screen.MousePointer = 0
End Sub

Private Sub CmdIndietro_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
    Me.cmdFine.SetFocus
    
End Sub

Private Sub Form_Load()
Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)

LINK_PERIODO_IVA = GET_LINK_PERIODO_IVA(DatePart("yyyy", Date))
LINK_ESERCIZIO = GET_LINK_ESERCIZIO(Date)
LINK_SEZIONALE = GET_LINK_SEZIONALE_DEFAULT
    
End Sub

Private Function GET_DESCRIZIONE_INTERVENTO(IDIntervento As Long) As String
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String


sSQL = "SELECT DataDocumento, NumeroDocumento "
sSQL = sSQL & "FROM RV_PO08_Intervento "
sSQL = sSQL & "WHERE IDRV_PO08_Intervento=" & IDIntervento


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_INTERVENTO = ""
Else
    GET_DESCRIZIONE_INTERVENTO = fnNotNullN(rs!NumeroDocumento) & " del " & fnNotNull(rs!DataDocumento)
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_LINK_ASSOCIAZIONE_INTERVENTO(rsTmp As DmtOleDbLib.adoResultset, TipoAssociazione As Long) As Long
Dim sSQL As String
Dim rsControllo As DmtOleDbLib.adoResultset
Dim LINK_STATO_INTERVENTO_DEFAULT As Long


If TipoAssociazione = 2 Then
    GET_LINK_ASSOCIAZIONE_INTERVENTO = CREA_NUOVO_INTERVENTO(rsTmp)
    Exit Function
End If

LINK_STATO_INTERVENTO_DEFAULT = GET_TIPO_ARTICOLO("IDStatoInterventoPerRicerca")


sSQL = "SELECT IDRV_PO08_Intervento, NumeroDocumento, DataDocumento "
sSQL = sSQL & "FROM RV_PO08_Intervento "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDRV_PO08_TipoIntervento=1 "
sSQL = sSQL & " AND IDRV_PO08_StatoIntervento=" & LINK_STATO_INTERVENTO_DEFAULT
sSQL = sSQL & " AND IDRV_PO08_ParcoNatanti=" & fnNotNullN(rsTmp!IDRV_PO08_ParcoNatanti)
sSQL = sSQL & " ORDER BY DataDocumento DESC, NumeroDocumento DESC"

Set rsControllo = CnDMT.OpenResultset(sSQL)

If rsControllo.EOF Then
    rsControllo.CloseResultset
    Set rsControllo = Nothing
    GET_LINK_ASSOCIAZIONE_INTERVENTO = CREA_NUOVO_INTERVENTO(rsTmp)
    Exit Function
Else
    GET_LINK_ASSOCIAZIONE_INTERVENTO = fnNotNullN(rsControllo!IDRV_PO08_Intervento)
End If

If Not (rsControllo Is Nothing) Then
    rsControllo.CloseResultset
    Set rsControllo = Nothing
End If
End Function
Private Function GET_LINK_PERIODO_IVA(Anno As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDPeriodoIVA FROM PeriodoIVA "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND Anno=" & Anno


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_PERIODO_IVA = 0
Else
    GET_LINK_PERIODO_IVA = fnNotNullN(rs!IDPeriodoIVA)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_IVA_ARTICOLO(IDArticolo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIvaVendita FROM Articolo "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDArticolo=" & IDArticolo



Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_ARTICOLO = 0
Else
    GET_LINK_IVA_ARTICOLO = fnNotNullN(rs!IDIvaVendita)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_ESERCIZIO(StringaData As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDEsercizio FROM Esercizio "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND DataInizio<=" & fnNormDate(StringaData)
sSQL = sSQL & " AND DataFine>=" & fnNormDate(StringaData)
sSQL = sSQL & " AND IDTipoEsercizio<>3"


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ESERCIZIO = 0
Else
    GET_LINK_ESERCIZIO = fnNotNullN(rs!IDEsercizio)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_SEZIONALE_DEFAULT() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSezionale FROM RV_PO08_ParametriSezionale "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.Branch
sSQL = sSQL & " AND IDRV_PO08_TipoIntervento=" & 1


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_SEZIONALE_DEFAULT = 0
Else
    GET_LINK_SEZIONALE_DEFAULT = fnNotNullN(rs!IDSezionale)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ALIQUOTA_IVA(IDIva As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "Select AliquotaIVA FROM IVA WHERE IDIVA=" & IDIva

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_ALIQUOTA_IVA = fnNotNullN(rs!AliquotaIva)
Else
    GET_ALIQUOTA_IVA = 0
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function CREA_NUOVO_INTERVENTO(rsTmp As DmtOleDbLib.adoResultset) As Long
Dim rs As ADODB.Recordset
Dim flx As DMTFlusso.cFlusso

Set rs = New ADODB.Recordset

rs.Open "RV_PO08_Intervento", CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic


rs.AddNew
    
    rs!IDRV_PO08_Intervento = fnGetNewKey("RV_PO08_Intervento", "IDRV_PO08_Intervento")
    rs!IDRV_PO08_TipoIntervento = 1
    rs!IDAzienda = TheApp.IDFirm
    rs!IDFiliale = TheApp.Branch
    rs!IDTipoOggetto = fnGetTipoOggetto("RV_PO08_Intervento")
    rs!IDRV_PO08_ParcoNatanti = fnNotNullN(rsTmp!IDRV_PO08_ParcoNatanti)
    rs!IDSezionale = LINK_SEZIONALE
    rs!IDEsercizio = LINK_ESERCIZIO
    rs!DataDocumento = Date
    rs!NumeroDocumento = GET_NUMERO_DOCUMENTO(LINK_SEZIONALE)
    rs!IDAnagrafica = fnNotNullN(rsTmp!IDAnagrafica)
    rs!IDAnagraficaFattura = GET_ANAGRAFICA_FATTURAZIONE(fnNotNullN(rsTmp!IDAnagrafica), fnNotNullN(rsTmp!IDRV_PO08_ParcoNatanti), fnNotNullN(rsTmp!IDRV_PO08_InterventoLavorazione))
    rs!IDRV_PO08_StatoIntervento = GET_TIPO_ARTICOLO("IDStatoInterventoPerRicerca")
    rs!IDRV_PO08_TipoAddebito = 0
    rs!IDRV_PO08_CategoriaChiamata = 0
    rs!IDUtenteInserimento = TheApp.IDUser
    rs!IDUtenteModifica = TheApp.IDUser
    rs!DataInserimento = Date
    rs!DataModifica = Date
    rs!IDListinoVendita = GET_VALORE_CAMPO_LAVORAZIONE(rsTmp!IDRV_PO08_InterventoLavorazione, "IDListinoVendita")
    rs!IDListinoAcquisto = GET_VALORE_CAMPO_LAVORAZIONE(rsTmp!IDRV_PO08_InterventoLavorazione, "IDListinoAcquisto")
    rs!IDMagazzino = GET_VALORE_CAMPO_LAVORAZIONE(rsTmp!IDRV_PO08_InterventoLavorazione, "IDMagazzino")
    rs!IDOggetto = GET_LINK_OGGETTO(LINK_SEZIONALE, rs!DataDocumento, rs!NumeroDocumento)
    rs!PrezzoLordoIva = GET_PREZZO_LORDO_IVA_CLIENTE(fnNotNullN(rs!IDAnagraficaFattura))
rs.Update


CREA_NUOVO_INTERVENTO = fnNotNullN(rs!IDRV_PO08_Intervento)

'INCREMENTO DEL NUMERO SEZIONALE
INCREMENTO_NUMERO_SEZIONALE LINK_SEZIONALE, LINK_PERIODO_IVA


'CREAZIONE FLUSSO DOCUMENTALE
Set flx = New DMTFlusso.cFlusso

flx.Connection = TheApp.Database.Connection

flx.IDFunzione = fncTrovaIDFunzione("RV_PO08_Intervento")

flx.IDOggetto = rs!IDOggetto

flx.IDTipoOggetto = fnGetTipoOggetto("RV_PO08_Intervento")

flx.OggettiCollegati.Add 10100, 0 'IDFlussoFunzione standard DMT di collegamento Scadenze à Primanota

flx.Insert


rs.Close
Set rs = Nothing





End Function
Private Function GET_NUMERO_DOCUMENTO(IDSezionale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT ProgressivoDisponibile FROM ProgressivoSezionale "
sSQL = sSQL & "WHERE IDSezionale=" & IDSezionale
sSQL = sSQL & " AND IDTipoModulo=1"
sSQL = sSQL & " AND IDPeriodoIVA=" & LINK_PERIODO_IVA

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    sSQL = "INSERT INTO ProgressivoSezionale ("
    sSQL = sSQL & "IDProgressivoSezionale, IDTipoModulo, IDPeriodoIVA, IDSezionale, "
    sSQL = sSQL & "ProgressivoDisponibile, DataUltimaVariazione, IDUtenteUltimaVariazione, "
    sSQL = sSQL & "VirtualDelete) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & fnGetNewKey("ProgressivoSezionale", "IDProgressivoSezionale") & ", "
    sSQL = sSQL & 1 & ", "
    sSQL = sSQL & LINK_PERIODO_IVA & ", "
    sSQL = sSQL & IDSezionale & ", "
    sSQL = sSQL & 1 & ", "
    sSQL = sSQL & fnNormDate(Date) & ", "
    sSQL = sSQL & TheApp.IDUser & ", "
    sSQL = sSQL & 0 & ")"
    
    CnDMT.Execute sSQL
    
    GET_NUMERO_DOCUMENTO = 1
Else
    GET_NUMERO_DOCUMENTO = fnNotNullN(rs!ProgressivoDisponibile)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ANAGRAFICA_FATTURAZIONE(IDAnagraficaIntestatario As Long, IDParcoNatanti As Long, IDLavorazione As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDAnagraficaFattura As Long


'''''''''''''PRENDE L'ANAGRAFICA DALLA LAVORAZIONE
sSQL = "SELECT IDAnagraficaFatturazione "
sSQL = sSQL & "FROM RV_PO08_InterventoLavorazione "
sSQL = sSQL & "WHERE IDRV_PO08_InterventoLavorazione=" & IDLavorazione


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    IDAnagraficaFattura = 0
Else
    IDAnagraficaFattura = fnNotNullN(rs!IDAnagraficaFatturazione)
End If

rs.CloseResultset
Set rs = Nothing

If IDAnagraficaFattura > 0 Then
    GET_ANAGRAFICA_FATTURAZIONE = IDAnagraficaFattura
    Exit Function
End If


''''''''''''''''''PRENDE L'ANAGRAFICA DAL PARCO NATANTI
sSQL = "SELECT IDAnagraficaPerFattura "
sSQL = sSQL & "FROM RV_PO08_ParcoNatanti "
sSQL = sSQL & "WHERE IDRV_PO08_ParcoNatanti=" & IDParcoNatanti

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    IDAnagraficaFattura = 0
Else
    IDAnagraficaFattura = fnNotNullN(rs!IDAnagraficaPerFattura)
End If

rs.CloseResultset
Set rs = Nothing

If IDAnagraficaFattura > 0 Then
    GET_ANAGRAFICA_FATTURAZIONE = IDAnagraficaFattura
    Exit Function
End If

''''''''''''''''''''''L'ANAGRAFICA DI FATTURAZIONE E' UGUALE ALL'INTESTATARIO
GET_ANAGRAFICA_FATTURAZIONE = IDAnagraficaIntestatario


End Function
Public Function GET_TIPO_ARTICOLO(NomeCampo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_PO08_ParametriAzienda "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_ARTICOLO = 0
Else
    GET_TIPO_ARTICOLO = fnNotNullN(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_OGGETTO(IDSezionale As Long, DataDocumento As String, NumeroDocumento As Long)
Dim IDOggetto As Long
Dim sSQL As String

IDOggetto = fnGetNewKey("Oggetto", "IDOggetto")

sSQL = "INSERT INTO Oggetto ("
sSQL = sSQL & "IDOggetto, IDTipoOggetto, IDAzienda, IDAttivitaAzienda, IDSezionale, "
sSQL = sSQL & "Oggetto, DataEmissione, Numero, DataUltimaVariazione, "
sSQL = sSQL & "IDUtenteUltimaVariazione, VirtualDelete, IDFunzione)"
sSQL = sSQL & " VALUES ("
sSQL = sSQL & IDOggetto & ", "
sSQL = sSQL & fnGetTipoOggetto("RV_PO08_Intervento") & ", "
sSQL = sSQL & TheApp.IDFirm & ", "
sSQL = sSQL & GetAttivitaAzienda(TheApp.IDFirm, TheApp.Branch) & ", "
sSQL = sSQL & IDSezionale & ", "
sSQL = sSQL & fnNormString(fncTrovaDescrizioneFunzione("RV_PO08_Intervento")) & ", "
sSQL = sSQL & fnNormDate(DataDocumento) & ", "
sSQL = sSQL & fnNormNumber(NumeroDocumento) & ", "
sSQL = sSQL & fnNormDate(Date) & ", "
sSQL = sSQL & TheApp.IDUser & ", "
sSQL = sSQL & 0 & ", "
sSQL = sSQL & fncTrovaIDFunzione("RV_PO08_Intervento") & ")"

CnDMT.Execute sSQL

GET_LINK_OGGETTO = IDOggetto

End Function

Private Function GetAttivitaAzienda(IDAzienda As Long, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AttivitaAzienda.IDAttivitaAzienda, Azienda.IDAzienda, Filiale.IDFiliale "
sSQL = sSQL & "FROM AttivitaAzienda INNER JOIN "
sSQL = sSQL & "Azienda ON AttivitaAzienda.IDAzienda = Azienda.IDAzienda INNER JOIN "
sSQL = sSQL & "Filiale ON AttivitaAzienda.IDAttivitaAzienda = Filiale.IDAttivitaAzienda "
sSQL = sSQL & "Where (Azienda.IDAzienda =" & IDAzienda & ") And (Filiale.IDFiliale = " & IDFiliale & ")"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GetAttivitaAzienda = 0
Else
    GetAttivitaAzienda = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function fnGetTipoOggetto(Optional Gestore As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    If Gestore = "" Then
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(App.EXEName)
    Else
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(Gestore)
    End If
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Function fncTrovaIDFunzione(Gestore As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Funzione.IDFunzione, Gestore.Gestore "
sSQL = sSQL & "FROM Gestore INNER JOIN "
sSQL = sSQL & "TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore INNER JOIN "
sSQL = sSQL & "Funzione ON TipoOggetto.IDTipoOggetto = Funzione.IDTipoOggetto "
sSQL = sSQL & "WHERE (Gestore.Gestore = " & fnNormString(Gestore) & ") "
sSQL = sSQL & "AND (Funzione.IDFunzione >= 10000)"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaIDFunzione = fnNotNullN(rs!IDFunzione)
Else
    fncTrovaIDFunzione = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function fncTrovaDescrizioneFunzione(Gestore As String) As String

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Funzione.Funzione, Gestore.Gestore "
sSQL = sSQL & "FROM Gestore INNER JOIN "
sSQL = sSQL & "TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore INNER JOIN "
sSQL = sSQL & "Funzione ON TipoOggetto.IDTipoOggetto = Funzione.IDTipoOggetto "
sSQL = sSQL & "WHERE (Gestore.Gestore = " & fnNormString(Gestore) & ") "
sSQL = sSQL & "AND (Funzione.IDFunzione >= 10000)"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaDescrizioneFunzione = fnNotNull(rs!Funzione)
Else
    fncTrovaDescrizioneFunzione = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub INSERISCI_MACRO_ATTIVITA(IDIntervento As Long, IDLavorazione As Long, IDIvaCLiente As Long)
Dim sSQL As String
Dim rsLav As DmtOleDbLib.adoResultset
Dim rsInt As ADODB.Recordset
Dim PrezzoLordoIvaCommessa As Long

sSQL = "SELECT * FROM RV_PO08_InterventoLavorazione "
sSQL = sSQL & "WHERE IDRV_PO08_InterventoLavorazione=" & IDLavorazione

Set rsLav = CnDMT.OpenResultset(sSQL)


If Not rsLav.EOF Then
    Set rsInt = New ADODB.Recordset
    
    sSQL = "SELECT * FROM RV_PO08_InterventoMacroAttivita "
    sSQL = sSQL & "WHERE IDRV_PO08_Intervento = " & IDIntervento
    
    rsInt.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    rsInt.AddNew
        rsInt!IDRV_PO08_InterventoMacroAttivita = fnGetNewKey("RV_PO08_InterventoMacroAttivita", "IDRV_PO08_InterventoMacroAttivita")
        rsInt!IDRV_PO08_Intervento = IDIntervento
        rsInt!IDRV_PO08_InterventoPreventivo = fnNotNullN(rsLav!IDRV_PO08_InterventoP)
        rsInt!IDRV_PO08_InterventoMacroAttivitaPreventivo = fnNotNullN(rsLav!IDRV_PO08_InterventoMacroAttivitaP)
        rsInt!IDArticolo = fnNotNullN(rsLav!IDArticolo)
        rsInt!CodiceArticolo = fnNotNull(rsLav!CodiceArticolo)
        rsInt!Articolo = fnNotNull(rsLav!Articolo)
        rsInt!IDUnitaDiMisura = fnNotNullN(rsLav!IDUnitaDiMisura)
        rsInt!IDRV_PO08_ModelloPreventivo = fnNotNullN(rsLav!IDRV_PO08_ModelloPreventivo)
        rsInt!IDRV_PO08_ElementoNatante = fnNotNullN(rsLav!IDRV_PO08_ParcoNatantiRighe)
        rsInt!ElementoNatante = fnNotNull(rsLav!ElementoNatante)
        rsInt!ElementoNatanteIdentificativo = fnNotNullN(rsLav!ElementoNatanteIdentificativo)
        rsInt!IDRV_PO08_GruppoElemento = fnNotNullN(rsLav!IDRV_PO08_GruppoElemento)
        rsInt!GruppoElemento = fnNotNull(rsLav!GruppoElemento)
        rsInt!CostoCalcolatoLordoIVA = fnNotNullN(rsLav!CostoLordoIVA)
        rsInt!CostoCalcolatoImponibile = fnNotNullN(rsLav!CostoNettoIVA)
        rsInt!CostoCalcolatoNeutro = fnNotNullN(rsLav!CostoNeutro)
        rsInt!CostoCalcolatoIva = fnNotNullN(rsLav!CostoIva)
        rsInt!RicavoCalcolatoLordoIVA = fnNotNullN(rsLav!RicavoLordoIVA)
        rsInt!RicavoCalcolatoImponibile = fnNotNullN(rsLav!RicavoNettoIVA)
        rsInt!RicavoCalcolatoNeutro = fnNotNullN(rsLav!RicavoNeutro)
        rsInt!RicavoCalcolatoIva = fnNotNullN(rsLav!RicavoIVA)
        'If IDIvaCLiente > 0 Then
        '    rsInt!IDIvaPrezzoConcordato = IDIvaCLiente
        'Else
        '    rsInt!IDIvaPrezzoConcordato = GET_LINK_IVA_ARTICOLO(rsLav!IDArticolo)
        'End If
        rsInt!IDIvaPrezzoConcordato = fnNotNullN(rsLav!IDIvaPrezzoConcordato)
        rsInt!AliquotaIvaPrezzoConcordato = fnNotNullN(rsLav!AliquotaIvaPrezzoConcordato)
        rsInt!ImportoIVAPrezzoConcordato = fnNotNullN(rsLav!ImportoIVAPrezzoConcordato)
        rsInt!PrezzoConcordatoLordoIVA = fnNotNullN(rsLav!PrezzoConcordatoLordoIVA)
        rsInt!PrezzoConcordatoImponibile = fnNotNullN(rsLav!PrezzoConcordatoImponibile)
        rsInt!PrezzoConcordatoNeutro = fnNotNullN(rsLav!PrezzoConcordatoNeutro)
        rsInt!ImportoDifferenza = fnNotNullN(rsLav!ImportoDifferenza)
        rsInt!ScontoPercentualeDiff = fnNotNullN(rsLav!ScontoPercentualeDiff)
        rsInt!Approvato = True
        rsInt!StampaDettagli = True
        
    rsInt.Update
    
    INSERISCI_DETTAGLIO_MACRO_ATTIVITA IDIntervento, IDLavorazione, rsInt!IDRV_PO08_InterventoMacroAttivita
    
    '''''''''''''''''''''CREAZIONE COLLEGAMENTO ALLA COMMESSA''''''''''''''''''''''''''''''''''''''
        sSQL = "UPDATE RV_PO08_InterventoLavorazione SET "
        If fnNotNullN(rsLav!IDRV_PO08_TipoIntervento) = 0 Then
            sSQL = sSQL & "IDRV_PO08_TipoIntervento=1" & ", "
        End If
        sSQL = sSQL & "IDRV_PO08_InterventoS=" & IDIntervento & ", "
        sSQL = sSQL & "IDRV_PO08_InterventoMacroAttivitaS=" & rsInt!IDRV_PO08_InterventoMacroAttivita
        sSQL = sSQL & "WHERE IDRV_PO08_InterventoLavorazione=" & IDLavorazione
    
        CnDMT.Execute sSQL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rsInt.Close
    Set rsInt = Nothing

End If

rsLav.CloseResultset
Set rsLav = Nothing


End Sub

Private Sub INSERISCI_DETTAGLIO_MACRO_ATTIVITA(IDIntervento As Long, IDLavorazione As Long, IDMacroAttivita As Long)
Dim sSQL As String
Dim rsLav As DmtOleDbLib.adoResultset
Dim rsInt As ADODB.Recordset

sSQL = "SELECT * FROM RV_PO08_InterventoLavorazioneRighe "
sSQL = sSQL & "WHERE IDRV_PO08_InterventoLavorazione=" & IDLavorazione

Set rsLav = CnDMT.OpenResultset(sSQL)
Set rsInt = New ADODB.Recordset
rsInt.Open "RV_PO08_InterventoMacroAttivitaRighe", CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsLav.EOF

    rsInt.AddNew
        rsInt!IDRV_PO08_InterventoMacroAttivitaRighe = fnGetNewKey("RV_PO08_InterventoMacroAttivitaRighe", "IDRV_PO08_InterventoMacroAttivitaRighe")
        rsInt!IDRV_PO08_InterventoMacroAttivita = IDMacroAttivita
        rsInt!IDRV_PO08_Intervento = IDIntervento
        rsInt!IDArticolo = rsLav!IDArticolo
        rsInt!IDGruppoEquivalenzaArticolo = rsLav!IDGruppoEquivalenzaArticolo
        rsInt!CodiceArticolo = rsLav!CodiceArticolo
        rsInt!Articolo = rsLav!Articolo
        rsInt!IDUnitaDiMisuraVendita = rsLav!IDUnitaDiMisuraVendita

        rsInt!IDRV_PO08_ElementoNatante = rsLav!IDRV_PO08_ParcoNatantiRighe
        rsInt!ElementoNatante = rsLav!ElementoNatante
        rsInt!ElementoNatanteIdentificativo = rsLav!ValoreIndentificativoElemento
        rsInt!IDRV_PO08_GruppoElemento = rsLav!IDRV_PO08_GruppoElemento
        rsInt!GruppoElemento = rsLav!GruppoElemento
        rsInt!DataInserimento = Date
        rsInt("Quantita").Value = rsLav("Quantita").Value
    
        ''''''''''''''''''''''''''''''''''''''''''''VENDITA'''''''''''''''''''''''''''''''''''''
        rsInt("IDListinoVendita").Value = rsLav("IDListinoVendita").Value
        rsInt("ImportoUnitarioVenditaNeutro").Value = rsLav("ImportoUnitarioVenditaNeutro").Value
        rsInt("ImportoUnitarioVenditaNettoIVA").Value = rsLav("ImportoUnitarioVenditaNettoIVA").Value
        rsInt("ImportoUnitarioVenditaLordoIVA").Value = rsLav("ImportoUnitarioVenditaLordoIVA").Value
        rsInt("Sconto1Vendita").Value = rsLav("Sconto1Vendita").Value
        rsInt("Sconto2Vendita").Value = rsLav("Sconto2Vendita").Value
        rsInt("Sconto3Vendita").Value = rsLav("Sconto3Vendita").Value
        rsInt("Sconto4Vendita").Value = rsLav("Sconto4Vendita").Value
        rsInt("Sconto5Vendita").Value = rsLav("Sconto5Vendita").Value
        
        rsInt("IDIvaVendita").Value = rsLav("IDIvaVendita").Value
        rsInt("AliquotaIvaVendita").Value = rsLav("AliquotaIvaVendita").Value
        rsInt("ImportoTotaleVenditaLordoIVA").Value = rsLav("ImportoTotaleVenditaLordoIVA").Value
        rsInt("ImportoTotaleVenditaNettoIVA").Value = rsLav("ImportoTotaleVenditaNettoIVA").Value
        rsInt("ImportoTotaleVenditaNeutro").Value = rsLav("ImportoTotaleVenditaNeutro").Value
        rsInt("ImportoIVA").Value = rsLav("ImportoIVA").Value
        rsInt("DaDataRimessaggio").Value = rsLav("DaDataRimessaggio").Value
        rsInt("ADataRimessaggio").Value = rsLav("ADataRimessaggio").Value
        ''''''''''''''''''''''''''''''''''''''''''''ACQUISTO''''''''''''''''''''''''''''''''''''
        rsInt("IDUnitaDiMisuraAcquisto").Value = rsLav("IDUnitaDiMisuraAcquisto").Value
        rsInt("IDListinoAcquisto").Value = rsLav("IDListinoAcquisto").Value
        rsInt("ImportoUnitarioAcquistoNettoIva").Value = rsLav("ImportoUnitarioAcquistoNettoIva").Value
        rsInt("ImportoUnitarioAcquistoLordoIva").Value = rsLav("ImportoUnitarioAcquistoLordoIva").Value
        rsInt("IDIvaAcquisto").Value = rsLav("IDIvaAcquisto").Value
        rsInt("AliquotaIvaAcquisto").Value = rsLav("AliquotaIvaAcquisto").Value
        rsInt("ImportoTotaleAcquistoNettoIVA").Value = rsLav("ImportoTotaleAcquistoNettoIVA").Value
        rsInt("ImportoTotaleAcquistoLordoIVA").Value = rsLav("ImportoTotaleAcquistoLordoIVA").Value
        rsInt("ImportoIvaAcquisto").Value = rsLav("ImportoIvaAcquisto").Value
        rsInt("Sconto1Acquisto").Value = rsLav("Sconto1Acquisto").Value
        rsInt("Sconto2Acquisto").Value = rsLav("Sconto2Acquisto").Value
        rsInt("Sconto3Acquisto").Value = rsLav("Sconto3Acquisto").Value
        rsInt("Sconto4Acquisto").Value = rsLav("Sconto4Acquisto").Value
        rsInt("Sconto5Acquisto").Value = rsLav("Sconto5Acquisto").Value
        rsInt("IDRV_PO08_TipoRigaVendita").Value = rsLav("IDRV_PO08_TipoRigaVendita").Value
    
    rsInt.Update
    

rsLav.MoveNext
Wend
rsInt.Close
Set rsInt = Nothing
rsLav.CloseResultset
Set rsLav = Nothing


End Sub

Private Sub GET_TOTALI_DOCUMENTO(IDIntervento As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsInt As ADODB.Recordset


sSQL = "SELECT Sum(PrezzoConcordatoImponibile) as TotaleImponibile, "
sSQL = sSQL & "Sum(ImportoIVAPrezzoConcordato) as TotaleImposta, "
sSQL = sSQL & "Sum(PrezzoConcordatoLordoIVA) as TotaleDocumento, "

sSQL = sSQL & "Sum(CostoCalcolatoImponibile) as TotaleCostoNettoIVA, "
sSQL = sSQL & "Sum(CostoCalcolatoLordoIVA) as TotaleCostoLordoIVA, "
sSQL = sSQL & "Sum(RicavoCalcolatoImponibile) as TotaleRicavoNettoIVA, "
sSQL = sSQL & "Sum(RicavoCalcolatoLordoIVA) as TotaleRicavoLordoIVA  "

sSQL = sSQL & "FROM RV_PO08_InterventoMacroAttivita "
sSQL = sSQL & "WHERE IDRV_PO08_Intervento=" & IDIntervento

Set rs = CnDMT.OpenResultset(sSQL)


sSQL = "SELECT * FROM RV_PO08_Intervento "
sSQL = sSQL & "WHERE IDRV_PO08_Intervento=" & IDIntervento
Set rsInt = New ADODB.Recordset




rsInt.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic



If Not rsInt.EOF Then

    rsInt!TotaleImponibileIntervento = fnNotNullN(rs!TotaleImponibile)
    rsInt!TotaleImposta = fnNotNullN(rs!TotaleImposta)
    rsInt!TotaleDocumento = fnNotNullN(rs!TotaleDocumento)
    rsInt!TotaleCostoNettoIVA = fnNotNullN(rs!TotaleCostoNettoIVA)
    rsInt!TotaleCostoLordoIVA = fnNotNullN(rs!TotaleCostoLordoIVA)
    rsInt!TotaleRicavoNettoIVA = fnNotNullN(rs!TotaleRicavoNettoIVA)
    rsInt!TotaleRicavoLordoIVA = fnNotNullN(rs!TotaleRicavoLordoIVA)
    rsInt!TotalePagato = GET_TOTALE_PAGATO(IDIntervento)
    rsInt!TotaleFatturato = GET_TOTALE_FATTURATO(IDIntervento)
    rsInt!Acconto = GET_TOTALE_ACCONTO(IDIntervento)
    rsInt!DaPagare = rsInt!TotaleDocumento - (rsInt!Acconto + rsInt!TotalePagato + rsInt!TotaleFatturato)
    rsInt.Update
    
    CREA_SCADENZA fnNotNullN(rsInt!IDOggetto), IDIntervento, rsInt!TotaleDocumento, rsInt!IDAnagraficaFattura, rsInt!NumeroDocumento, rsInt!DataDocumento, rsInt!IDSezionale
    
    
    rsInt.Close
    Set rsInt = Nothing
End If


rs.CloseResultset
Set rs = Nothing
End Sub

Private Function GET_TOTALE_PAGATO(IDIntervento As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT sum(ImportoPagato) AS TotaleImporto "
sSQL = sSQL & "FROM RV_PO08_IncassiRighe "
sSQL = sSQL & "WHERE IDRV_PO08_Intervento = " & IDIntervento
sSQL = sSQL & " AND IDRV_PO08_TipoIncasso <= 2"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_PAGATO = 0
Else
    GET_TOTALE_PAGATO = fnNotNullN(rs!TotaleImporto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_TOTALE_FATTURATO(IDIntervento As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT sum(ImportoPagato) AS TotaleImporto "
sSQL = sSQL & "FROM RV_PO08_IncassiRighe "
sSQL = sSQL & "WHERE IDRV_PO08_Intervento = " & IDIntervento
sSQL = sSQL & " AND IDRV_PO08_TipoIncasso = 4"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_FATTURATO = 0
Else
    GET_TOTALE_FATTURATO = fnNotNullN(rs!TotaleImporto)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_TOTALE_ACCONTO(IDIntervento As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT sum(ImportoPagato) AS TotaleImporto "
sSQL = sSQL & "FROM RV_PO08_IncassiRighe "
sSQL = sSQL & "WHERE IDRV_PO08_Intervento = " & IDIntervento
sSQL = sSQL & " AND IDRV_PO08_TipoIncasso = 3"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_ACCONTO = 0
Else
    GET_TOTALE_ACCONTO = fnNotNullN(rs!TotaleImporto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREA_SCADENZA(IDOggettoIntervento As Long, IDIntervento As Long, TotaleDocumento As Double, IDAnagrafica As Long, NumeroDocumento As Long, DataDocumento As String, IDSezionale As Long)
Dim flx As DMTFlusso.cFlusso
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim VAR_IMPORTO_PAGATO As Double
Dim VAR_IMPORTO_FATTURATO As Double
Dim VAR_TOTALE_DOCUMENTO_NETTO As Double
Dim VAR_TOTALE_ACCONTO As Double
Dim RIMANENZA As Double
Dim LINK_TESTATA_SCADENZA As Long

LINK_TESTATA_SCADENZA = GET_LINK_TESTATA_SCADENZA(IDOggettoIntervento)


If LINK_TESTATA_SCADENZA > 0 Then
    
    'Importo incassato
    VAR_IMPORTO_PAGATO = GET_IMPORTO_PAGATO(IDIntervento)
    
    
    'Eliminazione delle rate pagate
    sSQL = "DELETE FROM DettaglioScadenza "
    sSQL = sSQL & "WHERE IDTestataScadenza=" & LINK_TESTATA_SCADENZA
        
    CnDMT.Execute sSQL
    
    RIMANENZA = GeneraRateIntervento(IDIntervento, LINK_TESTATA_SCADENZA, VAR_IMPORTO_PAGATO, FormatNumber(TotaleDocumento, 2), 0, NumeroDocumento, DataDocumento)
    
    'Aggiornamento importo complessivo della testata scadenza
    sSQL = "UPDATE TestataScadenza SET "
    sSQL = sSQL & "IDAnagrafica_CF=" & IDAnagrafica & ", "
    sSQL = sSQL & "ImportoComplessivo=" & fnNormNumber(FormatNumber(TotaleDocumento, 2))
    sSQL = sSQL & " WHERE IDTestataScadenza=" & LINK_TESTATA_SCADENZA
        
    CnDMT.Execute sSQL
    
Else
    
    LINK_TESTATA_SCADENZA = GeneraNuovaTestataScadenza(LINK_TESTATA_SCADENZA, TotaleDocumento, IDOggettoIntervento, IDAnagrafica, DataDocumento, IDSezionale, NumeroDocumento)
    GeneraRateIntervento IDIntervento, LINK_TESTATA_SCADENZA, VAR_IMPORTO_PAGATO, FormatNumber(TotaleDocumento, 2), 0, NumeroDocumento, DataDocumento

End If

End Sub
Private Function GET_IMPORTO_PAGATO(IDIntervento) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDMovimentoContabile As Long
Dim rsMov As DmtOleDbLib.adoResultset
Dim IDOggettoScade As Long

'SOMMA INCASSI DI POSEIDON
sSQL = "SELECT SUM(ImportoPagato) AS Somma_Importo_Pagato "
sSQL = sSQL & "FROM RV_PO08_IncassiRighe "
sSQL = sSQL & "WHERE IDRV_PO08_Intervento = " & IDIntervento

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_IMPORTO_PAGATO = 0
Else
    GET_IMPORTO_PAGATO = fnNotNullN(rs!Somma_Importo_Pagato)
End If

rs.CloseResultset
Set rs = Nothing

GET_IMPORTO_PAGATO = FormatNumber(GET_IMPORTO_PAGATO, 2)

End Function
Private Function GET_LINK_TESTATA_SCADENZA(IDOggettoIntervento As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDOggettoScadenza As Long

sSQL = "SELECT IDOggettoCollegato "
sSQL = sSQL & "FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=10100"
sSQL = sSQL & " AND IDOggetto=" & IDOggettoIntervento
sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto("RV_PO08_Intervento")
sSQL = sSQL & " AND IDTipoOggettoCollegato=131"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_TESTATA_SCADENZA = 0
    
Else
    IDOggettoScadenza = fnNotNullN(rs!IDOggettoCollegato)
End If

rs.CloseResultset
Set rs = Nothing


If IDOggettoScadenza = 0 Then
    GET_LINK_TESTATA_SCADENZA = 0
Else
    sSQL = "SELECT IDTestataScadenza FROM TestataScadenza "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoScadenza
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_LINK_TESTATA_SCADENZA = 0
    Else
        GET_LINK_TESTATA_SCADENZA = fnNotNullN(rs!IDTestataScadenza)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If


End Function
Private Function GeneraRateIntervento(IDIntervento As Long, IDTestataScadenza As Long, ImportoIncassato As Double, ImportoIntervento As Double, Acconto As Double, NumeroDocumento As Long, DataDocumento As String) As Double
Dim rs As ADODB.Recordset
Dim rsNew As ADODB.Recordset
Dim sSQL As String
Dim IRata As Integer
Dim RIMANENZA As Double


sSQL = "SELECT * FROM DettaglioScadenza "

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

RIMANENZA = ImportoIntervento

If ImportoIncassato > 0 Then
    If (RIMANENZA - ImportoIncassato) <> 0 Then
        rsNew.AddNew
            rsNew!IDTestataScadenza = IDTestataScadenza
            rsNew!IDDettaglioScadenza = fnGetNewKey("DettaglioScadenza", "IDDettaglioScadenza")
            rsNew!DataScadenza = Date
            rsNew!ImportoScadenza = RIMANENZA
            rsNew!ImportoPagato = ImportoIncassato
            rsNew!DataPagamento = Date
            rsNew!IDTipoStatoScadenza = 1
            rsNew!IDTipoPosizioneScadenza = 2
            rsNew!IDTipoOggetto = 0
            rsNew!IDTipoPagamento = 1
            rsNew!RaggruppamentoScadenza = False
            rsNew!RIBA = False
            rsNew!TrasferitoOutlook = False
            rsNew!Contabilizzata = False
            rsNew!Note = Mid(GET_ANNOTAZIONI_INCASSI(IDIntervento), 1, 255)
            rsNew!NumeroRata = 1
            rsNew!DataUltimaVariazione = Date
            rsNew!IDUtenteUltimaVariazione = TheApp.IDUser
            rsNew!VirtualDelete = 0
        rsNew.Update
        RIMANENZA = RIMANENZA - ImportoIncassato
    Else
        rsNew.AddNew
            rsNew!IDTestataScadenza = IDTestataScadenza
            rsNew!IDDettaglioScadenza = fnGetNewKey("DettaglioScadenza", "IDDettaglioScadenza")
            rsNew!DataScadenza = Date
            rsNew!ImportoScadenza = RIMANENZA
            rsNew!ImportoPagato = ImportoIncassato
            rsNew!DataPagamento = Date
            rsNew!IDTipoStatoScadenza = 1
            rsNew!IDTipoPosizioneScadenza = 2
            rsNew!IDTipoOggetto = 0
            rsNew!IDTipoPagamento = 1
            rsNew!RaggruppamentoScadenza = False
            rsNew!RIBA = False
            rsNew!TrasferitoOutlook = False
            rsNew!Contabilizzata = False
            rsNew!Note = Mid(GET_ANNOTAZIONI_INCASSI(IDIntervento), 1, 255)
            rsNew!NumeroRata = 1
            rsNew!DataUltimaVariazione = Date
            rsNew!IDUtenteUltimaVariazione = TheApp.IDUser
            rsNew!VirtualDelete = 0
        rsNew.Update
        
        
        Exit Function
    End If
    
    
End If
    
    rsNew.AddNew
        rsNew!IDTestataScadenza = IDTestataScadenza
        rsNew!IDDettaglioScadenza = fnGetNewKey("DettaglioScadenza", "IDDettaglioScadenza")
        rsNew!DataScadenza = Date
        rsNew!ImportoScadenza = RIMANENZA
        rsNew!ImportoPagato = 0
        If RIMANENZA = 0 Then
            rsNew!IDTipoStatoScadenza = 1
            rsNew!IDTipoPosizioneScadenza = 2
        Else
            rsNew!IDTipoStatoScadenza = 1
            rsNew!IDTipoPosizioneScadenza = 3
        End If
        rsNew!IDTipoOggetto = 0
        rsNew!IDTipoPagamento = 1
        rsNew!RaggruppamentoScadenza = False
        rsNew!RIBA = False
        rsNew!TrasferitoOutlook = False
        rsNew!Contabilizzata = False
        rsNew!Note = "Intervento numero " & NumeroDocumento & " del " & DataDocumento
        rsNew!NumeroRata = 1
        rsNew!DataUltimaVariazione = Date
        rsNew!IDUtenteUltimaVariazione = TheApp.IDUser
        rsNew!VirtualDelete = 0
    rsNew.Update
    
    
    
    rsNew.Close
    Set rsNew = Nothing


GeneraRateIntervento = RIMANENZA
End Function
Private Function GET_ANNOTAZIONI_INCASSI(IDIntervento As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsMov As DmtOleDbLib.adoResultset
Dim IDOggettoScade As Long

GET_ANNOTAZIONI_INCASSI = ""


'ANNOTAZIONI INCASSI POSEIDON

sSQL = "SELECT RV_PO08_Incassi.Annotazioni, RV_PO08_IncassiRighe.ImportoPagato, RV_PO08_Incassi.DataIncasso, "
sSQL = sSQL & "RV_PO08_IncassiRighe.IDRV_PO08_TipoIncasso , RV_PO08_IncassiRighe.IDRV_PO08_Intervento, "
sSQL = sSQL & "RV_PO08_Incassi.IDRV_PO08_Incassi, RV_PO08_IncassiRighe.DescrizioneDocumento  "
sSQL = sSQL & "FROM RV_PO08_IncassiRighe INNER JOIN "
sSQL = sSQL & "RV_PO08_Incassi ON RV_PO08_IncassiRighe.IDRV_PO08_Incassi = RV_PO08_Incassi.IDRV_PO08_Incassi "
sSQL = sSQL & "WHERE (RV_PO08_IncassiRighe.IDRV_PO08_Intervento = " & IDIntervento & ") "

Set rs = CnDMT.OpenResultset(sSQL)
While Not rs.EOF
    Select Case fnNotNullN(rs!IDRV_PO08_TipoIncasso)
        Case 1
            GET_ANNOTAZIONI_INCASSI = GET_ANNOTAZIONI_INCASSI & "Inc. € " & FormatNumber(fnNotNullN(rs!ImportoPagato), 2) & " del " & fnNotNull(rs!DataIncasso)
        Case 2
            GET_ANNOTAZIONI_INCASSI = GET_ANNOTAZIONI_INCASSI & "Abb. € " & FormatNumber(fnNotNullN(rs!ImportoPagato), 2) & " del " & fnNotNull(rs!DataIncasso)
        Case 3
            GET_ANNOTAZIONI_INCASSI = GET_ANNOTAZIONI_INCASSI & "Acc. € " & FormatNumber(fnNotNullN(rs!ImportoPagato), 2) & " del " & fnNotNull(rs!DataIncasso)
        Case 4
            GET_ANNOTAZIONI_INCASSI = GET_ANNOTAZIONI_INCASSI & "Doc. € " & FormatNumber(fnNotNullN(rs!ImportoPagato), 2) & " " & fnNotNull(rs!DescrizioneDocumento)
            
    End Select
        If Len(Trim(rs!Annotazioni)) > 0 Then
            GET_ANNOTAZIONI_INCASSI = GET_ANNOTAZIONI_INCASSI & " (" & fnNotNull(rs!Annotazioni) & ")"
        End If
        
        GET_ANNOTAZIONI_INCASSI = GET_ANNOTAZIONI_INCASSI & vbCrLf
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing





End Function

Private Function GeneraNuovaTestataScadenza(IDTestataScadenza As Long, ImportoComplessivoScadenza As Double, IDOggettoIntervento As Long, IDAnagrafica As Long, DataDocumento As String, IDSezionale As Long, NumeroDocumento As Long) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim Link_Oggetto As Long
'CREAZIONE FLUSSO
Dim flx As DMTFlusso.cFlusso

Link_Oggetto = GET_LINK_OGGETTO_SCADENZA(DataDocumento, IDSezionale, NumeroDocumento)

If Link_Oggetto > 0 Then
        
    Set rs = New ADODB.Recordset
    
    rs.Open "TestataScadenza", CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    rs.AddNew
        GeneraNuovaTestataScadenza = fnGetNewKey("TestataScadenza", "IDTestataScadenza")
        rs!IDTestataScadenza = GeneraNuovaTestataScadenza
        rs!IDOggetto = Link_Oggetto
        rs!IDTipoOggetto = 131
        rs!IDFiliale = TheApp.Branch
        rs!IDNaturaScadenza = 6
        rs!IDAnagrafica_CF = IDAnagrafica
        rs!IDTipoAnagrafica_CF = 2
        rs!IDAzienda_CF = TheApp.IDFirm
        rs!IDAzienda = TheApp.IDFirm
        rs!IDPagamento = 52
        rs!IDRegistroIva = 1
        rs!IDTipoStatoScadenza = 1
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
        
    rs.Update
    
    
    
    rs.Close
    Set rs = Nothing
    
    
    


 

    Set flx = New DMTFlusso.cFlusso

    flx.Connection = CnDMT

    flx.IDFunzione = 88         'IDFunzione standard DMT

    flx.IDOggetto = Link_Oggetto

    flx.IDTipoOggetto = 131

    flx.OggettiCollegati.Add 13, 0 'IDFlussoFunzione standard DMT di collegamento Scadenze à Primanota

    flx.Insert

    
    'AGGIORNAMENTO FLUSSO OGGETTI COLLEGATI
    sSQL = "INSERT INTO FlussoOggettiCollegati ("
    sSQL = sSQL & "IDFlussoFunzione, IDTipoOggetto, IDOggetto, IDOggettoCollegato, IDTipoOggettoCollegato)"
    sSQL = sSQL & " VALUES ("
    sSQL = sSQL & 10100 & ", "
    sSQL = sSQL & fnGetTipoOggetto("RV_PO08_Intervento") & ", "
    sSQL = sSQL & IDOggettoIntervento & ", "
    
    sSQL = sSQL & Link_Oggetto & ", "
    sSQL = sSQL & 131 & ")"
    CnDMT.Execute sSQL
    
    'AGGIORNAMENTO FLUSSO FUNZIONE COLLEGATO
    sSQL = "UPDATE FlussoFunzioneCollegato SET "
    sSQL = sSQL & "FlussoFunzioneCollegato=2 "
    sSQL = sSQL & "WHERE IDFlussoFunzione=10100"
    sSQL = sSQL & " AND IDOggetto=" & IDOggettoIntervento
    sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto("RV_PO08_Intervento")
    CnDMT.Execute sSQL
    
    
    

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
    rs!IDAttivitaAzienda = GetAttivitaAzienda(TheApp.IDFirm, TheApp.Branch)
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
Private Function GET_VALORE_CAMPO_LAVORAZIONE(IDLavorazione As Long, Campo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & Campo
sSQL = sSQL & " FROM RV_PO08_InterventoLavorazione"
sSQL = sSQL & " WHERE IDRV_PO08_InterventoLavorazione=" & IDLavorazione


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_VALORE_CAMPO_LAVORAZIONE = 0
Else
    GET_VALORE_CAMPO_LAVORAZIONE = fnNotNullN(rs.adoColumns(Campo).Value)
End If


End Function
Private Sub INCREMENTO_NUMERO_SEZIONALE(IDSezionale As Long, IDPeriodoIVA As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset
    sSQL = "SELECT ProgressivoDisponibile FROM ProgressivoSezionale "
    sSQL = sSQL & "WHERE IDSezionale=" & IDSezionale
    sSQL = sSQL & " AND IDPeriodoIVA=" & IDPeriodoIVA
    sSQL = sSQL & " AND IDTipoModulo=" & 1
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, CnDMT.InternalConnection, adOpenDynamic, adLockPessimistic
    
    If Not rs.EOF Then
        rs!ProgressivoDisponibile = fnNotNullN(rs!ProgressivoDisponibile) + 1
        rs.Update
    End If
rs.Close
Set rs = Nothing
End Sub
Private Sub MovimentazioneDocumento(IDTipoOggetto As Long, IDOggetto As Long)
Dim rsPar As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim rsMacro As ADODB.Recordset
Dim rsDettaglio As ADODB.Recordset
Dim OLDCursor As Long
Set Mov = New DmtMovim.cMovimentazione

Set Mov.Connection = TheApp.Database.Connection


OLDCursor = CnDMT.CursorLocation
CnDMT.CursorLocation = adUseClient

'''''''''''ELIMINAZIONE MOVIMENTI CREATI DAL DOCUMENTO''''''''''''''''''''''''''''
If EliminaMovimenti(IDOggetto, IDTipoOggetto) = False Then
    MsgBox "Non è riuscita l'eliminazione dei movimenti di questo documento", vbCritical, "Eliminazione movimenti"
    Set Mov = Nothing
    Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


sSQL = "SELECT * FROM RV_PO08_ParametriMagazzino "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto

Set rsPar = CnDMT.OpenResultset(sSQL)

While Not rsPar.EOF
    MOVIMENTAZIONE_RIGHE_DOCUMENTO IDOggetto, fnNotNullN(rsPar!IDTipoOggetto), IDOggetto, fnNotNullN(rsPar!IDFunzione), fncTrovaDescrizioneFunzione("RV_PO08_Intervento"), fnNotNullN(rsPar!IDMagazzino)
rsPar.MoveNext
Wend

rsPar.CloseResultset
Set rsPar = Nothing

Set Mov = Nothing
CnDMT.CursorLocation = OLDCursor

End Sub
Private Function GeneraMovimento(IDOggetto As Long, IDTipoOggetto As Long, IDFunzione As Long, IDMagazzino As Long, _
IDArticolo As Long, Articolo As String, IDAnagrafica As Long, IDUnitaDiMisura As Long, Quantita As Double, _
NumeroDocumento As Long, DataDocumento As String, ImportoTotaleVendita As Double, ImportoUnitarioVendita As Double, _
IDValoriOggettoDettaglio As Long, IDRigaRiferimentoMacroAttivita As Long, Funzione As String, IDAnagraficaOperatore As Long, _
TotaleCostoRiga As Double, TotaleRicavoRiga As Double, TotaleConcordatoRiga As Double, _
TotaleCostoMacroAttivita As Double, TotaleRicavoMacroAttivita As Double, TotaleConcordatoMacroAttivita As Double, _
IDParcoNatanti As Long, IDParcoNatantiRighe As Long, IDRiferimentoPreventivo As Long, IDRiferimentoIntervento As Long, DataMovimento As String) As Boolean

Dim DataPerMovimento As String



If (DataMovimento = "0.00.00") Or (Len(DataMovimento) = 0) Then
    DataPerMovimento = Date
Else
    DataPerMovimento = DataMovimento
End If

Mov.DataMovimento = DataPerMovimento
Mov.FattoreDiConversione = Null

Mov.GestioneMatricole = False
Mov.IDEsercizio = LINK_ESERCIZIO
Mov.IDTipoOggetto = IDTipoOggetto
Mov.IDOggetto = IDOggetto
Mov.IDFunzione = IDFunzione
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoEntrata = IDMagazzino
Mov.IDMagazzinoUscita = IDMagazzino
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", IDAnagrafica
Mov.Field "IDTipoAnagrafica", 2
Mov.Field "IDArticolo", IDArticolo
Mov.Field "IDUnitaDiMisura", IDUnitaDiMisura
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", Articolo
Mov.Field "QuantitaTotale", Quantita
Mov.Field "Importo", ImportoTotaleVendita
Mov.Field "DataDocumento", DataDocumento
Mov.Field "NumeroDocumento", NumeroDocumento
Mov.Field "Oggetto", Funzione
Mov.Field "IDTipoMovimento", 1
Mov.Field "PrezzoUnitario", ImportoUnitarioVendita
Mov.Field "IDValoriOggettoDettaglio", IDValoriOggettoDettaglio

Mov.Field "TipoRiga", trcNessuno
Mov.Field "RV_PO08_IDRiferimentoMacroAttivita", IDRigaRiferimentoMacroAttivita
Mov.Field "RV_PO08_IDMacroAttivitaPreventivo", IDRiferimentoPreventivo
Mov.Field "RV_PO08_IDMacroAttivitaIntervento", IDRiferimentoIntervento
Mov.Field "RV_PO08_CostoRiga", TotaleCostoRiga
Mov.Field "RV_PO08_RicavoRiga", TotaleRicavoRiga
Mov.Field "RV_PO08_ConcordatoRiga", TotaleConcordatoRiga
Mov.Field "RV_PO08_CostoMacroAttivita", TotaleCostoMacroAttivita
Mov.Field "RV_PO08_RicavoMacroAttivita", TotaleRicavoMacroAttivita
Mov.Field "RV_PO08_ConcordatoMacroAttivita", TotaleConcordatoMacroAttivita
Mov.Field "RV_PO08_IDParcoNatanti", IDParcoNatanti
Mov.Field "RV_PO08_IDParcoNatantiRighe", IDParcoNatantiRighe
Mov.Field "RV_PO08_IDAnagraficaOperatore", IDAnagraficaOperatore

GeneraMovimento = Mov.Insert





End Function

Private Function EliminaMovimenti(IDOggetto As Long, IDTipoOggetto As Long) As Boolean
Dim sSQL As String

Mov.IDTipoOggetto = IDTipoOggetto
Mov.IDOggetto = IDOggetto


EliminaMovimenti = Mov.Delete






End Function
Private Sub MOVIMENTAZIONE_RIGHE_DOCUMENTO(IDIntervento As Long, IDTipoOggetto As Long, IDOggetto As Long, IDFunzione As Long, Funzione As String, IDMagazzino As Long)
Dim rsAtt As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim PrezzoConcordatoRiga As Double



sSQL = "SELECT RV_PO08_InterventoMacroAttivita.*, RV_PO08_Intervento.IDRV_PO08_ParcoNatanti,  RV_PO08_Intervento.IDAnagraficaFattura, "
sSQL = sSQL & "RV_PO08_Intervento.NumeroDocumento, RV_PO08_Intervento.DataDocumento "
sSQL = sSQL & "FROM RV_PO08_Intervento INNER JOIN "
sSQL = sSQL & "RV_PO08_InterventoMacroAttivita ON RV_PO08_Intervento.IDRV_PO08_Intervento = RV_PO08_InterventoMacroAttivita.IDRV_PO08_Intervento "
sSQL = sSQL & "WHERE RV_PO08_InterventoMacroAttivita.IDRV_PO08_Intervento=" & IDIntervento

'sSQL = sSQL & " AND Approvato=" & fnNormBoolean(1)

Set rsAtt = New ADODB.Recordset

rsAtt.Open sSQL, CnDMT.InternalConnection

While Not rsAtt.EOF

''''''''''''''''''''''''''''''''''''''''GENERAZIONE MOVIMENTO DELLA MACRO ATTIVITA'
    GeneraMovimento IDOggetto, IDTipoOggetto, IDFunzione, IDMagazzino, fnNotNullN(rsAtt!IDArticolo), fnNotNull(rsAtt!Articolo), _
                    fnNotNullN(rsAtt!IDAnagraficaFattura), fnNotNullN(rsAtt!IDUnitaDiMisura), 1, fnNotNullN(rsAtt!NumeroDocumento), fnNotNull(rsAtt!DataDocumento), _
                    fnNotNullN(rsAtt!RicavoCalcolatoImponibile), fnNotNullN(rsAtt!RicavoCalcolatoImponibile), fnNotNullN(rsAtt!IDRV_PO08_InterventoMacroAttivita), 0, Funzione, 0, _
                    fnNotNullN(rsAtt!CostoCalcolatoImponibile), fnNotNullN(rsAtt!RicavoCalcolatoImponibile), fnNotNullN(rsAtt!PrezzoConcordatoImponibile), _
                    0, 0, 0, fnNotNullN(rsAtt!IDRV_PO08_ParcoNatanti), fnNotNullN(rsAtt!IDRV_PO08_ElementoNatante), 0, 0, Date


''''''''''''''''''''''''''''''''''''''''DETTAGLIO MACRO ATTIVITA
    sSQL = "SELECT * FROM RV_PO08_InterventoMacroAttivitaRighe "
    sSQL = sSQL & "WHERE IDRV_PO08_InterventoMacroAttivita=" & fnNotNullN(rsAtt!IDRV_PO08_InterventoMacroAttivita)
    sSQL = sSQL & " AND NonMovimentare=" & fnNormBoolean(0)
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, CnDMT.InternalConnection

    

    While Not rs.EOF
            
        '''''''GENERAZIONE MOVIMENTO DELLA RIGA DI DETTAGLIO
        If fnNotNullN(rsAtt!RicavoCalcolatoImponibile) = 0 Then
            PrezzoConcordatoRiga = 0
        Else
            PrezzoConcordatoRiga = ((fnNotNullN(rs!ImportoTotaleVenditaNettoIva) / fnNotNullN(rsAtt!RicavoCalcolatoImponibile)) * fnNotNullN(rsAtt!PrezzoConcordatoImponibile))
        End If
        
        GeneraMovimento IDOggetto, IDTipoOggetto, IDFunzione, IDMagazzino, fnNotNullN(rs!IDArticolo), fnNotNull(rs!Articolo), _
                        fnNotNullN(rsAtt!IDAnagraficaFattura), fnNotNullN(rs!IDUnitaDiMisuraVendita), fnNotNullN(rs!Quantita), _
                        fnNotNullN(rsAtt!NumeroDocumento), fnNotNull(rsAtt!DataDocumento), fnNotNullN(rs!ImportoTotaleVenditaNettoIva), _
                        fnNotNullN(rs!ImportoUnitarioVenditaNettoIva), fnNotNullN(rs!IDRV_PO08_InterventoMacroAttivitaRighe), _
                        fnNotNullN(rsAtt!IDRV_PO08_InterventoMacroAttivita), Funzione, 0, fnNotNullN(rs!ImportoTotaleAcquistoNettoIVA), _
                        fnNotNullN(rs!ImportoTotaleVenditaNettoIva), PrezzoConcordatoRiga, _
                        fnNotNullN(rsAtt!CostoCalcolatoImponibile), fnNotNullN(rsAtt!RicavoCalcolatoImponibile), fnNotNullN(rsAtt!PrezzoConcordatoImponibile), _
                        fnNotNullN(rsAtt!IDRV_PO08_ParcoNatanti), fnNotNullN(rs!IDRV_PO08_ElementoNatante), 0, 0, fnNotNull(rs!DataInserimento)
                    
    rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    
rsAtt.MoveNext
Wend

rsAtt.Close
Set rsAtt = Nothing



End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmInizio.Show
End Sub
Private Sub CHIUDI_LAVORAZIONE(IDLavorazione As Long)
Dim sSQL As String
Dim LINK_LAVORAZIONE_CHIUSO As Long


LINK_LAVORAZIONE_CHIUSO = GET_TIPO_ARTICOLO("IDRV_PO08_StatoAttivitaPerFineLavori")

If LINK_LAVORAZIONE_CHIUSO = 0 Then Exit Sub

sSQL = "UPDATE RV_PO08_InterventoLavorazione SET "
sSQL = sSQL & "IDRV_PO08_StatoAttivita=" & LINK_LAVORAZIONE_CHIUSO & ", "
sSQL = sSQL & "DataStatoAttivita=" & fnNormDate(Date) & ", "
sSQL = sSQL & "StatoAttivita=" & fnNormString(GET_DESCRIZIONE_STATO_ATTIVITA(LINK_LAVORAZIONE_CHIUSO)) & " "
sSQL = sSQL & "WHERE IDRV_PO08_InterventoLavorazione=" & IDLavorazione

CnDMT.Execute sSQL

End Sub
Private Function GET_DESCRIZIONE_STATO_ATTIVITA(IDStatoAttivita As Long) As String
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT StatoAttivita FROM RV_PO08_StatoAttivita "
sSQL = sSQL & "WHERE IDRV_PO08_StatoAttivita=" & IDStatoAttivita

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_STATO_ATTIVITA = ""
Else
    GET_DESCRIZIONE_STATO_ATTIVITA = fnNotNull(rs!StatoAttivita)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_CAMBIO_PREZZO_LORDO_IVA(IDIntervento As Long, PrezzoLordoIva As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset

''''''''''''''''''RICALCOLO DEL DETTAGLIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PO08_InterventoMacroAttivitaRighe "
sSQL = sSQL & "WHERE IDRV_PO08_Intervento=" & IDIntervento

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic


While Not rs.EOF
    If PrezzoLordoIva = 0 Then
        rs!ImportoUnitarioVenditaNeutro = rs!ImportoUnitarioVenditaNettoIva
        rs!ImportoTotaleVenditaNeutro = rs!ImportoTotaleVenditaNettoIva
    Else
        rs!ImportoUnitarioVenditaNeutro = rs!ImportoUnitarioVenditaLordoIva
        rs!ImportoTotaleVenditaNeutro = rs!ImportoTotaleVenditaLordoIva
    End If
    rs.Update
rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''RICALCOLO DELLA MACRO ATTIVITA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PO08_InterventoMacroAttivita "
sSQL = sSQL & "WHERE IDRV_PO08_Intervento=" & IDIntervento

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic


While Not rs.EOF
    If PrezzoLordoIva = 0 Then
        rs!CostoCalcolatoNeutro = rs!CostoCalcolatoImponibile
        rs!RicavoCalcolatoNeutro = rs!RicavoCalcolatoImponibile
        rs!PrezzoConcordatoNeutro = rs!PrezzoConcordatoImponibile
    Else
        rs!CostoCalcolatoNeutro = rs!CostoCalcolatoLordoIVA
        rs!RicavoCalcolatoNeutro = rs!RicavoCalcolatoLordoIVA
        rs!PrezzoConcordatoNeutro = rs!PrezzoConcordatoLordoIVA
    End If
    rs.Update
rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


End Sub
Private Function GET_PREZZO_LORDO_IVA_COMMESSA(IDIntervento As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PrezzoLordoIva FROM RV_PO08_Intervento "
sSQL = sSQL & "WHERE IDRV_PO08_Intervento=" & IDIntervento


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREZZO_LORDO_IVA_COMMESSA = 0
Else
    GET_PREZZO_LORDO_IVA_COMMESSA = fnNotNullN(rs!PrezzoLordoIva)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PREZZO_LORDO_IVA_CLIENTE(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'''''''''''''''''PREZZI LORDO IVA''''''
sSQL = "SELECT PrezziNettiIvaInStampa "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREZZO_LORDO_IVA_CLIENTE = 0
Else
    GET_PREZZO_LORDO_IVA_CLIENTE = Abs(fnNotNullN(PrezziNettiIvaInStampa))
End If

rs.CloseResultset
Set rs = Nothing
End Function
