VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAzienda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_IDTipoAnagraficaTecnicoIntRif  As Long
Private m_IDRV_POStatoInterventoChiuso As Long
Private m_IDRV_POStatoInterventoInserimento As Long
Private m_IDTipoAnagraficaTecnicoFaseRif As Long
Private m_IDRV_POStatoFaseChiusa As Long
Private m_IDRV_POStatoFaseInserimento As Long
Private m_IDRV_POTipoFaseInterventoEla As Long
Private m_IDRV_POTipoFaseInterventoMan As Long
Private m_IDPagamentoFatt As Long
Private m_IDTipoOggettoFatt As Long
Private m_IDValutaFatt As Long
Private m_IDSezionaleFatt As Long
Private m_IDMagazzinoFatt As Long
Private m_IDArticoloFatt As Long
Private m_IDTipoAnagraficaAmministratore As Long
Private m_IDRV_POTipoGestioneBuono As Long
Private m_IDRV_POTipoNumerazioneBuono As Long
Private m_IDRV_POTipoTasca As Long
Private m_UtilizzaVoispeed As Boolean
Private m_IDRV_POTipoGenerazioneInt As Long
Private m_RiportaTecContratto As Long
Private m_IDTipoAnagraficaContratto As Long
Private m_TipoAddebitoObbligatorio As Long
Private m_TipoClasseObbligatorio As Long
Private m_VisualizzaImportiProdContratto As Long
Private m_ChiudiAltriInterventi As Boolean
Private m_GestioneProdottiPerIntervento As Boolean
Private m_IDSezionaleRateContratto As Long
Private m_VisualizzaNoteClienteAut As Boolean
Private m_GenAppAutAgendaContratto As Boolean
Private m_GenAppAutOutlookContratto As Boolean
Private m_ServerSMTP As String
Private m_NomeUtenteEmail As String
Private m_PasswordUtenteEmail As String
Private m_PortaSMTP As Long
Private m_IDTipoAnagraficaInstallatore As Long
Private m_IDClienteInterventoPredefinito As Long

Private m_IDAzienda As Long
Private m_IDFiliale As Long

Public Property Get IDTipoAnagraficaTecnicoIntRif() As Long
    IDTipoAnagraficaTecnicoIntRif = m_IDTipoAnagraficaTecnicoIntRif
End Property
Public Property Get IDRV_POStatoInterventoChiuso() As Long
    IDRV_POStatoInterventoChiuso = m_IDRV_POStatoInterventoChiuso
End Property
Public Property Get IDRV_POStatoInterventoInserimento() As Long
    IDRV_POStatoInterventoInserimento = m_IDRV_POStatoInterventoInserimento
End Property
Public Property Get IDTipoAnagraficaTecnicoFaseRif() As Long
    IDRV_POStatoInterventoInserimento = m_IDTipoAnagraficaTecnicoFaseRif
End Property
Public Property Get IDRV_POStatoFaseInserimento() As Long
    IDRV_POStatoFaseInserimento = m_IDRV_POStatoFaseInserimento
End Property
Public Property Get IDRV_POStatoFaseChiusa() As Long
    IDRV_POStatoFaseChiusa = m_IDRV_POStatoFaseChiusa
End Property

Public Property Get IDRV_POTipoFaseInterventoEla() As Long
    IDRV_POTipoFaseInterventoEla = m_IDRV_POTipoFaseInterventoEla
End Property
Public Property Get IDRV_POTipoFaseInterventoMan() As Long
    IDRV_POTipoFaseInterventoMan = m_IDRV_POTipoFaseInterventoMan
End Property
Public Property Get IDPagamentoFatt() As Long
    IDPagamentoFatt = m_IDPagamentoFatt
End Property
Public Property Get IDTipoOggettoFatt() As Long
    IDTipoOggettoFatt = m_IDTipoOggettoFatt
End Property
Public Property Get IDValutaFatt() As Long
    IDValutaFatt = m_IDValutaFatt
End Property
