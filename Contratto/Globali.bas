Attribute VB_Name = "Globali"
Option Explicit

'Declares
Public Declare Function fnAnsi2Jet Lib "Diamante.dll" Alias "fnAnsi2jet" (ByVal sSQL As String) As String
Public Declare Sub sbOpenURL Lib "Diamante.dll" (ByVal hwnd As Long, ByVal sURL As String)
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOP = 0
Public Const WM_SETREDRAW = &HB

'Costanti globali
Public Const TOTAL_CONTROLS_NUMBER = 10
Public Const SPLITLIMIT = 1000
Public Const SRCNEXT = 1
Public Const SRCPREVIOUS = 2
Public Const HELP_FINDER = &HB
Public Const HELP_CONTEXT = &H1
Public Const URL_DIAMANTE = "http://www.diamante.it"

'*** Costanti per la gestione della Attivazione-Disattivazione Menu e ToolBar
Public Const BTN_NEW = 1
Public Const BTN_SAVE = 2
Public Const BTN_PRINT = 4
Public Const BTN_PREVIEW = 8
Public Const BTN_CUT = 16
Public Const BTN_COPY = 32
Public Const BTN_PASTE = 64
Public Const BTN_DELETE = 128
Public Const BTN_CLEAR = 256
Public Const BTN_FIND = 512
Public Const BTN_SEARCH = 1024
Public Const BTN_VIEWMODE = 2048
Public Const BTN_PREVIOUS = 4096
Public Const BTN_NEXT = 8192
Public Const BTN_WORD = 16384
Public Const BTN_EXCEL = 32768
Public Const BTN_HTML = 65536
Public Const BTN_SEARCHFORM = 131072
Public Const BTN_SEARCHTABLE = 262144
Public Const BTN_FILTER = 262144 * 2
Public Const BTN_TOOLS = BTN_FILTER * 2
Public Const BTN_PDF = BTN_TOOLS * 2
Public Const BTN_EXPORT = BTN_PDF * 2
Public Const BTN_ALL = BTN_EXPORT * 2 - 1

'Il nome della ToolBar dell'Anteprima di stampa
Public Const BAND_CLOSE_PREVIEW = "Band_ClosePreview"

'Elenco errori
Public Const ERR_TABLE_STRUCT = vbObjectError + 10000
Public Const ERR_NO_DEFAULT_TABLEVIEW = vbObjectError + 10001
Public Const ERR_NO_PROCESSES = vbObjectError + 10002
Public Const ERR_NDELFILTER = vbObjectError + 2500

'La variabile globale TheApp mantiene un riferimento all'oggetto
'applicazione che viene utilizzato per eseguire le funzionalità
'ed i relativi processi del gestore.
Public TheApp As Application

'La variabile globale gResource mantiene un riferimento all'oggetto
'utilizzato per l'accesso alle risorse stringa, icon e bitmap di Diamante
Public gResource As Resource

Public Cn As DmtOleDbLib.adoConnection
Public Db As DMTDataLayer.Database


Public REGISTRY_KEY As String

'Variabili per il piano dei conti del contratto
Public VarIDEsercizio As Long
Public Link_ContoPDC As Long
Public oPDC As DmtPDC.PDCServices
Public Link_PianoDeiConti As Long
Public Link_Agente As Long
Public Link_Commesso As Long

Public RIPORTA_SERVIZI As Boolean

'''''VARIABILI PER CODA''''''''''''''''''''
Public APERTURA_FORM_CODA As Boolean
Public NOME_GESTORE As String
''''''''''''''''''''''''''''''''''''''''''

''''''VARIABILI PARAMETRI AZIENDA'''''''''''''''''''''''
Public LINK_TIPO_ANA_TEC_INT As Long
Public LINK_STATO_INT_NUOVO As Long
Public LINK_STATO_INT_CHIUSO As Long

Public LINK_TIPO_ANA_TEC_FASE As Long
Public LINK_STATO_FASE_NUOVA As Long
Public LINK_STATO_FASE_CHIUSA As Long
Public LINK_TIPO_FASE_ELA As Long
Public LINK_TIPO_FASE_MANUALE As Long

Public LINK_TIPO_ANA_AMM As Long
Public LINK_TIPO_GENERA_INTERVENTO As Long

Public LINK_TIPO_TASCA As Long
Public LINK_TIPO_ANA_TEC_CONTRATTO As Long
Public FLAG_RIPORTA_TEC_CONTRATTO As Long
Public LINK_TIPO_GESTIONE_BUONI As Long
Public LINK_TIPO_NUMERAZIONE_BUONO As Long
Public VISUALIZZA_IMPORTI_PROD As Long
Public LINK_SEZIONALE_RATE As Long
Public SEL_PROD_NON_CONTRATTO As Long
Public NON_MODIFICA_CONTRATTO As Long


Public LINK_AZIENDA_DOC As Long 'Identificativo dell'azienda
Public LINK_CLIENTE_DOC As Long 'Identificativo del cliente
Public LINK_CONTRATTO_DOC As Long 'Identificativo del contratto
Public LINK_CONTRATTO_DOC_PADRE As Long 'Identificativo del contratto
Public LINK_INTERVENTO_PADRE_DOC As Long 'Identificativo dell'intervento padre
Public LINK_INTERVENTO_DOC As Long 'Identificativo dell'intervento
Public LINK_STARTUP_DOC As Long 'Identificativo il programma di avvio della funzione
'1 = Azienda
'2 = Cliente
'3 = Contratto
'4 = Intervento


'Oggetto utilizzato per gestire l'inserimento / variazione del documento (DmtDocs.Dll)
Public oDoc As DmtDocs.cDocument
'Variabile utilizzata per ottenere il nome della tabella di testata del documento
Public sTabellaTestata As String
'Variabile utilizzata per ottenere il nome della tabella di dettaglio del documento
Public sTabellaDettaglio As String
'Variabile utilizzata per ottenere il nome della tabella delle scadenze del documento
Public sTabellaScadenze As String
'Variabile utilizzata per ottenere il nome della tabella del castelletto IVA del documento
Public sTabellaIVA As String

Public LINK_LISTINO_CLIENTE As Long
Public LINK_LISTINO_AZIENDA As Long

Public LINK_ANAGRAFICA_TEL As Long
Public LINK_SITO_PER_ANAGRAFICA_TEL As Long
Public LINK_GRUPPO As Long

Public CREA_APPUNTAMENTO_AGENDA As Long

Public FILTRO_PRODOTTO_ASSOCIATO As Long

Public LINK_TIPO_IMPOSTAZIONE As Long
Public LINK_UM_PERIODO_AZIENDA As Long

Public LINK_PRODOTTO_SEL As Long
Public LINK_PRODOTTO_RIGA_SEL As Long

Public LINK_UM_PERIODO_SEL As Long
Public DESCRIZIONE_UM_PERIODO_SEL As String

Public LINK_OGGETTO_CONTRATTO_SEL As Long
Public LINK_TIPO_OGGETTO_CONTRATTO_SEL As Long
Public OPERAZIONE_ESEGUITA_ACCONTO As Long

Public LINK_ARTICOLO_SERVIZIO As Long
Public LINK_ARTICOLO_ADDEBITO As Long


'GLOBAL LICENZA
Public MODULO_ATTIVATO As Long
Public MODULO_DESCRIZIONE As String
Public Const MODULO_CODICE As String = "LS001"


Public MODULO_ATTIVATO_INT As Long
Public MODULO_DESCRIZIONE_INT As String
Public Const MODULO_CODICE_INT As String = "LS002"

Public MODULO_ATTIVATO_NOL As Long
Public MODULO_DESCRIZIONE_NOL As String
Public Const MODULO_CODICE_NOL As String = "LS004"

Public MODULO_ATTIVATO_CONT As Long
Public MODULO_DESCRIZIONE_CONT As String
Public Const MODULO_CODICE_CONT As String = "LS005"

Public VIS_FORM_GEN_INT_DA_PROD As Long
Public FORM_GEN_INT_DA_PROD As Long

Public NO_CALCOLO_PERIODO_FATT As Long

Public IMPORTO_CONTRATTO_ISTAT As Double

Public AGGIORNA_DA_ISTAT As Long

Public IMPORTO_REG_IMP As Double
Public LINK_ARTICOLO_REG_IMP As Double

Public NON_IMPOSTARE_FILTRI As Long
Public DISDETTO_NONFATTURARE As Long

Public COPIA_NUOVA_RIGA_PROD As Long
Public OBBLIGATORIO_COLLEGAMENTO_INT As Long
Public NON_VISUALIZZARE_ALTRI_DATI As Long


