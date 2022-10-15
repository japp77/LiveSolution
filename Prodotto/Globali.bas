Attribute VB_Name = "Globali"
Option Explicit

'Declares
Public Declare Function fnAnsi2Jet Lib "Diamante.dll" Alias "fnAnsi2jet" (ByVal sSQL As String) As String
Public Declare Sub sbOpenURL Lib "Diamante.dll" (ByVal hwnd As Long, ByVal sURL As String)
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Declare Function PathFindExtension Lib "shlwapi" Alias "PathFindExtensionA" (ByVal pPath As String) As Long
  
Public Declare Function lstrcpy Lib "Kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Public Declare Function lstrlen Lib "Kernel32" Alias "lstrlenA" (ByVal Ptr As Any) As Long

Public Const VER_PLATFORM_WIN32_NT = 2
Public Type OSVERSIONINFO
     
     dwOSVersionInfoSize As Long
     dwMajorVersion As Long
     dwMinorVersion As Long
     dwBuildNumber As Long
     dwPlatformId As Long
     szCSDVersion As String * 128
End Type
Public Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetFileNameFromBrowseW Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, ByVal lpstrFile As Long, ByVal nMaxFile As Long, ByVal lpstrInitialDir As Long, ByVal lpstrDefExt As Long, ByVal lpstrFilter As Long, ByVal lpstrTitle As Long) As Long
Public Declare Function GetFileNameFromBrowseA Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, ByVal lpstrFile As String, ByVal nMaxFile As Long, ByVal lpstrInitialDir As String, ByVal lpstrDefExt As String, ByVal lpstrFilter As String, ByVal lpstrTitle As String) As Long



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

Public cn As DmtOleDbLib.adoConnection
Public Db As DMTDataLayer.Database


Public REGISTRY_KEY As String

Public NUOVA_RIGA As Boolean
Public rsGriglia As ADODB.Recordset

Public LINK_AZIENDA_DOC As Long 'Identificativo dell'azienda
Public LINK_CLIENTE_DOC As Long 'Identificativo del cliente
Public LINK_CONTRATTO_DOC As Long 'Identificativo del contratto
Public LINK_INTERVENTO_PADRE_DOC As Long 'Identificativo dell'intervento padre
Public LINK_INTERVENTO_DOC As Long 'Identificativo dell'intervento
Public LINK_STARTUP_DOC As Long 'Identificativo il programma di avvio della funzione

Public LINK_TIPO_ANA_INSTALLATORE As Long
Public LINK_PRODOTTO As Long
Public LINK_ARTICOLO As Long

Public PRODOTTO_DISMESSO As Boolean


'1 = Azienda
'2 = Cliente
'3 = Contratto
'4 = Intervento


'''''VARIABILI PER IMPOSTAZIONE UTENTE'''''''''''''''''''''''
Public LINK_TECNICO_UTENTE_RIF As Long
Public FLAG_INSERISCI_BUONI As Boolean
Public FLAG_VISUALIZZA_BUONI As Boolean
Public FLAG_MODIFICA_BUONI As Boolean
Public VISUALIZZA_IMPORTI As Boolean
Public SECONDI_DI_NOTIFICA As Long
Public LINK_TIPO_DURATA_NOTIFICA As Long
Public FLAG_NOTIFICA_SOLO_TEC As Boolean
Public LINK_TIPO_CATEGORIA_DEFAULT As Long
Public LINK_TIPO_CATEGORIA_NOTIFICA As Long
Public FLAG_RIPORTA_TEC_RIF As Boolean
Public FLAG_RIPORTA_TEC_OPE As Long
Public FLAG_ELIMINA_INTERVENTO As Long
Public FLAG_SEL_DOC_VEND_BUONO As Long
Public FLAG_MOD_BUONO_DOPO_FATT As Long
Public COSTO_ORARIO_OPERATORE As Double

Public FLAG_GEN_AUT_APP_OUTLOOK As Long
Public FLAG_GEN_AUT_APP_AGENDA As Long
Public FLAG_CHIEDI_CONF_APP As Long

Public LINK_GRUPPO_ANA As Long
Public FEEDBACK_UTENTE As Long
Public LINK_CLIENT_POSTA As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public LINK_TIPO_GESTIONE_BUONI As Long
Public LINK_TIPO_NUMERAZIONE_BUONO As Long



