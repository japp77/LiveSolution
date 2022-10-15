Attribute VB_Name = "Globali"
Option Explicit

'API di uso comune.
Public Declare Function fnAnsi2Jet Lib "Diamante.dll" Alias "fnAnsi2jet" (ByVal sSQL As String) As String
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long)
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long


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
Public Const NATIVE_LANGUAGE = 1
Public Const URL_DIAMANTE = "http://www.diamante.it"


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

'Oggetto Semaforo usato per gestire i conflitti di multiutenza.
Public gSemaphore As Semaforo.dmtSemaphore



'La variabile globale Application_Name è valorizzata nella Sub Main.
Public Application_Name As String


'La variabile globale Current_Process_ID è valorizzata nella Sub Main
'e rappresenta l'ID del processo in esecuzione.
Public Current_Process_ID As Long

'////////////////////////////////////////////////////////
'Impostare questa costante con il nome del
'processo previsto per questa manutenzione
'////////////////////////////////////////////////////////
Public Const PROCESS_NAME = "Manutenzione"


'Variabile di connessione
Public CnDMT As DmtOleDbLib.adoConnection




'**********************VARIABILI GLOBALI AZIENDA**************************
    Public VarIDAzienda As Long
    Public VarIDAttivitaAzienda As Long
    Public VarIDFiliale As Long
    Public VarIDEsercizio As Long
    Public VarIDUtente As Long
    
'*************************************************************************

Public NuovoRaggruppamento As Boolean

Public APERTURA_FORM_ASSOCIAZIONE_INTERVENTO As Boolean
Public TIPO_INSERIMENTO As Long

