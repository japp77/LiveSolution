Attribute VB_Name = "Main"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                   (ByVal hwnd As Long, ByVal lpszOp As String, _
                    ByVal lpszFile As String, ByVal lpszParams As String, _
                    ByVal LpszDir As String, ByVal FsShowCmd As Long) _
                    As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long


Public Const SW_SHOWNORMAL = 1

Public Const SE_ERR_FNF = 2&
Public Const SE_ERR_PNF = 3&
Public Const SE_ERR_ACCESSDENIED = 5&
Public Const SE_ERR_OOM = 8&
Public Const SE_ERR_DLLNOTFOUND = 32&
Public Const SE_ERR_SHARE = 26&
Public Const SE_ERR_ASSOCINCOMPLETE = 27&
Public Const SE_ERR_DDETIMEOUT = 28&
Public Const SE_ERR_DDEFAIL = 29&
Public Const SE_ERR_DDEBUSY = 30&
Public Const SE_ERR_NOASSOC = 31&
Public Const ERROR_BAD_FORMAT = 11&


Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOP = 0
Public Const WM_SETREDRAW = &HB

Public NomePrg As String
Public gResource As Resource

'Varibile di connessione
Public CnDMT As DmtOleDbLib.adoConnection



Public SETUP_SERVER As Integer
Public SETUP_CONSOLE As Integer
'**********************VARIABILI GLOBALI AZIENDA**************************
Public VarIDAzienda As Long
Public VarIDAttivitaAzienda As Long
Public VarIDFiliale As Long
Public VarIDAnagraficaAzienda As Long

Public f As FileSystemObject
Public Const Percorso_BIN As String = "\BIN"
Public Const Percorso_REPORT As String = "\REPORT"

Public Const Percorso_Tabelle As String = "\SCRIPT\TABELLE"
Public Const Percorso_Config_Tabelle As String = "\SCRIPT\CONFIG"
Public Const Percorso_Viste As String = "\SCRIPT\VISTE"


Public Const IdentificativoProgramma As Long = 51
Public Password As String
Public Utente As String

Public PARTITA_IVA_LICENZA As String

Public LINK_ANAGRAFICA As Long

Public CODICE_SBLOCCO_PRODOTTO As String
Public CODICE_SBLOCCO_DMT As String

Public RAGIONE_SOCIALE As String
Public INDIRIZZO As String
Public NUMERO_CIVICO As String
Public COMUNE As String
Public LOCALITA As String
Public PROVINCIA As String
Public CAP As String


'*************************************************************************
'*********************VARIABILI GLOBALI DELLA GESTIONE ERRORI*************
Public fncErrore As String
Public Errore As String
'*************************************************************************
Private Const NUMERICTYPE = 0
Private Const STRINGTYPE = 1
Private Const DATETYPE = 2

'record inseriti da DMT e i record inseriti dall'Utente
Public Const IDMAXSEP = 10000
'Oggetto principale per la generazione di applicazioni
Public oAppGen As cgAppGenerator
'Oggetto principale per la generazione dei tipo oggetti
Public oDocTypes As cDocTypes

Public REGISTRY_KEY As String

Public Declare Function fnAnsi2Jet Lib "diamante.dll" Alias "fnAnsi2jet" (ByVal sSQL As String) As String

''''CARTELLA DATA APPLICAZIONI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Public Const CSIDL_COMMON_APPDATA = &H23
Public Const CSIDL_LOCAL_APPDATA = &H1C&
Public Const MAX_PATH = 260
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




Public Function ConnessioneADODBLib() As Boolean
On Error GoTo ERR_ConnessioneADODBLib
Dim StringaDiConnessione As String
Dim Motore As String
Dim Catalogo As String

    
    
    Set gResource = New Resource
    
    'If TipoAvvio = 0 Then
        
        REGISTRY_KEY = Trim(gResource.GetMessage(LBL_REGISTRY_KEY))
        
        Utente = GetSetting(REGISTRY_KEY, "MenuSettings", "LASTUSER")
        Password = fnCryptString(GetSetting(REGISTRY_KEY, "MenuSettings", "LASTUSERPWD"))
        
        If Right(Trim(MenuOptions.ConnectionString), 1) = ";" Then
            StringaDiConnessione = MenuOptions.ConnectionString
        Else
            StringaDiConnessione = MenuOptions.ConnectionString & ";"
        End If
    'Else
    '    Utente = GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "NomeUtente")
    '    Password = fnCryptString(GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Password"))
    '    Motore = GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Motore")
    '    Catalogo = GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Catalogo")

    '    StringaDiConnessione = "Provider=SQLOLEDB.1;Initial Catalog=" & Catalogo & ";Data Source=" & Motore & ";"
    'End If
    
    
    Set CnDMT = DmtOleDbLib.adoEnvironments(0).OpenConnection((StringaDiConnessione & "User Id=" & Utente & ";Password=" & Password))
    
    ConnessioneADODBLib = True
    
Exit Function

ERR_ConnessioneADODBLib:
    Errore = Err.Description
    ConnessioneADODBLib = False
End Function
Public Function fnGetNewKey(Tabella As String, CampoKey As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    
    'Monta la query SQL per trovare il massimo valore della chiave primaria
    sSQL = "SELECT MAX (" & CampoKey & ") AS MaxID FROM " & Tabella
    
    'Apertura del recordset
    Set rs = CnDMT.OpenResultset(fnAnsi2Jet(sSQL))
    
    'Determina il primo progressivo disponibile
    fnGetNewKey = fnNotNullN(rs.adoColumns("MaxID")) + 1
    If fnGetNewKey <= 0 Then fnGetNewKey = 1

    'Chiude il recordset e distrugge l'oggetto.
    rs.CloseResultset
    Set rs = Nothing
    
End Function



