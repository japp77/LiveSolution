Attribute VB_Name = "Main"
Public NomePrg As String
Public gResource As Resource

'Varibile di connessione
Public CnDMT As DmtOleDbLib.adoConnection

Public cn As ADODB.Connection

Public SETUP_SERVER As Integer
Public SETUP_CONSOLE As Integer
'**********************VARIABILI GLOBALI AZIENDA**************************
Public VarIDAzienda As Long
Public VarIDAttivitaAzienda As Long
Public VarIDFiliale As Long
Public VarIDAnagraficaAzienda As Long

Public F As FileSystemObject
Public Const Percorso_BIN As String = "\BIN"
Public Const Percorso_REPORT As String = "\REPORT"

Public Const Percorso_Tabelle As String = "\SCRIPT\TABELLE"
Public Const Percorso_Config_Tabelle As String = "\SCRIPT\CONFIG"
Public Const Percorso_Viste As String = "\SCRIPT\VISTE"

Public Const NomeProgramma As String = "Contratti professional"
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


Public Const REGISTRY_KEY_PERS As String = "Point Office Company srl"
Public Const SECTION_REGISTRY_KEY_PERS As String = "ContrattiPRO"
Public TipoAvvio As Long


Public Declare Function fnAnsi2Jet Lib "diamante.dll" Alias "fnAnsi2jet" (ByVal sSQL As String) As String
Public Function ConnessioneADODBLib() As Boolean
On Error GoTo ERR_ConnessioneADODBLib
Dim StringaDiConnessione As String
Dim Motore As String
Dim Catalogo As String

    
    TipoAvvio = fnNotNullN(GetSetting(REGISTRY_KEY_PERS, SECTION_REGISTRY_KEY_PERS, "Tipo Avvio"))
    
    Set gResource = New Resource
    
    'If TipoAvvio = 0 Then
        
        REGISTRY_KEY = Trim(gResource.GetMessage(LBL_REGISTRY_KEY))
        
        Utente = GetSetting(REGISTRY_KEY, "MenuSettings", "LASTUSER")
        Password = fnCryptString(GetSetting(REGISTRY_KEY, "MenuSettings", "LASTUSERPWD"))
        
        If right(Trim(MenuOptions.ConnectionString), 1) = ";" Then
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
    
    
    Set cn = New ADODB.Connection
    
    cn = CnDMT.InternalConnection
    
    
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



Public Sub PrelevaAzienda()
Dim TmpFiliale As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
    TmpFiliale = GetSetting(REGISTRY_KEY, "MenuSettings", "LASTBRANCH")
    
    sSQL = "SELECT Anagrafica.IDAnagrafica, Azienda.IDAzienda, Anagrafica.Anagrafica, AttivitaAzienda.IDAttivitaAzienda, AttivitaAzienda.AttivitaAzienda, Filiale.IDFiliale, Filiale.Filiale"
    sSQL = sSQL & " FROM (Anagrafica INNER JOIN Azienda ON Anagrafica.IDAnagrafica = Azienda.IDAnagrafica) INNER JOIN (Filiale INNER JOIN AttivitaAzienda ON Filiale.IDAttivitaAzienda = AttivitaAzienda.IDAttivitaAzienda) ON Azienda.IDAzienda = AttivitaAzienda.IDAzienda"
    sSQL = sSQL & " WHERE (((Filiale.IDFiliale)=" & MenuOptions.LastBranch & "))"
    
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        VarIDAzienda = fnNotNullN(rs!IDAzienda)
        VarIDAttivitaAzienda = fnNotNullN(rs!IDAttivitaAzienda)
        VarIDFiliale = fnNotNullN(rs!IDFiliale)
        VarIDUtente = MenuOptions.LastUserID
        VarIDAnagraficaAzienda = fnNotNullN(rs!IDAnagrafica)
    rs.CloseResultset
    Set rs = Nothing
    End If
    
End Sub

Public Function fnGetNewKeyPerTipoOggetto(Tabella As String, CampoKey As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    
    'Monta la query SQL per trovare il massimo valore della chiave primaria
    sSQL = "SELECT MAX (" & CampoKey & ") AS MaxID FROM " & Tabella
   
    'Apertura del recordset
    Set rs = CnDMT.OpenResultset(fnAnsi2Jet(sSQL))
    
    If fnNotNullN(rs.adoColumns("MaxID")) < 10000 Then
        fnGetNewKeyPerTipoOggetto = 10000
    Else
        fnGetNewKeyPerTipoOggetto = fnNotNullN(rs.adoColumns("MaxID")) + 1
    End If

    rs.CloseResultset
    Set rs = Nothing
    
    
End Function
