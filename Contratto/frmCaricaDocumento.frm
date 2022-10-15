VERSION 5.00
Begin VB.Form frmCaricaDocumento 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Carica documenti"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7170
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
   ScaleHeight     =   5775
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTipo 
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtNomeFile 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3240
      Width           =   5655
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtFileSelezionato 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2640
      Width           =   6975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Archivia"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txtAnnotazioni 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3840
      Width           =   6975
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Nome file"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   5655
   End
   Begin VB.Label Label2 
      Caption         =   "Percorso selezionato"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   "Annotazioni"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   6975
   End
End
Attribute VB_Name = "frmCaricaDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Declare Function FindExecutable Lib "shell32" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal sResult As String) As Long
Private Const MAX_PATH As Long = 260
Private Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_FILE_SUCCESS As Long = 32 'costante personale
Private Const ERROR_BAD_FORMAT As Long = 11

Private Const HKEY_CLASSES_ROOT = &H80000000

Private Declare Function PathFindExtension _
    Lib "shlwapi" _
    Alias "PathFindExtensionA" _
    (ByVal pPath As String) As Long
  
Private Declare Function lstrcpy _
    Lib "Kernel32" _
    Alias "lstrcpyA" _
    (ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Private Declare Function lstrlen _
    Lib "Kernel32" _
    Alias "lstrlenA" _
    (ByVal Ptr As Any) As Long
 
Public Function GetExt(ByVal sPath As String) As String
   GetExt = Mid$(GetStrFromPtr(PathFindExtension(sPath)), 2)
End Function

Private Function GetStrFromPtr(ByVal lpsz As Long) As String
   GetStrFromPtr = String$(lstrlen(ByVal lpsz), 0)
   lstrcpy ByVal GetStrFromPtr, ByVal lpsz
End Function

Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim sts As ADODB.Stream
If Len(Trim(Me.txtNomeFile.Text)) = "" Then
    MsgBox "Selezionare un documento", vbCritical, "Carica documento"
    Exit Sub
End If

sSQL = "SELECT * FROM RV_PODocumentazione "

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
Screen.MousePointer = 11

rsNew.AddNew
    rsNew!IDRV_PODocumentazione = fnGetNewKey("RV_PODocumentazione", "IDRV_PODocumentazione")
    rsNew!IDAzienda = LINK_AZIENDA_DOC
    rsNew!IDAnagrafica = LINK_CLIENTE_DOC
    rsNew!IDRV_POContratto = LINK_CONTRATTO_DOC
    rsNew!IDRV_POInterventoPadre = LINK_INTERVENTO_PADRE_DOC
    rsNew!IDRV_POIntervento = LINK_INTERVENTO_DOC
    rsNew!PercorsoOriginale = Me.File1.Path
    rsNew!PercorsoCompleto = Me.txtFileSelezionato.Text
    rsNew!NomeFile = Me.txtNomeFile.Text
    rsNew!TipoFile = Me.txtTipo.Text
    rsNew!Annotazioni = Me.txtAnnotazioni.Text
    rsNew!Archiviato = Me.Check1.Value
    rsNew!IDUtenteInserimento = LINK_UTENTE
    rsNew!MacchinaInserimento = GET_NOMECOMPUTER
    rsNew!UtenteMacchinaInserimento = GET_NOMEUTENTE
    If Abs(rsNew!Archiviato) = 1 Then
        Set sts = New ADODB.Stream
        sts.Type = ADODB.adTypeBinary
        sts.Open
        sts.LoadFromFile Me.txtFileSelezionato.Text
        rsNew!Documento = sts.Read
    End If
rsNew.Update


rsNew.Close
Set rsNew = Nothing

If Not (sts Is Nothing) Then
    sts.Close
    Set sts = Nothing
End If

Screen.MousePointer = 0
Unload Me

Exit Sub
ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "Carica documento"
    Screen.MousePointer = 0
End Sub

Private Sub Dir1_Change()
On Error GoTo ERR_Dir1_Change
    Me.File1.Path = Me.Dir1.Path
    Me.File1.Refresh
Exit Sub
ERR_Dir1_Change:
    MsgBox Err.Description, vbCritical, "Carica documento"
End Sub

Private Sub Drive1_Change()
On Error GoTo ERR_Drive1_Change
    Me.Dir1.Path = Me.Drive1
Exit Sub
ERR_Drive1_Change:
    MsgBox Err.Description, vbCritical, "Carica documento"
End Sub

Private Sub File1_Click()
On Error GoTo ERR_File1_Click
    If Mid(Me.File1.Path, Len(Me.File1.Path), 1) = "\" Then
        Me.txtFileSelezionato.Text = Me.File1.Path & Me.File1.FileName
    Else
        Me.txtFileSelezionato.Text = Me.File1.Path & "\" & Me.File1.FileName
    End If
    If Len(Trim(Me.txtFileSelezionato.Text)) > 0 Then
        Me.txtNomeFile.Text = GET_NOME_FILE(Me.txtFileSelezionato.Text)
        Me.txtTipo.Text = GetExt(Me.txtNomeFile.Text)
    End If
Exit Sub
ERR_File1_Click:
    MsgBox Err.Description, vbCritical, "Info file"
End Sub

Private Function GET_NOME_FILE(Percorso As String)
On Error GoTo ERR_GET_NOME_FILE
Dim ArrayPercorso() As String

ArrayPercorso = Split(Percorso, "\")

GET_NOME_FILE = ArrayPercorso(UBound(ArrayPercorso))

Exit Function
ERR_GET_NOME_FILE:
    GET_NOME_FILE = ""
End Function
Public Function GET_NOMECOMPUTER() As String
Dim dwLen As Long
Dim strString As String
Const MAX_COMPUTERNAME_LENGTH As Long = 31
    
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    'Show the computer name
    GET_NOMECOMPUTER = strString
End Function

Function GET_NOMEUTENTE() As String
    Dim strString As String
    Dim lunghezzaStringa As Long
    lunghezzaStringa = 32
    strString = String(lunghezzaStringa, " ")
    GetUserName strString, lunghezzaStringa
    strString = Left(strString, lunghezzaStringa)
    GET_NOMEUTENTE = strString
    GET_NOMEUTENTE = Mid(GET_NOMEUTENTE, 1, Len(GET_NOMEUTENTE) - 1)
End Function

