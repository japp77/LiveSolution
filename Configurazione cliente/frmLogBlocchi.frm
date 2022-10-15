VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmLogBlocchi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17700
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogBlocchi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   17700
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17655
      _ExtentX        =   31141
      _ExtentY        =   8705
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableMove      =   0   'False
      ColumnsHeaderHeight=   20
   End
End
Attribute VB_Name = "frmLogBlocchi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private rsGriglia As ADODB.Recordset

Private Sub Form_Load()
    GET_GRIGLIA
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
sSQL = "SELECT * FROM RV_POIEConfigurazioneClienteLog "
sSQL = sSQL & "WHERE IDRV_POConfigurazioneCliente=" & LINK_TESTA_OGGETTO
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
    Set rsGriglia = New ADODB.Recordset
    rsGriglia.CursorLocation = adUseClient
    rsGriglia.Open sSQL, Cn.InternalConnection
    
    With Me.Griglia
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "ID", "ID", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POConfigurazioneCliente", "IDRV_POConfigurazioneCliente", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDUtente", "IDUtente", dgNumeric, False, 500, dgAlignRight
            .ColumnsHeader.Add "Bloccato", "Bloccato", dgBoolean, True, 1500, dgAligncenter
            .ColumnsHeader.Add "DataBlocco", "Data blocco", dgDate, True, 2000, dgAlignleft
            .ColumnsHeader.Add "AnnotazioniBlocco", "Annotazioni del blocco", dgchar, True, 3500, dgAlignleft
            .ColumnsHeader.Add "Utente", "Utente", dgchar, True, 3500, dgAlignleft
            .ColumnsHeader.Add "DataInserimento", "Data modifica", dgDateAndTime, True, 2000, dgAlignleft
        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With
    
    Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "Reperimento dati log"
    
End Sub
