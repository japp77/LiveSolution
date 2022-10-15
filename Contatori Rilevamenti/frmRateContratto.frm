VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmRateContratto 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Seleziona"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   6885
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   12091
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
Attribute VB_Name = "frmRateContratto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset


Private Sub GET_GRIGLIA()
On Error GoTo ERR_fnGrigliaTipoPagamento
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    sSQL = "SELECT * FROM RV_PORateContratto "
    sSQL = sSQL & "WHERE IDRV_POContratto=" & Link_Contratto
    sSQL = sSQL & " AND Fatturata=0"
    sSQL = sSQL & " AND NonFatturare=0"
    
    sSQL = sSQL & " ORDER BY DataRata"
            
    Set rsGriglia = New ADODB.Recordset
    rsGriglia.CursorLocation = adUseClient
    rsGriglia.Open sSQL, Cn.InternalConnection

    With Me.GrigliaCorpo
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear
        With Me.GrigliaCorpo.ColumnsHeader
            .Add "IDRV_PORateContratto", "IDRV_PORateContratto", dgInteger, False, 500, dgAlignRight
            .Add "IDRV_POContratto", "IDRV_PORateContratto", dgInteger, False, 500, dgAlignRight
            .Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignRight
            .Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignRight
            
            '.Add "Fatturata", "Fatturata", dgBoolean, True, 1000, dgAligncenter, True, True, False
            .Add "NumeroRata", "N°", dgInteger, True, 1000, dgAlignRight
            .Add "DataRata", "Data", dgDate, True, 2000, dgAlignleft

        End With
        Set .Recordset = rsGriglia
        .LoadUserSettings
        .Refresh
        
        
    End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaTipoPagamento:
    MsgBox Err.Description, vbCritical, "Storico rate del contratto"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    GET_GRIGLIA
End Sub
Private Sub GrigliaCorpo_DblClick()
On Error GoTo ERR_GrigliaCorpo_DblClick

    frmMain.txtDataFatturazione.Value = Me.GrigliaCorpo.AllColumns("DataRata").Value
    Unload Me
    
Exit Sub
ERR_GrigliaCorpo_DblClick:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_DblClick"
End Sub
