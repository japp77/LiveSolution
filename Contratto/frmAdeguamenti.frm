VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmAdeguamenti 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Adeguamenti storici accorpati"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNoteAdeg 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   6000
      Width           =   10215
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9975
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
   Begin VB.Label Label1 
      Caption         =   "Annotazioni adeguamento"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   9975
   End
End
Attribute VB_Name = "frmAdeguamenti"
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
    
    sSQL = "SELECT * FROM RV_POIEAdeguamentiContratto "
    sSQL = sSQL & "WHERE IDRV_POContrattoPadre=" & frmMain.txtIDContrattoPadre.Value
    sSQL = sSQL & " AND IDRV_POContratto<>" & Link_Contratto
    sSQL = sSQL & " AND IDRV_POTipoAdeguamento=1"
            
    Set rsGriglia = New ADODB.Recordset
    rsGriglia.CursorLocation = adUseClient
    rsGriglia.Open sSQL, Cn.InternalConnection

    With Me.GrigliaCorpo
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDRV_POContrattoAdeguamento", "IDRV_POContrattoAdeguamento", dgInteger, False, 500, dgAlignRight, True, True, False
            .ColumnsHeader.Add "IDRV_POContratto", "IDRV_POContratto", dgInteger, False, 500, dgAlignRight, True, True, False
            .ColumnsHeader.Add "IDRV_POContrattoPadre", "IDRV_POContrattoPadre", dgInteger, False, 500, dgAlignRight, True, True, False
            .ColumnsHeader.Add "IDRV_POTipoAdeguamento", "IDRV_POTipoAdeguamento", dgInteger, False, 500, dgAlignRight, True, True, False
            .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignRight, True, True, False
            .ColumnsHeader.Add "NumeroAdeguamento", "Numero", dgInteger, True, 1200, dgAlignRight, True, True, False
            .ColumnsHeader.Add "DataStipula", "Data stipula adeg.", dgDate, True, 1500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "DataDecorrenza", "Data decorrenza adeg.", dgDate, True, 1500, dgAlignleft, True, True, False
            Set cl = .ColumnsHeader.Add("Importo", "Importo adeg.", dgDouble, True, 2000, dgAlignRight, True, True, False)
                cl.Editable = True
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .ColumnsHeader.Add "NumeroRinnovo", "Numero rinnovo", dgInteger, True, 1500, dgAlignRight, True, True, False
            Set cl = .ColumnsHeader.Add("ImportoContrattoAttuale", "Importo contratto", dgDouble, True, 2000, dgAlignRight, True, True, False)
                cl.Editable = True
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .ColumnsHeader.Add "Annotazioni", "Annotazioni", dgchar, False, 1500, dgAlignleft, True, True, False
            
        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaTipoPagamento:
    MsgBox Err.Description, vbCritical, "Adeguamenti storici del contratto"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    GET_GRIGLIA
End Sub

Private Sub GrigliaCorpo_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
On Error GoTo ERR_GrigliaCorpo_Reposition
    Me.txtNoteAdeg.Text = fnNotNull(Me.GrigliaCorpo.AllColumns("Annotazioni").Value)
Exit Sub
ERR_GrigliaCorpo_Reposition:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_Reposition"
End Sub
