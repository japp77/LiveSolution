VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmMatricoleArt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Matricole"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9510
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
   ScaleHeight     =   4575
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8070
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
Attribute VB_Name = "frmMatricoleArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset


Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    sSQL = "SELECT * FROM MatricolaArticolo "
    sSQL = sSQL & " WHERE IDArticolo=" & LINK_ARTICOLO
    sSQL = sSQL & " AND DataScadenza>=" & fnNormDate(Date)
    sSQL = sSQL & " AND Sospeso=" & fnNormBoolean(0)
    
    OLDCursor = cn.CursorLocation
    cn.CursorLocation = 3

    Set rsGriglia = New ADODB.Recordset
    rsGriglia.CursorLocation = adUseClient
    rsGriglia.Open sSQL, cn.InternalConnection
        
    With Me.Griglia
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear
        
        .ColumnsHeader.Add "IDMatricolaArticolo", "IDMatricolaArticolo", dgInteger, False, 500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "Codice", "Codice", dgchar, True, 3500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "MatricolaArticolo", "Descrizione", dgchar, False, 3500, dgAlignleft, True, True, False
        .ColumnsHeader.Add "DataScadenza", "Data scadenza", dgDate, True, 1500, dgAlignleft, True, True, False

                
        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With
    
    cn.CursorLocation = OLDCursor


Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"

End Sub

Private Sub Form_Load()
    
    GET_GRIGLIA
    
End Sub

Private Sub Griglia_DblClick()
    frmMain.txtIDMatricola.Value = Me.Griglia.AllColumns("IDMatricolaArticolo").Value
    Unload Me
End Sub
