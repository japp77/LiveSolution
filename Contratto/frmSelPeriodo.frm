VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmSelPeriodo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleziona periodo"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelPeriodo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5530
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
Attribute VB_Name = "frmSelPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private Sub Form_Load()
    LINK_UM_PERIODO_SEL = 0
    DESCRIZIONE_UM_PERIODO_SEL = ""
    GET_GRIGLIA
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    Set rsGriglia = New ADODB.Recordset
    rsGriglia.CursorLocation = adUseClient
    
    sSQL = "SELECT * FROM RV_POUnitaDiMisuraPeriodo "
    sSQL = sSQL & "WHERE VisualizzaInContatore=1"
    
    rsGriglia.Open sSQL, Cn.InternalConnection
    
    
    With Me.GrigliaCorpo
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear

            .ColumnsHeader.Add "IDRV_POUnitaDiMisuraPeriodo", "IDRV_POUnitaDiMisuraPeriodo", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Descrizione", "Descrizione", dgchar, True, 3500, dgAlignleft

        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With
    Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "Reperimento dati U.M."

Exit Sub

End Sub


Private Sub GrigliaCorpo_DblClick()
On Error GoTo ERR_GrigliaCorpo_DblClick
    LINK_UM_PERIODO_SEL = Me.GrigliaCorpo.AllColumns("IDRV_POUnitaDiMisuraPeriodo").Value
    DESCRIZIONE_UM_PERIODO_SEL = Me.GrigliaCorpo.AllColumns("Descrizione").Value
    Unload Me

Exit Sub
ERR_GrigliaCorpo_DblClick:
    MsgBox Err.Description, vbCritical, "GrigliaCorpo_DblClick"
End Sub
