VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.3#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Begin VB.Form frmTipoContatto 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalva 
      Caption         =   "Salva"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuovo 
      Caption         =   "Nuovo"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin DmtGridCtl.DmtGrid GrigliaDMT 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3413
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin DMTDataCmb.DMTCombo cboTipocontatto 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblNominativo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo contatto"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "frmTipoContatto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs As DmtOleDbLib.adoResultset
Private Nuovo As Integer

Private Sub cmdElimina_Click()
    Dim sSQL As String
    sSQL = "DELETE FROM RV_POTipoContattoDettaglio "
    sSQL = sSQL & "WHERE IDRV_POTipoContattoDettaglio=" & Me.GrigliaDMT.AllColumns("IDRV_POTipoContattoDettaglio")
    Cn.Execute sSQL
    SettaggioGriglia
End Sub

Private Sub cmdNuovo_Click()
    Me.cboTipocontatto.WriteOn 0
    Nuovo = 0
End Sub

Private Sub cmdSalva_Click()
    Dim sSQL As String
    
    If Nuovo = 0 Then
        sSQL = "INSERT INTO RV_POTipoContattoDettaglio ("
        sSQL = sSQL & "IDRV_POTipoContattoDettaglio, IDRV_POContactDettaglio, IDRV_POTipoContatto) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnGetNewKey("RV_POTipoContattoDettaglio", "IDRV_POTipoContattoDettaglio") & ", "
        sSQL = sSQL & ID_ContactDettaglio & ", "
        sSQL = sSQL & Me.cboTipocontatto.CurrentID & ")"
    Else
        sSQL = "UPDATE RV_POTipoContattoDettaglio SET "
        sSQL = sSQL & "IDRV_POTipoContatto=" & Me.cboTipocontatto.CurrentID
        sSQL = sSQL & "WHERE IDRV_POTipoContattoDettaglio=" & Me.GrigliaDMT.AllColumns("IDRV_POTipoContattoDettaglio")
    End If
    
    Cn.Execute sSQL
    SettaggioGriglia
End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    Me.Caption = frmMain.CDCliente.Description & " " & frmMain.CDCliente.Code
    Me.lblNominativo.Caption = frmMain.GridDettaglio.AllColumns("Nominativo").Value
    fncTipoContatto
    SettaggioGriglia
    
    
    
End Sub
Private Sub SettaggioGriglia()

    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    
    sSQL = "SELECT RV_POTipoContattoDettaglio.IDRV_POTipoContattoDettaglio, RV_POTipoContattoDettaglio.IDRV_POContactDettaglio, "
sSQL = sSQL & "RV_POTipoContattoDettaglio.IDRV_POTipoContatto , RV_POTipoContatto.TipoContatto "
sSQL = sSQL & "FROM RV_POTipoContattoDettaglio LEFT OUTER JOIN "
sSQL = sSQL & "RV_POTipoContatto ON dbo.RV_POTipoContattoDettaglio.IDRV_POTipoContatto = dbo.RV_POTipoContatto.IDRV_POTipoContatto "
sSQL = sSQL & "WHERE IDRV_POContactDettaglio=" & ID_ContactDettaglio
    
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
        Set rs = Cn.OpenResultset(sSQL)
            Set rsEvent = rs.Data
        
        With Me.GrigliaDMT
            .ColumnsHeader.Clear
                    
                    .ColumnsHeader.Add "TipoContatto", "Tipo Contatto", dgchar, True, 3000, dgAlignleft
                    
            Set .Recordset = rs.Data
            .Refresh
            
        End With
    
    Cn.CursorLocation = OLDCursor
    
    If rs.EOF Then
        Nuovo = 0
        Me.cboTipocontatto.WriteOn 0
    Else
        Nuovo = 1
    End If
End Sub
Private Sub GrigliaDMT_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    Me.cboTipocontatto.WriteOn Me.GrigliaDMT.AllColumns("IDRV_POTipoContatto")
    Nuovo = 1
End Sub
Private Sub fncTipoContatto()
    With Me.cboTipocontatto
        Set .Database = Cn
        .AddFieldKey "IDRV_POTipoContatto"
        .DisplayField = "TipoContatto"
        .SQL = "SELECT * FROM RV_POTipoContatto"
        .Fill
    End With

End Sub
