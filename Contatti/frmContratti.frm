VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.3#0"; "DmtGridCtl.ocx"
Begin VB.Form frmContratti 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GrigliaDMT 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      BackColor       =   16777215
      ForeColor       =   4194304
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
Attribute VB_Name = "frmContratti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs As DmtOleDbLib.adoResultset
Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    Me.Caption = "Contratti attivi del cliente " & frmMain.CDCliente.Description & " " & frmMain.CDCliente.Code
    If frmMain.cboFiliale.CurrentID > 0 Then
        Me.Caption = Me.Caption & " Filiale di " & frmMain.cboFiliale.Text
    End If
    SettaggioGriglia
    
End Sub
Private Sub SettaggioGriglia()
On Error GoTo ERR_SettaggioGriglia
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    sSQL = "SELECT RV_POTipoContratto.TipoContratto, RV_POContratto.DataScadenzaPerRinnovo, RV_PORateizzazione.Rateizzazione, "
    sSQL = sSQL & "RV_POContratto.ImportoContrattoAttuale , RV_POContratto.IDAnagrafica, RV_POContratto.IDSitoPerAnagrafica "
    sSQL = sSQL & "FROM RV_POContratto LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PORateizzazione ON RV_POContratto.IDRateizzazione = RV_PORateizzazione.IDRV_PORateizzazione LEFT OUTER JOIN "
    sSQL = sSQL & "RV_POTipoContratto ON RV_POContratto.IDTipoContratto = RV_POTipoContratto.IDRV_POTipoContratto "
    sSQL = sSQL & "WHERE (RV_POContratto.IDAnagrafica=" & frmMain.CDCliente.KeyFieldID & ") AND (RV_POContratto.IDSitoPerAnagrafica=" & frmMain.cboFiliale.CurrentID & ")"
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    Set rs = Cn.OpenResultset(sSQL)
        Set rsEvent = rs.Data
    
    With Me.GrigliaDMT
        .ColumnsHeader.Clear
                
        .ColumnsHeader.Add "TipoContratto", "Contratto", dgchar, True, 2000, dgAlignleft
        .ColumnsHeader.Add "DataScadenzaPerRinnovo", "Data scad. per rinnovo", dgDate, True, 2000, dgAlignleft
        .ColumnsHeader.Add "Rateizzazione", "Rateizzazione", dgchar, True, 2000, dgAlignleft
        Set cl = .ColumnsHeader.Add("ImportoContrattoAttuale", "Importo contratto", dgCurrency, True, 3000, 0, True, True, False)
            cl.FormatOptions.FormatNumericRegionalSettings = False
            cl.FormatOptions.UseFormatControlSettings = False
            cl.FormatOptions.FormatNumericCurSymbol = "€  "
            cl.FormatOptions.FormatNumericDecSep = ","
            cl.FormatOptions.FormatNumericDecimals = 2
            cl.FormatOptions.FormatNumericThousandSep = "."

        Set .Recordset = rs.Data
        .Refresh
        
    End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGriglia:
    MsgBox Err.Description, vbCritical, "SettaggioGriglia"
End Sub
