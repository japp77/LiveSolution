VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmRateStorico 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Storico delle rate del contratto"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13650
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
   ScaleHeight     =   5985
   ScaleWidth      =   13650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Rate non fatturate"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   9763
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
Attribute VB_Name = "frmRateStorico"
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
    
    sSQL = "SELECT * FROM RV_POIEContrattoRate "
    sSQL = sSQL & "WHERE IDRV_POContrattoPadre=" & frmMain.txtIDContrattoPadre.Value
    sSQL = sSQL & " AND IDRV_POContratto<>" & Link_Contratto
    
    If (Me.Check1.Value = vbChecked) Then
        sSQL = sSQL & " AND Fatturata=0"
    End If
    
    sSQL = sSQL & " ORDER BY NumeroRinnovo, DataRata"
    
            
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
            
            .Add "Fatturata", "Fatturata", dgBoolean, True, 1000, dgAligncenter, True, True, False
            .Add "NumeroRinnovo", "N° Rinnovo contratto", dgInteger, False, 1000, dgAlignRight
            .Add "AnnoContratto", "Anno contratto", dgInteger, False, 1000, dgAlignRight
            .Add "NumeroContratto", "N° contratto", dgInteger, False, 1000, dgAlignRight
            
            
            .Add "NumeroRata", "N°", dgInteger, True, 500, 0, True, True, False
            .Add "DataRata", "Data", dgDate, True, 1500, 0, True, True, False
            .Add "IDPagamentoRata", "IDPagamentoRata", dgInteger, False, 500, dgAlignRight
            .Add "Pagamento", "Tipo pagamento", dgchar, True, 2500, 0, True, True, False
            Set cl = .Add("ImportoRata", "Importo", dgDouble, True, 2000, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            
            .Add "IDOggettoCollegato", "IDOggettoCollegato", dgInteger, False, 500, dgAlignRight
            .Add "Oggetto", "Documento collegato", dgchar, False, 3000, dgAlignleft
            .Add "Numero", "N° doc. fatt.", dgInteger, False, 1000, dgAlignRight
            .Add "DataEmissione", "Data doc. fatt.", dgDate, False, 2000, dgAlignleft
            
            '.Add "ProvieneDaAdeguamento", "Adeguamento", dgBoolean, True, 1000, dgAligncenter, True, True, False
            .Add "NumeroAdeguamento", "Adeguamento", dgInteger, False, 2000, dgAlignRight
            
            .Add "IDRV_POProdotto", "IDRV_POProdotto", dgInteger, False, 500, dgAlignRight
            .Add "DescrizioneProdotto", "Prodotto", dgchar, False, 3000, dgAlignleft
            .Add "Matricola", "Matricola", dgInteger, dgchar, False, 3000, dgAlignleft
            
            .Add "IDRV_POContrattoProdotti", "IDRV_POContrattoProdotti", dgInteger, False, 500, dgAlignRight
            .Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignRight
            
            .Add "CodiceArticoloProdContr", "Cod. art. prod. contr.", dgchar, False, 3000, dgAlignleft
            .Add "ArticoloProdContr", "Descr. art. prod. contr.", dgInteger, dgchar, False, 3000, dgAlignleft
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

Private Sub Check1_Click()
    GET_GRIGLIA
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    GET_GRIGLIA
End Sub
