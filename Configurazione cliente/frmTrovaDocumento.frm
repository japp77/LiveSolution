VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Begin VB.Form frmTrovaDocumento 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "TROVA DOCUMENTO"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8415
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
   ScaleHeight     =   5025
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DMTDataCmb.DMTCombo cboTipoDocumento 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DmtGridCtl.DmtGrid GrigliaDoc 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7646
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
      Caption         =   "Tipo documento"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmTrovaDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGrigliaDoc As ADODB.Recordset


'Oggetto utilizzato per gestire l'inserimento / variazione del documento (DmtDocs.Dll)
Private oDoc As DmtDocs.cDocument
'Variabile utilizzata per ottenere il nome della tabella di testata del documento
Private sTabellaTestata As String
'Variabile utilizzata per ottenere il nome della tabella di dettaglio del documento
Private sTabellaDettaglio As String
'Variabile utilizzata per ottenere il nome della tabella delle scadenze del documento
Private sTabellaScadenze As String
'Variabile utilizzata per ottenere il nome della tabella del castelletto IVA del documento
Private sTabellaIVA As String

Private Sub GET_GRIGLIA()
On Error GoTo ERR_fnGrigliaTipoPagamento
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    
    
    If Len(sTabellaTestata) > 0 Then
    
        sSQL = "SELECT " & sTabellaTestata & ".Link_Nom_anagrafica, Oggetto.IDOggetto, Oggetto.IDTipoOggetto, Oggetto.Oggetto, Oggetto.DataEmissione, "
        sSQL = sSQL & "Oggetto.Numero , Oggetto.IDAzienda, " & sTabellaTestata & ".Tot_documento_corr "
        sSQL = sSQL & "FROM Oggetto INNER JOIN "
        sSQL = sSQL & sTabellaTestata & " ON Oggetto.IDOggetto = " & sTabellaTestata & ".IDOggetto AND "
        sSQL = sSQL & "Oggetto.IDTipoOggetto = " & sTabellaTestata & ".IDTipoOggetto "
        sSQL = sSQL & " WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
        sSQL = sSQL & " AND " & sTabellaTestata & ".Link_Nom_Anagrafica=" & frmMain.CDCliente.KeyFieldID
        sSQL = sSQL & " ORDER BY Oggetto.DataEmissione DESC, Oggetto.Numero DESC"
    
    Else
        
        sSQL = "SELECT ValoriOggettoPerTipo0002.Link_Nom_anagrafica, Oggetto.IDOggetto, Oggetto.IDTipoOggetto, Oggetto.Oggetto, Oggetto.DataEmissione, "
        sSQL = sSQL & "Oggetto.Numero , Oggetto.IDAzienda, ValoriOggettoPerTipo0002.Tot_documento_corr "
        sSQL = sSQL & "FROM Oggetto INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0002 ON Oggetto.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto AND "
        sSQL = sSQL & "Oggetto.IDTipoOggetto = ValoriOggettoPerTipo0002.IDTipoOggetto "
        sSQL = sSQL & " WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
        sSQL = sSQL & " AND ValoriOggettoPerTipo0002.Link_Nom_Anagrafica=" & 0
        sSQL = sSQL & " ORDER BY Oggetto.DataEmissione DESC, Oggetto.Numero DESC"
    
    End If
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
    
        
        Set rsGrigliaDoc = New ADODB.Recordset
        rsGrigliaDoc.CursorLocation = adUseClient
        rsGrigliaDoc.Open sSQL, Cn.InternalConnection
        
        With Me.GrigliaDoc
            .EnableMove = True
            .UpdatePosition = True
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
                    .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "Link_nom_anagrafica", "IDAnagrafica", dgInteger, False, 500, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "Oggetto", "Documento collegato", dgchar, True, 2500, dgAlignleft, True, True, False
                    .ColumnsHeader.Add "DataEmissione", "Data doc.", dgDate, True, 1800, dgAlignRight, True, True, False
                    .ColumnsHeader.Add "Numero", "Numero doc.", dgchar, True, 1000, dgAlignleft, True, True, False
                    Set cl = .ColumnsHeader.Add("Tot_documento_corr", "Totale documento", dgDouble, True, 2000, dgAlignRight, True, True, False)
                        cl.Editable = True
                        cl.FormatOptions.FormatNumericRegionalSettings = False
                        cl.FormatOptions.UseFormatControlSettings = False
                        'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                        cl.FormatOptions.FormatNumericDecSep = ","
                        cl.FormatOptions.FormatNumericDecimals = 2
                        cl.FormatOptions.FormatNumericThousandSep = "."
                    
            Set .Recordset = rsGrigliaDoc
            .Refresh
            .LoadUserSettings
        End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaTipoPagamento:
    MsgBox Err.Description, vbCritical, "Reperimento dati carico ticket cliente"

End Sub

Private Sub cboTipoDocumento_Click()

    If Me.cboTipoDocumento.CurrentID > 0 Then
        If oDoc Is Nothing Then
            Set oDoc = New DmtDocs.cDocument
            Set oDoc.Connection = TheApp.Database.Connection
            
            oDoc.SetTipoOggetto Me.cboTipoDocumento.CurrentID
            
            oDoc.IDFunzione = GET_LINK_FUNZIONE(Me.cboTipoDocumento.CurrentID)
            
            oDoc.TablesNames oDoc.IDTipoOggetto, sTabellaTestata, sTabellaDettaglio, sTabellaIVA, sTabellaScadenze
            
            Set oDoc = Nothing
        End If
    Else
        sTabellaTestata = ""
    End If
    
    GET_GRIGLIA
    
End Sub

Private Sub Form_Load()
    'tipo documento
    With Me.cboTipoDocumento
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDTipoOggetto"
        .DisplayField = "Oggetto"
        .SQL = "SELECT * FROM IERepTipoDocumento "
        .SQL = .SQL & "WHERE IDTipoAnagrafica=" & 2
        .SQL = .SQL & " ORDER BY Oggetto"
        .Fill
    End With
    
    
    
    Me.cboTipoDocumento.WriteOn 107
    
    
    LINK_OGGETTO_COLLEGATO = 0
    LINK_TIPO_OGGETTO_COLLEGATO = 0
    
    

End Sub

Private Sub GrigliaDoc_DblClick()
    LINK_TIPO_OGGETTO_COLLEGATO = Me.GrigliaDoc.AllColumns("IDTipoOggetto").Value
    LINK_OGGETTO_COLLEGATO = Me.GrigliaDoc.AllColumns("IDOggetto").Value
    Unload Me
End Sub
Private Function GET_LINK_FUNZIONE(IDTipoOggetto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDFunzione FROM Funzione "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_FUNZIONE = 0
Else
    GET_LINK_FUNZIONE = fnNotNullN(rs!IDFunzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
