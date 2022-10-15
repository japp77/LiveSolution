VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form FrmMovimenti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anteprima fatturazione"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19470
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMovimenti.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   19470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GrigliaAnagrafica 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10095
      _ExtentX        =   17806
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
      UpdatePosition  =   0   'False
      UseUserSettings =   0   'False
      ColumnsHeaderHeight=   20
   End
   Begin VB.CommandButton cmdIndietro 
      Caption         =   "Indietro"
      Height          =   375
      Left            =   15120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   16560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Esci"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18000
      TabIndex        =   2
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CheckBox chkRaggruppaFatture 
      Caption         =   "Raggruppa rate per cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin DmtGridCtl.DmtGrid GrigliaBuoni 
      Height          =   5655
      Left            =   10200
      TabIndex        =   1
      Top             =   360
      Width           =   9135
      _ExtentX        =   16113
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
      UpdatePosition  =   0   'False
      UseUserSettings =   0   'False
      ColumnsHeaderHeight=   20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "BUONI DA FATTURARE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   10320
      TabIndex        =   7
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ANAGRAFICA CLIENTE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9975
   End
End
Attribute VB_Name = "FrmMovimenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAvanti_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    CREA_TABELLA_TEMPORANEA_ANAGRAFICA
    CREA_TABELLA_TEMPORANEA_BUONI
    
    GRIGLIA_ANAGRAFICA
    
    
End Sub

Private Sub Form_Activate()
    'Me.Caption = TheApp.FunctionName & " (Passo 2 di 4)"
    Me.SetFocus
End Sub

Private Sub CREA_TABELLA_TEMPORANEA_ANAGRAFICA()
On Error GoTo ERR_CREA_TABELLA_TEMPORANEA_ANAGRAFICA


If Not (rsAnaVis Is Nothing) Then
    If (rsAnaVis.State > 0) Then rsAnaVis.Close
    
    Set rsAnaVis = Nothing
    
End If

Set rsAnaVis = New ADODB.Recordset

rsAnaVis.CursorLocation = adUseClient

rsAnaVis.Fields.Append "IDAnagraficaCliente", adInteger, , adFldIsNullable
rsAnaVis.Fields.Append "AnagraficaCliente", adVarChar, 250, adFldIsNullable
rsAnaVis.Fields.Append "NomeCliente", adVarChar, 250, adFldIsNullable
rsAnaVis.Fields.Append "RitenutaAcconto", adSmallInt, , adFldIsNullable
rsAnaVis.Fields.Append "IDSitoPerAnagrafica", adInteger, , adFldIsNullable
rsAnaVis.Fields.Append "SitoPerAnagrafica", adVarChar, 250, adFldIsNullable
rsAnaVis.Fields.Append "Indirizzo", adVarChar, 250, adFldIsNullable
rsAnaVis.Fields.Append "Cap", adVarChar, 250, adFldIsNullable
rsAnaVis.Fields.Append "Comune", adVarChar, 250, adFldIsNullable
rsAnaVis.Fields.Append "Provincia", adVarChar, 250, adFldIsNullable

rsAnaVis.Open , , adOpenKeyset, adLockBatchOptimistic

If Not (rsAna.EOF And rsAna.BOF) Then
    rsAna.MoveFirst
    While Not rsAna.EOF
        rsAnaVis.AddNew
            rsAnaVis!IDAnagraficaCliente = rsAna!IDAnagraficaCliente
            rsAnaVis!RitenutaAcconto = rsAna!RitenutaAcconto
            rsAnaVis!IDSitoPerAnagrafica = rsAna!IDSitoPerAnagrafica
            rsAnaVis!AnagraficaCliente = GET_ANAGRAFICA(rsAna!IDAnagraficaCliente)
            rsAnaVis!SitoPerAnagrafica = GET_SITO_PER_ANAGRAFICA(rsAna!IDSitoPerAnagrafica, rsAnaVis)
        rsAnaVis.Update
    rsAna.MoveNext
    Wend
rsAna.MoveFirst
End If
Exit Sub
ERR_CREA_TABELLA_TEMPORANEA_ANAGRAFICA:
    MsgBox Err.Description, vbCritical, "CREA_TABELLA_TEMPORANEA_ANAGRAFICA"
End Sub
Private Sub CREA_TABELLA_TEMPORANEA_BUONI()
On Error GoTo ERR_CREA_TABELLA_TEMPORANEA_BUONI
Dim I As Long

If Not (rsBuoniVis Is Nothing) Then
    If (rsBuoniVis.State > 0) Then rsBuoniVis.Close
    
    Set rsBuoniVis = Nothing
    
End If


Set rsBuoniVis = New ADODB.Recordset
rsBuoniVis.CursorLocation = adUseClient

For I = 0 To rsnew.Fields.Count - 1
    rsBuoniVis.Fields.Append rsnew.Fields(I).Name, rsnew.Fields(I).Type, rsnew.Fields(I).DefinedSize, rsnew.Fields(I).Attributes
Next

rsBuoniVis.Open , , adOpenKeyset, adLockBatchOptimistic

If Not (rsnew.EOF And rsnew.BOF) Then
    rsnew.MoveFirst
    While Not rsnew.EOF
        rsBuoniVis.AddNew
            For I = 0 To rsnew.Fields.Count - 1
                rsBuoniVis.Fields(rsnew.Fields(I).Name).Value = rsnew.Fields(I).Value
            Next
        rsBuoniVis.Update
    
    rsnew.MoveNext
    Wend
End If

rsnew.MoveFirst


Exit Sub
ERR_CREA_TABELLA_TEMPORANEA_BUONI:
    MsgBox Err.Description, vbCritical, "CREA_TABELLA_TEMPORANEA_BUONI"
End Sub
Private Sub GRIGLIA_ANAGRAFICA()
On Error GoTo ERR_SettaggioGrigliaRate
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
    With Me.GrigliaAnagrafica
        .EnableMove = True
        .UpdatePosition = False
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear

        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDAnagraficaCliente", "IDAnagraficaCliente", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "AnagraficaCliente", "Anagrafica cliente", dgVarChar, True, 1700, dgAlignleft
            .ColumnsHeader.Add "RitenutaAcconto", "Ritenuta", dgBoolean, True, 1700, dgAligncenter
            If (FLAG_RAGGR_ALTRA_DEST = 1) Then
                .ColumnsHeader.Add "IDSitoPerAnagrafica", "IDSitoPerAnagrafica", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "SitoPerAnagrafica", "Altra destinazione", dgVarChar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "Indirizzo", "Indirizzo", dgVarChar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "Comune", "Comune", dgVarChar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "Provincia", "Provincia", dgVarChar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "Cap", "Cap", dgVarChar, True, 1700, dgAlignleft
            End If
                
        Set .Recordset = rsAnaVis
        
        .Refresh
        
    End With
    
    CnDMT.CursorLocation = OLDCursor
    
Exit Sub
ERR_SettaggioGrigliaRate:
    MsgBox Err.Description, vbCritical, "Settaggio griglia anagrafica da fatturare"

End Sub

Private Sub GRIGLIA_BUONI_DA_FATTURARE(IDAnagraficaCliente As Long, RitenutaAcconto As Long, IDSitoPerAnagrafica As Long)
On Error GoTo ERR_SettaggioGrigliaRate
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
    rsBuoniVis.Filter = "IDAnagraficaCliente=" & IDAnagraficaCliente
    rsBuoniVis.Filter = rsBuoniVis.Filter & " AND RitenutaAcconto=" & RitenutaAcconto
    
    If (FLAG_RAGGR_ALTRA_DEST = 1) Then
        rsBuoniVis.Filter = rsBuoniVis.Filter & " AND IDSitoPerAnagraficaIntervento=" & IDSitoPerAnagrafica
    End If
    
    With Me.GrigliaBuoni
        .EnableMove = True
        .UpdatePosition = False
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear

        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDRV_POInterventoRigheDett", "IDRV_POInterventoRigheDett", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POInterventoRighe", "IDRV_POInterventoRighe", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POIntervento", "IDRV_POIntervento", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "NumeroDocumento", "Numero", dgInteger, True, 1000, dgAlignRight
            .ColumnsHeader.Add "DataDocumento", "Data", dgDate, True, 1500, dgAlignleft
            .ColumnsHeader.Add "RitenutaAcconto", "Ritenuta acconto", dgBoolean, True, 1500, dgAligncenter
            Set cl = .ColumnsHeader.Add("TotaleRigaNettoIva", "Imponibile", dgDouble, True, 2000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."


            Set cl = .ColumnsHeader.Add("ImportoDiFatturazione", "Importo Fatt.", dgDouble, True, 2000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
            .ColumnsHeader.Add "NumeroIntervento", "Numero Int.", dgInteger, True, 1000, dgAlignRight
            .ColumnsHeader.Add "DataChiamata", "Data chiamata Int.", dgDate, True, 1500, dgAlignleft

            
            
        Set .Recordset = rsBuoniVis
        .Refresh
    End With
    
    CnDMT.CursorLocation = OLDCursor

Exit Sub
ERR_SettaggioGrigliaRate:
    MsgBox Err.Description, vbCritical, "Settaggio griglia buoni da fatturare"

End Sub

Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)

End Sub
Private Function GET_ANAGRAFICA(IDAnagrafica As Long) As String
On Error GoTo ERR_GET_ANAGRAFICA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_ANAGRAFICA = ""

sSQL = "SELECT IDAnagrafica, Anagrafica, Nome "
sSQL = sSQL & "FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_ANAGRAFICA = fnNotNull(rs!Anagrafica) & " " & fnNotNull(rs!Nome)

End If


rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_ANAGRAFICA:
    
End Function
Private Function GET_SITO_PER_ANAGRAFICA(IDSitoAnagrafica As Long, rstmp As ADODB.Recordset) As String
On Error GoTo ERR_GET_SITO_PER_ANAGRAFICA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_SITO_PER_ANAGRAFICA = ""

rstmp!Indirizzo = ""
rstmp!Cap = ""
rstmp!Comune = ""
rstmp!Provincia = ""

sSQL = "SELECT SitoPerAnagrafica.IDSitoPerAnagrafica, SitoPerAnagrafica.SitoPerAnagrafica, SitoPerAnagrafica.Indirizzo, SitoPerAnagrafica.Cap, SitoPerAnagrafica.Telefono, SitoPerAnagrafica.Fax,  "
sSQL = sSQL & "SitoPerAnagrafica.Referente , Comune.Comune, Provincia.Provincia, SitoPerAnagrafica.IDComune "
sSQL = sSQL & "FROM Provincia RIGHT OUTER JOIN "
sSQL = sSQL & "Comune ON Provincia.IDProvincia = Comune.IDProvincia RIGHT OUTER JOIN "
sSQL = sSQL & "SitoPerAnagrafica ON Comune.IDComune = SitoPerAnagrafica.IDComune "
sSQL = sSQL & "WHERE IDSitoPerAnagrafica=" & IDSitoAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_SITO_PER_ANAGRAFICA = fnNotNull(rs!SitoPerAnagrafica)
    
    rstmp!Indirizzo = fnNotNull(rs!Indirizzo)
    rstmp!Cap = fnNotNull(rs!Cap)
    rstmp!Comune = fnNotNull(rs!Comune)
    rstmp!Provincia = fnNotNull(rs!Provincia)
    
End If


rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_SITO_PER_ANAGRAFICA:
    
End Function

Private Sub GrigliaAnagrafica_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    GRIGLIA_BUONI_DA_FATTURARE Me.GrigliaAnagrafica("IDAnagraficaCliente").Value, Me.GrigliaAnagrafica.AllColumns("RitenutaAcconto").Value, Me.GrigliaAnagrafica.AllColumns("IDSitoPerAnagrafica").Value

End Sub
