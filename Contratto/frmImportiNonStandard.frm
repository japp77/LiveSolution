VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmImportiNonStandard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importi non standard per gli anni successivi"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10140
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
   ScaleHeight     =   5400
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DMTEDITNUMLib.dmtNumber txtSconto 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      _StockProps     =   253
      Text            =   "0"
      BackColor       =   16777215
      Appearance      =   1
      DecFinalZeros   =   -1  'True
      AllowEmpty      =   0   'False
   End
   Begin VB.TextBox txtAnnotazioni 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1080
      Width           =   8535
   End
   Begin VB.CheckBox chkAdeguamentoIstat 
      Caption         =   "Adeguamento I.S.T.A.T."
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Top             =   480
      Width           =   3015
   End
   Begin DMTDATETIMELib.dmtDate txtData 
      Height          =   285
      Left            =   8640
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      _StockProps     =   253
      BackColor       =   16777215
      Appearance      =   1
   End
   Begin DMTEDITNUMLib.dmtCurrency txtImporto 
      Height          =   285
      Left            =   4800
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      _StockProps     =   253
      Text            =   " 0"
      BackColor       =   16777215
      Appearance      =   1
      UseSeparator    =   -1  'True
      CurrencySymbol  =   ""
      AllowEmpty      =   0   'False
      DecFinalZeros   =   -1  'True
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   375
      Left            =   8760
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalva 
      Caption         =   "Salva"
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuovo 
      Caption         =   "Nuovo"
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin DmtGridCtl.DmtGrid GrigliaImportiNonStandard 
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4048
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
   Begin DMTEDITNUMLib.dmtCurrency txtImportoPrec 
      Height          =   285
      Left            =   6240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      _StockProps     =   253
      Text            =   " 0"
      BackColor       =   16777215
      Appearance      =   1
      UseSeparator    =   -1  'True
      CurrencySymbol  =   ""
      AllowEmpty      =   0   'False
      DecFinalZeros   =   -1  'True
   End
   Begin DMTEDITNUMLib.dmtNumber txtNumeroRinnovo 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   503
      _StockProps     =   253
      Text            =   "0"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Appearance      =   1
      AllowEmpty      =   0   'False
   End
   Begin DMTEDITNUMLib.dmtCurrency txtScontoAImporto 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   480
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      _StockProps     =   253
      Text            =   " 0"
      BackColor       =   16777215
      Appearance      =   1
      UseSeparator    =   -1  'True
      CurrencySymbol  =   ""
      AllowEmpty      =   0   'False
      DecFinalZeros   =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Differenza"
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   18
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "N° Rinnovo"
      Height          =   255
      Index           =   31
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Importo prec."
      Height          =   255
      Index           =   4
      Left            =   6240
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "% Sc."
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   15
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Annotazioni"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Importo attuale"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Data rinnovo"
      Height          =   255
      Index           =   0
      Left            =   8640
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmImportiNonStandard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsImporti As DmtOleDbLib.adoResultset
Private Nuovo As Integer
Private Link_ImportiNonStd As Long
Private NuovaData As String
Private b_loading As Boolean



Private Sub cmdElimina_Click()
On Error GoTo ERR_cmdElimina_Click
Dim Testo As String
Dim sSQL As String

If Nuovo = 0 Then Exit Sub
    Testo = "Sei sicuro di voler eliminare la riga?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione") = vbNo Then Exit Sub
    
    If fnNotNullN(Me.GrigliaImportiNonStandard.AllColumns("NumeroRinnovo").Value) > frmMain.txtNumeroRinnovo.Value Then
        sSQL = "DELETE FROM RV_POContrattoNonStandard "
        sSQL = sSQL & "WHERE IDRV_POContrattoNonStandard=" & Link_ImportiNonStd
        Cn.Execute sSQL
        SettaggioGriglia
    Else
        MsgBox "Impossibile eliminare, poichè l'importo è stato già utilizzato per rinnovi precedenti.", vbInformation, "Impossibile eliminare"
    End If
Exit Sub
ERR_cmdElimina_Click:
    MsgBox Err.Description, vbCritical, "cmdElimina_Click"
    
End Sub

Private Sub cmdNuovo_Click()

Me.txtData.Value = 0
Me.txtNumeroRinnovo.Value = GET_NUOVO_NUMERO_RINNOVO(frmMain.txtIDContrattoPadre.Value)
Me.txtImportoPrec.Value = GET_ULTIMO_CONTRATTO_NON_STANDARD(frmMain.txtIDContrattoPadre.Value)
Me.txtImporto.Value = 0
Me.txtScontoAImporto.Value = 0

Me.txtSconto.Value = 0
Me.chkAdeguamentoIstat.Value = frmMain.chkAdeguamentoIstat.Value

Nuovo = 0

Me.txtData.Enabled = True
Me.txtImportoPrec.Enabled = True
Me.txtImporto.Enabled = True
Me.chkAdeguamentoIstat.Enabled = True
Me.txtAnnotazioni.Enabled = True

Me.txtSconto.Enabled = True
Me.cmdSalva.Enabled = True
Me.cmdElimina.Enabled = True

Me.txtAnnotazioni.Locked = False

'If b_loading = True Then
    'Me.txtData.SetFocus
'End If
End Sub

Private Sub cmdSalva_Click()
On Error GoTo ERR_cmdSalva_Click
Dim sSQL As String
        
    If Me.txtImporto.Value = 0 Then
        MsgBox "L'importo non può essere uguale a zero", vbCritical, TheApp.FunctionName
        Me.txtImporto.SetFocus
        Exit Sub
    End If
    
    
    If Nuovo = 0 Then
        sSQL = "INSERT INTO RV_POContrattoNonStandard ("
        sSQL = sSQL & "IDRV_POContrattoNonStandard, IDRV_POContratto, "
        sSQL = sSQL & "Importo, PercentualeSconto,ImportoPrecedente, "
        sSQL = sSQL & "AdeguamentoIstat, Annotazioni, Disponibile, NumeroRinnovo, IDRV_POContrattoPadre, ScontoAImporto) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnGetNewKey("RV_POContrattoNonStandard", "IDRV_POContrattoNonStandard") & ", "
        sSQL = sSQL & Link_Contratto & ", "
        sSQL = sSQL & fnNormNumber(Me.txtImporto.Value) & ", "
        sSQL = sSQL & fnNormNumber(Me.txtSconto.Value) & ", "
        sSQL = sSQL & fnNormNumber(Me.txtImportoPrec.Value) & ", "
        sSQL = sSQL & Me.chkAdeguamentoIstat.Value & ", "
        sSQL = sSQL & fnNormString(Me.txtAnnotazioni.Text) & ", "
        sSQL = sSQL & "0" & ", "
        sSQL = sSQL & Me.txtNumeroRinnovo.Value & ", "
        sSQL = sSQL & frmMain.txtIDContrattoPadre.Value & ", "
        sSQL = sSQL & fnNormNumber(Me.txtScontoAImporto.Value) & ")"
    Else
        sSQL = "UPDATE RV_POContrattoNonStandard SET "
        sSQL = sSQL & "Importo=" & fnNormNumber(Me.txtImporto.Value) & ", "
        sSQL = sSQL & "ImportoPrecedente=" & fnNormNumber(Me.txtImportoPrec.Value) & ", "
        sSQL = sSQL & "PercentualeSconto=" & fnNormNumber(Me.txtSconto.Value) & ", "
        sSQL = sSQL & "AdeguamentoIstat=" & Me.chkAdeguamentoIstat.Value & ", "
        sSQL = sSQL & "Annotazioni=" & fnNormString(Me.txtAnnotazioni.Text) & ", "
        sSQL = sSQL & "ScontoAImporto=" & fnNormNumber(Me.txtScontoAImporto.Value) & " "
        sSQL = sSQL & "WHERE IDRV_POContrattoNonStandard=" & Link_ImportiNonStd
        
    End If
    
    Cn.Execute sSQL
    
    SettaggioGriglia
Exit Sub
ERR_cmdSalva_Click:
    MsgBox Err.Description, vbCritical, "cmdSalva_Click"
End Sub

Private Sub Form_Activate()
    
    cmdNuovo_Click
    
    b_loading = True

End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    SettaggioGriglia
    
    
    b_loading = False
End Sub


Private Sub SettaggioGriglia()
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    
sSQL = "Select * FROM RV_POContrattoNonStandard "
sSQL = sSQL & "WHERE IDRV_POContrattoPadre=" & frmMain.txtIDContrattoPadre.Value
sSQL = sSQL & " ORDER BY NumeroRinnovo DESC"


OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3

    Set rsImporti = Cn.OpenResultset(sSQL)
        Set rsEvent = rsImporti.Data
    
    With Me.GrigliaImportiNonStandard
        .ColumnsHeader.Clear
        
            .ColumnsHeader.Add "IDRV_POContrattoNonStandard", "IDRV_POContrattoNonStandard", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POContratto", "IDRV_POContratto", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "NumeroRinnovo", "Numero rinnovo", dgInteger, True, 1500, dgAlignRight
            
            Set cl = .ColumnsHeader.Add("PercentualeSconto", "% Sc.", dgDouble, True, 1000, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("ScontoAImporto", "Sconto a importo", dgDouble, True, 1000, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Importo", "Importo", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("ImportoPrecedente", "Importo Prec.", dgDouble, True, 1300, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
        Set .Recordset = rsImporti.Data
        .Refresh
        .LoadUserSettings
        
    End With

If rsImporti.EOF Then
    cmdNuovo_Click
End If
Cn.CursorLocation = OLDCursor
End Sub



Private Sub Form_Unload(Cancel As Integer)
    b_loading = False
End Sub

Private Sub GrigliaImportiNonStandard_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
Dim Abilita As Boolean

    Me.txtNumeroRinnovo.Value = fnNotNullN(Me.GrigliaImportiNonStandard.AllColumns("NumeroRinnovo").Value)
    Me.txtData.Value = fnNotNullN(Me.GrigliaImportiNonStandard.AllColumns("Data").Value)
    Me.txtImportoPrec.Value = fnNotNullN(Me.GrigliaImportiNonStandard.AllColumns("ImportoPrecedente").Value)
    Me.txtImporto.Value = fnNotNullN(Me.GrigliaImportiNonStandard.AllColumns("Importo").Value)
    Me.chkAdeguamentoIstat.Value = Abs(fnNotNullN(Me.GrigliaImportiNonStandard.AllColumns("AdeguamentoIstat").Value))
    Me.txtAnnotazioni.Text = fnNotNull(Me.GrigliaImportiNonStandard.AllColumns("Annotazioni"))
    Me.txtSconto.Value = fnNotNullN(Me.GrigliaImportiNonStandard.AllColumns("PercentualeSconto").Value)
    Me.txtScontoAImporto.Value = fnNotNullN(Me.GrigliaImportiNonStandard.AllColumns("ScontoAImporto").Value)
    Link_ImportiNonStd = Me.GrigliaImportiNonStandard.AllColumns("IDRV_POContrattoNonStandard")
    
    Nuovo = 1
    
    If Me.txtNumeroRinnovo.Value <= frmMain.txtNumeroRinnovo.Value Then
        Abilita = False
    Else
        Abilita = True
    End If
    
    Me.txtData.Enabled = Abilita
    Me.txtImportoPrec.Enabled = Abilita
    Me.txtImporto.Enabled = Abilita
    Me.chkAdeguamentoIstat.Enabled = Abilita
    Me.txtSconto.Enabled = Abilita
    Me.cmdSalva.Enabled = Abilita
    Me.cmdElimina.Enabled = Abilita
    
    If Abilita = False Then
        Me.txtAnnotazioni.Locked = True
    Else
        Me.txtAnnotazioni.Locked = False
    End If

End Sub
    
Private Function GET_ULTIMO_CONTRATTO_NON_STANDARD(IDContrattoPadre As Long) As Double
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT Importo, NumeroRinnovo "
sSQL = sSQL & " FROM RV_POContrattoNonStandard"
sSQL = sSQL & " WHERE IDRV_POContrattoPadre=" & IDContrattoPadre
sSQL = sSQL & " ORDER BY NumeroRinnovo DESC"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

If rs.EOF Then
    GET_ULTIMO_CONTRATTO_NON_STANDARD = frmMain.txtImportoAttuale.Value
Else
    If fnNotNullN(rs!NumeroRinnovo) < frmMain.txtNumeroRinnovo.Value Then
        GET_ULTIMO_CONTRATTO_NON_STANDARD = frmMain.txtImportoAttuale.Value
    Else
        GET_ULTIMO_CONTRATTO_NON_STANDARD = fnNotNullN(rs!Importo)
    End If
End If

rs.Close
Set rs = Nothing
End Function

Private Sub txtImporto_LostFocus()

'    If Me.txtSconto.Value > 0 Then
'        Me.txtSconto.Value = 100 - ((Me.txtImporto.Value / Me.txtImportoPrec.Value) * 100)
'    End If

    Me.txtScontoAImporto.Value = Me.txtImporto.Value - Me.txtImportoPrec.Value
    
End Sub

Private Sub txtImportoPrec_LostFocus()
    Me.txtImporto.Value = Me.txtImportoPrec.Value + Me.txtScontoAImporto.Value
    txtSconto_LostFocus
End Sub

Private Sub txtSconto_LostFocus()
    Me.txtImporto.Value = Me.txtImportoPrec.Value + Me.txtScontoAImporto.Value - ((Me.txtImportoPrec.Value / 100) * Me.txtSconto.Value)
End Sub
Private Function GET_NUOVO_NUMERO_RINNOVO(IDContrattoPadre As Long)
On Error GoTo ERR_GET_NUOVO_NUMERO_RINNOVO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(NumeroRinnovo) AS NumeroRinnovo "
sSQL = sSQL & "FROM RV_POContrattoNonStandard "
sSQL = sSQL & "WHERE IDRV_POContrattoPadre=" & IDContrattoPadre


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUOVO_NUMERO_RINNOVO = 0
Else
    GET_NUOVO_NUMERO_RINNOVO = fnNotNullN(rs!NumeroRinnovo)
End If

rs.CloseResultset
Set rs = Nothing

If GET_NUOVO_NUMERO_RINNOVO < frmMain.txtNumeroRinnovo.Value Then
    GET_NUOVO_NUMERO_RINNOVO = frmMain.txtNumeroRinnovo.Value + 1
Else
    GET_NUOVO_NUMERO_RINNOVO = GET_NUOVO_NUMERO_RINNOVO + 1
End If
Exit Function
ERR_GET_NUOVO_NUMERO_RINNOVO:
    MsgBox Err.Description, vbCritical, "GET_NUOVO_NUMERO_RINNOVO"
End Function



Private Sub txtScontoAImporto_LostFocus()
    txtSconto_LostFocus
End Sub
