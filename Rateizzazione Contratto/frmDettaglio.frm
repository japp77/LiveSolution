VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmDettaglio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERSONALIZZA RATEIZZAZIONE"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDettaglio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEliminaTutto 
      Caption         =   "ELIMINA TUTTO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   3720
      Width           =   2535
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   3615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   6376
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
Attribute VB_Name = "frmDettaglio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub CREA_RECORDSET()
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim I As Long
Dim J As Long

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

rsGriglia.Fields.Append "NumeroRata", adInteger, , adFldIsNullable
rsGriglia.Fields.Append "PercentualeRata", adDouble, , adFldIsNullable
rsGriglia.Fields.Append "PagamentoAnticipato", adSmallInt, , adFldIsNullable
rsGriglia.Fields.Append "N", adInteger, , adFldIsNullable

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic

sSQL = "SELECT * FROM RV_PORateizzazioneRighe "
sSQL = sSQL & "WHERE IDRV_PORateizzazione=" & LINK_RATEIZZAZIONE

Set rs = Cn.OpenResultset(sSQL)

I = 1

While Not rs.EOF
    rsGriglia.AddNew
        rsGriglia!N = I
        rsGriglia!NumeroRata = fnNotNullN(rs!NumeroRata)
        rsGriglia!PercentualeRata = fnNotNullN(rs!PercentualeRata)
        rsGriglia!PagamentoAnticipato = fnNotNullN(rs!PagamentoAnticipato)
    rsGriglia.Update
    I = I + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

For J = I To 20
    rsGriglia.AddNew
        rsGriglia!N = I
    rsGriglia.Update
Next

GET_GRIGLIA

End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_fnGrigliaTipoPagamento
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3

    With Me.Griglia
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "N", "N", dgInteger, False, 500, dgAlignRight, True, True, False
            Set cl = .ColumnsHeader.Add("NumeroRata", "N° rata", dgNumeric, True, 1500, dgAlignRight)
                cl.Editable = True
            Set cl = .ColumnsHeader.Add("PercentualeRata", "% Rata", dgDouble, True, 2000, dgAlignRight)
                cl.Editable = True
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("PagamentoAnticipato", "Pag. anticipato", dgBoolean, True, 2000, dgAligncenter, True, True, False)
                cl.Editable = True
        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With
    
    Cn.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaTipoPagamento:
    MsgBox Err.Description, vbCritical, "Adeguamenti storici del contratto"
End Sub

Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdEliminaTutto_Click
Dim sSQL As String
Dim rsNew As ADODB.Recordset
    
    rsGriglia.Update
    
    Griglia.UpdatePosition = False
    
    rsGriglia.MoveFirst
        
    sSQL = "DELETE FROM RV_PORateizzazioneRighe "
    sSQL = sSQL & " WHERE IDRV_PORateizzazione=" & LINK_RATEIZZAZIONE
    Cn.Execute sSQL
    
    sSQL = "SELECT * FROM RV_PORateizzazioneRighe "
    sSQL = sSQL & " WHERE IDRV_PORateizzazione=" & LINK_RATEIZZAZIONE
    
    Set rsNew = New ADODB.Recordset
    rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    While Not rsGriglia.EOF
        If (fnNotNullN(rsGriglia!NumeroRata) > 0) Then
            rsNew.AddNew
                rsNew!IDRV_PORateizzazione = LINK_RATEIZZAZIONE
                rsNew!NumeroRata = rsGriglia!NumeroRata
                rsNew!PercentualeRata = rsGriglia!PercentualeRata
                rsNew!PagamentoAnticipato = rsGriglia!PagamentoAnticipato
            rsNew.Update
        End If
    rsGriglia.MoveNext
    Wend
    
    rsNew.Close
    Set rsNew = Nothing
    
    rsGriglia.Close
    Set rsGriglia = Nothing
    
    Griglia.UpdatePosition = True
    Unload Me
    
Exit Sub
ERR_cmdEliminaTutto_Click:
    Griglia.UpdatePosition = True
    MsgBox Err.Description, vbCritical, "cmdEliminaTutto_Click"

End Sub

Private Sub cmdEliminaTutto_Click()
On Error GoTo ERR_cmdEliminaTutto_Click
    
    Griglia.UpdatePosition = False
    
    rsGriglia.MoveFirst
    
    While Not rsGriglia.EOF
        rsGriglia!NumeroRata = Null
        rsGriglia!PercentualeRata = Null
        rsGriglia!PagamentoAnticipato = Null
    rsGriglia.MoveNext
    Wend
    
    rsGriglia.MoveFirst
    
    Me.Griglia.Refresh
    
    Griglia.UpdatePosition = True
    
Exit Sub
ERR_cmdEliminaTutto_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaTutto_Click"
End Sub

Private Sub Form_Load()
    CREA_RECORDSET
End Sub

Private Sub Griglia_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_Griglia_KeyPress
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.Griglia.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(fnNotNullN(rsGriglia.Fields("PagamentoAnticipato").Value))
        End If
    End If

Exit Sub
ERR_Griglia_KeyPress:
    MsgBox Err.Description, vbCritical, "Griglia_KeyPress"
End Sub

Private Sub Griglia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"

'    'Controlla se l'utente ha cliccato su una riga valida
'    If Griglia.HitTest(X, Y) > 0 Then
'        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
'        If X > 0 And (X * Screen.TwipsPerPixelX) < Griglia.ColumnsHeader("PagamentoAnticipato").Width Then
'            'Se non siamo in modalità filtri
'            If Griglia.GuiMode = dgNormal Then
'                'Abilitiamo o disabilitiamo il check in base allo stato corrente
'                sbSelectSelectedRow Not CBool(fnNotNullN(rsGriglia.Fields("PagamentoAnticipato").Value))
'            End If
'        End If
'    End If

End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean)
On Error GoTo ERR_sbSelectSelectedRow
    If Not rsGriglia.EOF And Not rsGriglia.BOF Then
    
        rsGriglia.Fields("PagamentoAnticipato").Value = Abs(CLng(Selected))
        
        Me.Griglia.Refresh
        
    End If
Exit Sub
ERR_sbSelectSelectedRow:
    MsgBox Err.Description, vbCritical, "sbSelectSelectedRow"
End Sub
