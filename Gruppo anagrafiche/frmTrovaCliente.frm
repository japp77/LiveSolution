VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmTrovaCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TROVA CLIENTE"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13260
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
   ScaleHeight     =   8820
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTrovaAnagrafica 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DESELEZIONA TUTTO"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   8280
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SELEZIONA TUTTO"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   8280
      Width           =   2295
   End
   Begin VB.CommandButton cmdRegistra 
      Caption         =   "REGISTRA"
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
      Left            =   10800
      TabIndex        =   1
      Top             =   8280
      Width           =   2415
   End
   Begin DmtGridCtl.DmtGrid GrigliaClienti 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   13573
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
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   8400
      Width           =   5655
   End
End
Attribute VB_Name = "frmTrovaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsClienti As ADODB.Recordset

Private Sub cmdRegistra_Click()
Dim sSQL As String

    rsClienti.Filter = "DaRiportare=1"
    
    If Not rsClienti.EOF Then
        
        While Not rsClienti.EOF
            lblInfo.Caption = fnNotNull(rsClienti!Anagrafica) & " " & fnNotNull(rsClienti!Nome)
            DoEvents
        
            sSQL = "INSERT INTO RV_POGruppoAnagraficheRighe ("
            sSQL = sSQL & "IDRV_POGruppoAnagraficheRighe, IDRV_POGruppoAnagrafiche, IDAnagrafica"
            sSQL = sSQL & ") VALUES ("
            sSQL = sSQL & fnGetNewKey("RV_POGruppoAnagraficheRighe", "IDRV_POGruppoAnagraficheRighe") & ", "
            sSQL = sSQL & LINK_TESTA_GRUPPO & ", "
            sSQL = sSQL & fnNotNullN(rsClienti!IDAnagrafica)
            sSQL = sSQL & ")"
            
            Cn.Execute sSQL
            
        rsClienti.MoveNext
        Wend
    End If
    
    Unload Me
    
End Sub

Private Sub Command1_Click()
    SELEZIONA_TUTTO 1
End Sub

Private Sub Command2_Click()
    SELEZIONA_TUTTO 0
End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    CREATE_TABELLA_TEMPORANEA
    INSERIMENTO_DATI
    GET_GRIGLIA
    
    
End Sub
Private Sub CREATE_TABELLA_TEMPORANEA()
    Set rsClienti = New ADODB.Recordset

    rsClienti.CursorLocation = adUseClient

    rsClienti.Fields.Append "DaRiportare", adSmallInt, , adFldIsNullable
    rsClienti.Fields.Append "IDAnagrafica", adInteger, , adFldIsNullable
    rsClienti.Fields.Append "Anagrafica", adVarChar, 250, adFldIsNullable
    rsClienti.Fields.Append "Nome", adVarChar, 250, adFldIsNullable
    rsClienti.Fields.Append "CodiceFiscale", adVarChar, 250, adFldIsNullable
    rsClienti.Fields.Append "PartitaIVA", adVarChar, 250, adFldIsNullable
    rsClienti.Fields.Append "Indirizzo", adVarChar, 250, adFldIsNullable
    rsClienti.Fields.Append "Comune", adVarChar, 250, adFldIsNullable
    rsClienti.Fields.Append "Provincia", adVarChar, 250, adFldIsNullable
    rsClienti.Fields.Append "Nazione", adVarChar, 250, adFldIsNullable
    rsClienti.Fields.Append "Codice", adVarChar, 250, adFldIsNullable
    
    rsClienti.Open , , adOpenKeyset, adLockBatchOptimistic
End Sub
Private Sub INSERIMENTO_DATI()

Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT * FROM IERepCliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & "ORDER BY Anagrafica, Nome"


Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    If GET_ESISTENZA_ANAGRAFICA(rs!IDAnagrafica) = 0 Then
        rsClienti.AddNew
        
            rsClienti!DaRiportare = 0
            rsClienti!IDAnagrafica = fnNotNullN(rs!IDAnagrafica)
            rsClienti!Anagrafica = fnNotNull(rs!Anagrafica)
            rsClienti!Nome = fnNotNull(rs!Nome)
            rsClienti!CodiceFiscale = fnNotNull(rs!CodiceFiscale)
            rsClienti!PartitaIva = fnNotNull(rs!PartitaIva)
            rsClienti!Indirizzo = fnNotNull(rs!Indirizzo)
            rsClienti!Comune = fnNotNull(rs!Comune)
            rsClienti!Provincia = fnNotNull(rs!NomeProvincia)
            rsClienti!Nazione = fnNotNull(rs!Nazione)
            rsClienti!codice = fnNotNull(rs!codice)
            
            
        rsClienti.Update
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

    
    
End Sub

Private Sub GET_GRIGLIA()
On Error GoTo ERR_SettaggioGrigliaRate
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
       
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3

        
    With Me.GrigliaClienti
        .EnableMove = True
        .UpdatePosition = False
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear

        .ColumnsHeader.Clear
            Set cl = .ColumnsHeader.Add("DaRiportare", "Riporta", dgBoolean, True, 1000, dgAligncenter)
                cl.Editable = True
            
            .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Codice", "Codice", dgchar, True, 1200, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Anagrafica", "Ragione sociale", dgchar, True, 2500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Nome", "Nome", dgchar, True, 1200, dgAlignleft, True, True, False
            .ColumnsHeader.Add "PartitaIva", "Partita I.V.A.", dgchar, True, 1800, dgAlignleft, True, True, False
            .ColumnsHeader.Add "CodiceFiscale", "Codice fiscale", dgchar, True, 1800, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Indirizzo", "Indirizzo", dgchar, False, 1800, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Comune", "Comune", dgchar, False, 1800, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Provincia", "Provincia", dgchar, False, 1800, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Nazione", "Nazione", dgchar, False, 1800, dgAlignleft, True, True, False
                
        Set .Recordset = rsClienti
        .Refresh
    End With
    
    Cn.CursorLocation = OLDCursor

Exit Sub
ERR_SettaggioGrigliaRate:
    MsgBox Err.Description, vbCritical, "Settaggio griglia"

End Sub
Private Sub GrigliaClienti_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GrigliaClienti.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsClienti.Fields("DaRiportare").Value), 2
        End If
    End If
End Sub

Private Sub GrigliaClienti_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If GrigliaClienti.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < GrigliaClienti.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If GrigliaClienti.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsClienti.Fields("DaRiportare").Value), 2
            End If
        End If
    End If
    
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    
        If Not rsClienti.EOF And Not rsClienti.BOF Then
            rsClienti.Fields("DaRiportare").Value = Abs(CLng(Selected))

            Me.GrigliaClienti.Refresh
        End If
        DoEvents

End Sub
Private Sub SELEZIONA_TUTTO(Valore As Long)

If Not rsClienti.EOF And Not rsClienti.BOF Then
    rsClienti.MoveFirst
    
    While Not rsClienti.EOF
        rsClienti!DaRiportare = Valore
    rsClienti.MoveNext
    Wend
    
    
End If

GET_GRIGLIA
End Sub
Private Function GET_ESISTENZA_ANAGRAFICA(IDAnagrafica As Long) As Long
On Error GoTo ERR_GET_ESISTENZA_ANAGRAFICA
rsRighe.Filter = "IDAnagrafica=" & IDAnagrafica

If rsRighe.EOF Then
    GET_ESISTENZA_ANAGRAFICA = 0
Else
    GET_ESISTENZA_ANAGRAFICA = 1
End If

rsRighe.Filter = vbNullString

Exit Function
ERR_GET_ESISTENZA_ANAGRAFICA:
    GET_ESISTENZA_ANAGRAFICA = 0
End Function

Private Sub txtTrovaAnagrafica_Change()
On Error GoTo ERR_txtTrovaAnagrafica_Change
    If Len(Me.txtTrovaAnagrafica.Text) > 0 Then
        rsClienti.Filter = "Anagrafica LIKE " & fnNormString(Me.txtTrovaAnagrafica.Text & "%")
    Else
        rsClienti.Filter = vbNullString
    End If
    GET_GRIGLIA
Exit Sub
ERR_txtTrovaAnagrafica_Change:
    MsgBox Err.Description, vbCritical, "txtTrovaAnagrafica_Change"
End Sub
