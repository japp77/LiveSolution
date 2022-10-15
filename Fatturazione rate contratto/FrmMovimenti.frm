VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Begin VB.Form FrmMovimenti 
   Caption         =   "Creazione documenti (Passo 2 di 4)"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19245
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
   ScaleHeight     =   8880
   ScaleWidth      =   19245
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   840
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   0
      ScaleHeight     =   8745
      ScaleWidth      =   19065
      TabIndex        =   7
      Top             =   0
      Width           =   19095
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   8535
         Left            =   120
         ScaleHeight     =   8505
         ScaleWidth      =   18825
         TabIndex        =   8
         Top             =   120
         Width           =   18855
         Begin VB.CommandButton cmdAvanti 
            Caption         =   "Avanti"
            Height          =   375
            Left            =   17280
            TabIndex        =   3
            Top             =   8040
            Width           =   1335
         End
         Begin VB.CommandButton cmdAnnulla 
            Caption         =   "Annulla"
            Height          =   375
            Left            =   15840
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   8040
            Width           =   1335
         End
         Begin VB.CommandButton cmdIndietro 
            Caption         =   "Indietro"
            Height          =   375
            Left            =   14400
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   8040
            Width           =   1335
         End
         Begin VB.CommandButton cmdVisStampa 
            Caption         =   "Visualizza stampe"
            Height          =   375
            Left            =   4200
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   8040
            Width           =   1935
         End
         Begin VB.CommandButton cmdSelezionaTutto 
            Caption         =   "Seleziona tutto"
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Top             =   8040
            Width           =   1935
         End
         Begin VB.CommandButton cmdDeselezionaTutto 
            Caption         =   "Deseleziona tutto"
            Height          =   375
            Left            =   2160
            TabIndex        =   2
            Top             =   8040
            Width           =   1935
         End
         Begin DmtGridCtl.DmtGrid GrigliaRate 
            Height          =   7335
            Left            =   120
            TabIndex        =   0
            Top             =   120
            Width           =   18615
            _ExtentX        =   32835
            _ExtentY        =   12938
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
            Left            =   240
            TabIndex        =   9
            Top             =   5280
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblTotaleDocumento 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   16560
            TabIndex        =   11
            Top             =   7560
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Totale da fatturare: "
            Height          =   255
            Left            =   14640
            TabIndex        =   10
            Top             =   7560
            Width           =   1815
         End
      End
      Begin VB.Frame FraStampa 
         Caption         =   "Stampa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   5895
         Begin VB.CommandButton cmdChiudiVisStampa 
            Height          =   325
            Left            =   4920
            Picture         =   "FrmMovimenti.frx":4781A
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "STAMPA"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdStampa 
            Height          =   325
            Left            =   5400
            Picture         =   "FrmMovimenti.frx":47BA4
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "STAMPA"
            Top             =   0
            Width           =   375
         End
         Begin DmtActBox.DmtActBoxCtl ActivityBox 
            Height          =   6975
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   12303
            BackColor       =   -2147483643
            ForeColor       =   -2147483630
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
      End
   End
End
Attribute VB_Name = "FrmMovimenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private rsRate As ADODB.Recordset

Private oReport As dmtReportLib.dmtReport
Private IDTipoOggettoPrg As Long
Private Riga_selezionata As Long
Private Aggiornamento As Long
Private VisualizzaStampe As Long



'----- Oggetti e variabili per la gestione del riquadro attività -----------
'***Reports                                                                -
Private WithEvents oReportsActivity As DmtActBoxLib.ReportsActivity       '-
Attribute oReportsActivity.VB_VarHelpID = -1
'***Filtri                                                                 -
Private WithEvents oFiltersActivity As DmtActBoxLib.FiltersActivity       '-
Attribute oFiltersActivity.VB_VarHelpID = -1
'***Viste tabellari                                                        -
Private WithEvents oTableViewsActivity As DmtActBoxLib.TableViewsActivity '-
Attribute oTableViewsActivity.VB_VarHelpID = -1
'***Esportazioni                                                           -
Private oExportActivity As DmtActBoxLib.ExportActivity                    '-
'***Supporto tecnico                                                       -
Private oSupportActivity As DmtActBoxLib.SupportActivity                  '-
'***Nome dell'attività predefinita del riquadro attività                   -
Private m_DefaultActivity As String                                       '-
'---------------------------------------------------------------------------

Private Sub AggiornaGriglia(Valore As Integer)
    Dim sSQL As String
    
    sSQL = "UPDATE RV_POTMPFatturazioneRate SET "
    sSQL = sSQL & "NonFatturare=" & fnNormBoolean(Valore) & " "
    sSQL = sSQL & "WHERE IDTMP=" & Me.GrigliaRate.AllColumns("IDTMP")
    
    CnDMT.Execute sSQL
    Aggiornamento = 1
        SettaggioGrigliaRate
    Aggiornamento = 0
    
    Me.GrigliaRate.SetFocus
End Sub

Private Sub cmdAnnulla_Click()
Dim Risposta As Integer
    Risposta = MsgBox("Vuoi abbandonare il wizard per il passaggio delle rate contratto in fatturazione?", vbInformation + vbYesNo, "Abbandono")
    If Risposta = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdAvanti_Click()
    Dim I As Integer
    Dim IDRata
    
    If ((rsRate.EOF) And (rsRate.BOF)) Then Exit Sub
    
    Me.GrigliaRate.UpdatePosition = False
    
    rsRate.MoveFirst
    I = 0
    While Not rsRate.EOF
        If rsRate!Fatturare = 1 Then
            rsRateDaFatt.AddNew

                For I = 0 To rsRate.Fields.Count - 1
'                    If (rsRate(I).Name = "EntePubblico") Then
'                        MsgBox "STOP"
'                    End If
                    Select Case rsRate.Fields(I).Type
                        Case adInteger, adDouble
                            rsRateDaFatt(rsRate(I).Name).Value = fnNotNullN(rsRate(I).Value)
                        Case adSmallInt, adTinyInt, adBoolean, 17
                            rsRateDaFatt(rsRate(I).Name).Value = fnNotNullN(rsRate(I).Value)
                        Case Else
                            rsRateDaFatt(rsRate(I).Name).Value = rsRate(I).Value
                    End Select
                    
                    
                Next
            rsRateDaFatt.Update
        End If
    rsRate.MoveNext
    Wend
    rsRate.Close
    Set rsRate = Nothing
    
    'RaggruppamentoAnagrafica = Me.chkRaggruppaFatture.Value
    
    Unload Me
    
End Sub

Private Sub cmdChiudiVisStampa_Click()
    cmdVisStampa_Click
End Sub

Private Sub cmdDeselezionaTutto_Click()
If Not (rsRate.BOF And rsRate.EOF) Then
    
    Me.GrigliaRate.UpdatePosition = False
    
    rsRate.MoveFirst
    While Not rsRate.EOF
'        If rsRate!Fatturare = 0 Then
'            sbSelectSelectedRow Not CBool(rsRate.Fields("Fatturare").Value), 2
'        End If

        rsRate!Fatturare = 0
    rsRate.MoveNext
    Wend
    Me.GrigliaRate.Refresh
    'Me.GrigliaRate.UpdatePosition = True
    
        
    TotaleDocumento = TotaleRateDaFatturare
    Me.lblTotaleDocumento.Caption = FormatCurrency(TotaleDocumento, 2)
End If
End Sub

Private Sub cmdIndietro_Click()
    Unload Me
End Sub


Private Sub cmdSelezionaTutto_Click()
If Not (rsRate.BOF And rsRate.EOF) Then
    
    Me.GrigliaRate.UpdatePosition = False
    
    rsRate.MoveFirst
    While Not rsRate.EOF
'        If rsRate!Fatturare = 0 Then
'            sbSelectSelectedRow Not CBool(rsRate.Fields("Fatturare").Value), 2
'        End If

        rsRate!Fatturare = 1
    rsRate.MoveNext
    Wend
    Me.GrigliaRate.Refresh
    Me.GrigliaRate.UpdatePosition = True
    
        
    TotaleDocumento = TotaleRateDaFatturare
    Me.lblTotaleDocumento.Caption = FormatCurrency(TotaleDocumento, 2)
End If

End Sub
Private Sub cmdStampa_Click()
Dim sSQL As String

AGGIORNA_TABELLA_PER_REPORT
        
IDTipoOggettoPrg = fnGetTipoOggetto("RV_POCreazioneDocumenti")

Set oReport = New dmtReportLib.dmtReport
Set oReport.Connection = CnDMT
oReport.Password = TheApp.Password
oReport.User = TheApp.User

oReport.BranchID = TheApp.Branch 'IDFiliale

oReport.DocTypeID = IDTipoOggettoPrg

If Len(oReportsActivity.SelectedReportName) > 0 Then
    IDReport = fncTrovaReport(oReportsActivity.SelectedReportName, IDTipoOggettoPrg)
Else
    IDReport = fncTrovaReport(oReportsActivity.DefaultReportName, IDTipoOggettoPrg)
End If

If IDReport > 0 Then
    fncImpostaDefaultReport IDReport, IDTipoOggettoPrg
    
    
    oReport.Preview 0, 0, 0
    
Else
    MsgBox "ATTENZIONE!!!!" & vbCrLf & "Il report non è stato trovato!", vbCritical, "Impossibile stampare"
End If

           
End Sub
Private Sub cmdVisStampa_Click()

    If Me.FraStampa.Visible = False Then
        Me.Picture2.Left = Me.FraStampa.Left + Me.FraStampa.Width + 10
        Me.FraStampa.Visible = True
        Me.ZOrder 0
        Me.Pic1.Width = Me.Picture2.Width + Me.FraStampa.Width + 240
        VisualizzaStampe = 1
    Else
        Me.Picture2.Left = Me.FraStampa.Left
        Me.FraStampa.Visible = False
        Me.ZOrder 1
        Me.Pic1.Width = 15975
        VisualizzaStampe = 0
    End If
    
    Form_Resize
End Sub

Private Sub Form_Load()
    Dim TotaleDocumento As Double
    
    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    If Not (Cls_Nom Is Nothing) Then
        Set Cls_Nom = Nothing
    End If
    
    Set Cls_Nom = New Collection
    ElaborazioneRateTemporanee
    SettaggioGrigliaRate
    TotaleDocumento = TotaleRateDaFatturare
    Me.lblTotaleDocumento.Caption = FormatCurrency(TotaleDocumento, 2)
    ConfigurazioneStampe
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If Me.cmdAvanti.Value = True Then
        FrmParametri.Show
        Exit Sub
    End If
    If cmdIndietro.Value = True Then
        FrmInizio.Show
        Exit Sub
    End If
    
End Sub

Private Sub SettaggioGrigliaRate()
On Error GoTo ERR_SettaggioGrigliaRate
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
        'Set rsRate = New ADODB.Recordset
        'rsRate.CursorLocation = adUseClient
        'rsRate.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockBatchOptimistic
        
        With Me.GrigliaRate
            .EnableMove = True
            .UpdatePosition = False
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectCell
            .ColumnsHeader.Clear

            .ColumnsHeader.Clear
            .ColumnsHeader.Add "Fatturare", "Registra", dgBoolean, True, 1500, dgAligncenter
            
            .ColumnsHeader.Add "IDRV_PORateContratto", "IDRV_PORateContratto", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POContratto", "IDRV_POContratto", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POContrattoPadre", "IDRV_POContrattoPadre", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POContrattoAdeguamento", "IDRV_POContrattoAdeguamento", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POContatoreRilevamenti", "IDRV_POContatoreRilevamenti", dgInteger, False, 500, dgAlignRight
            
            .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "Anagrafica", "Cliente", dgchar, True, 3000, dgAlignleft
            .ColumnsHeader.Add "IDTipo", "IDTipo", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "DescrizioneTipo", "Descrizione add.", dgchar, True, 2500, dgAlignleft

            .ColumnsHeader.Add "ContrattoAttuale", "Contratto attuale", dgBoolean, False, 1500, dgAligncenter
            .ColumnsHeader.Add "AnnoContratto", "Anno contratto", dgInteger, True, 1500, dgAlignRight
            .ColumnsHeader.Add "NumeroContratto", "N° contratto", dgInteger, True, 1500, dgAlignRight
            
            .ColumnsHeader.Add "IDAnagraficaFatturazione", "IDAnagraficaFatturazione", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "AnagraficaFatturazione", "Anagrafica Fatt.", dgchar, False, 3000, dgAlignleft
            .ColumnsHeader.Add "NomeFatturazione", "Nome ana. Fatt.", dgchar, False, 1500, dgAlignleft
            .ColumnsHeader.Add "EntePubblico", "Ente Pubblico", dgBoolean, False, 1500, dgAligncenter
            
            .ColumnsHeader.Add "IDSitoPerAnagrafica", "IDSitoPerAnagrafica", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "SitoPerAnagrafica", "Altra sede", dgchar, False, 3000, dgAlignleft
        
            .ColumnsHeader.Add "IDTipoContratto", "IDTipoContratto", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "TipoContratto", "Tipo contratto", dgchar, True, 2000, dgAlignleft
            
            .ColumnsHeader.Add "DataRata", "Data Rata", dgDate, True, 2000, dgAlignleft
            .ColumnsHeader.Add "NumeroRata", "N° rata", dgInteger, True, 1000, dgAlignRight
            .ColumnsHeader.Add "Adeguamento", "Adeguamento", dgBoolean, True, 1000, dgAligncenter
            
            .ColumnsHeader.Add "Periodo", "Annotazioni", dgchar, True, 2000, dgAlignleft
            
            .ColumnsHeader.Add "IDPagamentoRata", "IDPagamentoRata", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "Pagamento", "Pagamento", dgchar, True, 2000, dgAlignleft
             
            .ColumnsHeader.Add "IDCategoriaAnagraficaCliente", "IDCategoriaAnagraficaCliente", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "CategoriaAnagraficaCliente", "Categoria cliente", dgchar, False, 2000, dgAlignleft
             
            .ColumnsHeader.Add "IDRaggruppamentoFatturatoCliente", "IDRaggruppamentoFatturatoCliente", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "RaggruppamentoFatturatoCliente", "Raggr. fatt. cliente", dgchar, False, 2000, dgAlignleft
             
             Set cl = .ColumnsHeader.Add("Quantita", "Q.tà", dgDouble, True, 1500, dgAlignleft)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                    
             Set cl = .ColumnsHeader.Add("ImportoUnitario", "Imp. Uni.", dgDouble, False, 1500, dgAlignleft)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
             Set cl = .ColumnsHeader.Add("Sconto1", "% Sc. 1", dgDouble, False, 1500, dgAlignleft)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
             Set cl = .ColumnsHeader.Add("Sconto2", "% Sc. 2", dgDouble, False, 1500, dgAlignleft)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                    
            Set cl = .ColumnsHeader.Add("ImportoRata", "Importo", dgDouble, True, 2000, dgAlignleft)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                        
            .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "CodiceArticolo", "Codice articolo rata", dgchar, False, 2000, dgAlignleft
            .ColumnsHeader.Add "Articolo", "Articolo rata", dgchar, False, 2000, dgAlignleft
            
            .ColumnsHeader.Add "IDArticoloTipoContratto", "IDArticoloTipoContratto", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "CodiceArticoloTipoContratto", "Codice articolo tipo contratto", dgchar, False, 2000, dgAlignleft
            .ColumnsHeader.Add "ArticoloTipoContratto", "Descr. art. da tipo contratto", dgchar, False, 2000, dgAlignleft
            
            .ColumnsHeader.Add "IDArticoloContratto", "IDArticoloContratto", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "CodiceArticoloContratto", "Codice articolo contratto", dgchar, False, 2000, dgAlignleft
            .ColumnsHeader.Add "ArticoloContratto", "Descr. art. da contratto", dgchar, False, 2000, dgAlignleft
                
            .ColumnsHeader.Add "IDArticoloContrattoProdotto", "IDArticoloDaContrattoProdotti", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "CodiceArticoloContrattoProdotto", "Codice articolo da prodotto", dgchar, False, 2000, dgAlignleft
            .ColumnsHeader.Add "ArticoloContrattoProdotto", "Descr. art. da prodotto", dgchar, False, 2000, dgAlignleft
            
            .ColumnsHeader.Add "IDRV_POProdotto", "IDRV_POProdotto", dgInteger, False, 500, dgAlignRight
            .ColumnsHeader.Add "DescrizioneProdotto", "Prodotto collegato", dgchar, False, 2000, dgAlignleft
            .ColumnsHeader.Add "ValoreIndentificativoContrattoProdotto", "Matricola prodotto collegato", dgchar, False, 2000, dgAlignleft
            .ColumnsHeader.Add "DescrizioneAggiuntivaContrattoProdotto", "Ubicazione prodotto collegato", dgchar, False, 2000, dgAlignleft
            Set .Recordset = rsRate
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
    If Me.GrigliaRate.Recordset.EOF = False Then
        If Riga_selezionata = 0 Then
            Me.GrigliaRate.Recordset.Move Riga_selezionata
        Else
            Me.GrigliaRate.Recordset.Move Riga_selezionata + 1
        End If
            
    End If
Exit Sub
ERR_SettaggioGrigliaRate:
    MsgBox Err.Description, vbCritical, "SettaggioGrigliaRate"
End Sub
Private Sub ElaborazioneRateTemporanee()
'On Error GoTo ERR_ElaborazioneRateTemporanee
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsVista As ADODB.Recordset
Dim I As Long
    
    NumeroRecordRighe = 0
    NumeroRecordTesta = 0
    
    sSQL = "SELECT * FROM RV_PO51IEFatturazioneRate "
    sSQL = sSQL & "WHERE IDRV_PORateContratto=0"
    
    Set rsVista = New ADODB.Recordset
    rsVista.Open sSQL, CnDMT.InternalConnection
    
    If Not (rsRate Is Nothing) Then
        If rsRate.State > 0 Then
            rsRate.Close
        End If
        Set rsRate = Nothing
    End If
    
    Set rsRate = New ADODB.Recordset
    rsRate.CursorLocation = adUseClient
    
    If Not (rsRateDaFatt Is Nothing) Then
        If rsRateDaFatt.State > 0 Then
            rsRateDaFatt.Close
        End If
        Set rsRateDaFatt = Nothing
    End If
    
    Set rsRateDaFatt = New ADODB.Recordset
    rsRateDaFatt.CursorLocation = adUseClient
    
    If Not (rsRateRpt Is Nothing) Then
        If rsRateRpt.State > 0 Then
            rsRateRpt.Close
        End If
        Set rsRateRpt = Nothing
    End If
    
    Set rsRateRpt = New ADODB.Recordset
    rsRateRpt.CursorLocation = adUseClient
    
    If Not (rsRateSomma Is Nothing) Then
        If rsRateSomma.State > 0 Then
            rsRateSomma.Close
        End If
        Set rsRateSomma = Nothing
    End If
    
    Set rsRateSomma = New ADODB.Recordset
    rsRateSomma.CursorLocation = adUseClient
    
    With rsVista
        For I = 0 To rsVista.Fields.Count - 1
            
            Select Case rsVista.Fields(I).Type
                Case adChar, adVarChar, adVarWChar, adWChar, 201
                    rsRate.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
                    rsRateDaFatt.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
                    rsRateRpt.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
                    rsRateSomma.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
                Case adNumeric, adBigInt, adCurrency, adDecimal, adDouble, adInteger, adLongVarBinary, adSingle
                    rsRate.Fields.Append .Fields(I).Name, adDouble, , adFldIsNullable
                    rsRateRpt.Fields.Append .Fields(I).Name, adDouble, , adFldIsNullable
                    rsRateDaFatt.Fields.Append .Fields(I).Name, adDouble, , adFldIsNullable
                    rsRateSomma.Fields.Append .Fields(I).Name, adDouble, , adFldIsNullable
                Case adDate, adDBTimeStamp, adDBDate
                    rsRate.Fields.Append .Fields(I).Name, adDBDate, , adFldIsNullable
                    rsRateRpt.Fields.Append .Fields(I).Name, adDBDate, , adFldIsNullable
                    rsRateDaFatt.Fields.Append .Fields(I).Name, adDBDate, , adFldIsNullable
                    rsRateSomma.Fields.Append .Fields(I).Name, adDBDate, , adFldIsNullable
                Case adSmallInt, adBoolean
                    rsRate.Fields.Append .Fields(I).Name, adSmallInt, , adFldIsNullable
                    rsRateRpt.Fields.Append .Fields(I).Name, adSmallInt, , adFldIsNullable
                    rsRateDaFatt.Fields.Append .Fields(I).Name, adSmallInt, , adFldIsNullable
                    rsRateSomma.Fields.Append .Fields(I).Name, adSmallInt, , adFldIsNullable
                Case Else
                    rsRate.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
                    rsRateRpt.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
                    rsRateDaFatt.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
                    rsRateSomma.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
            End Select
        Next
        
        rsRate.Fields.Append "Fatturare", adSmallInt, , adFldIsNullable
        rsRate.Fields.Append "Quantita", adDouble, , adFldIsNullable
        rsRate.Fields.Append "Sconto1", adDouble, , adFldIsNullable
        rsRate.Fields.Append "Sconto2", adDouble, , adFldIsNullable
        rsRate.Fields.Append "IDTipo", adInteger, , adFldIsNullable
        rsRate.Fields.Append "DescrizioneTipo", adVarChar, 250, adFldIsNullable
        rsRate.Fields.Append "ImportoRataUnitaria", adDouble, , adFldIsNullable
        rsRate.Fields.Append "IDRV_POContatoreRilevamenti", adDouble, , adFldIsNullable
        rsRate.Fields.Append "DaFatturare", adSmallInt, , adFldIsNullable
        rsRate.Fields.Append "ImportoDaFatturare", adDouble, , adFldIsNullable
        
        
        
        rsRateDaFatt.Fields.Append "Fatturare", adSmallInt, , adFldIsNullable
        rsRateDaFatt.Fields.Append "Quantita", adDouble, , adFldIsNullable
        rsRateDaFatt.Fields.Append "Sconto1", adDouble, , adFldIsNullable
        rsRateDaFatt.Fields.Append "Sconto2", adDouble, , adFldIsNullable
        rsRateDaFatt.Fields.Append "IDTipo", adInteger, , adFldIsNullable
        rsRateDaFatt.Fields.Append "DescrizioneTipo", adVarChar, 250, adFldIsNullable
        rsRateDaFatt.Fields.Append "ImportoRataUnitaria", adDouble, , adFldIsNullable
        rsRateDaFatt.Fields.Append "IDRV_POContatoreRilevamenti", adDouble, , adFldIsNullable
        rsRateDaFatt.Fields.Append "DaFatturare", adSmallInt, , adFldIsNullable
        rsRateDaFatt.Fields.Append "ImportoDaFatturare", adDouble, , adFldIsNullable
        
        rsRateRpt.Fields.Append "Fatturare", adSmallInt, , adFldIsNullable
        rsRateRpt.Fields.Append "Quantita", adDouble, , adFldIsNullable
        rsRateRpt.Fields.Append "Sconto1", adDouble, , adFldIsNullable
        rsRateRpt.Fields.Append "Sconto2", adDouble, , adFldIsNullable
        rsRateRpt.Fields.Append "IDTipo", adInteger, , adFldIsNullable
        rsRateRpt.Fields.Append "DescrizioneTipo", adVarChar, 250, adFldIsNullable
        rsRateRpt.Fields.Append "ImportoRataUnitaria", adDouble, , adFldIsNullable
        rsRateRpt.Fields.Append "IDRV_POContatoreRilevamenti", adDouble, , adFldIsNullable
        rsRateRpt.Fields.Append "DaFatturare", adSmallInt, , adFldIsNullable
        rsRateRpt.Fields.Append "ImportoDaFatturare", adDouble, , adFldIsNullable
        
        
        rsRateSomma.Fields.Append "Fatturare", adSmallInt, , adFldIsNullable
        rsRateSomma.Fields.Append "Quantita", adDouble, , adFldIsNullable
        rsRateSomma.Fields.Append "Sconto1", adDouble, , adFldIsNullable
        rsRateSomma.Fields.Append "Sconto2", adDouble, , adFldIsNullable
        rsRateSomma.Fields.Append "IDTipo", adInteger, , adFldIsNullable
        rsRateSomma.Fields.Append "DescrizioneTipo", adVarChar, 250, adFldIsNullable
        rsRateSomma.Fields.Append "ImportoRataUnitaria", adDouble, , adFldIsNullable
        rsRateSomma.Fields.Append "IDRV_POContatoreRilevamenti", adDouble, , adFldIsNullable
        rsRateSomma.Fields.Append "DaFatturare", adSmallInt, , adFldIsNullable
        rsRateSomma.Fields.Append "ImportoDaFatturare", adDouble, , adFldIsNullable
        
    End With
    
    rsRate.Open , , adOpenKeyset, adLockBatchOptimistic
    rsRateRpt.Open , , adOpenKeyset, adLockBatchOptimistic
    rsRateDaFatt.Open , , adOpenKeyset, adLockBatchOptimistic
    rsRateSomma.Open , , adOpenKeyset, adLockBatchOptimistic
    
    If ((TIPO_SELEZIONE = 0) Or (TIPO_SELEZIONE = 1)) Then
        sSQL = "SELECT * FROM RV_PO51IEFatturazioneRate "
        sSQL = sSQL & " WHERE (IDAzienda=" & TheApp.IDFirm & ") "
        sSQL = sSQL & " AND ((Fatturata=0) OR (Fatturata IS NULL))"
        sSQL = sSQL & " AND ((Offerta=0" & ") OR (Offerta IS NULL))"
        
        If LINK_CONTRATTO_SELEZIONATO > 0 Then
            sSQL = sSQL & " AND IDRV_POContratto=" & LINK_CONTRATTO_SELEZIONATO
        Else
            If (Len(VAR_A_DATA) > 0) Then
                sSQL = sSQL & " AND (DataRata <=" & fnNormDate(VAR_A_DATA) & ") "
            End If
            If LINK_TIPO_CONTRATTO > 0 Then
                sSQL = sSQL & " AND (IDTipoContratto=" & LINK_TIPO_CONTRATTO & ") "
            End If
            
            If LINK_CLIENTE > 0 Then
                sSQL = sSQL & " AND (IDAnagraficaFatturazione=" & LINK_CLIENTE & ") "
            End If
            If LINK_AMMINISTRATORE > 0 Then
                sSQL = sSQL & " AND (IDAnagraficaAmministratore=" & LINK_AMMINISTRATORE & ") "
            End If
            If LINK_RAGGR_FATT_CLIENTE > 0 Then
                sSQL = sSQL & " AND (IDRaggruppamentoFatturatoCliente=" & LINK_RAGGR_FATT_CLIENTE & ") "
            End If
            If LINK_CAT_ANA_CLIENTE > 0 Then
                sSQL = sSQL & " AND (IDCategoriaAnagraficaCliente=" & LINK_CAT_ANA_CLIENTE & ") "
            End If
        End If
        
        sSQL = sSQL & "ORDER BY DataRata, Anagrafica"
        
        Set rs = New ADODB.Recordset
        rs.Open sSQL, CnDMT.InternalConnection
        
        While Not rs.EOF
            If ((fnNotNullN(rs!NonFatturare) = 0) And (fnNotNullN(rs!NonFatturareContratto) = 0)) Then
                rsRate.AddNew
                rsRateRpt.AddNew
                rsRateSomma.AddNew
                    For I = 0 To rs.Fields.Count - 1
                        rsRate.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
                        rsRateRpt.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
                        rsRateSomma.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
                    Next
                    rsRate!Fatturare = 1
                    rsRate!Quantita = 1
                    rsRate!Sconto1 = 0
                    rsRate!Sconto2 = 0
                    rsRate!IDTipo = 1
                    rsRate!DescrizioneTipo = "Rate contratto"
                    rsRate!IDRV_POContatoreRilevamenti = 0
                    rsRate!DaFatturare = 1
                    rsRate!ImportoRataUnitaria = rs!ImportoRata
                    rsRate!ImportoDaFatturare = rs!ImportoRata
                    
                    rsRateRpt!Fatturare = 1
                    rsRateRpt!Quantita = 1
                    rsRateRpt!Sconto1 = 0
                    rsRateRpt!Sconto2 = 0
                    rsRateRpt!IDTipo = 1
                    rsRateRpt!DescrizioneTipo = "Rate contratto"
                    rsRateRpt!IDRV_POContatoreRilevamenti = 0
                    rsRateRpt!DaFatturare = 1
                    rsRateRpt!ImportoRataUnitaria = rs!ImportoRata
                    rsRateRpt!ImportoDaFatturare = rs!ImportoRata
                    
                    
                    rsRateSomma!Fatturare = 1
                    rsRateSomma!Quantita = 1
                    rsRateSomma!Sconto1 = 0
                    rsRateSomma!Sconto2 = 0
                    rsRateSomma!IDTipo = 1
                    rsRateSomma!DescrizioneTipo = "Rate contratto"
                    rsRateSomma!IDRV_POContatoreRilevamenti = 0
                    rsRateSomma!DaFatturare = 1
                    rsRateSomma!ImportoRataUnitaria = rs!ImportoRata
                    rsRateSomma!ImportoDaFatturare = rs!ImportoRata
                rsRate.Update
                rsRateRpt.Update
                rsRateSomma.Update
                NumeroRecordRighe = NumeroRecordRighe + 1
            End If
        rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
    End If
    
    
    If ((TIPO_SELEZIONE = 0) Or (TIPO_SELEZIONE = 2)) Then
        GET_RILEVAMENTI_DA_FATTURARE
    End If

Exit Sub
ERR_ElaborazioneRateTemporanee:
    MsgBox Err.Description, vbCritical, "ElaborazioneRateTemporanee"
End Sub
Private Sub GrigliaRate_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    'Me.chkImporta.Value = IIf((Me.GrigliaRate.AllColumns("NonFatturare") = False), 0, 1)
    If Aggiornamento = 0 Then
        Riga_selezionata = Me.GrigliaRate.ListIndex - 1
    End If
End Sub
Private Function TotaleRateDaFatturare() As Double

'rsRateSomma.Filter = "Fatturare=1"
'TotaleRateDaFatturare = 0
'If rsRateSomma.EOF Then
'    TotaleRateDaFatturare = 0
'Else
'    rsRateSomma.MoveFirst
'    While Not rsRateSomma.EOF
'        TotaleRateDaFatturare = TotaleRateDaFatturare + fnNotNullN(rsRateSomma!ImportoRata)
'        DoEvents
'    rsRateSomma.MoveNext
'    Wend
'
'End If
'rsRateSomma.Filter = vbNullString

Me.GrigliaRate.UpdatePosition = False

rsRate.Filter = "Fatturare=1"
TotaleRateDaFatturare = 0
If rsRate.EOF Then
    TotaleRateDaFatturare = 0
Else
    rsRate.MoveFirst
    While Not rsRate.EOF
        TotaleRateDaFatturare = TotaleRateDaFatturare + fnNotNullN(rsRate!ImportoRata)
        DoEvents
    rsRate.MoveNext
    Wend
    
End If
rsRate.Filter = vbNullString
Me.GrigliaRate.Refresh

End Function
Private Function fncIDTipoOggettoPrg() As Long
    Dim rs As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto, Gestore.Gestore"
    sSQL = sSQL & " FROM Gestore INNER JOIN TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore"
    sSQL = sSQL & " WHERE (((Gestore.Gestore)=" & fnNormString(App.EXEName) & "))"
    
    Set rs = CnDMT.OpenResultset(sSQL)
        
    If rs.EOF = False Then
        fncIDTipoOggettoPrg = rs!IDTipoOggetto
    Else
        fncIDTipoOggettoPrg = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
End Function
Private Sub GrigliaRate_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GrigliaRate.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsRate.Fields("Fatturare").Value), 2
        End If
    End If
    
        
    TotaleDocumento = TotaleRateDaFatturare
    Me.lblTotaleDocumento.Caption = FormatCurrency(TotaleDocumento, 2)
End Sub

Private Sub GrigliaRate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If GrigliaRate.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < GrigliaRate.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If GrigliaRate.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsRate.Fields("Fatturare").Value), 2
            End If
        End If
    End If

    TotaleDocumento = TotaleRateDaFatturare
    Me.lblTotaleDocumento.Caption = FormatCurrency(TotaleDocumento, 2)
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    
        If Not rsRate.EOF And Not rsRate.BOF Then
            rsRate.Fields("Fatturare").Value = Abs(CLng(Selected))
            Me.GrigliaRate.Refresh
        
            'AGGIORNA_RECORDSET fnNotNullN(rsRate.Fields("IDRV_PORateContratto").Value), fnNotNullN(rsRate.Fields("Fatturare").Value)
            'TotaleDocumento = TotaleRateDaFatturare
            
        
        End If
        DoEvents

End Sub

Private Function fncImpostaDefaultReport(ByVal IDReportDefault As Long, IDTipoOggetto As Long)
On Error GoTo ERR_fncImpostaDefaultReport
    Dim sSQL As String
    
    sSQL = "UPDATE DefaultFilialePerTipoOggetto SET "
    sSQL = sSQL & "IDReportTipoOggetto=" & IDReportDefault
    sSQL = sSQL & " WHERE IDTipoOggetto = " & IDTipoOggetto
    sSQL = sSQL & " AND IDFiliale = " & TheApp.Branch
    
    CnDMT.Execute sSQL
    
Exit Function
ERR_fncImpostaDefaultReport:
    MsgBox Err.Description, vbCritical, "Settaggio report di default"
End Function
Private Function fncTrovaReport(NomeReport As String, IDTipoOggetto As Long) As Long
On Error GoTo ERR_fncTrovaReport
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDReportTipoOggetto FROM ReportTipoOggetto "
sSQL = sSQL & "WHERE ((ReportTipoOggetto=" & fnNormString(NomeReport) & ") AND (IDTipoOggetto=" & IDTipoOggetto & "))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaReport = rs!IDReportTipoOggetto
Else
    fncTrovaReport = 0
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_fncTrovaReport:
    MsgBox Err.Description, vbCritical, "Trova report per stampa"
    fncTrovaReport = 0

End Function



Private Sub VScroll1_Change()
   Me.Pic1.Top = -VScroll1.Value
End Sub
Private Sub VScroll1_Scroll()
   Me.Pic1.Top = -VScroll1.Value
End Sub
Private Sub HScroll1_Change()
   Me.Pic1.Left = -HScroll1.Value
End Sub
Private Sub HScroll1_Scroll()
   Me.Pic1.Left = -HScroll1.Value
End Sub
Private Sub ConfigurazioneStampe()
Dim oActivity As IActivity
Dim o As Activity
Dim oFilter As Filter

    'Inizializzazione del riquadro attività
    With ActivityBox
        .Activities.Clear
        
        'Aggiunge l'attività dei reports
        Set oActivity = .Activities.Add("DmtActBoxLib.ReportsActivity", "Reports")
        Set oActivity.Connection = CnDMT.InternalConnection
        
        oActivity.Load fnGetTipoOggetto("RV_POCreazioneDocumenti"), TheApp.IDFirm
        Set o = oActivity
        Set oReportsActivity = o.InternalClass
        
        'Imposta quale attività deve essere attivata per default
        If m_DefaultActivity <> "" Then
            Set .CurrentActivity = .Activities(m_DefaultActivity)
        End If
        
        'ridisegna il controllo
        .Redraw = True
        
        oReportsActivity.Is4DlgPrint = False
    End With

End Sub
Private Function fnGetTipoOggetto(Optional Gestore As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    If Gestore = "" Then
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(App.EXEName)
    Else
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(Gestore)
    End If
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function


Private Sub Form_Resize()
On Error GoTo ERR_Form_Resize
  If Me.WindowState <> 1 Then
        If VisualizzaStampe = 0 Then
            If Me.Width > 19365 Then
                Me.Pic1.Width = Me.Width - 270
                Me.Picture2.Width = Me.Pic1.Width - 240
                Me.GrigliaRate.Width = Me.Picture2.Width - 240
                
                Me.lblTotaleDocumento.Left = Me.GrigliaRate.Width - 120 - Me.lblTotaleDocumento.Width
                Me.Label1.Left = Me.lblTotaleDocumento.Left - 60 - Me.Label1.Width
                Me.cmdAvanti.Left = Me.GrigliaRate.Width - 120 - Me.cmdAvanti.Width
                Me.cmdAnnulla.Left = Me.cmdAvanti.Left - 60 - Me.cmdAnnulla.Width
                Me.cmdIndietro.Left = Me.cmdAnnulla.Left - 60 - Me.cmdIndietro.Width
                
            End If
        End If
        If Me.Height > 9345 Then
            Me.Pic1.Height = Me.Height - 580
            Me.Picture2.Height = Me.Pic1.Height - 240
            
            Me.cmdAvanti.Top = Me.Picture2.Height - 240 - Me.cmdAvanti.Height
            Me.cmdAnnulla.Top = Me.cmdAvanti.Top
            Me.cmdIndietro.Top = Me.cmdAvanti.Top
            Me.cmdSelezionaTutto.Top = Me.cmdAvanti.Top
            Me.cmdDeselezionaTutto.Top = Me.cmdAvanti.Top
            Me.cmdVisStampa.Top = Me.cmdAvanti.Top
            
            Me.lblTotaleDocumento.Top = Me.cmdAvanti.Top - 240 - Me.lblTotaleDocumento.Height
            Me.Label1.Top = Me.lblTotaleDocumento.Top
            Me.GrigliaRate.Height = Me.lblTotaleDocumento.Top - 240
        End If

        If Me.ScaleWidth < Me.Pic1.ScaleWidth Then
            Me.HScroll1.Visible = True
            Me.HScroll1.Top = Me.ScaleHeight - Me.HScroll1.Height
            Me.HScroll1.Left = 0
            
        Else
            Me.HScroll1.Visible = False
        End If
        
        If Me.ScaleHeight < Me.Pic1.ScaleHeight Then
            Me.VScroll1.Visible = True
            Me.VScroll1.Top = 0
            Me.VScroll1.Left = Me.ScaleWidth - Me.VScroll1.Width
            
        Else
            Me.VScroll1.Visible = False
        End If
        
        If (VScroll1.Visible = True) And (HScroll1.Visible = False) Then
            Me.VScroll1.Height = Me.ScaleHeight '- Me.HScroll1.Height
        Else
            Me.VScroll1.Height = Me.ScaleHeight - Me.HScroll1.Height
        End If
        
        If (HScroll1.Visible = True) And (HScroll1.Visible = True) Then
            Me.HScroll1.Width = Me.ScaleWidth '- Me.VScroll1.Width
        Else
            Me.HScroll1.Width = Me.ScaleWidth - Me.VScroll1.Width
        End If
            
        With HScroll1
            .Max = (Pic1.ScaleWidth - Me.ScaleWidth + Me.VScroll1.Width)
            If .Max > 0 Then
                .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With

        With VScroll1
            .Max = (Pic1.ScaleHeight - Me.ScaleHeight + Me.HScroll1.Height)
            If .Max > 0 Then
                 .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With
        
        
    End If
Exit Sub
ERR_Form_Resize:
    MsgBox Err.Description, vbCritical, "Form_Resize"
End Sub

Private Sub AGGIORNA_RECORDSET(IDRataContratto As Long, Fatturare As Long)
    rsRateRpt.Filter = "IDRV_PORateContratto=" & IDRataContratto
    rsRateSomma.Filter = "IDRV_PORateContratto=" & IDRataContratto
    
    rsRateRpt!Fatturare = Fatturare
    rsRateSomma!Fatturare = Fatturare
    
    
    rsRateRpt.Update
    rsRateSomma.Update
    
    rsRateRpt.Filter = vbNullString
    rsRateSomma.Filter = vbNullString
    
End Sub
Private Sub AGGIORNA_TABELLA_PER_REPORT()
On Error GoTo ERR_AGGIORNA_TABELLA_PER_REPORT
Dim sSQL As String

sSQL = "DELETE FROM RV_POTMPFatturazioneRate "
CnDMT.Execute sSQL

Screen.MousePointer = 11
rsRateRpt.Filter = vbNullString

If Not rsRateRpt.EOF Then
    While Not rsRateRpt.EOF
        sSQL = "INSERT INTO RV_POTMPFatturazioneRate ("
        sSQL = sSQL & "IDRV_PORateContratto, DaFatturare) "
        sSQL = sSQL & " VALUES ("
        sSQL = sSQL & fnNotNullN(rsRateRpt!IDRV_PORateContratto) & ", "
        sSQL = sSQL & fnNotNullN(rsRateRpt!Fatturare) & ")"
        CnDMT.Execute sSQL
        DoEvents
    rsRateRpt.MoveNext
    Wend
End If

rsRateRpt.Filter = vbNullString
Screen.MousePointer = 0
Exit Sub
ERR_AGGIORNA_TABELLA_PER_REPORT:
    MsgBox Err.Description, vbCritical, "AGGIORNA_TABELLA_PER_REPORT"
End Sub
Private Sub GET_RILEVAMENTI_DA_FATTURARE()
On Error GoTo ERR_GET_RILEVAMENTI_DA_FATTURARE
Dim sSQL As String
Dim rs As ADODB.Recordset

    sSQL = "SELECT * FROM RV_POIEContatoreRilevamentiFatt "
    sSQL = sSQL & " WHERE (IDAzienda=" & TheApp.IDFirm & ") "
    sSQL = sSQL & " AND (DataRata <=" & fnNormDate(VAR_A_DATA) & ") "
    sSQL = sSQL & " AND (DaFatturare=1) " ' OR (NonFatturareContratto=0)) "
    sSQL = sSQL & " AND ((Fatturata=0) OR (Fatturata IS NULL))"
    
    If LINK_CONTRATTO_SELEZIONATO > 0 Then
        sSQL = sSQL & " AND IDRV_POContratto=" & LINK_CONTRATTO_SELEZIONATO
    Else
    
        If LINK_TIPO_CONTRATTO > 0 Then
            sSQL = sSQL & " AND (IDTipoContratto=" & LINK_TIPO_CONTRATTO & ") "
        End If
        
        If LINK_CLIENTE > 0 Then
            sSQL = sSQL & " AND (IDAnagraficaFatturazione=" & LINK_CLIENTE & ") "
        End If
        If LINK_AMMINISTRATORE > 0 Then
            sSQL = sSQL & " AND (IDAnagraficaAmministratore=" & LINK_AMMINISTRATORE & ") "
        End If
    End If
    
    sSQL = sSQL & "ORDER BY DataRata, Anagrafica"
    
    Set rs = New ADODB.Recordset
    rs.Open sSQL, CnDMT.InternalConnection
    
    While Not rs.EOF
        If (rs!NonFatturareContratto = 0) Then
            rsRate.AddNew
            rsRateRpt.AddNew
            rsRateSomma.AddNew
                For I = 0 To rs.Fields.Count - 1
    '                If rs.Fields(I).Name = "IDPagamentoRata" Then
    '                    MsgBox "STOP"
    '                End If
                    rsRate.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
                    rsRateRpt.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
                    rsRateSomma.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
                Next
                
                rsRate!Fatturare = 1
                rsRate!IDTipo = 2
                rsRate!DescrizioneTipo = "Rilevamento eccedenze contatori"
    
                
                rsRateRpt!Fatturare = 1
                rsRateRpt!IDTipo = 2
                rsRateRpt!DescrizioneTipo = "Rilevamento eccedenze contatori"
                
                rsRateSomma!Fatturare = 1
                rsRateSomma!IDTipo = 2
                rsRateSomma!DescrizioneTipo = "Rilevamento eccedenze contatori"
    
                
            rsRate.Update
            rsRateRpt.Update
            rsRateSomma.Update
            NumeroRecordRighe = NumeroRecordRighe + 1
        End If
    rs.MoveNext
    Wend

    
    
rs.Close
Set rs = Nothing
Exit Sub
ERR_GET_RILEVAMENTI_DA_FATTURARE:
    MsgBox Err.Description, vbCritical, "GET_RILEVAMENTI_DA_FATTURARE"
End Sub
