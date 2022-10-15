VERSION 5.00
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.9#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmAltriDati 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Altri dati del contratto"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6870
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
   ScaleHeight     =   8250
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   38
      Top             =   7560
      Width           =   6615
   End
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.Frame Frame2 
         Caption         =   "Inserimento/Ultima modifica"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1455
         Left            =   120
         TabIndex        =   29
         Top             =   5640
         Width           =   6375
         Begin DMTDATETIMELib.dmtDate txtDataInserimento 
            Height          =   315
            Left            =   4320
            TabIndex        =   30
            Top             =   465
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   556
            _StockProps     =   253
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboUtenteInserimento 
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Top             =   465
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin DMTDataCmb.DMTCombo cboUtenteModifica 
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Top             =   1050
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin DMTDATETIMELib.dmtDate txtDataModifica 
            Height          =   315
            Left            =   4320
            TabIndex        =   33
            Top             =   1050
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   556
            _StockProps     =   253
            BackColor       =   16777215
            Enabled         =   0   'False
            Appearance      =   1
         End
         Begin VB.Label Label7 
            Caption         =   "Utente ultima modifica"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   4095
         End
         Begin VB.Label Label6 
            Caption         =   "Data di inserimento"
            Height          =   255
            Left            =   4320
            TabIndex        =   36
            Top             =   225
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Utente di inserimento"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label8 
            Caption         =   "Data ultima modifica"
            Height          =   255
            Left            =   4320
            TabIndex        =   34
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.Frame FraRapprLegali 
         Caption         =   "Rappresentanti"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1575
         Left            =   120
         TabIndex        =   20
         Top             =   4080
         Width           =   6375
         Begin VB.ComboBox cboRapprAzienda 
            Height          =   315
            ItemData        =   "frmAltriDati.frx":0000
            Left            =   120
            List            =   "frmAltriDati.frx":0002
            TabIndex        =   24
            Top             =   480
            Width           =   3375
         End
         Begin VB.ComboBox cboRuoloRapprAzienda 
            Height          =   315
            ItemData        =   "frmAltriDati.frx":0004
            Left            =   3600
            List            =   "frmAltriDati.frx":0006
            TabIndex        =   23
            Top             =   480
            Width           =   2655
         End
         Begin VB.ComboBox cboRapprCliente 
            Height          =   315
            ItemData        =   "frmAltriDati.frx":0008
            Left            =   120
            List            =   "frmAltriDati.frx":000A
            TabIndex        =   22
            Top             =   1080
            Width           =   3375
         End
         Begin VB.ComboBox cboRuoloRapprCliente 
            Height          =   315
            ItemData        =   "frmAltriDati.frx":000C
            Left            =   3600
            List            =   "frmAltriDati.frx":000E
            TabIndex        =   21
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Rappresentante azienda"
            Height          =   255
            Index           =   44
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label1 
            Caption         =   "Ruolo"
            Height          =   255
            Index           =   45
            Left            =   3600
            TabIndex        =   27
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Rappresentante cliente"
            Height          =   255
            Index           =   46
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label Label1 
            Caption         =   "Ruolo"
            Height          =   255
            Index           =   47
            Left            =   3600
            TabIndex        =   25
            Top             =   840
            Width           =   2655
         End
      End
      Begin VB.TextBox txtCodiceConto 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtDescrizioneConto 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2400
         Width           =   4695
      End
      Begin DMTDataCmb.DMTCombo cboClassificazioneContratto 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   2985
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboRaggrFatturato 
         Height          =   315
         Left            =   3240
         TabIndex        =   3
         Top             =   3000
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboAccordoCommerciale 
         Height          =   315
         Left            =   3360
         TabIndex        =   5
         Top             =   1155
         Width           =   3015
         _ExtentX        =   5318
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
      Begin DMTDataCmb.DMTCombo cboContrattoBancario 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1155
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DmtCodDescCtl.DmtCodDesc CDArticoloContratto 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   1085
         PropCodice      =   $"frmAltriDati.frx":0010
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmAltriDati.frx":0060
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmAltriDati.frx":00CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin DmtCodDescCtl.DmtCodDesc CDAnaFatt 
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   1085
         PropCodice      =   $"frmAltriDati.frx":0124
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmAltriDati.frx":0185
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmAltriDati.frx":01D5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroLicenze 
         Height          =   315
         Left            =   3600
         TabIndex        =   14
         Top             =   3555
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtCurrency txtMaggiorazioneIstat 
         Height          =   315
         Left            =   2160
         TabIndex        =   15
         Top             =   3555
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   " 0"
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
         UseSeparator    =   -1  'True
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTDataCmb.DMTCombo cboIstat 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   3555
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "N° licenze"
         Height          =   255
         Index           =   15
         Left            =   3600
         TabIndex        =   19
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "I.S.T.A.T."
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   18
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Maggiorazione"
         Height          =   255
         Index           =   33
         Left            =   2160
         TabIndex        =   17
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label lblPianodeiDeiConti 
         Caption         =   "Piano dei conti"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Contratto bancario"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Accordo commerciale"
         Height          =   255
         Index           =   39
         Left            =   3360
         TabIndex        =   8
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Raggruppamento fatturato"
         Height          =   255
         Index           =   34
         Left            =   3240
         TabIndex        =   4
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Classificazione"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   2
         Top             =   2760
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmAltriDati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CDAnaFatt_ChangeElement()
    
    INIT_CONTROLLI_ANA_FATT Me.CDAnaFatt.KeyFieldID
    
End Sub



Private Sub cmdConferma_Click()
    
    CONFERMA_MOD_ALTRI_DATI
    
    CONFERMA_MODIFICA = 1
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    INIT_CONTROLLI
    
    CONFERMA_MODIFICA = 0
    
    INIT_VALORI_CONTROLLI
    
End Sub
Private Sub INIT_VALORI_CONTROLLI()
    
    Me.CDAnaFatt.Load IDClienteFatturazione
    Me.CDArticoloContratto.Load IDArticoloContratto

    Me.cboContrattoBancario.WriteOn IDContrattoBancario
    Me.cboRaggrFatturato.WriteOn IDRaggrFattContratto
    Me.cboClassificazioneContratto.WriteOn IDClassContratto
    Me.cboAccordoCommerciale.WriteOn IDAccordoCommerciale
    
    Link_ContoPDC = IDPDCContratto
    Me.txtCodiceConto.Text = CodicePDCContratto
    Me.txtDescrizioneConto.Text = DescrPDCContratto
    
    Me.cboIstat.WriteOn IDIstatContratto
    Me.txtMaggiorazioneIstat.Value = MaggIstatContratto
    Me.txtNumeroLicenze.Value = NumeroLicenze
    
    
    Me.cboRapprAzienda.Text = RapprLegaleAzienda
    Me.cboRuoloRapprAzienda.Text = RuoloRapprAzienda
    Me.cboRapprCliente.Text = RapprLegaleCliente
    Me.cboRuoloRapprCliente.Text = RuoloRapprCliente
    
    Me.cboUtenteInserimento.WriteOn IDUtenteInserimento
    Me.txtDataInserimento.Text = DataInsContratto
    
    Me.cboUtenteModifica.WriteOn IDUtenteUltMod
    Me.txtDataModifica.Text = DataUltMod
    
End Sub


Private Sub lblPianodeiDeiConti_Click(Index As Integer)
        If frmMain.txtDataDecorrenza.Value = 0 Then
            VarIDEsercizio = fnGetEsercizio(Date)
        Else
            VarIDEsercizio = fnGetEsercizio(frmMain.txtDataDecorrenza.Text)
        End If
        
        Link_PianoDeiConti = GetPianoDeiConti
    
        SetPDCProperties
        
        'oPDC.ShowSearchDialog
        
        'ShowNodeProperties oPDC.SelectedNode
End Sub
        
Private Function GetPianoDeiConti() As Long
'On Error GoTo ERR_GetPianoDeiConti
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    sSQL = "SELECT IDPianoDeiConti FROM PianoDeiConti WHERE ("
    sSQL = sSQL & "(IDAzienda = " & TheApp.Branch & ") AND "
    sSQL = sSQL & "(TipoPDC = " & 1 & ") AND "
    sSQL = sSQL & "(IDEsercizio= " & VarIDEsercizio & "))"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        GetPianoDeiConti = fnNotNullN(rs!IDPianoDeiConti)
    Else
        GetPianoDeiConti = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
Exit Function
ERR_GetPianoDeiConti:
    MsgBox Err.Description, vbCritical, "Errore piano dei conti"
End Function

Private Sub SetPDCProperties()
Dim oNode As DmtPDC.INode
Dim oBranch As DmtPDC.Branch
Dim oNode1 As DmtPDC.INode
    Set oPDC = New DmtPDC.PDCServices
    'Imposta le proprietà dell'oggetto PDCServices
    With oPDC
        'Viene fornita al controllo la connessione al database DMT.
        'La connessione è di tipo ADO.Connection quindi viene
        'passata la proprietà InternalConnection dell'oggetto Database
        Set .Connection = TheApp.Database.InternalConnection
        'Indica l'identificativo del Piano dei conti da visualizzare
        .IDPDC = Link_PianoDeiConti
        .HideAccounts = False
        .BranchType = btcAllBranchs
        '.BranchType = .BranchType + btcRevenuesBranch
        .AccountType = atcAllAccounts
        If Len(Me.txtCodiceConto.Text) > 0 Then
            Set oNode = .SearchNodeExtended(Me.txtCodiceConto.Text)
        Else
            Set oNode = .SearchNodeExtended(, Me.txtDescrizioneConto.Text)
        End If
        If .RecordFounded = 1 Then
                If TypeName(oNode) = "Account" Then
                
                    Link_ContoPDC = oNode.ID
                    'Codifica completa del Conto o del Ramo
                    Me.txtCodiceConto = oNode.CompletedCode
                    Me.txtDescrizioneConto.Text = oNode.Description
                
                Else
                
                    If Len(Me.txtCodiceConto.Text) > 0 Then
                        .SelectedNode.CompletedCode = Me.txtCodiceConto.Text
                    Else
                        .SelectedNode.Description = Me.txtDescrizioneConto.Text
                    End If
                        
                    .ShowSearchDialog
                    
                    ShowNodeProperties oPDC.SelectedNode
                End If
            
        ElseIf .RecordFounded > 1 Then
            If Len(Me.txtCodiceConto.Text) > 0 Then
                .SelectedNode.CompletedCode = Me.txtCodiceConto.Text
            Else
                .SelectedNode.Description = Me.txtDescrizioneConto.Text
            End If

            .ShowSearchDialog
            
            ShowNodeProperties oPDC.SelectedNode
            
        Else
            If Len(Me.txtCodiceConto.Text) > 0 Then
                .SelectedNode.CompletedCode = Me.txtCodiceConto.Text
            Else
                .SelectedNode.Description = Me.txtDescrizioneConto.Text
            End If
        
        
            .ShowSearchDialog
            
            ShowNodeProperties oPDC.SelectedNode
            
        End If
    End With
    
    Set oPDC = Nothing
    
End Sub
Private Sub ShowNodeProperties(ByVal oNode As DmtPDC.INode)
    'Rappresenta un conto
    Dim oAccount As DmtPDC.Account
    'Rappresenta un ramo
    Dim oBranch As DmtPDC.Branch
    
    'Vengono visualizzati nei campi appositi tutte
    'le caratteristiche del conto o ramo selezionato
        
    'Controlla se è stato passato un elemento valido
    If Not oNode Is Nothing Then
        'Riporta i dati comuni del conto o del ramo
        
        'Identificativo unico del Conto o del Ramo
        Link_ContoPDC = oNode.ID
        'Codifica completa del Conto o del Ramo
        Me.txtCodiceConto = oNode.CompletedCode
        Me.txtDescrizioneConto.Text = oNode.Description
    End If
End Sub

Private Sub txtCodiceConto_LostFocus()
    If Len(Me.txtCodiceConto.Text) > 0 Then
        lblPianodeiDeiConti_Click 0
    End If
End Sub



Private Sub txtDescrizioneConto_LostFocus()
    If Len(Me.txtDescrizioneConto.Text) > 0 Then
        lblPianodeiDeiConti_Click 0
    End If
End Sub

Private Sub INIT_CONTROLLI()

    'Anagrafica di fatturazione
    With Me.CDAnaFatt
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Anagrafica"
        .DescriptionField = "Nome"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Anagrafica"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Nome"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Anagrafica"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Nome"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 29 'Anagrafica
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    With Me.CDArticoloContratto
       Set .Application = TheApp
       Set .Database = TheApp.Database
       .HwndContainer = Me.hwnd
       .CodeField = "CodiceArticolo"
       .DescriptionField = "Articolo"
       .KeyField = "IDArticolo"
       .TableName = "Articolo"
       .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm
       .MenuFunctions("EseguiGestione").Enabled = True
       .PropCodice.Caption = "Codice"
       .PropDescrizione.Caption = "Descrizione"
       .CodeCaption4Find = "Codice Articolo"
       .DescriptionCaption4Find = "Descrizione Articolo"
       .IDExecuteFunction = 6 'Articoli
       .CodeIsNumeric = False
    End With
    
    
    'Raggruppamento fatturato
    With Me.cboRaggrFatturato
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRaggruppamentoFatturato"
        .DisplayField = "RaggruppamentoFatturato"
        .SQL = "SELECT * FROM RaggruppamentoFatturato ORDER BY RaggruppamentoFatturato"
        .Fill
    End With

    'Classificazione tipo contratto
    With Me.cboClassificazioneContratto
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoClassificazioneContratto"
        .DisplayField = "TipoClassificazioneContratto"
        .SQL = "SELECT * FROM RV_POTipoClassificazioneContratto ORDER BY TipoClassificazioneContratto"
        .Fill
    End With


    'Istat per contratto
    With Me.cboIstat
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POIstat"
        .DisplayField = "Istat"
        .SQL = "SELECT * FROM RV_POIstat"
        .Fill
    End With
    
    'Utente di inserimento
    With Me.cboUtenteInserimento
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDUtente"
        .DisplayField = "Utente"
        .SQL = "SELECT * FROM Utente"
        .Fill
    End With

    'Utente di ultima modifica
    With Me.cboUtenteModifica
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDUtente"
        .DisplayField = "Utente"
        .SQL = "SELECT * FROM Utente"
        .Fill
    End With
    
    
    INIT_RAPPR_AZIENDA
    INIT_RAPPR_CLIENTE
    INIT_RAPPR_RUOLI
    
    
End Sub

Private Sub INIT_CONTROLLI_ANA_FATT(IDAnagrafica As Long)
Dim sSQL As String
    
With Me.cboContrattoBancario
    Set .Database = TheApp.Database.Connection
    .AddFieldKey "IDBancaPerAnagrafica"
    .DisplayField = "BancaPerAnagrafica"
    .SQL = "SELECT IDBancaPerAnagrafica, BancaPerAnagrafica "
    .SQL = .SQL & "FROM BancaPerAnagrafica "
    .SQL = .SQL & "WHERE IDAnagrafica=" & IDAnagrafica
    .SQL = .SQL & " AND IDAzienda=" & TheApp.IDFirm
    .Fill
End With
    
    
'Accordi commerciali
With Me.cboAccordoCommerciale
    Set .Database = TheApp.Database.Connection
    .AddFieldKey "IDAccordiCommerciali"
    .DisplayField = "Descrizione"
    .SQL = "SELECT * FROM AccordiCommerciali "
    .SQL = .SQL & "WHERE IDAzienda=" & TheApp.IDFirm
    .SQL = .SQL & " AND IDAnagrafica=" & IDAnagrafica
    .SQL = .SQL & " AND IDTipoAnagrafica=2"
    .Fill
End With
    
End Sub
Private Sub INIT_RAPPR_AZIENDA()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Descrizione FROM RV_PORappresentantiLegaliAna "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica IS NULL "

Set rs = Cn.OpenResultset(sSQL)

Me.cboRapprAzienda.Clear

While Not rs.EOF
    Me.cboRapprAzienda.AddItem fnNotNull(rs!Descrizione)
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub INIT_RAPPR_CLIENTE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Descrizione FROM RV_PORappresentantiLegaliAna "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & frmMain.CDCliente.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)

Me.cboRapprCliente.Clear

While Not rs.EOF
    Me.cboRapprCliente.AddItem fnNotNull(rs!Descrizione)
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub INIT_RAPPR_RUOLI()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Descrizione FROM RV_PORappresentantiLegaliRuolo "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

Me.cboRuoloRapprCliente.Clear
Me.cboRuoloRapprAzienda.Clear

While Not rs.EOF
    Me.cboRuoloRapprCliente.AddItem fnNotNull(rs!Descrizione)
    Me.cboRuoloRapprAzienda.AddItem fnNotNull(rs!Descrizione)
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
