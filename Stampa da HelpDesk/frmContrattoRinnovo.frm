VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmContrattiPerRinnovo 
   Caption         =   "Rinnovi contratti (Passo 2 di 3)"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15870
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
   ScaleHeight     =   7710
   ScaleWidth      =   15870
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2160
      TabIndex        =   41
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      Height          =   7695
      Left            =   0
      ScaleHeight     =   7635
      ScaleWidth      =   15795
      TabIndex        =   20
      Top             =   0
      Width           =   15855
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Left            =   120
         ScaleHeight     =   7425
         ScaleWidth      =   15585
         TabIndex        =   21
         Top             =   120
         Width           =   15615
         Begin DmtGridCtl.DmtGrid GrigliaDMT 
            Height          =   6375
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   15375
            _ExtentX        =   27120
            _ExtentY        =   11245
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
         Begin VB.Frame Frame1 
            Height          =   2775
            Left            =   120
            TabIndex        =   23
            Top             =   3960
            Width           =   13335
            Begin VB.CheckBox chkRinnovoAutomatico 
               Caption         =   "Rinnovo automatico"
               Height          =   255
               Left            =   10440
               TabIndex        =   11
               Top             =   480
               Width           =   2775
            End
            Begin VB.CheckBox chkAdeguamentoIstat 
               Caption         =   "Adeguamento I.S.T.A.T"
               Height          =   255
               Left            =   10440
               TabIndex        =   12
               Top             =   840
               Width           =   2775
            End
            Begin VB.CheckBox chkUsaAdeguamentoIstat 
               Caption         =   "Usa adeguamento I.S.T.A.T"
               Height          =   255
               Left            =   10440
               TabIndex        =   13
               Top             =   1200
               Width           =   2775
            End
            Begin VB.CheckBox chkContrattoDaRinnovare 
               Caption         =   "Contratto da rinnovare"
               Height          =   255
               Left            =   10440
               TabIndex        =   14
               Top             =   1560
               Width           =   2415
            End
            Begin VB.CommandButton cmdAggiorna 
               Caption         =   "AGGIORNA"
               Enabled         =   0   'False
               Height          =   375
               Left            =   8520
               TabIndex        =   24
               Top             =   2280
               Width           =   1695
            End
            Begin DMTEDITNUMLib.dmtCurrency txtImportoContratto 
               Height          =   315
               Left            =   6480
               TabIndex        =   8
               Top             =   1080
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
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
               Appearance      =   1
               UseSeparator    =   -1  'True
               CurrencySymbol  =   ""
               AllowEmpty      =   0   'False
               DecFinalZeros   =   -1  'True
            End
            Begin DMTDataCmb.DMTCombo cboPagamento 
               Height          =   315
               Left            =   3240
               TabIndex        =   7
               Top             =   1080
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
            Begin DMTDataCmb.DMTCombo cboRateizzazione 
               Height          =   315
               Left            =   120
               TabIndex        =   6
               Top             =   1080
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
            Begin DMTDATETIMELib.dmtDate txtDataScadenzaRinnovo 
               Height          =   315
               Left            =   8520
               TabIndex        =   5
               Top             =   480
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   556
               _StockProps     =   253
               BackColor       =   16777215
               Appearance      =   1
            End
            Begin DMTDataCmb.DMTCombo CboRinnovoContratto 
               Height          =   315
               Left            =   6120
               TabIndex        =   4
               Top             =   480
               Width           =   2295
               _ExtentX        =   4048
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
            Begin DMTDATETIMELib.dmtDate txtDataScadenza 
               Height          =   315
               Left            =   4320
               TabIndex        =   3
               Top             =   480
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   556
               _StockProps     =   253
               BackColor       =   16777215
               Appearance      =   1
            End
            Begin DMTDATETIMELib.dmtDate txtDataDecorrenza 
               Height          =   315
               Left            =   120
               TabIndex        =   1
               Top             =   480
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   556
               _StockProps     =   253
               BackColor       =   16777215
               Appearance      =   1
            End
            Begin DMTDataCmb.DMTCombo cboDurataContratto 
               Height          =   315
               Left            =   1920
               TabIndex        =   2
               Top             =   480
               Width           =   2295
               _ExtentX        =   4048
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
            Begin DMTDATETIMELib.dmtDate txtDataAssistenza 
               Height          =   315
               Left            =   3240
               TabIndex        =   10
               Top             =   1680
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   556
               _StockProps     =   253
               BackColor       =   16777215
               Appearance      =   1
            End
            Begin DMTDataCmb.DMTCombo cboDurataAssistenza 
               Height          =   315
               Left            =   120
               TabIndex        =   9
               Top             =   1680
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
            Begin VB.Label Label2 
               Caption         =   "Durata contratto"
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   35
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label Label2 
               Caption         =   "Data decorrenza"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   34
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label2 
               Caption         =   "Scadenza contratto"
               Height          =   255
               Index           =   2
               Left            =   4320
               TabIndex        =   33
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label2 
               Caption         =   "Rinnovo contratto"
               Height          =   255
               Index           =   3
               Left            =   6120
               TabIndex        =   32
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label Label2 
               Caption         =   "Scadenza rinnovo"
               Height          =   255
               Index           =   4
               Left            =   8520
               TabIndex        =   31
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label2 
               Caption         =   "Rateizzazione"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   30
               Top             =   840
               Width           =   3015
            End
            Begin VB.Label Label2 
               Caption         =   "Pagamento"
               Height          =   255
               Index           =   6
               Left            =   3240
               TabIndex        =   29
               Top             =   840
               Width           =   2175
            End
            Begin VB.Label Label2 
               Caption         =   "Importo contratto"
               Height          =   255
               Index           =   7
               Left            =   6480
               TabIndex        =   28
               Top             =   840
               Width           =   1815
            End
            Begin VB.Image Image1 
               Height          =   480
               Left            =   120
               Picture         =   "frmContrattoRinnovo.frx":0000
               Top             =   2160
               Width           =   480
            End
            Begin VB.Label lblInfo 
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   840
               TabIndex        =   27
               Top             =   2280
               Width           =   7575
            End
            Begin VB.Label Label2 
               Caption         =   "Scadenza assistenza"
               Height          =   255
               Index           =   8
               Left            =   3240
               TabIndex        =   26
               Top             =   1440
               Width           =   1935
            End
            Begin VB.Label Label2 
               Caption         =   "Durata assistenza"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   25
               Top             =   1440
               Width           =   3015
            End
         End
         Begin VB.CommandButton cmdAvanti 
            Caption         =   "Avanti"
            Height          =   375
            Left            =   14400
            TabIndex        =   19
            Top             =   6840
            Width           =   1095
         End
         Begin VB.CommandButton cmdIndietro 
            Caption         =   "Indietro"
            Height          =   375
            Left            =   10800
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   6840
            Width           =   1095
         End
         Begin VB.CommandButton cmdAnnulla 
            Caption         =   "Annulla"
            Height          =   375
            Left            =   12000
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   6840
            Width           =   1095
         End
         Begin VB.CommandButton Fine 
            Caption         =   "Fine"
            Enabled         =   0   'False
            Height          =   375
            Left            =   13200
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   6840
            Width           =   1095
         End
         Begin VB.CommandButton cmdVisStampa 
            Caption         =   "Stampa resoconto"
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
            Left            =   120
            TabIndex        =   15
            Top             =   6840
            Width           =   2175
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   135
            Left            =   2400
            TabIndex        =   22
            Top             =   7080
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   238
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "CONTRATTI DA RINNOVARE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   0
            Width           =   15255
         End
         Begin VB.Label lblInfoEla 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   2520
            TabIndex        =   36
            Top             =   6840
            Width           =   8175
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
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   5895
         Begin VB.CommandButton cmdStampa 
            Height          =   325
            Left            =   5400
            Picture         =   "frmContrattoRinnovo.frx":0582
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "STAMPA"
            Top             =   0
            Width           =   375
         End
         Begin DmtActBox.DmtActBoxCtl ActivityBox 
            Height          =   6975
            Left            =   120
            TabIndex        =   40
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
Attribute VB_Name = "FrmContrattiPerRinnovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsContratti As ADODB.Recordset


Private ChangeDurataContratto As Boolean
Private ChangeTipoRinnovo As Boolean
Private ChangeTipoRateizzazione As Boolean
Private ImportoNonStandard As Integer
Private Link_ContrattoNonStandard As Long
Private Var_AdeguamentoIstat As Boolean

Private oReport As dmtReportLib.dmtReport
Private IDTipoOggettoPrg As Long

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

Private Sub ElaborazioniContratti()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim DataInizioNew As String
Dim DataFineNew As String



'''''CALCOLO NUMERO CONTRATTI DA RINNOVARE''''''''''''''''''''''''''''''''''
sSQL = "SELECT COUNT(IDRV_PO08_Contratto) AS NumeroRecord "
sSQL = sSQL & "FROM RV_PO08_Contratto"
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND DataFineContratto <=" & fnNormDate(Var_ADataRinnovo)
sSQL = sSQL & " AND Chiuso=" & fnNormBoolean(0)
sSQL = sSQL & " AND Attuale=" & fnNormBoolean(1)
sSQL = sSQL & " AND Confermato=" & fnNormBoolean(1)
sSQL = sSQL & " AND DaRinnovare=" & fnNormBoolean(1)
sSQL = sSQL & " AND ContrattoRinnovato=" & fnNormBoolean(0)
sSQL = sSQL & " AND TotaleContratto>0"
sSQL = sSQL & " AND Tipo=1"

If Link_cliente_Ric > 0 Then
    sSQL = sSQL & " AND IDAnagraficaFatturazione=" & Link_cliente_Ric
End If
If Link_Tipo_Contratto_Ric > 0 Then
    sSQL = sSQL & " AND IDRV_PO08_TipoContratto=" & Link_Tipo_Contratto_Ric
End If

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

If NumeroRecord = 0 Then Exit Sub

Me.ProgressBar1.Value = 0
Unita_Progresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 4)

sSQL = "SELECT * FROM RV_PO08_IEContratto "
'sSQL = sSQL & "  DataScadenzaPerRinnovo >=" & fnNormDate(Var_DaDataRinnovo)
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND DataFineContratto <=" & fnNormDate(Var_ADataRinnovo)
sSQL = sSQL & " AND Chiuso=" & fnNormBoolean(0)
sSQL = sSQL & " AND Attuale=" & fnNormBoolean(1)
sSQL = sSQL & " AND Confermato=" & fnNormBoolean(1)
sSQL = sSQL & " AND DaRinnovare=" & fnNormBoolean(1)
sSQL = sSQL & " AND ContrattoRinnovato=" & fnNormBoolean(0)
sSQL = sSQL & " AND TotaleContratto>0 "
sSQL = sSQL & " AND Tipo=1"


If Link_cliente_Ric > 0 Then
    sSQL = sSQL & " AND IDAnagraficaFatturazione=" & Link_cliente_Ric
End If
If Link_Tipo_Contratto_Ric > 0 Then
    sSQL = sSQL & " AND IDRV_PO08_TipoContratto=" & Link_Tipo_Contratto_Ric
End If


CREA_RECORDSET
    
Set rs = CnDMT.OpenResultset(sSQL)
    
    NumeroRecordEla = 1
    
    While Not rs.EOF
        GET_DATE_RINNOVO rs!DataInizioContratto, rs!DataFineContratto, DataInizioNew, DataFineNew
        
        rsContratti.AddNew
            rsContratti!IDRV_PO08_Contratto = fnNotNullN(rs!IDRV_PO08_Contratto)
            If ((DateDiff("d", Var_DaDataRinnovo, rs!DataFineContratto) >= 0) And (DateDiff("d", Var_ADataRinnovo, rs!DataFineContratto) <= 0)) Then
            
                rsContratti!Registra = 1
            Else
                rsContratti!Registra = 0
            End If
            
            
            'rsContratti!Registra = 0 'Funzione che restituisce se è nel range di date selezionato
            rsContratti!Conferma = 0
            rsContratti!DataInizioContratto = DataInizioNew ' rs!DataInizioContratto
            rsContratti!DataFineContratto = DataFineNew 'rs!DataFineContratto
            rsContratti!NumeroRinnovo = 2 'Funzione che calcola il numero di rinnovo
            rsContratti!IDAnagraficaFatturazione = rs!IDAnagraficaFatturazione
            rsContratti!AnagraficaFatturazione = rs!AnagraficaFatturazione
           
            rsContratti!IDRV_PO08_TipoContratto = rs!IDRV_PO08_TipoContratto
            rsContratti!TipoContratto = rs!TipoContratto

            rsContratti!IDRV_PO08_TipoRateizzazione = rs!IDRV_PO08_TipoRateizzazione
            rsContratti!TipoRateizzazione = rs!TipoRateizzazione


            rsContratti!IDRV_PO08_ParcoNatanti = rs!IDRV_PO08_ParcoNatanti
            rsContratti!NomeImbarcazione = rs!NomeNatante

            rsContratti!TotaleDocumento = rs!TotaleContratto
            
        rsContratti.Update
        
        'NumeroRecord = NumeroRecord + 1
        
        If (Me.ProgressBar1.Value + Unita_Progresso) >= Me.ProgressBar1.Max Then
            Me.ProgressBar1.Value = Me.ProgressBar1.Max
        Else
            Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_Progresso
        End If
        
        Me.lblInfoEla.Caption = NumeroRecordEla & " di " & NumeroRecord
        NumeroRecordEla = NumeroRecordEla + 1
        
        DoEvents
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
End Sub
Private Sub SettaggioGrigliaContratti()
'On Error GoTo ERR_SettaggioGrigliaContratti
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
        With Me.GrigliaDMT
            .EnableMove = True
            .UpdatePosition = False
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectCell
            .ColumnsHeader.Clear
            
            Set cl = .ColumnsHeader.Add("Registra", "Da rinnovare", dgBoolean, True, 1000, dgAligncenter)
                cl.Editable = True
            Set cl = .ColumnsHeader.Add("Conferma", "Conferma", dgBoolean, True, 1000, dgAligncenter)
                cl.Editable = True
            .ColumnsHeader.Add "IDRV_PO08_Contratto", "IDRV_POContratto", dgchar, False, 500, dgAlignRight
            
            .ColumnsHeader.Add "NumeroRinnovo", "Numero rinnovo", dgInteger, True, 1000, dgAlignRight
            .ColumnsHeader.Add "IDAnagraficaFatturazione", "IDAnagraficaFatturazione", dgchar, False, 500, dgAlignRight
            .ColumnsHeader.Add "AnagraficaFatturazione", "Cliente", dgchar, True, 2500, dgAlignleft
            
            .ColumnsHeader.Add "IDRV_PO08_TipoContratto", "IDRV_PO08_TipoContratto", dgchar, False, 500, dgAlignRight
            .ColumnsHeader.Add "TipoContratto", "Tipo Contratto", dgchar, True, 2500, dgAlignleft
            
            .ColumnsHeader.Add "IDRV_PO08_ParcoNatanti", "IDRV_PO08_ParcoNatanti", dgchar, False, 500, dgAlignRight
            .ColumnsHeader.Add "NomeImbarcazione", "Imbarcazione", dgchar, True, 2500, dgAlignleft
            
            
            .ColumnsHeader.Add "DataInizioContratto", "Data inizio", dgDate, True, 2000, dgAligncenter
            .ColumnsHeader.Add "DataFineContratto", "Data fine", dgDate, True, 2000, dgAligncenter

            Set cl = .ColumnsHeader.Add("TotaleDocumento", "Importo contratto", dgDouble, True, 2000, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
                    
            Set .Recordset = rsContratti
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGrigliaContratti:
    MsgBox Err.Description, vbCritical, "SettaggioGrigliaContratti"
End Sub


Private Sub cmdAnnulla_Click()
    If MsgBox("Sei sicuro di voler chiudere l'applicazione?", vbInformation + vbYesNo, "Chiusura importazione dati") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdAvanti_Click()
Dim I As Integer
    
    
    
    If Not ((rsContratti.EOF) And (rsContratti.BOF)) Then
    
        rsContratti.Filter = "Registra=1"
        
        If Not ((rsContratti.EOF) And (rsContratti.BOF)) Then
        
            CREA_RECORDSET_FINALE
            NumeroRecordContratti = 0
            rsContratti.MoveFirst
            
            While Not rsContratti.EOF
                
                NumeroRecordContratti = NumeroRecordContratti + 1

                rsContrattiReg.AddNew
                    For I = 0 To rsContratti.Fields.Count - 1
                        rsContrattiReg(rsContratti.Fields(I).Name).Value = rsContratti.Fields(I).Value
                    Next
                rsContrattiReg.Update
            rsContratti.MoveNext
            Wend
            
            Unload Me
        End If
    End If


    
End Sub

Private Sub cmdIndietro_Click()
    Unload Me
End Sub
Private Sub cmdStampa_Click()
Dim sSQL As String

IDTipoOggettoPrg = fnGetTipoOggetto("RV_PORinnovoPreStampa")

AVVIA_PROCEDURA_STAMPA



Set oReport = New dmtReportLib.dmtReport
Set oReport.Connection = CnDMT
oReport.Password = TheApp.Password
oReport.User = TheApp.User


'Imposta l'idfiliale di appartenenza del documento da stampare
oReport.BranchID = TheApp.Branch 'IDFiliale
'Imposta l'identificativo del tipo di documento
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
    Else
        Me.Picture2.Left = Me.FraStampa.Left
        Me.FraStampa.Visible = False
        Me.ZOrder 1
        Me.Pic1.Width = 15855
    End If
    
    Form_Resize

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    ElaborazioniContratti
    SettaggioGrigliaContratti
End Sub

Private Sub Form_Load()
    With HScroll1
      .Max = (Pic1.ScaleWidth)
      .LargeChange = .Max \ 10
      .SmallChange = .Max \ 10
      
    End With

    With VScroll1
      .Max = (Pic1.ScaleHeight)
      .LargeChange = .Max \ 10
      .SmallChange = .Max \ 10
    End With


    Me.Icon = gResource.GetIcon(IDI_DIAMANTE32)

    ConfigurazioneStampe
End Sub


Private Sub Form_Resize()
  If Me.WindowState <> 1 Then
    

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
End Sub



Private Sub Form_Unload(Cancel As Integer)
If Me.cmdAvanti.Value = True Then
    FrmFine.Show
    Exit Sub
End If
If cmdAnnulla.Value = True Then
    ChiusuraConnessione
    Exit Sub
    
End If
If Me.cmdIndietro.Value = True Then
    FrmInizio.Show
    Exit Sub
End If

ChiusuraConnessione
End Sub

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
        
        oActivity.Load fnGetTipoOggetto("RV_PO08_ContrattoRinnovo"), TheApp.IDFirm
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

Private Sub GrigliaDMT_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    'If KeyAscii = vbKeySpace Then
    '    'Se non siamo in modalità filtri
    '    If Me.GrigliaDMT.GuiMode = dgNormal Then
    '    'Abilitiamo o disabilitiamo il check in base allo stato corrente
    '        sbSelectSelectedRow Not CBool(rsContratti.Fields("Registra").Value), 2
    '    End If
    'End If
    
End Sub

Private Sub GrigliaDMT_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If GrigliaDMT.HitTest(x, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If x > 0 And (x * Screen.TwipsPerPixelX) < GrigliaDMT.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If GrigliaDMT.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsContratti.Fields("Registra").Value), "Registra"
            End If
        End If
        If ((x > 0) And ((x * Screen.TwipsPerPixelX) < (GrigliaDMT.ColumnsHeader(2).Width * 2)) And ((x * Screen.TwipsPerPixelX) > GrigliaDMT.ColumnsHeader(1).Width)) Then
            'Se non siamo in modalità filtri
            If GrigliaDMT.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsContratti.Fields("Conferma").Value), "Conferma"
                If rsContratti.Fields("Registra").Value = 0 Then
                    sbSelectSelectedRow Not CBool(rsContratti.Fields("Registra").Value), "Registra"
                End If
            End If
        End If
    End If
    

End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, NomeCampo As String)
    
    If Not rsContratti.EOF And Not rsContratti.BOF Then
        rsContratti.Fields(NomeCampo).Value = Abs(CLng(Selected))
        rsContratti.UpdateBatch
        Me.GrigliaDMT.Refresh
    End If
End Sub
Private Sub CREA_RECORDSET()

If Not (rsContratti Is Nothing) Then
    Set rsContratti = Nothing
End If

Set rsContratti = New ADODB.Recordset
rsContratti.CursorLocation = adUseClient


rsContratti.Fields.Append "IDRV_PO08_Contratto", adInteger, , adFldIsNullable
rsContratti.Fields.Append "Registra", adBoolean, adFldIsNullable
rsContratti.Fields.Append "Conferma", adBoolean, adFldIsNullable
rsContratti.Fields.Append "DataInizioContratto", adDBTimeStamp, adFldIsNullable
rsContratti.Fields.Append "DataFineContratto", adDBTimeStamp, adFldIsNullable
rsContratti.Fields.Append "NumeroRinnovo", adInteger, , adFldIsNullable

rsContratti.Fields.Append "IDAnagraficaFatturazione", adInteger, , adFldIsNullable
rsContratti.Fields.Append "AnagraficaFatturazione", adVarChar, 250, adFldIsNullable

rsContratti.Fields.Append "IDRV_PO08_TipoContratto", adInteger, , adFldIsNullable
rsContratti.Fields.Append "TipoContratto", adVarChar, 250, adFldIsNullable

rsContratti.Fields.Append "IDRV_PO08_TipoRateizzazione", adInteger, , adFldIsNullable
rsContratti.Fields.Append "TipoRateizzazione", adVarChar, 250, adFldIsNullable

rsContratti.Fields.Append "TotaleDocumento", adDouble, , adFldIsNullable

rsContratti.Fields.Append "IDRV_PO08_ParcoNatanti", adInteger, , adFldIsNullable
rsContratti.Fields.Append "NomeImbarcazione", adVarChar, 250, adFldIsNullable


rsContratti.Open , , adOpenKeyset, adLockBatchOptimistic









End Sub
Private Sub CREA_RECORDSET_FINALE()
If Not (rsContrattiReg Is Nothing) Then
    Set rsContrattiReg = Nothing
End If

Set rsContrattiReg = New ADODB.Recordset
rsContrattiReg.CursorLocation = adUseClient


rsContrattiReg.Fields.Append "IDRV_PO08_Contratto", adInteger, , adFldIsNullable
rsContrattiReg.Fields.Append "Registra", adBoolean, adFldIsNullable
rsContrattiReg.Fields.Append "Conferma", adBoolean, adFldIsNullable
rsContrattiReg.Fields.Append "DataInizioContratto", adDBTimeStamp, adFldIsNullable
rsContrattiReg.Fields.Append "DataFineContratto", adDBTimeStamp, adFldIsNullable
rsContrattiReg.Fields.Append "NumeroRinnovo", adInteger, , adFldIsNullable

rsContrattiReg.Fields.Append "IDAnagraficaFatturazione", adInteger, , adFldIsNullable
rsContrattiReg.Fields.Append "AnagraficaFatturazione", adVarChar, 250, adFldIsNullable

rsContrattiReg.Fields.Append "IDRV_PO08_TipoContratto", adInteger, , adFldIsNullable
rsContrattiReg.Fields.Append "TipoContratto", adVarChar, 250, adFldIsNullable

rsContrattiReg.Fields.Append "IDRV_PO08_TipoRateizzazione", adInteger, , adFldIsNullable
rsContrattiReg.Fields.Append "TipoRateizzazione", adVarChar, 250, adFldIsNullable

rsContrattiReg.Fields.Append "TotaleDocumento", adDouble, , adFldIsNullable

rsContrattiReg.Fields.Append "IDRV_PO08_ParcoNatanti", adInteger, , adFldIsNullable
rsContrattiReg.Fields.Append "NomeImbarcazione", adVarChar, 250, adFldIsNullable


rsContrattiReg.Open , , adOpenKeyset, adLockBatchOptimistic
End Sub
Private Sub GET_DATE_RINNOVO(DataInizioContratto As String, DataFineContratto As String, DataInizioContrattoNew As String, DataFineContrattoNew As String)

    Dim AnnoInizio As String
    Dim AnnoFine As String
    
    Dim Mese As String
    Dim Giorno As String
    Dim GiorniContratto As Long
    Dim DataFine As String
    
    
    AnnoInizio = Year(DataInizioContratto)
    AnnoFine = Year(DataFineContratto)
    
    If AnnoInizio = AnnoFine Then
        AnnoInizio = Year(DataFineContratto) + 1
    Else
        AnnoInizio = Year(DataFineContratto)
    End If
    
    Mese = Month(DataInizioContratto)
    Giorno = Day(DataInizioContratto)
    GiorniContratto = DateDiff("d", DataInizioContratto, DataFineContratto)
    
    If Len(Mese) = 1 Then
        Mese = "0" & Mese
    End If
    
    If Len(Giorno) = 1 Then
        Giorno = "0" & Giorno
    End If
    
    DataInizioContrattoNew = Giorno & "/" & Mese & "/" & AnnoInizio
    DataFine = DateAdd("d", GiorniContratto, DataInizioContrattoNew)
    
    AnnoFine = Year(DataFine)
    Mese = Month(DataFineContratto)
    Giorno = Day(DataFineContratto)
    
    If Len(Mese) = 1 Then
        Mese = "0" & Mese
    End If
    
    If Len(Giorno) = 1 Then
        Giorno = "0" & Giorno
    End If
    
    DataFineContrattoNew = Giorno & "/" & Mese & "/" & AnnoFine
    
    
    Exit Sub
ERR_Command1_Click:
        MsgBox Err.Description, vbCritical, "Recupero dati delle date di rinnovo"
End Sub
Private Sub AVVIA_PROCEDURA_STAMPA()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsRighe As DmtOleDbLib.adoResultset
Dim rsRigheTMP As ADODB.Recordset
Dim DataInizioRighe As String
Dim DataFineRighe As String

    sSQL = "DELETE FROM RV_PO08_TMPContratto"
    CnDMT.Execute sSQL
    
    sSQL = "DELETE FROM RV_PO08_TMPContrattoRighe"
    CnDMT.Execute sSQL
    
    
    sSQL = "SELECT * FROM RV_PO08_TMPContratto"
    Set rs = New ADODB.Recordset
    rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    
    sSQL = "SELECT * FROM RV_PO08_TMPContrattoRighe"
    Set rsRigheTMP = New ADODB.Recordset
    rsRigheTMP.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    rsContratti.Filter = "Registra=1"
    
    If Not ((rsContratti.EOF) And (rsContratti.BOF)) Then
        
        While Not rsContratti.EOF
            rs.AddNew
                rs!IDRV_PO08_Contratto = rsContratti!IDRV_PO08_Contratto
                rs!NumeroRinnovo = rsContratti!NumeroRinnovo
                rs!DataInizioContratto = rsContratti!DataInizioContratto
                rs!DataFineContratto = rsContratti!DataFineContratto
                rs!Rinnovato = Abs(rsContratti!Registra)
                rs!Confermato = Abs(rsContratti!Conferma)
            rs.Update
            
            sSQL = "SELECT * FROM RV_PO08_ContrattoRighe "
            sSQL = sSQL & "WHERE IDRV_PO08_Contratto=" & fnNotNullN(rsContratti!IDRV_PO08_Contratto)
            
            Set rsRighe = CnDMT.OpenResultset(sSQL)
            
            While Not rsRighe.EOF
                GET_DATE_RINNOVO rsRighe!DataInizio, rsRighe!DataFine, DataInizioRighe, DataFineRighe
                rsRigheTMP.AddNew
                    rsRigheTMP!IDRV_PO08_ContrattoRighe = rsRighe!IDRV_PO08_ContrattoRighe
                    rsRigheTMP!IDRV_PO08_Contratto = rsRighe!IDRV_PO08_Contratto
                    rsRigheTMP!DataInizio = DataInizioRighe
                    rsRigheTMP!DataFine = DataFineRighe
                rsRigheTMP.Update
            rsRighe.MoveNext
            Wend
            
            rsRighe.CloseResultset
            Set rsRighe = Nothing
        rsContratti.MoveNext
        Wend
        
        
    End If

    rs.Close
    Set rs = Nothing
    
    rsRigheTMP.Close
    Set rsRigheTMP = Nothing
    


End Sub
