VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmContrattiPerRinnovo 
   Caption         =   "Rinnovi contratti (Passo 2 di 3)"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmContrattiPerRinnovo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   19080
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
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   0
      ScaleHeight     =   8625
      ScaleWidth      =   18945
      TabIndex        =   20
      Top             =   0
      Width           =   18975
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   8415
         Left            =   120
         ScaleHeight     =   8385
         ScaleWidth      =   18705
         TabIndex        =   21
         Top             =   120
         Width           =   18735
         Begin DmtGridCtl.DmtGrid GrigliaDMT 
            Height          =   7455
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   18495
            _ExtentX        =   32623
            _ExtentY        =   13150
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
               Picture         =   "FrmContrattiPerRinnovo.frx":4781A
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
            Left            =   17520
            TabIndex        =   19
            Top             =   7920
            Width           =   1095
         End
         Begin VB.CommandButton cmdIndietro 
            Caption         =   "Indietro"
            Height          =   375
            Left            =   13920
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   7920
            Width           =   1095
         End
         Begin VB.CommandButton cmdAnnulla 
            Caption         =   "Annulla"
            Height          =   375
            Left            =   15120
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   7920
            Width           =   1095
         End
         Begin VB.CommandButton Fine 
            Caption         =   "Fine"
            Enabled         =   0   'False
            Height          =   375
            Left            =   16320
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   7920
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
            Top             =   7920
            Width           =   2175
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   135
            Left            =   2400
            TabIndex        =   22
            Top             =   8160
            Width           =   11415
            _ExtentX        =   20135
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
            Width           =   18495
         End
         Begin VB.Label lblInfoEla 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   2400
            TabIndex        =   36
            Top             =   6840
            Width           =   6135
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
            Picture         =   "FrmContrattiPerRinnovo.frx":47D9C
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
Dim Importo As Double
Dim DataDecorrenza As String
Dim DataScadenza As String
Dim DataPerRinnovo As String
Dim DataFineAssistenza As String
Dim NumeroRecord As Long
Dim Unita_Progresso As Double
Dim NumeroRecordEla As Long
Dim NumeroRinnovo As Long
Dim ImportoMaggiorazioneIstat As Double
Dim IDDurataContrattoProx As Long
Dim FineContratto As Long
Dim FinePrimoContratto As Long

''''''''''''ELIMINAZIONE TABELLA TEMPORANEA'''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POTMPRinnovoContratto"
CnDMT.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''CALCOLO NUMERO CONTRATTI DA RINNOVARE''''''''''''''''''''''''''''''''''
sSQL = "SELECT COUNT(IDRV_POContratto) AS NumeroRecord "
sSQL = sSQL & "FROM RV_POContratto"
'sSQL = sSQL & "  DataScadenzaPerRinnovo >=" & fnNormDate(Var_DaDataRinnovo)
sSQL = sSQL & " WHERE DataScadenzaPerRinnovo <=" & fnNormDate(Var_ADataRinnovo)
sSQL = sSQL & " AND Disdetta=" & 0
sSQL = sSQL & " AND RinnovoAutomatico=" & 1
sSQL = sSQL & " AND ContrattoAttuale=1"
sSQL = sSQL & " AND Offerta=0"
sSQL = sSQL & " AND IDRV_POTipoImpostazioneContratto<>3"
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

'sSQL = sSQL & " AND NonFatturare=0"

If Link_cliente_Ric > 0 Then
    sSQL = sSQL & " AND IDAnagrafica=" & Link_cliente_Ric
End If
If Link_Tipo_Contratto_Ric > 0 Then
    sSQL = sSQL & " AND IDTipoContratto=" & Link_Tipo_Contratto_Ric
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

sSQL = "SELECT * FROM RV_POViewContratto"
'sSQL = sSQL & "  DataScadenzaPerRinnovo >=" & fnNormDate(Var_DaDataRinnovo)
sSQL = sSQL & " WHERE DataScadenzaPerRinnovo <=" & fnNormDate(Var_ADataRinnovo)
sSQL = sSQL & " AND Disdetta=" & 0
sSQL = sSQL & " AND RinnovoAutomatico=" & 1
sSQL = sSQL & " AND ContrattoAttuale=1"
sSQL = sSQL & " AND Offerta=0"
sSQL = sSQL & " AND IDRV_POTipoImpostazioneContratto<>3"
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
'sSQL = sSQL & " AND NonFatturare=0"

If Link_cliente_Ric > 0 Then
    sSQL = sSQL & " AND IDAnagrafica=" & Link_cliente_Ric
End If
If Link_Tipo_Contratto_Ric > 0 Then
    sSQL = sSQL & " AND IDTipoContratto=" & Link_Tipo_Contratto_Ric
End If
    
Set rs = CnDMT.OpenResultset(sSQL)
    
    NumeroRecordEla = 1
    
    While Not rs.EOF
        
        NumeroRinnovo = fnNotNullN(rs!NumeroRinnovo) + 1
        
        If fnNotNullN(rs!AnnoPrecedenteTipoRinnovo) = 0 Then
            DataDecorrenza = DateAdd("d", 1, fnNotNull(rs!DataScadenzaPerRinnovo))
        Else
            DataDecorrenza = DateAdd("yyyy", 1, fnNotNull(rs!DataDecorrenza))
        End If
        
        FinePrimoContratto = 0
        FineContratto = 0
        
        DataPerRinnovo = DateAdd("m", fnNotNullN(rs!MesiRinnovoContratto), DataDecorrenza) - 1
        DataPerRinnovo = DateAdd("d", fnNotNullN(rs!GiorniRinnovoContratto), DataPerRinnovo)
        
        DataScadenza = fnNotNull(rs!DataScadenza)
        IDDurataContrattoProx = fnNotNullN(rs!IDDurataContratto)
        
        If (DataPerRinnovo = rs!DataScadenza) Then
            FinePrimoContratto = 1
        End If
        
        If (DataPerRinnovo > rs!DataScadenza) Then
            If fnNotNullN(rs!IDDurataContrattoProssimoRinnovo) = 0 Then
                DataScadenza = DateAdd("m", fnNotNullN(rs!MesiDurataContratto), DataDecorrenza) - 1
                DataScadenza = DateAdd("d", fnNotNullN(GiorniDurataContratto), DataScadenza)
            Else
                If fnNotNullN(rs!IDDurataContrattoProssimoRinnovo) = fnNotNullN(rs!IDDurataContratto) Then
                    DataScadenza = DateAdd("m", fnNotNullN(rs!MesiDurataContratto), DataDecorrenza) - 1
                    DataScadenza = DateAdd("d", fnNotNullN(GiorniDurataContratto), DataScadenza)
                    FineContratto = 1
                Else
                    IDDurataContrattoProx = fnNotNullN(rs!IDDurataContrattoProssimoRinnovo)
                    DataScadenza = GET_DURATA_SCADENZA_PROX(fnNotNullN(rs!IDDurataContrattoProssimoRinnovo), DataDecorrenza)
                    FineContratto = 1
                    If DataScadenza = "" Then
                        IDDurataContrattoProx = fnNotNullN(rs!IDDurataContratto)
                        DataScadenza = DateAdd("m", fnNotNullN(rs!MesiDurataContratto), DataDecorrenza) - 1
                        DataScadenza = DateAdd("d", fnNotNullN(GiorniDurataContratto), DataScadenza)
                        
                    End If
                    
                End If
            End If
        End If
        
        DataFineAssistenza = DateAdd("m", fnNotNullN(rs!MesiDurataAssistenza), DataDecorrenza) - 1
        DataFineAssistenza = DateAdd("d", fnNotNullN(rs!GiorniDurataAssistenza), DataFineAssistenza)
        
        sSQL = "INSERT INTO RV_POTMPRinnovoContratto ("
        sSQL = sSQL & "IDRV_POContratto, ImportoContratto, DataDecorrenza, IDDurataContratto, "
        sSQL = sSQL & "IDTipoRinnovo, IDRateizzazione, IDPagamento, DataScadenzaContratto, "
        sSQL = sSQL & "DataScadenzaPerRinnovo, RinnovoAutomatico, AdeguamentoIstat, "
        sSQL = sSQL & "DaRinnovare, UsaAdeguamentoIstat, ImportoNonStandard, IDContrattoNonStandard,IDAzienda, IDFiliale, NonFatturare, NoteFatturazione,"
        sSQL = sSQL & "IDIstat, Percentuale, Maggiorazione, IDRV_POTipoDurataAssistenza, TipoDurataAssistenza, DataFineAssistenza, "
        sSQL = sSQL & "IDRV_POContrattoPadre, NumeroRinnovo, IDRaggruppamentoFatturato, IDRV_POTipoClassificazioneContratto, IDArticoloContratto, "
        sSQL = sSQL & "IDRV_POTipoImpostazioneContratto, TipoImpostazioneContratto, NumeroGiorniPrimaRata, GeneraRatePerProdotto, FineContratto, FinePrimoContratto"
        sSQL = sSQL & ") "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnNotNullN(rs!IDRV_POContratto) & ", "
        
        Importo = GetImportoContratto(fnNotNullN(rs!IDRV_POContratto), fnNotNullN(rs!ImportoContrattoAttuale), fnNotNullN(rs!AdeguamentoIstat), NumeroRinnovo, ImportoMaggiorazioneIstat)
        
        sSQL = sSQL & fnNormNumber(Importo) & ", "
        sSQL = sSQL & fnNormDate(DataDecorrenza) & ", "
        
        'sSQL = sSQL & fnNotNullN(rs!IDDurataContratto) & ", "
        sSQL = sSQL & fnNotNullN(IDDurataContrattoProx) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDTipoRinnovo) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDRateizzazione) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDPagamentoRate) & ", "
        sSQL = sSQL & fnNormDate(DataScadenza) & ", "

        sSQL = sSQL & fnNormDate(DataPerRinnovo) & ", "
        sSQL = sSQL & fnNotNullN(rs!RinnovoAutomatico) & ", "
        sSQL = sSQL & fnNotNullN(Var_AdeguamentoIstat) & ", "
        If DateDiff("d", rs!DataScadenzaPerRinnovo, Var_DaDataRinnovo) > 0 Then
            sSQL = sSQL & 0 & ", "
        Else
            sSQL = sSQL & 1 & ", "
        End If
        sSQL = sSQL & 1 & ", "
        sSQL = sSQL & fnNotNullN(ImportoNonStandard) & ", "
        sSQL = sSQL & fnNormNumber(Link_ContrattoNonStandard) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDAzienda) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDFiliale) & ", "
        sSQL = sSQL & fnNotNullN(rs!NonFatturare) & ", "
        sSQL = sSQL & fnNormString(rs!NoteFattura) & ", "
        sSQL = sSQL & Link_Istat & ", "
        sSQL = sSQL & fnNormNumber(PercentualeIstat) & ", "
        'sSQL = sSQL & fnNormNumber((fnNotNullN(rs!ImportoContrattoAttuale) / 100) * PercentualeIstat) & ", "
        sSQL = sSQL & fnNormNumber(ImportoMaggiorazioneIstat) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDRV_POTipoDurataAssistenza) & ", "
        sSQL = sSQL & fnNormString(Descrizione_Tipo_Assistenza) & ", "
        sSQL = sSQL & fnNormDate(DataFineAssistenza) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDRV_POContrattoPadre) & ", "
        sSQL = sSQL & NumeroRinnovo & ", "
        sSQL = sSQL & fnNotNullN(rs!IDRaggruppamentoFatturato) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDRV_POTipoClassificazioneContratto) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDArticoloContratto) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDRV_POTipoImpostazioneContratto) & ", "
        sSQL = sSQL & fnNormString(rs!TipoImpostazioneContratto) & ", "
        sSQL = sSQL & fnNotNullN(rs!NumeroGiorniPrimaRata) & ", "
        sSQL = sSQL & fnNotNullN(rs!GeneraRatePerProdotto) & ", "
        sSQL = sSQL & fnNotNullN(FineContratto) & ", "
        sSQL = sSQL & fnNotNullN(FinePrimoContratto)
        sSQL = sSQL & ")"
        
        
        CnDMT.Execute sSQL
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
    
    
    sSQL = "SELECT RV_POTMPRinnovoContratto.IDRV_POContratto, RV_POTMPRinnovoContratto.ImportoContratto, "
    sSQL = sSQL & "RV_POTMPRinnovoContratto.IDDurataContratto, RV_POTMPRinnovoContratto.IDTipoRinnovo,"
    sSQL = sSQL & "RV_POTMPRinnovoContratto.IDRateizzazione, RV_POTMPRinnovoContratto.RinnovoAutomatico,"
    sSQL = sSQL & "RV_POTMPRinnovoContratto.AdeguamentoIstat, RV_POTMPRinnovoContratto.DaRinnovare,"
    sSQL = sSQL & "RV_POTMPRinnovoContratto.UsaAdeguamentoIstat, RV_POTipoContratto.TipoContratto, Anagrafica.Anagrafica, Anagrafica.Nome,"
    sSQL = sSQL & "RV_PORateizzazione.Rateizzazione, RV_POTipoRinnovo.TipoRinnovo, RV_PODurataContratto.DurataContratto,"
    sSQL = sSQL & "Pagamento.Pagamento , RV_POTMPRinnovoContratto.IDPagamento, RV_POTMPRinnovoContratto.DataDecorrenza, "
    sSQL = sSQL & "RV_POTMPRinnovoContratto.DataScadenzaContratto, RV_POTMPRinnovoContratto.DataScadenzaPerRinnovo,"
    sSQL = sSQL & "RV_POTMPRinnovoContratto.MesiDurataContratto, RV_POTMPRinnovoContratto.ImportoNonStandard, RV_POContratto.DataScadenza AS DataScadenzaPrecedente, "
    sSQL = sSQL & "RV_POContratto.DataScadenzaPerRinnovo AS DataScadenzaPerRinnovoPrecedente, "
    sSQL = sSQL & "RV_POTMPRinnovoContratto.IDIstat, RV_POTMPRinnovoContratto.Percentuale, RV_POTMPRinnovoContratto.Maggiorazione, "
    sSQL = sSQL & "RV_POTMPRinnovoContratto.IDRV_POTipoDurataAssistenza, RV_POTMPRinnovoContratto.TipoDurataAssistenza, RV_POTMPRinnovoContratto.DataFineAssistenza, "
    sSQL = sSQL & "RV_POTMPRinnovoContratto.IDRV_POTipoImpostazioneContratto, RV_POTMPRinnovoContratto.TipoImpostazioneContratto, "
    sSQL = sSQL & "RV_POTMPRinnovoContratto.GeneraRatePerProdotto, RV_POTMPRinnovoContratto.NumeroGiorniPrimaRata, RV_POTMPRinnovoContratto.FineContratto, RV_POTMPRinnovoContratto.FinePrimoContratto "
    
    
    sSQL = sSQL & "FROM RV_POContratto INNER JOIN "
    sSQL = sSQL & "RV_POTMPRinnovoContratto ON RV_POContratto.IDRV_POContratto = RV_POTMPRinnovoContratto.IDRV_POContratto LEFT OUTER JOIN "
    sSQL = sSQL & "Pagamento ON RV_POTMPRinnovoContratto.IDPagamento = Pagamento.IDPagamento LEFT OUTER JOIN "
    sSQL = sSQL & "RV_POTipoContratto ON RV_POContratto.IDTipoContratto = RV_POTipoContratto.IDRV_POTipoContratto LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PODurataContratto ON "
    sSQL = sSQL & "RV_POTMPRinnovoContratto.IDDurataContratto = RV_PODurataContratto.IDRV_PODurataContratto LEFT OUTER JOIN "
    sSQL = sSQL & "RV_PORateizzazione ON RV_POTMPRinnovoContratto.IDRateizzazione = RV_PORateizzazione.IDRV_PORateizzazione LEFT OUTER JOIN "
    sSQL = sSQL & "RV_POTipoRinnovo ON RV_POTMPRinnovoContratto.IDTipoRinnovo = RV_POTipoRinnovo.IDRV_POTipoRinnovo LEFT OUTER JOIN "
    sSQL = sSQL & "Anagrafica ON RV_POContratto.IDAnagrafica = Anagrafica.IDAnagrafica "
    sSQL = sSQL & " ORDER BY RV_POTMPRinnovoContratto.DaRinnovare, Anagrafica.Anagrafica"
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
        Set rsContratti = New ADODB.Recordset
        rsContratti.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockBatchOptimistic
        
        With Me.GrigliaDMT
            .EnableMove = True
            .UpdatePosition = False
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectCell
            .ColumnsHeader.Clear
            
            Set cl = .ColumnsHeader.Add("DaRinnovare", "Da rinnovare", dgBoolean, True, 1000, dgAligncenter)
                cl.Editable = True
            Set cl = .ColumnsHeader.Add("FinePrimoContratto", "Fine I° contratto", dgBoolean, True, 1000, dgAligncenter)
            Set cl = .ColumnsHeader.Add("FineContratto", "Inizio II° contratto", dgBoolean, True, 1000, dgAligncenter)
            
            .ColumnsHeader.Add "IDRV_POContratto", "IDRV_POContratto", dgchar, False, 500, dgAlignRight
            .ColumnsHeader.Add "IDRV_POTipoImpostazioneContratto", "IDRV_POTipoImpostazioneContratto", dgchar, False, 500, dgAlignRight
            .ColumnsHeader.Add "TipoImpostazioneContratto", "Tipo impostazione", dgchar, True, 3000, dgAlignleft
            
            .ColumnsHeader.Add "Anagrafica", "Cliente", dgchar, True, 3000, dgAlignleft
            .ColumnsHeader.Add "TipoContratto", "Tipo contratto", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "DataScadenzaContratto", "Data scadenza", dgDate, True, 2000, dgAlignleft
            .ColumnsHeader.Add "DataDecorrenza", "Data decorrenza", dgDate, True, 2000, dgAlignleft
            .ColumnsHeader.Add "DurataContratto", "Durata contratto", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "TipoRinnovo", "Tipo rinnovo periodo", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "DataScadenzaPerRinnovo", "Rinnovo entro", dgDate, True, 2000, dgAlignleft
            .ColumnsHeader.Add "TipoDurataAssistenza", "Tipo durata assistenza", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "DataFineAssistenza", "Data fine assistenza", dgDate, True, 2000, dgAlignleft
            Set cl = .ColumnsHeader.Add("NumeroGiorniPrimaRata", "N° GG prima rata", dgDouble, True, 2000, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 0
                cl.FormatOptions.FormatNumericThousandSep = "."
            .ColumnsHeader.Add "GeneraRatePerProdotto", "Gen. rate per prod.", dgBoolean, True, 2000, dgAligncenter
            .ColumnsHeader.Add "Rateizzazione", "Tipo di rateizzazione", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "Pagamento", "Modalità di pagamento", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "Percentuale", "% Istat", dgDouble, True, 2000, dgAlignleft

            Set cl = .ColumnsHeader.Add("Maggiorazione", "Maggiorazione", dgDouble, True, 2000, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("ImportoContratto", "Importo contratto", dgCurrency, True, 2000, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."

                
            Set .Recordset = rsContratti
            
            .LoadUserSettings
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGrigliaContratti:
    MsgBox Err.Description, vbCritical, "SettaggioGrigliaContratti"
End Sub

Private Sub cboDurataAssistenza_Click()
    DurataAssistenza Me.cboDurataAssistenza.CurrentID
    
    If Me.cboDurataAssistenza.CurrentID <> fnNotNullN(Me.GrigliaDMT.AllColumns("IDRV_POTipoDurataAssistenza").Value) Then
        Me.txtDataAssistenza.Text = DateAdd("m", Mesi_Tipo_Assistenza, Me.txtDataDecorrenza.Text) - 1
        Me.txtDataAssistenza.Text = DateAdd("d", Giorni_Tipo_Assistenza, Me.txtDataAssistenza.Text)
    End If
End Sub

Private Sub cboDurataContratto_Click()
    DurataContratto Me.cboDurataContratto.CurrentID
    
    If ChangeDurataContratto = True Then
        Me.txtDataScadenza.Text = DateAdd("m", Mesi_Durata_Contratto, Me.txtDataDecorrenza.Text) - 1
        Me.txtDataScadenza.Text = DateAdd("d", Giorni_Durata_Contratto, Me.txtDataScadenza.Text)
    
    Else
        If Me.cboDurataContratto.CurrentID <> Me.GrigliaDMT.AllColumns("IDDurataContratto") Then
            Me.txtDataScadenza.Text = DateAdd("m", Mesi_Durata_Contratto, Me.txtDataDecorrenza.Text) - 1
            Me.txtDataScadenza.Text = DateAdd("d", Giorni_Durata_Contratto, Me.txtDataScadenza.Text)
            
            ChangeDurataContratto = True
        End If
    End If
    
    
End Sub




Private Sub cboRateizzazione_Click()
    TipoRateizzazione Me.cboRateizzazione.CurrentID
    If ChangeTipoRateizzazione = True Then
        'If Mesi_Rinnovo_Contratto > 0 Then
        '    If Mesi_Rate > Mesi_Rinnovo_Contratto Then
        '        MsgBox "Impossibile inserire questo tipo di rateizzazione", vbInformation, "Inserimento rinnovo contratto"
        '        Me.cboRateizzazione.WriteOn 0
        '    End If
        'End If
    Else
        'If Me.cboRateizzazione.CurrentID <> Me.GrigliaDMT.AllColumns("IDRateizzazione") Then
        '    If Mesi_Rinnovo_Contratto > 0 Then
        '        If Mesi_Rate > Mesi_Rinnovo_Contratto Then
        '            MsgBox "Impossibile inserire questo tipo di rateizzazione", vbInformation, "Inserimento rinnovo contratto"
        '            Me.cboRateizzazione.WriteOn 0
        '        End If
        '    End If
        '    ChangeTipoRateizzazione = True
        'End If
    End If

End Sub

Private Sub CboRinnovoContratto_Click()
    TipoRinnovo Me.CboRinnovoContratto.CurrentID
    If ChangeTipoRinnovo = True Then
        If Mesi_Rinnovo_Contratto > Mesi_Durata_Contratto Then
            MsgBox "Impossibile inserire questo tipo rinnovo", vbInformation, "Inserimento rinnovo contratto"
            Me.CboRinnovoContratto.WriteOn 0
        Else
            Me.txtDataScadenzaRinnovo.Text = DateAdd("m", Mesi_Rinnovo_Contratto, Me.txtDataDecorrenza.Value) - 1
            Me.txtDataScadenzaRinnovo.Text = DateAdd("d", Giorni_Rinnovo_Contratto, Me.txtDataScadenzaRinnovo.Value)
        
        End If

    Else
        If Me.CboRinnovoContratto.CurrentID <> Me.GrigliaDMT.AllColumns("IDTipoRinnovo") Then
            If Mesi_Rinnovo_Contratto > Mesi_Durata_Contratto Then
                MsgBox "Impossibile inserire questo tipo rinnovo", vbInformation, "Inserimento rinnovo contratto"
                Me.CboRinnovoContratto.WriteOn 0
            Else
                Me.txtDataScadenzaRinnovo.Text = DateAdd("m", Mesi_Rinnovo_Contratto, Me.txtDataDecorrenza.Value) - 1
                Me.txtDataScadenzaRinnovo.Text = DateAdd("d", Giorni_Rinnovo_Contratto, Me.txtDataScadenzaRinnovo.Value)
            End If
            
            ChangeTipoRinnovo = True
        End If
    End If
End Sub

Private Sub cmdAggiorna_Click()
Dim sSQL As String
Dim NumeroRecord As Long


If Mesi_Rinnovo_Contratto > Mesi_Durata_Contratto Then
    MsgBox "Il tipo di rinnovo non è congruente il tipo di durara del contratto", vbInformation, "Impossibile aggiornare"
    Me.CboRinnovoContratto.SetFocus
    Exit Sub
End If
If Mesi_Rinnovo_Contratto > 0 Then
'    If Mesi_Rate > Mesi_Rinnovo_Contratto Then
'        MsgBox "Impossibile inserire questo tipo di rateizzazione", vbInformation, "Impossibile aggiornare"
'        Me.cboRateizzazione.SetFocus
'        Exit Sub
'    End If
End If

sSQL = "UPDATE RV_POTMPRinnovoContratto SET "
sSQL = sSQL & "DataDecorrenza=" & fnNormDate(Me.txtDataDecorrenza.Text) & ", "
sSQL = sSQL & "ImportoContratto=" & fnNormNumber(Me.txtImportoContratto.Value) & ", "
sSQL = sSQL & "IDDurataContratto=" & Me.cboDurataContratto.CurrentID & ", "
sSQL = sSQL & "IDTipoRinnovo=" & Me.CboRinnovoContratto.CurrentID & ", "
sSQL = sSQL & "IDRateizzazione=" & Me.cboRateizzazione.CurrentID & ", "
sSQL = sSQL & "IDPagamento=" & Me.cboPagamento.CurrentID & ", "
sSQL = sSQL & "IDRV_POTipoDurataAssistenza=" & Me.cboDurataAssistenza.CurrentID & ", "
sSQL = sSQL & "DataFineAssistenza=" & fnNormDate(Me.txtDataAssistenza.Text) & ", "
sSQL = sSQL & "DataScadenzaContratto=" & fnNormDate(Me.txtDataScadenza.Text) & ", "
sSQL = sSQL & "DataScadenzaPerRinnovo=" & fnNormDate(Me.txtDataScadenzaRinnovo.Text) & ", "
sSQL = sSQL & "RinnovoAutomatico=" & fnNotNullN(Me.chkRinnovoAutomatico.Value) & ", "
sSQL = sSQL & "AdeguamentoIstat=" & fnNotNullN(Me.chkAdeguamentoIstat.Value) & ", "
sSQL = sSQL & "DaRinnovare=" & fnNotNullN(Me.chkContrattoDaRinnovare.Value) & ", "
sSQL = sSQL & "UsaAdeguamentoIstat=" & fnNotNullN(Me.chkUsaAdeguamentoIstat.Value) & " "
sSQL = sSQL & "WHERE IDRV_POContratto=" & Me.GrigliaDMT.AllColumns("IDRV_POContratto")

CnDMT.Execute sSQL
NumeroRecord = Me.GrigliaDMT.ListIndex - 1

SettaggioGrigliaContratti

If Not (Me.GrigliaDMT.Recordset.BOF And Me.GrigliaDMT.Recordset.EOF) Then
    Me.GrigliaDMT.Recordset.Move NumeroRecord
End If
End Sub

Private Sub cmdAnnulla_Click()
    If MsgBox("Sei sicuro di voler chiudere l'applicazione?", vbInformation + vbYesNo, "Chiusura importazione dati") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdAvanti_Click()
    Unload Me
End Sub

Private Sub cmdIndietro_Click()
    Unload Me
End Sub
Private Sub cmdStampa_Click()
Dim sSQL As String

IDTipoOggettoPrg = fnGetTipoOggetto("RV_PORinnovoPreStampa")

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
        Me.Pic1.Width = 18975
    End If
    
    Form_Resize

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
    
    fncTipoDurataContratto
    fncTipoRinnovo
    fncTipoRateizzazione
    fncPagamento
    fncTipoDurataAssistenza
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
Private Function GetImportoContratto(IDContratto As Long, ImportoContratto As Double, AdeguamentoIstatPrecedente As Boolean, NumeroRinnovoOLD As Long, ImportoAdeguamentoIstat As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsAdeg As DmtOleDbLib.adoResultset
Dim ImportoAdeguamento As Double
Dim ImportoContrattoLocal As Double


    ImportoAdeguamentoIstat = 0

    sSQL = "SELECT IDRV_POContrattoNonStandard, Importo, AdeguamentoIstat, PercentualeSconto, ScontoAImporto "
    sSQL = sSQL & "FROM RV_POContrattoNonStandard "
    sSQL = sSQL & "WHERE IDRV_POContratto = " & IDContratto
    sSQL = sSQL & " AND NumeroRinnovo=" & NumeroRinnovoOLD
    

    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        ImportoNonStandard = 1
        
        Link_ContrattoNonStandard = fnNotNullN(rs!IDRV_POContrattoNonStandard)
        
        Var_AdeguamentoIstat = fnNotNullN(rs!AdeguamentoIstat)
        
        'ImportoContrattoLocal = fnNotNullN(rs!Importo)
        
        ImportoContrattoLocal = ImportoContratto
        
        ImportoContrattoLocal = ImportoContrattoLocal - ((ImportoContrattoLocal / 100) * fnNotNullN(rs!PercentualeSconto))
        ImportoContrattoLocal = ImportoContrattoLocal + fnNotNullN(rs!ScontoAImporto)

    Else
        ImportoContrattoLocal = ImportoContratto '+ ImportoAdeguamento
        ImportoNonStandard = 0
        Link_ContrattoNonStandard = 0
        
        Var_AdeguamentoIstat = AdeguamentoIstatPrecedente
    
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    '''ADEGUAMENTI DA ACCORPARE'''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT Sum(Importo) AS TotaleImporto "
    sSQL = sSQL & "FROM RV_POContrattoAdeguamento "
    sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
    sSQL = sSQL & " AND RiportaProssimoRinnovo=1 "
    sSQL = sSQL & " AND IDRV_POTipoAdeguamento=1 "
    
    Set rsAdeg = CnDMT.OpenResultset(sSQL)
    
    If rsAdeg.EOF Then
        ImportoAdeguamento = 0
    Else
        ImportoAdeguamento = fnNotNullN(rsAdeg!TotaleImporto)
    End If
    
    rsAdeg.CloseResultset
    Set rsAdeg = Nothing
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ImportoContrattoLocal = ImportoContrattoLocal + ImportoAdeguamento
    
    If Var_AdeguamentoIstat = True Then
        ImportoAdeguamentoIstat = (ImportoContrattoLocal / 100) * PercentualeIstat
        GetImportoContratto = ImportoContrattoLocal + ((ImportoContrattoLocal / 100) * PercentualeIstat)
    Else
        GetImportoContratto = ImportoContrattoLocal
    End If
    
    GetImportoContratto = FormatNumber(GetImportoContratto, 2)
    
End Function



Private Sub GrigliaDMT_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
            
    ChangeDurataContratto = False
    ChangeTipoRinnovo = False
    ChangeTipoRateizzazione = False
            
    Me.txtDataDecorrenza.Text = Me.GrigliaDMT.AllColumns("DataDecorrenza")
    Me.cboDurataContratto.WriteOn Me.GrigliaDMT.AllColumns("IDDurataContratto")
    Me.txtDataScadenza.Text = Me.GrigliaDMT.AllColumns("DataScadenzaContratto")
    Me.CboRinnovoContratto.WriteOn Me.GrigliaDMT.AllColumns("IDTipoRinnovo")
    Me.txtDataScadenzaRinnovo.Text = Me.GrigliaDMT.AllColumns("DataScadenzaPerRinnovo")
    Me.cboRateizzazione.WriteOn Me.GrigliaDMT.AllColumns("IDRateizzazione")
    Me.cboPagamento.WriteOn Me.GrigliaDMT.AllColumns("IDPagamento")
    Me.chkAdeguamentoIstat.Value = Abs(Me.GrigliaDMT.AllColumns("AdeguamentoIstat"))
    Me.chkContrattoDaRinnovare.Value = Abs(Me.GrigliaDMT.AllColumns("DaRinnovare"))
    Me.chkRinnovoAutomatico.Value = Abs(Me.GrigliaDMT.AllColumns("RinnovoAutomatico"))
    Me.chkUsaAdeguamentoIstat.Value = Abs(Me.GrigliaDMT.AllColumns("UsaAdeguamentoIstat"))
    Me.txtImportoContratto.Value = Me.GrigliaDMT.AllColumns("ImportoContratto")
    Me.cboDurataAssistenza.WriteOn fnNotNullN(Me.GrigliaDMT.AllColumns("IDRV_POTipoDurataAssistenza").Value)
    Me.txtDataAssistenza.Value = fnNotNullN(Me.GrigliaDMT.AllColumns("DataFineAssistenza").Value)
    
    If DateDiff("d", Me.GrigliaDMT.AllColumns("DataScadenzaPerRinnovoPrecedente"), Me.GrigliaDMT.AllColumns("DataScadenzaPrecedente")) > 0 Then
        Me.cboDurataContratto.Enabled = False
        Me.txtDataScadenza.Enabled = False
    Else
        Me.cboDurataContratto.Enabled = True
        Me.txtDataScadenza.Enabled = True
    End If
    If Me.GrigliaDMT.AllColumns("ImportoNonStandard") = True Then
        Me.Image1.Visible = True
        Me.lblInfo.Caption = "L'importo visualizzato è un importo non standard"
    Else
        Me.Image1.Visible = False
        Me.lblInfo.Caption = ""
    End If
    
End Sub
Private Sub fncTipoDurataContratto()
    
    With Me.cboDurataContratto
        Set .Database = CnDMT
        .AddFieldKey "IDRV_PODurataContratto"
        .DisplayField = "DurataContratto"
        .Sql = "SELECT * FROM RV_PODurataContratto"
        .Fill
    End With

End Sub
Private Sub fncTipoRinnovo()
    With Me.CboRinnovoContratto
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoRinnovo"
        .DisplayField = "TipoRinnovo"
        .Sql = "SELECT * FROM RV_POTipoRinnovo"
        .Fill
    End With
    
    
End Sub
Private Sub fncTipoRateizzazione()
    With Me.cboRateizzazione
        Set .Database = CnDMT
        .AddFieldKey "IDRV_PORateizzazione"
        .DisplayField = "Rateizzazione"
        .Sql = "SELECT * FROM RV_PORateizzazione"
        .Fill
    End With

End Sub
Private Sub fncTipoDurataAssistenza()
    
    With Me.cboDurataAssistenza
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoDurataAssistenza"
        .DisplayField = "TipoDurataAssistenza"
        .Sql = "SELECT * FROM RV_POTipoDurataAssistenza"
        .Fill
    End With
    
End Sub
Private Sub fncPagamento()
    With Me.cboPagamento
        Set .Database = CnDMT
        .AddFieldKey "IDPagamento"
        .DisplayField = "Pagamento"
        .Sql = "SELECT * FROM Pagamento"
        .Fill
    End With

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
        
        oActivity.Load fnGetTipoOggetto("RV_PORinnovoPreStampa"), TheApp.IDFirm
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
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GrigliaDMT.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsContratti.Fields("DaRinnovare").Value), 2
        End If
    End If
End Sub

Private Sub GrigliaDMT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If GrigliaDMT.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < GrigliaDMT.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If GrigliaDMT.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsContratti.Fields("DaRinnovare").Value), 2
            End If
        End If
    End If
    

End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    
    If Not rsContratti.EOF And Not rsContratti.BOF Then
        rsContratti.Fields("DaRinnovare").Value = Abs(CLng(Selected))
        rsContratti.UpdateBatch
        Me.GrigliaDMT.Refresh
    End If
End Sub

Private Function GET_DURATA_SCADENZA_PROX(IDDurataContratto As Long, DataPartenza As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_DURATA_SCADENZA_PROX = ""

sSQL = "SELECT * FROM RV_PODurataContratto "
sSQL = sSQL & "WHERE IDRV_PODurataContratto=" & IDDurataContratto

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_DURATA_SCADENZA_PROX = DateAdd("m", fnNotNullN(rs!Mesi), DataPartenza) - 1
    GET_DURATA_SCADENZA_PROX = DateAdd("d", fnNotNullN(rs!Giorni), GET_DURATA_SCADENZA_PROX)
End If


End Function

