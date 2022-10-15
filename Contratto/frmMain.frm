VERSION 5.00
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Object = "{7A1D73E4-F461-11D0-8F01-004033A00AF2}#1.0#0"; "DmtWheel.ocx"
Object = "{5C67DC8E-40E7-11D3-AF44-00105A2FBE61}#3.0#0"; "DmtPrnDlgCtl.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{9385BB2E-6637-11D1-850D-002018802E11}#3.1#0"; "Dmtsplit.ocx"
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{E1215E52-40E1-11D3-AF44-00105A2FBE61}#5.1#0"; "DMTLblLinkCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   13800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   25365
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   13800
   ScaleWidth      =   25365
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin ActiveBar3LibraryCtl.ActiveBar3 BarMenu 
      Height          =   13455
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   25365
      _LayoutVersion  =   2
      _ExtentX        =   44741
      _ExtentY        =   23733
      _DataPath       =   ""
      Bands           =   "frmMain.frx":4781A
      Begin DMTSPLIT.DMTSplitBar DMTSplitBar1 
         Height          =   510
         Left            =   10920
         TabIndex        =   137
         Top             =   120
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
      End
      Begin VB.PictureBox PicForm 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   13395
         Left            =   0
         ScaleHeight     =   13365
         ScaleWidth      =   25245
         TabIndex        =   54
         Top             =   0
         Width           =   25275
         Begin VB.PictureBox PicForm2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   13095
            Left            =   120
            ScaleHeight     =   13065
            ScaleWidth      =   24945
            TabIndex        =   55
            Top             =   120
            Width           =   24975
            Begin VB.Frame FraTesta 
               Caption         =   "Intestazione"
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
               Height          =   5175
               Left            =   120
               TabIndex        =   138
               Top             =   0
               Width           =   11655
               Begin VB.CheckBox chkFineContratto 
                  Caption         =   "Fine contratto"
                  Height          =   195
                  Left            =   2760
                  TabIndex        =   300
                  TabStop         =   0   'False
                  Top             =   4800
                  Width           =   2415
               End
               Begin VB.CheckBox chkAttivoPassivo 
                  Caption         =   "Contratto cliente"
                  Height          =   195
                  Left            =   2760
                  TabIndex        =   295
                  TabStop         =   0   'False
                  Top             =   4800
                  Visible         =   0   'False
                  Width           =   2415
               End
               Begin VB.CheckBox chkContrattoAttuale 
                  Caption         =   "Contratto attuale"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   294
                  TabStop         =   0   'False
                  Top             =   4200
                  Width           =   2775
               End
               Begin VB.CheckBox chkNonFatturare 
                  Caption         =   "Non fatturare"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   293
                  TabStop         =   0   'False
                  Top             =   3960
                  Width           =   2535
               End
               Begin VB.CheckBox chkAdeguamentoIstat 
                  Caption         =   "Adeguamento I.S.T.A.T."
                  Height          =   195
                  Left            =   120
                  TabIndex        =   292
                  TabStop         =   0   'False
                  Top             =   3405
                  Width           =   2415
               End
               Begin VB.CheckBox chkRinnovoAutomatico 
                  Caption         =   "Rinnovo automatico"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   291
                  TabStop         =   0   'False
                  Top             =   3675
                  Width           =   2535
               End
               Begin VB.CheckBox chkRitAcconto 
                  Caption         =   "Ritenuta di acconto"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   290
                  TabStop         =   0   'False
                  Top             =   4245
                  Width           =   2535
               End
               Begin VB.CheckBox chkFatturazioneRic 
                  Caption         =   "Fatturazione ricorrente"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   120
                  TabIndex        =   289
                  TabStop         =   0   'False
                  Top             =   4515
                  Width           =   2535
               End
               Begin VB.CheckBox chkTotDaiProdotti 
                  Caption         =   "Totale contratti dai prodotti"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   288
                  TabStop         =   0   'False
                  Top             =   3360
                  Width           =   2775
               End
               Begin VB.CheckBox chkChiuso 
                  Caption         =   "Chiuso"
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
                  Left            =   2760
                  TabIndex        =   287
                  TabStop         =   0   'False
                  Top             =   3915
                  Width           =   2775
               End
               Begin VB.CheckBox chkGeneraRateProd 
                  Caption         =   "Genera rate per prodotto"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   286
                  TabStop         =   0   'False
                  Top             =   3615
                  Width           =   2775
               End
               Begin VB.CheckBox chkOfferta 
                  Caption         =   "Offerta"
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   285
                  TabStop         =   0   'False
                  Top             =   4485
                  Width           =   2775
               End
               Begin VB.CheckBox chkGeneraAccontiSaldo 
                  Caption         =   "Gestione Acconti/Saldo"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   120
                  TabIndex        =   284
                  TabStop         =   0   'False
                  Top             =   4800
                  Visible         =   0   'False
                  Width           =   2655
               End
               Begin DMTLblLinkCtl.LabelLink LabelLink1 
                  Height          =   135
                  Left            =   4800
                  TabIndex        =   270
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   238
                  Name            =   "LabelLink"
               End
               Begin VB.Frame fraAltriDati 
                  Caption         =   "Altri dati"
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
                  Height          =   3495
                  Left            =   5880
                  TabIndex        =   238
                  Top             =   1560
                  Width           =   5655
                  Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
                     Height          =   3135
                     Left            =   120
                     TabIndex        =   248
                     Top             =   240
                     Width           =   5415
                     _ExtentX        =   9551
                     _ExtentY        =   5530
                     _Version        =   393216
                     FixedRows       =   0
                     FixedCols       =   0
                     RowHeightMin    =   300
                     BackColor       =   -2147483633
                     ForeColorFixed  =   -2147483640
                     BackColorSel    =   -2147483633
                     ForeColorSel    =   0
                     BackColorBkg    =   -2147483633
                     ScrollTrack     =   -1  'True
                     GridLinesFixed  =   1
                     SelectionMode   =   1
                     AllowUserResizing=   3
                     BorderStyle     =   0
                     Appearance      =   0
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
                  Begin VB.TextBox txtAltriDati 
                     Appearance      =   0  'Flat
                     BackColor       =   &H8000000F&
                     Height          =   195
                     Left            =   120
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   242
                     Top             =   360
                     Width           =   5415
                  End
               End
               Begin VB.CommandButton Command3 
                  Height          =   300
                  Left            =   5280
                  Picture         =   "frmMain.frx":479EA
                  Style           =   1  'Graphical
                  TabIndex        =   183
                  ToolTipText     =   "Recapiti telefonici dell'amministratore"
                  Top             =   2880
                  Width           =   375
               End
               Begin VB.CommandButton Command2 
                  Height          =   300
                  Left            =   5280
                  Picture         =   "frmMain.frx":47F74
                  Style           =   1  'Graphical
                  TabIndex        =   182
                  ToolTipText     =   "Recapiti telefonici dell'agente"
                  Top             =   2310
                  Width           =   375
               End
               Begin DMTDataCmb.DMTCombo cboSitoPerAnagrafica 
                  Height          =   315
                  Left            =   5760
                  TabIndex        =   0
                  Top             =   570
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
               Begin DmtCodDescCtl.DmtCodDesc CDAgente 
                  Height          =   615
                  Left            =   120
                  TabIndex        =   14
                  Top             =   2085
                  Width           =   5415
                  _ExtentX        =   9551
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":484FE
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":4854E
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":4859E
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
               Begin VB.CommandButton Command5 
                  Height          =   280
                  Left            =   7680
                  Picture         =   "frmMain.frx":485F8
                  Style           =   1  'Graphical
                  TabIndex        =   184
                  ToolTipText     =   "Recapiti telefonici del cliente"
                  Top             =   300
                  Width           =   375
               End
               Begin VB.CommandButton cmdAnaContratto 
                  Height          =   280
                  Left            =   5280
                  Picture         =   "frmMain.frx":48B82
                  Style           =   1  'Graphical
                  TabIndex        =   180
                  ToolTipText     =   "Recapiti telefonici del cliente"
                  Top             =   300
                  Width           =   375
               End
               Begin VB.CommandButton Command1 
                  Height          =   300
                  Left            =   5280
                  Picture         =   "frmMain.frx":4910C
                  Style           =   1  'Graphical
                  TabIndex        =   181
                  ToolTipText     =   "Recapiti telefonici del tecnico"
                  Top             =   1740
                  Width           =   375
               End
               Begin VB.TextBox txtDescrizioneContratto 
                  Height          =   315
                  Left            =   6120
                  TabIndex        =   4
                  Top             =   1120
                  Width           =   5415
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDContrattoPadre 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   19
                  Top             =   5865
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
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
               Begin DMTDataCmb.DMTCombo cboTipoContratto 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   1
                  Top             =   1125
                  Width           =   4215
                  _ExtentX        =   7435
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
               Begin DMTDataCmb.DMTCombo cboTipoImpostazione 
                  Height          =   315
                  Left            =   8160
                  TabIndex        =   204
                  Top             =   570
                  Width           =   3375
                  _ExtentX        =   5953
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
               Begin DmtCodDescCtl.DmtCodDesc CDTecnico 
                  Height          =   615
                  Left            =   120
                  TabIndex        =   227
                  Top             =   1500
                  Width           =   5415
                  _ExtentX        =   9551
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":49696
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":496F6
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":49746
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
               Begin DmtCodDescCtl.DmtCodDesc CDAmministratore 
                  Height          =   615
                  Left            =   120
                  TabIndex        =   228
                  Top             =   2640
                  Width           =   5415
                  _ExtentX        =   9551
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":497A0
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":497F8
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":49848
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
               Begin DmtCodDescCtl.DmtCodDesc CDCliente 
                  Height          =   615
                  Left            =   120
                  TabIndex        =   296
                  Top             =   315
                  Width           =   5535
                  _ExtentX        =   9763
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":498A2
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":49903
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":49953
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
               Begin DMTDATETIMELib.dmtDate txtDataStipula 
                  Height          =   315
                  Left            =   4440
                  TabIndex        =   2
                  Top             =   1125
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin VB.Label Label1 
                  Caption         =   "Data stipula"
                  Height          =   255
                  Index           =   14
                  Left            =   4440
                  TabIndex        =   297
                  Top             =   900
                  Width           =   1455
               End
               Begin VB.Label Label1 
                  Caption         =   "Tipo impostazione"
                  Height          =   255
                  Index           =   40
                  Left            =   8160
                  TabIndex        =   205
                  Top             =   360
                  Width           =   2535
               End
               Begin VB.Label Label1 
                  Caption         =   "Descrizione aggiuntiva contratto"
                  Height          =   255
                  Index           =   10
                  Left            =   6120
                  TabIndex        =   154
                  Top             =   900
                  Width           =   3375
               End
               Begin VB.Label Label4 
                  Caption         =   "Identificativo padre"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   153
                  Top             =   5655
                  Width           =   1695
               End
               Begin VB.Label Label1 
                  Caption         =   "Tipo contratto"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   140
                  Top             =   900
                  Width           =   2775
               End
               Begin VB.Label Label1 
                  Caption         =   "Altra sede"
                  Height          =   255
                  Index           =   6
                  Left            =   5760
                  TabIndex        =   139
                  Top             =   345
                  Width           =   2415
               End
            End
            Begin VB.Frame FraDate 
               Caption         =   "Proprietà"
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
               Height          =   5175
               Left            =   11760
               TabIndex        =   141
               Top             =   0
               Width           =   9015
               Begin VB.TextBox txtNumeroProtocollo 
                  Height          =   315
                  Left            =   3840
                  TabIndex        =   7
                  Top             =   1080
                  Width           =   4935
               End
               Begin VB.Frame fraRateizzazione 
                  Caption         =   "Rateizzazione"
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
                  Height          =   2415
                  Left            =   5160
                  TabIndex        =   249
                  Top             =   1400
                  Width           =   3615
                  Begin DMTDataCmb.DMTCombo cboTipoRateizzazione 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   250
                     Top             =   240
                     Width           =   3375
                     _ExtentX        =   5953
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
                  Begin DMTEDITNUMLib.dmtNumber txtNGGPrimaRata 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   251
                     Top             =   825
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
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
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataPrimaRata 
                     Height          =   315
                     Left            =   1560
                     TabIndex        =   252
                     Top             =   825
                     Width           =   1935
                     _Version        =   65536
                     _ExtentX        =   3413
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDataCmb.DMTCombo cboPagamentoRate 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   253
                     Top             =   1380
                     Width           =   3375
                     _ExtentX        =   5953
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
                  Begin DMTDataCmb.DMTCombo cboTipoRateizzazioneProx 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   306
                     Top             =   1875
                     Width           =   3375
                     _ExtentX        =   5953
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
                  Begin VB.Label Label1 
                     Caption         =   "Rateizzazione al II° contratto"
                     Height          =   255
                     Index           =   39
                     Left            =   120
                     TabIndex        =   307
                     Top             =   1680
                     Width           =   2895
                  End
                  Begin VB.Label Label1 
                     Caption         =   "N° GG"
                     Height          =   255
                     Index           =   41
                     Left            =   120
                     TabIndex        =   256
                     ToolTipText     =   "Numero giorni prima rata"
                     Top             =   615
                     Width           =   735
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Data prima rata"
                     Height          =   255
                     Index           =   42
                     Left            =   1560
                     TabIndex        =   255
                     ToolTipText     =   "Numero giorni prima rata"
                     Top             =   600
                     Width           =   1455
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Modalità pagamento delle rate"
                     Height          =   255
                     Index           =   5
                     Left            =   120
                     TabIndex        =   254
                     Top             =   1140
                     Width           =   3015
                  End
               End
               Begin DMTEDITNUMLib.dmtCurrency txtImportoStipula 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   5
                  Top             =   480
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   " 0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTDataCmb.DMTCombo cboDurataContratto 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   8
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
               Begin DMTDATETIMELib.dmtDate txtDataDecorrenza 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   6
                  Top             =   480
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboTipoRinnovo 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   10
                  Top             =   2250
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
               Begin DMTDATETIMELib.dmtDate txtDataScadenza 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   9
                  Top             =   1680
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataScadenzaPerRinnovo 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   11
                  Top             =   2250
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboDurataAssistenza 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   12
                  Top             =   2850
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
               Begin DMTDATETIMELib.dmtDate txtDataFineAssistenza 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   13
                  Top             =   2850
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtCurrency txtImportoAttuale 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   236
                  Top             =   1080
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   " 0"
                  ForeColor       =   0
                  BackColor       =   65535
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtNumber txtNumeroRinnovo 
                  Height          =   315
                  Left            =   7680
                  TabIndex        =   243
                  Top             =   480
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   556
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
               Begin DMTEDITNUMLib.dmtNumber txtAnnoContratto 
                  Height          =   315
                  Left            =   5160
                  TabIndex        =   244
                  Top             =   480
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
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
               Begin DMTEDITNUMLib.dmtNumber txtNumeroContratto 
                  Height          =   315
                  Left            =   6360
                  TabIndex        =   245
                  Top             =   480
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
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
               Begin DMTEDITNUMLib.dmtCurrency txtImportoTotAdeg 
                  Height          =   315
                  Left            =   1920
                  TabIndex        =   258
                  Top             =   1080
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   " 0"
                  ForeColor       =   0
                  BackColor       =   65535
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
               Begin DMTDataCmb.DMTCombo cboDurataContrattoProx 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   277
                  Top             =   3400
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
               Begin DMTDATETIMELib.dmtDate txtDataScadSecContr 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   298
                  Top             =   3400
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataPrimaDecorr 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   3
                  Top             =   480
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboTipoRinnovoProx 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   304
                  Top             =   3915
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
               Begin DMTDATETIMELib.dmtDate txtDataScadPerRinnovoProx 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   308
                  Top             =   3915
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTLblLinkCtl.LabelLink lblLinkFattContratto 
                  Height          =   135
                  Left            =   0
                  TabIndex        =   311
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   238
                  Caption         =   "Fatturazione contratti"
                  Name            =   "LabelLink"
               End
               Begin VB.Label Label1 
                  Caption         =   "Rinn. entro II° contr."
                  Height          =   255
                  Index           =   43
                  Left            =   3240
                  TabIndex        =   309
                  Top             =   3720
                  Width           =   1815
               End
               Begin VB.Label Label1 
                  Caption         =   "Tipo rinnovo periodo II° contratto"
                  Height          =   255
                  Index           =   34
                  Left            =   120
                  TabIndex        =   305
                  Top             =   3720
                  Width           =   2895
               End
               Begin VB.Label Label1 
                  Caption         =   "Scad. II° contratto"
                  Height          =   255
                  Index           =   15
                  Left            =   3240
                  TabIndex        =   299
                  Top             =   3200
                  Width           =   1815
               End
               Begin VB.Label Label1 
                  Caption         =   "Durata II° contratto"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   278
                  Top             =   3200
                  Width           =   2895
               End
               Begin VB.Label Label1 
                  Caption         =   "Tot. contr. + Adeg."
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
                  Index           =   35
                  Left            =   1920
                  TabIndex        =   259
                  Top             =   840
                  Width           =   1695
               End
               Begin VB.Label Label1 
                  Caption         =   "N° Rinnovo"
                  Height          =   255
                  Index           =   31
                  Left            =   7680
                  TabIndex        =   247
                  Top             =   280
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Numero contratto"
                  Height          =   255
                  Index           =   38
                  Left            =   5160
                  TabIndex        =   246
                  Top             =   280
                  Width           =   1695
               End
               Begin VB.Label Label1 
                  Caption         =   "Importo contratto"
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
                  Index           =   9
                  Left            =   120
                  MouseIcon       =   "frmMain.frx":499AD
                  MousePointer    =   99  'Custom
                  TabIndex        =   237
                  ToolTipText     =   "Inserimento importi non standard"
                  Top             =   840
                  Width           =   1695
               End
               Begin VB.Label Label1 
                  Caption         =   "N° Protocollo"
                  Height          =   255
                  Index           =   16
                  Left            =   3840
                  TabIndex        =   151
                  Top             =   840
                  Width           =   4815
               End
               Begin VB.Label Label1 
                  Caption         =   "Rinnovo entro"
                  Height          =   255
                  Index           =   13
                  Left            =   3240
                  TabIndex        =   150
                  Top             =   2040
                  Width           =   1815
               End
               Begin VB.Label Label1 
                  Caption         =   "Durata assistenza"
                  Height          =   255
                  Index           =   18
                  Left            =   120
                  TabIndex        =   149
                  Top             =   2640
                  Width           =   1695
               End
               Begin VB.Label Label1 
                  Caption         =   "Data di fine ass."
                  Height          =   255
                  Index           =   19
                  Left            =   3240
                  TabIndex        =   148
                  Top             =   2640
                  Width           =   1455
               End
               Begin VB.Label Label1 
                  Caption         =   "Scad. I° contratto"
                  Height          =   255
                  Index           =   7
                  Left            =   3240
                  TabIndex        =   147
                  Top             =   1485
                  Width           =   1815
               End
               Begin VB.Label Label1 
                  Caption         =   "Tipo rinnovo periodo"
                  Height          =   255
                  Index           =   12
                  Left            =   120
                  TabIndex        =   146
                  Top             =   2040
                  Width           =   1935
               End
               Begin VB.Label Label1 
                  Caption         =   "Decorrenza attuale"
                  Height          =   255
                  Index           =   2
                  Left            =   3240
                  TabIndex        =   145
                  Top             =   285
                  Width           =   1815
               End
               Begin VB.Label Label1 
                  Caption         =   "Durata I° contratto"
                  Height          =   255
                  Index           =   11
                  Left            =   120
                  TabIndex        =   144
                  Top             =   1485
                  Width           =   2895
               End
               Begin VB.Label Label1 
                  Caption         =   "Prima decorrenza"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   143
                  Top             =   280
                  Width           =   1575
               End
               Begin VB.Label Label1 
                  Caption         =   "Importo iniziale"
                  Height          =   255
                  Index           =   8
                  Left            =   1800
                  TabIndex        =   142
                  Top             =   280
                  Width           =   1455
               End
            End
            Begin VB.Frame FraAnnotazioni 
               Caption         =   "Annotazioni"
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
               Height          =   1815
               Left            =   20760
               TabIndex        =   260
               Top             =   3360
               Width           =   4095
               Begin VB.TextBox txtAnnotazioni 
                  Height          =   1455
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   261
                  Top             =   240
                  Width           =   3855
               End
            End
            Begin TabDlg.SSTab SSTab1 
               Height          =   7695
               Left            =   120
               TabIndex        =   58
               Top             =   5280
               Width           =   24735
               _ExtentX        =   43630
               _ExtentY        =   13573
               _Version        =   393216
               Tabs            =   5
               Tab             =   3
               TabsPerRow      =   5
               TabHeight       =   520
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Rateizzazione"
               TabPicture(0)   =   "frmMain.frx":49CB7
               Tab(0).ControlEnabled=   0   'False
               Tab(0).Control(0)=   "LabelLink3"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "cmdEliminaRif"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "Command8"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "txtIDProdottoRata"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).Control(4)=   "cmdSelProdContratto"
               Tab(0).Control(4).Enabled=   0   'False
               Tab(0).Control(5)=   "txtProdottoRata"
               Tab(0).Control(5).Enabled=   0   'False
               Tab(0).Control(6)=   "txtNoteNonFattRata"
               Tab(0).Control(6).Enabled=   0   'False
               Tab(0).Control(7)=   "chkNonFatturareRata"
               Tab(0).Control(7).Enabled=   0   'False
               Tab(0).Control(8)=   "txtIDOggettoRata"
               Tab(0).Control(8).Enabled=   0   'False
               Tab(0).Control(9)=   "txtIDTipoOggettoRata"
               Tab(0).Control(9).Enabled=   0   'False
               Tab(0).Control(10)=   "cmdTrovaFattura"
               Tab(0).Control(10).Enabled=   0   'False
               Tab(0).Control(11)=   "txtOggettoCollegato"
               Tab(0).Control(11).Enabled=   0   'False
               Tab(0).Control(12)=   "txtPeriodo"
               Tab(0).Control(12).Enabled=   0   'False
               Tab(0).Control(13)=   "chkRataFatturata"
               Tab(0).Control(13).Enabled=   0   'False
               Tab(0).Control(14)=   "cmdEliminaRata"
               Tab(0).Control(14).Enabled=   0   'False
               Tab(0).Control(15)=   "cmdSalvaRata"
               Tab(0).Control(15).Enabled=   0   'False
               Tab(0).Control(16)=   "cmdNuovaRata"
               Tab(0).Control(16).Enabled=   0   'False
               Tab(0).Control(17)=   "txtImportoRata"
               Tab(0).Control(17).Enabled=   0   'False
               Tab(0).Control(18)=   "cboPagamentoRataContratto"
               Tab(0).Control(18).Enabled=   0   'False
               Tab(0).Control(19)=   "txtDataRata"
               Tab(0).Control(19).Enabled=   0   'False
               Tab(0).Control(20)=   "GrigliaRateContratto"
               Tab(0).Control(20).Enabled=   0   'False
               Tab(0).Control(21)=   "cboIVARateContratto"
               Tab(0).Control(21).Enabled=   0   'False
               Tab(0).Control(22)=   "txtDataInizioPer"
               Tab(0).Control(22).Enabled=   0   'False
               Tab(0).Control(23)=   "txtDataFinePer"
               Tab(0).Control(23).Enabled=   0   'False
               Tab(0).Control(24)=   "txtNumeroRata"
               Tab(0).Control(24).Enabled=   0   'False
               Tab(0).Control(25)=   "txtIDOggettoCollegato"
               Tab(0).Control(25).Enabled=   0   'False
               Tab(0).Control(26)=   "txtIDProdRifContr"
               Tab(0).Control(26).Enabled=   0   'False
               Tab(0).Control(27)=   "Label2(12)"
               Tab(0).Control(27).Enabled=   0   'False
               Tab(0).Control(28)=   "Label13(1)"
               Tab(0).Control(28).Enabled=   0   'False
               Tab(0).Control(29)=   "Label2(11)"
               Tab(0).Control(29).Enabled=   0   'False
               Tab(0).Control(30)=   "Label2(9)"
               Tab(0).Control(30).Enabled=   0   'False
               Tab(0).Control(31)=   "Label2(8)"
               Tab(0).Control(31).Enabled=   0   'False
               Tab(0).Control(32)=   "Label13(0)"
               Tab(0).Control(32).Enabled=   0   'False
               Tab(0).Control(33)=   "Label2(4)"
               Tab(0).Control(33).Enabled=   0   'False
               Tab(0).Control(34)=   "Label2(3)"
               Tab(0).Control(34).Enabled=   0   'False
               Tab(0).Control(35)=   "Label2(2)"
               Tab(0).Control(35).Enabled=   0   'False
               Tab(0).Control(36)=   "Label2(1)"
               Tab(0).Control(36).Enabled=   0   'False
               Tab(0).Control(37)=   "Label2(0)"
               Tab(0).Control(37).Enabled=   0   'False
               Tab(0).ControlCount=   38
               TabCaption(1)   =   "Ricorrenze Servizi"
               TabPicture(1)   =   "frmMain.frx":49CD3
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Label1(20)"
               Tab(1).Control(0).Enabled=   0   'False
               Tab(1).Control(1)=   "Label1(21)"
               Tab(1).Control(1).Enabled=   0   'False
               Tab(1).Control(2)=   "Label1(22)"
               Tab(1).Control(2).Enabled=   0   'False
               Tab(1).Control(3)=   "Label1(23)"
               Tab(1).Control(3).Enabled=   0   'False
               Tab(1).Control(4)=   "Label1(24)"
               Tab(1).Control(4).Enabled=   0   'False
               Tab(1).Control(5)=   "Label1(25)"
               Tab(1).Control(5).Enabled=   0   'False
               Tab(1).Control(6)=   "Label1(26)"
               Tab(1).Control(6).Enabled=   0   'False
               Tab(1).Control(7)=   "Label1(27)"
               Tab(1).Control(7).Enabled=   0   'False
               Tab(1).Control(8)=   "Label1(28)"
               Tab(1).Control(8).Enabled=   0   'False
               Tab(1).Control(9)=   "Label1(29)"
               Tab(1).Control(9).Enabled=   0   'False
               Tab(1).Control(10)=   "Label1(30)"
               Tab(1).Control(10).Enabled=   0   'False
               Tab(1).Control(11)=   "Label1(17)"
               Tab(1).Control(11).Enabled=   0   'False
               Tab(1).Control(12)=   "Label1(33)"
               Tab(1).Control(12).Enabled=   0   'False
               Tab(1).Control(13)=   "cboTipoAnnoFineRicorr"
               Tab(1).Control(13).Enabled=   0   'False
               Tab(1).Control(14)=   "cboTipoAnnoInizioRicorr"
               Tab(1).Control(14).Enabled=   0   'False
               Tab(1).Control(15)=   "CDServizio"
               Tab(1).Control(15).Enabled=   0   'False
               Tab(1).Control(16)=   "txtNumeroRicorrenze"
               Tab(1).Control(16).Enabled=   0   'False
               Tab(1).Control(17)=   "txtMeseFissoFineRic"
               Tab(1).Control(17).Enabled=   0   'False
               Tab(1).Control(18)=   "txtGiornoFissoFineRic"
               Tab(1).Control(18).Enabled=   0   'False
               Tab(1).Control(19)=   "cboTipoDataFineRic"
               Tab(1).Control(19).Enabled=   0   'False
               Tab(1).Control(20)=   "txtMeseFissoInizioRic"
               Tab(1).Control(20).Enabled=   0   'False
               Tab(1).Control(21)=   "txtGiornoFissoInizioRic"
               Tab(1).Control(21).Enabled=   0   'False
               Tab(1).Control(22)=   "cboTipoDataInizioRic"
               Tab(1).Control(22).Enabled=   0   'False
               Tab(1).Control(23)=   "txtOgniNumeroSettimane"
               Tab(1).Control(23).Enabled=   0   'False
               Tab(1).Control(24)=   "txtOgniNumeroMesi"
               Tab(1).Control(24).Enabled=   0   'False
               Tab(1).Control(25)=   "cboCriterioRicorrenza"
               Tab(1).Control(25).Enabled=   0   'False
               Tab(1).Control(26)=   "txtOgniNumeroGiorni"
               Tab(1).Control(26).Enabled=   0   'False
               Tab(1).Control(27)=   "GrigliaServizi"
               Tab(1).Control(27).Enabled=   0   'False
               Tab(1).Control(28)=   "cmdNuovoServizio"
               Tab(1).Control(28).Enabled=   0   'False
               Tab(1).Control(29)=   "cmdSalvaServizio"
               Tab(1).Control(29).Enabled=   0   'False
               Tab(1).Control(30)=   "cmdEliminaServizio"
               Tab(1).Control(30).Enabled=   0   'False
               Tab(1).Control(31)=   "cmdConfiguraServizi"
               Tab(1).Control(31).Enabled=   0   'False
               Tab(1).Control(32)=   "cmdGeneraInterventi"
               Tab(1).Control(32).Enabled=   0   'False
               Tab(1).Control(33)=   "cmdGeneraIntSing"
               Tab(1).Control(33).Enabled=   0   'False
               Tab(1).Control(34)=   "fraRiepilogoServizio"
               Tab(1).Control(34).Enabled=   0   'False
               Tab(1).Control(35)=   "Command9"
               Tab(1).Control(35).Enabled=   0   'False
               Tab(1).ControlCount=   36
               TabCaption(2)   =   "Interventi"
               TabPicture(2)   =   "frmMain.frx":49CEF
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "GrigliaInterventi"
               Tab(2).Control(0).Enabled=   0   'False
               Tab(2).ControlCount=   1
               TabCaption(3)   =   "Adeguamenti contrattuali"
               TabPicture(3)   =   "frmMain.frx":49D0B
               Tab(3).ControlEnabled=   -1  'True
               Tab(3).Control(0)=   "Label9"
               Tab(3).Control(0).Enabled=   0   'False
               Tab(3).Control(1)=   "Label10"
               Tab(3).Control(1).Enabled=   0   'False
               Tab(3).Control(2)=   "Label11(0)"
               Tab(3).Control(2).Enabled=   0   'False
               Tab(3).Control(3)=   "Label12(0)"
               Tab(3).Control(3).Enabled=   0   'False
               Tab(3).Control(4)=   "Label2(5)"
               Tab(3).Control(4).Enabled=   0   'False
               Tab(3).Control(5)=   "Label2(6)"
               Tab(3).Control(5).Enabled=   0   'False
               Tab(3).Control(6)=   "Label2(7)"
               Tab(3).Control(6).Enabled=   0   'False
               Tab(3).Control(7)=   "Label25(0)"
               Tab(3).Control(7).Enabled=   0   'False
               Tab(3).Control(8)=   "Label2(10)"
               Tab(3).Control(8).Enabled=   0   'False
               Tab(3).Control(9)=   "Label28"
               Tab(3).Control(9).Enabled=   0   'False
               Tab(3).Control(10)=   "Label12(1)"
               Tab(3).Control(10).Enabled=   0   'False
               Tab(3).Control(11)=   "Label18(11)"
               Tab(3).Control(11).Enabled=   0   'False
               Tab(3).Control(12)=   "Label6"
               Tab(3).Control(12).Enabled=   0   'False
               Tab(3).Control(13)=   "Label11(1)"
               Tab(3).Control(13).Enabled=   0   'False
               Tab(3).Control(14)=   "Label25(1)"
               Tab(3).Control(14).Enabled=   0   'False
               Tab(3).Control(15)=   "txtNAdegIniz"
               Tab(3).Control(15).Enabled=   0   'False
               Tab(3).Control(16)=   "txtImportoAdegRinn"
               Tab(3).Control(16).Enabled=   0   'False
               Tab(3).Control(17)=   "txtDataScadenzaAdeg"
               Tab(3).Control(17).Enabled=   0   'False
               Tab(3).Control(18)=   "cboTipoRateizzazioneAdeg"
               Tab(3).Control(18).Enabled=   0   'False
               Tab(3).Control(19)=   "cboIstatAdeg"
               Tab(3).Control(19).Enabled=   0   'False
               Tab(3).Control(20)=   "CDArticoloAdeg"
               Tab(3).Control(20).Enabled=   0   'False
               Tab(3).Control(21)=   "cboTipoAdeguamento"
               Tab(3).Control(21).Enabled=   0   'False
               Tab(3).Control(22)=   "cboIvaAdeg"
               Tab(3).Control(22).Enabled=   0   'False
               Tab(3).Control(23)=   "txtDataDecorrenzaAdeg"
               Tab(3).Control(23).Enabled=   0   'False
               Tab(3).Control(24)=   "GrigliaAdeg"
               Tab(3).Control(24).Enabled=   0   'False
               Tab(3).Control(25)=   "cmdEliminaAdeg"
               Tab(3).Control(25).Enabled=   0   'False
               Tab(3).Control(26)=   "cmdSalvaAdeg"
               Tab(3).Control(26).Enabled=   0   'False
               Tab(3).Control(27)=   "cmdNuovoAdeg"
               Tab(3).Control(27).Enabled=   0   'False
               Tab(3).Control(28)=   "txtDataStipulaAdeg"
               Tab(3).Control(28).Enabled=   0   'False
               Tab(3).Control(29)=   "txtImportoAdeg"
               Tab(3).Control(29).Enabled=   0   'False
               Tab(3).Control(30)=   "txtAnnotazioniAdeg"
               Tab(3).Control(30).Enabled=   0   'False
               Tab(3).Control(31)=   "txtProtAdeg"
               Tab(3).Control(31).Enabled=   0   'False
               Tab(3).Control(32)=   "FraImpAdeg"
               Tab(3).Control(32).Enabled=   0   'False
               Tab(3).Control(33)=   "txtNumeroAdeguamento"
               Tab(3).Control(33).Enabled=   0   'False
               Tab(3).Control(34)=   "txtMaggiorazioneAdeg"
               Tab(3).Control(34).Enabled=   0   'False
               Tab(3).Control(35)=   "txtDescrFattAde"
               Tab(3).Control(35).Enabled=   0   'False
               Tab(3).ControlCount=   36
               TabCaption(4)   =   "Prodotti/Addebiti"
               TabPicture(4)   =   "frmMain.frx":49D27
               Tab(4).ControlEnabled=   0   'False
               Tab(4).Control(0)=   "Label14(0)"
               Tab(4).Control(0).Enabled=   0   'False
               Tab(4).Control(1)=   "Label15"
               Tab(4).Control(1).Enabled=   0   'False
               Tab(4).Control(2)=   "Label16"
               Tab(4).Control(2).Enabled=   0   'False
               Tab(4).Control(3)=   "Label18(0)"
               Tab(4).Control(3).Enabled=   0   'False
               Tab(4).Control(4)=   "Label22"
               Tab(4).Control(4).Enabled=   0   'False
               Tab(4).Control(5)=   "Label23(0)"
               Tab(4).Control(5).Enabled=   0   'False
               Tab(4).Control(6)=   "Label18(1)"
               Tab(4).Control(6).Enabled=   0   'False
               Tab(4).Control(7)=   "Label3"
               Tab(4).Control(7).Enabled=   0   'False
               Tab(4).Control(8)=   "Label21"
               Tab(4).Control(8).Enabled=   0   'False
               Tab(4).Control(9)=   "Label20"
               Tab(4).Control(9).Enabled=   0   'False
               Tab(4).Control(10)=   "Label19"
               Tab(4).Control(10).Enabled=   0   'False
               Tab(4).Control(11)=   "Label17"
               Tab(4).Control(11).Enabled=   0   'False
               Tab(4).Control(12)=   "Label30(0)"
               Tab(4).Control(12).Enabled=   0   'False
               Tab(4).Control(13)=   "Label18(2)"
               Tab(4).Control(13).Enabled=   0   'False
               Tab(4).Control(14)=   "Label18(3)"
               Tab(4).Control(14).Enabled=   0   'False
               Tab(4).Control(15)=   "Label18(4)"
               Tab(4).Control(15).Enabled=   0   'False
               Tab(4).Control(16)=   "Label18(5)"
               Tab(4).Control(16).Enabled=   0   'False
               Tab(4).Control(17)=   "Label18(6)"
               Tab(4).Control(17).Enabled=   0   'False
               Tab(4).Control(18)=   "Label18(7)"
               Tab(4).Control(18).Enabled=   0   'False
               Tab(4).Control(19)=   "Label18(8)"
               Tab(4).Control(19).Enabled=   0   'False
               Tab(4).Control(20)=   "Label30(1)"
               Tab(4).Control(20).Enabled=   0   'False
               Tab(4).Control(21)=   "Line5"
               Tab(4).Control(21).Enabled=   0   'False
               Tab(4).Control(22)=   "Label18(9)"
               Tab(4).Control(22).Enabled=   0   'False
               Tab(4).Control(23)=   "Label18(10)"
               Tab(4).Control(23).Enabled=   0   'False
               Tab(4).Control(24)=   "cboTipoRateizzazioneProd"
               Tab(4).Control(24).Enabled=   0   'False
               Tab(4).Control(25)=   "cboAnaOperatoreProd"
               Tab(4).Control(25).Enabled=   0   'False
               Tab(4).Control(26)=   "txtQuantitaEffettiva"
               Tab(4).Control(26).Enabled=   0   'False
               Tab(4).Control(27)=   "LblLinkRil"
               Tab(4).Control(27).Enabled=   0   'False
               Tab(4).Control(28)=   "txtImportoIvaProd"
               Tab(4).Control(28).Enabled=   0   'False
               Tab(4).Control(29)=   "txtAliquotaIvaProd"
               Tab(4).Control(29).Enabled=   0   'False
               Tab(4).Control(30)=   "cboIvaProd"
               Tab(4).Control(30).Enabled=   0   'False
               Tab(4).Control(31)=   "cboTipoPeriodo"
               Tab(4).Control(31).Enabled=   0   'False
               Tab(4).Control(32)=   "txtOraFineProd"
               Tab(4).Control(32).Enabled=   0   'False
               Tab(4).Control(33)=   "cboUMArtProd"
               Tab(4).Control(33).Enabled=   0   'False
               Tab(4).Control(34)=   "txtQtaPeriodo"
               Tab(4).Control(34).Enabled=   0   'False
               Tab(4).Control(35)=   "txtTotaleRigaProd"
               Tab(4).Control(35).Enabled=   0   'False
               Tab(4).Control(36)=   "cboUMPeriodoProd"
               Tab(4).Control(36).Enabled=   0   'False
               Tab(4).Control(37)=   "txtImponibileProd"
               Tab(4).Control(37).Enabled=   0   'False
               Tab(4).Control(38)=   "txtScontoImpProd"
               Tab(4).Control(38).Enabled=   0   'False
               Tab(4).Control(39)=   "txtSconto2Prod"
               Tab(4).Control(39).Enabled=   0   'False
               Tab(4).Control(40)=   "txtSconto1Prod"
               Tab(4).Control(40).Enabled=   0   'False
               Tab(4).Control(41)=   "txtImpUniProd"
               Tab(4).Control(41).Enabled=   0   'False
               Tab(4).Control(42)=   "txtQtaArtProd"
               Tab(4).Control(42).Enabled=   0   'False
               Tab(4).Control(43)=   "LabelLink2"
               Tab(4).Control(43).Enabled=   0   'False
               Tab(4).Control(44)=   "cboListinoProd"
               Tab(4).Control(44).Enabled=   0   'False
               Tab(4).Control(45)=   "txtDataFineProd"
               Tab(4).Control(45).Enabled=   0   'False
               Tab(4).Control(46)=   "txtDataInizioProd"
               Tab(4).Control(46).Enabled=   0   'False
               Tab(4).Control(47)=   "CDArticoloProd"
               Tab(4).Control(47).Enabled=   0   'False
               Tab(4).Control(48)=   "cmdNuovoProd"
               Tab(4).Control(48).Enabled=   0   'False
               Tab(4).Control(49)=   "cmdSalvaProd"
               Tab(4).Control(49).Enabled=   0   'False
               Tab(4).Control(50)=   "cmdEliminaProd"
               Tab(4).Control(50).Enabled=   0   'False
               Tab(4).Control(51)=   "txtValIdentProd"
               Tab(4).Control(51).Enabled=   0   'False
               Tab(4).Control(52)=   "txtDescrProdotto"
               Tab(4).Control(52).Enabled=   0   'False
               Tab(4).Control(53)=   "txtNoteProdotto"
               Tab(4).Control(53).Enabled=   0   'False
               Tab(4).Control(54)=   "txtQtaProdotto"
               Tab(4).Control(54).Enabled=   0   'False
               Tab(4).Control(55)=   "txtProdotto"
               Tab(4).Control(55).Enabled=   0   'False
               Tab(4).Control(56)=   "cmdSelServizio"
               Tab(4).Control(56).Enabled=   0   'False
               Tab(4).Control(57)=   "txtIDProdotto"
               Tab(4).Control(57).Enabled=   0   'False
               Tab(4).Control(58)=   "Frame2"
               Tab(4).Control(58).Enabled=   0   'False
               Tab(4).Control(59)=   "chkProdottoGenerico"
               Tab(4).Control(59).Enabled=   0   'False
               Tab(4).Control(60)=   "cmdNuovoProdotto"
               Tab(4).Control(60).Enabled=   0   'False
               Tab(4).Control(61)=   "GrigliaProd"
               Tab(4).Control(61).Enabled=   0   'False
               Tab(4).Control(62)=   "txtOraInizioProd"
               Tab(4).Control(62).Enabled=   0   'False
               Tab(4).Control(63)=   "Command7"
               Tab(4).Control(63).Enabled=   0   'False
               Tab(4).Control(64)=   "cmdAttivaContProd"
               Tab(4).Control(64).Enabled=   0   'False
               Tab(4).Control(65)=   "cmdContatori"
               Tab(4).Control(65).Enabled=   0   'False
               Tab(4).Control(66)=   "cmdInfoAggArtProd"
               Tab(4).Control(66).Enabled=   0   'False
               Tab(4).Control(67)=   "chkEscludiGiorniFestivi"
               Tab(4).Control(67).Enabled=   0   'False
               Tab(4).Control(68)=   "chkEscludiSabato"
               Tab(4).Control(68).Enabled=   0   'False
               Tab(4).Control(69)=   "chkConducente"
               Tab(4).Control(69).Enabled=   0   'False
               Tab(4).Control(70)=   "chkACorpo"
               Tab(4).Control(70).Enabled=   0   'False
               Tab(4).Control(71)=   "fraTotaliProdotti"
               Tab(4).Control(71).Enabled=   0   'False
               Tab(4).Control(72)=   "cmdStampaProdotti"
               Tab(4).Control(72).Enabled=   0   'False
               Tab(4).Control(73)=   "cmdgenIntDaProd"
               Tab(4).Control(73).Enabled=   0   'False
               Tab(4).Control(74)=   "txtNoteProdInt"
               Tab(4).Control(74).Enabled=   0   'False
               Tab(4).Control(75)=   "cmdGeneraScadenzeProd"
               Tab(4).Control(75).Enabled=   0   'False
               Tab(4).Control(76)=   "chkRinnovareProd"
               Tab(4).Control(76).Enabled=   0   'False
               Tab(4).Control(77)=   "chkGeneraUnaRataProd"
               Tab(4).Control(77).Enabled=   0   'False
               Tab(4).Control(78)=   "cmdNuovoProdDaEsis"
               Tab(4).Control(78).Enabled=   0   'False
               Tab(4).ControlCount=   79
               Begin VB.CommandButton cmdNuovoProdDaEsis 
                  Caption         =   "Copia come nuovo"
                  Height          =   495
                  Left            =   -52200
                  TabIndex        =   310
                  Top             =   6000
                  Width           =   1815
               End
               Begin VB.CheckBox chkGeneraUnaRataProd 
                  Caption         =   "Genera una sola rata"
                  Height          =   315
                  Left            =   -55920
                  TabIndex        =   281
                  Top             =   2040
                  Width           =   3495
               End
               Begin VB.CheckBox chkRinnovareProd 
                  Caption         =   "Non rinnovare"
                  Height          =   315
                  Left            =   -55920
                  TabIndex        =   280
                  Top             =   1800
                  Width           =   3615
               End
               Begin VB.CommandButton cmdGeneraScadenzeProd 
                  Caption         =   "Genera scadenze"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   735
                  Left            =   -72000
                  TabIndex        =   279
                  Top             =   6840
                  Width           =   2655
               End
               Begin VB.TextBox txtNoteProdInt 
                  Height          =   285
                  Left            =   -54000
                  TabIndex        =   275
                  Top             =   2400
                  Visible         =   0   'False
                  Width           =   1575
               End
               Begin VB.CommandButton Command9 
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
                  Left            =   -74880
                  Picture         =   "frmMain.frx":49D43
                  Style           =   1  'Graphical
                  TabIndex        =   274
                  ToolTipText     =   "Generazione interventi di tutti i servizi del contratto"
                  Top             =   1440
                  Width           =   615
               End
               Begin VB.CommandButton cmdgenIntDaProd 
                  Caption         =   "Genera interventi"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   735
                  Left            =   -74880
                  TabIndex        =   273
                  Top             =   6840
                  Width           =   2655
               End
               Begin DMTLblLinkCtl.LabelLink LabelLink3 
                  Height          =   135
                  Left            =   -68880
                  TabIndex        =   272
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   2535
                  _ExtentX        =   4471
                  _ExtentY        =   238
                  Name            =   "LabelLink"
               End
               Begin VB.CommandButton cmdEliminaRif 
                  Height          =   285
                  Left            =   -65400
                  Picture         =   "frmMain.frx":4A2CD
                  Style           =   1  'Graphical
                  TabIndex        =   29
                  ToolTipText     =   "Trova documento da collegare"
                  Top             =   1440
                  Width           =   375
               End
               Begin VB.CommandButton Command8 
                  Height          =   285
                  Left            =   -65760
                  Picture         =   "frmMain.frx":4A857
                  Style           =   1  'Graphical
                  TabIndex        =   271
                  ToolTipText     =   "Vai al documento"
                  Top             =   1440
                  Width           =   375
               End
               Begin VB.CommandButton cmdStampaProdotti 
                  Caption         =   "Stampa da modelli"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   735
                  Left            =   -69120
                  TabIndex        =   269
                  TabStop         =   0   'False
                  Top             =   6840
                  Width           =   2655
               End
               Begin VB.Frame fraTotaliProdotti 
                  Caption         =   "Riepilogo totali"
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
                  Height          =   855
                  Left            =   -57600
                  TabIndex        =   262
                  Top             =   6720
                  Width           =   5295
                  Begin DMTEDITNUMLib.dmtNumber txtImponibileProdTot 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   263
                     Top             =   480
                     Width           =   1575
                     _Version        =   65536
                     _ExtentX        =   2778
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648447
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
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtTotaleRigaProdTot 
                     Height          =   315
                     Left            =   3480
                     TabIndex        =   264
                     Top             =   480
                     Width           =   1695
                     _Version        =   65536
                     _ExtentX        =   2990
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648447
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
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtImportoIvaProdTot 
                     Height          =   315
                     Left            =   1800
                     TabIndex        =   265
                     Top             =   480
                     Width           =   1575
                     _Version        =   65536
                     _ExtentX        =   2778
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   12648447
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
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin VB.Label Label5 
                     Caption         =   "Imponibile"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   268
                     Top             =   240
                     Width           =   1575
                  End
                  Begin VB.Label Label30 
                     Caption         =   "Totale"
                     Height          =   255
                     Index           =   3
                     Left            =   3480
                     TabIndex        =   267
                     Top             =   240
                     Width           =   1575
                  End
                  Begin VB.Label Label30 
                     Caption         =   "Totale I.V.A."
                     Height          =   255
                     Index           =   2
                     Left            =   1800
                     TabIndex        =   266
                     Top             =   240
                     Width           =   1575
                  End
               End
               Begin VB.CheckBox chkACorpo 
                  Caption         =   "A corpo"
                  Height          =   315
                  Left            =   -55920
                  TabIndex        =   235
                  Top             =   1560
                  Width           =   1695
               End
               Begin VB.CheckBox chkConducente 
                  Caption         =   "Conducente"
                  Height          =   315
                  Left            =   -55920
                  TabIndex        =   234
                  Top             =   600
                  Width           =   1455
               End
               Begin VB.CheckBox chkEscludiSabato 
                  Caption         =   "Escludi sabato"
                  Height          =   315
                  Left            =   -55920
                  TabIndex        =   233
                  Top             =   1320
                  Width           =   1575
               End
               Begin VB.CheckBox chkEscludiGiorniFestivi 
                  Caption         =   "Escludi domenica e festivi"
                  Height          =   315
                  Left            =   -55920
                  TabIndex        =   232
                  Top             =   1080
                  Width           =   2775
               End
               Begin VB.CommandButton cmdInfoAggArtProd 
                  Height          =   315
                  Left            =   -74880
                  Picture         =   "frmMain.frx":4ADE1
                  Style           =   1  'Graphical
                  TabIndex        =   231
                  ToolTipText     =   "Descrizioni aggiuntiva"
                  Top             =   1320
                  Width           =   375
               End
               Begin VB.CommandButton cmdContatori 
                  Height          =   315
                  Left            =   -69360
                  Picture         =   "frmMain.frx":4B36B
                  Style           =   1  'Graphical
                  TabIndex        =   229
                  ToolTipText     =   "Rilevamenti contatori"
                  Top             =   720
                  Width           =   375
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDProdottoRata 
                  Height          =   255
                  Left            =   -63840
                  TabIndex        =   225
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   735
                  _Version        =   65536
                  _ExtentX        =   1296
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin VB.CommandButton cmdAttivaContProd 
                  Height          =   315
                  Left            =   -69720
                  Picture         =   "frmMain.frx":4B8F5
                  Style           =   1  'Graphical
                  TabIndex        =   224
                  ToolTipText     =   "Attiva contatori"
                  Top             =   720
                  Width           =   375
               End
               Begin VB.CommandButton cmdSelProdContratto 
                  Height          =   315
                  Left            =   -57360
                  Picture         =   "frmMain.frx":4BE7F
                  Style           =   1  'Graphical
                  TabIndex        =   220
                  ToolTipText     =   "Seleziona prodotto"
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   375
               End
               Begin VB.CommandButton Command7 
                  Caption         =   "Nuovo addebito"
                  Height          =   495
                  Left            =   -52200
                  TabIndex        =   129
                  Top             =   3480
                  Width           =   1815
               End
               Begin DMTDATETIMELib.dmtTime txtOraInizioProd 
                  Height          =   315
                  Left            =   -63960
                  TabIndex        =   111
                  Top             =   1320
                  Width           =   735
                  _Version        =   65536
                  _ExtentX        =   1296
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DmtGridCtl.DmtGrid GrigliaProd 
                  Height          =   4215
                  Left            =   -74880
                  TabIndex        =   131
                  TabStop         =   0   'False
                  Top             =   2400
                  Width           =   22575
                  _ExtentX        =   39820
                  _ExtentY        =   7435
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
                  UseUserSettings =   0   'False
                  ColumnsHeaderHeight=   20
               End
               Begin VB.TextBox txtProdottoRata 
                  Height          =   315
                  Left            =   -64800
                  Locked          =   -1  'True
                  TabIndex        =   206
                  Top             =   840
                  Width           =   7455
               End
               Begin VB.Frame fraRiepilogoServizio 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   1575
                  Left            =   -54840
                  TabIndex        =   195
                  Top             =   360
                  Width           =   2775
                  Begin VB.CommandButton Command6 
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
                     Left            =   2760
                     Picture         =   "frmMain.frx":4C409
                     Style           =   1  'Graphical
                     TabIndex        =   202
                     ToolTipText     =   "Associati prodotti al servizio"
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   615
                  End
                  Begin VB.CommandButton cmdInterventiServizio 
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
                     Left            =   2040
                     Picture         =   "frmMain.frx":4C993
                     Style           =   1  'Graphical
                     TabIndex        =   199
                     ToolTipText     =   "Interventi del servizio"
                     Top             =   1080
                     Visible         =   0   'False
                     Width           =   615
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtNProdAss 
                     Height          =   375
                     Left            =   120
                     TabIndex        =   197
                     Top             =   360
                     Width           =   1815
                     _Version        =   65536
                     _ExtentX        =   3201
                     _ExtentY        =   661
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   12
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
                  Begin VB.CommandButton cmdProdottiServizio 
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
                     Left            =   2040
                     Picture         =   "frmMain.frx":4CF1D
                     Style           =   1  'Graphical
                     TabIndex        =   196
                     ToolTipText     =   "Prodotti associati al prodotto"
                     Top             =   360
                     Width           =   615
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtNInterventiServ 
                     Height          =   375
                     Left            =   120
                     TabIndex        =   200
                     Top             =   1080
                     Width           =   1815
                     _Version        =   65536
                     _ExtentX        =   3201
                     _ExtentY        =   661
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   12
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
                  Begin VB.Label Label1 
                     Caption         =   "N° interventi"
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
                     Index           =   37
                     Left            =   120
                     TabIndex        =   201
                     Top             =   840
                     Width           =   1815
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Prodotti associati"
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
                     Index           =   36
                     Left            =   120
                     TabIndex        =   198
                     Top             =   120
                     Width           =   2055
                  End
               End
               Begin VB.CommandButton cmdNuovoProdotto 
                  Height          =   315
                  Left            =   -74880
                  Picture         =   "frmMain.frx":4D4A7
                  Style           =   1  'Graphical
                  TabIndex        =   100
                  ToolTipText     =   "Vai al prodotto..."
                  Top             =   720
                  Width           =   375
               End
               Begin VB.CheckBox chkProdottoGenerico 
                  Caption         =   "Prodotto generico"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   -64080
                  TabIndex        =   105
                  Top             =   720
                  Width           =   1815
               End
               Begin VB.Frame Frame2 
                  Caption         =   "Dismesso"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   1815
                  Left            =   -52200
                  TabIndex        =   193
                  Top             =   480
                  Width           =   1815
                  Begin DMTDATETIMELib.dmtDate txtDataDismesso 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   133
                     Top             =   840
                     Width           =   1575
                     _Version        =   65536
                     _ExtentX        =   2778
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin VB.CheckBox chkDismessoProd 
                     Caption         =   "Dismesso"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   132
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.Label Label23 
                     Caption         =   "Data dismissione"
                     Height          =   255
                     Index           =   1
                     Left            =   120
                     TabIndex        =   194
                     Top             =   600
                     Width           =   1575
                  End
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDProdotto 
                  Height          =   255
                  Left            =   -70560
                  TabIndex        =   192
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin VB.CommandButton cmdSelServizio 
                  Height          =   315
                  Left            =   -70080
                  Picture         =   "frmMain.frx":4DA31
                  Style           =   1  'Graphical
                  TabIndex        =   102
                  ToolTipText     =   "Seleziona prodotto"
                  Top             =   720
                  Width           =   375
               End
               Begin VB.TextBox txtProdotto 
                  Height          =   315
                  Left            =   -74520
                  Locked          =   -1  'True
                  TabIndex        =   101
                  Top             =   720
                  Width           =   4455
               End
               Begin VB.TextBox txtNoteNonFattRata 
                  ForeColor       =   &H00000000&
                  Height          =   385
                  Left            =   -73320
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   190
                  Top             =   1980
                  Width           =   8280
               End
               Begin VB.CheckBox chkNonFatturareRata 
                  Caption         =   "Non fatturare"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   189
                  Top             =   1980
                  Width           =   1455
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDOggettoRata 
                  Height          =   255
                  Left            =   -57840
                  TabIndex        =   186
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDTipoOggettoRata 
                  Height          =   255
                  Left            =   -57840
                  TabIndex        =   185
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtQtaProdotto 
                  Height          =   315
                  Left            =   -65280
                  TabIndex        =   104
                  Top             =   720
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin VB.CommandButton cmdGeneraIntSing 
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
                  Left            =   -74880
                  Picture         =   "frmMain.frx":4DFBB
                  Style           =   1  'Graphical
                  TabIndex        =   179
                  ToolTipText     =   "Generazione interventi del servizio selezionato"
                  Top             =   960
                  Width           =   615
               End
               Begin VB.TextBox txtDescrFattAde 
                  Height          =   1095
                  Left            =   4800
                  MaxLength       =   250
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   92
                  Top             =   1320
                  Width           =   5055
               End
               Begin DMTEDITNUMLib.dmtNumber txtMaggiorazioneAdeg 
                  Height          =   315
                  Left            =   22920
                  TabIndex        =   90
                  Top             =   720
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtNumeroAdeguamento 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   78
                  Top             =   720
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   556
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
               Begin VB.TextBox txtNoteProdotto 
                  Height          =   735
                  Left            =   -74520
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   135
                  Top             =   3000
                  Width           =   1575
               End
               Begin VB.TextBox txtDescrProdotto 
                  Height          =   315
                  Left            =   -62160
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   106
                  Top             =   720
                  Width           =   6015
               End
               Begin VB.TextBox txtValIdentProd 
                  Height          =   315
                  Left            =   -68880
                  TabIndex        =   103
                  Top             =   720
                  Width           =   3495
               End
               Begin VB.CommandButton cmdEliminaProd 
                  Caption         =   "Elimina"
                  Height          =   495
                  Left            =   -52200
                  TabIndex        =   130
                  TabStop         =   0   'False
                  Top             =   5160
                  Width           =   1815
               End
               Begin VB.CommandButton cmdSalvaProd 
                  Caption         =   "Salva"
                  Height          =   495
                  Left            =   -52200
                  TabIndex        =   127
                  Top             =   4320
                  Width           =   1815
               End
               Begin VB.CommandButton cmdNuovoProd 
                  Caption         =   "Nuovo prodotto"
                  Height          =   495
                  Left            =   -52200
                  TabIndex        =   128
                  Top             =   2640
                  Width           =   1815
               End
               Begin VB.Frame FraImpAdeg 
                  Caption         =   "Impostazioni"
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
                  Height          =   1215
                  Left            =   9960
                  TabIndex        =   166
                  Top             =   1200
                  Width           =   3855
                  Begin VB.CheckBox chkNoCalcPeriodoFatt 
                     Caption         =   "Non calcolare il periodo di fatturazione"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   301
                     TabStop         =   0   'False
                     Top             =   960
                     Width           =   3615
                  End
                  Begin VB.CheckBox chkIstatAdeguamento 
                     Caption         =   "Adeguamento I.S.T.A.T."
                     Height          =   195
                     Left            =   120
                     TabIndex        =   95
                     TabStop         =   0   'False
                     Top             =   720
                     Width           =   2655
                  End
                  Begin VB.CheckBox chkAdegContrProx 
                     Caption         =   "Rinnovo automatico"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   94
                     TabStop         =   0   'False
                     Top             =   480
                     Width           =   3135
                  End
                  Begin VB.CheckBox chkAdeguaContrAttuale 
                     Caption         =   "Crea nuove rate del contratto attuale"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   93
                     TabStop         =   0   'False
                     Top             =   240
                     Width           =   3495
                  End
               End
               Begin VB.TextBox txtProtAdeg 
                  Height          =   315
                  Left            =   17520
                  TabIndex        =   88
                  Top             =   720
                  Width           =   2535
               End
               Begin VB.TextBox txtAnnotazioniAdeg 
                  Height          =   1095
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   91
                  Top             =   1320
                  Width           =   4575
               End
               Begin DMTEDITNUMLib.dmtNumber txtImportoAdeg 
                  Height          =   315
                  Left            =   8640
                  TabIndex        =   83
                  Top             =   720
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTDATETIMELib.dmtDate txtDataStipulaAdeg 
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   79
                  Top             =   720
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin VB.CommandButton cmdNuovoAdeg 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   22920
                  TabIndex        =   97
                  Top             =   2760
                  Width           =   1455
               End
               Begin VB.CommandButton cmdSalvaAdeg 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   22920
                  TabIndex        =   96
                  Top             =   3720
                  Width           =   1455
               End
               Begin VB.CommandButton cmdEliminaAdeg 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   22920
                  TabIndex        =   98
                  TabStop         =   0   'False
                  Top             =   4680
                  Width           =   1455
               End
               Begin VB.CommandButton cmdTrovaFattura 
                  Height          =   285
                  Left            =   -66120
                  Picture         =   "frmMain.frx":4E545
                  Style           =   1  'Graphical
                  TabIndex        =   28
                  ToolTipText     =   "Trova documento da collegare"
                  Top             =   1440
                  Width           =   375
               End
               Begin VB.TextBox txtOggettoCollegato 
                  Appearance      =   0  'Flat
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   -73320
                  Locked          =   -1  'True
                  TabIndex        =   27
                  Top             =   1440
                  Width           =   7200
               End
               Begin DmtGridCtl.DmtGrid GrigliaInterventi 
                  Height          =   7095
                  Left            =   -74880
                  TabIndex        =   77
                  Top             =   480
                  Width           =   24495
                  _ExtentX        =   43206
                  _ExtentY        =   12515
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
               Begin VB.CommandButton cmdGeneraInterventi 
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
                  Left            =   -74880
                  Picture         =   "frmMain.frx":4EACF
                  Style           =   1  'Graphical
                  TabIndex        =   76
                  ToolTipText     =   "Generazione interventi del servizio selezionato"
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   615
               End
               Begin VB.CommandButton cmdConfiguraServizi 
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
                  Left            =   -74880
                  Picture         =   "frmMain.frx":4F059
                  Style           =   1  'Graphical
                  TabIndex        =   75
                  ToolTipText     =   "Configurazione servizi (servizi configurati dal tipo di contratto)"
                  Top             =   480
                  Width           =   615
               End
               Begin VB.CommandButton cmdEliminaServizio 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -51960
                  TabIndex        =   49
                  TabStop         =   0   'False
                  Top             =   4320
                  Width           =   1455
               End
               Begin VB.CommandButton cmdSalvaServizio 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -51960
                  TabIndex        =   47
                  Top             =   3360
                  Width           =   1455
               End
               Begin VB.CommandButton cmdNuovoServizio 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -51960
                  TabIndex        =   48
                  Top             =   2400
                  Width           =   1455
               End
               Begin VB.TextBox txtPeriodo 
                  Height          =   855
                  Left            =   -64800
                  MaxLength       =   250
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   25
                  Top             =   1440
                  Width           =   7815
               End
               Begin VB.CheckBox chkRataFatturata 
                  Caption         =   "Fatturata"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   26
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.CommandButton cmdEliminaRata 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -51960
                  TabIndex        =   32
                  TabStop         =   0   'False
                  Top             =   4440
                  Width           =   1455
               End
               Begin VB.CommandButton cmdSalvaRata 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -51960
                  TabIndex        =   30
                  Top             =   3480
                  Width           =   1455
               End
               Begin VB.CommandButton cmdNuovaRata 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -51960
                  TabIndex        =   31
                  Top             =   2520
                  Width           =   1455
               End
               Begin DMTEDITNUMLib.dmtCurrency txtImportoRata 
                  Height          =   315
                  Left            =   -69960
                  TabIndex        =   22
                  Top             =   840
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   " 0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTDataCmb.DMTCombo cboPagamentoRataContratto 
                  Height          =   315
                  Left            =   -72600
                  TabIndex        =   21
                  Top             =   840
                  Width           =   2535
                  _ExtentX        =   4471
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
               Begin DMTDATETIMELib.dmtDate txtDataRata 
                  Height          =   315
                  Left            =   -74280
                  TabIndex        =   20
                  Top             =   840
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DmtGridCtl.DmtGrid GrigliaRateContratto 
                  Height          =   5175
                  Left            =   -74880
                  TabIndex        =   33
                  TabStop         =   0   'False
                  Top             =   2400
                  Width           =   22695
                  _ExtentX        =   40031
                  _ExtentY        =   9128
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
                  UseUserSettings =   0   'False
                  ColumnsHeaderHeight=   20
               End
               Begin DmtGridCtl.DmtGrid GrigliaServizi 
                  Height          =   5535
                  Left            =   -74880
                  TabIndex        =   50
                  TabStop         =   0   'False
                  Top             =   2040
                  Width           =   22815
                  _ExtentX        =   40243
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
                  UseUserSettings =   0   'False
                  ColumnsHeaderHeight=   20
               End
               Begin DMTEDITNUMLib.dmtNumber txtOgniNumeroGiorni 
                  Height          =   315
                  Left            =   -66480
                  TabIndex        =   35
                  Top             =   780
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboCriterioRicorrenza 
                  Height          =   315
                  Left            =   -68640
                  TabIndex        =   34
                  Top             =   780
                  Width           =   2055
                  _ExtentX        =   3625
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
               Begin DMTEDITNUMLib.dmtNumber txtOgniNumeroMesi 
                  Height          =   315
                  Left            =   -65160
                  TabIndex        =   36
                  Top             =   780
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtOgniNumeroSettimane 
                  Height          =   315
                  Left            =   -63840
                  TabIndex        =   37
                  Top             =   780
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboTipoDataInizioRic 
                  Height          =   315
                  Left            =   -74040
                  TabIndex        =   38
                  Top             =   1500
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
               Begin DMTEDITNUMLib.dmtNumber txtGiornoFissoInizioRic 
                  Height          =   315
                  Left            =   -70800
                  TabIndex        =   39
                  Top             =   1500
                  Width           =   615
                  _Version        =   65536
                  _ExtentX        =   1085
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtMeseFissoInizioRic 
                  Height          =   315
                  Left            =   -70080
                  TabIndex        =   40
                  Top             =   1500
                  Width           =   615
                  _Version        =   65536
                  _ExtentX        =   1085
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboTipoDataFineRic 
                  Height          =   315
                  Left            =   -66120
                  TabIndex        =   42
                  Top             =   1500
                  Width           =   3255
                  _ExtentX        =   5741
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
               Begin DMTEDITNUMLib.dmtNumber txtGiornoFissoFineRic 
                  Height          =   315
                  Left            =   -62760
                  TabIndex        =   43
                  Top             =   1500
                  Width           =   615
                  _Version        =   65536
                  _ExtentX        =   1085
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtMeseFissoFineRic 
                  Height          =   315
                  Left            =   -62040
                  TabIndex        =   44
                  Top             =   1500
                  Width           =   615
                  _Version        =   65536
                  _ExtentX        =   1085
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtNumeroRicorrenze 
                  Height          =   315
                  Left            =   -58080
                  TabIndex        =   46
                  Top             =   1500
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DmtGridCtl.DmtGrid GrigliaAdeg 
                  Height          =   5055
                  Left            =   120
                  TabIndex        =   99
                  TabStop         =   0   'False
                  Top             =   2520
                  Width           =   22575
                  _ExtentX        =   39820
                  _ExtentY        =   8916
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
                  UseUserSettings =   0   'False
                  ColumnsHeaderHeight=   20
               End
               Begin DMTDATETIMELib.dmtDate txtDataDecorrenzaAdeg 
                  Height          =   315
                  Left            =   2640
                  TabIndex        =   80
                  Top             =   720
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboIVARateContratto 
                  Height          =   315
                  Left            =   -68760
                  TabIndex        =   159
                  Top             =   2880
                  Width           =   2535
                  _ExtentX        =   4471
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
               Begin DMTDataCmb.DMTCombo cboIvaAdeg 
                  Height          =   315
                  Left            =   8280
                  TabIndex        =   162
                  Top             =   2880
                  Width           =   2535
                  _ExtentX        =   4471
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
               Begin DMTDataCmb.DMTCombo cboTipoAdeguamento 
                  Height          =   315
                  Left            =   14760
                  TabIndex        =   87
                  Top             =   720
                  Width           =   2655
                  _ExtentX        =   4683
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
               Begin DmtCodDescCtl.DmtCodDesc CDArticoloAdeg 
                  Height          =   615
                  Left            =   12600
                  TabIndex        =   86
                  Top             =   480
                  Width           =   2175
                  _ExtentX        =   3836
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":4F5E3
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":4F643
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":4F6AA
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
               Begin DmtCodDescCtl.DmtCodDesc CDArticoloProd 
                  Height          =   615
                  Left            =   -74520
                  TabIndex        =   107
                  Top             =   1080
                  Width           =   5655
                  _ExtentX        =   9975
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":4F704
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":4F753
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":4F7B3
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
               Begin DMTDATETIMELib.dmtDate txtDataInizioProd 
                  Height          =   315
                  Left            =   -65280
                  TabIndex        =   110
                  Top             =   1320
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataFineProd 
                  Height          =   315
                  Left            =   -61800
                  TabIndex        =   113
                  Top             =   1320
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboListinoProd 
                  Height          =   315
                  Left            =   -58200
                  TabIndex        =   115
                  Top             =   1320
                  Width           =   2055
                  _ExtentX        =   3625
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
               Begin DMTDATETIMELib.dmtDate txtDataInizioPer 
                  Height          =   315
                  Left            =   -68280
                  TabIndex        =   23
                  Top             =   840
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataFinePer 
                  Height          =   315
                  Left            =   -66600
                  TabIndex        =   24
                  Top             =   840
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboIstatAdeg 
                  Height          =   315
                  Left            =   20160
                  TabIndex        =   89
                  Top             =   720
                  Width           =   2655
                  _ExtentX        =   4683
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
               Begin DMTEDITNUMLib.dmtNumber txtNumeroRata 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   187
                  Top             =   840
                  Width           =   495
                  _Version        =   65536
                  _ExtentX        =   873
                  _ExtentY        =   556
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
               Begin DMTEDITNUMLib.dmtNumber txtIDOggettoCollegato 
                  Height          =   255
                  Left            =   -66720
                  TabIndex        =   136
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   975
                  _Version        =   65536
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTLblLinkCtl.LabelLink LabelLink2 
                  Height          =   255
                  Left            =   -74520
                  TabIndex        =   203
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   4455
                  _ExtentX        =   7858
                  _ExtentY        =   450
                  Caption         =   "Prodotto"
                  Name            =   "LabelLink"
               End
               Begin DMTEDITNUMLib.dmtNumber txtQtaArtProd 
                  Height          =   315
                  Left            =   -67440
                  TabIndex        =   120
                  Top             =   1920
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtImpUniProd 
                  Height          =   315
                  Left            =   -66120
                  TabIndex        =   121
                  Top             =   1920
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecimalPlaces   =   5
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtSconto1Prod 
                  Height          =   315
                  Left            =   -64440
                  TabIndex        =   122
                  Top             =   1920
                  Width           =   735
                  _Version        =   65536
                  _ExtentX        =   1296
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtSconto2Prod 
                  Height          =   315
                  Left            =   -63600
                  TabIndex        =   123
                  Top             =   1920
                  Width           =   735
                  _Version        =   65536
                  _ExtentX        =   1296
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtScontoImpProd 
                  Height          =   315
                  Left            =   -62760
                  TabIndex        =   134
                  Top             =   1920
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtImponibileProd 
                  Height          =   315
                  Left            =   -61080
                  TabIndex        =   124
                  Top             =   1920
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   12648447
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
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboUMPeriodoProd 
                  Height          =   315
                  Left            =   -67080
                  TabIndex        =   109
                  Top             =   1320
                  Width           =   1695
                  _ExtentX        =   2990
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
               Begin DMTEDITNUMLib.dmtNumber txtTotaleRigaProd 
                  Height          =   315
                  Left            =   -57720
                  TabIndex        =   126
                  Top             =   1920
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   12648447
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
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtQtaPeriodo 
                  Height          =   315
                  Left            =   -63120
                  TabIndex        =   112
                  Top             =   1320
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboUMArtProd 
                  Height          =   315
                  Left            =   -72240
                  TabIndex        =   117
                  Top             =   1920
                  Width           =   1695
                  _ExtentX        =   2990
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
               Begin DMTDATETIMELib.dmtTime txtOraFineProd 
                  Height          =   315
                  Left            =   -60480
                  TabIndex        =   114
                  Top             =   1320
                  Width           =   735
                  _Version        =   65536
                  _ExtentX        =   1296
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboTipoPeriodo 
                  Height          =   315
                  Left            =   -68880
                  TabIndex        =   108
                  Top             =   1320
                  Width           =   1695
                  _ExtentX        =   2990
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
               Begin DMTDataCmb.DMTCombo cboIvaProd 
                  Height          =   315
                  Left            =   -70440
                  TabIndex        =   118
                  Top             =   1920
                  Width           =   2175
                  _ExtentX        =   3836
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
               Begin DMTEDITNUMLib.dmtNumber txtAliquotaIvaProd 
                  Height          =   315
                  Left            =   -68280
                  TabIndex        =   119
                  Top             =   1920
                  Width           =   735
                  _Version        =   65536
                  _ExtentX        =   1296
                  _ExtentY        =   556
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
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtImportoIvaProd 
                  Height          =   315
                  Left            =   -59400
                  TabIndex        =   125
                  Top             =   1920
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   12648447
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
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDProdRifContr 
                  Height          =   255
                  Left            =   -63120
                  TabIndex        =   226
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   735
                  _Version        =   65536
                  _ExtentX        =   1296
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTLblLinkCtl.LabelLink LblLinkRil 
                  Height          =   255
                  Left            =   -72840
                  TabIndex        =   230
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  Caption         =   "Rilevamenti"
                  Name            =   "LabelLink"
               End
               Begin DMTEDITNUMLib.dmtNumber txtQuantitaEffettiva 
                  Height          =   315
                  Left            =   -59640
                  TabIndex        =   239
                  Top             =   1320
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
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
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboAnaOperatoreProd 
                  Height          =   315
                  Left            =   -54360
                  TabIndex        =   241
                  Top             =   600
                  Width           =   2055
                  _ExtentX        =   3625
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
               Begin DmtCodDescCtl.DmtCodDesc CDServizio 
                  Height          =   615
                  Left            =   -74040
                  TabIndex        =   257
                  Top             =   540
                  Width           =   5295
                  _ExtentX        =   9340
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":4F80D
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":4F85C
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":4F8B3
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
               Begin DMTDataCmb.DMTCombo cboTipoRateizzazioneProd 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   116
                  Top             =   1920
                  Width           =   2535
                  _ExtentX        =   4471
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
               Begin DMTDataCmb.DMTCombo cboTipoRateizzazioneAdeg 
                  Height          =   315
                  Left            =   6000
                  TabIndex        =   82
                  Top             =   720
                  Width           =   2535
                  _ExtentX        =   4471
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
               Begin DMTDATETIMELib.dmtDate txtDataScadenzaAdeg 
                  Height          =   315
                  Left            =   4320
                  TabIndex        =   81
                  Top             =   720
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboTipoAnnoInizioRicorr 
                  Height          =   315
                  Left            =   -69360
                  TabIndex        =   41
                  Top             =   1500
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
               Begin DMTDataCmb.DMTCombo cboTipoAnnoFineRicorr 
                  Height          =   315
                  Left            =   -61320
                  TabIndex        =   45
                  Top             =   1500
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
               Begin DMTEDITNUMLib.dmtNumber txtImportoAdegRinn 
                  Height          =   315
                  Left            =   10200
                  TabIndex        =   84
                  Top             =   720
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtNAdegIniz 
                  Height          =   315
                  Left            =   11760
                  TabIndex        =   85
                  Top             =   720
                  Width           =   735
                  _Version        =   65536
                  _ExtentX        =   1296
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin VB.Label Label25 
                  Caption         =   "N° iniz."
                  Height          =   255
                  Index           =   1
                  Left            =   11760
                  TabIndex        =   313
                  ToolTipText     =   "Numero della rata iniziale"
                  Top             =   480
                  Width           =   735
               End
               Begin VB.Label Label11 
                  Caption         =   "Importo al rinn."
                  Height          =   255
                  Index           =   1
                  Left            =   10200
                  TabIndex        =   312
                  ToolTipText     =   "Importo al rinnovo"
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.Label Label1 
                  Caption         =   "Anno fine ricorrenza"
                  Height          =   255
                  Index           =   33
                  Left            =   -61320
                  TabIndex        =   303
                  Top             =   1260
                  Width           =   3135
               End
               Begin VB.Label Label1 
                  Caption         =   "Anno inizio ricorrenza"
                  Height          =   255
                  Index           =   17
                  Left            =   -69360
                  TabIndex        =   302
                  Top             =   1260
                  Width           =   3135
               End
               Begin VB.Label Label6 
                  Caption         =   "Data scadenza"
                  Height          =   255
                  Left            =   4320
                  TabIndex        =   283
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.Label Label18 
                  Caption         =   "Rateizzazione"
                  Height          =   255
                  Index           =   11
                  Left            =   6000
                  TabIndex        =   282
                  Top             =   480
                  Width           =   2535
               End
               Begin VB.Label Label18 
                  Caption         =   "Rateizzazione"
                  Height          =   255
                  Index           =   10
                  Left            =   -74880
                  TabIndex        =   276
                  Top             =   1680
                  Width           =   2535
               End
               Begin VB.Label Label18 
                  Caption         =   "Q.tà effettiva"
                  Height          =   255
                  Index           =   9
                  Left            =   -59640
                  TabIndex        =   240
                  Top             =   1080
                  Width           =   1215
               End
               Begin VB.Line Line5 
                  BorderColor     =   &H00FF8080&
                  BorderWidth     =   2
                  X1              =   -56040
                  X2              =   -56040
                  Y1              =   720
                  Y2              =   2280
               End
               Begin VB.Label Label30 
                  Caption         =   "Totale I.V.A."
                  Height          =   255
                  Index           =   1
                  Left            =   -59400
                  TabIndex        =   223
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.Label Label18 
                  Caption         =   "% I.V.A."
                  Height          =   255
                  Index           =   8
                  Left            =   -68280
                  TabIndex        =   222
                  Top             =   1680
                  Width           =   735
               End
               Begin VB.Label Label18 
                  Caption         =   "Aliquota I.V.A."
                  Height          =   255
                  Index           =   7
                  Left            =   -70440
                  TabIndex        =   221
                  Top             =   1680
                  Width           =   3015
               End
               Begin VB.Label Label18 
                  Caption         =   "Tipo Periodo"
                  Height          =   255
                  Index           =   6
                  Left            =   -68880
                  TabIndex        =   219
                  Top             =   1080
                  Width           =   1695
               End
               Begin VB.Label Label18 
                  Caption         =   "Unità di misura"
                  Height          =   255
                  Index           =   5
                  Left            =   -72240
                  TabIndex        =   218
                  Top             =   1680
                  Width           =   1815
               End
               Begin VB.Label Label18 
                  Caption         =   "Listino"
                  Height          =   255
                  Index           =   4
                  Left            =   -58200
                  TabIndex        =   217
                  Top             =   1080
                  Width           =   1935
               End
               Begin VB.Label Label18 
                  Caption         =   "Q.tà periodo"
                  Height          =   255
                  Index           =   3
                  Left            =   -63120
                  TabIndex        =   216
                  Top             =   1080
                  Width           =   1095
               End
               Begin VB.Label Label18 
                  Caption         =   "Periodo"
                  Height          =   255
                  Index           =   2
                  Left            =   -67080
                  TabIndex        =   215
                  Top             =   1080
                  Width           =   1455
               End
               Begin VB.Label Label30 
                  Caption         =   "Totale"
                  Height          =   255
                  Index           =   0
                  Left            =   -57720
                  TabIndex        =   214
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.Label Label17 
                  Caption         =   "Importo unitario"
                  Height          =   255
                  Left            =   -66120
                  TabIndex        =   213
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.Label Label19 
                  Caption         =   "% Sc. 1"
                  Height          =   255
                  Left            =   -64440
                  TabIndex        =   212
                  Top             =   1680
                  Width           =   735
               End
               Begin VB.Label Label20 
                  Caption         =   "% Sc. 2"
                  Height          =   255
                  Left            =   -63600
                  TabIndex        =   211
                  Top             =   1680
                  Width           =   735
               End
               Begin VB.Label Label21 
                  Caption         =   "Sconto"
                  Height          =   255
                  Left            =   -62760
                  TabIndex        =   210
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.Label Label3 
                  Caption         =   "Imponibile"
                  Height          =   255
                  Left            =   -61080
                  TabIndex        =   209
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.Label Label18 
                  Caption         =   "Q.tà articolo"
                  Height          =   255
                  Index           =   1
                  Left            =   -67440
                  TabIndex        =   208
                  Top             =   1680
                  Width           =   1095
               End
               Begin VB.Label Label2 
                  Caption         =   "Prodotto"
                  Height          =   255
                  Index           =   12
                  Left            =   -64800
                  TabIndex        =   207
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.Label Label13 
                  Caption         =   "Annotazioni"
                  Height          =   255
                  Index           =   1
                  Left            =   -73320
                  TabIndex        =   191
                  Top             =   1760
                  Width           =   7575
               End
               Begin VB.Label Label2 
                  Caption         =   "N°"
                  Height          =   255
                  Index           =   11
                  Left            =   -74880
                  TabIndex        =   188
                  Top             =   600
                  Width           =   495
               End
               Begin VB.Label Label12 
                  Caption         =   "Note di fatturazione"
                  Height          =   255
                  Index           =   1
                  Left            =   4800
                  TabIndex        =   178
                  Top             =   1080
                  Width           =   5055
               End
               Begin VB.Label Label28 
                  Caption         =   "Maggiorazione"
                  Height          =   255
                  Left            =   22920
                  TabIndex        =   177
                  Top             =   480
                  Width           =   1695
               End
               Begin VB.Label Label2 
                  Caption         =   "Istat utilizzato al rinnovo"
                  Height          =   255
                  Index           =   10
                  Left            =   20160
                  TabIndex        =   176
                  Top             =   480
                  Width           =   2535
               End
               Begin VB.Label Label2 
                  Caption         =   "Fine periodo"
                  Height          =   255
                  Index           =   9
                  Left            =   -66600
                  TabIndex        =   175
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.Label Label2 
                  Caption         =   "Inizio periodo"
                  Height          =   255
                  Index           =   8
                  Left            =   -68280
                  TabIndex        =   174
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.Label Label25 
                  Caption         =   "Numero"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   173
                  Top             =   480
                  Width           =   855
               End
               Begin VB.Label Label23 
                  Caption         =   "Data e ora fine"
                  Height          =   255
                  Index           =   0
                  Left            =   -61800
                  TabIndex        =   172
                  Top             =   1080
                  Width           =   2055
               End
               Begin VB.Label Label22 
                  Caption         =   "Data e ora inizio"
                  Height          =   255
                  Left            =   -65280
                  TabIndex        =   171
                  Top             =   1080
                  Width           =   2055
               End
               Begin VB.Label Label18 
                  Caption         =   "Quantità"
                  Height          =   255
                  Index           =   0
                  Left            =   -65280
                  TabIndex        =   170
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.Label Label16 
                  Caption         =   "Annotazioni"
                  Height          =   255
                  Left            =   -74520
                  TabIndex        =   169
                  Top             =   2760
                  Width           =   3975
               End
               Begin VB.Label Label15 
                  Caption         =   "Ubicazione"
                  Height          =   255
                  Left            =   -62160
                  TabIndex        =   168
                  Top             =   480
                  Width           =   6015
               End
               Begin VB.Label Label14 
                  Caption         =   "Matricola"
                  Height          =   255
                  Index           =   0
                  Left            =   -68880
                  TabIndex        =   167
                  Top             =   480
                  Width           =   3495
               End
               Begin VB.Label Label2 
                  Caption         =   "Numero del protocollo"
                  Height          =   255
                  Index           =   7
                  Left            =   17520
                  TabIndex        =   165
                  Top             =   480
                  Width           =   2535
               End
               Begin VB.Label Label2 
                  Caption         =   "Tipo adeguamento al rinnovo"
                  Height          =   255
                  Index           =   6
                  Left            =   14760
                  TabIndex        =   164
                  Top             =   480
                  Width           =   2535
               End
               Begin VB.Label Label2 
                  Caption         =   "Aliquota I.V.A. fatturazione"
                  Height          =   255
                  Index           =   5
                  Left            =   8280
                  TabIndex        =   163
                  Top             =   2640
                  Width           =   2535
               End
               Begin VB.Label Label13 
                  Caption         =   "Documento di fatturazione"
                  Height          =   255
                  Index           =   0
                  Left            =   -73320
                  TabIndex        =   161
                  Top             =   1200
                  Width           =   6615
               End
               Begin VB.Label Label2 
                  Caption         =   "Aliquota I.V.A. fatturazione"
                  Height          =   255
                  Index           =   4
                  Left            =   -68760
                  TabIndex        =   160
                  Top             =   2640
                  Width           =   2535
               End
               Begin VB.Label Label12 
                  Caption         =   "Annotazioni"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   158
                  Top             =   1080
                  Width           =   1215
               End
               Begin VB.Label Label11 
                  Caption         =   "Importo"
                  Height          =   255
                  Index           =   0
                  Left            =   8640
                  TabIndex        =   157
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.Label Label10 
                  Caption         =   "Data decorrenza"
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   156
                  Top             =   480
                  Width           =   1575
               End
               Begin VB.Label Label9 
                  Caption         =   "Data stipula"
                  Height          =   255
                  Left            =   1080
                  TabIndex        =   155
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.Label Label1 
                  Caption         =   "Ogni n° sett."
                  Height          =   255
                  Index           =   30
                  Left            =   -63840
                  TabIndex        =   74
                  Top             =   540
                  Width           =   1215
               End
               Begin VB.Label Label1 
                  Caption         =   "Ogni n° mesi"
                  Height          =   255
                  Index           =   29
                  Left            =   -65160
                  TabIndex        =   73
                  Top             =   540
                  Width           =   1215
               End
               Begin VB.Label Label1 
                  Caption         =   "Ogni n° giorni"
                  Height          =   255
                  Index           =   28
                  Left            =   -66480
                  TabIndex        =   72
                  Top             =   540
                  Width           =   1215
               End
               Begin VB.Label Label1 
                  Caption         =   "Criterio di ricorrenza"
                  Height          =   255
                  Index           =   27
                  Left            =   -68640
                  TabIndex        =   71
                  Top             =   540
                  Width           =   1815
               End
               Begin VB.Label Label1 
                  Caption         =   "Numero ricorrenze"
                  Height          =   255
                  Index           =   26
                  Left            =   -58080
                  TabIndex        =   70
                  Top             =   1260
                  Width           =   1815
               End
               Begin VB.Label Label1 
                  Caption         =   "Mese"
                  Height          =   255
                  Index           =   25
                  Left            =   -62040
                  TabIndex        =   69
                  Top             =   1260
                  Width           =   615
               End
               Begin VB.Label Label1 
                  Caption         =   "Giorno"
                  Height          =   255
                  Index           =   24
                  Left            =   -62760
                  TabIndex        =   68
                  Top             =   1260
                  Width           =   615
               End
               Begin VB.Label Label1 
                  Caption         =   "Data fine ricorrenza"
                  Height          =   255
                  Index           =   23
                  Left            =   -66120
                  TabIndex        =   67
                  Top             =   1260
                  Width           =   2295
               End
               Begin VB.Label Label1 
                  Caption         =   "Mese"
                  Height          =   255
                  Index           =   22
                  Left            =   -70080
                  TabIndex        =   66
                  Top             =   1260
                  Width           =   615
               End
               Begin VB.Label Label1 
                  Caption         =   "Giorno"
                  Height          =   255
                  Index           =   21
                  Left            =   -70800
                  TabIndex        =   65
                  Top             =   1260
                  Width           =   615
               End
               Begin VB.Label Label1 
                  Caption         =   "Data inizio ricorrenza"
                  Height          =   255
                  Index           =   20
                  Left            =   -74040
                  TabIndex        =   64
                  Top             =   1260
                  Width           =   2175
               End
               Begin VB.Label Label2 
                  Caption         =   "Descrizione per fatturazione"
                  Height          =   255
                  Index           =   3
                  Left            =   -64800
                  TabIndex        =   62
                  Top             =   1200
                  Width           =   7695
               End
               Begin VB.Label Label2 
                  Caption         =   "Importo rata"
                  Height          =   255
                  Index           =   2
                  Left            =   -69960
                  TabIndex        =   61
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  Caption         =   "Modalità di pagamento"
                  Height          =   255
                  Index           =   1
                  Left            =   -72600
                  TabIndex        =   60
                  Top             =   600
                  Width           =   2175
               End
               Begin VB.Label Label2 
                  Caption         =   "Data fatturazione"
                  Height          =   255
                  Index           =   0
                  Left            =   -74280
                  TabIndex        =   59
                  Top             =   600
                  Width           =   1575
               End
            End
            Begin VB.Frame FraNoteFatt 
               Caption         =   "Note di fatturazione"
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
               Left            =   20760
               TabIndex        =   152
               Top             =   1920
               Width           =   4095
               Begin VB.TextBox txtNoteFatturazione 
                  Height          =   1215
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   15
                  Top             =   240
                  Width           =   3855
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Disdetta"
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
               Height          =   2055
               Left            =   20760
               TabIndex        =   56
               Top             =   0
               Width           =   4095
               Begin VB.TextBox txtDescrizioneDisdetta 
                  Height          =   1215
                  Left            =   120
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   18
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   3855
               End
               Begin VB.CheckBox chkDisdetta 
                  Caption         =   "Disdetta"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   16
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   1095
               End
               Begin DMTDATETIMELib.dmtDate txtDataDisdetta 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   17
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin VB.Label Label1 
                  Caption         =   "Data disdetta"
                  Height          =   255
                  Index           =   3
                  Left            =   1320
                  TabIndex        =   57
                  Top             =   120
                  Width           =   1215
               End
            End
         End
         Begin DmtGridCtl.DmtGrid BrwMain 
            Height          =   735
            Left            =   0
            TabIndex        =   314
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1296
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
            ColumnsHeaderHeight=   20
         End
      End
      Begin DmtPrnDlgCtl.DMTDialog DmtPrnDlg 
         Left            =   480
         Top             =   1290
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin VB.PictureBox picSplitter 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   960
         ScaleHeight     =   4935
         ScaleWidth      =   60
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin DmtActBox.DmtActBoxCtl ActivityBox 
         Height          =   9075
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   16007
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
      Begin DMTWheelCtrl.SpareWheel SpareWheel 
         Left            =   945
         Top             =   660
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.Image imgSplitter 
         Height          =   4695
         Left            =   2580
         MousePointer    =   9  'Size W E
         Top             =   0
         Width           =   60
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   1800
         X2              =   6720
         Y1              =   3360
         Y2              =   3360
      End
   End
   Begin MSComctlLib.StatusBar stbStatusbar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   51
      Top             =   13455
      Width           =   25365
      _ExtentX        =   44741
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'L'applicazione corrente
Private WithEvents m_App As DMTRunAppLib.Application
Attribute m_App.VB_VarHelpID = -1
'Il processo corrente
Private m_Process As DMTRunAppLib.Process
'Il tipo di documento corrente
Private m_DocType As DmtDocManLib.DBFormDocType
'Il documento corrente
Private WithEvents m_Document As DmtDocManLib.DBFormDocument
Attribute m_Document.VB_VarHelpID = -1
'La vista tabellare attiva
Private m_ActiveTableView As DmtDocManLib.TableView
'Il filtro attivo
Private m_ActiveFilter As DmtDocManLib.Filter
'Il report da stampare
Private m_Report As DmtDocManLib.Report
'La collezione dei campi del documento
'collegati ai controlli del form
Private m_FormFields As FormFields
'Il campo con la proprietà TabIndex uguale a 0
Private m_ControlTabIndex0 As Control
'La variabile  m_Semaphore mantiene un riferimento all'oggetto
'Semaphore che gestisce i conflitti di multiutenza
Private m_Semaphore As Semaforo.dmtSemaphore
'Indica se all'evento KeyPress del Form il tasto deve essere annullato
Private m_EatKey As Boolean
'Indica se l'utente ha modificato uno dei campi del documento
Private m_Changed As Boolean
'Indica se i valori dei campi del documento sono stati salvati
Private m_Saved As Boolean
'Indica se è in corso la definizione di una ricerca
Private m_Search As Boolean
'Indica se uno dei filtri è stato selezionato
Private m_FilterSelected As Boolean
'Indica lo stato di visibilità della vista tabellare
'prima dell'inizio della fase di esecuzione della
'anteprima di stampa
Private m_TabMode As Boolean
'Indica se si sta muovendo lo splitter
Private m_SplitterMoving As Boolean
'Nome dell'eventuale database esteso
Private m_ExtendedDatabase As String
'Processo "Shell su evento OnSave" nome del campo collegato
Private m_LinkedField As String
'Handle della finestra della anteprima di stampa
Private m_PreviewWindowHandle As Long
'Flag che permette l'esecuzione di Form_Activate soltanto all'avvio del programma
Private m_bOnFirstTime As Boolean
'Impedisce il Reposition della browse.
Private m_bAvoidReposition As Boolean
'Consente l'esecuzione del codice contenuto in BrwMain_OnChangeGuiMode()
Private bEnableGuiEvent As Boolean
'Indica se è stato attivato un link
Private m_LinkActive As Boolean




Public BLoadingContratto As Boolean

Private bLoadingProdotti As Boolean


'cbcx
'Oggetto adibito alla gestione del processo On_Extend
'Private m_ExtendApplication As DmtExtendAppLib.ExtendApplication


'rif1
'L'oggetto per la gestione dei sottodocumenti
'RATEIZZAZIONE
Private WithEvents m_DocumentsLink As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink.VB_VarHelpID = -1
'SERVIZI
Private WithEvents m_DocumentsLink1 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink1.VB_VarHelpID = -1
'ADEGUAMENTI
Private WithEvents m_DocumentsLink2 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink2.VB_VarHelpID = -1
'PRODOTTI
Private WithEvents m_DocumentsLink3 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink3.VB_VarHelpID = -1


'Costanti che rappresentano le modalità di visualizzazione
Private Enum neVisualModality
    Insert          'Modalità INSERIMENTO
    Modify          'Modalità VARIAZIONE
    Find            'Modalità TROVA
    Browse          'Modalità ELENCO
    Preview         'Modalità ANTEPRIMA
End Enum

'Costanti usate da SetStatus4Modality per l'apertura/chiusura dell'anteprima di stampa
Private Enum nePreviewModality
    OpenPrw
    ClosePrw
End Enum

Private m_iNumeroCopieDefault As Integer
Private m_OrientamentoDefault As OrientationConsts


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

Public bNotReturnValue As Boolean

'///////////////////////////////////////////////////////////////////////////////////
' ATTENZIONE:
' Occorre impostare questa costante!
' (ed eventualmente personalizzare il codice della funzione Caption2Display
'///////////////////////////////////////////////////////////////////////////////////
' Costante che identifica il campo più significativo del documento, il cui valore
' verrà visualizzato nella Caption del Form ed in quei messaggi in cui è mostrato
' il contenuto del campo principale del documento attivo.
' La costante può essere una stringa tipo "NomeCampo" o un intero che funge da indice
' nella collection m_Document.Fields().
'(Se l'applicazione può essere chiamata da un link occorre impostare anche la variabile
'sMessage1 presente nel metodo FormUnload.)
Private Const CAMPO_PER_CAPTION = "Anagrafica"


'Versione del controllo ActiveBar
Private Const BARMENUVERSION = "3.0"
'Variabile per la gestione degli shortcut del Menu
Private aryShortCut(1) As New ActiveBar3LibraryCtl.ShortCut

Private BLoading As Long



'****************************VARIABILI CONTRATTO**********************************
Public NuovaRata As Integer
Public ALIQUOTA_IVA_PRODOTTO As Double



Public rsGrigliaInt As ADODB.Recordset
Public rsGrigliaAcc As ADODB.Recordset
'******************************************************************************
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long


Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set m_App = NewValue
End Property

Public Property Get Application() As DMTRunAppLib.Application
    Set Application = m_App
End Property



'**+
'Nome: ChangeStringsLanguage
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Legge le stringe dal file di risorse per gestire l'opzione multilingue.
'Qui vanno inserite tutte le stringhe aggiunte in frmMain solo se si vuole
'gestire l'opzione multilingua
'**/
Public Sub ChangeStringsLanguage()
    '//////////////////////////////////////////////////////////////////////////////
    'ATTENZIONE
    'Inserire qui il codice per la lettura dal file di risorse di tutte le stringhe
    'per le quali si vuole gestire l'opzione multilingue.
    '//////////////////////////////////////////////////////////////////////////////
End Sub

'**+
'Nome: ChangeToolBarLanguage
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Legge dal file di risorse le stringe delle ToolTipText e dei suggerimenti da visualizzare
'sulla Statusbar per gestire l'opzione multilingue
'**/
Public Sub ChangeToolBarLanguage()

    'New
    BarMenu.Bands("Standard").Tools("New").ToolTipText = GetToolTipText4ToolBar("New")
    BarMenu.Bands("Standard").Tools("New").Description = GetDescription4StatusBar("New")
    
    'Save
    BarMenu.Bands("Standard").Tools("Save").ToolTipText = GetToolTipText4ToolBar("Save")
    BarMenu.Bands("Standard").Tools("Save").Description = GetDescription4StatusBar("Save")

    'Print
    BarMenu.Bands("Standard").Tools("Print").ToolTipText = GetToolTipText4ToolBar("Print")
    BarMenu.Bands("Standard").Tools("Print").Description = GetDescription4StatusBar("Print")

    'PrePrint
    BarMenu.Bands("Standard").Tools("PrePrint").ToolTipText = GetToolTipText4ToolBar("PrePrint")
    BarMenu.Bands("Standard").Tools("PrePrint").Description = GetDescription4StatusBar("PrePrint")

    'Cut
    BarMenu.Bands("Standard").Tools("Cut").ToolTipText = GetToolTipText4ToolBar("Cut")
    BarMenu.Bands("Standard").Tools("Cut").Description = GetDescription4StatusBar("Cut")

    'Copy
    BarMenu.Bands("Standard").Tools("Copy").ToolTipText = GetToolTipText4ToolBar("Copy")
    BarMenu.Bands("Standard").Tools("Copy").Description = GetDescription4StatusBar("Copy")

    'Paste
    BarMenu.Bands("Standard").Tools("Paste").ToolTipText = GetToolTipText4ToolBar("Paste")
    BarMenu.Bands("Standard").Tools("Paste").Description = GetDescription4StatusBar("Paste")

    'Delete
    BarMenu.Bands("Standard").Tools("Delete").ToolTipText = GetToolTipText4ToolBar("Delete")
    BarMenu.Bands("Standard").Tools("Delete").Description = GetDescription4StatusBar("Delete")

    'Clear
    BarMenu.Bands("Standard").Tools("Clear").ToolTipText = GetToolTipText4ToolBar("Clear")
    BarMenu.Bands("Standard").Tools("Clear").Description = GetDescription4StatusBar("Clear")

    'NewSearch
    BarMenu.Bands("Standard").Tools("NewSearch").ToolTipText = GetToolTipText4ToolBar("NewSearch")
    BarMenu.Bands("Standard").Tools("NewSearch").Description = GetDescription4StatusBar("NewSearch")

    'ExecuteSearch
    BarMenu.Bands("Standard").Tools("ExecuteSearch").ToolTipText = GetToolTipText4ToolBar("ExecuteSearch")
    BarMenu.Bands("Standard").Tools("ExecuteSearch").Description = GetDescription4StatusBar("ExecuteSearch")

    'ChangeView
    BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
    BarMenu.Bands("Standard").Tools("ChangeView").Description = GetDescription4StatusBar("ChangeView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").ToolTipText = GetToolTipText4ToolBar("ChangeView")

    'SearchPrevious
    BarMenu.Bands("Standard").Tools("SearchPrevious").ToolTipText = GetToolTipText4ToolBar("SearchPrevious")
    BarMenu.Bands("Standard").Tools("SearchPrevious").Description = GetDescription4StatusBar("SearchPrevious")

    'SearchNext
    BarMenu.Bands("Standard").Tools("SearchNext").ToolTipText = GetToolTipText4ToolBar("SearchNext")
    BarMenu.Bands("Standard").Tools("SearchNext").Description = GetDescription4StatusBar("SearchNext")

    'ExportWord
    BarMenu.Bands("Standard").Tools("ExportWord").ToolTipText = GetToolTipText4ToolBar("ExportWord")
    BarMenu.Bands("Standard").Tools("ExportWord").Description = GetDescription4StatusBar("ExportWord")

    'ExportExcel
    BarMenu.Bands("Standard").Tools("ExportExcel").ToolTipText = GetToolTipText4ToolBar("ExportExcel")
    BarMenu.Bands("Standard").Tools("ExportExcel").Description = GetDescription4StatusBar("ExportExcel")

    'ExportHtml
    BarMenu.Bands("Standard").Tools("ExportHtml").ToolTipText = GetToolTipText4ToolBar("ExportHtml")
    BarMenu.Bands("Standard").Tools("ExportHtml").Description = GetDescription4StatusBar("ExportHtml")

    'Assistente
    BarMenu.Bands("Standard").Tools("ViewAssistant").ToolTipText = GetToolTipText4ToolBar("ViewAssistant")
    BarMenu.Bands("Standard").Tools("ViewAssistant").Description = GetDescription4StatusBar("ViewAssistant")

    'Help
    BarMenu.Bands("Standard").Tools("Help").ToolTipText = GetToolTipText4ToolBar("Help")
    BarMenu.Bands("Standard").Tools("Help").Description = GetDescription4StatusBar("Help")

    BarMenu.RecalcLayout
End Sub


'**+
'Nome: ChangeMenuLanguage
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Legge dal file di risorse le stringe delle Caption e dei suggerimenti da visualizzare
'sulla Statusbar per gestire l'opzione multilingue
'**/
Public Sub ChangeMenuLanguage()

    '--- Menù PopUp del pulsante "ChangeView" della Toolbar ---
    'ChangeView - Form
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Caption = GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    
    'ChangeView - Tabella
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Caption = GetCaption4MenuBar("Mnu_TableView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    
    'ChangeView - Filtro
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Caption = GetCaption4MenuBar("Mnu_SearchFilter")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Description = GetDescription4StatusBar("Mnu_SearchFilter")
    '---                           ---                      ---
    

    'File
    BarMenu.Bands("Band_Menù").Tools("File").Caption = GetCaption4MenuBar("File")
    BarMenu.Bands("Band_Menù").Tools("File").Description = GetDescription4StatusBar("File")

    'File-New
    BarMenu.Bands("Band_File").Tools("Mnu_New").Caption = GetCaption4MenuBar("Mnu_New")
    BarMenu.Bands("Band_File").Tools("Mnu_New").Description = GetDescription4StatusBar("Mnu_New")
    
    'File-Save
    BarMenu.Bands("Band_File").Tools("Mnu_Save").Caption = GetCaption4MenuBar("Mnu_Save")
    BarMenu.Bands("Band_File").Tools("Mnu_Save").Description = GetDescription4StatusBar("Mnu_Save")
    
    'File-PrePrint
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Caption = GetCaption4MenuBar("Mnu_PrePrint")
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Description = GetDescription4StatusBar("Mnu_PrePrint")
    
    'File-Print
    BarMenu.Bands("Band_File").Tools("Mnu_Print").Caption = GetCaption4MenuBar("Mnu_Print")
    BarMenu.Bands("Band_File").Tools("Mnu_Print").Description = GetDescription4StatusBar("Mnu_Print")
    
    'File-Exit
    BarMenu.Bands("Band_File").Tools("Mnu_Exit").Caption = GetCaption4MenuBar("Mnu_Exit")
    BarMenu.Bands("Band_File").Tools("Mnu_Exit").Description = GetDescription4StatusBar("Mnu_Exit")
    
    'Edit
    BarMenu.Bands("Band_Menù").Tools("Edit").Caption = GetCaption4MenuBar("Edit")
    BarMenu.Bands("Band_Menù").Tools("Edit").Description = GetDescription4StatusBar("Edit")
    
    'Edit-Delete
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Caption = GetCaption4MenuBar("Mnu_Delete")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Description = GetDescription4StatusBar("Mnu_Delete")
    
    'Edit-Clear
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Caption = GetCaption4MenuBar("Mnu_Clear")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Description = GetDescription4StatusBar("Mnu_Clear")
    
    'Edit-Cut
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Caption = GetCaption4MenuBar("Mnu_Cut")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Description = GetDescription4StatusBar("Mnu_Cut")
    
    'Edit-Copy
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Caption = GetCaption4MenuBar("Mnu_Copy")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Description = GetDescription4StatusBar("Mnu_Copy")
    
    'Edit-Paste
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Caption = GetCaption4MenuBar("Mnu_Paste")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Description = GetDescription4StatusBar("Mnu_Paste")
    
    'Edit-NewSearch
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Caption = GetCaption4MenuBar("Mnu_NewSearch")
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Description = GetDescription4StatusBar("Mnu_NewSearch")
    
    'Edit-ExecuteSearch
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Caption = GetCaption4MenuBar("Mnu_ExecuteSearch")
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Description = GetDescription4StatusBar("Mnu_ExecuteSearch")
    
    'Edit-SearchPrevious
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Caption = GetCaption4MenuBar("Mnu_SearchPrevious")
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Description = GetDescription4StatusBar("Mnu_SearchPrevious")
    
    'Edit-SearchNext
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Caption = GetCaption4MenuBar("Mnu_SearchNext")
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Description = GetDescription4StatusBar("Mnu_SearchNext")
    
    'View
    BarMenu.Bands("Band_Menù").Tools("View").Caption = GetCaption4MenuBar("View")
    BarMenu.Bands("Band_Menù").Tools("View").Description = GetDescription4StatusBar("View")
    
    'View-FormView
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").Caption = GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    
    'View-TableView
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").Caption = GetCaption4MenuBar("Mnu_TableView")
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    
    'View-SearchFilter
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Caption = GetCaption4MenuBar("Mnu_SearchFilter")
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Description = GetDescription4StatusBar("Mnu_SearchFilter")
    
    'View-Folders
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Caption = GetCaption4MenuBar("Mnu_Folders")
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Description = GetDescription4StatusBar("Mnu_Folders")
    
    'View-ToolBar
    BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Caption = GetCaption4MenuBar("Mnu_ToolBar")
    BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Description = GetDescription4StatusBar("Mnu_ToolBar")
    
    'Tools
    BarMenu.Bands("Band_Menù").Tools("Tools").Caption = GetCaption4MenuBar("Tools")
    BarMenu.Bands("Band_Menù").Tools("Tools").Description = GetDescription4StatusBar("Tools")
    
    'Tools-Export
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Caption = GetCaption4MenuBar("Mnu_Export")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Description = GetDescription4StatusBar("Mnu_Export")
    
    'Tools-Options
    BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Caption = GetCaption4MenuBar("Mnu_Options")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Description = GetDescription4StatusBar("Mnu_Options")
    
    'Tools-Export-ExportWord
    BarMenu.Bands("Band_Export").Tools("Mnu_ExportWord").Caption = GetCaption4MenuBar("Mnu_ExportWord")
    BarMenu.Bands("Band_Export").Tools("Mnu_ExportWord").Description = GetDescription4StatusBar("Mnu_ExportWord")
    
    'Tools-Export-ExportExcel
    BarMenu.Bands("Band_Export").Tools("Mnu_ExportExcel").Caption = GetCaption4MenuBar("Mnu_ExportExcel")
    BarMenu.Bands("Band_Export").Tools("Mnu_ExportExcel").Description = GetDescription4StatusBar("Mnu_ExportExcel")
    
    'Tools-Export-ExportHtml
    BarMenu.Bands("Band_Export").Tools("Mnu_ExportHtml").Caption = GetCaption4MenuBar("Mnu_ExportHtml")
    BarMenu.Bands("Band_Export").Tools("Mnu_ExportHtml").Description = GetDescription4StatusBar("Mnu_ExportHtml")

    'Help
    BarMenu.Bands("Band_Menù").Tools("Help").Caption = GetCaption4MenuBar("Help")
    BarMenu.Bands("Band_Menù").Tools("Help").Description = GetDescription4StatusBar("Help")
    
    'Help-HelpOnLine
    BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").Caption = GetCaption4MenuBar("Mnu_HelpOnLine")
    BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").Description = GetDescription4StatusBar("Mnu_HelpOnLine")
    
    'Help-Summary
    BarMenu.Bands("Band_Help").Tools("Mnu_Summary").Caption = GetCaption4MenuBar("Mnu_Summary")
    BarMenu.Bands("Band_Help").Tools("Mnu_Summary").Description = GetDescription4StatusBar("Mnu_Summary")
    
    'Help-FastHelp
    BarMenu.Bands("Band_Help").Tools("Mnu_FastHelp").Caption = GetCaption4MenuBar("Mnu_FastHelp")
    BarMenu.Bands("Band_Help").Tools("Mnu_FastHelp").Description = GetDescription4StatusBar("Mnu_FastHelp")
    
    'Help-Arg
    BarMenu.Bands("Band_Help").Tools("Mnu_Arg").Caption = GetCaption4MenuBar("Mnu_Arg")
    BarMenu.Bands("Band_Help").Tools("Mnu_Arg").Description = GetDescription4StatusBar("Mnu_Arg")
    
    'Help-Web
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").Caption = GetCaption4MenuBar("Mnu_Web")
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").Description = GetDescription4StatusBar("Mnu_Web")
    
    'Help-Assistant
    BarMenu.Bands("Band_Help").Tools("Mnu_Assistant").Caption = GetCaption4MenuBar("Mnu_Assistant")
    BarMenu.Bands("Band_Help").Tools("Mnu_Assistant").Description = GetDescription4StatusBar("Mnu_Assistant")
    
    'Help-Info
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Caption = GetCaption4MenuBar("Mnu_Info")
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Description = GetDescription4StatusBar("Mnu_Info")
    
    'Help-Web-Web1
    BarMenu.Bands("Band_Web").Tools("Mnu_Web1").Caption = GetCaption4MenuBar("Mnu_Web1")
    
    'Help-Assistant-ChoiceAssistant
    BarMenu.Bands("Band_Assistant").Tools("Mnu_ChoiceAssistant").Caption = GetCaption4MenuBar("Mnu_ChoiceAssistant")
    BarMenu.Bands("Band_Assistant").Tools("Mnu_ChoiceAssistant").Description = GetDescription4StatusBar("Mnu_ChoiceAssistant")
    
    'Help-Assistant-ViewAssistant
    BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAssistant").Caption = GetCaption4MenuBar("Mnu_ViewAssistant")
    BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAssistant").Description = GetDescription4StatusBar("Mnu_ViewAssistant")
    
    'Help-Assistant-ViewAnimation
    BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAnimation").Caption = GetCaption4MenuBar("Mnu_ViewAnimation")
    BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAnimation").Description = GetDescription4StatusBar("Mnu_ViewAnimation")
    
    'PopUp-RunApplication
    BarMenu.Bands("Band_PopUp").Tools("Mnu_RunApplication").Caption = GetCaption4MenuBar("Mnu_RunApplication")
    
    'PopUp-SearchObject
    BarMenu.Bands("Band_PopUp").Tools("Mnu_SearchObject").Caption = GetCaption4MenuBar("Mnu_SearchObject")
    
    BarMenu.RecalcLayout
End Sub


'**+
'Nome: SetStatusBarVisibility
'
'Parametri:Boolean che valorizzerà la proprietà Visible della StatusBar
'
'Valori di ritorno:
'
'Funzionalità:
'Su richiesta di frmOption, Mostra/Nasconde la Statusbar
'**/
Public Sub SetStatusBarVisibility(ByVal bVisible As Boolean)
    stbStatusbar.Visible = bVisible
End Sub

'**+
'Nome: SetToolBarIcons
'
'Parametri:
'LargeIcons - Il tipo di icona da usare per i bottoni,
'grandi o piccole
'
'Valori di ritorno:
'
'Funzionalità:
'Cambia il tipo di icona della ToolBar standard
'**/
Public Sub SetToolBarIcons(ByVal LargeIcons As Boolean)
    Dim iPicture As Integer

    BarMenu.LargeIcons = LargeIcons
    If LargeIcons Then
        BarMenu.Bands("Standard").Tools("New").SetPicture 0, gResource.GetBitmap(IDB_STD_NEW32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Save").SetPicture 0, gResource.GetBitmap(IDB_STD_SAVE32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Print").SetPicture 0, gResource.GetBitmap(IDB_STD_PRINT32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("PrePrint").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIEW32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Cut").SetPicture 0, gResource.GetBitmap(IDB_STD_CUT32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Copy").SetPicture 0, gResource.GetBitmap(IDB_STD_COPY32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Paste").SetPicture 0, gResource.GetBitmap(IDB_STD_PASTE32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Delete").SetPicture 0, gResource.GetBitmap(IDB_STD_DELETE32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Clear").SetPicture 0, gResource.GetBitmap(IDB_STD_CLEAR32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("NewSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_FIND32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("ExecuteSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_EXECUTE32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("SearchPrevious").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIOUS32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("SearchNext").SetPicture 0, gResource.GetBitmap(IDB_STD_NEXT32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Export").SetPicture 0, gResource.GetBitmap(IDB_EXPORT_32), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportWord").SetPicture 0, gResource.GetBitmap(IDB_STD_WORD32), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportExcel").SetPicture 0, gResource.GetBitmap(IDB_STD_EXCEL32), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportHtml").SetPicture 0, gResource.GetBitmap(IDB_STD_HTML32), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportPDF").SetPicture 0, gResource.GetBitmap(IDB_ACROBAT_32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Web").SetPicture 0, gResource.GetBitmap(IDB_DMT_WEB32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Agg_Web").SetPicture 0, gResource.GetBitmap(IDB_AGG_WEB32), &HC0C0C0
        
        'cbc - L'icona del pulsante "ChangeView" dipende dalla modalità attuale
        iPicture = IIf(BrwMain.Visible, IDB_STD_FORM32, IDB_STD_GRID32)
        BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
        
        BarMenu.LargeIcons = False
    Else
        BarMenu.Bands("Standard").Tools("New").SetPicture 0, gResource.GetBitmap(IDB_STD_NEW16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Save").SetPicture 0, gResource.GetBitmap(IDB_STD_SAVE16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Print").SetPicture 0, gResource.GetBitmap(IDB_STD_PRINT16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("PrePrint").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIEW16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Cut").SetPicture 0, gResource.GetBitmap(IDB_STD_CUT16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Copy").SetPicture 0, gResource.GetBitmap(IDB_STD_COPY16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Paste").SetPicture 0, gResource.GetBitmap(IDB_STD_PASTE16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Delete").SetPicture 0, gResource.GetBitmap(IDB_STD_DELETE16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Clear").SetPicture 0, gResource.GetBitmap(IDB_STD_CLEAR16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("NewSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_FIND16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("ExecuteSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_EXECUTE16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("SearchPrevious").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIOUS16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("SearchNext").SetPicture 0, gResource.GetBitmap(IDB_STD_NEXT16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Export").SetPicture 0, gResource.GetBitmap(IDB_EXPORT_16), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportWord").SetPicture 0, gResource.GetBitmap(IDB_STD_WORD16), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportExcel").SetPicture 0, gResource.GetBitmap(IDB_STD_EXCEL16), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportHtml").SetPicture 0, gResource.GetBitmap(IDB_STD_HTML16), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportPDF").SetPicture 0, gResource.GetBitmap(IDB_ACROBAT_16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Web").SetPicture 0, gResource.GetBitmap(IDB_DMT_WEB16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Agg_Web").SetPicture 0, gResource.GetBitmap(IDB_AGG_WEB16), &HC0C0C0
    
        'L'icona del pulsante "ChangeView" dipende dalla modalità attuale
        iPicture = IIf(BrwMain.Visible, IDB_STD_FORM16, IDB_STD_GRID16)
        BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
    End If
    BarMenu.RecalcLayout
End Sub

'**+
'Nome: SetVisibilityIDFields
'
'Parametri: Optional IDVisible As Variant (Boolean)
'           Se IDVisible è presente (chiamata da frmOption) viene usato il suo valore
'           per settare la visibilità dei campi ID, altrimenti viene letta l'impostazione
'           del registry
'
'Valori di ritorno:
'
'Funzionalità: Mostra/Nasconde i campi ID della Browse
'**/
Public Sub SetVisibilityIDFields(Optional ByVal IDVisible As Variant)
    Dim Col As DmtGridCtl.dgColumnHeader
    Dim bValue As Boolean

    'Legge le impostazioni dal registry
    bValue = IIf(IsMissing(IDVisible), AppOptions.IDFieldsVisibility, IDVisible)

    For Each Col In BrwMain.ColumnsHeader
        If Left(Col.FieldName, 2) = "ID" Then
            Col.Visible = bValue
        End If
    Next Col
    BrwMain.LoadUserSettings
    'L'aspetto della browse viene ridisegnato.
    BrwMain.Refresh
End Sub


'ATTENZIONE: Nella funzione GetDescription4StatusBar vanno impostati tutti i
'            suggerimenti dei pulsanti della Toolbar e delle voci di menu
'            da visualizzare sulla Statusbar.
'            Per gestire l'opzione multilingua occorre inserire nel file di risorse
'            tutte le stringhe occorrenti.

'**+
'Nome:   GetDescription4StatusBar
'
'Parametri: sToolName è il nome del pulsante o della voce di menu per i quali
'           si vuole ottenere il messaggio sulla Statusbar
'
'Valori di ritorno: La stringa da visualizzare sulla StatusBar
'
'Funzionalità: Restituisce la stringa del suggerimento associato ad un bottone
'              della toolbar o ad una voce di menu
'**/
Private Function GetDescription4StatusBar(ByVal sToolName As String) As String
    Dim sApplicationName As String
    Dim sTipoOggetto As String
    Dim sStr As String
    Dim sTemp As String

    
    sApplicationName = m_App.FunctionName
    sTipoOggetto = m_DocType.Name
    
    Select Case sToolName
    
        Case "File"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_App.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_FILE)
        
        Case "New", "Mnu_New"
            sStr = "Crea un nuovo " & sTipoOggetto
            
        Case "Save", "Mnu_Save"
            sStr = "Memorizza il " & sTipoOggetto & " corrente"
        
        Case "Print", "Mnu_Print"
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sStr = "Stampa i " & sTipoOggetto & " correnti"
            Else
                'Si è in modalità form
                sStr = "Stampa il " & sTipoOggetto & " corrente"
            End If
            
        Case "PrePrint", "Mnu_PrePrint"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                sTemp = sTipoOggetto & " correnti"
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_SETPREVIEW)
            Else
                'Si è in modalità form
                sTemp = sTipoOggetto & " corrente"
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_SETPREVIEW)
            End If
    
        Case "Mnu_Exit"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add TheApp.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_EXIT)
            
        Case "Edit"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_App.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_MODIFY)
    
        Case "Cut", "Mnu_Cut"
            sStr = gResource.GetCustomizedMessage(IDS_SB_CUT)
            
        Case "Copy", "Mnu_Copy"
            sStr = gResource.GetCustomizedMessage(IDS_SB_COPY)
            
        Case "Paste", "Mnu_Paste"
            sStr = gResource.GetCustomizedMessage(IDS_SB_PASTE)
            
        Case "Delete", "Mnu_Delete"
            sStr = "Elimina il " & sTipoOggetto & " corrente"
            
        Case "Clear", "Mnu_Clear"
            sStr = gResource.GetCustomizedMessage(IDS_SB_CLEAR)
            
        Case "NewSearch", "Mnu_NewSearch"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_SEARCHWINDOW)
            
        Case "ExecuteSearch", "Mnu_ExecuteSearch"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_SEARCHEXECUTE)
            
        Case "Mnu_FormView"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_FORM)
            
        Case "Mnu_TableView"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_TABLE)
            
        Case "Mnu_SearchFilter"
            sStr = "Espone " & m_DocType.Name & " in modo <filtri>."
            
        Case "ChangeView"
            If BrwMain.Visible And BrwMain.GuiMode = dgNormal Then
                'Si è in modalità tabellare
                gResource.CustomStrings.Clear
                gResource.CustomStrings.Add m_DocType.Name, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_FORM)
            Else
                'Si è in modalità form
                gResource.CustomStrings.Clear
                gResource.CustomStrings.Add m_DocType.Name, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_TABLE)
            End If
            
        Case "View"
            sStr = gResource.GetMessage(IDS_SB_DISPLAY)
            
        Case "SearchPrevious", "Mnu_SearchPrevious"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_SEARCHPREVIOUS)
            
        Case "SearchNext", "Mnu_SearchNext"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_SEARCHNEXT)
            
        Case "Mnu_Folders"
            sStr = "Riquadro attività"
            
        Case "Mnu_ToolBar"
            sStr = gResource.GetMessage(IDS_SB_TOOLBAR)
            
        Case "Tools"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add TheApp.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_TOOLS)
            
        Case "Mnu_Export"
            sStr = gResource.GetMessage(IDS_SB_EXPORT)
            
        Case "Mnu_Options"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add TheApp.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_OPTION)

            
        Case "ExportWord", "Mnu_ExportWord"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sTemp = " i " & sTipoOggetto & " correnti "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTWORD)
            Else
                'Si è in modalità form
                sTemp = " il " & sTipoOggetto & " corrente "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTWORD)
            End If
        
        Case "ExportExcel", "Mnu_ExportExcel"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sTemp = " i " & sTipoOggetto & " correnti "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTEXCEL)
            Else
                'Si è in modalità form
                sTemp = " il " & sTipoOggetto & " corrente "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTEXCEL)
            End If
        
        Case "ExportHtml", "Mnu_ExportHtml"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sTemp = " i " & sTipoOggetto & " correnti "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTHTML)
            Else
                'Si è in modalità form
                sTemp = " il " & sTipoOggetto & " corrente "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTHTML)
            End If
        
        Case "ExportPDF", "Mnu_ExportPDF"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sTemp = " i " & sTipoOggetto & " correnti "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTACROBAT)
            Else
                'Si è in modalità form
                sTemp = " il " & sTipoOggetto & " corrente "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTACROBAT)
            End If
        
        Case "Mnu_HelpOnLine"
            sStr = gResource.GetMessage(IDS_SB_SUMMARY)
            
        Case "Mnu_Arg"
            sStr = gResource.GetMessage(IDS_SB_ARG)
            
        Case "Mnu_Web"
            sStr = gResource.GetMessage(IDS_SB_WEB)
            
        Case "Mnu_Info"
            sStr = gResource.GetMessage(IDS_SB_INFO)
            
        Case "Mnu_Web", "Web"
            sStr = gResource.GetMessage(IDS_SB_WEB)
        
        Case "Mnu_Agg_Web", "Agg_Web"
            sStr = gResource.GetMessage(IDS_SB_AGG_WEB)
    End Select
    
    GetDescription4StatusBar = sStr
End Function
'//////////////////////////////////////////////////////////////////////////////////
'ATTENZIONE: Nella funzione GetToolTipText4ToolBar vanno impostate tutte le
'            stringhe dei ToolTipText dei pulsanti della Toolbar.
'            Per gestire l'opzione multilingua occorre inserire nel file di risorse
'            tutte le stringhe occorrenti.
'//////////////////////////////////////////////////////////////////////////////////
'**+
'Nome:   GetToolTipText4ToolBar
'
'Parametri: sToolName è il nome del pulsante per il quale
'           si vuole ottenere la stringa per la proprietà ToolTipText
'
'Valori di ritorno: La stringa ToolTipText
'
'Funzionalità: Restituisce la stringa del suggerimento associato ad un bottone
'              della toolbar (ToolTipext)
'**/
Private Function GetToolTipText4ToolBar(ByVal sToolName As String) As String
    Dim sStr As String
    
    gResource.CustomStrings.Clear
    
    Select Case sToolName
    
        Case "New"
            sStr = gResource.GetMessage(TT_NEW)
            
        Case "Save"
            sStr = gResource.GetMessage(TT_SAVE)
        
        Case "Print"
            sStr = gResource.GetMessage(TT_PRINT)
            
        Case "PrePrint"
            sStr = gResource.GetMessage(TT_PREVIEW)
    
        Case "Cut"
            sStr = gResource.GetMessage(TT_CUT)
            
        Case "Copy"
            sStr = gResource.GetMessage(TT_COPY)
            
        Case "Paste"
            sStr = gResource.GetMessage(TT_PASTE)
            
        Case "Delete"
            sStr = gResource.GetMessage(TT_DELETE)
            
        Case "Clear"
            sStr = gResource.GetMessage(TT_CLEAR)
            
        Case "NewSearch"
            sStr = gResource.GetMessage(TT_SEARCH)
            
        Case "ExecuteSearch"
            sStr = gResource.GetMessage(TT_SEARCHEXECUTE)
            
        Case "ChangeView"
            If BrwMain.Visible And BrwMain.GuiMode = dgNormal Then
                'Si è in modalità tabellare
                sStr = gResource.GetMessage(TT_FORM)
            Else
                'Si è in modalità form
                sStr = gResource.GetMessage(TT_SEARCHRESULT)
            End If
            
        Case "SearchPrevious"
            sStr = gResource.GetMessage(TT_SEARCHPREVIOUS)
            
        Case "SearchNext"
            sStr = gResource.GetMessage(TT_SEARCHNEXT)
            
        Case "ExportWord"
            sStr = gResource.GetMessage(TT_WORD)
        
        Case "ExportExcel"
            sStr = gResource.GetMessage(TT_EXCEL)
        
        Case "ExportHtml"
            sStr = gResource.GetMessage(TT_HTML)
        
        Case "ViewAssistant" 'toolbar
            sStr = gResource.GetMessage(TT_SHOW_ASSISTANT)
            
        Case "Help" 'toolbar e menu
            sStr = gResource.GetMessage(TT_HELP)

    End Select
    
    GetToolTipText4ToolBar = sStr
End Function

'//////////////////////////////////////////////////////////////////////////////////
'ATTENZIONE: Nella funzione GetCaption4MenuBar vanno impostate tutte le
'            stringhe delle Caption delle voci di menu.
'            Per gestire l'opzione multilingua occorre inserire nel file di risorse
'            tutte le stringhe occorrenti.
'//////////////////////////////////////////////////////////////////////////////////
'**+
'Nome:   GetCaption4MenuBar
'
'Parametri: sToolName è il nome della voce di menu per la quale
'           si vuole ottenere la stringa per la Caption
'
'Valori di ritorno: La stringa da visualizzare nella Caption del menu
'
'Funzionalità: Restituisce la stringa della Caption di una voce di menu
'**/
Private Function GetCaption4MenuBar(ByVal sToolName As String) As String
    Dim sStr As String
    
    gResource.CustomStrings.Clear
    
    Select Case sToolName
    
        Case "File"
            sStr = gResource.GetMessage(MNU_FILE)
        
        Case "Mnu_New"
            sStr = gResource.GetMessage(MNU_NEW)
            aryShortCut(1).Value = "Control+N"
            BarMenu.Bands("Band_File").Tools("Mnu_New").ShortCuts = aryShortCut
            
        Case "Mnu_Save"
            If m_App.Language <> 1 Then
                sStr = gResource.GetMessage(MNU_SAVE)
                aryShortCut(1).Value = "Control+S"
                BarMenu.Bands("Band_File").Tools("Mnu_Save").ShortCuts = aryShortCut
            Else
                sStr = gResource.GetMessage(MNU_SAVE)
                aryShortCut(1).Value = "Shift+F12"
                BarMenu.Bands("Band_File").Tools("Mnu_Save").ShortCuts = aryShortCut
            End If
        
        Case "Mnu_PrePrint"
            sStr = gResource.GetMessage(MNU_PREVIEW)
        
        Case "Mnu_Print"
            If m_App.Language <> 1 Then
                sStr = gResource.GetMessage(MNU_PRINT) & "..."
                aryShortCut(1).Value = "Control+P"
                BarMenu.Bands("Band_File").Tools("Mnu_Print").ShortCuts = aryShortCut
            Else
                sStr = gResource.GetMessage(MNU_PRINT) & "..."
                aryShortCut(1).Value = "Control+Shift+F12"
                BarMenu.Bands("Band_File").Tools("Mnu_Print").ShortCuts = aryShortCut
            End If
    
        Case "Mnu_Exit"
            sStr = gResource.GetMessage(MNU_EXIT)
            
        Case "Edit"
            sStr = gResource.GetMessage(MNU_MODIFY)
    
        Case "Mnu_Delete"
            sStr = gResource.GetMessage(MNU_DELETE)
            aryShortCut(1).Value = "Delete"
            BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").ShortCuts = aryShortCut
            
        Case "Mnu_Clear"
            sStr = gResource.GetMessage(MNU_CLEAR)
    
        Case "Mnu_Cut"
            sStr = gResource.GetMessage(MNU_CUT)
            aryShortCut(1).Value = "Control+X"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").ShortCuts = aryShortCut
            
        Case "Mnu_Copy"
            sStr = gResource.GetMessage(MNU_COPY)
            aryShortCut(1).Value = "Control+C"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").ShortCuts = aryShortCut
            
        Case "Mnu_Paste"
            sStr = gResource.GetMessage(MNU_PASTE)
            aryShortCut(1).Value = "Control+V"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").ShortCuts = aryShortCut
            
        Case "Mnu_NewSearch"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_FIND)
            aryShortCut(1).Value = "Control+T"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").ShortCuts = aryShortCut
            
        Case "Mnu_ExecuteSearch"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_EXECUTE_SEARCH)
            aryShortCut(1).Value = "Control+E"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").ShortCuts = aryShortCut
            
        Case "Mnu_SearchPrevious"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_PREVIOUS_SEARCH)
            aryShortCut(1).Value = "Control+P"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").ShortCuts = aryShortCut
            
        Case "Mnu_SearchNext"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_NEXT_SEARCH)
            aryShortCut(1).Value = "Control+S"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").ShortCuts = aryShortCut
            
        Case "View"
            sStr = gResource.GetMessage(MNU_DISPLAY)
            
        Case "Mnu_FormView"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_FORM)
            aryShortCut(1).Value = "Control+F"
            frmMain.BarMenu.Bands("Band_View").Tools("Mnu_FormView").ShortCuts = aryShortCut
            
        Case "Mnu_TableView"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_TABLE)
            aryShortCut(1).Value = "Control+M"
            frmMain.BarMenu.Bands("Band_View").Tools("Mnu_TableView").ShortCuts = aryShortCut
            
        Case "Mnu_SearchFilter"
            sStr = "Mo&dalità filtri"
            aryShortCut(1).Value = "Control+Shift+T"
            frmMain.BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").ShortCuts = aryShortCut
            
        Case "Mnu_Folders"
            sStr = "&Riquadro attività"
            
        Case "Mnu_ToolBar"
            sStr = gResource.GetMessage(MNU_TOOLBAR)
            
        Case "Tools"
            sStr = gResource.GetMessage(MNU_TOOL)
            
        Case "Mnu_Export"
            sStr = gResource.GetMessage(MNU_EXPORT)
            
        Case "Mnu_Options"
            sStr = gResource.GetMessage(MNU_OPTION)
            
        Case "Mnu_ExportWord"
                sStr = gResource.GetMessage(MNU_EXPORT_WORD)
        
        Case "Mnu_ExportExcel"
                sStr = gResource.GetMessage(MNU_EXPORT_EXCEL)
        
        Case "Mnu_ExportHtml"
                sStr = gResource.GetMessage(MNU_EXPORT_HTML)
        
        Case "Mnu_ExportPDF"
                sStr = gResource.GetMessage(MNU_EXPORT_ACROBAT)
        
        Case "Help" 'toolbar e menu
            sStr = "&?"

        Case "Mnu_HelpOnLine"
            sStr = gResource.GetMessage(MNU_HELP)
            aryShortCut(1).Value = "F1"
            frmMain.BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").ShortCuts = aryShortCut
            
        Case "Mnu_Arg"
            sStr = gResource.GetMessage(MNU_ARG)
            aryShortCut(1).Value = "Shift+F1"
            frmMain.BarMenu.Bands("Band_Help").Tools("Mnu_Arg").ShortCuts = aryShortCut
            
        Case "Mnu_Web"
            sStr = gResource.GetMessage(MNU_WEB)
            
        Case "Mnu_Agg_Web"
            sStr = gResource.GetMessage(MNU_AGG_WEB)
            
        Case "Mnu_Info"
            sStr = gResource.GetMessage(MNU_INFO)
            
        Case "Mnu_RunApplication"
            sStr = gResource.GetMessage(MNU_EXE_GEST)
            aryShortCut(1).Value = "Control+G"
            frmMain.BarMenu.Bands("Band_PopUp").Tools("Mnu_RunApplication").ShortCuts = aryShortCut
        
        Case "Mnu_SearchObject"
            sStr = gResource.GetMessage(MNU_SEARCH)
            aryShortCut(1).Value = "Control+R"
            frmMain.BarMenu.Bands("Band_PopUp").Tools("Mnu_SearchObject").ShortCuts = aryShortCut
            
    End Select
    
    GetCaption4MenuBar = sStr
End Function


'**+
'Nome:   RefreshDescriptions4StatusBar
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Reimposta i messaggi da visualizzare sulla StatusBar per quelle
'              voci che dipendono dalla modalità di visualizzazione (Form/Tabella).
'
'**/
Private Sub RefreshDescriptions4StatusBar()
    'ATTENZIONE:
    'Inserire qui tutte le voci di menu ed i pulsanti della toolbar per i quali si
    'vuole cambiare il suggerimento sulla StatusBar in funzione della modalità di
    'visualizzazione. Ad esempio è possibile avere dei messaggi al SINGOLARE per
    'la modalità form e PLURALE per la modalità tabellare.
    'La funzione GetDescription4StatusBar si occupa di determinare la frase esatta.
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Description = GetDescription4StatusBar("Mnu_PrePrint")
    BarMenu.Bands("Band_File").Tools("Mnu_Print").Description = GetDescription4StatusBar("Mnu_Print")
    BarMenu.Bands("Standard").Tools("Print").Description = GetDescription4StatusBar("Print")
    BarMenu.Bands("Standard").Tools("PrePrint").Description = GetDescription4StatusBar("PrePrint")
    BarMenu.Bands("Standard").Tools("ChangeView").Description = GetDescription4StatusBar("ChangeView")
    BarMenu.Bands("Band_Export").Tools("ExportWord").Description = GetDescription4StatusBar("ExportWord")
    BarMenu.Bands("Band_Export").Tools("ExportExcel").Description = GetDescription4StatusBar("ExportExcel")
    BarMenu.Bands("Band_Export").Tools("ExportHtml").Description = GetDescription4StatusBar("ExportHtml")
    BarMenu.Bands("Band_Export").Tools("ExportPDF").Description = GetDescription4StatusBar("ExportPDF")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Description = GetDescription4StatusBar("ExportWord")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Description = GetDescription4StatusBar("ExportExcel")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Description = GetDescription4StatusBar("ExportHtml")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Description = GetDescription4StatusBar("ExportPDF")

End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 25/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: Caption2Display
'
'Parametri:
'  Boolean ReadFromGrid - determina se le stringhe per la costruzione della caption devono essere lette direttamente
'  dai campi del documento o dalla collection AllColumns della BrwMain.
'
'Valori di ritorno: String
'
'Funzionalità:
'                  ///////////////////////////////////////////////////////////////////////////////////////////////////////
'                  In questa funzione va inserito il codice per la determinazione della caption del form principale
'                  per le modalità Modify e Browse in base alle esigenze specifiche.
'                  Di default viene usato esclusivamente il campo del documento individuato dalla
'                  costante CAMPO_PER_CAPTION.
'                  ///////////////////////////////////////////////////////////////////////////////////////////////////////
'
'**/
Private Function Caption2Display(Optional ByVal ReadFromGrid As Boolean) As String
    If Not m_Document.EOF And Not m_Document.BOF Then
        If Not ReadFromGrid Then
            
            Caption2Display = m_App.Caption & ": " & fnNotNull(m_Document.Fields(CAMPO_PER_CAPTION).Value) & " [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
        Else
            
            Caption2Display = m_App.Caption & ": " & fnNotNull(BrwMain.AllColumns(CAMPO_PER_CAPTION).Value) & " [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
        End If
    Else
        Caption2Display = m_App.FunctionName & " [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    End If
End Function



'**+
'Nome :SetStatus4Modality
'
'Parametri:NewModality rappresenta la modalità di visualizzazione
'          che si vuole ottenere.
'          ModePreview è uno switch per apertura o chiusura anteprima di stampa.
'
'Valori di ritorno:
'
'Funzionalità: Abilita i pulsanti della Toolbar e le voci di menu in funzione
'              di una determinata modalità di visualizzazione.
'              (disabilita tutti i rimanenti pulsanti e voci di menu)
'              Imposta la Caption del form in funzione della modalità di visualizzazione
'**/
Private Sub SetStatus4Modality(ByVal NewModality As neVisualModality, _
                                Optional ByVal ModePreview As nePreviewModality)
    Dim KeyON As Currency
    Dim KeyOFF As Currency
    Dim iPicture As Integer
   
    
    'Indica lo stato di visibilità della ToolBar standard
    'prima della visualizzazione della ToolBar della anteprima
    'di stampa
    Static bToolBarStandardVisible As Boolean
    
    'Indica lo stato di attivazione dei bottoni della ToolBar
    'standard prima della visualizzazione della ToolBar della
    'anteprima di stampa
    Static curToolBarStandardStatus As Currency
    
    
    'Elimina l'acceleratore CUT
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Caption = gResource.GetMessage(MNU_DELETE)
    'Rimuove lo shortcut "Delete"
    aryShortCut(1).Clear
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").ShortCuts = aryShortCut
    
    'Imposta i pulsanti e le voci di menu
    Select Case NewModality
    
        Case Insert
            KeyOFF = BTN_SAVE + BTN_PRINT + BTN_PREVIEW + BTN_DELETE + BTN_SEARCH
            KeyOFF = KeyOFF + BTN_SEARCHTABLE + BTN_SEARCHFORM + BTN_VIEWMODE
            KeyOFF = KeyOFF + BTN_FILTER
            KeyOFF = KeyOFF + BTN_PREVIOUS + BTN_NEXT
            KeyOFF = KeyOFF + BTN_WORD + BTN_EXCEL + BTN_HTML + BTN_PDF
            KeyON = BTN_ALL - KeyOFF
            Me.Caption = m_App.Caption
            oFiltersActivity.AbortNewFilter
            
            If BrwMain.GuiMode = dgFilterDefinition Then
                bEnableGuiEvent = False
                BrwMain.GuiMode = dgNormal
                bEnableGuiEvent = True
            End If
            
            m_Search = False
            
        Case Modify
            KeyOFF = BTN_SAVE + BTN_CLEAR + BTN_SEARCH + BTN_SEARCHFORM
            KeyON = BTN_ALL - KeyOFF
            'in modalità variazione si è necessariamente in modalità form
            'pertanto il pulsante ChangeView della toolbar deve visualizzare
            'l'icona della griglia
            iPicture = IIf(GetSetting(REGISTRY_KEY, TheApp.Name & "Settings", "LargeIcon", False), IDB_STD_GRID32, IDB_STD_GRID16)
            BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
            BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
            If Not (m_Document.EOF = True Or m_Document.BOF = True) Then
                'Monta la caption del form principale
                Me.Caption = Caption2Display(False)
            End If
            oFiltersActivity.AbortNewFilter
                        
            m_Search = False
            
        Case Find
        
            'Solo se esiste almeno un elemento nel data manager.
            If Not (m_Document.EOF = True And m_Document.BOF = True) Then
                KeyON = BTN_VIEWMODE + BTN_SEARCHTABLE + BTN_SEARCHFORM
            End If
            KeyON = KeyON + BTN_NEW + BTN_CUT + BTN_COPY + BTN_PASTE
            KeyON = KeyON + BTN_CLEAR + BTN_SEARCH
            KeyOFF = BTN_ALL - KeyON
            'In modalità Find verrà proposto il pulsante per andare in modalità tabella
            'pertanto il pulsante ChangeView della toolbar deve visualizzare
            'l'icona della griglia
            iPicture = IIf(GetSetting(REGISTRY_KEY, TheApp.Name & "Settings", "LargeIcon", False), IDB_STD_GRID32, IDB_STD_GRID16)
            BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
            BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
            BarMenu.Bands("Standard").Tools("ChangeView").Description = GetDescription4StatusBar("ChangeView")
            Me.Caption = gResource.GetMessage(TT_SEARCH) & " - " & m_App.Caption
            
            oFiltersActivity.AbortNewFilter
                
            'Cancella eventuali blocchi su qualsiasi azione.
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            
        Case Browse
            KeyOFF = BTN_SAVE + BTN_CLEAR + BTN_SEARCH + BTN_PREVIOUS + BTN_NEXT
            KeyOFF = KeyOFF + BTN_SEARCHTABLE + BTN_CUT + BTN_COPY + BTN_PASTE
            KeyON = BTN_ALL - KeyOFF
            'Seleziona l'icona grande o piccola in base alle impostazioni correnti
            iPicture = IIf(GetSetting(REGISTRY_KEY, TheApp.Name & "Settings", "LargeIcon", False), IDB_STD_FORM32, IDB_STD_FORM16)
            BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
            BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
            If Not (m_Document.EOF = True Or m_Document.BOF = True) Then
                'Monta la caption del form principale
                Me.Caption = Caption2Display(False)
            End If
            'Inserisce l'acceleratore CUT
            BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Caption = gResource.GetMessage(MNU_DELETE)
            'Inserisce lo shortcut "Delete"
            aryShortCut(1).Value = "Delete"
            BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").ShortCuts = aryShortCut
            
            'Questo controllo si è reso necessario per evitare un loop infinito
            'con la gestione dell'evento BrwMain_OnChangeGuiMode() quando dal
            'Menu della browse si va in modalità tabellare.
            If BrwMain.GuiMode <> dgNormal Then
                BrwMain.GuiMode = dgNormal
            End If
            
            'Se il filtro attivo è un filtro temporaneo viene abilitato il pulsante
            'Salva Filtro del DocTypeExplorer per poterlo rendere permanente.
            If m_ActiveFilter.ID = -1 Then
                oFiltersActivity.NewFilterBegin   'Abilita il pulsante Salva Filtro
            Else
                oFiltersActivity.AbortNewFilter   'Disabilita il pulsante Salva Filtro
            End If
            ActivityBox.Redraw = True
            
            m_Search = False
            
            'Cancella eventuali blocchi su qualsiasi azione.
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            
            
        Case Preview
            If ModePreview = OpenPrw Then
                bToolBarStandardVisible = BarMenu.Bands("Standard").Visible
                curToolBarStandardStatus = GetStatusToolBar(True)
                KeyON = BTN_PRINT + BTN_EXCEL + BTN_WORD + BTN_HTML + BTN_PDF
                KeyOFF = BTN_ALL - KeyON
                BarMenu.Bands("Band_View").Tools("Mnu_Folders").Enabled = False
                BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Enabled = False
                BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Enabled = False
                BarMenu.Bands("Standard").Visible = False
                BarMenu.Bands(BAND_CLOSE_PREVIEW).Visible = True
                BarMenu.RecalcLayout
            Else
                BarMenu.Bands(BAND_CLOSE_PREVIEW).Visible = False
                BarMenu.Bands("Standard").Visible = bToolBarStandardVisible
                ActivateBarButtons curToolBarStandardStatus, True
                ActivateBarButtons BTN_ALL - curToolBarStandardStatus, False
                BarMenu.Bands("Band_View").Tools("Mnu_Folders").Enabled = True
                BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Enabled = True
                BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Enabled = True
                BarMenu.RecalcLayout
            End If
            
    End Select
    
    'Attiva/disattiva i pulsanti e le voci di menu
    ActivateBarButtons KeyON, True
    ActivateBarButtons KeyOFF, False
End Sub


'**+
'Autore                 : Diamante S.p.a
'
'Nome                   : PermissionToSave
'
'Parametri:
'
'Valori di ritorno: True se il documento può essere salvato, False altrimenti.
'
'Funzionalità: Controlli da effettuare PRIMA di salvare il documento corrente
'
'**/
Private Function PermissionToSave() As Boolean
Dim Testo As String


    PermissionToSave = True
    
    If Me.cboTipoImpostazione.CurrentID = 0 Then
        MsgBox "Manca il tipo impostazione del contratto", vbInformation, "Impossibile salvare"
        PermissionToSave = False
        If Me.cboTipoImpostazione.Enabled = True Then Me.cboTipoImpostazione.SetFocus
        Exit Function
    End If
    
    If Me.CDCliente.KeyFieldID = 0 Then
        MsgBox "Manca l'anagrafica intestataria del contratto", vbInformation, "Impossibile salvare"
        PermissionToSave = False
        If Me.CDCliente.Enabled = True Then Me.CDCliente.SetFocus
        Exit Function
        
    End If
    
    
    
    If IDClienteFatturazione = 0 Then
        MsgBox "Manca l'anagrafica di fatturazione  del contratto", vbInformation, "Impossibile salvare"
        PermissionToSave = False
        
        Exit Function
    End If
    
    If Me.cboTipoImpostazione.CurrentID <= 2 Then
        If Me.chkContrattoAttuale.Value = vbUnchecked Then
            Testo = "Attenzione!!!!" & vbCrLf
            Testo = Testo & "Il documento non risulta essere il contratto attuale, modificando questi dati potrebbero esserci delle discordanze con i rinnovi successivi già effettuati" & vbCrLf
            Testo = Testo & "Vuoi continuare?"
            
            'MsgBox "Impossibile salvare poichè non risulta che il contratto sia quello attuale", vbInformation, "Impossibile salvare"
            If MsgBox(Testo, vbQuestion + vbYesNo, TheApp.FunctionName) = vbNo Then
            
                PermissionToSave = False
                Exit Function
            End If
        End If
        
        If Me.cboTipoContratto.CurrentID = 0 Then
            MsgBox "Inserire il tipo contratto", vbInformation, "Impossibile salvare"
            PermissionToSave = False
            Me.cboTipoContratto.SetFocus
            Exit Function
        End If
        
        If Me.txtDataDecorrenza.Value = 0 Then
            MsgBox "Manca la data di decorrenza del contratto", vbInformation, "Impossibile salvare"
            PermissionToSave = False
            Me.txtDataDecorrenza.SetFocus
            Exit Function
        End If
    
        If Me.cboDurataContratto.CurrentID = 0 Then
            MsgBox "Manca il tipo di durata del contratto", vbInformation, "Impossibile salvare"
            PermissionToSave = False
            cboDurataContratto.SetFocus
            Exit Function
        End If
    
        If Me.cboTipoRateizzazione.CurrentID = 0 Then
            MsgBox "Manca il tipo di rateizzazione del contratto", vbInformation, "Impossibile salvare"
            PermissionToSave = False
            cboTipoRateizzazione.SetFocus
            Exit Function
        End If
    
        If Me.cboTipoRinnovo.CurrentID = 0 Then
            MsgBox "Manca il tipo di rinnovo del contratto", vbInformation, "Impossibile salvare"
            PermissionToSave = False
            cboTipoRinnovo.SetFocus
            Exit Function
        End If
    
        If Me.txtDataScadenza.Value = 0 Then
            MsgBox "Manca la data di scadenza del contratto", vbInformation, "Impossibile salvare"
            PermissionToSave = False
            txtDataScadenza.SetFocus
            Exit Function
        End If
    
        If Me.txtDataScadenzaPerRinnovo.Value = 0 Then
            MsgBox "Manca la data di scadenza di rinnovo del contratto", vbInformation, "Impossibile salvare"
            PermissionToSave = False
            txtDataScadenzaPerRinnovo.SetFocus
            Exit Function
        End If
        
        If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
            If Me.cboDurataAssistenza.CurrentID = 0 Then
                MsgBox "Manca il tipo durata assistenza", vbInformation, "Impossibile salvare"
                PermissionToSave = False
                cboDurataAssistenza.SetFocus
                Exit Function
            End If
        
            If Me.txtDataFineAssistenza.Value = 0 Then
                MsgBox "Manca la data di fine assistenza", vbInformation, "Impossibile salvare"
                PermissionToSave = False
                txtDataFineAssistenza.SetFocus
                Exit Function
            End If
        End If
        
        If IDClienteFatturazione = 0 Then
            MsgBox "Inserire il cliente di fatturazione", vbInformation, "Impossibile salvare"
            PermissionToSave = False
            
            Exit Function
        End If
        
        If Me.txtDataDecorrenza.Value >= Me.txtDataScadenza.Value Then
            MsgBox "La data di decorrenza è maggiore della data di scadenza del contratto", vbInformation, "Impossibile salvare"
            PermissionToSave = False
            
            Exit Function
        End If
        If Me.txtDataDecorrenza.Value >= Me.txtDataScadenzaPerRinnovo.Value Then
            MsgBox "La data di decorrenza è maggiore della data di scadenza del rinnovo del contratto", vbInformation, "Impossibile salvare"
            PermissionToSave = False
            
            Exit Function
        End If
        If Me.txtDataDecorrenza.Value >= Me.txtDataFineAssistenza.Value Then
            MsgBox "La data di decorrenza è maggiore della data di scadenza dell'assistenza", vbInformation, "Impossibile salvare"
            PermissionToSave = False
            
            Exit Function
        End If
        
        If Me.cboPagamentoRate.CurrentID = 0 Then
            MsgBox "Manca il tipo di pagamento delle rate del contratto", vbInformation, "Impossibile salvare"
            PermissionToSave = False
            cboPagamentoRate.SetFocus
            Exit Function
        End If
    End If
    
    

End Function


'**+
'Nome: SearchNext
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Posizionamento al record successivo
'**/
Private Sub SearchNext()
    
    m_Document.MoveNext
    
    If m_Document.EOF Then
        'Si era già sull'ultimo record (prima di MoveNext).
        
        'Si annulla l'operazione
        m_Document.MovePrevious
        sbMsgInfo gResource.GetMessage(MESS_NO_NEXT_ELEMENTS), m_App.FunctionName
        Exit Sub
    Else
        'Controlla la presenza di eventuali conflitti nel caso di multiutenza.
        
        If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
            m_Document.MovePrevious
        Else
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
        End If
    End If
    
End Sub

'**+
'Nome: SearchPrevious
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Posizionamento al record precedente
'**/
Private Sub SearchPrevious()

    m_Document.MovePrevious
    
    If m_Document.BOF Then
        'Si era già sul primo record (prima di MovePrevious).
        
        'Si annulla l'operazione
        m_Document.MoveNext
        sbMsgInfo gResource.GetMessage(MESS_NO_PREVIOUS_ELEMENTS), m_App.FunctionName
        Exit Sub
    Else
        'Controlla la presenza di eventuali conflitti nel caso di multiutenza.
        
        If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
            m_Document.MoveNext
        Else
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
        End If
    End If
End Sub

'**+
'Nome: BrowseReposition
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni da compiere al riposizionamento del record corrente
'**/
Private Sub BrowseReposition()

    'Dopo un Save del documento avviene un Refresh della Browse ma in tal caso
    'è inutile effettuare il refresh del form.
    If Not m_bAvoidReposition Then
    
        'Refresh dei campi del form
        RefreshFormFields
        
        'Refresh della caption del Form
        If Not (m_Document.EOF = True Or m_Document.BOF = True) Then
            'Monta la caption del form principale
            Me.Caption = Caption2Display(False)
        End If
        
    End If
 
    'Refresh delle variabili di stato
    m_Changed = False
    m_Saved = False
    m_Search = False
    
    'Annullamento di un eventuale inizio di inserimento di un nuovo record
    m_Document.AbortNew
    
End Sub



'**+
'Nome: NewRecord
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni su richiesta nuovo record
'**/
Private Sub NewRecord()
 BLoading = 1

'--------------------------------------------------------------------------------------------
'NOTA:
'Il gruppo di istruzioni sottostanti e la riga  'Imposta il blocco su inserimento'
'sono state commentate per far si che la manutenzione NON imposti alcun blocco per
'l'azione Inserimento.
'Pertanto 2 o più utenti potranno effettuare contemporaneamente la suddetta azione.
'Se si intende impedire questa possibilità sarà sufficiente ripristinare le righe commentate.
'--------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
'    'Controllo se ho il permesso di salvare ( nel caso di conflitti di multiutenza )
'    If Not m_Semaphore.IsActionAvailable(m_DocType.ID, SemAllObjects, SemInsertAction) Then
'        'C'è un altro utente in modalità inserimento che blocca la medesima azione per
'        'tutti gli altri utenti. Pertanto annullo l'operazione di inserimento ed esco.
'        Exit Sub
'    End If




    'Ho il permesso per l'azione inserimento.
    '
    'Cancella il blocco precedente
    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
    
    '--------------------------------
    'Imposta il blocco su inserimento
'    m_Semaphore.SetObjectAction m_DocType.ID, SemAllObjects, SemInsertAction
    '
    'A questo punto nessun altro utente potrà effettuare una operazione di inserimento
    'finchè non verrà cancellato il blocco su inserimento.


    'Annulla una eventuale operazione precedente.
    If m_Document.TableNew Then
        m_Document.AbortNew
    End If

    'Creazione buffers vuoti
    m_Document.NewDoc
    
    'Refresh delle variabili di stato
    m_Search = False
    m_Changed = False
    m_Saved = False
    
    'Refresh della toolbar in modalità inserimento
    SetStatus4Modality Insert
    
    'Ripristina la vista del Form
    
    Me.txtIDContrattoPadre.Value = 0
    Me.txtNumeroRinnovo.Value = 1
'    Me.cboUtenteInserimento.WriteOn TheApp.IDUser
'    Me.cboUtenteModifica.WriteOn TheApp.IDUser
'    Me.txtDataInserimento.Value = Date
'    Me.txtDataModifica.Value = Date
    Me.cboTipoImpostazione.WriteOn LINK_TIPO_IMPOSTAZIONE
    cboTipoImpostazione_Click
'    Me.chkContrattoAttuale.Value = vbChecked
'    Me.chkAttivoPassivo.Value = vbChecked
'    Me.chkAdeguamentoIstat.Value = vbChecked
'    Me.chkRinnovoAutomatico.Value = vbChecked
    
    
    ActivateBarButtons BTN_SAVE, False

    m_Changed = False
    'Me.cboContrattoBancario.SetFocus
    SetFocusTabIndex0
    
    BLoading = 0
    BrwMain.Visible = False
    DoEvents
End Sub

'**+
'Nome                   : ClearControl
'
'Parametri              : ctrControl As Control - controllo da pulire
'
'Valori di ritorno      :
'
'Funzionalità           : Pulisce un controllo sulla base del tipo del controllo stesso
'
'**/
Private Sub ClearControl(ByVal ctrControl As Control)
    Dim sType As String

    sType = TypeName(ctrControl)
    
    If sType = "fpDateTime" Or sType = "TextBox" Or sType = "fpText" Or sType = "fpLongInteger" Or sType = "fpCurrency" Or sType = "fpDoubleSingle" Or sType = "dmtDate" Then
        ctrControl.Text = ""
    ElseIf sType = "CheckBox" Then
        ctrControl.Value = 0
    ElseIf sType = "fpBoolean" Then
        ctrControl.Value = 0
    ElseIf sType = "ComboBox" Then
        ctrControl.ListIndex = -1
    ElseIf sType = "DMTCombo" Then
        ctrControl.WriteOn 0
    ElseIf sType = "ListBox" Then
        ctrControl.ListIndex = -1
    ElseIf sType = "ListView" Then
        ctrControl.ListItems.Clear
    ElseIf sType = "TreeView" Then
        ctrControl.Nodes.Clear
    ElseIf sType = "Town" Then
        ctrControl.Reset
    ElseIf sType = "DmtSearchACS" Then
        ctrControl.IDAnagrafica = 0
        ctrControl.Description = ""
        ctrControl.SecondDescription = ""
        
    ElseIf sType = "dmtCurrency" Or sType = "dmtNumber" Then
        ctrControl.Value = 0
        ctrControl.Text = ""
    ElseIf sType = "DmtSearchACS" Then
        ctrControl.IDNode = 0
    ElseIf sType = "DmtFirmGerarchy" Then
        ctrControl.LoadActivity 0
    ElseIf sType = "DMTProgControl" Then
        'Queste istruzioni forzano il refresh
        'e il reset del componente
        ctrControl.IDArticolo = 0
        ctrControl.Show
    ElseIf sType = "DmtCodDesc" Then
        ctrControl.Load 0
    End If
End Sub


'**+
'Nome: ClearFormFields
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Pulisce il contenuto dei campi di input del Form
'**/
Private Sub ClearFormFields()
    Dim cField As FormField
    
    For Each cField In m_FormFields
        'Viene ripulito il campo di immissione.
        ClearControl cField.Control
    Next
End Sub

'**+
'Nome: ExecuteMenuCommand
'
'Parametri:
'sToolName - Nome del comando selezionato
'
'Valori di ritorno:
'
'Funzionalità:
'Gestione dei comandi generati dal controllo ActiveBar
'**/
Private Sub ExecuteMenuCommand(ByVal sToolName As String)
    Dim iAnswer As Integer

    'cbcxn
    'Notifica alla (eventuale) applicazione che gestisce il processo On_Extend la
    'pressione di un Tool. Se l'applicazione chiamata restituisce True viene annullata l'operazione.
    'If m_ExtendApplication.BeforeCommandClick(sToolName) Then Exit Sub

    Select Case sToolName
        Case "Cut", "Mnu_Cut"
            SendKeys ("+{DEL}")
            
        Case "Copy", "Mnu_Copy"
            SendKeys ("^{INSERT}")
            
        Case "Paste", "Mnu_Paste"
            SendKeys ("+{INSERT}")
            
        Case "Mnu_Folders"
            OnFolders
            
        Case "ClosePreview"
            ClosePreview
            
        Case "Save", "Mnu_Save"
            OnSave
            
        Case "Mnu_Exit"
            Unload frmMain
            
        Case "Delete", "Mnu_Delete"
            OnDelete
            
        Case "Clear", "Mnu_Clear"
            OnClear
            
        Case "ExecuteSearch", "Mnu_ExecuteSearch"
            OnExecuteSearch
        
        Case "SearchNext", "Mnu_SearchNext"
            
            OnMoveCurrentRecord SRCNEXT, sToolName
        
        Case "SearchPrevious", "Mnu_SearchPrevious"
            
            OnMoveCurrentRecord SRCPREVIOUS, sToolName
        
        Case "ChangeView", "Mnu_FormView", "Mnu_TableView"
            OnChangeView sToolName
            
        Case "Mnu_ToolBar"
            OnToolBarOptions
            
        Case "Mnu_Options"
            OnOptions
            
        Case "Mnu_Info"
            OnInfo
            
        Case "PrePrint", "Mnu_PrePrint", "Print", "Mnu_Print", "ExportPDF", "Mnu_ExportPDF", "MailPDF", "ExportWord", "Mnu_ExportWord", "MailWord", "ExportExcel", "Mnu_ExportExcel", "MailExcel", "ExportHtml", "Mnu_ExportHtml", "MailMHTL"
            OnPrint sToolName
            
        Case "NewSearch", "Mnu_NewSearch", "Mnu_SearchFilter"
            OnNewSearch
            
        Case "New", "Mnu_New"
            
            OnNew sToolName
           
        Case "Mnu_RunApplication", "Mnu_SearchObject"
            OnRunApplication sToolName
        Case "Mnu_Summary"
            OnSummary
        Case "Mnu_FastHelp", "Help"
            OnFastHelp
        Case "Mnu_HelpOnLine"
            OnHelpOnLine
        Case "Mnu_Arg"
             OnArg
        Case "Mnu_Web1"
             sbOpenURL hwnd, URL_DIAMANTE
        
        Case "cmdStoricoRate"
            cmdRateStorico_Click
        Case "cmdStoriaAdeg"
            cmdStoriaAdeguamenti_Click
        Case "cmdDocumentazione"
            cmdDocumentazione_Click
        Case "cmdClausole"
            
        Case "cmdAltreInfo"
            cmdApri_Click
        Case "cmdStampaModelliTesta"
            Command4_Click
        Case "cmdAdeguamentoIstat"
            CALCOLA_ISTAT_CONTRATTO
        Case "cmdFatturazione"
            AVVIA_FATTURAZIONE_CONTRATTO fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
         Case "cmdGeneraRate"
            cmdGeneraRate_Click
         Case "RichiediListaInterventi"
            RICHIEDI_INTERVENTI
    End Select
    
    
    
    'cbcxn
    'Notifica alla (eventuale) applicazione che gestisce il processo On_Extend la
    'pressione di un Tool DOPO avere eseguito l'operazione ad esso associata
    'm_ExtendApplication.AfterCommandClick sToolName
    
End Sub

'**+
'Nome: RefreshFormFields
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Riempie i valori dei campi del Form con i valori
'del documento.
'**/
Private Sub RefreshFormFields()

    'rif3 start
    
    Dim Fields As DmtDocManLib.Fields
    Dim Control As Control
    Dim Field As FormField
    
    On Error Resume Next
    
    'In questi casi non si deve far nulla
    If Not (m_Document.EOF = True Or m_Document.BOF = True) Then

       'Passa alla collezione Fields dell'oggetto
        'Document i valori da salvare
        For Each Field In m_FormFields
    
            Select Case TypeName(Field.Control)
                Case "TextBox"
                    Field.Control.Text = fnNotNull(m_Document.Fields(Field.Name).Value)
                Case "DMTCombo"
                    Field.Control.WriteOn fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "Town"
                    If Field.Name = "IDComune" Then
                        Field.Control.TownID = fnNotNullN(m_Document.Fields(Field.Name).Value)
                    ElseIf Field.Name = "Cap" Then
                        Field.Control.Zip = fnNotNull(m_Document.Fields(Field.Name).Value)
                    End If
                Case "dmtDate"
                    Field.Control.Value = fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "dmtNumber"
                    Field.Control.Value = fnNotNullN((m_Document.Fields(Field.Name).Value))
                    
                Case "dmtCurrency"
                    Field.Control.Value = fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "dmtTime"
                    Field.Control.Text = fnNormDate(m_Document.Fields(Field.Name).Value)
                Case "DmtSearchACS"
                        Field.Control.Description = fnNotNull(m_Document.Fields("Anagrafica").Value)
                        Field.Control.SecondDescription = fnNotNull(m_Document.Fields("Nome").Value)
                        Field.Control.IDAnagrafica = fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "CheckBox"
                    Field.Control.Value = Abs(fnNotNullN(m_Document.Fields(Field.Name).Value))
                Case "DmtCodDesc"
                    Field.Control.Load fnNotNullN(m_Document.Fields(Field.Name).Value)
            End Select
           
        Next

  
    End If

    'rif3 end
    
End Sub

'**+
'Nome: ClearFields
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Azzera i valori dei campi del documento
'**/
Private Sub ClearFields()
    Dim Field As DmtDocManLib.Field
    
    For Each Field In m_Document.Fields
        Field.Value = Empty
    Next
End Sub

'**+
'Nome: Change
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni su variazione di un campo del Form
'**/
Private Sub Change()
    'Se si è in modalità tabellare non deve essere eseguita perchè
    'altrimenti al Click della Browse si attiverebbe il pulsante Salva
    If Not m_Search And Not BrwMain.Visible Then
        ActivateBarButtons BTN_SAVE, True
    
        m_Changed = True
        m_Saved = False
        m_Search = False
    End If
End Sub

'**+
'Nome: CreateFormFields
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Crea la collezione FormFields che associa i campi del
'documento con i controlli di input del Form. Vengono
'anche creati i controlli del Form necessari e calcolato
'il layout del Form.
'**/
Private Sub CreateFormFields()
    Dim Field As FormField
        
        
    'rif2 start
    
    'Se non esiste il documento aperto non si può creare la collezione
    If m_Document Is Nothing Then Exit Sub
    
    'Se la collezione è già stata creata esce
    If Not m_FormFields Is Nothing Then Exit Sub
    
    'Istanzia la collezione.  Il codice sottostante viene eseguito soltanto la prima volta
    Set m_FormFields = New FormFields
    
    'rif2   End
    
    
    'Anagrafica
    Set Field = New FormField
    Set Field.Control = Me.CDCliente
    Field.Name = "IDAnagrafica"
    Field.Visible = True
    Me.CDCliente.Tag = Field.Name
    m_FormFields.Add Field

    'Tipo contratto
    Set Field = New FormField
    Set Field.Control = Me.cboTipoContratto
    Field.Name = "IDTipoContratto"
    Field.Visible = True
    Me.cboTipoContratto.Tag = Field.Name
    m_FormFields.Add Field
    
    'Data stipula
    Set Field = New FormField
    Set Field.Control = Me.txtDataStipula
    Field.Name = "DataStipula"
    Field.Visible = True
    Me.txtDataStipula.Tag = Field.Name
    m_FormFields.Add Field
    

   
    'Data decorrenza
    Set Field = New FormField
    Set Field.Control = Me.txtDataDecorrenza
    Field.Name = "DataDecorrenza"
    Field.Visible = True
    Me.txtDataDecorrenza.Tag = Field.Name
    m_FormFields.Add Field
    
    'Data scadenza
    Set Field = New FormField
    Set Field.Control = Me.txtDataScadenza
    Field.Name = "DataScadenza"
    Field.Visible = True
    Me.txtDataScadenza.Tag = Field.Name
    m_FormFields.Add Field
    
    'Disdetta
    Set Field = New FormField
    Set Field.Control = Me.chkDisdetta
    Field.Name = "Disdetta"
    Field.Visible = True
    Me.chkDisdetta.Tag = Field.Name
    m_FormFields.Add Field
    
    'Data disdetta
    Set Field = New FormField
    Set Field.Control = Me.txtDataDisdetta
    Field.Name = "DataDisdetta"
    Field.Visible = True
    Me.txtDataDisdetta.Tag = Field.Name
    m_FormFields.Add Field
    
    'AttivoPassivo
    Set Field = New FormField
    Set Field.Control = Me.chkAttivoPassivo
    Field.Name = "AttivoPassivo"
    Field.Visible = True
    Me.chkAttivoPassivo.Tag = Field.Name
    m_FormFields.Add Field
    
    'Rinnovo automatico
    Set Field = New FormField
    Set Field.Control = Me.chkRinnovoAutomatico
    Field.Name = "RinnovoAutomatico"
    Field.Visible = True
    Me.chkRinnovoAutomatico.Tag = Field.Name
    m_FormFields.Add Field

    'Adeguamento Istat
    Set Field = New FormField
    Set Field.Control = Me.chkAdeguamentoIstat
    Field.Name = "AdeguamentoIstat"
    Field.Visible = True
    Me.chkAdeguamentoIstat.Tag = Field.Name
    m_FormFields.Add Field

    'Non fatturare
    Set Field = New FormField
    Set Field.Control = Me.chkNonFatturare
    Field.Name = "NonFatturare"
    Field.Visible = True
    Me.chkNonFatturare.Tag = Field.Name
    m_FormFields.Add Field

    'IDRateizzazione
    Set Field = New FormField
    Set Field.Control = Me.cboTipoRateizzazione
    Field.Name = "IDRateizzazione"
    Field.Visible = True
    Me.cboTipoRateizzazione.Tag = Field.Name
    m_FormFields.Add Field
        
    'IDPagamentoRate
    Set Field = New FormField
    Set Field.Control = Me.cboPagamentoRate
    Field.Name = "IDPagamentoRate"
    Field.Visible = True
    Me.cboPagamentoRate.Tag = Field.Name
    m_FormFields.Add Field
        
    'IDDurataContratto
    Set Field = New FormField
    Set Field.Control = Me.cboDurataContratto
    Field.Name = "IDDurataContratto"
    Field.Visible = True
    Me.cboDurataContratto.Tag = Field.Name
    m_FormFields.Add Field
        
    'IDTipoRinnovo
    Set Field = New FormField
    Set Field.Control = Me.cboTipoRinnovo
    Field.Name = "IDTipoRinnovo"
    Field.Visible = True
    Me.cboTipoRinnovo.Tag = Field.Name
    m_FormFields.Add Field
        
        
    'IDSitoPerAnagrafica
    Set Field = New FormField
    Set Field.Control = Me.cboSitoPerAnagrafica
    Field.Name = "IDSitoPerAnagrafica"
    Field.Visible = True
    Me.cboSitoPerAnagrafica.Tag = Field.Name
    m_FormFields.Add Field
    
    'Importo contratto stipula
    Set Field = New FormField
    Set Field.Control = Me.txtImportoStipula
    Field.Name = "ImportoContrattoStipula"
    Field.Visible = True
    Me.txtImportoStipula.Tag = Field.Name
    m_FormFields.Add Field
    
    'Importo contratto attuale
    Set Field = New FormField
    Set Field.Control = Me.txtImportoAttuale
    Field.Name = "ImportoContrattoAttuale"
    Field.Visible = True
    Me.txtImportoAttuale.Tag = Field.Name
    m_FormFields.Add Field
    
    'Data scadenza per rinnovo
    Set Field = New FormField
    Set Field.Control = Me.txtDataScadenzaPerRinnovo
    Field.Name = "DataScadenzaPerRinnovo"
    Field.Visible = True
    Me.txtDataScadenzaPerRinnovo.Tag = Field.Name
    m_FormFields.Add Field
    
    'Annotazioni
    Set Field = New FormField
    Set Field.Control = Me.txtAnnotazioni
    Field.Name = "Annotazioni"
    Field.Visible = True
    Me.txtAnnotazioni.Tag = Field.Name
    m_FormFields.Add Field
    
    'Note fatturazione
    Set Field = New FormField
    Set Field.Control = Me.txtNoteFatturazione
    Field.Name = "NoteFattura"
    Field.Visible = True
    Me.txtNoteFatturazione.Tag = Field.Name
    m_FormFields.Add Field



    'IDentificativo dell'Anagrafica agente
    Set Field = New FormField
    Set Field.Control = Me.CDAgente
    Field.Name = "IDAnagraficaAgente"
    Field.Visible = True
    Me.CDAgente.Tag = Field.Name
    m_FormFields.Add Field

    'Nome agente
    'Set Field = New FormField
    'Set Field.Control = Me.txtNomeAgente
    'Field.Name = "NomeAgente"
    'Field.Visible = True
    'Me.txtNomeAgente.Tag = Field.Name
    'm_FormFields.Add Field

    'Identificativo dell'anagrafica commesso
    Set Field = New FormField
    Set Field.Control = Me.CDTecnico
    Field.Name = "IDAnagraficaCommesso"
    Field.Visible = True
    Me.CDTecnico.Tag = Field.Name
    m_FormFields.Add Field

    'Identificativo dell'anagrafica commesso
    Set Field = New FormField
    Set Field.Control = Me.CDAmministratore
    Field.Name = "IDAnagraficaAmministratore"
    Field.Visible = True
    Me.CDAmministratore.Tag = Field.Name
    m_FormFields.Add Field
    
    'Nome commesso
    'Set Field = New FormField
    'Set Field.Control = Me.txtNomeTecnico
    'Field.Name = "NomeCommesso"
    'Field.Visible = True
    'Me.txtNomeTecnico.Tag = Field.Name
    'm_FormFields.Add Field

    'Numero protocollo
    Set Field = New FormField
    Set Field.Control = Me.txtNumeroProtocollo
    Field.Name = "NumeroProtocollo"
    Field.Visible = True
    Me.txtNumeroProtocollo.Tag = Field.Name
    m_FormFields.Add Field



    'Identificativo della durata assistenza
    Set Field = New FormField
    Set Field.Control = Me.cboDurataAssistenza
    Field.Name = "IDRV_POTipoDurataAssistenza"
    Field.Visible = True
    Me.cboDurataAssistenza.Tag = Field.Name
    m_FormFields.Add Field

    'Data fine assistenza
    Set Field = New FormField
    Set Field.Control = Me.txtDataFineAssistenza
    Field.Name = "DataFineAssistenza"
    Field.Visible = True
    Me.txtDataFineAssistenza.Tag = Field.Name
    m_FormFields.Add Field

    'Ritenuta di acconto
    Set Field = New FormField
    Set Field.Control = Me.chkRitAcconto
    Field.Name = "RitenutaAcconto"
    Field.Visible = True
    Me.chkRitAcconto.Tag = Field.Name
    m_FormFields.Add Field

    'Contratto padre
    Set Field = New FormField
    Set Field.Control = Me.txtIDContrattoPadre
    Field.Name = "IDRV_POContrattoPadre"
    Field.Visible = True
    Me.txtIDContrattoPadre.Tag = Field.Name
    m_FormFields.Add Field

    'Numero rinnovo
    Set Field = New FormField
    Set Field.Control = Me.txtNumeroRinnovo
    Field.Name = "NumeroRinnovo"
    Field.Visible = True
    Me.txtNumeroRinnovo.Tag = Field.Name
    m_FormFields.Add Field



    'Contratto attuale
    Set Field = New FormField
    Set Field.Control = Me.chkContrattoAttuale
    Field.Name = "ContrattoAttuale"
    Field.Visible = True
    Me.chkContrattoAttuale.Tag = Field.Name
    m_FormFields.Add Field

    'Descrizione tipo contratto
    Set Field = New FormField
    Set Field.Control = Me.txtDescrizioneContratto
    Field.Name = "DescrizioneTipoContratto"
    Field.Visible = True
    Me.txtDescrizioneContratto.Tag = Field.Name
    m_FormFields.Add Field

    
    'DescrizioneDisdetta
    Set Field = New FormField
    Set Field.Control = Me.txtDescrizioneDisdetta
    Field.Name = "DescrizioneDisdetta"
    Field.Visible = True
    Me.txtDescrizioneDisdetta.Tag = Field.Name
    m_FormFields.Add Field
    
 
    

    
    'AnnoContratto
    Set Field = New FormField
    Set Field.Control = Me.txtAnnoContratto
    Field.Name = "AnnoContratto"
    Field.Visible = True
    Me.txtAnnoContratto.Tag = Field.Name
    m_FormFields.Add Field
    
    'NumeroContratto
    Set Field = New FormField
    Set Field.Control = Me.txtNumeroContratto
    Field.Name = "NumeroContratto"
    Field.Visible = True
    Me.txtNumeroContratto.Tag = Field.Name
    m_FormFields.Add Field
 
    
    'FatturazioneRicorrente
    Set Field = New FormField
    Set Field.Control = Me.chkFatturazioneRic
    Field.Name = "FatturazioneRicorrente"
    Field.Visible = True
    Me.chkFatturazioneRic.Tag = Field.Name
    m_FormFields.Add Field
    
    'TotaleContrattoDaProdotti
    Set Field = New FormField
    Set Field.Control = Me.chkTotDaiProdotti
    Field.Name = "TotaleContrattoDaProdotti"
    Field.Visible = True
    Me.chkTotDaiProdotti.Tag = Field.Name
    m_FormFields.Add Field
    
    'Chiuso
    Set Field = New FormField
    Set Field.Control = Me.chkChiuso
    Field.Name = "Chiuso"
    Field.Visible = True
    Me.chkChiuso.Tag = Field.Name
    m_FormFields.Add Field

    'IDRV_POTipoImpostazioneContratto
    Set Field = New FormField
    Set Field.Control = Me.cboTipoImpostazione
    Field.Name = "IDRV_POTipoImpostazioneContratto"
    Field.Visible = True
    Me.cboTipoImpostazione.Tag = Field.Name
    m_FormFields.Add Field
    
    'chkGeneraRateProd
    Set Field = New FormField
    Set Field.Control = Me.chkGeneraRateProd
    Field.Name = "GeneraRatePerProdotto"
    Field.Visible = True
    Me.chkGeneraRateProd.Tag = Field.Name
    m_FormFields.Add Field
    
    'NumeroGiorniPrimaRata
    Set Field = New FormField
    Set Field.Control = Me.txtNGGPrimaRata
    Field.Name = "NumeroGiorniPrimaRata"
    Field.Visible = True
    Me.txtNGGPrimaRata.Tag = Field.Name
    m_FormFields.Add Field
    
    'Offerta
    Set Field = New FormField
    Set Field.Control = Me.chkOfferta
    Field.Name = "Offerta"
    Field.Visible = True
    Me.chkOfferta.Tag = Field.Name
    m_FormFields.Add Field
    
    'GestioneAccontoSaldo
    Set Field = New FormField
    Set Field.Control = Me.chkGeneraAccontiSaldo
    Field.Name = "GestioneAccontoSaldo"
    Field.Visible = True
    Me.chkGeneraAccontiSaldo.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDDurataContrattoProssimoRinnovo
    Set Field = New FormField
    Set Field.Control = Me.cboDurataContrattoProx
    Field.Name = "IDDurataContrattoProssimoRinnovo"
    Field.Visible = True
    Me.cboDurataContrattoProx.Tag = Field.Name
    m_FormFields.Add Field

    'DataPrimaDecorrenza
    Set Field = New FormField
    Set Field.Control = Me.txtDataPrimaDecorr
    Field.Name = "DataPrimaDecorrenza"
    Field.Visible = True
    Me.txtDataPrimaDecorr.Tag = Field.Name
    m_FormFields.Add Field

    'FineContratto
    Set Field = New FormField
    Set Field.Control = Me.chkFineContratto
    Field.Name = "FineContratto"
    Field.Visible = True
    Me.chkFineContratto.Tag = Field.Name
    m_FormFields.Add Field

    'DataScadenzaSecondoContratto
    Set Field = New FormField
    Set Field.Control = Me.txtDataScadSecContr
    Field.Name = "DataScadenzaSecondoContratto"
    Field.Visible = True
    Me.txtDataScadSecContr.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDTipoRinnovoProssimoRinnovo
    Set Field = New FormField
    Set Field.Control = Me.cboTipoRinnovoProx
    Field.Name = "IDTipoRinnovoProssimoRinnovo"
    Field.Visible = True
    Me.cboTipoRinnovoProx.Tag = Field.Name
    m_FormFields.Add Field
    
    'DataScadanzaRinnovoProssimoRinnovo
    Set Field = New FormField
    Set Field.Control = Me.txtDataScadPerRinnovoProx
    Field.Name = "DataScadanzaRinnovoProssimoRinnovo"
    Field.Visible = True
    Me.txtDataScadPerRinnovoProx.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDTipoRateizzazioneProssimoRinnovo
    Set Field = New FormField
    Set Field.Control = Me.cboTipoRateizzazioneProx
    Field.Name = "IDTipoRateizzazioneProssimoRinnovo"
    Field.Visible = True
    Me.cboTipoRateizzazioneProx.Tag = Field.Name
    m_FormFields.Add Field
    
End Sub


'**+
'Nome: ClosePreview
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Chiude la finestra della anteprima di stampa
'**/
Private Sub ClosePreview()
    Dim myDate
    
    On Error GoTo errHandler
        
    If m_Report.ClosePreview Then
        m_PreviewWindowHandle = 0
        PicForm.Visible = True
        BrwMain.Visible = m_TabMode
        ActivityBox.Visible = BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked
        FormRecalcLayout
        Set m_Report = Nothing
        SetStatus4Modality Preview, ClosePrw
    End If
    Exit Sub
errHandler:
    'Se si verifica un errore "SQL server in use"
    'la subroutine entra in un ciclo di attesa per
    '3 secondi prima di tentare nuovamente la chiusura
    myDate = Now
    If Err.Description = "SQL server in use" Then
        While Not (Now = DateAdd("s", 3, myDate))
        Wend
        Resume
    End If
    Err.Raise Err.Number, , Err.Description
End Sub




'**+
'Nome: ShortCut
'
'Parametri:
'KeyCode - Codice del tasto
'Shift - Stato del tasto Shift
'
'Valori di ritorno:
'
'Funzionalità:
'Gestione degli accelleratori da tastiera
'**/
'**+
'Nome: ShortCut
'
'Parametri:
'KeyCode - Codice del tasto
'Shift - Stato del tasto Shift
'
'Valori di ritorno:
'
'Funzionalità:
'Gestione degli accelleratori da tastiera
'**/
Private Function ShortCut(KeyCode As Integer, Shift As Integer) As Boolean
    Dim bCtrlDown As Boolean
    Dim bShiftDown As Boolean
    Dim bAltDown As Boolean
    
    bShiftDown = (Shift And vbShiftMask) > 0
    bCtrlDown = (Shift And vbCtrlMask) > 0
    bAltDown = (Shift And vbAltMask) > 0
    
    Select Case KeyCode
         Case vbKeyF12
            If bShiftDown Then
                If bCtrlDown Then
                    If BarMenu.Bands("Band_File").Tools("Mnu_Print").Enabled Then
                        ExecuteMenuCommand ("Mnu_Print")
                        ShortCut = True
                    End If
                Else
                    If BarMenu.Bands("Band_File").Tools("Mnu_Save").Enabled Then
                    
                        'Forza il lostfocus ed attende l'esecuzione di eventuali eventi associati
                        AutoLostFocus
                        
                        ExecuteMenuCommand ("Mnu_Save")
                        ShortCut = True
                    End If
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyF1
            SendMessage hwnd, WM_SETREDRAW, 0, 0
            'SendKeys ("{ESC}")
            DoEvents
            SendMessage hwnd, WM_SETREDRAW, 1, 0
            If bShiftDown Then
                'case shift F1
                ExecuteMenuCommand ("Mnu_Arg")
                ShortCut = True
                KeyCode = 0
                Shift = 0
            Else
                ExecuteMenuCommand ("Mnu_HelpOnLine")
                ShortCut = True
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyN
            If bCtrlDown Then
                If BarMenu.Bands("Band_File").Tools("Mnu_New").Enabled Then
                    ExecuteMenuCommand ("Mnu_New")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyX
'            If bCtrlDown Then
'                If BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Enabled Then
'                    ExecuteMenuCommand ("Mnu_Cut")
'                    ShortCut = True
'                End If
''                KeyCode = 0
''                Shift = 0
'            End If
            
        Case vbKeyC
            If bCtrlDown Then
'                If BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Enabled Then
'                    ExecuteMenuCommand ("Mnu_Copy")
'                    ShortCut = True
'                End If
                KeyCode = 0
                Shift = 0
            End If
            If bAltDown And BarMenu.Bands(BAND_CLOSE_PREVIEW).Visible Then
                ClosePreview
                ShortCut = True
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyV
            If bCtrlDown Then
'                If BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Enabled Then
'                    ExecuteMenuCommand ("Mnu_Paste")
'                    ShortCut = True
'                End If
                'KeyCode = 0
                'Shift = 0
            End If
            
        Case vbKeyT
            If bCtrlDown And bShiftDown = False Then   'CTRL + T
                If BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Enabled Then
                    ExecuteMenuCommand ("Mnu_NewSearch")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            If bCtrlDown And bShiftDown = True Then     'CTRL + MAIUSC + T
                If BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Enabled Then
                    ExecuteMenuCommand ("Mnu_SearchFilter")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyE
            If bCtrlDown Then
                If BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Enabled Then
                    ExecuteMenuCommand "Mnu_ExecuteSearch"
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyP
            If bCtrlDown Then
                If BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Enabled Then
                    ExecuteMenuCommand ("Mnu_SearchPrevious")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyS
            If bCtrlDown Then
                If BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Enabled Then
                    ExecuteMenuCommand ("Mnu_SearchNext")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyM
            If bCtrlDown Then
                If BarMenu.Bands("Band_View").Tools("Mnu_TableView").Enabled Then
                    ExecuteMenuCommand ("Mnu_TableView")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyF
            If bCtrlDown Then
                If BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled Then
                    ExecuteMenuCommand ("Mnu_FormView")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyDelete
            'Il tasto Canc ha effetto solo se il controllo attivo è la browse principale.
            If ActiveControl.Name = "BrwMain" And Not bShiftDown Then
                If BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Enabled Then
                    If BrwMain.Visible Then
                        ExecuteMenuCommand ("Mnu_Delete")
                        ShortCut = True
                        KeyCode = 0
                        Shift = 0
                    End If
                End If
            End If
            
        Case vbKeyR
            If bCtrlDown Then
                ExecuteMenuCommand ("Mnu_SearchObject")
                ShortCut = True
                'La condizione sottostante è necessaria per attivare l'acceleratore CTRL+R dalla modalità
                'filtri della DmrGrid
                If Not BrwMain.Visible Or (BrwMain.Visible And BrwMain.GuiMode = dgNormal) Then
                    KeyCode = 0
                    Shift = 0
                End If
            End If
            
        Case vbKeyG
            If bCtrlDown Then
                ExecuteMenuCommand ("Mnu_RunApplication")
                ShortCut = True
                KeyCode = 0
                Shift = 0
            End If
    
        Case vbKeyEscape
             If Not ActiveControl Is Nothing Then
                    If TypeName(ActiveControl) = "DmtGrid" Then
                        If BrwMain.GuiMode = dgFilterDefinition Then
                            If Not (m_Document.EOF = True And m_Document.BOF = True) Then
                                BrwMain.GuiMode = dgNormal
                                ExecuteMenuCommand "Mnu_TableView"
                                ShortCut = True
                            Else
                                'Ripulisce il contenuto delle condizioni.
                                BrwMain.Conditions.ClearValues
                                'Imposta la modalità FilterDefinition
                                BrwMain.GuiMode = dgFilterDefinition
                                ShortCut = True
                            End If
                        End If
                    End If
                KeyCode = 0
                Shift = 0
            End If
    
    End Select

End Function

'**+
'Nome: ShowErrorLog
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Mostra il dialogo di informazioni su l'ultimo errore
'bloccante verificatosi durante l'esecuzione del programma
'**/
Private Sub ShowErrorLog()
    Load frmErrorLog
    frmErrorLog.DMTErrorContol.MainProgram.Comments = App.Comments
    frmErrorLog.DMTErrorContol.MainProgram.Company = App.CompanyName
    frmErrorLog.DMTErrorContol.MainProgram.Copyright = App.LegalCopyright
    frmErrorLog.DMTErrorContol.MainProgram.Description = App.FileDescription
    frmErrorLog.DMTErrorContol.MainProgram.FileName = App.EXEName
    frmErrorLog.DMTErrorContol.MainProgram.Version = App.Major & "." & App.Minor & "." & App.Revision
    frmErrorLog.DMTErrorContol.ErrorNumber = Err.Number
    frmErrorLog.DMTErrorContol.ErrorDescription = Err.Description
    frmErrorLog.DMTErrorContol.Show
    frmErrorLog.Show vbModal
    End
End Sub


'**+
'Nome: OnBeforeOpenDoc
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Inizializzazioni da effettuare prima dell'apertura del documento.
'**/
Private Sub OnBeforeOpenDoc()
Dim rs As DmtOleDbLib.adoResultset
    'Inserire qui le
    'inizializzazioni da effettuare prima dell'apertura del documento.
    
    
    'rif6 begin
    
    Dim NewLink As DmtDocManLib.Link
    
'**************************SOTTO DOCUMENTO DELLA RATEIZZAZIONE*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_PORateContratto"
    
    Set m_DocumentsLink = m_Document.AddDocumentsLink("RV_PORateContratto")
          
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink.PrimaryKey = "IDRV_PORateContratto" '<-- Specifica il campo chiave primaria
    
    'Crea un Link LEFT JOIN sul campo "PagamentoRata"
    Set NewLink = m_DocumentsLink.AddLink("IDPagamentoRata", "Pagamento", ltLeft, "IDPagamento")
    NewLink.AddLinkColumn "Pagamento.Pagamento"
    
    'Crea un Link LEFT JOIN sul campo "RV_POContrattoAdeguamento"
    Set NewLink = m_DocumentsLink.AddLink("IDRV_POContrattoAdeguamento", "RV_POContrattoAdeguamento", ltLeft, "IDRV_POContrattoAdeguamento")
    NewLink.AddLinkColumn "RV_POContrattoAdeguamento.DescrizioneAdeguamento"

    'Crea un Link LEFT JOIN sul campo "RV_POContrattoProdotti"
    Set NewLink = m_DocumentsLink.AddLink("IDRV_POContrattoProdotti", "RV_POIEContrattoProdotti", ltLeft, "IDRV_POContrattoProdotti")
    NewLink.AddLinkColumn "RV_POIEContrattoProdotti.IDArticolo", "IDArtProdContratto"
    NewLink.AddLinkColumn "RV_POIEContrattoProdotti.CodiceArticolo", "CodiceArticoloProdContratto"
    NewLink.AddLinkColumn "RV_POIEContrattoProdotti.Articolo", "ArticoloProdContratto"
    NewLink.AddLinkColumn "RV_POIEContrattoProdotti.ValoreIndentificativo", "MatricolaProdottoDaContratto"
    
    'Crea un Link LEFT JOIN sul campo "RV_POProdotto"
    Set NewLink = m_DocumentsLink.AddLink("IDRV_POProdotto", "RV_POProdotto", ltLeft, "IDRV_POProdotto")
    NewLink.AddLinkColumn "RV_POProdotto.Descrizione", "DescrizioneProdotto"
    NewLink.AddLinkColumn "RV_POProdotto.Matricola", "MatricolaProdotto"

    'Crea un Link LEFT JOIN sul campo "Articolo"
    Set NewLink = m_DocumentsLink.AddLink("IDArticolo", "Articolo", ltLeft, "IDArticolo")
    NewLink.AddLinkColumn "Articolo.CodiceArticolo"
    NewLink.AddLinkColumn "Articolo.Articolo"
    
    
'**********************************************************************************************
    
'**************************SOTTO DOCUMENTO*****************************************************

    Set m_DocumentsLink1 = m_Document.AddDocumentsLink("RV_POContrattoServizi")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink1.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink1.PrimaryKey = "IDRV_POContrattoServizi" '<-- Specifica il campo chiave primaria
    
    'Crea un Link LEFT JOIN sul campo "IDArticolo"
    Set NewLink = m_DocumentsLink1.AddLink("IDArticolo", "Articolo", ltLeft, "IDArticolo")
    NewLink.AddLinkColumn "Articolo.CodiceArticolo"
    NewLink.AddLinkColumn "Articolo.Articolo"
    
    'Crea un Link LEFT JOIN sul campo "IDRV_POCriterioRicorrenza"
    Set NewLink = m_DocumentsLink1.AddLink("IDRV_POCriterioRicorrenza", "RV_POCriterioRicorrenza", ltLeft, "IDRV_POCriterioRicorrenza")
    NewLink.AddLinkColumn "RV_POCriterioRicorrenza.CriterioRicorrenza"
    
    'Crea un Link LEFT JOIN sul campo "IDRV_POTipoDataInizioRicorrenza"
    Set NewLink = m_DocumentsLink1.AddLink("IDRV_POTipoDataInizioRicorrenza", "RV_POTipoDataInizioRicorrenza", ltLeft, "IDRV_POTipoDataInizioRicorrenza")
    NewLink.AddLinkColumn "RV_POTipoDataInizioRicorrenza.TipoDataInizioRicorrenza"

    'Crea un Link LEFT JOIN sul campo "IDRV_POTipoDataFineRicorrenza"
    Set NewLink = m_DocumentsLink1.AddLink("IDRV_POTipoDataFineRicorrenza", "RV_POTipoDataFineRicorrenza", ltLeft, "IDRV_POTipoDataFineRicorrenza")
    NewLink.AddLinkColumn "RV_POTipoDataFineRicorrenza.TipoDataFineRicorrenza"

    'Crea un Link LEFT JOIN sul campo "IDRV_POTipoAnnoInizioRicorrenza"
    Set NewLink = m_DocumentsLink1.AddLink("IDRV_POTipoAnnoInizioRicorrenza", "RV_POTipoAnno", ltLeft, "IDRV_POTipoAnno")
    NewLink.AddLinkColumn "RV_POTipoAnno.TipoAnno", "TipoAnnoInizioRicorrenza"
    
    'Crea un Link LEFT JOIN sul campo "IDRV_POTipoAnnoFineRicorrenza"
    Set NewLink = m_DocumentsLink1.AddLink("IDRV_POTipoAnnoFineRicorrenza", "RV_POIE_TipoAnno", ltLeft, "IDRV_POTipoAnno")
    NewLink.AddLinkColumn "RV_POTipoAnno.TipoAnno", "TipoAnnoFineRicorrenza"
    
'***************************************************************************************************
    
'**************************SOTTO DOCUMENTO ADEGUAMENTI*********************************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_PORateContratto"
    
    Set m_DocumentsLink2 = m_Document.AddDocumentsLink("RV_POContrattoAdeguamento")
          
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink2.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink2.PrimaryKey = "IDRV_POContrattoAdeguamento" '<-- Specifica il campo chiave primaria
    
    'Crea un Link LEFT JOIN sul campo "Iva"
    Set NewLink = m_DocumentsLink2.AddLink("IDArticolo", "Articolo", ltLeft, "IDArticolo")
    NewLink.AddLinkColumn "Articolo.Articolo"
    NewLink.AddLinkColumn "Articolo.CodiceArticolo"
    
    'Crea un Link LEFT JOIN sul campo "Iva"
    Set NewLink = m_DocumentsLink2.AddLink("IDRV_POTipoAdeguamento", "RV_POTipoAdeguamento", ltLeft, "IDRV_POTipoAdeguamento")
    NewLink.AddLinkColumn "RV_POTipoAdeguamento.TipoAdeguamento"
    
    'Crea un Link LEFT JOIN sul campo "Rateizzazione"
    Set NewLink = m_DocumentsLink2.AddLink("IDRateizzazione", "RV_PORateizzazione", ltLeft, "IDRV_PORateizzazione")
    NewLink.AddLinkColumn "RV_PORateizzazione.Rateizzazione"
    
'**********************************************************************************************

'**************************SOTTO DOCUMENTO PRODOTTI*********************************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_PORateContratto"
    
    Set m_DocumentsLink3 = m_Document.AddDocumentsLink("RV_POContrattoProdotti")
          
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink3.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink3.PrimaryKey = "IDRV_POContrattoProdotti" '<-- Specifica il campo chiave primaria
    
    'Crea un Link LEFT JOIN sul campo "Prodotto"
    Set NewLink = m_DocumentsLink3.AddLink("IDRV_POProdotto", "RV_POProdotto", ltLeft, "IDRV_POProdotto")
    NewLink.AddLinkColumn "RV_POProdotto.Descrizione", "Prodotto"
    NewLink.AddLinkColumn "RV_POProdotto.Matricola"
    NewLink.AddLinkColumn "RV_POProdotto.ProdottoGenerico"
    
    'Crea un Link LEFT JOIN sul campo "Articolo"
    Set NewLink = m_DocumentsLink3.AddLink("IDArticolo", "Articolo", ltLeft, "IDArticolo")
    NewLink.AddLinkColumn "Articolo.CodiceArticolo"
    NewLink.AddLinkColumn "Articolo.Articolo", "DescrizioneArticolo"

    'Crea un Link LEFT JOIN sul campo "RV_POUnitaDiMisuraPeriodo"
    Set NewLink = m_DocumentsLink3.AddLink("IDRV_POUnitaDiMisuraPeriodo", "RV_POUnitaDiMisuraPeriodo", ltLeft, "IDRV_POUnitaDiMisuraPeriodo")
    NewLink.AddLinkColumn "RV_POUnitaDiMisuraPeriodo.Descrizione", "UnitaDiMisuraPeriodo"

    'Crea un Link LEFT JOIN sul campo "UnitaDiMisura"
    Set NewLink = m_DocumentsLink3.AddLink("IDUnitaDiMisuraArticolo", "UnitaDiMisura", ltLeft, "IDUnitaDiMisura")
    NewLink.AddLinkColumn "UnitaDiMisura.UnitaDiMisura", "UnitaDiMisuraArticolo"

    'Crea un Link LEFT JOIN sul campo "IDListino"
    Set NewLink = m_DocumentsLink3.AddLink("IDListino", "Listino", ltLeft, "IDListino")
    NewLink.AddLinkColumn "Listino.Listino"
    
    'Crea un Link LEFT JOIN sul campo "IDRateizzazione"
    Set NewLink = m_DocumentsLink3.AddLink("IDRateizzazione", "RV_PORateizzazione", ltLeft, "IDRV_PORateizzazione")
    NewLink.AddLinkColumn "RV_PORateizzazione.Rateizzazione"
    
'**********************************************************************************************

    
    'rif6 end



End Sub


'**+
'Autore: Carlo B. Collovà
'Data creazione: 20/11/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: InitExtensions
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Inizializza la componente adibita alla gestione dell'evento On_Extend
'
'**/
'Private Sub InitExtensions()
    'cbcx
    
    'Istanzia l'oggetto
    'Set m_ExtendApplication = New DmtExtendApp.ExtendApplication
    'Set m_ExtendApplication = New DmtExtendAppLib.ExtendApplication
    
    'Assegna un riferimento all'oggetto Application.
    'In questo modo la maggior parte dei parametri di inizializzazione vengono
    'letti da quest'ultimo
    'Set m_ExtendApplication.Application = m_App
    
    'Assegna un riferimento al controllo ActiveBar affinchè la classe
    'che gestisce i dati aggiuntivi possa interagire con la user interface
    'della manutenzione.
    'Set m_ExtendApplication.MenuBar = BarMenu
        
    'Se la funzione correntemente in esecuzione prevede l'evento On_Extend
    'vengono effettuate tutte le inizializzazioni del caso (come l'aggiunta di bottoni
    ' e menu alla BarMenu, ecc.) altrimenti la classe ExtendApplication non effettua
    'alcuna operazione.
    'm_ExtendApplication.Initialize
    
    'NOTA:
    '-----------------------------------------------------------------------------------------------------
    'Tutte le proprietà di m_ExtendApplication presenti anche nell'interfaccia IExtendApplication ed impostate
    'dopo la chiamata al metodo Initialize saranno settate anche in cContactPlus
    '-----------------------------------------------------------------------------------------------------
    
    'Assegna un riferimento del documento corrente
    'Set m_ExtendApplication.CurrentDocument = m_Document

'End Sub


'**+
'Nome: Start
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Inizializzazione e procedura di avvio
'**/
Private Sub Start()
    Dim OLDCursor As Integer
    Dim ToolID As Integer
    Dim Field As DmtDocManLib.Field
    Dim oActivity As IActivity
    Dim o As Activity
    Dim oFilter As Filter

        
    
    'Apertura del documento
    If Len(m_ExtendedDatabase) > 0 Then
        'Apre un nuovo documento usando il database esteso
        Set m_Document = m_App.OpenDocument(m_DocType, m_ExtendedDatabase)
    Else
        'Apre un nuovo documento usando il database diamante
        Set m_Document = m_App.OpenDocument(m_DocType)
    End If
    
    
    'NOTA: Con la sottostante proprietà settata a TRUE i metodi OnXXXDocumentsLink()
    'non sono più necessari in quanto il modello ad oggetti si occupa della gestione
    'dei sottodocumenti.
    '
    'Abilita la gestione automatica degli eventuali DocumentsLink
    m_Document.EnableRefreshDocumentsLinks = True
    
    
    'Clessidra
    OLDCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    Caption = m_App.Caption
    
    'Inizializzazione del controllo ActiveBar
    InitMenuBar ToolID
    InitToolBar ToolID
    ActivateBarButtons BTN_ALL, True
    
    'Inizializzazione del riquadro attività
    With ActivityBox
        .Activities.Clear
        
        'Aggiunge l'attività dei reports
        Set oActivity = .Activities.Add("DmtActBoxLib.ReportsActivity", "Reports")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load m_DocType.ID, TheApp.IDFirm
        Set o = oActivity
        Set oReportsActivity = o.InternalClass
        
        'Aggiunge l'attività dei filtri
        Set oActivity = .Activities.Add("DmtActBoxLib.FiltersActivity", "Filters")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load m_DocType
        Set o = oActivity
        Set oFiltersActivity = o.InternalClass
        
        'Aggiunge l'attività delle viste tabellari
        Set oActivity = .Activities.Add("DmtActBoxLib.TableViewsActivity", "TableViews")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load m_DocType.ID
        Set o = oActivity
        Set oTableViewsActivity = o.InternalClass

        'Aggiunge l'attività delle esportazioni
        Set oActivity = .Activities.Add("DmtActBoxLib.ExportActivity", "Export")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load m_DocType.ID
        Set o = oActivity
        Set oExportActivity = o.InternalClass
        
        'Aggiunge l'attività del supporto tecnico
        Set oActivity = .Activities.Add("DmtActBoxLib.SupportActivity", "Support")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load
        Set o = oActivity
        Set oSupportActivity = o.InternalClass
        
        'attiva/disattiva la visualizzazione delle attività
        EnableDOMActivitiesItems
        
        'imposta quale attività deve essere attivata per default
        If m_DefaultActivity <> "" Then
            Set .CurrentActivity = .Activities(m_DefaultActivity)
        End If
        
        'ridisegna il controllo
        .Redraw = True
    End With


    'Lettura impostazioni dal registry
    ReadRegistrySettings
        
    'Aggiunge due filtri temporanei, uno per le ricerche temporanee
    'e uno per la stampa in modalità form
    m_DocType.AddFilter "Temp"
    m_DocType.AddFilter "Form"
    
    

    'Connessione di tipo DMTADODBLib
        ConnessioneDiamanteADO
        GET_PARAMETRI_AZIENDA TheApp.IDFirm
        'GET_CONTROLLO_LICENZA
        GET_MODULO_ATTIVATO MODULO_CODICE, 51
        GET_MODULO_ATTIVATO_INT MODULO_CODICE_INT, 51
        GET_MODULO_ATTIVATO_NOL MODULO_CODICE_NOL, 51
        GET_MODULO_ATTIVATO_CONT MODULO_CODICE_CONT, 51
        
        NON_MODIFICA_CONTRATTO = GET_PARAMETRO_UTENTE(TheApp.IDUser, TheApp.IDFirm, "ModificaContratto")
        GET_PARAMETRI_STRINGA_FATT
        
    'Inizializzazioni da fare prima dell'apertura del documento
        OnBeforeOpenDoc
        
    'rif12
    'Altre inizializzazioni
        OnStart
    
    If Len(m_App.Caller) > 0 And m_App.CallerFieldValue > 0 Then
        '-------------------------------------------------
        '     Il programma è stato chiamato da un link.
        '-------------------------------------------------
        
        'In tal caso occorre mostrare in modalità variazione il record richiesto dal programma client.
        
        'Ripulisce la collezione Fields dell'oggetto DocType.
        For Each Field In m_DocType.Fields
            Field.Value = Empty
        Next
                
        'Imposta una condizione di ricerca basata sull'ID del record richiesto dal programma client.
        m_DocType.Fields("ID" & m_App.TableName).Value = m_App.CallerFieldValue
        
        'Rimuove il filtro precedente
        m_DocType.RemoveFilter "Temp"
        
        'Crea un nuovo filtro temporaneo a partire dalle condizioni di ricerca
        'e viene reso filtro attivo
        Set m_ActiveFilter = m_DocType.AddFilterWithConditions("Temp")
        
        'Inidica, nel caso di esegui gestione, se riportare il valore corrente al chiamante
        bNotReturnValue = CBool(Val(GetSetting(REGISTRY_KEY, App.EXEName, "NoReturnValue", "0")))
    Else
        '---------------------------------------------------
        '     Il programma non è stato chiamato da un link.
        '---------------------------------------------------
    
        'Il filtro attivo alla partenza è quello predefinito
        For Each oFilter In m_DocType.Filters
            If oFilter.ID = oFiltersActivity.DefaultFilterID Then
                Set m_ActiveFilter = m_DocType.Filters(oFilter.Name)
                Exit For
            End If
        Next
        
        m_Document.Dataset.Recordset.Sort = "Anagrafica, TipoContratto, NumeroRinnovo DESC"
    End If
    
    'Si comunica al documento quale filtro eseguire all'avvio.
    Set m_Document.ActiveFilter = m_ActiveFilter
    '
    'Set Me.BrwMain.Recordset = m_Document.Dataset.Recordset
    'Prima di aprire il documento occorre comunicargli qual'è il campo chiave primaria.
    m_Document.PrimaryKey = "ID" & m_Document.TableName
    'Apertura del documento.
    m_Document.OpenDoc
    
    'Questa impostazione serve per conservare le impostazioni grafiche
    BrwMain.IDUser = m_App.IDUser
    'Permette di gestire l'evento BrwMain_OnApplyFilter
    BrwMain.AutoFiltering = False
    'Con questa impostazione la dmtGrid NON effettua mai il Move sul documento.
    'Questo pertanto andrà forzato in BrwMain_DblClick e BrwMain_KeyDown.
    BrwMain.EnableMove = False
    'Inizializza le colonne da visualizzare nella griglia
    If m_DocType.DefaultTableView Is Nothing Then
        Err.Raise ERR_NO_DEFAULT_TABLEVIEW, , "Default TableView not found"
    Else
        BrwMain.LoadColumns m_DocType.DefaultTableView
        SetVisibilityIDFields
    End If
    
    
    'Crea i campi per la ricerca.
    CreateBrowserConditions
    'Assegnazione del riferimento alla fonte dati (binding sul recordset del documento)
    
    'rif14

    
    'Set BrwMain.Recordset = m_Document.Dataset.Recordset
    Set BrwMain.Recordset = m_Document.Data
    
    
            
     'Viene inizializzato il dialogo di stampa
    With DmtPrnDlg
        Set .Application = m_App
        Set .DocType = m_DocType
    End With
    
    
    
    'Ripulisco la tabella semaforo.
    'Se era avvenuto un crash di sistema questo garantisce il ripristino della situazione.
    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
    
    'Evita il blocco della toolbar
    'BarMenu.ResetHooks
    
    Screen.MousePointer = OLDCursor
End Sub




'**+
'Nome: ConditionType
'
'Parametri: DBType è il valore di DMTDocManLib.Field.DBType e rappresenta
'           il tipo di dato corrispondente all'oggetto Field in base dati.
'
'Valori di ritorno: una costante di tipo ConditionTypeConstants usata dalla Browse
'                   per costruire una condizione di ricerca.
'
'Funzionalità: Trasforma una costante DBType in una costante compatibile ConditionTypeConstants
'**/
Private Function ConditionType(ByVal DBType As Integer) As DmtGridCtl.ConditionTypeConstants
    Select Case DBType

        'dbTypeCHAR, dbTypeVARCHAR, dbTypeWCHAR, dbTypeWVARCHAR
        Case 1, 12, -8, -9
            ConditionType = dgCondTypeText
       
        'dbTypeNUMERIC, dbTypeDECIMAL, dbTypeINTEGER, dbTypeSMALLINT, dbTypeFLOAT
        'dbTypeREAL, dbTypeDOUBLE, dbTypeBIGINT, dbTypeTINYINT
        Case 2, 3, 4, 5, 6, 7, 8, -5, -6
            ConditionType = dgCondTypeNumber
            
        'dbTypeTIMESTAMP  ////NOTA: Se si desidera un campo dmCondTypeTime occorre gestirlo ad Hoc.
        Case 135
            ConditionType = dgCondTypeDate
    
        'dbTypeBIT
        Case -7, 11
            ConditionType = dgCondTypeBoolean
            
    End Select
End Function

'**+
'Nome: CreateBrowserConditions
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Crea automaticamente i campi per la ricerca (modalità DefineFilter)
'              a partire dai campi non ID del documento.
'**/
Private Sub CreateBrowserConditions()
    Dim Field As DmtDocManLib.Field
    Dim Cond As DmtGridCtl.dgCondition
    
    If Right(Trim(MenuOptions.ConnectionString), 1) = ";" Then
        Me.BrwMain.ConnectionString = MenuOptions.ConnectionString & "User Id=" & m_App.User & ";Password=" & m_App.Password
    Else
        Me.BrwMain.ConnectionString = MenuOptions.ConnectionString & ";" & "User Id=" & m_App.User & ";Password=" & m_App.Password
    End If
    
    
    
    'Vengono creati automaticamente i campi per la ricerca.
    'In una applicazione specifica questo codice andrà sostituito integralmente per definire
    'dei campi di ricerca ad hoc.
    
    'Non viene visualizzata la Check Intervallo perchè attualmente
    'il modello ad oggetti non prevede la gestione di filtri con
    'clausole BETWEEN.
    
    BrwMain.Conditions.Clear
    BrwMain.Conditions.WidthConditions = 300
    BrwMain.Conditions.WidthFields = 200
    BrwMain.Conditions.WidthIntervals = 100



        Set Cond = BrwMain.Conditions.Add("Anagrafica", "Cliente", m_DocType.TableName, True, False, , dgCondTypeText)
        Set Cond = BrwMain.Conditions.Add("AnagraficaFatturazione", "Cliente per fatturazione", m_DocType.TableName, True, False, , dgCondTypeText)
        
        Set Cond = BrwMain.Conditions.Add("IDTipoContratto", "Tipo contratto", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.RecordSource = "SELECT * FROM RV_POTipoContratto ORDER BY TipoContratto"
        Cond.DisplayField = "TipoContratto"
        Cond.KeyField = "IDRV_POTipoContratto"
        
        Set Cond = BrwMain.Conditions.Add("TipoContrattoFestivo", "Festivo", m_DocType.TableName, False, True, , dgCondTypeBoolean)
        
        Set Cond = BrwMain.Conditions.Add("AnnoContratto", "Anno contratto", m_DocType.TableName, False, True, , dgCondTypeNumber)
        Set Cond = BrwMain.Conditions.Add("NumeroContratto", "Numero contratto", m_DocType.TableName, False, True, , dgCondTypeNumber)
        
        
        
        Set Cond = BrwMain.Conditions.Add("DataStipula", "Data stipula", m_DocType.TableName, False, True, , dgCondTypeDate)
        Set Cond = BrwMain.Conditions.Add("DataPrimaDecorrenza", "Prima decorrenza", m_DocType.TableName, False, True, , dgCondTypeDate)
        Set Cond = BrwMain.Conditions.Add("DataDecorrenza", "Data decorrenza", m_DocType.TableName, False, True, , dgCondTypeDate)
        Set Cond = BrwMain.Conditions.Add("DataScadenza", "Scadenza I° contratto", m_DocType.TableName, False, True, , dgCondTypeDate)
        Set Cond = BrwMain.Conditions.Add("DataScadenzaSecondoContratto", "Scadenza II° contratto", m_DocType.TableName, False, True, , dgCondTypeDate)
        Set Cond = BrwMain.Conditions.Add("DataScadenzaPerRinnovo", "Rinnovo entro", m_DocType.TableName, False, True, , dgCondTypeDate)
        Set Cond = BrwMain.Conditions.Add("DataFineAssistenza", "Fine assistenza", m_DocType.TableName, False, True, , dgCondTypeDate)
        
        Set Cond = BrwMain.Conditions.Add("IDDurataContratto", "Durata I° contratto", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.RecordSource = "SELECT * FROM RV_PODurataContratto ORDER BY DurataContratto"
        Cond.DisplayField = "DurataContratto"
        Cond.KeyField = "IDRV_PODurataContratto"
           
        Set Cond = BrwMain.Conditions.Add("IDTipoRinnovo", "Tipo rinnovo periodo", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.RecordSource = "SELECT * FROM RV_POTipoRinnovo ORDER BY TipoRinnovo"
        Cond.DisplayField = "TipoRinnovo"
        Cond.KeyField = "IDRV_POTipoRinnovo"

        Set Cond = BrwMain.Conditions.Add("IDDurataContrattoProssimoRinnovo", "Durata II° contratto", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.RecordSource = "SELECT * FROM RV_PODurataContratto ORDER BY DurataContratto"
        Cond.DisplayField = "DurataContratto"
        Cond.KeyField = "IDRV_PODurataContratto"

        Set Cond = BrwMain.Conditions.Add("IDRV_POTipoDurataAssistenza", "Tipo Durata assistenza", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.RecordSource = "SELECT * FROM RV_POTipoDurataAssistenza ORDER BY TipoDurataAssistenza"
        Cond.DisplayField = "TipoDurataAssistenza"
        Cond.KeyField = "IDRV_POTipoDurataAssistenza"

        Set Cond = BrwMain.Conditions.Add("IDRateizzazione", "Tipo rateizzazione", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.RecordSource = "SELECT * FROM RV_PORateizzazione ORDER BY Rateizzazione"
        Cond.DisplayField = "Rateizzazione"
        Cond.KeyField = "IDRV_PORateizzazione"

        Set Cond = BrwMain.Conditions.Add("IDRaggruppamentoFatturato", "Raggruppamento fatturato", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.RecordSource = "SELECT * FROM RaggruppamentoFatturato ORDER BY RaggruppamentoFatturato"
        Cond.DisplayField = "RaggruppamentoFatturato"
        Cond.KeyField = "IDRaggruppamentoFatturato"

        Set Cond = BrwMain.Conditions.Add("IDRV_POTipoClassificazioneContratto", "Classificazione contratto", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.RecordSource = "SELECT * FROM RV_POTipoClassificazioneContratto ORDER BY TipoClassificazioneContratto"
        Cond.DisplayField = "TipoClassificazioneContratto"
        Cond.KeyField = "IDRV_POTipoClassificazioneContratto"



        Set Cond = BrwMain.Conditions.Add("CodiceArticoloContratto", "Codice articolo contratto", m_DocType.TableName, True, False, , dgCondTypeText)
        Set Cond = BrwMain.Conditions.Add("ArticoloContratto", "Articolo contratto", m_DocType.TableName, True, False, , dgCondTypeText)

        Set Cond = BrwMain.Conditions.Add("SitoPerAnagrafica", "Filiale", m_DocType.TableName, True, False, , dgCondTypeText)
        Set Cond = BrwMain.Conditions.Add("AnagraficaAgente", "Agente", m_DocType.TableName, True, False, , dgCondTypeText)
        Set Cond = BrwMain.Conditions.Add("NomeAgente", "Nome agente", m_DocType.TableName, False, False, , dgCondTypeText)
                
        Set Cond = BrwMain.Conditions.Add("IDAnagraficaCommesso", "Tecnico", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.RecordSource = "SELECT * FROM RV_POIEAnagraficaPerTipo WHERE IDAzienda=" & TheApp.IDFirm & " AND IDTipoAnagrafica=" & LINK_TIPO_ANA_TEC_INT & " ORDER BY Anagrafica"
        Cond.DisplayField = "Anagrafica"
        Cond.KeyField = "IDAnagrafica"

        Set Cond = BrwMain.Conditions.Add("NomeCommesso", "Nome tecnico", m_DocType.TableName, False, False, , dgCondTypeText)

        Set Cond = BrwMain.Conditions.Add("IDAnagraficaAmministratore", "Amministratore", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.RecordSource = "SELECT * FROM RV_POIEAnagraficaPerTipo WHERE IDAzienda=" & TheApp.IDFirm & " AND IDTipoAnagrafica=" & LINK_TIPO_ANA_AMM & " ORDER BY Anagrafica"
        Cond.DisplayField = "Anagrafica"
        Cond.KeyField = "IDAnagrafica"

        Set Cond = BrwMain.Conditions.Add("NomeAmministratore", "Nome amministratore", m_DocType.TableName, False, False, , dgCondTypeText)
        
        Set Cond = BrwMain.Conditions.Add("NonFatturare", "Non fatturare", m_DocType.TableName, False, False, , dgCondTypeBoolean)
        Set Cond = BrwMain.Conditions.Add("AdeguamentoIstat", "Adeguamente Istat", m_DocType.TableName, False, False, , dgCondTypeBoolean)
        Set Cond = BrwMain.Conditions.Add("RinnovoAutomatico", "Rinnovo automatico", m_DocType.TableName, False, False, , dgCondTypeBoolean)
        Set Cond = BrwMain.Conditions.Add("RitenutaAcconto", "RitenutaAcconto", m_DocType.TableName, False, False, , dgCondTypeBoolean)

        Set Cond = BrwMain.Conditions.Add("Disdetta", "Disdetta", m_DocType.TableName, False, False, , dgCondTypeBoolean)
            If NON_IMPOSTARE_FILTRI = 0 Then Cond.FromValue = "NO"
        Set Cond = BrwMain.Conditions.Add("ContrattoAttuale", "Contratto attuale", m_DocType.TableName, False, False, , dgCondTypeBoolean)
            If NON_IMPOSTARE_FILTRI = 0 Then Cond.FromValue = "SI"
        Set Cond = BrwMain.Conditions.Add("Offerta", "Offerta", m_DocType.TableName, False, False, , dgCondTypeBoolean)
            If NON_IMPOSTARE_FILTRI = 0 Then Cond.FromValue = "NO"
        
        Set Cond = BrwMain.Conditions.Add("FineContratto", "Fine contratto", m_DocType.TableName, False, False, , dgCondTypeBoolean)
        Set Cond = BrwMain.Conditions.Add("Chiuso", "Contratto chiuso", m_DocType.TableName, False, False, , dgCondTypeBoolean)
        
        Set Cond = BrwMain.Conditions.Add("DataDisdetta", "Data disdetta", m_DocType.TableName, False, True, , dgCondTypeDate)
        Set Cond = BrwMain.Conditions.Add("Annotazioni", "Annotazioni", m_DocType.TableName, False, False, , dgCondTypeText)
        
End Sub

'**+
'Nome: Export
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue l'esportazione del documento con controllo di errore
'**/
Private Sub ExportDocument(ByVal Appl As Long)
    On Error GoTo errHandler
    
    Dim OLDCursor As Integer
    
    OLDCursor = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    m_Document.Export m_Report, Appl
    
    Screen.MousePointer = OLDCursor
    Exit Sub
errHandler:
    Screen.MousePointer = OLDCursor
    
    If Err.Number = 20507 Then
        'Errore "Invalid file Name" generato quando non è possibile trovare il file .rpt
        sbMsgInfo "File di report non trovato", m_App.FunctionName
    Else
        sbMsgInfo Err.Description, m_App.FunctionName
    End If
End Sub

'**+
'Nome: PrintDocument
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue la stampa del documento con controllo di errore per nessuna stampante
'definita
'**/
Private Sub PrintDocument(ByVal ToolName As String)
    On Error GoTo errHandler
    
    Dim OLDCursor As Integer
    
    '**+ Riferimento al cursore corrente
    OLDCursor = Screen.MousePointer
    
    '**+ Inizializzazione selezioni di stampa
    m_Report.Copies = 1
    m_Report.Orientation = ocPortrait
    m_Report.PrinterName = ""
    
    If ToolName = "Mnu_Print" Then
        '**+ stampa con dialogo
        Set DmtPrnDlg.Report = m_Report
        DmtPrnDlg.Show
        If Not DmtPrnDlg.Cancel Then
            Screen.MousePointer = vbHourglass
            m_Document.DoPrint m_Report
        End If
    Else
        'Stampa diretta
        Screen.MousePointer = vbHourglass
        m_Document.DoPrint m_Report
    End If
    
    Screen.MousePointer = OLDCursor
    Exit Sub

errHandler:
    Screen.MousePointer = OLDCursor
    If Err.Number = vbObjectError + 36 Then
        ' errore generato all'interno della DMTDocManLib per nessuna stampante
        sbMsgInfo "Non è possibile ottenere informazioni sulla stampante." & Chr(13) & "Controllare che sia installata correttamente", m_App.FunctionName
    ElseIf Err.Number = vbObjectError + 4 Then
        'Si è annullata la stampa.
    Else
        sbMsgInfo Err.Description, m_App.FunctionName
    End If
    
End Sub

'**+
'Nome: DoNewDocument
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Procedure per la richiesta di un nuovo documento
'**/
Private Function DoNewDocument() As Integer
    
    '------------------------------------------------
    'Inserire qui se occorre del codice specifico
    'per la manutenzione.
    '------------------------------------------------
    
    DoNewDocument = ChooseAboutSaving
End Function

'**+
'Nome: WriteStatusBar
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Scrive una stringa di testo nella StatusBar
'**/
Private Sub WriteStatusBar(ByVal sTesto As String)
    If stbStatusbar.Style = sbrSimple Then
        stbStatusbar.SimpleText = sTesto
    Else
        stbStatusbar.Panels(1).Text = sTesto
    End If
End Sub

'**+
'Nome: FormUnload
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue i controlli alla richiesta di abbandono del form
'**/
Private Function FormUnload() As Integer
    Dim sMessage As String
    Dim sMessage1 As String
    Dim lIDField As Long
    Dim Testo As String
    
    
    If Me.BrwMain.Visible = False Then
        If Me.cboTipoImpostazione.CurrentID <> 3 Then
            If (GET_CONTROLLO_IMPORTO_RATE(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) <> Me.txtImportoAttuale.Value) Then
                Testo = "ATTENZIONE!!!" & vbCrLf
                Testo = Testo & "Il totale delle rate non è uguale al totale del contratto" & vbCrLf
                Testo = Testo & "Vuoi continuare?"
                If MsgBox(Testo, vbCritical + vbYesNo, "Controllo dati contratto") = vbNo Then
                    FormUnload = 1
                    Exit Function
                End If
            End If
        End If
    End If
    
    If m_Changed Then
        Select Case ChooseAboutSaving
            Case vbCancel
                FormUnload = 1
                Exit Function
            Case vbYes
                OnSave
                'Se la registrazione non è andata a buon fine
                'esce e non chiude il programma
                If Not m_Saved Then
                    FormUnload = 1
                    Exit Function
                End If
        End Select
    End If
        
    
    If m_PreviewWindowHandle > 0 Then
        ClosePreview
    End If
    
    SaveRegistrySettings
    
    'Se il programma è stato chiamato da un link occorre restituire l'ID del record attivo
    'all'applicazione chiamante.
    If Len(m_App.Caller) > 0 Then
        'Il programma è stato chiamato da un link.
        
        'Se non verrà correttamente selezionato un elemento sarà restituito il valore -1 all'applicazione client.
        lIDField = -1
        
        'Se il documento è vuoto non si deve far nulla.
        'Se la browse è in modalità Filter Definition non formula la domanda di riporto dei dati nel programma chiamante.
        If (Not (m_Document.EOF And m_Document.BOF)) And (BrwMain.GuiMode <> dgFilterDefinition) Then
        
            'ATTENZIONE: La stringa sMessage1 deve essere personalizzata a seconda dei casi!!!
'            sMessage1 = " il " & m_DocType.Name
'            sMessage = sMessage1 & " """ & m_Document.Fields(CAMPO_PER_CAPTION).Value & """"
'
'            gResource.CustomStrings.Clear
'            gResource.CustomStrings.Add sMessage, 1
'
'            'Viene chiesto se si intende riportare il record corrente al programma chiamante.
'            'If fnMsgQuestion(gResource.GetCustomizedMessage(MESS_QUERYPASTE), m_App.FunctionName) = vbYes Then
'                'Legge l'ID del record corrente affinchè venga riportato all'applicazione chiamante.
'                lIDField = m_Document.Fields("ID" & m_App.TableName).Value
'            'End If
            
        End If
        
        'Scrive sul registry l'ID da passare all'aplicazione chiamante.
        SaveSetting REGISTRY_KEY, m_App.Caller, "IDField", lIDField
                      
    End If
    
End Function

'**+
'Nome: FormRecalcLayout
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Ricalcolo del layout del form
'**/
Private Sub FormRecalcLayout()
    Dim Height As Single
    Dim Width As Single

    'Se il form è minimizzato non serve il ricalcolo del layout
    If WindowState <> vbMinimized Then
        ActivityBox.Top = BarMenu.ClientAreaTop
        ActivityBox.Left = BarMenu.ClientAreaLeft
        ActivityBox.Height = IIf(BarMenu.ClientAreaHeight > 0, BarMenu.ClientAreaHeight, 0)
        
        imgSplitter.Visible = ActivityBox.Visible
        imgSplitter.Top = ActivityBox.Top
        imgSplitter.Height = ActivityBox.Height
        
        If ActivityBox.Visible Then
            imgSplitter.Left = ActivityBox.Width + ActivityBox.Left
            picSplitter.Left = imgSplitter.Left
        End If
        
        PicForm.Top = BarMenu.ClientAreaTop
        
        If ActivityBox.Visible Then
            PicForm.Left = imgSplitter.Left + imgSplitter.Width
        Else
            PicForm.Left = BarMenu.ClientAreaLeft
        End If
        


        Width = BarMenu.ClientAreaWidth - IIf(ActivityBox.Visible, ActivityBox.Width + imgSplitter.Width, 0)
        Height = BarMenu.ClientAreaHeight
        
        PicForm.Width = IIf(Width < 100, 100, Width)
        PicForm.Height = IIf(Height < 100, 100, Height)
        
        'RIDIMENSIONA LA SPLIT BAR IN BASE ALLA DIMENSIONE DEL FORM
        DMTSplitBar1.Move PicForm.Left, PicForm.Top, PicForm.Width, PicForm.Height
        'INIZIALIZZA LA SPLIT BAR
        DMTSplitBar1.SetSplitBar Height, Width, Me.PicForm2.Height, Me.PicForm2.Width
        
        'PicForm.Top = BarMenu.ClientAreaTop
        
        'If ActivityBox.Visible Then
        '    PicForm.Left = imgSplitter.Left + imgSplitter.Width
        'Else
        '    PicForm.Left = BarMenu.ClientAreaLeft
        'End If
        
        BrwMain.Top = PicForm.ScaleTop
        BrwMain.Left = PicForm.ScaleLeft
        BrwMain.Width = PicForm.ScaleWidth
        BrwMain.Height = PicForm.ScaleHeight
        

    End If
End Sub

'**+
'Nome: GetStatusToolBar
'
'Parametri:
'Enabled - Stato di abilitazione da controllare
'
'Valori di ritorno:
'
'Funzionalità:
'Calcola lo stato dei bottoni della ToolBar standard
'**/
Private Function GetStatusToolBar(ByVal Enabled As Boolean) As Currency
    Dim Valore As Currency

    Valore = 0
    If BarMenu.Bands("Standard").Tools("New").Enabled = Enabled Then Valore = Valore Or BTN_NEW
    If BarMenu.Bands("Standard").Tools("Save").Enabled = Enabled Then Valore = Valore Or BTN_SAVE
    If BarMenu.Bands("Standard").Tools("Print").Enabled = Enabled Then Valore = Valore Or BTN_PRINT
    If BarMenu.Bands("Standard").Tools("PrePrint").Enabled = Enabled Then Valore = Valore Or BTN_PREVIEW
    If BarMenu.Bands("Standard").Tools("Cut").Enabled = Enabled Then Valore = Valore Or BTN_CUT
    If BarMenu.Bands("Standard").Tools("Copy").Enabled = Enabled Then Valore = Valore Or BTN_COPY
    If BarMenu.Bands("Standard").Tools("Paste").Enabled = Enabled Then Valore = Valore Or BTN_PASTE
    If BarMenu.Bands("Standard").Tools("Delete").Enabled = Enabled Then Valore = Valore Or BTN_DELETE
    If BarMenu.Bands("Standard").Tools("Clear").Enabled = Enabled Then Valore = Valore Or BTN_CLEAR
    If BarMenu.Bands("Standard").Tools("NewSearch").Enabled = Enabled Then Valore = Valore Or BTN_FIND
    If BarMenu.Bands("Standard").Tools("ExecuteSearch").Enabled = Enabled Then Valore = Valore Or BTN_SEARCH
    If BarMenu.Bands("Standard").Tools("ChangeView").Enabled = Enabled Then Valore = Valore Or BTN_VIEWMODE
    If BarMenu.Bands("Standard").Tools("SearchPrevious").Enabled = Enabled Then Valore = Valore Or BTN_PREVIOUS
    If BarMenu.Bands("Standard").Tools("SearchNext").Enabled = Enabled Then Valore = Valore Or BTN_NEXT
    If BarMenu.Bands("Standard").Tools("Export").Enabled = Enabled Then Valore = Valore Or BTN_EXPORT
    If BarMenu.Bands("Band_Export").Tools("ExportWord").Enabled = Enabled Then Valore = Valore Or BTN_WORD
    If BarMenu.Bands("Band_Export").Tools("ExportExcel").Enabled = Enabled Then Valore = Valore Or BTN_EXCEL
    If BarMenu.Bands("Band_Export").Tools("ExportHtml").Enabled = Enabled Then Valore = Valore Or BTN_HTML
    If BarMenu.Bands("Band_Export").Tools("ExportPDF").Enabled = Enabled Then Valore = Valore Or BTN_PDF
    
    If BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled = Enabled Then Valore = Valore Or BTN_SEARCHFORM
    If BarMenu.Bands("Band_View").Tools("Mnu_TableView").Enabled = Enabled Then Valore = Valore Or BTN_SEARCHTABLE
    If BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Enabled = Enabled Then Valore = Valore Or BTN_FILTER

    If BarMenu.Bands("Band_File").Tools("Mnu_New").Enabled = Enabled Then Valore = Valore Or BTN_NEW
    If BarMenu.Bands("Band_File").Tools("Mnu_Save").Enabled = Enabled Then Valore = Valore Or BTN_SAVE
    If BarMenu.Bands("Band_File").Tools("Mnu_Print").Enabled = Enabled Then Valore = Valore Or BTN_PRINT
    If BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Enabled = Enabled Then Valore = Valore Or BTN_PREVIEW

    If BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Enabled = Enabled Then Valore = Valore Or BTN_CUT
    If BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Enabled = Enabled Then Valore = Valore Or BTN_COPY
    If BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Enabled = Enabled Then Valore = Valore Or BTN_PASTE
    If BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Enabled = Enabled Then Valore = Valore Or BTN_DELETE
    If BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Enabled = Enabled Then Valore = Valore Or BTN_CLEAR
    If BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Enabled = Enabled Then Valore = Valore Or BTN_FIND
    If BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Enabled = Enabled Then Valore = Valore Or BTN_SEARCH
    If BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Enabled = Enabled Then Valore = Valore Or BTN_PREVIOUS
    If BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Enabled = Enabled Then Valore = Valore Or BTN_NEXT

    If BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Enabled = Enabled Then Valore = Valore Or BTN_TOOLS
    If BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled = Enabled Then Valore = Valore Or BTN_SEARCHFORM
    If BarMenu.Bands("Band_View").Tools("Mnu_TableView").Enabled = Enabled Then Valore = Valore Or BTN_SEARCHTABLE

    If BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Enabled = Enabled Then Valore = Valore Or BTN_EXPORT
    If BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Enabled = Enabled Then Valore = Valore Or BTN_WORD
    If BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Enabled = Enabled Then Valore = Valore Or BTN_EXCEL
    If BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Enabled = Enabled Then Valore = Valore Or BTN_HTML
    If BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Enabled = Enabled Then Valore = Valore Or BTN_PDF

    GetStatusToolBar = Valore
End Function

'**+
'Nome: ReadRegistrySettings
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Legge i valori registrati nel registry relativi allo stato
'dei controlli del Form
'
'
'**/
Private Sub ReadRegistrySettings()
    Dim Index As Integer
    Dim FormHeight As Single
    Dim FormWidth As Single
    Dim NomeBanda As String
    Dim lngIDLanguage As Long
    Dim bFoldersVisible As Boolean
    Dim lValue As Long
           
           
    'Lettura file di help
    App.HelpFile = MenuOptions.ProgramsPath & "\Diamante.chm"
           
    ' Legge dal Registry le impostazioni sulla lingua
    lngIDLanguage = AppOptions.IDLanguage
           
    ' Modifica tutte le stringhe nel linguaggio corrente ( se <> da default )
    If lngIDLanguage <> NATIVE_LANGUAGE Then
        gResource.IDCurrentLanguage = lngIDLanguage
        'Setta i nuovi ToolTipText della Toolbar
        'e le Caption dei menu
        ChangeMenuLanguage
        ChangeToolBarLanguage
        'Traduce tutte le stringhe presenti sul form
        '(Solo se ChangeStringsLanguage è gestita dal programmatore !!!)
        ChangeStringsLanguage
    End If
    
    'Settaggio per la statusbar
    stbStatusbar.Visible = AppOptions.StatusBarVisibility
        
        
    '**+ settaggi per la barra degli strumenti
    With BarMenu
    
        '**+ E' necessario verificare la versione dell'activebar xchè nella nuova vesione 3.0
        'sono stati cambiati i valori di impostazione della proprietà DockingArea
        If AppOptions.BARMENUVERSION = BARMENUVERSION Then
    
            For Index = 0 To .Bands.Count - 1
            
                'Settaggi sulle toolbar (ancoraggio e dimensioni)
                If .Bands(Index).Type <> ddBTPopup Then
                    With .Bands(Index)
                        If AppOptions.ToolbarDockingArea(Index) > -1 Then
                            .DockingArea = AppOptions.ToolbarDockingArea(Index)
                            .DockLine = AppOptions.ToolbarDockLine(Index)
                            lValue = AppOptions.ToolbarHeight(Index)
                            If lValue > 0 Then .Height = lValue
                            lValue = AppOptions.ToolbarWidth(Index)
                            If lValue > 0 Then .Width = lValue
                            '**+ Attenzione le impostazioni del Left e Top devono essere effettuate dopo
                            'quelle dell'Height e del Width xchè se siamo in presenza di valori superiori
                            'a quelli della ClientArea azzera il left e top impostati in precedenza **/
                            lValue = AppOptions.ToolbarLeft(Index)
                            If lValue > 0 Then .Left = lValue
                            lValue = AppOptions.ToolbarTop(Index)
                            If lValue > 0 Then .Top = lValue
                            .DockingOffset = AppOptions.ToolbarDockingOffset(Index)
                        End If
                    End With
                End If
            
                'Settaggi sulla visibilità delle toolbar.
                If .Bands(Index).Type = ddBTNormal And .Bands(Index).Name <> BAND_CLOSE_PREVIEW Then
                     NomeBanda = .Bands(Index).Name
                     .Bands(NomeBanda).Visible = AppOptions.ToolbarVisibility(NomeBanda)
                End If
        
            Next Index
            
        End If
        
        'Settaggio sulla visualizzazione dei tooltip.
        .DisplayToolTips = AppOptions.DisplayTooltip
    End With
        
    
    'Dimensione delle icone della ToolBar
    SetToolBarIcons AppOptions.LargeIcon
    
    BarMenu.RecalcLayout
   
    bFoldersVisible = AppOptions.FoldersVisibility
   
    'Settaggi del riquadro attività
    ActivityBox.Visible = bFoldersVisible
    ActivityBox.Width = AppOptions.FoldersWidth
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked = bFoldersVisible
    m_DefaultActivity = AppOptions.DefaultActivity
    
    '**+ settaggi per la finestra principale del programma
    WindowState = AppOptions.WindowState
    If WindowState = 0 Then
        FormHeight = AppOptions.FormHeight
        If FormHeight <> -1 Then
            Height = FormHeight
        End If
        FormWidth = AppOptions.FormWidth
        If FormWidth <> -1 Then
            Width = FormWidth
        End If
    End If
End Sub

'**+
'Nome: SaveRegistrySettings
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Salva i valori relativi allo stato dei controlli del Form
'nel registry
'
'**/
Private Sub SaveRegistrySettings()
    Dim I As Integer

    '**+ Salva le impostazioni relative alle toolbar
    With AppOptions
        
        For I = 0 To BarMenu.Bands.Count - 1
            If BarMenu.Bands(I).Type <> ddBTPopup Then
                    .ToolbarDockingArea(I) = BarMenu.Bands(I).DockingArea
                    .ToolbarDockLine(I) = BarMenu.Bands(I).DockLine
                    .ToolbarLeft(I) = BarMenu.Bands(I).Left
                    .ToolbarTop(I) = BarMenu.Bands(I).Top
                    .ToolbarHeight(I) = BarMenu.Bands(I).Height
                    .ToolbarWidth(I) = BarMenu.Bands(I).Width
                    .ToolbarDockingOffset(I) = BarMenu.Bands(I).DockingOffset
            End If
        Next I
        .BARMENUVERSION = BARMENUVERSION
        
        'Salva le impostazioni relative alla finestra principale.
        If WindowState <> vbMinimized Then
            .FormHeight = Height
            .FormWidth = Width
            .WindowState = WindowState
        End If
        
        'Salva le impostazioni del riquadro attività
        .FoldersWidth = ActivityBox.Width
        .FoldersVisibility = ActivityBox.Visible
        .DefaultActivity = ActivityBox.CurrentActivityKey
    End With
End Sub


'**+
'Nome: ChangeView
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Cambia modalità di visualizzazione dei dati tra Form e vista tabellare
'
'**/
Private Sub ChangeView(Optional ByVal sToolName As Variant)

    'Se non vi sono record presenti nel browser
    'la modalità di visualizzazione non cambia e si esce.
    If (m_Document.EOF = True And m_Document.BOF = True) Then Exit Sub

    'Se si proviene dalla modalità tabellare
    '( o dalla modalità filtro provenendo dalla modalità tabellare )
    'potrebbe essere necessario allineare il documento con l'ultima selezione fatta nella browse.
    If BrwMain.Visible = True Then
        If BrwMain.ListIndex > 0 Then
            m_Document.Move BrwMain.ListIndex - 1
        End If
    End If
    

    If IsMissing(sToolName) Then sToolName = "ChangeView"

    'Cambia la visibiltà del browser
    If sToolName = "ChangeView" Then
        BrwMain.Visible = IIf(BrwMain.Visible And BrwMain.GuiMode = dgNormal, False, True)
    Else
        BrwMain.Visible = IIf((sToolName = "Mnu_FormView"), False, True)
    End If
    
    'Se si va in modalità form ed il record è bloccato si torna in modalità tabellare
    'impedendo di effettuare modifiche su quel record.
    'Quando si va in modalità tabellare il controllo non è necessario.
    If Not BrwMain.Visible Then

        If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
            'Il record è bloccato - si va in modalità tabellare
            
            BrwMain.Visible = True

            'Input Focus al browser
            'BrwMain.SetFocus

            'Refresh dello stato dei bottoni della ToolBar standard e dei menu
            SetStatus4Modality Browse

            Exit Sub
        Else
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
        End If

    End If
    
    
    
    'Se si era in fase di immissione di un nuovo record viene annullata
    m_Document.AbortNew
    
    If BrwMain.Visible Then 'Modalità tabellare
        
        'Input Focus al browser
        'BrwMain.SetFocus
        
        'Refresh dello stato dei bottoni della ToolBar standard e dei menu
        SetStatus4Modality Browse
        
    Else 'Modalità form
        
        'Refresh dello stato dei bottoni della ToolBar standard e dei menu
        SetStatus4Modality Modify
        
        'Input Focus al primo campo del form
        SetFocusTabIndex0
    End If
       
    'Imposta i suggerimenti da visualizzare sulla Statusbar in funzione
    'della modalità di visualizzazione corrente.
    'Ad esempio in alcuni casi le frasi sono al Singolare/Plurare.
    'La funzione GetDescription4StatusBar si occupa di determinare la frase esatta.
    'La Sub RefreshDescriptions4StatusBar deve essere chiamata anche in Execute_Search()--> Vedi.
    RefreshDescriptions4StatusBar
End Sub

'**+
'Nome: InitMenuBar
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Inizializzazione della MenuBar
'
'**/
Private Sub InitMenuBar(ByRef ToolID As Integer)
    BarMenu.Bands.Add "Band_Menu"
    BarMenu.Bands("Band_Menu").WrapTools = True
    BarMenu.Bands("Band_Menu").Type = ddBTMenuBar
    BarMenu.Bands("Band_Menu").DockLine = 1
    BarMenu.Bands("Band_Menu").Flags = ddBFDockTop Or ddBFDockLeft Or ddBFFloat Or ddBFDockRight Or ddBFDockBottom
    BarMenu.Bands("Band_Menu").GrabHandleStyle = ddGSNormal

    'File
    BarMenu.Bands.Add "Band_File"
    BarMenu.Bands("Band_File").Type = ddBTPopup
    BarMenu.Bands("Band_File").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "File"
    BarMenu.Bands("Band_Menu").Tools("File").SubBand = "Band_File"
    BarMenu.Bands("Band_Menu").Tools("File").Caption = GetCaption4MenuBar("File")
    BarMenu.Bands("Band_Menu").Tools("File").Description = GetDescription4StatusBar("File")

    'File-New
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_New"
    BarMenu.Bands("Band_File").Tools("Mnu_New").SetPicture 0, gResource.GetBitmap(IDB_STD_NEW16), &HC0C0C0
    BarMenu.Bands("Band_File").Tools("Mnu_New").Caption = GetCaption4MenuBar("Mnu_New")
    BarMenu.Bands("Band_File").Tools("Mnu_New").Description = GetDescription4StatusBar("Mnu_New")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "SepMnu_Save"
    BarMenu.Bands("Band_File").Tools("SepMnu_Save").ControlType = ddTTSeparator
    
    'File-Save
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_Save"
    BarMenu.Bands("Band_File").Tools("Mnu_Save").SetPicture 0, gResource.GetBitmap(IDB_STD_SAVE16), &HC0C0C0
    If m_App.Language <> 1 Then
        BarMenu.Bands("Band_File").Tools("Mnu_Save").Caption = GetCaption4MenuBar("Mnu_Save")
    Else
        BarMenu.Bands("Band_File").Tools("Mnu_Save").Caption = GetCaption4MenuBar("Mnu_Save")
    End If
    BarMenu.Bands("Band_File").Tools("Mnu_Save").Description = GetDescription4StatusBar("Mnu_Save")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "SepMnu_PrePrint"
    BarMenu.Bands("Band_File").Tools("SepMnu_PrePrint").ControlType = ddTTSeparator
    
    'File-PrePrint
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_PrePrint"
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIEW16), &HC0C0C0
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Caption = GetCaption4MenuBar("Mnu_PrePrint")
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Description = GetDescription4StatusBar("Mnu_PrePrint")
    
    'File-Print
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_Print"
    BarMenu.Bands("Band_File").Tools("Mnu_Print").SetPicture 0, gResource.GetBitmap(IDB_STD_PRINT16), &HC0C0C0
    If m_App.Language <> 1 Then
        BarMenu.Bands("Band_File").Tools("Mnu_Print").Caption = GetCaption4MenuBar("Mnu_Print")
    Else
        BarMenu.Bands("Band_File").Tools("Mnu_Print").Caption = GetCaption4MenuBar("Mnu_Print")
    End If
    BarMenu.Bands("Band_File").Tools("Mnu_Print").Description = GetDescription4StatusBar("Mnu_Print")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "SepMnu_Exit"
    BarMenu.Bands("Band_File").Tools("SepMnu_Exit").ControlType = ddTTSeparator
    
    'File-Exit
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_Exit"
    BarMenu.Bands("Band_File").Tools("Mnu_Exit").Caption = GetCaption4MenuBar("Mnu_Exit")
    BarMenu.Bands("Band_File").Tools("Mnu_Exit").Description = GetDescription4StatusBar("Mnu_Exit")
    
    'Edit
    BarMenu.Bands.Add "Band_Edit"
    BarMenu.Bands("Band_Edit").Type = ddBTPopup
    BarMenu.Bands("Band_Edit").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "Edit"
    BarMenu.Bands("Band_Menu").Tools("Edit").SubBand = "Band_Edit"
    BarMenu.Bands("Band_Menu").Tools("Edit").Caption = GetCaption4MenuBar("Edit")
    BarMenu.Bands("Band_Menu").Tools("Edit").Description = GetDescription4StatusBar("Edit")
    
    'Edit-Delete
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Delete"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").SetPicture 0, gResource.GetBitmap(IDB_STD_DELETE16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Caption = GetCaption4MenuBar("Mnu_Delete")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Description = GetDescription4StatusBar("Mnu_Delete")
    
    'Edit-Clear
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Clear"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").SetPicture 0, gResource.GetBitmap(IDB_STD_CLEAR16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Caption = GetCaption4MenuBar("Mnu_Clear")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Description = GetDescription4StatusBar("Mnu_Clear")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "SepMnu_Cut"
    BarMenu.Bands("Band_Edit").Tools("SepMnu_Cut").ControlType = ddTTSeparator
    
    'Edit-Cut
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Cut"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").SetPicture 0, gResource.GetBitmap(IDB_STD_CUT16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Caption = GetCaption4MenuBar("Mnu_Cut")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Description = GetDescription4StatusBar("Mnu_Cut")
    
    'Edit-Copy
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Copy"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").SetPicture 0, gResource.GetBitmap(IDB_STD_COPY16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Caption = GetCaption4MenuBar("Mnu_Copy")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Description = GetDescription4StatusBar("Mnu_Copy")
    
    'Edit-Paste
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Paste"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").SetPicture 0, gResource.GetBitmap(IDB_STD_PASTE16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Caption = GetCaption4MenuBar("Mnu_Paste")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Description = GetDescription4StatusBar("Mnu_Paste")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "SepMnu_NewSearch"
    BarMenu.Bands("Band_Edit").Tools("SepMnu_NewSearch").ControlType = ddTTSeparator
    
    'Edit-NewSearch
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_NewSearch"
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_FIND16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Caption = GetCaption4MenuBar("Mnu_NewSearch")
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Description = GetDescription4StatusBar("Mnu_NewSearch")
    
    'Edit-ExecuteSearch
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_ExecuteSearch"
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_EXECUTE16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Caption = GetCaption4MenuBar("Mnu_ExecuteSearch")
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Description = GetDescription4StatusBar("Mnu_ExecuteSearch")
    
    'Edit-SearchPrevious
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_SearchPrevious"
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIOUS16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Caption = GetCaption4MenuBar("Mnu_SearchPrevious")
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Description = GetDescription4StatusBar("Mnu_SearchPrevious")
    
    'Edit-SearchNext
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_SearchNext"
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").SetPicture 0, gResource.GetBitmap(IDB_STD_NEXT16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Caption = GetCaption4MenuBar("Mnu_SearchNext")
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Description = GetDescription4StatusBar("Mnu_SearchNext")
    
    'View
    BarMenu.Bands.Add "Band_View"
    BarMenu.Bands("Band_View").Type = ddBTPopup
    BarMenu.Bands("Band_View").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "View"
    BarMenu.Bands("Band_Menu").Tools("View").SubBand = "Band_View"
    BarMenu.Bands("Band_Menu").Tools("View").Caption = GetCaption4MenuBar("View")
    BarMenu.Bands("Band_Menu").Tools("View").Description = GetDescription4StatusBar("View")
    
    'View-FormView
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_FormView"
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").SetPicture 0, gResource.GetBitmap(IDB_STD_FORM16), &HC0C0C0
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").Caption = GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    
    'View-TableView
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_TableView"
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").SetPicture 0, gResource.GetBitmap(IDB_STD_GRID16), &HC0C0C0
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").Caption = GetCaption4MenuBar("Mnu_TableView")
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    
    'View - SearchFilter
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_SearchFilter"
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").SetPicture 0, gResource.GetBitmap(IDB_FILTRO16), &HC0C0C0
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Caption = GetCaption4MenuBar("Mnu_SearchFilter")
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Description = GetDescription4StatusBar("Mnu_SearchFilter")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "SepMnu_Folders"
    BarMenu.Bands("Band_View").Tools("SepMnu_Folders").ControlType = ddTTSeparator
    
    'View-Folders
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_Folders"
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Caption = GetCaption4MenuBar("Mnu_Folders")
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Description = GetDescription4StatusBar("Mnu_Folders")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "SepMnu_ToolBar"
    BarMenu.Bands("Band_View").Tools("SepMnu_ToolBar").ControlType = ddTTSeparator
    
    'View-ToolBar
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_ToolBar"
    BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Caption = GetCaption4MenuBar("Mnu_ToolBar")
    BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Description = GetDescription4StatusBar("Mnu_ToolBar")
    
    'Tools
    BarMenu.Bands.Add "Band_Tools"
    BarMenu.Bands("Band_Tools").Type = ddBTPopup
    BarMenu.Bands("Band_Tools").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "Tools"
    BarMenu.Bands("Band_Menu").Tools("Tools").SubBand = "Band_Tools"
    BarMenu.Bands("Band_Menu").Tools("Tools").Caption = GetCaption4MenuBar("Tools")
    BarMenu.Bands("Band_Menu").Tools("Tools").Description = GetDescription4StatusBar("Tools")
    
    'Tools-Export
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Tools").Tools.Add ToolID, "Mnu_Export"
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").ControlType = ddTTLabel
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").SubBand = "Mnu_Band_Export"
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Caption = GetCaption4MenuBar("Mnu_Export")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Description = GetDescription4StatusBar("Mnu_Export")
    
    'Tools-Options
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Tools").Tools.Add ToolID, "Mnu_Options"
    BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Caption = GetCaption4MenuBar("Mnu_Options")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Description = GetDescription4StatusBar("Mnu_Options")
    
    'Tools-Export
    BarMenu.Bands.Add "Mnu_Band_Export"
    BarMenu.Bands("Mnu_Band_Export").Type = ddBTPopup
    BarMenu.Bands("Mnu_Band_Export").DockingArea = ddDAPopup
    
    'Tools-Export-ExportPDF
    ToolID = ToolID + 1
    BarMenu.Bands("Mnu_Band_Export").Tools.Add ToolID, "Mnu_ExportPDF"
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").SetPicture 0, gResource.GetBitmap(IDB_ACROBAT_16), &HC0C0C0
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Caption = GetCaption4MenuBar("Mnu_ExportPDF")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Description = GetDescription4StatusBar("Mnu_ExportPDF")
    
    'Tools-Export-ExportWord
    ToolID = ToolID + 1
    BarMenu.Bands("Mnu_Band_Export").Tools.Add ToolID, "Mnu_ExportWord"
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").SetPicture 0, gResource.GetBitmap(IDB_STD_WORD16), &HC0C0C0
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Caption = GetCaption4MenuBar("Mnu_ExportWord")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Description = GetDescription4StatusBar("Mnu_ExportWord")
    
    'Tools-Export-ExportExcel
    ToolID = ToolID + 1
    BarMenu.Bands("Mnu_Band_Export").Tools.Add ToolID, "Mnu_ExportExcel"
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").SetPicture 0, gResource.GetBitmap(IDB_STD_EXCEL16), &HC0C0C0
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Caption = GetCaption4MenuBar("Mnu_ExportExcel")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Description = GetDescription4StatusBar("Mnu_ExportExcel")
    
    'Tools-Export-ExportHtml
    ToolID = ToolID + 1
    BarMenu.Bands("Mnu_Band_Export").Tools.Add ToolID, "Mnu_ExportHtml"
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").SetPicture 0, gResource.GetBitmap(IDB_STD_HTML16), &HC0C0C0
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Caption = GetCaption4MenuBar("Mnu_ExportHtml")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Description = GetDescription4StatusBar("Mnu_ExportHtml")

    'Help
    BarMenu.Bands.Add "Band_Help"
    BarMenu.Bands("Band_Help").Type = ddBTPopup
    BarMenu.Bands("Band_Help").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "Help"
    BarMenu.Bands("Band_Menu").Tools("Help").Caption = GetCaption4MenuBar("Help")
    BarMenu.Bands("Band_Menu").Tools("Help").Description = GetDescription4StatusBar("Help")
    BarMenu.Bands("Band_Menu").Tools("Help").SubBand = "Band_Help"
    
    'Help-HelpOnLine
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_HelpOnLine"
    BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").Caption = GetCaption4MenuBar("Mnu_HelpOnLine")
    BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").Description = GetDescription4StatusBar("Mnu_HelpOnLine")
    
    'Help-Arg
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_Arg"
    BarMenu.Bands("Band_Help").Tools("Mnu_Arg").Caption = GetCaption4MenuBar("Mnu_Arg")
    BarMenu.Bands("Band_Help").Tools("Mnu_Arg").Description = GetDescription4StatusBar("Mnu_Arg")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "SepMnu_Web"
    BarMenu.Bands("Band_Help").Tools("SepMnu_Web").ControlType = ddTTSeparator
    
    'Help-Web
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_Web"
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").SetPicture 0, gResource.GetBitmap(IDB_DMT_WEB16), &HC0C0C0
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").Caption = GetCaption4MenuBar("Mnu_Web")
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").Description = GetDescription4StatusBar("Mnu_Web")
    
    'Help-Blog
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_Agg_Web"
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").SetPicture 0, gResource.GetBitmap(IDB_AGG_WEB16), &HC0C0C0
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").Caption = GetCaption4MenuBar("Mnu_Agg_Web")
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").Description = GetDescription4StatusBar("Mnu_Agg_Web")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "SepMnu_Info"
    BarMenu.Bands("Band_Help").Tools("SepMnu_Info").ControlType = ddTTSeparator
    
    'Help-Info
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_Info"
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Caption = GetCaption4MenuBar("Mnu_Info")
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Description = GetDescription4StatusBar("Mnu_Info")
    
    'PopUp
    BarMenu.Bands.Add "Band_PopUp"
    BarMenu.Bands("Band_PopUp").Type = ddBTPopup
    BarMenu.Bands("Band_PopUp").DockingArea = ddDAPopup
    
    'PopUp-RunApplication
    ToolID = ToolID + 1
    BarMenu.Bands("Band_PopUp").Tools.Add ToolID, "Mnu_RunApplication"
    BarMenu.Bands("Band_PopUp").Tools("Mnu_RunApplication").Caption = GetCaption4MenuBar("Mnu_RunApplication")
    
    'PopUp-SearchObject
    ToolID = ToolID + 1
    BarMenu.Bands("Band_PopUp").Tools.Add ToolID, "Mnu_SearchObject"
    BarMenu.Bands("Band_PopUp").Tools("Mnu_SearchObject").Caption = GetCaption4MenuBar("Mnu_SearchObject")
    
    BarMenu.RecalcLayout
End Sub

'**+
'Nome: InitToolBar
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Inizializzazione della ToolBar
'
'**/
Private Sub InitToolBar(ByRef ToolID As Integer)


    BarMenu.Bands.Add "Standard"
    BarMenu.Bands("Standard").DockLine = 2
    BarMenu.Bands("Standard").Type = ddBTNormal
    BarMenu.Bands("Standard").Flags = ddBFDockTop Or ddBFDockLeft Or ddBFFloat Or ddBFDockRight Or ddBFDockBottom
    BarMenu.Bands("Standard").GrabHandleStyle = ddGSNormal
    BarMenu.Bands.Add BAND_CLOSE_PREVIEW
    BarMenu.Bands(BAND_CLOSE_PREVIEW).DockLine = 2
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Type = ddBTMenuBar
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Caption = "Chiudi"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).DockingArea = ddDATop
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Visible = False

    'New
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "New"
    BarMenu.Bands("Standard").Tools("New").ToolTipText = GetToolTipText4ToolBar("New")
    BarMenu.Bands("Standard").Tools("New").Description = GetDescription4StatusBar("New")
    
    'Save
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Save"
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep2"
    BarMenu.Bands("Standard").Tools("Sep2").ControlType = ddTTSeparator
    BarMenu.Bands("Standard").Tools("Save").ToolTipText = GetToolTipText4ToolBar("Save")
    BarMenu.Bands("Standard").Tools("Save").Description = GetDescription4StatusBar("Save")
    
    'Print
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Print"
    BarMenu.Bands("Standard").Tools("Print").ToolTipText = GetToolTipText4ToolBar("Print")
    BarMenu.Bands("Standard").Tools("Print").Description = GetDescription4StatusBar("Print")
    
    'PrePrint
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "PrePrint"
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep3"
    BarMenu.Bands("Standard").Tools("Sep3").ControlType = ddTTSeparator
    BarMenu.Bands("Standard").Tools("PrePrint").ToolTipText = GetToolTipText4ToolBar("PrePrint")
    BarMenu.Bands("Standard").Tools("PrePrint").Description = GetDescription4StatusBar("PrePrint")
    
    'Cut
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Cut"
    BarMenu.Bands("Standard").Tools("Cut").ToolTipText = GetToolTipText4ToolBar("Cut")
    BarMenu.Bands("Standard").Tools("Cut").Description = GetDescription4StatusBar("Cut")
    
    'Copy
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Copy"
    BarMenu.Bands("Standard").Tools("Copy").ToolTipText = GetToolTipText4ToolBar("Copy")
    BarMenu.Bands("Standard").Tools("Copy").Description = GetDescription4StatusBar("Copy")
    
    'Paste
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Paste"
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep"
    BarMenu.Bands("Standard").Tools("Sep").ControlType = ddTTSeparator
    BarMenu.Bands("Standard").Tools("Paste").ToolTipText = GetToolTipText4ToolBar("Paste")
    BarMenu.Bands("Standard").Tools("Paste").Description = GetDescription4StatusBar("Paste")
    
    'Delete
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Delete"
    BarMenu.Bands("Standard").Tools("Delete").ToolTipText = GetToolTipText4ToolBar("Delete")
    BarMenu.Bands("Standard").Tools("Delete").Description = GetDescription4StatusBar("Delete")
    
    'Clear
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Clear"
    BarMenu.Bands("Standard").Tools("Clear").ToolTipText = GetToolTipText4ToolBar("Clear")
    BarMenu.Bands("Standard").Tools("Clear").Description = GetDescription4StatusBar("Clear")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "SepNewSearch"
    BarMenu.Bands("Standard").Tools("SepNewSearch").ControlType = ddTTSeparator
    
    'NewSearch
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "NewSearch"
    BarMenu.Bands("Standard").Tools("NewSearch").ToolTipText = GetToolTipText4ToolBar("NewSearch")
    BarMenu.Bands("Standard").Tools("NewSearch").Description = GetDescription4StatusBar("NewSearch")
    
    'ExecuteSearch
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "ExecuteSearch"
    BarMenu.Bands("Standard").Tools("ExecuteSearch").ToolTipText = GetToolTipText4ToolBar("ExecuteSearch")
    BarMenu.Bands("Standard").Tools("ExecuteSearch").Description = GetDescription4StatusBar("ExecuteSearch")
    
    'ChangeView
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "ChangeView"
    BarMenu.Bands("Standard").Tools("ChangeView").ControlType = ddTTButtonDropDown
    BarMenu.Bands("Standard").Tools("ChangeView").SubBand = "Band_ChangeView"
    BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
    BarMenu.Bands("Standard").Tools("ChangeView").Description = GetDescription4StatusBar("ChangeView")
    BarMenu.Bands.Add "Band_ChangeView"
    BarMenu.Bands("Band_ChangeView").Type = ddBTPopup
    BarMenu.Bands("Band_ChangeView").DockingArea = ddDATop
    
    'ChangeView - Form
    ToolID = ToolID + 1
    BarMenu.Bands("Band_ChangeView").Tools.Add ToolID, "Mnu_FormView"
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").SetPicture 0, gResource.GetBitmap(IDB_STD_FORM16), &HC0C0C0
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Caption = GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    
    'ChangeView - Tabella
    ToolID = ToolID + 1
    BarMenu.Bands("Band_ChangeView").Tools.Add ToolID, "Mnu_TableView"
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").SetPicture 0, gResource.GetBitmap(IDB_STD_GRID16), &HC0C0C0
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Caption = GetCaption4MenuBar("Mnu_TableView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    
     'ChangeView - Filtro
    ToolID = ToolID + 1
    BarMenu.Bands("Band_ChangeView").Tools.Add ToolID, "Mnu_SearchFilter"
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").SetPicture 0, gResource.GetBitmap(IDB_FILTRO16), &HC0C0C0
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Caption = GetCaption4MenuBar("Mnu_SearchFilter")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Description = GetDescription4StatusBar("Mnu_SearchFilter")
    
    'SearchPrevious
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "SearchPrevious"
    BarMenu.Bands("Standard").Tools("SearchPrevious").ToolTipText = GetToolTipText4ToolBar("SearchPrevious")
    BarMenu.Bands("Standard").Tools("SearchPrevious").Description = GetDescription4StatusBar("SearchPrevious")
    
    'SearchNext
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "SearchNext"
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep4"
    BarMenu.Bands("Standard").Tools("Sep4").ControlType = ddTTSeparator
    BarMenu.Bands("Standard").Tools("SearchNext").ToolTipText = GetToolTipText4ToolBar("SearchNext")
    BarMenu.Bands("Standard").Tools("SearchNext").Description = GetDescription4StatusBar("SearchNext")
        
    'Export
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Export"
    BarMenu.Bands("Standard").Tools("Export").ControlType = ddTTButtonDropDown
    BarMenu.Bands("Standard").Tools("Export").SubBand = "Band_Export"
    BarMenu.Bands("Standard").Tools("Export").ToolTipText = GetToolTipText4ToolBar("Export")
    BarMenu.Bands("Standard").Tools("Export").Description = GetDescription4StatusBar("Mnu_Export")
    BarMenu.Bands.Add "Band_Export"
    BarMenu.Bands("Band_Export").Type = ddBTPopup
    BarMenu.Bands("Band_Export").DockingArea = ddDATop
    
    'ExportPDF
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Export").Tools.Add ToolID, "ExportPDF"
    BarMenu.Bands("Band_Export").Tools("ExportPDF").Caption = GetCaption4MenuBar("Mnu_ExportPDF")
    BarMenu.Bands("Band_Export").Tools("ExportPDF").ToolTipText = GetToolTipText4ToolBar("ExportPDF")
    BarMenu.Bands("Band_Export").Tools("ExportPDF").Description = GetDescription4StatusBar("ExportPDF")
    
    'ExportWord
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Export").Tools.Add ToolID, "ExportWord"
    BarMenu.Bands("Band_Export").Tools("ExportWord").Caption = GetCaption4MenuBar("Mnu_ExportWord")
    BarMenu.Bands("Band_Export").Tools("ExportWord").ToolTipText = GetToolTipText4ToolBar("ExportWord")
    BarMenu.Bands("Band_Export").Tools("ExportWord").Description = GetDescription4StatusBar("ExportWord")
    
    'ExportExcel
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Export").Tools.Add ToolID, "ExportExcel"
    BarMenu.Bands("Band_Export").Tools("ExportExcel").Caption = GetCaption4MenuBar("Mnu_ExportExcel")
    BarMenu.Bands("Band_Export").Tools("ExportExcel").ToolTipText = GetToolTipText4ToolBar("ExportExcel")
    BarMenu.Bands("Band_Export").Tools("ExportExcel").Description = GetDescription4StatusBar("ExportExcel")
    
    'ExportHtml
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Export").Tools.Add ToolID, "ExportHtml"
    BarMenu.Bands("Band_Export").Tools("ExportHtml").Caption = GetCaption4MenuBar("Mnu_ExportHtml")
    BarMenu.Bands("Band_Export").Tools("ExportHtml").ToolTipText = GetToolTipText4ToolBar("ExportHtml")
    BarMenu.Bands("Band_Export").Tools("ExportHtml").Description = GetDescription4StatusBar("ExportHtml")
    
    'Bottone chiusura anteprima
    ToolID = ToolID + 1
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools.Add ToolID, "ClosePreview"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").Style = ddSText
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").Caption = "&Chiudi"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").ToolTipText = "Chiudi anteprima"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").Description = "Esci da modalità Anteprima di stampa"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").ControlType = ddTTButton
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").Visible = True
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep5"
    BarMenu.Bands("Standard").Tools("Sep5").ControlType = ddTTSeparator

    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Web"
    BarMenu.Bands("Standard").Tools("Web").ToolTipText = GetToolTipText4ToolBar("Web")
    BarMenu.Bands("Standard").Tools("Web").Description = GetDescription4StatusBar("Web")

    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Agg_Web"
    BarMenu.Bands("Standard").Tools("Agg_Web").ToolTipText = GetToolTipText4ToolBar("Agg_Web")
    BarMenu.Bands("Standard").Tools("Agg_Web").Description = GetDescription4StatusBar("Agg_Web")
    
'    BarMenu.Bands.Add "Band_Contratto"
'    BarMenu.Bands("Band_Contratto").Type = ddBTNormal
'    BarMenu.Bands("Band_Contratto").DockingArea = ddDARight
'    BarMenu.Bands("Band_Contratto").DockLine = 3
'    BarMenu.Bands("Band_Contratto").Flags = ddBFDockTop Or ddBFDockLeft Or ddBFFloat Or ddBFDockRight Or ddBFDockBottom
'    BarMenu.Bands("Band_Contratto").GrabHandleStyle = ddGSNormal
'
    'Menù contratto - Storia adeguamento - Form
'    ToolID = ToolID + 1
'    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep6"
'    BarMenu.Bands("Standard").Tools("Sep6").ControlType = ddTTSeparator
    
    BarMenu.Bands.Add "StandardPO"
    BarMenu.Bands("StandardPO").DockLine = 3
    BarMenu.Bands("StandardPO").Type = ddBTNormal
    BarMenu.Bands("StandardPO").Flags = ddBFDockTop Or ddBFDockLeft Or ddBFFloat Or ddBFDockRight Or ddBFDockBottom
    BarMenu.Bands("StandardPO").GrabHandleStyle = ddGSNormal
    
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdAltreInfo"
    BarMenu.Bands("StandardPO").Tools("cmdAltreInfo").SetPicture 0, gResource.GetBitmap(IDB_CONF_OBJ16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdAltreInfo").ToolTipText = "Altre informazioni del contratto" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdAltreInfo").Description = "Altre informazioni"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdAltreInfo").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdAltreInfo").Caption = "Informazioni aggiuntive"  'GetDescription4StatusBar("Mnu_FormView")
    
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep8"
    BarMenu.Bands("StandardPO").Tools("Sep8").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdDocumentazione"
    BarMenu.Bands("StandardPO").Tools("cmdDocumentazione").SetPicture 0, gResource.GetBitmap(IDB_NEWWITHWIZARD16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdDocumentazione").ToolTipText = "Documentazione del contratto" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdDocumentazione").Description = "Documentazione"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdDocumentazione").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdDocumentazione").Caption = "Documentazione"  'GetDescription4StatusBar("Mnu_FormView")
    

    
'    ToolID = ToolID + 1
'    BarMenu.Bands("Standard").Tools.Add ToolID, "cmdClausole"
'    BarMenu.Bands("Standard").Tools("cmdClausole").SetPicture 0, gResource.GetBitmap(IDB_ACT_ORDER_CUSTOMER_16), &HC0C0C0
'    BarMenu.Bands("Standard").Tools("cmdClausole").ToolTipText = "Clausole del contratto" '"GetCaption4MenuBar("Mnu_FormView")
'    BarMenu.Bands("Standard").Tools("cmdClausole").Description = "Clausole del contratto"  'GetDescription4StatusBar("Mnu_FormView")
'
'    ToolID = ToolID + 1
'    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep10"
'    BarMenu.Bands("Standard").Tools("Sep10").ControlType = ddTTSeparator
    



    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep14"
    BarMenu.Bands("StandardPO").Tools("Sep14").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdGeneraRate"
    BarMenu.Bands("StandardPO").Tools("cmdGeneraRate").SetPicture 0, gResource.GetBitmap(IDB_AGG_PROGRESSIVI16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdGeneraRate").ToolTipText = "Genera rate del contratto" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdGeneraRate").Description = "Genera rate del contratto"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdGeneraRate").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdGeneraRate").Caption = "Genera rate del contratto"  'GetDescription4StatusBar("Mnu_FormView")


    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep12"
    BarMenu.Bands("StandardPO").Tools("Sep12").ControlType = ddTTSeparator
    
    
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdFatturazione"
    BarMenu.Bands("StandardPO").Tools("cmdFatturazione").SetPicture 0, gResource.GetBitmap(IDB_INCASSO16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdFatturazione").ToolTipText = "Fatturazione delle rate del contratto" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdFatturazione").Description = "Fatturazione delle rate del contratto"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdFatturazione").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdFatturazione").Caption = "Fatturazione"  'GetDescription4StatusBar("Mnu_FormView")

    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep13"
    BarMenu.Bands("StandardPO").Tools("Sep13").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdAdeguamentoIstat"
    BarMenu.Bands("StandardPO").Tools("cmdAdeguamentoIstat").SetPicture 0, gResource.GetBitmap(IDB_AGG_PROGRESSIVI16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdAdeguamentoIstat").ToolTipText = "Elaborazione adeguamento istat" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdAdeguamentoIstat").Description = "Elaborazione adeguamento istat"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdAdeguamentoIstat").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdAdeguamentoIstat").Caption = "Adeguamento istat"  'GetDescription4StatusBar("Mnu_FormView")
    
    

    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep11"
    BarMenu.Bands("StandardPO").Tools("Sep11").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdStampaModelliTesta"
    BarMenu.Bands("StandardPO").Tools("cmdStampaModelliTesta").SetPicture 0, gResource.GetBitmap(IDB_ACT_INVOICE_BILL_PAYMENT_16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdStampaModelliTesta").ToolTipText = "Stampa contratto da modello" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdStampaModelliTesta").Description = "Stampa contratto da modello"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdStampaModelliTesta").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdStampaModelliTesta").Caption = "Stampa da modello"  'GetDescription4StatusBar("Mnu_FormView")


    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep9"
    BarMenu.Bands("StandardPO").Tools("Sep9").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdStoriaAdeg"
    BarMenu.Bands("StandardPO").Tools("cmdStoriaAdeg").SetPicture 0, gResource.GetBitmap(IDB_CONFIG_REFRESH16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdStoriaAdeg").ToolTipText = "Storia adeguamento" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdStoriaAdeg").Description = "Storia adeguamento"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdStoriaAdeg").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdStoriaAdeg").Caption = "Storia adeguamento"  'GetDescription4StatusBar("Mnu_FormView")
    
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep7"
    BarMenu.Bands("StandardPO").Tools("Sep7").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdStoricoRate"
    BarMenu.Bands("StandardPO").Tools("cmdStoricoRate").SetPicture 0, gResource.GetBitmap(IDB_NOTA16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdStoricoRate").ToolTipText = "Rate storico" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdStoricoRate").Description = "Rate storico"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdStoricoRate").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdStoricoRate").Caption = "Storico rate"  'GetDescription4StatusBar("Mnu_FormView")

    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep20"
    BarMenu.Bands("StandardPO").Tools("Sep20").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "RichiediListaInterventi"
    BarMenu.Bands("StandardPO").Tools("RichiediListaInterventi").SetPicture 0, gResource.GetBitmap(IDB_NOTA16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("RichiediListaInterventi").ToolTipText = "Richiedi lista interventi" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("RichiediListaInterventi").Description = "Richiedi lista interventi"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("RichiediListaInterventi").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("RichiediListaInterventi").Caption = "Interventi"  'GetDescription4StatusBar("Mnu_FormView")
    
    BarMenu.RecalcLayout
End Sub





'**+
'Nome: ChooseAboutSaving
'
'Parametri:
'Ritorna i valori vbYes, vbNo o vbCancel a seconda della risposta data
'
'Valori di ritorno:
'
'Funzionalità:
'Richiesta della registrazione di un record
'**/
Private Function ChooseAboutSaving() As Integer
Dim Testo As String
On Error Resume Next
    If Me.cboTipoImpostazione.CurrentID <> 3 Then
        If (GET_CONTROLLO_IMPORTO_RATE(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) <> Me.txtImportoAttuale.Value) Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "Il totale delle rate non è uguale al totale del contratto" & vbCrLf
            Testo = Testo & "Vuoi continuare?"
            If MsgBox(Testo, vbCritical + vbYesNo, "Controllo dati contratto") = vbNo Then
                ChooseAboutSaving = vbCancel
                Exit Function
            End If
        End If
    End If


    If m_Changed Then
        gResource.CustomStrings.Clear
        gResource.CustomStrings.Add Chr(34) & m_App.FunctionName & Chr(34), 1

        ChooseAboutSaving = fnMsgQuestionWithCancel((gResource.GetCustomizedMessage(MESS_QUERYSAVE)), TheApp.FunctionName)
    End If
End Function

'**+
'Nome: ChooseAboutSavingOkCancel
'
'Parametri:
'
'Valori di ritorno:
'Ritorna i valori vbOK o vbCancel a seconda della risposta data
'
'Funzionalità:
'Come ChooseAboutSaving ma con pulsanti Ok e Annulla
'**/
Private Function ChooseAboutSavingOkCancel() As Integer
    Dim sRecord As String

    sRecord = IIf(m_Document.Fields(CAMPO_PER_CAPTION).Value <> Empty, m_Document.Fields(CAMPO_PER_CAPTION).Value, TheApp.FunctionName)
  
    gResource.CustomStrings.Clear
    gResource.CustomStrings.Add Chr(34) & sRecord & Chr(34), 1
    ChooseAboutSavingOkCancel = fnMsgQuestionOKCancel((gResource.GetCustomizedMessage(MESS_QUERYSAVE)), m_App.FunctionName)
    
End Function

'**+
'Nome: ActivateBarButtons
'
'Parametri:
'Buttons - Variabile lunga 8 byte con la combinazione
'della maschera di bit che indica di quali bottoni cambiare
'lo stato di abilitazione
'Enable - Valore booleano che indica lo stato di abilitazione
'da applicare
'
'Valori di ritorno:
'
'Funzionalità:
'Abilita o meno gruppi di bottoni e voci di menu
'**/
Private Sub ActivateBarButtons(ByVal Buttons As Currency, ByVal Enable As Boolean)

    'Pulsanti della Toolbar
    '----------------------
    If (Buttons And BTN_NEW) Then BarMenu.Bands("Standard").Tools("New").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_SAVE) Then BarMenu.Bands("Standard").Tools("Save").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_PRINT) Then BarMenu.Bands("Standard").Tools("Print").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_PREVIEW) Then BarMenu.Bands("Standard").Tools("PrePrint").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_CUT) Then BarMenu.Bands("Standard").Tools("Cut").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_COPY) Then BarMenu.Bands("Standard").Tools("Copy").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_PASTE) Then BarMenu.Bands("Standard").Tools("Paste").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_DELETE) Then BarMenu.Bands("Standard").Tools("Delete").Enabled = CheckRights("Cancellazione", Enable)
    If (Buttons And BTN_CLEAR) Then BarMenu.Bands("Standard").Tools("Clear").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_FIND) Then BarMenu.Bands("Standard").Tools("NewSearch").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_SEARCH) Then BarMenu.Bands("Standard").Tools("ExecuteSearch").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_VIEWMODE) Then BarMenu.Bands("Standard").Tools("ChangeView").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_PREVIOUS) Then BarMenu.Bands("Standard").Tools("SearchPrevious").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_NEXT) Then BarMenu.Bands("Standard").Tools("SearchNext").Enabled = CheckRights("Selezione", Enable)
    If Not oExportActivity Is Nothing Then
        If (Buttons And BTN_EXPORT) Then oExportActivity.EnableItems CheckRights("Stampa", Enable)
    End If
    If (Buttons And BTN_EXPORT) Then BarMenu.Bands("Standard").Tools("Export").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_WORD) Then BarMenu.Bands("Band_Export").Tools("ExportWord").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_EXCEL) Then BarMenu.Bands("Band_Export").Tools("ExportExcel").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_HTML) Then BarMenu.Bands("Band_Export").Tools("ExportHtml").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_PDF) Then BarMenu.Bands("Band_Export").Tools("ExportPDF").Enabled = CheckRights("Stampa", Enable)
    
    If (Buttons And BTN_SEARCHFORM) Then
        BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Enabled = CheckRights("Selezione", Enable)
        BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Checked = Not Enable
    End If
    
    If (Buttons And BTN_SEARCHTABLE) Then
        BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Enabled = CheckRights("Selezione", Enable)
        BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Checked = Not Enable
    End If
    
    If (Buttons And BTN_FILTER) Then BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Enabled = CheckRights("Selezione", Enable)
    
    'VOCI DI MENU
    '------------
    
    'Menu File
    '---------
    If (Buttons And BTN_NEW) Then BarMenu.Bands("Band_File").Tools("Mnu_New").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_SAVE) Then BarMenu.Bands("Band_File").Tools("Mnu_Save").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_PRINT) Then BarMenu.Bands("Band_File").Tools("Mnu_Print").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_PREVIEW) Then BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Enabled = CheckRights("Stampa", Enable)
    
    'Menu Edit
    '---------
    If (Buttons And BTN_CUT) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_COPY) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_PASTE) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_DELETE) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Enabled = CheckRights("Cancellazione", Enable)
    If (Buttons And BTN_CLEAR) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_FIND) Then BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_SEARCH) Then BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_PREVIOUS) Then BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_NEXT) Then BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Enabled = CheckRights("Selezione", Enable)
    
    'Menu Visualizza
    '---------------
    If (Buttons And BTN_FILTER) Then BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_SEARCHFORM) Then BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_SEARCHTABLE) Then BarMenu.Bands("Band_View").Tools("Mnu_TableView").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_VIEWMODE) Then
        BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled = CheckRights("Selezione", Enable)
    End If

    'Menu Export
    '-----------
    If Not oExportActivity Is Nothing Then
        If (Buttons And BTN_EXPORT) Then oExportActivity.EnableItems CheckRights("Stampa", Enable)
    End If
    If (Buttons And BTN_EXPORT) Then BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_WORD) Then BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_EXCEL) Then BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_HTML) Then BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_PDF) Then BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Enabled = CheckRights("Stampa", Enable)
End Sub


'**+
'Nome: NewSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni da compiere in caso di richiesta di una nuova ricerca
'**/
Private Sub NewSearch()

    'Refresh dello stato del Form
    m_Changed = False
    m_Saved = False
    m_Search = True
    
    'Annulla una eventuale operazione di inserimento di un nuovo record
    If m_Document.TableNew Then
        m_Document.AbortNew
        RefreshFormFields
    End If
    
    'Ripristina la vista del Form
    BrwMain.Visible = True
    
    'Predispone la modalità DefineFilter della Browse
    BrwMain.AbortFilterEdit = False
    BrwMain.GuiMode = dgFilterDefinition
    BrwMain.SetFocus
    
    'Refresh dello stato dei bottoni delle barre dei menu per la modalità ricerca
    SetStatus4Modality Find
    
End Sub

'**+
'Nome: ExecuteSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue la ricerca impostata.
'
'**/
Private Sub ExecuteSearch()
    Dim Cond As DmtGridCtl.dgCondition
    Dim Field As DmtDocManLib.Field
    Dim OLDCursor As Integer
    Dim sWhere As String
    
    
    'Gestione della clessidra
    OLDCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    
    'Se non è stato selezionato nessun filtro dal controllo DocTypeExplorer
    'viene creato un filtro temporaneo in memoria e reso il filtro attivo
    If Not m_FilterSelected Then
        
        'Comunica all'oggetto DocType i valori da usare per la ricerca
        sWhere = fnFillDocTypeCondition
        
        'Rimuove il filtro precedente
        m_DocType.RemoveFilter "Temp"
        
        'Crea un nuovo filtro temporaneo a partire dalle condizioni di ricerca
        'e viene reso filtro attivo
        Set m_ActiveFilter = m_DocType.AddFilterWithConditions("Temp")
        
        'Aggiunge al filtro eventuali condizioni aggiuntive restituite dalla funzione fnFillDocTypeCondition
        If sWhere <> "" Then m_ActiveFilter.AddCondition sWhere
        
    End If
    
    'Comunica al documento il nuovo filtro da usare
    Set m_Document.ActiveFilter = m_ActiveFilter
    
    'Viene effettuata la ricerca
    m_Document.OpenDoc
    
    
    'Assegnazione del riferimento alla fonte dati (binding sul recordset del documento)
    
    'rif13
    
    'Set BrwMain.Recordset = m_Document.Dataset.Recordset
    Set BrwMain.Recordset = m_Document.Data
    
    
    
    'Ripristina il cursore
    Screen.MousePointer = OLDCursor
    
    'Operazioni da effettuare dopo l'esecuzione della ricerca.
    AfterExecuteSearch
End Sub

'**+
'Nome: AfterExecuteSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Determina quali operazioni compiere dopo ExecuteSearch
'              in funzione dell'esito della ricerca.
'
'**/
Private Sub AfterExecuteSearch()

    If Not (m_Document.EOF = True And m_Document.BOF = True) Then
        'La ricerca ha avuto esito positivo
        'Attiva la vista tabellare
        BrwMain.Visible = True
        BrwMain.SetFocus

        'Imposta i menu e la toolbar per la modalità tabellare
        SetStatus4Modality Browse

        'Attiva le procedure di creazione di un nuovo filtro solo se l'ExecuteSearch
        'non è stata chiamata da una selezione del DocTypeExplorer

        'Se l'ExecuteSearch non è stata chiamata da un filtro del riquadro attività
        'si permette di salvare il nuovo filtro ed aggiungerlo nel ramo dei filtri.
        If Not m_FilterSelected Then
            oFiltersActivity.NewFilterBegin
        End If

        'Imposta i suggerimenti da visualizzare sulla Statusbar in funzione
        'della modalità di visualizzazione corrente.
        'Ad esempio in alcuni casi le frasi sono al Singolare/Plurare.
        'Le impostazioni sottostanti servono soltanto all'avvio del programma dopo la prima
        'ricerca. (in quanto ChangeView non è stata ancora eseguita)
        'La Sub RefreshDescriptions4StatusBar deve essere chiamata anche in ChangeView()--> Vedi.
        RefreshDescriptions4StatusBar

        m_Search = False
    Else
        'La ricerca ha avuto esito negativo. Viene mostrato un messaggio
        'e si torna in modalità ricerca.

        'Per questioni estetiche viene subito mostrata la modalità FilterDefinition
        'al posto della browse vuota e quindi viene mostrato il messaggio.
        BrwMain.GuiMode = dgFilterDefinition

        'Se si è selezionato il filtro "Nessun record" non occorre
        'visualizzare il messaggio
        If m_ActiveFilter.NothingSelected = False Or m_FilterSelected = False Then
            'Messaggio  "Nessun elemento trovato"
            sbMsgInfo gResource.GetMessage(MESS_NORECFOUND), m_App.FunctionName
        End If

        'Si torna in modalità form (modalità ricerca)
        OnNewSearch
    End If
    
End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 26/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: fnFillDocTypeCondition
'
'Parametri:
'
'Valori di ritorno: String - in base alle esigenze specifiche di una manutenzione è possibile montare ad hoc
'                                 una clausola WHERE che potrà poi essere presa in considerazione nel filtro di selezione
'                                 con il metodo AddCondition dell'oggetto DmtDocManLib.Filter
'
'Funzionalità: Comunica all'oggetto DocType i valori da usare per la ricerca
'
'**/
Private Function fnFillDocTypeCondition() As String
    Dim Field As DmtDocManLib.Field
    Dim Cond As DmtGridCtl.dgCondition
    Dim sWhere As String
    
    
    'NOTA per l'uso dei campi RANGE
    '--------------------------------------------------------------------------------------------------
    'E' consentito l'inserimento, nella modalità filtri e nel caso di campi di tipo range, del solo il valore iniziale
    '(in questo caso vengono filtrati tutti gli elementi maggiori o uguali a quello inserito)
    'o solo quello finale (in questo vengono filtrati tutti gli elementi minori o uguali a quello inserito).
    'Questo funzionamento vale per tutte le tipologie di campo.
    
    'Nel caso di condizione RANGE la sintassi da usare è del tipo della riga sotto:
    'm_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
    '--------------------------------------------------------------------------------------------------
    
    sWhere = vbNullString
    
    'Ripulisce la collezione Fields dell'oggetto DocType
    For Each Field In m_DocType.Fields
        Field.Value = Empty
    Next
    
    m_DocType.Fields("IDFiliale").Value = m_App.Branch
    
    
    
    'Comunica all'oggetto DocType i valori da usare per la ricerca
    For Each Cond In BrwMain.Conditions
        
        Select Case Cond.ConditionType
            'Condizione boolean
            Case dgCondTypeBoolean
                m_DocType.Fields(Cond.FieldName).Value = IIf(IsEmpty(Cond.FromValue), Empty, Abs(CDbl(Cond.FromValue = "SI")))
                
            'Condizione associata ad una combo box
            Case dgCondTypeComboDB
                m_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValueID
            
            'Condizione di tipo text, numeric, data, time
            Case dgCondTypeText
                If Cond.RangeChecked = True Then
                    m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                Else
                    If Len(Cond.FromValue) > 0 Then
                        m_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValue
                    End If
                End If
            Case dgCondTypeNumber
                If Cond.RangeChecked = True Then
                    m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                Else
                    If Len(Cond.FromValue) > 0 Then
                        m_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValue
                    End If
                End If
            
            Case dgCondTypeDate
                If Cond.RangeChecked = True Then
                    m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                Else
                    If Len(Cond.FromValue) > 0 Then
                        m_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValue
                    End If
                End If
            
            Case dgCondTypeTime
                If Cond.RangeChecked = True Then
                    m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                Else
                    If Len(Cond.FromValue) > 0 Then
                        m_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValue
                    End If
                End If
          
            'Altre condizioni
            Case Else
                m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
  
        End Select
       
    Next Cond
        
    fnFillDocTypeCondition = sWhere
End Function



'**+
'Nome: CheckRights
'
'Parametri:
'ActionName - Nome della azione
'Enable - Valore da modificare o ritornare inalterato
'
'Valori di ritorno:
'Il valore in Enable o False se l'azione non è abilitata
'per il tipo di documento
'
'Funzionalità:
'Controlla se l'azione passata è abilitata per il tipo documento
'**/
Private Function CheckRights(ByVal ActionName As String, ByVal Enable As Boolean) As Boolean
    Dim Action As DmtDocManLib.Action
    Dim Dummy As String
    
    If m_DocType.Actions.Count = 0 Then
        CheckRights = Enable
        Exit Function
    End If
    For Each Action In m_DocType.Actions
        If Action.Name = "TUTTE LE AZIONI" Then
            CheckRights = Enable
            Exit Function
        End If
    Next
    On Error GoTo ActionNotFound
    Dummy = m_DocType.Actions(ActionName).Name
    CheckRights = Enable
    Exit Function
ActionNotFound:
    CheckRights = False
End Function

'**+
'Nome: SetFocusTabIndex0
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Da l'input focus al campo con TabIndex uguale a 0.
'**/
Private Sub SetFocusTabIndex0()
    On Error GoTo SetFocusTabIndex0_Error
    
    Dim ControlObject As Control
    Dim iIndex As Long
    Dim bError As Boolean
    
    If m_ControlTabIndex0 Is Nothing Then
        For Each ControlObject In frmMain.Controls
            iIndex = ControlObject.TabIndex
            If bError Then
                '**+ Controllo corrente non ha proprietà TabIndex,
                '    quindi va saltato.
                bError = False
            Else
                If ControlObject.TabIndex = 0 Then
                    Set m_ControlTabIndex0 = ControlObject
                    Exit For
                End If
            End If
        Next
    End If
    m_ControlTabIndex0.SetFocus

    Exit Sub
SetFocusTabIndex0_Error:
    bError = True
    Resume Next
End Sub

'**+
'Nome: IsFieldInput
'
'Parametri:
'Control - un oggetto Control da controllare
'
'Valori di ritorno:
'Se il controllo è abilitato all'input torna vero altrimenti falso
'
'Funzionalità:
'Controllo se un certo controllo è usabile come campo
'di input dei dati del Form
'**/
Private Function IsFieldInput(ByVal Control As Control) As Boolean
    'Controlla se il Controllo è di Immissione
    IsFieldInput = IsFieldInput Or TypeName(Control) = "TextBox"
    IsFieldInput = IsFieldInput Or TypeName(Control) = "CheckBox"
    IsFieldInput = IsFieldInput Or TypeName(Control) = "ComboBox"
    IsFieldInput = IsFieldInput Or TypeName(Control) = "OptionButton"
    
    'rif5 begin
    
    IsFieldInput = IsFieldInput Or TypeName(Control) = "DMTCombo"
    IsFieldInput = IsFieldInput Or TypeName(Control) = "Town"
    
    'rif5 end
    
    
End Function

'**+
'Nome: FieldPresent
'
'Parametri:
'Name - Nome del campo
'
'Valori di ritorno:
'Se il campo specificato nel parametro è presente nella
'collezione FormFields torna vero altrimenti torna falso
'
'Funzionalità:
'Controlla la presenza di un campo nella collezione FormFields
'**/
Private Function FieldPresent(ByVal Name As String) As Boolean
    Dim Field As FormField

    For Each Field In m_FormFields
        FieldPresent = (Name = Field.Name)
        If FieldPresent Then Exit For
    Next
End Function

'**+
'Nome: OnStart
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Altre inizializzazioni dopo quelle predefinite
'**/
Private Sub OnStart()
Dim sSQL As String
  
'SETTARE LE GRIGLIE DEI SOTTODOCUMENTI
    Dim cl As DmtGridCtl.dgColumnHeader

    'Inizializzazione della griglia adibita alla visualizzazione tabellare dei sotto-documenti
    '-------------------------------------------------------------------------------
   
    If Me.GrigliaRateContratto.ColumnsHeader.Count = 0 Then
        With Me.GrigliaRateContratto.ColumnsHeader
            .Add "IDRV_PORateContratto", "IDRV_PORateContratto", dgInteger, False, 500, dgAlignRight
            .Add "IDRV_POContratto", "IDRV_PORateContratto", dgInteger, False, 500, dgAlignRight
            .Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignRight
            .Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignRight
            .Add "NumeroRata", "N°", dgInteger, True, 500, 0, True, True, False
            .Add "DataRata", "Data", dgDate, True, 1500, 0, True, True, False
            .Add "DataInizioPeriodo", "Data inizio periodo", dgDate, False, 1500, 0
            .Add "DataFinePeriodo", "Data fine periodo", dgDate, False, 1500, 0
            
            .Add "IDPagamentoRata", "IDPagamentoRata", dgInteger, False, 500, dgAlignRight
            .Add "Pagamento", "Tipo pagamento", dgchar, True, 2500, 0, True, True, False
            Set cl = .Add("ImportoRata", "Importo", dgDouble, True, 2000, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .Add "Fatturata", "Fatturata", dgBoolean, True, 1000, dgAligncenter
            .Add "NonFatturare", "Non fatturare", dgBoolean, True, 1000, dgAligncenter
            .Add "IDOggettoCollegato", "IDOggettoCollegato", dgInteger, False, 500, dgAlignRight
            .Add "DescrizioneAdeguamento", "Adeguamento", dgchar, False, 3000, dgAlignleft
            .Add "IDRV_POProdotto", "IDRV_POProdotto", dgInteger, False, 500, dgAlignRight
            .Add "DescrizioneProdotto", "Prodotto", dgchar, False, 3000, dgAlignleft
            .Add "MatricolaProdotto", "Matricola", dgchar, False, 3000, dgAlignleft
            .Add "MatricolaProdottoDaContratto", "Matricola da contratto", dgchar, False, 3000, dgAlignleft
            
            .Add "IDRV_POContrattoProdotti", "IDRV_POContrattoProdotti", dgInteger, False, 500, dgAlignRight
            .Add "IDArtProdContratto", "IDArtProdContratto", dgInteger, False, 500, dgAlignRight
            .Add "CodiceArticoloProdContratto", "Cod. art. contr. prod.", dgchar, False, 3000, dgAlignleft
            .Add "ArticoloProdContratto", "Descr. art. contr. prod.", dgchar, False, 3000, dgAlignleft
            
            .Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignRight
            .Add "CodiceArticolo", "Codice articolo", dgchar, False, 3000, dgAlignleft
            .Add "Articolo", "Descrizione articolo", dgchar, False, 3000, dgAlignleft
            
            
        End With
    End If
    Me.GrigliaRateContratto.EnableMove = True
    
    If Me.GrigliaServizi.ColumnsHeader.Count = 0 Then
        With Me.GrigliaServizi.ColumnsHeader
            .Add "IDRV_POContrattoServizi", "IDRV_POContrattoServizi", dgInteger, False, 500, 0, True, True, False
            .Add "IDRV_POContratto", "IDRV_POContratto", dgInteger, False, 500, 0, True, True, False
            .Add "IDRV_POStoriaContratto", "IDRV_POStoriaContratto", dgInteger, False, 500, 0, True, True, False
            
            .Add "IDArticolo", "IDArticolo", dgInteger, False, 500, 0, True, True, False
            .Add "CodiceArticolo", "Codice articolo", dgchar, True, 1500, 0, True, True, False
            .Add "Articolo", "Articolo", dgchar, True, 2500, 0, True, True, False
            .Add "IDRV_POCriterioRicorrenza", "IDRV_POCriterioRicorrenza", dgInteger, False, 500, 0, True, True, False
            .Add "CriterioRicorrenza", "Criterio di ricorrenza", dgchar, True, 1500, 0, True, True, False
            
            .Add "OgniNumeroGiorni", "Ogni n° giorni", dgInteger, True, 1000, 0, True, True, False
            .Add "OgniNumeroMesi", "Ogni n° mesi", dgInteger, True, 1000, 0, True, True, False
            .Add "OgniNumeroSettimane", "Ogni n° settimane", dgInteger, True, 1000, 0, True, True, False
            
            .Add "IDRV_POTipoDataInizioRicorrenza", "IDRV_POTipoDataInizioRicorrenza", dgInteger, False, 500, 0, True, True, False
            .Add "TipoDataInizioRicorrenza", "Data inizio ric.", dgchar, True, 1500, 0, True, True, False
            .Add "GiornoInizioRicorrenza", "Giorno inizio ric.", dgInteger, True, 1000, 0, True, True, False
            .Add "MeseInizioRicorrenza", "Mese inizio ric.", dgInteger, True, 1000, 0, True, True, False
            .Add "TipoAnnoInizioRicorrenza", "Anno inizio ric.", dgVarChar, True, 2500, 0, True, True, False
            
            .Add "IDRV_POTipoDataFineRicorrenza", "IDRV_POTipoDataFineRicorrenza", dgInteger, False, 500, 0, True, True, False
            .Add "TipoDataFineRicorrenza", "Data fine ric.", dgchar, True, 1500, 0, True, True, False
            .Add "GiornoFineRicorrenza", "Giorno fine ric.", dgInteger, True, 1000, 0, True, True, False
            .Add "MeseFineRicorrenza", "Mese fine ric.", dgInteger, True, 1000, 0, True, True, False
            .Add "TipoAnnoFineRicorrenza", "Anno fine ric.", dgVarChar, True, 2500, 0, True, True, False
            
            .Add "NumeroRicorrenze", "N° ricorrenze", dgInteger, True, 1000, 0, True, True, False

            
            
        End With
    End If
    Me.GrigliaServizi.EnableMove = True

    If Me.GrigliaAdeg.ColumnsHeader.Count = 0 Then
        With Me.GrigliaAdeg.ColumnsHeader
            .Add "IDRV_POContrattoAdeguamento", "IDRV_POContrattoAdeguamento", dgInteger, False, 500, 0, True, True, False
            .Add "IDRV_POContratto", "IDRV_POContratto", dgInteger, False, 500, 0, True, True, False
            .Add "IDRV_POContrattoPadre", "IDRV_POContrattoPadre", dgInteger, False, 500, 0, True, True, False
            .Add "DataStipula", "Data stipula", dgDate, False, 1500, 0, True, True, False
            .Add "DataDecorrenza", "Data decorrenza", dgDate, True, 1500, 0, True, True, False
            .Add "DataFineAdeguamento", "Data scadenza", dgDate, True, 1500, 0, True, True, False
            Set cl = .Add("Importo", "Importo", dgCurrency, True, 2500, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .Add "RiportaProssimoRinnovo", "Rip. prox rinn.", dgBoolean, True, 1000, dgAligncenter, True, True, False
            .Add "AdeguaContrattoAttuale", "Adegua rate", dgBoolean, True, 1000, dgAligncenter, True, True, False
            .Add "RiportaAnnotazioni", "Rip. note prox rinn.", dgBoolean, True, 1000, dgAligncenter, True, True, False
            .Add "IDRV_POTipoAdeguamento", "IDRV_POTipoAdeguamento", dgInteger, False, 500, dgAlignRight
            .Add "TipoAdeguamento", "Tipo adeguamento", dgchar, True, 2500, dgAlignleft
            .Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignRight
            .Add "CodiceArticolo", "Codice articolo", dgchar, False, 2500, dgAlignleft
            .Add "Articolo", "Articolo", dgchar, False, 2500, dgAlignleft
            .Add "IDArticoloServizio", "IDArticoloServizio", dgInteger, False, 500, dgAlignRight
            .Add "Annotazioni", "Annotazioni", dgchar, False, 1500, 0, True, True, False
            .Add "DescrizionePerFatturazione", "Descr. Fatt.", dgchar, False, 1500, 0, True, True, False
            .Add "IDRateizzazione", "IDRateizzazione", dgInteger, False, 500, 0, True, True, False
            .Add "Rateizzazione", "Tipo di rateizzazione", dgchar, False, 2500, dgAlignleft
            Set cl = .Add("ImportoAlRinnovo", "Importo al rinnovo", dgCurrency, True, 2500, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .Add "NumeroPartenza", "N° prog. iniz.", dgInteger, False, 500, 0, True, True, False
        End With
    End If
    Me.GrigliaAdeg.EnableMove = True

    If Me.GrigliaProd.ColumnsHeader.Count = 0 Then
        With Me.GrigliaProd.ColumnsHeader
            .Add "IDRV_POContrattoProdotti", "IDRV_POContrattoProdotti", dgInteger, False, 500, 0, True, True, False
            .Add "IDRV_POContratto", "IDRV_POContratto", dgInteger, False, 500, 0, True, True, False
            .Add "IDRV_POContrattoPadre", "IDRV_POContrattoPadre", dgInteger, False, 500, 0, True, True, False
            .Add "IDRV_POProdotto", "IDRV_POProdotto", dgInteger, False, 500, 0, True, True, False
            .Add "Prodotto", "Prodotto", dgchar, True, 2000, dgAlignleft
            .Add "ValoreIndentificativo", "Matricola", dgchar, True, 2500, dgAlignleft
            .Add "ProdottoGenerico", "Prodotto generico", dgBoolean, True, 1500, dgAligncenter
            Set cl = .Add("Quantita", "Quantità", dgDouble, True, 1800, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .Add "DescrizioneAggiuntiva", "Ubicazione", dgchar, True, 3000, dgAlignleft
            .Add "Annotazioni", "Annotazioni", dgchar, False, 3000, dgAlignleft
            .Add "Dismesso", "Dismesso", dgBoolean, True, 1500, dgAligncenter
            .Add "DataDismesso", "Data dismesso", dgDate, False, dgAlignleft
            .Add "IDArticolo", "IDArticolo", dgInteger, False, 500, 0, True, True, False
            .Add "CodiceArticolo", "Codice articolo", dgchar, False, 2000, dgAlignleft
            .Add "DescrizioneArticolo", "Articolo", dgchar, False, 2500, dgAlignleft
            .Add "IDRV_POUnitaDiMisuraPeriodo", "IDRV_POUnitaDiMisuraPeriodo", dgInteger, False, 500, 0, True, True, False
            .Add "UnitaDiMisuraPeriodo", "U.M. Periodo", dgchar, False, 2000, dgAlignleft
            .Add "DataInizioPeriodo", "Data inizio", dgDate, False, 2000, dgAligncenter
            .Add "OraInizioPeriodo", "Ora inizio", dgchar, False, 2000, dgAligncenter
            .Add "DataFinePeriodo", "Data fine", dgDate, False, 2000, dgAligncenter
            .Add "OraFinePeriodo", "Ora fine", dgchar, False, 2000, dgAligncenter
            Set cl = .Add("QuantitaPeriodo", "Q.tà periodo", dgDouble, False, 1800, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .Add "IDRateizzazione", "IDRateizzazione", dgInteger, False, 500, dgAlignRight
            .Add "Rateizzazione", "Rateizzazione", dgchar, False, 2000, dgAlignleft
            .Add "IDUnitaDiMisuraArticolo", "IDUnitaDiMisuraArticolo", dgInteger, False, 500, 0, True, True, False
            .Add "UnitaDiMisuraArticolo", "U.M. articolo", dgchar, False, 2000, dgAlignleft
            Set cl = .Add("QuantitaArticolo", "Q.tà articolo", dgDouble, False, 1800, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("ImportoUnitario", "Imp. uni.", dgDouble, False, 1800, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
             Set cl = .Add("Sconto1", "% sc. 1", dgDouble, False, 1800, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
             Set cl = .Add("Sconto2", "% sc. 2", dgDouble, False, 1800, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
             Set cl = .Add("Imponibile", "Imponibile riga", dgDouble, False, 1800, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."

             Set cl = .Add("ScontoAImporto", "Sc. importo", dgDouble, False, 1800, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."

             Set cl = .Add("TotaleRiga", "Totale", dgDouble, False, 1800, dgAlignRight, True, True, False)
                'cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."

        End With
    End If
    Me.GrigliaProd.EnableMove = True

'''''''''''''''''''''''''CONTROLLI STANDARD''''''''''''''''''''''''''''''''''''

    'Anagrafica cliente
    With Me.CDCliente
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Anagrafica"
        .DescriptionField = "Nome"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
        .Filter = "IDAzienda = " & m_App.IDFirm
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



    'Anagrafica del tecnico di riferimento
    With Me.CDAmministratore
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Anagrafica"
        .DescriptionField = "Nome"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIEAnagraficaPerTipo"
        .Filter = "IDAzienda = " & TheApp.IDFirm & " AND IDTipoAnagrafica=" & LINK_TIPO_ANA_AMM
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Cognome"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Nome"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Cognome"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Nome"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 29 'Anagrafica
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    'Anagrafica del tecnico di riferimento
    With Me.CDTecnico
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Anagrafica"
        .DescriptionField = "Nome"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIEAnagraficaPerTipo"
        .Filter = "IDAzienda = " & m_App.IDFirm & " AND IDTipoAnagrafica=" & LINK_TIPO_ANA_TEC_CONTRATTO
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Cognome"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Nome"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Cognome"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Nome"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 29 'Anagrafica
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    'Anagrafica dell'agente di riferimento
    With Me.CDAgente
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Anagrafica"
        .DescriptionField = "Nome"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepAgente"
        .Filter = "IDAzienda = " & m_App.IDFirm & " AND IDTipoAgente=1"
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Cognome"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Nome"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Cognome"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Nome"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 29 'Anagrafica
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    'Tipo contratto
    With Me.cboTipoContratto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoContratto"
        .DisplayField = "TipoContratto"
        .SQL = "SELECT * FROM RV_POTipoContratto ORDER BY TipoContratto"
        .Fill
    End With
    
    'Durata del contratto
    With Me.cboDurataContratto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PODurataContratto"
        .DisplayField = "DurataContratto"
        .SQL = "SELECT * FROM RV_PODurataContratto ORDER BY DurataContratto"
        .Fill
    End With
    
    'Durata del contratto prox rinnovo
    With Me.cboDurataContrattoProx
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PODurataContratto"
        .DisplayField = "DurataContratto"
        .SQL = "SELECT * FROM RV_PODurataContratto ORDER BY DurataContratto"
        .Fill
    End With
    
    'Tipo rinnovo del contratto
    With Me.cboTipoRinnovo
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoRinnovo"
        .DisplayField = "TipoRinnovo"
        .SQL = "SELECT * FROM RV_POTipoRinnovo ORDER BY TipoRinnovo"
        .Fill
    End With
    
    'Tipo rinnovo del contratto
    With Me.cboTipoRinnovoProx
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoRinnovo"
        .DisplayField = "TipoRinnovo"
        .SQL = "SELECT * FROM RV_POTipoRinnovo ORDER BY TipoRinnovo"
        .Fill
    End With
    
    'Pagamento rate
    With Me.cboPagamentoRate
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDPagamento"
        .DisplayField = "Pagamento"
        .SQL = "SELECT * FROM Pagamento ORDER BY Pagamento"
        .Fill
    End With

    With Me.cboPagamentoRataContratto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDPagamento"
        .DisplayField = "Pagamento"
        .SQL = "SELECT * FROM Pagamento ORDER BY Pagamento"
        .Fill
    End With

    'Tipo rateizzazione
    With Me.cboTipoRateizzazione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PORateizzazione"
        .DisplayField = "Rateizzazione"
        .SQL = "SELECT * FROM RV_PORateizzazione ORDER BY Rateizzazione"
        .Fill
    End With
    
    'Tipo rateizzazione
    With Me.cboTipoRateizzazioneProx
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PORateizzazione"
        .DisplayField = "Rateizzazione"
        .SQL = "SELECT * FROM RV_PORateizzazione ORDER BY Rateizzazione"
        .Fill
    End With
    
    'Tipo durata assistenza
    With Me.cboDurataAssistenza
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoDurataAssistenza"
        .DisplayField = "TipoDurataAssistenza"
        .SQL = "SELECT * FROM RV_POTipoDurataAssistenza ORDER BY TipoDurataAssistenza"
        .Fill
    End With

    

    With Me.CDServizio
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        
        If LINK_ARTICOLO_SERVIZIO = 0 Then
            .TableName = "RV_POIE_ConfigurazioneServizio"
            .Filter = "IDAzienda = " & m_App.IDFirm ' & " AND GestioneRicorrenze=1"
        Else
            .TableName = "IDArticolo"
            .Filter = "IDAzienda = " & m_App.IDFirm & " AND IDGruppoEquivalenzaArticolo=" & LINK_ARTICOLO_SERVIZIO
        End If
        
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 6 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    'Criterio di ricorrenza
    With Me.cboCriterioRicorrenza
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POCriterioRicorrenza"
        .DisplayField = "CriterioRicorrenza"
        .SQL = "SELECT * FROM RV_POCriterioRicorrenza ORDER BY IDRV_POCriterioRicorrenza"
        .Fill
    End With

    'Tipo data inizio ricorrenza
    With Me.cboTipoDataInizioRic
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoDataInizioRicorrenza"
        .DisplayField = "TipoDataInizioRicorrenza"
        .SQL = "SELECT * FROM RV_POTipoDataInizioRicorrenza ORDER BY IDRV_POTipoDataInizioRicorrenza"
        .Fill
    End With

    'Tipo data fine ricorrenza
    With Me.cboTipoDataFineRic
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoDataFineRicorrenza"
        .DisplayField = "TipoDataFineRicorrenza"
        .SQL = "SELECT * FROM RV_POTipoDataFineRicorrenza ORDER BY IDRV_POTipoDataFineRicorrenza"
        .Fill
    End With



    'IVA rate contratto
    With Me.cboIVARateContratto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Iva"
        .SQL = "SELECT * FROM Iva ORDER BY Iva"
        .Fill
    End With

    'IVA rate contratto
    With Me.cboIvaAdeg
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Iva"
        .SQL = "SELECT * FROM Iva ORDER BY Iva"
        .Fill
    End With

    'Tipo adeguamento per rinnovo
    With Me.cboTipoAdeguamento
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoAdeguamento"
        .DisplayField = "TipoAdeguamento"
        .SQL = "SELECT * FROM RV_POTipoAdeguamento"
        .Fill
    End With
    
    With Me.CDArticoloAdeg
       Set .Application = m_App
       Set .Database = m_App.Database
       .HwndContainer = Me.hwnd
       .CodeField = "CodiceArticolo"
       .DescriptionField = "Articolo"
       .KeyField = "IDArticolo"
       .TableName = "Articolo"
       .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm
       .MenuFunctions("EseguiGestione").Enabled = True
       .PropCodice.Caption = "Codice"
       .PropDescrizione.Caption = "Descrizione"
       .CodeCaption4Find = "Codice Articolo"
       .DescriptionCaption4Find = "Descrizione Articolo"
       .IDExecuteFunction = 6 'Articoli
       .CodeIsNumeric = False
    End With

    With CDArticoloProd
       Set .Application = m_App
       Set .Database = m_App.Database
       .HwndContainer = Me.hwnd
       .CodeField = "CodiceArticolo"
       .DescriptionField = "Articolo"
       .KeyField = "IDArticolo"
       .TableName = "Articolo"
       .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm
       .MenuFunctions("EseguiGestione").Enabled = True
       .PropCodice.Caption = "Codice"
       .PropDescrizione.Caption = "Descrizione"
       .CodeCaption4Find = "Codice Articolo"
       .DescriptionCaption4Find = "Descrizione Articolo"
       .IDExecuteFunction = 6 'Articoli
       .CodeIsNumeric = False
    End With

'    With Me.CDServizioProdotto
'       Set .Application = m_App
'       Set .Database = m_App.Database
'       .HwndContainer = Me.hwnd
'       .CodeField = "CodiceArticolo"
'       .DescriptionField = "Articolo"
'       .KeyField = "IDArticolo"
'       .TableName = "Articolo"
'       .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm
'       .MenuFunctions("EseguiGestione").Enabled = True
'       .PropCodice.Caption = "Codice"
'       .PropDescrizione.Caption = "Descrizione"
'       .CodeCaption4Find = "Codice Articolo"
'       .DescriptionCaption4Find = "Descrizione Articolo"
'       .IDExecuteFunction = 6 'Articoli
'       .CodeIsNumeric = False
'    End With
    
'    'IVA prodotti
'    With Me.cboIvaProd
'        Set .Database = m_App.Database.Connection
'        .AddFieldKey "IDIva"
'        .DisplayField = "Iva"
'        .SQL = "SELECT * FROM Iva ORDER BY Iva"
'        .Fill
'    End With

    'Istat per adeguamento
    With Me.cboIstatAdeg
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POIstat"
        .DisplayField = "Istat"
        .SQL = "SELECT * FROM RV_POIstat"
        .Fill
    End With
    
    'Impostazioni contratto
    With Me.cboTipoImpostazione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoImpostazioneContratto"
        .DisplayField = "Descrizione"
        .SQL = "SELECT * FROM RV_POTipoImpostazioneContratto"
        .Fill
    End With
    
    'Unità di misura del periodo dei prodotti
    With Me.cboUMPeriodoProd
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POUnitaDiMisuraPeriodo"
        .DisplayField = "Descrizione"
        .SQL = "SELECT * FROM RV_POUnitaDiMisuraPeriodo"
        .Fill
    End With
    
    'Unità di misura articolo
    With Me.cboUMArtProd
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDUnitaDiMisura"
        .DisplayField = "UnitaDiMisura"
        .SQL = "SELECT * FROM UnitaDiMisura"
        .Fill
    End With
    
    'Listino articolo del prodotto
    With Me.cboListinoProd
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDListino"
        .DisplayField = "Listino"
        .SQL = "SELECT * FROM Listino "
        .SQL = .SQL & "WHERE IDAzienda=" & TheApp.IDFirm
        .SQL = .SQL & " AND TipoListino=0"
        .Fill
    End With
    
    With Me.cboTipoPeriodo
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoPeriodo"
        .DisplayField = "Descrizione"
        .SQL = "SELECT * FROM RV_POTipoPeriodo"
        .Fill
    End With
    
    
    With Me.cboIvaProd
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Iva"
        .SQL = "SELECT * FROM Iva ORDER BY Iva"
        .Fill
    End With

    With Me.cboAnaOperatoreProd
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDAnagrafica"
        .DisplayField = "NomeCompleto"
        .SQL = "SELECT IDAnagrafica, NomeCompleto "
        .SQL = .SQL & "FROM RV_POIEAnagraficaPerTipo "
        .SQL = .SQL & "WHERE IDAzienda=" & TheApp.IDFirm
        .SQL = .SQL & " AND IDTipoAnagrafica=" & LINK_TIPO_ANA_TEC_FASE
        .SQL = .SQL & "ORDER BY NomeCompleto"
        .Fill
    End With
    
    'Tipo rateizzazione prodotti
    With Me.cboTipoRateizzazioneProd
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PORateizzazione"
        .DisplayField = "Rateizzazione"
        .SQL = "SELECT * FROM RV_PORateizzazione ORDER BY Rateizzazione"
        .Fill
    End With
    
    'Tipo rateizzazione adeguamento
    With Me.cboTipoRateizzazioneAdeg
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PORateizzazione"
        .DisplayField = "Rateizzazione"
        .SQL = "SELECT * FROM RV_PORateizzazione ORDER BY Rateizzazione"
        .Fill
    End With
    
    'Inizializza la DmtDocs
    If oDoc Is Nothing Then
        'Crea una istanza dell'oggetto cDocument della DmtDocs
        
        Set oDoc = New DmtDocs.cDocument
        Set oDoc.Connection = m_App.Database.Connection
        oDoc.SetTipoOggetto 2 'Documento di trasporto
        oDoc.IDFunzione = 105 'Documento di trasporto
        
        oDoc.TablesNames oDoc.IDTipoOggetto, sTabellaTestata, sTabellaDettaglio, sTabellaIVA, sTabellaScadenze
        
        oDoc.IDAzienda = TheApp.IDFirm
        oDoc.IDFiliale = TheApp.Branch
        oDoc.IDAttivitaAzienda = GetAttivitaAzienda(TheApp.IDFirm, TheApp.Branch)
        oDoc.IDTipoAnagrafica = 2 'Cliente
        
        oDoc.IDUtente = TheApp.IDUser
        
    End If
    
    'anno inizio ricorrenza
    With Me.cboTipoAnnoInizioRicorr
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoAnno"
        .DisplayField = "TipoAnno"
        .SQL = "SELECT * FROM RV_POTipoAnno"
        .Fill
    End With
    
    'Anno fine ricorrenza
    With Me.cboTipoAnnoFineRicorr
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoAnno"
        .DisplayField = "TipoAnno"
        .SQL = "SELECT * FROM RV_POTipoAnno"
        .Fill
    End With
    
    
    
    Set Me.LabelLink2.Application = TheApp
    Me.LabelLink2.WindowHandleClient = Me.hwnd
    Me.LabelLink2.PopMenuItems("Mnu_SearchObject").Enabled = False
    
    Set Me.LblLinkRil.Application = TheApp
    Me.LblLinkRil.WindowHandleClient = Me.hwnd
    Me.LblLinkRil.PopMenuItems("Mnu_SearchObject").Enabled = False
    
    Set Me.LabelLink1.Application = TheApp
    Me.LabelLink1.WindowHandleClient = Me.hwnd
    Me.LabelLink1.PopMenuItems("Mnu_SearchObject").Enabled = False
    
    Set Me.LabelLink3.Application = TheApp
    Me.LabelLink3.WindowHandleClient = Me.hwnd
    Me.LabelLink3.PopMenuItems("Mnu_SearchObject").Enabled = False
    
    Set Me.lblLinkFattContratto.Application = TheApp
    Me.lblLinkFattContratto.WindowHandleClient = Me.hwnd
    Me.lblLinkFattContratto.PopMenuItems("Mnu_SearchObject").Enabled = False
End Sub

'**+
'Nome: OnSave
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Save
'**/
Private Sub OnSave()
'On Error GoTo ERR_OnSave
Dim Field As DmtDocManLib.Field
Dim DocLink As DmtDocManLib.DocumentsLink
Dim NuovoContratto As Boolean
Dim NuovaRateizzazione As Boolean
Dim RiSviluppoRate As Boolean
Dim NuovoDocumento As Integer
Dim sSQL As String
Dim NumeroInserimenti As Long
Dim X As Long
Dim ErroreCoda As Boolean
Dim OLD_Cursor As Long
Dim m_condition As DmtDocManLib.Condition

    
    If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub

    NuovoContratto = False
    
    If m_Document(m_Document.PrimaryKey).Value < 0 Then
        NuovoContratto = True
        NuovoDocumento = 1
    Else
        NuovoDocumento = 0
    End If
    
    If Not PermissionToSave Then
        Exit Sub
    End If
    
    If (Me.cboTipoImpostazione.CurrentID = 1) Then
        If MODULO_ATTIVATO = 0 Then
            If Len(MODULO_DESCRIZIONE) > 0 Then
                MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
            Else
                MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
            End If
        Exit Sub
        End If
    End If
    If (Me.cboTipoImpostazione.CurrentID > 1) Then
        If MODULO_ATTIVATO_NOL = 0 Then
            If Len(MODULO_DESCRIZIONE_NOL) > 0 Then
                MsgBox "Il modulo " & MODULO_DESCRIZIONE_NOL & " non è stato abilitato", vbInformation, TheApp.FunctionName
            Else
                MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
            End If
        Exit Sub
        End If
    End If
    
    
    RiSviluppoRate = False
    
    If Me.cboTipoImpostazione.CurrentID = 1 Then
        If AGGIORNA_DA_ISTAT = 0 Then
            If NuovoContratto = False Then
                RiSviluppoRate = ControlloCambioTesta
            End If
        End If
    End If
    Screen.MousePointer = 11
    DoEvents
    
    
    SCRIVI_CODA fnNotNullN(m_Document(m_Document.PrimaryKey)), m_DocType.ID
    APERTURA_FORM_CODA = False
    NOME_GESTORE = App.EXEName
    
    'Passa alla collezione Fields dell'oggetto
    'Document i valori da salvare
    For Each Field In m_Document.Fields
        'Sul campo chiave primaria non si deve far nulla
        If Not Field.PrimaryKey Then
            If FieldPresent(Field.Name) Then
            
                'rif4 begin

                Select Case TypeName(m_FormFields(Field.Name).Control)
                    Case "TextBox"
                        Field.Value = m_FormFields(Field.Name).Control.Text
                    Case "DmtCodDesc"
                        Field.Value = m_FormFields(Field.Name).Control.KeyFieldID
                    Case "DMTCombo"
                        Field.Value = m_FormFields(Field.Name).Control.CurrentID
                    Case "Town"
                        If Field.Name = "IDComune" Then
                            Field.Value = m_FormFields(Field.Name).Control.CityID
                        ElseIf Field.Name = "Cap" Then
                            Field.Value = m_FormFields(Field.Name).Control.Zip
                        End If
                    Case "dmtDate"
                        If (m_FormFields(Field.Name).Control.Text = "") Or (IsNull(m_FormFields(Field.Name).Control.Value)) Then
                            Field.Value = Null
                        Else
                            Field.Value = m_FormFields(Field.Name).Control.Value
                        End If
                    Case "dmtNumber"
                        Field.Value = m_FormFields(Field.Name).Control.Value
                    Case "dmtCurrency"
                        Field.Value = m_FormFields(Field.Name).Control.Value
                    
                    Case "dmtTime"
                        Field.Value = m_FormFields(Field.Name).Control.Value
                    Case "DmtSearchACS"
                        Field.Value = m_FormFields(Field.Name).Control.IDAnagrafica
                    Case "CheckBox"
                        Field.Value = fnNormBoolean(m_FormFields(Field.Name).Control.Value)
                
                End Select
                
                'rif4 end
                
            Else
                If Field.Name = "IDAzienda" Then
                    Field.Value = m_App.IDFirm
                End If
                
                If Field.Name = "IDFiliale" Then
                    Field.Value = m_App.Branch
                End If
                If Field.Name = "IDCodiceConto" Then
                    Field.Value = Link_ContoPDC
                End If
                If Field.Name = "AnagraficaAgente" Then
                    Field.Value = Me.CDAgente.Code
                End If
                If Field.Name = "NomeAgente" Then
                    Field.Value = Me.CDAgente.Description
                End If
                If Field.Name = "AnagraficaCommesso" Then
                    Field.Value = Me.CDTecnico.Code
                End If
                If Field.Name = "NomeCommesso" Then
                    Field.Value = Me.CDTecnico.Description
                End If
                If Field.Name = "IDTipoAnagraficaCommesso" Then
                    Field.Value = LINK_TIPO_ANA_TEC_CONTRATTO
                End If
                If Field.Name = "AnagraficaAmministratore" Then
                    Field.Value = Me.CDAmministratore.Code
                End If
                If Field.Name = "NomeAmministratore" Then
                    Field.Value = Me.CDAmministratore.Description
                End If
                If Field.Name = "IDTipoAnagraficaAmministratore" Then
                    Field.Value = LINK_TIPO_ANA_AMM
                End If
                
                If Field.Name = "IDAnagraficaFatturazione" Then
                    Field.Value = IDClienteFatturazione
                End If
                                
                If Field.Name = "IDContrattoBancario" Then
                    Field.Value = IDContrattoBancario
                End If
                
                If Field.Name = "IDAccordoCommerciale" Then
                    Field.Value = IDAccordoCommerciale
                End If

                If Field.Name = "IDArticoloContratto" Then
                    Field.Value = IDArticoloContratto
                End If
                
                If Field.Name = "IDRaggruppamentoFatturato" Then
                    Field.Value = IDRaggrFattContratto
                End If
                
                If Field.Name = "IDRV_POTipoClassificazioneContratto" Then
                    Field.Value = IDClassContratto
                End If
                
                If Field.Name = "CodiceConto" Then
                    Field.Value = CodicePDCContratto
                End If
                
                If Field.Name = "DescrizioneConto" Then
                    Field.Value = DescrPDCContratto
                End If
                
                If Field.Name = "RiferimentoAzienda" Then
                    Field.Value = RapprLegaleAzienda
                End If
                
                If Field.Name = "RiferimentoCliente" Then
                    Field.Value = RapprLegaleCliente
                End If
                
                If Field.Name = "RuoloRifAzienda" Then
                    Field.Value = RuoloRapprAzienda
                End If
                
                If Field.Name = "RuoloRifCliente" Then
                    Field.Value = RuoloRapprCliente
                End If
                
                If Field.Name = "IDIstat" Then
                    Field.Value = IDIstatContratto
                End If
                
                If Field.Name = "Maggiorazione" Then
                    Field.Value = MaggIstatContratto
                End If
                
                
                If m_Document(m_Document.PrimaryKey).Value < 0 Then
                    If Field.Name = "IDUtentePerInserimento" Then
                        Field.Value = m_App.IDUser
                    End If
                    If Field.Name = "DataInserimento" Then
                        Field.Value = Date
                    End If
                Else
                    If Field.Name = "IDUtentePerModifica" Then
                        Field.Value = m_App.IDUser
                    End If
                    If Field.Name = "DataModifica" Then
                        Field.Value = Date
                    End If
                End If
                'Se il processo in corso è "Manutenzione da Shell"
                'la variabile m_LinkedField contiene il nome del
                'campo collegato alla applicazione chiamante
                'quindi il campo relativo deve essere valorizzato
                'con il valore ricevuto dalla applicazione chiamante
                'nCBC+
                If Field.Name = m_LinkedField Then
                    Field.Value = m_App.CallerFieldValue
                End If
            End If
        End If
    Next
    

    ''''''''''''''''''''''''''''''CONTROLLA LA CODA DEI SALVATAGGI'''''''''''''''''''''''''''''
    X = 0
    ErroreCoda = False
    Do
        X = GET_NUMERO_DOCUMENTO(IIf((m_Document(m_Document.PrimaryKey).Value <= 0), True, False), m_DocType.ID)
        If X = -1 Then
            X = 1
            ErroreCoda = True
        End If
    Loop Until X = 1
    
    If ErroreCoda = True Then
        X = -1
    End If
    
    If X = -1 Then
        Me.Enabled = True
        Me.SetFocus
        Me.Caption = Caption2Display
        Screen.MousePointer = 0
        ELIMINA_RIFERIMENTI_CODA m_DocType.ID
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Me.Enabled = True
    'Me.SetFocus
    'Me.Caption = Caption2Display
    
    OLD_Cursor = Cn.CursorLocation
    Cn.CursorLocation = adUseClient
    
    
    frmAttesa.Show
    Me.Enabled = False
    
    DoEvents
    
    Me.Caption = "SALVATAGGIO IN CORSO..................."
    DoEvents
    
    
    frmAttesa.lblInfo = Me.Caption
    DoEvents
    
    Cn.BeginTrans
    m_Document.SaveDocument
    Cn.CommitTrans

    If NuovoDocumento = 1 Then
        Me.txtIDContrattoPadre.Value = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        m_Document("IDRV_POContrattoPadre") = Me.txtIDContrattoPadre.Value
        Cn.BeginTrans
        m_Document.SaveDocument
        Cn.CommitTrans
    End If
    
    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
    m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
    
    'Refresh delle variabili di stato
    m_Changed = False
    m_Search = False
    m_Saved = True
    
    'Refresh dello stato della ToolBar standard in modalità variazione
    SetStatus4Modality Modify
    
    
    
    If NuovoContratto = True Then
        If Me.cboTipoImpostazione.CurrentID = 1 Then
            SviluppoRateContratto fnNotNullN(m_Document("IDRV_POContratto").Value)
        End If
    Else
        If AGGIORNA_DA_ISTAT = 0 Then
            If Me.cboTipoImpostazione.CurrentID = 1 Then
                If RiSviluppoRate = True Then
                    If MsgBox("Vuoi ricalcolare le rate?", vbQuestion + vbYesNo, "Sviluppo rate contratto") = vbYes Then
                        SviluppoRateContratto (m_Document("IDRV_POContratto").Value)
                    End If
                End If
            End If
        Else
            AGGIORNA_RATE_DA_ISTAT fnNotNullN(m_Document(m_Document.PrimaryKey).Value), Me.cboTipoRateizzazione.CurrentID, MaggIstatContratto
        End If
    End If
    
    If Me.cboTipoImpostazione.CurrentID = 3 Then
        If Me.chkGeneraAccontiSaldo.Value = vbChecked Then
            If fnNotNullN(m_Document("IDOggetto").Value) = 0 Then
            Dim Testo As String
                Testo = "ATTENZIONE!!!" & vbCrLf
                Testo = Testo & "Non è stato creato l'oggetto per poter creare un collegamento tra il contratto e il flusso documentale" & vbCrLf
                Testo = Testo & "Se il problema persiste contattare l'assistenza"
                MsgBox Testo, vbCritical, "Contratto in oggetto"
                Exit Sub
            End If
    
            Me.txtImportoTotAdeg.Value = Me.txtImportoAttuale.Value - GET_SALDO_CONTRATTO(fnNotNullN(m_Document("IDOggetto").Value))
    
    
            frmAttesa.lblInfo.Caption = "CREAZIONE SCADENZA..."
            DoEvents
            CREA_SCADENZA_CONTRATTO
        End If
    End If
    
    m_DocumentsLink.Refresh
    
    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display

    ELIMINA_RIFERIMENTI_CODA m_DocType.ID
    
    If Me.cboTipoImpostazione.CurrentID <= 2 Then
        Me.txtImportoTotAdeg.Value = GET_TOTALE_ADEGUAMENTI_DETTAGLIO(fnNotNullN(m_Document(m_Document.PrimaryKey).Value))
    End If
    If Me.cboTipoImpostazione.CurrentID = 3 Then
        'Me.txtImportoTotAdeg.Value = Me.txtImportoAttuale.Value - GET_TOTALI_ACCONTI_CONTRATTO(fnNotNullN(m_Document("IDOggetto").Value), m_DocType.ID)
        If Me.chkGeneraAccontiSaldo.Value = vbChecked Then
            Me.txtImportoTotAdeg.Value = Me.txtImportoAttuale.Value - GET_SALDO_CONTRATTO(fnNotNullN(m_Document("IDOggetto").Value))
        Else
            Me.txtImportoTotAdeg.Value = GET_TOTALE_ADEGUAMENTI_DETTAGLIO(fnNotNullN(m_Document(m_Document.PrimaryKey).Value))
        End If
    End If
    
    Screen.MousePointer = 0
    
    InitVariabiliContrSalvato fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    'Me.txtAltriDati.Text = GET_CARATTERISTICHE_RISORSA(Me.MSFlexGrid1) 'GET_DESCRIZIONE_ALTRI_DATI
    GET_CARATTERISTICHE_RISORSA Me.MSFlexGrid1
    
    CREA_RAPPR_LEG_AZIENDA TheApp.IDFirm, RapprLegaleAzienda
    CREA_RAPPR_LEG_CLIENTE TheApp.IDFirm, RapprLegaleCliente, Me.CDCliente.KeyFieldID
    CREA_RUOLO TheApp.IDFirm, RuoloRapprAzienda
    CREA_RUOLO TheApp.IDFirm, RuoloRapprCliente
    
    
    If m_Document(m_Document.PrimaryKey).Value > 0 Then
        Me.CDCliente.Enabled = False
        Me.cboTipoImpostazione.Enabled = False
    End If
    
    DoEvents
Exit Sub
ERR_OnSave:

    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    
    MsgBox Err.Description, vbCritical, "OnSave"

    Cn.RollbackTrans
    ELIMINA_RIFERIMENTI_CODA m_DocType.ID
    Cn.CursorLocation = OLD_Cursor
    
    Me.Caption = Caption2Display(False)


End Sub

'**+
'Nome: OnSaveDocumentsLink
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando OnSaveDocumentsLink
'**/
Private Sub OnSaveDocumentsLink(ByVal DocumentLink As DmtDocManLib.DocumentsLink)

End Sub

'**+
'Nome: OnDelete
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Delete
'**/
Private Sub OnDelete()
On Error GoTo ERR_OnDelete
    Dim sToRemove As String
    Dim DocLink As DmtDocManLib.DocumentsLink
    Dim Link_Contratto_local As Long
    Dim Link_Contratto_Prec As Long
    Dim Link_Contratto_Padre_local As Long
    Dim sSQL As String
    Dim rsListaProdContr As ADODB.Recordset
    Dim Link_Oggetto_Contratto As Long
    Dim IDOggettoScadenza As Long
    
    Dim Testo As String
    
    
    'Se si è in modalità tabellare potrebbe essere necessario sincronizzare
    'il documento con il record evidenziato nella browse
    If BrwMain.Visible = True Then
        If Not (m_Document.EOF = True And m_Document.BOF = True) Then
            m_Document.Move BrwMain.ListIndex - 1
        End If
    End If
    
    
    If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemDeleteAction) Then
        Exit Sub
    End If
    
    'Se in fase di inserimento di un nuovo
    'record non c'è niente da fare
    If m_Document.TableNew Then
        Exit Sub
    End If
    
    'If Me.txtNumeroRinnovo.Value > 1 Then
    '    MsgBox "Impossibile eliminare questo contratto poichè risultano dei rinnovi", vbCritical, TheApp.FunctionName
    '    Exit Sub
    'End If
    
    If Me.chkContrattoAttuale.Value = vbUnchecked Then
        MsgBox "Impossibile eliminare questo contratto poichè non risulta essere un contratto attuale", vbCritical, TheApp.FunctionName
        Exit Sub
    End If
        
    If GET_RATA_PAGATA(m_Document(m_Document.PrimaryKey).Value) = True Then
        MsgBox "Impossibile eliminare questo contratto poichè alcune rate risultano collegate ad un documento di vendita", vbCritical, TheApp.FunctionName
        Exit Sub
    End If

    If GET_RILEVAMENTO_PAGATO(m_Document(m_Document.PrimaryKey).Value) = True Then
        MsgBox "Impossibile eliminare questo contratto poichè alcuni rilevamenti di eccedenza dei contatori risultano collegate ad un documento di vendita", vbCritical, TheApp.FunctionName
        Exit Sub
    End If
    
    If GET_ESISTENZA_INTERVENTI_CONTRATTO(m_Document(m_Document.PrimaryKey).Value) = True Then
        MsgBox "Impossibile eliminare questo contratto poichè risultano collegamenti con uno più interventi gestiti manualmente dall'operatore", vbCritical, TheApp.FunctionName
        Exit Sub
    End If
    If Me.cboTipoImpostazione.CurrentID = 3 Then
        If GET_CONTROLLO_ESISTENZA_ACCONTI(fnNotNullN(m_Document("IDOggetto").Value)) = True Then
            MsgBox "Impossibile eliminare questo contratto poichè risultano collegamenti con uno più documenti di acconto/saldo", vbCritical, TheApp.FunctionName
            Exit Sub
        End If
    End If
    'Conferma della cancellazione
    gResource.CustomStrings.Clear
    sToRemove = fnNotNull(m_Document.Fields(CAMPO_PER_CAPTION).Value)
    gResource.CustomStrings.Add Chr(34) & sToRemove & Chr(34), 1
    
    If fnMsgQuestion(gResource.GetCustomizedMessage(MESS_QUERYREMOVE), m_App.FunctionName) = vbYes Then
        

        
        If Not (m_Document.EOF Or m_Document.BOF) Then
            'Cancella l'eventuale blocco sul record da cancellare.
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
        End If
        'rif16
        
        Link_Contratto_local = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        Link_Contratto_Padre_local = fnNotNullN(m_Document("IDRV_POContrattoPadre").Value)
        Link_Contratto_Prec = GET_LINK_CONTRATTO_PRECEDENTE(Link_Contratto_local, Link_Contratto_Padre_local)
        Link_Oggetto_Contratto = fnNotNullN(m_Document("IDOggetto").Value)
        
        GET_LISTA_PRODOTTI_CONTRATTI rsListaProdContr, fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        Screen.MousePointer = 11
        'Cancellazione
        m_Document.DeleteDocument
        
        EliminazioneContrattoNonStandard Link_Contratto_local
                       
        ELIMINA_FLUSSO_RATE_DA_DOCUMENTO Link_Contratto_local
        
        ELIMINA_INTERVENTO_CONTRATTO Link_Contratto_local
        
        ELIMINA_CONFIGURAZIONE_CONTATORI rsListaProdContr, Link_Contratto_local
        
        
        
        
        IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(Link_Oggetto_Contratto, m_DocType.ID, 0)
        
        If IDOggettoScadenza > 0 Then
            ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, Link_Oggetto_Contratto, m_DocType.ID
            ELIMINA_SCADENZA IDOggettoScadenza
        End If
        
        Screen.MousePointer = 0
        
        If Link_Contratto_Prec > 0 Then AGGIORNA_CONTRATTO_PRECEDENTE Link_Contratto_Prec
        
        If (m_Document.EOF = True And m_Document.BOF = True) Then
            'Se è stato cancellato l'ultimo record si va in modalità inserimento
            OnNewSearch
        Else
            'Refresh dello stato della ToolBar standard e dei menu
            If BrwMain.Visible Then
                'Va in modalità tabellare
                SetStatus4Modality Browse
            Else
                'Essendo in modalità variazione occorre controllare se il record su cui
                'ci si è posizionati è bloccato.
                'Se non lo è lo si blocca e si procede altrimenti si andrà in modalità tabellare.
                If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
                    'Il record è bloccato.
                    'Va in modalità tabellare
                    BrwMain.Visible = True
                    SetStatus4Modality Browse
                Else
                    'Il record non è bloccato.
                    
                    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
                    m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
                    
                    'Va in modalità variazione
                    SetStatus4Modality Modify
                End If
            
                 RefreshDescriptions4StatusBar
            End If
        End If
        
        'Refresh delle variabili di stato
        m_Changed = False
        m_Saved = True
        m_Search = False
        
    End If

Exit Sub
ERR_OnDelete:
    Screen.MousePointer = 0
    
    MsgBox Err.Description, vbCritical, "OnDelete"
    
    
End Sub

'**+
'Nome: OnDeleteDocumentsLink
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni relative ai documents link sul comando Delete
'**/
Private Sub OnDeleteDocumentsLink(ByVal DocumentLink As DmtDocManLib.DocumentsLink)
End Sub

'**+
'Nome: OnClear
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Clear
'**/
Private Sub OnClear()
'Se si è in modalità Filtro occorre ripulire i campi di immissione altrimenti,
'se si è in modalità Form, si cancella il contenuto di tutti i controlli
    
    
    If BrwMain.Visible And BrwMain.GuiMode = dgFilterDefinition Then
        '---Modalità Filtro---
        'Ripulisce i campi di immissione delle condizioni di ricerca.
        BrwMain.Conditions.ClearValues
    Else
        '---Modalità Form---
        'Ripulisce i campi del form
        ClearFormFields
        SetFocusTabIndex0
        
        'Se si era in modalità Nuovo viene disabilitato il pulsante Salva
        'e si ripristina la modalità stessa.
        If m_Document.TableNew Then
            ActivateBarButtons BTN_SAVE, False
            m_Changed = False
            m_Saved = True
        End If
    End If
End Sub

'**+
'Nome: OnExecuteSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando ExecuteSearch
'**/
Private Sub OnExecuteSearch()
    
    'Nota: utilizzo la chiamata al metodo ApplyFilter della dmtGrid piuttosto
    'che la chiamata diretta di ExecuteSearch perchè in questo modo la dmtGrid
    'può gestire internamente le conditions di ricerca.
    'Verrà generato l'evento BrwMain_OnApplyFilter()
    '
    'ExecuteSearch
    '
    BrwMain.ApplyFilter
    
        
End Sub

'**+
'Nome: OnMoveCurrentRecord
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando di riposizionamento del record corrente
'**/
Private Sub OnMoveCurrentRecord(ByVal Tipo As Integer, ByVal sToolName As String)
    Dim iResponse As Integer
    
    iResponse = ChooseAboutSaving
    If iResponse = vbYes Then
        OnSave
        'Se la registrazione non è andata a buon fine esce
        If Not m_Saved Then
            Exit Sub
        End If
    End If
    If iResponse <> vbCancel Then
       Select Case Tipo
           Case SRCNEXT
               SearchNext
           Case SRCPREVIOUS
               SearchPrevious
       End Select
       m_Changed = False
       ActivateBarButtons BTN_SAVE, False
    End If
End Sub

'**+
'Nome: OnRepositionDocumentsLink
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando di riposizionamento del record corrente
'per i DocumentsLink
'**/
Private Sub OnRepositionDocumentsLink(ByVal DocumentsLink As DmtDocManLib.DocumentsLink)

End Sub

'**+
'Nome: OnChangeView
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando ChangeView
'**/
Private Sub OnChangeView(ByVal sToolName As String)
    Dim iResponse   As Integer
    
    If Not BrwMain.Visible And m_Changed Then
        iResponse = ChooseAboutSaving
        
        If iResponse = vbYes Then
            OnSave
            'Se la registrazione non è andata a buon fine esce
            If Not m_Saved Then
                Exit Sub
            End If
        End If
        
        If iResponse <> vbCancel Then
            'cbc 20/04/1999
            'se si è scelto NO ripulisce i campi e va in modalità tabellare annullando
            'le ultime modifiche
            RefreshFormFields
            ChangeView sToolName
            m_Changed = False
        End If
    Else
        ChangeView sToolName
    End If
    
End Sub

'**+
'Nome: OnToolBarOptions
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando ToolBar
'
'
'**/
Private Sub OnToolBarOptions()
    Dim dlgToolBars As frmToolBars
    Dim bVisible As Boolean
    
    On Error Resume Next
    
    Set dlgToolBars = New frmToolBars
    'Imposta un riferimento al form chiamante
    Set dlgToolBars.FormClient = Me
    dlgToolBars.Show vbModal, Me
    Set dlgToolBars = Nothing
    
    'All'uscita dal form di dialogo la visibilità della toolbar dei filtri dipende dalla
    'visibilità del Riquadro attività e dall'impostazione fatta nel dialogo.
    bVisible = GetSetting(REGISTRY_KEY, TheApp.Name & "Settings", "Riquadro attività", True)
    BarMenu.RecalcLayout
End Sub

'**+
'Nome: OnOptions
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Option
'**/
Private Sub OnOptions()
    Dim dlgOption As frmOption
    
    Set dlgOption = New frmOption
    Set dlgOption.FormClient = Me
    dlgOption.Show vbModal, Me
    
    
    Set dlgOption = Nothing
    
    'Impedisce il 'blocco' della Toolbar alla chiusura di un form di dialogo.
    
End Sub

'**+
'Nome: OnInfo
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Info
'**/
Private Sub OnInfo()
    Dim dlgInfo As frmInformazioni
    
    Set dlgInfo = New frmInformazioni
    dlgInfo.Show vbModal, Me
    
    'Impedisce il 'blocco' della Toolbar alla chiusura di un form di dialogo.
    
End Sub

'**+
'Nome: OnPrint
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Print
'**/
Private Sub OnPrint(ByVal ToolName As String)
    Dim lFlags As Long
    Dim OLDCursor As Integer
    Dim sStr As String
    Dim Field As DmtDocManLib.Field
    
    
    OLDCursor = Screen.MousePointer
    
    'Se il filtro attivo è "Nessun record" è possibile eseguire una stampa/esportazione soltanto se
    'si è in modalità form. In tal caso, infatti, verrà passato al Crystals Reports un filtro
    'creato ad hoc sull'ID del record attuale.
    If m_ActiveFilter.NothingSelected And BrwMain.Visible Then
        sStr = "Impossibile effettuare l'operazione richiesta." & vbCrLf
        sStr = sStr & "Prima di procedere occorre eseguire un filtro."
        sbMsgInfo sStr, m_App.FunctionName
        Screen.MousePointer = OLDCursor
        Exit Sub
    End If
    
    'Se non esiste un report attivo occorre annullare l'operazione.
    If Len(oReportsActivity.SelectedReportName) > 0 Then
        Set m_Report = m_DocType.Reports.Item(oReportsActivity.SelectedReportName)
    End If
    If m_Report Is Nothing Then
        sbMsgError "Impossibile eseguire - Nessun report predefinito.", m_App.FunctionName
        GoTo OnPrint_Exit
    End If
    m_iNumeroCopieDefault = m_Report.Copies
    m_OrientamentoDefault = m_Report.Orientation
    
    
    
    
    'Se è attivo il pulsante Salva deve essere visualizzato un messaggio di avviso
    'con i pulsanti OK e annulla (occorre salvare PRIMA della stampa)
    'Se è attivo il pulsante Salva deve essere visualizzato un messaggio di avviso
    'con i pulsanti OK e annulla (occorre salvare PRIMA della stampa)
    If m_Changed Then
        Select Case ChooseAboutSavingOkCancel
            Case vbOK
                OnSave
                'Se la registrazione non è andata a buon fine esce
                If Not m_Saved Then
                    GoTo OnPrint_Exit
                End If
                
            Case vbCancel
                GoTo OnPrint_Exit
        End Select
    End If
    
    If Not BrwMain.Visible Then
        'Modalità Form - deve stampare solo il record corrente
        
        'Ripulisce la collezione Fields dell'oggetto DocType.
        For Each Field In m_DocType.Fields
            Field.Value = Empty
        Next
        
        'Viene inserita la condizione di ricerca basata sull'ID del record corrente.
        m_DocType.Fields("ID" & m_App.TableName).Value = m_Document.Fields("ID" & m_App.TableName).Value
        
        'Viene creato un filtro temporaneo per il Crystals Reports.
        m_DocType.RemoveFilter "Form"
        Set m_Report.Filter = m_DocType.AddFilterWithConditions("Form")
    Else
        'Modalità vista tabellare
        
        'Viene passato il filtro corrente al Crystals Reports.
        Set m_Report.Filter = m_ActiveFilter
    End If
            
    
    Select Case ToolName
    
        Case "PrePrint", "Mnu_PrePrint"
            On Error GoTo ErrorHandler
            
            Screen.MousePointer = vbHourglass
            
            m_TabMode = BrwMain.Visible
            PicForm.Visible = False
            BrwMain.Visible = False
            ActivityBox.Visible = False
            
            SetStatus4Modality Preview, OpenPrw
            Refresh
            
            m_PreviewWindowHandle = m_Document.Preview(m_Report, "", hwnd, CInt(BarMenu.ClientAreaLeft / Screen.TwipsPerPixelX), CInt(BarMenu.ClientAreaTop / Screen.TwipsPerPixelY), CInt(BarMenu.ClientAreaWidth / Screen.TwipsPerPixelX), CInt(BarMenu.ClientAreaHeight / Screen.TwipsPerPixelY), False)
            lFlags = SWP_NOSIZE Or SWP_NOREPOSITION Or SWP_NOMOVE
            SetWindowPos m_PreviewWindowHandle, HWND_TOP, 0, 0, 0, 0, lFlags
            
        Case "Print", "Mnu_Print"
            PrintDocument ToolName
            
        Case "ExportWord", "Mnu_ExportWord"
            ExportDocument ecWord
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), Word, TheApp.Name

        Case "ExportExcel", "Mnu_ExportExcel"
            ExportDocument ecExcel
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), Excel, TheApp.Name
            
        Case "ExportHtml", "Mnu_ExportHtml"
            ExportDocument ecHtml
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), HTML, TheApp.Name
        
        Case "ExportPDF", "Mnu_ExportPDF"
            ExportDocument ecPdf
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), PDF, TheApp.Name
        
        Case "MailWord"
            SendDocument ecWord
            
        Case "MailExcel"
            SendDocument ecExcel
            
        Case "MailHtml"
            SendDocument ecHtml
        
        Case "MailPDF"
            SendDocument ecPdf
    End Select
    
   
OnPrint_Exit:
    Set Field = Nothing
    Screen.MousePointer = OLDCursor
    Exit Sub
    
ErrorHandler:
    Const ERROR_PRINTING_ABORTED = 3
    Const ERROR_PRINTING_CANCELLED = 4
    Select Case Err.Number
        Case 20507
            'Errore "Invalid file Name" generato quando non è possibile trovare il file .rpt
            sbMsgInfo "File di report non trovato", m_App.FunctionName
        Case ERROR_PRINTING_ABORTED, ERROR_PRINTING_CANCELLED
            'non deve far niente, è stato già segnalato da CrystalReport
        Case Else
            If Len(Trim(Err.Description)) > 0 Then
                sbMsgInfo Err.Description, m_App.FunctionName
            End If
    End Select

    'Si è verificato un errore durante la procedura di anteprima.
    Screen.MousePointer = OLDCursor
    
    'Ripristina la situazione del form
    m_PreviewWindowHandle = 0
    PicForm.Visible = True

    BrwMain.Visible = m_TabMode
    ActivityBox.Visible = BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked
    FormRecalcLayout
    SetStatus4Modality Preview, ClosePrw
        
    Set Field = Nothing
End Sub

'**+
'Nome: OnNewSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando NewSearch
'**/
Private Sub OnNewSearch()
    Dim iResponse As Integer

    m_FilterSelected = False
    
    If Not m_Changed Then
        NewSearch
    Else
        'cbc 20/04/1999
        'deve mostrare il messaggio con Si, No, Annulla
        iResponse = ChooseAboutSaving
        If iResponse = vbYes Then
            OnSave
            'Se la registrazione non è andata a buon fine esce
            If Not m_Saved Then
                Exit Sub
            End If
        End If
        If iResponse <> vbCancel Then
            'se si è scelto NO ripristina i dati precedenti annullando le ultime modifiche
            'e predispone la modalità ricerca.
            RefreshFormFields
            NewSearch
            m_Changed = False
        End If
    
    End If
    
    
    
End Sub

'**+
'Nome: OnNew
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando New
'**/
Private Sub OnNew(ByVal sToolName As String)
    
    Select Case DoNewDocument
        Case vbYes
            'Si è risposto affermativamente alla
            'richiesta di Update delle modifiche apportate
            OnSave
            'Se la registrazione non è andata a buon fine esce
            If Not m_Saved Then
                Exit Sub
            End If
            
            NewRecord
            
            If (NON_MODIFICA_CONTRATTO = 1) Then
                NewSearch
            End If
        Case vbCancel
            'Si è risposto Annulla alla richiesta di Update
            Exit Sub
            
        Case Else
        
            NewRecord
            'Si è premuto il tasto <No> alla richiesta di Update
            If (NON_MODIFICA_CONTRATTO = 1) Then
                NewSearch
            End If
    End Select
End Sub

'**+
'Nome: OnNewDocumentsLink
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando New per i documentslink
'**/
Private Sub OnNewDocumentsLink(ByVal DocumentsLink As DmtDocManLib.DocumentsLink)

End Sub

'**+
'Nome: OnSummary
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Summary
'**/
Private Sub OnSummary()
    Dim lRes As Long
    
    lRes = WinHelp(hwnd, App.HelpFile, HELP_FINDER, 0)
End Sub

'**+
'Nome: OnFastHelp
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando FastHelp
'**/
Private Sub OnFastHelp()
    frmMain.WhatsThisMode
End Sub

'**+
'Nome: OnHelpOnLine
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando HelpOnLine
'**/
Private Sub OnHelpOnLine()
    Dim lRes As Long
    
    If Not ActiveControl Is Nothing Then
        If ActiveControl.HelpContextID <> 0 Then
            lRes = WinHelp(hwnd, App.HelpFile, HELP_CONTEXT, ActiveControl.HelpContextID)
        Else
            ExecuteMenuCommand "Mnu_Arg"
        End If
    End If
End Sub

'**+
'Nome: OnArg
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Arg
'**/
Private Sub OnArg()
    Dim lRes As Long
    
    If m_App.ContextHelpID <> 0 Then
        lRes = WinHelp(hwnd, App.HelpFile, HELP_CONTEXT, m_App.ContextHelpID)
    Else
        ExecuteMenuCommand "Mnu_Summary"
    End If
End Sub

'**+
'Nome: OnViewAssistant
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando ViewAssistant
'**/
Private Sub OnViewAssistant()
    BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAssistant").Checked = Not BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAssistant").Checked
    If BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAssistant").Checked Then
    Else
    End If
End Sub

'**/
'Autore                 : Diamante S.p.a
'
'Nome                   : OnFolders
'
'Parametri:
'
'
'Valori di ritorno:
'
'Funzionalità:
'Permette la visualizzazione o meno del DocTypeExplorer e della relativa toolbar.
'**/
Private Sub OnFolders()
    ActivityBox.Visible = Not ActivityBox.Visible
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked = ActivityBox.Visible
    FormRecalcLayout
    
    BarMenu.RecalcLayout
End Sub


'**+
'Nome: OnRunApplication
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando RunApplication
'**/
Private Sub OnRunApplication(ByVal sToolName As String)


End Sub




Private Sub BrwMain_ConditionEdit(ByVal Name As String, Value As Variant)
    Dim oSearch As dmtFind.Find
    Dim sSQL As String
    Dim oRes As DmtOleDbLib.adoResultset

    'Crea un'istanza dell'oggetto Find
    Set oSearch = New dmtFind.Find
    
    'Assegna la connessione aperta
    oSearch.Database = TheApp.Database.Connection



If Name = "Cliente" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Cliente"

    oSearch.AddDisplayField "Anagrafica", "Anagrafica", 1 'STRINGTYPE
    oSearch.AddDisplayField "Nome", "Nome", 1   'STRINGTYPE
 
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome "
    sSQL = sSQL & "FROM Anagrafica INNER JOIN "
    sSQL = sSQL & "Cliente ON Anagrafica.IDAnagrafica = Cliente.IDAnagrafica "
    sSQL = sSQL & "WHERE Cliente.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " ORDER BY Anagrafica"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("Anagrafica")
                
    End If
End If

If Name = "Agente" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Agente"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Anagrafica" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Anagrafica", "Anagrafica", 1 'STRINGTYPE
    oSearch.AddDisplayField "Nome", "Nome", 1   'STRINGTYPE
 
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome "
    sSQL = sSQL & "FROM Anagrafica INNER JOIN "
    sSQL = sSQL & "Agente ON Anagrafica.IDAnagrafica = Agente.IDAnagrafica "
    sSQL = sSQL & "WHERE Agente.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND Agente.IDTipoAgente=1"
    sSQL = sSQL & " ORDER BY Anagrafica"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("Anagrafica")
                
    End If
End If
If Name = "Tecnico" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Tecnico"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Anagrafica" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Anagrafica", "Anagrafica", 1 'STRINGTYPE
    oSearch.AddDisplayField "Nome", "Nome", 1   'STRINGTYPE
 
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT IDAnagrafica, Anagrafica, Nome "
    sSQL = sSQL & "FROM RV_POIEAnagraficaPerTipo "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDTipoAnagrafica=" & LINK_TIPO_ANA_TEC_INT
    sSQL = sSQL & " ORDER BY Anagrafica"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("Anagrafica")
                
    End If
End If
If Name = "Amministratore" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Amministratore"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Anagrafica" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Anagrafica", "Anagrafica", 1 'STRINGTYPE
    oSearch.AddDisplayField "Nome", "Nome", 1   'STRINGTYPE
 
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT IDAnagrafica, Anagrafica, Nome "
    sSQL = sSQL & "FROM RV_POIEAnagraficaPerTipo "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDTipoAnagrafica=" & LINK_TIPO_ANA_AMM
    sSQL = sSQL & " ORDER BY Anagrafica"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("Anagrafica")
                
    End If
End If

If Name = "Filiale" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Filiale"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Anagrafica" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Filiale", "SitoPerAnagrafica", 1 'STRINGTYPE
    oSearch.AddDisplayField "Cliente", "Anagrafica", 1   'STRINGTYPE
 
    'Assegna la condizione iniziale
    oSearch.Start = Value
    
    
    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT SitoPerAnagrafica.IDSitoPerAnagrafica, SitoPerAnagrafica.SitoPerAnagrafica, Anagrafica.Anagrafica "
    sSQL = sSQL & "FROM SitoPerAnagrafica INNER JOIN "
    sSQL = sSQL & "Cliente ON SitoPerAnagrafica.IDAnagrafica = Cliente.IDAnagrafica INNER JOIN "
    sSQL = sSQL & "Anagrafica ON Cliente.IDAnagrafica = Anagrafica.IDAnagrafica "
    sSQL = sSQL & "WHERE Cliente.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " ORDER BY Anagrafica"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("SitoPerAnagrafica")
                
    End If
End If

If Name = "Codice articolo contratto" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Codice articolo contratto"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Anagrafica" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Codice", "CodiceArticolo", 1 'STRINGTYPE
    oSearch.AddDisplayField "Descrizione", "Articolo", 1   'STRINGTYPE
 
    'Assegna la condizione iniziale
    oSearch.Start = Value
    
    
    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT IDArticolo, CodiceArticolo, Articolo "
    sSQL = sSQL & "FROM Articolo "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " ORDER BY CodiceArticolo"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("CodiceArticolo")
                
    End If
End If

If Name = "Articolo contratto" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Articolo contratto"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Anagrafica" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    
    oSearch.AddDisplayField "Descrizione", "Articolo", 1   'STRINGTYPE
    oSearch.AddDisplayField "Codice", "CodiceArticolo", 1 'STRINGTYPE
    'Assegna la condizione iniziale
    oSearch.Start = Value
    
    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT IDArticolo, CodiceArticolo, Articolo "
    sSQL = sSQL & "FROM Articolo "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " ORDER BY Articolo"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("Articolo")
                
    End If
End If


End Sub




Private Sub cboDurataAssistenza_Click()
On Error GoTo ERR_cboDurataAssistenza_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If Me.cboDurataAssistenza.CurrentID = 0 Then Exit Sub
sSQL = "SELECT Mesi, Giorni "
sSQL = sSQL & "FROM RV_POTipoDurataAssistenza "
sSQL = sSQL & "WHERE IDRV_POTipoDurataAssistenza=" & Me.cboDurataAssistenza.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Mesi_Durata_Assistenza = fnNotNullN(rs!Mesi)
    Giorni_Durata_Assistenza = fnNotNullN(rs!Giorni)
End If

rs.CloseResultset
Set rs = Nothing

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    Me.txtDataFineAssistenza.Value = DateAdd("m", Mesi_Durata_Assistenza, Me.txtDataDecorrenza.Text) - 1
    Me.txtDataFineAssistenza.Value = DateAdd("d", Giorni_Durata_Assistenza, Me.txtDataFineAssistenza.Text)
Else
    If fnNotNullN(m_Document("IDRV_POTipoDurataAssistenza").Value) <> Me.cboDurataAssistenza.CurrentID Then
        Me.txtDataFineAssistenza.Value = DateAdd("m", Mesi_Durata_Assistenza, Me.txtDataDecorrenza.Text) - 1
        Me.txtDataFineAssistenza.Value = DateAdd("d", Giorni_Durata_Assistenza, Me.txtDataFineAssistenza.Text)
    End If
End If

If Not (BrwMain.Visible) Then Change

Exit Sub
ERR_cboDurataAssistenza_Click:
    MsgBox Err.Description, vbCritical, "cboDurataAssistenza_Click"
End Sub

Private Sub cboDurataContratto_Click()
On Error GoTo ERR_cboDurataContratto_Click
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    

    sSQL = "SELECT Mesi, Giorni "
    sSQL = sSQL & "FROM RV_PODurataContratto "
    sSQL = sSQL & "WHERE IDRV_PODurataContratto=" & Me.cboDurataContratto.CurrentID
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        Mesi_Durata_Contratto = fnNotNullN(rs!Mesi)
        Giorni_Durata_Contratto = fnNotNullN(rs!Giorni)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    If fnNotNullN(m_Document("IDRV_POContratto")) <= 0 Then
        If fnNotNullN(m_Document("IDDurataContratto").Value) = 0 Then
            If Me.txtDataDecorrenza.Text <> "" Then
                Me.txtDataScadenza.Text = DateAdd("m", Mesi_Durata_Contratto, Me.txtDataDecorrenza.Text) - 1
                Me.txtDataScadenza.Text = DateAdd("d", Giorni_Durata_Contratto, Me.txtDataScadenza.Text)
            End If
        Else
            If fnNotNullN(m_Document("IDDurataContratto").Value) <> Me.cboDurataContratto.CurrentID Then
                Me.txtDataScadenza.Text = DateAdd("m", Mesi_Durata_Contratto, Me.txtDataDecorrenza.Text) - 1
                Me.txtDataScadenza.Text = DateAdd("d", Giorni_Durata_Contratto, Me.txtDataScadenza.Text)
                   
                Me.cboTipoRinnovo.WriteOn 0
                Me.cboTipoRateizzazione.WriteOn 0

            End If
        End If
        
        Me.cboDurataContrattoProx.WriteOn Me.cboDurataContratto.CurrentID
    Else
        If m_Document("IDDurataContratto") <> Me.cboDurataContratto.CurrentID Then
            Me.txtDataScadenza.Text = DateAdd("m", Mesi_Durata_Contratto, Me.txtDataDecorrenza.Text) - 1
            Me.txtDataScadenza.Text = DateAdd("d", Giorni_Durata_Contratto, Me.txtDataScadenza.Text)
            
            Me.cboTipoRinnovo.WriteOn 0
            Me.cboTipoRateizzazione.WriteOn 0
        Else
            If (Me.txtDataScadenza.Text = "") And (Me.cboDurataContratto.CurrentID > 0) Then
                Me.txtDataScadenza.Text = DateAdd("m", Mesi_Durata_Contratto, Me.txtDataDecorrenza.Text) - 1
                Me.txtDataScadenza.Text = DateAdd("d", Giorni_Durata_Contratto, Me.txtDataScadenza.Text)
            End If
        End If
    
    End If
    

    
    If Not (BrwMain.Visible) Then Change

Exit Sub
ERR_cboDurataContratto_Click:
    MsgBox Err.Description, vbCritical, "cboDurataContratto_Click"
End Sub
Private Sub cboDurataContrattoProx_Click()
On Error GoTo ERR_cboDurataContrattoProx_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Mesi, Giorni "
sSQL = sSQL & "FROM RV_PODurataContratto "
sSQL = sSQL & "WHERE IDRV_PODurataContratto=" & Me.cboDurataContrattoProx.CurrentID

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF = False Then
    Mesi_Durata_Contratto_prox = fnNotNullN(rs!Mesi)
    Giorni_Durata_Contratto_prox = fnNotNullN(rs!Giorni)
End If

rs.CloseResultset
Set rs = Nothing

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    If (Me.txtDataScadenza.Value) > 0 Then
        Me.txtDataScadSecContr.Text = DateAdd("m", Mesi_Durata_Contratto_prox + Mesi_Durata_Contratto, Me.txtDataDecorrenza.Text) - 1
        Me.txtDataScadSecContr.Text = DateAdd("d", Giorni_Durata_Contratto_prox + Giorni_Durata_Contratto, Me.txtDataScadSecContr.Text)
    End If
Else
    If (fnNotNullN(m_Document("IDDurataContrattoProssimoRinnovo").Value) <> Me.cboDurataContrattoProx.CurrentID) Then
        If (Me.txtDataScadenza.Value) > 0 Then
            Me.txtDataScadSecContr.Text = DateAdd("m", Mesi_Durata_Contratto_prox + Mesi_Durata_Contratto, Me.txtDataDecorrenza.Text) - 1
            Me.txtDataScadSecContr.Text = DateAdd("d", Giorni_Durata_Contratto_prox + Giorni_Durata_Contratto, Me.txtDataScadSecContr.Text)
        End If
    End If
End If

If Not (BrwMain.Visible) Then Change

Exit Sub
ERR_cboDurataContrattoProx_Click:
    MsgBox Err.Description, vbCritical, "cboDurataContrattoProx_Click"
End Sub

Private Sub cboIvaProd_Click()
On Error Resume Next

If (m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0 Then

    Me.txtAliquotaIvaProd.Value = GET_ALIQUOTA_IVA(Me.cboIvaProd.CurrentID)

    GET_TOTALI_RIGA_DETTAGLIO
Else
    If (fnNotNullN(m_DocumentsLink3("IDIva").Value)) <> Me.cboIvaProd.CurrentID Then
        Me.txtAliquotaIvaProd.Value = GET_ALIQUOTA_IVA(Me.cboIvaProd.CurrentID)
    
        GET_TOTALI_RIGA_DETTAGLIO
    End If
End If

End Sub
Private Sub cboListinoProd_Click()
Dim Testo As String
On Error Resume Next

If (m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0 Then
    GET_PREZZO_ARTICOLO Me.CDArticoloProd.KeyFieldID, Me.cboListinoProd.CurrentID, LINK_LISTINO_AZIENDA, Me.CDCliente.KeyFieldID
    
    GET_TOTALI_RIGA_DETTAGLIO
Else
    If (m_DocumentsLink3("IDListino").Value) <> Me.cboListinoProd.CurrentID Then
        Testo = "Vuoi ricalcolare l'importo dell'articolo?"
        
        If MsgBox(Testo, vbQuestion + vbYesNo, "Prezzo da listino") = vbNo Then Exit Sub
        
        GET_PREZZO_ARTICOLO Me.CDArticoloProd.KeyFieldID, Me.cboListinoProd.CurrentID, LINK_LISTINO_AZIENDA, Me.CDCliente.KeyFieldID
        
        GET_TOTALI_RIGA_DETTAGLIO
    End If
End If
End Sub

Private Sub cboPagamentoRate_Click()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub cboTipoContratto_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT DescrizioneAggiuntiva "
sSQL = sSQL & "FROM RV_POTipoContratto "
sSQL = sSQL & "WHERE IDRV_POTipoContratto=" & Me.cboTipoContratto.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Descrizione_Tipo_Contratto = ""
Else
    Descrizione_Tipo_Contratto = fnNotNull(rs!DescrizioneAggiuntiva)
End If

rs.CloseResultset
Set rs = Nothing

If Not (BrwMain.Visible) Then Change

End Sub


Private Sub cboTipoImpostazione_Click()
On Error GoTo ERR_cboTipoImpostazione_Click

    Me.Label1(35).Caption = "Tot. contr. + Adeg."
    
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    SSTab1.TabEnabled(4) = True
    'SSTab1.TabEnabled(5) = True
    'SSTab1.TabEnabled(6) = True
    
    SSTab1.Tab = 0
    
    Select Case cboTipoImpostazione.CurrentID
        Case 1
            txtDataStipula.Enabled = True
            txtImportoStipula.Enabled = True
            txtDataDecorrenza.Enabled = True
            txtNumeroRinnovo.Enabled = True
            txtImportoAttuale.Enabled = True
            cboTipoRateizzazione.Enabled = True
            cboDurataContratto.Enabled = True
            txtDataScadenza.Enabled = True
            cboTipoRinnovo.Enabled = True
            txtDataScadenzaPerRinnovo.Enabled = True
            cboDurataAssistenza.Enabled = True
            txtDataFineAssistenza.Enabled = True
            cboDurataContrattoProx.Enabled = True
            cboTipoRinnovoProx.Enabled = True
            txtDataScadSecContr.Enabled = True
            txtDataScadPerRinnovoProx.Enabled = True
            cboTipoRateizzazioneProx.Enabled = True
            'cmdGeneraRate.Enabled = True
            'cmdAcconti.Enabled = False
            'cmdSaldo.Enabled = False
            
            chkRinnovoAutomatico.Enabled = True
            chkAdeguamentoIstat.Enabled = True
            chkGeneraRateProd.Enabled = False
            chkGeneraAccontiSaldo.Enabled = False
            
            cmdgenIntDaProd.Enabled = False
            cmdGeneraScadenzeProd.Enabled = False
            
            'SSTab1.TabEnabled(5) = False
            
        Case 2
            txtDataStipula.Enabled = True
            txtImportoStipula.Enabled = True
            txtDataDecorrenza.Enabled = True
            txtNumeroRinnovo.Enabled = True
            txtImportoAttuale.Enabled = False
            cboTipoRateizzazione.Enabled = True
            cboDurataContratto.Enabled = True
            txtDataScadenza.Enabled = True
            cboTipoRinnovo.Enabled = True
            txtDataScadenzaPerRinnovo.Enabled = True
            cboDurataAssistenza.Enabled = True
            txtDataFineAssistenza.Enabled = True

            cboDurataContrattoProx.Enabled = True
            cboTipoRinnovoProx.Enabled = True
            txtDataScadSecContr.Enabled = True
            txtDataScadPerRinnovoProx.Enabled = True
            cboTipoRateizzazioneProx.Enabled = True

            'cmdGeneraRate.Enabled = True
            'cmdAcconti.Enabled = False
            'cmdSaldo.Enabled = False

            chkRinnovoAutomatico.Enabled = True
            chkAdeguamentoIstat.Enabled = True
            chkGeneraRateProd.Enabled = True
            chkGeneraAccontiSaldo.Enabled = False
            
            cmdgenIntDaProd.Enabled = False
            cmdGeneraScadenzeProd.Enabled = False
            
            SSTab1.TabEnabled(3) = False
            'SSTab1.TabEnabled(5) = False
            
            SSTab1.Tab = 4
            
            If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
                Me.txtImportoAttuale.Value = 0
            End If
            
            
        Case 3
            txtDataStipula.Enabled = True
            txtImportoStipula.Enabled = False
            txtDataDecorrenza.Enabled = True
            txtNumeroRinnovo.Enabled = False
            txtImportoAttuale.Enabled = False
            cboTipoRateizzazione.Enabled = False
            cboDurataContratto.Enabled = False
            txtDataScadenza.Enabled = False
            cboTipoRinnovo.Enabled = False
            txtDataScadenzaPerRinnovo.Enabled = False
            cboDurataAssistenza.Enabled = False
            txtDataFineAssistenza.Enabled = False
            cboDurataContrattoProx.Enabled = False
            cboTipoRinnovoProx.Enabled = False
            txtDataScadSecContr.Enabled = False
            txtDataScadPerRinnovoProx.Enabled = False
            cboTipoRateizzazioneProx.Enabled = False
            
            'Me.Label1(35).Caption = "Saldo"
            'cmdGeneraRate.Enabled = False
            'cmdAcconti.Enabled = False
            'cmdSaldo.Enabled = False

            chkRinnovoAutomatico.Enabled = False
            chkAdeguamentoIstat.Enabled = False
            chkGeneraRateProd.Enabled = False
            chkGeneraAccontiSaldo.Enabled = True
            
            cmdgenIntDaProd.Enabled = True
            cmdGeneraScadenzeProd.Enabled = True

            'SSTab1.TabEnabled(0) = False
            SSTab1.TabEnabled(1) = False
            'SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            'SSTab1.TabEnabled(6) = False
            If Me.chkGeneraAccontiSaldo.Value = vbUnchecked Then
                SSTab1.TabEnabled(0) = True
                'SSTab1.TabEnabled(5) = False
            Else
                SSTab1.TabEnabled(0) = False
                'SSTab1.TabEnabled(5) = True
            End If
            
            SSTab1.Tab = 4
            
            If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
                Me.txtImportoAttuale.Value = 0
            End If
        Case Else
        
            txtDataStipula.Enabled = True
            txtImportoStipula.Enabled = True
            txtDataDecorrenza.Enabled = True
            txtNumeroRinnovo.Enabled = True
            txtImportoAttuale.Enabled = True
            cboTipoRateizzazione.Enabled = True
            cboDurataContratto.Enabled = True
            txtDataScadenza.Enabled = True
            cboTipoRinnovo.Enabled = True
            txtDataScadenzaPerRinnovo.Enabled = True
            cboDurataAssistenza.Enabled = True
            txtDataFineAssistenza.Enabled = True
            cboDurataContrattoProx.Enabled = True
            cboTipoRinnovoProx.Enabled = True
            txtDataScadSecContr.Enabled = True
            txtDataScadPerRinnovoProx.Enabled = True
            cboTipoRateizzazioneProx.Enabled = True
            
            'cmdGeneraRate.Enabled = True
            'cmdAcconti.Enabled = False
            'cmdSaldo.Enabled = False
            
            chkRinnovoAutomatico.Enabled = True
            chkAdeguamentoIstat.Enabled = True
            chkGeneraRateProd.Enabled = False
            chkGeneraAccontiSaldo.Enabled = False
            
            cmdgenIntDaProd.Enabled = False
            cmdGeneraScadenzeProd.Enabled = False
            
            'SSTab1.TabEnabled(5) = False
    End Select


    If m_Document(m_Document.PrimaryKey).Value > 0 Then
        
        If Me.cboTipoImpostazione.CurrentID = 3 Then
            Me.chkGeneraAccontiSaldo.Enabled = False
        End If
        
        Exit Sub
    End If
    Select Case cboTipoImpostazione.CurrentID
        
        Case 1
            Me.chkContrattoAttuale.Value = vbChecked
            Me.chkAttivoPassivo.Value = vbChecked
            Me.chkAdeguamentoIstat.Value = vbChecked
            Me.chkRinnovoAutomatico.Value = vbChecked
            Me.chkFatturazioneRic.Value = vbChecked
            Me.chkTotDaiProdotti.Value = vbUnchecked
            Me.chkGeneraRateProd.Value = vbUnchecked
        Case 2
            Me.chkContrattoAttuale.Value = vbChecked
            Me.chkAttivoPassivo.Value = vbChecked
            Me.chkAdeguamentoIstat.Value = vbChecked
            Me.chkRinnovoAutomatico.Value = vbChecked
            Me.chkFatturazioneRic.Value = vbChecked
            Me.chkTotDaiProdotti.Value = vbChecked
            Me.chkGeneraRateProd.Value = vbChecked
            
        Case 3
            Me.chkContrattoAttuale.Value = vbChecked
            Me.chkAttivoPassivo.Value = vbChecked
            Me.chkAdeguamentoIstat.Value = vbUnchecked
            Me.chkRinnovoAutomatico.Value = vbUnchecked
            Me.chkFatturazioneRic.Value = vbUnchecked
            Me.chkTotDaiProdotti.Value = vbChecked
            Me.chkGeneraRateProd.Value = vbUnchecked
            
            If txtDataStipula.Value = 0 Then txtDataStipula.Value = Date
            If txtDataDecorrenza.Value = 0 Then txtDataDecorrenza.Value = Date
            
            
        Case Else
            Me.chkContrattoAttuale.Value = vbChecked
            Me.chkAttivoPassivo.Value = vbChecked
            Me.chkAdeguamentoIstat.Value = vbChecked
            Me.chkRinnovoAutomatico.Value = vbChecked
            Me.chkFatturazioneRic.Value = vbChecked
            Me.chkTotDaiProdotti.Value = vbUnchecked
            Me.chkGeneraRateProd.Value = vbUnchecked
    End Select

Exit Sub
ERR_cboTipoImpostazione_Click:
    MsgBox Err.Description, vbCritical, "cboTipoImpostazione_Click"
End Sub

Private Sub cboTipoPeriodo_Click()

If bLoadingProdotti = True Then Exit Sub

    Me.txtDataInizioProd.Enabled = True
    Me.txtOraInizioProd.Enabled = False
    Me.txtQtaPeriodo.Enabled = True

    If Me.cboTipoPeriodo.CurrentID = 2 Then
        Me.txtDataInizioProd.Enabled = False
        Me.txtOraInizioProd.Enabled = False
        Me.txtQtaPeriodo.Enabled = True
        Me.txtDataInizioProd.Value = Me.txtDataDecorrenza.Value
        Me.txtDataFineProd.Value = Me.txtDataScadenzaPerRinnovo.Value
        'Me.txtQtaPeriodo.Value = 1
    Else
        Me.txtDataInizioProd.Enabled = True
        Me.txtOraInizioProd.Enabled = True
        Me.txtQtaPeriodo.Enabled = True
        Me.txtDataInizioProd.Value = Date
        Me.txtDataFineProd.Value = Date
    End If
    

    cboUMPeriodoProd_Click
    
End Sub

Private Sub cboTipoRateizzazione_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT * FROM RV_PORateizzazione "
    sSQL = sSQL & "WHERE IDRV_PORateizzazione=" & Me.cboTipoRateizzazione.CurrentID
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        Mesi_Rate = fnNotNullN(rs!Mesi)
        Numero_Rate = fnNotNullN(rs!numerorate)
        Pagamento_Anticipato_Periodo = Abs(fnNotNullN(rs!PagamentoInizioPeriodo))
        Rata_Iniziale = fnNotNullN(rs!RataInizialeRataFinale)
        Anno_Solare = fnNotNullN(rs!AnnoSolare)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub cboTipoRateizzazioneProx_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboTipoRinnovo_Click()
On Error GoTo ERR_cboTipoRinnovo_Click
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    

    sSQL = "SELECT * FROM RV_POTipoRinnovo "
    sSQL = sSQL & "WHERE IDRV_POTipoRinnovo=" & Me.cboTipoRinnovo.CurrentID
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        Mesi_Rinnovo_Contratto = fnNotNullN(rs!Mesi)
        Giorni_Rinnovo_Contratto = fnNotNullN(rs!Giorni)
        AnnoPrecedente_Rinnovo_Contratto = fnNotNullN(rs!AnnoPrecedente)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        If Mesi_Rinnovo_Contratto > Mesi_Durata_Contratto Then
            MsgBox "Impossibile inserire questo tipo rinnovo", vbInformation, "Inserimento rinnovo contratto"
            Me.cboTipoRinnovo.WriteOn 0
            Me.txtDataScadenzaPerRinnovo.Value = 0
        Else
            Me.txtDataScadenzaPerRinnovo.Text = DateAdd("m", Mesi_Rinnovo_Contratto, Me.txtDataDecorrenza.Value) - 1
            Me.txtDataScadenzaPerRinnovo.Text = DateAdd("d", Giorni_Rinnovo_Contratto, Me.txtDataScadenzaPerRinnovo.Value)
        End If
    Else
        If fnNotNullN(m_Document("IDTipoRinnovo").Value) <> Me.cboTipoRinnovo.CurrentID Then
            Me.txtDataScadenzaPerRinnovo.Text = DateAdd("m", Mesi_Rinnovo_Contratto, Me.txtDataDecorrenza.Value) - 1
            Me.txtDataScadenzaPerRinnovo.Text = DateAdd("d", Giorni_Rinnovo_Contratto, Me.txtDataScadenzaPerRinnovo.Value)
        End If
    End If
        
    
    If Not (BrwMain.Visible) Then Change

Exit Sub
ERR_cboTipoRinnovo_Click:
    MsgBox Err.Description, vbCritical, "cboTipoRinnovo_Click"
End Sub

Private Sub cboTipoRinnovoProx_Click()
On Error GoTo ERR_cboTipoRinnovoProx_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Mesi As Long
Dim Giorni As Long

    sSQL = "SELECT * FROM RV_POTipoRinnovo "
    sSQL = sSQL & "WHERE IDRV_POTipoRinnovo=" & Me.cboTipoRinnovoProx.CurrentID
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        Mesi_Rinnovo_Contratto_prox = fnNotNullN(rs!Mesi)
        Giorni_Rinnovo_Contratto_prox = fnNotNullN(rs!Giorni)
        AnnoPrecedente_Rinnovo_Contratto_prox = fnNotNullN(rs!AnnoPrecedente)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        If Mesi_Rinnovo_Contratto_prox > Mesi_Durata_Contratto_prox Then
            MsgBox "Impossibile inserire questo tipo rinnovo", vbInformation, "Inserimento rinnovo contratto per il secondo contratto"
            Me.cboTipoRinnovoProx.WriteOn 0
            Me.txtDataScadPerRinnovoProx.Value = 0
        Else
            Me.txtDataScadPerRinnovoProx.Text = DateAdd("m", Mesi_Rinnovo_Contratto_prox + Mesi_Rinnovo_Contratto, Me.txtDataDecorrenza.Value) - 1
            Me.txtDataScadPerRinnovoProx.Text = DateAdd("d", Giorni_Rinnovo_Contratto_prox + Giorni_Rinnovo_Contratto_prox, Me.txtDataScadPerRinnovoProx.Value)
        End If
    Else
        If fnNotNullN(m_Document("IDTipoRinnovoProssimoRinnovo").Value) <> Me.cboTipoRinnovoProx.CurrentID Then
            Me.txtDataScadPerRinnovoProx.Text = DateAdd("m", Mesi_Rinnovo_Contratto_prox + Mesi_Rinnovo_Contratto, Me.txtDataDecorrenza.Value) - 1
            Me.txtDataScadPerRinnovoProx.Text = DateAdd("d", Giorni_Rinnovo_Contratto_prox + Giorni_Rinnovo_Contratto_prox, Me.txtDataScadPerRinnovoProx.Value)
        End If
    End If
        
    
    If Not (BrwMain.Visible) Then Change

Exit Sub
ERR_cboTipoRinnovoProx_Click:
    MsgBox Err.Description, vbCritical, "cboTipoRinnovoProx"
End Sub

Private Sub cboUMPeriodoProd_Click()
    On Error GoTo ERR_cboUMPeriodoProd_Click
    
    
    If bLoadingProdotti = True Then Exit Sub
    
'    Me.txtDataInizioProd.Enabled = True
'    Me.txtOraInizioProd.Enabled = False
'    Me.txtQtaPeriodo.Enabled = True
    
    
    Me.txtDataFineProd.Enabled = False
    Me.txtQtaPeriodo.Enabled = True
    'Me.txtOraInizioProd.Enabled = False
    
    If Me.cboTipoPeriodo.CurrentID = 1 Then
        Me.txtDataInizioProd.Enabled = True
    End If
    If Me.cboTipoPeriodo.CurrentID <= 1 Then
        If fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0 Then
            If Me.txtDataDecorrenza.Value = 0 Then
                Me.txtDataInizioProd.Value = Date
            Else
                Me.txtDataInizioProd.Value = Me.txtDataDecorrenza.Value
            End If
            
            Me.txtOraInizioProd.Value = 0
            Me.txtOraFineProd.Value = 0
        End If
    End If
        
    Select Case Me.cboUMPeriodoProd.CurrentID
        Case 1
            Me.txtDataInizioProd.Enabled = False
            Me.txtOraInizioProd.Enabled = True
        Case 2
            Me.txtQtaPeriodo.Value = DateDiff("d", Me.txtDataInizioProd.Text, Me.txtDataFineProd.Text) + 1
            If Me.cboTipoPeriodo.CurrentID = 1 Then Me.txtDataFineProd.Enabled = True
        Case 3
            Me.txtQtaPeriodo.Value = DateDiff("ww", Me.txtDataInizioProd.Text, Me.txtDataFineProd.Text)
        Case 4
            If (Day(Me.txtDataInizioProd) = 1) Then
                Me.txtQtaPeriodo.Value = DateDiff("m", Me.txtDataInizioProd.Text, Me.txtDataFineProd.Text) + 1
            Else
                Me.txtQtaPeriodo.Value = DateDiff("m", Me.txtDataInizioProd.Text, Me.txtDataFineProd.Text)
            End If
        Case 5
            If (Day(Me.txtDataInizioProd) = 1) Then
                Me.txtQtaPeriodo.Value = DateDiff("yyyy", Me.txtDataInizioProd.Text, Me.txtDataFineProd.Text) + 1
            Else
                Me.txtQtaPeriodo.Value = DateDiff("yyyy", Me.txtDataInizioProd.Text, Me.txtDataFineProd.Text)
            End If
        Case 6
            Me.txtDataInizioProd.Value = Me.txtDataDecorrenza.Value
            Me.txtDataFineProd.Value = Me.txtDataScadenzaPerRinnovo.Value
            Me.txtQtaPeriodo.Value = 1
            Me.txtDataInizioProd.Enabled = False
            Me.txtDataFineProd.Enabled = False
            Me.txtQtaPeriodo.Enabled = False
            Me.txtOraInizioProd.Enabled = False
    End Select
           
    txtQtaPeriodo_LostFocus
       
'    End If
Exit Sub
ERR_cboUMPeriodoProd_Click:
    
    
End Sub

Private Sub CDAgente_ChangeElement()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub CDAmministratore_ChangeElement()




If Not (BrwMain.Visible) Then Change
End Sub



Private Sub CDArticoloAdeg_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim LINK_IVA_ARTICOLO As Long
Dim Testo As String
On Error Resume Next

If BLoading = 1 Then Exit Sub
If fnNotNullN(m_DocumentsLink2(m_DocumentsLink2.PrimaryKey).Value) > 0 Then Exit Sub
    
sSQL = "SELECT IDIvaVendita FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & Me.CDArticoloAdeg.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    LINK_IVA_ARTICOLO = 0
Else
    LINK_IVA_ARTICOLO = fnNotNullN(rs!IDIvaVendita)
End If

rs.CloseResultset
Set rs = Nothing
    
If ((Me.cboIvaAdeg.CurrentID > 0) And (LINK_IVA_ARTICOLO > 0)) Then
    If LINK_IVA_ARTICOLO <> Me.cboIvaAdeg.CurrentID Then
        Testo = "L'iva dell'articolo è diversa da quella inserita" & vbCrLf
        Testo = Testo & "Vuoi cambiarla?"
            
        If MsgBox(Testo, vbQuestion + vbYesNo, "Inserimento articolo") = vbNo Then Exit Sub
        
        Me.cboIvaAdeg.WriteOn LINK_IVA_ARTICOLO
    End If
Else
    Me.cboIvaAdeg.WriteOn LINK_IVA_ARTICOLO
End If

End Sub


Private Sub CDArticoloProd_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

On Error Resume Next

If BLoading = 1 Then Exit Sub

If bLoadingProdotti = True Then Exit Sub


If fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0 Then

    sSQL = "SELECT IDUnitaDiMisuraVendita, IDIvaVendita FROM Articolo "
    sSQL = sSQL & "WHERE IDArticolo=" & Me.CDArticoloProd.KeyFieldID
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        Me.cboUMArtProd.WriteOn 0
        Me.cboIvaProd.WriteOn 0
    Else
        Me.cboUMArtProd.WriteOn fnNotNullN(rs!IDUnitaDiMisuraVendita)
        Me.cboIvaProd.WriteOn fnNotNullN(rs!IDIvaVendita)
    End If
    
    
    GET_PREZZO_ARTICOLO Me.CDArticoloProd.KeyFieldID, Me.cboListinoProd.CurrentID, LINK_LISTINO_AZIENDA, Me.CDCliente.KeyFieldID
    
    GET_TOTALI_RIGA_DETTAGLIO
    
Else

    If fnNotNullN(m_DocumentsLink3("IDArticolo").Value <> Me.CDArticoloProd.KeyFieldID) Then

        sSQL = "SELECT IDUnitaDiMisuraVendita, IDIvaVendita FROM Articolo "
        sSQL = sSQL & "WHERE IDArticolo=" & Me.CDArticoloProd.KeyFieldID
        
        Set rs = Cn.OpenResultset(sSQL)
        
        If rs.EOF Then
            Me.cboUMArtProd.WriteOn 0
            Me.cboIvaProd.WriteOn 0
        Else
            Me.cboUMArtProd.WriteOn fnNotNullN(rs!IDUnitaDiMisuraVendita)
            Me.cboIvaProd.WriteOn fnNotNullN(rs!IDIvaVendita)
        End If
        
        
        GET_PREZZO_ARTICOLO Me.CDArticoloProd.KeyFieldID, Me.cboListinoProd.CurrentID, LINK_LISTINO_AZIENDA, Me.CDCliente.KeyFieldID
        
        GET_TOTALI_RIGA_DETTAGLIO
        
    End If
    
End If

End Sub


Private Sub CDCliente_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    CHANGE_ANAGRAFICA_FATTURAZIONE Me.CDCliente.KeyFieldID
    
    'Me.txtAltriDati.Text = GET_CARATTERISTICHE_RISORSA(Me.MSFlexGrid1) 'GET_DESCRIZIONE_ALTRI_DATI
    
    GET_CARATTERISTICHE_RISORSA Me.MSFlexGrid1
End If

fncSitoPerAnagrafica

If Not (BrwMain.Visible) Then Change

End Sub




Private Sub CDServizio_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
On Error Resume Next

If BLoading = 1 Then Exit Sub

If fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value) <= 0 Then
    
    sSQL = "SELECT * FROM RV_POTipoContrattoServizi "
    sSQL = sSQL & "WHERE IDArticolo=" & Me.CDServizio.KeyFieldID
    sSQL = sSQL & " AND IDTipoContratto=" & Me.cboTipoContratto.CurrentID
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        rs.CloseResultset
        Set rs = Nothing
        
        sSQL = "SELECT * FROM RV_POConfigurazioneServizio "
        sSQL = sSQL & "WHERE IDArticolo=" & Me.CDServizio.KeyFieldID
        sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
        
        Set rs = Cn.OpenResultset(sSQL)
        If rs.EOF Then
            'Me.CDServizio.Load 0
            Me.cboCriterioRicorrenza.WriteOn 0
            Me.txtOgniNumeroGiorni.Value = 0
            Me.txtOgniNumeroMesi.Value = 0
            Me.txtOgniNumeroSettimane.Value = 0
            Me.cboTipoDataInizioRic.WriteOn 0
            Me.txtGiornoFissoInizioRic.Value = 0
            Me.txtMeseFissoInizioRic.Value = 0
            Me.cboTipoDataFineRic.WriteOn 0
            Me.txtGiornoFissoFineRic.Value = 0
            Me.txtMeseFissoFineRic.Value = 0
            Me.txtNumeroRicorrenze.Value = 0
            Me.cboTipoAnnoInizioRicorr.WriteOn 0
            Me.cboTipoAnnoFineRicorr.WriteOn 0
        Else
            Me.cboCriterioRicorrenza.WriteOn fnNotNullN(rs("IDRV_POCriterioRicorrenza").Value)
            Me.txtOgniNumeroGiorni.Value = fnNotNullN(rs("OgniNumeroGiorni").Value)
            Me.txtOgniNumeroMesi.Value = fnNotNullN(rs("OgniNumeroMesi").Value)
            Me.txtOgniNumeroSettimane.Value = fnNotNullN(rs("OgniNumeroSettimane").Value)
            Me.cboTipoDataInizioRic.WriteOn fnNotNullN(rs("IDRV_POTipoDataInizioRicorrenza").Value)
            Me.txtGiornoFissoInizioRic.Value = fnNotNullN(rs("GiornoInizioRicorrenza").Value)
            Me.txtMeseFissoInizioRic.Value = fnNotNullN(rs("MeseInizioRicorrenza").Value)
            Me.cboTipoDataFineRic.WriteOn fnNotNullN(rs("IDRV_POTipoDataFineRicorrenza").Value)
            Me.txtGiornoFissoFineRic.Value = fnNotNullN(rs("GiornoFineRicorrenza").Value)
            Me.txtMeseFissoFineRic.Value = fnNotNullN(rs("MeseFineRicorrenza").Value)
            Me.txtNumeroRicorrenze.Value = fnNotNullN(rs("NumeroRicorrenze").Value)
            Me.cboTipoAnnoInizioRicorr.WriteOn fnNotNullN(rs("IDRV_POTipoAnnoInizioRicorrenza").Value)
            Me.cboTipoAnnoFineRicorr.WriteOn fnNotNullN(rs("IDRV_POTipoAnnoFineRicorrenza").Value)
        
        End If
        
        rs.CloseResultset
        Set rs = Nothing
    Else
        'Me.CDServizio.Load fnNotNullN(rs("IDArticolo").Value)
        Me.cboCriterioRicorrenza.WriteOn fnNotNullN(rs("IDRV_POCriterioRicorrenza").Value)
        Me.txtOgniNumeroGiorni.Value = fnNotNullN(rs("OgniNumeroGiorni").Value)
        Me.txtOgniNumeroMesi.Value = fnNotNullN(rs("OgniNumeroMesi").Value)
        Me.txtOgniNumeroSettimane.Value = fnNotNullN(rs("OgniNumeroSettimane").Value)
        Me.cboTipoDataInizioRic.WriteOn fnNotNullN(rs("IDRV_POTipoDataInizioRicorrenza").Value)
        Me.txtGiornoFissoInizioRic.Value = fnNotNullN(rs("GiornoInizioRicorrenza").Value)
        Me.txtMeseFissoInizioRic.Value = fnNotNullN(rs("MeseInizioRicorrenza").Value)
        Me.cboTipoDataFineRic.WriteOn fnNotNullN(rs("IDRV_POTipoDataFineRicorrenza").Value)
        Me.txtGiornoFissoFineRic.Value = fnNotNullN(rs("GiornoFineRicorrenza").Value)
        Me.txtMeseFissoFineRic.Value = fnNotNullN(rs("MeseFineRicorrenza").Value)
        Me.txtNumeroRicorrenze.Value = fnNotNullN(rs("NumeroRicorrenze").Value)
        Me.cboTipoAnnoInizioRicorr.WriteOn fnNotNullN(rs("IDRV_POTipoAnnoInizioRicorrenza").Value)
        Me.cboTipoAnnoFineRicorr.WriteOn fnNotNullN(rs("IDRV_POTipoAnnoFineRicorrenza").Value)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    
End If
    cboCriterioRicorrenza_Click
    cboTipoDataFineRic_Click
    cboTipoDataInizioRic_Click
End Sub

Private Sub CDTecnico_ChangeElement()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub chkACorpo_Click()
    If bLoadingProdotti = True Then Exit Sub
    
        
    GET_TOTALI_RIGA_DETTAGLIO
End Sub

Private Sub chkAdeguamentoIstat_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkAnnoSolareInt_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkAttivoPassivo_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkChiuso_Click()
    
    m_DocumentsLink3_OnReposition
        
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkConducente_Click()
        
    
    
    If Me.chkConducente.Value = vbUnchecked Then
        Me.cboAnaOperatoreProd.WriteOn 0
        Me.cboAnaOperatoreProd.Enabled = False
    Else
        Me.cboAnaOperatoreProd.Enabled = True
    End If
End Sub

Private Sub chkDisdetta_Click()
    If DISDETTO_NONFATTURARE = 1 Then
        Me.chkNonFatturare.Value = Me.chkDisdetta.Value
    End If

    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkEscludiGiorniFestivi_Click()
    If bLoadingProdotti = True Then Exit Sub
    
    If Me.cboUMPeriodoProd.CurrentID = 2 Then
        Me.txtQuantitaEffettiva.Value = GET_CALCOLO_QUANTITA_EFFETTIVA(Me.txtDataInizioProd.Text, Me.txtDataFineProd.Text)
    End If
    
    GET_TOTALI_RIGA_DETTAGLIO
End Sub

Private Sub chkEscludiSabato_Click()
    
    If bLoadingProdotti = True Then Exit Sub
    
        
    If Me.cboUMPeriodoProd.CurrentID = 2 Then
        Me.txtQuantitaEffettiva.Value = GET_CALCOLO_QUANTITA_EFFETTIVA(Me.txtDataInizioProd.Text, Me.txtDataFineProd.Text)
    End If
    
    GET_TOTALI_RIGA_DETTAGLIO
    
End Sub

Private Sub chkFineContratto_Click()
        If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkGeneraAccontiSaldo_Click()

'    If chkGeneraAccontiSaldo.Value = vbChecked Then
'        cmdAcconti.Enabled = True
'        cmdSaldo.Enabled = True
'        Me.Label1(35).Caption = "Saldo"
'        SSTab1.TabEnabled(0) = False
'        SSTab1.TabEnabled(5) = True
'
'    Else
'        cmdAcconti.Enabled = False
'        cmdSaldo.Enabled = False
'        Me.Label1(35).Caption = "Tot. contr. + Adeg."
'        SSTab1.TabEnabled(0) = True
'        SSTab1.TabEnabled(5) = False
'
'    End If
    
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkGeneraRateProd_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkNonFatturare_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkOfferta_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkProdottoGenerico_Click()
    If Me.chkProdottoGenerico.Value = vbUnchecked Then
        
        Me.txtQtaProdotto.Enabled = False
    Else
        Me.txtQtaProdotto.Enabled = True
    End If
End Sub

Private Sub chkRinnovoAutomatico_Click()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub chkRitAcconto_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cmdAcconti_Click()

    If NON_MODIFICA_CONTRATTO = 1 Then Exit Sub

    If (m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    If Me.txtImportoTotAdeg.Value <= 0 Then Exit Sub
    
    
    Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    LINK_OGGETTO_CONTRATTO_SEL = fnNotNullN(m_Document("IDOggetto").Value)
    LINK_TIPO_OGGETTO_CONTRATTO_SEL = m_DocType.ID
    OPERAZIONE_ESEGUITA_ACCONTO = 0
    
    frmAcconto.Show vbModal
    
    'FUNZIONE CHE INSERISCI GLI ACCONTI IN GRIGLIA
    'GET_GRIGLIA_ACCONTI
    
    If OPERAZIONE_ESEGUITA_ACCONTO = 1 Then
        
        OnSave
    End If

End Sub

Private Sub cmdAnaContratto_Click()
    If Me.CDCliente.KeyFieldID = 0 Then Exit Sub
    
    LINK_SITO_PER_ANAGRAFICA_TEL = 0
    LINK_ANAGRAFICA_TEL = Me.CDCliente.KeyFieldID
    
    frmRiferimentiTelefonici.Show vbModal
End Sub

Private Sub cmdApri_Click()
On Error GoTo ERR_cmdApri_Click
    frmAltriDati.Show vbModal
    
    If CONFERMA_MODIFICA = 1 Then
        'Me.txtAltriDati.Text = GET_CARATTERISTICHE_RISORSA(Me.MSFlexGrid1) 'GET_DESCRIZIONE_ALTRI_DATI
        GET_CARATTERISTICHE_RISORSA Me.MSFlexGrid1
        If Not (BrwMain.Visible) Then Change

    End If

Exit Sub
ERR_cmdApri_Click:
    MsgBox Err.Description, vbCritical, "cmdApri_Click"
End Sub

Private Sub cmdAttivaContProd_Click()
If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di procedere ad inserire i servizi legati al tipo di contratto", vbInformation, "Salvataggio documento"
    Exit Sub
End If
If fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare la riga prima di procedere ad inserire la configurazione dei contatori", vbInformation, "Salvataggio documento"
    Exit Sub
End If

If fnNotNullN(m_DocumentsLink3("IDRV_POProdotto").Value) <= 0 Then
    Exit Sub
End If

If MODULO_ATTIVATO_CONT = 0 Then
    If Len(MODULO_DESCRIZIONE_CONT) > 0 Then
        MsgBox "Il modulo " & MODULO_DESCRIZIONE_CONT & " non è stato abilitato", vbInformation, TheApp.FunctionName
    Else
        MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
    End If
Exit Sub
End If

Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
LINK_PRODOTTO_SEL = fnNotNullN(m_DocumentsLink3("IDRV_POProdotto").Value)
LINK_PRODOTTO_RIGA_SEL = fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)

frmContatori.Show vbModal


End Sub

Private Sub cmdConfiguraServizi_Click()
On Error GoTo ERR_cmdConfiguraServizi_Click
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        MsgBox "Salvare il documento prima di procedere ad inserire i servizi legati al tipo di contratto", vbInformation, "Salvataggio documento"
        Exit Sub
    End If
    
    Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

    frmConfiguraServizi.Show vbModal
    
    m_DocumentsLink1.Refresh
   
Exit Sub
ERR_cmdConfiguraServizi_Click:
    MsgBox Err.Description, vbCritical, "cmdConfiguraServizi_Click"
End Sub

Private Sub cmdContatori_Click()
On Error GoTo ERR_cmdContatori_Click
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    If fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0 Then Exit Sub
    
    If Me.txtIDProdotto.Value = 0 Then Exit Sub
    
    If MODULO_ATTIVATO_CONT = 0 Then
        If Len(MODULO_DESCRIZIONE_CONT) > 0 Then
            MsgBox "Il modulo " & MODULO_DESCRIZIONE_CONT & " non è stato abilitato", vbInformation, TheApp.FunctionName
        Else
            MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
        End If
    Exit Sub
    End If
    
    
    LblLinkRil.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POContatoreRilevamenti"))
    Me.LblLinkRil.IDReturn = fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)
    Me.LblLinkRil.RunApplication
    
Exit Sub
ERR_cmdContatori_Click:
    MsgBox Err.Description, vbCritical, "cmdContatori_Click"
End Sub

Private Sub cmdDocumentazione_Click()
On Error GoTo ERR_cmdDocumentazione_Click
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

    LINK_AZIENDA_DOC = TheApp.IDFirm
    LINK_CLIENTE_DOC = Me.CDCliente.KeyFieldID
    LINK_CONTRATTO_DOC = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    LINK_CONTRATTO_DOC_PADRE = Me.txtIDContrattoPadre.Value
    LINK_INTERVENTO_PADRE_DOC = 0
    LINK_INTERVENTO_DOC = 0
    LINK_STARTUP_DOC = 3
    frmDocumentazione.Show vbModal

Exit Sub
ERR_cmdDocumentazione_Click:
    MsgBox Err.Description, vbCritical, "cmdDocumentazione_Click"
    
End Sub

Private Sub cmdEliminaAcconto_Click()
Dim Testo As String
Dim IDOggettoVend As Long
Dim IDTipoOggettoVend As Long

'IDOggettoVend = fnNotNullN(Me.GrigliaAcconti.AllColumns("IDOggetto").Value)
'IDTipoOggettoVend = fnNotNullN(Me.GrigliaAcconti.AllColumns("IDTipoOggetto").Value)
'
'If IDOggettoVend = 0 Then Exit Sub
'
'
'If Me.chkChiuso.Value = vbChecked Then
'    MsgBox "Il contratto è chiuso", vbInformation, "Controllo dati"
'    Exit Sub
'End If
'
'
'Testo = "Sei sicuro di voler eliminare il riferimento al documento selezionato?"
'
'If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione rif. acconto/saldo") = vbNo Then Exit Sub
'
'ELIMINA_FLUSSO_DOCUMENTALE_ACCONTO IDTipoOggettoVend, IDOggettoVend, fnNotNullN(m_Document("IDOggetto").Value), m_DocType.ID, fnNotNull(Me.GrigliaAcconti.AllColumns("DescrizioneFlusso").Value)
'
'OnSave



End Sub

Private Sub cmdEliminaAdeg_Click()
On Error GoTo ERR_cmdEliminaAdeg_Click



Dim LINK_LOCAL_ADEGUAMENTO As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDOggettoScadenza As Long
Dim IDOggettoRata As Long
Dim IDTipoOggettoRata As Long
Dim Testo As String


If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
If fnNotNullN(m_DocumentsLink2(m_DocumentsLink2.PrimaryKey).Value) <= 0 Then Exit Sub

Testo = "Vuoi eliminare l'adeguamento selezionato?"

If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione adeguamento") = vbNo Then Exit Sub

If GET_RATA_PAGATA_ADEGUAMENTO(fnNotNullN(m_DocumentsLink2(m_DocumentsLink2.PrimaryKey).Value), fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) = True Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Impossibile eliminare poichè una o più rate di questo adeguamento risultano fatturate"
    MsgBox Testo, vbInformation, "Eliminazione adeguamento"
    Exit Sub
End If

Screen.MousePointer = 11
LINK_LOCAL_ADEGUAMENTO = fnNotNullN(m_DocumentsLink2(m_DocumentsLink2.PrimaryKey).Value)

m_DocumentsLink2.Delete

''''CLICLO DI ELIMINAZIONE IDOGGETTI DELLE RATE DI ADEGUAMENTO'''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento=" & LINK_LOCAL_ADEGUAMENTO

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    ''''''ELIMINAZIONE SCADENZA DEL CONTRATTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), 0)
    
    If IDOggettoScadenza > 0 Then
        ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)
        ELIMINA_SCADENZA IDOggettoScadenza
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ELIMINA_LINK_OGGETTO_RATA fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''ELIMINAZIONE DELLE RATE DI ADEGUAMENTO'''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContrattoAdeguamento=" & LINK_LOCAL_ADEGUAMENTO
Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
m_DocumentsLink.Refresh
m_DocumentsLink2.Refresh
Screen.MousePointer = 0

    Me.txtImportoTotAdeg.Value = GET_TOTALE_ADEGUAMENTI_DETTAGLIO(fnNotNullN(m_Document(m_Document.PrimaryKey).Value))

Exit Sub
ERR_cmdEliminaAdeg_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaAdeg_Click"
    Screen.MousePointer = 0
End Sub

Private Sub cmdEliminaProd_Click()
On Error GoTo ERR_cmdEliminaProd_Click
Dim Testo As String
Dim sSQL As String
Dim Link_Riga As Long
Dim LINK_PRODOTTO As Long
Dim Elimina_Intervento As Long


If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

If fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0 Then Exit Sub


Elimina_Intervento = 0

If OBBLIGATORIO_COLLEGAMENTO_INT = 1 Then
    If GET_CONTROLLO_ESISTENZA_RIGA_PRODOTTO(fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)) = True Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Questa riga del prodotto è associato a uno o più interventi" & vbCrLf
        Testo = Testo & "Impossibile continuare"
        
        MsgBox Testo, vbCritical, "Eliminazione dati"
        Exit Sub
        
    End If
End If

If GET_CONTROLLO_ESISTENZA_PRODOTTO_SERVIZIO(fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)) = True Then
    
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Questa riga del prodotto è associato a uno o più servizi di questo contratto" & vbCrLf
    Testo = Testo & "Per continuare con questo comando bisogna eliminare il collegamento con il servizio"
    Testo = Testo & "Impossibile continuare"
    
    MsgBox Testo, vbCritical, "Eliminazione dati"
    Exit Sub
End If

If Me.cboTipoImpostazione.CurrentID = 2 Then
    If GET_CONTROLLO_RIGA_PROD_FATT(fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)) = True Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Una o più rate del contratto risultano fatturate, pertanto è impossibile eliminare il prodotto la riga selezionata " & vbCrLf
        Testo = Testo & "Impossibile continuare"
    
        MsgBox Testo, vbCritical, "Eliminazione dati"
        Exit Sub
    End If
End If

If (Me.cboTipoImpostazione.CurrentID = 3) Then

    If GET_CONTROLLO_RIGA_PROD_FATT(fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)) = True Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Una o più rate di questa riga del contratto risultano fatturate, pertanto è impossibile eliminare o modificare la riga selezionata " & vbCrLf
        Testo = Testo & "Impossibile continuare"
        MsgBox Testo, vbCritical, "Validazione dati"
        Exit Sub
    End If

    If fnNotNullN(m_DocumentsLink3("IDRV_POInterventoRigheDett").Value) > 0 Then
        If GET_CONTROLLO_ESISTENZA_ADD_PROD(fnNotNullN(m_DocumentsLink3("IDRV_POInterventoRigheDett").Value)) = True Then
            MsgBox "Impossibile eliminare una riga generata dalla gestione degli addebiti intervento", vbInformation, "Controllo dati"
            Exit Sub
        End If
    End If
    
    If fnNotNullN(m_DocumentsLink3("IDRV_POContatoreRilevamenti").Value) > 0 Then
        If GET_CONTROLLO_ESISTENZA_RIL_PROD(fnNotNullN(m_DocumentsLink3("IDRV_POContatoreRilevamenti").Value)) = True Then
            MsgBox "Impossibile eliminare una riga generata dalla gestione rilevamenti contatore", vbInformation, "Controllo dati"
            Exit Sub
        End If
    End If
    

    
    
End If

Testo = "Sei sicuro di voler eliminare la riga?"

If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riga") = vbNo Then Exit Sub

Link_Riga = fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)
LINK_PRODOTTO = txtIDProdotto.Value

Screen.MousePointer = 11

m_DocumentsLink3.Delete

AGGIORNA_INTERVENTI_PRODOTTO_PER_ELIMINAZIONE LINK_PRODOTTO, Link_Riga

ELIMINAZIONE_ASSOCIAZIONE_PRODOTTO_SERVIZIO Link_Riga

ELIMINA_INTERVENTI_COLLEGATI Link_Riga

Screen.MousePointer = 0

If Me.cboTipoImpostazione.CurrentID <> 1 Then
    Me.txtImportoAttuale.Value = GET_TOTALE_CONTRATTO(m_Document(m_Document.PrimaryKey).Value)
    OnSave
End If


GET_TOTALE_PRODOTTI fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

Exit Sub
ERR_cmdEliminaProd_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaProd_Click"
    Screen.MousePointer = 0
    
End Sub

Private Sub cmdEliminaRata_Click()
On Error GoTo ERR_cmdEliminaRata_Click
Dim sSQL As String
Dim IDOggettoRata As Long
Dim IDTipoOggettoRata As Long
Dim IDOggettoVend As Long
Dim IDTipoOggettoVend As Long
Dim IDOggettoScadenza As Long
Dim Testo As String

If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

Testo = "Sei sicuro di voler eliminare la rata selezionata?"

If MsgBox(Testo, vbQuestion + vbYesNo, TheApp.FunctionName) = vbNo Then Exit Sub

If Abs(fnNotNullN(m_DocumentsLink("Fatturata").Value)) = 1 Then
    If MsgBox("Questa rata risulta fatturata!" & vbCrLf & "Vuoi continuare?", vbInformation + vbYesNo, "Eliminazione dato") = vbNo Then Exit Sub
End If

Screen.MousePointer = 11

IDOggettoRata = fnNotNullN(m_DocumentsLink("IDOggetto").Value)
IDTipoOggettoRata = fnNotNullN(m_DocumentsLink("IDTipoOggetto").Value)
IDOggettoVend = fnNotNullN(m_DocumentsLink("IDOggettoCollegato").Value)
IDTipoOggettoVend = fncIDDocumentoAllegato(IDOggettoVend)


''''''ELIMINAZIONE SCADENZA DEL CONTRATTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(IDOggettoRata, IDTipoOggettoRata, 0)

If IDOggettoScadenza > 0 Then
    ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, IDOggettoRata, IDTipoOggettoRata
    ELIMINA_SCADENZA IDOggettoScadenza
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''ELIMINAZIONE COLLEGAMENTO AL FLUSSO DOCUMENTALE DEL CONTRATTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ELIMINA_LINK_OGGETTO_RATA IDOggettoRata, IDTipoOggettoRata

If IDOggettoVend > 0 Then
    ELIMINA_FLUSSO_DOCUMENTALE IDTipoOggettoVend, IDOggettoVend, IDOggettoRata, IDTipoOggettoRata
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

m_DocumentsLink.Delete

Screen.MousePointer = 0

Exit Sub
ERR_cmdEliminaRata_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaRata_Click"
    Screen.MousePointer = 0
End Sub
Private Sub cmdEliminaRif_Click()
Dim Testo As String

If Me.txtIDOggettoCollegato.Value > 0 Then
    Testo = "ATTENZIONE!!!!" & vbCrLf
    Testo = Testo & "Si sta tentando di eliminare un collegamento di una rata del contratto ad un documento di vendita" & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    If MsgBox(Testo, vbQuestion + vbYesNo, "Elimina riferimento") = vbNo Then Exit Sub
    
    
    Me.txtIDOggettoCollegato.Value = 0
    Me.chkRataFatturata.Value = 0
End If

End Sub
Private Sub cmdEliminaServizio_Click()
Dim Testo As String
Dim Link_Riga As Long
Dim sSQL As String

If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

If fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value) <= 0 Then Exit Sub

If GET_CONTROLLO_ESISTENZA_RIGA_SERVIZIO(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value) = True Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Il servizio per questo contratto risulta collegato ad uno o più interventi lavorati" & vbCrLf
    Testo = Testo & "Impossibile continuare"
        
    MsgBox Testo, vbCritical, "Eliminazione dati"
    Exit Sub
End If


If MsgBox("Sei sicuro di eliminare il servizio?", vbQuestion + vbYesNo, "Eliminazione dati") = vbNo Then
    Exit Sub
End If

Link_Riga = fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value)

m_DocumentsLink1.Delete


sSQL = "DELETE FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POContrattoServizi=" & Link_Riga
Cn.Execute sSQL

sSQL = "DELETE FROM RV_POContrattoServiziProdotti "
sSQL = sSQL & "WHERE IDRV_POContrattoServizi=" & Link_Riga
Cn.Execute sSQL



End Sub



Private Sub cmdGeneraInterventi_Click()
'On Error GoTo ERR_cmdGeneraInterventi_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim TestoMessaggio As String
Dim AvviaProcedura As Boolean
Dim rsInterventoTesta As ADODB.Recordset
Dim X As Long
Dim ErroreCoda As Boolean
Dim OLD_Cursor As Long


If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di procedere ad inserire i servizi legati al tipo di contratto", vbInformation, "Salvataggio documento"
    Exit Sub
End If
If fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il la riga del servizio prima di procedere alla generazione degli interventi", vbInformation, "Salvataggio documento"
    Exit Sub
End If

Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
Link_Contratto_Servizio = fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value)

frmInterventiDaServ.Show vbModal

Me.txtNProdAss.Value = GET_NUMERO_PRODOTTI_PER_SERVIZIO(fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))
Me.txtNInterventiServ.Value = GET_NUMERO_INTERVENTI_PER_SERVIZIO(fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))

GET_GRIGLIA_INTERVENTI

Exit Sub



'CONTROLLO PER ELIMINAZIONE E RIGENERAZIONE DEGLI INTERVENTI PER CONTRATTI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
TestoMessaggio = "ATTENZIONE!!!" & vbCrLf
TestoMessaggio = TestoMessaggio & "Questo comando eliminerà tutti gli interventi elaborati per questo contratto per poi rigenerarli" & vbCrLf
TestoMessaggio = TestoMessaggio & "Vuoi continuare?"

sSQL = "SELECT IDArticolo FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POContratto=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
sSQL = sSQL & " AND Elaborata=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    'SE NON SONO PRESENTI INTERVENTI ELABORATI DAL CONTRATTO SI VA AVANTI AD ELABORARE
    AvviaProcedura = True

    rs.CloseResultset
    Set rs = Nothing

Else
    'SE SONO PRESENTI INTERVENTI ELABORATI DAL CONTRATTI SARA' L'UTENTE A DECIDERE COSA FARE
    If MsgBox(TestoMessaggio, vbQuestion + vbYesNo, "generazione interventi da contratto") = vbNo Then
        'SE RISPONDE NO ALLA DOMANDA NON SI ELABORA NULLA
        AvviaProcedura = False
        
        rs.CloseResultset
        Set rs = Nothing
        
    Else
        rs.CloseResultset
        Set rs = Nothing
        
        AvviaProcedura = True
        
        'SE VIENE RISPOSTO SI E QUINDI CONTINUARE AL PRIMO COMANDO SI ESEGUE IL SECONDO CONTROLLO
        'PER ELIMINAZIONE E RIGENERAZIONE DEGLI INTERVENTI PER CONTRATTI DOVE QUALCHE INTERVENTO'''
        'E' STATO MODIFICATO DOPO L'ELABORAZIONE                                                            '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        TestoMessaggio = "ATTENZIONE!!!" & vbCrLf
        TestoMessaggio = TestoMessaggio & "Sono presenti interventi generati dal contratto e modificati in un secondo momento" & vbCrLf
        TestoMessaggio = TestoMessaggio & "Se si dovesse continuare con questo comando tutte le modifiche andranno perse" & vbCrLf
        TestoMessaggio = TestoMessaggio & "Vuoi continuare?"
        
        sSQL = "SELECT IDArticolo FROM RV_POIntervento "
        sSQL = sSQL & "WHERE IDRV_POContratto=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        sSQL = sSQL & " AND Elaborata=1"
        sSQL = sSQL & " AND Manuale=1"
        
        Set rs = Cn.OpenResultset(sSQL)
        
        If rs.EOF Then
            AvviaProcedura = True
        Else
            If MsgBox(TestoMessaggio, vbQuestion + vbYesNo, "generazione interventi da contratto") = vbNo Then
                AvviaProcedura = False
            Else
                AvviaProcedura = True
            End If
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
    
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If AvviaProcedura = False Then Exit Sub


''''''''''''PROCEDURA DI ELIMINAZIONE DEGLI INTERVENTI GENERATI DAL CONTRATTO'''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POContratto=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
'sSQL = sSQL & " AND IDRV_POStoriaContratto=" & Link_StoriaContratto
sSQL = sSQL & " AND Elaborata=1"

Set rsInterventoTesta = New ADODB.Recordset

rsInterventoTesta.Open sSQL, Cn.InternalConnection

If Not rsInterventoTesta.EOF Then
    While Not rsInterventoTesta.EOF
        
        '''ELIMINAZIONE BUONI INTERVENTO
        sSQL = "DELETE FROM RV_POInterventoRigheDett "
        sSQL = sSQL & "WHERE IDRV_POIntervento=" & fnNotNullN(rsInterventoTesta!IDRV_POIntervento)
        Cn.Execute sSQL
        
        '''ELIMINAZIONE BUONI INTERVENTO
        sSQL = "DELETE FROM RV_PODocumentazione "
        sSQL = sSQL & "WHERE IDRV_POIntervento=" & fnNotNullN(rsInterventoTesta!IDRV_POIntervento)
        Cn.Execute sSQL

        '''ELIMINAZIONE BUONI INTERVENTO
        sSQL = "DELETE FROM RV_POInterventoEmail "
        sSQL = sSQL & "WHERE IDRV_POIntervento=" & fnNotNullN(rsInterventoTesta!IDRV_POIntervento)
        Cn.Execute sSQL
        
        sSQL = "DELETE FROM Appuntamento "
        sSQL = sSQL & "WHERE RV_POIDIntervento=" & fnNotNullN(rsInterventoTesta!IDRV_POIntervento)
        Cn.Execute sSQL
        
        
    rsInterventoTesta.MoveNext
    Wend
End If

rsInterventoTesta.Close
Set rsInterventoTesta = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''ELIMINAZIONE DEGLI INTERVENTI DI TESTA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POContratto=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
'sSQL = sSQL & " AND IDRV_POStoriaContratto=" & Link_StoriaContratto
sSQL = sSQL & " AND Elaborata=1"
Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


SCRIVI_CODA fnNotNullN(m_Document(m_Document.PrimaryKey)), fnGetTipoOggetto("RV_POIntervento")
APERTURA_FORM_CODA = False
NOME_GESTORE = "RV_POIntervento"

''''''''''''''''''''''''''''''CONTROLLA LA CODA DEI SALVATAGGI'''''''''''''''''''''''''''''
    X = 0
    ErroreCoda = False
    Do
        X = GET_NUMERO_DOCUMENTO(False, fnGetTipoOggetto("RV_POIntervento"))
        If X = -1 Then
            X = 1
            ErroreCoda = True
        End If
    Loop Until X = 1
    
    If ErroreCoda = True Then
        X = -1
    End If
    
    If X = -1 Then
        Me.Enabled = True
        Me.SetFocus
        Me.Caption = Caption2Display
        Screen.MousePointer = 0
        ELIMINA_RIFERIMENTI_CODA fnGetTipoOggetto("RV_POIntervento")
        Exit Sub
    End If

    
    OLD_Cursor = Cn.CursorLocation
    Cn.CursorLocation = adUseClient
    
    
    frmAttesa.Show
    Me.Enabled = False
    
    DoEvents
    
    Me.Caption = "SALVATAGGIO IN CORSO..................."
    DoEvents
    
    frmAttesa.lblInfo = Me.Caption
    DoEvents
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''PROCEDURA DI INSERIMENTO AUTOMATICO DEI SERVIZI E LORO CALCOLO''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POContrattoServizi "
sSQL = sSQL & "WHERE IDRV_POContratto=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF

    ELABORAZIONE_INTERVENTI_PER_SERVIZIO fnNotNullN(rs!IDRV_POContrattoServizi), Me.CDCliente.KeyFieldID, Me.CDTecnico.KeyFieldID, fnNotNullN(m_Document(m_Document.PrimaryKey).Value), _
    Me.txtIDContrattoPadre.Value, fnNotNullN(rs!IDArticolo), fnNotNullN(rs!IDRV_POCriterioRicorrenza), fnNotNullN(rs!OgniNumeroGiorni), fnNotNullN(rs!OgniNumeroMesi), _
    fnNotNullN(rs!OgniNumeroSettimane), fnNotNullN(rs!IDRV_POTipoDataInizioRicorrenza), fnNotNullN(rs!GiornoInizioRicorrenza), fnNotNullN(rs!MeseInizioRicorrenza), _
    fnNotNullN(rs!IDRV_POTipoDataFineRicorrenza), fnNotNullN(rs!GiornoFineRicorrenza), fnNotNullN(rs!MeseFineRicorrenza), fnNotNullN(rs!NumeroRicorrenze), _
    LINK_TIPO_ANA_TEC_INT, LINK_STATO_INT_NUOVO, LINK_STATO_FASE_NUOVA, LINK_TIPO_FASE_ELA, GET_DESCRIZIONE_ARTICOLO(fnNotNullN(rs!IDArticolo)), LINK_TIPO_ANA_TEC_FASE, 0, IDClienteFatturazione
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Unload frmAttesa
Me.Enabled = True
Me.SetFocus
Me.Caption = Caption2Display


ELIMINA_RIFERIMENTI_CODA fnGetTipoOggetto("RV_POIntervento")

Screen.MousePointer = 0
GET_GRIGLIA_INTERVENTI
SSTab1.Tab = 2

Exit Sub

ERR_cmdGeneraInterventi_Click:

    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    
    MsgBox Err.Description, vbCritical, "OnSave"

    Cn.RollbackTrans
    ELIMINA_RIFERIMENTI_CODA fnGetTipoOggetto("RV_POIntervento")
    Cn.CursorLocation = OLD_Cursor
    
    Me.Caption = Caption2Display(False)

End Sub



Private Sub cmdGeneraIntSing_Click()
On Error GoTo ERR_cmdGeneraInterventi_Click
'On Error GoTo ERR_cmdGeneraInterventi_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim TestoMessaggio As String
Dim AvviaProcedura As Boolean
Dim rsInterventoTesta As ADODB.Recordset
Dim X As Long
Dim ErroreCoda As Boolean
Dim OLD_Cursor As Long


If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di procedere ad inserire i servizi legati al tipo di contratto", vbInformation, "Salvataggio documento"
    Exit Sub
End If
If fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare la riga del servizio prima di procedere alla generazione degli interventi", vbInformation, "Salvataggio documento"
    Exit Sub
End If

If MODULO_ATTIVATO_INT = 0 Then
    If Len(MODULO_DESCRIZIONE_INT) > 0 Then
        MsgBox "Il modulo " & MODULO_DESCRIZIONE_INT & " non è stato abilitato", vbInformation, TheApp.FunctionName
    Else
        MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
    End If
Exit Sub
End If


Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
Link_Contratto_Servizio = fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value)
Elabora_Tutti_Servizi_Contratto = 0

frmInterventiDaServ.Show vbModal

Me.txtNProdAss.Value = GET_NUMERO_PRODOTTI_PER_SERVIZIO(fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))
Me.txtNInterventiServ.Value = GET_NUMERO_INTERVENTI_PER_SERVIZIO(fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))

GET_GRIGLIA_INTERVENTI

Exit Sub
ERR_cmdGeneraInterventi_Click:
    MsgBox Err.Description, vbCritical, "ERR_cmdGeneraInterventi_Click"
End Sub

Private Sub cmdGeneraRate_Click()
On Error GoTo ERR_cmdGeneraRate_Click
Dim Testo As String

    If NON_MODIFICA_CONTRATTO = 1 Then Exit Sub

    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    If Me.cboTipoImpostazione.CurrentID = 3 Then Exit Sub
    
    If (m_Changed = True) Then
        MsgBox "Salvare il documento prima di procedere all'operazione", vbInformation, "Generazione rate"
        Exit Sub
    End If
    
    
    If Not ((Me.GrigliaRateContratto.Recordset.EOF) And (Me.GrigliaRateContratto.Recordset.BOF)) Then
        Testo = "Sei sicuro di voler procedere con questo comando?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then Exit Sub
    End If
    
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Una o più rate del contratto risultano fatturate, pertanto continuando con questo comando verranno eseguite le seguenti operazioni: " & vbCrLf
    Testo = Testo & "1. Eliminazione di tutte le rate che non risultano fatturate" & vbCrLf
    Testo = Testo & "2. Generazione delle nuove rate con data e numerazione successivva all'ultima rata fatturata" & vbCrLf
    Testo = Testo & "3. L'ultima rata del contratto sarà calcolata per differenza tra il totale del contratto e il totale delle rate precedentemente elaborate" & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    
    If GET_RATA_PAGATA(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) = True Then
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then Exit Sub
    End If
        
    Screen.MousePointer = 11
    
    frmAttesa.Show
    Me.Enabled = False
    DoEvents
    
    frmAttesa.lblInfo.Caption = "SVILUPPO RATE DEL CONTRATTO..."
    DoEvents
    
    If Me.cboTipoImpostazione.CurrentID = 2 Then
        SviluppoRateContratto fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    End If
    
    If Me.cboTipoImpostazione.CurrentID = 1 Then
        'If ControlloCambioTesta = True Then
            SviluppoRateContratto fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        'End If
    End If
    
    Screen.MousePointer = 0
    
    Unload frmAttesa
    Me.Enabled = True
    
    
    m_DocumentsLink.Refresh
    
    SSTab1.Tab = 0
    
Exit Sub
ERR_cmdGeneraRate_Click:
    Unload frmAttesa
    Me.Enabled = True
    
    MsgBox Err.Description, vbCritical, "ERR_cmdGeneraRate_Click"
    
    Screen.MousePointer = 0
End Sub

Private Sub cmdGeneraScadenzeProd_Click()
On Error GoTo ERR_cmdGeneraScadenzeProd_Click
    If Me.cboTipoImpostazione.CurrentID <= 3 Then Exit Sub
    
    If Me.chkGeneraAccontiSaldo.Value = vbChecked Then Exit Sub
    
    SviluppoRateContrattoProdotto fnNotNullN(m_Document(m_Document.PrimaryKey).Value), fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)
    
    
Exit Sub
ERR_cmdGeneraScadenzeProd_Click:
    MsgBox Err.Description, vbCritical, "cmdGeneraScadenzeProd_Click"
End Sub

Private Sub cmdgenIntDaProd_Click()
On Error GoTo ERR_cmdgenIntDaProd_Click
    If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub
    If (Me.chkOfferta.Value = vbChecked) Then Exit Sub
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey)) <= 0 Then
        MsgBox "Salvare il documento prima di procedere ad un inserimento di una rata manuale", vbCritical, "Nuova riga documento"
        Exit Sub
    End If
    
    If fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0 Then Exit Sub

    If fnNotNullN(m_DocumentsLink3("IDRV_POProdotto").Value) = 0 Then Exit Sub
    
    Link_Contratto_Prodotto = fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)
    Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    Link_Contratto_padre = fnNotNullN(m_Document("IDRV_POContrattoPadre").Value)
    Note_Prodotto_Intervento = Me.txtNoteProdInt.Text
    
    frmGeneraIntDaProd.Show vbModal
    
    Me.txtNoteProdInt.Text = Note_Prodotto_Intervento
    
    FORM_GEN_INT_DA_PROD = 1
    
    cmdSalvaProd_Click
    
    FORM_GEN_INT_DA_PROD = 0
    
    GET_GRIGLIA_INTERVENTI
    
Exit Sub
ERR_cmdgenIntDaProd_Click:
    MsgBox Err.Description, vbCritical, "cmdgenIntDaProd_Click"
End Sub

Private Sub cmdInfoAggArtProd_Click()
    frmNoteArtProd.Show vbModal
    
End Sub

Private Sub cmdNuovaRata_Click()
On Error GoTo ERR_cmdNuovaRata_Click
    If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub

    If fnNotNullN(m_Document(m_Document.PrimaryKey)) <= 0 Then
        MsgBox "Salvare il documento prima di procedere ad un inserimento di una rata manuale", vbCritical, "Nuova riga documento"
        Exit Sub
    End If
    
    If m_DocumentsLink.TableNew Then
        m_DocumentsLink.AbortNewRow
    End If
    
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink.NewRow
    
    NuovaRata = 1
    
    Me.txtNumeroRata.Value = GET_NUMERO_RATA(fnNotNullN(m_Document(m_Document.PrimaryKey).Value))
    
    Me.cboPagamentoRataContratto.WriteOn Me.cboPagamentoRate.CurrentID
    
    Me.txtDataRata.SetFocus
Exit Sub
ERR_cmdNuovaRata_Click:
    MsgBox Err.Description, vbCritical, "cmdNuovaRata_Click"
End Sub
Private Function fncListinoDefault() As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "Select IDListinoDiBase From ConfigurazioneVendite Where IDAzienda=" & m_App.IDFirm
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fncListinoDefault = rs!IDListinoDiBase
    
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Sub cmdNuovoAdeg_Click()
On Error GoTo ERR_cmdNuovoAdeg_Click
    If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey)) <= 0 Then
        MsgBox "Salvare il documento prima di procedere ad un inserimento di un servizio legato al tipo di contratto", vbCritical, "Nuova riga documento"
        Exit Sub
    End If

    If Me.chkContrattoAttuale.Value = vbUnchecked Then Exit Sub


    If m_DocumentsLink2.TableNew Then
        m_DocumentsLink2.AbortNewRow
    End If
    
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink2.NewRow
    
    Me.chkIstatAdeguamento.Value = Me.chkAdeguamentoIstat.Value
    Me.chkAdegContrProx.Value = Me.chkRinnovoAutomatico.Value
    Me.chkAdeguaContrAttuale.Value = vbChecked
    Me.txtNumeroAdeguamento.Value = GET_NUMERO_ADEGUAMENTO(Me.txtIDContrattoPadre.Value)
    Me.txtDataStipulaAdeg.SetFocus
Exit Sub
ERR_cmdNuovoAdeg_Click:
    MsgBox Err.Description, vbCritical, "cmdNuovoAdeg_Click"
End Sub

Private Sub cmdNuovoProd_Click()
On Error GoTo ERR_cmdNuovoProd_Click
    
    If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey)) <= 0 Then
        MsgBox "Salvare il documento prima di procedere ad un inserimento", vbCritical, "Nuova riga documento"
        Exit Sub
    End If
    
    If Me.chkContrattoAttuale.Value = vbUnchecked Then Exit Sub
    
    If m_DocumentsLink3.TableNew Then
        m_DocumentsLink3.AbortNewRow
    End If
    
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink3.NewRow
    
    
    Me.txtQtaArtProd.Value = 1
    
    If COPIA_NUOVA_RIGA_PROD = 0 Then cmdSelServizio_Click

Exit Sub
ERR_cmdNuovoProd_Click:
    MsgBox Err.Description, vbCritical, "cmdNuovoProd_Click"
End Sub

Private Sub cmdNuovoProdDaEsis_Click()
On Error GoTo ERR_cmdNuovoProdDaEsis_Click
Dim IDRigaDaCopiare As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
If fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0 Then Exit Sub

IDRigaDaCopiare = fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)

If IDRigaDaCopiare <= 0 Then Exit Sub

COPIA_NUOVA_RIGA_PROD = 1

cmdNuovoProd_Click

sSQL = "SELECT * FROM RV_POContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDRigaDaCopiare

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtIDProdotto.Value = fnNotNullN(rs!IDRV_POProdotto)
    Me.txtDescrProdotto.Text = fnNotNull(rs!DescrizioneAggiuntiva)
    Me.CDArticoloProd.Load fnNotNullN(rs!IDArticolo)
    Me.cboTipoPeriodo.WriteOn fnNotNullN(rs!IDRV_POTipoPeriodo)
    Me.cboUMPeriodoProd.WriteOn fnNotNullN(rs!IDRV_POUnitaDiMisuraPeriodo)
    Me.txtDataInizioProd.Value = fnNotNullN(rs!DataFinePeriodo) + 1
    Me.txtQtaPeriodo.Value = fnNotNullN(rs!QuantitaPeriodo)
    Me.cboListinoProd.WriteOn fnNotNullN(rs!IDListino)
    Me.cboTipoRateizzazioneProd.WriteOn fnNotNullN(rs!IDRateizzazione)
    Me.cboUMArtProd.WriteOn fnNotNullN(rs!IDUnitaDiMisuraArticolo)
    Me.cboIvaProd.WriteOn fnNotNullN(rs!IDIva)
    Me.txtQtaArtProd.Value = fnNotNullN(rs!QuantitaArticolo)
    Me.txtImpUniProd.Value = fnNotNullN(rs!ImportoUnitario)
    Me.txtSconto1Prod.Value = fnNotNullN(rs!Sconto1)
    Me.txtSconto2Prod.Value = fnNotNullN(rs!Sconto2)
    Me.txtScontoImpProd.Value = fnNotNullN(rs!ScontoAImporto)
    
    Me.chkConducente.Value = Abs(fnNotNullN(rs!Conducente))
    Me.chkEscludiGiorniFestivi.Value = Abs(fnNotNullN(rs!EscludiGiorniFestivi))
    Me.chkEscludiSabato.Value = Abs(fnNotNullN(rs!EscludiSabato))
    Me.chkACorpo.Value = Abs(fnNotNullN(rs!ACorpo))
    Me.cboAnaOperatoreProd.WriteOn fnNotNullN(rs!IDAnagraficaOperatore)
    Me.chkRinnovareProd.Value = Abs(fnNotNullN(rs!NonRinnovare))
    
    Me.txtNoteProdotto.Text = fnNotNull(rs!AnnotazioniPerIntervento)
    
    txtQtaPeriodo_LostFocus
    Me.txtQuantitaEffettiva.Value = Me.txtQtaPeriodo.Value
    
    If Me.cboUMPeriodoProd.CurrentID = 2 Then
        Me.txtQuantitaEffettiva.Value = GET_CALCOLO_QUANTITA_EFFETTIVA(Me.txtDataInizioProd.Text, Me.txtDataFineProd.Text)
        
    End If
    
    GET_TOTALI_RIGA_DETTAGLIO
    
    
End If

rs.CloseResultset
Set rs = Nothing

COPIA_NUOVA_RIGA_PROD = 0
Exit Sub
ERR_cmdNuovoProdDaEsis_Click:
    MsgBox Err.Description, vbCritical, "ERR_cmdNuovoProdDaEsis_Click"
    COPIA_NUOVA_RIGA_PROD = 0
End Sub

Private Sub cmdNuovoProdotto_Click()
On Error GoTo ERR_cmdNuovoProdotto_Click
    If fnNotNullN(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) <= 0 Then Exit Sub
    If Me.txtIDProdotto.Value = 0 Then Exit Sub
    
    LabelLink2.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POProdotto"))
    Me.LabelLink2.IDReturn = Me.txtIDProdotto.Value
    Me.LabelLink2.RunApplication
Exit Sub
ERR_cmdNuovoProdotto_Click:
    MsgBox Err.Description, vbCritical, "cmdNuovoProdotto_Click"
End Sub

Private Sub cmdNuovoServizio_Click()
    
    If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey)) <= 0 Then
        MsgBox "Salvare il documento prima di procedere ad un inserimento di un servizio legato al tipo di contratto", vbCritical, "Nuova riga documento"
        Exit Sub
    End If
        
    If Me.chkContrattoAttuale.Value = vbUnchecked Then Exit Sub
        
    If m_DocumentsLink1.TableNew Then
        m_DocumentsLink1.AbortNewRow
    End If
    
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink1.NewRow
    

    
    Me.CDServizio.SetFocus
End Sub

Private Sub cmdProdottiServizio_Click()
    
    If fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value) <= 0 Then Exit Sub
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    Link_Contratto_Servizio = fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value)
    Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    FILTRO_PRODOTTO_ASSOCIATO = 0
    
    frmSelProdottiServizi.Show vbModal

    Me.txtNProdAss.Value = GET_NUMERO_PRODOTTI_PER_SERVIZIO(fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))
    Me.txtNInterventiServ.Value = GET_NUMERO_INTERVENTI_PER_SERVIZIO(fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))
    
End Sub

Private Sub cmdRateStorico_Click()
On Error GoTo ERR_cmdRateStorico_Click

    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    frmRateStorico.Show vbModal

Exit Sub
ERR_cmdRateStorico_Click:
MsgBox Err.Description, vbCritical, "cmdRateStorico_Click"
End Sub

Private Sub cmdSaldo_Click()
    
    If NON_MODIFICA_CONTRATTO = 1 Then Exit Sub
    
    If (m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    If Me.txtImportoTotAdeg.Value <= 0 Then Exit Sub
    
    Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    LINK_OGGETTO_CONTRATTO_SEL = fnNotNullN(m_Document("IDOggetto").Value)
    LINK_TIPO_OGGETTO_CONTRATTO_SEL = m_DocType.ID
    OPERAZIONE_ESEGUITA_ACCONTO = 0
    
    frmSaldo.Show vbModal
    
    'FUNZIONE CHE INSERISCI GLI ACCONTI IN GRIGLIA
       
    If OPERAZIONE_ESEGUITA_ACCONTO = 1 Then
        Me.chkChiuso.Value = vbChecked
        OnSave
    End If
    
End Sub

Private Sub cmdSalvaAdeg_Click()
On Error GoTo ERR_cmdSalvaServizio_Click
Dim Attiva_Crea_Scadenza As Long
Dim Testo As String

    
    If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub
    
    Attiva_Crea_Scadenza = 1
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        MsgBox "Salvare il documento prima di procedere ad inserire gli adeguamente contrattuali", vbInformation, "Salvataggio documento"
        Exit Sub
    End If
    
'    If Me.chkContrattoAttuale.Value = vbUnchecked Then
'        MsgBox ""
'        Exit Sub
'    End If
    
    If Me.txtDataStipulaAdeg.Value = 0 Then
        MsgBox "Inserire la data di stipula dell'adeguamento", vbInformation, "Permesso salvataggio"
        Exit Sub
    End If
    
    If Me.txtDataDecorrenzaAdeg.Value = 0 Then
        MsgBox "Inserire la data di decorrenza dell'adeguamento", vbInformation, "Permesso salvataggio"
        Exit Sub
    End If
    
    If Me.txtImportoAdeg.Value = 0 Then
        MsgBox "Inserire l'importo dell'adeguamento", vbInformation, "Permesso salvataggio"
        Exit Sub
    End If

    If Me.chkAdegContrProx.Value = vbChecked Then
        If Me.cboTipoAdeguamento.CurrentID = 0 Then
            MsgBox "Inserire il tipo di adeguamento per il prossimo rinnovo", vbInformation, "Permesso salvataggio"
            Exit Sub
        End If
    End If
    
    If Me.txtDataScadenzaAdeg.Value = 0 Then
        If Me.txtDataDecorrenzaAdeg.Value > Me.txtDataScadenzaPerRinnovo.Value + 1 Then
            MsgBox "La data di decorrenza dell'adeguamento deve essere minore o uguale alla data di scadenza per il rinnovo del contratto", vbInformation, "Permesso salvataggio"
            Exit Sub
        End If
    
        If Me.txtDataDecorrenzaAdeg.Value < Me.txtDataDecorrenza.Value Then
            MsgBox "La data di decorrenza dell'adeguamento deve essere maggiore o uguale alla data di deccorenza del contratto", vbInformation, "Permesso salvataggio"
            Exit Sub
        End If
    End If
    
    If ((Me.txtDataScadenzaAdeg.Value > 0) And (Me.cboTipoRateizzazioneAdeg.CurrentID = 0)) Then
        MsgBox "Se si inserisce la data di scadenza dell'adeguamento, bisogna inserire il tipo di rateizzazione", vbInformation, "Permesso salvataggio"
        Exit Sub
    End If
    
    If ((Me.txtDataScadenzaAdeg.Value = 0) And (Me.cboTipoRateizzazioneAdeg.CurrentID > 0)) Then
        MsgBox "Se si inserisce il tipo di rateizzazione dell'adeguamento, bisogna inserire la data di scadenza", vbInformation, "Permesso salvataggio"
        Exit Sub
    End If
    
    If ((Me.txtDataScadenzaAdeg.Value > 0) And (Me.cboTipoRateizzazioneAdeg.CurrentID > 0)) Then
        If (Me.txtDataScadenzaAdeg.Value < Me.txtDataDecorrenzaAdeg.Value) Then
            MsgBox "La data di decorrenza non può essere maggiore della data scadenza dell'adeguamento ", vbInformation, "Permesso salvataggio"
            Exit Sub
        End If
    End If
    
    If fnNotNullN(m_DocumentsLink2(m_DocumentsLink2.PrimaryKey).Value) > 0 Then
        If GET_RATA_PAGATA_ADEGUAMENTO(fnNotNullN(m_DocumentsLink2(m_DocumentsLink2.PrimaryKey).Value), fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) = True Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "Una o più rate dell'adeguamento risultano fatturate pertanto se si continua nell'operazione non verranno generate altre rate e ci potrebbe essere un disiallineamento " & vbCrLf
            Testo = Testo & "Vuoi continuare?"
            
            If MsgBox(Testo, vbQuestion + vbYesNo, "Permesso salvataggio") = vbNo Then Exit Sub
            
            Attiva_Crea_Scadenza = 0

        End If
    End If
        
    m_DocumentsLink2("DataStipula").Value = Me.txtDataStipulaAdeg.Value
    m_DocumentsLink2("DataDecorrenza").Value = Me.txtDataDecorrenzaAdeg.Value
    m_DocumentsLink2("Importo").Value = Me.txtImportoAdeg.Value
    m_DocumentsLink2("Annotazioni").Value = Me.txtAnnotazioniAdeg.Text
    m_DocumentsLink2("RiportaProssimoRinnovo").Value = Me.chkAdegContrProx.Value
    m_DocumentsLink2("AdeguaContrattoAttuale").Value = Me.chkAdeguaContrAttuale.Value
    m_DocumentsLink2("IDRV_POContrattoPadre").Value = Me.txtIDContrattoPadre.Value
    
    m_DocumentsLink2("IDRV_POTipoAdeguamento").Value = Me.cboTipoAdeguamento.CurrentID
    m_DocumentsLink2("IDArticolo").Value = Me.CDArticoloAdeg.KeyFieldID
    m_DocumentsLink2("NumeroProtocollo").Value = Me.txtProtAdeg.Text
    m_DocumentsLink2("AdeguamentoIstat").Value = Me.chkIstatAdeguamento.Value
    m_DocumentsLink2("NumeroAdeguamento").Value = Me.txtNumeroAdeguamento.Value
    m_DocumentsLink2("DescrizioneAdeguamento").Value = "Adeguamento numero " & Me.txtNumeroAdeguamento.Value & " - " & Me.txtNumeroRinnovo.Value
    m_DocumentsLink2("DescrizionePerFatturazione").Value = Me.txtDescrFattAde.Text
    m_DocumentsLink2("IDRateizzazione").Value = Me.cboTipoRateizzazioneAdeg.CurrentID
    m_DocumentsLink2("NoCalcPeriodoFatt").Value = chkNoCalcPeriodoFatt.Value
    
    m_DocumentsLink2("ImportoAlRinnovo").Value = txtImportoAdegRinn.Value
    m_DocumentsLink2("NumeroPartenza").Value = txtNAdegIniz.Value
    
    If Me.txtDataScadenzaAdeg.Value = 0 Then
        m_DocumentsLink2("DataFineAdeguamento").Value = Null
    Else
        m_DocumentsLink2("DataFineAdeguamento").Value = Me.txtDataScadenzaAdeg.Value
    End If
    
    If fnNotNullN(m_DocumentsLink2(m_DocumentsLink2.PrimaryKey).Value) <= 0 Then
        m_DocumentsLink2("DataInserimento").Value = Date
        m_DocumentsLink2("IDUtenteInserimento").Value = TheApp.IDUser
        m_DocumentsLink2("OraInserimento").Value = GET_ORARIO(Now)
        m_DocumentsLink2("PCInserimento").Value = GET_NOMECOMPUTER
        m_DocumentsLink2("UtentePCInserimento").Value = GET_NOMEUTENTE
    End If
    
    Screen.MousePointer = 11
    Cn.BeginTrans
    
    m_DocumentsLink2.Save
    
    Cn.CommitTrans

    m_DocumentsLink2.Move Me.GrigliaAdeg.ListIndex - 1
    
    
    If Attiva_Crea_Scadenza = 1 Then
        If Me.chkAdeguaContrAttuale.Value = vbChecked Then
            If ((Me.txtDataScadenzaAdeg.Value = 0) And (Me.cboTipoRateizzazioneAdeg.CurrentID = 0)) Then
                If Me.txtDataDecorrenzaAdeg.Value <> Me.txtDataDecorrenza.Value Then
                    GENERA_RATE_ADEGUAMENTO fnNotNullN(m_DocumentsLink2(m_DocumentsLink2.PrimaryKey).Value), Me.cboTipoRinnovo.CurrentID, _
                    fnNotNullN(m_Document(m_Document.PrimaryKey).Value), Me.txtIDContrattoPadre.Value, DatePart("yyyy", Me.txtDataDecorrenza.Text), _
                    Me.txtDataDecorrenzaAdeg.Text, Me.txtDataScadenzaPerRinnovo.Text, Me.txtImportoAdeg.Value, Me.cboPagamentoRate.CurrentID, _
                    Me.txtNumeroAdeguamento.Value, Me.txtProtAdeg.Text, Me.CDArticoloAdeg.KeyFieldID, Me.txtDescrFattAde.Text
                Else
                    SviluppoRateContratto fnNotNullN(m_Document(m_Document.PrimaryKey).Value), fnNotNullN(m_DocumentsLink2(m_DocumentsLink2.PrimaryKey).Value)
                End If
            Else
                SviluppoRateContratto fnNotNullN(m_Document(m_Document.PrimaryKey).Value), fnNotNullN(m_DocumentsLink2(m_DocumentsLink2.PrimaryKey).Value)
                
            End If
            m_DocumentsLink.Refresh
        End If
    End If
    Screen.MousePointer = 0
    Me.txtImportoTotAdeg.Value = GET_TOTALE_ADEGUAMENTI_DETTAGLIO(fnNotNullN(m_Document(m_Document.PrimaryKey).Value))

    Exit Sub
    
ERR_cmdSalvaServizio_Click:
    MsgBox Err.Description, vbCritical, "Salva adeguamenti"
    Cn.RollbackTrans
    Screen.MousePointer = 0
    
End Sub

Private Sub cmdSalvaProd_Click()
On Error GoTo ERR_cmdSalvaProd_Click
Dim Testo As String
Dim Dismesso As Boolean
Dim matricolaOLD As String

    If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub

    Dismesso = False

    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        MsgBox "Salvare il documento prima di procedere ad inserire i prodotti del contratto", vbInformation, "Salvataggio documento"
        Exit Sub
    End If
    
    If ((fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) > 0) And (Me.chkProdottoGenerico.Value = vbUnchecked)) Then
        If ((m_DocumentsLink3("Dismesso").Value = 1) And (Me.chkDismessoProd.Value = vbUnchecked)) Then
            If GET_CONTROLLO_PRODOTTO_ALTRO_CONTRATTO(Me.txtIDProdotto.Value, fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) = True Then
                Testo = "ATTENZIONE!!!" & vbCrLf
                Testo = Testo & "Il prodotto risulta associato ad un altro contratto " & vbCrLf
                Testo = Testo & "Impossibile continuare nell'operazione"
                MsgBox Testo, vbCritical, "Salvataggio dati"
                
                Exit Sub
            End If
        End If
    Else
        If Me.chkProdottoGenerico.Value = vbChecked Then
            Dismesso = True
        End If
    End If
    
    If Me.cboTipoImpostazione.CurrentID > 1 Then
        If ((fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) > 0)) Then
            
            If Me.CDArticoloProd.KeyFieldID = 0 Then
                MsgBox "Inserire l'articolo", vbInformation, "Controllo dati"
                Exit Sub
            End If
            
        End If
'
'        If fnNotNullN(m_DocumentsLink3("IDRV_POInterventoRigheDett").Value) > 0 Then
'            MsgBox "Impossibile modificare i dati ad una riga generata dalla gestione degli addebiti intervento", vbInformation, "Controllo dati"
'            Exit Sub
'        End If
'
'        If fnNotNullN(m_DocumentsLink3("IDRV_POContatoreRilevamenti").Value) > 0 Then
'            MsgBox "Impossibile modificare i dati ad una riga generata dalla gestione rilevamenti dei contatori", vbInformation, "Controllo dati"
'            Exit Sub
'        End If

        If fnNotNullN(m_DocumentsLink3("IDRV_POInterventoRigheDett").Value) > 0 Then
            If GET_CONTROLLO_ESISTENZA_ADD_PROD(fnNotNullN(m_DocumentsLink3("IDRV_POInterventoRigheDett").Value)) = True Then
                MsgBox "Impossibile eliminare una riga generata dalla gestione degli addebiti intervento", vbInformation, "Controllo dati"
                Exit Sub
            End If
        End If
        
        If fnNotNullN(m_DocumentsLink3("IDRV_POContatoreRilevamenti").Value) > 0 Then
            If GET_CONTROLLO_ESISTENZA_RIL_PROD(fnNotNullN(m_DocumentsLink3("IDRV_POContatoreRilevamenti").Value)) = True Then
                MsgBox "Impossibile eliminare una riga generata dalla gestione rilevamenti contatore", vbInformation, "Controllo dati"
                Exit Sub
            End If
        End If
        If Me.cboTipoImpostazione.CurrentID = 3 Then
            If GET_CONTROLLO_RIGA_PROD_FATT(fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)) = True Then
                Testo = "ATTENZIONE!!!" & vbCrLf
                Testo = Testo & "Una o più rate di questa riga del contratto risultano fatturate, pertanto è impossibile eliminare o modificare la riga selezionata " & vbCrLf
                Testo = Testo & "Impossibile continuare"
                MsgBox Testo, vbCritical, "Validazione dati"
                Exit Sub
            End If
        End If
    End If
    
    If (fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0) Then
        If (Me.txtIDProdotto.Value > 0) Then
            If Me.chkProdottoGenerico.Value = vbUnchecked Then
                If GET_ESISTENZA_PROD_CONTR(Me.txtIDProdotto.Value, Me.txtDataInizioProd.Text, Me.txtDataFineProd.Text) = True Then
                    Testo = "ATTENZIONE!!!" & vbCrLf
                    Testo = Testo & "Il prodotto risulta in un altro contratto nel periodo selezionato" & vbCrLf
                    Testo = Testo & "Vuoi continuare?"
                    
                    If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then Exit Sub
                    
                End If
            End If
        End If
    End If
    
    GET_TOTALI_RIGA_DETTAGLIO
    
    matricolaOLD = fnNotNull(m_DocumentsLink3("ValoreIndentificativo").Value)
    
    m_DocumentsLink3("ValoreIndentificativo").Value = Me.txtValIdentProd.Text
    m_DocumentsLink3("DescrizioneAggiuntiva").Value = Me.txtDescrProdotto.Text
    m_DocumentsLink3("Annotazioni").Value = Me.txtNoteProdotto.Text
    m_DocumentsLink3("Quantita").Value = Me.txtQtaProdotto.Value
    m_DocumentsLink3("Dismesso").Value = Me.chkDismessoProd.Value
    If Me.txtDataDismesso.Value > 0 Then m_DocumentsLink3("DataDismesso").Value = Me.txtDataDismesso.Value Else m_DocumentsLink3("DataDismesso").Value = Null
    If ((Me.chkDismessoProd.Value = vbUnchecked) And (txtDataDismesso.Value > 0)) Then
        m_DocumentsLink3("DataDismesso").Value = Null
    End If
    m_DocumentsLink3("IDRV_POProdotto").Value = Me.txtIDProdotto.Value
    m_DocumentsLink3("IDRV_POContrattoPadre").Value = Me.txtIDContrattoPadre.Value
    
    m_DocumentsLink3("IDArticolo").Value = Me.CDArticoloProd.KeyFieldID
    m_DocumentsLink3("IDRV_POTipoPeriodo").Value = Me.cboTipoPeriodo.CurrentID
    m_DocumentsLink3("IDRV_POUnitaDiMisuraPeriodo").Value = Me.cboUMPeriodoProd.CurrentID
    If (Me.txtDataInizioProd.Value) > 0 Then
        m_DocumentsLink3("DataInizioPeriodo").Value = Me.txtDataInizioProd.Value
    End If
    m_DocumentsLink3("OraInizioPeriodo").Value = Me.txtOraInizioProd.Text
    If (Me.txtDataFineProd.Value) > 0 Then
        m_DocumentsLink3("DataFinePeriodo").Value = Me.txtDataFineProd.Value
    End If
    m_DocumentsLink3("OraFinePeriodo").Value = Me.txtOraFineProd.Text
        
    m_DocumentsLink3("QuantitaPeriodo").Value = Me.txtQtaPeriodo.Value
        
    m_DocumentsLink3("IDListino").Value = Me.cboListinoProd.CurrentID
        
    m_DocumentsLink3("IDUnitaDiMisuraArticolo").Value = Me.cboUMArtProd.CurrentID
    m_DocumentsLink3("QuantitaArticolo").Value = Me.txtQtaArtProd.Value
        
    m_DocumentsLink3("ImportoUnitario").Value = Me.txtImpUniProd.Value
    m_DocumentsLink3("Sconto1").Value = Me.txtSconto1Prod.Value
    m_DocumentsLink3("Sconto2").Value = Me.txtSconto2Prod.Value
    m_DocumentsLink3("Imponibile").Value = Me.txtImponibileProd.Value
    m_DocumentsLink3("ScontoAImporto").Value = Me.txtScontoImpProd.Value
    m_DocumentsLink3("TotaleRiga").Value = Me.txtTotaleRigaProd.Value
    
    m_DocumentsLink3("IDIva").Value = Me.cboIvaProd.CurrentID
    m_DocumentsLink3("AliquotaIva").Value = Me.txtAliquotaIvaProd.Value
    m_DocumentsLink3("ImportoIva").Value = Me.txtImportoIvaProd.Value
    
    m_DocumentsLink3("QuantitaEffettiva").Value = Me.txtQuantitaEffettiva.Value
    m_DocumentsLink3("EscludiGiorniFestivi").Value = Me.chkEscludiGiorniFestivi.Value
    m_DocumentsLink3("EscludiSabato").Value = Me.chkEscludiSabato.Value
    m_DocumentsLink3("Conducente").Value = Me.chkConducente.Value
    m_DocumentsLink3("ACorpo").Value = Me.chkACorpo.Value
    m_DocumentsLink3("IDAnagraficaOperatore").Value = Me.cboAnaOperatoreProd.CurrentID
    m_DocumentsLink3("IDRateizzazione").Value = Me.cboTipoRateizzazioneProd.CurrentID
    m_DocumentsLink3("NonRinnovare").Value = Me.chkRinnovareProd.Value
    m_DocumentsLink3("NonRateizzare").Value = Me.chkGeneraUnaRataProd.Value
    
'    If Len(Trim(Me.txtNoteProdotto.Text)) = 0 Then
'        If (Me.txtIDProdotto.Value > 0) Then
'            Me.txtNoteProdotto.Text = "Noleggio dal " & Me.txtDataInizioProd.Text
'            If (Me.txtOraInizioProd.Value > 0) Then
'                Me.txtNoteProdotto.Text = Me.txtNoteProdotto.Text & " " & Me.txtOraInizioProd.Text
'            End If
'            Me.txtNoteProdotto.Text = Me.txtNoteProdotto.Text & " al " & Me.txtDataFineProd.Text
'            If (Me.txtOraFineProd.Value > 0) Then
'                Me.txtNoteProdotto.Text = Me.txtNoteProdotto.Text & " " & Me.txtOraFineProd.Text
'            End If
'
'            If ((Me.chkEscludiGiorniFestivi.Value = vbChecked) And (Me.chkEscludiSabato.Value = vbChecked)) Then
'                Me.txtNoteProdotto.Text = Me.txtNoteProdotto.Text & " (esclusi i sabati, le domeniche e i giorni festivi) "
'            Else
'                If ((Me.chkEscludiGiorniFestivi.Value = vbChecked)) Then
'                    Me.txtNoteProdotto.Text = Me.txtNoteProdotto.Text & " (escluse le domeniche e i giorni festivi) "
'                End If
'
'                If ((Me.chkEscludiSabato.Value = vbChecked)) Then
'                    Me.txtNoteProdotto.Text = Me.txtNoteProdotto.Text & " (esclusi i sabati) "
'                End If
'            End If
'            m_DocumentsLink3("Annotazioni").Value = Me.txtNoteProdotto.Text
'        End If
'    End If
'
    If Me.cboTipoImpostazione.CurrentID = 3 Then
        If Me.chkGeneraAccontiSaldo.Value = vbUnchecked Then
            m_DocumentsLink3("ImportoComplessivo").Value = Me.txtImponibileProd.Value
        Else
            m_DocumentsLink3("ImportoComplessivo").Value = Me.txtTotaleRigaProd.Value
        End If
    End If
    
    If Me.cboTipoImpostazione.CurrentID = 2 Then
        m_DocumentsLink3("ImportoComplessivo").Value = Me.txtImponibileProd.Value
    End If
    
    m_DocumentsLink3("AnnotazioniPerIntervento").Value = Me.txtNoteProdInt.Text
    
    Screen.MousePointer = 11
    Cn.BeginTrans
    
    m_DocumentsLink3.Save
    
    Cn.CommitTrans
    
    m_DocumentsLink3.Move Me.GrigliaProd.ListIndex - 1
    
    Screen.MousePointer = 0
    
    If m_DocumentsLink3("Dismesso").Value = 0 Then
        AGGIORNA_INTERVENTI_PRODOTTO m_DocumentsLink3("IDRV_POProdotto").Value, m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value
    Else
        AGGIORNA_INTERVENTI_PRODOTTO_PER_ELIMINAZIONE m_DocumentsLink3("IDRV_POProdotto").Value, m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value
        ELIMINAZIONE_ASSOCIAZIONE_PRODOTTO_SERVIZIO fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)
    End If
    
    If (Me.cboTipoImpostazione.CurrentID = 3) Then
        If Me.chkGeneraAccontiSaldo.Value = vbUnchecked Then
            If FORM_GEN_INT_DA_PROD = 0 Then
                SviluppoRateContrattoProdotto fnNotNullN(m_Document(m_Document.PrimaryKey).Value), fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)
            End If
        End If
    End If
    If (Len(matricolaOLD) > 0) Then
        If (matricolaOLD <> Me.txtValIdentProd.Text) Then
            AGGIORNA_DESCR_RATE_PROD fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value), matricolaOLD, Me.txtValIdentProd.Text
        End If
    End If
    m_DocumentsLink1.Refresh
    m_DocumentsLink.Refresh
    
    GET_GRIGLIA_INTERVENTI
    
    If Me.cboTipoImpostazione.CurrentID <> 1 Then
        Me.txtImportoAttuale.Value = GET_TOTALE_CONTRATTO(m_Document(m_Document.PrimaryKey).Value)
        OnSave
    End If
    
    GET_TOTALE_PRODOTTI fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    If (Me.chkOfferta.Value = vbUnchecked) Then
        If VIS_FORM_GEN_INT_DA_PROD = 1 Then
            If FORM_GEN_INT_DA_PROD = 0 Then
                cmdgenIntDaProd_Click
            End If
        End If
    End If
    Exit Sub
    
ERR_cmdSalvaProd_Click:
    MsgBox Err.Description, vbCritical, "Salva prodotti"
    Cn.RollbackTrans
    Screen.MousePointer = 0

End Sub

Private Sub cmdSalvaRata_Click()
'On Error GoTo ERR_cmdSalvaRata_Click
Dim I As Long
Dim numerorata As Long
Dim OLDCursor As Long
Dim Testo As String
Dim IDTipoOggettoRata As Long
Dim IDOggettoRata As Long
Dim IDOggettoVend As Long
Dim IDTipoOggettoVend As Long
Dim IDOggettoScadenza As Long

    If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub

    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        MsgBox "Salvare il documento prima di procedere ad inserire le rate manualmente", vbInformation, "Salvataggio documento"
        Exit Sub
    End If
    
    If Me.chkContrattoAttuale.Value = 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "La rata del contratto che si sta tentando di modificare non è del contratto attuale" & vbCrLf
        Testo = Testo & "Vuoi continuare?"
        
        'MsgBox "Impossibile modificare le rate poichè non risulta essere il contratto attuale", vbInformation, TheApp.FunctionName
        If MsgBox(Testo, vbQuestion + vbYesNo, "Salvataggio rate del documento") = vbNo Then
            Exit Sub
        End If
    End If
    
    If Me.txtDataRata.Value = 0 Then
        MsgBox "Inserire la data della rata del contratto", vbInformation, TheApp.FunctionName
        Me.txtDataRata.SetFocus
        Exit Sub
    End If
    
    If Me.cboPagamentoRataContratto.CurrentID = 0 Then
        MsgBox "Inserire la modalità del pagamento della rata del contratto", vbInformation, TheApp.FunctionName
        Me.cboPagamentoRataContratto.SetFocus
        Exit Sub
    End If
    
    If Me.chkRataFatturata.Value = vbChecked Then
        If Me.txtIDOggettoCollegato.Value = 0 Then
            Testo = "La rata risulta fatturata, ma il collegamento al documento non è presente" & vbCrLf
            Testo = Testo & "Vuoi continuare?"
            If MsgBox(Testo, vbQuestion + vbYesNo, TheApp.FunctionName) = vbNo Then Exit Sub
        End If
    Else
        If Me.txtIDOggettoCollegato.Value > 0 Then
            Testo = "La rata risulta collegata, ma il flag 'Rata fatturata' risulta non vistato" & vbCrLf
            Testo = Testo & "Vuoi continuare?"
            If MsgBox(Testo, vbQuestion + vbYesNo, TheApp.FunctionName) = vbNo Then Exit Sub
        End If
        
    End If
        
    'If Me.txtIDOggettoCollegato.Value > 0 Then
    '    Testo = "Impossibile salvare la rata poichè risulta collegata ad un documento di vendita"
    '    MsgBox Testo, vbCritical, TheApp.FunctionName
    '    Exit Sub
    'End If
    
    If Me.chkNonFatturareRata.Value = vbChecked Then
        If Len(Trim(Me.txtNoteNonFattRata.Text)) = 0 Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "Inserire una descrizione del perchè questa rata non deve essere fatturata"
            MsgBox Testo, vbCritical, TheApp.FunctionName
            Exit Sub
        End If
    End If

    IDOggettoVend = fnNotNullN(m_DocumentsLink("IDOggettoCollegato").Value)
    IDTipoOggettoVend = fncIDDocumentoAllegato(IDOggettoVend)
    
    Me.txtIDTipoOggettoRata.Value = fnGetTipoOggetto("RV_PORateContratto")
    Me.txtIDOggettoRata.Value = GET_LINK_OGGETTO(fnNotNullN(m_DocumentsLink("IDOggetto").Value), Me.txtIDTipoOggettoRata.Value, Me.txtNumeroRata.Value, Me.txtDataRata.Text)

    m_DocumentsLink("DataRata") = Me.txtDataRata.Value
    m_DocumentsLink("ImportoRata") = Me.txtImportoRata.Value
    m_DocumentsLink("IDPagamentoRata") = Me.cboPagamentoRataContratto.CurrentID
    m_DocumentsLink("Fatturata") = fnNormBoolean(Me.chkRataFatturata.Value)
    m_DocumentsLink("IDOggettoCollegato") = Me.txtIDOggettoCollegato.Value
    m_DocumentsLink("Mese") = DatePart("m", Me.txtDataRata.Text)
    m_DocumentsLink("Anno") = DatePart("yyyy", Me.txtDataRata.Text)
    
    If Me.txtDataInizioPer.Value > 0 Then
        m_DocumentsLink("DataInizioPeriodo") = Me.txtDataInizioPer.Value
    End If
    If Me.txtDataFinePer.Value > 0 Then
        m_DocumentsLink("DataFinePeriodo") = Me.txtDataFinePer.Value
    End If
    
    If NuovaRata = 1 Then
        numerorata = numerorate(m_Document("IDRV_POContratto"))
        m_DocumentsLink("NumeroRata") = numerorate(m_Document("IDRV_POContratto"))
        m_DocumentsLink("Manuale") = True
        m_DocumentsLink("ContrattoAttuale") = True
        m_DocumentsLink("Adeguamento") = False
    Else
        m_DocumentsLink("Manuale") = Me.GrigliaRateContratto.AllColumns("Manuale").Value
        m_DocumentsLink("Adeguamento") = Me.GrigliaRateContratto.AllColumns("Adeguamento").Value
        m_DocumentsLink("ContrattoAttuale") = Me.GrigliaRateContratto.AllColumns("ContrattoAttuale").Value
    End If
    

    m_DocumentsLink("IDTipoOggetto").Value = Me.txtIDTipoOggettoRata.Value
    m_DocumentsLink("IDOggetto").Value = Me.txtIDOggettoRata.Value
    m_DocumentsLink("NonFatturare").Value = Me.chkNonFatturareRata.Value
    m_DocumentsLink("AnnotazioniNonFatturare").Value = Me.txtNoteNonFattRata.Text
    m_DocumentsLink("IDRV_POProdotto").Value = Me.txtIDProdottoRata.Value
    m_DocumentsLink("IDRV_POContrattoProdotti").Value = Me.txtIDProdRifContr.Value
    
    
    If fnNotNullN(m_DocumentsLink("IDOggetto").Value) = 0 Then
        MsgBox "Errore flusso documentale", vbCritical, TheApp.FunctionName
        Exit Sub
    End If
        
    Screen.MousePointer = 11
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
    'Me.lblInfoSalvaRiga.Caption = "SALVATAGGIO RIGA IN CORSO........."
    DoEvents
    
    Cn.BeginTrans
    
    m_DocumentsLink.Save
    
    Cn.CommitTrans


    If fnNotNullN(m_DocumentsLink("IDRV_POContrattoAdeguamento").Value) = 0 Then
        If NO_CALCOLO_PERIODO_FATT = 0 Then
            If Len(Trim(Me.txtPeriodo.Text)) > 0 Then
                m_DocumentsLink("Periodo") = Me.txtPeriodo.Text
            Else
                m_DocumentsLink("Periodo") = GET_STRINGA_PERIODO_ADEG(1, m_App.Branch, fnNotNullN(m_Document(m_Document.PrimaryKey).Value), fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value), fnNotNullN(m_DocumentsLink("IDRV_POContrattoAdeguamento").Value), Me.txtIDProdotto.Value)
            End If
        Else
            If fnNotNullN(m_DocumentsLink("IDRV_POProdotto").Value) > 0 Then
                If Len(Trim(Me.txtPeriodo.Text)) > 0 Then
                    m_DocumentsLink("Periodo") = Me.txtPeriodo.Text
                Else
                    m_DocumentsLink("Periodo") = GET_STRINGA_PERIODO_ADEG(1, m_App.Branch, fnNotNullN(m_Document(m_Document.PrimaryKey).Value), fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value), fnNotNullN(m_DocumentsLink("IDRV_POContrattoAdeguamento").Value), Me.txtIDProdotto.Value)
                End If
            Else
                m_DocumentsLink("Periodo") = Me.txtPeriodo.Text
            End If
        End If
    End If
    
    m_DocumentsLink.Save

    'Me.lblInfoSalvaRiga.Caption = "RICALCOLO DELLE RIGHE COLLEGATE IN CORSO....."
    
    DoEvents
    If Me.txtIDOggettoCollegato.Value = 0 Then
        ELIMINA_FLUSSO_DOCUMENTALE IDTipoOggettoVend, IDOggettoVend, fnNotNullN(m_DocumentsLink("IDOggetto").Value), fnNotNullN(m_DocumentsLink("IDTipoOggetto").Value)
    Else
        If Me.txtIDOggettoCollegato.Value <> IDOggettoVend Then
            ELIMINA_FLUSSO_DOCUMENTALE IDTipoOggettoVend, IDOggettoVend, fnNotNullN(m_DocumentsLink("IDOggetto").Value), fnNotNullN(m_DocumentsLink("IDTipoOggetto").Value)
        End If
        CREA_FLUSSO_DOCUMENTALE fncIDDocumentoAllegato(Me.txtIDOggettoCollegato.Value), Me.txtIDOggettoCollegato.Value, fnNotNullN(m_DocumentsLink("IDOggetto").Value), fnNotNullN(m_DocumentsLink("IDTipoOggetto").Value)
    End If
    
    Me.Caption = "CREAZIONE SCADENZA E FLUSSO DOCUMENTALE..."
    DoEvents

    If LINK_SEZIONALE_RATE > 0 Then
        IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(Me.txtIDOggettoRata.Value, Me.txtIDTipoOggettoRata.Value, 0)
        
        If IDOggettoScadenza > 0 Then
            ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, Me.txtIDOggettoRata.Value, Me.txtIDTipoOggettoRata.Value
            ELIMINA_SCADENZA IDOggettoScadenza
        End If
        
        If Me.txtIDOggettoCollegato.Value = 0 Then
            IDOggettoScadenza = GET_LINK_SCADENZA(Me.txtImportoRata.Value, IDClienteFatturazione, Me.txtNumeroRata.Text, Me.txtDataRata.Text, LINK_SEZIONALE_RATE, Me.txtPeriodo.Text)
            CREA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, Me.txtIDOggettoRata.Value, Me.txtIDTipoOggettoRata.Value
        End If
    End If
    
    
    Cn.CursorLocation = OLDCursor
    Me.Caption = Caption2Display
    Screen.MousePointer = 0
    
    m_DocumentsLink.Move Me.GrigliaRateContratto.ListIndex - 1
    
    

    'Me.lblInfoSalvaRiga.Caption = ""
    DoEvents
    'Si sposta sull'ultima riga
    'If Not (BrwMain.Visible) Then Change
    
    NuovaRata = 0
    

    DISEGNA_CONTROLLI GET_CONTROLLO_RATE_PAGATE(fnNotNullN(m_Document(m_Document.PrimaryKey).Value))
    
    Exit Sub

ERR_cmdSalvaRata_Click:

    MsgBox Err.Description, vbCritical, "Salvataggio dati"
    
    Cn.RollbackTrans
    Screen.MousePointer = 0
End Sub
Private Function GET_ESISTENZA_RATA(IDRata) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POStoriaRateContratto "
sSQL = sSQL & "WHERE IDRiferimentoRata=" & IDRata

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_ESISTENZA_RATA = False
Else
    GET_ESISTENZA_RATA = True
End If
rs.CloseResultset
Set rs = Nothing
End Function

Private Sub cmdSalvaServizio_Click()
On Error GoTo ERR_cmdSalvaServizio_Click

    If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        MsgBox "Salvare il documento prima di procedere ad inserire li servizi legati al tipo di contratto", vbInformation, "Salvataggio documento"
        Exit Sub
    End If
    
    If PermissionToSaveRiga = False Then
        Exit Sub
    End If
    
    m_DocumentsLink1("IDArticolo").Value = Me.CDServizio.KeyFieldID
    m_DocumentsLink1("IDRV_POCriterioRicorrenza").Value = Me.cboCriterioRicorrenza.CurrentID
    m_DocumentsLink1("OgniNumeroGiorni").Value = Me.txtOgniNumeroGiorni.Value
    m_DocumentsLink1("OgniNumeroMesi").Value = Me.txtOgniNumeroMesi.Value
    m_DocumentsLink1("OgniNumeroSettimane").Value = Me.txtOgniNumeroSettimane.Value
    m_DocumentsLink1("IDRV_POTipoDataInizioRicorrenza").Value = Me.cboTipoDataInizioRic.CurrentID
    m_DocumentsLink1("GiornoInizioRicorrenza").Value = Me.txtGiornoFissoInizioRic.Value
    m_DocumentsLink1("MeseInizioRicorrenza").Value = Me.txtMeseFissoInizioRic.Value
    m_DocumentsLink1("IDRV_POTipoDataFineRicorrenza").Value = Me.cboTipoDataFineRic.CurrentID
    m_DocumentsLink1("GiornoFineRicorrenza").Value = Me.txtGiornoFissoFineRic.Value
    m_DocumentsLink1("MeseFineRicorrenza").Value = Me.txtMeseFissoFineRic.Value
    m_DocumentsLink1("NumeroRicorrenze").Value = Me.txtNumeroRicorrenze.Value
    m_DocumentsLink1("IDRV_POContrattoPadre").Value = Me.txtIDContrattoPadre.Value
    m_DocumentsLink1("IDRV_POTipoAnnoInizioRicorrenza").Value = Me.cboTipoAnnoInizioRicorr.CurrentID
    m_DocumentsLink1("IDRV_POTipoAnnoFineRicorrenza").Value = Me.cboTipoAnnoFineRicorr.CurrentID
    
    Screen.MousePointer = 11
    Cn.BeginTrans
    
    m_DocumentsLink1.Save
    
    Cn.CommitTrans
    
    m_DocumentsLink1.Move Me.GrigliaServizi.ListIndex - 1
    
    Screen.MousePointer = 0
    
    Exit Sub
    
ERR_cmdSalvaServizio_Click:
    MsgBox Err.Description, vbCritical, "Salva Servizio"
    Cn.RollbackTrans
    Screen.MousePointer = 0
    
End Sub

Private Sub cmdSelServizio_Click()
On Error GoTo ERR_cmdSelServizio_Click
Dim IDUMPeriodoProd As Long
Dim IDTipoRateizzazioneProd As Long

    If (m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

    frmSelProdotti.Show vbModal
    
    If Me.cboTipoImpostazione.CurrentID = 1 Then
        m_DocumentsLink3.Refresh
    End If
    
    If Me.cboTipoImpostazione.CurrentID = 3 Then
        If fnNotNullN((m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)) <= 0 Then
            If Me.txtIDProdotto.Value > 0 Then
                IDUMPeriodoProd = GET_LINK_UM_PERIODO_PRED(Me.txtIDProdotto.Value)
                IDTipoRateizzazioneProd = GET_LINK_RATEIZZAZIONE_PRED(Me.txtIDProdotto.Value)
                
                If IDUMPeriodoProd > 0 Then Me.cboUMPeriodoProd.WriteOn IDUMPeriodoProd
                If IDTipoRateizzazioneProd > 0 Then Me.cboTipoRateizzazioneProd.WriteOn IDTipoRateizzazioneProd
                
                txtQtaPeriodo.Value = 1
                txtQtaPeriodo_LostFocus
            End If
        End If
    End If
    
Exit Sub
ERR_cmdSelServizio_Click:
    MsgBox Err.Description, vbCritical, "cmdSelServizio_Click"
End Sub

Private Sub cmdStampaProdotti_Click()
    On Error GoTo ERR_cmdCondizioniGenerali_Click

    If fnNotNullN(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) <= 0 Then Exit Sub
    
    If fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0 Then Exit Sub
    
    SaveSetting REGISTRY_KEY, App.EXEName, "IDField", "2"
    SaveSetting REGISTRY_KEY, "RV_POSnapProgetto", "Client", App.EXEName
    
    SaveSetting REGISTRY_KEY, "RV_POSnapProgetto", "IDContratto", fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    SaveSetting REGISTRY_KEY, "RV_POSnapProgetto", "IDContrattoProdotto", fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)
    SaveSetting REGISTRY_KEY, "RV_POSnapProgetto", "IDTipoVisualizzazione", 1
    SaveSetting REGISTRY_KEY, "RV_POSnapProgetto", "IDTipoLista", 2
    
    Shell App.Path & "\RV_POSnapProgetto.exe"
    
Exit Sub
ERR_cmdCondizioniGenerali_Click:
    MsgBox Err.Description, vbCritical, "cmdCondizioniGenerali_Click"
End Sub

Private Sub cmdStoriaAdeguamenti_Click()
On Error GoTo ERR_cmdStoriaAdeguamenti_Click
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    frmAdeguamenti.Show vbModal

Exit Sub
ERR_cmdStoriaAdeguamenti_Click:
    MsgBox Err.Description, vbCritical, "cmdStoriaAdeguamenti_Click"
End Sub

Private Sub cmdTrovaFattura_Click()
    frmTrovaFattura.Show vbModal
    If Me.txtIDOggettoCollegato.Value > 0 Then
        Me.chkRataFatturata.Value = vbChecked
    End If
End Sub

Private Sub Command1_Click()
    If Me.CDTecnico.KeyFieldID = 0 Then Exit Sub

    LINK_SITO_PER_ANAGRAFICA_TEL = 0
    LINK_ANAGRAFICA_TEL = Me.CDTecnico.KeyFieldID
    
    frmRiferimentiTelefonici.Show vbModal
End Sub

Private Sub Command2_Click()
    If Me.CDAgente.KeyFieldID = 0 Then Exit Sub

    LINK_SITO_PER_ANAGRAFICA_TEL = 0
    LINK_ANAGRAFICA_TEL = Me.CDAgente.KeyFieldID
    
    frmRiferimentiTelefonici.Show vbModal
End Sub

Private Sub Command3_Click()
    If Me.CDAmministratore.KeyFieldID = 0 Then Exit Sub

    LINK_SITO_PER_ANAGRAFICA_TEL = 0
    LINK_ANAGRAFICA_TEL = Me.CDAmministratore.KeyFieldID
    
    frmRiferimentiTelefonici.Show vbModal
End Sub

Private Sub Command4_Click()
    On Error GoTo ERR_cmdCondizioniGenerali_Click

    If fnNotNullN(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) <= 0 Then Exit Sub
    

    'Me.LabelLink2.IDReturn = fnNotNullN(fnNotNullN(m_Document(m_Document.PrimaryKey).Value))
    'Me.LabelLink2.RunApplication
    
    SaveSetting REGISTRY_KEY, App.EXEName, "IDField", "2"
    SaveSetting REGISTRY_KEY, "RV_POSnapProgetto", "Client", App.EXEName
    
    SaveSetting REGISTRY_KEY, "RV_POSnapProgetto", "IDContratto", fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    SaveSetting REGISTRY_KEY, "RV_POSnapProgetto", "IDContrattoProdotto", 0
        
    SaveSetting REGISTRY_KEY, "RV_POSnapProgetto", "IDTipoVisualizzazione", 1
    SaveSetting REGISTRY_KEY, "RV_POSnapProgetto", "IDTipoLista", 1


    'Shell "C:\Program Files (x86)\Diamante spa\DMT Professional v3.7\Bin\RV_POSnapProgetto.exe"
    Shell App.Path & "\RV_POSnapProgetto.exe"
Exit Sub
ERR_cmdCondizioniGenerali_Click:
    MsgBox Err.Description, vbCritical, "cmdCondizioniGenerali_Click"
End Sub

Private Sub Command5_Click()
    If Me.cboSitoPerAnagrafica.CurrentID = 0 Then Exit Sub
    
    LINK_SITO_PER_ANAGRAFICA_TEL = Me.cboSitoPerAnagrafica.CurrentID
    LINK_ANAGRAFICA_TEL = Me.CDCliente.KeyFieldID
    
    frmRiferimentiTelefonici.Show vbModal
End Sub

Private Sub Command6_Click()
Link_Contratto_Servizio = fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value)
Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
FILTRO_PRODOTTO_ASSOCIATO = 2


frmSelProdottiServizi.Show vbModal


Me.txtNProdAss.Value = GET_NUMERO_PRODOTTI_PER_SERVIZIO(fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))
Me.txtNInterventiServ.Value = GET_NUMERO_INTERVENTI_PER_SERVIZIO(fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))

End Sub

Private Sub Command7_Click()
On Error GoTo ERR_Command7_Click
    
    If (NON_MODIFICA_CONTRATTO = 1) Then Exit Sub
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey)) <= 0 Then
        MsgBox "Salvare il documento prima di procedere ad un inserimento", vbCritical, "Nuova riga documento"
        Exit Sub
    End If
    
    If Me.chkContrattoAttuale.Value = vbUnchecked Then Exit Sub
    
    If m_DocumentsLink3.TableNew Then
        m_DocumentsLink3.AbortNewRow
    End If
    
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink3.NewRow
    
    Me.txtQtaArtProd.Value = 1
    
    If (Me.cboTipoImpostazione.CurrentID <= 2) Then
        Me.cboTipoPeriodo.WriteOn 2
    Else
        Me.cboTipoPeriodo.WriteOn 1
    End If
    
    Me.cboUMPeriodoProd.WriteOn LINK_UM_PERIODO_AZIENDA

    
    'Me.txtQtaPeriodo.Value = 1
    txtQtaPeriodo_LostFocus
    
    Me.cboListinoProd.WriteOn GET_LINK_LISTINO(IDClienteFatturazione, TheApp.IDFirm)
    
    Me.CDArticoloProd.SetFocus
    
    
    'cmdSelServizio_Click
    
Exit Sub
ERR_Command7_Click:
    MsgBox Err.Description, vbCritical, "Command7_Click"
End Sub



Private Sub Command8_Click()
    If fnNotNullN(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) <= 0 Then Exit Sub
    
    If ((m_DocumentsLink.EOF) And (m_DocumentsLink.BOF)) Then Exit Sub
    
    'If Me.txtIDProdotto.Value = 0 Then Exit Sub
    
    If Me.txtIDOggettoCollegato.Value = 0 Then Exit Sub
    
    
    LabelLink3.IDFunction = GET_FUNZIONE_DA_IDOGGETTO(Me.txtIDOggettoCollegato.Value)  'GET_FUNZIONE(fnNotNull(m_DocumentsLink("IDTipoOggetto").Value))
    Me.LabelLink3.IDReturn = Me.txtIDOggettoCollegato.Value
    Me.LabelLink3.RunApplication
    
    
End Sub

Private Sub Command9_Click()
On Error GoTo ERR_cmdGeneraInterventi_Click
'On Error GoTo ERR_cmdGeneraInterventi_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim TestoMessaggio As String
Dim AvviaProcedura As Boolean
Dim rsInterventoTesta As ADODB.Recordset
Dim X As Long
Dim ErroreCoda As Boolean
Dim OLD_Cursor As Long


If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare il documento prima di procedere ad inserire i servizi legati al tipo di contratto", vbInformation, "Salvataggio documento"
    Exit Sub
End If
If fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value) <= 0 Then
    MsgBox "Salvare la riga del servizio prima di procedere alla generazione degli interventi", vbInformation, "Salvataggio documento"
    Exit Sub
End If

If MODULO_ATTIVATO_INT = 0 Then
    If Len(MODULO_DESCRIZIONE_INT) > 0 Then
        MsgBox "Il modulo " & MODULO_DESCRIZIONE_INT & " non è stato abilitato", vbInformation, TheApp.FunctionName
    Else
        MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
    End If
Exit Sub
End If

Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
Link_Contratto_Servizio = fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value)


Elabora_Tutti_Servizi_Contratto = 1

frmInterventiDaServ.Show vbModal

'Me.txtNProdAss.Value = GET_NUMERO_PRODOTTI_PER_SERVIZIO(fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))
'Me.txtNInterventiServ.Value = GET_NUMERO_INTERVENTI_PER_SERVIZIO(fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))

GET_GRIGLIA_INTERVENTI

Exit Sub
ERR_cmdGeneraInterventi_Click:
    MsgBox Err.Description, vbCritical, "ERR_cmdGeneraInterventi_Click"
End Sub

Private Sub Form_Activate()
    'Il codice di Form_Activate deve essere eseguito soltanto la prima volta,
    'all'avvio del programma.
    '
    'La variabile m_bOnFirstTime è usata per evitare di eseguire il codice seguente
    'quando si chiude un Form di dialogo e si riattiva frmMain.
    '
    'Queste inizializzazioni non sono state effettuate nella Sub Main() per evitare di
    'rendere visibili le variabili m_DocType, m_Document e m_Changed.
    If m_bOnFirstTime = True Then
    
        m_bOnFirstTime = False

        'Se il filtro di default restituisce dei record si va in modalità variazione
        'ma solo se il primo record non è bloccato altrimenti si va in modalità tabellare
        If Not (m_Document.EOF = True And m_Document.BOF = True) Then
            'Il filtro ha restituito almeno un record
             
            'Controlla se il primo record su cui si dovrebbe andare in variazione è bloccato.
            If m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
                'Il primo record NON è bloccato
                'allora si effettua il blocco e si va in modalità Variazione
                
                m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
                    
                'La vista alla partenza deve essere quella del Form
                BrwMain.Visible = False

                'Imposta la modalità variazione
                SetStatus4Modality Modify
                
            Else
                'Il primo record è bloccato
                'allora si parte in modalità tabellare
                
                BrwMain.Visible = True
                
                SetStatus4Modality Browse
            End If

            RefreshDescriptions4StatusBar
        Else
            'Il filtro di default non ha restituito nessun record.
            'Si va in modalità inserimento nuovo record
            NewRecord
            
            If (NON_MODIFICA_CONTRATTO = 1) Then
                NewSearch
            End If
        
        End If
               
    End If
    
  
    
End Sub

Private Sub Form_Initialize()
    ActivityBox.Visible = True
    
    'Impostazione iniziale del flag
    m_bOnFirstTime = True
    
    bEnableGuiEvent = True
End Sub

Private Sub Form_Load()
    'La vista tabellare deve trovarsi sopra tutti gli altri controlli
    BrwMain.ZOrder
    
    'IMPOSTA IL CONTROLLO CHE CONTIENTE I TUTTI GLI ALTRI CONTROLLI
    Set DMTSplitBar1.dContainer = Me.PicForm2
    'IMPOSTA L'UNITA' DI MISURA DEL FORM
    DMTSplitBar1.ScaleMode = DMTSplit_Twips
    'INIZIALIZZA LA SPLIT BAR
    DMTSplitBar1.SetSplitBar Me.ScaleHeight, Me.ScaleWidth, Me.PicForm.ScaleHeight, Me.PicForm.ScaleWidth

End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If ActivityBox.Visible Then
        imgSplitter.Left = ActivityBox.Width + ActivityBox.Left
    End If
    If m_PreviewWindowHandle > 0 Then
        MoveWindow m_PreviewWindowHandle, CInt(BarMenu.ClientAreaLeft / Screen.TwipsPerPixelX), CInt(BarMenu.ClientAreaTop / Screen.TwipsPerPixelY), CInt(BarMenu.ClientAreaWidth / Screen.TwipsPerPixelY), CInt(BarMenu.ClientAreaHeight / Screen.TwipsPerPixelX), True
    Else
        FormRecalcLayout
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim bHandled As Boolean
    
    m_EatKey = False
    
    ShortCut KeyCode, Shift
    
    If KeyCode = 0 And Shift = 0 Then
        m_EatKey = True
    Else
        m_EatKey = False
    End If
    
    
    Select Case KeyCode
        Case vbKeyPageDown
            DMTSplitBar1.ScrollDown
        Case vbKeyPageUp
            DMTSplitBar1.ScrollUp
    End Select

    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If m_EatKey Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' ATTENZIONE
    '-------------------------------------------------------------------------------------
    ' In questo metodo qualsiasi riferimento a proprietà o metodi di un oggetto dovrebbe
    ' essere 'protetto' dal test
    '
    '                                         If obj Is Nothing then .....
    '
    ' perchè il form potrebbe essere scaricato prima che l'oggetto stesso vengana istanziato.
    '-------------------------------------------------------------------------------------
    
    'chiude e distrugge il riferimento alle connessioni
        'CloseConnection
    'Distrugge il riferimento al recordset
    Set BrwMain.Recordset = Nothing
    
    Cancel = FormUnload
End Sub

Private Sub Form_Terminate()

    'Distrugge tutti gli oggetti allocati e provvede ad eliminare gli eventuali blocchi
    'effettuati dalla Semaforo.
    '(Inserire in DestroyObjects il codice per la distruzione degli oggetti allocati)
    DestroyObjects

End Sub

Private Sub brwMain_DblClick()
    
    
'-------------------------------------------------------------------
    'Il documento si sincronizza con la browse
'    If BrwMain.ListIndex > 0 Then
'        m_Document.Move BrwMain.ListIndex - 1
'    End If
'-------------------------------------------------------------------
'NOTA: La versione attuale della dmtGrid effettua automaticamente il
'      Move sul documento.
'-------------------------------------------------------------------
    
    'Se si è in modalità FilterDefinition il DblClick e la pressione
    'di Invio non devono avere alcun effetto
    If BrwMain.GuiMode <> dgFilterDefinition Then
    
        ChangeView
        BrowseReposition
        
        m_Document.AbortNew
        m_Changed = False
        ActivateBarButtons BTN_SAVE, False

        

    End If
End Sub

Private Sub brwMain_KeyDown(KeyCode As Integer, Shift As Integer)

    'Alla pressione del tasto INVIO dalla modalità tabellare si passa in modalità form.
    If KeyCode = vbKeyReturn And BrwMain.GuiMode = dgNormal And BrwMain.Visible Then
        brwMain_DblClick
    End If
    
    'Viene intercettata la pressione del tasto CANC
    'e la si comunica al form.
    If KeyCode = vbKeyDelete Then
    
        'Prima di cancellare sincronizzo il documento con la selezione
        'fatta nella browse
        If BrwMain.GuiMode = dgNormal And BrwMain.ListIndex > 0 Then
            m_Document.Move BrwMain.ListIndex - 1
        End If
    
        ShortCut KeyCode, Shift
    End If
    
End Sub


'Quando si selezionano i documenti dalla modalità tabellare la Caption del form
'va costruita leggendo i valori direttamente dalla riga selezionata nella griglia
'e non da un campo del documento perchè in modalità tabellare non viene eseguito
'il Move sul documento.
Private Sub BrwMain_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    
    If Not (m_Document.EOF = True And m_Document.BOF = True) Then
        'Monta la caption del form principale
        Me.Caption = Caption2Display(True)
    End If

End Sub


Private Sub BrwMain_OnChangeGuiMode()
    'Se si cambia modalità tramite il menù presente nel controllo
    'dmtGrid occorre effettuare delle impostazioni preliminari nella UserInterface
    
    If bEnableGuiEvent Then
    
        'Modalità FilterDefinition
        If BrwMain.GuiMode = dgFilterDefinition Then
            'Annulla una eventuale operazione di inserimento di un nuovo record
            If m_Document.TableNew Then
                m_Document.AbortNew
            End If
            
            'Impostazioni per la modalità Ricerca
            SetStatus4Modality Find
        End If
        
        'Modalità tabellare
        If BrwMain.GuiMode = dgNormal Then
            'Se si è premuto il pulsante "Visualizzazione tabellare" dalla browse
            'in modalità FilterDefinition e con il recordset vuoto, non si deve andare in
            'modalità tabellare (browse vuota) ma si deve restare in modalità ricerca.
            If (m_Document.EOF = True And m_Document.BOF = True) Then
                BrwMain.GuiMode = dgFilterDefinition
            Else
                'Impostazioni per la modalità tabellare
                SetStatus4Modality Browse
            End If
        End If
    
    End If
End Sub

'Scatenato prima che venga visualizzata la Toolbar della DmtGrid
Private Sub BrwMain_BeforeShowActions()
    
    'Quando si è in modalità FilterDefinition si può andare in
    'modalità tabellare solo se il documento contiene almeno un record.
    If BrwMain.GuiMode = dgFilterDefinition Then
        'Abilita/disabilita il pulsante Modalità Tabellare della dmtGrid
        BrwMain.Actions("TableMode").Enabled = (m_Document.EOF <> True And m_Document.BOF <> True)
    End If
End Sub

'Scatenato quando dalla Browse ( in modalità FilterDefinition ) si clicca su esegui ricerca.
Private Sub BrwMain_OnApplyFilter(ByVal Filter As String)
    ExecuteSearch
End Sub

Private Sub FraImportiProd_DblClick()

End Sub





Private Sub GrigliaAcconti_DblClick()
'    If fnNotNullN(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) <= 0 Then Exit Sub
'
'    If ((Me.GrigliaAcconti.Recordset.EOF) And (Me.GrigliaAcconti.Recordset.BOF)) Then Exit Sub
'
'    'If Me.txtIDProdotto.Value = 0 Then Exit Sub
'
'    LabelLink1.IDFunction = GET_FUNZIONE(fnNotNull(Me.GrigliaAcconti.AllColumns("IDTipoOggetto").Value))
'    Me.LabelLink1.IDReturn = fnNotNull(Me.GrigliaAcconti.AllColumns("IDOggetto").Value)
'    Me.LabelLink1.RunApplication
    
End Sub

Private Sub Label1_Click(Index As Integer)

If Me.cboTipoImpostazione.CurrentID <> 1 Then Exit Sub

If m_Document("IDRV_POContratto") > 0 Then
    If Index = 9 Then
        Link_Contratto = m_Document("IDRV_POContratto").Value
        frmImportiNonStandard.Show vbModal
    End If
    If Index = 35 Then
        Link_Contratto = m_Document("IDRV_POContratto").Value
        frmRiepilogoImporto.Show vbModal
    End If
End If

End Sub





Private Sub lblLinkFattContratto_AfterRunServerApplication(ByVal lIDResultKey As Long)
On Error GoTo ERR_lblLinkFattContratto_AfterRunServerApplication
    
    DeleteSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POCreazioneDocumenti", "IDContratto"
    
    m_DocumentsLink.Refresh

Exit Sub
ERR_lblLinkFattContratto_AfterRunServerApplication:
    MsgBox Err.Description, vbCritical, "lblLinkFattContratto_AfterRunServerApplication"
End Sub

Private Sub lblLinkFattContratto_BeforeRunServerApplication()
On Error GoTo ERR_lblLinkFattContratto_BeforeRunServerApplication

    SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POCreazioneDocumenti", "IDContratto", fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

Exit Sub
ERR_lblLinkFattContratto_BeforeRunServerApplication:
    MsgBox Err.Description, vbCritical, "lblLinkFattContratto_BeforeRunServerApplication"
End Sub

Private Sub m_App_OnRun(ByVal Proc As Process)
    Dim Parameter As DMTRunAppLib.Parameter

    On Error GoTo ErrorHandler
    
    Set m_Process = Proc
    Set m_DocType = m_Process.IDocType
    
    
    '.................................................................................................................................
    '.................................................................................................................................
    'Gestione preliminare della Semaforo per il controllo dei conflitti di multiutenza
    
    
    'Inizializza la Semaforo
    InitSemaphore
    
    ' Verifica se l'applicazione corrente è bloccata da altri gestori.
    ' (Il controllo avviene sul Tipo Oggetto correntemente trattato.)
    If Not m_Semaphore.IsActionAvailable(m_DocType.ID, SemAllObjects, SemAllActions) Then
        '-------------------------------------------------------------
        'Il programma è bloccato da un'altra manutenzione in esecuzione.
        '-------------------------------------------------------------
        
        'Scarica il form
        Unload Me
       
        'Prima di terminare il programma è bene distruggere tutti gli oggetti allocati
        DestroyObjects
       
        'Termina il programma
        End
    End If
    
    '----------------------------------------------------
    'Il programma non è bloccato e prosegue normalmente.
    '----------------------------------------------------
    
    'Ripulisce la tabella semaforo.
    'Se era avvenuto un crash di sistema questo garantisce il ripristino della situazione.
    SemaphoreUnlock
    
    'Imposta gli eventuali blocchi (semaforo) su altre manutenzioni.
    SemaphoreLock
    '.................................................................................................................................
    '.................................................................................................................................
    
    
    
    Select Case Proc.Name
        '*
        'Inserire il codice per la gestione del processo
        '*
        Case "Manutenzione"
        '   For Each Parameter In Proc.Parameters
        '       Select Case Parameter.Name
        '       *
        '       Inserire il codice per la gestione del parametro
        '       *
        '       Case ParameterName??????
        '       End Select
        '   next
           Start 'di solito
    
    Case Else
    
        'cbcx
        'QUESTA PARTE DEVE ESSERE RIVISTA
        '-----------------------------------------------------------------
        
'''''        Dim ErrorMsg As String
'''''
'''''        ErrorMsg = "No processes to execute" & vbCrLf
'''''        ErrorMsg = ErrorMsg & "This application is able to execute these processes:" & vbCrLf
'''''        '*
'''''        'Inserire i processi che l'applicazione sa eseguire
'''''        '*
'''''        'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE & vbCrLf
'''''        'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE_EXTENDED_DATABASE & vbCrLf
'''''        'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE_DA_SHELL & vbCrLf
'''''        Err.Raise ERR_NO_PROCESSES, , ErrorMsg


    End Select
    Exit Sub
ErrorHandler:
    SemaphoreUnlock
    ShowErrorLog
End Sub

Private Sub m_Document_OnReposition()
Dim m_condition As DmtDocManLib.Condition

    'Viene creata (se non è già stato fatto) la collezione FormFields
    CreateFormFields
    BLoading = 1
    
    If Not m_Document.TableNew Then
        'Se EOF = true o BOF = true vuol dire che si è andati oltre l'ultimo o
        'prima del primo record. In tal caso non si deve fare il refresh dei
        'controlli del form.
        If Not (m_Document.EOF Or m_Document.BOF) Then
            BrowseReposition
            
            'cbcx
            '---------------------------------------------
            'Gestione processo On_Extend
'            If Not m_ExtendApplication Is Nothing Then
'                'Notifica l'identificativo unico del documento corrente
'                m_ExtendApplication.PrimaryID = m_Document.Fields("ID" & m_App.TableName).Value
'            End If
            
        End If
    Else
        'Nel caso di inserimento nuovo record ripulisce i campi del form
        ClearFormFields
    End If
    
    'rif11 begin
    

    
   'If Me.GrigliaFasiIntervento.ColumnsHeader.Count > 0 Then
       On Error Resume Next
        'Binding mediante le proprietà DataMember e DataSource.
        'Me.GrigliaFasiIntervento.DataMember = m_DocumentsLink2.TableName
        'Set Me.GrigliaFasiIntervento.DataSource = m_Document

        'Binding mediante la proprietà Recordset
       
        Set Me.GrigliaRateContratto.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink.TableName).Data
        Set Me.GrigliaServizi.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink1.TableName).Data
        Set Me.GrigliaAdeg.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink2.TableName).Data
        Set Me.GrigliaProd.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink3.TableName).Data

    'End If
    
    Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    'Link_StoriaContratto = ContrattoAttualeStorico(fnNotNullN(m_Document(m_Document.PrimaryKey).Value))
    
    If m_Document(m_Document.PrimaryKey).Value > 0 Then
        Me.CDCliente.Enabled = False
        Me.cboTipoImpostazione.Enabled = False
        Link_ContoPDC = fnNotNullN(m_Document("IDContoPDC"))

'        If Me.cboTipoImpostazione.CurrentID <> 3 Then
'            DISEGNA_CONTROLLI GET_CONTROLLO_RATE_PAGATE(fnNotNullN(m_Document(m_Document.PrimaryKey).Value))
'        End If
    Else
        Me.CDCliente.Enabled = True
        Me.cboTipoImpostazione.Enabled = True
        Link_ContoPDC = 0
    End If
    
    'rif11 end

        
    If Me.cboTipoImpostazione.CurrentID <> 3 Then
        DISEGNA_CONTROLLI GET_CONTROLLO_RATE_PAGATE(fnNotNullN(m_Document(m_Document.PrimaryKey).Value))
    End If

    LINK_LISTINO_AZIENDA = GET_LINK_LISTINO_AZIENDA(TheApp.IDFirm)
    
'    If VISUALIZZA_IMPORTI_PROD = 1 Then
'        Me.FraImportiProd.Visible = True
'    Else
'        Me.FraImportiProd.Visible = False
'    End If
    'Adeguamento dei prodotti
'    With Me.cboAdegProd
'        Set .Database = m_App.Database.Connection
'        .AddFieldKey "IDRV_POContrattoAdeguamento"
'        .DisplayField = "DescrizioneAdeguamento"
'        .SQL = "SELECT * FROM RV_POContrattoAdeguamento "
'        .SQL = .SQL & "WHERE IDRV_POContrattoPadre=" & Me.txtIDContrattoPadre.Value
'        .SQL = .SQL & " ORDER BY NumeroAdeguamento "
'    End With
'
    If Me.cboTipoImpostazione.CurrentID <= 2 Then
        Me.txtImportoTotAdeg.Value = GET_TOTALE_ADEGUAMENTI_DETTAGLIO(fnNotNullN(m_Document(m_Document.PrimaryKey).Value))
    End If
    If Me.cboTipoImpostazione.CurrentID = 3 Then
        'Me.txtImportoTotAdeg.Value = Me.txtImportoAttuale.Value - GET_TOTALI_ACCONTI_CONTRATTO(fnNotNullN(m_Document("IDOggetto").Value), m_DocType.ID)
        If Me.chkGeneraAccontiSaldo.Value = vbChecked Then
            Me.txtImportoTotAdeg.Value = Me.txtImportoAttuale.Value - GET_SALDO_CONTRATTO(fnNotNullN(m_Document("IDOggetto").Value))
        Else
            Me.txtImportoTotAdeg.Value = GET_TOTALE_ADEGUAMENTI_DETTAGLIO(fnNotNullN(m_Document(m_Document.PrimaryKey).Value))
        End If
    End If
    
    'GET_GRIGLIA_INTERVENTI
    'GET_GRIGLIA_ACCONTI
    Set Me.GrigliaInterventi.Recordset = Null
    
    InitVariabiliAltriDati m_Document
    'Me.txtAltriDati.Text = GET_CARATTERISTICHE_RISORSA(Me.MSFlexGrid1) 'GET_DESCRIZIONE_ALTRI_DATI
    If NON_VISUALIZZARE_ALTRI_DATI = 1 Then
        Me.fraAltriDati.Visible = False
    Else
        Me.fraAltriDati.Visible = True
        GET_CARATTERISTICHE_RISORSA Me.MSFlexGrid1
    End If
    
    GET_TOTALE_PRODOTTI fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    Screen.MousePointer = 0
    BLoading = 0
    
End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 07/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: AutoLostFocus
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Forza un LostFocus del controllo attivo ed attende la gestione di eventuali eventi associati.
'                  Alla fine ripristina il fuoco sul controllo iniziale.
'                  Usata quando si clicca sulla toolbar e quando si utilizza l'acceleratore per il salvataggio SHIFT + F12
'                  (in tal caso infatti non viene scatenato l'evento BarMenu_Click)
'
'**/
Private Sub AutoLostFocus()
    Dim Ctr As Control

    
    'Se si è in modalità FilterDefinition non si deve spostare il fuoco
    'altrimenti Taglia, Copia e Incolla (dalla toolbar) non possono funzionare
    If BrwMain.GuiMode <> dgFilterDefinition And Not Me.ActiveControl Is Nothing Then
    
        'Memorizza il controllo che ha il fuoco
        Set Ctr = Me.ActiveControl
    
        'Forza il lost focus del controllo attivo
        Globali.SetFocus PicForm.hwnd
        
        'Vengono gestiti gli eventi LostFocus (se previsti)
        DoEvents
        
        'Ripristina il fuoco sul controllo.
        Ctr.SetFocus
        
    End If

End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 11/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: InitSemaphore
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Inizializzazione del semaforo per la gestione
'                  dei conflitti in caso di multiutenza

'
'**/
Private Sub InitSemaphore()

    Set m_Semaphore = New Semaforo.dmtSemaphore
    Set m_Semaphore.Database = m_App.Database.Connection
    Set m_Semaphore.objRes = gResource
    
    m_Semaphore.IDUser = m_App.IDUser
    m_Semaphore.IDBranch = m_App.Branch
    m_Semaphore.IDFunction = m_App.FunctionID
    
End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 12/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: SemaphoreLock
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'                 ////////////////////////////////////////////////////////////////////////
'                     Impostare qui gli eventuali blocchi sulle altre manutenzioni
'                 ////////////////////////////////////////////////////////////////////////
'**/
Private Sub SemaphoreLock()
    If Not m_Semaphore Is Nothing Then
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        'Personalizzare, se necessario, le righe sottostanti
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        
'        m_Semaphore.SetObjectAction TO_TIPO_OGGETTO_XXX, SemAllObjects, SemAllActions
'        m_Semaphore.SetObjectAction TO_TIPO_OGGETTO_YYY, SemAllObjects, SemAllActions
'        m_Semaphore.SetObjectAction TO_TIPO_OGGETTO_ZZZ, SemAllObjects, SemAllActions

    End If
End Sub

'**+
'Autore: Diamante s.p.a
'Data creazione: 12/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: SemaphoreUnlock
'
'Parametri:
'
'Valori di ritorno:

'Funzionalità:
'                 //////////////////////////////////////////////////////////////////////////////////////////////////
'                     Sbloccare qui le altre manutenzioni (bloccate precedentemente in SemaphoreLock)
'                 //////////////////////////////////////////////////////////////////////////////////////////////////
'
'**/
Private Sub SemaphoreUnlock()
    If Not m_Semaphore Is Nothing Then
    
        'Ripulisce la tabella semaforo per quanto riguarda il Tipo Oggetto e l'utente correnti
        m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
        
        
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        'Personalizzare, se necessario, le righe sottostanti
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        
        'Sblocca le manutenzioni bloccate precedentemente
'        m_Semaphore.ClearObjectAction TO_TIPO_OGGETTO_XXX, SemAllObjects, SemAllActions
'        m_Semaphore.ClearObjectAction TO_TIPO_OGGETTO_YYY, SemAllObjects, SemAllActions
'        m_Semaphore.ClearObjectAction TO_TIPO_OGGETTO_ZZZ, SemAllObjects, SemAllActions
    
    End If
End Sub




'**+
'Autore: Diamante s.p.a
'Data creazione: 11/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: DestroyObjects
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'                  ////////////////////////////////////////////////////////////////////////////////////////////////////
'                  /         Inserire qui il codice per distruggere (prima che venga terminato il programma)     /
'                  /         tutti gli oggetti allocati                                                                              /
'                  ////////////////////////////////////////////////////////////////////////////////////////////////////
'
'**/
Private Sub DestroyObjects()
    
    'Sblocca gli eventuali gestori bloccati da questa manutenzione
    SemaphoreUnlock

    Set m_FormFields = Nothing
    Set m_Report = Nothing
    Set m_ActiveFilter = Nothing
    Set m_Document = Nothing
    Set m_Process = Nothing
    Set m_App = Nothing
    Set m_Semaphore = Nothing
    
    'cbcx
    'Set m_ExtendApplication = Nothing
End Sub



'rif8 begin

'**+
'Autore: Diamante s.p.a
'Data creazione: 03/11/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: m_DocumentsLink_OnReposition
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Operazioni da effettuare al Reposition del sottodocumento.
'
'**/
Private Sub m_DocumentsLink_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink.BOF And m_DocumentsLink.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        Me.txtNumeroRata.Value = fnNotNullN(m_DocumentsLink("NumeroRata").Value)
        Me.txtDataRata.Value = fnNotNullN(m_DocumentsLink("DataRata").Value)
        Me.txtImportoRata.Value = fnNotNullN(m_DocumentsLink("ImportoRata").Value)
        Me.cboPagamentoRataContratto.WriteOn fnNotNullN(m_DocumentsLink("IDPagamentoRata").Value)
                
        Me.txtIDOggettoCollegato.Value = fnNotNullN(m_DocumentsLink("IDOggettoCollegato").Value)
        Me.chkRataFatturata.Value = Abs(fnNotNullN(m_DocumentsLink("Fatturata").Value))
        
        Me.txtPeriodo.Text = fnNotNull(m_DocumentsLink("Periodo"))
        Me.cboIVARateContratto.WriteOn fnNotNullN(m_DocumentsLink("IDIvaFatturazione"))
        Me.txtDataInizioPer.Value = fnNotNullN(m_DocumentsLink("DataInizioPeriodo").Value)
        Me.txtDataFinePer.Value = fnNotNullN(m_DocumentsLink("DataFinePeriodo").Value)
        Me.txtIDOggettoRata.Value = fnNotNullN(m_DocumentsLink("IDOggetto").Value)
        Me.txtIDTipoOggettoRata.Value = fnNotNullN(m_DocumentsLink("IDTipoOggetto").Value)
        
        Me.chkNonFatturareRata.Value = Abs(fnNotNullN(m_DocumentsLink("NonFatturare").Value))
        Me.txtNoteNonFattRata.Text = fnNotNull(m_DocumentsLink("AnnotazioniNonFatturare").Value)
        Me.txtIDProdottoRata.Value = fnNotNullN(m_DocumentsLink("IDRV_POProdotto").Value)
        Me.txtIDProdRifContr.Value = fnNotNullN(m_DocumentsLink("IDRV_POContrattoProdotti").Value)
        
        bValue = True
        
        
        '----------------------------------------------------------------------------
        'Popola i controlli associati al sottodocumento con i valori presenti
        'nell'oggetto DocumentsLink
        '----------------------------------------------------------------------------
       
    Else
        'Il DocumentsLink è vuoto - non contiene righe.
        Me.txtNumeroRata.Value = 0
        Me.txtDataRata.Value = 0
        Me.txtImportoRata.Value = 0
        Me.cboPagamentoRataContratto.WriteOn 0
        Me.chkRataFatturata.Value = 0
        Me.txtIDOggettoCollegato.Value = 0
        Me.txtPeriodo.Text = ""
        Me.cboIVARateContratto.WriteOn 0
        Me.txtDataInizioPer.Value = 0
        Me.txtDataFinePer.Value = 0
        Me.txtIDOggettoRata.Value = 0
        Me.txtIDTipoOggettoRata.Value = 0
        Me.chkNonFatturareRata.Value = 0
        Me.txtNoteNonFattRata.Text = ""
        Me.txtIDProdottoRata.Value = 0
        Me.txtIDProdRifContr.Value = 0
        '---------------------------------------------
        'Ripulisce i controlli associati al sottodocumento
        '---------------------------------------------

        bValue = False
    End If
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
   
    Me.txtDataRata.Enabled = bValue
    Me.txtImportoRata.Enabled = bValue
    Me.cboPagamentoRataContratto.Enabled = bValue
    Me.chkRataFatturata.Enabled = bValue
    Me.cboIVARateContratto.Enabled = bValue
    Me.txtDataInizioPer.Enabled = bValue
    Me.txtDataFinePer.Enabled = bValue
    'Me.txtIDOggettoRata.Enabled = bValue
    'Me.txtIDTipoOggettoRata.Enabled = bValue
    Me.chkNonFatturareRata.Enabled = bValue
    Me.txtNoteNonFattRata.Enabled = bValue

    Me.cmdNuovaRata.Enabled = True
    Me.cmdSalvaRata.Enabled = bValue
    Me.cmdEliminaRata.Enabled = bValue

    If Me.chkContrattoAttuale.Value = vbUnchecked Then
        Me.cmdNuovaRata.Enabled = False
        
        Me.cmdEliminaRata.Enabled = False
    End If
        
    NuovaRata = 0
    
End Sub
Public Sub ConnessioneDiamanteADO()
On Error GoTo ERR_ConnessioneDiamanteADO
    '------------------------------
    'APERTURA DELLA CONNESSIONE
    '------------------------------
    
    'Leggiamo il tipo di database utilizzato (Access o SQL Server)
    'Apriamo la connessione in base al tipo di database rilevato
    '(MenuOptions.DBType restituisce il valore del DBType)
    'Select Case MenuOptions.DBType
    '    Case 0 'CONNESSIONE_SQL_SERVER            'Microsoft SQL Server
    '        Set Cn = adoEngine.adoEnvironments(0).OpenConnection("", , , "DSN=Diamante;UID=sa;PWD=")
    '    Case 1 'CONNESSIONE_ACCESS               'Microsoft ACCESS
    '        Set Cn = adoEngine.adoEnvironments(0).OpenConnection("", , , "DSN=Diamante;UID=admin;PWD=dmt192981046")
    '    Case -1
            'Se la voce DBType non viene trovata nel file di registro
            'vuol dire che Diamante non è stato installato correttamente
    '        MsgBox "Impossibile avviare il programma. Diamante non è stato installatto correttamente!", vbCritical, "Aggiornamento scadenze"
    '        End
    'End Select
    
    Set Cn = m_App.Database.Connection
    
Exit Sub
ERR_ConnessioneDiamanteADO:
    MsgBox Err.Description, vbCritical, "Connessione Diamante di tipo ADO"
End Sub

Private Sub m_DocumentsLink1_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink1.BOF And m_DocumentsLink1.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        
        Me.CDServizio.Load fnNotNullN(m_DocumentsLink1("IDArticolo").Value)
        Me.cboCriterioRicorrenza.WriteOn fnNotNullN(m_DocumentsLink1("IDRV_POCriterioRicorrenza").Value)
        Me.txtOgniNumeroGiorni.Value = fnNotNullN(m_DocumentsLink1("OgniNumeroGiorni").Value)
        Me.txtOgniNumeroMesi.Value = fnNotNullN(m_DocumentsLink1("OgniNumeroMesi").Value)
        Me.txtOgniNumeroSettimane.Value = fnNotNullN(m_DocumentsLink1("OgniNumeroSettimane").Value)
        Me.cboTipoDataInizioRic.WriteOn fnNotNullN(m_DocumentsLink1("IDRV_POTipoDataInizioRicorrenza").Value)
        Me.txtGiornoFissoInizioRic.Value = fnNotNullN(m_DocumentsLink1("GiornoInizioRicorrenza").Value)
        Me.txtMeseFissoInizioRic.Value = fnNotNullN(m_DocumentsLink1("MeseInizioRicorrenza").Value)
        Me.cboTipoDataFineRic.WriteOn fnNotNullN(m_DocumentsLink1("IDRV_POTipoDataFineRicorrenza").Value)
        Me.txtGiornoFissoFineRic.Value = fnNotNullN(m_DocumentsLink1("GiornoFineRicorrenza").Value)
        Me.txtMeseFissoFineRic.Value = fnNotNullN(m_DocumentsLink1("MeseFineRicorrenza").Value)
        Me.txtNumeroRicorrenze.Value = fnNotNullN(m_DocumentsLink1("NumeroRicorrenze").Value)
        Me.cboTipoAnnoInizioRicorr.WriteOn fnNotNullN(m_DocumentsLink1("IDRV_POTipoAnnoInizioRicorrenza").Value)
        Me.cboTipoAnnoFineRicorr.WriteOn fnNotNullN(m_DocumentsLink1("IDRV_POTipoAnnoFineRicorrenza").Value)
        
        bValue = True
        
        
        '----------------------------------------------------------------------------
        'Popola i controlli associati al sottodocumento con i valori presenti
        'nell'oggetto DocumentsLink
        '----------------------------------------------------------------------------
       
    Else
        'Il DocumentsLink è vuoto - non contiene righe.
        Me.CDServizio.Load 0
        Me.cboCriterioRicorrenza.WriteOn 0
        Me.txtOgniNumeroGiorni.Value = 0
        Me.txtOgniNumeroMesi.Value = 0
        Me.txtOgniNumeroSettimane.Value = 0
        Me.cboTipoDataInizioRic.WriteOn 0
        Me.txtGiornoFissoInizioRic.Value = 0
        Me.txtMeseFissoInizioRic.Value = 0
        Me.cboTipoDataFineRic.WriteOn 0
        Me.txtGiornoFissoFineRic.Value = 0
        Me.txtMeseFissoFineRic.Value = 0
        Me.txtNumeroRicorrenze.Value = 0
        Me.txtNProdAss.Value = 0
        Me.txtNInterventiServ.Value = 0
        Me.cboTipoAnnoInizioRicorr.WriteOn 0
        Me.cboTipoAnnoFineRicorr.WriteOn 0
        
        '---------------------------------------------
        'Ripulisce i controlli associati al sottodocumento
        '---------------------------------------------
        
        bValue = False
    End If
    

    

    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
    Me.CDServizio.Enabled = bValue
    Me.cboCriterioRicorrenza.Enabled = bValue
    Me.txtOgniNumeroGiorni.Enabled = bValue
    Me.txtOgniNumeroMesi.Enabled = bValue
    Me.txtOgniNumeroSettimane.Enabled = bValue
    Me.cboTipoDataInizioRic.Enabled = bValue
    Me.txtGiornoFissoInizioRic.Enabled = bValue
    Me.txtMeseFissoInizioRic.Enabled = bValue
    Me.cboTipoDataFineRic.Enabled = bValue
    Me.txtGiornoFissoFineRic.Enabled = bValue
    Me.txtMeseFissoFineRic.Enabled = bValue
    Me.txtNumeroRicorrenze.Enabled = bValue
    Me.cboTipoAnnoInizioRicorr.Enabled = bValue
    Me.cboTipoAnnoFineRicorr.Enabled = bValue
  
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
    Me.cmdNuovoServizio.Enabled = True
    Me.cmdSalvaServizio.Enabled = bValue
    Me.cmdEliminaServizio.Enabled = bValue
    
    
    
    cboCriterioRicorrenza_Click
    cboTipoDataFineRic_Click
    cboTipoDataInizioRic_Click
        
        
        
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
            Me.cmdNuovoServizio.Enabled = True
            Me.cmdSalvaServizio.Enabled = bValue
            Me.cmdEliminaServizio.Enabled = bValue
    Else
    
        If Me.chkContrattoAttuale.Value = vbUnchecked Then
            Me.cmdNuovoServizio.Enabled = False
            Me.cmdSalvaServizio.Enabled = False
            Me.cmdEliminaServizio.Enabled = False
        Else
            Me.cmdNuovoServizio.Enabled = True
            Me.cmdSalvaServizio.Enabled = bValue
            Me.cmdEliminaServizio.Enabled = bValue
            
        End If
    End If
    
    Me.txtNProdAss.Value = GET_NUMERO_PRODOTTI_PER_SERVIZIO(fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))
    Me.txtNInterventiServ.Value = GET_NUMERO_INTERVENTI_PER_SERVIZIO(fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))
    
End Sub
Private Sub m_DocumentsLink3_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    
    bLoadingProdotti = True

    If Not (m_DocumentsLink3.BOF And m_DocumentsLink3.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        Me.txtIDProdotto.Value = fnNotNullN(m_DocumentsLink3("IDRV_POProdotto").Value)
        Me.txtDescrProdotto.Text = fnNotNull(m_DocumentsLink3("DescrizioneAggiuntiva").Value)
        Me.txtValIdentProd.Text = fnNotNull(m_DocumentsLink3("ValoreIndentificativo").Value)
        Me.txtNoteProdotto.Text = fnNotNull(m_DocumentsLink3("Annotazioni").Value)
        Me.txtQtaProdotto.Value = fnNotNullN(m_DocumentsLink3("Quantita").Value)
        Me.chkDismessoProd.Value = fnNotNullN(m_DocumentsLink3("Dismesso").Value)
        Me.txtDataDismesso.Value = fnNotNullN(m_DocumentsLink3("DataDismesso").Value)
        
        Me.CDArticoloProd.Load fnNotNullN(m_DocumentsLink3("IDArticolo").Value)
        Me.cboTipoPeriodo.WriteOn fnNotNullN(m_DocumentsLink3("IDRV_POTipoPeriodo").Value)
        Me.cboUMPeriodoProd.WriteOn fnNotNullN(m_DocumentsLink3("IDRV_POUnitaDiMisuraPeriodo").Value)
        Me.txtDataInizioProd.Value = fnNotNullN(m_DocumentsLink3("DataInizioPeriodo").Value)
        Me.txtOraInizioProd.Text = fnNotNull(m_DocumentsLink3("OraInizioPeriodo").Value)
        
        Me.txtDataFineProd.Value = fnNotNullN(m_DocumentsLink3("DataFinePeriodo").Value)
        Me.txtOraFineProd.Text = fnNotNull(m_DocumentsLink3("OraFinePeriodo").Value)
        Me.txtQtaPeriodo.Value = fnNotNullN(m_DocumentsLink3("QuantitaPeriodo").Value)
        
        Me.cboListinoProd.WriteOn fnNotNullN(m_DocumentsLink3("IDListino").Value)
        
        Me.cboUMArtProd.WriteOn fnNotNullN(m_DocumentsLink3("IDUnitaDiMisuraArticolo").Value)
        Me.cboIvaProd.WriteOn fnNotNullN(m_DocumentsLink3("IDIva").Value)
        Me.txtAliquotaIvaProd.Value = fnNotNullN(m_DocumentsLink3("AliquotaIva").Value)
        
        Me.txtQtaArtProd.Value = fnNotNullN(m_DocumentsLink3("QuantitaArticolo").Value)
        
        Me.txtImpUniProd.Value = fnNotNullN(m_DocumentsLink3("ImportoUnitario").Value)
        Me.txtSconto1Prod.Value = fnNotNullN(m_DocumentsLink3("Sconto1").Value)
        Me.txtSconto2Prod.Value = fnNotNullN(m_DocumentsLink3("Sconto2").Value)
        Me.txtImponibileProd.Value = fnNotNullN(m_DocumentsLink3("Imponibile").Value)
        Me.txtScontoImpProd.Value = fnNotNullN(m_DocumentsLink3("ScontoAImporto").Value)
        Me.txtImportoIvaProd.Value = fnNotNullN(m_DocumentsLink3("ImportoIva").Value)
        Me.txtTotaleRigaProd.Value = fnNotNullN(m_DocumentsLink3("TotaleRiga").Value)
        
        Me.txtQuantitaEffettiva.Value = fnNotNullN(m_DocumentsLink3("QuantitaEffettiva").Value)
        Me.chkEscludiGiorniFestivi.Value = fnNotNullN(m_DocumentsLink3("EscludiGiorniFestivi").Value)
        Me.chkEscludiSabato.Value = fnNotNullN(m_DocumentsLink3("EscludiSabato").Value)
        Me.chkConducente.Value = fnNotNullN(m_DocumentsLink3("Conducente").Value)
        Me.chkACorpo.Value = fnNotNullN(m_DocumentsLink3("ACorpo").Value)
        Me.cboAnaOperatoreProd.WriteOn fnNotNullN(m_DocumentsLink3("IDAnagraficaOperatore").Value)
        Me.txtNoteProdInt.Text = fnNotNull(m_DocumentsLink3("AnnotazioniPerIntervento").Value)
        Me.cboTipoRateizzazioneProd.WriteOn fnNotNullN(m_DocumentsLink3("IDRateizzazione").Value)
        Me.chkRinnovareProd.Value = fnNotNullN(m_DocumentsLink3("NonRinnovare").Value)
        Me.chkGeneraUnaRataProd.Value = fnNotNullN(m_DocumentsLink3("NonRateizzare").Value)


        
        bValue = True
        
        '----------------------------------------------------------------------------
        'Popola i controlli associati al sottodocumento con i valori presenti
        'nell'oggetto DocumentsLink
        '----------------------------------------------------------------------------
       
    Else
        'Il DocumentsLink è vuoto - non contiene righe.
        '---------------------------------------------
        'Ripulisce i controlli associati al sottodocumento
        '---------------------------------------------
        Me.txtIDProdotto.Value = 0
        Me.txtDescrProdotto.Text = ""
        Me.txtValIdentProd.Text = ""
        Me.txtNoteProdotto.Text = ""
        'Me.txtQtaProdotto.Value = 0
        Me.chkDismessoProd.Value = vbUnchecked
        Me.txtDataDismesso.Value = 0

        Me.CDArticoloProd.Load 0
        Me.cboTipoPeriodo.WriteOn 0
        Me.cboUMPeriodoProd.WriteOn 0
        
        Me.txtDataInizioProd.Value = 0
        Me.txtOraInizioProd.Text = ""
        
        Me.txtDataFineProd.Value = 0
        Me.txtOraFineProd.Text = ""
        
        Me.txtQtaPeriodo.Value = 0
        
        Me.cboListinoProd.WriteOn 0
        
        Me.cboUMArtProd.WriteOn 0
        Me.cboIvaProd.WriteOn 0
        Me.txtAliquotaIvaProd.Value = 0

        Me.txtQtaArtProd.Value = 0
        
        Me.txtImpUniProd.Value = 0
        Me.txtSconto1Prod.Value = 0
        Me.txtSconto2Prod.Value = 0
        Me.txtImponibileProd.Value = 0
        Me.txtScontoImpProd.Value = 0
        Me.txtImportoIvaProd.Value = 0
        Me.txtTotaleRigaProd.Value = 0
        
        Me.txtQuantitaEffettiva.Value = 0
        Me.chkEscludiGiorniFestivi.Value = 0
        Me.chkEscludiSabato.Value = 0
        Me.chkConducente.Value = 0
        Me.chkACorpo.Value = 0
        Me.cboAnaOperatoreProd.WriteOn 0
        
        Me.txtNoteProdInt.Text = ""
        Me.cboTipoRateizzazioneProd.WriteOn 0
        Me.chkRinnovareProd.Value = 0
        Me.chkGeneraUnaRataProd.Value = 0
        

        bValue = False
    End If
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
        Me.txtIDProdotto.Enabled = bValue
        
        Me.txtDescrProdotto.Enabled = bValue
        Me.txtNoteProdotto.Enabled = bValue
        Me.txtValIdentProd.Enabled = bValue
        'Me.txtQtaProdotto.Enabled = bValue
        Me.chkDismessoProd.Enabled = bValue
        Me.txtDataDismesso.Enabled = bValue
        
        Me.CDArticoloProd.Enabled = bValue
        Me.cboTipoPeriodo.Enabled = bValue
        Me.cboUMPeriodoProd.Enabled = bValue
        
        'Me.txtDataInizioProd.Enabled = bValue
        'Me.txtOraInizioProd.Enabled = bValue
        
        'Me.txtDataFineProd.Enabled = bValue
        'Me.txtOraFineProd.Enabled = bValue
        Me.txtQtaPeriodo.Enabled = bValue
        
        Me.cboListinoProd.Enabled = bValue
        
        Me.cboUMArtProd.Enabled = bValue
        Me.cboIvaProd.Enabled = bValue
        
        Me.txtQtaArtProd.Enabled = bValue
        
        Me.txtImpUniProd.Enabled = bValue
        Me.txtSconto1Prod.Enabled = bValue
        Me.txtSconto2Prod.Enabled = bValue
        
        'Me.txtImponibileProd.Enabled = bValue
        Me.txtScontoImpProd.Enabled = bValue
        If Me.cboTipoImpostazione.CurrentID <= 2 Then
            Me.cboTipoRateizzazioneProd.Enabled = False
            Me.chkRinnovareProd.Enabled = bValue
            If chkGeneraRateProd.Value = vbChecked Then
                Me.chkGeneraUnaRataProd.Enabled = bValue
            Else
                Me.chkGeneraUnaRataProd.Enabled = False
            End If
        Else
            Me.cboTipoRateizzazioneProd.Enabled = bValue
            Me.chkRinnovareProd.Enabled = False
            Me.chkGeneraUnaRataProd.Enabled = False
        End If
        
        'Me.txtTotaleRigaProd.Enabled = bValue
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
    Me.cmdNuovoProd.Enabled = True
    Me.cmdSalvaProd.Enabled = bValue
    Me.cmdEliminaProd.Enabled = bValue
    Me.cmdNuovoProdDaEsis.Enabled = bValue
    If Me.cboTipoImpostazione.CurrentID < 3 Then
        Me.cmdNuovoProdDaEsis.Enabled = False
    End If

    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
            Me.cmdNuovoProd.Enabled = True
            Me.Command7.Enabled = True
            Me.cmdSalvaProd.Enabled = bValue
            Me.cmdEliminaProd.Enabled = bValue
            Me.cmdNuovoProdDaEsis.Enabled = bValue
            If Me.cboTipoImpostazione.CurrentID < 3 Then
                Me.cmdNuovoProdDaEsis.Enabled = False
            End If
    Else
        If Me.chkChiuso.Value = vbChecked Then
            Me.cmdNuovoProd.Enabled = False
            Me.Command7.Enabled = False
            Me.cmdSalvaProd.Enabled = False
            Me.cmdEliminaProd.Enabled = False
            Me.cmdNuovoProdDaEsis.Enabled = False
        Else
            Me.cmdNuovoProd.Enabled = True
            Me.Command7.Enabled = True
            Me.cmdSalvaProd.Enabled = bValue
            Me.cmdEliminaProd.Enabled = bValue
            Me.cmdNuovoProdDaEsis.Enabled = bValue
            If Me.cboTipoImpostazione.CurrentID < 3 Then
                Me.cmdNuovoProdDaEsis.Enabled = False
            End If
        End If
    End If
    
    bLoadingProdotti = False
End Sub



Private Sub txtAnnotazioni_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataDecorrenza_Change()
    txtNGGPrimaRata_Change
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataDisdetta_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataFineAssistenza_Change()
    If Not (BrwMain.Visible) Then Change
End Sub


Private Sub txtDataFineProd_LostFocus()
    If Me.cboUMPeriodoProd.CurrentID = 2 Then
        Me.txtQtaPeriodo.Value = DateDiff("d", Me.txtDataInizioProd.Text, Me.txtDataFineProd.Text) + 1
        txtQtaPeriodo_LostFocus
    End If
End Sub


Private Sub txtDataInizioProd_LostFocus()
    txtQtaPeriodo_LostFocus
End Sub

Private Sub txtDataPrimaDecorr_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataPrimaRata_LostFocus()
On Error GoTo ERR_txtDataPrimaRata_LostFocus
    

    If Me.txtDataDecorrenza.Value = 0 Then Exit Sub
    

    If Me.txtDataPrimaRata.Value = 0 Then
        txtNGGPrimaRata.Value = 0
        Exit Sub
    End If
    
    If Me.txtDataPrimaRata.Value < Me.txtDataDecorrenza.Value Then
        txtNGGPrimaRata.Value = 0
        Exit Sub
    End If
    

    If Me.txtDataPrimaRata.Value > 0 Then
        txtNGGPrimaRata.Value = DateDiff("d", Me.txtDataDecorrenza.Text, Me.txtDataPrimaRata.Text) + 1
    End If

    

Exit Sub
ERR_txtDataPrimaRata_LostFocus:
    MsgBox Err.Description, vbCritical, "txtDataPrimaRata_LostFocus"
End Sub

Private Sub txtDataScadenza_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataScadenzaPerRinnovo_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataStipula_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataStipula_LostFocus()
On Error GoTo ERR_txtDataStipula_LostFocus
    If fnNotNullN(m_Document(m_Document.PrimaryKey)) <= 0 Then
            
        IDAccordoCommerciale = GET_LINK_ACCORDO_COMMERCIALE(IDClienteFatturazione, Me.txtDataStipula.Text)
    
        If Me.txtDataDecorrenza.Value = 0 Then Me.txtDataDecorrenza.Value = Me.txtDataStipula.Value
        
        GET_CARATTERISTICHE_RISORSA Me.MSFlexGrid1
    End If

Exit Sub

ERR_txtDataStipula_LostFocus:
    MsgBox Err.Description, vbCritical, "txtDataStipula_LostFocus"
    
End Sub
Private Sub txtDescrizioneContratto_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtIDOggettoCollegato_Change()

    Me.txtOggettoCollegato.Text = fncDocumentoAllegato(Me.txtIDOggettoCollegato.Value)
    
    'If Me.txtIDOggettoCollegato.Value = 0 Then
    '    Me.chkRataFatturata.Value = 0
    'Else
    '    Me.chkRataFatturata.Value = 1
    'End If
    
End Sub

Private Sub txtIDProdotto_Change()
On Error GoTo ERR_txtIDProdotto_Change
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset



sSQL = "SELECT * FROM RV_POProdotto "
sSQL = sSQL & " WHERE IDRV_POProdotto=" & txtIDProdotto.Value

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.chkProdottoGenerico.Value = vbUnchecked
    Me.txtValIdentProd.Text = ""
    Me.txtProdotto.Text = ""
    
Else
    Me.chkProdottoGenerico.Value = Abs(fnNotNullN(rs!ProdottoGenerico))
    If fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0 Then
        Me.txtValIdentProd.Text = fnNotNull(rs!Matricola)
    End If
    Me.txtProdotto.Text = fnNotNull(rs!Descrizione)
    
    If fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value) <= 0 Then
        Me.txtQtaProdotto.Value = 1
        
        If (Me.cboTipoImpostazione.CurrentID <= 2) Then
            Me.cboTipoPeriodo.WriteOn 2
        Else
            Me.cboTipoPeriodo.WriteOn 1
        End If
        
        Me.cboUMPeriodoProd.WriteOn LINK_UM_PERIODO_AZIENDA
        If (fnNotNullN(rs!IDRV_POUnitaDiMisuraPeriodo)) > 0 Then
            Me.cboUMPeriodoProd.WriteOn fnNotNullN(rs!IDRV_POUnitaDiMisuraPeriodo)
        End If
        
        'Me.txtQtaPeriodo.Value = 1
        txtQtaPeriodo_LostFocus
        
        Me.cboListinoProd.WriteOn GET_LINK_LISTINO(IDClienteFatturazione, TheApp.IDFirm)
        Me.CDArticoloProd.Load fnNotNullN(rs!IDArticolo)
        
        
    End If
End If


rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_txtIDProdotto_Change:
    MsgBox Err.Description, vbCritical, "txtIDProdotto_Change"
End Sub

Private Sub txtIDProdottoRata_Change()
    Me.txtProdottoRata.Text = GET_PRODOTTO(Me.txtIDProdottoRata.Value)
End Sub

Private Sub txtIDProdRifContr_Change()
    cmdSelProdContratto.Enabled = True
    
    If txtIDProdRifContr.Value > 0 Then
        cmdSelProdContratto.Enabled = False
    End If
    
    Me.txtProdottoRata.Text = GET_PRODOTTO_CONTRATTO(Me.txtIDProdRifContr.Value)
End Sub

Private Sub txtImportoAttuale_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtImportoStipula_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Function fncDocumentoAllegato(IDOggetto As Long) As String
On Error GoTo ERR_fncDocumentoAllegato
    Dim sSQL As String
    Dim rsOgg As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT Oggetto, DataEmissione, Numero "
    sSQL = sSQL & "FROM Oggetto "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
    
    Set rsOgg = Cn.OpenResultset(sSQL)
    
    If rsOgg.EOF = True Then
        fncDocumentoAllegato = ""
    Else
        fncDocumentoAllegato = fnNotNull(rsOgg!Oggetto) & " N° " & fnNotNull(rsOgg!numero) & " del " & fnNotNull(rsOgg!DataEmissione)
    End If
    
    rsOgg.CloseResultset
    Set rsOgg = Nothing
    Exit Function
    
ERR_fncDocumentoAllegato:
    Me.txtIDOggettoCollegato.Text = ""
End Function
Private Function fncIDDocumentoAllegato(IDOggetto As Long) As Long
On Error GoTo ERR_fncDocumentoAllegato
Dim sSQL As String
Dim rsOgg As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoOggetto "
sSQL = sSQL & "FROM Oggetto "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto

Set rsOgg = Cn.OpenResultset(sSQL)

If rsOgg.EOF = True Then
    fncIDDocumentoAllegato = 0
Else
    fncIDDocumentoAllegato = fnNotNullN(rsOgg!IDTipoOggetto)
End If

rsOgg.CloseResultset
Set rsOgg = Nothing
Exit Function
    
ERR_fncDocumentoAllegato:
    fncIDDocumentoAllegato = 0
End Function



Private Sub txtImportoStipula_LostFocus()
    If fnNotNullN(m_Document(m_Document.PrimaryKey)) <= 0 Then
        If Me.cboTipoImpostazione.CurrentID = 1 Then
            If Me.txtImportoAttuale.Value = 0 Then Me.txtImportoAttuale.Value = Me.txtImportoStipula.Value
        End If
    End If
End Sub

Private Sub txtImpUniProd_LostFocus()
GET_TOTALI_RIGA_DETTAGLIO
End Sub

Private Sub txtNGGPrimaRata_Change()
On Error GoTo ERR_txtNGGPrimaRata_Change
    txtDataPrimaRata.Value = 0
    If Me.txtDataDecorrenza.Value = 0 Then Exit Sub
    
    If Me.txtNGGPrimaRata.Value > 0 Then
        Me.txtDataPrimaRata.Text = DateAdd("d", Me.txtNGGPrimaRata.Value, Me.txtDataDecorrenza.Text) - 1
    End If
    
    If Not (BrwMain.Visible) Then Change

Exit Sub
ERR_txtNGGPrimaRata_Change:
    MsgBox Err.Description, vbCritical, "txtNGGPrimaRata_Change"
End Sub

Private Sub txtNoteFatturazione_Change()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub txtNumeroProtocollo_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub AggiornaRataStoricaSingola(IDRiferimento As Long)
    Dim sSQL As String
    
    sSQL = "UPDATE RV_POStoriaRateContratto SET "
    sSQL = sSQL & "DataRata=" & fnNormDate(Me.txtDataRata.Text) & ", "
    sSQL = sSQL & "ImportoRata=" & fnNormNumber(Me.txtImportoRata.Text) & ", "
    sSQL = sSQL & "IDPagamentoRate=" & Me.cboPagamentoRataContratto.CurrentID & ", "
    sSQL = sSQL & "Mese=" & fnNormNumber(m_DocumentsLink("Mese").Value) & ", "
    sSQL = sSQL & "Anno=" & fnNormNumber(m_DocumentsLink("Anno").Value) & ", "
    sSQL = sSQL & "Periodo=" & fnNormString(m_DocumentsLink("Periodo").Value) & ", "
    sSQL = sSQL & "Adeguamento=" & fnNormBoolean(m_DocumentsLink("Adeguamento")) & ", "
    sSQL = sSQL & "Manuale=" & fnNormBoolean(m_DocumentsLink("Manuale")) & ", "
    sSQL = sSQL & "Fatturata=" & fnNormBoolean(m_DocumentsLink("Fatturata")) & ", "
    sSQL = sSQL & "ContrattoAttuale=" & fnNormBoolean(m_DocumentsLink("ContrattoAttuale")) & " "
    sSQL = sSQL & "WHERE IDRiferimentoRata=" & IDRiferimento
    
    Cn.Execute sSQL
    
End Sub
Private Sub INSERISCI_RATA_STORICA()
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

''''''''''''''''''''''TROVA CONTRATTO STORICO''''''''''''''''''''''''''''''''''''''''''''''
Link_StoriaContratto = ContrattoAttualeStorico(m_Document(m_Document.PrimaryKey).Value)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''ELIMINA RATE STORICHE'''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POStoriaRateContratto "
sSQL = sSQL & "WHERE IDRV_POStoriaContratto=" & Link_StoriaContratto
Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''INSERISCI RATE STORICHE''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
sSQL = sSQL & " ORDER BY NumeroRata ASC"

Set rs = Cn.OpenResultset(sSQL)

    While Not rs.EOF
                     
        sSQL = "INSERT INTO RV_POStoriaRateContratto ("
        sSQL = sSQL & "IDRV_POStoriaRateContratto, IDRV_POStoriaContratto, IDRiferimentoRata, NumeroRata, DataRata, "
        sSQL = sSQL & "IDPagamentoRate, ImportoRata, Mese, Anno, Periodo, "
        sSQL = sSQL & "Adeguamento, Manuale, ContrattoAttuale, IDOggettoCollegato, Fatturata) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnGetNewKey("RV_POStoriaRateContratto", "IDRV_POStoriaRateContratto") & ", "
        sSQL = sSQL & Link_StoriaContratto & ", "
        sSQL = sSQL & fnNotNullN(rs!IDRV_PORateContratto) & ", "
        sSQL = sSQL & fnNotNullN(rs!numerorata) & ", "
        sSQL = sSQL & fnNormDate(rs!DataRata) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDPagamentoRata) & ", "
        sSQL = sSQL & fnNormNumber(rs!ImportoRata) & ", "
        sSQL = sSQL & fnNormNumber(rs!mese) & ", "
        sSQL = sSQL & fnNormNumber(rs!Anno) & ", "
        sSQL = sSQL & fnNormString(rs!Periodo) & ", "
        sSQL = sSQL & fnNotNullN(rs!Adeguamento) & ", "
        sSQL = sSQL & fnNotNullN(rs!Manuale) & ", "
        sSQL = sSQL & fnNotNullN(rs!ContrattoAttuale) & ", "
        sSQL = sSQL & fnNotNullN(rs!IDOggettoCollegato) & ", "
        sSQL = sSQL & fnNotNullN(rs!Fatturata) & ")"
        
        Cn.Execute sSQL
        
    rs.MoveNext
    Wend

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub EliminaRataStoricaSingola()
Dim sSQL As String

sSQL = "DELETE FROM RV_POStoriaRateContratto "
sSQL = sSQL & "WHERE IDRiferimentoRata=" & fnNotNullN(Me.GrigliaRateContratto.AllColumns("IDRV_PORateContratto").Value)

Cn.Execute sSQL

End Sub
Private Sub fncSitoPerAnagrafica()
    Dim sSQL As String
    'Dim Link_Filiale As Long
    
    With Me.cboSitoPerAnagrafica
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT IDSitoPerAnagrafica, SitoPerAnagrafica FROM SitoPerAnagrafica WHERE IDAnagrafica=" & Me.CDCliente.KeyFieldID
        .Fill
    End With
    
    Me.cboSitoPerAnagrafica.WriteOn fnNotNullN(m_Document("IDSitoPerAnagrafica"))

End Sub
Private Function ControlloCambioTesta() As Boolean
    
    ControlloCambioTesta = False

    'If GET_RATA_PAGATA(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) = True Then
    '    ControlloCambioTesta = False
    '    Exit Function
    'End If

    If m_Document("IDRateizzazione") <> Me.cboTipoRateizzazione.CurrentID Then
        ControlloCambioTesta = True
        Exit Function
    End If
    
    If m_Document("IDTipoRinnovo") <> Me.cboTipoRinnovo.CurrentID Then
        ControlloCambioTesta = True
        Exit Function
    End If

    If m_Document("IDDurataContratto") <> Me.cboDurataContratto.CurrentID Then
        ControlloCambioTesta = True
        Exit Function
    End If
    
    If m_Document("DataDecorrenza") <> Me.txtDataDecorrenza.Text Then
        ControlloCambioTesta = True
        Exit Function
    End If

    If m_Document("ImportoContrattoAttuale") <> Me.txtImportoAttuale.Value Then
        ControlloCambioTesta = True
        Exit Function
    End If

    If m_Document("DataScadenza") <> Me.txtDataScadenza.Text Then
        ControlloCambioTesta = True
        Exit Function
    End If
    If m_Document("DataScadenzaPerRinnovo") <> Me.txtDataScadenzaPerRinnovo.Text Then
        ControlloCambioTesta = True
        Exit Function
    End If
    If m_Document("IDPagamentoRate") <> Me.cboPagamentoRate.CurrentID Then
        ControlloCambioTesta = True
        Exit Function
    End If
    If fnNotNullN(m_Document("NumeroGiorniPrimaRata").Value) <> Me.txtNGGPrimaRata.Value Then
        ControlloCambioTesta = True
        Exit Function
    End If

    
End Function
Private Sub Ripristina_Valori_Iniziali()
    Me.cboDurataContratto.WriteOn m_Document("IDDurataContratto")
    Me.txtDataDecorrenza.Text = m_Document("DataDecorrenza")
    Me.cboTipoRinnovo.WriteOn m_Document("IDTipoRinnovo")
    Me.cboTipoRateizzazione.WriteOn m_Document("IRateizzazione")
    Me.txtImportoAttuale.Value = m_Document("ImportoContrattoAttuale")
End Sub
Private Sub EliminazioneContrattoStorico()
'On Error GoTo ERR_EliminazioneContrattoStorico
    Dim sSQL As String
    
    sSQL = "DELETE FROM RV_POStoriaContratto WHERE IDRV_POStoriaContratto=" & Link_StoriaContratto
    
    Cn.Execute sSQL
Exit Sub
ERR_EliminazioneContrattoStorico:
    MsgBox Err.Description, vbCritical, "Eliminazione contratto storico"
End Sub
Private Sub EliminazioneRateContrattoStorico()
'On Error GoTo ERR_EliminazioneRateContrattoStorico
    Dim sSQL As String
    
    sSQL = "DELETE FROM RV_POStoriaRateContratto WHERE IDRV_POStoriaContratto=" & Link_StoriaContratto
    
    Cn.Execute sSQL
Exit Sub
ERR_EliminazioneRateContrattoStorico:
    MsgBox Err.Description, vbCritical, "Eliminazione rate contratto storico"
End Sub

Private Sub EliminazioneContrattoNonStandard(IDContratto As Long)
'On Error GoTo ERR_EliminazioneContrattoNonStandard
    Dim sSQL As String
    
    sSQL = "DELETE FROM RV_POContrattoNonStandard WHERE IDRV_POContratto=" & IDContratto
    Cn.Execute sSQL
Exit Sub
ERR_EliminazioneContrattoNonStandard:
    MsgBox Err.Description, vbCritical, "Eliminazione contratto non standard"
End Sub
'Private Function GetPianoDeiConti() As Long
''On Error GoTo ERR_GetPianoDeiConti
'    Dim sSQL As String
'    Dim rs As DmtOleDbLib.adoResultset
'    sSQL = "SELECT IDPianoDeiConti FROM PianoDeiConti WHERE ("
'    sSQL = sSQL & "(IDAzienda = " & m_App.Branch & ") AND "
'    sSQL = sSQL & "(TipoPDC = " & 1 & ") AND "
'    sSQL = sSQL & "(IDEsercizio= " & VarIDEsercizio & "))"
'
'    Set rs = Cn.OpenResultset(sSQL)
'
'    If rs.EOF = False Then
'        GetPianoDeiConti = fnNotNullN(rs!IDPianoDeiConti)
'    Else
'        GetPianoDeiConti = 0
'    End If
'
'    rs.CloseResultset
'    Set rs = Nothing
'Exit Function
'ERR_GetPianoDeiConti:
'    MsgBox Err.Description, vbCritical, "Errore piano dei conti"
'End Function
'
'Private Sub SetPDCProperties()
'Dim oNode As DmtPDC.INode
'Dim oBranch As DmtPDC.Branch
'Dim oNode1 As DmtPDC.INode
'    Set oPDC = New DmtPDC.PDCServices
'    'Imposta le proprietà dell'oggetto PDCServices
'    With oPDC
'        'Viene fornita al controllo la connessione al database DMT.
'        'La connessione è di tipo ADO.Connection quindi viene
'        'passata la proprietà InternalConnection dell'oggetto Database
'        Set .Connection = m_App.Database.InternalConnection
'        'Indica l'identificativo del Piano dei conti da visualizzare
'        .IDPDC = Link_PianoDeiConti
'        .HideAccounts = False
'        .BranchType = btcAllBranchs
'        '.BranchType = .BranchType + btcRevenuesBranch
'        .AccountType = atcAllAccounts
'        If Len(Me.txtCodiceConto.Text) > 0 Then
'            Set oNode = .SearchNodeExtended(Me.txtCodiceConto.Text)
'        Else
'            Set oNode = .SearchNodeExtended(, Me.txtDescrizioneConto.Text)
'        End If
'        If .RecordFounded = 1 Then
'                If TypeName(oNode) = "Account" Then
'
'                    Link_ContoPDC = oNode.ID
'                    'Codifica completa del Conto o del Ramo
'                    Me.txtCodiceConto = oNode.CompletedCode
'                    Me.txtDescrizioneConto.Text = oNode.Description
'
'                Else
'
'                    If Len(Me.txtCodiceConto.Text) > 0 Then
'                        .SelectedNode.CompletedCode = Me.txtCodiceConto.Text
'                    Else
'                        .SelectedNode.Description = Me.txtDescrizioneConto.Text
'                    End If
'
'                    .ShowSearchDialog
'
'                    ShowNodeProperties oPDC.SelectedNode
'                End If
'
'        ElseIf .RecordFounded > 1 Then
'            If Len(Me.txtCodiceConto.Text) > 0 Then
'                .SelectedNode.CompletedCode = Me.txtCodiceConto.Text
'            Else
'                .SelectedNode.Description = Me.txtDescrizioneConto.Text
'            End If
'
'            .ShowSearchDialog
'
'            ShowNodeProperties oPDC.SelectedNode
'
'        Else
'            If Len(Me.txtCodiceConto.Text) > 0 Then
'                .SelectedNode.CompletedCode = Me.txtCodiceConto.Text
'            Else
'                .SelectedNode.Description = Me.txtDescrizioneConto.Text
'            End If
'
'
'            .ShowSearchDialog
'
'            ShowNodeProperties oPDC.SelectedNode
'
'        End If
'    End With
'
'    Set oPDC = Nothing
'
'End Sub
'Private Sub ShowNodeProperties(ByVal oNode As DmtPDC.INode)
'    'Rappresenta un conto
'    Dim oAccount As DmtPDC.Account
'    'Rappresenta un ramo
'    Dim oBranch As DmtPDC.Branch
'
'    'Vengono visualizzati nei campi appositi tutte
'    'le caratteristiche del conto o ramo selezionato
'
'    'Controlla se è stato passato un elemento valido
'    If Not oNode Is Nothing Then
'        'Riporta i dati comuni del conto o del ramo
'
'        'Identificativo unico del Conto o del Ramo
'        Link_ContoPDC = oNode.ID
'        'Codifica completa del Conto o del Ramo
'        Me.txtCodiceConto = oNode.CompletedCode
'        Me.txtDescrizioneConto.Text = oNode.Description
'    End If
'End Sub
Private Function fnUtentePerInserimentoContratto() As String
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT Utente FROM Utente WHERE IDUtente=" & fnNotNullN(m_Document("IDUtentePerInserimento").Value)
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        fnUtentePerInserimentoContratto = "Il contratto è stato inserito dall'utente " & fnNotNull(rs!Utente) & " in data " & m_Document("DataInserimento").Value
    Else
        fnUtentePerInserimentoContratto = "Non è stato possibile recuperare le informazioni sull'utente che ha inserito il contratto"
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Function fnUtentePerModificaContratto() As String
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT Utente FROM Utente WHERE IDUtente=" & fnNotNullN(m_Document("IDUtentePerModifica").Value)
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        fnUtentePerModificaContratto = "L'ultimo aggiornamento del contratto è stato effettuato dall'utente " & fnNotNull(rs!Utente) & " in data " & m_Document("DataModifica").Value
    Else
        fnUtentePerModificaContratto = "Non è stato possibile recuperare le informazioni sull'utente dell'ultimo aggiornamento del contratto"
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function






Private Sub BarMenu_BandClose(ByVal Band As ActiveBar3LibraryCtl.Band)
     'Se la banda è una Toolbar allora viene registrata la chiusura.
    If Band.Type = ddBTNormal And Band.Name <> BAND_CLOSE_PREVIEW Then
        
        'Salva nel registry l'impostazione sulla visibilità della toolbar
        AppOptions.ToolbarVisibility(Band.Name) = False
        
    End If
End Sub

Private Sub BarMenu_BandMove(ByVal Band As ActiveBar3LibraryCtl.Band)
    Form_Resize
End Sub

Private Sub BarMenu_BandOpen(ByVal Band As ActiveBar3LibraryCtl.Band, ByVal Cancel As ActiveBar3LibraryCtl.ReturnBool)
     'Se la banda è una Toolbar allora viene registrata l'apertura.
    If Band.Type = ddBTNormal And Band.Name <> BAND_CLOSE_PREVIEW Then
        AppOptions.ToolbarVisibility(Band.Name) = True
    End If
End Sub

Private Sub BarMenu_MenuItemEnter(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    WriteStatusBar Tool.Description
End Sub

Private Sub BarMenu_MenuItemExit(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    WriteStatusBar ""
End Sub

Private Sub BarMenu_MouseEnter(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    WriteStatusBar Tool.Description
End Sub

Private Sub BarMenu_MouseExit(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    WriteStatusBar ""
End Sub
Private Sub BarMenu_QueryUnload(Cancel As Integer)
    Cancel = True
End Sub

Private Sub BarMenu_Resize(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    Form_Resize
End Sub

Private Sub BarMenu_ToolClick(ByVal Tool As ActiveBar3LibraryCtl.Tool)
On Error Resume Next
    Dim iKeyCode As Integer
    Dim iShift As Integer
    Dim bContinue As Boolean
    
    'On Error GoTo BarMenu_ClickError
        
    'Forza il lostfocus ed attende l'esecuzione di eventuali eventi associati
    AutoLostFocus
        
    bContinue = True
    iShift = GetShift(Tool)
    iKeyCode = GetKeyCode(Tool)
    
    If iKeyCode <> 0 Or iShift <> 0 Then
        bContinue = Not ShortCut(iKeyCode, iShift)
        If bContinue Then
            SendKeys GetSendKeys(Tool) & "(" & GetKey(Tool) & ")"      '"^(R)"
        End If
    Else
        ExecuteMenuCommand Tool.Name
    End If
    
    Exit Sub
    
BarMenu_ClickError:
    Screen.MousePointer = 0
    
    If Err.Number = ERR_NDELFILTER Then
        'In seguito a particolari sequenze di eventi può risultare abilitato il cancella filtro sul
        'filtro di default. Se si esegue la cancellazione viene sollevata una eccezione.
        sbMsgError "Non è possibile eliminare il filtro di default.", m_App.FunctionName
    Else
        sbMsgError Err.Description, m_App.FunctionName
    End If
    
    Resume Next
End Sub
'**+
'Nome: SendDocument
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue l'esportazione del documento con controllo di errore
'**/
Private Sub SendDocument(ByVal Appl As Long)
    On Error GoTo errHandler
    Dim EmailCliente As String
    
    Dim OLDCursor As Integer
    
    OLDCursor = Screen.MousePointer
    
    EmailCliente = ""
    
    Screen.MousePointer = vbHourglass
    EmailCliente = GET_INDIRIZZO_EMAIL_CLIENTE(Me.CDCliente.KeyFieldID)
    
    If Len(Trim(fnNotNull(EmailCliente))) = 0 Then
        m_Report.SendMail Appl
    Else
        m_Report.SendMailTo Appl, EmailCliente
    End If
    
    'm_Document.SendMail m_Report, Appl
    Screen.MousePointer = OLDCursor
    Exit Sub
errHandler:
    Screen.MousePointer = OLDCursor
    
    If Err.Number = 20507 Then
        'Errore "Invalid file Name" generato quando non è possibile trovare il file .rpt
        sbMsgInfo "File di report non trovato", m_App.FunctionName
    Else
        sbMsgInfo Err.Description, m_App.FunctionName
    End If
End Sub
'**+
'Autore                     : Diamante s.p.a
'Data creazione             :
'Nome                       : InitSemaphore
'
'Parametri                  :
'
'Funzionalità               : Attiva/disattiva le attività del Riquadro attività
'
'**/
Private Sub EnableDOMActivitiesItems()
    oFiltersActivity.EnableItems (BrwMain.GuiMode = dgNormal And BrwMain.Visible)
    oTableViewsActivity.EnableItems (BrwMain.GuiMode = dgNormal And BrwMain.Visible)
    
    ActivityBox.Redraw = True
End Sub
'**+
'Autore                     : Diamante s.p.a
'Data creazione             :
'Nome                       : ActivityBox_CloseButtonPressed
'
'Parametri                  :
'
'Funzionalità               : Gestione della chiusura del Riquadro attività
'
'**/
Private Sub ActivityBox_CloseButtonPressed()
    OnFolders
End Sub

'**+
'Autore                     : Diamante s.p.a
'Data creazione             :
'Nome                       : ActivityBox_ItemSelected
'
'Parametri                  :
'
'Funzionalità               : Gestione della selezione delle voci del Riquadro attività
'
'**/
Private Sub ActivityBox_ItemSelected(ByVal Item As DmtActBoxTlb.Item, NeedRedraw As Boolean)
    Dim oFilter As Filter
    Dim oTableView As TableView
    
    If BrwMain.Visible And BrwMain.GuiMode = dgNormal Then
        Select Case ActivityBox.CurrentActivity.Caption
            Case "Filtri"
                For Each oFilter In m_DocType.Filters
                    If oFilter.ID = Val(Item.Tag) Then
                        Set m_ActiveFilter = m_DocType.Filters(oFilter.Name)
                        Exit For
                    End If
                Next
                'Flag usato per specificare che deve essere eseguito un filtro permanente.
                m_FilterSelected = True
                
                '---Modalità Filtro---
                'Ripulisce i campi di immissione delle condizioni di ricerca.
                BrwMain.Conditions.ClearValues
                
                'Se attivo, viene disabilitato il pulsante Salva Filtro del DocTypeExplorer.
                oFiltersActivity.AbortNewFilter
                ActivityBox.Redraw = True
                
                'Viene eseguita la ricerca basata sul nuovo filtro.
                ExecuteSearch

            Case "Viste tabellari"
                For Each oTableView In m_DocType.TableViews
                    If oTableView.ID = Val(Item.Tag) Then
                        Set m_ActiveTableView = m_DocType.TableViews(oTableView.Name)
                        Exit For
                    End If
                Next
                BrwMain.LoadColumns m_ActiveTableView
                SetVisibilityIDFields
        End Select
    End If
    If ActivityBox.CurrentActivity.Caption = "Esportazioni" Then
        If Item.Hyperlink Then
            Select Case Item.Name
                Case "E" & ExportConstants.PDF
                    ExecuteMenuCommand "ExportPDF"
                Case "E" & ExportConstants.Word
                    ExecuteMenuCommand "ExportWord"
                Case "E" & ExportConstants.Excel
                    ExecuteMenuCommand "ExportExcel"
                Case "E" & ExportConstants.HTML
                    ExecuteMenuCommand "ExportHtml"
                Case "S" & ExportConstants.PDF
                    ExecuteMenuCommand "MailPDF"
                Case "S" & ExportConstants.Word
                    ExecuteMenuCommand "MailWord"
                Case "S" & ExportConstants.Excel
                    ExecuteMenuCommand "MailExcel"
                Case "S" & ExportConstants.HTML
                    ExecuteMenuCommand "MailHtml"
            End Select
        End If
    End If
End Sub
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width, ActivityBox.Height
        picSplitter.AutoRedraw = True
    End With
    picSplitter.Visible = True
    m_SplitterMoving = True
    picSplitter.ZOrder
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single

    If m_SplitterMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < SPLITLIMIT Then
            picSplitter.Left = SPLITLIMIT
        ElseIf sglPos > BarMenu.ClientAreaWidth - SPLITLIMIT Then
            picSplitter.Left = BarMenu.ClientAreaWidth - SPLITLIMIT
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActivityBox.Width = picSplitter.Left - ActivityBox.Left
    FormRecalcLayout
    picSplitter.Visible = False
    m_SplitterMoving = False
End Sub

Private Sub GET_CONTROLLO_LICENZA()
Dim sSQL As String
Dim Codice_Diamante As String
Dim Codice_Prodotto_Calcolato  As String
Dim Codice_Attivazione As String

Dim Partita_Iva_Licenza As String

Codice_Diamante = GET_CODICE_DIAMANTE
Partita_Iva_Licenza = GET_PARTITA_IVA
Codice_Attivazione = GET_CODICE_SBLOCCO_ATTIVAZIONE


Codice_Prodotto_Calcolato = GET_CODICE_SBLOCCO(Codice_Diamante, Partita_Iva_Licenza, "51")

If Codice_Attivazione = Codice_Prodotto_Calcolato Then
    DEMO = False
Else
    DEMO = True
End If

    
End Sub
Private Function GET_CODICE_DIAMANTE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Descrizione FROM ComponenteSwAbilitata "
sSQL = sSQL & "WHERE NomeCompSW=" & fnNormString("*IDSW___")

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICE_DIAMANTE = ""
Else
    GET_CODICE_DIAMANTE = Trim(fnNotNull(rs!Descrizione))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PARTITA_IVA() As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PartitaIva FROM RV_POProgramma "
sSQL = sSQL & "WHERE IDRV_POProgramma=51"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PARTITA_IVA = ""
Else
    GET_PARTITA_IVA = fnCryptString(Trim(fnNotNull(rs!PartitaIVA)))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CODICE_SBLOCCO_ATTIVAZIONE() As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CodiceSblocco FROM RV_POProgramma "
sSQL = sSQL & "WHERE IDRV_POProgramma=51"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICE_SBLOCCO_ATTIVAZIONE = ""
Else
    GET_CODICE_SBLOCCO_ATTIVAZIONE = fnCryptString(Trim(fnNotNull(rs!CodiceSblocco)))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_NUMERO_INSERIMENTI() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Count(IDRV_POContratto) As NumeroInserimenti "
sSQL = sSQL & "FROM RV_POContratto "
sSQL = sSQL & " WHERE IDAzienda=" & m_App.IDFirm


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_INSERIMENTI = 0
Else
    GET_NUMERO_INSERIMENTI = fnNotNullN(rs!NumeroInserimenti)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub cboCriterioRicorrenza_Click()

If BLoading = 1 Then Exit Sub

    Select Case Me.cboCriterioRicorrenza.CurrentID
        Case 1
            Me.txtOgniNumeroGiorni.Enabled = True
            Me.txtOgniNumeroMesi.Enabled = True
            Me.txtOgniNumeroSettimane.Enabled = False
            
            Me.txtOgniNumeroSettimane.Value = 0
        Case 2
            Me.txtOgniNumeroGiorni.Enabled = False
            Me.txtOgniNumeroMesi.Enabled = True
            Me.txtOgniNumeroSettimane.Enabled = False
            
            Me.txtOgniNumeroGiorni.Value = 0
            Me.txtOgniNumeroSettimane.Value = 0
        Case 3
            Me.txtOgniNumeroGiorni.Enabled = False
            Me.txtOgniNumeroMesi.Enabled = False
            Me.txtOgniNumeroSettimane.Enabled = True
            
            Me.txtOgniNumeroGiorni.Value = 0
            Me.txtOgniNumeroMesi.Value = 0
        Case Else
            Me.txtOgniNumeroGiorni.Enabled = True
            Me.txtOgniNumeroMesi.Enabled = True
            Me.txtOgniNumeroSettimane.Enabled = True
            
            Me.txtOgniNumeroGiorni.Value = 0
            Me.txtOgniNumeroMesi.Value = 0
            Me.txtOgniNumeroSettimane.Value = 0
    End Select


    

End Sub

Private Sub cboTipoDataFineRic_Click()
On Error GoTo ERR_cboTipoDataFineRic_Click

If BLoading = 1 Then Exit Sub
    Me.txtGiornoFissoFineRic.Enabled = False
    Me.txtMeseFissoFineRic.Enabled = False
    Me.cboTipoAnnoFineRicorr.Enabled = False
    
    If Me.cboTipoDataFineRic.CurrentID = 4 Then
        Me.txtGiornoFissoFineRic.Enabled = True
        Me.txtMeseFissoFineRic.Enabled = True
        Me.cboTipoAnnoFineRicorr.Enabled = True
    Else
        Me.txtGiornoFissoFineRic.Value = 0
        Me.txtMeseFissoFineRic.Value = 0
        Me.cboTipoAnnoFineRicorr.WriteOn 0
    End If

Exit Sub
ERR_cboTipoDataFineRic_Click:
    MsgBox Err.Description, vbCritical, "cboTipoDataFineRic_Click"
End Sub

Private Sub cboTipoDataInizioRic_Click()
On Error GoTo ERR_cboTipoDataInizioRic_Click

If BLoading = 1 Then Exit Sub
    Me.txtGiornoFissoInizioRic.Enabled = False
    Me.txtMeseFissoInizioRic.Enabled = False
    Me.cboTipoAnnoInizioRicorr.Enabled = False
    
    If Me.cboTipoDataInizioRic.CurrentID = 3 Then
        Me.txtGiornoFissoInizioRic.Enabled = True
        Me.txtMeseFissoInizioRic.Enabled = True
        Me.cboTipoAnnoInizioRicorr.Enabled = True
    Else
        Me.txtGiornoFissoInizioRic.Value = 0
        Me.txtMeseFissoInizioRic.Value = 0
        Me.cboTipoAnnoInizioRicorr.WriteOn 0
    End If
   
Exit Sub

ERR_cboTipoDataInizioRic_Click:
    MsgBox Err.Description, vbCritical, "cboTipoDataInizioRic_Click"
   
End Sub
Private Function PermissionToSaveRiga() As Boolean

    '///////////////////////////////////////////////////////////////////
    'Inserire qui il codice di controllo sulla validità e consistenza
    'dei dati da salvare.
    '///////////////////////////////////////////////////////////////////
    
    PermissionToSaveRiga = True
    If Me.CDServizio.KeyFieldID = 0 Then
        MsgBox "Inserire un codice articolo", vbCritical, "Controllo salvataggio dati"
        PermissionToSaveRiga = False
        Me.CDServizio.SetFocus
        Exit Function
    End If
    
    If Me.chkContrattoAttuale.Value = vbUnchecked Then
        MsgBox "Impossibile salvare poichè non risulta che il contratto sia quello attuale", vbInformation, "Impossibile salvare"
        PermissionToSaveRiga = False
        Exit Function
    End If
    Select Case Me.cboCriterioRicorrenza.CurrentID
        Case 1
            If Me.txtOgniNumeroGiorni.Value = 0 Then
                MsgBox "Il numero dei giorni di ricorrenza deve essere valorizzato", vbCritical, "Controllo salvataggio dati"
                PermissionToSaveRiga = False
                Me.txtOgniNumeroGiorni.SetFocus
                Exit Function
            End If
        Case 2
            If Me.txtOgniNumeroMesi.Value = 0 Then
                MsgBox "Il numero di mesi di ricorrenza deve essere valorizzato", vbCritical, "Controllo salvataggio dati"
                PermissionToSaveRiga = False
                Me.txtOgniNumeroMesi.SetFocus
                Exit Function
            End If
        Case 3
            If Me.txtOgniNumeroSettimane.Value = 0 Then
                MsgBox "Il numero di settimane di ricorrenza deve essere valorizzato", vbCritical, "Controllo salvataggio dati"
                PermissionToSaveRiga = False
                Me.txtOgniNumeroSettimane.SetFocus
                Exit Function
            End If
    End Select
    
    If Me.cboTipoImpostazione.CurrentID <> 3 Then
'        If Me.cboTipoDataInizioRic.CurrentID = 0 Then
'            MsgBox "Il tipo data di inizio ricorrenza deve essere valorizzato", vbCritical, "Controllo salvataggio dati"
'            PermissionToSaveRiga = False
'            Me.cboTipoDataInizioRic.SetFocus
'            Exit Function
'        End If
    End If
    
    If Me.cboTipoDataInizioRic.CurrentID = 3 Then
        If Me.txtGiornoFissoInizioRic.Value = 0 Then
            MsgBox "Il giorno fisso di inizio ricorrenza deve essere valorizzato", vbCritical, "Controllo salvataggio dati"
            PermissionToSaveRiga = False
            Me.txtGiornoFissoInizioRic.SetFocus
            Exit Function
        End If
        If Me.txtMeseFissoInizioRic.Value = 0 Then
            MsgBox "Il mese fisso di inizio ricorrenza deve essere valorizzato", vbCritical, "Controllo salvataggio dati"
            PermissionToSaveRiga = False
            Me.txtMeseFissoInizioRic.SetFocus
            Exit Function
        End If
    End If
    If Me.cboTipoDataFineRic.CurrentID = 4 Then
        If Me.txtGiornoFissoFineRic.Value = 0 Then
            MsgBox "Il giorno fisso di fine ricorrenza deve essere valorizzato", vbCritical, "Controllo salvataggio dati"
            PermissionToSaveRiga = False
            Me.txtGiornoFissoFineRic.SetFocus
            Exit Function
        End If
        If Me.txtMeseFissoFineRic.Value = 0 Then
            MsgBox "Il mese fisso di fine ricorrenza deve essere valorizzato", vbCritical, "Controllo salvataggio dati"
            PermissionToSaveRiga = False
            Me.txtMeseFissoFineRic.SetFocus
            Exit Function
        End If
    End If

    If Me.cboTipoDataFineRic.CurrentID = 0 Then
        If Me.txtNumeroRicorrenze.Value = 0 Then
            MsgBox "Il numero di ricorrenza deve essere valorizzato se non si specifica un tipo di data di fine ricorrenza", vbCritical, "Controllo salvataggio dati"
            PermissionToSaveRiga = False
            Me.txtNumeroRicorrenze.SetFocus
            Exit Function
             
        End If
    End If
    

End Function

Private Sub SCRIVI_CODA(IDOggetto As Long, IDTipoOggetto As Long)
Dim rs As ADODB.Recordset
Dim sSQL As String

'''''''''''''''''ELIMINAZIONE DATI UTENTE PER IL TIPO OGGETTO'''''''''''''''''''

sSQL = "DELETE FROM RV_POTMP "
sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
'sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID

Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set rs = New ADODB.Recordset

rs.Open "RV_POTMP", Cn.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    'rs!IDSessione = fnGetNewKey("RV_POTMP", "IDSessione")
    rs!IDUtente = m_App.IDUser
    rs!IDTipoOggetto = IDTipoOggetto
    rs!IDOggetto = IDOggetto
    rs!Utente = m_App.User
rs.Update

rs.Close
Set rs = Nothing

End Sub
Private Function GET_NUMERO_DOCUMENTO(NuovoDocumento As Boolean, IDTipoOggetto As Long) As Long
On Error GoTo ERR_GET_NUMERO_DOCUMENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim X_FRM As Form
Dim OLD_Cursor As Long
Dim CodiceLotto_OLD As String
Dim AnnoContratto As String

GET_NUMERO_DOCUMENTO = 0

sSQL = "SELECT * FROM RV_POTMP "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " ORDER BY IDSessione, IDUtente"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!IDUtente) = m_App.IDUser Then
        Me.Caption = "SALVATAGGIO IN CORSO.........."
        
        DoEvents
        
        
        If APERTURA_FORM_CODA = True Then
            Unload frmCoda
            APERTURA_FORM_CODA = False
        End If
        
        
        
        If NuovoDocumento = True Then
            If Me.txtDataDecorrenza.Value > 0 Then
                Me.txtAnnoContratto.Value = Year(Me.txtDataDecorrenza.Text)
                Me.txtNumeroContratto.Value = GET_NUMERO_CONTRATTO(Me.txtAnnoContratto.Value)
            Else
                Me.txtAnnoContratto.Value = Year(Date)
                Me.txtNumeroContratto.Value = GET_NUMERO_CONTRATTO(Me.txtAnnoContratto.Value)
            End If
            
            AnnoContratto = Me.txtAnnoContratto.Text & "-" & Me.txtNumeroContratto.Text
            
            
            m_Document("IDOggetto").Value = GET_LINK_OGGETTO_CONTRATTO(fnNotNullN(m_Document("IDOggetto").Value), m_DocType.ID, AnnoContratto, Me.txtDataDecorrenza)
            m_Document("AnnoContratto").Value = Me.txtAnnoContratto.Value
            m_Document("NumeroContratto").Value = Me.txtNumeroContratto.Value
            
        Else
            If fnNotNullN(m_Document("IDOggetto").Value) = 0 Then
                AnnoContratto = Me.txtAnnoContratto.Text & "-" & Me.txtNumeroContratto.Text
                m_Document("IDOggetto").Value = GET_LINK_OGGETTO_CONTRATTO(fnNotNullN(m_Document("IDOggetto").Value), m_DocType.ID, AnnoContratto, txtDataDecorrenza.Text)
            End If
            
        End If
        
        GET_NUMERO_DOCUMENTO = 1
        
        rs.CloseResultset
        Set rs = Nothing
    Else
        rs.CloseResultset
        Set rs = Nothing
    
        If APERTURA_FORM_CODA = False Then
            APERTURA_FORM_CODA = True
            Me.Enabled = False
            frmCoda.Show
        End If
        
        Me.Caption = "ATTENDERE......."
        DoEvents
        'GET_NUMERO_DOCUMENTO NuovoDocumento
        
    End If
End If
Exit Function

ERR_GET_NUMERO_DOCUMENTO:
    MsgBox Err.Description, vbCritical, "Errore coda"
    GET_NUMERO_DOCUMENTO = -1
    Unload frmCoda
End Function
Private Sub ELABORAZIONE_INTERVENTI_PER_SERVIZIO(IDContrattoServizio As Long, IDAnagraficaCliente As Long, IDAnagraficaTecnicoRif As Long, IDContratto As Long, IDContrattoPadre As Long, IDArticolo As Long, _
IDTipoRicorrenza As Integer, GiornoRic As String, MeseRic As String, SettimanaRic As String, _
IDTipoDataInizioRic As Integer, GiorniInizioRic As String, MeseInizioRic As String, _
IDTipoDataFineRic As String, GiornoFineRic As String, MeseFineRic As String, _
NumeroRicorrenze As String, IDTipoAnaTecRif As Long, IDTipoStatoInt As Long, IDTipoStatoFase As Long, _
IDTipoFase As Long, DescrizioneArticolo As String, IDTipoAnaTecOpe As Long, IDCategoriaIntervento As Long, IDAnagraficaFatturazione As Long)


'''''''''''''''DICHIARAZIONE DELLE VARIABILI''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim DataInizioServizio As String
Dim DataInizioPersonalizzata As String
Dim DataFineServizio As String
Dim DataFinePersonalizzata As String
Dim X_Ricorrenza As Long
Dim I As Integer
Dim oItem As MSComctlLib.ListItem
Dim NumeroIntervento As Long

Dim sSQL As String
Dim rsIntervento As ADODB.Recordset
Dim rsFase As ADODB.Recordset
Dim LINK_INTERVENTO As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'CALCOLO DELLA DATA INIZIO RICORRENZA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Select Case IDTipoDataInizioRic
    Case 1
        DataInizioServizio = DateAdd("m", IIf((MeseRic = ""), 0, MeseRic), Me.txtDataDecorrenza.Text)
        DataInizioServizio = DateAdd("d", IIf((GiornoRic = ""), 0, GiornoRic), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((SettimanaRic = ""), 0, SettimanaRic), DataInizioServizio)
        DataInizioServizio = DateAdd("d", -1, DataInizioServizio)
    Case 2
        DataInizioServizio = DateAdd("m", IIf((MeseRic = ""), 0, MeseRic), Me.txtDataDecorrenza.Text)
        DataInizioServizio = DateAdd("d", IIf((GiornoRic = ""), 0, GiornoRic), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((SettimanaRic = ""), 0, SettimanaRic), DataInizioServizio)
    Case 3
        DataInizioPersonalizzata = GET_COSTRUZIONE_DATA_PERS(GiorniInizioRic, MeseInizioRic) & Year(Me.txtDataDecorrenza.Text)
        If (GiorniInizioRic = GiornoFineRic) And (MeseInizioRic = MeseFineRic) Then
            DataInizioPersonalizzata = GET_COSTRUZIONE_DATA_PERS(GiorniInizioRic, MeseInizioRic) & Year(Me.txtDataFineAssistenza.Text)
        End If
        DataInizioServizio = DateAdd("m", IIf((MeseRic = ""), 0, MeseRic), DataInizioPersonalizzata)
        DataInizioServizio = DateAdd("d", IIf((GiornoRic = ""), 0, GiornoRic), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((SettimanaRic = ""), 0, SettimanaRic), DataInizioServizio)
    Case Else
        DataInizioServizio = DateAdd("m", IIf((MeseRic = ""), 0, MeseRic), Me.txtDataDecorrenza.Text)
        DataInizioServizio = DateAdd("d", IIf((GiornoRic = ""), 0, GiornoRic), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((SettimanaRic = ""), 0, SettimanaRic), DataInizioServizio)
        DataInizioServizio = DateAdd("d", -1, DataInizioServizio)
End Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'CALCOLO DELLA DATA DI FINE RICORRENZA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Select Case IDTipoDataFineRic
    Case 1
        DataFineServizio = Me.txtDataScadenza.Text
    Case 2
        DataFineServizio = Me.txtDataScadenzaPerRinnovo.Text
        
    Case 3
        DataFineServizio = Me.txtDataFineAssistenza.Text
    Case 4
        DataFineServizio = GET_COSTRUZIONE_DATA_PERS(GiornoFineRic, MeseFineRic) & Year(Me.txtDataFineAssistenza.Text)
        
    Case Else
        DataFineServizio = ""
End Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CALCOLO NUMERO RICORRENZE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If (Len(NumeroRicorrenze) = 0) Or (NumeroRicorrenze = "0") Then
    X_Ricorrenza = 0
Else
    X_Ricorrenza = NumeroRicorrenze
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set rsIntervento = New ADODB.Recordset
Set rsFase = New ADODB.Recordset

sSQL = "SELECT * FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=0"

rsIntervento.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic


If X_Ricorrenza > 0 Then
    For I = 1 To X_Ricorrenza
        
        If DataFineServizio <> "" Then
            If DateDiff("d", DataFineServizio, DataInizioServizio) > 0 Then
                Exit For
            End If
        End If
        
        rsIntervento.AddNew
            LINK_INTERVENTO = fnGetNewKey("RV_POIntervento", "IDRV_POIntervento")
            rsIntervento!IDRV_POIntervento = LINK_INTERVENTO
            rsIntervento!IDRV_POInterventoPadre = LINK_INTERVENTO
            rsIntervento!IDRV_POContrattoServizi = IDContrattoServizio
            rsIntervento!IDRV_POContratto = IDContratto
            rsIntervento!IDRV_POContrattoPadre = IDContrattoPadre
            rsIntervento!IDAnagraficaFatturazione = IDAnagraficaFatturazione
            rsIntervento!IDAnagraficaCliente = IDAnagraficaCliente
            rsIntervento!IDAzienda = TheApp.IDFirm
            rsIntervento!IDFiliale = TheApp.Branch
            rsIntervento!NumeroIntervento = GET_NUMERO_INTERVENTO(Year(Date))
            rsIntervento!AnnoIntervento = Year(Date)
            rsIntervento!IDArticolo = IDArticolo
            rsIntervento!IDAnagraficaTecnicoRif = IDAnagraficaTecnicoRif
            rsIntervento!IDTipoAnagraficaTecnicoRif = IDTipoAnaTecRif
            rsIntervento!IDRV_POStatoIntervento = IDTipoStatoInt
            rsIntervento!InterventoChiuso = 0
            rsIntervento!DataInserimento = Date
            rsIntervento!OraInserimento = GET_ORARIO(Now)
            rsIntervento!IDUtenteInserimento = TheApp.IDUser
            rsIntervento!DataUltimaModifica = Date
            rsIntervento!OraUltimaModifica = GET_ORARIO(Now)
            rsIntervento!IDUtenteUltimaModifica = TheApp.IDUser
            rsIntervento!Elaborata = 1
            rsIntervento!Manuale = 0
            rsIntervento!Richiesta = DescrizioneArticolo
            rsIntervento!NomeComputerInserimento = GET_NOMECOMPUTER
            rsIntervento!UtenteComputerInserimento = GET_NOMEUTENTE
            rsIntervento!NomeComputerModifica = GET_NOMECOMPUTER
            rsIntervento!UtenteComputerModifica = GET_NOMEUTENTE
            rsIntervento!NumeroFase = 1
            rsIntervento!IDRV_POTipoFaseIntervento = IDTipoFase
            rsIntervento!IDAnagraficaTecnicoOperativo = IDAnagraficaTecnicoRif
            rsIntervento!IDTipoAnagraficaTecnicoOpe = IDTipoAnaTecOpe
            rsIntervento!DataAppuntamento = DataInizioServizio
            rsIntervento!OraAppuntamento = "09.00"
            rsIntervento!LavoroEseguito = DescrizioneArticolo
            rsIntervento!Annotazioni = ""
            
            rsIntervento!IDRV_POStagione = GET_LINK_STAGIONE(DataInizioServizio)
            rsIntervento!IDRV_POCategoriaIntervento = GET_PARAMETRI_TEC_OPE(fnNotNullN(rsIntervento!IDAnagraficaTecnicoOperativo), "IDRV_POCategoriaFase")
            rsIntervento!IDRV_POTipoAddebito = GET_PARAMETRI_TEC_OPE(fnNotNullN(rsIntervento!IDAnagraficaTecnicoOperativo), "IDRV_POTipoAddebito")
            rsIntervento!IDRV_POTipoClasseIntervento = GET_PARAMETRI_TEC_OPE(fnNotNullN(rsIntervento!IDAnagraficaTecnicoOperativo), "IDRV_POTipoClasseIntervento")
            If fnNotNullN(rsIntervento!IDRV_POCategoriaIntervento) = 0 Then
                rsIntervento!IDRV_POCategoriaIntervento = IDCategoriaIntervento
            End If
            
        rsIntervento.Update
        
        If CREA_APPUNTAMENTO_AGENDA = 1 Then
            SCRIVI_APPUNTAMENTO fnNotNull(rsIntervento!DataAppuntamento), fnNotNull(rsIntervento!OraAppuntamento), LINK_INTERVENTO
        End If
        
        
        DataInizioServizio = DateAdd("m", IIf((MeseRic = ""), 0, MeseRic), DataInizioServizio)
        DataInizioServizio = DateAdd("d", IIf((GiornoRic = ""), 0, GiornoRic), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((SettimanaRic = ""), 0, SettimanaRic), DataInizioServizio)
        DataInizioServizio = DateAdd("d", -1, DataInizioServizio)
        
    Next
End If

If (X_Ricorrenza = 0) And (Len(DataFineServizio) > 0) Then
    While Not DateDiff("d", DataFineServizio, DataInizioServizio) > 0
        'Set oItem = Me.LVElaborazione.ListItems.Add
        'oItem.Text = Me.LVElaborazione.ListItems.Count
        'oItem.SubItems(1) = Servizio
        'oItem.SubItems(2) = NumeroIntervento
        'oItem.SubItems(3) = DataInizioServizio
        
        'NumeroIntervento = NumeroIntervento + 1
        
        'INSERIMENTO INTERVENTO TESTA
        rsIntervento.AddNew
            LINK_INTERVENTO = fnGetNewKey("RV_POIntervento", "IDRV_POIntervento")
            rsIntervento!IDRV_POIntervento = LINK_INTERVENTO
            rsIntervento!IDRV_POInterventoPadre = LINK_INTERVENTO
            rsIntervento!IDRV_POContrattoServizi = IDContrattoServizio
            rsIntervento!IDRV_POContratto = IDContratto
            rsIntervento!IDRV_POContrattoPadre = IDContrattoPadre
            rsIntervento!IDAnagraficaFatturazione = IDAnagraficaFatturazione
            rsIntervento!IDAnagraficaCliente = IDAnagraficaCliente
            rsIntervento!IDAzienda = TheApp.IDFirm
            rsIntervento!IDFiliale = TheApp.Branch
            rsIntervento!NumeroIntervento = GET_NUMERO_INTERVENTO(Year(Date))
            rsIntervento!AnnoIntervento = Year(Date)
            rsIntervento!IDArticolo = IDArticolo
            rsIntervento!IDAnagraficaTecnicoRif = IDAnagraficaTecnicoRif
            rsIntervento!IDTipoAnagraficaTecnicoRif = IDTipoAnaTecRif
            rsIntervento!IDRV_POStatoIntervento = IDTipoStatoInt
            rsIntervento!InterventoChiuso = 0
            rsIntervento!DataInserimento = Date
            rsIntervento!OraInserimento = GET_ORARIO(Now)
            rsIntervento!IDUtenteInserimento = TheApp.IDUser
            rsIntervento!DataUltimaModifica = Date
            rsIntervento!OraUltimaModifica = GET_ORARIO(Now)
            rsIntervento!IDUtenteUltimaModifica = TheApp.IDUser
            rsIntervento!Elaborata = 1
            rsIntervento!Manuale = 0
            rsIntervento!Richiesta = DescrizioneArticolo
            rsIntervento!NomeComputerInserimento = GET_NOMECOMPUTER
            rsIntervento!UtenteComputerInserimento = GET_NOMEUTENTE
            rsIntervento!NomeComputerModifica = GET_NOMECOMPUTER
            rsIntervento!UtenteComputerModifica = GET_NOMEUTENTE
            rsIntervento!NumeroFase = 1
            rsIntervento!IDRV_POTipoFaseIntervento = IDTipoFase
            rsIntervento!IDAnagraficaTecnicoOperativo = IDAnagraficaTecnicoRif
            rsIntervento!IDTipoAnagraficaTecnicoOpe = IDTipoAnaTecOpe
            rsIntervento!DataAppuntamento = DataInizioServizio
            rsIntervento!OraAppuntamento = "09.00"
            rsIntervento!LavoroEseguito = DescrizioneArticolo
            rsIntervento!Annotazioni = ""
            rsIntervento!IDRV_POStagione = GET_LINK_STAGIONE(DataInizioServizio)
            rsIntervento!IDRV_POCategoriaIntervento = GET_PARAMETRI_TEC_OPE(fnNotNullN(rsIntervento!IDAnagraficaTecnicoOperativo), "IDRV_POCategoriaFase")
            rsIntervento!IDRV_POTipoAddebito = GET_PARAMETRI_TEC_OPE(fnNotNullN(rsIntervento!IDAnagraficaTecnicoOperativo), "IDRV_POTipoAddebito")
            rsIntervento!IDRV_POTipoClasseIntervento = GET_PARAMETRI_TEC_OPE(fnNotNullN(rsIntervento!IDAnagraficaTecnicoOperativo), "IDRV_POTipoClasseIntervento")
            If fnNotNullN(rsIntervento!IDRV_POCategoriaIntervento) = 0 Then
                rsIntervento!IDRV_POCategoriaIntervento = IDCategoriaIntervento
            End If
        rsIntervento.Update

        If CREA_APPUNTAMENTO_AGENDA = 1 Then
            SCRIVI_APPUNTAMENTO fnNotNull(rsIntervento!DataAppuntamento), fnNotNull(rsIntervento!OraAppuntamento), LINK_INTERVENTO
        End If
        
        DataInizioServizio = DateAdd("m", IIf((MeseRic = ""), 0, MeseRic), DataInizioServizio)
        DataInizioServizio = DateAdd("d", IIf((GiornoRic = ""), 0, GiornoRic), DataInizioServizio)
        DataInizioServizio = DateAdd("ww", IIf((SettimanaRic = ""), 0, SettimanaRic), DataInizioServizio)
    Wend
End If
rsIntervento.Close
Set rsIntervento = Nothing



End Sub

Private Function GET_COSTRUZIONE_DATA_PERS(Giorno As String, mese As String) As String
Dim GiornoInizio As String
Dim MeseInizio As String

If Len(Giorno) = 1 Then
    GiornoInizio = "0" & Giorno
ElseIf Len(Giorno) = 0 Then
    GiornoInizio = "01" '& Giorno
Else
    GiornoInizio = Giorno
End If
    
If Len(mese) = 1 Then
    MeseInizio = "0" & mese
ElseIf Len(mese) = 0 Then
    MeseInizio = "01" '& Giorno
Else
    MeseInizio = mese
End If

GET_COSTRUZIONE_DATA_PERS = GiornoInizio & "/" & MeseInizio & "/"
End Function


Private Function GET_ORARIO(StringaData As String) As String
Dim Ora As String
Dim Minuti As String
Dim Secondi As String

If Len(DatePart("h", StringaData)) = 1 Then
    Ora = "0" & DatePart("h", StringaData)
Else
    Ora = DatePart("h", StringaData)
End If
If Len(DatePart("n", StringaData)) = 1 Then
    Minuti = "0" & DatePart("n", StringaData)
Else
    Minuti = DatePart("n", StringaData)
End If
If Len(DatePart("s", StringaData)) = 1 Then
    Secondi = "0" & DatePart("s", StringaData)
Else
    Secondi = DatePart("s", StringaData)
End If

GET_ORARIO = Ora & "." & Minuti & "." & Secondi


End Function
Private Function GET_NUMERO_INTERVENTO(Anno As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(NumeroIntervento) as Numero "
sSQL = sSQL & "FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND AnnoIntervento=" & Anno

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_INTERVENTO = 1
Else
    GET_NUMERO_INTERVENTO = fnNotNullN(rs!numero) + 1
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function ELIMINA_RIFERIMENTI_CODA(IDTipoOggetto As Long)
Dim sSQL As String

''''''''ELIMINAZIONE RIFERIMENTO CODA'''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POTMP "
sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Private Function fnGetTipoOggetto(NomeGestore) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(NomeGestore)
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    
End Function
Private Sub GET_PARAMETRI_AZIENDA(IDAzienda As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POParametriAzienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    LINK_TIPO_ANA_TEC_INT = 0
    LINK_STATO_INT_NUOVO = 0
    LINK_STATO_INT_CHIUSO = 0
    
    LINK_TIPO_ANA_TEC_FASE = 0
    LINK_STATO_FASE_NUOVA = 0
    LINK_STATO_FASE_CHIUSA = 0
    LINK_TIPO_FASE_ELA = 0
    LINK_TIPO_FASE_MANUALE = 0
    LINK_TIPO_ANA_AMM = 0
    If LINK_TIPO_TASCA = 0 Then
        LINK_TIPO_TASCA = 0
    End If

    FLAG_RIPORTA_TEC_CONTRATTO = 0
    LINK_TIPO_ANA_TEC_CONTRATTO = 0
    VISUALIZZA_IMPORTI_PROD = 0
    LINK_SEZIONALE_RATE = 0
    CREA_APPUNTAMENTO_AGENDA = 0
    LINK_TIPO_IMPOSTAZIONE = 0
    
    LINK_UM_PERIODO_AZIENDA = 0
    LINK_ARTICOLO_SERVIZIO = 0
    LINK_ARTICOLO_ADDEBITO = 0
    SEL_PROD_NON_CONTRATTO = 0
    VIS_FORM_GEN_INT_DA_PROD = 0
    NON_IMPOSTARE_FILTRI = 0
    DISDETTO_NONFATTURARE = 0
    OBBLIGATORIO_COLLEGAMENTO_INT = 0
    NON_VISUALIZZARE_ALTRI_DATI = 0
Else
    LINK_TIPO_ANA_TEC_INT = fnNotNullN(rs!IDTipoAnagraficaTecnicoIntRif)
    LINK_STATO_INT_NUOVO = fnNotNullN(rs!IDRV_POStatoInterventoInserimento)
    LINK_STATO_INT_CHIUSO = fnNotNullN(rs!IDRV_POStatoInterventoChiuso)
    
    LINK_TIPO_ANA_TEC_FASE = fnNotNullN(rs!IDTipoAnagraficaTecnicoFaseRif)
    LINK_STATO_FASE_NUOVA = fnNotNullN(rs!IDRV_POStatoFaseInserimento)
    LINK_STATO_FASE_CHIUSA = fnNotNullN(rs!IDRV_POStatoFaseChiusa)
    LINK_TIPO_FASE_ELA = fnNotNullN(rs!IDRV_POTipoFaseInterventoEla)
    LINK_TIPO_FASE_MANUALE = fnNotNullN(rs!IDRV_POTipoFaseInterventoMan)
    If LINK_TIPO_TASCA = 0 Then
        LINK_TIPO_TASCA = fnNotNullN(rs!IDRV_POTipoTasca)
    End If
    
    FLAG_RIPORTA_TEC_CONTRATTO = fnNotNullN(rs!RiportaTecContratto)
    LINK_TIPO_ANA_TEC_CONTRATTO = fnNotNullN(rs!IDTipoAnagraficaContratto)
    LINK_TIPO_ANA_AMM = fnNotNullN(rs!IDTipoAnagraficaAmministratore)
    VISUALIZZA_IMPORTI_PROD = fnNotNullN(rs!VisualizzaImportiProdContratto)
    LINK_SEZIONALE_RATE = fnNotNullN(rs!IDSezionaleRateContratto)
    CREA_APPUNTAMENTO_AGENDA = fnNotNullN(rs!GenAppAutAgendaContratto)
    
    LINK_TIPO_IMPOSTAZIONE = fnNotNullN(rs!IDRV_POTipoImpostazioneContratto)
    If LINK_TIPO_IMPOSTAZIONE = 0 Then LINK_TIPO_IMPOSTAZIONE = 1
    
    LINK_UM_PERIODO_AZIENDA = fnNotNullN(rs!IDRV_POUnitaDiMisuraPeriodo)
    If LINK_UM_PERIODO_AZIENDA = 0 Then LINK_UM_PERIODO_AZIENDA = 2
    LINK_ARTICOLO_SERVIZIO = fnNotNullN(rs!IDGruppoEquivalenzaArticoloIntervento)
    LINK_ARTICOLO_ADDEBITO = fnNotNullN(rs!IDGruppoEquivalenzaArticoloAddebito)
    SEL_PROD_NON_CONTRATTO = fnNotNullN(rs!SelProdNonContratto)
    VIS_FORM_GEN_INT_DA_PROD = fnNotNullN(rs!VisualizzaAutGenIntDaProd)
    NON_IMPOSTARE_FILTRI = fnNotNullN(rs!NonImpostareFiltriRicercaContratto)
    DISDETTO_NONFATTURARE = fnNotNullN(rs!ContrattoDisdettoImpostaNonFatturare)
    OBBLIGATORIO_COLLEGAMENTO_INT = fnNotNullN(rs!NonEliminareConCollIntInContrProd)
    NON_VISUALIZZARE_ALTRI_DATI = fnNotNullN(rs!NonVisualizzareAltriDatiContratto)
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Public Function GET_NOMECOMPUTER() As String
Dim dwLen As Long
Dim strString As String
Const MAX_COMPUTERNAME_LENGTH As Long = 31
    
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    'Show the computer name
    GET_NOMECOMPUTER = strString
End Function

Function GET_NOMEUTENTE() As String
    Dim strString As String
    Dim lunghezzaStringa As Long
    lunghezzaStringa = 32
    strString = String(lunghezzaStringa, " ")
    GetUserName strString, lunghezzaStringa
    strString = Left(strString, lunghezzaStringa)
    GET_NOMEUTENTE = strString
    GET_NOMEUTENTE = Mid(GET_NOMEUTENTE, 1, Len(GET_NOMEUTENTE) - 1)
End Function

Private Function GET_DESCRIZIONE_ARTICOLO(IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Articolo FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_ARTICOLO = ""
Else
    GET_DESCRIZIONE_ARTICOLO = fnNotNull(rs!Articolo)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub GET_GRIGLIA_INTERVENTI()
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    
    sSQL = "SELECT * FROM RV_POIEIntervento "
    sSQL = sSQL & "WHERE IDRV_POContratto=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    sSQL = sSQL & " ORDER BY AnnoIntervento, NumeroIntervento"
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
    If Not (rsGrigliaInt Is Nothing) Then
        If rsGrigliaInt.State > 0 Then
            rsGrigliaInt.Close
        End If
        Set rsGrigliaInt = Nothing
    End If
    
    Set rsGrigliaInt = New ADODB.Recordset
    rsGrigliaInt.CursorLocation = adUseClient
    rsGrigliaInt.Open sSQL, Cn.InternalConnection
        
    With Me.GrigliaInterventi
        'Set .PaintNotifyObj = gPaintNotify
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear

        .ColumnsHeader.Add "IDRV_POIntervento", "IDRV_POIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDFiliale", "IDFiliale", dgInteger, False, 500, dgAlignleft
        '''''''''''''''''''''''''''''''''''''''''DATI INTERVENTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .ColumnsHeader.Add "InterventoChiuso", "Chiuso", dgBoolean, True, 1000, dgAligncenter
        .ColumnsHeader.Add "DataInserimento", "Data inserimento", dgDate, False, 1500, dgAlignleft
        .ColumnsHeader.Add "IDRV_POTipoFaseIntervento", "IDRV_POTipoFaseIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "TipoFaseIntervento", "Tipo fase", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "OraInserimento", "Ora inserimento", dgchar, False, 1000, dgAlignleft
        .ColumnsHeader.Add "AnnoIntervento", "Anno Int.", dgInteger, False, 1200, dgAlignRight
        .ColumnsHeader.Add "NumeroIntervento", "N° Int.", dgInteger, True, 1200, dgAlignRight
        .ColumnsHeader.Add "NumeroInterventoSub", "Sub", dgInteger, False, 1200, dgAlignRight
        .ColumnsHeader.Add "NumeroFase", "N° fase", dgInteger, True, 500, dgAlignRight
        .ColumnsHeader.Add "IDUtenteInserimento", "IDUtenteInserimento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "UtenteInserimento", "Utente Ins.", dgchar, False, 1500, dgAlignleft
        .ColumnsHeader.Add "DataAppuntamento", "Data appuntamento", dgDate, True, 1500, dgAligncenter
        .ColumnsHeader.Add "OraAppuntamento", "Ora appuntamento", dgchar, True, 1500, dgAlignRight
        
        .ColumnsHeader.Add "IDAnagraficaCliente", "IDAnagraficaCliente", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaCliente", "Cliente", dgchar, False, 2500, dgAlignleft
        .ColumnsHeader.Add "NomeAnagraficaCliente", "Nome cliente", dgchar, False, 1500, dgAlignleft
        .ColumnsHeader.Add "IDAnagraficaTecnicoRif", "IDAnagraficaTecnicoRiferimento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDTipoAnagraficaTecnicoRif", "IDTipoAnagraficaTecnicoRif", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaTecnico", "Riferimento interno", dgchar, True, 2500, dgAlignleft
        .ColumnsHeader.Add "NomeTecnico", "Nome riferimento interno", dgchar, False, 1500, dgAlignleft
        .ColumnsHeader.Add "IDAnagraficaTecnicoOperativo", "IDAnagraficaTecnicoOperativo", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDTipoAnagraficaTecnicoOpe", "IDTipoAnagraficaTecnicoOpe", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaTecnicoOperativo", "Tecnico operativo", dgchar, False, 2500, dgAlignleft
        .ColumnsHeader.Add "NomeTecnicoOperativo", "Nome Tecnico operativo", dgchar, False, 1500, dgAlignleft
        
        .ColumnsHeader.Add "IDRV_POCategoriaIntervento", "IDRV_POCategoriaIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "CategoriaIntervento", "Categoria Int.", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "IDRV_POStatoIntervento", "IDRV_POStatoIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "StatoIntervento", "Stato Int.", dgchar, True, 1500, dgAlignleft
        '.ColumnsHeader.Add "Richiesta", "Richiesta", dgchar, True, 4000, dgAlignleft
        .ColumnsHeader.Add "Annotazioni", "Annotazioni", dgchar, False, 4000, dgAlignleft
        .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "CodiceArticolo", "Codice articolo int.", dgchar, True, 1500, dgAlignleft
        .ColumnsHeader.Add "Articolo", "Articolo int.", dgchar, True, 2500, dgAlignleft
        '.ColumnsHeader.Add "LavoroEseguito", "Lavoro Eseguito", dgchar, True, 4000, dgAlignleft
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'        .EnableRowColors = True
'        .RowColors.Clear
'        .RowColors.Add "InterventoPadre", "NumeroFase=1", &HD5FDED
'
'        .RowColors.Add "AltriInterventi", "NumeroFase>1", &HFFFFCC
        
        Set .Recordset = rsGrigliaInt
        .LoadUserSettings
        .Refresh
            
    End With
    Cn.CursorLocation = OLDCursor

    Me.GrigliaInterventi.Refresh
    Me.GrigliaInterventi.LoadUserSettings

Exit Sub

ERR_fnGrigliaAssegnazione:
    MsgBox Err.Description, vbCritical, "Reperimento dati assegnazione"

End Sub

Private Sub GET_RECORDSET_OGGETTO_SINGOLO(IDRigaIntervento As Long, IDIntervento As Long, IDAzienda As Long, IDFiliale As Long)
Dim sSQL As String
Dim rsOgg As ADODB.Recordset
Dim rsInt As DmtOleDbLib.adoResultset
Dim rsRiga As DmtOleDbLib.adoResultset

If (IDRigaIntervento = 0) And (IDIntervento > 0) Then
    GET_RECORDSET_OGGETTO_ALL IDIntervento, TheApp.IDFirm, TheApp.Branch
End If

'''''''''''''''''''''RECORDSET TESTA INTERVENTO'''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIEIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=" & IDIntervento

Set rsInt = Cn.OpenResultset(sSQL)

If rsInt.EOF Then
    rsInt.CloseResultset
    Set rsInt = Nothing
    Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''RECORDSET RIGA INTERVENTO''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIEInterventoRighe "
sSQL = sSQL & "WHERE IDRV_POInterventoRighe=" & IDRigaIntervento

Set rsRiga = Cn.OpenResultset(sSQL)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''RECORDSET OGGETTO''''''''''''''''''''''''''''''''''''''''''''''''
Set rsOgg = New ADODB.Recordset


sSQL = "SELECT * FROM RV_POOggettoIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=" & IDIntervento
sSQL = sSQL & " AND IDRV_POInterventoRighe=" & IDRigaIntervento

rsOgg.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsOgg.EOF Then
    rsOgg.AddNew
    rsOgg!IDRV_POOggettoIntervento = fnGetNewKey("RV_POOggettoIntervento", "IDRV_POOggettoIntervento")
End If

    rsOgg!IDAzienda = IDAzienda
    rsOgg!IDFiliale = IDFiliale
    rsOgg!IDRV_POIntervento = IDIntervento
    rsOgg!IDRV_POInterventoRighe = IDRigaIntervento
    
    ''''''''TESTA INTERVENTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rsOgg!IDAnagraficaCliente = fnNotNullN(rsInt!IDAnagraficaCliente)
    rsOgg!AnagraficaCliente = fnNotNull(rsInt!AnagraficaCliente)
    rsOgg!NomeAnagraficaCliente = fnNotNull(rsInt!NomeCliente)
    rsOgg!NumeroIntervento = fnNotNullN(rsInt!NumeroIntervento)
    rsOgg!AnnoIntervento = fnNotNullN(rsInt!AnnoIntervento)
    rsOgg!IDArticoloInt = fnNotNullN(rsInt!IDArticolo)
    rsOgg!CodiceArticoloInt = fnNotNull(rsInt!CodiceArticolo)
    rsOgg!ArticoloInt = fnNotNull(rsInt!Articolo)
    
    rsOgg!IDAnagraficaTecnicoRifInt = fnNotNullN(rsInt!IDAnagraficaTecnicoRif)
    rsOgg!AnagraficaTecnicoRifInt = fnNotNull(rsInt!AnagraficaTecnico)
    rsOgg!NomeAnagraficaTecnicoRifInt = fnNotNull(rsInt!NomeTecnico)
    
    rsOgg!IDRV_POStatoIntervento = fnNotNullN(rsInt!IDRV_POStatoIntervento)
    rsOgg!StatoIntervento = fnNotNull(rsInt!StatoIntervento)
    rsOgg!InterventoChiuso = Abs(fnNotNullN(rsInt!InterventoChiuso))
    rsOgg!IDRV_POCategoriaIntervento = fnNotNullN(rsInt!IDRV_POCategoriaIntervento)
    rsOgg!CategoriaIntervento = fnNotNull(rsInt!CategoriaIntervento)
    rsOgg!Richiesta = fnNotNull(rsInt!Richiesta)
    rsOgg!Annotazioni = fnNotNull(rsInt!Annotazioni)
    rsOgg!DataInserimento = rsRiga!DataInserimento
    rsOgg!OraInserimento = fnNotNull(rsRiga!OraInserimento)
    rsOgg!IDRV_POContratto = fnNotNullN(rsInt!IDRV_POContratto)
    rsOgg!IDRV_POStoriaContratto = fnNotNullN(rsInt!IDRV_POStoriaContratto)
    rsOgg!IDTipoContratto = fnNotNullN(rsInt!IDTipoContratto)
    rsOgg!TipoContratto = fnNotNull(rsInt!TipoContratto)
    rsOgg!IDSitoPerAnagraficaContratto = fnNotNullN(rsInt!IDSitoPerAnagrafica)
    rsOgg!SitoPerAnagraficaContratto = fnNotNull(rsInt!SitoPerAnagrafica)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rsOgg!IDRV_POTipoIntestazione = 2
    rsOgg!Intestazione = "FASE INTERVENTO"
    rsOgg!NumeroFaseIntervento = fnNotNullN(rsRiga!NumeroFaseIntervento)
    rsOgg!DataAppuntamento = rsRiga!DataAppuntamento
    rsOgg!OraAppuntamento = fnNotNull(rsRiga!OraAppuntamento)
    rsOgg!IDRV_POStatoFase = fnNotNullN(rsRiga!IDRV_POStatoFase)
    rsOgg!StatoFase = fnNotNull(rsRiga!StatoFase)
    rsOgg!IDRV_POTipoFaseIntervento = fnNotNullN(rsRiga!IDRV_POTipoFaseIntervento)
    rsOgg!TipoFaseIntervento = fnNotNull(rsRiga!TipoFaseIntervento)
    rsOgg!AnnotazioniFase = fnNotNull(rsRiga!Annotazioni)
    rsOgg!FaseChiusa = Abs(fnNotNullN(rsRiga!FaseChiusa))
    rsOgg!IDArticoloFase = fnNotNullN(rsRiga!IDArticolo)
    rsOgg!CodiceArticoloFase = fnNotNull(rsRiga!CodiceArticolo)
    rsOgg!ArticoloFase = fnNotNull(rsRiga!Articolo)
    rsOgg!IDAnagraficaTecnicoOpe = fnNotNullN(rsRiga!IDAnagraficaTecnicoOperativo)
    rsOgg!AnagraficaTecnicoOpe = fnNotNull(rsRiga!AnagraficaTecnicoOperativo)
    rsOgg!NomeAnagraficaTecnicoOpe = fnNotNull(rsRiga!NomeTecnicoOperativo)
    rsOgg!IDAnagraficaTecnicoRiferimento = fnNotNullN(rsRiga!IDAnagraficaTecnicoRif)
    rsOgg!AnagraficaTecnicoRiferimento = fnNotNull(rsRiga!AnagraficaTecnicoRif)
    rsOgg!NomeAnagraficaTecnicoRiferimento = fnNotNull(rsRiga!NomeTecnicoRif)
    rsOgg!LavoroEseguitoFase = fnNotNull(rsRiga!LavoroEseguito)
    rsOgg!IDRV_POCategoriaFase = fnNotNullN(rsRiga!IDRV_POCategoriaFase)
    rsOgg!CategoriaFase = fnNotNull(rsRiga!CategoriaFase)
    rsOgg!IDUtenteInserimento = fnNotNullN(rsRiga!IDUtenteInserimento)
    rsOgg!UtenteInserimento = fnNotNull(rsRiga!UtenteInserimento)

rsOgg.Update

rsOgg.Close
Set rsOgg = Nothing

rsRiga.CloseResultset
Set rsRiga = Nothing

rsInt.CloseResultset
Set rsInt = Nothing
End Sub


Private Sub GET_RECORDSET_OGGETTO_ALL(IDIntervento As Long, IDAzienda As Long, IDFiliale As Long)
Dim sSQL As String
Dim rsOgg As ADODB.Recordset
Dim rsInt As DmtOleDbLib.adoResultset
Dim rsRiga As DmtOleDbLib.adoResultset
Dim rsOggRighe As ADODB.Recordset


''''ELIMINAZIONE DATI'''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POOggettoIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=" & IDIntervento
Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'SCRITTURA SOLO RECORDSET INTERVENTO
'''''''''''''''''''''RECORDSET TESTA INTERVENTO'''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIEIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=" & IDIntervento

Set rsInt = Cn.OpenResultset(sSQL)

If rsInt.EOF Then
    rsInt.CloseResultset
    Set rsInt = Nothing
    Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''RECORDSET OGGETTO''''''''''''''''''''''''''''''''''''''''''''
Set rsOgg = New ADODB.Recordset


sSQL = "SELECT * FROM RV_POOggettoIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=" & IDIntervento

rsOgg.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

rsOgg.AddNew
    rsOgg!IDRV_POOggettoIntervento = fnGetNewKey("RV_POOggettoIntervento", "IDRV_POOggettoIntervento")
    
    rsOgg!IDAzienda = IDAzienda
    rsOgg!IDFiliale = IDFiliale
    rsOgg!IDRV_POIntervento = IDIntervento
    rsOgg!IDRV_POInterventoRighe = 0
    
    ''''''''TESTA INTERVENTO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rsOgg!IDAnagraficaCliente = fnNotNullN(rsInt!IDAnagraficaCliente)
    rsOgg!AnagraficaCliente = fnNotNull(rsInt!AnagraficaCliente)
    rsOgg!NomeAnagraficaCliente = fnNotNull(rsInt!NomeCliente)
    rsOgg!NumeroIntervento = fnNotNullN(rsInt!NumeroIntervento)
    rsOgg!AnnoIntervento = fnNotNullN(rsInt!AnnoIntervento)
    rsOgg!IDArticoloInt = fnNotNullN(rsInt!IDArticolo)
    rsOgg!CodiceArticoloInt = fnNotNull(rsInt!CodiceArticolo)
    rsOgg!ArticoloInt = fnNotNull(rsInt!Articolo)
    rsOgg!IDAnagraficaTecnicoRiferimento = fnNotNullN(rsInt!IDAnagraficaTecnicoRif)
    rsOgg!AnagraficaTecnicoRiferimento = fnNotNull(rsInt!AnagraficaTecnico)
    rsOgg!NomeAnagraficaTecnicoRiferimento = fnNotNull(rsInt!NomeTecnico)
    rsOgg!IDRV_POStatoIntervento = fnNotNullN(rsInt!IDRV_POStatoIntervento)
    rsOgg!StatoIntervento = fnNotNull(rsInt!StatoIntervento)
    rsOgg!InterventoChiuso = Abs(fnNotNullN(rsInt!InterventoChiuso))
    rsOgg!IDRV_POCategoriaIntervento = fnNotNullN(rsInt!IDRV_POCategoriaIntervento)
    rsOgg!CategoriaIntervento = fnNotNull(rsInt!CategoriaIntervento)
    rsOgg!Richiesta = fnNotNull(rsInt!Richiesta)
    rsOgg!Annotazioni = fnNotNull(rsInt!Annotazioni)
    rsOgg!DataInserimento = rsInt!DataInserimento
    rsOgg!OraInserimento = fnNotNull(rsInt!OraInserimento)
    rsOgg!IDRV_POTipoIntestazione = 1
    rsOgg!Intestazione = "INTERVENTO"
    rsOgg!NumeroFaseIntervento = 0
    rsOgg!IDUtenteInserimento = fnNotNullN(rsInt!IDUtenteInserimento)
    rsOgg!UtenteInserimento = fnNotNull(rsInt!UtenteInserimentoIntervento)
    rsOgg!IDRV_POContratto = fnNotNullN(rsInt!IDRV_POContratto)
    rsOgg!IDRV_POStoriaContratto = fnNotNullN(rsInt!IDRV_POStoriaContratto)
    rsOgg!IDTipoContratto = fnNotNullN(rsInt!IDTipoContratto)
    rsOgg!TipoContratto = fnNotNull(rsInt!TipoContratto)
    rsOgg!IDSitoPerAnagraficaContratto = fnNotNullN(rsInt!IDSitoPerAnagrafica)
    rsOgg!SitoPerAnagraficaContratto = fnNotNull(rsInt!SitoPerAnagrafica)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

rsOgg.Update


rsOgg.Close
Set rsOgg = Nothing

'''''''''''''''''''''RECORDSET RIGA INTERVENTO''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIEInterventoRighe "
sSQL = sSQL & "WHERE IDRV_POIntervento=" & IDIntervento

Set rsRiga = Cn.OpenResultset(sSQL)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set rsOgg = New ADODB.Recordset


sSQL = "SELECT * FROM RV_POOggettoIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=" & IDIntervento

rsOgg.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

'''''''''RECORDSET OGGETTO''''''''''''''''''''''''''''''''''''''''''''


While Not rsRiga.EOF

    rsOgg.AddNew
        rsOgg!IDRV_POOggettoIntervento = fnGetNewKey("RV_POOggettoIntervento", "IDRV_POOggettoIntervento")
        
        
        rsOgg!IDAzienda = IDAzienda
        rsOgg!IDFiliale = IDFiliale
        rsOgg!IDRV_POIntervento = IDIntervento
        rsOgg!IDRV_POInterventoRighe = fnNotNullN(rsRiga!IDRV_POInterventoRighe)
        
        ''''''''TESTA INTERVENTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        rsOgg!IDAnagraficaCliente = fnNotNullN(rsInt!IDAnagraficaCliente)
        rsOgg!AnagraficaCliente = fnNotNull(rsInt!AnagraficaCliente)
        rsOgg!NomeAnagraficaCliente = fnNotNull(rsInt!NomeCliente)
        rsOgg!NumeroIntervento = fnNotNullN(rsInt!NumeroIntervento)
        rsOgg!AnnoIntervento = fnNotNullN(rsInt!AnnoIntervento)
        rsOgg!IDArticoloInt = fnNotNullN(rsInt!IDArticolo)
        rsOgg!CodiceArticoloInt = fnNotNull(rsInt!CodiceArticolo)
        rsOgg!ArticoloInt = fnNotNull(rsInt!Articolo)
        
        rsOgg!IDAnagraficaTecnicoRifInt = fnNotNullN(rsInt!IDAnagraficaTecnicoRif)
        rsOgg!AnagraficaTecnicoRifInt = fnNotNull(rsInt!AnagraficaTecnico)
        rsOgg!NomeAnagraficaTecnicoRifInt = fnNotNull(rsInt!NomeTecnico)
        
        rsOgg!IDRV_POStatoIntervento = fnNotNullN(rsInt!IDRV_POStatoIntervento)
        rsOgg!StatoIntervento = fnNotNull(rsInt!StatoIntervento)
        rsOgg!InterventoChiuso = Abs(fnNotNullN(rsInt!InterventoChiuso))
        rsOgg!IDRV_POCategoriaIntervento = fnNotNullN(rsInt!IDRV_POCategoriaIntervento)
        rsOgg!CategoriaIntervento = fnNotNull(rsInt!CategoriaIntervento)
        rsOgg!Richiesta = fnNotNull(rsInt!Richiesta)
        rsOgg!Annotazioni = fnNotNull(rsInt!Annotazioni)
        rsOgg!DataInserimento = rsRiga!DataInserimento
        rsOgg!OraInserimento = fnNotNull(rsRiga!OraInserimento)
        rsOgg!IDRV_POContratto = fnNotNullN(rsInt!IDRV_POContratto)
        rsOgg!IDRV_POStoriaContratto = fnNotNullN(rsInt!IDRV_POStoriaContratto)
        rsOgg!IDTipoContratto = fnNotNullN(rsInt!IDTipoContratto)
        rsOgg!TipoContratto = fnNotNull(rsInt!TipoContratto)
        rsOgg!IDSitoPerAnagraficaContratto = fnNotNullN(rsInt!IDSitoPerAnagrafica)
        rsOgg!SitoPerAnagraficaContratto = fnNotNull(rsInt!SitoPerAnagrafica)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
        rsOgg!IDRV_POTipoIntestazione = 2
        rsOgg!Intestazione = "FASE INTERVENTO"
        rsOgg!NumeroFaseIntervento = fnNotNullN(rsRiga!NumeroFaseIntervento)
        rsOgg!DataAppuntamento = rsRiga!DataAppuntamento
        rsOgg!OraAppuntamento = fnNotNull(rsRiga!OraAppuntamento)
        rsOgg!IDRV_POStatoFase = fnNotNullN(rsRiga!IDRV_POStatoFase)
        rsOgg!StatoFase = fnNotNull(rsRiga!StatoFase)
        rsOgg!IDRV_POTipoFaseIntervento = fnNotNullN(rsRiga!IDRV_POTipoFaseIntervento)
        rsOgg!TipoFaseIntervento = fnNotNull(rsRiga!TipoFaseIntervento)
        rsOgg!AnnotazioniFase = fnNotNull(rsRiga!Annotazioni)
        rsOgg!FaseChiusa = Abs(fnNotNullN(rsRiga!FaseChiusa))
        rsOgg!IDArticoloFase = fnNotNullN(rsRiga!IDArticolo)
        rsOgg!CodiceArticoloFase = fnNotNull(rsRiga!CodiceArticolo)
        rsOgg!ArticoloFase = fnNotNull(rsRiga!Articolo)
        rsOgg!IDAnagraficaTecnicoOpe = fnNotNullN(rsRiga!IDAnagraficaTecnicoOperativo)
        rsOgg!AnagraficaTecnicoOpe = fnNotNull(rsRiga!AnagraficaTecnicoOperativo)
        rsOgg!NomeAnagraficaTecnicoOpe = fnNotNull(rsRiga!NomeTecnicoOperativo)
        rsOgg!IDAnagraficaTecnicoRiferimento = fnNotNullN(rsRiga!IDAnagraficaTecnicoRif)
        rsOgg!AnagraficaTecnicoRiferimento = fnNotNull(rsRiga!AnagraficaTecnicoRif)
        rsOgg!NomeAnagraficaTecnicoRiferimento = fnNotNull(rsRiga!NomeTecnicoRif)
        rsOgg!LavoroEseguitoFase = fnNotNull(rsRiga!LavoroEseguito)
        rsOgg!IDRV_POCategoriaFase = fnNotNullN(rsRiga!IDRV_POCategoriaFase)
        rsOgg!CategoriaFase = fnNotNull(rsRiga!CategoriaFase)
        rsOgg!IDUtenteInserimento = fnNotNullN(rsRiga!IDUtenteInserimento)
        rsOgg!UtenteInserimento = fnNotNull(rsRiga!UtenteInserimento)

    rsOgg.Update
    
rsRiga.MoveNext
Wend

rsOgg.Close
Set rsOgg = Nothing
rsRiga.CloseResultset
Set rsRiga = Nothing
rsInt.CloseResultset
Set rsInt = Nothing
End Sub
Private Function GET_PARAMETRI_TEC_OPE(IDAnagrafica As Long, Nome As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & Nome
sSQL = sSQL & " FROM RV_POConfigurazioneTecnicoOpe "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PARAMETRI_TEC_OPE = 0
Else
    GET_PARAMETRI_TEC_OPE = fnNotNullN(rs.adoColumns(Nome).Value)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_LINK_STAGIONE(VarData As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POStagione FROM RV_POStagione "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND DataInizio<=" & fnNormDate(VarData)
sSQL = sSQL & " AND DataFine>=" & fnNormDate(VarData)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_STAGIONE = 0
Else
    GET_LINK_STAGIONE = fnNotNullN(rs!IDRV_POStagione)
End If

rs.CloseResultset
Set rs = Nothing

End Function







Private Sub m_DocumentsLink2_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink2.BOF And m_DocumentsLink2.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        Me.txtDataStipulaAdeg.Value = fnNotNullN(m_DocumentsLink2("DataStipula").Value)
        Me.txtDataDecorrenzaAdeg.Value = fnNotNullN(m_DocumentsLink2("DataDecorrenza").Value)
        Me.txtImportoAdeg.Value = fnNotNullN(m_DocumentsLink2("Importo").Value)
        Me.txtAnnotazioniAdeg.Text = fnNotNull(m_DocumentsLink2("Annotazioni").Value)
        Me.chkAdegContrProx.Value = Abs(fnNotNullN(m_DocumentsLink2("RiportaProssimoRinnovo").Value))
        Me.chkAdeguaContrAttuale.Value = Abs(fnNotNullN(m_DocumentsLink2("AdeguaContrattoAttuale").Value))
        'Me.cboIvaAdeg.WriteOn fnNotNullN(m_DocumentsLink2("IDIvaFatturazione").Value)
        Me.cboTipoAdeguamento.WriteOn fnNotNullN(m_DocumentsLink2("IDRV_POTipoAdeguamento").Value)
        Me.CDArticoloAdeg.Load fnNotNullN(m_DocumentsLink2("IDArticolo").Value)
        Me.txtProtAdeg.Text = fnNotNull(m_DocumentsLink2("NumeroProtocollo").Value)
        Me.chkIstatAdeguamento.Value = Abs(fnNotNullN(m_DocumentsLink2("AdeguamentoIstat").Value))
        Me.txtNumeroAdeguamento.Value = fnNotNullN(m_DocumentsLink2("NumeroAdeguamento").Value)
        Me.cboIstatAdeg.WriteOn fnNotNullN(m_DocumentsLink2("IDIstat").Value)
        Me.txtMaggiorazioneAdeg.Value = fnNotNullN(m_DocumentsLink2("MaggiorazioneIstat").Value)
        Me.txtDescrFattAde.Text = fnNotNull(m_DocumentsLink2("DescrizionePerFatturazione").Value)
        Me.cboTipoRateizzazioneAdeg.WriteOn fnNotNullN(m_DocumentsLink2("IDRateizzazione").Value)
        Me.txtDataScadenzaAdeg.Value = fnNotNullN(m_DocumentsLink2("DataFineAdeguamento").Value)
        chkNoCalcPeriodoFatt.Value = fnNotNullN(m_DocumentsLink2("NoCalcPeriodoFatt").Value)
        txtImportoAdegRinn.Value = fnNotNullN(m_DocumentsLink2("ImportoAlRinnovo").Value)
        txtNAdegIniz.Value = fnNotNullN(m_DocumentsLink2("NumeroPartenza").Value)
        bValue = True
        
        
        '----------------------------------------------------------------------------
        'Popola i controlli associati al sottodocumento con i valori presenti
        'nell'oggetto DocumentsLink
        '----------------------------------------------------------------------------
       
    Else
        'Il DocumentsLink è vuoto - non contiene righe.
        '---------------------------------------------
        'Ripulisce i controlli associati al sottodocumento
        '---------------------------------------------
        Me.txtDataStipulaAdeg.Value = 0
        Me.txtDataDecorrenzaAdeg.Value = 0
        Me.txtImportoAdeg.Value = 0
        Me.txtAnnotazioniAdeg.Text = ""
        Me.chkAdegContrProx.Value = 0
        Me.chkAdeguaContrAttuale.Value = 0
        Me.cboIvaAdeg.WriteOn 0
        Me.cboTipoAdeguamento.WriteOn 0
        Me.CDArticoloAdeg.Load 0
        Me.txtProtAdeg.Text = ""
        Me.chkIstatAdeguamento.Value = 0
        Me.txtNumeroAdeguamento.Value = 0
        Me.cboIstatAdeg.WriteOn 0
        Me.txtMaggiorazioneAdeg.Value = 0
        Me.txtDescrFattAde.Text = ""
        Me.cboTipoRateizzazioneAdeg.WriteOn 0
        Me.txtDataScadenzaAdeg.Value = 0
        chkNoCalcPeriodoFatt.Value = 0
        txtImportoAdegRinn.Value = 0
        txtNAdegIniz.Value = 0
        bValue = False
    End If
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
    Me.txtDataStipulaAdeg.Enabled = bValue
    Me.txtDataDecorrenzaAdeg.Enabled = bValue
    Me.txtImportoAdeg.Enabled = bValue
    Me.txtAnnotazioniAdeg.Enabled = bValue
    Me.chkAdegContrProx.Enabled = bValue
    Me.chkAdeguaContrAttuale.Enabled = bValue
    Me.cboIvaAdeg.Enabled = bValue
    Me.cboTipoAdeguamento.Enabled = bValue
    Me.CDArticoloAdeg.Enabled = bValue
    Me.txtProtAdeg.Enabled = bValue
    Me.chkIstatAdeguamento.Enabled = bValue
    Me.txtDescrFattAde.Enabled = bValue
    Me.cboTipoRateizzazioneAdeg.Enabled = bValue
    Me.txtDataScadenzaAdeg.Enabled = bValue
    chkNoCalcPeriodoFatt.Enabled = bValue
    txtImportoAdegRinn.Enabled = bValue
    txtNAdegIniz.Enabled = bValue
    'Me.txtNumeroAdeguamento.Enabled = bValue

    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
    Me.cmdNuovoAdeg.Enabled = True
    Me.cmdSalvaAdeg.Enabled = bValue
    Me.cmdEliminaAdeg.Enabled = bValue


'    'Adeguamento dei prodotti
'    With Me.cboAdegProd
'        Set .Database = m_App.Database.Connection
'        .AddFieldKey "IDRV_POContrattoAdeguamento"
'        .DisplayField = "DescrizioneAdeguamento"
'        .SQL = "SELECT * FROM RV_POContrattoAdeguamento "
'        .SQL = .SQL & "WHERE IDRV_POContrattoPadre=" & Me.txtIDContrattoPadre.Value
'        .SQL = .SQL & "ORDER BY NumeroAdeguamento "
'        .Fill
'    End With
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
            Me.cmdNuovoServizio.Enabled = True
            Me.cmdSalvaServizio.Enabled = bValue
            Me.cmdEliminaServizio.Enabled = bValue
    Else
        If fnNotNullN(m_Document("ContrattoAttuale").Value) = 0 Then
            Me.cmdNuovoAdeg.Enabled = False
            Me.cmdSalvaAdeg.Enabled = False
            Me.cmdEliminaAdeg.Enabled = False
        Else
            Me.cmdNuovoAdeg.Enabled = True
            Me.cmdSalvaAdeg.Enabled = bValue
            Me.cmdEliminaAdeg.Enabled = bValue
        End If
    End If
End Sub

Private Sub GENERA_RATE_ADEGUAMENTO(IDAdeguamentoContratto As Long, IDDurataContratto As Long, IDContratto As Long, IDContrattoPadre As Long, AnnoContratto As Long, DataDecorrenzaAdeg As String, DataScadenzaContratto As String, ImportoAdeg As Double, IDPagamentoRate As Long, NumeroAdeguamento As Long, ProtocolloAdeg As String, IDArticolo As Long, DescrizioneFatturazione As String)
On Error GoTo ERR_GENERA_RATE_ADEGUAMENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim NumeroGiorniContratto As Long
Dim NumeroGiorniRimanenti As Double
Dim ImportoGiornalieroAdeg As Double
Dim ImportoAdeguamentoTotale As Double
Dim NumeroRateDaPagare As Long
Dim ImportoAdegPerRata As Double
Dim NumeroRataElaborata As Long
Dim ImportoAdegProgressivo As Double
Dim ImportoDaRegistrare As Double
Dim ImportoAdegPerRataComp As Double
Dim ImportoAdegProgressivoComp As Double
Dim NumeroRataAdeguamento As Long

Dim IDOggettoScadenza As Long
Dim Periodo As String
Dim ArrayRate As String
Dim SplitArrayRate() As String
Dim I As Integer
Dim Errore As String

ArrayRate = ""
Errore = "Elaborazione"

''''CLICLO DI ELIMINAZIONE IDOGGETTI DELLE RATE DI ADEGUAMENTO'''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento=" & IDAdeguamentoContratto

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    
    IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), 0)
    
    If IDOggettoScadenza > 0 Then
        ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)
        ELIMINA_SCADENZA IDOggettoScadenza
    End If
    
    ELIMINA_LINK_OGGETTO_RATA fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)

rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ELIMINAZIONE RATE DELL'ADEGUAMENTO DEL CONTRATTO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento=" & IDAdeguamentoContratto
Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Year(DataScadenzaContratto) Mod 4 <> 0 Then
    DataScadenzaContratto = DateAdd("d", 1, DataScadenzaContratto)
    If Year(DataDecorrenzaAdeg) Mod 4 <> 0 Then
        DataDecorrenzaAdeg = DataDecorrenzaAdeg
    Else
        If DateDiff("d", DataDecorrenzaAdeg, "29/02/" & Year(DataDecorrenzaAdeg)) >= 0 Then
            DataDecorrenzaAdeg = DateAdd("d", 1, DataDecorrenzaAdeg)
        Else
            DataDecorrenzaAdeg = DataDecorrenzaAdeg
        End If
    End If
    
    NumeroGiorniContratto = GET_NUMERO_GIORNI_CONTRATTO(IDDurataContratto, AnnoContratto)
    
    NumeroGiorniRimanenti = DateDiff("d", DataDecorrenzaAdeg, DataScadenzaContratto)

    ImportoGiornalieroAdeg = ImportoAdeg / NumeroGiorniContratto
    
    If DatePart("yyyy", DataScadenzaContratto) Mod 4 = 0 Then
        ImportoAdeguamentoTotale = ImportoGiornalieroAdeg * (NumeroGiorniRimanenti)
    Else
        ImportoAdeguamentoTotale = ImportoGiornalieroAdeg * NumeroGiorniRimanenti
    End If
Else
    If DateDiff("d", DataScadenzaContratto, "29/02/" & Year(DataScadenzaContratto)) <= 0 Then
    
        DataScadenzaContratto = DataScadenzaContratto
        
        NumeroGiorniContratto = GET_NUMERO_GIORNI_CONTRATTO(IDDurataContratto, AnnoContratto)
        
        NumeroGiorniRimanenti = DateDiff("d", DataDecorrenzaAdeg, DataScadenzaContratto)
    
        ImportoGiornalieroAdeg = ImportoAdeg / NumeroGiorniContratto
        
        If DatePart("yyyy", DataScadenzaContratto) Mod 4 = 0 Then
            ImportoAdeguamentoTotale = ImportoGiornalieroAdeg * (NumeroGiorniRimanenti)
        Else
            ImportoAdeguamentoTotale = ImportoGiornalieroAdeg * NumeroGiorniRimanenti
        End If
    Else
    DataScadenzaContratto = DateAdd("d", 1, DataScadenzaContratto)
    If Year(DataDecorrenzaAdeg) Mod 4 <> 0 Then
        DataDecorrenzaAdeg = DataDecorrenzaAdeg
    Else
        If DateDiff("d", DataDecorrenzaAdeg, "29/02/" & Year(DataDecorrenzaAdeg)) >= 0 Then
            DataDecorrenzaAdeg = DateAdd("d", 1, DataDecorrenzaAdeg)
        Else
            DataDecorrenzaAdeg = DataDecorrenzaAdeg
        End If
    End If
    
    NumeroGiorniContratto = GET_NUMERO_GIORNI_CONTRATTO(IDDurataContratto, AnnoContratto)
    
    NumeroGiorniRimanenti = DateDiff("d", DataDecorrenzaAdeg, DataScadenzaContratto)

    ImportoGiornalieroAdeg = ImportoAdeg / NumeroGiorniContratto
    
    If DatePart("yyyy", DataScadenzaContratto) Mod 4 = 0 Then
        ImportoAdeguamentoTotale = ImportoGiornalieroAdeg * (NumeroGiorniRimanenti)
    Else
        ImportoAdeguamentoTotale = ImportoGiornalieroAdeg * NumeroGiorniRimanenti
    End If
    End If
    
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT COUNT(IDRV_PORateContratto) AS NumeroRecord "
sSQL = sSQL & "FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Fatturata=0"
sSQL = sSQL & " AND DataRata>=" & fnNormDate(DataDecorrenzaAdeg)
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    NumeroRateDaPagare = 0
Else
    NumeroRateDaPagare = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=0"

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
NumeroRataAdeguamento = 1

If NumeroRateDaPagare > 0 Then
    ImportoAdegPerRataComp = ImportoAdeguamentoTotale / NumeroRateDaPagare
    ImportoAdegPerRata = FormatNumber(ImportoAdegPerRataComp, 2)
    ImportoAdegProgressivo = ImportoAdegPerRata
    ImportoAdegProgressivoComp = ImportoAdegPerRataComp
    
    sSQL = "SELECT * FROM RV_PORateContratto "
    sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
    sSQL = sSQL & " AND Fatturata=0"
    sSQL = sSQL & " AND DataRata>=" & fnNormDate(DataDecorrenzaAdeg)
    sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"
    sSQL = sSQL & " ORDER BY DataRata "
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        If ImportoAdeguamentoTotale > 0 Then
            If ImportoAdegProgressivoComp >= ImportoAdeguamentoTotale Then
                ImportoDaRegistrare = FormatNumber((ImportoAdeguamentoTotale - (ImportoAdegProgressivo - ImportoAdegPerRataComp)), 2)
            Else
                ImportoDaRegistrare = ImportoAdegPerRata
            End If
        Else
            If ImportoAdegProgressivoComp <= ImportoAdeguamentoTotale Then
                ImportoDaRegistrare = FormatNumber((ImportoAdeguamentoTotale - (ImportoAdegProgressivo - ImportoAdegPerRataComp)), 2)
            Else
                ImportoDaRegistrare = ImportoAdegPerRata
            End If
        End If
        rsNew.AddNew
            rsNew!IDRV_PORateContratto = fnGetNewKey("RV_PORateContratto", "IDRV_PORateContratto")
            rsNew!IDRV_POContratto = IDContratto
            rsNew!numerorata = fnNotNullN(rs!numerorata)
            rsNew!DataRata = rs!DataRata
            rsNew!IDPagamentoRata = fnNotNullN(rs!IDPagamentoRata)
            rsNew!ImportoRata = ImportoDaRegistrare
            rsNew!Fatturata = 0
            rsNew!IDOggettoCollegato = 0
            rsNew!IDTipoOggettoCollegato = 0
            rsNew!mese = DatePart("m", fnNotNull(rs!DataRata))
            rsNew!Anno = DatePart("yyyy", fnNotNull(rs!DataRata))
            rsNew!Periodo = Mid(DescrizioneFatturazione, 1, 250)
            rsNew!Manuale = 0
            rsNew!ContrattoAttuale = 1
            rsNew!IDRV_POContrattoPadre = IDContrattoPadre
            rsNew!IDRV_POContrattoAdeguamento = IDAdeguamentoContratto
            rsNew!IDArticolo = IDArticolo
            rsNew!IDTipoOggetto = fnGetTipoOggetto("RV_PORateContratto")
            rsNew!IDOggetto = GET_LINK_OGGETTO(0, rsNew!IDTipoOggetto, rsNew!numerorata, rsNew!DataRata)
            If NumeroRataAdeguamento = 1 Then
                rsNew!DataInizioPeriodo = DataDecorrenzaAdeg
            Else
                rsNew!DataInizioPeriodo = rs!DataInizioPeriodo
            End If
            rsNew!DataFinePeriodo = rs!DataFinePeriodo
            rsNew!NonFatturare = 0
            rsNew!AnnotazioniNonFatturare = ""
            
            If Len(ArrayRate) > 0 Then
                ArrayRate = ArrayRate & "|"
            End If
            ArrayRate = ArrayRate & fnNotNullN(rsNew!IDRV_PORateContratto)
            
            
        rsNew.Update
        
        
        
        ''''''''''''''''FLUSSO DOCUMENTALE SCADENZARIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(IDOggettoRata, IDTipoOggettoRata, 0)
        If LINK_SEZIONALE_RATE > 0 Then
            IDOggettoScadenza = GET_LINK_SCADENZA(rsNew!ImportoRata, IDClienteFatturazione, rsNew!numerorata, rsNew!DataRata, LINK_SEZIONALE_RATE, rsNew!Periodo)
            CREA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, rsNew!IDOggetto, rsNew!IDTipoOggetto
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        

        ImportoAdegProgressivo = ImportoAdegProgressivo + ImportoAdegPerRata
        ImportoAdegProgressivoComp = ImportoAdegProgressivoComp + ImportoAdegPerRataComp
        NumeroRataAdeguamento = NumeroRataAdeguamento + 1
        DoEvents
        
    rs.MoveNext
    Wend
    rs.CloseResultset
    Set rs = Nothing
       
Else
    
    rsNew.AddNew
        rsNew!IDRV_PORateContratto = fnGetNewKey("RV_PORateContratto", "IDRV_PORateContratto")
        rsNew!IDRV_POContratto = IDContratto
        rsNew!numerorata = GET_NUMERO_RATA_CONTRATTO(IDContratto)
        rsNew!DataRata = DataDecorrenzaAdeg
        rsNew!IDPagamentoRata = IDPagamentoRate
        rsNew!ImportoRata = ImportoAdeguamentoTotale
        rsNew!Fatturata = 0
        rsNew!IDOggettoCollegato = 0
        rsNew!IDTipoOggettoCollegato = 0
        rsNew!mese = DatePart("m", DataDecorrenzaAdeg)
        rsNew!Anno = DatePart("yyyy", DataDecorrenzaAdeg)
        rsNew!Periodo = Mid(DescrizioneFatturazione, 1, 250)
        rsNew!Manuale = 0
        rsNew!ContrattoAttuale = 1
        rsNew!IDRV_POContrattoPadre = IDContrattoPadre
        rsNew!IDRV_POContrattoAdeguamento = IDAdeguamentoContratto
        rsNew!IDArticolo = IDArticolo
        rsNew!IDTipoOggetto = fnGetTipoOggetto("RV_PORateContratto")
        rsNew!IDOggetto = GET_LINK_OGGETTO(0, rsNew!IDTipoOggetto, rsNew!numerorata, rsNew!DataRata)
        rsNew!DataInizioPeriodo = DataDecorrenzaAdeg
        rsNew!DataFinePeriodo = DataScadenzaContratto
        rsNew!NonFatturare = 0
        rsNew!AnnotazioniNonFatturare = ""
        
        If Len(ArrayRate) > 0 Then
            ArrayRate = ArrayRate & "|"
        End If
        
        ArrayRate = ArrayRate & fnNotNullN(rsNew!IDRV_PORateContratto)

    rsNew.Update
    

    If LINK_SEZIONALE_RATE > 0 Then

        ''''''''''''''''FLUSSO DOCUMENTALE SCADENZARIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
        'IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(IDOggettoRata, IDTipoOggettoRata, 0)
         
        IDOggettoScadenza = GET_LINK_SCADENZA(rsNew!ImportoRata, IDClienteFatturazione, rsNew!numerorata, rsNew!DataRata, LINK_SEZIONALE_RATE, rsNew!Periodo)
        
        CREA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, rsNew!IDOggetto, rsNew!IDTipoOggetto
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
End If

rsNew.Close
Set rsNew = Nothing

Errore = "Stringa per fatturazione"

SplitArrayRate = Split(ArrayRate, "|")

For I = 0 To UBound(SplitArrayRate)
    
    If Me.chkNoCalcPeriodoFatt.Value = 0 Then
        Periodo = GET_STRINGA_PERIODO_ADEG(2, TheApp.Branch, m_Document(m_Document.PrimaryKey), CLng(SplitArrayRate(I)), GET_LINK_ADEGUAMENTO_DA_RATA(CLng(SplitArrayRate(I))), 0)
    Else
        Periodo = Me.CDArticoloAdeg.Description
    End If
    
    Periodo = Mid(Periodo, 1, 250)
    sSQL = "UPDATE RV_PORateContratto SET "
    sSQL = sSQL & "Periodo=" & fnNormString(Periodo)
    sSQL = sSQL & "WHERE IDRV_PORateContratto=" & SplitArrayRate(I)
    Cn.Execute sSQL
        
Next

Exit Sub
ERR_GENERA_RATE_ADEGUAMENTO:
    MsgBox Err.Description, vbCritical, "GENERA_RATE_ADEGUAMENTO (" & Errore & ")"
    

End Sub

Private Function GET_NUMERO_GIORNI_CONTRATTO(IDDurataContratto As Long, AnnoContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroGiorniAnno As Long

If AnnoContratto Mod 4 <> 0 Then
    NumeroGiorniAnno = 365
Else
    NumeroGiorniAnno = 365
End If

sSQL = "SELECT Mesi FROM RV_POTipoRinnovo "
sSQL = sSQL & "WHERE IDRV_POTipoRinnovo=" & IDDurataContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_GIORNI_CONTRATTO = NumeroGiorniAnno
Else
    GET_NUMERO_GIORNI_CONTRATTO = (NumeroGiorniAnno / (12 / fnNotNullN(rs!Mesi)))
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_NUMERO_RATA_CONTRATTO(IDContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(NumeroRata) as MaxNumeroRata "
sSQL = sSQL & "FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_RATA_CONTRATTO = 1
Else
    GET_NUMERO_RATA_CONTRATTO = fnNotNullN(rs!MaxNumeroRata) + 1
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_RATA_PAGATA_ADEGUAMENTO(IDAdeguamentoContratto As Long, IDContratto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Count(IDRV_PORateContratto) AS NumeroRate "
sSQL = sSQL & "FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContrattoAdeguamento=" & IDAdeguamentoContratto
sSQL = sSQL & " AND IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Fatturata=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_RATA_PAGATA_ADEGUAMENTO = False
Else
    If fnNotNullN(rs!numerorate) > 0 Then
        GET_RATA_PAGATA_ADEGUAMENTO = True
    Else
        GET_RATA_PAGATA_ADEGUAMENTO = False
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_NUMERO_ADEGUAMENTO(IDContrattoPadre As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT Max(NumeroAdeguamento) as NumeroRecord "
sSQL = sSQL & "FROM RV_POContrattoAdeguamento "
sSQL = sSQL & "WHERE IDRV_POContrattoPadre=" & IDContrattoPadre

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_ADEGUAMENTO = 1
Else
    GET_NUMERO_ADEGUAMENTO = fnNotNullN(rs!NumeroRecord) + 1
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub GET_PREZZO_ARTICOLO(IDArticolo As Long, IDListinoCliente As Long, IDListinoAzienda As Long, IDAnagraficaCliente As Long)

oDoc.ClearValues

oDoc.Tables(sTabellaDettaglio).SetActiveRetail oDoc.Tables(sTabellaDettaglio).NumRetails
oDoc.ReadDataFromCliFo IDAnagraficaCliente
oDoc.DataEmissione = Me.txtDataInizioProd.Value
oDoc.Field "Doc_data", oDoc.DataEmissione, sTabellaTestata
oDoc.Field "Link_Val_valuta", 9, sTabellaTestata

oDoc.ReadDataFromArticle IDArticolo, sTabellaDettaglio
oDoc.Field "Link_Doc_listino", IDListinoCliente, sTabellaTestata
oDoc.Field "Link_Doc_listino_base", IDListinoAzienda, sTabellaTestata

oDoc.Field "Art_quantita_totale", Me.txtQtaArtProd.Value, sTabellaDettaglio
        
oDoc.ReadDataFromPriceList IDListinoCliente
oDoc.ReadDataFromDiscountsList
        
Me.txtImpUniProd.Value = fnNotNullN(oDoc.Field("Art_prezzo_unitario_neutro", , sTabellaDettaglio))
        
Me.txtSconto1Prod.Value = fnNotNullN(oDoc.Field("Art_sco_in_percentuale_1", , sTabellaDettaglio))
Me.txtSconto2Prod.Value = fnNotNullN(oDoc.Field("Art_sco_in_percentuale_2", , sTabellaDettaglio))

GET_TOTALI_RIGA_DETTAGLIO

End Sub
Private Function GetAttivitaAzienda(IDAzienda As Long, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AttivitaAzienda.IDAttivitaAzienda, Azienda.IDAzienda, Filiale.IDFiliale "
sSQL = sSQL & "FROM AttivitaAzienda INNER JOIN "
sSQL = sSQL & "Azienda ON AttivitaAzienda.IDAzienda = Azienda.IDAzienda INNER JOIN "
sSQL = sSQL & "Filiale ON AttivitaAzienda.IDAttivitaAzienda = Filiale.IDAttivitaAzienda "
sSQL = sSQL & "WHERE Azienda.IDAzienda =" & IDAzienda
sSQL = sSQL & " AND Filiale.IDFiliale = " & IDFiliale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GetAttivitaAzienda = 0
Else
    GetAttivitaAzienda = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_LISTINO(IDAnagraficaCliente As Long, IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
If IDAnagraficaCliente > 0 Then
    sSQL = "Select IDListinoDefault From Cliente "
    sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
    sSQL = sSQL & " AND IDAnagrafica=" & IDAnagraficaCliente
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        GET_LINK_LISTINO = fnNotNullN(rs!IDListinoDefault)
        If GET_LINK_LISTINO = 0 Then
            rs.CloseResultset
            Set rs = Nothing
            
            sSQL = "SELECT IDListinoDiBase FROM ConfigurazioneVendite WHERE IDAzienda=" & IDAzienda
            
            Set rs = Cn.OpenResultset(sSQL)
            
            If rs.EOF = False Then
                GET_LINK_LISTINO = fnNotNullN(rs!IDListinoDiBase)
            End If
            
            rs.CloseResultset
            Set rs = Nothing
           
        End If
        
    Else
        rs.CloseResultset
        Set rs = Nothing
        
        sSQL = "SELECT IDListinoDiBase FROM ConfigurazioneVendite WHERE IDAzienda=" & IDAzienda
        Set rs = Cn.OpenResultset(sSQL)
        
        If rs.EOF = False Then
            GET_LINK_LISTINO = fnNotNullN(rs!IDListinoDiBase)
        
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        
    End If
    
    
Else

    sSQL = "SELECT IDListinoDiBase FROM ConfigurazioneVendite WHERE IDAzienda=" & IDAzienda
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        GET_LINK_LISTINO = fnNotNullN(rs!IDListinoDiBase)
    
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If
End Function
Private Function GET_LINK_LISTINO_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDListinoDiBase FROM ConfigurazioneVendite WHERE IDAzienda=" & IDAzienda

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF = False Then
    GET_LINK_LISTINO_AZIENDA = fnNotNullN(rs!IDListinoDiBase)
Else
    GET_LINK_LISTINO_AZIENDA = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_TOTALI_RIGA_DETTAGLIO()
    
    GET_TOTALE_RIGA
    
    
End Sub
Private Function GET_ALIQUOTA_IVA(IDIva As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AliquotaIva FROM Iva "
sSQL = sSQL & "WHERE IDIva=" & IDIva

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ALIQUOTA_IVA = 0
Else
    GET_ALIQUOTA_IVA = fnNotNullN(rs!AliquotaIva)
End If


rs.CloseResultset
Set rs = Nothing
End Function

Private Sub txtOraInizioProd_LostFocus()
    txtQtaPeriodo_LostFocus
End Sub

Private Sub txtQtaArtProd_LostFocus()
    GET_TOTALI_RIGA_DETTAGLIO
End Sub

Private Sub txtQtaPeriodo_LostFocus()
On Error GoTo ERR_txtQtaPeriodo_LostFocus
    
    'If Me.cboTipoPeriodo.CurrentID = 2 Then Exit Sub
    
    Me.txtOraFineProd.Value = 0
    'Me.txtOraInizioProd.Value = 0
    'If Me.cboTipoPeriodo.CurrentID = 1 Then
    '    Me.txtDataInizioProd.Value = Date
    'End If
    
    Select Case Me.cboUMPeriodoProd.CurrentID
        Case 1
            Dim Datafinale As String
            Dim ore As Long
            Dim MinutiArray() As String
            Dim Minuti As Long
            
            ore = Me.txtQtaPeriodo.Value
            Minuti = 0
            
            If (Me.txtQtaPeriodo.Value - ore) > 0 Then
                Minuti = 60 / (1 / (Me.txtQtaPeriodo.Value - ore))
            End If
            
            Datafinale = DateAdd("h", ore, Me.txtDataInizioProd.Text + " " + Me.txtOraInizioProd.Text)
            Datafinale = DateAdd("n", Minuti, Datafinale)
            
            Me.txtDataFineProd.Value = Me.txtDataInizioProd.Value
            Me.txtOraFineProd.Text = DatePart("h", Datafinale) & ":" & DatePart("n", Datafinale)
            
            
        Case 2
            Me.txtDataFineProd.Value = DateAdd("d", Me.txtQtaPeriodo.Value, Me.txtDataInizioProd.Text) - 1
            
            
            
        Case 3
            Me.txtDataFineProd.Value = DateAdd("ww", Me.txtQtaPeriodo.Value, Me.txtDataInizioProd.Text)
            
        Case 4
            Me.txtDataFineProd.Value = DateAdd("m", Me.txtQtaPeriodo.Value, Me.txtDataInizioProd.Text)
            Me.txtDataFineProd.Value = DateAdd("d", -1, Me.txtDataFineProd.Text)
        Case 5
            Me.txtDataFineProd.Value = DateAdd("yyyy", Me.txtQtaPeriodo.Value, Me.txtDataInizioProd.Text)
            Me.txtDataFineProd.Value = DateAdd("d", -1, Me.txtDataFineProd.Text)
        Case 6
            
    End Select
    
    Me.txtQuantitaEffettiva.Value = Me.txtQtaPeriodo.Value
    
    If Me.cboUMPeriodoProd.CurrentID = 2 Then
        Me.txtQuantitaEffettiva.Value = GET_CALCOLO_QUANTITA_EFFETTIVA(Me.txtDataInizioProd.Text, Me.txtDataFineProd.Text)
        
    End If
    
    GET_TOTALI_RIGA_DETTAGLIO
    
    Exit Sub
ERR_txtQtaPeriodo_LostFocus:
    
    MsgBox Err.Description, vbCritical, "txtQtaPeriodo_LostFocus"


    
End Sub

Private Sub txtQtaProdotto_LostFocus()
    GET_TOTALI_RIGA_DETTAGLIO
End Sub

Private Sub txtQuantitaEffettiva_Change()
    If bLoadingProdotti = True Then Exit Sub
    GET_TOTALI_RIGA_DETTAGLIO
End Sub

Private Sub txtSconto1Prod_LostFocus()
    GET_TOTALI_RIGA_DETTAGLIO
End Sub

Private Sub txtSconto2Prod_LostFocus()
    GET_TOTALI_RIGA_DETTAGLIO
End Sub
Private Function GET_LINK_OGGETTO(IDOggetto As Long, IDTipoOggetto As Long, numerorata As Long, DataScadenzaRata As String) As Long
On Error GoTo ERR_GET_LINK_OGGETTO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim IDOggettoLocal As Long
Dim IDFunzione As Long

GET_LINK_OGGETTO = 0

IDFunzione = GET_LINK_FUNZIONE(IDTipoOggetto)

If IDFunzione = 0 Then Exit Function

sSQL = "SELECT * FROM Oggetto "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = True Then
    rs.AddNew
        rs!IDTipoOggetto = IDTipoOggetto
        rs!IDFunzione = IDFunzione
        rs!IDAzienda = TheApp.IDFirm
        rs!IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.Branch)
        rs!IDSezionale = 0
        rs!Oggetto = GET_DESCRIZIONE_FUNZIONE(IDFunzione)
        rs!DataEmissione = DataScadenzaRata
        rs!numero = numerorata
        rs!DataUltimaVariazione = Date
        rs!IDUtenteUltimaVariazione = TheApp.IDUser
        rs!VirtualDelete = 0
        rs!IDOggetto = fnGetNewKey("Oggetto", "IDOggetto")
        GET_LINK_OGGETTO = rs!IDOggetto
    rs.Update
Else
    rs!DataEmissione = DataScadenzaRata
    rs!numero = numerorata
    rs!DataUltimaVariazione = Date
    rs!IDUtenteUltimaVariazione = TheApp.IDUser
    rs!VirtualDelete = 0
    GET_LINK_OGGETTO = rs!IDOggetto
    rs.Update
End If

rs.Close
Set rs = Nothing

Exit Function

ERR_GET_LINK_OGGETTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_OGGETTO"
    GET_LINK_OGGETTO = 0
End Function
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
Private Function GET_DESCRIZIONE_FUNZIONE(IDFunzione As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Funzione FROM Funzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzione

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_FUNZIONE = ""
Else
    GET_DESCRIZIONE_FUNZIONE = fnNotNull(rs!Funzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ELIMINA_FLUSSO_DOCUMENTALE(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long)
On Error GoTo ERR_ELIMINA_FLUSSO_DOCUMENTALE
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim IDFunzioneVend As Long
Dim IDFunzioneRata As Long
Dim IDFlussoGruppo As Long
Dim IDFlussoFunzione As Long

IDFunzioneVend = GET_LINK_FUNZIONE(IDTipoOggettoVend)
IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)


'''''''''''''''''''''''''''''''''GRUPPO FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoGruppo "
sSQL = sSQL & "WHERE Descrizione=" & fnNormString("Documento di vendita -> Rata contratto")
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFLussoGruppo")
        rsNew!Descrizione = "Documento di vendita -> Rata contratto"
    rsNew.Update
End If

IDFlussoGruppo = fnNotNullN(rsNew!IDFlussoGruppo)

rsNew.Close
Set rsNew = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IDFunzioneVend > 0 Then
    sSQL = "SELECT * FROM FlussoFunzione "
    sSQL = sSQL & "WHERE IDFunzione=" & IDFunzioneVend
    sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneRata
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDFlussoFunzione = fnGetNewKeyTipoOggetto("FlussoFunzione", "IDFlussoFunzione")
            rsNew!IDFunzione = IDFunzioneVend
            rsNew!IDFunzioneSuccessiva = IDFunzioneRata
            rsNew!Cardinalita = 3
            rsNew!TipoAutomatismo = 1
            rsNew!Attributo = 14
            rsNew!TipoDipendenza = 1
            rsNew!IDFlussoGruppo = IDFlussoGruppo
        rsNew.Update
    End If
    
    IDFlussoFunzione = fnNotNullN(rsNew!IDFlussoFunzione)
    
    rsNew.Close
    Set rsNew = Nothing
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''FLUSSO FUNZIONE COLLEGATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggetto=" & IDOggettoVend
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggettoCollegato<>" & IDOggettoRata

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    sSQL = "DELETE FROM FlussoFunzioneCollegato "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoVend
    sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoVend
    sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
    Cn.Execute sSQL
End If
rsNew.Close
Set rsNew = Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''FLUSSO OGGETTI COLLEGATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggetto=" & IDOggettoVend
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoRata
Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Sub
ERR_ELIMINA_FLUSSO_DOCUMENTALE:
    MsgBox Err.Description, vbCritical, "ELIMINA_FLUSSO_DOCUMENTALE"
    
End Sub
Private Sub CREA_FLUSSO_DOCUMENTALE(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long)
On Error GoTo ERR_CREA_FLUSSO_DOCUMENTALE
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim IDFunzioneVend As Long
Dim IDFunzioneRata As Long
Dim IDFlussoGruppo As Long
Dim IDFlussoFunzione As Long

IDFunzioneVend = GET_LINK_FUNZIONE(IDTipoOggettoVend)
IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)

If IDFunzioneVend = 0 Then Exit Sub
If IDFunzioneRata = 0 Then Exit Sub
'''''''''''''''''''''''''''''''''GRUPPO FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoGruppo "
sSQL = sSQL & "WHERE Descrizione=" & fnNormString("Documento di vendita -> Rata contratto")
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFlussoGruppo")
        rsNew!Descrizione = "Documento di vendita -> Rata contratto"
    rsNew.Update
End If

IDFlussoGruppo = fnNotNullN(rsNew!IDFlussoGruppo)

rsNew.Close
Set rsNew = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoFunzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzioneVend
sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneRata
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoFunzione = fnGetNewKeyTipoOggetto("FlussoFunzione", "IDFlussoFunzione")
        rsNew!IDFunzione = IDFunzioneVend
        rsNew!IDFunzioneSuccessiva = IDFunzioneRata
        rsNew!Cardinalita = 3
        rsNew!TipoAutomatismo = 1
        rsNew!Attributo = 14
        rsNew!TipoDipendenza = 1
        rsNew!IDFlussoGruppo = IDFlussoGruppo
    rsNew.Update
End If

IDFlussoFunzione = fnNotNullN(rsNew!IDFlussoFunzione)

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE COLLEGATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoFunzioneCollegato "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoVend
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoFunzione = IDFlussoFunzione
        rsNew!IDOggetto = IDOggettoVend
        rsNew!IDTipoOggetto = IDTipoOggettoVend
End If

rsNew!FlussoFunzioneCollegato = 2
rsNew.Update

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FLUSSO OGGETTI COLLEGATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggetto=" & IDOggettoVend
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoRata

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
    rsNew!IDFlussoFunzione = IDFlussoFunzione
    rsNew!IDOggetto = IDOggettoVend
    rsNew!IDTipoOggetto = IDTipoOggettoVend
    rsNew!IDTipoOggettoCollegato = IDTipoOggettoRata
    rsNew!IDOggettoCollegato = IDOggettoRata
    rsNew.Update
End If

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Sub
ERR_CREA_FLUSSO_DOCUMENTALE:
MsgBox Err.Description, vbCritical, "CREA_FLUSSO_DOCUMENTALE"
End Sub
Private Function GET_LINK_ATTIVITA_AZIENDA(IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAttivitaAzienda FROM Filiale "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ATTIVITA_AZIENDA = 0
Else
    GET_LINK_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function ELIMINA_LINK_OGGETTO_RATA(IDOggetto As Long, IDTipoOggetto As Long)
On Error GoTo ERR_ELIMINA_LINK_OGGETTO_RATA
Dim sSQL As String

sSQL = "DELETE FROM Oggetto "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
Cn.Execute sSQL

Exit Function
ERR_ELIMINA_LINK_OGGETTO_RATA:
    MsgBox Err.Description, vbCritical, "ELIMINA_LINK_OGGETTO_RATA"

End Function
Private Function GET_RATA_PAGATA(IDContratto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Count(IDRV_PORateContratto) AS NumeroRate "
sSQL = sSQL & "FROM RV_PORateContratto "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Fatturata=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_RATA_PAGATA = False
Else
    If fnNotNullN(rs!numerorate) > 0 Then
        GET_RATA_PAGATA = True
    Else
        GET_RATA_PAGATA = False
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_LINK_OGGETTO_SCADENZA_COLLEGATA(IDOggettoRata As Long, IDTipoOggettoRata As Long, IDSezionale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDOggettoScadenza As Long
Dim IDFunzioneRata As Long
Dim IDFunzioneScadenza As Long
Dim IDFlussoFunzione As Long

IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)
IDFunzioneScadenza = GET_LINK_FUNZIONE(131)
IDFlussoFunzione = GET_LINK_FLUSSO_FUNZIONE(IDFunzioneRata, IDFunzioneScadenza)

If IDFlussoFunzione = 0 Then
    GET_LINK_OGGETTO_SCADENZA_COLLEGATA = 0
    Exit Function
End If

sSQL = "SELECT IDOggettoCollegato "
sSQL = sSQL & "FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=131"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_OGGETTO_SCADENZA_COLLEGATA = 0
Else
    GET_LINK_OGGETTO_SCADENZA_COLLEGATA = fnNotNullN(rs!IDOggettoCollegato)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_FLUSSO_FUNZIONE(IDFunzione As Long, IDFunzioneSucc As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDFlussoFunzione FROM FlussoFunzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzione
sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneSucc

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_FLUSSO_FUNZIONE = 0
Else
    GET_LINK_FLUSSO_FUNZIONE = fnNotNullN(rs!IDFlussoFunzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long)
On Error GoTo ERR_ELIMINA_FLUSSO_DOCUMENTALE
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim IDFunzioneVend As Long
Dim IDFunzioneRata As Long
Dim IDFlussoGruppo As Long
Dim IDFlussoFunzione As Long

IDFunzioneVend = GET_LINK_FUNZIONE(IDTipoOggettoVend)
IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)

'''''''''''''''''''''''''''''''''GRUPPO FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoGruppo "
sSQL = sSQL & "WHERE Descrizione=" & fnNormString("Rata contratto -> Scadenza")
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFLussoGruppo")
        rsNew!Descrizione = "Rata contratto -> Scadenza"
    rsNew.Update
End If

IDFlussoGruppo = fnNotNullN(rsNew!IDFlussoGruppo)

rsNew.Close
Set rsNew = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IDFunzioneVend > 0 Then
    sSQL = "SELECT * FROM FlussoFunzione "
    sSQL = sSQL & "WHERE IDFunzione=" & IDFunzioneRata
    sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneVend
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDFlussoFunzione = fnGetNewKeyTipoOggetto("FlussoFunzione", "IDFlussoFunzione")
            rsNew!IDFunzione = IDFunzioneVend
            rsNew!IDFunzioneSuccessiva = IDFunzioneRata
            rsNew!Cardinalita = 3
            rsNew!TipoAutomatismo = 1
            rsNew!Attributo = 14
            rsNew!TipoDipendenza = 1
            rsNew!IDFlussoGruppo = IDFlussoGruppo
        rsNew.Update
    End If
    
    IDFlussoFunzione = fnNotNullN(rsNew!IDFlussoFunzione)
    
    rsNew.Close
    Set rsNew = Nothing
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE COLLEGATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggettoCollegato<>" & IDOggettoVend

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    sSQL = "DELETE FROM FlussoFunzioneCollegato "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoRata
    sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
    sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
    Cn.Execute sSQL
End If
rsNew.Close
Set rsNew = Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''FLUSSO OGGETTI COLLEGATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoVend
Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Sub
ERR_ELIMINA_FLUSSO_DOCUMENTALE:
    MsgBox Err.Description, vbCritical, "ELIMINA_FLUSSO_DOCUMENTALE"
    
End Sub
Private Sub CREA_FLUSSO_DOCUMENTALE_SCADENZA(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long)
On Error GoTo ERR_CREA_FLUSSO_DOCUMENTALE_SCADENZA
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim IDFunzioneVend As Long
Dim IDFunzioneRata As Long
Dim IDFlussoGruppo As Long
Dim IDFlussoFunzione As Long

IDFunzioneVend = GET_LINK_FUNZIONE(IDTipoOggettoVend)
IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)

If IDFunzioneVend = 0 Then Exit Sub
If IDFunzioneRata = 0 Then Exit Sub
'''''''''''''''''''''''''''''''''GRUPPO FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoGruppo "
sSQL = sSQL & "WHERE Descrizione=" & fnNormString("Rata contratto -> Scadenza")
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFlussoGruppo")
        rsNew!Descrizione = "Rata contratto -> Scadenza"
    rsNew.Update
End If

IDFlussoGruppo = fnNotNullN(rsNew!IDFlussoGruppo)

rsNew.Close
Set rsNew = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoFunzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzioneRata
sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneVend
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoFunzione = fnGetNewKeyTipoOggetto("FlussoFunzione", "IDFlussoFunzione")
        rsNew!IDFunzione = IDFunzioneRata
        rsNew!IDFunzioneSuccessiva = IDFunzioneVend
        rsNew!Cardinalita = 3
        rsNew!TipoAutomatismo = 1
        rsNew!Attributo = 14
        rsNew!TipoDipendenza = 1
        rsNew!IDFlussoGruppo = IDFlussoGruppo
    rsNew.Update
End If

IDFlussoFunzione = fnNotNullN(rsNew!IDFlussoFunzione)

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE COLLEGATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoFunzioneCollegato "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoFunzione = IDFlussoFunzione
        rsNew!IDOggetto = IDOggettoRata
        rsNew!IDTipoOggetto = IDTipoOggettoRata
End If

rsNew!FlussoFunzioneCollegato = 2
rsNew.Update

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FLUSSO OGGETTI COLLEGATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoVend

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
    rsNew!IDFlussoFunzione = IDFlussoFunzione
    rsNew!IDOggetto = IDOggettoRata
    rsNew!IDTipoOggetto = IDTipoOggettoRata
    rsNew!IDTipoOggettoCollegato = IDTipoOggettoVend
    rsNew!IDOggettoCollegato = IDOggettoVend
    rsNew.Update
End If

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Sub
ERR_CREA_FLUSSO_DOCUMENTALE_SCADENZA:
MsgBox Err.Description, vbCritical, "CREA_FLUSSO_DOCUMENTALE_SCADENZA"
End Sub

Private Sub ELIMINA_SCADENZA(IDOggettoScadenza As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDTestataScadenza As Long


sSQL = "SELECT IDTestataScadenza FROM TestataScadenza "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoScadenza

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    IDTestataScadenza = 0
Else
    IDTestataScadenza = fnNotNullN(rs!IDTestataScadenza)
End If


rs.CloseResultset
Set rs = Nothing

If IDTestataScadenza > 0 Then
    sSQL = "DELETE FROM TestataScadenza "
    sSQL = sSQL & "WHERE IDTestataScadenza=" & IDTestataScadenza
    Cn.Execute sSQL

    sSQL = "DELETE FROM DettaglioScadenza "
    sSQL = sSQL & "WHERE IDTestataScadenza=" & IDTestataScadenza
    Cn.Execute sSQL

    sSQL = "DELETE FROM Oggetto "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoScadenza
    Cn.Execute sSQL
    
End If

End Sub

Private Function GET_LINK_SCADENZA(ImportoComplessivoScadenza As Double, IDAnagrafica As Long, NumeroDocumento As String, DataDocumento As String, IDSezionale As Long, Periodo As String) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim Link_Oggetto As Long
Dim IDTestataScadenza As Long


Link_Oggetto = GET_LINK_OGGETTO_SCADENZA(DataDocumento, IDSezionale, NumeroDocumento)

If Link_Oggetto > 0 Then
        
    Set rs = New ADODB.Recordset
    sSQL = "SELECT * FROM TestataScadenza "
    sSQL = sSQL & "WHERE IDOggetto=" & Link_Oggetto
    
    rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    rs.AddNew
        rs!IDTestataScadenza = fnGetNewKey("TestataScadenza", "IDTestataScadenza")
        rs!IDOggetto = Link_Oggetto
        rs!IDTipoOggetto = 131
        rs!IDFiliale = TheApp.Branch
        rs!IDNaturaScadenza = 6
        rs!IDAnagrafica_CF = IDAnagrafica
        rs!IDTipoAnagrafica_CF = 2
        rs!IDAzienda_CF = TheApp.IDFirm
        rs!IDAzienda = TheApp.IDFirm
        rs!IDPagamento = 52
        rs!IDRegistroIva = GET_LINK_REGISTRO_IVA(IDSezionale)
        rs!IDTipoStatoScadenza = 2
        rs!IDSezionale = IDSezionale
        rs!NumeroDocumento = NumeroDocumento
        rs!DataDocumento = DataDocumento
        rs!DataInizioScadenza = DataDocumento
        rs!ImportoComplessivo = ImportoComplessivoScadenza
        rs!ImportoIva = 0
        rs!ScadenzaAttivaPassiva = 0
        rs!GiornoScadenzaFissa = 0
        rs!IvaSuScadenza = 2
        rs!NumeroRataIVA = 0
        rs!NonRipartireIva = False
        rs!DataUltimaVariazione = Date
        rs!IDUtenteUltimaVariazione = TheApp.IDUser
        rs!VirtualDelete = 0
        
        IDTestataScadenza = fnNotNullN(rs!IDTestataScadenza)
    rs.Update

    rs.Close
    Set rs = Nothing
    If IDTestataScadenza > 0 Then
        GENERA_DETTAGLIO_SCADENZA IDTestataScadenza, ImportoComplessivoScadenza, 1, DataDocumento, Periodo
    End If
    
    
    GET_LINK_SCADENZA = Link_Oggetto
End If
End Function

Private Function GET_LINK_OGGETTO_SCADENZA(DataDocumento As String, IDSezionale As Long, NumeroDocumento As String) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

GET_LINK_OGGETTO_SCADENZA = 0

sSQL = "SELECT * FROM Oggetto "
sSQL = sSQL & "WHERE IDOggetto=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    rs!IDOggetto = fnGetNewKey("Oggetto", "IDOggetto")
    rs!IDTipoOggetto = 131
    rs!IDFunzione = 88
    rs!Oggetto = "Testata Scadenza del " & DataDocumento
    rs!IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.Branch)
    rs!IDAzienda = TheApp.IDFirm
    rs!IDSezionale = IDSezionale
    rs!DataEmissione = DataDocumento
    rs!numero = NumeroDocumento
    rs!DataUltimaVariazione = Date
    rs!IDUtenteUltimaVariazione = TheApp.IDUser
    rs!VirtualDelete = 0
    
rs.Update

GET_LINK_OGGETTO_SCADENZA = rs!IDOggetto

rs.Close
Set rs = Nothing
End Function

Private Function GENERA_DETTAGLIO_SCADENZA(IDTestataScadenza As Long, ImportoComplessivo As Double, NumeroDocumento As Long, DataDocumento As String, Periodo As String) As Double
Dim rs As ADODB.Recordset
Dim rsNew As ADODB.Recordset
Dim sSQL As String
Dim IRata As Integer
Dim RIMANENZA As Double
Dim IDTipoStatoScadenza As Long


sSQL = "SELECT * FROM DettaglioScadenza "
sSQL = sSQL & "WHERE IDTestataScadenza=" & IDTestataScadenza

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

rsNew.AddNew
    rsNew!IDTestataScadenza = IDTestataScadenza
    rsNew!IDDettaglioScadenza = fnGetNewKey("DettaglioScadenza", "IDDettaglioScadenza")
    rsNew!DataScadenza = DataDocumento
    rsNew!ImportoScadenza = ImportoComplessivo
    rsNew!IDTipoStatoScadenza = 2
    rsNew!IDTipoPosizioneScadenza = 3
    rsNew!IDTipoOggetto = 0
    rsNew!IDTipoPagamento = 1
    rsNew!RaggruppamentoScadenza = False
    rsNew!RIBA = False
    rsNew!TrasferitoOutlook = False
    rsNew!Contabilizzata = False
    rsNew!Note = Periodo
    rsNew!numerorata = NumeroDocumento
    rsNew!DataUltimaVariazione = Date
    rsNew!IDUtenteUltimaVariazione = TheApp.IDUser
    rsNew!VirtualDelete = 0
rsNew.Update

End Function



Private Function GET_LINK_REGISTRO_IVA(IDSezionale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRegistroIva FROM Sezionale "
sSQL = sSQL & "WHERE IDSezionale=" & IDSezionale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_REGISTRO_IVA = 1
Else
    GET_LINK_REGISTRO_IVA = fnNotNullN(rs!IDRegistroIva)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub ELIMINA_FLUSSO_RATE_DA_DOCUMENTO(IDContratto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDOggettoScadenza As Long

sSQL = "SELECT * FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), 0)
    
    If IDOggettoScadenza > 0 Then
        ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA 131, IDOggettoScadenza, fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto)
        ELIMINA_SCADENZA IDOggettoScadenza
    End If

rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_NUMERO_RATA(IDContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(NumeroRata) as Numero "
sSQL = sSQL & "FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_RATA = 1
Else
    GET_NUMERO_RATA = fnNotNullN(rs!numero) + 1
End If


rs.CloseResultset
Set rs = Nothing

End Function
Private Sub ELIMINAZIONE_APPUNTAMENTO_AGENDA(IDIntervento As Long)
Dim sSQL As String

sSQL = "DELETE FROM Appuntamento WHERE RV_POIDIntervento=" & IDIntervento

Cn.Execute sSQL

End Sub
Private Sub CREAZIONE_APPUNTAMENTO_AGENDA(DataAppuntamento As String, Orario As String, IDIntervento As Long, Oggetto As String, Luogo As String, Descrizione As String)
Dim sSQL As String
Dim ValoreGuid As String
Dim ValoreUtente As String
Dim DataFineApputamento As String

ValoreGuid = GET_GUID
ValoreUtente = GET_RISORSE
DataFineApputamento = DateAdd("h", 1, DataAppuntamento & " " & Orario)

sSQL = "INSERT INTO Appuntamento ("
sSQL = sSQL & "IDAppuntamento,Tipo, DataInizio, DataFine, TuttoIlGiorno, Oggetto, Luogo, Descrizione,"
sSQL = sSQL & "Stato, Etichetta, Risorse,  Privato, IDUtente, RV_POIDIntervento, "
sSQL = sSQL & "OptimisticLockField"
sSQL = sSQL & ") VALUES ("
sSQL = sSQL & "'" & ValoreGuid & "', "
sSQL = sSQL & 0 & ", "
sSQL = sSQL & fnNormDate(DataAppuntamento & " " & Orario) & ", "
sSQL = sSQL & fnNormDate(DataFineApputamento) & ", "
sSQL = sSQL & 0 & ", "
sSQL = sSQL & fnNormString(Oggetto) & ", "
sSQL = sSQL & fnNormString(Luogo) & ", "
sSQL = sSQL & fnNormString(Descrizione) & ", "
sSQL = sSQL & 2 & ", "
sSQL = sSQL & 0 & ", "
sSQL = sSQL & fnNormString(ValoreUtente) & ", "
sSQL = sSQL & 0 & ", "
sSQL = sSQL & TheApp.IDUser & ", "
sSQL = sSQL & IDIntervento & ", "
sSQL = sSQL & 0 & ")"

Cn.Execute sSQL


End Sub
Private Sub SCRIVI_APPUNTAMENTO(DataApputamento As String, Orario As String, IDIntervento As Long)
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Testo As String
Dim OutlookAppIntervento As Long
Dim AgendaAppIntervento As Long
Dim Oggetto As String
Dim Luogo As String
Dim Descrizione As String

If DataApputamento = "" Then Exit Sub
If Orario = "" Then Exit Sub


GET_DESCRIZIONI_APP_INTERVENTO IDIntervento, Oggetto, Luogo, Descrizione

ELIMINAZIONE_APPUNTAMENTO_AGENDA IDIntervento

CREAZIONE_APPUNTAMENTO_AGENDA DataApputamento, Orario, IDIntervento, Oggetto, Luogo, Descrizione

AgendaAppIntervento = 1

'AGGIORNAMENTO TABELLA INTERVENTO'''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "UPDATE RV_POIntervento SET "
    sSQL = sSQL & "AppuntamentoOutlook=" & 0 & ", "
    sSQL = sSQL & "AppuntamentoAgenda=" & AgendaAppIntervento
    sSQL = sSQL & "WHERE IDRV_POIntervento=" & IDIntervento
    Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub
Private Function GET_GUID() As String
Dim GetGUID As String
Dim GetGUID1 As String
Dim GetGUID2 As String
Dim GetGUID3 As String
Dim GetGUID4 As String
Dim GetGUID5 As String
Dim udtGUID As GUID
GetGUID = ""
If (CoCreateGuid(udtGUID) = 0) Then

    GetGUID = _
        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
        String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
        String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
        IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
        IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
        IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
        IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
        IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
        
       
        GetGUID1 = Mid(GetGUID, 1, 8)
        GetGUID2 = Mid(GetGUID, 9, 4)
        GetGUID3 = Mid(GetGUID, 13, 4)
        GetGUID4 = Mid(GetGUID, 17, 4)
        GetGUID5 = Mid(GetGUID, 21, 12)
        
        GetGUID = GetGUID1 & "-" & GetGUID2 & "-" & GetGUID3 & "-" & GetGUID4 & "-" & GetGUID5
    
End If

GET_GUID = GetGUID
End Function
Private Function GET_RISORSE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_RISORSE = "<ResourceIds>" & vbCrLf

sSQL = "SELECT * FROM Utente ORDER BY IDUtente "

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    GET_RISORSE = GET_RISORSE & "<ResourceId Type=" & Chr(34) & "System.Int32" & Chr(34) & " Value=" & Chr(34) & fnNotNullN(rs!IDUtente) & Chr(34) & " />" & vbCrLf
rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing

GET_RISORSE = GET_RISORSE & "</ResourceIds>"
End Function
Private Sub GET_PASSAGGGIO_APP_PREC(IDIntervento As Long, OutlookAppIntervento As Long, AgendaAppIntervento As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POIntervento, AppuntamentoOutlook, AppuntamentoAgenda "
sSQL = sSQL & "FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=" & IDIntervento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    OutlookAppIntervento = 0
    AgendaAppIntervento = 0
Else
    OutlookAppIntervento = fnNotNullN(rs!AppuntamentoOutlook)
    AgendaAppIntervento = fnNotNullN(rs!AppuntamentoAgenda)
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub GET_DESCRIZIONI_APP_INTERVENTO(IDIntervento As Long, Oggetto As String, Luogo As String, Descrizione As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POIEIntervento "
sSQL = sSQL & "WHERE IDRV_POIntervento=" & IDIntervento

Set rs = Cn.OpenResultset(sSQL)


If Not rs.EOF Then
    'LUOGO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Luogo = ""
    If fnNotNullN(rs!AppuntamentoPressoCliente) = 1 Then
        Luogo = GET_INDIRIZZO_CLIENTE(fnNotNullN(rs!IDAnagraficaCliente), fnNotNullN(rs!IDSitoPerAnagraficaIntervento))
        If fnNotNullN(rs!IDSitoPerAnagraficaIntervento) = 0 Then
            Luogo = Luogo & " (" & fnNotNull(rs!AnagraficaCliente) & " " & fnNotNull(rs!NomeCliente) & ")"
        Else
            Luogo = Luogo & " (" & fnNotNull(rs!SitoPerAnagraficaIntervento) & " - " & fnNotNull(rs!AnagraficaCliente) & " " & fnNotNull(rs!NomeCliente) & ")"
        End If
    Else
        Luogo = fnNotNull(rs!AnagraficaCliente) & " " & fnNotNull(rs!NomeCliente)
    End If
    
    Luogo = Mid(Luogo, 1, 50)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'OGGETTO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Oggetto = ""
    Oggetto = fnNotNull(rs!Articolo) & " (" & fnNotNull(rs!AnagraficaTecnicoOperativo) & " " & fnNotNull(rs!NomeTecnicoOperativo) & ")"
    Oggetto = Mid(Oggetto, 1, 100)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'DESCRIZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Descrizione = ""
    If Len(Trim(fnNotNull(rs!Richiesta))) > 0 Then
        Descrizione = Descrizione & "RIHIESTA" & vbCrLf
        Descrizione = Descrizione & fnNotNull(rs!Richiesta)
    End If
    If Len(Trim(fnNotNull(rs!Annotazioni))) > 0 Then
        Descrizione = Descrizione & vbCrLf
        Descrizione = Descrizione & "ANNOTAZIONI" & vbCrLf
        Descrizione = Descrizione & fnNotNull(rs!Annotazioni)
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Function GET_INDIRIZZO_CLIENTE(IDAnagrafica As Long, IDSitoPerAnagrafica As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If IDSitoPerAnagrafica = 0 Then 'Presso il cliente
    sSQL = "SELECT * FROM IERepAnagrafica "
    sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_INDIRIZZO_CLIENTE = ""
    Else
        GET_INDIRIZZO_CLIENTE = fnNotNull(rs!INDIRIZZO)
        If Len(Trim(rs!Comune)) > 0 Then
            GET_INDIRIZZO_CLIENTE = fnNotNull(rs!INDIRIZZO) & " - " & fnNotNull(rs!Comune)
        End If
        
    End If
    
    rs.CloseResultset
    Set rs = Nothing
Else 'Presso una sede del cliente

    sSQL = "SELECT dbo.SitoPerAnagrafica.IDAnagrafica, dbo.SitoPerAnagrafica.Cap, dbo.Comune.Comune, dbo.SitoPerAnagrafica.Email, dbo.SitoPerAnagrafica.Fax, "
    sSQL = sSQL & "dbo.SitoPerAnagrafica.Indirizzo, dbo.Nazione.Nazione, dbo.SitoPerAnagrafica.Referente, dbo.SitoPerAnagrafica.Telefono, dbo.TipoSito.TipoSito, "
    sSQL = sSQL & "dbo.SitoPerAnagrafica.SitoPerAnagrafica , dbo.Anagrafica.CodiceFiscale, dbo.Anagrafica.PartitaIVA, dbo.SitoPerAnagrafica.IDSitoPerAnagrafica "
    sSQL = sSQL & "FROM dbo.SitoPerAnagrafica LEFT OUTER JOIN "
    sSQL = sSQL & "dbo.Comune ON dbo.SitoPerAnagrafica.IDComune = dbo.Comune.IDComune LEFT OUTER JOIN "
    sSQL = sSQL & "dbo.Nazione ON dbo.SitoPerAnagrafica.IDNazione = dbo.Nazione.IDNazione INNER JOIN "
    sSQL = sSQL & "dbo.TipoSito ON dbo.SitoPerAnagrafica.IDTipoSito = dbo.TipoSito.IDTipoSito INNER JOIN "
    sSQL = sSQL & "dbo.Anagrafica ON dbo.SitoPerAnagrafica.IDAnagrafica = dbo.Anagrafica.IDAnagrafica "
    sSQL = sSQL & "WHERE dbo.SitoPerAnagrafica.IDSitoPerAnagrafica=" & IDSitoPerAnagrafica
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_INDIRIZZO_CLIENTE = ""
    Else
        GET_INDIRIZZO_CLIENTE = fnNotNull(rs!INDIRIZZO)
        If Len(Trim(rs!Comune)) > 0 Then
            GET_INDIRIZZO_CLIENTE = fnNotNull(rs!INDIRIZZO) & " - " & fnNotNull(rs!Comune)
        End If
        
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If




End Function


Private Function GET_TOTALE_ADEGUAMENTI_DETTAGLIO(IDContratto As Long) As Double
On Error GoTo ERR_GET_TOTALE_ADEGUAMENTI_DETTAGLIO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_TOTALE_ADEGUAMENTI_DETTAGLIO = 0


sSQL = "SELECT SUM(Importo) AS TotaleAdeguamenti "
sSQL = sSQL & "FROM RV_POIEAdeguamentiContratto "
'sSQL = sSQL & "WHERE IDRV_POContrattoPadre=" & frmMain.txtIDContrattoPadre.Value
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
'sSQL = sSQL & " AND IDRV_POTipoAdeguamento=2"


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_ADEGUAMENTI_DETTAGLIO = 0
Else
    GET_TOTALE_ADEGUAMENTI_DETTAGLIO = fnNotNullN(rs!TotaleAdeguamenti)
End If

rs.CloseResultset
Set rs = Nothing
GET_TOTALE_ADEGUAMENTI_DETTAGLIO = GET_TOTALE_ADEGUAMENTI_DETTAGLIO + Me.txtImportoAttuale.Value
Exit Function
ERR_GET_TOTALE_ADEGUAMENTI_DETTAGLIO:
    MsgBox Err.Description, vbCritical, "GET_TOTALE_ADEGUAMENTI_DETTAGLIO"
    
End Function
Private Function GET_LINK_ADEGUAMENTO_DA_RATA(IDRataContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POContrattoAdeguamento FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_PORateContratto=" & IDRataContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ADEGUAMENTO_DA_RATA = 0
Else
    GET_LINK_ADEGUAMENTO_DA_RATA = fnNotNullN(rs!IDRV_POContrattoAdeguamento)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_NUMERO_PRODOTTI_PER_SERVIZIO(IDRigaServizio As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT COUNT (IDRV_POContrattoServiziProdotti) as NumeroRecord "
sSQL = sSQL & "FROM RV_POContrattoServiziProdotti"
sSQL = sSQL & " WHERE IDRV_POContrattoServizi=" & IDRigaServizio
sSQL = sSQL & " AND Eliminato=" & fnNormBoolean(0)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_PRODOTTI_PER_SERVIZIO = 0
Else
    GET_NUMERO_PRODOTTI_PER_SERVIZIO = fnNotNullN(rs!NumeroRecord)
End If


rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_NUMERO_INTERVENTI_PER_SERVIZIO(IDRigaServizio As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT COUNT (IDRV_POIntervento) as NumeroRecord "
sSQL = sSQL & "FROM RV_POIntervento"
sSQL = sSQL & " WHERE IDRV_POContrattoServizi=" & IDRigaServizio

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_INTERVENTI_PER_SERVIZIO = 0
Else
    GET_NUMERO_INTERVENTI_PER_SERVIZIO = fnNotNullN(rs!NumeroRecord)
End If


rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_NUMERO_CONTRATTO(AnnoContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT COUNT(NumeroContratto) AS Numero "
sSQL = sSQL & "FROM RV_POContratto "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND AnnoContratto=" & AnnoContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_CONTRATTO = 1
Else
    GET_NUMERO_CONTRATTO = fnNotNullN(rs!numero) + 1
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub AGGIORNA_INTERVENTI_PRODOTTO(IDProdotto As Long, IDRigaProdottoContratto As Long)
Dim sSQL As String

sSQL = "UPDATE RV_POIntervento SET "
sSQL = sSQL & "IDRV_POContratto=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value) & ", "
sSQL = sSQL & "IDAnagraficaCliente=" & Me.CDCliente.KeyFieldID & ", "
sSQL = sSQL & "IDRV_POContrattoPadre=" & Me.txtIDContrattoPadre.Value & ", "
sSQL = sSQL & "IDAnagraficaFatturazione=" & IDClienteFatturazione & ", "
sSQL = sSQL & "IDRV_POContrattoProdotti=" & IDRigaProdottoContratto

sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDRV_POProdotto=" & IDProdotto
sSQL = sSQL & " AND GeneratoDa=1"
sSQL = sSQL & " AND DataAppuntamento>" & fnNormDate(Date)
'sSQL = sSQL & " AND InterventoChiuso=0"

Cn.Execute sSQL

End Sub
Private Sub AGGIORNA_INTERVENTI_PRODOTTO_PER_ELIMINAZIONE(IDProdotto As Long, IDRigaProdottoContratto As Long)
Dim sSQL As String
Dim Link_Cliente_Predefinito As Long

Link_Cliente_Predefinito = GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDClienteInterventoPredefinito")

If Link_Cliente_Predefinito = 0 Then Exit Sub

sSQL = "UPDATE RV_POIntervento SET "
sSQL = sSQL & "IDRV_POContratto=0" & ", "
sSQL = sSQL & "IDAnagraficaCliente=" & Link_Cliente_Predefinito & ", "
sSQL = sSQL & "IDRV_POContrattoPadre=0" & ", "
sSQL = sSQL & "IDAnagraficaFatturazione=" & Link_Cliente_Predefinito & ", "
sSQL = sSQL & "IDRV_POContrattoProdotti=Null"

sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDRV_POProdotto=" & IDProdotto
sSQL = sSQL & " AND GeneratoDa=1"
sSQL = sSQL & " AND DataAppuntamento>" & fnNormDate(Date)
'sSQL = sSQL & " AND InterventoChiuso=0"

Cn.Execute sSQL

End Sub
Private Function GET_CONTROLLO_ESISTENZA_RIGA_PRODOTTO(IDRigaProdottoContratto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT IDRV_POIntervento FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDRigaProdottoContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_ESISTENZA_RIGA_PRODOTTO = False
Else
    GET_CONTROLLO_ESISTENZA_RIGA_PRODOTTO = True
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CONTROLLO_ESISTENZA_PRODOTTO_SERVIZIO(IDRigaProdottoContratto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POContrattoServiziProdotti FROM RV_POContrattoServiziProdotti "
sSQL = sSQL & " WHERE IDRV_POContrattoProdotti=" & IDRigaProdottoContratto
sSQL = sSQL & " AND Eliminato=" & fnNormBoolean(0)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_ESISTENZA_PRODOTTO_SERVIZIO = False
Else
    GET_CONTROLLO_ESISTENZA_PRODOTTO_SERVIZIO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CONTROLLO_PRODOTTO_ALTRO_CONTRATTO(IDProdotto As Long, IDContratto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POContrattoProdotti FROM RV_POContrattoProdotti "
sSQL = sSQL & " WHERE IDRV_POProdotto=" & IDProdotto
sSQL = sSQL & " AND Dismesso=0"
sSQL = sSQL & " AND IDRV_POContratto<>" & IDContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_PRODOTTO_ALTRO_CONTRATTO = False
Else
    GET_CONTROLLO_PRODOTTO_ALTRO_CONTRATTO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ELIMINAZIONE_ASSOCIAZIONE_PRODOTTO_SERVIZIO(IDRigaProdottoContratto As Long)
Dim sSQL As String

sSQL = "DELETE FROM RV_POContrattoServiziProdotti "
sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDRigaProdottoContratto
Cn.Execute sSQL

End Sub
Private Function GET_CONTROLLO_ESISTENZA_RIGA_SERVIZIO(IDRigaServizioContratto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POIntervento FROM RV_POIntervento "
sSQL = sSQL & " WHERE IDRV_POContrattoServizi=" & IDRigaServizioContratto
sSQL = sSQL & " AND Manuale=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_ESISTENZA_RIGA_SERVIZIO = False
Else
    GET_CONTROLLO_ESISTENZA_RIGA_SERVIZIO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PARAMETRO_AZIENDA_LONG(IDFiliale As Long, nomeCampo As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & nomeCampo
sSQL = sSQL & " FROM RV_POParametriAzienda "
sSQL = sSQL & " WHERE IDFiliale=" & IDFiliale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PARAMETRO_AZIENDA_LONG = 0
Else
    GET_PARAMETRO_AZIENDA_LONG = fnNotNullN(rs.adoColumns(nomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_CONTRATTO_PRECEDENTE(IDContratto As Long, IDCotrattoPadre As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POContratto FROM RV_POContratto "
sSQL = sSQL & " WHERE IDRV_POContrattoPadre=" & IDCotrattoPadre
sSQL = sSQL & " AND IDRV_POContratto<>" & IDContratto
sSQL = sSQL & " ORDER BY NumeroRinnovo DESC"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_CONTRATTO_PRECEDENTE = 0
Else
    GET_LINK_CONTRATTO_PRECEDENTE = fnNotNullN(rs!IDRV_POContratto)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_ESISTENZA_INTERVENTI_CONTRATTO(IDContratto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POIntervento FROM RV_POIntervento "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Manuale=1 "

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_INTERVENTI_CONTRATTO = False
Else
    GET_ESISTENZA_INTERVENTI_CONTRATTO = True
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub AGGIORNA_CONTRATTO_PRECEDENTE(IDContrattoPrecedente As Long)
On Error GoTo ERR_AGGIORNA_CONTRATTO_PRECEDENTE
Dim sSQL As String

sSQL = "UPDATE RV_POContratto SET"
sSQL = sSQL & " ContrattoAttuale=1, "
sSQL = sSQL & " Chiuso=0"
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContrattoPrecedente
Cn.Execute sSQL

Exit Sub
ERR_AGGIORNA_CONTRATTO_PRECEDENTE:
    MsgBox Err.Description, vbCritical, "AGGIORNA_CONTRATTO_PRECEDENTE"
    
End Sub
Private Sub ELIMINA_INTERVENTO_CONTRATTO(IDContratto As Long)
On Error GoTo ERR_ELIMINA_INTERVENTO_CONTRATTO
Dim sSQL As String

sSQL = "DELETE FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
Cn.Execute sSQL

Exit Sub
ERR_ELIMINA_INTERVENTO_CONTRATTO:
    MsgBox Err.Description, vbCritical, "ELIMINA_INTERVENTO_CONTRATTO"
    
End Sub
Private Function GET_FUNZIONE_DA_IDOGGETTO(IDOggetto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto, IDTipoOggetto "
sSQL = sSQL & "FROM Oggetto  "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_FUNZIONE_DA_IDOGGETTO = 0
Else
    GET_FUNZIONE_DA_IDOGGETTO = GET_FUNZIONE(fnNotNullN(rs!IDTipoOggetto))
End If

rs.CloseResultset
Set rs = Nothing
End Function


Private Function GET_FUNZIONE(IDTipoOggetto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDFunzione "
sSQL = sSQL & "FROM Funzione  "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_FUNZIONE = 0
Else
    GET_FUNZIONE = fnNotNullN(rs!IDFunzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_TIPO_OGGETTO(NomeGestore) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT TipoOggetto.IDTipoOggetto, Gestore.Gestore "
sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
sSQL = sSQL & "WHERE (Gestore.Gestore = " & fnNormString(NomeGestore) & ")"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_OGGETTO = 0
Else
    GET_TIPO_OGGETTO = fnNotNullN(rs!IDTipoOggetto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_ACCORDO_COMMERCIALE(IDAnagrafica As Long, DataDocumento As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_ACCORDO_COMMERCIALE = 0

sSQL = "SELECT IDAccordiCommerciali, Descrizione FROM AccordiCommerciali "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDTipoAnagrafica=2"
sSQL = sSQL & " AND DataInizio<=" & fnNormDate(DataDocumento)
sSQL = sSQL & " AND DataFine>=" & fnNormDate(DataDocumento)
sSQL = sSQL & " AND Chiuso=" & fnNormBoolean(0)

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_ACCORDO_COMMERCIALE = fnNotNullN(rs!IDAccordiCommerciali)
    DescrAccordoCommerciale = fnNotNull(rs!Descrizione)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub GET_TOTALE_RIGA()
Dim Imponibile As Double
Dim importoTotale As Double
Dim ImportoIva As Double

Imponibile = Me.txtImpUniProd.Value
Imponibile = Imponibile - ((Imponibile / 100) * Me.txtSconto1Prod.Value)
Imponibile = Imponibile - ((Imponibile / 100) * Me.txtSconto2Prod.Value)
Imponibile = Imponibile * Me.txtQtaArtProd.Value

If Me.chkACorpo.Value = vbUnchecked Then
    Imponibile = Imponibile * Me.txtQuantitaEffettiva.Value
End If

Imponibile = Imponibile - Me.txtScontoImpProd.Value
Imponibile = fnRoundChange(Imponibile, 0.01, 3)

importoTotale = Imponibile * ((Me.txtAliquotaIvaProd.Value / 100) + 1)

importoTotale = fnRoundChange(importoTotale, 0.01, 3)

ImportoIva = importoTotale - Imponibile

Me.txtImponibileProd.Value = Imponibile
Me.txtImportoIvaProd.Value = ImportoIva
Me.txtTotaleRigaProd.Value = importoTotale

End Sub
Private Sub txtScontoImpProd_LostFocus()
    GET_TOTALI_RIGA_DETTAGLIO
End Sub
Private Function GET_TOTALE_PRODOTTI(IDContratto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Me.txtImponibileProdTot.Value = 0
Me.txtImportoIvaProdTot.Value = 0
Me.txtTotaleRigaProdTot.Value = 0

sSQL = "SELECT SUM(TotaleRiga) AS TotaleRiga, "
sSQL = sSQL & "SUM (Imponibile) AS TotaleImponibile, "
sSQL = sSQL & "SUM (ImportoIva) AS TotaleIva "
sSQL = sSQL & "FROM RV_POContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtImponibileProdTot.Value = fnNotNullN(rs!TotaleImponibile)
    Me.txtImportoIvaProdTot.Value = fnNotNullN(rs!TotaleIva)
    Me.txtTotaleRigaProdTot.Value = fnNotNullN(rs!TotaleRiga)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_TOTALE_CONTRATTO(IDContratto As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT SUM(ImportoComplessivo) AS TotaleRiga "
sSQL = sSQL & "FROM RV_POContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_CONTRATTO = 0
Else
    GET_TOTALE_CONTRATTO = fnNotNullN(rs!TotaleRiga)
End If

rs.CloseResultset
Set rs = Nothing

End Function



Private Function GET_LINK_OGGETTO_CONTRATTO(IDOggetto As Long, IDTipoOggetto As Long, numero As String, DataDecorrenza As String) As Long
On Error GoTo ERR_GET_LINK_OGGETTO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim IDOggettoLocal As Long
Dim IDFunzione As Long

GET_LINK_OGGETTO_CONTRATTO = 0

IDFunzione = GET_LINK_FUNZIONE(IDTipoOggetto)

If IDFunzione = 0 Then Exit Function

sSQL = "SELECT * FROM Oggetto "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF = True Then
    rs.AddNew
        rs!IDTipoOggetto = IDTipoOggetto
        rs!IDFunzione = IDFunzione
        rs!IDAzienda = TheApp.IDFirm
        rs!IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.Branch)
        rs!IDSezionale = 0
        rs!Oggetto = GET_DESCRIZIONE_FUNZIONE(IDFunzione)
        rs!DataEmissione = DataDecorrenza
        rs!numero = numero
        rs!DataUltimaVariazione = Date
        rs!IDUtenteUltimaVariazione = TheApp.IDUser
        rs!VirtualDelete = 0
        rs!IDOggetto = fnGetNewKey("Oggetto", "IDOggetto")
        GET_LINK_OGGETTO_CONTRATTO = rs!IDOggetto
    rs.Update
Else
    rs!DataEmissione = DataDecorrenza
    rs!numero = numero
    rs!DataUltimaVariazione = Date
    rs!IDUtenteUltimaVariazione = TheApp.IDUser
    rs!VirtualDelete = 0
    GET_LINK_OGGETTO_CONTRATTO = rs!IDOggetto
    rs.Update
End If

rs.Close
Set rs = Nothing

Exit Function

ERR_GET_LINK_OGGETTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_OGGETTO"
    GET_LINK_OGGETTO_CONTRATTO = 0
End Function


Private Sub CREA_FLUSSO_DOCUMENTALE_CONTRATTO(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long, Descrizione As String)
On Error GoTo ERR_CREA_FLUSSO_DOCUMENTALE_SCADENZA
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim IDFunzioneVend As Long
Dim IDFunzioneRata As Long
Dim IDFlussoGruppo As Long
Dim IDFlussoFunzione As Long

IDFunzioneVend = GET_LINK_FUNZIONE(IDTipoOggettoVend)
IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)

If IDFunzioneVend = 0 Then Exit Sub
If IDFunzioneRata = 0 Then Exit Sub
'''''''''''''''''''''''''''''''''GRUPPO FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoGruppo "
sSQL = sSQL & "WHERE Descrizione=" & fnNormString(Descrizione)
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFlussoGruppo")
        rsNew!Descrizione = Descrizione
    rsNew.Update
End If

IDFlussoGruppo = fnNotNullN(rsNew!IDFlussoGruppo)

rsNew.Close
Set rsNew = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoFunzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzioneRata
sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneVend
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoFunzione = fnGetNewKeyTipoOggetto("FlussoFunzione", "IDFlussoFunzione")
        rsNew!IDFunzione = IDFunzioneRata
        rsNew!IDFunzioneSuccessiva = IDFunzioneVend
        rsNew!Cardinalita = 3
        rsNew!TipoAutomatismo = 1
        rsNew!Attributo = 14
        rsNew!TipoDipendenza = 1
        rsNew!IDFlussoGruppo = IDFlussoGruppo
    rsNew.Update
End If

IDFlussoFunzione = fnNotNullN(rsNew!IDFlussoFunzione)

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE COLLEGATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoFunzioneCollegato "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoFunzione = IDFlussoFunzione
        rsNew!IDOggetto = IDOggettoRata
        rsNew!IDTipoOggetto = IDTipoOggettoRata
End If

rsNew!FlussoFunzioneCollegato = 2
rsNew.Update

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FLUSSO OGGETTI COLLEGATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoVend

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
    rsNew!IDFlussoFunzione = IDFlussoFunzione
    rsNew!IDOggetto = IDOggettoRata
    rsNew!IDTipoOggetto = IDTipoOggettoRata
    rsNew!IDTipoOggettoCollegato = IDTipoOggettoVend
    rsNew!IDOggettoCollegato = IDOggettoVend
    rsNew.Update
End If

rsNew.Close
Set rsNew = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Sub
ERR_CREA_FLUSSO_DOCUMENTALE_SCADENZA:
MsgBox Err.Description, vbCritical, "CREA_FLUSSO_DOCUMENTALE_SCADENZA"
End Sub
Private Sub ELIMINA_FLUSSO_DOCUMENTALE_CONTRATTO(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long, Descrizione As String)
On Error GoTo ERR_ELIMINA_FLUSSO_DOCUMENTALE
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim IDFunzioneVend As Long
Dim IDFunzioneRata As Long
Dim IDFlussoGruppo As Long
Dim IDFlussoFunzione As Long

IDFunzioneVend = GET_LINK_FUNZIONE(IDTipoOggettoVend)
IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)

'''''''''''''''''''''''''''''''''GRUPPO FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoGruppo "
sSQL = sSQL & "WHERE Descrizione=" & fnNormString(Descrizione)
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFLussoGruppo")
        rsNew!Descrizione = Descrizione
    rsNew.Update
End If

IDFlussoGruppo = fnNotNullN(rsNew!IDFlussoGruppo)

rsNew.Close
Set rsNew = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IDFunzioneVend > 0 Then
    sSQL = "SELECT * FROM FlussoFunzione "
    sSQL = sSQL & "WHERE IDFunzione=" & IDFunzioneRata
    sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneVend
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDFlussoFunzione = fnGetNewKeyTipoOggetto("FlussoFunzione", "IDFlussoFunzione")
            rsNew!IDFunzione = IDFunzioneVend
            rsNew!IDFunzioneSuccessiva = IDFunzioneRata
            rsNew!Cardinalita = 3
            rsNew!TipoAutomatismo = 1
            rsNew!Attributo = 14
            rsNew!TipoDipendenza = 1
            rsNew!IDFlussoGruppo = IDFlussoGruppo
        rsNew.Update
    End If
    
    IDFlussoFunzione = fnNotNullN(rsNew!IDFlussoFunzione)
    
    rsNew.Close
    Set rsNew = Nothing
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE COLLEGATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggettoCollegato<>" & IDOggettoVend

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    sSQL = "DELETE FROM FlussoFunzioneCollegato "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoRata
    sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
    sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
    Cn.Execute sSQL
End If
rsNew.Close
Set rsNew = Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''FLUSSO OGGETTI COLLEGATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoVend
Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Sub
ERR_ELIMINA_FLUSSO_DOCUMENTALE:
    MsgBox Err.Description, vbCritical, "ELIMINA_FLUSSO_DOCUMENTALE"
    
End Sub

Private Sub CREA_SCADENZA_CONTRATTO()
Dim IDOggettoScadenza As Long
Dim AnnoContratto As String

    If LINK_SEZIONALE_RATE > 0 Then
        AnnoContratto = txtAnnoContratto.Text & "-" & txtNumeroContratto.Text
        
        IDOggettoScadenza = GET_LINK_OGGETTO_SCADENZA_COLLEGATA(fnNotNullN(m_Document("IDOggetto").Value), m_DocType.ID, 0)
        
        If IDOggettoScadenza > 0 Then
            ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA_C 131, IDOggettoScadenza, fnNotNullN(m_Document("IDOggetto").Value), m_DocType.ID
            ELIMINA_SCADENZA IDOggettoScadenza
        End If
        
        If Me.txtImportoAttuale.Value > 0 Then
            If Me.txtImportoTotAdeg.Value <> 0 Then
                IDOggettoScadenza = GET_LINK_SCADENZA(Me.txtImportoTotAdeg.Value, IDClienteFatturazione, AnnoContratto, Me.txtDataDecorrenza.Text, LINK_SEZIONALE_RATE, "")
                CREA_FLUSSO_DOCUMENTALE_CONTRATTO 131, IDOggettoScadenza, fnNotNullN(m_Document("IDOggetto").Value), m_DocType.ID, "Contratto -> Scadenza"
            End If
        End If
    End If
    
End Sub
Private Function GET_CONTROLLO_RATE_PAGATE(IDContratto As Long) As Boolean
On Error GoTo ERR_GET_CONTROLLO_RATE_PAGATE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim numero As Long

sSQL = "SELECT COUNT(IDRV_PORateContratto) AS NRatePag"
sSQL = sSQL & " FROM RV_PORateContratto "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Adeguamento=0"
sSQL = sSQL & " AND Manuale=0"
sSQL = sSQL & " AND ContrattoAttuale=1"
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"
sSQL = sSQL & " AND Fatturata=1 "


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    numero = 0
Else
    numero = fnNotNullN(rs!NRatePag)
End If

rs.CloseResultset
Set rs = Nothing

If numero = 0 Then
    GET_CONTROLLO_RATE_PAGATE = True
Else
    GET_CONTROLLO_RATE_PAGATE = False
End If

Exit Function

ERR_GET_CONTROLLO_RATE_PAGATE:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_RATE_PAGATE"
    GET_CONTROLLO_RATE_PAGATE = False
End Function
Private Function GET_CONTROLLO_IMPORTO_RATE(IDContratto As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim numero As Double

GET_CONTROLLO_IMPORTO_RATE = 0

sSQL = "SELECT SUM(ImportoRata) AS TotaleRata"
sSQL = sSQL & " FROM RV_PORateContratto "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND ((Adeguamento=0) OR (Adeguamento IS NULL))"
sSQL = sSQL & " AND Manuale=0"
sSQL = sSQL & " AND ContrattoAttuale=1"
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_IMPORTO_RATE = 0
Else
    GET_CONTROLLO_IMPORTO_RATE = FormatNumber(fnNotNullN(rs!TotaleRata), 2)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub DISEGNA_CONTROLLI(Abilita As Boolean)

Me.cboTipoRateizzazione.Enabled = Abilita
Me.cboDurataContratto.Enabled = Abilita
Me.cboTipoRinnovo.Enabled = Abilita
Me.cboDurataAssistenza.Enabled = Abilita
Me.txtDataScadenza.Enabled = Abilita
Me.txtDataScadenzaPerRinnovo.Enabled = Abilita
Me.txtDataFineAssistenza.Enabled = Abilita
Me.chkGeneraRateProd.Enabled = Abilita

End Sub
Private Function GET_PRODOTTO(IDProdotto As Long) As String
On Error GoTo ERR_GET_PRODOTTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POIEProdotto "
sSQL = sSQL & "WHERE IDRV_POProdotto=" & IDProdotto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PRODOTTO = ""
Else
    GET_PRODOTTO = fnNotNull(rs!Descrizione)
    If Len(Trim(fnNotNull(rs!Matricola))) > 0 Then
        GET_PRODOTTO = GET_PRODOTTO & " (" & fnNotNull(rs!Matricola) & ")"
    End If
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_PRODOTTO:
    MsgBox Err.Description, vbCritical, "GET_PRODOTTO"
End Function
Private Function GET_PRODOTTO_CONTRATTO(IDProdottoContratto As Long) As String
On Error GoTo ERR_GET_PRODOTTO_CONTRATTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POIEContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDProdottoContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    
    GET_PRODOTTO_CONTRATTO = ""

Else
    
    GET_PRODOTTO_CONTRATTO = fnNotNull(rs!Descrizione)
    
    If Len(Trim(fnNotNull(rs!ValoreIndentificativo))) > 0 Then
        GET_PRODOTTO_CONTRATTO = GET_PRODOTTO_CONTRATTO & " (" & fnNotNull(rs!ValoreIndentificativo) & ")"
    End If
    If fnNotNullN(rs!IDArticolo) > 0 Then
        GET_PRODOTTO_CONTRATTO = GET_PRODOTTO_CONTRATTO & " - " & fnNotNull(rs!CodiceArticolo) & ""
    End If
    
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_PRODOTTO_CONTRATTO:
    MsgBox Err.Description, vbCritical, "GET_PRODOTTO_CONTRATTO"
End Function
Private Function GET_TOTALI_ACCONTI_CONTRATTO(IDOggettoContratto As Long, IDTipoOggettoContratto As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(Tot_documento_corr) AS TotaliAcconti "
sSQL = sSQL & "FROM RV_POIEContrattoAcconti "
sSQL = sSQL & "WHERE IDTipoOggettoCollegato=" & IDTipoOggettoContratto
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALI_ACCONTI_CONTRATTO = 0
Else
    GET_TOTALI_ACCONTI_CONTRATTO = fnNotNullN(rs!TotaliAcconti)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub GET_GRIGLIA_ACCONTI()
'Dim sSQL As String
'Dim OLDCursor As Long
'Dim cl As dgColumnHeader
'
'
''    sSQL = "SELECT * FROM RV_POIEContrattoAcconti "
''    sSQL = sSQL & "WHERE IDOggettoCollegato=" & fnNotNullN(m_Document("IDOggetto").Value)
''    sSQL = sSQL & " AND IDTipoOggettoCollegato=" & m_DocType.ID
''    sSQL = sSQL & " ORDER BY DataEmissione, Numero"
'
'    OLDCursor = Cn.CursorLocation
'    Cn.CursorLocation = 3
'
''    If Not (rsGrigliaAcc Is Nothing) Then
''        If rsGrigliaAcc.State > 0 Then
''            rsGrigliaAcc.Close
''        End If
''        Set rsGrigliaAcc = Nothing
''    End If
'
''    Set rsGrigliaAcc = New ADODB.Recordset
''    rsGrigliaAcc.CursorLocation = adUseClient
''    rsGrigliaAcc.Open sSQL, Cn.InternalConnection
'
'    With Me.GrigliaAcconti
'        'Set .PaintNotifyObj = gPaintNotify
'        .EnableMove = True
'        .UpdatePosition = True
'        .BooleanType = dgGraphic
'        .SelectionMode = dgSelectRow
'        .ColumnsHeader.Clear
'
'        .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgInteger, False, 500, dgAlignleft
'
'        .ColumnsHeader.Add "IDFlussoFunzione", "IDFlussoFunzione", dgInteger, False, 500, dgAlignleft
'        .ColumnsHeader.Add "IDFlussoGruppo", "IDFlussoGruppo", dgInteger, False, 500, dgAlignleft
'        .ColumnsHeader.Add "DescrizioneFlusso", "Flusso", dgchar, True, 3500, dgAlignleft
'
'        .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignleft
'        .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
'
'        .ColumnsHeader.Add "IDTipoOggettoCollegato", "IDTipoOggettoCollegato", dgInteger, False, 500, dgAlignleft
'        .ColumnsHeader.Add "IDOggettoCollegato", "IDOggettoCollegato", dgInteger, False, 500, dgAlignleft
'
'        .ColumnsHeader.Add "Oggetto", "Tipo documento", dgchar, True, 3500, dgAlignleft
'        .ColumnsHeader.Add "DataEmissione", "Data doc.", dgDate, False, 2500, dgAlignleft
'        .ColumnsHeader.Add "Numero", "Numero doc.", dgchar, True, 2500, dgAlignRight
'
'        Set cl = .ColumnsHeader.Add("NettoAPagare", "Da pagare", dgDouble, True, 2000, dgAlignRight)
'            cl.FormatOptions.FormatNumericRegionalSettings = False
'            cl.FormatOptions.UseFormatControlSettings = False
'            cl.FormatOptions.FormatNumericDecSep = ","
'            cl.FormatOptions.FormatNumericDecimals = 2
'            cl.FormatOptions.FormatNumericThousandSep = "."
'
'        Set cl = .ColumnsHeader.Add("TotaleDocumento", "Totale", dgDouble, False, 2000, dgAlignRight)
'            cl.FormatOptions.FormatNumericRegionalSettings = False
'            cl.FormatOptions.UseFormatControlSettings = False
'            cl.FormatOptions.FormatNumericDecSep = ","
'            cl.FormatOptions.FormatNumericDecimals = 2
'            cl.FormatOptions.FormatNumericThousandSep = "."
'
'        Set cl = .ColumnsHeader.Add("TotaleNetto", "Totale imp.", dgDouble, False, 2000, dgAlignRight)
'            cl.FormatOptions.FormatNumericRegionalSettings = False
'            cl.FormatOptions.UseFormatControlSettings = False
'            cl.FormatOptions.FormatNumericDecSep = ","
'            cl.FormatOptions.FormatNumericDecimals = 2
'            cl.FormatOptions.FormatNumericThousandSep = "."
'
'        Set cl = .ColumnsHeader.Add("TotaleIva", "Totale I.V.A.", dgDouble, False, 2000, dgAlignRight)
'            cl.FormatOptions.FormatNumericRegionalSettings = False
'            cl.FormatOptions.UseFormatControlSettings = False
'            cl.FormatOptions.FormatNumericDecSep = ","
'            cl.FormatOptions.FormatNumericDecimals = 2
'            cl.FormatOptions.FormatNumericThousandSep = "."
'
'
'
'        Set .Recordset = rsGrigliaAcc
'        .LoadUserSettings
'        .Refresh
'
'    End With
'    Cn.CursorLocation = OLDCursor
'
'Exit Sub
'
'ERR_fnGrigliaAssegnazione:
'    MsgBox Err.Description, vbCritical, "Reperimento dati acconti"

End Sub
Private Sub ELIMINA_FLUSSO_DOCUMENTALE_ACCONTO(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long, Descrizione As String)
On Error GoTo ERR_ELIMINA_FLUSSO_DOCUMENTALE
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim IDFunzioneVend As Long
Dim IDFunzioneRata As Long
Dim IDFlussoGruppo As Long
Dim IDFlussoFunzione As Long

IDFunzioneVend = GET_LINK_FUNZIONE(IDTipoOggettoVend)
IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)

'''''''''''''''''''''''''''''''''GRUPPO FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoGruppo "
sSQL = sSQL & "WHERE Descrizione=" & fnNormString(Descrizione)
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFLussoGruppo")
        rsNew!Descrizione = Descrizione
    rsNew.Update
End If

IDFlussoGruppo = fnNotNullN(rsNew!IDFlussoGruppo)

rsNew.Close
Set rsNew = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IDFunzioneVend > 0 Then
    sSQL = "SELECT * FROM FlussoFunzione "
    sSQL = sSQL & "WHERE IDFunzione=" & IDFunzioneVend
    sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneRata
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDFlussoFunzione = fnGetNewKeyTipoOggetto("FlussoFunzione", "IDFlussoFunzione")
            rsNew!IDFunzione = IDFunzioneVend
            rsNew!IDFunzioneSuccessiva = IDFunzioneRata
            rsNew!Cardinalita = 3
            rsNew!TipoAutomatismo = 1
            rsNew!Attributo = 14
            rsNew!TipoDipendenza = 1
            rsNew!IDFlussoGruppo = IDFlussoGruppo
        rsNew.Update
    End If
    
    IDFlussoFunzione = fnNotNullN(rsNew!IDFlussoFunzione)
    
    rsNew.Close
    Set rsNew = Nothing
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE COLLEGATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggetto=" & IDOggettoVend
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggettoCollegato<>" & IDOggettoRata

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    sSQL = "DELETE FROM FlussoFunzioneCollegato "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoVend
    sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoVend
    sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
    Cn.Execute sSQL
End If
rsNew.Close
Set rsNew = Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''FLUSSO OGGETTI COLLEGATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggetto=" & IDOggettoVend
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoRata
Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Sub
ERR_ELIMINA_FLUSSO_DOCUMENTALE:
    MsgBox Err.Description, vbCritical, "ELIMINA_FLUSSO_DOCUMENTALE"
    
End Sub

Private Function GET_SALDO_CONTRATTO(IDOggettoContratto As Long) As Double
On Error GoTo ERR_GET_SALDO_CONTRATTO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Long
Dim TotaleDocumenti As Double


If Not (rsGrigliaAcc Is Nothing) Then
    If rsGrigliaAcc.State > 0 Then
        rsGrigliaAcc.Close
    End If

    Set rsGrigliaAcc = Nothing
End If

Set rsGrigliaAcc = New ADODB.Recordset
rsGrigliaAcc.CursorLocation = adUseClient

rsGrigliaAcc.Fields.Append "IDFlussoFunzione", adInteger, , adFldIsNullable
rsGrigliaAcc.Fields.Append "IDFlussoGruppo", adInteger, , adFldIsNullable
rsGrigliaAcc.Fields.Append "DescrizioneFlusso", adVarChar, 250, adFldIsNullable
rsGrigliaAcc.Fields.Append "IDTipoOggetto", adInteger, , adFldIsNullable
rsGrigliaAcc.Fields.Append "IDOggetto", adInteger, , adFldIsNullable
rsGrigliaAcc.Fields.Append "IDTipoOggettoCollegato", adInteger, , adFldIsNullable
rsGrigliaAcc.Fields.Append "IDOggettoCollegato", adInteger, , adFldIsNullable
rsGrigliaAcc.Fields.Append "Oggetto", adVarChar, 250, adFldIsNullable
rsGrigliaAcc.Fields.Append "DataEmissione", adDBDate, , adFldIsNullable
rsGrigliaAcc.Fields.Append "Numero", adVarChar, 250, adFldIsNullable
rsGrigliaAcc.Fields.Append "IDAzienda", adInteger, , adFldIsNullable
rsGrigliaAcc.Fields.Append "Entita", adVarChar, 250, adFldIsNullable
rsGrigliaAcc.Fields.Append "TotaleDocumento", adDouble, , adFldIsNullable
rsGrigliaAcc.Fields.Append "TotaleNetto", adDouble, , adFldIsNullable
rsGrigliaAcc.Fields.Append "TotaleIva", adDouble, , adFldIsNullable
rsGrigliaAcc.Fields.Append "NettoAPagare", adDouble, , adFldIsNullable

rsGrigliaAcc.Open , , adOpenKeyset, adLockBatchOptimistic


TotaleDocumenti = 0

sSQL = "SELECT * FROM RV_POIEContrattoSaldo "
sSQL = sSQL & "WHERE IDTipoOggettoCollegato=" & m_DocType.ID
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoContratto

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

While Not rs.EOF
    rsGrigliaAcc.AddNew
        rsGrigliaAcc!IDFlussoFunzione = rs!IDFlussoFunzione
        rsGrigliaAcc!IDFlussoGruppo = rs!IDFlussoGruppo
        rsGrigliaAcc!DescrizioneFlusso = rs!DescrizioneFlusso
        rsGrigliaAcc!IDTipoOggetto = rs!IDTipoOggetto
        rsGrigliaAcc!IDOggetto = rs!IDOggetto
        rsGrigliaAcc!IDTipoOggettoCollegato = rs!IDTipoOggettoCollegato
        rsGrigliaAcc!IDOggettoCollegato = rs!IDOggettoCollegato
        rsGrigliaAcc!Oggetto = rs!Oggetto
        rsGrigliaAcc!DataEmissione = rs!DataEmissione
        rsGrigliaAcc!numero = rs!numero
        rsGrigliaAcc!IDAzienda = rs!IDAzienda
        rsGrigliaAcc!Entita = rs!Entita
        GET_TOTALE_DOC_COLLEGATO rsGrigliaAcc!IDOggetto, rsGrigliaAcc!Entita, rsGrigliaAcc
        
        TotaleDocumenti = TotaleDocumenti + rsGrigliaAcc!NettoAPagare
        
    rsGrigliaAcc.Update
rs.MoveNext
Wend


rs.Close
Set rs = Nothing


GET_SALDO_CONTRATTO = TotaleDocumenti

'GET_GRIGLIA_ACCONTI

Exit Function
ERR_GET_SALDO_CONTRATTO:
    MsgBox Err.Description, vbCritical, "GET_SALDO_CONTRATTO"

End Function
Private Sub GET_TOTALE_DOC_COLLEGATO(IDOggetto As Long, Entita As String, rstmp As ADODB.Recordset)
On Error GoTo ERR_GET_TOTALE_DOC_COLLEGATO
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String


rstmp!TotaleDocumento = 0
rstmp!TotaleNetto = 0
rstmp!TotaleIva = 0
rstmp!NettoAPagare = 0

sSQL = "SELECT IDOggetto, Tot_documento_corr, Tot_imponibile_corr, Tot_imposta_corr, Tot_netto_a_pagare_corr "
sSQL = sSQL & " FROM " & Entita
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    rstmp!TotaleDocumento = fnNotNullN(rs!Tot_documento_corr)
    rstmp!TotaleNetto = fnNotNullN(rs!Tot_imponibile_corr)
    rstmp!TotaleIva = fnNotNullN(rs!Tot_imposta_corr)
    rstmp!NettoAPagare = fnNotNullN(rs!Tot_netto_a_pagare_corr)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_GET_TOTALE_DOC_COLLEGATO:

End Sub
Private Function GET_RILEVAMENTO_PAGATO(IDContratto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Count(IDRV_POContatoreRilevamenti) AS NumeroRate "
sSQL = sSQL & "FROM RV_POContatoreRilevamenti "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Fatturata=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_RILEVAMENTO_PAGATO = False
Else
    If fnNotNullN(rs!numerorate) > 0 Then
        GET_RILEVAMENTO_PAGATO = True
    Else
        GET_RILEVAMENTO_PAGATO = False
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub GET_LISTA_PRODOTTI_CONTRATTI(rstmp As ADODB.Recordset, IDContratto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


Set rstmp = New ADODB.Recordset
rstmp.CursorLocation = adUseClient

rstmp.Fields.Append "IDRV_POContrattoProdotti", adInteger, , adFldIsNullable

rstmp.Open , , adOpenKeyset, adLockBatchOptimistic

sSQL = "SELECT IDRV_POContrattoProdotti FROM RV_POContrattoProdotti "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    rstmp.AddNew
        rstmp!IDRV_POContrattoProdotti = rs!IDRV_POContrattoProdotti
    rstmp.Update
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing



End Sub
Private Sub ELIMINA_CONFIGURAZIONE_CONTATORI(rstmp As ADODB.Recordset, IDContratto As Long)
Dim sSQL As String

If ((rstmp.EOF) And (rstmp.BOF)) Then Exit Sub

rstmp.MoveFirst

While Not rstmp.EOF
    sSQL = "DELETE FROM RV_POContrattoServiziProdotti "
    sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & fnNotNullN(rstmp!IDRV_POContrattoProdotti)
    Cn.Execute sSQL
    
    
    sSQL = "DELETE FROM RV_POContrattoProdottiContatori "
    sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & fnNotNullN(rstmp!IDRV_POContrattoProdotti)
    Cn.Execute sSQL
rstmp.MoveNext
Wend

rstmp.Close
Set rstmp = Nothing

'ELIMINAZIONE RILEVAMENTI
sSQL = "DELETE FROM RV_POContatoreRilevamenti "
sSQL = sSQL & "WHERE IDRV_POContratto=" & IDContratto
Cn.Execute sSQL

End Sub
Private Function GET_CONTROLLO_ESISTENZA_ACCONTI(IDOggettoContratto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POIEContrattoSaldo "
sSQL = sSQL & "WHERE IDTipoOggettoCollegato=" & m_DocType.ID
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoContratto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_ESISTENZA_ACCONTI = False
Else
    GET_CONTROLLO_ESISTENZA_ACCONTI = True
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Sub GENERA_INTERVENTO_DA_PRODOTTO(IDContrattoProdotto As Long, IDArticoloConsegna As Long, IDArticoloRiconsegna As Long)




End Sub
Private Sub CHANGE_ANAGRAFICA_FATTURAZIONE(IDAnagrafica As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) < 0 Then
    
    IDClienteFatturazione = IDAnagrafica
    ClienteFatturazione = Me.CDCliente.Code & " " & Me.CDCliente.Description
    
    IDAccordoCommerciale = GET_LINK_ACCORDO_COMMERCIALE(IDAnagrafica, txtDataStipula.Text)
    
    
    
    
    sSQL = "SELECT IDPagamentoDefault, CalcolaRitenutaAcconto "
    sSQL = sSQL & "FROM Cliente "
    sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        Me.cboPagamentoRate.WriteOn fnNotNullN(rs!IDPagamentoDefault)
        Me.chkRitAcconto.Value = Abs(fnNotNullN(rs!CalcolaRitenutaAcconto))
    End If

    rs.CloseResultset
    Set rs = Nothing
    
    
    
    
End If

LINK_LISTINO_CLIENTE = GET_LINK_LISTINO(IDAnagrafica, TheApp.IDFirm)

End Sub

Private Sub CREA_RAPPR_LEG_AZIENDA(IDAzienda As Long, Descrizione As String)
Dim sSQL As String
Dim rs As ADODB.Recordset


If (Len(Trim(fnNotNull(Descrizione))) = 0) Then Exit Sub

sSQL = "SELECT * FROM RV_PORappresentantiLegaliAna "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND IDAnagrafica IS NULL "
sSQL = sSQL & " AND Descrizione=" & fnNormString(Descrizione)

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
        rs!IDAzienda = IDAzienda
        rs!Descrizione = Descrizione
    rs.Update
End If

rs.Close
Set rs = Nothing

End Sub
Private Sub CREA_RAPPR_LEG_CLIENTE(IDAzienda As Long, Descrizione As String, IDAnagrafica As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset

If (Len(Trim(fnNotNull(Descrizione))) = 0) Then Exit Sub

sSQL = "SELECT * FROM RV_PORappresentantiLegaliAna "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND Descrizione=" & fnNormString(Descrizione)

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
        rs!IDAzienda = IDAzienda
        rs!IDAnagrafica = IDAnagrafica
        rs!Descrizione = Descrizione
    rs.Update
End If

rs.Close
Set rs = Nothing

End Sub
Private Sub CREA_RUOLO(IDAzienda As Long, Descrizione As String)
Dim sSQL As String
Dim rs As ADODB.Recordset

If (Len(Trim(fnNotNull(Descrizione))) = 0) Then Exit Sub



sSQL = "SELECT * FROM RV_PORappresentantiLegaliRuolo "
sSQL = sSQL & " WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND Descrizione=" & fnNormString(Descrizione)

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
        rs!IDAzienda = IDAzienda
        rs!Descrizione = Descrizione
    rs.Update
End If

rs.Close
Set rs = Nothing


End Sub

Private Function GET_CALCOLO_QUANTITA_EFFETTIVA(DataInizioPeriodo As String, DataFinePeriodo As String) As Long
On Error GoTo ERR_GET_CALCOLO_QUANTITA_EFFETTIVA
Dim DataElaborata As String
Dim Giorni As Long
Dim I As Long
Dim NumeroGiorni As Long
Dim Incrementa As Boolean
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim DataPasqua As String
Dim DataPasquaFinePeriodo As String
Dim DataPasquetta As String
Dim DataPasquettaFinePeriodo As String
Dim MeseDataEla As Long
Dim GiornoDataEla As Long



DataPasqua = ""
DataPasquetta = ""
DataPasquaFinePeriodo = ""
DataPasquettaFinePeriodo = ""

If (Year(DataInizioPeriodo) = Year(DataFinePeriodo)) Then
    DataPasqua = CalcolaPasqua(Year(DataInizioPeriodo))
    DataPasquetta = DateAdd("d", 1, DataPasqua)
Else
    DataPasqua = CalcolaPasqua(Year(DataInizioPeriodo))
    DataPasquetta = DateAdd("d", 1, DataPasqua)
    DataPasquaFinePeriodo = CalcolaPasqua(Year(DataFinePeriodo))
    DataPasquettaFinePeriodo = DateAdd("d", 1, DataPasquaFinePeriodo)
End If

sSQL = "SELECT * FROM RV_POFestivita"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

Giorni = DateDiff("d", DataInizioPeriodo, DataFinePeriodo) + 1
DataElaborata = DataInizioPeriodo
NumeroGiorni = 0

For I = 1 To Giorni
    Incrementa = True
    
    MeseDataEla = Month(DataElaborata)
    GiornoDataEla = Day(DataElaborata)
    
    If Me.chkEscludiGiorniFestivi.Value = vbChecked Then
        If DatePart("w", DataElaborata) = 1 Then
            Incrementa = False
        End If
        
        'FESTIVITA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        rs.Filter = "Mese=" & MeseDataEla
        rs.Filter = rs.Filter & " AND Giorno=" & GiornoDataEla
        rs.Filter = rs.Filter & " AND FestivitaNazionale=1"
        
        If Not rs.EOF Then
            Incrementa = False
        End If
        
        rs.Filter = vbNullString
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'PASQUA'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Len(DataPasqua) > 0 Then
            If DateDiff("d", DataElaborata, DataPasqua) = 0 Then
                Incrementa = False
            End If
        End If
        
         If Len(DataPasquetta) > 0 Then
            If DateDiff("d", DataElaborata, DataPasquetta) = 0 Then
                Incrementa = False
            End If
        End If
        
         If Len(DataPasquaFinePeriodo) > 0 Then
            If DateDiff("d", DataElaborata, DataPasquaFinePeriodo) = 0 Then
                Incrementa = False
            End If
        End If
        
        If Len(DataPasquettaFinePeriodo) > 0 Then
            If DateDiff("d", DataElaborata, DataPasquaFinePeriodo) = 0 Then
                Incrementa = False
            End If
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
    
    If Me.chkEscludiSabato.Value = vbChecked Then
        If DatePart("w", DataElaborata) = 7 Then
            Incrementa = False
        End If
    End If
    If Incrementa = True Then
        NumeroGiorni = NumeroGiorni + 1
    End If
    
    DataElaborata = DateAdd("d", 1, DataElaborata)
Next I

rs.Close
Set rs = Nothing

GET_CALCOLO_QUANTITA_EFFETTIVA = NumeroGiorni

Exit Function

ERR_GET_CALCOLO_QUANTITA_EFFETTIVA:
    MsgBox Err.Description, vbCritical, "GET_CALCOLO_QUANTITA_EFFETTIVA"
End Function

Private Function CalcolaPasqua(Anno As Integer) As String
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim m As Double
Dim n As Double
Dim Giorno As Double
Dim mese As Double
   
   If (Anno <= 2099) Then
      m = 24
      n = 5
   ElseIf (Anno <= 2199) Then
      m = 24
      n = 6
   ElseIf (Anno <= 2299) Then
      m = 25
      n = 0
   ElseIf (Anno <= 2399) Then
      m = 26
      n = 1
   ElseIf (Anno <= 2499) Then
      m = 25
      n = 1
   End If
 
   a = Anno Mod 19
   b = Anno Mod 4
   c = Anno Mod 7
   d = ((19 * a) + m) Mod 30
   e = ((2 * b) + (4 * c) + (6 * d) + n) Mod 7
 
   If ((d + e) < 10) Then
      Giorno = d + e + 22
      mese = 3
   Else
      Giorno = d + e - 9
      mese = 4
   End If
 
   If (Giorno = 26 And mese = 4) Then
      Giorno = 19
      mese = 4
   End If
 
   If (Giorno = 25 And mese = 4 And d = 28 And e = 6 And a > 10) Then
      Giorno = 18
      mese = 4
   End If
 
   CalcolaPasqua = DateSerial(Anno, mese, Giorno)
End Function
Private Sub ELIMINA_FLUSSO_DOCUMENTALE_SCADENZA_C(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long)
On Error GoTo ERR_ELIMINA_FLUSSO_DOCUMENTALE
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim IDFunzioneVend As Long
Dim IDFunzioneRata As Long
Dim IDFlussoGruppo As Long
Dim IDFlussoFunzione As Long

IDFunzioneVend = GET_LINK_FUNZIONE(IDTipoOggettoVend)
IDFunzioneRata = GET_LINK_FUNZIONE(IDTipoOggettoRata)

'''''''''''''''''''''''''''''''''GRUPPO FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoGruppo "
sSQL = sSQL & "WHERE Descrizione=" & fnNormString("Contratto -> Scadenza")
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFLussoGruppo")
        rsNew!Descrizione = "Contratto -> Scadenza"
    rsNew.Update
End If

IDFlussoGruppo = fnNotNullN(rsNew!IDFlussoGruppo)

rsNew.Close
Set rsNew = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IDFunzioneVend > 0 Then
    sSQL = "SELECT * FROM FlussoFunzione "
    sSQL = sSQL & "WHERE IDFunzione=" & IDFunzioneRata
    sSQL = sSQL & " AND IDFunzioneSuccessiva=" & IDFunzioneVend
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDFlussoFunzione = fnGetNewKeyTipoOggetto("FlussoFunzione", "IDFlussoFunzione")
            rsNew!IDFunzione = IDFunzioneVend
            rsNew!IDFunzioneSuccessiva = IDFunzioneRata
            rsNew!Cardinalita = 3
            rsNew!TipoAutomatismo = 1
            rsNew!Attributo = 14
            rsNew!TipoDipendenza = 1
            rsNew!IDFlussoGruppo = IDFlussoGruppo
        rsNew.Update
    End If
    
    IDFlussoFunzione = fnNotNullN(rsNew!IDFlussoFunzione)
    
    rsNew.Close
    Set rsNew = Nothing
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FLUSSO FUNZIONE COLLEGATO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggettoCollegato<>" & IDOggettoVend

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    sSQL = "DELETE FROM FlussoFunzioneCollegato "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoRata
    sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
    sSQL = sSQL & " AND IDFlussoFunzione=" & IDFlussoFunzione
    Cn.Execute sSQL
End If
rsNew.Close
Set rsNew = Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''FLUSSO OGGETTI COLLEGATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM FlussoOggettiCollegati "
sSQL = sSQL & "WHERE IDFlussoFunzione=" & IDFlussoFunzione
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggettoRata
sSQL = sSQL & " AND IDOggetto=" & IDOggettoRata
sSQL = sSQL & " AND IDTipoOggettoCollegato=" & IDTipoOggettoVend
sSQL = sSQL & " AND IDOggettoCollegato=" & IDOggettoVend
Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Sub
ERR_ELIMINA_FLUSSO_DOCUMENTALE:
    MsgBox Err.Description, vbCritical, "ELIMINA_FLUSSO_DOCUMENTALE"
    
End Sub
Private Sub ELIMINA_PRODOTTI_SERVIZI(rstmp As ADODB.Recordset)
Dim sSQL As String

    If ((rstmp.EOF) And (rstmp.BOF)) Then Exit Sub
    
    rstmp.MoveFirst
    
    While Not rstmp.EOF
        sSQL = "DELETE FROM RV_POContrattoServiziProdotti "
        sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & fnNotNullN(rstmp!IDRV_POContrattoProdotti)
        Cn.Execute sSQL
    rstmp.MoveNext
    Wend
    
    rstmp.Close
    Set rstmp = Nothing
    
End Sub
Private Function GET_ESISTENZA_PROD_CONTR(IDProdotto As Long, DataInizio As String, DataFine As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_ESISTENZA_PROD_CONTR = False

sSQL = "SELECT IDRV_POContrattoProdotti FROM RV_POIEContrattoProdotti "
sSQL = sSQL & "WHERE DataInizioPeriodo<=" & fnNormDate(DataFine)
sSQL = sSQL & " AND DataFinePeriodo>=" & fnNormDate(DataInizio)
sSQL = sSQL & " AND Chiuso=0"
sSQL = sSQL & " AND ContrattoAttuale=1"
sSQL = sSQL & " AND Offerta=0"
sSQL = sSQL & " AND IDRV_POProdotto=" & IDProdotto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_ESISTENZA_PROD_CONTR = True
End If


rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_CONTROLLO_ESISTENZA_ADD_PROD(IDAddebito As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


GET_CONTROLLO_ESISTENZA_ADD_PROD = False

sSQL = "SELECT IDRV_POInterventoRigheDett FROM RV_POInterventoRigheDett "
sSQL = sSQL & "WHERE IDRV_POInterventoRigheDett=" & IDAddebito

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_ESISTENZA_ADD_PROD = True
End If


rs.CloseResultset
Set rs = Nothing

End Function


Private Function GET_CONTROLLO_ESISTENZA_RIL_PROD(IDRilevamento As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


GET_CONTROLLO_ESISTENZA_RIL_PROD = False

sSQL = "SELECT IDRV_POContatoreRilevamenti FROM RV_POContatoreRilevamenti "
sSQL = sSQL & "WHERE IDRV_POContatoreRilevamenti=" & IDRilevamento

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_ESISTENZA_RIL_PROD = True
End If


rs.CloseResultset
Set rs = Nothing

End Function


Private Sub GET_MODULO_ATTIVATO(Codice As String, IdentificativoProgramma As Long)
On Error GoTo ERR_GET_MODULO_ATTIVATO

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Attivato, DescrizioneModulo FROM RV_POProgrammaModulo "
sSQL = sSQL & "WHERE CodiceModulo=" & fnNormString(Codice)
sSQL = sSQL & " AND IdentificazioneProgramma=" & IdentificativoProgramma

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    MODULO_ATTIVATO = 0
    MODULO_DESCRIZIONE = ""
Else
    MODULO_ATTIVATO = Abs(fnNotNullN(rs!Attivato))
    MODULO_DESCRIZIONE = fnNotNull(rs!DescrizioneModulo)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_GET_MODULO_ATTIVATO:
    MODULO_ATTIVATO = 0
    MODULO_DESCRIZIONE = ""
End Sub

Private Sub GET_MODULO_ATTIVATO_INT(Codice As String, IdentificativoProgramma As Long)
On Error GoTo ERR_GET_MODULO_ATTIVATO

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Attivato, DescrizioneModulo FROM RV_POProgrammaModulo "
sSQL = sSQL & "WHERE CodiceModulo=" & fnNormString(Codice)
sSQL = sSQL & " AND IdentificazioneProgramma=" & IdentificativoProgramma

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    MODULO_ATTIVATO_INT = 0
    MODULO_DESCRIZIONE_INT = ""
Else
    MODULO_ATTIVATO_INT = Abs(fnNotNullN(rs!Attivato))
    MODULO_DESCRIZIONE_INT = fnNotNull(rs!DescrizioneModulo)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_GET_MODULO_ATTIVATO:
    MODULO_ATTIVATO_INT = 0
    MODULO_DESCRIZIONE_INT = ""
End Sub
Private Sub GET_MODULO_ATTIVATO_NOL(Codice As String, IdentificativoProgramma As Long)
On Error GoTo ERR_GET_MODULO_ATTIVATO

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Attivato, DescrizioneModulo FROM RV_POProgrammaModulo "
sSQL = sSQL & "WHERE CodiceModulo=" & fnNormString(Codice)
sSQL = sSQL & " AND IdentificazioneProgramma=" & IdentificativoProgramma

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    MODULO_ATTIVATO_NOL = 0
    MODULO_DESCRIZIONE_NOL = ""
Else
    MODULO_ATTIVATO_NOL = Abs(fnNotNullN(rs!Attivato))
    MODULO_DESCRIZIONE_NOL = fnNotNull(rs!DescrizioneModulo)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_GET_MODULO_ATTIVATO:
    MODULO_ATTIVATO_NOL = 0
    MODULO_DESCRIZIONE_NOL = ""
End Sub
Private Sub GET_MODULO_ATTIVATO_CONT(Codice As String, IdentificativoProgramma As Long)
On Error GoTo ERR_GET_MODULO_ATTIVATO

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Attivato, DescrizioneModulo FROM RV_POProgrammaModulo "
sSQL = sSQL & "WHERE CodiceModulo=" & fnNormString(Codice)
sSQL = sSQL & " AND IdentificazioneProgramma=" & IdentificativoProgramma

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    MODULO_ATTIVATO_CONT = 0
    MODULO_DESCRIZIONE_CONT = ""
Else
    MODULO_ATTIVATO_CONT = Abs(fnNotNullN(rs!Attivato))
    MODULO_DESCRIZIONE_CONT = fnNotNull(rs!DescrizioneModulo)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_GET_MODULO_ATTIVATO:
    MODULO_ATTIVATO_CONT = 0
    MODULO_DESCRIZIONE_CONT = ""
End Sub
Private Function GET_PARAMETRO_UTENTE(IDUtente As Long, IDAzienda As Long, nomeCampo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & nomeCampo & " FROM RV_POParametriUtente "
sSQL = sSQL & "WHERE IDUtente=" & IDUtente
sSQL = sSQL & " AND IDAzienda=" & IDAzienda

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PARAMETRO_UTENTE = 0
Else
    GET_PARAMETRO_UTENTE = fnNotNullN(rs.adoColumns(nomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_INDIRIZZO_EMAIL_CLIENTE(IDAnagrafica As Long) As String
On Error GoTo ERR_GET_INDIRIZZO_EMAIL_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica, EmailInternet "
sSQL = sSQL & "FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_INDIRIZZO_EMAIL_CLIENTE = ""
Else
    GET_INDIRIZZO_EMAIL_CLIENTE = fnNotNull(rs!EMailInternet)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_INDIRIZZO_EMAIL_CLIENTE:
    MsgBox Err.Description, vbCritical, "GET_INDIRIZZO_EMAIL_CLIENTE"
End Function
Private Function GET_LINK_UM_PERIODO_PRED(IDProdotto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_UM_PERIODO_PRED = 0

sSQL = "SELECT IDRV_POUnitaDiMisuraPeriodo "
sSQL = sSQL & "FROM RV_POProdotto "
sSQL = sSQL & "WHERE IDRV_POProdotto=" & IDProdotto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_UM_PERIODO_PRED = fnNotNullN(rs!IDRV_POUnitaDiMisuraPeriodo)
End If

If GET_LINK_UM_PERIODO_PRED > 0 Then Exit Function

sSQL = "SELECT IDRV_POUnitaDiMisuraPeriodo FROM RV_POParametriAzienda "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch


Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_UM_PERIODO_PRED = fnNotNullN(rs!IDRV_POUnitaDiMisuraPeriodo)
End If

End Function
Private Function GET_LINK_RATEIZZAZIONE_PRED(IDProdotto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_RATEIZZAZIONE_PRED = 0

sSQL = "SELECT IDRateizzazioneProdotto "
sSQL = sSQL & "FROM RV_POProdotto "
sSQL = sSQL & "WHERE IDRV_POProdotto=" & IDProdotto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_RATEIZZAZIONE_PRED = fnNotNullN(rs!IDRateizzazioneProdotto)
End If

If GET_LINK_RATEIZZAZIONE_PRED > 0 Then Exit Function

sSQL = "SELECT IDRateizzazioneProdotto FROM RV_POParametriAzienda "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch


Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_RATEIZZAZIONE_PRED = fnNotNullN(rs!IDRateizzazioneProdotto)
End If

End Function
Private Sub GET_PARAMETRI_STRINGA_FATT()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

NO_CALCOLO_PERIODO_FATT = 0

sSQL = "SELECT NoCalcoloPeriodoFattSenzaProdotto "
sSQL = sSQL & " FROM RV_POStringaPeriodoTesta "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    NO_CALCOLO_PERIODO_FATT = fnNotNullN(rs!NoCalcoloPeriodoFattSenzaProdotto)
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub CALCOLA_ISTAT_CONTRATTO()
On Error GoTo ERR_CALCOLA_ISTAT_CONTRATTO
Dim Testo As String

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

If Me.cboTipoImpostazione.CurrentID = 3 Then Exit Sub

If (IDIstatContratto > 0) Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Al contratto è stato già calcolato un adeguamento Istat" & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then Exit Sub
    
End If

IMPORTO_CONTRATTO_ISTAT = fnNotNullN(m_Document("ImportoContrattoAttuale").Value)

frmIstat.Show vbModal

If AGGIORNA_DA_ISTAT = 1 Then
    GET_CARATTERISTICHE_RISORSA Me.MSFlexGrid1
    
    OnSave
    
    AGGIORNA_DA_ISTAT = 0
End If

Exit Sub
ERR_CALCOLA_ISTAT_CONTRATTO:
    MsgBox Err.Description, vbCritical, "CALCOLA_ISTAT_CONTRATTO"
End Sub
Private Sub AGGIORNA_RATE_DA_ISTAT(IDContratto As Long, IDTipoRateizzazione As Long, MaggiorazioneContratto As Double)
On Error GoTo ERR_AGGIORNA_RATE_DA_ISTAT
Dim sSQL As String
Dim numerorate As Long
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoMaggRata As Double
Dim ImportoElaborato As Double
Dim DataPrimaRataDisp As String
Dim ICont As Long
Dim rsAgg As ADODB.Recordset
Dim NumeroRatePagate As Long
Dim DifferenzaRatePrec As Double
Dim ImportoRataMaggDiff As Double
Dim IFor As Long
Dim rsNew As ADODB.Recordset
Dim NumeroNuovaRata As Long

numerorate = 0

sSQL = "SELECT * FROM RV_PORateizzazione "
sSQL = sSQL & "WHERE IDRV_PORateizzazione=" & IDTipoRateizzazione

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    numerorate = fnNotNullN(rs!numerorate)
    NumeroNuovaRata = numerorate + 1
End If

rs.CloseResultset
Set rs = Nothing


numerorate = GET_NUMERO_RATE_CONTRATTO(IDContratto)

If numerorate = 0 Then Exit Sub

ImportoMaggRata = FormatNumber((MaggiorazioneContratto / numerorate), 2)

'AGGIORNAMENTO RATE NON PAGATE'''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PORateContratto "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Adeguamento=0"
sSQL = sSQL & " AND Manuale=0"
sSQL = sSQL & " AND ContrattoAttuale=1"
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"
sSQL = sSQL & " AND Fatturata=0"

Set rsAgg = New ADODB.Recordset

rsAgg.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
ImportoElaborato = 0
ICont = 1
While Not rsAgg.EOF
    If ICont = 1 Then
        DataPrimaRataDisp = rsAgg!DataRata
    End If
    If ICont < numerorate Then
        rsAgg!ImportoRata = rsAgg!ImportoRata + ImportoMaggRata
    Else
        rsAgg!ImportoRata = rsAgg!ImportoRata + (MaggiorazioneContratto - ImportoElaborato)
    End If
    rsAgg.Update
    
    ImportoElaborato = ImportoElaborato + ImportoMaggRata
    ICont = ICont + 1
rsAgg.MoveNext
Wend

rsAgg.Close
Set rsAgg = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''INSERIMENTO NUOVE RATE PER DIFFERENZA CON LE RATE PRECEDENTI PAGATE''''''''''''''


NumeroRatePagate = GET_NUMERO_RATE_PAGATE(IDContratto)

If NumeroRatePagate = 0 Then Exit Sub

DifferenzaRatePrec = MaggiorazioneContratto - ImportoElaborato
ImportoRataMaggDiff = FormatNumber((DifferenzaRatePrec / NumeroRatePagate), 2)

sSQL = "SELECT * FROM RV_PORateContratto "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Adeguamento=0"
sSQL = sSQL & " AND Manuale=0"
sSQL = sSQL & " AND ContrattoAttuale=1"
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"
sSQL = sSQL & " AND Fatturata=1"

Set rs = Cn.OpenResultset(sSQL)

sSQL = "SELECT * FROM RV_PORateContratto "
sSQL = sSQL & "WHERE IDRV_POContratto=0"

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
ICont = 1
ImportoElaborato = 0
While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_PORateContratto = fnGetNewKey("RV_PORateContratto", "IDRV_PORateContratto")
        rsNew!IDRV_POContratto = IDContratto
        rsNew!numerorata = NumeroNuovaRata
        rsNew!DataRata = DataPrimaRataDisp
        rsNew!IDPagamentoRata = Me.cboPagamentoRate.CurrentID
        If ICont < NumeroRatePagate Then
            rsNew!ImportoRata = ImportoRataMaggDiff
        Else
            rsNew!ImportoRata = DifferenzaRatePrec - ImportoElaborato
        End If
        rsNew!Fatturata = 0
        rsNew!mese = Month(DataPrimaRataDisp)
        rsNew!Anno = Year(DataPrimaRataDisp)
        rsNew!Periodo = "Adeguamento Istat riferimento dal " & rs!DataInizioPeriodo & " al " & rs!DataFinePeriodo
        rsNew!Manuale = 0
        rsNew!ContrattoAttuale = Me.chkContrattoAttuale.Value
        rsNew!IDRV_POContrattoPadre = Me.txtIDContrattoPadre.Value
        rsNew!DataInizioPeriodo = rs!DataInizioPeriodo
        rsNew!DataFinePeriodo = rs!DataFinePeriodo
        rsNew!NonFatturare = 0
        rsNew!IDTipoOggetto = fnGetTipoOggetto("RV_PORateContratto")
        rsNew!IDOggetto = GET_LINK_OGGETTO(0, rsNew!IDTipoOggetto, fnNotNullN(rsNew!numerorata), rsNew!DataRata)
        rsNew!Adeguamento = 0
        
    rsNew.Update
    ImportoElaborato = ImportoElaborato + ImportoRataMaggDiff
    NumeroNuovaRata = NumeroNuovaRata + 1
    ICont = ICont + 1
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If ((IMPORTO_REG_IMP > 0) And (LINK_ARTICOLO_REG_IMP > 0)) Then
    sSQL = "SELECT * FROM RV_PORateContratto "
    sSQL = sSQL & "WHERE IDRV_POContratto=0"
    
    Set rsNew = New ADODB.Recordset
    rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    rsNew.AddNew
        rsNew!IDRV_PORateContratto = fnGetNewKey("RV_PORateContratto", "IDRV_PORateContratto")
        rsNew!IDRV_POContratto = IDContratto
        rsNew!numerorata = NumeroNuovaRata
        rsNew!DataRata = DataPrimaRataDisp
        rsNew!IDPagamentoRata = Me.cboPagamentoRate.CurrentID
        rsNew!ImportoRata = IMPORTO_REG_IMP
        rsNew!Fatturata = 0
        rsNew!mese = Month(DataPrimaRataDisp)
        rsNew!Anno = Year(DataPrimaRataDisp)
        rsNew!Periodo = "Registro d'imposta riferito all'anno precedente"
        rsNew!Manuale = 1
        rsNew!ContrattoAttuale = Me.chkContrattoAttuale.Value
        rsNew!IDRV_POContrattoPadre = Me.txtIDContrattoPadre.Value
        rsNew!DataInizioPeriodo = DataPrimaRataDisp
        rsNew!DataFinePeriodo = DataPrimaRataDisp
        rsNew!NonFatturare = 0
        rsNew!IDTipoOggetto = fnGetTipoOggetto("RV_PORateContratto")
        rsNew!IDOggetto = GET_LINK_OGGETTO(0, rsNew!IDTipoOggetto, fnNotNullN(rsNew!numerorata), rsNew!DataRata)
        rsNew!Adeguamento = 0
        rsNew!IDArticolo = LINK_ARTICOLO_REG_IMP
    rsNew.Update
    
    rsNew.Close
    Set rsNew = Nothing
End If

Exit Sub
ERR_AGGIORNA_RATE_DA_ISTAT:
    MsgBox Err.Description, vbCritical, "AGGIORNA_RATE_DA_ISTAT"
End Sub
Private Function GET_NUMERO_RATE_PAGATE(IDContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT COUNT(IDRV_PORateContratto) AS NumeroRecord "
sSQL = sSQL & "FROM RV_PORateContratto "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Adeguamento=0"
sSQL = sSQL & " AND Manuale=0"
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"
sSQL = sSQL & " AND Fatturata=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_RATE_PAGATE = 0
Else
    GET_NUMERO_RATE_PAGATE = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing
    
End Function
Private Function GET_NUMERO_RATE_CONTRATTO(IDContratto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT COUNT(IDRV_PORateContratto) AS NumeroRecord "
sSQL = sSQL & "FROM RV_PORateContratto "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto
sSQL = sSQL & " AND Adeguamento=0"
sSQL = sSQL & " AND Manuale=0"
sSQL = sSQL & " AND ContrattoAttuale=1"
sSQL = sSQL & " AND IDRV_POContrattoAdeguamento IS NULL"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_RATE_CONTRATTO = 0
Else
    GET_NUMERO_RATE_CONTRATTO = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing


End Function

Private Sub AVVIA_FATTURAZIONE_CONTRATTO(IDContratto As Long)
        
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    Link_Contratto = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    lblLinkFattContratto.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POCreazioneDocumenti"))
    Me.lblLinkFattContratto.IDReturn = 0
    Me.lblLinkFattContratto.RunApplication
    
End Sub

Private Function GET_CONTROLLO_RIGA_PROD_FATT(ID As Long) As Boolean
On Error GoTo ERR_GET_CONTROLLO_RIGA_PROD_FATT
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_RIGA_PROD_FATT = False

sSQL = "SELECT IDRV_PORateContratto, IDRV_POContrattoProdotti "
sSQL = sSQL & "FROM RV_PORateContratto "
sSQL = sSQL & " WHERE IDRV_POContrattoProdotti=" & ID
sSQL = sSQL & " AND Fatturata=1"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_RIGA_PROD_FATT = True
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_CONTROLLO_RIGA_PROD_FATT:
    'MsgBox Err.Description, vbCritical, "GET_CONTROLLO_RIGA_PROD_FATT"
End Function
Private Sub ELIMINA_INTERVENTI_COLLEGATI(IDRigaProdottoContratto As Long)
On Error GoTo ERR_ELIMINA_INTERVENTI_COLLEGATI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POIntervento FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDRigaProdottoContratto

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    sSQL = "DELETE FROM RV_POInterventoEmail "
    sSQL = sSQL & "WHERE IDRV_POIntervento=" & fnNotNullN(rs!IDRV_POIntervento)
    Cn.Execute sSQL
    
    sSQL = "DELETE FROM RV_POInterventoRigheDett "
    sSQL = sSQL & "WHERE IDRV_POIntervento=" & fnNotNullN(rs!IDRV_POIntervento)
    Cn.Execute sSQL
    
    sSQL = "DELETE FROM RV_PODocumentazione "
    sSQL = sSQL & "WHERE IDRV_POIntervento=" & fnNotNullN(rs!IDRV_POIntervento)
    Cn.Execute sSQL
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

sSQL = "DELETE FROM RV_POIntervento "
sSQL = sSQL & "WHERE IDRV_POContrattoProdotti=" & IDRigaProdottoContratto
Cn.Execute sSQL
Exit Sub
ERR_ELIMINA_INTERVENTI_COLLEGATI:
    MsgBox Err.Description, vbCritical, "ELIMINA_INTERVENTI_COLLEGATI"

End Sub
Private Sub RICHIEDI_INTERVENTI()
Screen.MousePointer = 11
GET_GRIGLIA_INTERVENTI
Screen.MousePointer = 0

End Sub
Private Sub AGGIORNA_DESCR_RATE_PROD(IDRigaProdotto As Long, matricolaOLD As String, matricolaNew As String)
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT Periodo FROM RV_PORateContratto "
sSQL = sSQL & " WHERE IDRV_POContrattoProdotti=" & IDRigaProdotto
sSQL = sSQL & " AND ((Fatturata IS NULL) OR (Fatturata=0))"

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rs.EOF
    rs!Periodo = Replace(rs!Periodo, matricolaOLD, matricolaNew)
    rs.Update
rs.MoveNext
Wend
rs.Close
Set rs = Nothing
End Sub
