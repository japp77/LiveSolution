VERSION 5.00
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmIstat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADEGUAMENTO ISTAT CONTRATTO"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14685
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIstat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   14685
   StartUpPosition =   2  'CenterScreen
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
      Left            =   12240
      TabIndex        =   8
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impostazione Istat"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   14535
      Begin DMTEDITNUMLib.dmtNumber txtPercIstat 
         Height          =   315
         Left            =   5040
         TabIndex        =   2
         Top             =   600
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTDataCmb.DMTCombo cboIstat 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
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
      Begin DMTEDITNUMLib.dmtNumber txtImportoContratto 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
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
      Begin DMTEDITNUMLib.dmtNumber txtNuovoImporto 
         Height          =   315
         Left            =   6000
         TabIndex        =   3
         Top             =   600
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
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
      Begin DMTEDITNUMLib.dmtNumber txtMaggiorazione 
         Height          =   315
         Left            =   7680
         TabIndex        =   4
         Top             =   600
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
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
      Begin DMTEDITNUMLib.dmtNumber txtPercRegImp 
         Height          =   315
         Left            =   9360
         TabIndex        =   5
         Top             =   600
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecimalPlaces   =   5
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DmtCodDescCtl.DmtCodDesc CDArticoloAdeg 
         Height          =   615
         Left            =   10680
         TabIndex        =   6
         Tag             =   "Articolo utilizzato per il registro d'imposta"
         Top             =   340
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         PropCodice      =   $"frmIstat.frx":4781A
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmIstat.frx":47875
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmIstat.frx":478DC
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
      Begin DMTEDITNUMLib.dmtNumber txtImportoRegImp 
         Height          =   315
         Left            =   12840
         TabIndex        =   7
         Top             =   600
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
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
      Begin VB.Label Label1 
         Caption         =   "Importo reg. imp."
         Height          =   255
         Index           =   6
         Left            =   12840
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "% Reg. imp."
         Height          =   255
         Index           =   5
         Left            =   9360
         TabIndex        =   15
         ToolTipText     =   "Percentuale registro d'imposta"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Maggiorazione"
         Height          =   255
         Index           =   4
         Left            =   7680
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nuovo importo"
         Height          =   255
         Index           =   3
         Left            =   6000
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Importo partenza"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Istat"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmIstat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboIstat_Click()
On Error GoTo ERR_cboIstat_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Me.txtPercIstat.Value = 0

sSQL = "SELECT Percentuale FROM RV_POIstat "
sSQL = sSQL & "WHERE IDRV_POIstat=" & Me.cboIstat.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtPercIstat.Value = fnNotNullN(rs!Percentuale)
End If

rs.CloseResultset
Set rs = Nothing

CALCOLA

Exit Sub
ERR_cboIstat_Click:
    MsgBox Err.Description, vbCritical, "cboIstat_Click"
End Sub

Private Sub cmdConferma_Click()

    frmMain.txtImportoAttuale.Value = Me.txtNuovoImporto.Value
    
    IDIstatContratto = Me.cboIstat.CurrentID
    DescrIstaContratto = Me.cboIstat.Text
    MaggIstatContratto = Me.txtMaggiorazione.Value
    IMPORTO_REG_IMP = Me.txtImportoRegImp.Value
    LINK_ARTICOLO_REG_IMP = Me.CDArticoloAdeg.KeyFieldID
    
    AGGIORNA_DA_ISTAT = 1
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    IMPORTO_REG_IMP = 0
    LINK_ARTICOLO_REG_IMP = 0
    
    AGGIORNA_DA_ISTAT = 0

    INIT_CONTROLLI
    
    INIT_VARIABILI
    
    CALCOLA
    
End Sub
Private Sub INIT_VARIABILI()
On Error GoTo ERR_INIT_VARIABILI

    Me.txtImportoContratto.Value = IMPORTO_CONTRATTO_ISTAT
    Me.cboIstat.WriteOn 0
    Me.txtPercIstat.Value = 0

Exit Sub
ERR_INIT_VARIABILI:
    MsgBox Err.Description, vbCritical, "INIT_VARIABILI"
End Sub
Private Sub INIT_CONTROLLI()
On Error GoTo ERR_INIT_CONTROLLI
    'Istat per contratto
    With Me.cboIstat
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POIstat"
        .DisplayField = "Istat"
        .SQL = "SELECT * FROM RV_POIstat"
        .Fill
    End With
    
    With Me.CDArticoloAdeg
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
Exit Sub
ERR_INIT_CONTROLLI:
    MsgBox Err.Description, vbCritical, "INIT_CONTROLLI"
End Sub
Private Sub CALCOLA()
On Error GoTo ERR_CALCOLA
Dim Importo As Double
Dim Maggiorazione As Double
Dim ImportoRegImp As Double

    Importo = Me.txtImportoContratto.Value + ((Me.txtImportoContratto.Value / 100) * Me.txtPercIstat.Value)

    Importo = FormatNumber(Importo, 2)
    Me.txtNuovoImporto.Value = Importo
    Maggiorazione = FormatNumber((Importo - Me.txtImportoContratto.Value), 2)
    Me.txtMaggiorazione.Value = Maggiorazione
    
    ImportoRegImp = FormatNumber(((Importo / 100) * Me.txtPercRegImp.Value), 2)
        
    Me.txtImportoRegImp.Value = ImportoRegImp
    
Exit Sub
ERR_CALCOLA:
    MsgBox Err.Description, vbCritical, "CALCOLA"
End Sub

Private Sub CONFERMA()
    
    
End Sub

Private Sub txtPercRegImp_Change()
       CALCOLA
End Sub
