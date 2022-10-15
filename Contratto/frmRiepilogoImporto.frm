VERSION 5.00
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmRiepilogoImporto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RIEPILOGO IMPORTI CONTRATTO"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3870
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
   ScaleHeight     =   2970
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "CHIUDI"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
   Begin DMTEDITNUMLib.dmtCurrency txtImportoContratto 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   330
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
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
   Begin DMTEDITNUMLib.dmtCurrency txtImportoAdeg 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
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
   Begin DMTEDITNUMLib.dmtCurrency txtImportoTotale 
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
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
   Begin VB.Label Label1 
      Caption         =   "Totale"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1590
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Adeguamenti"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   990
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Contratto"
      Height          =   255
      Index           =   35
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmRiepilogoImporto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.txtImportoContratto.Value = GET_IMPORTO_CONTRATTO(Link_Contratto)
    Me.txtImportoAdeg.Value = GET_IMPORTO_ADEGUAMENTO(Link_Contratto)
    Me.txtImportoTotale.Value = Me.txtImportoContratto.Value + Me.txtImportoAdeg.Value
End Sub
    
Private Function GET_IMPORTO_CONTRATTO(IDContratto As Long) As Double
On Error GoTo ERR_GET_TOTALE_ADEGUAMENTI_DETTAGLIO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_IMPORTO_CONTRATTO = 0

sSQL = "SELECT ImportoContrattoAttuale  "
sSQL = sSQL & "FROM RV_POContratto "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto



Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_IMPORTO_CONTRATTO = 0
Else
    GET_IMPORTO_CONTRATTO = fnNotNullN(rs!ImportoContrattoAttuale)
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_TOTALE_ADEGUAMENTI_DETTAGLIO:
    MsgBox Err.Description, vbCritical, "GET_IMPORTO_CONTRATTO"
End Function
Private Function GET_IMPORTO_ADEGUAMENTO(IDContratto As Long) As Double
On Error GoTo ERR_GET_TOTALE_ADEGUAMENTI_DETTAGLIO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_IMPORTO_ADEGUAMENTO = 0


sSQL = "SELECT SUM(Importo) AS TotaleAdeguamenti "
sSQL = sSQL & "FROM RV_POIEAdeguamentiContratto "
sSQL = sSQL & " WHERE IDRV_POContratto=" & IDContratto



Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_IMPORTO_ADEGUAMENTO = 0
Else
    GET_IMPORTO_ADEGUAMENTO = fnNotNullN(rs!TotaleAdeguamenti)
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_TOTALE_ADEGUAMENTI_DETTAGLIO:
    MsgBox Err.Description, vbCritical, "GET_IMPORTO_ADEGUAMENTO"
End Function

