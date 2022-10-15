VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.9#0"; "DmtCodDesc.ocx"
Begin VB.Form FrmInizio 
   Caption         =   "Rinnovi contratti (Passo 1 di 3)"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInizio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Fine 
      Caption         =   "Fine"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdIndietro 
      Caption         =   "Indietro"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Avanti"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin DMTDATETIMELib.dmtDate txtAData 
      Height          =   315
      Left            =   5160
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   253
      BackColor       =   16777215
      Appearance      =   1
   End
   Begin DMTDATETIMELib.dmtDate txtDaData 
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   253
      BackColor       =   16777215
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
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
      Height          =   4815
      Left            =   0
      Picture         =   "FrmInizio.frx":4781A
      ScaleHeight     =   4815
      ScaleWidth      =   2775
      TabIndex        =   4
      Top             =   0
      Width           =   2775
   End
   Begin DMTDataCmb.DMTCombo cboIstat 
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Top             =   2880
      Width           =   3975
      _ExtentX        =   7011
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
   Begin DMTDataCmb.DMTCombo cboTipoContratto 
      Height          =   315
      Left            =   2880
      TabIndex        =   12
      Top             =   1560
      Width           =   3975
      _ExtentX        =   7011
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
   Begin DmtCodDescCtl.DmtCodDesc CDCliente 
      Height          =   615
      Left            =   2880
      TabIndex        =   14
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1085
      PropCodice      =   $"FrmInizio.frx":71B5C
      BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PropDescrizione =   $"FrmInizio.frx":71BBD
      BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MenuFunctions   =   $"FrmInizio.frx":71C0D
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
   Begin VB.Label Label2 
      Caption         =   "Tipo contratto"
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   13
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "I.S.T.A.T."
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   11
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "A data scadenza"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Da data scadenza"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   9
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RINNOVO CONTRATTI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmInizio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents m_App As DMTRunAppLib.Application
Attribute m_App.VB_VarHelpID = -1

Public Sub ConnessioneADO()
    If Not (CnDMT Is Nothing) Then
        CnDMT.CloseConnection
        Set CnDMT = Nothing
    End If
    
    Set CnDMT = m_App.Database.Connection
  
    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
    
End Sub
Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set m_App = NewValue
End Property
Public Property Get Application() As DMTRunAppLib.Application
    Set Application = m_App
End Property
Private Sub cmdAvanti_Click()
On Error GoTo ERR_cmdAvanti_Click
    If DateDiff("d", Me.txtDaData.Text, Me.txtAData.Text) < 0 Then
        MsgBox "Le date non sono congruenti", vbInformation, "Impossibile continuare"
    Else
        Unload Me
    End If
Exit Sub
ERR_cmdAvanti_Click:
MsgBox Err.Description, vbCritical, "cmdAvanti_Click"
End Sub

Private Sub Form_Load()
    'ConnessioneADODBLib
    'PrelevaAzienda
    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    If b_Loading = True Then
        InitControlli
    End If
End Sub
Public Sub InitControlli()
    With Me.cboIstat
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POIstat"
        .DisplayField = "Istat"
        .Sql = "SELECT * FROM RV_POIstat ORDER BY IDRV_POIstat DESC"
        .Fill
    End With
    With Me.cboTipoContratto
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoContratto"
        .DisplayField = "TipoContratto"
        .Sql = "SELECT * FROM RV_POTipoContratto ORDER BY TipoContratto"
        .Fill
    End With
    
    'Anagrafica cliente
    With Me.CDCliente
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hWnd
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

End Sub
Private Sub Form_Unload(Cancel As Integer)

If Me.cmdAvanti.Value = True Then
    Var_DaDataRinnovo = Me.txtDaData.Text
    Var_ADataRinnovo = Me.txtAData.Text
    Link_cliente_Ric = Me.CDCliente.KeyFieldID
    Link_Tipo_Contratto_Ric = Me.cboTipoContratto.CurrentID
    
    FrmContrattiPerRinnovo.Show
    Exit Sub
End If
If cmdAnnulla.Value = True Then
    ChiusuraConnessione
    Exit Sub
    
End If



Exit Sub
End Sub
Private Sub cboIstat_Click()
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT Percentuale FROM RV_POIStat WHERE IDRV_POIstat=" & Me.cboIstat.CurrentID
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        PercentualeIstat = rs!Percentuale
        Link_Istat = Me.cboIstat.CurrentID
    Else
        PercentualeIstat = 0
         Link_Istat = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Sub
Private Sub INIT_CONTROLLI()


    
End Sub
