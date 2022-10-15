VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#9.1#0"; "DMTDataCmb.OCX"
Begin VB.Form FrmParametri 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creazione documenti  (Passo 3 di 4)"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   Icon            =   "FrmFine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   0
      Picture         =   "FrmFine.frx":030A
      ScaleHeight     =   4395
      ScaleWidth      =   2235
      TabIndex        =   15
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Avanti"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton CmdFine 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton CmdIndietro 
      Caption         =   "Indietro"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo di documento"
      Height          =   3615
      Left            =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin DMTDataCmb.DMTCombo CboPagamento 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   3120
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo CboMagazzino 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDATETIMELib.dmtDate DataDoc 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo CboValuta 
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo CboSezionale 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo CboTipoOggetto 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Pagamento di Default"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Magazzino"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Data documento"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Valuta"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   7
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Sezionale"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3495
      End
   End
End
Attribute VB_Name = "FrmParametri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Indica se ci sono parametri di default per l'azienda
'per l'importazione dei dati
'0 = non ci sono dati
'1 = ci sono i dati
Public NewRecDefault As Integer

Private Function ControlloInserimento() As String
    ControlloInserimento = ""
    
    If Me.CboTipoOggetto.Text = "" Then
        ControlloInserimento = ControlloInserimento & "Inserire il tipo di documento" & vbCrLf
    Else
        Link_TipoOggetto = Me.CboTipoOggetto.CurrentID
        Var_TipoOggetto = Me.CboTipoOggetto.Text
    End If
    If Me.CboSezionale.CurrentID = 0 Then
        ControlloInserimento = ControlloInserimento & "Inserire il tipo di sezionale del documento" & vbCrLf
    Else
        Link_Sezionale = Me.CboSezionale.CurrentID
    End If
    If Me.CboValuta.CurrentID = 0 Then
        ControlloInserimento = ControlloInserimento & "Inserire il tipo di valuta del documento" & vbCrLf
    Else
        Link_Valuta = Me.CboValuta.CurrentID
    End If
    If Me.CboMagazzino.CurrentID = 0 Then
        ControlloInserimento = ControlloInserimento & "Inserire il magazzino" & vbCrLf
    Else
        Link_Magazzino = Me.CboMagazzino.CurrentID
    End If
    If Me.CboPagamento.CurrentID = 0 Then
        ControlloInserimento = ControlloInserimento & "Inserire il tipo di pagamento di default" & vbCrLf
    Else
        Link_PagamentoDefault = Me.CboPagamento.CurrentID
    End If
    
    Data_Documento = Me.DataDoc.Text
    
End Function
Private Sub cmdAvanti_Click()
    If ControlloInserimento = "" Then
        'fncPassaggioDocumenti
        SalvataggioParametriDefault
        Unload Me
    Else
        MsgBox ControlloInserimento, vbInformation, "Parametri mancanti"
    End If
End Sub

Private Sub CmdFine_Click()
    Dim Risposta As Integer
    Risposta = MsgBox("Vuoi abbandonare lo wizard per il passaggio degli interventi in contabilità?", vbInformation + vbYesNo, "Abbandono")
    If Risposta = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdIndietro_Click()
    
    Unload Me
    
End Sub
Private Sub SalvataggioParametriDefault()
On Error GoTo ERR_SalvataggioParametriDefault
    Dim sSQL As String
    
    If NewRecDefault = 0 Then
    'NUOVO RECORD DI DEFAULT
        sSQL = "INSERT INTO RV_POParametriDefault "
        sSQL = sSQL & "(IDAzienda, IDTipoOggetto, IDSezionale, IDValuta, IDMagazzino, IDPagamento)"
        sSQL = sSQL & " VALUES ("
        sSQL = sSQL & VarIDAzienda & ", "
        sSQL = sSQL & Me.CboTipoOggetto.CurrentID & ", "
        sSQL = sSQL & Me.CboSezionale.CurrentID & ", "
        sSQL = sSQL & Me.CboValuta.CurrentID & ", "
        sSQL = sSQL & Me.CboMagazzino.CurrentID & ", "
        sSQL = sSQL & Me.CboPagamento.CurrentID & ")"
        
        
        
    Else
    'AGGIORNA IL RECORD DI DEFAULT
        sSQL = "UPDATE RV_POParametriDefault SET "
        sSQL = sSQL & "IDTipoOggetto=" & Me.CboTipoOggetto.CurrentID & ", "
        sSQL = sSQL & "IDSezionale=" & Me.CboSezionale.CurrentID & ", "
        sSQL = sSQL & "IDValuta=" & Me.CboValuta.CurrentID & ", "
        sSQL = sSQL & "IDMagazzino=" & Me.CboMagazzino.CurrentID & ", "
        sSQL = sSQL & "IDPagamento=" & Me.CboPagamento.CurrentID
        sSQL = sSQL & " WHERE IDAzienda= " & VarIDAzienda
        
    End If
    
    CnDMT.Execute sSQL
    
Exit Sub

ERR_SalvataggioParametriDefault:
    MsgBox Err.Description, vbCritical, "Salvataggio Parametri Default"
End Sub
Private Sub DefaultImportazione()
    Dim rs As DmtADOLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT * From RV_POParametriDefault Where IDAzienda=" & VarIDAzienda
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        Me.CboSezionale.WriteOn IIf(IsNull(rs!IDSezionale), 0, rs!IDSezionale)
        Me.CboTipoOggetto.WriteOn IIf(IsNull(rs!IDTipoOggetto), 0, rs!IDTipoOggetto)
        Me.CboValuta.WriteOn IIf(IsNull(rs!IDValuta), 0, rs!IDValuta)
        Me.CboMagazzino.WriteOn IIf(IsNull(rs!IDMagazzino), 0, rs!IDMagazzino)
        Me.CboPagamento.WriteOn IIf(IsNull(rs!IDPagamento), 0, rs!IDPagamento)
        NewRecDefault = 1
    Else
        Me.CboSezionale.WriteOn 0
        Me.CboTipoOggetto.WriteOn 0
        Me.CboValuta.WriteOn 0
        Me.CboMagazzino.WriteOn 0
        Me.CboPagamento.WriteOn 0
        
        NewRecDefault = 0
    End If
    
    
End Sub
Private Sub Form_Load()
On Error GoTo ERR_Form_Load
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    
    fncTipoOggetto
    fncSezionale
    fncValuta
    fncMagazzino
    fncPagamento
    Me.DataDoc.Text = Date
    
    
    DefaultImportazione
Exit Sub

ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERR_Form_Unload
    If Me.cmdAvanti.Value = True Then
        frmCreazioneDocumenti.Show
        Exit Sub
    End If
    
    
    If Me.cmdIndietro.Value = True Then
        FrmMovimenti.Show
        Exit Sub
        
    End If

    CnDMT.CloseConnection
    Set CnDMT = Nothing
Exit Sub

ERR_Form_Unload:
    MsgBox Err.Description, vbCritical, "Form_Unload"
    
    
End Sub
Private Function SettaggioProgressBar() As Long

    
End Function
Private Sub fncTipoOggetto()
    Dim sSQL As String
    
    sSQL = "SELECT IDTipoOggetto, Oggetto"
    sSQL = sSQL & " FROM TipoOggetto"
    sSQL = sSQL & " WHERE IDGestore=15"
    sSQL = sSQL & " ORDER BY Oggetto"
    

    With Me.CboTipoOggetto
    Set .Database = CnDMT
    .DisplayField = "Oggetto"
    .AddFieldKey "IDTipoOggetto"
    .SQL = sSQL
    .Refresh
    End With
    
End Sub
    
    
Private Sub fncSezionale()
    Dim sSQL As String
    
    sSQL = "SELECT IDSezionale, Sezionale"
    sSQL = sSQL & " FROM Sezionale"
    sSQL = sSQL & " WHERE ((IDFiliale=" & VarIDFiliale & ") AND (IDRegistroIva = 1))"
    sSQL = sSQL & " ORDER BY Sezionale"

    With Me.CboSezionale
    Set .Database = CnDMT
    .DisplayField = "Sezionale"
    .AddFieldKey "IDSezionale"
    .SQL = sSQL
    .Refresh
    End With
    
End Sub
Private Sub fncValuta()
    Dim sSQL As String
    'Dim sSQLValuta As String
    'Dim rs As DmtADOLib.adoResultset
    
    
    sSQL = "SELECT IDValuta, Valuta"
    sSQL = sSQL & " FROM Valuta"
    'sSQL = sSQL & " WHERE ((IDFiliale=" & VarIDFiliale & ") AND (IDRegistroIva = 1))"
    sSQL = sSQL & " ORDER BY Valuta"

    With Me.CboValuta
    Set .Database = CnDMT
    .DisplayField = "Valuta"
    .AddFieldKey "IDValuta"
    .SQL = sSQL
    .Refresh
    End With
    
End Sub
Private Sub fncMagazzino()
    Dim sSQL As String
    'Dim sSQLValuta As String
    'Dim rs As DmtADOLib.adoResultset
    
    
    sSQL = "SELECT IDMagazzino, Magazzino"
    sSQL = sSQL & " FROM Magazzino"
    sSQL = sSQL & " WHERE IDAzienda = " & VarIDAzienda
    sSQL = sSQL & " ORDER BY Magazzino"

    With Me.CboMagazzino
    Set .Database = CnDMT
    .DisplayField = "Magazzino"
    .AddFieldKey "IDMagazzino"
    .SQL = sSQL
    .Refresh
    End With

End Sub
Private Sub fncPagamento()
    Dim sSQL As String

    
    
    sSQL = "SELECT IDPagamento, Pagamento"
    sSQL = sSQL & " FROM Pagamento"
    sSQL = sSQL & " ORDER BY Pagamento"

    With Me.CboPagamento
    Set .Database = CnDMT
    .DisplayField = "Pagamento"
    .AddFieldKey "IDPagamento"
    .SQL = sSQL
    .Refresh
    End With

End Sub
