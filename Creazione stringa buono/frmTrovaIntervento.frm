VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.0#0"; "DMTDataCmb.OCX"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.7#0"; "DmtCodDesc.ocx"
Begin VB.Form frmTrovaIntervento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TROVA COMMESSA DA ASSOCIARE"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12990
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
   ScaleHeight     =   6060
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAssociaPrimaDisponibile 
      Caption         =   "Crea una nuova commessa o associa alla prima disponibile"
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
      Left            =   9120
      TabIndex        =   12
      Top             =   1680
      Width           =   3615
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   6376
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
      Caption         =   "Parametri di ricerca"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12735
      Begin VB.CommandButton cmdNuovaScheda 
         Caption         =   "Associa ad una nuova commessa"
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
         Left            =   4680
         TabIndex        =   11
         Top             =   1680
         Width           =   3615
      End
      Begin VB.CommandButton cmdRicerca 
         Caption         =   "RICERCA"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   3615
      End
      Begin DmtCodDescCtl.DmtCodDesc CDParcoNatanti 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1085
         PropCodice      =   $"frmTrovaIntervento.frx":0000
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmTrovaIntervento.frx":004D
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmTrovaIntervento.frx":0091
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
      Begin DmtCodDescCtl.DmtCodDesc CDAnagrafica 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1085
         PropCodice      =   $"frmTrovaIntervento.frx":00EB
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"frmTrovaIntervento.frx":012F
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"frmTrovaIntervento.frx":0183
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
      Begin DMTDataCmb.DMTCombo cboCategoriaIntervento 
         Height          =   315
         Left            =   8160
         TabIndex        =   4
         Top             =   480
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
      Begin DMTDataCmb.DMTCombo CboTipoAddebitoIntervento 
         Height          =   315
         Left            =   10440
         TabIndex        =   5
         Top             =   480
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
      Begin DMTDataCmb.DMTCombo cboStatoIntervento 
         Height          =   315
         Left            =   5880
         TabIndex        =   6
         Top             =   480
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
      Begin VB.Label Label4 
         Caption         =   "Categoria commessa"
         Height          =   255
         Index           =   2
         Left            =   8160
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo addebito commessa"
         Height          =   255
         Index           =   1
         Left            =   10440
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Stato commessa"
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmTrovaIntervento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub INIT_CONTROLLI()
Dim sSQL As String
    
    sSQL = "SELECT IDRV_PO08_StatoIntervento, StatoIntervento"
    sSQL = sSQL & " FROM RV_PO08_StatoIntervento "
    sSQL = sSQL & " ORDER BY StatoIntervento"
    

    With Me.cboStatoIntervento
    Set .Database = TheApp.Database.Connection
    .DisplayField = "StatoIntervento"
    .AddFieldKey "IDRV_PO08_StatoIntervento"
    .Sql = sSQL
    .Refresh
    End With

    sSQL = "SELECT IDRV_PO08_TipoAddebitoIntervento, TipoAddebitoIntervento"
    sSQL = sSQL & " FROM RV_PO08_TipoAddebitoIntervento"
    sSQL = sSQL & " ORDER BY TipoAddebitoIntervento"

    With Me.CboTipoAddebitoIntervento
    Set .Database = TheApp.Database.Connection
    .DisplayField = "TipoAddebitoIntervento"
    .AddFieldKey "IDRV_PO08_TipoAddebitoIntervento"
    .Sql = sSQL
    .Refresh
    End With
    
    
     'Categoria intervento
    With Me.cboCategoriaIntervento
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_PO08_CategoriaChiamata"
        .DisplayField = "CategoriaChiamata"
        .Sql = "SELECT * FROM RV_PO08_CategoriaChiamata"
        .Fill
    End With

    
    
   'Parco natanti
    With Me.CDParcoNatanti
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hWnd
        .CodeField = "NomeNatante"
        .DescriptionField = "Targa"
        .KeyField = "IDRV_PO08_ParcoNatanti"
        .TableName = "RV_PO08_ParcoNatanti"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Nome natante"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Targa"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Nome natante"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Targa"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("RV_PO08_ParcoNatanti")
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With


   'Proprietario
    With Me.CDAnagrafica
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hWnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Ragione sociale"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Ragione sociale"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 29 'TestataDescrizioneAggiuntiva
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
End Sub

Private Sub cmdAssociaPrimaDisponibile_Click()
Dim Testo As String
Dim sSQL As String
    
    
    Testo = "Questo comando esegue le seguenti operazioni: " & vbCrLf
    Testo = Testo & "    1° Associa questa attività ad una commessa con lo stato uguale a quello dei parametri azienda" & vbCrLf
    Testo = Testo & "    2° Se non esiste una commessa del tipo descritto nel punto 1 associa questa attività ad una nuova commessa"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Associazione commessa") = vbNo Then Exit Sub

    sSQL = "UPDATE RV_PO08_TMPPassaggioAttivita SET "
    sSQL = sSQL & "TipoAssociazione=1" & ", "
    sSQL = sSQL & "IDRV_PO08_InterventoAssociato=" & 0 & ", "
    sSQL = sSQL & "NumeroDocumentoAssociato=" & 0 & ", "
    sSQL = sSQL & "DataDocumentoAssociato=" & fnNormDate("")
    sSQL = sSQL & " WHERE IDRV_PO08_TMPPassaggioAttivita=" & FrmInizio.GridDisponibili("IDRV_PO08_TMPPassaggioAttivita").Value
    
    CnDMT.Execute sSQL
    
    Unload Me
End Sub

Private Sub cmdNuovaScheda_Click()
Dim Testo As String
Dim sSQL As String
    
    
    Testo = "Questo comando esegue le seguenti operazioni: " & vbCrLf
    Testo = Testo & "    1° Associa questa attività ad una nuova commessa"
    
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Associazione commessa") = vbNo Then Exit Sub
    
    
    sSQL = "UPDATE RV_PO08_TMPPassaggioAttivita SET "
    sSQL = sSQL & "TIpoAssociazione=2" & ", "
    sSQL = sSQL & "IDRV_PO08_InterventoAssociato=" & 0 & ", "
    sSQL = sSQL & "NumeroDocumentoAssociato=" & 0 & ", "
    sSQL = sSQL & "DataDocumentoAssociato=" & fnNormDate("")
    sSQL = sSQL & " WHERE IDRV_PO08_TMPPassaggioAttivita=" & FrmInizio.GridDisponibili("IDRV_PO08_TMPPassaggioAttivita").Value
    
    CnDMT.Execute sSQL


    
    Unload Me
End Sub

Private Sub cmdRicerca_Click()
    GET_GRIGLIA_ATTIVITA
End Sub

Private Sub Form_Load()
INIT_CONTROLLI


Me.CDAnagrafica.Load FrmInizio.GridDisponibili("IDAnagrafica").Value
Me.CDParcoNatanti.Load FrmInizio.GridDisponibili("IDRV_PO08_ParcoNatanti")
Me.CDAnagrafica.Enabled = False
Me.CDParcoNatanti.Enabled = False
Me.cboStatoIntervento.WriteOn GET_TIPO_ARTICOLO("IDStatoInterventoPerRicerca")


GET_GRIGLIA_ATTIVITA


End Sub
Public Function GET_TIPO_ARTICOLO(NomeCampo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_PO08_ParametriAzienda "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_ARTICOLO = 0
Else
    GET_TIPO_ARTICOLO = fnNotNullN(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub GET_GRIGLIA_ATTIVITA()
'On Error GoTo ERR_fnGrigliaAssegnazione
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader



    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3

    sSQL = "SELECT * FROM RV_PO08_IEIntervento "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDRV_PO08_TipoIntervento=1"
    sSQL = sSQL & " AND IDRV_PO08_ParcoNatanti=" & Me.CDParcoNatanti.KeyFieldID
    sSQL = sSQL & " AND IDAnagrafica=" & Me.CDAnagrafica.KeyFieldID
    If Me.cboStatoIntervento.CurrentID > 0 Then
        sSQL = sSQL & " AND IDRV_PO08_StatoIntervento=" & Me.cboStatoIntervento.CurrentID
    End If
    If Me.CboTipoAddebitoIntervento.CurrentID > 0 Then
        sSQL = sSQL & " AND IDRV_PO08_TipoAddebito=" & Me.CboTipoAddebitoIntervento.CurrentID
    End If
    If Me.cboCategoriaIntervento.CurrentID > 0 Then
        sSQL = sSQL & " AND IDRV_PO08_CategoriaChiamata=" & Me.cboCategoriaIntervento.CurrentID
    End If
    
    
    sSQL = sSQL & " ORDER BY DataDocumento DESC "
    
    Set rsGriglia = New ADODB.Recordset
    rsGriglia.CursorLocation = adUseClient
    rsGriglia.Open sSQL, CnDMT.InternalConnection
    
        With Me.Griglia
            .EnableMove = True
            .UpdatePosition = True
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear


            With Me.Griglia.ColumnsHeader
                
                .Add "IDRV_PO08_Intervento", "IDRV_PO08_Intervento", dgInteger, False, 500, dgAlignleft
                .Add "IDAzienda", "IDAzienda", dgInteger, False, 500, dgAlignleft
                .Add "IDUtente", "IDUtente", dgInteger, False, 500, dgAlignleft
                
                .Add "NumeroDocumento", "N° commessa", dgNumeric, True, 1250, dgAlignleft
                .Add "DataDocumento", "Data commessa", dgDate, True, 1500, dgAlignlef
                .Add "Anagrafica", "Anagrafica", dgchar, True, 2500, dgAlignleft
                .Add "NomeNatante", "Nome imbarcazione", dgchar, True, 2500, dgAlignleft
                .Add "StatoIntervento", "Stato intervento", dgDate, True, 1500, dgAlignlef
                .Add "TipoAddebitoIntervento", "Tipo addebito", dgchar, True, 2500, dgAlignleft
                .Add "CategoriaChiamata", "Categoria", dgchar, True, 2500, dgAlignleft
                
            End With
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaAssegnazione:
    MsgBox Err.Description, vbCritical, "Reperimento dati assegnazione"
End Sub


Private Sub Griglia_DblClick()
Dim Testo As String
Dim sSQL As String
    
    If fnNotNullN(FrmInizio.GridDisponibili("IDRV_PO08_InterventoAssociato").Value) <> fnNotNullN(Me.Griglia("IDRV_PO08_Intervento").Value) Then
    
        Testo = "Sei sicuro di voler cambiare l'associazione alla scheda di lavorazione?"
        
        If MsgBox(Testo, vbQuestion + vbYesNo, "Associazione scheda di lavorazione") = vbNo Then Exit Sub
        
        
        sSQL = "UPDATE RV_PO08_TMPPassaggioAttivita SET "
        sSQL = sSQL & "TIpoAssociazione=0" & ", "
        sSQL = sSQL & "IDRV_PO08_InterventoAssociato=" & fnNotNullN(Me.Griglia("IDRV_PO08_Intervento").Value) & ", "
        sSQL = sSQL & "NumeroDocumentoAssociato=" & fnNotNullN(Me.Griglia("NumeroDocumento").Value) & ", "
        sSQL = sSQL & "DataDocumentoAssociato=" & fnNormDate(Me.Griglia("DataDocumento").Value)
        sSQL = sSQL & " WHERE IDRV_PO08_TMPPassaggioAttivita=" & FrmInizio.GridDisponibili("IDRV_PO08_TMPPassaggioAttivita").Value
        
        CnDMT.Execute sSQL

    End If
    
    Unload Me
End Sub
