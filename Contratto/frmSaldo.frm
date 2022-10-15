VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmSaldo 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Saldo contratto"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
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
   ScaleHeight     =   8190
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Opzioni"
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
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Width           =   4455
      Begin VB.CheckBox chkStampa 
         Alignment       =   1  'Right Justify
         Caption         =   "Stampa immediata del documento"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   4215
      End
      Begin VB.ComboBox cboStampante 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1560
         Width           =   4215
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroCopie 
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   840
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecimalPlaces   =   0
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Numero di copie per stampa"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Stampante"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   2415
      End
   End
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
      Left            =   120
      TabIndex        =   16
      Top             =   7560
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo di documento"
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
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CheckBox chkEntePubblico 
         Caption         =   "Ente pubblico"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   4200
         Width           =   4215
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroDocumento 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   2160
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTDataCmb.DMTCombo CboPagamento 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   3600
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
      Begin DMTDataCmb.DMTCombo CboMagazzino 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   2880
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
      Begin DMTDATETIMELib.dmtDate txtDataDoc 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   2160
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
         Left            =   3120
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
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
      Begin DMTDataCmb.DMTCombo CboSezionale 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   960
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
      Begin DMTDataCmb.DMTCombo CboTipoOggetto 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   240
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
      Begin DMTDataCmb.DMTCombo cboSezionalePA 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1560
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
      Begin VB.Label Label1 
         Caption         =   "Sezionale"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Valuta"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   14
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Data documento"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Magazzino"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Pagamento di Default"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Sezionale per fatturazione PA"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "N° Doc."
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   9
         Top             =   1920
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ObjDoc As DmtDocs.cDocument
Private oReport As dmtReportLib.dmtReport

'Variabile che contiene il numero documento creato
Private VarNumeroDoc As String

Private Const NOMETABELLAPIANA = "ValoriOggettoPerTipo"
Private Const NOMETABELLADETTAGLIO = "ValoriOggettoDettaglio"
Private LINK_IVA_CLIENTE As Long


Private Sub CboTipoOggetto_Click()
On Error GoTo ERR_CboTipoOggetto_Click
    With Me.CboSezionale

        Set .Database = TheApp.Database.Connection

        .AddFieldKey "IDSezionale"

        .DisplayField = "Sezionale"

        .SQL = "SELECT Sezionale.IDSezionale, Sezionale.Sezionale "
        .SQL = .SQL & "FROM RegistroIvaPerTipoOggetto INNER JOIN "
        .SQL = .SQL & "Sezionale ON RegistroIvaPerTipoOggetto.IDRegistroIva = Sezionale.IDRegistroIva AND "
        .SQL = .SQL & "RegistroIvaPerTipoOggetto.IDFiliale = Sezionale.IDFiliale LEFT OUTER JOIN "
        .SQL = .SQL & "TipoOggetto ON RegistroIvaPerTipoOggetto.IDTipoOggetto = TipoOggetto.IDTipoOggetto "
        .SQL = .SQL & "WHERE RegistroIvaPerTipoOggetto.IDTipoOggetto = " & Me.CboTipoOggetto.CurrentID
        .SQL = .SQL & " AND RegistroIvaPerTipoOggetto.IDFiliale = " & TheApp.Branch
        
    End With
    
    GET_DEFAULT_SEZ_TIPO_OGGETTO Me.CboTipoOggetto.CurrentID
    
Exit Sub
ERR_CboTipoOggetto_Click:
    MsgBox Err.Description, vbCritical, "CboTipoOggetto_Click"
End Sub

Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click
If Permesso = False Then Exit Sub


Settaggio Me.chkEntePubblico.Value

fncTestata IDClienteFatturazione, frmMain.cboSitoPerAnagrafica.CurrentID, frmMain.chkRitAcconto.Value

fncRighe

If InserimentoDMT = True Then

    CREA_FLUSSO_DOCUMENTALE ObjDoc.IDTipoOggetto, ObjDoc.IDOggetto, LINK_OGGETTO_CONTRATTO_SEL, LINK_TIPO_OGGETTO_CONTRATTO_SEL, "Contratto -> Saldo"
    
    OPERAZIONE_ESEGUITA_ACCONTO = 1
    
    If chkStampa.Value = 1 Then
        DoEvents
        ObjDoc.Prepare2Print TheApp.IDFirm, TheApp.IDUser, ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto
        StampaDocumento
        DoEvents
    End If
    
    
    Unload Me

End If

Exit Sub

ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "cmdConferma_Click"
End Sub

Private Sub Settaggio(EntePubblico As Integer)
On Error GoTo ERR_Settaggio
    Set ObjDoc = New cDocument

    
    With ObjDoc
        Set .Connection = Cn
        .IDAzienda = TheApp.IDFirm
        .IDAttivitaAzienda = GetAttivitaAzienda(TheApp.IDFirm, TheApp.Branch)
        .IDFiliale = TheApp.Branch
        .SetTipoOggetto Me.CboTipoOggetto.CurrentID
        .IDFunzione = GET_LINK_FUNZIONE(Me.CboTipoOggetto.CurrentID)
        .UseAutomation = True
        .IDEsercizio = GET_LINK_ESERCIZIO(Me.txtDataDoc.Text)
        If EntePubblico = 0 Then
            .IDSezionale = Me.CboSezionale.CurrentID
        Else
            If Me.cboSezionalePA.CurrentID = 0 Then
                .IDSezionale = Me.CboSezionale.CurrentID
            Else
                .IDSezionale = Me.cboSezionalePA.CurrentID
            End If
        End If
        .IDTipoAnagrafica = 2
        .IDUtente = TheApp.IDUser
        .Descrizione = GET_DESCRIZIONE_FUNZIONE(Me.CboTipoOggetto.CurrentID)
        .DataEmissione = Date
        .Numero = Me.txtNumeroDocumento.Value
        If .Tables.Count = 0 Then
        'Se Tables.Count = 0 vuol dire che l'oggetto
        'DmtDocs non è mai stato inizializzato
            .Clear
            .SetTipoOggetto Me.CboTipoOggetto.CurrentID
        Else
            .ClearValues
        End If
    
    End With

Exit Sub
ERR_Settaggio:
    MsgBox Err.Description, vbCritical, "Settaggio"
End Sub

Private Function fncTestata(IDAnagrafica As Long, IDSitoPerAnagrafica As Long, RitenutaAcconto As Long) As Boolean
On Error GoTo ERR_fncTestata
Dim IDLetteraIntento As Long

With ObjDoc.Tables

'Imposta la riga attiva per la tabella di testata
    
    ObjDoc.Tables(NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)).SetActiveRetail 1
    'Dati generici del documento
    .Field "Link_Val_cambio", Null, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Doc_data", ObjDoc.DataEmissione, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Doc_numero", ObjDoc.Numero, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Link_Doc_sezionale", ObjDoc.IDSezionale, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Link_Doc_magazzino", Me.CboMagazzino.CurrentID, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)

    'TrovaAnagrafica IDRata
    ObjDoc.ReadDataFromCliFo IDAnagrafica, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.ReadDataFromCliFoSite IDSitoPerAnagrafica, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        
    ObjDoc.ReadDataFromPayment Me.CboPagamento.CurrentID, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Doc_perc_rit_acc", ObjDoc.DBDefaults.PercentualeRitenutaAcconto, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Nom_calcola_rit_acc", RitenutaAcconto, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
    .Field "Doc_prezzi_lordo_IVA", 1, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    .Field "Doc_spese_lordo_IVA", 1, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
    If IDRaggrFattContratto > 0 Then
        ObjDoc.Field "Link_Nom_raggrup_fatturato", IDRaggrFattContratto, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    If IDAccordoCommerciale > 0 Then
        ObjDoc.Field "Link_Nom_accordi_commerciali", IDAccordoCommerciale, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    If IDContrattoBancario > 0 Then
        ObjDoc.Field "Link_Nom_contratto_bancario", IDContrattoBancario, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    

    If fnNotNullN(.Field("Link_Doc_pagamento", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) = 0 Then
        .Field "Link_Doc_pagamento", Me.CboPagamento.CurrentID, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    
    
    If fnNotNullN(.Field("Link_Val_valuta", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) = 0 Then
        .Field "Link_Val_valuta", 9, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    
End With

fncTestata = True
     
Exit Function
ERR_fncTestata:
MsgBox Err.Description, vbCritical, "ERR_fncTestata"

End Function
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


Private Function fncRighe() As Boolean
'On Error GoTo ERR_fncRighe
VARErroreFunzione = "fncRighe"

Dim I As Integer
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoUnitarioLordo As Double



I = 1

    sSQL = "SELECT * FROM RV_POContrattoProdotti "
    sSQL = sSQL & "WHERE IDRV_POContratto=" & Link_Contratto
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
    
        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
        
        With ObjDoc.Tables
            ObjDoc.ReadDataFromArticle fnNotNullN(rs!IDArticolo)
            
            If (Len(Trim(fnNotNull(rs!Annotazioni))) > 0) Then
                If (fnNotNullN(rs!IDRV_POProdotto) = 0) Then
                    .Field "Art_descrizione", Mid(fnNotNull(rs!Annotazioni), 1, 255), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                End If
            End If
            
            ImportoUnitarioLordo = fnNotNullN(rs!TotaleRiga) / (fnNotNullN(rs!QuantitaArticolo) * fnNotNullN(rs!QuantitaPeriodo))
            
            If (fnNotNullN(rs!ACorpo)) = 0 Then
                .Field "Art_quantita_totale", fnNotNullN(rs!QuantitaArticolo) * fnNotNullN(rs!QuantitaPeriodo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            Else
                .Field "Art_quantita_totale", fnNotNullN(rs!QuantitaArticolo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            End If
            
            .Field "Art_prezzo_unitario_neutro", ImportoUnitarioLordo, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Link_Art_IVA", fnNotNullN(rs!IDIva), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_aliquota_IVA", fnNotNullN(rs!AliquotaIva), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
            .Field "Art_data_inizio_competenza", fnNotNull(rs!DataInizioPeriodo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_data_fine_competenza", fnNotNull(rs!DataFinePeriodo), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
           
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
            I = I + 1
            If (fnNotNullN(rs!IDRV_POProdotto) > 0) Then
                If (Len(Trim(fnNotNull(rs!Annotazioni))) > 0) Then
                    ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                    
                    .Field "Art_descrizione", Mid(fnNotNull(rs!Annotazioni), 1, 255), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                    .Field "Art_quantita_totale", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                    .Field "Art_prezzo_unitario_neutro", 0, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    
                    .Field "Link_Art_IVA", fnNotNullN(rs!IDIva), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                    .Field "Art_aliquota_IVA", fnNotNullN(rs!AliquotaIva), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                               
                    I = I + 1
                End If
            End If
        
        End With
        
    rs.MoveNext
    
    Wend
       
    rs.CloseResultset
    Set rs = Nothing
    
    
    
    sSQL = "SELECT * FROM RV_POIEContrattoAcconti "
    sSQL = sSQL & "WHERE IDOggettoCollegato=" & LINK_OGGETTO_CONTRATTO_SEL
    sSQL = sSQL & " AND IDTipoOggettoCollegato = " & LINK_TIPO_OGGETTO_CONTRATTO_SEL
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
    
        ObjDoc.Tables(NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
    
        With ObjDoc.Tables
            'ObjDoc.ReadDataFromArticle fnNotNullN(rs!IDArticolo)
            
            .Field "Art_descrizione", "Rif. fattura di acconto n. " & fnNotNullN(rs!Doc_Numero) & " del " & fnNotNull(rs!Doc_data), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_quantita_totale", fnNotNullN(rs!Art_quantita_totale), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_prezzo_unitario_neutro", -fnNotNullN(rs!Art_prezzo_unitario_neutro), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            
            If fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
                .Field "Link_art_IVA", fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA(.Field("Link_art_IVA", , NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)))
            Else
                .Field "Link_Art_IVA", fnNotNullN(rs!Link_Art_IVA), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                .Field "Art_aliquota_IVA", fnNotNullN(rs!Art_aliquota_IVA), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            End If
                
            .Field "Art_data_inizio_competenza", fnNotNull(rs!Art_data_inizio_competenza), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            .Field "Art_data_fine_competenza", fnNotNull(rs!Art_data_fine_competenza), NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ObjDoc.AppendDettaglioFlusso2Detail fnNotNullN(rs!ID_Art_dettaglio_order), fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), I
        
        End With
                
        
        I = I + 1
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    
fncRighe = True
Exit Function
ERR_fncRighe:
    fncRighe = False


    MsgBox Err.Description, vbCritical, "ERR_fncRighe"
    
End Function
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

Private Function InserimentoDMT() As Boolean
On Error GoTo ERR_InserimentoDMT
VARErroreFunzione = "InserimentoDMT"


Screen.MousePointer = vbHourglass
    
    Set ObjDoc.Scadenze = Nothing
    ObjDoc.PerformDocument Nothing
    
    
    
    VarNumeroDoc = ObjDoc.Insert
    
    If VarNumeroDoc > 0 Then
        InserimentoDMT = True
    Else
        InserimentoDMT = False
    End If
        
    
    
Screen.MousePointer = vbDefault
    
Exit Function

ERR_InserimentoDMT:
    InserimentoDMT = False
    MsgBox Err.Description, vbCritical, "Creazione fattura"
End Function


Private Sub CREA_FLUSSO_DOCUMENTALE(IDTipoOggettoVend As Long, IDOggettoVend As Long, IDOggettoRata As Long, IDTipoOggettoRata As Long, DescrizioneFunzione As String)
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
'''''''''''''''''''''''''''''''''GRUPPO FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM FlussoGruppo "
sSQL = sSQL & "WHERE Descrizione=" & fnNormString(DescrizioneFunzione)
Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDFlussoGruppo = fnGetNewKeyTipoOggetto("FlussoGruppo", "IDFlussoGruppo")
        rsNew!Descrizione = DescrizioneFunzione
    rsNew.Update
End If

IDFlussoGruppo = fnNotNullN(rsNew!IDFlussoGruppo)

rsNew.Close
Set rsNew = Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''FLUSSO FUNZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Function GET_LINK_ESERCIZIO(DataDocumento As String) As Long
    
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    
    GET_LINK_ESERCIZIO = 0
    
    sSQL = "Select IDEsercizio, Esercizio"
    sSQL = sSQL & " FROM Esercizio"
    sSQL = sSQL & " WHERE IDAzienda = " & TheApp.IDFirm
    sSQL = sSQL & " AND DataInizio<=" & fnNormDate(DataDocumento)
    sSQL = sSQL & " AND DataFine>=" & fnNormDate(DataDocumento)
    
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        GET_LINK_ESERCIZIO = fnNotNullN(rs!IDEsercizio)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function


Private Sub INIT_CONTROLLI()
Dim sSQL As String


    sSQL = "SELECT IDTipoOggetto, Oggetto"
    sSQL = sSQL & " FROM TipoOggetto"
    sSQL = sSQL & " WHERE ((IDGestore=15) OR (IDGestore=237))"
    sSQL = sSQL & " ORDER BY Oggetto"
    

    With Me.CboTipoOggetto
        Set .Database = Cn
        .DisplayField = "Oggetto"
        .AddFieldKey "IDTipoOggetto"
        .SQL = sSQL
        .Refresh
    End With

    With Me.CboSezionale
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .SQL = "SELECT Sezionale.IDSezionale, Sezionale.Sezionale "
        .SQL = .SQL & "FROM RegistroIvaPerTipoOggetto INNER JOIN "
        .SQL = .SQL & "Sezionale ON RegistroIvaPerTipoOggetto.IDRegistroIva = Sezionale.IDRegistroIva AND "
        .SQL = .SQL & "RegistroIvaPerTipoOggetto.IDFiliale = Sezionale.IDFiliale LEFT OUTER JOIN "
        .SQL = .SQL & "TipoOggetto ON RegistroIvaPerTipoOggetto.IDTipoOggetto = TipoOggetto.IDTipoOggetto "
        .SQL = .SQL & "WHERE RegistroIvaPerTipoOggetto.IDTipoOggetto = " & 339
        .SQL = .SQL & " AND RegistroIvaPerTipoOggetto.IDFiliale = " & TheApp.Branch
        .Fill
    End With
    

    
    'SEZIONALE PA
    sSQL = "SELECT IDSezionale, Sezionale"
    sSQL = sSQL & " FROM Sezionale"
    sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch
    sSQL = sSQL & " AND FatturaElettronica=" & fnNormBoolean(1)
    sSQL = sSQL & " AND IDRegistroIva=1"
    sSQL = sSQL & " ORDER BY Sezionale"
    
    With Me.cboSezionalePA
        Set .Database = Cn
        .DisplayField = "Sezionale"
        .AddFieldKey "IDSezionale"
        .SQL = sSQL
        .Refresh
    End With
    
    
    sSQL = "SELECT IDValuta, Valuta"
    sSQL = sSQL & " FROM Valuta"
    'sSQL = sSQL & " WHERE ((IDFiliale=" & theapp.branch & ") AND (IDRegistroIva = 1))"
    sSQL = sSQL & " ORDER BY Valuta"

    With Me.CboValuta
        Set .Database = Cn
        .DisplayField = "Valuta"
        .AddFieldKey "IDValuta"
        .SQL = sSQL
        .Refresh
    End With
    
    
    sSQL = "SELECT IDPagamento, Pagamento"
    sSQL = sSQL & " FROM Pagamento"
    sSQL = sSQL & " ORDER BY Pagamento"

    With Me.CboPagamento
        Set .Database = Cn
        .DisplayField = "Pagamento"
        .AddFieldKey "IDPagamento"
        .SQL = sSQL
        .Refresh
    End With
    
    sSQL = "SELECT IDMagazzino, Magazzino"
    sSQL = sSQL & " FROM Magazzino"
    sSQL = sSQL & " WHERE IDAzienda = " & TheApp.IDFirm
    sSQL = sSQL & " ORDER BY Magazzino"

    With Me.CboMagazzino
        Set .Database = Cn
        .DisplayField = "Magazzino"
        .AddFieldKey "IDMagazzino"
        .SQL = sSQL
        .Refresh
    End With
    
    
    
    GET_DEFAULT_TIPOOGGETTO IDClienteFatturazione
    
    Me.CboValuta.WriteOn 9
    Me.txtDataDoc.Value = Date
    
    GET_DEFAULT_SEZ_TIPO_OGGETTO Me.CboTipoOggetto.CurrentID
    
    GET_DEFAULT_MAGAZZINO
    
    GET_DEFAULT_PAGAMENTO IDClienteFatturazione
    
    Me.chkEntePubblico.Value = GET_ENTE_PUBBLICO(IDClienteFatturazione)
        
        
End Sub
Private Sub GET_DEFAULT_PAGAMENTO(IDAnagrafica As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDPagamento As Long

IDPagamento = 0

sSQL = "SELECT IDPagamentoDefault "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    IDPagamento = fnNotNullN(rs!IDPagamentoDefault)
End If

rs.CloseResultset
Set rs = Nothing

If IDPagamento = 0 Then

    sSQL = "SELECT IDPagamentoDocDefault "
    sSQL = sSQL & "FROM PersonalizzazionePerFiliale "
    sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        IDPagamento = fnNotNullN(rs!IDPagamentoDocDefault)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If

Me.CboPagamento.WriteOn IDPagamento


End Sub
Private Function GET_ENTE_PUBBLICO(IDAnagrafica As Long) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    Dim EntePubblico As Long
    
    EntePubblico = 0
    
    sSQL = "SELECT EntePubblico "
    sSQL = sSQL & "FROM Anagrafica "
    sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        EntePubblico = fnNotNullN(rs!EntePubblico)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    GET_ENTE_PUBBLICO = Abs(EntePubblico)
    
End Function
Private Sub GET_DEFAULT_SEZ_TIPO_OGGETTO(IDTipoOggetto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

    sSQL = "SELECT IDSezionale "
    sSQL = sSQL & "FROM DefaultFilialePerTipoOggetto "
    sSQL = sSQL & "WHERE (IDTipoOggetto = " & IDTipoOggetto & ") And (IDSezionale > 0) And (IDFiliale = " & TheApp.Branch & ")"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        Me.CboSezionale.WriteOn fnNotNullN(rs!IDSezionale)
    Else
        Me.CboSezionale.WriteOn 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing

End Sub

Private Function Permesso() As Boolean
Dim Testo As String

Permesso = False

If Me.CboTipoOggetto.CurrentID = 0 Then
    MsgBox "Selezionare il tipo di documento", vbInformation, "Controllo dati"
    Exit Function
End If

If Me.CboSezionale.CurrentID = 0 Then
    MsgBox "Selezionare il sezionale", vbInformation, "Controllo dati"
    Exit Function
End If

If Me.chkEntePubblico.Value = vbChecked Then
    If Me.cboSezionalePA.CurrentID = 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Il sezionale per la fatturazione elettronica non è stato selezionato " & vbCrLf
        Testo = Testo & "Se si continua con questo comando verrà utilizzato il sezionale predefinito del tipo documento selezionato " & vbCrLf
        Testo = Testo & "Vuoi continuare?"
    
        If MsgBox(Testo, vbInformation, "Controllo dati") = vbNo Then Exit Function
        Exit Function
    End If
End If

If Me.CboMagazzino.CurrentID = 0 Then
    MsgBox "Selezionare il magazzino", vbInformation, "Controllo dati"
    Exit Function
End If

If Me.CboValuta.CurrentID = 0 Then
    MsgBox "Selezionare la valuta", vbInformation, "Controllo dati"
    Exit Function
End If

If Me.CboPagamento.CurrentID = 0 Then
    MsgBox "Selezionare il pagamento", vbInformation, "Controllo dati"
    Exit Function
End If


Permesso = True


End Function


Private Sub Form_Load()
    INIT_CONTROLLI
    fncStampantiPedana
    Me.chkStampa.Value = vbChecked
    Me.txtNumeroCopie.Value = 1
    
End Sub
Private Sub fncStampantiPedana()
Dim prn As Printer

Me.cboStampante.Clear

For Each prn In Printers
    Me.cboStampante.AddItem prn.DeviceName
Next

End Sub
Private Sub GET_DEFAULT_TIPOOGGETTO(IDAnagrafica As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDTipoDocumento As Long

IDTipoDocumento = 0

sSQL = "SELECT IDTipoOggettoDocEvasione "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    IDTipoDocumento = fnNotNullN(rs!IDTipoOggettoDocEvasione)
End If

rs.CloseResultset
Set rs = Nothing

If IDTipoDocumento = 0 Then

    sSQL = "SELECT IDTipoOggettoDocEvasione "
    sSQL = sSQL & "FROM PersonalizzazionePerFiliale "
    sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
       IDTipoDocumento = fnNotNullN(rs!IDTipoOggettoDocEvasione)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If

Me.CboTipoOggetto.WriteOn IDTipoDocumento


End Sub
Private Sub GET_DEFAULT_MAGAZZINO()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDTipoDocumento As Long

IDMagazzino = 0

If IDMagazzino = 0 Then

    sSQL = "SELECT IDMagazzino "
    sSQL = sSQL & "FROM PersonalizzazionePerFiliale "
    sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
       IDMagazzino = fnNotNullN(rs!IDMagazzino)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If

Me.CboMagazzino.WriteOn IDMagazzino


End Sub
Private Function GET_DESCRIZIONE_FUNZIONE(IDTipoOggetto As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDFunzione, Funzione FROM Funzione "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_FUNZIONE = 0
Else
    GET_DESCRIZIONE_FUNZIONE = fnNotNull(rs!Funzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Public Sub StampaDocumento()
On Error GoTo ERR_StampaDocumento

Dim IDReport As Long

        Set oReport = New dmtReportLib.dmtReport
            Set oReport.Connection = Cn
            
            If MenuOptions.DBType = 1 Then
                'parametri di accesso al database ACCESS
                oReport.Password = "dmt192981046"
                oReport.User = "admin"
            Else
                'parametri di accesso al database SQL Server
                oReport.Password = TheApp.Password
                oReport.User = TheApp.User
            End If
        
        
        'Imposta l'idfiliale di appartenenza del documento da stampare
            oReport.BranchID = TheApp.Branch 'IDFiliale
        'Imposta l'identificativo del tipo di documento
            oReport.DocTypeID = ObjDoc.IDTipoOggetto
            'oReport.Where = "IDOggetto = 873" '& Val(Me.Txt_Reg_IDRegistro)
            oReport.Where = "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto) & ".IDOggetto = " & ObjDoc.IDOggetto
            oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
            
            oReport.Copies = txtNumeroCopie.Text
            
            If (Len(Me.cboStampante.Text)) = 0 Then
                oReport.DoPrint Printer.DeviceName
            Else
                oReport.DoPrint Me.cboStampante.Text
            End If
   
Exit Sub
ERR_StampaDocumento:
    MsgBox Err.Description, vbCritical, "Stampa Documento"
End Sub


