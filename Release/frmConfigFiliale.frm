VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Begin VB.Form frmConfigFiliale 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Configurazione parametri filiale"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9555
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
   ScaleHeight     =   6810
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   4335
   End
   Begin VB.TextBox txtControllo 
      Height          =   6615
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   4935
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin DMTDataCmb.DMTCombo cboFiliale 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
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
   Begin DMTDataCmb.DMTCombo cboMagLav 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
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
   Begin DMTDataCmb.DMTCombo cboTipoImpostazione 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   4335
      _ExtentX        =   7646
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
   Begin VB.Label Label1 
      Caption         =   "Tipo impostazione contratto"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Magazzino principale "
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Filiale"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
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
      TabIndex        =   4
      Top             =   6240
      Width           =   4335
   End
End
Attribute VB_Name = "frmConfigFiliale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private cnConfig As ADODB.Connection
Private raParametri As ADODB.Recordset
Private LINK_PARAMETRO_FILIALE As Long
Private LINK_PARAMETRO_FILIALE_AZIENDA As Long
Private LINK_PARAMETRO_FILIALE_FARMACIA As Long
Private LINK_PARAMETRO_FILIALE_LIQ As Long

Private LINK_CLIENTE_ORDINE_PRED As Long

Private NUMERO_DOCUMENTO_ORDINE As Long

Private LINK_CLIENTE_LAV_IVGAMMA As Long
Private LINK_ORDINE_LAV_IV_GAMMA As Long
Private LINK_ORDINE_MERCE_GIACENZA As Long

Private LINK_TIPO_PEDANA_PREDEFINITA As Long

Private LINK_DEMO_CLIENTE As Long
Private LINK_DEMO_SOCIO As Long
Private LINK_DEMO_FORNITORE As Long

Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click
Dim sSQL As String
Dim NumeroOrdineCliente As String

If Me.cboFiliale.CurrentID = 0 Then Exit Sub
If Me.cboMagLav.CurrentID = 0 Then Exit Sub
If Me.cboTipoImpostazione.CurrentID = 0 Then Exit Sub

'GENERALE IN DMT''''''''''''''''''''''''''''''''''''''
CREAZIONE_UNITA_DI_MISURA
CREA_TIPO_ANAGRAFICA
CREAZIONE_SEZIONALI VarIDFiliale, VarIDAzienda
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'TABELLE DEI CONTRATTI''''''''''''''''''''''''''''''
CREAZIONE_TIPO_DURATA
CREAZIONE_TIPO_RATEIZZAZIONE
CREAZIONE_TIPO_DURATA_RINNOVO
CREAZIONE_TIPO_DURATA_ASSITENZA
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'TABELLE DEGLI INTERVENTI'''''''''''''''''''''''''''''
CREAZIONE_TIPO_INTERVENTO
''''''''''''''''''''''''''''''''''''''''''''''''''''''

CREAZIONE_CAUSALI_MAGAZZINO

INSERIMENTO_ARTICOLO VarIDAzienda, VarIDFiliale

GET_LINK_ANAGRAFICA_CLIENTE_DEMO "Cliente Demo", VarIDAzienda, 2, "Cliente"
GET_LINK_ANAGRAFICA_CLIENTE_DEMO "Tecnico Riferimento Demo", VarIDAzienda, 11, "AnagraficaTipo1"
GET_LINK_ANAGRAFICA_CLIENTE_DEMO "Tecnico operativo Demo", VarIDAzienda, 12, "AnagraficaTipo2"
GET_LINK_ANAGRAFICA_CLIENTE_DEMO "Amministratore Demo", VarIDAzienda, 13, "AnagraficaTipo3"
GET_LINK_ANAGRAFICA_CLIENTE_DEMO "Installatore Demo", VarIDAzienda, 14, "AnagraficaTipo4"

LINK_PARAMETRO_FILIALE = GET_LINK_PARAMETRO_FILIALE(Me.cboFiliale.CurrentID, VarIDAzienda)

Me.lblInfo.Caption = "OPERAZIONE COMPLETATA"

MsgBox "OPERAZIONE COMPLETATA", vbInformation, "Configurazione iniziale della filiale"

Unload Me

Exit Sub

ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "ERR_cmdConferma_Click"
End Sub
Private Function CONNESSIONE_CONFIGURAZIONE() As Boolean
On Error GoTo ERR_CONNESSIONE_CONFIGURAZIONE

    CONNESSIONE_CONFIGURAZIONE = False
    
    If Not (cnConfig Is Nothing) Then
        cnConfig.Close
        Set cnConfig = Nothing
    End If

    Set cnConfig = New ADODB.Connection
    
    cnConfig.ConnectionString = "Provider=Microsoft.Jet.Oledb.4.0;Data source=" & App.Path & "\ConfigLiveSolution.mdb"
    cnConfig.Open

    CONNESSIONE_CONFIGURAZIONE = True
    
Exit Function

ERR_CONNESSIONE_CONFIGURAZIONE:
    MsgBox Err.Description, vbCritical, "CONNESSIONE_CONFIGURAZIONE"
    CONNESSIONE_CONFIGURAZIONE = False
End Function
Private Sub Form_Load()
    If CONNESSIONE_CONFIGURAZIONE = False Then
        Me.cmdConferma.Enabled = False
        Exit Sub
    End If
    
    INIT_CONTROLLI

End Sub
Private Sub INIT_CONTROLLI()
    'Filiale
    With Me.cboFiliale
        Set .Database = CnDMT
        .AddFieldKey "IDFiliale"
        .DisplayField = "Filiale"
        .Sql = "SELECT * FROM Filiale WHERE IDAttivitaAzienda=" & GET_LINK_ATTIVITA_AZIENDA(VarIDAzienda)
        .Fill
    End With
    
    'Magazzino di lavorazione
    With Me.cboMagLav
        Set .Database = CnDMT
        .AddFieldKey "IDMagazzino"
        .DisplayField = "Magazzino"
        .Sql = "SELECT * FROM Magazzino WHERE IDAzienda=" & VarIDAzienda
        .Fill
    End With
    
    'Impostazioni contratto
    With Me.cboTipoImpostazione
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoImpostazioneContratto"
        .DisplayField = "Descrizione"
        .Sql = "SELECT * FROM RV_POTipoImpostazioneContratto"
        .Fill
    End With
'
'    'Magazzino di vendita
'    With Me.cboMagVend
'        Set .Database = CnDMT
'        .AddFieldKey "IDMagazzino"
'        .DisplayField = "Magazzino"
'        .Sql = "SELECT * FROM Magazzino WHERE IDAzienda=" & VarIDAzienda
'        .Fill
'    End With
    
    Me.cboFiliale.WriteOn VarIDFiliale
End Sub
Private Function GET_LINK_ATTIVITA_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAttivitaAzienda FROM AttivitaAzienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ATTIVITA_AZIENDA = 0
Else
    GET_LINK_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub CREAZIONE_CAUSALI_MAGAZZINO()
On Error GoTo ERR_CREAZIONE_CAUSALI_MAGAZZINO
Dim sSQL As String
Dim rsConfig As ADODB.Recordset
Dim Link_Funzione As Long

Me.lblInfo.Caption = "CREAZIONE CAUSALE DI MAGAZZINO"
Me.txtControllo.Text = Me.txtControllo.Text & "CAUSALI DI MAGAZZINO" & vbCrLf


sSQL = "SELECT * FROM Funzione "

Set rsConfig = New ADODB.Recordset

rsConfig.Open sSQL, cnConfig

While Not rsConfig.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsConfig!Funzione) & vbCrLf
    DoEvents
    GET_LINK_FUNZIONE rsConfig!Funzione, rsConfig!IDFunzione
    
rsConfig.MoveNext
Wend

rsConfig.Close
Set rsConfig = Nothing
Exit Sub
ERR_CREAZIONE_CAUSALI_MAGAZZINO:
    MsgBox Err.Description, vbCritical, "CREAZIONE_CAUSALI_MAGAZZINO"
    Me.txtControllo.Text = Me.txtControllo.Text & " (" & Err.Description & ")" & vbCrLf
End Sub

Private Function GET_ESISTENZA_FUNZIONE(Funzione As String) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT IDFunzione FROM Funzione "
sSQL = sSQL & "WHERE Funzione=" & fnNormString(Funzione)

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    GET_ESISTENZA_FUNZIONE = 0
Else
    GET_ESISTENZA_FUNZIONE = fnNotNullN(rs!IDFunzione)
End If


rs.Close
Set rs = Nothing
End Function
Private Function GET_LINK_FUNZIONE(Funzione As String, IDFunzioneConfig As Long)
Dim sSQL As String
Dim Link_Funzione As Long
Dim Link_processo_per_funzione As Long

Link_Funzione = GET_ESISTENZA_FUNZIONE(Funzione)

If Link_Funzione = 0 Then
    'CREAZIONE FUNZIONE
    Link_Funzione = fnGetNewKeyPerTipoOggetto("Funzione", "IDFunzione")
    
    sSQL = "INSERT INTO Funzione (IDFunzione, Funzione, IDTipoOggetto, "
    sSQL = sSQL & "DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & Link_Funzione & ", "
    sSQL = sSQL & fnNormString(Funzione) & ", "
    sSQL = sSQL & 9 & ", "
    sSQL = sSQL & fnNormDate(Date) & ", "
    sSQL = sSQL & 1 & ", "
    sSQL = sSQL & 0 & ")"
    
    CnDMT.Execute sSQL
End If
Link_processo_per_funzione = CREA_PROCESSO_PER_FUNZIONE(Link_Funzione, IDFunzioneConfig)



End Function

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
    If Not rs.EOF Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Function GET_ESISTENZA_CONTATORE_PER_PROCESSO(IDContatore As Long, IDProcessoFunzione As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM ContatorePerProcesso "
sSQL = sSQL & "WHERE IDContatoreArticolo=" & IDContatore
sSQL = sSQL & " AND IDProcessoPerFunzione=" & IDProcessoFunzione
Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_CONTATORE_PER_PROCESSO = False
Else
    GET_ESISTENZA_CONTATORE_PER_PROCESSO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function CREA_PROCESSO_PER_FUNZIONE(IDFunzione As Long, IDFunzioneConfig As Long) As Long
Dim sSQL As String
Dim IDLocal As Long
Dim rsConfig As ADODB.Recordset

sSQL = "SELECT * FROM ProcessoPerFunzione "
sSQL = sSQL & " WHERE IDFunzione=" & IDFunzioneConfig
sSQL = sSQL & " ORDER BY Sequenza "

Set rsConfig = New ADODB.Recordset

rsConfig.Open sSQL, cnConfig

While Not rsConfig.EOF
    IDLocal = GET_LINK_PROCESSO_FUNZIONE(rsConfig!IDProcesso, IDFunzione)
    
    If IDLocal = 0 Then
        IDLocal = fnGetNewKeyPerTipoOggetto("ProcessoPerFunzione", "IDProcessoPerFunzione")
        
        sSQL = "INSERT INTO ProcessoPerFunzione ("
        sSQL = sSQL & "IDProcessoPerFunzione, IDFunzione, IDProcesso, Sequenza, "
        sSQL = sSQL & "DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & IDLocal & ", "
        sSQL = sSQL & IDFunzione & ", "
        sSQL = sSQL & rsConfig!IDProcesso & ", "
        sSQL = sSQL & rsConfig!Sequenza & ", "
        sSQL = sSQL & fnNormDate(Date) & ", "
        sSQL = sSQL & 1 & ", "
        sSQL = sSQL & 0 & ")"
        
        CnDMT.Execute sSQL
    End If
    
    If ((fnNotNullN(rsConfig!IDProcesso) = 49) Or (fnNotNullN(rsConfig!IDProcesso) = 63)) Then
        CREA_PROCESSO_PER_FUNZIONE = IDLocal
        CREA_CONTATORI_PER_PROCESSO fnNotNullN(rsConfig!IDProcessoPerFunzione), IDLocal
    End If

rsConfig.MoveNext
Wend

rsConfig.Close
Set rsConfig = Nothing
End Function
Private Function GET_LINK_PROCESSO_FUNZIONE(IDProcesso As Long, IDFunzione As Long) As Long

Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT IDProcessoPerFunzione FROM ProcessoPerFunzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzione
sSQL = sSQL & " AND IDProcesso=" & IDProcesso

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    GET_LINK_PROCESSO_FUNZIONE = 0
Else
    GET_LINK_PROCESSO_FUNZIONE = fnNotNullN(rs!IDProcessoPerFunzione)
End If


rs.Close
Set rs = Nothing
End Function
Private Function GET_LINK_CONTATORE_ARTICOLO(ContatoreArticolo As String) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT IDContatoreArticolo FROM ContatoreArticolo "
sSQL = sSQL & "WHERE ContatoreArticolo=" & fnNormString(ContatoreArticolo)


Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    GET_LINK_CONTATORE_ARTICOLO = 0
Else
    GET_LINK_CONTATORE_ARTICOLO = fnNotNullN(rs!IDContatoreArticolo)
End If


rs.Close
Set rs = Nothing

End Function
Private Function GET_ESISTENZA_CONTATORE_MAGAZZINO(IDContatoreArticolo As Long, IDMagazzino As Long) As Boolean
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * FROM ContatoreArticoloPerMagazzino "
sSQL = sSQL & "WHERE IDContatoreArticolo=" & IDContatoreArticolo
sSQL = sSQL & " AND IDMagazzino=" & IDMagazzino

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    GET_ESISTENZA_CONTATORE_MAGAZZINO = False
Else
    GET_ESISTENZA_CONTATORE_MAGAZZINO = True
End If


rs.Close
Set rs = Nothing

End Function

Private Sub CREA_CONTATORI_PER_PROCESSO(IDProcessoPerFunzioneConfig As Long, IDProcessoPerFunzione As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim Link_Contatore As Long
Dim rsCont As ADODB.Recordset

sSQL = "SELECT * FROM ContatorePerProcesso "
sSQL = sSQL & "WHERE IDProcessoPerFunzione=" & IDProcessoPerFunzioneConfig

Set rs = New ADODB.Recordset
rs.Open sSQL, cnConfig

If Not rs.EOF Then
    sSQL = "SELECT * FROM ContatoreArticolo "
    sSQL = sSQL & " WHERE IDContatoreArticolo=" & fnNotNullN(rs!IDContatoreArticolo)
    
    Set rsCont = New ADODB.Recordset
    rsCont.Open sSQL, cnConfig
    
    If Not rsCont.EOF Then
        Link_Contatore = GET_LINK_CONTATORE_ARTICOLO(fnNotNull(rsCont!ContatoreArticolo))
        'CREAZIONE CONTATORE ARTICOLO
        If Link_Contatore = 0 Then
            Link_Contatore = fnGetNewKey("ContatoreArticolo", "IDContatoreArticolo")
            
            sSQL = "INSERT INTO ContatoreArticolo (IDContatoreArticolo, ContatoreArticolo, "
            sSQL = sSQL & "PartecipaPrezzoMedio, PartecipaCostoMedio, VariaCostoUltimo, VariaCostoPrecedente, Predefinito) "
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & Link_Contatore & ", "
            sSQL = sSQL & fnNormString(rsCont!ContatoreArticolo) & ", "
            sSQL = sSQL & fnNormString(rsCont!PartecipaPrezzoMedio) & ", "
            sSQL = sSQL & fnNormString(rsCont!PartecipaCostoMedio) & ", "
            sSQL = sSQL & fnNormBoolean(rsCont!VariaCostoUltimo) & ", "
            sSQL = sSQL & fnNormBoolean(rsCont!VariaCostoPrecedente) & ", "
            sSQL = sSQL & fnNormBoolean(rsCont!Predefinito) & ")"
            CnDMT.Execute sSQL
         End If
             If Link_Contatore = 0 Then
                'CREAZIONE CONTATORE PER PROCESSO
                sSQL = "INSERT INTO ContatorePerProcesso (IDContatoreArticolo, IDProcessoPerFunzione, "
                sSQL = sSQL & "Numero, Quantita, Valore, DataVariazione, "
                sSQL = sSQL & "DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete) "
                sSQL = sSQL & "VALUES ("
                sSQL = sSQL & Link_Contatore & ", "
                sSQL = sSQL & IDProcessoPerFunzione & ", "
                sSQL = sSQL & fnNormString(rs!Numero) & ", "
                sSQL = sSQL & fnNormString(rs!Quantita) & ", "
                sSQL = sSQL & fnNormString(rs!Valore) & ", "
                sSQL = sSQL & fnNormBoolean(1) & ", "
                sSQL = sSQL & fnNormDate(Date) & ", "
                sSQL = sSQL & 1 & ", "
                sSQL = sSQL & 0 & ")"
                CnDMT.Execute sSQL
            Else
                If GET_ESISTENZA_CONTATORE_PER_PROCESSO(Link_Contatore, IDProcessoPerFunzione) = False Then
                    'CREAZIONE CONTATORE PER PROCESSO
                    sSQL = "INSERT INTO ContatorePerProcesso (IDContatoreArticolo, IDProcessoPerFunzione, "
                    sSQL = sSQL & "Numero, Quantita, Valore, DataVariazione, "
                    sSQL = sSQL & "DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete) "
                    sSQL = sSQL & "VALUES ("
                    sSQL = sSQL & Link_Contatore & ", "
                    sSQL = sSQL & IDProcessoPerFunzione & ", "
                    sSQL = sSQL & fnNormString(rs!Numero) & ", "
                    sSQL = sSQL & fnNormString(rs!Quantita) & ", "
                    sSQL = sSQL & fnNormString(rs!Valore) & ", "
                    sSQL = sSQL & fnNormBoolean(1) & ", "
                    sSQL = sSQL & fnNormDate(Date) & ", "
                    sSQL = sSQL & 1 & ", "
                    sSQL = sSQL & 0 & ")"
                    CnDMT.Execute sSQL
                End If
            End If

        CREA_CONTATORE_PER_MAGAZZINO Link_Contatore, rsCont!IDContatoreArticolo
    End If
    
    rsCont.Close
    Set rsCont = Nothing
End If

rs.Close
Set rs = Nothing
End Sub
Private Sub CREA_CONTATORE_PER_MAGAZZINO(IDContatore As Long, IDContatoreConfig As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * FROM ContatoreArticoloPerMagazzino "
sSQL = sSQL & "WHERE IDContatoreArticolo=" & IDContatoreConfig

Set rs = New ADODB.Recordset

rs.Open sSQL, cnConfig

If Not rs.EOF Then
    If GET_ESISTENZA_CONTATORE_MAGAZZINO(IDContatore, Me.cboMagLav.CurrentID) = False Then
        sSQL = "INSERT INTO ContatoreArticoloPerMagazzino("
        sSQL = sSQL & "IDMagazzino, IDContatoreArticolo, PartecipaGiacenza, PartecipaDisponibilita, "
        sSQL = sSQL & "DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & Me.cboMagLav.CurrentID & ", "
        sSQL = sSQL & IDContatore & ", "
        sSQL = sSQL & fnNormString(rs!PartecipaGiacenza) & ", "
        sSQL = sSQL & fnNormString(rs!PartecipaDisponibilita) & ", "
        sSQL = sSQL & fnNormDate(Date) & ", "
        sSQL = sSQL & 1 & ", "
        sSQL = sSQL & 0
        sSQL = sSQL & ")"
        CnDMT.Execute sSQL
    End If
End If

rs.Close
Set rs = Nothing
End Sub
Private Sub CREAZIONE_TIPO_DURATA()
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "TIPO DURATA"
Me.txtControllo.Text = Me.txtControllo.Text & "TIPO DURATA" & vbCrLf
DoEvents

''''''''RECUPERO DATI DURATA CONTRATTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PODurataContratto "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PODurataContratto "
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!DurataContratto) & vbCrLf
    DoEvents
    If GET_ESISTENZA_DURATA_CONTRATTO(fnNotNull(rsArc!DurataContratto)) = False Then
        rsNew.AddNew
            rsNew!IDRV_PODurataContratto = fnGetNewKey("RV_PODurataContratto", "IDRV_PODurataContratto")
            rsNew!DurataContratto = fnNotNull(rsArc!DurataContratto)
            rsNew!Mesi = fnNotNullN(rsArc!Mesi)
            rsNew!Giorni = fnNotNullN(rsArc!Giorni)
        rsNew.Update
    End If
rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing

End Sub
Private Function GET_ESISTENZA_DURATA_CONTRATTO(DurataContratto As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_PODurataContratto FROM RV_PODurataContratto "
sSQL = sSQL & " WHERE DurataContratto=" & fnNormString(DurataContratto)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_DURATA_CONTRATTO = False
Else
    GET_ESISTENZA_DURATA_CONTRATTO = True
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREAZIONE_SEZIONALI(IDFiliale As Long, IDAzienda As Long)
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "SEZIONALI PER RATA CONTRATTO"
Me.txtControllo.Text = Me.txtControllo.Text & "SEZIONALI PER RATA CONTRATTO" & vbCrLf
DoEvents

''''''''RECUPERO SEZIONALI ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM Sezionale "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM Sezionale "
sSQL = sSQL & " WHERE IDFiliale=" & IDFiliale
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!Sezionale) & vbCrLf
    DoEvents
    If GET_ESISTENZA_SEZIONALE(fnNotNull(rsArc!Sezionale), IDFiliale) = False Then
        rsNew.AddNew
            rsNew!IDSezionale = fnGetNewKey("Sezionale", "IDSezionale")
            rsNew!IDFiliale = VarIDFiliale
            rsNew!IDRegistroIva = fnNotNullN(rsArc!IDRegistroIva)
            rsNew!Sezionale = fnNotNull(rsArc!Sezionale)
            rsNew!DataUltimaVariazione = Date
            rsNew!IDUtenteUltimaVariazione = 1
            rsNew!VirtualDelete = 0
            rsNew!DescrizioneInFattDiff = ""
            rsNew!Prefisso = ""
        rsNew.Update
    End If
rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing

End Sub
Private Function GET_ESISTENZA_SEZIONALE(Sezionale As String, IDFiliale As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSezionale FROM Sezionale "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & " AND Sezionale=" & fnNormString(Sezionale)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_SEZIONALE = False
Else
    GET_ESISTENZA_SEZIONALE = True
End If


rs.CloseResultset
Set rs = Nothing
End Function
Public Function GET_ATTIVITA_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "SELECT IDAttivitaAzienda "
    sSQL = sSQL & "FROM AttivitaAzienda "
    sSQL = sSQL & " WHERE IDAzienda = " & IDAzienda
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        GET_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
    Else
        GET_ATTIVITA_AZIENDA = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing

End Function
Private Function GET_LINK_PARAMETRO_FILIALE(IDFiliale As Long, IDAzienda As Long) As Long
On Error GoTo ERR_GET_LINK_PARAMETRO_FILIALE
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim AvviaImpostazioniTabellari As Boolean
Dim IDParametro As Long


sSQL = "SELECT * FROM RV_POParametriAzienda "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

Me.lblInfo.Caption = "CREAZIONE PARAMETRI FILIALE"
Me.txtControllo.Text = Me.txtControllo.Text & Me.lblInfo.Caption & vbCrLf
DoEvents

If rs.EOF Then
    rs.AddNew
        rs!IDRV_POParametriAzienda = fnGetNewKey("RV_POParametriAzienda", "IDRV_POParametriAzienda")
        rs!IDAzienda = IDAzienda
        rs!IDFiliale = IDFiliale
        rs!IDTipoAnagraficaTecnicoIntRif = 11
        rs!IDRV_POStatoInterventoChiuso = 2
        rs!IDRV_POStatoInterventoInserimento = 1
        rs!IDTipoAnagraficaTecnicoFaseRif = 12
        rs!IDRV_POStatoFaseChiusa = 2
        rs!IDRV_POStatoFaseInserimento = 1
        rs!IDRV_POTipoFaseInterventoEla = GET_LINK_TIPO_INTERVENTO("Ordinario")
        rs!IDRV_POTipoFaseInterventoMan = GET_LINK_TIPO_INTERVENTO("Straordinario")
        rs!IDTipoAnagraficaAmministratore = 13
        rs!IDRV_POTipoGestioneBuono = 2
        rs!IDRV_POTipoNumerazioneBuono = 2
        rs!RiportaTecContratto = 0
        rs!IDTipoAnagraficaContratto = 11
        rs!TipoAddebitoObbligatorio = 0
        rs!TipoClasseObbligatorio = 0
        rs!VisualizzaImportiProdContratto = 0
        rs!ChiudiAltriInterventi = 0
        rs!IDSezionaleRateContratto = 0
        rs!VisualizzaNoteClienteAut = 0
        rs!GenAppAutAgendaContratto = 0
        rs!GenAppAutOutlookContratto = 0
        rs!IDTipoAnagraficaInstallatore = 14
        rs!IDFunzioneScaricoAddebito = 0
        rs!IDFunzioneCaricoAddebito = 0
        rs!IDRV_POTipoImpostazioneContratto = Me.cboTipoImpostazione.CurrentID
        If rs!IDRV_POTipoImpostazioneContratto = 3 Then rs!SalvaAddebitoInContratto = 1
        rs!ObbligatorioDataAppuntamento = 0
        rs!ObbligatorioRifContratto = 0

    rs.Update
    
    
End If


Exit Function

ERR_GET_LINK_PARAMETRO_FILIALE:
    MsgBox Err.Description, vbCritical, "GET_LINK_PARAMETRO_FILIALE"
    GET_LINK_PARAMETRO_FILIALE = 0
    Me.txtControllo.Text = Me.txtControllo.Text & " (" & Err.Description & ")" & vbCrLf
    
End Function
Private Function GET_LINK_SEZIONALE(Sezionale As String, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSezionale FROM Sezionale "
sSQL = sSQL & " WHERE Sezionale=" & fnNormString(Sezionale)
sSQL = sSQL & " AND IDFiliale=" & IDFiliale

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_SEZIONALE = 0
Else
    GET_LINK_SEZIONALE = fnNotNullN(rs!IDSezionale)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_LISTINO_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsAzienda As DmtOleDbLib.adoResultset
Dim Link_Listino_Imballo As Long

GET_LINK_LISTINO_AZIENDA = 0


'''''''''''''''''''''''LISTINO AZIENDA'''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDListinoDiBase "
sSQL = sSQL & "FROM ConfigurazioneVendite "
sSQL = sSQL & " WHERE IDAzienda=" & IDAzienda

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_LISTINO_AZIENDA = 0
Else
    GET_LINK_LISTINO_AZIENDA = fnNotNullN(rs!IDListinoDiBase)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function

Private Sub CREAZIONE_UNITA_DI_MISURA()
On Error GoTo ERR_UNITA_DI_MISURA_COOP

Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "UNITA DI MISURA"
Me.txtControllo.Text = Me.txtControllo.Text & "UNITA DI MISURA" & vbCrLf
DoEvents

''''''''RECUPERO DATI TIPO PRODOTTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM UnitaDiMisura "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM UnitaDiMisura "
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!UnitaDiMisura) & vbCrLf
    DoEvents
    rsNew.Filter = "UnitaDiMisura=" & fnNormString(rsArc!UnitaDiMisura)
    If rsNew.EOF Then
        rsNew.AddNew
        rsNew!IDUnitaDiMisura = fnGetNewKey("UnitaDiMisura", "IDUnitaDiMisura")
        rsNew!UnitaDiMisura = fnNotNull(rsArc!UnitaDiMisura)
        rsNew!DescrizioneFattura = Trim(fnNotNull(rsArc!DescrizioneFattura))

    End If
    
    
    rsNew.Update
    rsNew.Filter = vbNullString

rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing
Exit Sub
ERR_UNITA_DI_MISURA_COOP:
    MsgBox Err.Description, vbCritical, "UNITA_DI_MISURA_COOP"
    
End Sub
Private Function GET_ESISTENZA_UNITA_DI_MISURA(UnitaDiMisura As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisura FROM UnitaDiMisura "
sSQL = sSQL & " WHERE UnitaDiMisura=" & fnNormString(UnitaDiMisura)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_UNITA_DI_MISURA = False
Else
    GET_ESISTENZA_UNITA_DI_MISURA = True
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_ANAGRAFICA_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica FROM Azienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANAGRAFICA_AZIENDA = 0
Else
    GET_LINK_ANAGRAFICA_AZIENDA = fnNotNull(rs!IDAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_CONTROLLO_ESISTENZA_ARTICOLO(CodiceArticolo As String, IDAzienda As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT IDArticolo FROM Articolo "
sSQL = sSQL & "WHERE CodiceArticolo=" & fnNormString(CodiceArticolo)
sSQL = sSQL & " AND IDAzienda=" & IDAzienda

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_ESISTENZA_ARTICOLO = False
Else
    GET_CONTROLLO_ESISTENZA_ARTICOLO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_UM_ARTICOLO(UM As String) As Long
Dim sSQL As String

Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisura FROM UnitaDiMisura "
sSQL = sSQL & "WHERE UnitaDiMisura=" & fnNormString(UM)


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_UM_ARTICOLO = 0
Else
    GET_LINK_UM_ARTICOLO = fnNotNullN(rs!IDUnitaDiMisura)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_IVA_ARTICOLO(AliquotaIva As Double) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT IDIva FROM Iva "
sSQL = sSQL & "WHERE AliquotaIva=" & fnNormNumber(AliquotaIva)
sSQL = sSQL & " AND IDIvaDetraibile IS NULL"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_ARTICOLO = 0
Else
    GET_LINK_IVA_ARTICOLO = fnNotNullN(rs!IDIva)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub INSERIMENTO_ARTICOLO(IDAzienda As Long, IDFiliale As Long)
On Error GoTo ERR_AVVIA_INSERIMENTO_ARTICOLO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Dim LINK_LISTINO As Long
Dim LINK_UM As Long
Dim LINK_IVA As Long
Dim LINK_TIPO_PRODOTTO As Long
Dim LINK_ARTICOLO_IMBALLO As Long

sSQL = "SELECT * FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=0"
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

sSQL = "SELECT * FROM Articolo "
Set rs = New ADODB.Recordset
rs.Open sSQL, cnConfig

While Not rs.EOF
    If GET_CONTROLLO_ESISTENZA_ARTICOLO(rs!Codice, IDAzienda) = False Then
        LINK_UM = GET_LINK_UM_ARTICOLO(fnNotNull(rs!UnitaDiMisura))
        LINK_IVA = GET_LINK_IVA_ARTICOLO(fnNotNullN(rs!Iva))
        
        rsNew.AddNew
            rsNew!IDArticolo = fnGetNewKey("Articolo", "IDArticolo")
            rsNew!IDAzienda = IDAzienda
            rsNew!IDIvaAcquisto = LINK_IVA
            rsNew!IDIvaVendita = LINK_IVA
            rsNew!CodiceArticolo = fnNotNull(rs!Codice)
            rsNew!Articolo = fnNotNull(rs!Articolo)
            rsNew!IDUnitaDiMisuraAcquisto = LINK_UM
            rsNew!IDUnitaDiMisuraVendita = LINK_UM
            rsNew!IDUtenteUltimaVariazione = 1
            rsNew!VirtualDelete = 0
            rsNew!DataUltimaVariazione = Date
        rsNew.Update

    End If
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rs.Close
Set rs = Nothing

Exit Sub
ERR_AVVIA_INSERIMENTO_ARTICOLO:
    MsgBox Err.Description, vbCritical, "AVVIA_INSERIMENTO_ARTICOLO"
End Sub
Private Function GET_LINK_ANAGRAFICA_CLIENTE_DEMO(AnagraficaCliente As String, IDAzienda As Long, IDTipoAnagrafica As Long, TabelleTipoAna As String) As Long
On Error GoTo ERR_GET_LINK_ANAGRAFICA_CLIENTE_DEMO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim objAna As dmtRegAna.CRegAnagrafica
Dim ris As Long

Me.lblInfo.Caption = "CREAZIONE ANAGRAFICA CLIENTE DEMO"
Me.txtControllo.Text = Me.txtControllo.Text & "CREAZIONE ANAGRAFICA CLIENTE DEMO" & vbCrLf
DoEvents

''''''''''''Controllo dell'esistenza dell'anagrafica''''''''''''''''''''''''''''''''
sSQL = "SELECT IDAnagrafica FROM Anagrafica "
'sSQL = sSQL & " WHERE IDAzienda=" & VarIDAzienda
sSQL = sSQL & " WHERE Anagrafica=" & fnNormString(AnagraficaCliente)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANAGRAFICA_CLIENTE_DEMO = 0
Else
    GET_LINK_ANAGRAFICA_CLIENTE_DEMO = fnNotNullN(rs!IDAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If GET_LINK_ANAGRAFICA_CLIENTE_DEMO > 0 Then Exit Function

'''''''CREAZIONE DELL'ANAGRAFICA'''''''''''''''''''''''''''''''

'Crea un'istanza dell'oggetto CRegAnagrafica
Set objAna = New dmtRegAna.CRegAnagrafica

'Assegna la connessione aperta all'oggetto CReganagrafica
objAna.Connection = CnDMT

objAna.Field "Anagrafica", AnagraficaCliente, "Anagrafica" 'Richiesto

objAna.Field "DataUltimaVariazione", Date, "Anagrafica" 'Richiesto
objAna.Field "IDUtenteUltimaVariazione", 1, "Anagrafica" 'Richiesto
objAna.Field "VirtualDelete", 0, "Anagrafica" 'Richiesto
'objAna.Field "PartitaIva", "09748321008", "Anagrafica" 'Richiesto
'objAna.Field "CodiceFiscale", "09748321008", "Anagrafica" 'Richiesto
'objAna.Field "Indirizzo", "Via orvieto 12", "Anagrafica" 'Richiesto
'objAna.Field "IDComune", 5898, "Anagrafica" 'Richiesto
'objAna.Field "IDNazione", 110, "Anagrafica" 'Richiesto

'Valorizzare i campi della tabella CLIENTE, utilizzando il metodo Field
objAna.Field "IDAzienda", IDAzienda, TabelleTipoAna 'Richiesto
objAna.Field "IDTipoAnagrafica", IDTipoAnagrafica, TabelleTipoAna 'Richiesto
objAna.Field "DataUltimaVariazione", Date, TabelleTipoAna 'Richiesto
objAna.Field "IDUtenteUltimaVariazione", 1, TabelleTipoAna 'Richiesto
objAna.Field "VirtualDelete", 0, TabelleTipoAna 'Richiesto

ris = objAna.Insert

If ris = 0 Then
    GET_LINK_ANAGRAFICA_CLIENTE_DEMO = fnNotNullN(objAna.Field("IDAnagrafica", , "Anagrafica"))
End If

Set objAna = Nothing

Exit Function

ERR_GET_LINK_ANAGRAFICA_CLIENTE_DEMO:
    MsgBox Err.Description, vbCritical, "GET_LINK_ANAGRAFICA_CLIENTE_DEMO"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Private Sub CREAZIONE_TIPO_RATEIZZAZIONE()
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "TIPO RATEIZZAZIONE"
Me.txtControllo.Text = Me.txtControllo.Text & "TIPO RATEIZZAZIONE" & vbCrLf
DoEvents

''''''''RECUPERO DATI DURATA CONTRATTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PORateizzazione "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PORateizzazione "
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!Rateizzazione) & vbCrLf
    DoEvents
    If GET_ESISTENZA_DURATA_RATEIZZAZIONE(fnNotNull(rsArc!Rateizzazione)) = False Then
        rsNew.AddNew
            rsNew!IDRV_PORateizzazione = fnGetNewKey("RV_PORateizzazione", "IDRV_PORateizzazione")
            rsNew!Rateizzazione = fnNotNull(rsArc!Rateizzazione)
            rsNew!NumeroRate = fnNotNullN(rsArc!NumeroRate)
            rsNew!PagamentoInizioPeriodo = fnNotNullN(rsArc!PagamentoInizioPeriodo)
            rsNew!Mesi = fnNotNullN(rsArc!Mesi)
            
        rsNew.Update
    End If
rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing

End Sub
Private Function GET_ESISTENZA_DURATA_RATEIZZAZIONE(Rateizzazione As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_PORateizzazione FROM RV_PORateizzazione "
sSQL = sSQL & " WHERE Rateizzazione=" & fnNormString(Rateizzazione)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_DURATA_RATEIZZAZIONE = False
Else
    GET_ESISTENZA_DURATA_RATEIZZAZIONE = True
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREAZIONE_TIPO_DURATA_RINNOVO()
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "DURATA RINNOVO"
Me.txtControllo.Text = Me.txtControllo.Text & "DURATA RINNOVO" & vbCrLf
DoEvents

''''''''RECUPERO DATI DURATA CONTRATTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POTipoRinnovo "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POTipoRinnovo "
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!TipoRinnovo) & vbCrLf
    DoEvents
    If GET_ESISTENZA_DURATA_TIPO_RINNOVO(fnNotNull(rsArc!TipoRinnovo)) = False Then
        rsNew.AddNew
            rsNew!IDRV_POTipoRinnovo = fnGetNewKey("RV_POTipoRinnovo", "IDRV_POTipoRinnovo")
            rsNew!TipoRinnovo = fnNotNull(rsArc!TipoRinnovo)
            rsNew!Mesi = fnNotNullN(rsArc!Mesi)
            rsNew!Giorni = fnNotNullN(rsArc!Giorni)
            rsNew!AnnoPrecedente = fnNotNullN(rsArc!AnnoPrecedente)
        rsNew.Update
    End If
rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing

End Sub
Private Function GET_ESISTENZA_DURATA_TIPO_RINNOVO(TipoRinnovo As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoRinnovo FROM RV_POTipoRinnovo "
sSQL = sSQL & " WHERE TipoRinnovo=" & fnNormString(TipoRinnovo)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_DURATA_TIPO_RINNOVO = False
Else
    GET_ESISTENZA_DURATA_TIPO_RINNOVO = True
End If


rs.CloseResultset
Set rs = Nothing
End Function

Private Sub CREAZIONE_TIPO_DURATA_ASSITENZA()
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "DURATA ASSISTENZA"
Me.txtControllo.Text = Me.txtControllo.Text & "DURATA ASSISTENZA" & vbCrLf
DoEvents

''''''''RECUPERO DATI DURATA CONTRATTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POTipoDurataAssistenza "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POTipoDurataAssistenza "
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!TipoDurataAssistenza) & vbCrLf
    DoEvents
    If GET_ESISTENZA_DURATA_ASSISTENZA(fnNotNull(rsArc!TipoDurataAssistenza)) = False Then
        rsNew.AddNew
            rsNew!IDRV_POTipoDurataAssistenza = fnGetNewKey("RV_POTipoDurataAssistenza", "IDRV_POTipoDurataAssistenza")
            rsNew!TipoDurataAssistenza = fnNotNull(rsArc!TipoDurataAssistenza)
            rsNew!Mesi = fnNotNullN(rsArc!Mesi)
            rsNew!Giorni = fnNotNullN(rsArc!Giorni)
        rsNew.Update
    End If
rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing

End Sub
Private Function GET_ESISTENZA_DURATA_ASSISTENZA(TipoDurataAssistenza As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoDurataAssistenza FROM RV_POTipoDurataAssistenza "
sSQL = sSQL & " WHERE TipoDurataAssistenza=" & fnNormString(TipoDurataAssistenza)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_DURATA_ASSISTENZA = False
Else
    GET_ESISTENZA_DURATA_ASSISTENZA = True
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREAZIONE_TIPO_INTERVENTO()
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "TIPO INTERVENTO"
Me.txtControllo.Text = Me.txtControllo.Text & "TIPO INTERVENTO" & vbCrLf
DoEvents

''''''''RECUPERO DATI DURATA CONTRATTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POTipoFaseIntervento "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_POTipoFaseIntervento "
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!TipoFaseIntervento) & vbCrLf
    DoEvents
    If GET_ESISTENZA_DURATA_TIPO_INTERVENTO(fnNotNull(rsArc!TipoFaseIntervento)) = False Then
        rsNew.AddNew
            rsNew!IDRV_POTipoFaseIntervento = fnGetNewKey("RV_POTipoFaseIntervento", "IDRV_POTipoFaseIntervento")
            rsNew!TipoFaseIntervento = fnNotNull(rsArc!TipoFaseIntervento)
        rsNew.Update
    End If
rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing

End Sub
Private Function GET_ESISTENZA_DURATA_TIPO_INTERVENTO(TipoFaseIntervento As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoFaseIntervento FROM RV_POTipoFaseIntervento "
sSQL = sSQL & " WHERE TipoFaseIntervento=" & fnNormString(TipoFaseIntervento)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_DURATA_TIPO_INTERVENTO = False
Else
    GET_ESISTENZA_DURATA_TIPO_INTERVENTO = True
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREA_TIPO_ANAGRAFICA()
Dim sSQL As String

sSQL = "UPDATE TipoAnagrafica SET "
sSQL = sSQL & "TipoAnagrafica=" & fnNormString("Tecnici di riferimento")
sSQL = sSQL & " WHERE TipoAnagrafica=" & fnNormString("Anagrafica Tipo1")
CnDMT.Execute sSQL

sSQL = "UPDATE TipoAnagrafica SET "
sSQL = sSQL & "TipoAnagrafica=" & fnNormString("Tecnici operativi")
sSQL = sSQL & " WHERE TipoAnagrafica=" & fnNormString("Anagrafica Tipo2")
CnDMT.Execute sSQL

sSQL = "UPDATE TipoAnagrafica SET "
sSQL = sSQL & "TipoAnagrafica=" & fnNormString("Amministratore")
sSQL = sSQL & " WHERE TipoAnagrafica=" & fnNormString("Anagrafica Tipo3")
CnDMT.Execute sSQL

sSQL = "UPDATE TipoAnagrafica SET "
sSQL = sSQL & "TipoAnagrafica=" & fnNormString("Installatori")
sSQL = sSQL & " WHERE TipoAnagrafica=" & fnNormString("Anagrafica Tipo4")
CnDMT.Execute sSQL

End Sub
Private Function GET_LINK_TIPO_INTERVENTO(TipoIntervento As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoFaseIntervento FROM RV_POTipoFaseIntervento "
sSQL = sSQL & "WHERE TipoFaseIntervento=" & fnNormString(TipoIntervento)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_TIPO_INTERVENTO = 0
Else
    GET_LINK_TIPO_INTERVENTO = fnNotNullN(rs!IDRV_POTipoFaseIntervento)
End If

rs.CloseResultset
Set rs = Nothing
End Function

