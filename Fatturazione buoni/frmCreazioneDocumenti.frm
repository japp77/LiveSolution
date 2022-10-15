VERSION 5.00
Object = "{FCA49525-5F72-11D2-B9EB-00201880103B}#18.1#0"; "DMTPrinterDialog.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmCreazioneDocumenti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fatturazione buoni (Passo 3 di 3)"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   Icon            =   "frmCreazioneDocumenti.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Avanti"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opzioni"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4560
      TabIndex        =   5
      Top             =   1440
      Width           =   3855
      Begin DMTEDITNUMLib.dmtNumber txtNumeroCopie 
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   960
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
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
         Appearance      =   1
         DecimalPlaces   =   0
         AllowEmpty      =   0   'False
      End
      Begin VB.CheckBox chkStampa 
         Alignment       =   1  'Right Justify
         Caption         =   "Stampa immediata del documento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Numero di copie per stampa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   0
      Picture         =   "frmCreazioneDocumenti.frx":4781A
      ScaleHeight     =   3465
      ScaleWidth      =   4425
      TabIndex        =   4
      Top             =   0
      Width           =   4455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Max             =   1,00000e5
   End
   Begin VB.CommandButton cmdIndietro 
      Caption         =   "Indietro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   4680
      Width           =   1335
   End
   Begin DMTPrinterDialog.DMTDialog DMTDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdFine 
      Caption         =   "Fine"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   8295
   End
End
Attribute VB_Name = "frmCreazioneDocumenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oReport As dmtReportLib.dmtReport
Private IDTipoOggettoPrg As Long
Private Sub cmdAnnulla_Click()
Dim Risposta As Integer
    Risposta = MsgBox("Vuoi abbandonare il wizard per il passaggio degli addebiti interventi in fatturazione?", vbInformation + vbYesNo, "Abbandono")
        
    If Risposta = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdAvanti_Click()
    Unload Me
End Sub

Private Sub CmdFine_Click()
    On Error GoTo ERR_CmdFine_Click
    CmdFine.Enabled = False
    
    fncPassaggioDocumenti
    
    If (LINK_CARICO_MAGAZZINO > 0) Then
        MOVIMENTAZIONE_RIGHE
    End If
    
    
    Me.cmdAvanti.Enabled = True
    Me.lblInfo.Caption = "OPERAZIONE COMPLETATA"
    
    Exit Sub
ERR_CmdFine_Click:
    MsgBox Err.Description, vbCritical, "CmdFine_Click"
    CmdFine.Enabled = True
    Me.cmdAvanti.Enabled = True
End Sub
Private Sub MOVIMENTAZIONE_RIGHE()
Dim Mov As DmtMovim.cMovimentazione
Dim DescrizioneFunzione As String
Dim UnitaProgresso As Double
Dim Numero As Long

If (NUMERO_ADDEBITI_DA_FATTURARE = 0) Then Exit Sub

lblInfo.Caption = "MOVIMENTAZIONE IN CORSO..."
DoEvents

ProgressBar1.Value = 0

ProgressBar1.Max = 100

Unita_Progresso = FormatNumber((ProgressBar1.Max / NUMERO_ADDEBITI_DA_FATTURARE), 4)
   
DescrizioneFunzione = fncTrovaDescrizioneFunzione("RV_POBuonoIntervento")

Set Mov = New DmtMovim.cMovimentazione
Set Mov.Connection = TheApp.Database.Connection

rsnew.MoveFirst
Numero = 0


While Not rsnew.EOF
    Numero = Numero + 1
    lblInfo.Caption = "MOVIMETAZIONE " & Numero & " di " & NUMERO_ADDEBITI_DA_FATTURARE
    DoEvents
    
    If Abs(fnNotNullN(rsnew!Movimentato)) = 1 Then
        
        Mov.DataMovimento = Data_Documento
        Mov.FattoreDiConversione = Null
        Mov.GestioneLotti = False
        Mov.GestioneMatricole = False
        Mov.IDEsercizio = fnGetEsercizio(Data_Documento)
        Mov.IDTipoOggetto = rsnew!IDTipoOggettoBuono
        Mov.IDOggetto = rsnew!IDOggettoBuono
        Mov.IDFunzione = LINK_CARICO_MAGAZZINO
        Mov.IDUtente = TheApp.IDUser
        Mov.IDMagazzinoEntrata = rsnew!IDMagazzino
        Mov.IDMagazzinoUscita = rsnew!IDMagazzino
        Mov.Cessione = 0
        Mov.Field "IDAzienda", TheApp.IDFirm
        Mov.Field "IDAnagrafica", rsnew!IDAnagraficaCliente
        Mov.Field "IDTipoAnagrafica", 2
        Mov.Field "IDArticolo", rsnew!IDArticolo
        Mov.Field "IDUnitaDiMisura", rsnew!IDUnitaDiMisura
        Mov.Field "IDcambio", Null
        Mov.Field "DescrizioneArticolo", rsnew!Articolo
        Mov.Field "QuantitaTotale", rsnew!Quantita
        Mov.Field "QuantitaMovimentata", rsnew!Quantita
        Mov.Field "Importo", rsnew!ImportoDiFatturazione
        Mov.Field "DataDocumento", rsnew!DataDocumento
        Mov.Field "NumeroDocumento", rsnew!NumeroDocumento
        Mov.Field "Oggetto", DescrizioneFunzione
        Mov.Field "IDTipoMovimento", 1
        Mov.Field "PrezzoUnitario", rsnew!ImportoUnitario
        Mov.Field "IDValoriOggettoDettaglio", rsnew!IDRV_POInterventoRigheDett
        Mov.Field "TipoRiga", trcNessuno
        
        GeneraMovimento = Mov.Insert
    End If
    
    If (Me.ProgressBar1.Value + UnitaProgresso) > Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + UnitaProgresso
    End If
    
    DoEvents
        
rsnew.MoveNext
Wend

Set Mov = Nothing

End Sub
Private Function fncTrovaDescrizioneFunzione(Gestore As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim X As Long

sSQL = "SELECT Funzione.IDFunzione, Gestore.Gestore, Funzione.Funzione "
sSQL = sSQL & "FROM Gestore INNER JOIN "
sSQL = sSQL & "TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore INNER JOIN "
sSQL = sSQL & "Funzione ON TipoOggetto.IDTipoOggetto = Funzione.IDTipoOggetto "
sSQL = sSQL & "WHERE (Gestore.Gestore = " & fnNormString(Gestore) & ") "
sSQL = sSQL & "AND (Funzione.IDFunzione >= 10000)"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaDescrizioneFunzione = fnNotNull(rs!Funzione)
Else
    fncTrovaDescrizioneFunzione = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub cmdIndietro_Click()
    Unload Me
End Sub


Private Sub cmdStampa_Click()
On Error GoTo ERR_cmdStampa_Click
Dim sSQL As String

        
    IDTipoOggettoPrg = fncIDTipoOggettoPrg
    
        Set oReport = New dmtReportLib.dmtReport
            Set oReport.Connection = CnDMT
            If MenuOptions.DBType = 1 Then
                'parametri di accesso al database ACCESS
                oReport.Password = "dmt192981046"
                oReport.User = "admin"
            Else
                'parametri di accesso al database SQL Server
                oReport.Password = Password
                oReport.User = Utente
            End If
        
        
        'Imposta l'idfiliale di appartenenza del documento da stampare
            oReport.BranchID = TheApp.Branch 'IDFiliale
        'Imposta l'identificativo del tipo di documento
            oReport.DocTypeID = IDTipoOggettoPrg
            
            
            
            
            
            IDReport = fncTrovaReport("RV_PORepFatturazioneRate_After.rpt", IDTipoOggettoPrg)
        
            If IDReport > 0 Then
                fncImpostaDefaultReport (IDReport)
                'Effettua l'anteprima di stampa
                'Settare il nome della stampante per questo tipo di stampa
                
                'oReport.PrinterName = fncTrovaStampante(IDReport)
                oReport.Preview 0, 0, 0
            Else
                MsgBox "ATTENZIONE!!!!" & vbCrLf & "Il report non è stato trovato!", vbCritical, "Impossibile stampare"
            End If
Exit Sub
ERR_cmdStampa_Click:
    MsgBox Err.Description, vbCritical, "Stampa"
End Sub

Private Sub Form_Load()
    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    Me.Caption = TheApp.FunctionName & " (Passo 3 di 3)"
   
    Me.txtNumeroCopie.Text = 1
    Me.chkStampa.Value = 1
    Me.cmdAvanti.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.cmdIndietro.Value = True Then
        FrmInizio.Show
        Exit Sub
    End If
    If Me.cmdAvanti.Value = True Then
        FrmInizio.Show
        Exit Sub
    End If
End Sub

Private Sub txtNumeroCopie_Change()
    If Me.txtNumeroCopie.Text = 0 Then
        Me.txtNumeroCopie.Text = 1
    End If
End Sub
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

Private Function fncImpostaDefaultReport(ByVal IDReportDefault As Long)
On Error GoTo ERR_fncImpostaDefaultReport
    Dim sSQL As String
    
    sSQL = "UPDATE DefaultFilialePerTipoOggetto SET "
    sSQL = sSQL & "IDReportTipoOggetto=" & IDReportDefault
    sSQL = sSQL & " WHERE IDTipoOggetto = " & IDTipoOggettoPrg & " AND IDFiliale = " & TheApp.Branch
    
    CnDMT.Execute sSQL
    
Exit Function
ERR_fncImpostaDefaultReport:
    MsgBox Err.Description, vbCritical, "Settaggio report di default"
End Function
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
Private Sub CREA_ARRAY_ANA()
Dim I As Integer
'Eliminazione dati dall'array che contiene gli identificativi della anagrafiche da fatturare
For I = 0 To 100
    ArrayAna(0) = 0
Next

'Ciclo che inserisce i dati degli identificativi della anagrafiche da fatturare
If Not (rsAnaFatt.EOF And rsAnaFatt.BOF) Then
    rsAnaFatt.MoveFirst
    I = 0
    NUMERO_ANAGRAFICHE = 0
    While Not rsAnaFatt.EOF
        ArrayAna(I) = fnNotNullN(rsAnaFatt!IDAnagraficaCliente)
        I = I + 1
        NUMERO_ANAGRAFICHE = NUMERO_ANAGRAFICHE + 1
    rsAnaFatt.MoveNext
    Wend
End If

End Sub
Private Sub CREA_BUONI_TMP_DA_FATTURARE()
Dim sSQL As String
Dim rs As ADODB.Recordset

'ELIMINAZIONE FATTURAZIONE BUONI''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POTMPFatturazioneBuoni"
CnDMT.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''INSERIMENTO BUONI DA FATTURARE'''''''''''''''''''''''''''''''''''''''
Set rs = New ADODB.Recordset

sSQL = "SELECT * FROM RV_POTMPFatturazioneBuoni"

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

rsBuoniFatt.Filter = ""

If Not (rsBuoniFatt.EOF And rsBuoniFatt.BOF) Then
    rsBuoniFatt.MoveFirst
    While Not rsBuoniFatt.EOF
        rs.AddNew
            rs!IDRV_POInterventoRigheDett = rsBuoniFatt!IDRV_POInterventoRigheDett
            rs!IDRV_POintervento = rsBuoniFatt!IDRV_POintervento
            rs!NumeroDocumento = rsBuoniFatt!NumeroBuono
            rs!DataDocumento = rsBuoniFatt!DataBuono
            rs!IDAnagraficaCliente = rsBuoniFatt!IDAnagraficaCliente
            rs!IDAnagraficaIntervento = rsBuoniFatt!IDAnagraficaIntervento
            rs!ImportoDiFatturazione = rsBuoniFatt!ImportoFatturazione
            rs!RitenutaAcconto = rsBuoniFatt!RitenutaAcconto
            rs!Quantita = rsBuoniFatt!Quantita
            rs!ImportoUnitario = rsBuoniFatt!ImportoUnitario
            rs!IDIva = rsBuoniFatt!IDIva
            rs!AliquotaIva = rsBuoniFatt!AliquotaIva
            rs!Sconto1 = rsBuoniFatt!Sconto1
            rs!Sconto2 = rsBuoniFatt!Sconto2
            rs!Sconto3 = rsBuoniFatt!Sconto3
        rs.Update
    rsBuoniFatt.MoveNext
    Wend
End If

rs.Close
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
