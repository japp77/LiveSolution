VERSION 5.00
Object = "{FCA49525-5F72-11D2-B9EB-00201880103B}#18.1#0"; "DMTPrinterDialog.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmCreazioneDocumenti 
   Caption         =   "Creazione documenti (passo 4 di 4)"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCreazioneDocumenti.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStampa 
      Caption         =   "Stampa"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   4560
      TabIndex        =   5
      Top             =   1200
      Width           =   3855
      Begin VB.CheckBox chkSalvaInPDF 
         Alignment       =   1  'Right Justify
         Caption         =   "Salva documenti in PDF"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   3615
      End
      Begin VB.ComboBox cboStampante 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1680
         Width           =   3615
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroCopie 
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   1080
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
      Begin VB.CheckBox chkStampa 
         Alignment       =   1  'Right Justify
         Caption         =   "Stampa immediata del documento"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Stampante"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Numero di copie per stampa"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   5640
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
      TabIndex        =   10
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

Private Sub cboStampante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Me.cboStampante.ListIndex = -1
        
    End If
    If KeyCode = vbKeyCancel Then
        Me.cboStampante.ListIndex = -1
        
    End If
End Sub

Private Sub cmdAnnulla_Click()
Dim Risposta As Integer
    Risposta = MsgBox("Vuoi abbandonare il wizard per il passaggio degli interventi in contabilità?", vbInformation + vbYesNo, "Abbandono")
        
    If Risposta = vbYes Then
        Unload Me
    End If
End Sub

Private Sub CmdFine_Click()
On Error GoTo ERR_CmdFine_Click

    Me.cmdFine.Enabled = False
    
    GET_PARAMETRI_STRINGA_FATT
    
    fncPassaggioDocumenti
    
    cmdFine.Enabled = False
    
    cmdStampa.Enabled = True
    
    Unload Me

Exit Sub
ERR_CmdFine_Click:
    MsgBox Err.Description, vbCritical, "CmdFine_Click"
End Sub


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
'    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    Me.txtNumeroCopie.Text = 1
    Me.chkStampa.Value = 1
    Me.chkSalvaInPDF.Value = 1
    Me.cmdStampa.Enabled = False
    fncStampantiPedana
    
End Sub
Private Sub fncStampantiPedana()
Dim prn As Printer

Me.cboStampante.Clear

For Each prn In Printers
    Me.cboStampante.AddItem prn.DeviceName
Next

End Sub
Private Sub Form_Unload(Cancel As Integer)

    If Me.cmdIndietro.Value = True Then
        FrmParametri.Show
        Exit Sub
    End If
    
    If Me.cmdFine.Value = True Then
        FrmInizio.Show
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
Private Sub GET_PARAMETRI_STRINGA_FATT()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

ARTICOLO_PIU_PERIODO_FATT = 0

sSQL = "SELECT ArticoloPiuPeriodoFatturazione "
sSQL = sSQL & " FROM RV_POStringaPeriodoTesta "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    ARTICOLO_PIU_PERIODO_FATT = fnNotNullN(rs!ArticoloPiuPeriodoFatturazione)
End If

rs.CloseResultset
Set rs = Nothing
End Sub
