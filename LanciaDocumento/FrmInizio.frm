VERSION 5.00
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Object = "{E1215E52-40E1-11D3-AF44-00105A2FBE61}#5.1#0"; "DMTLblLinkCtl.ocx"
Begin VB.Form FrmInizio 
   Caption         =   "STAMPE"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin DMTLblLinkCtl.LabelLink LabelLink1 
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   8760
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Name            =   "LabelLink"
   End
   Begin VB.CheckBox chkStampaSingola 
      Caption         =   "Stampa il record selezionato"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   7680
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.CommandButton cmdStampaAnteprima 
      Caption         =   "ANTEPRIMA"
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
      Height          =   570
      Left            =   120
      Picture         =   "FrmInizio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "STAMPA"
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton cmdStampa 
      Caption         =   "STAMPA DIRETTA"
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
      Height          =   570
      Left            =   3600
      Picture         =   "FrmInizio.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "STAMPA"
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Frame FraStampa 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5895
      Begin DmtActBox.DmtActBoxCtl ActivityBox 
         Height          =   6855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   12091
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
   End
   Begin VB.Label lblFunzione 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "FrmInizio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents m_App As DMTRunAppLib.Application
Attribute m_App.VB_VarHelpID = -1
Private oReport As dmtReportLib.dmtReport
Private IDTipoOggettoPrg As Long

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


'------Variabili dal Registry-----------------------------------------------
Private IDTipoOggettoReg As Long
Private sqlWhere As String
Private ChiavePrimariaReg As String
Private ValoreChiavePrimaria As Long
Private FunzioneReg As String

Private IDOggetto As Long
Private IDFunzione As Long





Public Sub ConnessioneADO()
    ConnessioneADODBLib
    
'    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
'
'    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
    Get_Chiavi_Registry
    
'    If IDTipoOggettoReg > 0 Then
'        ConfigurazioneStampe IDTipoOggettoReg
'        lblFunzione = FunzioneReg
'        cmdStampa.Enabled = True
'        cmdStampaAnteprima.Enabled = True
'    End If
    
        'L'Applicazione.

    
    Set Me.LabelLink1.Application = TheApp
    Me.LabelLink1.WindowHandleClient = Me.hWnd
    Me.LabelLink1.PopMenuItems("Mnu_SearchObject").Enabled = False
    

    LabelLink1.IDFunction = 109 'IDFunzione
    Me.LabelLink1.IDReturn = 329887 'IDOggetto
    Me.LabelLink1.RunApplication

    
    
    
End Sub
Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set m_App = NewValue
End Property
Public Property Get Application() As DMTRunAppLib.Application
    Set Application = m_App
End Property

Private Sub cmdStampa_Click()
On Error GoTo ERR_cmdStampaAnteprima_Click
Dim sSQL As String
Dim IDReportOLD As Long
Dim IDReport As Long

IDTipoOggettoPrg = IDTipoOggettoReg


Set oReport = New dmtReportLib.dmtReport
Set oReport.Connection = CnDMT
oReport.Password = Password
oReport.User = Utente

'Imposta l'idfiliale di appartenenza del documento da stampare
oReport.BranchID = LINK_FILIALE 'IDFiliale
'Imposta l'identificativo del tipo di documento
oReport.DocTypeID = IDTipoOggettoPrg


'If Me.chkStampaSingola.Value = Unchecked Then
'    oReport.Where = sqlWhere
'Else
'    oReport.Where = ChiavePrimariaReg & "=" & ValoreChiavePrimaria
'End If

oReport.Where = "IDUtente=" & LINK_UTENTE

IDReportOLD = fncTrovaReport(oReportsActivity.DefaultReportName, IDTipoOggettoPrg)

If Len(oReportsActivity.SelectedReportName) > 0 Then
    IDReport = fncTrovaReport(oReportsActivity.SelectedReportName, IDTipoOggettoPrg)
Else
    IDReport = fncTrovaReport(oReportsActivity.DefaultReportName, IDTipoOggettoPrg)
End If

If IDReport > 0 Then
    fncImpostaDefaultReport IDReport, IDTipoOggettoPrg
    
    oReport.DoPrint oReport.PrinterName
    If (IDReport <> IDReportOLD) Then
        fncImpostaDefaultReport IDReportOLD, IDTipoOggettoPrg
    End If
Else
    MsgBox "ATTENZIONE!!!!" & vbCrLf & "Il report non è stato trovato!", vbCritical, "Impossibile stampare"
End If

Exit Sub
ERR_cmdStampaAnteprima_Click:
    MsgBox Err.Description, vbCritical, "Stampa documento"
End Sub

Private Sub cmdStampaAnteprima_Click()
'On Error GoTo ERR_cmdStampaAnteprima_Click
Dim sSQL As String
Dim IDReportOLD As Long
Dim IDReport As Long

IDTipoOggettoPrg = IDTipoOggettoReg


Set oReport = New dmtReportLib.dmtReport
Set oReport.Connection = CnDMT
oReport.Password = Password
oReport.User = Utente


'Imposta l'idfiliale di appartenenza del documento da stampare
oReport.BranchID = LINK_FILIALE 'IDFiliale
'Imposta l'identificativo del tipo di documento
oReport.DocTypeID = IDTipoOggettoPrg


oReport.Where = "IDUtente=" & LINK_UTENTE

IDReportOLD = fncTrovaReport(oReportsActivity.DefaultReportName, IDTipoOggettoPrg)

If Len(oReportsActivity.SelectedReportName) > 0 Then
    IDReport = fncTrovaReport(oReportsActivity.SelectedReportName, IDTipoOggettoPrg)
Else
    IDReport = fncTrovaReport(oReportsActivity.DefaultReportName, IDTipoOggettoPrg)
End If

If IDReport > 0 Then
    fncImpostaDefaultReport IDReport, IDTipoOggettoPrg
    
    oReport.Preview 0, 0, 0
    If (IDReport <> IDReportOLD) Then
        fncImpostaDefaultReport IDReportOLD, IDTipoOggettoPrg
    End If
Else
    MsgBox "ATTENZIONE!!!!" & vbCrLf & "Il report non è stato trovato!", vbCritical, "Impossibile stampare"
End If

Exit Sub
ERR_cmdStampaAnteprima_Click:
    MsgBox Err.Description, vbCritical, "Stampa documento"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        STRINGA_SQL = sqlWhere
        FrmFine.Show vbModal
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load

Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub

Private Sub ConfigurazioneStampe(IDTipoOggetto As Long)
Dim oActivity As IActivity
Dim o As Activity
Dim oFilter As Filter

    'Inizializzazione del riquadro attività
    With ActivityBox
        .Activities.Clear
        
        'Aggiunge l'attività dei reports
        Set oActivity = .Activities.Add("DmtActBoxLib.ReportsActivity", "Reports")
        Set oActivity.Connection = CnDMT.InternalConnection
        
        oActivity.Load IDTipoOggettoReg, LINK_AZIENDA
        Set o = oActivity
        Set oReportsActivity = o.InternalClass
        
        'Imposta quale attività deve essere attivata per default
        If m_DefaultActivity <> "" Then
            Set .CurrentActivity = .Activities(m_DefaultActivity)
        End If
        
        'ridisegna il controllo
        .Redraw = True
        
        oReportsActivity.Is4DlgPrint = False
    End With

End Sub
Private Function fnGetTipoOggetto(IDTipoOggetto As Long) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto "
    sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub Get_Chiavi_Registry()

    

    IDFunzione = GetSetting(gResource.GetMessage(LBL_REGISTRY_KEY), App.EXEName, "IDFunzioneHD", 0)
    IDOggetto = GetSetting(gResource.GetMessage(LBL_REGISTRY_KEY), App.EXEName, "IDOggettoHD", 0)
    
End Sub
Private Function fncImpostaDefaultReport(ByVal IDReportDefault As Long, IDTipoOggetto As Long)
On Error GoTo ERR_fncImpostaDefaultReport
    Dim sSQL As String
    
    sSQL = "UPDATE DefaultFilialePerTipoOggetto SET "
    sSQL = sSQL & "IDReportTipoOggetto=" & IDReportDefault
    sSQL = sSQL & " WHERE IDTipoOggetto = " & IDTipoOggetto
    sSQL = sSQL & " AND IDFiliale = " & LINK_FILIALE
    
    CnDMT.Execute sSQL
    
Exit Function
ERR_fncImpostaDefaultReport:
    MsgBox Err.Description, vbCritical, "Settaggio report di default"
End Function
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

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    DeleteSetting gResource.GetMessage(LBL_REGISTRY_KEY), App.EXEName
    
End Sub

