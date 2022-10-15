VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form FrmMovimenti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fatturato (Passo 2 di 2)"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStampa 
      Caption         =   "Stampa"
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
      Left            =   10800
      TabIndex        =   4
      Top             =   6720
      Width           =   1335
   End
   Begin DmtGridCtl.DmtGrid GrigliaRate 
      Height          =   5895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   10398
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
      Left            =   7920
      TabIndex        =   1
      Top             =   6720
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
      Left            =   9360
      TabIndex        =   0
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CheckBox chkRaggruppaFatture 
      Caption         =   "Raggruppa rate per cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "FrmMovimenti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsRate As DmtOleDbLib.adoResultset
Private oReport As dmtReportLib.dmtReport
Private IDTipoOggettoPrg As Long
Private Riga_selezionata As Long
Private Aggiornamento As Long

Private Sub cmdAnnulla_Click()
Dim Risposta As Integer
    Risposta = MsgBox("Vuoi abbandonare il wizard per la stampa del fatturato?", vbInformation + vbYesNo, "Abbandono")
        
    If Risposta = vbYes Then
        Unload Me
    End If
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
                oReport.Password = TheApp.Password
                oReport.User = TheApp.User
            End If
        
        
        'Imposta l'idfiliale di appartenenza del documento da stampare
            oReport.BranchID = TheApp.Branch 'IDFiliale
        'Imposta l'identificativo del tipo di documento
            oReport.DocTypeID = IDTipoOggettoPrg
            
            sSQL = "Mese >= " & VAR_DA_MESE
            sSQL = sSQL & " AND Mese <= " & VAR_A_MESE
            sSQL = sSQL & " AND ContrattoAttuale=1"
            sSQL = sSQL & " AND ImportoRata>0"
            sSQL = sSQL & " AND Disdetta=0"
            
            oReport.Where = sSQL
            
            IDReport = fncTrovaReport("RV_PORepFatturato.rpt", IDTipoOggettoPrg)
        
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
    MsgBox Err.Description, vbCritical, "Stampa fatturato"
End Sub



Private Sub Form_Load()
    Dim TotaleDocumento As Double
    
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    SettaggioGrigliaRate
    
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If cmdIndietro.Value = True Then
        FrmInizio.Show
        Exit Sub
    End If
        
    
End Sub

Private Sub SettaggioGrigliaRate()
On Error GoTo ERR_SettaggioGrigliaRate
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    
    sSQL = "SELECT * FROM RV_PORepFatturato "
    sSQL = sSQL & "WHERE Disdetta=0 "
    sSQL = sSQL & " AND NonFatturare=0 "
    sSQL = sSQL & " AND Mese>=" & VAR_DA_MESE
    sSQL = sSQL & " AND Mese<=" & VAR_A_MESE
    sSQL = sSQL & " AND ContrattoAttuale=1 "
    sSQL = sSQL & " AND ImportoRata>0"
    sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
    sSQL = sSQL & " ORDER BY Anagrafica"
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
        Set rsRate = CnDMT.OpenResultset(sSQL)
            Set rsEvent = rsRate.Data
        
        With Me.GrigliaRate
            .ColumnsHeader.Clear
            .ColumnsHeader.Add "Anagrafica", "Cliente", dgchar, True, 3000, dgAlignleft
            .ColumnsHeader.Add "TipoContratto", "Tipo contratto", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "DataRata", "Data Rata", dgDate, True, 2000, dgAlignleft
            Set cl = .ColumnsHeader.Add("ImportoRata", "Importo", dgCurrency, True, 2000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                    
                    
            Set .Recordset = rsRate.Data
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
    If Me.GrigliaRate.Recordset.EOF = False Then
        If Riga_selezionata = 0 Then
            Me.GrigliaRate.Recordset.Move Riga_selezionata
        Else
            Me.GrigliaRate.Recordset.Move Riga_selezionata + 1
        End If
            
    End If
Exit Sub
ERR_SettaggioGrigliaRate:
    MsgBox Err.Description, vbCritical, "SettaggioGrigliaRate"
End Sub



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

