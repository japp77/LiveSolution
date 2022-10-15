VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmDocumentazione 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DOCUMENTAZIONE"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12405
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
   ScaleHeight     =   6915
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   375
      Left            =   11400
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   11040
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   0
      ScaleHeight     =   6795
      ScaleWidth      =   12315
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.CommandButton cmdCompleto 
         Caption         =   "COMPLETO"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   6240
         Width           =   2055
      End
      Begin VB.CommandButton cmdInterventoPadre 
         Caption         =   "INTERVENTO"
         Height          =   495
         Left            =   5880
         TabIndex        =   9
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdIntervento 
         Caption         =   "FASE"
         Height          =   495
         Left            =   7800
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdContratto 
         Caption         =   "CONTRATTO"
         Height          =   495
         Left            =   3960
         TabIndex        =   7
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdCliente 
         Caption         =   "CLIENTE"
         Height          =   495
         Left            =   2040
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdAzienda 
         Caption         =   "AZIENDA"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdCaricaDocumento 
         Caption         =   "CARICA DOCUMENTO"
         Height          =   495
         Left            =   9720
         TabIndex        =   4
         Top             =   120
         Width           =   2535
      End
      Begin DmtGridCtl.DmtGrid GrigliaDocumentazione 
         Height          =   5415
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   9551
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
   End
End
Attribute VB_Name = "frmDocumentazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Public LINK_AZIENDA_DOC As Long 'Identificativo dell'azienda
''Public LINK_CLIENTE_DOC As Long 'Identificativo del cliente
''Public LINK_CONTRATTO_DOC As Long 'Identificativo del contratto
''Public LINK_INTERVENTO_PADRE_DOC As Long 'Identificativo dell'intervento padre
''Public LINK_INTERVENTO_DOC As Long 'Identificativo dell'intervento
''Public LINK_STARTUP_DOC As Long 'Identificativo il programma di avvio della funzione
''''1 = Azienda
''''2 = Cliente
''''3 = Contratto
''''4 = Intervento


Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                   (ByVal hwnd As Long, ByVal lpszOp As String, _
                    ByVal lpszFile As String, ByVal lpszParams As String, _
                    ByVal LpszDir As String, ByVal FsShowCmd As Long) _
                    As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long


Private Const SW_SHOWNORMAL = 1
Private Const SE_ERR_FNF = 2&
Private Const SE_ERR_PNF = 3&
Private Const SE_ERR_ACCESSDENIED = 5&
Private Const SE_ERR_OOM = 8&
Private Const SE_ERR_DLLNOTFOUND = 32&
Private Const SE_ERR_SHARE = 26&
Private Const SE_ERR_ASSOCINCOMPLETE = 27&
Private Const SE_ERR_DDETIMEOUT = 28&
Private Const SE_ERR_DDEFAIL = 29&
Private Const SE_ERR_DDEBUSY = 30&
Private Const SE_ERR_NOASSOC = 31&
Private Const ERROR_BAD_FORMAT = 11&
Private Const CSIDL_COMMON_APPDATA = &H23
Private Const MAX_PATH = 260

Private rsGriglia As ADODB.Recordset

Private LINK_STARTUP_DOC_LOCAL As Long


Private Sub GET_DOCUMENTAZIONE(IDStartUpFunzione As Long, Generale As Long)
'La variabile IDStartUpFunzione indica quale pulsante è stato cliccato e con
'la variabile globale LINK_STARTUP_DOC possiamo scrivere una vista ad hoc in base alle nostre necessità
Dim GET_SQL As String

Select Case LINK_STARTUP_DOC
    
    Case 1 'VISUALIZZATO DAI PARAMETRI AZIENDA
        
        Select Case IDStartUpFunzione
            Case 1 'Pulsante azienda
                GET_SQL = "WHERE IDAzienda=" & LINK_AZIENDA_DOC
                GET_SQL = GET_SQL & " AND IDAnagrafica=0"
                GET_SQL = GET_SQL & " AND IDRV_POContratto=0"
                GET_SQL = GET_SQL & " AND IDRV_POInterventoPadre=0"
                GET_SQL = GET_SQL & " AND IDRV_POIntervento=0"
                GET_SQL = GET_SQL & " ORDER BY NomeFile"
            Case 2 'Pulsante cliente
                Exit Sub
            Case 3 'Pulsante contratto
                Exit Sub
            Case 4 'Pulsante Intervento
                Exit Sub
            Case 5 'Pulsante fase intervento
                Exit Sub
        End Select
    Case 2 'VISUALIZZATO DALLA CONFIGURAZIONE DEL CLIENTE
        
        Select Case IDStartUpFunzione
            Case 1 'Pulsante azienda
                GET_SQL = "WHERE IDAzienda=" & LINK_AZIENDA_DOC
                GET_SQL = GET_SQL & " AND IDAnagrafica=0"
                GET_SQL = GET_SQL & " AND IDRV_POContratto=0"
                GET_SQL = GET_SQL & " AND IDRV_POInterventoPadre=0"
                GET_SQL = GET_SQL & " AND IDRV_POIntervento=0"
                GET_SQL = GET_SQL & " ORDER BY NomeFile"
            Case 2 'Pulsante cliente
                GET_SQL = "WHERE IDAzienda=" & LINK_AZIENDA_DOC
                GET_SQL = GET_SQL & " AND IDAnagrafica=" & LINK_CLIENTE_DOC
                If Generale = 1 Then
                    GET_SQL = GET_SQL & " AND IDRV_POContratto=0"
                    GET_SQL = GET_SQL & " AND IDRV_POInterventoPadre=0"
                    GET_SQL = GET_SQL & " AND IDRV_POIntervento=0"
                End If
                GET_SQL = GET_SQL & " ORDER BY NomeFile"
            Case 3 'Pulsante contratto
                GET_SQL = "WHERE IDAzienda=" & LINK_AZIENDA_DOC
                GET_SQL = GET_SQL & " AND IDAnagrafica=" & LINK_CLIENTE_DOC
                GET_SQL = GET_SQL & " AND IDRV_POContratto>0"
                If Generale = 1 Then
                    GET_SQL = GET_SQL & " AND IDRV_POInterventoPadre=0"
                    GET_SQL = GET_SQL & " AND IDRV_POIntervento=0"
                    
                End If
                GET_SQL = GET_SQL & " ORDER BY TipoContratto, NomeFile"
            Case 4 'Pulsante Intervento
                Exit Sub
            Case 5 'Pulsante fase intervento
                Exit Sub
        End Select
    Case 3 'VISUALIZZATO DAL CONTRATTO
        Select Case IDStartUpFunzione
            
            Case 1 'Pulsante azienda
                GET_SQL = "WHERE IDAzienda=" & LINK_AZIENDA_DOC
                GET_SQL = GET_SQL & " AND IDAnagrafica=0"
                GET_SQL = GET_SQL & " AND IDRV_POContratto=0"
                GET_SQL = GET_SQL & " AND IDRV_POInterventoPadre=0"
                GET_SQL = GET_SQL & " AND IDRV_POIntervento=0"
                GET_SQL = GET_SQL & " ORDER BY NomeFile"
            Case 2 'Pulsante cliente
                GET_SQL = "WHERE IDAzienda=" & LINK_AZIENDA_DOC
                GET_SQL = GET_SQL & " AND IDAnagrafica=" & LINK_CLIENTE_DOC
                If Generale = 1 Then
                    GET_SQL = GET_SQL & " AND IDRV_POContratto=0"
                    GET_SQL = GET_SQL & " AND IDRV_POInterventoPadre=0"
                    GET_SQL = GET_SQL & " AND IDRV_POIntervento=0"
                End If
                GET_SQL = GET_SQL & " ORDER BY NomeFile"

            Case 3 'Pulsante contratto
                GET_SQL = "WHERE IDAzienda=" & LINK_AZIENDA_DOC
                GET_SQL = GET_SQL & " AND IDAnagrafica=" & LINK_CLIENTE_DOC
                GET_SQL = GET_SQL & " AND IDRV_POContratto=" & LINK_CONTRATTO_DOC
                If Generale = 1 Then
                    GET_SQL = GET_SQL & " AND IDRV_POInterventoPadre=0"
                    GET_SQL = GET_SQL & " AND IDRV_POIntervento=0"
                End If
                GET_SQL = GET_SQL & " ORDER BY TipoContratto, NomeFile"

            Case 4 'Pulsante Intervento
                Exit Sub
            Case 5 'Pulsante fase intervento
                Exit Sub
        End Select
        
    Case 4 'VISUALIZZATO DALL'INTERVENTO
        Select Case IDStartUpFunzione
           
            Case 1 'Pulsante azienda
                GET_SQL = "WHERE IDAzienda=" & LINK_AZIENDA_DOC
                GET_SQL = GET_SQL & " AND IDAnagrafica=0"
                GET_SQL = GET_SQL & " AND IDRV_POContratto=0"
                GET_SQL = GET_SQL & " AND IDRV_POInterventoPadre=0"
                GET_SQL = GET_SQL & " AND IDRV_POIntervento=0"
                GET_SQL = GET_SQL & " ORDER BY NomeFile"
            Case 2 'Pulsante cliente
                GET_SQL = "WHERE IDAzienda=" & LINK_AZIENDA_DOC
                GET_SQL = GET_SQL & " AND IDAnagrafica=" & LINK_CLIENTE_DOC
                If Generale = 1 Then
                    GET_SQL = GET_SQL & " AND IDRV_POContratto=0"
                    GET_SQL = GET_SQL & " AND IDRV_POInterventoPadre=0"
                    GET_SQL = GET_SQL & " AND IDRV_POIntervento=0"
                    
                End If
                GET_SQL = GET_SQL & " ORDER BY NomeFile"
            Case 3 'Pulsante contratto
                GET_SQL = "WHERE IDAzienda=" & LINK_AZIENDA_DOC
                GET_SQL = GET_SQL & " AND IDAnagrafica=" & LINK_CLIENTE_DOC
                GET_SQL = GET_SQL & " AND IDRV_POContratto=" & LINK_CONTRATTO_DOC
                If Generale = 1 Then
                    GET_SQL = GET_SQL & " AND IDRV_POInterventoPadre=0"
                    GET_SQL = GET_SQL & " AND IDRV_POIntervento=0"
                End If
                GET_SQL = GET_SQL & " ORDER BY NomeFile"

            Case 4 'Pulsante Intervento
                GET_SQL = "WHERE IDAzienda=" & LINK_AZIENDA_DOC
                GET_SQL = GET_SQL & " AND IDAnagrafica=" & LINK_CLIENTE_DOC
                'GET_SQL = GET_SQL & " AND IDRV_POContratto=" & LINK_CONTRATTO_DOC
                GET_SQL = GET_SQL & " AND IDRV_POInterventoPadre=" & LINK_INTERVENTO_PADRE_DOC
                'GET_SQL = GET_SQL & " AND IDRV_POIntervento>0"
                GET_SQL = GET_SQL & " ORDER BY AnnoIntervento DESC, NumeroIntervento DESC, NumeroFase DESC, NomeFile"
            Case 5 'Pulsante fase intervento
                GET_SQL = "WHERE IDAzienda=" & LINK_AZIENDA_DOC
                GET_SQL = GET_SQL & " AND IDAnagrafica=" & LINK_CLIENTE_DOC
                'GET_SQL = GET_SQL & " AND IDRV_POContratto=" & LINK_CONTRATTO_DOC
                GET_SQL = GET_SQL & " AND IDRV_POInterventoPadre=" & LINK_INTERVENTO_PADRE_DOC
                GET_SQL = GET_SQL & " AND IDRV_POIntervento=" & LINK_INTERVENTO_DOC
                GET_SQL = GET_SQL & " ORDER BY AnnoIntervento DESC, NumeroIntervento DESC, NumeroFase DESC, NomeFile"

        End Select
        
End Select

GET_GRIGLIA GET_SQL

End Sub
Private Sub GET_GRIGLIA(sSQL_WHERE As String)
'On Error GoTo ERR_GET_GRIGLIA
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
sSQL = "SELECT * FROM RV_POIEDocumentazione "
sSQL = sSQL & sSQL_WHERE
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
    Set rsGriglia = New ADODB.Recordset
    rsGriglia.CursorLocation = adUseClient
    rsGriglia.Open sSQL, Cn.InternalConnection
    
    With Me.GrigliaDocumentazione
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectRow
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDRV_PODocumentazione", "IDRV_PODocumentazione", dgInteger, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDRV_POContratto", "IDRV_POContratto", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDRV_POStoriaContratto", "IDRV_POStoriaContratto", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDRV_POInterventoPadre", "IDRV_POInterventoPadre", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDRV_POIntervento", "IDRV_POIntervento", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "AnagraficaAzienda", "Azienda", dgchar, False, 2500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "AnagraficaCliente", "Cliente", dgchar, True, 2500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "NomeCliente", "Nome cliente", dgchar, True, 2500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "TipoContratto", "Tipo contratto", dgchar, True, 2500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "AnnoIntervento", "Anno intervento", dgInteger, True, 1000, dgAlignRight, True, True, False
            .ColumnsHeader.Add "NumeroIntervento", "N° Int.", dgInteger, True, 1000, dgAlignRight, True, True, False
            .ColumnsHeader.Add "NumeroFase", "N° fase", dgInteger, True, 1000, dgAlignRight, True, True, False
            .ColumnsHeader.Add "NomeFile", "Nome file", dgchar, True, 2500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "TipoFile", "Ext", dgchar, False, 1000, dgAlignleft, True, True, False
            .ColumnsHeader.Add "PercorsoOriginale", "Percorso interno", dgchar, False, 2500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Annotazioni", "Annotazioni", dgchar, False, 2500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "Archiviato", "Archiviato", dgBoolean, False, 2500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "IDUtenteInserimento", "IDUtenteInserimento", dgNumeric, False, 500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "UtenteInserimento", "Utente ins.", dgchar, False, 2500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "MacchinaInserimento", "Macchina ins.", dgchar, False, 2500, dgAlignleft, True, True, False
            .ColumnsHeader.Add "UtenteMacchinaInserimento", "Utente macchina ins.", dgBoolean, False, 2500, dgAlignleft, True, True, False
        Set .Recordset = rsGriglia
        .Refresh
        .LoadUserSettings
    End With

    
    Cn.CursorLocation = OLDCursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "Reperimento dati documentazione"
    
End Sub

Private Sub cmdAzienda_Click()
    Me.Caption = "DOCUMENTAZIONE AZIENDA"
    LINK_STARTUP_DOC_LOCAL = 1
    GET_DOCUMENTAZIONE LINK_STARTUP_DOC_LOCAL, 1
End Sub

Private Sub cmdCaricaDocumento_Click()
On Error GoTo ERR_cmdCaricaDocumento_Click
Dim FileName As String
Dim PercorsoFile As String


     Dim sSave As String
     sSave = Space(255)
     If IsWinNT Then
         GetFileNameFromBrowseW Me.hwnd, StrPtr(sSave), 255, StrPtr(""), StrPtr(""), StrPtr("All files (*.*)" + Chr$(0) + "*.*" + Chr$(0)), StrPtr("Selezionare un file")
     Else
         GetFileNameFromBrowseA Me.hwnd, sSave, 255, "", "", "All files (*.*)" + Chr$(0) + "*.*" + Chr$(0), "Selezionare un file"
     End If
    
     If Len(Trim(sSave)) = 0 Then Exit Sub
     
    FileName = sSave
    
    sSQL = "SELECT * FROM RV_PODocumentazione "
    sSQL = sSQL & "WHERE IDRV_PODocumentazione=0"
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    Screen.MousePointer = 11
    
    'For I = 1 To Data.Files.Count
        rsNew.AddNew
            rsNew!IDRV_PODocumentazione = fnGetNewKey("RV_PODocumentazione", "IDRV_PODocumentazione")
            rsNew!IDAzienda = LINK_AZIENDA_DOC
            rsNew!IDAnagrafica = LINK_CLIENTE_DOC
            rsNew!IDRV_POContratto = LINK_CONTRATTO_DOC
            rsNew!IDRV_POInterventoPadre = LINK_INTERVENTO_PADRE_DOC
            rsNew!IDRV_POIntervento = LINK_INTERVENTO_DOC
            rsNew!PercorsoOriginale = Trim(FileName)
            rsNew!PercorsoCompleto = Trim(FileName)
            rsNew!NomeFile = GET_NOME_FILE(Trim(FileName))
            rsNew!TipoFile = GetExt(rsNew!NomeFile)
            rsNew!Annotazioni = ""
            rsNew!Archiviato = 1
            rsNew!IDUtenteInserimento = LINK_UTENTE
            rsNew!MacchinaInserimento = GET_NOMECOMPUTER
            rsNew!UtenteMacchinaInserimento = GET_NOMEUTENTE
            If Abs(rsNew!Archiviato) = 1 Then
                Set sts = New ADODB.Stream
                sts.Type = ADODB.adTypeBinary
                sts.Open
                sts.LoadFromFile Trim(FileName)
                rsNew!Documento = sts.Read
            End If
        rsNew.Update
    'Next
    
    rsNew.Close
    Set rsNew = Nothing
    
    If Not (sts Is Nothing) Then
        sts.Close
        Set sts = Nothing
    End If
    
    Screen.MousePointer = 0
    
        Select Case LINK_STARTUP_DOC
            Case 1
                cmdAzienda_Click
            Case 2
                cmdCliente_Click
            Case 3
                cmdContratto_Click
            Case 4
                cmdIntervento_Click
        End Select



Exit Sub
ERR_cmdCaricaDocumento_Click:
    MsgBox Err.Description, vbCritical, "cmdCaricaDocumento_Click"
End Sub

Private Sub cmdCliente_Click()
    Me.Caption = "DOCUMENTAZIONE CLIENTE"
    LINK_STARTUP_DOC_LOCAL = 2
    GET_DOCUMENTAZIONE LINK_STARTUP_DOC_LOCAL, 1
End Sub

Private Sub cmdCompleto_Click()
    GET_DOCUMENTAZIONE LINK_STARTUP_DOC_LOCAL, 0
End Sub

Private Sub cmdContratto_Click()
    Me.Caption = "DOCUMENTAZIONE CONTRATTO"
    LINK_STARTUP_DOC_LOCAL = 3
    GET_DOCUMENTAZIONE LINK_STARTUP_DOC_LOCAL, 1
End Sub

Private Sub cmdIntervento_Click()
    Me.Caption = "DOCUMENTAZIONE FASE INTERVENTO"
    LINK_STARTUP_DOC_LOCAL = 5
    GET_DOCUMENTAZIONE LINK_STARTUP_DOC_LOCAL, 1
End Sub

Private Sub cmdInterventoPadre_Click()
    Me.Caption = "DOCUMENTAZIONE INTERVENTO"
    LINK_STARTUP_DOC_LOCAL = 4
    GET_DOCUMENTAZIONE LINK_STARTUP_DOC_LOCAL, 1
End Sub

Private Sub Form_Load()
    GET_CONTROLLO_PULSANTI

    Select Case LINK_STARTUP_DOC
        Case 1
            cmdAzienda_Click
        Case 2
            cmdCliente_Click
        Case 3
            cmdContratto_Click
        Case 4
            cmdIntervento_Click
    End Select
    
End Sub


Private Sub GrigliaDocumentazione_DblClick()
On Error GoTo ERR_GrigliaDocumentazione_DblClick
Dim NomeCartella As String
Dim IDDocumentazione As Long
Dim NomeFile As String
Dim f As FileSystemObject
If (rsGriglia.EOF = True) And (rsGriglia.BOF = True) Then Exit Sub
    
    IDDocumentazione = fnNotNullN(Me.GrigliaDocumentazione.AllColumns("IDRV_PODocumentazione").Value)
    NomeFile = fnNotNull(Me.GrigliaDocumentazione.AllColumns("NomeFile").Value)
    
    If Me.GrigliaDocumentazione.AllColumns("Archiviato").Value = 1 Then
        NomeCartella = TrovaCartella(CSIDL_COMMON_APPDATA) & "Contratti Professional\"
        Set f = New FileSystemObject
            If f.FolderExists(NomeCartella) = False Then
                f.CreateFolder NomeCartella
            End If
        Set f = Nothing
    Else
        NomeCartella = fnNotNull(Me.GrigliaDocumentazione.AllColumns("PercorsoOriginale").Value)
    End If
    
    AVVIA_FILE IDDocumentazione, NomeCartella, NomeFile, fnNotNullN(Me.GrigliaDocumentazione.AllColumns("Archiviato").Value)
Exit Sub
ERR_GrigliaDocumentazione_DblClick:
    MsgBox Err.Description, vbCritical, "GrigliaDocumentazione_DblClick"
End Sub
Public Function TrovaCartella(IDLCartella As Long) As String

    TrovaCartella = String$(MAX_PATH, 0)
    
    Call SHGetSpecialFolderPath(ByVal 0&, TrovaCartella, IDLCartella, ByVal 0&)
    
    TrovaCartella = Left$(TrovaCartella, InStr(1, TrovaCartella, Chr$(0)) - 1)
    
    If Len(TrovaCartella) > 0 And Right$(TrovaCartella, 1) <> "\" Then TrovaCartella = TrovaCartella & "\"
End Function
Private Sub AVVIA_FILE(IDDocumentazione As Long, Percorso As String, NomeFile As String, Archiviato As Long)
On Error GoTo ERR_AVVIA_FILE
Dim rs As ADODB.Recordset
Dim sts As ADODB.Stream
Dim Scr_hDC As Long
Dim X
Dim msg As String
Dim PercorsoDocArchiviati As String

Scr_hDC = GetDesktopWindow()

If Archiviato = 1 Then

    
    sSQL = "SELECT Documento FROM RV_PODocumentazione "
    sSQL = sSQL & "WHERE IDRV_PODocumentazione = " & IDDocumentazione
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection

    Set sts = New ADODB.Stream
    sts.Type = ADODB.adTypeBinary
    sts.Open
    sts.Write rs.Fields("Documento").Value
    sts.SaveToFile Percorso & NomeFile, adSaveCreateOverWrite

    rs.Close
    Set rs = Nothing
    
    sts.Close
    Set sts = Nothing

End If


X = ShellExecute(Scr_hDC, "Open", NomeFile, "", Percorso, SW_SHOWNORMAL)

If X <= 32 Then
    'There was an error
    Select Case X
        Case SE_ERR_FNF
            msg = "File not found"
        Case SE_ERR_PNF
            msg = "Path not found"
        Case SE_ERR_ACCESSDENIED
            msg = "Access denied"
        Case SE_ERR_OOM
            msg = "Out of memory"
        Case SE_ERR_DLLNOTFOUND
            msg = "DLL not found"
        Case SE_ERR_SHARE
            msg = "A sharing violation occurred"
        Case SE_ERR_ASSOCINCOMPLETE
            msg = "Incomplete or invalid file association"
        Case SE_ERR_DDETIMEOUT
            msg = "DDE Time out"
        Case SE_ERR_DDEFAIL
            msg = "DDE transaction failed"
        Case SE_ERR_DDEBUSY
            msg = "DDE busy"
        Case SE_ERR_NOASSOC
            msg = "No association for file extension"
        Case ERROR_BAD_FORMAT
            msg = "Invalid EXE file or error in EXE image"
        Case Else
            msg = "Unknown error"
    End Select
    
    'MsgBox msg, vbInformation, "Apertura file"

End If
Exit Sub
ERR_AVVIA_FILE:
    MsgBox Err.Description, vbCritical, "AVVIA_FILE"
End Sub
Private Sub GET_CONTROLLO_PULSANTI()
    
    Select Case LINK_STARTUP_DOC
        Case 1 'DA AZIENDA
            Me.cmdCliente.Enabled = False
            Me.cmdContratto.Enabled = False
            Me.cmdInterventoPadre.Enabled = False
            Me.cmdIntervento.Enabled = False
    
        Case 2 'DA CLIENTE
            Me.cmdInterventoPadre.Enabled = False
            Me.cmdIntervento.Enabled = False
            
        Case 3 'DA CONTRATTO
            Me.cmdInterventoPadre.Enabled = False
            Me.cmdIntervento.Enabled = False
        
        Case 4 'DA INTERVENTO
            Me.cmdCliente.Enabled = True
            Me.cmdContratto.Enabled = True
            Me.cmdInterventoPadre.Enabled = True
            Me.cmdIntervento.Enabled = True
            If LINK_CONTRATTO_DOC = 0 Then
                Me.cmdContratto.Enabled = False
            End If
    End Select
End Sub
Private Function GET_NOME_FILE(Percorso As String)
On Error GoTo ERR_GET_NOME_FILE
Dim ArrayPercorso() As String

ArrayPercorso = Split(Percorso, "\")

GET_NOME_FILE = ArrayPercorso(UBound(ArrayPercorso))

Exit Function
ERR_GET_NOME_FILE:
    GET_NOME_FILE = ""
End Function
Public Function GetExt(ByVal sPath As String) As String
   GetExt = Mid$(GetStrFromPtr(PathFindExtension(sPath)), 2)
End Function

Private Function GetStrFromPtr(ByVal lpsz As Long) As String
   GetStrFromPtr = String$(lstrlen(ByVal lpsz), 0)
   lstrcpy ByVal GetStrFromPtr, ByVal lpsz
End Function

Public Function GET_NOMECOMPUTER() As String
Dim dwLen As Long
Dim strString As String
Const MAX_COMPUTERNAME_LENGTH As Long = 31
    
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    'Show the computer name
    GET_NOMECOMPUTER = strString
End Function

Function GET_NOMEUTENTE() As String
    Dim strString As String
    Dim lunghezzaStringa As Long
    lunghezzaStringa = 32
    strString = String(lunghezzaStringa, " ")
    GetUserName strString, lunghezzaStringa
    strString = Left(strString, lunghezzaStringa)
    GET_NOMEUTENTE = strString
    GET_NOMEUTENTE = Mid(GET_NOMEUTENTE, 1, Len(GET_NOMEUTENTE) - 1)
End Function
Public Function IsWinNT() As Boolean
     Dim myOS As OSVERSIONINFO
     myOS.dwOSVersionInfoSize = Len(myOS)
     GetVersionEx myOS
     IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function
