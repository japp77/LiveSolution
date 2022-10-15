VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmCoda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Coda dei salvataggi"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9045
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
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   360
      Top             =   2520
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   255
      Left            =   7200
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
   End
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4260
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
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2310
      Left            =   0
      Picture         =   "frmCoda.frx":0000
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmCoda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsArt As DmtOleDbLib.adoResultset

Private Sub SettaggioGriglia()
On Error GoTo ERR_SettaggioGriglia
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
    sSQL = "SELECT RV_POTMP.IDSessione, RV_POTMP.IDUtente, RV_POTMP.Utente, RV_POTMP.IDTipoOggetto, RV_POTMP.IDOggetto, TipoOggetto.Oggetto "
    sSQL = sSQL & "FROM RV_POTMP LEFT OUTER JOIN "
    sSQL = sSQL & "TipoOggetto ON RV_POTMP.IDTipoOggetto = TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "ORDER BY IDSessione, IDUtente "
    
    Set rsArt = CnDMT.OpenResultset(sSQL)
        Set rsEvent = rsArt.Data
    
    With Me.Griglia
        .ColumnsHeader.Clear
            .ColumnsHeader.Add "IDUtente", "IDUtente", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDSessione", "Sessione", dgNumeric, True, 1000, dgAlignleft
            .ColumnsHeader.Add "Utente", "Utente", dgchar, True, 3000, dgAlignleft
            .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgNumeric, False, 1000, dgAlignleft
            .ColumnsHeader.Add "Oggetto", "Oggetto", dgchar, True, 3000, dgAlignleft
            
            
        Set .Recordset = rsArt.Data
        .Refresh
    End With
    
    CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_SettaggioGriglia:
    MsgBox Err.Description, vbCritical, "Settaggio griglia"
End Sub

Private Sub cmdElimina_Click()
On Error GoTo ERR_cmdElimina_Click
Dim sSQL As String
Dim Testo As String

'If TheApp.IDUser <> Me.Griglia.AllColumns("IDUtente") Then
    Me.Timer1.Enabled = False
    Testo = "ATTENZIONE!!!!" & vbCrLf
    Testo = Testo & "Eliminando l'utente dalla coda di salvataggio si potrebbero avere degli effetti indesiderati." & vbCrLf
    Testo = Testo & "Continuare con il comando di eliminazione?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione dati di coda") = vbYes Then
        sSQL = "DELETE FROM RV_POTMP "
        sSQL = sSQL & "WHERE IDUtente=" & fnNotNullN(Me.Griglia("IDUtente").Value)
        sSQL = sSQL & " AND IDSessione=" & fnNotNullN(Me.Griglia("IDSessione").Value)
        CnDMT.Execute sSQL
        
        SettaggioGriglia
        
        Me.Timer1.Enabled = True
    End If
'End If


Exit Sub
ERR_cmdElimina_Click:
    MsgBox Err.Description, vbCritical, "cmdElimina_Click"
End Sub

Private Sub Form_Activate()
    SettaggioGriglia
    'Me.cmdElimina.Enabled = False
    DoEvents
    Me.Timer1.Enabled = True
End Sub

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
    If rs.EOF = False Then
        fnGetTipoOggetto = rs!IDTipoOggetto
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyInsert Then
'    If Me.cmdElimina.Enabled = False Then
'        Me.cmdElimina.Enabled = True
'    Else
'        Me.cmdElimina.Enabled = False
'    End If
'End If
End Sub

Private Sub Form_Load()
    Me.Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (rsArt Is Nothing) Then
    rsArt.CloseResultset
    Set rsArt = Nothing
End If

Me.Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    SettaggioGriglia
    DoEvents
End Sub
