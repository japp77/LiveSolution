VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPresentazione 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DOCUMENTAZIONE"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FrmPresentazione.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LV 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   11456
      View            =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":0BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":1158
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":16F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":184C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":1C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":1FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":22D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":2724
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":2A3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":2E90
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":32E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":35FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":3756
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":45A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":49FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":6704
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":6A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":6D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":718A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":75DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":78F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":79FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":7E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":8168
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":8D3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":918C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":95DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":9A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":9E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":A2D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":1056E
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":10888
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":10BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":10EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":111D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":11628
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":11BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":1215C
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":126F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":12C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":1322A
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":137C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPresentazione.frx":13D5E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPresentazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private CARTELLA_PROGRAMMA As String

Private Sub Form_Load()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

CARTELLA_PROGRAMMA = ""

If ConnessioneADODBLib = True Then

    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)

    'PrelevaAzienda
    CARTELLA_PROGRAMMA = GET_RECUPERA_NOME_PROGRAMMA
    
    If Len(CARTELLA_PROGRAMMA) > 0 Then
        GET_DOCUMENTAZIONE
    End If
            
            
End If
Me.Caption = Me.Caption & " " & CARTELLA_PROGRAMMA
End Sub

Private Function GET_RECUPERA_NOME_PROGRAMMA() As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT * FROM RV_POProgramma "
sSQL = sSQL & "WHERE IDRV_POProgramma=" & IdentificativoProgramma

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_RECUPERA_NOME_PROGRAMMA = fnNotNull(rs!Programma)
Else
    GET_RECUPERA_NOME_PROGRAMMA = ""
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub GET_DOCUMENTAZIONE()
Dim f As FileSystemObject
Dim Percorso As String

'Lista Help
Set f = New FileSystemObject
Percorso = MenuOptions.ProgramsPath & "\Help_" & CARTELLA_PROGRAMMA
If f.FolderExists(Percorso) Then
    Set VFolder = f.GetFolder(Percorso)
    
    Set VFile = VFolder.Files
    
    For Each CountFile In VFile
        If Right(CountFile.Name, 3) = "pdf" Then
            Me.LV.ListItems.Add , "H_" & CountFile.Name, CountFile.Name, 2, 2
        End If
    Next
End If
End Sub

Private Sub LV_DblClick()


On Error GoTo ERR_FileDocumentazione_DblClick
Dim Scr_hDC As Long
Dim X
Dim msg As String
Scr_hDC = GetDesktopWindow()

X = ShellExecute(Scr_hDC, "Open", Me.LV.SelectedItem.Text, "", MenuOptions.ProgramsPath & "\Help_" & CARTELLA_PROGRAMMA, SW_SHOWNORMAL)

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
    
    MsgBox msg, vbInformation, "Apertura file"
    MsgBox Me.LV.SelectedItem.Text
     
    MsgBox MenuOptions.ProgramsPath & "/Help_" & CARTELLA_PROGRAMMA
End If

Exit Sub

ERR_FileDocumentazione_DblClick:
    MsgBox Err.Description, vbCritical, "Apertura file"
    MsgBox Me.LV.SelectedItem.Text
     
    MsgBox MenuOptions.ProgramsPath & "/Help_" & CARTELLA_PROGRAMMA
End Sub
