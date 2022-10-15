VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.3#0"; "DmtGridCtl.ocx"
Begin VB.Form frmVisualizzazione 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CREAZIONE CONTATTI   (Passo 3 di 3)"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid GrigliaDMT 
      Height          =   4695
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8281
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   4800
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
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
      Left            =   9240
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
   End
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
      Left            =   8040
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdIIndietro 
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
      Left            =   6840
      TabIndex        =   1
      Top             =   5280
      Width           =   1095
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
      Left            =   5640
      TabIndex        =   0
      Top             =   5280
      Width           =   1095
   End
End
Attribute VB_Name = "frmVisualizzazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsDMT As DmtOleDbLib.adoResultset

Private Sub cmdAnnulla_Click()
    If MsgBox("Vuoi abbandonare la procedura?", vbInformation + vbYesNo, "Chiusura applicazioni") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdFine_Click()
On Error Resume Next
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
'''''''VARIABILI PER OUTLOOK''''''''''''''''''''
    Dim Out 'As Outlook.Application
    Dim Name 'As Outlook.NameSpace
    Dim f 'As Outlook.MAPIFolder 'MAIN MAPIFOLDER
    Dim G 'As Outlook.MAPIFolder 'CONTATCT MAPIFOLDER
    Dim H 'As Outlook.MAPIFolder 'CARTELLA POINT MAPIFOLDER
    Dim Rubrica ' As Outlook.AddressList
    Dim Cont 'As Outlook.ContactItem
    Dim RCP 'As Outlook.Recipient
    Dim Lista 'As Outlook.DistListItem
    Dim NomeUtente 'As Outlook.AddressEntries
    
    
    
    
    Dim NumeroContattiOLD As Long
    Dim sContactFolder As String
    Dim EmailCount As Long
    Dim ToSave As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case optOperazione
    
        Case 0
            sSQL = "SELECT Anagrafica, Cap, IDSitoPerAnagrafica, Indirizzo, Comune, Provincia, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Fax, Email, Nominativo "
            sSQL = sSQL & "FROM RV_POTMPContact "
            sSQL = sSQL & "GROUP BY IDAnagrafica, Anagrafica, Cap, IDSitoPerAnagrafica, Indirizzo, Comune, Provincia, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Fax, Email, Nominativo "
        Case 1
            sSQL = "SELECT Anagrafica, Cap, IDSitoPerAnagrafica, Indirizzo, Comune, Provincia, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Email, Fax ,Nominativo "
            sSQL = sSQL & "FROM RV_POTMPContact "
            sSQL = sSQL & "GROUP BY IDAnagrafica, Anagrafica, Cap, IDSitoPerAnagrafica, Indirizzo, Comune, Provincia, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Email, Fax, Nominativo "
                    
        Case 2
            sSQL = "SELECT Anagrafica, Cap, IDSitoPerAnagrafica, Indirizzo, Comune, Provincia, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Nominativo "
            sSQL = sSQL & "FROM RV_POTMPContact "
            sSQL = sSQL & "GROUP BY IDAnagrafica, Anagrafica, Cap, IDSitoPerAnagrafica, Indirizzo, Comune, Provincia, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Nominativo "
    End Select
        
    Set rs = CnDMT.OpenResultset(sSQL)
    
    Set Out = CreateObject("Outlook.Application")
    Set Name = Out.GetNamespace("MAPI")
    sMainFolder = "Cartelle personali"
    sContactFolder = Name.GetDefaultFolder(10)

    Set f = Name.Folders(1)

    Set G = f.Folders(sContactFolder)

    Set H = G.Folders("CONTACT POINT")

    If H Is Nothing Then
        G.Folders.Add "CONTACT POINT"
        Set H = G.Folders("CONTACT POINT")
        H.ShowAsOutlookAB = True
    Else
        Set H = G.Folders("CONTACT POINT")
        NumeroContattiOLD = H.Items.Count
        For i = 1 To NumeroContattiOLD
            H.Items.Remove 1
        Next
   
        'For Each Cont In H.Items
        '    email = Cont.Email1Address
        '    Cont.Delete
        'Next
    End If
    
    
    If rs.EOF = False Then
    Me.ProgressBar1.Max = NumeroRecord * NumeroRecord
    Me.ProgressBar1.Value = 0
    EmailCount = 1
    While Not rs.EOF
            If IsNull(rs!Nominativo) = False Then
            
                Set Cont = Out.CreateItem(2)
                Cont.FirstName = rs!Nominativo
                Cont.FullName = rs!Nominativo
                Cont.CompanyName = rs!Anagrafica
                If IDSitoPerAnagrafica > 0 Then
                    Cont.CompanyName = Cont.CompanyName & " filiale di " & rs!SitoPerAnagrafica
                End If
                If optOperazione = 0 Then
                    If IsNull(rs!fax) Then
                        ToSave = False
                    Else
                        If Len(rs!fax) > 0 Then
                            Cont.BusinessFaxNumber = "+39 " & rs!fax
                            'Cont.Email1Address = "X@X"
                            ToSave = True
                        Else
                            ToSave = False
                        End If
                    End If
                End If
                
                If optOperazione = 1 Then
                    If IsNull(rs!Email) Then
                        ToSave = False
                    Else
                        If Len(rs!Email) > 0 Then
                            Cont.Email1Address = rs!Email
                            ToSave = True
                        Else
                            ToSave = False
                        End If
                    End If
                End If
                
                
                
               
                'Cont.Email1DisplayName = rs!nominativo
                                
                If ToSave = True Then Cont.Save
                If ToSave = True Then
                    Set CopyContact = Cont.Copy
                    CopyContact.Move H
                    Cont.Delete
                End If

                
            End If
        EmailCount = EmailCount + 1
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + NumeroRecord
        rs.MoveNext
        Wend
    
    End If
    
    If Me.ProgressBar1.Value < Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    End If
    
    Set Lista = Out.CreateItem(7)
    
    For i = 1 To H.Items.Count
       
        
        If optOperazione = 0 Then
            v = Name.AddressLists.Item("CONTACT POINT").AddressEntries.Item(i).Name
            
                Set RCP = Out.Session.CreateRecipient(H.Items(i))
                                            
                Risultato = RCP.Resolve
                
                
                Lista.AddMember RCP
                
            
            
        Else
            Set RCP = Out.Session.CreateRecipient(H.Items(i))
            Risultato = RCP.Resolve
            
            Lista.AddMember RCP
        End If
    Next
    
    If optOperazione = 0 Then
        Lista.DLName = "_Fax"
    Else
        Lista.DLName = "_Email"
    End If
    Lista.Save
    If Lista.MemberCount = 0 Then
        testo = "ATTENZIONE!!!" & vbCrLf
        testo = testo & "La lista di distribuzione "
        If optOperazione = 0 Then
            testo = testo & "_Fax"
        Else
            testo = testo & "_Email"
        End If
        testo = testo & " non contiene nessun membro." & vcbrlf
        testo = testo & "Pertanto per il corretto funzionamento della procedura includere a mano i membri della lista creati in " & H.FolderPath
        MsgBox testo, vbInformation, "Creazione contatti"
    End If
    Set CopyContact = Lista.Copy
    CopyContact.Move H
    Lista.Delete
    
   
    'Out.Quit
    Set Out = Nothing
    Set Name = Nothing
    Set Cont = Nothing
    Set f = Nothing
    Set G = Nothing
    Set H = Nothing
    rs.CloseResultset
    Set rs = Nothing
    
    
    cmdFine.Enabled = False
    
    
End Sub

Private Sub cmdIIndietro_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    DinamicNumber = 2
    ControlloComandi
    
    GrigliaContatti
    
End Sub
Private Sub ControlloComandi()
    Me.cmdAvanti.Enabled = False
End Sub
Private Sub GrigliaContatti()
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As DmtGridCtl.dgColumnHeader
    
    Select Case optOperazione
    
        Case 0
            sSQL = "SELECT Anagrafica, Cap, IDSitoPerAnagrafica, Indirizzo, Comune, Provincia, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Fax, Nominativo "
            sSQL = sSQL & "FROM RV_POTMPContact "
            sSQL = sSQL & "GROUP BY IDAnagrafica, Anagrafica, Cap, IDSitoPerAnagrafica, Indirizzo, Comune, Provincia, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Fax, Nominativo "
        Case 1
            sSQL = "SELECT Anagrafica, Cap, IDSitoPerAnagrafica, Indirizzo, Comune, Provincia, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Email, Nominativo "
            sSQL = sSQL & "FROM RV_POTMPContact "
            sSQL = sSQL & "GROUP BY IDAnagrafica, Anagrafica, Cap, IDSitoPerAnagrafica, Indirizzo, Comune, Provincia, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Email, Nominativo "
                    
        Case 2
            sSQL = "SELECT Anagrafica, Cap, IDSitoPerAnagrafica, Indirizzo, Comune, Provincia, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Nominativo "
            sSQL = sSQL & "FROM RV_POTMPContact "
            sSQL = sSQL & "GROUP BY IDAnagrafica, Anagrafica, Cap, IDSitoPerAnagrafica, Indirizzo, Comune, Provincia, Filiale, IndirizzoFiliale, ComuneFiliale, CapFiliale, ProvinciaFiliale, "
            sSQL = sSQL & "Nominativo "
    End Select
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
        Set rsDMT = CnDMT.OpenResultset(sSQL)
            'Set rsEvent = rsDMT.Data
        
        With Me.GrigliaDMT
            .ColumnsHeader.Clear
                    If optOperazione = 0 Then
                        .ColumnsHeader.Add "Anagrafica", "Azienda", dgchar, True, 3000, dgAlignleft
                        .ColumnsHeader.Add "Filiale", "Filiale", dgchar, True, 3000, dgAlignleft
                        .ColumnsHeader.Add "Nominativo", "Nominativo", dgchar, True, 3000, dgAlignleft
                        .ColumnsHeader.Add "Fax", "Fax", dgchar, True, 3000, dgAlignleft
                    End If
                    If optOperazione = 1 Then
                        .ColumnsHeader.Add "Anagrafica", "Azienda", dgchar, True, 3000, dgAlignleft
                        .ColumnsHeader.Add "Filiale", "Filiale", dgchar, True, 3000, dgAlignleft
                        .ColumnsHeader.Add "Nominativo", "Nominativo", dgchar, True, 3000, dgAlignleft
                        .ColumnsHeader.Add "Email", "E-mail", dgchar, True, 3000, dgAlignleft
                    End If
                    If optOperazione = 2 Then
                        .ColumnsHeader.Add "Anagrafica", "Azienda", dgchar, True, 3000, dgAlignleft
                        .ColumnsHeader.Add "Filiale", "Filiale", dgchar, True, 3000, dgAlignleft
                    End If
                    

            Set .Recordset = rsDMT.Data
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.cmdIIndietro.Value = True Then
        frmParametri.Show
        Exit Sub
    End If
    
End Sub
