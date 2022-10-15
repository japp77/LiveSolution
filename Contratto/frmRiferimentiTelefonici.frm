VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.7#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRiferimentiTelefonici 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Riferimenti telefonici"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4035
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
   ScaleHeight     =   4605
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraNuovoContatto 
      Caption         =   "Nuovo contatto"
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
      Height          =   4575
      Left            =   4080
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CommandButton cmdVisGriglia 
         Height          =   320
         Left            =   3360
         Picture         =   "frmRiferimentiTelefonici.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Nuovo contatto"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdElimina 
         Caption         =   "ELIMINA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   38
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton cmdConfermaContatto 
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
         Height          =   255
         Left            =   8520
         TabIndex        =   37
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txtPosizioneAziendale 
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox TxtNominativo 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtAnnotazioni 
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3960
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdProprieta 
         Height          =   320
         Left            =   2880
         Picture         =   "frmRiferimentiTelefonici.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Gestione proprieta"
         Top             =   0
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3735
         Left            =   3960
         TabIndex        =   30
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   6588
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   16711680
         ForeColorFixed  =   16777215
         BackColorBkg    =   -2147483633
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   1
         BorderStyle     =   0
      End
      Begin DMTDataCmb.DMTCombo cboFiliale 
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   3615
         _ExtentX        =   6376
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
      Begin VB.Line Line1 
         X1              =   3840
         X2              =   3840
         Y1              =   240
         Y2              =   4440
      End
      Begin VB.Label Label2 
         Caption         =   "Posizione aziendale"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Nominativo"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Annotazioni"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Altra sede"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   4095
      End
   End
   Begin VB.Frame FraTrovaContatto 
      Caption         =   "Trova contatto"
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
      Height          =   3855
      Left            =   4080
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox txtPosizioneRic 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   3015
      End
      Begin DMTDataCmb.DMTCombo cboFilialeRic 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   3015
         _ExtentX        =   5318
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
      Begin VB.TextBox txtNominativoRic 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Posizione"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Filiale"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Nominativo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3015
      End
   End
   Begin DmtGridCtl.DmtGrid GrigliaTel 
      Height          =   4575
      Left            =   4080
      TabIndex        =   15
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8070
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
   Begin VB.Frame FraTelAna 
      Caption         =   "Recapiti principali"
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
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      Begin VB.CommandButton cmdNuovoContatto 
         Height          =   320
         Left            =   2400
         Picture         =   "frmRiferimentiTelefonici.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Nuovo contatto"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdTrovaContatti 
         Height          =   320
         Left            =   2880
         Picture         =   "frmRiferimentiTelefonici.frx":0E9E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Trova contatto"
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdTelContatti 
         Height          =   320
         Left            =   3360
         Picture         =   "frmRiferimentiTelefonici.frx":1428
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Contatti"
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtEmailInternet 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox txtAltriFax 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox txtEmailInterna 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtAltriTelefono 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtTelefono 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Sito web"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Altri fax"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "E-mail interna"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Fax"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Altri telefoni"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Telefono"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtIndirizzo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Height          =   765
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmRiferimentiTelefonici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGrigliaTel As ADODB.Recordset
Private CONST_WIDTH_FORM As Long
Private LINK_DETTAGLIO_CONTATTO As Long
Private BLoading As Boolean



Private Sub cboFilialeRic_Click()
GET_GRIGLIA
End Sub

Private Sub cmdConfermaContatto_Click()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim NomeProprieta As String

sSQL = "SELECT * FROM RV_POContactDettaglio "
sSQL = sSQL & "WHERE IDRV_POContactDettaglio=" & LINK_DETTAGLIO_CONTATTO

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
    rs!IDRV_POContactDettaglio = fnGetNewKey("RV_POContactDettaglio", "IDRV_POContactDettaglio")
    rs!IDAnagrafica = LINK_ANAGRAFICA_TEL
End If
rs!Nominativo = Me.TxtNominativo.Text
rs!IDSitoPerAnagrafica = Me.cboFiliale.CurrentID
rs!Posizione = Me.txtPosizioneAziendale.Text
rs!Annotazioni = Me.txtAnnotazioni.Text

rs.Update

AGGIORNA_PROPRIETA_CONTATTO rs!IDRV_POContactDettaglio, LINK_ANAGRAFICA_TEL, LINK_ANAGRAFICA_TEL, Me.cboFiliale.CurrentID


rs.Close
Set rs = Nothing

CREA_RECORDSET_TMP
INSERISCI_RECAPITI_CLIENTE LINK_ANAGRAFICA_TEL, LINK_SITO_PER_ANAGRAFICA_TEL
GET_GRIGLIA

Me.GrigliaTel.Visible = True
Me.FraNuovoContatto.Visible = False
End Sub

Private Sub cmdElimina_Click()
Dim Testo As String
Dim sSQL As String

If LINK_DETTAGLIO_CONTATTO = 0 Then Exit Sub
Testo = "Sei sicuro di voler eliminare il nominativo?"

If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione nominativo") = vbNo Then Exit Sub

sSQL = "DELETE FROM RV_POContactDettaglio "
sSQL = sSQL & " WHERE IDRV_POContactDettaglio=" & LINK_DETTAGLIO_CONTATTO
Cn.Execute sSQL


CREA_RECORDSET_TMP
INSERISCI_RECAPITI_CLIENTE LINK_ANAGRAFICA_TEL, LINK_SITO_PER_ANAGRAFICA_TEL
GET_GRIGLIA

Me.FraNuovoContatto.Visible = False
Me.GrigliaTel.Visible = True

End Sub

Private Sub cmdNuovoContatto_Click()
    If Me.FraTrovaContatto.Visible = True Then Exit Sub
    
    If Me.Width = CONST_WIDTH_FORM Then
        cmdTelContatti_Click
    End If
    If Me.FraNuovoContatto.Visible = False Then
        Me.FraNuovoContatto.Visible = True
        Me.GrigliaTel.Visible = False
    Else
        If LINK_DETTAGLIO_CONTATTO = 0 Then
            Me.FraNuovoContatto.Visible = False
            Me.GrigliaTel.Visible = True
        Else
            Me.FraNuovoContatto.Visible = True
            Me.GrigliaTel.Visible = False
        End If
    End If
    Form_Resize
    LINK_DETTAGLIO_CONTATTO = 0
    GET_PROPRIETA_CONTATTO LINK_DETTAGLIO_CONTATTO
    GET_CARATTERISTICHE_RISORSA LINK_ANAGRAFICA_TEL, LINK_SITO_PER_ANAGRAFICA_TEL, LINK_DETTAGLIO_CONTATTO
    If LINK_SITO_PER_ANAGRAFICA_TEL > 0 Then
        Me.cboFiliale.WriteOn LINK_SITO_PER_ANAGRAFICA_TEL
    End If
    If Me.FraNuovoContatto.Visible = True Then
        Me.TxtNominativo.SetFocus
    End If
    
End Sub

Private Sub cmdProprieta_Click()
    frmProprieta.Show vbModal

    GET_CARATTERISTICHE_RISORSA LINK_ANAGRAFICA_TEL, LINK_SITO_PER_ANAGRAFICA_TEL, LINK_DETTAGLIO_CONTATTO

    
End Sub

Private Sub cmdTelContatti_Click()
    If Me.Width = CONST_WIDTH_FORM Then
        Me.Width = frmMain.PicForm2.Width
        Me.FraNuovoContatto.Visible = False
        Me.FraTrovaContatto.Visible = False
        Me.GrigliaTel.Visible = True
    Else
        Me.Width = CONST_WIDTH_FORM
        Me.FraNuovoContatto.Visible = False
        Me.FraTrovaContatto.Visible = False
        Me.GrigliaTel.Visible = True
    End If
    
    Form_Resize
End Sub

Private Sub cmdTrovaContatti_Click()
    If Me.FraNuovoContatto.Visible = True Then Exit Sub
    If Me.Width = CONST_WIDTH_FORM Then
        cmdTelContatti_Click
    End If
    If Me.FraTrovaContatto.Visible = False Then
        Me.FraTrovaContatto.Visible = True
    Else
        Me.FraTrovaContatto.Visible = False
    End If
    Form_Resize
End Sub

Private Sub cmdVisGriglia_Click()
    Me.FraNuovoContatto.Visible = False
    Me.GrigliaTel.Visible = True
End Sub

Private Sub Form_Load()
    Me.Left = frmMain.PicForm2.Left
    Me.Top = frmMain.PicForm2.Top
    
    
    
    CONST_WIDTH_FORM = Me.Width
    
    GET_TEL_ANA LINK_ANAGRAFICA_TEL, LINK_SITO_PER_ANAGRAFICA_TEL
    
    INIT_CONTROLLI
    CREA_RECORDSET_TMP
    INSERISCI_RECAPITI_CLIENTE LINK_ANAGRAFICA_TEL, LINK_SITO_PER_ANAGRAFICA_TEL
    If LINK_SITO_PER_ANAGRAFICA_TEL > 0 Then
        Me.cboFilialeRic.WriteOn LINK_SITO_PER_ANAGRAFICA_TEL
    End If
    
    GET_GRIGLIA
    
    
End Sub
Private Sub GET_TEL_ANA(IDAnagrafica As Long, IDSitoPerAnagrafica As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Provincia As String

If IDSitoPerAnagrafica = 0 Then
    sSQL = "SELECT Comune.Comune, Provincia.NomeProvincia, Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome, Anagrafica.Indirizzo, Anagrafica.Cap, "
    sSQL = sSQL & "Anagrafica.Telefono, Anagrafica.Fax, Anagrafica.EMailInterno, Anagrafica.EMailInternet, Anagrafica.AltriTelefoni, Anagrafica.AltriFax, "
    sSQL = sSQL & "Anagrafica.Referente , Nazione.Nazione, Anagrafica.IDNazione, Provincia.Provincia "
    sSQL = sSQL & "FROM Provincia RIGHT OUTER JOIN "
    sSQL = sSQL & "Comune ON Provincia.IDProvincia = Comune.IDProvincia RIGHT OUTER JOIN "
    sSQL = sSQL & "Anagrafica ON Comune.IDComune = Anagrafica.IDComune LEFT OUTER JOIN "
    sSQL = sSQL & "Nazione ON Anagrafica.IDNazione = Nazione.IDNazione "
    sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
Else
    sSQL = "SELECT SitoPerAnagrafica.IDSitoPerAnagrafica, SitoPerAnagrafica.SitoPerAnagrafica, SitoPerAnagrafica.Indirizzo, SitoPerAnagrafica.Cap, "
    sSQL = sSQL & "SitoPerAnagrafica.Telefono, SitoPerAnagrafica.Fax, SitoPerAnagrafica.Referente, SitoPerAnagrafica.Email, SitoPerAnagrafica.IDComune,"
    sSQL = sSQL & "COMUNE.COMUNE , PROVINCIA.PROVINCIA "
    sSQL = sSQL & "FROM SitoPerAnagrafica LEFT OUTER JOIN "
    sSQL = sSQL & "Comune ON SitoPerAnagrafica.IDComune = Comune.IDComune LEFT OUTER JOIN "
    sSQL = sSQL & "Provincia ON Comune.IDProvincia = Provincia.IDProvincia "
    sSQL = sSQL & " WHERE IDSitoPerAnagrafica=" & IDAnagrafica
End If

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtTelefono.Text = ""
    Me.txtAltriTelefono.Text = ""
    Me.txtFax.Text = ""
    Me.txtAltriFax.Text = ""
    Me.txtEmailInterna.Text = ""
    Me.txtEmailInternet.Text = ""
Else
    If IDTipoAna > 0 Then
        Me.txtIndirizzo.Text = fnNotNull(rs!INDIRIZZO) & vbCrLf
        Me.txtIndirizzo.Text = Me.txtIndirizzo.Text & fnNotNull(rs!CAP) & " " & fnNotNull(rs!Comune)
        If Len(Trim(fnNotNull(rs!Provincia))) > 0 Then
            Me.txtIndirizzo.Text = Me.txtIndirizzo.Text & " (" & fnNotNull(rs!Provincia) & ")"
        End If
        Me.txtTelefono.Text = fnNotNull(rs!Telefono)
        Me.txtAltriTelefono.Text = fnNotNull(rs!AltriTelefoni)
        Me.txtFax.Text = fnNotNull(rs!Fax)
        Me.txtAltriFax.Text = fnNotNull(rs!AltriFax)
        Me.txtEmailInterna.Text = fnNotNull(rs!EMailInterno)
        Me.txtEmailInternet.Text = fnNotNull(rs!EMailInternet)
    Else
        Me.txtIndirizzo.Text = fnNotNull(rs!INDIRIZZO) & vbCrLf
        Me.txtIndirizzo.Text = Me.txtIndirizzo.Text & fnNotNull(rs!CAP) & " " & fnNotNull(rs!Comune)
       
        If Len(Trim(fnNotNull(rs!Provincia))) > 0 Then
            Me.txtIndirizzo.Text = Me.txtIndirizzo.Text & " (" & fnNotNull(rs!Provincia) & ")"
        End If
        Me.txtTelefono.Text = fnNotNull(rs!Telefono)
        Me.txtFax.Text = fnNotNull(rs!Fax)
        Me.txtEmailInterna.Text = fnNotNull(rs!EMailInterno)
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub CREA_RECORDSET_TMP()
Dim sSQL As String
Dim I As Integer
Dim rs As DmtOleDbLib.adoResultset


If Not (rsGrigliaTel Is Nothing) Then
    If rsGrigliaTel.State > 0 Then
        rsGrigliaTel.Close
    End If
    
    Set rsGrigliaTel = Nothing
End If

Set rsGrigliaTel = New ADODB.Recordset

rsGrigliaTel.Fields.Append "IDRV_POContactDettaglio", adInteger, , adFldIsNullable
rsGrigliaTel.Fields.Append "IDAnagrafica", adInteger, , adFldIsNullable
rsGrigliaTel.Fields.Append "IDSitoPerAnagrafica", adInteger, , adFldIsNullable
rsGrigliaTel.Fields.Append "SitoPerAnagrafica", adVarChar, 250, adFldIsNullable
rsGrigliaTel.Fields.Append "Nominativo", adVarChar, 250, adFldIsNullable
rsGrigliaTel.Fields.Append "Posizione", adVarChar, 250, adFldIsNullable

sSQL = "SELECT * FROM RV_POContactProprieta "
sSQL = sSQL & "WHERE Gestisci=1"
Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    rsGrigliaTel.Fields.Append fnNotNull(rs!NomeCampo), adVarChar, 250, adFldIsNullable
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

rsGrigliaTel.Open , , adOpenKeyset, adLockBatchOptimistic

End Sub
Private Sub INSERISCI_RECAPITI_CLIENTE(IDAnagraficaCliente As Long, IDSitoPerAnagrafica As Long)
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim ICont As Long
Dim rsProp As DmtOleDbLib.adoResultset

BLoading = False

sSQL = "SELECT * FROM RV_POContactDettaglio "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaCliente
If IDSitoPerAnagrafica > 0 Then
    sSQL = sSQL & " AND IDSitoPerAnagrafica=" & IDSitoPerAnagrafica
End If
sSQL = sSQL & " ORDER BY IDSitoPerAnagrafica, Nominativo"

Set rs = Cn.OpenResultset(sSQL)

ICont = 1

If rs.EOF Then
    rs.CloseResultset
    Set rs = Nothing
    Me.cmdTelContatti.Visible = False
    Me.cmdTrovaContatti.Visible = False
    
    Exit Sub
End If

Me.cmdTelContatti.Visible = True
Me.cmdTrovaContatti.Visible = True

While Not rs.EOF
    rsGrigliaTel.AddNew
        rsGrigliaTel!IDRV_POContactDettaglio = fnNotNullN(rs!IDRV_POContactDettaglio)
        rsGrigliaTel!IDAnagrafica = fnNotNullN(rs!IDAnagrafica)
        rsGrigliaTel!IDSitoPerAnagrafica = fnNotNullN(rs!IDSitoPerAnagrafica)
        rsGrigliaTel!Nominativo = fnNotNull(rs!Nominativo)
        rsGrigliaTel!Posizione = fnNotNull(rs!Posizione)
        If fnNotNullN(rs!IDSitoPerAnagrafica) > 0 Then
            rsGrigliaTel!SitoPerAnagrafica = GET_SITO_PER_ANAGRAFICA(fnNotNullN(rs!IDSitoPerAnagrafica))
        Else
            rsGrigliaTel!SitoPerAnagrafica = ""
        End If
        
        sSQL = "SELECT RV_POContactValoriProp.IDRV_POContactValoriProp, RV_POContactValoriProp.IDRV_POContactDettaglio, "
        sSQL = sSQL & "RV_POContactValoriProp.IDRV_POContactProprieta, RV_POContactValoriProp.IDRV_POContactGruppo, RV_POContactValoriProp.IDRV_POContact, "
        sSQL = sSQL & "RV_POContactValoriProp.IDAnagrafica, RV_POContactValoriProp.IDSitoPerAnagrafica, RV_POContactValoriProp.Valore, "
        sSQL = sSQL & "RV_POContactProprieta.Visualizza , RV_POContactProprieta.Gestisci, RV_POContactProprieta.NomeCampo, RV_POContactProprieta.Proprieta "
        sSQL = sSQL & "FROM RV_POContactValoriProp INNER JOIN "
        sSQL = sSQL & "RV_POContactProprieta ON RV_POContactValoriProp.IDRV_POContactProprieta = RV_POContactProprieta.IDRV_POContactProprieta "
        sSQL = sSQL & "WHERE IDRV_POContactDettaglio=" & fnNotNullN(rs!IDRV_POContactDettaglio)
        sSQL = sSQL & "  AND Gestisci=1"
        
        Set rsProp = Cn.OpenResultset(sSQL)
        
        While Not rsProp.EOF
            rsGrigliaTel.Fields(fnNotNull(rsProp!NomeCampo)).Value = fnNotNull(rsProp!Valore)
        rsProp.MoveNext
        Wend
    rsGrigliaTel.Update
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim OLD_Cursor As Long
Dim sSQL As String
Dim cl As DmtGridCtl.dgColumnHeader
Dim rs As DmtOleDbLib.adoResultset
OLD_Cursor = Cn.CursorLocation
Cn.CursorLocation = 3

sSQL = ""

If Len(Trim(Me.txtNominativoRic.Text)) > 0 Then
    If Me.txtNominativoRic.Text <> "%" Then
        sSQL = "Nominativo LIKE " & fnNormString(Me.txtNominativoRic.Text & "%")
    End If
End If
If Me.cboFilialeRic.CurrentID > 0 Then
    If sSQL = "" Then
        sSQL = "IDSitoPerAnagrafica=" & Me.cboFilialeRic.CurrentID
    Else
        sSQL = " AND IDSitoPerAnagrafica=" & Me.cboFilialeRic.CurrentID
    End If
End If
If Len(Trim(Me.txtPosizioneRic.Text)) > 0 Then
    If Me.txtPosizioneRic.Text <> "%" Then
        If sSQL = "" Then
            sSQL = "Posizione LIKE " & fnNormString(Me.txtPosizioneRic.Text & "%")
        Else
            sSQL = " AND Posizione=" & fnNormString(Me.txtPosizioneRic.Text & "%")
        End If
    End If
End If

rsGrigliaTel.Filter = sSQL

With Me.GrigliaTel
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectRow
    .ColumnsHeader.Clear
        .ColumnsHeader.Add "IDNumero", "ID", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "SitoPerAnagrafica", "Altra sede", dgchar, True, 1800, dgAlignleft
        .ColumnsHeader.Add "Nominativo", "Nominativo", dgchar, True, 1800, dgAlignleft
        sSQL = "SELECT * FROM RV_POContactProprieta "
        sSQL = sSQL & "WHERE Gestisci=1"
        Set rs = Cn.OpenResultset(sSQL)
        While Not rs.EOF
            .ColumnsHeader.Add fnNotNull(rs!NomeCampo), fnNotNull(rs!Proprieta), dgchar, True, 1800, dgAlignleft
        rs.MoveNext
        Wend
        rs.CloseResultset
        Set rs = Nothing
    Set .Recordset = rsGrigliaTel
    .LoadUserSettings
    .Refresh
End With

Cn.CursorLocation = OLD_Cursor

Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "Griglia recapiti"
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.Width > CONST_WIDTH_FORM Then
        If Me.FraTrovaContatto.Visible = False Then
            Me.GrigliaTel.Left = Me.FraTelAna.Width + Me.FraTelAna.Left + 120
            Me.GrigliaTel.Width = Me.Width - Me.FraTelAna.Width + Me.FraTelAna.Left - 500
        Else
            Me.GrigliaTel.Left = Me.FraTelAna.Width + Me.FraTelAna.Left + Me.FraTrovaContatto.Width + 120
            Me.GrigliaTel.Width = Me.Width - Me.FraTelAna.Width + Me.FraTelAna.Left - 500 - Me.FraTrovaContatto.Width
        End If
        
    Else
        Me.GrigliaTel.Width = 300
    End If
        
    Me.FraNuovoContatto.Left = Me.FraTelAna.Width + Me.FraTelAna.Left + 120
    Me.FraNuovoContatto.Width = Me.Width - Me.FraTelAna.Width + Me.FraTelAna.Left - 500
    Me.MSFlexGrid1.Width = Me.FraNuovoContatto.Width - Me.Line1.X1
    Me.cmdConfermaContatto.Left = Me.FraNuovoContatto.Width - 120 - Me.cmdConfermaContatto.Width
    RESIZE_GRIGLIA_RECAPITI Me.MSFlexGrid1.Width
    
End Sub
Private Function GET_SITO_PER_ANAGRAFICA(IDSitoPerAnagrafica As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SitoPerAnagrafica FROM SitoPerAnagrafica "
sSQL = sSQL & "WHERE IDSitoPerAnagrafica=" & IDSitoPerAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_SITO_PER_ANAGRAFICA = ""
Else
    GET_SITO_PER_ANAGRAFICA = fnNotNull(rs!SitoPerAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    
    If Not (rsGrigliaTel Is Nothing) Then
        If rsGrigliaTel.State > 0 Then
            rsGrigliaTel.Close
        End If
        
        Set rsGrigliaTel = Nothing
    End If

End Sub

Private Sub GET_CARATTERISTICHE_RISORSA(IDAnagrafica As Long, IDSitoPerAnagrafica As Long, IDDettaglio As Long)
On Error Resume Next
Dim sSQL As String
Dim rsDett As DmtOleDbLib.adoResultset
Dim NumeroColonne As Long
Dim NumeroRighe As Long
Dim N_Col As Long
Dim N_Row As Long
Dim IDTipoRisorsa As Long
LoadingGriglia = True
Me.MSFlexGrid1.Visible = False
N_Col = 0
N_Row = 0

BLoading = True
Me.Text1.Visible = False
NumeroColonne = 3
NumeroRighe = GET_NUMERO_RIGHE_CAMPI_RISORSA


Me.MSFlexGrid1.Clear
Me.MSFlexGrid1.Cols = 1
Me.MSFlexGrid1.Rows = 1
DoEvents

With Me.MSFlexGrid1
    .Cols = NumeroColonne
    .Rows = NumeroRighe + 1
End With

''''INTESTAZIONI'''''''''''''''''''''''''''''''''''''''''''''''
Me.MSFlexGrid1.TextMatrix(N_Row, 0) = "Proprietà"
Me.MSFlexGrid1.TextMatrix(N_Row, 1) = "Valore"
Me.MSFlexGrid1.TextMatrix(N_Row, 2) = "Gruppo"

Me.MSFlexGrid1.Col = 0
Me.MSFlexGrid1.Row = N_Row
Me.MSFlexGrid1.CellFontBold = True
Me.MSFlexGrid1.Col = 1
Me.MSFlexGrid1.Row = N_Row
Me.MSFlexGrid1.CellFontBold = True
Me.MSFlexGrid1.Col = 2
Me.MSFlexGrid1.Row = N_Row
Me.MSFlexGrid1.CellFontBold = True

'Me.MSFlexGrid1.ColWidth(0) = 1800
'Me.MSFlexGrid1.ColWidth(1) = 4000
'Me.MSFlexGrid1.ColWidth(2) = 1720
RESIZE_GRIGLIA_RECAPITI Me.MSFlexGrid1.Width

N_Row = 1

sSQL = "SELECT RV_POContactProprieta.IDRV_POContactProprieta, RV_POContactProprieta.IDRV_POContactGruppo, RV_POContactProprieta.Proprieta, "
sSQL = sSQL & "RV_POContactGruppo.ContactGruppo "
sSQL = sSQL & "FROM RV_POContactProprieta LEFT OUTER JOIN "
sSQL = sSQL & "RV_POContactGruppo ON RV_POContactProprieta.IDRV_POContactGruppo = RV_POContactGruppo.IDRV_POContactGruppo "
sSQL = sSQL & "WHERE Gestisci=1 "
sSQL = sSQL & "ORDER BY ContactGruppo "

Set rsDett = Cn.OpenResultset(sSQL)

While Not rsDett.EOF
            
    Me.MSFlexGrid1.TextMatrix(N_Row, 0) = fnNotNull(rsDett!Proprieta)
    Me.MSFlexGrid1.TextMatrix(N_Row, 1) = GET_VALORE_PROPRIETA(IDDettaglio, fnNotNullN(rsDett!IDRV_POContactProprieta))
    Me.MSFlexGrid1.TextMatrix(N_Row, 2) = GET_DESCRIZIONE_GRUPPO_PROPRIETA(fnNotNullN(rsDett!IDRV_POContactGruppo))
    Me.MSFlexGrid1.RowData(N_Row) = fnNotNullN(rsDett!IDRV_POContactProprieta)
        
    Me.MSFlexGrid1.Col = 1
    Me.MSFlexGrid1.Row = N_Row
    Me.MSFlexGrid1.CellFontBold = True
    Me.MSFlexGrid1.CellAlignment = 7
        
    N_Row = N_Row + 1
    DoEvents
rsDett.MoveNext
Wend

rsDett.CloseResultset
Set rsDett = Nothing
Me.MSFlexGrid1.Visible = True
BLoading = False

End Sub
Private Function GET_NUMERO_RIGHE_CAMPI_RISORSA()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Count(IDRV_POContactProprieta) as NumeroRecord "
sSQL = sSQL & "FROM RV_POContactProprieta "
sSQL = sSQL & "WHERE Gestisci = 1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_RIGHE_CAMPI_RISORSA = 1
Else
    GET_NUMERO_RIGHE_CAMPI_RISORSA = fnNotNullN(rs!NumeroRecord)
End If
rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_VALORE_PROPRIETA(IDDettaglio As Long, IDProprieta As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Valore FROM RV_POContactValoriProp "
sSQL = sSQL & "WHERE IDRV_POContactDettaglio=" & IDDettaglio
sSQL = sSQL & " AND IDRV_POContactProprieta=" & IDProprieta

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_VALORE_PROPRIETA = ""
Else
    GET_VALORE_PROPRIETA = fnNotNull(rs!Valore)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_GRUPPO_PROPRIETA(IDGruppo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POContactGruppo "
sSQL = sSQL & "WHERE IDRV_POContactGruppo=" & IDGruppo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_GRUPPO_PROPRIETA = ""
Else
    GET_DESCRIZIONE_GRUPPO_PROPRIETA = fnNotNull(rs!ContactGruppo)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_GRUPPO_PROPRIETA(IDProprieta As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POContactProprieta "
sSQL = sSQL & "WHERE IDRV_POContactProprieta=" & IDProprieta

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_GRUPPO_PROPRIETA = 0
Else
    GET_LINK_GRUPPO_PROPRIETA = fnNotNullN(rs!IDRV_POContactGruppo)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GrigliaTel_DblClick()
    If ((Me.GrigliaTel.Recordset.EOF) And (Me.GrigliaTel.Recordset.BOF)) Then Exit Sub
    
    Me.GrigliaTel.Visible = False
    Me.FraNuovoContatto.Visible = True
    LINK_DETTAGLIO_CONTATTO = fnNotNullN(Me.GrigliaTel.AllColumns("IDRV_POContactDettaglio").Value)
    GET_PROPRIETA_CONTATTO LINK_DETTAGLIO_CONTATTO
    GET_CARATTERISTICHE_RISORSA LINK_ANAGRAFICA_TEL, LINK_SITO_PER_ANAGRAFICA_TEL, LINK_DETTAGLIO_CONTATTO
    
End Sub

Private Sub MSFlexGrid1_EnterCell()
    If BLoading = True Then Exit Sub
    Me.MSFlexGrid1.ColSel = Me.MSFlexGrid1.Col
    Me.MSFlexGrid1.RowSel = Me.MSFlexGrid1.Row
    
    If Me.MSFlexGrid1.Col <> 1 Then
        Me.Text1.Visible = False
        Exit Sub
    Else
        Me.Text1.Visible = True
    End If
    
    Me.Text1.Visible = True
    Me.Text1.Move Me.MSFlexGrid1.CellLeft + Me.MSFlexGrid1.Left, Me.MSFlexGrid1.CellTop + Me.MSFlexGrid1.Top, Me.MSFlexGrid1.CellWidth, Me.MSFlexGrid1.CellHeight
    Me.Text1.Text = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.RowSel, Me.MSFlexGrid1.ColSel)
    Me.Text1.ZOrder 0
    Me.Text1.SetFocus
    
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim NumeroProp As Long

    If KeyCode = vbKeyReturn Then
        Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.RowSel, Me.MSFlexGrid1.ColSel) = Me.Text1.Text
      
        NumeroProp = GET_NUMERO_RIGHE_CAMPI_RISORSA
        
        Me.Text1.Visible = False
        
        If Me.MSFlexGrid1.Row >= NumeroProp Then
            Me.MSFlexGrid1.Col = Me.MSFlexGrid1.Col
            Me.MSFlexGrid1.Row = 1
        Else
            Me.MSFlexGrid1.Col = Me.MSFlexGrid1.Col
            Me.MSFlexGrid1.Row = Me.MSFlexGrid1.Row + 1

        End If
        
        'Me.Text1.Visible = True
        'Me.Text1.Move Me.MSFlexGrid1.CellLeft + Me.MSFlexGrid1.Left, Me.MSFlexGrid1.CellTop + Me.MSFlexGrid1.Top, Me.MSFlexGrid1.CellWidth, Me.MSFlexGrid1.CellHeight
        'Me.Text1.Text = Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.RowSel, Me.MSFlexGrid1.ColSel)
        'Me.Text1.ZOrder 0
        'Me.Text1.SetFocus
    End If
End Sub
Private Sub AGGIORNA_PROPRIETA_CONTATTO(IDDettaglio As Long, IDTesta As Long, IDAnagrafica As Long, IDSitoPerAnagrafica As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I_Row As Long
Dim Link_Prop As Long

'''ELIMINAZIONE DELLE PROPRIETA DEL CONTATTO''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POContactValoriProp "
sSQL = sSQL & " WHERE IDRV_POContactDettaglio=" & IDDettaglio
Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
I_Row = 1

sSQL = "SELECT * FROM RV_POContactValoriProp "
sSQL = sSQL & "WHERE IDRV_POContactDettaglio=" & IDDettaglio
Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

For I_Row = 1 To Me.MSFlexGrid1.Rows - 1
    Link_Prop = Me.MSFlexGrid1.RowData(I_Row)
    rs.AddNew
        rs!IDRV_POContactValoriProp = fnGetNewKey("RV_POContactValoriProp", "IDRV_POContactValoriProp")
        rs!IDRV_POContactDettaglio = IDDettaglio
        rs!IDRV_POContactProprieta = Link_Prop
        rs!IDRV_POContactGruppo = GET_LINK_GRUPPO_PROPRIETA(Link_Prop)
        rs!IDRV_POContact = IDTesta
        rs!IDAnagrafica = IDAnagrafica
        rs!IDSitoPerAnagrafica = IDSitoPerAnagrafica
        rs!Valore = fnNotNull(Me.MSFlexGrid1.TextMatrix(I_Row, 1))
    rs.Update
Next

End Sub
Private Sub GET_TIPO_ANAGRAFICA(IDAnagrafica As Long)
'Dim sSQL As String
'Dim rs As DmtOleDbLib.adoResultset
'Dim VarTop As Long
'Const VarTopADD As Long = 360
'Dim IControl As Long

'sSQL = "SELECT AnagraficaPerTipo.IDAnagrafica, AnagraficaPerTipo.IDTipoAnagrafica, TipoAnagrafica.TipoAnagrafica "
'sSQL = sSQL & "FROM AnagraficaPerTipo INNER JOIN "
'sSQL = sSQL & "TipoAnagrafica ON AnagraficaPerTipo.IDTipoAnagrafica = TipoAnagrafica.IDTipoAnagrafica "
'sSQL = sSQL & "GROUP BY AnagraficaPerTipo.IDAnagrafica, AnagraficaPerTipo.IDTipoAnagrafica, TipoAnagrafica.TipoAnagrafica "
'sSQL = sSQL & "HAVING AnagraficaPerTipo.IDAnagrafica = " & IDAnagrafica

'Set rs = Cn.OpenResultset(sSQL)

'For IControl = 1 To Me.lblTipoAnagrafica.Count - 1
'    Unload Me.lblTipoAnagrafica(IControl)
'Next

'VarTop = 360
'IControl = 1

'While Not rs.EOF
'    Load Me.lblTipoAnagrafica(IControl)
'    With Me.lblTipoAnagrafica(IControl)
'        .Top = VarTop
'        .Visible = True
'        .ZOrder 0
'        .Caption = fnNotNull(rs!TipoAnagrafica)
'        .ToolTipText = GET_AZIENDE_TIPO_ANAGRAFICA(IDAnagrafica, fnNotNullN(rs!IDTipoAnagrafica))
'    End With
'    IControl = IControl + 1
'    VarTop = VarTop + VarTopADD
'rs.MoveNext
'Wend

'rs.CloseResultset
'Set rs = Nothing
End Sub
Private Function GET_AZIENDE_TIPO_ANAGRAFICA(IDAnagrafica As Long, IDTipoAnagrafica As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroRecord As Long

GET_AZIENDE_TIPO_ANAGRAFICA = ""

sSQL = "SELECT AnagraficaPerTipo.IDAnagrafica, AnagraficaPerTipo.IDTipoAnagrafica, AnagraficaPerTipo.IDAzienda, Anagrafica.Anagrafica "
sSQL = sSQL & "FROM Anagrafica INNER JOIN "
sSQL = sSQL & "Azienda ON Anagrafica.IDAnagrafica = Azienda.IDAnagrafica INNER JOIN "
sSQL = sSQL & "AnagraficaPerTipo ON Azienda.IDAzienda = AnagraficaPerTipo.IDAzienda "
sSQL = sSQL & "WHERE AnagraficaPerTipo.IDAnagrafica = " & IDAnagrafica
sSQL = sSQL & "AND AnagraficaPerTipo.IDTipoAnagrafica = " & IDTipoAnagrafica

Set rs = Cn.OpenResultset(sSQL)
NumeroRecord = 1
While Not rs.EOF
    If NumeroRecord = 1 Then
        GET_AZIENDE_TIPO_ANAGRAFICA = GET_AZIENDE_TIPO_ANAGRAFICA & fnNotNull(rs!Anagrafica)
    Else
        GET_AZIENDE_TIPO_ANAGRAFICA = vbCrLf & GET_AZIENDE_TIPO_ANAGRAFICA & fnNotNull(rs!Anagrafica)
    End If
NumeroRecord = NumeroRecord + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_ANAGRAFICA_AZIENDA(IDAzienda As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Anagrafica.Anagrafica, Azienda.IDAzienda "
sSQL = sSQL & "FROM Anagrafica INNER JOIN "
sSQL = sSQL & "Azienda ON Anagrafica.IDAnagrafica = Azienda.IDAnagrafica "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ANAGRAFICA_AZIENDA = ""
Else
    GET_ANAGRAFICA_AZIENDA = fnNotNull(rs!Anagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub GET_PROPRIETA_CONTATTO(IDDettaglio As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POContactDettaglio "
sSQL = sSQL & "WHERE IDRV_POContactDettaglio=" & IDDettaglio

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.TxtNominativo.Text = ""
    Me.cboFiliale.WriteOn 0
    Me.txtPosizioneAziendale.Text = ""
    Me.txtAnnotazioni.Text = ""
Else
    Me.TxtNominativo.Text = fnNotNull(rs!Nominativo)
    Me.cboFiliale.WriteOn fnNotNullN(rs!IDSitoPerAnagrafica)
    Me.txtPosizioneAziendale.Text = fnNotNull(rs!Posizione)
    Me.txtAnnotazioni.Text = fnNotNull(rs!Annotazioni)
End If
rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub RESIZE_GRIGLIA_RECAPITI(Larghezza As Long)
Dim L_Colonna1 As Long
Dim L_Colonna2 As Long
Dim L_Colonna3 As Long
Const PercColonna_1 As Long = 20
Const PercColonna_2 As Long = 60
Const PercColonna_3 As Long = 15

L_Colonna1 = (Larghezza / 100) * PercColonna_1
L_Colonna2 = (Larghezza / 100) * PercColonna_2
L_Colonna3 = (Larghezza / 100) * PercColonna_3


Me.MSFlexGrid1.ColWidth(0) = L_Colonna1
Me.MSFlexGrid1.ColWidth(1) = L_Colonna2
Me.MSFlexGrid1.ColWidth(2) = L_Colonna3
End Sub
Private Sub INIT_CONTROLLI()
    With Me.cboFiliale
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT IDSitoPerAnagrafica, SitoPerAnagrafica FROM SitoPerAnagrafica "
        .SQL = .SQL & "WHERE IDAnagrafica=" & LINK_ANAGRAFICA_TEL
        .Fill
    End With
    
    With Me.cboFilialeRic
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT IDSitoPerAnagrafica, SitoPerAnagrafica FROM SitoPerAnagrafica "
        .SQL = .SQL & "WHERE IDAnagrafica=" & LINK_ANAGRAFICA_TEL
        .Fill
    End With
    
End Sub

Private Sub txtNominativoRic_Change()
GET_GRIGLIA
End Sub

Private Sub txtPosizioneRic_Change()
GET_GRIGLIA
End Sub
Private Function GET_NOME_CAMPO_PROPRIETA(Nominativo As String) As String
Dim I As Long
Dim ReturnString As String

ReturnString = ""
For I = 1 To Len(Nominativo)
    If Mid(Nominativo, I, 1) <> " " Then
        ReturnString = ReturnString & Mid(Nominativo, I, 1)
    End If
Next

ReturnString = Replace(ReturnString, "-", "")
ReturnString = Replace(ReturnString, "/", "")
ReturnString = Replace(ReturnString, "\", "")
ReturnString = Replace(ReturnString, "(", "")
ReturnString = Replace(ReturnString, ")", "")
ReturnString = Replace(ReturnString, "%", "")
ReturnString = Replace(ReturnString, "$", "")
ReturnString = Replace(ReturnString, "#", "")
ReturnString = Replace(ReturnString, "@", "")
ReturnString = Replace(ReturnString, "é", "e")
ReturnString = Replace(ReturnString, "ò", "o")
ReturnString = Replace(ReturnString, "à", "à")
ReturnString = Replace(ReturnString, "ù", "u")
ReturnString = Replace(ReturnString, "ì", "ì")
ReturnString = Replace(ReturnString, ".", "")
ReturnString = Replace(ReturnString, ":", "")
ReturnString = Replace(ReturnString, "=", "")
ReturnString = Replace(ReturnString, "?", "")
ReturnString = Replace(ReturnString, "^", "")
ReturnString = Replace(ReturnString, "<", "")
ReturnString = Replace(ReturnString, ">", "")
ReturnString = Replace(ReturnString, "|", "")
ReturnString = Replace(ReturnString, "!", "")
ReturnString = Replace(ReturnString, Chr(34), "")
ReturnString = Replace(ReturnString, "£", "")
ReturnString = Replace(ReturnString, "€", "")
ReturnString = Replace(ReturnString, "ç", "")
ReturnString = Replace(ReturnString, "[", "")
ReturnString = Replace(ReturnString, "]", "")
ReturnString = Replace(ReturnString, "+", "")
ReturnString = Replace(ReturnString, "*", "")
ReturnString = Replace(ReturnString, "'", "")
ReturnString = Replace(ReturnString, ";", "")

GET_NOME_CAMPO_PROPRIETA = ReturnString

End Function

