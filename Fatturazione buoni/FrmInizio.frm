VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.9#0"; "DmtCodDesc.ocx"
Begin VB.Form FrmInizio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fatturazione buoni (Passo 1 di 3)"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInizio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   17940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDeSelTutto 
      Caption         =   "Deseleziona tutto"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton cmdSelTutto 
      Caption         =   "Seleziona tutto"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   9000
      Width           =   1815
   End
   Begin DmtGridCtl.DmtGrid GrigliaBuoniFiltro 
      Height          =   6015
      Left            =   2520
      TabIndex        =   1
      Top             =   2880
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   10610
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
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   2520
      TabIndex        =   12
      Top             =   0
      Width           =   15375
      Begin VB.Frame frmImpostazioni 
         Caption         =   "Impostazioni"
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
         Height          =   2535
         Left            =   10680
         TabIndex        =   25
         Top             =   120
         Width           =   4575
         Begin VB.CheckBox Check1 
            Caption         =   "Sovrascrivi descrizione articolo con descrizione addebito per la fatturazione"
            Height          =   435
            Left            =   120
            TabIndex        =   29
            Top             =   1360
            Width           =   4215
         End
         Begin VB.CheckBox chkRaggrCorpoAltraDest 
            Caption         =   "Raggruppa corpo per altra destinazione"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   4215
         End
         Begin VB.CheckBox chkRaggrCorpoInt 
            Caption         =   "Raggruppa corpo per intervento"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   4215
         End
         Begin VB.CheckBox chkRaggrSitoAna 
            Caption         =   "Raggruppa per altra destinazione"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   4095
         End
      End
      Begin DMTDataCmb.DMTCombo cboInterventoChiuso 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
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
      Begin VB.CommandButton cmdRicerca 
         Caption         =   "Ricerca"
         Height          =   375
         Left            =   6120
         TabIndex        =   0
         Top             =   2280
         Width           =   1335
      End
      Begin DmtCodDescCtl.DmtCodDesc CDCliente 
         Height          =   615
         Left            =   3840
         TabIndex        =   8
         Top             =   225
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1085
         PropCodice      =   $"FrmInizio.frx":4781A
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"FrmInizio.frx":47873
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"FrmInizio.frx":478C3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtDaNumeroBuono 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtDate txtAData 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtDate txtDaData 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtNumber txtANumeroBuono 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtNumber txtDaNumeroIntervento 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtNumber txtANumeroIntervento 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DmtCodDescCtl.DmtCodDesc CDTecnicoOpeFase 
         Height          =   615
         Left            =   3840
         TabIndex        =   23
         Top             =   830
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1085
         PropCodice      =   $"FrmInizio.frx":4791D
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"FrmInizio.frx":47978
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"FrmInizio.frx":479C8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin DmtCodDescCtl.DmtCodDesc CDAmministratore 
         Height          =   615
         Left            =   3840
         TabIndex        =   24
         Top             =   1430
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1085
         PropCodice      =   $"FrmInizio.frx":47A22
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"FrmInizio.frx":47A7A
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"FrmInizio.frx":47ACA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   3600
         X2              =   3600
         Y1              =   240
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   10560
         X2              =   10560
         Y1              =   240
         Y2              =   2520
      End
      Begin VB.Label Label4 
         Caption         =   "Intervento chiuso"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "A n° intervento"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Da n° intervento"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "A n° addebito"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Da n° addebito"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Da data addebito"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "A data addebito"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   1815
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
      Height          =   9375
      Left            =   0
      Picture         =   "FrmInizio.frx":47B24
      ScaleHeight     =   9345
      ScaleWidth      =   2385
      TabIndex        =   11
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton CmdFine 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   15000
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton CmdAvanti 
      Caption         =   "Avanti"
      Height          =   375
      Left            =   16560
      TabIndex        =   9
      Top             =   9000
      Width           =   1335
   End
End
Attribute VB_Name = "FrmInizio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents m_App As DMTRunAppLib.Application
Attribute m_App.VB_VarHelpID = -1
Private rsGriglia As DmtOleDbLib.adoResultset


Public Sub ConnessioneADO()
    If Not (CnDMT Is Nothing) Then
        'CnDMT.CloseConnection
        Set CnDMT = Nothing
    End If
  
  Set CnDMT = m_App.Database.Connection
  
  Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"

End Sub

Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set m_App = NewValue
End Property

Public Property Get Application() As DMTRunAppLib.Application
    Set Application = m_App
End Property
Public Sub InitControlli()
    
    CREA_TABELLE
    
    GET_PARAMETRI_AZIENDA m_App.IDFirm

    'Anagrafica cliente
    With Me.CDCliente
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hWnd
        .CodeField = "Anagrafica"
        .DescriptionField = "Nome"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
        .Filter = "IDAzienda = " & m_App.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Anagrafica"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Nome"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Anagrafica"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Nome"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 29 'Anagrafica
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
      
      
    'Intervento chiuso
    With Me.cboInterventoChiuso
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POSINO"
        .DisplayField = "SINO"
        .Sql = "SELECT * FROM RV_POSiNo"
        .Fill
        .WriteOn 1
    End With
    
      

    'Anagrafica del tecnico operativo della fase
    With Me.CDTecnicoOpeFase
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hWnd
        .CodeField = "Anagrafica"
        .DescriptionField = "Nome"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIEAnagraficaPerTipo"
        .Filter = "IDAzienda = " & m_App.IDFirm & " AND IDTipoAnagrafica=" & LINK_TIPO_ANA_TEC_FASE
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Cognome"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Nome"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Cognome"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Nome"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 29 'Anagrafica
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With


    'Anagrafica del tecnico di riferimento
    With Me.CDAmministratore
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hWnd
        .CodeField = "Anagrafica"
        .DescriptionField = "Nome"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIEAnagraficaPerTipo"
        .Filter = "IDAzienda = " & TheApp.IDFirm & " AND IDTipoAnagrafica=" & GET_PARAMETRO_AZIENDA_LONG(TheApp.Branch, "IDTipoAnagraficaAmministratore")
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Cognome"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Nome"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Cognome"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Nome"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 29 'Anagrafica
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    GET_GRIGLIA
    
End Sub

Private Sub chkRaggrCorpoAltraDest_Click()
    If Me.chkRaggrCorpoAltraDest.Value = vbChecked Then
        Me.chkRaggrSitoAna.Value = vbUnchecked
        Me.chkRaggrSitoAna.Enabled = False
    Else
        Me.chkRaggrSitoAna.Enabled = True
    End If
End Sub

Private Sub chkRaggrSitoAna_Click()
    If Me.chkRaggrSitoAna.Value = vbChecked Then
        Me.chkRaggrCorpoAltraDest.Value = vbUnchecked
        Me.chkRaggrCorpoAltraDest.Enabled = False
    Else
        Me.chkRaggrCorpoAltraDest.Enabled = True
    End If
End Sub

Private Sub cmdAvanti_Click()
'On Error GoTo ERR_cmdAvanti_Click
    Screen.MousePointer = 11
        FLAG_RAGGR_ALTRA_DEST = Me.chkRaggrSitoAna.Value
        FLAG_RAGGR_CORPO_ALTRA_DEST = Me.chkRaggrCorpoAltraDest.Value
        FLAG_RAGGR_CORPO_INT = Me.chkRaggrCorpoInt.Value
        FLAG_SOVRAS_DESCRIZIONE = Me.Check1.Value
    
        SALVA_DATI_DA_REGISTRARE
    Screen.MousePointer = 0
    Unload Me
Exit Sub
ERR_cmdAvanti_Click:
    MsgBox Err.Description, vbCritical, "cmdAvanti_Click"
    
End Sub

Private Sub cmdDeSelTutto_Click()
On Error GoTo ERR_cmdDeSelTutto_Click
If ((rsFiltro.EOF) And (rsFiltro.BOF)) Then Exit Sub
rsFiltro.MoveFirst

While Not rsFiltro.EOF
    rsFiltro!Registra = 0
    rsFiltro.Update
rsFiltro.MoveNext
Wend


Me.GrigliaBuoniFiltro.Refresh
Exit Sub
ERR_cmdDeSelTutto_Click:
    MsgBox Err.Description, vbCritical, "cmdDeSelTutto_Click"
    
End Sub

Private Sub CmdFine_Click()
Dim Risposta As Integer
    Risposta = MsgBox("Vuoi abbandonare il wizard per il passaggio degli interventi in fatturazione?", vbInformation + vbYesNo, "Abbandono")
    If Risposta = vbYes Then
        Unload Me
    End If
End Sub
Private Sub SALVA_DATI_DA_REGISTRARE()

If ((rsFiltro.EOF) And (rsFiltro.BOF)) Then Exit Sub

rsFiltro.Update

NUMERO_ANAGRAFICHE = 0
NUMERO_ADDEBITI_DA_FATTURARE = 0

rsFiltro.Filter = "Registra=1"
    While Not rsFiltro.EOF
        rsnew.AddNew
            For I = 0 To rsFiltro.Fields.Count - 1
                Select Case rsnew.Fields(I).Type
                    
                    Case adInteger
                        rsnew.Fields(rsFiltro.Fields(I).Name).Value = fnNotNullN(rsFiltro.Fields(I).Value)
                    Case Else
                        rsnew.Fields(rsFiltro.Fields(I).Name).Value = rsFiltro.Fields(I).Value
                End Select

                
            Next
            
        rsnew.Update
        
        NUMERO_ADDEBITI_DA_FATTURARE = NUMERO_ADDEBITI_DA_FATTURARE + 1
        
        'CLIENTI'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        rsAna.Filter = "IDAnagraficaCliente=" & fnNotNullN(rsFiltro.Fields("IDAnagraficaFatturazione").Value)
        rsAna.Filter = rsAna.Filter & " AND RitenutaAcconto=" & fnNotNullN(rsFiltro.Fields("RitenutaAcconto").Value)
        If FLAG_RAGGR_ALTRA_DEST = 1 Then
            rsAna.Filter = rsAna.Filter & " AND IDSitoPerAnagrafica=" & fnNotNullN(rsFiltro.Fields("IDSitoPerAnagraficaIntervento").Value)
        End If
        
        
        If (rsAna.EOF) Then
            
            rsAna.AddNew
                rsAna.Fields("IDAnagraficaCliente").Value = rsFiltro.Fields("IDAnagraficaFatturazione").Value
                rsAna.Fields("RitenutaAcconto").Value = rsFiltro.Fields("RitenutaAcconto").Value
                rsAna.Fields("IDSitoPerAnagrafica").Value = fnNotNullN(rsFiltro.Fields("IDSitoPerAnagraficaIntervento").Value)
            rsAna.Update
            
            NUMERO_ANAGRAFICHE = NUMERO_ANAGRAFICHE + 1
        
        End If
        
        rsAna.Filter = vbNullString
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    rsFiltro.MoveNext
Wend

rsFiltro.Close
Set rsFiltro = Nothing

End Sub

Private Sub cmdRicerca_Click()
    GET_GRIGLIA
End Sub



Private Sub cmdSelTutto_Click()
On Error GoTo ERR_cmdSelTutto_Click
If ((rsFiltro.EOF) And (rsFiltro.BOF)) Then Exit Sub
rsFiltro.MoveFirst

While Not rsFiltro.EOF
    rsFiltro!Registra = 1
    rsFiltro.Update
rsFiltro.MoveNext
Wend


Me.GrigliaBuoniFiltro.Refresh
Exit Sub
ERR_cmdSelTutto_Click:
    MsgBox Err.Description, vbCritical, TheApp.FunctionName
End Sub

Private Sub Form_Activate()
ConnessioneADO

Me.Caption = TheApp.FunctionName & " (Passo 1 di 3)"

GET_PARAMETRI_FILIALE_BUONI

DefaultImportazione



If b_Loading = True Then
    InitControlli
End If

End Sub

Private Sub Form_Load()
On Error GoTo ERR_Form_Load

Exit Sub
ERR_Form_Load:
    MsgBox Err.Description, vbCritical, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERR_Form_Unload
Dim Risposta As Integer

    If Me.cmdAvanti.Value = True Then
        FrmParametri.Show

        Set rsFiltro = Nothing
    End If

Exit Sub

ERR_Form_Unload:
    MsgBox Err.Description, vbCritical, "Form_Unload"
End Sub
Private Sub GET_PARAMETRI_AZIENDA(IDAzienda As Long)


Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POParametriAzienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    LINK_TIPO_ANA_TEC_INT = 0
    LINK_STATO_INT_NUOVO = 0
    LINK_STATO_INT_CHIUSO = 0
    
    LINK_TIPO_ANA_TEC_FASE = 0
    LINK_STATO_FASE_NUOVA = 0
    LINK_STATO_FASE_CHIUSA = 0
    LINK_TIPO_FASE_ELA = 0
    LINK_TIPO_FASE_MANUALE = 0
    LINK_CARICO_MAGAZZINO = 0
Else
    LINK_TIPO_ANA_TEC_INT = fnNotNullN(rs!IDTipoAnagraficaTecnicoIntRif)
    LINK_STATO_INT_NUOVO = fnNotNullN(rs!IDRV_POStatoInterventoInserimento)
    LINK_STATO_INT_CHIUSO = fnNotNullN(rs!IDRV_POStatoInterventoChiuso)
    
    LINK_TIPO_ANA_TEC_FASE = fnNotNullN(rs!IDTipoAnagraficaTecnicoFaseRif)
    LINK_STATO_FASE_NUOVA = fnNotNullN(rs!IDRV_POStatoFaseInserimento)
    LINK_STATO_FASE_CHIUSA = fnNotNullN(rs!IDRV_POStatoFaseChiusa)
    LINK_TIPO_FASE_ELA = fnNotNullN(rs!IDRV_POTipoFaseInterventoEla)
    LINK_TIPO_FASE_MANUALE = fnNotNullN(rs!IDRV_POTipoFaseInterventoMan)
    LINK_CARICO_MAGAZZINO = fnNotNullN(rs!IDFunzioneCaricoAddebito)
    
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_GRIGLIA()
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader

    CREA_TABELLA_GRIGLIA


    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
     
    With Me.GrigliaBuoniFiltro
        .EnableMove = True
        .UpdatePosition = False
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
        
        Set cl = .ColumnsHeader.Add("Registra", "Registra", dgBoolean, True, 1000, dgAligncenter)
            cl.Editable = True
        Set cl = .ColumnsHeader.Add("RitenutaAcconto", "Ritenuta acconto", dgBoolean, True, 1000, dgAligncenter)
            cl.Editable = True
        .ColumnsHeader.Add "IDRV_POInterventoRigheDett", "IDRV_POInterventoRigheDett", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDRV_POInterventoRighe", "IDRV_POInterventoRighe", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "IDRV_POIntervento", "IDRV_POIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "NumeroDocumento", "Numero", dgInteger, True, 1000, dgAlignRight
        .ColumnsHeader.Add "DataDocumento", "Data", dgDate, True, 1500, dgAlignleft
        
        .ColumnsHeader.Add "NumeroIntervento", "N° Intervento", dgInteger, True, 1000, dgAlignRight
        .ColumnsHeader.Add "AnnoIntervento", "Anno intervento", dgInteger, True, 1000, dgAlignRight
        .ColumnsHeader.Add "NumeroFase", "N° Fase", dgInteger, True, 1000, dgAlignRight
        
        .ColumnsHeader.Add "IDAnagraficaTecnicoOperativo", "IDAnagraficaTecnicoOperativo", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaTecnicoOperativo", "Tecnico operativo", dgVarChar, True, 1700, dgAlignleft
        .ColumnsHeader.Add "NomeTecnicoOperativo", "Nome tecnico operativo", dgVarChar, True, 1700, dgAlignleft
        
        .ColumnsHeader.Add "IDAnagraficaCliente", "IDAnagraficaCliente", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaCliente", "Anagrafica cliente", dgVarChar, True, 1700, dgAlignleft
        .ColumnsHeader.Add "NomeCliente", "Nome cliente", dgVarChar, True, 1700, dgAlignleft
        
        .ColumnsHeader.Add "IDAnagraficaFatturazione", "IDAnagraficaFatturazione", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "AnagraficaFatturazione", "Anagrafica fatturazione", dgVarChar, True, 1700, dgAlignleft
        .ColumnsHeader.Add "NomeAnagraficaFatturazione", "Nome anagrafica fatturazione", dgVarChar, True, 1700, dgAlignleft
        
        .ColumnsHeader.Add "IDSitoPerAnagraficaIntervento", "IDSitoPerAnagraficaIntervento", dgInteger, False, 500, dgAlignleft
        .ColumnsHeader.Add "SitoPerAnagraficaIntervento", "Altra sede", dgVarChar, False, 1700, dgAlignleft
        
        If LINK_TIPO_GESTIONE_BUONI <= 1 Then
            Set cl = .ColumnsHeader.Add("ImportoDiFatturazione", "Importo Fatt.", dgDouble, True, 2000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
        Else
            Set cl = .ColumnsHeader.Add("Quantita", "Q.tà", dgDouble, True, 2000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("ImportoUnitario", "Importo unitario", dgDouble, True, 2000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("IDIva", "IDIva", dgDouble, False, 500, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("AliquotaIva", "Aliquota", dgDouble, True, 1000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Sconto1", "Sc. 1", dgDouble, True, 1000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Sconto2", "Sc. 2", dgDouble, True, 1000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("Sconto3", "Sc. 3", dgDouble, True, 1000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("TotaleRigaNettoIva", "Imponibile", dgDouble, True, 1000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("TotaleRigaIva", "I.V.A.", dgDouble, True, 1000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .ColumnsHeader.Add("TotaleRigaLordoIva", "Totale riga", dgDouble, True, 1000, dgAlignRight)
            cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
        End If

        
                
        Set .Recordset = rsFiltro
        .LoadUserSettings
        .Refresh
    End With
    
    CnDMT.CursorLocation = OLDCursor

End Sub
Private Sub CREA_TABELLA_TEMPORANEA()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Set rsnew = New ADODB.Recordset

rsnew.CursorLocation = adUseClient

rsnew.Fields.Append "IDRV_POInterventoRigheDett", adInteger, , adFldIsNullable
rsnew.Fields.Append "IDRV_POIntervento", adInteger, , adFldIsNullable
rsnew.Fields.Append "NumeroBuono", adInteger, , adFldIsNullable
rsnew.Fields.Append "DataBuono", adVarChar, 250, adFldIsNullable

rsnew.Fields.Append "NumeroIntervento", adInteger, , adFldIsNullable
rsnew.Fields.Append "AnnoIntervento", adInteger, , adFldIsNullable

rsnew.Fields.Append "IDAnagraficaTecnico", adInteger, , adFldIsNullable
rsnew.Fields.Append "AnagraficaTecnico", adVarChar, 250, adFldIsNullable
rsnew.Fields.Append "NomeTecnico", adVarChar, 250, adFldIsNullable

rsnew.Fields.Append "IDAnagraficaCliente", adInteger, , adFldIsNullable
rsnew.Fields.Append "AnagraficaCliente", adVarChar, 250, adFldIsNullable
rsnew.Fields.Append "NomeCliente", adVarChar, 250, adFldIsNullable

rsnew.Fields.Append "IDAnagraficaIntervento", adInteger, , adFldIsNullable
rsnew.Fields.Append "AnagraficaIntervento", adVarChar, 250, adFldIsNullable
rsnew.Fields.Append "NomeAnagraficaIntervento", adVarChar, 250, adFldIsNullable

rsnew.Fields.Append "ImportoFatturazione", adDouble, , adFldIsNullable
rsnew.Fields.Append "Quantita", adDouble, , adFldIsNullable
rsnew.Fields.Append "ImportoUnitario", adDouble, , adFldIsNullable
rsnew.Fields.Append "IDIva", adInteger, , adFldIsNullable
rsnew.Fields.Append "AliquotaIva", adDouble, , adFldIsNullable
rsnew.Fields.Append "Sconto1", adDouble, , adFldIsNullable
rsnew.Fields.Append "Sconto2", adDouble, , adFldIsNullable
rsnew.Fields.Append "Sconto3", adDouble, , adFldIsNullable
rsnew.Fields.Append "TotaleRigaNettoIva", adDouble, , adFldIsNullable
rsnew.Fields.Append "TotaleRigaIva", adDouble, , adFldIsNullable
rsnew.Fields.Append "TotaleRigaLordoIva", adDouble, , adFldIsNullable
rsnew.Fields.Append "RitenutaAcconto", adSmallInt, , adFldIsNullable
rsnew.Fields.Append "Registra", adSmallInt, , adFldIsNullable
rsnew.Fields.Append "IDSitoPerAnagrafica", adInteger, , adFldIsNullable
rsnew.Fields.Append "SitoPerAnagrafica", adInteger, , adFldIsNullable
rsnew.Fields.Append "EntePubblico", adSmallInt, , adFldIsNullable
rsnew.Fields.Append "IDAccordoCommerciale", adInteger, , adFldIsNullable
rsnew.Fields.Append "IDRaggruppamentoFatturato", adInteger, , adFldIsNullable
rsnew.Fields.Append "IDMagazzino", adInteger, adFldIsNullable

rsnew.Open , , adOpenKeyset, adLockBatchOptimistic

End Sub
Private Sub CREA_TABELLA_GRIGLIA()
On Error GoTo ERR_CREA_TABELLA_GRIGLIA
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim I As Long


If Not (rsFiltro Is Nothing) Then
    If rsFiltro.State > 0 Then
        rsFiltro.Close
    End If
    Set rsFiltro = Nothing
End If

Set rsFiltro = New ADODB.Recordset

rsFiltro.CursorLocation = adUseClient

sSQL = "SELECT * FROM RV_POIEBuoniIntervervento "
sSQL = sSQL & " WHERE IDRV_POInterventoRigheDett=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

With rs
    For I = 0 To rs.Fields.Count - 1
        Select Case rs.Fields(I).Type
            Case adChar, adVarChar, adVarWChar, adWChar, 201
                rsFiltro.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
            Case adNumeric, adBigInt, adCurrency, adDecimal, adDouble, adInteger, adLongVarBinary, adSingle
                rsFiltro.Fields.Append .Fields(I).Name, adDouble, , adFldIsNullable
            Case adDate, adDBTimeStamp, adDBDate
                rsFiltro.Fields.Append .Fields(I).Name, adDBDate, , adFldIsNullable
            Case adSmallInt, adBoolean
                rsFiltro.Fields.Append .Fields(I).Name, adSmallInt, , adFldIsNullable
            Case Else
                rsFiltro.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
        End Select
    Next
End With

rsFiltro.Fields.Append "Registra", adSmallInt, , adFldIsNullable
rsFiltro.Open , , adOpenKeyset, adLockBatchOptimistic

rs.Close
Set rs = Nothing


sSQL = "SELECT * FROM RV_POIEBuoniIntervervento"
sSQL = sSQL & " WHERE IDAzienda=" & m_App.IDFirm
sSQL = sSQL & " AND DaFatturare=1"
sSQL = sSQL & " AND Fatturata=0"
sSQL = sSQL & " AND Preventivo=" & fnNormBoolean(0)

If Me.cboInterventoChiuso.CurrentID > 0 Then
    If Me.cboInterventoChiuso.CurrentID = 1 Then
        sSQL = sSQL & " AND InterventoChiuso=1"
    End If
    If Me.cboInterventoChiuso.CurrentID = 2 Then
        sSQL = sSQL & " AND InterventoChiuso=0"
    End If
End If

If Me.txtDaNumeroBuono.Value > 0 Then
    If Me.txtANumeroBuono.Value = 0 Then
        sSQL = sSQL & " AND NumeroDocumento=" & Me.txtDaNumeroBuono.Value
    Else
        sSQL = sSQL & " AND NumeroDocumento>=" & Me.txtDaNumeroBuono.Value
        sSQL = sSQL & " AND NumeroDocumento<=" & Me.txtANumeroBuono.Value
    End If
End If

If Me.txtDaNumeroIntervento.Value > 0 Then
    If Me.txtANumeroIntervento.Value = 0 Then
        sSQL = sSQL & " AND NumeroIntervento=" & Me.txtDaNumeroIntervento.Value
    Else
        sSQL = sSQL & " AND NumeroIntervento>=" & Me.txtDaNumeroIntervento.Value
        sSQL = sSQL & " AND NumeroIntervento<=" & Me.txtANumeroIntervento.Value
    End If
End If

If (Me.txtDaData.Value > 0) And (Me.txtAData.Value = 0) Then
    sSQL = sSQL & " AND DataDocumento=" & fnNormDate(Me.txtDaData.Text)
End If
If (Me.txtDaData.Value > 0) And (Me.txtAData.Value > 0) Then
    sSQL = sSQL & " AND DataDocumento>=" & fnNormDate(Me.txtDaData.Text)
    sSQL = sSQL & " AND DataDocumento<=" & fnNormDate(Me.txtAData.Text)
End If

If Me.CDCliente.KeyFieldID > 0 Then
    sSQL = sSQL & " AND IDAnagraficaFatturazione=" & Me.CDCliente.KeyFieldID
End If

If Me.CDTecnicoOpeFase.KeyFieldID > 0 Then
    sSQL = sSQL & " AND IDAnagraficaTecnicoOperativo=" & Me.CDTecnicoOpeFase.KeyFieldID
End If
If Me.CDAmministratore.KeyFieldID > 0 Then
    sSQL = sSQL & " AND IDAmmninistratore=" & Me.CDAmministratore.KeyFieldID
End If

sSQL = sSQL & " ORDER BY DataDocumento DESC, NumeroDocumento DESC"


Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

While Not rs.EOF
    If fnNotNullN(GET_BLOCCO_CLIENTE(rs!IDAnagraficaFatturazione)) = 0 Then
        rsFiltro.AddNew
            For I = 0 To rs.Fields.Count - 1
                rsFiltro(rs.Fields(I).Name).Value = rs.Fields(I).Value
            Next
            rsFiltro!Registra = 1
        rsFiltro.Update
    End If
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

rsFiltro.Sort = "AnnoIntervento DESC, NumeroIntervento DESC"


Exit Sub
ERR_CREA_TABELLA_GRIGLIA:
    MsgBox Err.Description, vbCritical, "CREA_TABELLA_GRIGLIA"
End Sub

Private Sub GrigliaBuoniFiltro_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_GrigliaBuoniFiltro_KeyPress
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If Me.GrigliaBuoniFiltro.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsFiltro.Fields("Registra").Value), 2
        End If
    End If

Exit Sub
ERR_GrigliaBuoniFiltro_KeyPress:
    MsgBox Err.Description, vbCritical, "GrigliaBuoniFiltro_KeyPress"
    
End Sub

Private Sub GrigliaBuoniFiltro_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_GrigliaBuoniFiltro_MouseUp
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If GrigliaBuoniFiltro.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < GrigliaBuoniFiltro.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If GrigliaBuoniFiltro.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsFiltro.Fields("Registra").Value), 2
            End If
        End If
        If ((X > 0) And ((X * Screen.TwipsPerPixelX) < (GrigliaBuoniFiltro.ColumnsHeader(2).Width * 2)) And ((X * Screen.TwipsPerPixelX) > GrigliaBuoniFiltro.ColumnsHeader(1).Width)) Then
            'Se non siamo in modalità filtri
            If GrigliaBuoniFiltro.GuiMode = dgNormal Then
            'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRowRitAcc Not CBool(rsFiltro.Fields("RitenutaAcconto").Value)
            End If
        End If
    End If
    

Exit Sub
ERR_GrigliaBuoniFiltro_MouseUp:
    MsgBox Err.Description, vbCritical, "GrigliaBuoniFiltro_MouseUp"

End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
On Error GoTo ERR_sbSelectSelectedRow
    If Not rsFiltro.EOF And Not rsFiltro.BOF Then
        rsFiltro.Fields("Registra").Value = Abs(CLng(Selected))
        Me.GrigliaBuoniFiltro.Refresh
    End If
    DoEvents
Exit Sub
ERR_sbSelectSelectedRow:
    MsgBox Err.Description, vbCritical, "sbSelectSelectedRow"
End Sub
Private Sub sbSelectSelectedRowRitAcc(ByVal Selected As Boolean)
On Error GoTo ERR_sbSelectSelectedRow
    If Not rsFiltro.EOF And Not rsFiltro.BOF Then
        rsFiltro.Fields("RitenutaAcconto").Value = Abs(CLng(Selected))
        Me.GrigliaBuoniFiltro.Refresh
    End If
    DoEvents
Exit Sub
ERR_sbSelectSelectedRow:
    MsgBox Err.Description, vbCritical, "sbSelectSelectedRowRitAcc"
End Sub

Private Sub txtDaNumeroBuono_LostFocus()
    GET_GRIGLIA
End Sub
Private Sub CREA_TABELLE()
Set rsnew = New ADODB.Recordset
rsnew.CursorLocation = adUseClient


sSQL = "SELECT * FROM RV_POIEBuoniIntervervento "
sSQL = sSQL & " WHERE IDRV_POInterventoRigheDett=0"

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

With rs
    For I = 0 To rs.Fields.Count - 1
        Select Case rs.Fields(I).Type
            Case adChar, adVarChar, adVarWChar, adWChar, 201
                rsnew.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
            Case adNumeric, adBigInt, adCurrency, adDecimal, adDouble, adLongVarBinary, adSingle
                rsnew.Fields.Append .Fields(I).Name, adDouble, , adFldIsNullable
            Case adInteger
                rsnew.Fields.Append .Fields(I).Name, adInteger, , adFldIsNullable
            Case adDate, adDBTimeStamp, adDBDate
                rsnew.Fields.Append .Fields(I).Name, adDBDate, , adFldIsNullable
            Case adSmallInt, adBoolean
                rsnew.Fields.Append .Fields(I).Name, adSmallInt, , adFldIsNullable
            Case Else
                rsnew.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
        End Select
    Next
End With

rsnew.Fields.Append "Registra", adSmallInt, , adFldIsNullable
rsnew.Open , , adOpenKeyset, adLockBatchOptimistic

rs.Close
Set rs = Nothing


Set rsAna = New ADODB.Recordset
rsAna.CursorLocation = adUseClient

rsAna.Fields.Append "IDAnagraficaCliente", adInteger, , adFldIsNullable
rsAna.Fields.Append "RitenutaAcconto", adSmallInt, , adFldIsNullable
rsAna.Fields.Append "IDSitoPerAnagrafica", adInteger, adFldIsNullable

rsAna.Open , , adOpenKeyset, adLockBatchOptimistic


End Sub

Private Function GET_PARAMETRO_AZIENDA_LONG(IDFiliale As Long, NomeCampo As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_POParametriAzienda "
sSQL = sSQL & " WHERE IDFiliale=" & IDFiliale

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PARAMETRO_AZIENDA_LONG = 0
Else
    GET_PARAMETRO_AZIENDA_LONG = fnNotNullN(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_BLOCCO_CLIENTE(IDAnaFatt As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT BloccoEmissioneDoc FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnaFatt
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_BLOCCO_CLIENTE = 0
Else
    GET_BLOCCO_CLIENTE = fnNotNullN(rs!BloccoEmissioneDoc)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub DefaultImportazione()
    Dim rs As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT * From RV_POParametriDefault Where IDAzienda=" & TheApp.IDFirm
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        
        FLAG_RAGGR_ALTRA_DEST = fnNotNullN(rs!RaggrAltraDestinazioneFattAdd)
        FLAG_RAGGR_CORPO_ALTRA_DEST = fnNotNullN(rs!RaggrCorpoAltraDestinazioneFattAdd)
        FLAG_RAGGR_CORPO_INT = fnNotNullN(rs!RaggrCorpoInterventoFattAdd)
        FLAG_SOVRAS_DESCRIZIONE = fnNotNullN(rs!SovrascriviDescrizioneArticolo)
    Else
        FLAG_RAGGR_ALTRA_DEST = 0
        FLAG_RAGGR_CORPO_ALTRA_DEST = 0
        FLAG_RAGGR_CORPO_INT = 0
        FLAG_SOVRAS_DESCRIZIONE = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    Me.chkRaggrCorpoAltraDest.Value = Abs(FLAG_RAGGR_CORPO_ALTRA_DEST)
    Me.chkRaggrCorpoInt.Value = Abs(FLAG_RAGGR_CORPO_INT)
    Me.chkRaggrSitoAna.Value = Abs(FLAG_RAGGR_ALTRA_DEST)
    Me.Check1.Value = Abs(FLAG_SOVRAS_DESCRIZIONE)
End Sub
