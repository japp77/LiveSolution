VERSION 5.00
Object = "{51FF1814-CD6C-11D2-A1DE-00A0244FF30F}#1.0#0"; "DMTPREFSTOOLBARSLIB.OCX"
Begin VB.Form frmToolBars 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Barre degli strumenti"
   ClientHeight    =   3585
   ClientLeft      =   2340
   ClientTop       =   2640
   ClientWidth     =   7005
   Icon            =   "Toolbars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3585
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin DMTPrefsToolbarsLib.DMTPrefsToolbars DMTPrefsToolbars 
      Height          =   3285
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   5794
   End
End
Attribute VB_Name = "frmToolBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************************
' frmToolBars utilizza il nuovo componente DmtPrefsToolbarsLib.OCX
'******************************************************************

Option Explicit

'Riferimento al form chiamante
Private m_FormClient As Form

Public Property Get FormClient() As Form
    Set FormClient = m_FormClient
End Property

Public Property Set FormClient(ByVal vNewValue As Form)
    Set m_FormClient = vNewValue
End Property


Private Sub Form_Load()
    Dim mBand As Band
    Dim bVisible As Boolean

    On Error Resume Next
    
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    
    'Visualizza o meno la toolbar del DocTypeExplorer in base all'impostazione attuale
    'nel registry, indipendentemente dalla visibilità del DocTypeExplorer.
    bVisible = GetSetting("Diamante", TheApp.Name & "Settings", "Proprietà tipo documento", True)
    m_FormClient.BarMenu.Bands("Proprietà tipo documento").Visible = bVisible
    m_FormClient.BarMenu.RecalcLayout

    
    ' Viene notificato il nome dell'applicazione
     DMTPrefsToolbars.ApplicationName = TheApp.Name
    
    ' inserisce nella Listbox "Barre degli strumenti"
    ' tutte le barre degli strumenti usate dal programma
    ' escludendo BAND_CLOSE_PREVIEW che è un caso a parte.
    For Each mBand In FormClient.BarMenu.Bands
        If mBand.Type = DDBTNormal And mBand.Name <> BAND_CLOSE_PREVIEW Then
        DMTPrefsToolbars.AddToolbar mBand.Name
        End If
    Next mBand
    
    ' Notifica il codice per il controllo sulla lingua corrente
    DMTPrefsToolbars.LanguageNative = NATIVE_LANGUAGE
    
    ' Inizializza il controllo
    DMTPrefsToolbars.Refresh
End Sub

' E' possibile scaricare il form
Private Sub DMTPrefsToolbars_CanCloseForm()
    Unload Me
End Sub

'Libera il riferimento al Form chiamante
Private Sub Form_Terminate()
    Set m_FormClient = Nothing
End Sub


' Occorre impostare la grandezza delle icone in base al valore restituito Value
Private Sub DMTPrefsToolbars_LargeIconClick(ByVal Value As Boolean)
    FormClient.SetToolBarIcons Value
End Sub

' Occorre settare la proprietà Visible della barra Bands(Name)
Private Sub DMTPrefsToolbars_ToolbarItemCheck(ByVal Name As String, ByVal Visible As Boolean)
    FormClient.BarMenu.Bands(Name).Visible = Visible
    FormClient.BarMenu.RecalcLayout
End Sub

' Occorre impostare la visibilità dei ToolTips
Private Sub DMTPrefsToolbars_ToolTipClick(ByVal Value As Boolean)
    FormClient.BarMenu.DisplayToolTips = Value
End Sub

