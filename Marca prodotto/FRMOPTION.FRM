VERSION 5.00
Object = "{00B55B47-CFD3-11D2-A1DE-00105ACC6020}#6.0#0"; "DMTOPTIONSLIB.OCX"
Begin VB.Form frmOption 
   Caption         =   "Opzioni"
   ClientHeight    =   4005
   ClientLeft      =   1725
   ClientTop       =   1560
   ClientWidth     =   3840
   Icon            =   "FRMOPTION.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4005
   ScaleWidth      =   3840
   StartUpPosition =   1  'CenterOwner
   Begin DmtOptionsLib.DmtOptions DmtOptions1 
      Height          =   3840
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   6773
      OpzioniTab      =   1
      InsertOption    =   0   'False
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Riferimento al form chiamante
Private m_FormClient As Form

'Opzioni gestite da DmtOptions
Private Enum neOptions
    neChangeLanguage = 1
    neStabusBarVisibility = 2
    neIDFieldsVisibility = 3
End Enum


Public Property Get FormClient() As Form
    Set FormClient = m_FormClient
End Property

Public Property Set FormClient(ByVal vNewValue As Form)
    Set m_FormClient = vNewValue
End Property

'**+
'Nome: Form_Load
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Carica tutte le impostazioni per il controllo sul
'form
'
'**/
Private Sub Form_Load()
    '------------------------------
    'Inizializzazioni di DmtOptions
    '------------------------------
    
    '-------------------------------------------------------------
    'NON viene visualizzata la check sulla visibilità dei campi ID
    DmtOptions1.IDFieldsOptionVisible = True
    
    DmtOptions1.ApplicationName = FormClient.Application.Name
    Set DmtOptions1.Database = FormClient.Application.Database.Connection
    DmtOptions1.NativeLanguage = NATIVE_LANGUAGE
    DmtOptions1.Refresh
    
    If DmtOptions1.IDLanguage <> DmtOptions1.NativeLanguage Then
        gResource.IDCurrentLanguage = DmtOptions1.IDLanguage
        Me.Caption = gResource.GetMessage(LBL_OPTIONS)
    End If

    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
End Sub


'**+
'Nome: DmtOptions1_OnCmdClick
'
'Parametri:
'DmtOptionsLib.dmtButton indica quale bottone è stato premuto.
'
'DmtOptionsLib.clsOptions è una collection di clsOption.
'Quest'ultimo oggetto contiene 3 informazioni:
' 1. IDOption (chiave che identifica l'opzione)
' 2. OldValue (valore dell'opzione prima della modifica)
' 3. NewValue (valore attuale dell'opzione).
'
'Valori di ritorno:
'
'Funzionalità:
'Controlla tutti i valori restituiti dal controllo e aggiorna le impostazioni
'del programma
'
'**/
Private Sub DmtOptions1_OnCmdClick(ByVal Button As DmtOptionsLib.dmtButton, ByVal OptionsCollection As DmtOptionsLib.clsOptions)
    Dim objOption As DmtOptionsLib.clsOption

    'Comunica il linguaggio da utilizzare nella lettura delle stringhe
    'dal file di risorse
    gResource.IDCurrentLanguage = DmtOptions1.IDLanguage

    For Each objOption In OptionsCollection
        'Se si preme Annulla si deve soltanto chiudere il form
        If Button <> btUndo Then
            Select Case objOption.IDOption
        
                Case neChangeLanguage  'Cambia lingua
                    'Setta i nuovi ToolTipText della Toolbar
                    'e le Caption dei menu
                    FormClient.ChangeMenuLanguage
                    FormClient.ChangeToolBarLanguage
                    'Reinizializza le stringhe del form principale
                    FormClient.ChangeStringsLanguage
                    Me.Caption = gResource.GetMessage(LBL_OPTIONS)
        
                Case neStabusBarVisibility  'Visibilità StatusBar
                    FormClient.SetStatusBarVisibility objOption.NewValue
        
                Case neIDFieldsVisibility  'Mostra/Nascondi campi ID
                    FormClient.SetVisibilityIDFields DmtOptions1.IDFieldsOptionValue
                    
            End Select
        End If
    Next objOption

    'Se non si è premuto il pulsante Applica viene scaricato il form
    If Button <> btApply Then
        Unload Me
    End If
End Sub



'Libera il riferimento al Form chiamante
Private Sub Form_Terminate()
    Set m_FormClient = Nothing
End Sub




