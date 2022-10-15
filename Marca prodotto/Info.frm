VERSION 5.00
Object = "{EA24283A-2440-11D1-8243-0040053260B6}#2.0#0"; "DMTAbout.ocx"
Begin VB.Form frmInformazioni 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informazioni su Diamante"
   ClientHeight    =   4905
   ClientLeft      =   5355
   ClientTop       =   2325
   ClientWidth     =   6480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4905
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin About.About_on Info 
      Height          =   4845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   8546
   End
End
Attribute VB_Name = "frmInformazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim sString As String
    sString = "Attenzione: questo programma è tutelato dalle leggi sul copyright,"
    sString = sString & " dalle leggi sui diritti d'autore e dalle disposizioni dei trattati"
    sString = sString & " internazionali. La riproduzione o distribuzione totale o parziale di questo "
    sString = sString & " programma, sarà perseguita civilmente e penalmente a termini di legge."
    Info.Comments = sString
    Info.CompanyName = App.CompanyName
    Info.Version = App.FileDescription & " " & App.Major & "." & App.Minor & "." & App.Revision
    Info.ProductName = App.ProductName
    Info.Show
End Sub

Private Sub Info_Caption()
    Me.Caption = Info.InfoCaption
End Sub

Private Sub Info_Unload()
    Unload Me
End Sub
