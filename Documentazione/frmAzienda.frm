VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.2#0"; "DmtGridCtl.ocx"
Begin VB.Form frmAzienda 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "AZIENDA"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DmtGridCtl.DmtGrid Griglia 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5106
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
Attribute VB_Name = "frmAzienda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Sub Form_Load()
    
    GET_GRIGLIA_AZIENDA

End Sub
Private Sub GET_GRIGLIA_AZIENDA()
On Error GoTo ERR_fnGrigliaAssegnazione
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
    
    sSQL = "SELECT Anagrafica.IDAnagrafica, Azienda.IDAzienda, Anagrafica.Anagrafica, Anagrafica.PartitaIva, Anagrafica.Indirizzo, Comune.Comune, Provincia.Provincia, "
    sSQL = sSQL & "Anagrafica.Cap "
    sSQL = sSQL & "FROM Provincia RIGHT OUTER JOIN "
    sSQL = sSQL & "Comune ON Provincia.IDProvincia = Comune.IDProvincia RIGHT OUTER JOIN "
    sSQL = sSQL & "Anagrafica INNER JOIN "
    sSQL = sSQL & "Azienda ON Anagrafica.IDAnagrafica = Azienda.IDAnagrafica ON Comune.IDComune = Anagrafica.IDComune "
    
    
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, CnDMT.InternalConnection
        
        With Me.Griglia
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
                .ColumnsHeader.Add "IDAnagrafica", "IDAnagrafica", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "Anagrafica", "Anagrafica", dgchar, True, 3000, dgAlignleft
                .ColumnsHeader.Add "Indirizzo", "Indirizzo", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "Comune", "Comune", dgchar, True, 1500, dgAlignleft
                .ColumnsHeader.Add "Provincia", "Provincia", dgchar, True, 1000, dgAlignleft
                .ColumnsHeader.Add "Cap", "Cap", dgchar, True, 1000, dgAlignleft
                
            
            Set .Recordset = rsGriglia
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaAssegnazione:
    MsgBox Err.Description, vbCritical, "Reperimento dati assegnazione"
End Sub


Private Sub Griglia_DblClick()
    LINK_ANAGRAFICA = fnNotNullN(Me.Griglia("IDAnagrafica").Value)
    FrmPresentazione.txtIDAnagrafica.Value = LINK_ANAGRAFICA
    FrmPresentazione.txtRagioneSociale.Text = Me.Griglia("Anagrafica").Value
    FrmPresentazione.lblPartitaIva.Caption = "Partita I.V.A.: " & Me.Griglia("PartitaIVA").Value
    PARTITA_IVA_LICENZA = Me.Griglia("PartitaIva").Value
    FrmPresentazione.lblIndirizzo.Caption = Me.Griglia("Indirizzo").Value
    FrmPresentazione.lblComune.Caption = Me.Griglia("Comune").Value & " (" & Me.Griglia("Provincia").Value & ") " & Me.Griglia("Cap").Value
    
    COMUNE = fnNotNull(Me.Griglia("Comune").Value)
    PROVINCIA = fnNotNull(Me.Griglia("Provincia").Value)
    CAP = fnNotNull(Me.Griglia("Cap").Value)
    Unload Me
        
End Sub
