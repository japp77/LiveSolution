VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#10.15#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#9.1#0"; "DMTDataCmb.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmImportiNonStandard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importi non standard per gli anni successivi"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DMTDATETIMELib.dmtDate txtData 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   503
      _StockProps     =   253
      BackColor       =   16777215
      Appearance      =   1
   End
   Begin DMTDataCmb.DMTCombo cboIstat 
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DMTEDITNUMLib.dmtCurrency txtImporto 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   503
      _StockProps     =   253
      Text            =   "€ 0"
      BackColor       =   16777215
      Appearance      =   1
      UseSeparator    =   -1  'True
      CurrencySymbol  =   "€"
      AllowEmpty      =   0   'False
      DecFinalZeros   =   -1  'True
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalva 
      Caption         =   "Salva"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuovo 
      Caption         =   "Nuovo"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin DmtGridCtl.DmtGrid GrigliaImportiNonStandard 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3413
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Label Label1 
      Caption         =   "Adeguamento istat "
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   9
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Importo"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Data rinnovo"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmImportiNonStandard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsImporti As DmtADOLib.adoResultset
Private Nuovo As Integer
Private Link_ImportiNonStd As Long
Private NuovaData As String


Private Sub cmdNuovo_Click()


Me.txtData.Text = ""

If rsImporti.EOF Then
    Me.txtData.Text = CDate(frmMain.txtDataScadenzaPerRinnovo) + 1
Else
    SettaggioGriglia
    rsImporti.MoveFirst
    Me.txtData.Text = DateAdd("m", Mesi_Rinnovo_Contratto, rsImporti!Data)
End If

Me.txtImporto.Text = 0
Me.cboIstat.WriteOn 0
Nuovo = 0
End Sub

Private Sub cmdSalva_Click()
    Dim sSQL As String
    If Nuovo = 0 Then
        sSQL = "INSERT INTO RV_POContrattoNonStandard ("
        sSQL = sSQL & "IDRV_POContrattoNonStandard, IDRV_POContratto, Data, Importo, Mese, Anno,"
        sSQL = sSQL & "IDAdeguamentoIstat) VALUES ("
        sSQL = sSQL & fnGetNewKey("RV_POContrattoNonStandard", "IDRV_POContrattoNonStandard") & ", "
        sSQL = sSQL & Link_Contratto & ", "
        sSQL = sSQL & fnNormDate(Me.txtData.Text) & ", "
        sSQL = sSQL & fnNormNumber(Me.txtImporto.Value) & ", "
        sSQL = sSQL & DatePart("m", Me.txtData.Text) & ", "
        sSQL = sSQL & DatePart("yyyy", Me.txtData.Text) & ", "
        sSQL = sSQL & Me.cboIstat.CurrentID & ")"
         
    Else
        sSQL = "UPDATE RV_POContrattoNonStandard SET "
        sSQL = sSQL & "Data=" & fnNormDate(Me.txtData.Text) & ", "
        sSQL = sSQL & "Importo=" & fnNormNumber(Me.txtImporto.Value) & ", "
        sSQL = sSQL & "Mese=" & DatePart("m", Me.txtData.Text) & ", "
        sSQL = sSQL & "Anno=" & DatePart("yyyy", Me.txtData.Text) & ", "
        sSQL = sSQL & "IDAdeguamentoIstat=" & Me.cboIstat.CurrentID & " "
        sSQL = sSQL & "WHERE IDRV_POContrattoNonStandard=" & Link_ImportiNonStd

    End If
    Cn.Execute sSQL
    
    SettaggioGriglia
    
End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    fncIstat
    SettaggioGriglia
    Nuovo = 0
End Sub
Private Sub fncIstat()

    With Me.cboIstat
        Set .Database = Cn
        .AddFieldKey "IDRV_POIstat"
        .DisplayField = "Istat"
        .SQL = "SELECT * FROM RV_POIstat"
        .Fill
    End With
    
End Sub

Private Sub SettaggioGriglia()

    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    
    sSQL = "Select * FROM RV_POContrattoNonStandard WHERE IDRV_POContratto=" & Link_Contratto
    sSQL = sSQL & " ORDER BY Data DESC"
    
    
    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
    
        Set rsImporti = Cn.OpenResultset(sSQL)
            Set rsEvent = rsImporti.Data
        
        With Me.GrigliaImportiNonStandard
            .ColumnsHeader.Clear
                    
                    .ColumnsHeader.Add "Data", "Data rinnovo", dgDate, True, 2000, dgAlignleft
                    .ColumnsHeader.Add "Importo", "Importo", dgDouble, True, 2000, dgAlignleft
                    .ColumnsHeader.Add "IDAdeguamentoIstat", "Istat", dgInteger, True, 1000, dgAlignleft
                    
            Set .Recordset = rsImporti.Data
            .Refresh
            If rsImporti.EOF Then
                Me.txtData.Text = CDate(frmMain.txtDataScadenzaPerRinnovo) + 1
            End If
            
        End With
    
    Cn.CursorLocation = OLDCursor
End Sub



Private Sub GrigliaImportiNonStandard_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    Me.txtData.Text = Me.GrigliaImportiNonStandard.AllColumns("Data")
    Me.txtImporto.Value = Me.GrigliaImportiNonStandard.AllColumns("Importo")
    Me.cboIstat.WriteOn Me.GrigliaImportiNonStandard.AllColumns("IDAdeguamentoIstat")
    Link_ImportiNonStd = Me.GrigliaImportiNonStandard.AllColumns("IDRV_POContrattoNonStandard")
    
    Nuovo = 1
End Sub
