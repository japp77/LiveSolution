VERSION 5.00
Begin VB.Form frmProva 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importazione lotti"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Procedi"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label lblInfo2 
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   5175
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Nome tabella"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
End
Attribute VB_Name = "frmProva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim IDArticolo As Long

sSQL = "SELECT * FROM " & Me.Text1.Text

Set rs = CnDMT.OpenResultset(sSQL)

sSQL = "SELECT * FROM LottoArticolo "
sSQL = sSQL & "WHERE IDLottoArticolo=0"
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rs.EOF
    DoEvents
    Me.lblInfo.Caption = fnNotNull(rs!F1) & " - " & fnNotNull(rs!F2)
    
    IDArticolo = GET_LINK_ARTICOLO(fnNotNull(rs!F1))
    If IDArticolo > 0 Then
        If GET_ESISTENZA_LOTTO_ARTICOLO(IDArticolo, fnNotNull(rs!F4)) = False Then
            rsNew.AddNew
                rsNew!IDLottoArticolo = fnGetNewKey("LottoArticolo", "IDLottoArticolo")
                rsNew!IDArticolo = IDArticolo
                rsNew!Codice = fnNotNull(rs!F4)
                rsNew!LottoArticolo = fnNotNull(rs!F4)
                If Len(Trim(fnNotNull(rs!F5))) = 0 Then
                    rsNew!DataScadenza = "31/12/2099"
                Else
                    rsNew!DataScadenza = rs!F5
                End If
                rsNew!Sospeso = 0
            rsNew.Update
        End If
        '''AGGIORNAMENTO LOTTI ARTICOLO'''''''''''''''''''''''''''
        sSQL = "UPDATE Articolo SET "
        sSQL = sSQL & "GestioneLotti=" & fnNormBoolean(1)
        sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
        CnDMT.Execute sSQL
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
    
    
    
DoEvents
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
Me.lblInfo2.Caption = "OPERAZIONE COMPLETATA"
Me.lblInfo.Caption = ""
End Sub
Private Function GET_LINK_ARTICOLO(CodiceArticolo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDArticolo FROM Articolo "
sSQL = sSQL & "WHERE CodiceArticolo=" & fnNormString(CodiceArticolo)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ARTICOLO = 0
Else
    GET_LINK_ARTICOLO = fnNotNullN(rs!IDArticolo)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_ESISTENZA_LOTTO_ARTICOLO(IDArticolo As Long, CodiceLotto As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDLottoArticolo FROM LottoArticolo "
sSQL = sSQL & "WHERE Codice=" & fnNormString(CodiceLotto)
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_LOTTO_ARTICOLO = False
Else
    GET_ESISTENZA_LOTTO_ARTICOLO = True
End If

rs.CloseResultset
Set rs = Nothing

End Function
