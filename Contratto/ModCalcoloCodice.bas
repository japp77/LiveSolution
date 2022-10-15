Attribute VB_Name = "ModCalcoloCodice"
Public DEMO As Boolean

Public Function GET_CODICE_SBLOCCO(CodiceDiamante As String, PartitaIVA As String, CodiceProdotto As String) As String
Dim CodiceDiamante_Local As String
Dim PartitaIva_Local As String
Dim Risultato As String


CodiceDiamante_Local = GET_CODICE_DIAMANTE(CodiceDiamante)
PartitaIva_Local = GET_CODICE_PARTITA_IVA(PartitaIVA)


Risultato = Risultato & Mid(CodiceProdotto, 1, 1)
Risultato = Risultato & Mid(CodiceDiamante_Local, 1, 1)
Risultato = Risultato & Mid(CodiceDiamante_Local, 2, 1)
Risultato = Risultato & Mid(PartitaIva_Local, 1, 1)
Risultato = Risultato & Mid(PartitaIva_Local, 2, 1)
Risultato = Risultato & Mid(PartitaIva_Local, 3, 1)
Risultato = Risultato & Mid(CodiceDiamante_Local, 3, 1)
Risultato = Risultato & Mid(CodiceProdotto, 2, 1)


GET_CODICE_SBLOCCO = Risultato

End Function
Private Function GET_CODICE_DIAMANTE(CodiceDiamante As String) As String
Dim I As Integer
Dim SommaLettere As Long
Dim SommaNumeri As Long
Dim Risultato As Long


GET_CODICE_DIAMANTE = ""
SommaLettere = 0
SommaNumeri = 0

For I = 1 To Len(CodiceDiamante)
    If IsNumeric(Mid(CodiceDiamante, I, 1)) Then
        SommaNumeri = SommaNumeri + Mid(CodiceDiamante, I, 1)
    Else
        SommaLettere = SommaLettere + GET_N_LETTERA(Mid(CodiceDiamante, I, 1))
    End If
    
Next I
Risultato = SommaLettere - SommaNumeri

GET_CODICE_DIAMANTE = Risultato

If Len(GET_CODICE_DIAMANTE) < 3 Then
    For I = Len(GET_CODICE_DIAMANTE) To 2
        GET_CODICE_DIAMANTE = "0" & GET_CODICE_DIAMANTE
    Next
End If


End Function

Private Function GET_CODICE_PARTITA_IVA(PartitaIVA As String) As String
Dim I As Integer
Dim SommaNumeriPari As Long
Dim SommaNumeriDispari As Long
Dim Risultato As Long

GET_CODICE_PARTITA_IVA = ""
SommaNumeriPari = 0
SommaNumeriDispari = 0


For I = 1 To Len(PartitaIVA)
    If I Mod 2 = 0 Then
        SommaNumeriPari = SommaNumeriPari + Mid(PartitaIVA, I, 1)
    Else
        SommaNumeriDispari = SommaNumeriDispari + Mid(PartitaIVA, I, 1)
    End If
    
Next I

Risultato = SommaNumeriPari + SommaNumeriDispari
GET_CODICE_PARTITA_IVA = Risultato
If Len(CStr(GET_CODICE_PARTITA_IVA)) < 3 Then
    For I = Len(CStr(GET_CODICE_PARTITA_IVA)) To 2
        GET_CODICE_PARTITA_IVA = "0" & GET_CODICE_PARTITA_IVA
    Next
End If


End Function


Private Function GET_N_LETTERA(Lettera As String) As Long

Select Case Lettera
    Case "A"
        GET_N_LETTERA = 1
    Case "B"
        GET_N_LETTERA = 2
    Case "C"
        GET_N_LETTERA = 3
    Case "D"
        GET_N_LETTERA = 4
    Case "E"
        GET_N_LETTERA = 5
    Case "F"
        GET_N_LETTERA = 6
    Case "G"
        GET_N_LETTERA = 7
    Case "H"
        GET_N_LETTERA = 8
    Case "I"
        GET_N_LETTERA = 9
    Case "J"
        GET_N_LETTERA = 10
    Case "K"
        GET_N_LETTERA = 11
    Case "L"
        GET_N_LETTERA = 12
    Case "M"
        GET_N_LETTERA = 13
    Case "N"
        GET_N_LETTERA = 14
    Case "O"
        GET_N_LETTERA = 15
    Case "P"
        GET_N_LETTERA = 16
    Case "Q"
        GET_N_LETTERA = 17
    Case "R"
        GET_N_LETTERA = 18
    Case "S"
        GET_N_LETTERA = 19
    Case "T"
        GET_N_LETTERA = 20
    Case "U"
        GET_N_LETTERA = 21
    Case "V"
        GET_N_LETTERA = 22
    Case "W"
        GET_N_LETTERA = 23
    Case "X"
        GET_N_LETTERA = 24
    Case "Y"
        GET_N_LETTERA = 25
    Case "Z"
        GET_N_LETTERA = 26
End Select


End Function
