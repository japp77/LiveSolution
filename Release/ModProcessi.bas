Attribute VB_Name = "ModProcessi"
Option Explicit
Private Const TH32CS_SNAPPROCESS = &H2
Private Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "Kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "Kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "Kernel32" (ByVal hObject As Long)

Public Function CONTROLLO_PROCESSO_ATTIVO(Processo As String) As Boolean
Dim hSnapShot As Long
Dim uProcess As PROCESSENTRY32
Dim lngRet As Long
   
hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
uProcess.dwSize = Len(uProcess)
lngRet = Process32First(hSnapShot, uProcess)
CONTROLLO_PROCESSO_ATTIVO = False
Do While lngRet
    If Processo = Left$(uProcess.szExeFile, InStr(1, uProcess.szExeFile, vbNullChar) - 1) Then
        CONTROLLO_PROCESSO_ATTIVO = True
    End If
    lngRet = Process32Next(hSnapShot, uProcess)
Loop
CloseHandle hSnapShot
End Function
