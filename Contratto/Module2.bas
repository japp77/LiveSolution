Attribute VB_Name = "Module2"
Option Explicit

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Public Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_RETURNONLYFSDIRS = 1

Private Const MAX_PATH = 260

Private StartFolder As String

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Function BrowseForFolder(DefaultFolder As String) As String
On Error GoTo ERR_BrowseForFolder
   Dim iNull As Integer, lpIDList As Long
   Dim sPath As String, udtBI As BrowseInfo

   StartFolder = DefaultFolder

   With udtBI
       'Set the owner window
       .hWndOwner = 0 'OwnerForm.hwnd
       'lstrcat appends the two strings and returns the memory address
       .lpszTitle = lstrcat("Select Search Path", "")
       '.ulFlags = BIF_RETURNONLYFSDIRS
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE
       'Return only if the user selected a directory
       .lpfnCallback = Address_Of(AddressOf BrowseCallbackProc)
   End With

   'Show the 'Browse for folder' dialog
   lpIDList = SHBrowseForFolder(udtBI)
   If lpIDList Then
       sPath = String$(MAX_PATH, 0)
       'Get the path from the IDList
       SHGetPathFromIDList lpIDList, sPath
       'free the block of memory
       CoTaskMemFree lpIDList
       iNull = InStr(sPath, vbNullChar)
       If iNull Then
           sPath = Left$(sPath, iNull - 1)
       End If
   End If
   BrowseForFolder = sPath

Exit Function
ERR_BrowseForFolder:
    MsgBox Err.Description, vbCritical, "BrowseForFolder"
End Function


Private Function Address_Of(ByVal n As Long) As Long
On Error GoTo ERR_Address_Of
   Address_Of = n

Exit Function
ERR_Address_Of:
    MsgBox Err.Description, vbCritical, "Address_Of"
End Function

Private Function BrowseCallbackProc(ByVal hwnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal lParam As Long, _
                                    ByVal lpData As Long) As Long
    On Error GoTo ERR_BrowseCallbackProc
   Const WM_USER = &H400&
   Const BFFM_INITIALIZED = 1
   Const BFFM_SETSELECTIONA = (WM_USER + 102)
   
   Dim default_path() As Byte
      
   If uMsg = BFFM_INITIALIZED Then
      default_path = StrConv(StartFolder, vbFromUnicode)
      SendMessage hwnd, BFFM_SETSELECTIONA, 1&, ByVal VarPtr(default_path(0))
   End If
   
   Exit Function
ERR_BrowseCallbackProc:
ERR_BrowseForFolder:
    MsgBox Err.Description, vbCritical, "BrowseCallbackProc"
End Function

