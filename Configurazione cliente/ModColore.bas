Attribute VB_Name = "ModColore"
'Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpChoosecolor As udtCHOOSECOLOR) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
  (lpChoosecolor As ChooseColorStruct) As Long


Private Const CC_RGBINIT = &H1&
Private Const CC_FULLOPEN = &H2&
Private Const CC_PREVENTFULLOPEN = &H4&
Private Const CC_SHOWHELP = &H8&
Private Const CC_ENABLEHOOK = &H10&
Private Const CC_ENABLETEMPLATE = &H20&
Private Const CC_ENABLETEMPLATEHANDLE = &H40&
Private Const CC_SOLIDCOLOR = &H80&
Private Const CC_ANYCOLOR = &H100&
Private Const CLR_INVALID = &HFFFF

Private Type udtCHOOSECOLOR
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Type ChooseColorStruct
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As Long
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

'Public Function GetColor() As Long
'  Dim rc As Long
'  Dim pChooseColor As udtCHOOSECOLOR'

'  'In the original source code, this is dimensioned as a dynamic
'  'array, but some experimenting yield this array as having a
'  'maximum of 64 elements (4 bytes per custom box color)
'  Dim CustomColors(63) As Byte

'  '********************************************
'  'This section defines the box colors using the Red, Green
'  'Blue value assignments as indicated. There are 16 custom
'  'color boxes so there will be 16 groups of 4 bytes as shown
'  '********************************************
'  'Define Colors For 1st Box, Top Row
'  CustomColors(0) = 110  'Red Value
'  CustomColors(1) = 34  'Green Value
'  CustomColors(2) = 255  'Blue Value
'  CustomColors(3) = 0  'Always zero
'
'  'Define Colors For 2nd Box, Top Row
'  CustomColors(4) = 230  'Red Value
'  CustomColors(5) = 255  'Green Value
'  CustomColors(6) = 10  'Blue Value
'  CustomColors(7) = 0  'Always zero
'
'  'Define Colors For 3rd Box, Top Row
'  '  etc.
'  '  etc.
''

'  With pChooseColor
'    .hwndOwner = 0
'    .hInstance = App.hInstance
'    .lpCustColors = StrConv(CustomColors, vbUnicode)
'    .flags = 0
'    .lStructSize = Len(pChooseColor)
'  End With

'  rc = ChooseColor(pChooseColor)
'
'  If rc Then
'    GetColor = pChooseColor.rgbResult
'  End If

'End Function
Public Function ShowColorDialog(Optional ByVal hParent As Long, _
  Optional ByVal bFullOpen As Boolean, Optional ByVal InitColor As Long) _
  As Long
  Dim CC As ChooseColorStruct
  Dim aColorRef(15) As Long
  Dim lInitColor As Long
 
  ' translate the initial OLE color to a long value
  If InitColor <> 0 Then
    If OleTranslateColor(InitColor, 0, lInitColor) Then
      lInitColor = CLR_INVALID
    End If
  End If
  
  'fill the ChooseColorStruct struct
  With CC
    .lStructSize = Len(CC)
    .hwndOwner = hParent
    .lpCustColors = VarPtr(aColorRef(0))
    .rgbResult = lInitColor
    .flags = CC_SOLIDCOLOR Or CC_ANYCOLOR Or CC_RGBINIT Or IIf(bFullOpen, _
      CC_FULLOPEN, 0)
  End With
  
  ' Show the dialog
  If ChooseColor(CC) Then
    'if not canceled, return the color
    ShowColorDialog = CC.rgbResult
  Else
    'else return -1
    ShowColorDialog = -1
  End If
End Function
