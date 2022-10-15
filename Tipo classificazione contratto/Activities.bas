Attribute VB_Name = "Activities"
Option Explicit

Public Function fnExtractNameFromTag(ByRef Tool As ActiveBar3LibraryCtl.Tool, ByVal ToolTag As ExportConstants, ByVal sRegistrySection As String) As String
    Dim oRes As Resource
    
    Set oRes = New Resource
    
    Select Case ToolTag
        Case ExportConstants.PDF
            fnExtractNameFromTag = "ExportPDF"
            Tool.TagVariant = ExportConstants.PDF
            If GetSetting(REGISTRY_KEY, sRegistrySection & "Settings", "LargeIcon", False) = False Then
                Tool.SetPicture 0, oRes.GetBitmap(IDB_ACROBAT_16), &HC0C0C0
            Else
                Tool.SetPicture 0, oRes.GetBitmap(IDB_ACROBAT_32), &HC0C0C0
            End If
        Case ExportConstants.Word
            fnExtractNameFromTag = "ExportWord"
            Tool.TagVariant = ExportConstants.Word
            If GetSetting(REGISTRY_KEY, sRegistrySection & "Settings", "LargeIcon", False) = False Then
                Tool.SetPicture 0, oRes.GetBitmap(IDB_STD_WORD16), &HC0C0C0
            Else
                Tool.SetPicture 0, oRes.GetBitmap(IDB_STD_WORD32), &HC0C0C0
            End If
        Case ExportConstants.Excel
            fnExtractNameFromTag = "ExportExcel"
            Tool.TagVariant = ExportConstants.Excel
            If GetSetting(REGISTRY_KEY, sRegistrySection & "Settings", "LargeIcon", False) = False Then
                Tool.SetPicture 0, oRes.GetBitmap(IDB_STD_EXCEL16), &HC0C0C0
            Else
                Tool.SetPicture 0, oRes.GetBitmap(IDB_STD_EXCEL32), &HC0C0C0
            End If
        Case ExportConstants.HTML
            fnExtractNameFromTag = "ExportHtml"
            Tool.TagVariant = ExportConstants.HTML
            If GetSetting(REGISTRY_KEY, sRegistrySection & "Settings", "LargeIcon", False) = False Then
                Tool.SetPicture 0, oRes.GetBitmap(IDB_STD_HTML16), &HC0C0C0
            Else
                Tool.SetPicture 0, oRes.GetBitmap(IDB_STD_HTML32), &HC0C0C0
            End If
    End Select
    
    Set oRes = Nothing
End Function


