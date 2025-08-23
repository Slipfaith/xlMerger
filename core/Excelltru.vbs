' VBS script to create a copy of an Excel file with the suffix "totr"
' This script is intended to be called from Python using cscript.exe

Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
    WScript.Quit 1
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set excel = CreateObject("Excel.Application")
excel.Visible = False

For Each filePath In objArgs
    If fso.FileExists(filePath) Then
        ext = LCase(fso.GetExtensionName(filePath))
        If ext = "xls" Or ext = "xlsx" Then
            Set wbOriginal = excel.Workbooks.Open(filePath)
            origFolder = fso.GetParentFolderName(filePath)
            origName = fso.GetBaseName(filePath)
            origExt = fso.GetExtensionName(filePath)
            newFileName = origFolder & "\" & origName & "_totr." & origExt

            Set wbNew = excel.Workbooks.Add
            Do While wbNew.Worksheets.Count > 1
                wbNew.Worksheets(1).Delete
            Loop

            For Each sheet In wbOriginal.Sheets
                sheet.Copy , wbNew.Sheets(wbNew.Sheets.Count)
            Next

            If wbNew.Sheets.Count > wbOriginal.Sheets.Count Then
                wbNew.Sheets(1).Delete
            End If

            wbNew.SaveAs newFileName
            wbNew.Close False
            wbOriginal.Close False
        End If
    End If
Next

excel.Quit
