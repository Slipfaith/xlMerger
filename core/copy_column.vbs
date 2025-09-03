' VBScript to copy a column from one Excel file to another
Set args = WScript.Arguments
If args.Count < 9 Then
    WScript.Echo "Usage: copy_column.vbs srcFile srcSheet srcCol destFile destSheet destCol headerRow copyByRow preserveFormatting"
    WScript.Quit 1
End If

srcFile = args(0)
srcSheet = args(1)
srcCol = CLng(args(2))
destFile = args(3)
destSheet = args(4)
destCol = CLng(args(5))
headerRow = CLng(args(6))
copyByRow = CLng(args(7))
preserveFmt = CLng(args(8))

Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

Set wbSrc = excel.Workbooks.Open(srcFile)
Set wsSrc = wbSrc.Sheets(srcSheet)
Set wbDest = excel.Workbooks.Open(destFile)
Set wsDest = wbDest.Sheets(destSheet)

lastRow = wsSrc.Cells(wsSrc.Rows.Count, srcCol).End(-4162).Row  ' xlUp

If headerRow > 0 Then
    For r = 1 To headerRow
        Set cellSrc = wsSrc.Cells(r, srcCol)
        val = cellSrc.Value
        If Not IsEmpty(val) And Trim(CStr(val)) <> "" Then
            destRow = r
            If copyByRow = 0 Then destRow = r + headerRow
            If preserveFmt = 1 Then
                cellSrc.Copy wsDest.Cells(destRow, destCol)
            Else
                wsDest.Cells(destRow, destCol).Value = val
            End If
        End If
    Next
End If

startRow = headerRow + 2
For r = startRow To lastRow
    Set cellSrc = wsSrc.Cells(r, srcCol)
    val = cellSrc.Value
    If Not IsEmpty(val) And Trim(CStr(val)) <> "" Then
        destRow = r
        If copyByRow = 0 Then destRow = r + headerRow
        If preserveFmt = 1 Then
            cellSrc.Copy wsDest.Cells(destRow, destCol)
        Else
            wsDest.Cells(destRow, destCol).Value = val
        End If
    End If
Next

wbDest.Save
wbSrc.Close False
wbDest.Close True
excel.Quit
