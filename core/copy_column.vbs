' VBScript to copy a column from one Excel file to another
' Copies contiguous blocks of data to improve performance

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

Sub CopyBlocks(startRow, endRow)
    Dim r, blockStart, blockEnd, val, destStart, srcRange, destCell
    r = startRow
    Do While r <= endRow
        ' Find start of next non-empty block
        Do While r <= endRow
            val = wsSrc.Cells(r, srcCol).Value
            If Not IsEmpty(val) And Trim(CStr(val)) <> "" Then Exit Do
            r = r + 1
        Loop
        If r > endRow Then Exit Do
        blockStart = r
        ' Find end of the block
        Do While r <= endRow
            val = wsSrc.Cells(r, srcCol).Value
            If IsEmpty(val) Or Trim(CStr(val)) = "" Then Exit Do
            r = r + 1
        Loop
        blockEnd = r - 1

        destStart = blockStart
        If copyByRow = 0 Then destStart = blockStart + headerRow

        Set srcRange = wsSrc.Range(wsSrc.Cells(blockStart, srcCol), wsSrc.Cells(blockEnd, srcCol))
        Set destCell = wsDest.Cells(destStart, destCol)

        If preserveFmt = 1 Then
            srcRange.Copy destCell
        Else
            destCell.Resize(srcRange.Rows.Count, 1).Value = srcRange.Value
        End If
    Loop
End Sub

If headerRow > 0 Then
    CopyBlocks 1, headerRow
End If

If lastRow >= headerRow + 2 Then
    CopyBlocks headerRow + 2, lastRow
End If

wbDest.Save
wbSrc.Close False
wbDest.Close True
excel.Quit

