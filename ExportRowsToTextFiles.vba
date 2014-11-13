Sub ExportRowsToTextFiles()
    Dim wb As Excel.Workbook, wbNew As Excel.Workbook
    Dim wsSource As Excel.Worksheet, wsTemp As Excel.Worksheet
    Dim row As Long, c As Long, columnCount As Long
    Dim filePath As String
    Dim fileName As String
    Dim rowRange As Range
    Dim cell As Range
    
    filePath = "C:\Temp\Epiphany\"
    columnCount = 29
    row = 2
    
    For Each cell In Range("A2", Range("A51"))
        Set rowRange = Range(cell.Address, Range(cell.Address).End(xlToRight))
    
        fileName = filePath & cell.Value  'set fileName value to be the value from the first column in source sheet
    
        Set wsSource = ThisWorkbook.Worksheets("Sheet1")
        
        ThisWorkbook.Worksheets.Add ThisWorkbook.Worksheets(1)
        Set wsTemp = ThisWorkbook.Worksheets(1)
        
        For d = 2 To columnCount ' iterate over the source sheet header row and put the values into the first column of target sheet
            wsTemp.Cells(d, 1).Value = wsSource.Cells(1, d).Value
        Next d

        For c = 2 To columnCount 'iterate over each cell in the source sheet row and put its data value into the column
            wsTemp.Cells(c, 2).Value = wsSource.Cells(row, c).Value
            If c = 13 Or c = 14 Or c = 16 Or c = 18 Or c = 20 Or c = 22 Or c = 24 Or c = 26 Or c = 28 Then
                wsTemp.Cells(c, 2).NumberFormat = "mm/dd/yyyy" ' Set date format on specific cells with date in them
            End If
        Next c

        wsTemp.Move
        Set wbNew = ActiveWorkbook
        Set wsTemp = wbNew.Worksheets(1)
        wbNew.SaveAs fileName & ".txt", xlTextWindows 'save as .txt
        wbNew.Close 'close the file
        ThisWorkbook.Activate ' go back to the main workbook
        row = row + 1 'increment the row
    Next

End Sub
