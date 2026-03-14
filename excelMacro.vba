' #############################################################################
' ### https://comate.baidu.com/zh/page/j81lk42dul8
Sub AddSerialNumbers()
    ' *** https://www.google.com/search?q=excel+VBA+save+and+restore+address
    ' --- 1. Declare variables to store the original location ---
    Dim rngOrigSelection As Range
    Dim sOrigWSName As String
    Dim sOrigWBName As String

    ' --- 2. Save the current location ---
    Set rngOrigSelection = Selection ' Stores the actual range object, including multiple selected cells
    sOrigWSName = ActiveSheet.Name   ' Stores the name of the original worksheet
    sOrigWBName = ActiveWorkbook.Name ' Stores the name of the original workbook

    Dim i As Integer
    Dim maxNumber As Integer

    Dim ws As Worksheet
    Set ws = ActiveSheet
    maxNumber = ws.UsedRange.Rows.Count - 1

    Range("A2").Select
    If maxNumber <= 0 Then
        Exit Sub
    End If

    For i = 1 To maxNumber
        ActiveCell.Value = i
        ActiveCell.Offset(1, 0).Activate ' move to next cell
    Next i

    ' --- 4. Restore the original location ---
    ' Ensure the correct workbook and worksheet are active before selecting the range
    Workbooks(sOrigWBName).Activate
    Worksheets(sOrigWSName).Activate
    rngOrigSelection.Select
End Sub

' #############################################################################
' ### https://www.google.com/search?q=excel+%E5%AE%8F+%E9%9A%94%E8%A1%8C+%E4%B8%8A%E8%89%B2
Sub HighlightAlternateRows()
    Dim i As Long
    Dim lastRow As Long
    
    ' Get the last row of the used range in the current worksheet
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Start looping from the second row, assuming the first row is the header
    For i = 2 To lastRow Step 2
        ' Set the fill colour to light grey (can be changed to vbYellow, vbRed, etc.)
        Rows(i).Interior.Color = RGB(220, 220, 220)
    Next i
End Sub

' #############################################################################
' ### https://www.google.com/search?q=excel+macro+make+hyperlink+with+content+in+selected+cells
Sub MakeSelectedHyperlinksHot()
    ' Macro to convert text URLs in selected cells to clickable hyperlinks
    
    Dim cell As Range
    
    ' Loop through each cell in the current selection
    For Each cell In Selection
        ' Check if the cell is not empty
        If cell.Value <> "" Then
            ' Add a hyperlink to the cell
            ' Anchor:=cell (the cell itself)
            ' Address:=cell.Value (the URL or file path from the cell content)
            ' TextToDisplay:=cell.Value (the visible text in the cell)
            ActiveSheet.Hyperlinks.Add Anchor:=cell, Address:=cell.Value, TextToDisplay:=cell.Value
        End If
    Next cell
End Sub

' ################################ END ########################################
