Attribute VB_Name = "Module11"
Sub init()
    Call DeleteTopAndBottomData
    Call SortBoxNo
    Call FourColumnFormat
    Call MergeBoxNo
    Call addHeaderRow
    Call AddBorder
End Sub


Sub DeleteTopAndBottomData()
Attribute DeleteTopAndBottomData.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+Shift+P
'
' Deleting Extra Data on top
    Range("A1:F8").Select
    Range("F8").Activate
    Selection.EntireRow.Delete
    
' Deleting ECIL AND YEAR Columns
    Range("D:D,E:E").Select
    Range("E1").Activate
    Selection.Delete Shift:=xlToLeft
    
' Deleting Extra Data at Bottom
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    thirdLastRow = lastRow - 2
    Range("A" & thirdLastRow & ":C" & lastRow).Select
    Range("C" & lastRow).Activate
    Selection.EntireRow.Delete
    
End Sub
Sub SortBoxNo()
'
' SortBoxNo Macro
'
' Keyboard Shortcut: Ctrl+Shift+O
'
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    Columns("C:C").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("C1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A1:C" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub FourColumnFormat()
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

    remainder = lastRow Mod 60
    loopCounter = lastRow + remainder
    rangeCounter = 1   'For range of data to be selected
    
    columnNumber = 1  'To change number in A1 then A61 then A
    
   Do While rangeCounter < loopCounter
        
        rangeStart = rangeCounter
        rangeEnd = rangeCounter + 60 - 1
        Range("A" & rangeStart & ":C" & rangeEnd).Select
        Selection.Cut
        Range("A" & columnNumber).Select
        ActiveSheet.Paste
        
        rangeStart = rangeEnd + 1
        rangeEnd = rangeStart + 60 - 1
        Range("A" & rangeStart & ":C" & rangeEnd).Select
        Selection.Cut
        Range("D" & columnNumber).Select
        ActiveSheet.Paste
        
        rangeStart = rangeEnd + 1
        rangeEnd = rangeStart + 60 - 1
        Range("A" & rangeStart & ":C" & rangeEnd).Select
        Selection.Cut
        Range("G" & columnNumber).Select
        ActiveSheet.Paste
        
        rangeStart = rangeEnd + 1
        rangeEnd = rangeStart + 60 - 1
        Range("A" & rangeStart & ":C" & rangeEnd).Select
        Selection.Cut
        Range("J" & columnNumber).Select
        ActiveSheet.Paste
        
        rangeCounter = rangeCounter + 240
        columnNumber = columnNumber + 60
         
    Loop
            
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
End Sub

Sub MergeBoxNo()
'
' MergeBoxNo Macro
'
' Keyboard Shortcut: Ctrl+Shift+I
'
    Range("C1:C10").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C1:C10").Select
    Selection.Copy
       
    Range("F1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("I1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("L1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    rangeCounter = 11
    
    Do While rangeCounter < lastRow
                Range("C" & rangeCounter).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
                Range("F" & rangeCounter).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Range("I" & rangeCounter).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Range("L" & rangeCounter).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
                
                rangeCounter = rangeCounter + 10
    Loop
End Sub

Sub AddBorder()
'
' AddBorder Macro
'

'
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    Range("A1:L" & lastRow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Sub addHeaderRow()
'
' addHeaderRow Macro
'
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    rowNoToInsert = lastRow - 60 + 1
    
    Do While rowNoToInsert > 0
            
            Range("A" & rowNoToInsert).Select
            Application.CutCopyMode = False
            Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
            rowNoToInsert = rowNoToInsert - 60
    Loop
End Sub
