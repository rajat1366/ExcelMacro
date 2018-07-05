Attribute VB_Name = "Module1"
Sub DeleteTopAndBottomData()
Attribute DeleteTopAndBottomData.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+Shift+P
'
    Range("A1:F8").Select
    Range("F8").Activate
    Selection.EntireRow.Delete
    
    
    Range("D:D,E:E").Select
    Range("E1").Activate
    Selection.Delete Shift:=xlToLeft
    
    
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    thirdLastRow = lastRow - 2
    Range("A" & thirdLastRow & ":C" & lastRow).Select
    Range("C" & lastRow).Activate
    Selection.EntireRow.Delete
    
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
    
End Sub

