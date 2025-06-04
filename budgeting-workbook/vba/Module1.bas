Attribute VB_Name = "Module1"
Sub BuildLayout()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BUILDING 1 New")

    Dim rawInput As String
    rawInput = Trim(ws.Range("B1").Text)

    If rawInput = "" Then
        MsgBox "Please enter a number between 1 and 50 in B1.", vbExclamation
        Exit Sub
    End If

    rawInput = Replace(rawInput, ",", "")
    If Not IsNumeric(rawInput) Then
        MsgBox "Entry in B1 must be a number between 1 and 50.", vbExclamation
        Exit Sub
    End If

    Dim numBuildings As Long
    numBuildings = CLng(rawInput)
    If numBuildings < 1 Or numBuildings > 50 Then
        MsgBox "Please enter a valid number of buildings (1–50) in B1.", vbExclamation
        Exit Sub
    End If

    ws.Range("D1:ZZ100").ClearContents
    ws.Range("D1:ZZ100").ClearFormats

    ' Header label
    ws.Range("A1").Value = "# of Buildings"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Interior.Color = RGB(255, 230, 153)
    ws.Range("B1").Value = numBuildings

    Dim i As Long, colOffset As Long
    Dim startCol As Long: startCol = 4 ' Column D
    Dim startRow As Long: startRow = 5

    For i = 1 To numBuildings
        colOffset = (i - 1) * 3
        Dim baseCol As Long: baseCol = startCol + colOffset

        ' # of Levels Label + Dropdown
        ws.Range(ws.Cells(1, baseCol), ws.Cells(1, baseCol + 1)).Merge
        ws.Cells(1, baseCol).Value = "# of Levels"
        ws.Cells(1, baseCol).HorizontalAlignment = xlCenter
        ws.Cells(1, baseCol).Font.Bold = True
        ws.Cells(1, baseCol).Interior.Color = RGB(255, 230, 153)

        With ws.Cells(1, baseCol + 2).Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=Join(Application.Transpose(Evaluate("ROW(1:30)")), ",")
            .IgnoreBlank = True
            .InCellDropdown = True
        End With

        ' Building Header
        ws.Range(ws.Cells(2, baseCol), ws.Cells(2, baseCol + 2)).Merge
        ws.Cells(2, baseCol).Value = "Building " & i
        ws.Cells(2, baseCol).Font.Bold = True
        ws.Cells(2, baseCol).HorizontalAlignment = xlCenter
        ws.Cells(2, baseCol).Interior.Color = RGB(255, 230, 153)

        ' Totals Row (now row 3)
        ws.Cells(3, baseCol).Value = "Total:"
        ws.Cells(3, baseCol + 1).Formula = "=SUM(" & ws.Cells(5, baseCol + 1).Address & ":" & ws.Cells(34, baseCol + 1).Address & ")"
        ws.Cells(3, baseCol + 2).Formula = "=SUM(" & ws.Cells(5, baseCol + 2).Address & ":" & ws.Cells(34, baseCol + 2).Address & ")"
        ws.Range(ws.Cells(3, baseCol), ws.Cells(3, baseCol + 2)).Font.Bold = True

        ' Table Headers (Row 4)
        ws.Cells(4, baseCol).Value = "Input"
        ws.Cells(4, baseCol + 1).Value = "Perimeter"
        ws.Cells(4, baseCol + 2).Value = "Area"
        ws.Range(ws.Cells(4, baseCol), ws.Cells(4, baseCol + 2)).Font.Bold = True
        ws.Range(ws.Cells(4, baseCol), ws.Cells(4, baseCol + 2)).Interior.Color = RGB(255, 230, 153)
    Next i

    ws.Columns("D:ZZ").AutoFit
End Sub

