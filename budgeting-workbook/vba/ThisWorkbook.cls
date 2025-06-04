VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub UpdateLevels()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BUILDING 1 New")

    Dim numBuildings As Long
    numBuildings = Val(ws.Range("B1").Value)
    If numBuildings < 1 Then Exit Sub

    Dim i As Long, maxLevels As Long: maxLevels = 30
    Dim startCol As Long: startCol = 4
    Dim startRow As Long: startRow = 4 ' Row 4 has headers, Level 1 starts at Row 5

    For i = 1 To numBuildings
        Dim colOffset As Long: colOffset = (i - 1) * 3
        Dim baseCol As Long: baseCol = startCol + colOffset
        Dim lvlCount As Long: lvlCount = Val(ws.Cells(1, baseCol + 2).Value)

        If lvlCount < 0 Then lvlCount = 0
        If lvlCount > maxLevels Then lvlCount = maxLevels

        ' Clear previous values and formatting from Level 1 to max row
        ws.Range(ws.Cells(startRow + 1, baseCol), ws.Cells(startRow + maxLevels, baseCol + 2)).ClearContents
        ws.Range(ws.Cells(startRow + 1, baseCol), ws.Cells(startRow + maxLevels, baseCol + 2)).Interior.ColorIndex = xlNone

        ' Populate Level labels and color fill for data entry cells
        Dim j As Long
        For j = 1 To lvlCount
            Dim r As Long: r = startRow + j
            ws.Cells(r, baseCol).Value = "Level " & j
            ws.Cells(r, baseCol).Font.Bold = True
            ws.Range(ws.Cells(r, baseCol + 1), ws.Cells(r, baseCol + 2)).Interior.Color = RGB(255, 242, 204)
        Next j

        ' Apply borders around the level section
        If lvlCount > 0 Then
            With ws.Range(ws.Cells(startRow + 1, baseCol), ws.Cells(startRow + lvlCount, baseCol + 2)).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End If
    Next i
End Sub

