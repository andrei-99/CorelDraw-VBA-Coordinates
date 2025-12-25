' VBA Macro for CorelDRAW 22+
' Collects center coordinates of objects on current layer in mm,
' rounds them, sorts snake-like, normalizes, and saves to text file.

Option Explicit

Sub SnakeSortExport()
    Dim doc As Document
    Dim lyr As Layer
    Dim s As Shape
    Dim arr() As Variant
    Dim i As Long, j As Long, count As Long

    ' Ensure measurement units are millimeters
    ActiveDocument.Unit = cdrMillimeter

    Set doc = ActiveDocument
    Set lyr = doc.ActiveLayer

    ' Count shapes
    count = 0
    For Each s In lyr.Shapes
        count = count + 1
    Next s

    If count = 0 Then
        MsgBox "Нет объектов на текущем слое."
        Exit Sub
    End If

    ReDim arr(1 To count, 1 To 2)

    ' Fill array with rounded centers
    i = 1
    For Each s In lyr.Shapes
        Dim cx As Double, cy As Double
        cx = RoundAway(s.CenterX)
        cy = RoundAway(s.CenterY)
        arr(i, 1) = cx
        arr(i, 2) = cy
        i = i + 1
    Next s

    ' Sort by Y ascending, then snake-sort rows by X
    Call SortByYThenSnake(arr, count)

    ' Normalize relative to the first (lowest-left) object
    Dim baseX As Double, baseY As Double
    baseX = arr(1, 1)
    baseY = arr(1, 2)

    For i = 1 To count
        arr(i, 1) = arr(i, 1) - baseX
        arr(i, 2) = arr(i, 2) - baseY
    Next i

    ' Save to text file next to the CDR file
    Dim filePath As String
    filePath = Left(doc.FullFileName, InStrRev(doc.FullFileName, ".") - 1) & ".txt"

    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f

    For i = 1 To count
        Print #f, CStr(arr(i, 1)) & ", " & CStr(arr(i, 2))
    Next i

    Close #f

    MsgBox "Файл сохранён: " & filePath
End Sub

' Helper: sort by Y ascending, then snake along X
Sub SortByYThenSnake(ByRef a() As Variant, ByVal n As Long)
    Dim i As Long, j As Long, t1 As Variant, t2 As Variant

    ' Primary sort by Y ascending, then X ascending
    For i = 1 To n - 1
        For j = i + 1 To n
            If a(j, 2) < a(i, 2) Or (a(j, 2) = a(i, 2) And a(j, 1) < a(i, 1)) Then
                t1 = a(i, 1): t2 = a(i, 2)
                a(i, 1) = a(j, 1): a(i, 2) = a(j, 2)
                a(j, 1) = t1: a(j, 2) = t2
            End If
        Next j
    Next i

    ' Apply snake pattern across rows (same rounded Y)
    Dim startRow As Long, endRow As Long, rowIndex As Long
    startRow = 1
    rowIndex = 0

    Do While startRow <= n
        endRow = startRow
        ' expand the row safely (no short-circuit issue)
        Do While endRow < n
            If a(endRow + 1, 2) <> a(startRow, 2) Then Exit Do
            endRow = endRow + 1
        Loop

        ' Reverse X order for every second row group
        If (rowIndex Mod 2) = 1 Then
            Call ReverseSegment(a, startRow, endRow)
        End If

        rowIndex = rowIndex + 1
        startRow = endRow + 1
    Loop
End Sub

Sub ReverseSegment(ByRef a() As Variant, ByVal s As Long, ByVal e As Long)
    Dim i As Long, j As Long
    Dim tx As Variant, ty As Variant
    i = s: j = e
    Do While i < j
        tx = a(i, 1): ty = a(i, 2)
        a(i, 1) = a(j, 1): a(i, 2) = a(j, 2)
        a(j, 1) = tx: a(j, 2) = ty
        i = i + 1: j = j - 1
    Loop
End Sub

' Round halves away from zero to the nearest integer (мм)
Public Function RoundAway(ByVal x As Double) As Long
    If x >= 0 Then
        RoundAway = CLng(Int(x + 0.5))
    Else
        RoundAway = -CLng(Int(-x + 0.5))
    End If
End Function
