Dim Bombs As Collection
Dim Running As Boolean
Dim Flags As Integer
Dim Size As Range
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
        Cancel = True
     If Selection.Count = 1 And Running Then
      
        If Not Intersect(Target, Size) Is Nothing And Not Target.Cells.Interior.ColorIndex = 5 Then
            Call checkCell(Target)
            
        End If
        
    End If
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
     If Selection.Count = 1 And Running Then
      
        If Not Intersect(Target, Size) Is Nothing Then
            If Target.Cells.Interior.ColorIndex = 15 And Flags > 0 Then
                Target.Cells.Interior.ColorIndex = 5
                Flags = Flags - 1
                Cells(2, 1).Value = Flags
            ElseIf Target.Cells.Interior.ColorIndex = 5 Then
                Target.Cells.Interior.ColorIndex = 15
                Flags = Flags + 1
                Cells(2, 1).Value = Flags
            End If
            
        End If
        
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Selection.Count = 1 Then
        If Not Intersect(Target, Range("A1")) Is Nothing Then
            Call ClearBoard
            Call GenerateRandom
            Call setCellNumbers(Size)
            Call Test
            Flags = Cells(4, 1).Value
            Cells(2, 1).Value = Flags
            Call Test
            Running = True
        ElseIf Not Intersect(Target, Range("A3")) Is Nothing Then
            If Cells(3, 1).Value = "9 x 9" Then
                Cells(3, 1).Value = "16 x 16"
                Cells(4, 1).Value = "40"
            ElseIf Cells(3, 1).Value = "16 x 16" Then
                Cells(3, 1).Value = "30 x 16"
                Cells(4, 1).Value = "99"
            ElseIf Cells(3, 1).Value = "30 x 16" Then
                Cells(3, 1).Value = "9 x 9"
                Cells(4, 1).Value = "10"
            End If
        End If
    End If
End Sub

Private Sub ClearBoard()
    Dim x As Integer
    Dim y As Integer
    y = CInt(Left(Cells(3, 1).Value, InStr(Cells(3, 1).Value, "x") - 1))
    x = CInt(Right(Cells(3, 1).Value, InStr(Cells(3, 1).Value, "x") - 1))
    Set Size = Range(Cells(2, 2), Cells(x + 1, y + 1))

    Worksheets("Sheet1").Protect "Password", UserInterfaceOnly:=True
    Worksheets("Sheet1").Range("B2:AH34").ClearContents
    Worksheets("Sheet1").Range("A1:AH34").Interior.ColorIndex = 2
    Worksheets("Sheet1").Range("A1:AH34").Borders.LineStyle = xlNone
    Worksheets("Sheet1").Range("B1:AH34").Cells.RowHeight = 25
    Worksheets("Sheet1").Range("B1:AH34").Cells.ColumnWidth = 5
    Worksheets("Sheet1").EnableOutlining = True
    Worksheets("Sheet1").Cells.Font.Name = "Arial Black"
    Worksheets("Sheet1").Range("A1:A10").Font.Name = "Calibri"
    Worksheets("Sheet1").Cells.HorizontalAlignment = xlCenter
    Worksheets("Sheet1").Cells.VerticalAlignment = xlCenter
    Size.Borders.LineStyle = xlContinuous
    Size.Interior.ColorIndex = 15
    Size.NumberFormat = ";;;"
    Size.FormulaHidden = True
    

End Sub

Private Sub Test()
    Debug.Print Bombs.Count & " - " & Cells(4, 1).Value
End Sub

Private Sub GenerateRandom()
    Dim c As Range
    Dim x As Integer
    Dim y As Integer
    Dim xChar As String
    Dim yChar As String
    Dim i As Integer
    Dim r As Range
    Dim Num As Integer
    
    Set Bombs = New Collection
    Num = Cells(4, 1).Value
    For i = 1 To Num
        x = (Size.Rows.Count - 1) * Rnd() + 2
        y = (Size.Columns.Count - 1) * Rnd() + 2
        If Cells(x, y).Value = "" Then
            Set r = Worksheets("Sheet1").Range(Cells(x, y).Address)
            r.Cells.Value = "B"
            Bombs.Add r
        Else
            Num = Num + 1
        End If
    Next
End Sub

Private Function containsBomb(r As Range) As Boolean
    Dim i As Integer
    Dim b As Range
    For Each b In Bombs
        If b.Cells.Address = r.Cells.Address Then
            containsBomb = True
            Exit Function
        End If
    Next
containsBomb = False
End Function

Private Function getVisibilityRange(r As Range) As Range
    Dim c0 As Range
    Dim c1 As Range
    Debug.Print (r.Cells.Row - 1) & ":" & (r.Cells.Column - 1) & " : " & (r.Cells.Row + 1) & ":" & (r.Cells.Column + 1)
    
    Set getVisibilityRange = Range(Cells(r.Cells.Row - 1, r.Cells.Column - 1), Cells(r.Cells.Row + 1, r.Cells.Column + 1))
    
End Function

Private Function setCellNumbers(r As Range)
    Dim c As Range
    Dim c1 As Range
    Dim visRange As Range
    Dim n As Integer
    
    For Each c In r
        n = isNextTo(c)
        If n = -1 Then
            c.Cells.Value = "B"
            c.Font.ColorIndex = 1
        ElseIf n = 0 Then
            c.Cells.Value = ""
        Else
            c.Cells.Value = n
            If n = 1 Then
                c.Font.ColorIndex = 5
            ElseIf n = 2 Then
                c.Font.ColorIndex = 10
            ElseIf n = 3 Then
                c.Font.ColorIndex = 6
            ElseIf n = 4 Then
                c.Font.ColorIndex = 46
            ElseIf n = 5 Then
                c.Font.ColorIndex = 53
            Else
                c.Font.ColorIndex = 3
            End If
        End If
    Next c
End Function

Private Function isNextTo(r As Range) As Integer
    Dim c As Range
    Dim n As Integer
    n = 0
    For Each c In Bombs
        If r.Cells.Column = c.Cells.Column And r.Cells.Row = c.Cells.Row Then
            isNextTo = -1
            Exit Function
        
        ElseIf r.Cells.Column - 1 = c.Cells.Column And r.Cells.Row - 1 = c.Cells.Row Then
            n = n + 1
            'Debug.Print CStr(r.Cells.Row) + ":" + CStr(r.Cells.Column) + " - " + CStr(c.Cells.Row) + ":" + CStr(c.Cells.Column)
        ElseIf r.Cells.Column + 1 = c.Cells.Column And r.Cells.Row + 1 = c.Cells.Row Then
            n = n + 1
            'Debug.Print CStr(r.Cells.Row) + ":" + CStr(r.Cells.Column) + " - " + CStr(c.Cells.Row) + ":" + CStr(c.Cells.Column)
        ElseIf r.Cells.Column + 1 = c.Cells.Column And r.Cells.Row - 1 = c.Cells.Row Then
            n = n + 1
            'Debug.Print CStr(r.Cells.Row) + ":" + CStr(r.Cells.Column) + " - " + CStr(c.Cells.Row) + ":" + CStr(c.Cells.Column)
        ElseIf r.Cells.Column - 1 = c.Cells.Column And r.Cells.Row = c.Cells.Row Then
            n = n + 1
            'Debug.Print CStr(r.Cells.Row) + ":" + CStr(r.Cells.Column) + " - " + CStr(c.Cells.Row) + ":" + CStr(c.Cells.Column)
        ElseIf r.Cells.Column - 1 = c.Cells.Column And r.Cells.Row + 1 = c.Cells.Row Then
            n = n + 1
            'Debug.Print CStr(r.Cells.Row) + ":" + CStr(r.Cells.Column) + " - " + CStr(c.Cells.Row) + ":" + CStr(c.Cells.Column)
        ElseIf r.Cells.Column = c.Cells.Column And r.Cells.Row - 1 = c.Cells.Row Then
            n = n + 1
            'Debug.Print CStr(r.Cells.Row) + ":" + CStr(r.Cells.Column) + " - " + CStr(c.Cells.Row) + ":" + CStr(c.Cells.Column)
        ElseIf r.Cells.Column = c.Cells.Column And r.Cells.Row + 1 = c.Cells.Row Then
            n = n + 1
            'Debug.Print CStr(r.Cells.Row) + ":" + CStr(r.Cells.Column) + " - " + CStr(c.Cells.Row) + ":" + CStr(c.Cells.Column)
        ElseIf r.Cells.Column + 1 = c.Cells.Column And r.Cells.Row = c.Cells.Row Then
            n = n + 1
            'Debug.Print CStr(r.Cells.Row) + ":" + CStr(r.Cells.Column) + " - " + CStr(c.Cells.Row) + ":" + CStr(c.Cells.Column)
        End If
    Next c
isNextTo = n
End Function


Private Function checkCell(Target As Range)
    Dim cRange As Collection
    Dim tempRange As Range
    Set cRange = New Collection
    If containsBomb(Target) Then
        Target.Interior.ColorIndex = 3
        Size.NumberFormat = "General"
        Running = False
    Else
        Target.NumberFormat = "General"
        Target.Interior.ColorIndex = 16
        If Target.Cells.Value = "" Then
            Call openAllEmpty
        End If
       
    End If
End Function



Private Function isOpen(r As Range) As Boolean
    If r.Cells.Interior.ColorIndex = 15 Then
        isOpen = False
        Exit Function
    End If
isOpen = True
End Function

Private Function showRange(Target As Range) As Boolean
    Dim tCell As Range
    Dim openedNew As Boolean
    openedNew = False
    
    For Each tCell In Target.Cells
        If Not Application.Intersect(tCell, Size) Is Nothing And Not isOpen(tCell) Then
            tCell.NumberFormat = "General"
            tCell.Interior.ColorIndex = 16
            openedNew = True
        End If
        
    Next tCell
showRange = openedNew
End Function

Private Function openAllEmpty()
    Dim opened As Boolean
    Dim r As Range
    Dim vRange As Range
    
    opened = True
    
    Do While opened
        opened = False
        For Each r In Size
            If isOpen(r) And r.Cells.Value = "" Then
                Set vRange = getVisibilityRange(r)
                If showRange(vRange) Then
                    opened = True
                End If
            End If
        Next r
    Loop
    
    
End Function
