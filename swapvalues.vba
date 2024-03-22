Sub SwapRangesDynamically()
    Dim range1 As Range
    Dim range2 As Range
    Dim tempValue As Variant
    Dim i As Integer
    Dim inputRange As String
    
    ' 사용자가 선택한 첫 번째 범위
    Set range1 = Selection
    
    ' 두 번째 범위 입력 요청
    inputRange = Application.InputBox("두 번째 범위 입력:", Type:=8)
    
    ' 입력받은 두 번째 범위 설정
    If TypeName(inputRange) = "Range" Then
        Set range2 = inputRange
    Else
        MsgBox "올바른 범위가 아님"
        Exit Sub
    End If
    
    ' 두 범위의 크기 비교
    If range1.Rows.Count <> range2.Rows.Count Or range1.Columns.Count <> range2.Columns.Count Then
        MsgBox "범위의 크기가 일치하지 않음."
        Exit Sub
    End If
    
    ' 값 교환
    For i = 1 To range1.Rows.Count
        tempValue = range1.Cells(i, 1).Value
        range1.Cells(i, 1).Value = range2.Cells(i, 1).Value
        range2.Cells(i, 1).Value = tempValue
    Next i

End Sub
