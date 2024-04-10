Sub UpdateAndLogOverlappingValues()
    Dim dictKeyMap As Object
    Dim dictColMap As Object
    Set dictKeyMap = CreateObject("Scripting.Dictionary")
    Set dictColMap = CreateObject("Scripting.Dictionary")
    
    ' keyArray_1를 딕셔너리에 매핑
    For k1 = LBound(keyArray_1, 1) To UBound(keyArray_1, 1)
        dictKeyMap(keyArray_1(k1, 1)) = k1
    Next k1
    
    ' colArray_1의 칼럼 이름을 딕셔너리에 매핑
    For col_1 = LBound(colArray_1, 2) To UBound(colArray_1, 2)
        dictColMap(colArray_1(1, col_1)) = col_1
    Next col_1
    
    ' 겹치는 값을 기록할 시트 설정
    Dim logSheet As Worksheet
    Set logSheet = ThisWorkbook.Sheets("Sheet4") ' 네 번째 시트 이름을 정확히 지정하세요
    Dim logRow As Long
    logRow = 1 ' 로깅 시작 행
    
    ' keyArray_2를 순회하며 일치하는 키 찾기
    For k2 = LBound(keyArray_2, 1) To UBound(keyArray_2, 1)
        If dictKeyMap.Exists(keyArray_2(k2, 1)) Then
            Dim i_1 As Long
            i_1 = dictKeyMap(keyArray_2(k2, 1))
            
            For col_2 = LBound(colArray_2, 2) To UBound(colArray_2, 2)
                If dictColMap.Exists(colArray_2(1, col_2)) Then
                    Dim j_1 As Long
                    j_1 = dictColMap(colArray_2(1, col_2))
                    
                    ' 겹치는 값이 있는 경우 로깅
                    If Not IsEmpty(valueArray_1(i_1, j_1)) And valueArray_1(i_1, j_1) <> "" Then
                        If valueArray_1(i_1, j_1) <> valueArray_2(k2, col_2) Then
                            ' 칼럼 이름 로깅
                            logSheet.Cells(logRow, 1).Value = colArray_1(1, j_1)
                            ' 원래 값과 새 값 비교하여 로깅
                            logSheet.Cells(logRow, 2).Value = valueArray_1(i_1, j_1)
                            logSheet.Cells(logRow, 3).Value = valueArray_2(k2, col_2)
                            logRow = logRow + 1
                        End If
                        
                        ' 우선순위에 따라 값을 할당할 로직을 여기에 추가하세요
                    End If
                    
                    ' 해당 위치가 비어있으면 값을 복사
                    If IsEmpty(valueArray_1(i_1, j_1)) Or valueArray_1(i_1, j_1) = "" Then
                        valueArray_1(i_1, j_1) = valueArray_2(k2, col_2)
                    End If
                End If
            Next col_2
        End If
    Next k2
End Sub
