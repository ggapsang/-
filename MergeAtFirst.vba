Sub MergeTableSetOrder()

    Dim wb As workbook
    Dim ws_1 As worksheet
    Dim ws_2 As worksheet
    Dim ws_result As worksheet
    
    Set wb = ActiveWorkbook
    
    Set ws_1 = wb.Sheets(1)
    Set ws_2 = wb.Sheets(2)
    Set ws_result = wb.Sheets(3)
    
    
    '첫 번째 시트에 기본 변수 선언 : 마지막 행, 마지막 열, 키, 칼럼, 값을 배열로 저장
    
    Dim lastRowWs_1 As Long '마지막 행
    Dim lastColsWs_1 As Long '마지막 열
    Dim strLastColsWs_1 As String '마지막 열 문자
    
    Dim keyArray_1() As Variant '키 배열
    Dim colArray_1() As Variant '열 배열
    Dim valueArray_1() As Variant '값 배열
    
    lastRowWs_1 = ws_1.Range("A1").End(xlDown).Row
    lastColsWs_1 = ws_1.Range("A1").End(xlToRight).Column
    strLastColsWs_1 = ConvertToLetter(lastColsWs_1)
    
    
    keyArray_1() = ws_1.Range("A2:A" & lastRowWs_1).value
    colArray_1() = ws_1.Range("B1:" & strLastColsWs_1 & 1).value
    valueArray_1() = ws_1.Range("B2:" & strLastColsWs_1 & lastRowWs_1).value
    
    
    '두 번째 시트에 기본 변수 선언 : 첫 번째 시트와 동일한 방식으로 저장
    
    Dim lastRowWs_2 As Long '마지막 행
    Dim lastColsWs_2 As Long '마지막 열
    Dim strLastColsWs_2 As String '마지막 열 문자
    
    Dim keyArray_2() As Variant '키 배열
    Dim colArray_2() As Variant '열 배열
    Dim valueArray_2() As Variant '값 배열
    
    lastRowWs_2 = ws_2.Range("A1").End(xlDown).Row
    lastColsWs_2 = ws_2.Range("A1").End(xlToRight).Column
    strLastColsWs_2 = ConvertToLetter(lastColsWs_2)
    
    
    keyArray_2() = ws_2.Range("A2:A" & lastRowWs_2).value
    colArray_2() = ws_2.Range("B1:" & strLastColsWs_2 & 1).value
    valueArray_2() = ws_2.Range("B2:" & strLastColsWs_2 & lastRowWs_2).value
    
    
    Dim i_1 As Long
    Dim i_2 As Long
    Dim j_1 As Long
    Dim j_2 As Long
                            
    
    'key-col을 기준으로 매핑하면서 우선순위에 따라 배열의 값을 수정
    
    For k2 = LBound(keyArray_2, 1) To UBound(keyArray_2, 1)
        '일치하는 키 값 찾기
        For k1 = LBound(keyArray_1, 1) To UBound(keyArray_1, 1)
            ' 키 값이 일치하는 경우
            If keyArray_2(k2, 1) = keyArray_1(k1, 1) Then
                i_1 = k1 '값 배열의 행 번호 찾음
                i_2 = k2
                '모든 열에 대해서 검사해야 하므로 내부 For 루프를 벗어나지 않음
                '일치하는 칼럼 값 찾기
                For col_2 = LBound(colArray_2, 2) To UBound(colArray_2, 2)
                    For col_1 = LBound(colArray_1, 2) To UBound(colArray_1, 2)
                        If colArray_1(1, col_1) = colArray_2(1, col_2) Then
                            j_1 = col_1
                            j_2 = col_2
                            '해당 위치가 비어있으면 값을 복사
                            If IsEmpty(valueArray_1(i_1, j_1)) Or valueArray_1(i_1, j_1) = "" Then
                                valueArray_1(i_1, j_1) = valueArray_2(i_2, j_2)
                            End If
                        End If
                    Next col_1
                Next col_2
            End If
        Next k1
    Next k2

    
    With ws_result
        
        .Range("B2").Resize(UBound(valueArray_1, 1), UBound(valueArray_1, 2)).value = valueArray_1
    
        .Range("A2:A" & lastRowWs_1).value = keyArray_1
        
        .Range("B1:" & strLastColsWs_1 & 1).value = colArray_1
    
        .Range("A1").value = ws_1.Range("A1").value
        
    End With
    
End Sub

Function ConvertToLetter(iCol As Long) As String
   Dim a As Long
   Dim b As Long
   a = iCol
   ConvertToLetter = ""
   Do While iCol > 0
      a = Int((iCol - 1) / 26)
      b = (iCol - 1) Mod 26
      ConvertToLetter = Chr(b + 65) & ConvertToLetter
      iCol = a
   Loop
End Function
