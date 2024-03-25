Sub ListDefinedNamesAndValues()
    Dim ws As Worksheet
    Dim name As Name
    Dim i As Long
    
    ' 새 시트를 추가하고 변수에 할당
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "DefinedNamesList" ' 새 시트의 이름 설정, 이미 존재하는 이름인 경우 에러가 발생할 수 있음

    ' 헤더 설정
    ws.Cells(1, 1).Value = "Name"
    ws.Cells(1, 2).Value = "Value/Reference"
    
    ' 모든 이름(Defined Names)을 순회하며 시트에 기록
    i = 2 ' 시작 행 번호
    For Each name In ThisWorkbook.Names
        ws.Cells(i, 1).Value = name.Name ' 이름 관리자의 이름
        ws.Cells(i, 2).Value = "'" & name.RefersTo ' 이름이 참조하는 값 또는 수식 (앞에 '를 붙여 문자열로 처리)
        i = i + 1
    Next name
End Sub
