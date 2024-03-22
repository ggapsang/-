Sub ProcessFileInLoop()
    Dim myPath As String
    Dim myFile As String
    Dim masterFilePath As String
    Dim myWorkbook As Workbook
    Dim masterWorkbook As Workbook
    Dim fd As FileDialog

    ' 폴더 선택 대화상자 초기화
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "폴더 선택"
        If .Show = -1 Then
            myPath = .SelectedItems(1) ' 폴더 선택
        Else
            MsgBox "선택 취소"
            Exit Sub
        End If
    End With

    ' 파일 선택 대화상자 초기화
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "마스터 파일 선택"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls"
        If .Show = -1 Then
            masterFilePath = .SelectedItems(1) ' 마스터 파일 선택
        Else
            MsgBox "선택 취소"
            Exit Sub
        End If
    End With

    ' 마스터 워크북 열기
    Set masterWorkbook = Workbooks.Open(masterFilePath)

    ' 폴더 경로 검증
    If Right(myPath, 1) <> "\" Then myPath = myPath & "\"

    ' 파일 확장자를 포함하여 모든 파일을 순회
    myFile = Dir(myPath & "*.xls*") ' .xlsx와 .xls 모두 포함

    Do While myFile <> ""
        Set myWorkbook = Workbooks.Open(Filename:=myPath & myFile)
        
        ' FindValuesInWorkbook에 마스터 워크북과 현재 워크북 전달
        FindValuesAndMove myWorkbook, masterWorkbook

        myWorkbook.Close False ' 변경 사항 없으므로 저장하지 않고 닫음
        myFile = Dir ' 다음 파일로 이동
    Loop

    ' 마스터 워크북 저장 및 닫기
    masterWorkbook.Save
    'masterWorkbook.Close

    MsgBox "끝"
End Sub


Public Sub FindValuesAndMove(sourceWorkbook As Workbook, masterWorkbook As Workbook)
    Dim masterSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim cellAddress As String
    Dim sourceAddress() As String
    Dim i As Long, j As Long
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim fileRow As Long
    Dim beginColumn As Long
    Dim workbookName As String


    ' 마스터 파일의 첫 번째 시트 설정
    Set masterSheet = masterWorkbook.Sheets(1)
    
    ' 마스터 파일에서 데이터를 찾을 마지막 행 결정
    lastRow = masterSheet.Cells(masterSheet.Rows.Count, "B").End(xlUp).Row
    
    ' 마스터 파일의 마지막 열 결정
    lastColumn = masterSheet.Cells(3, Columns.Count).End(xlToLeft).Column

    ' 소스 파일의 이름 결정
    workbookName = sourceWorkbook.Name
    
    ' 마스터 파일의 네 번째 행부터 마지막 행까지 순회하면서 소스 파일(데이터시트)이 있는 워크북의 행 번호 찾기
    fileRow = 0
    For i = 4 To lastRow
        Dim fileNameCell As String
        fileNameCell = masterSheet.Range("B" & i).Value
        
        
        If fileNameCell = workbookName Then
            fileRow = i
            Exit For ' 파일을 찾으면 반복 중지
        End If
    Next i

    If fileRow = 0 Then
        Debug.Print sourceWorkbook.Name
        Exit Sub ' 파일을 찾지 못했으므로 서브 프로시저 종료
    End If


    ' 데이터가 들어가는 열(beginColumn)부터 시작하여 마지막 열까지 순회
    beginColumn = 5
        
    For j = beginColumn To lastColumn
        cellAddress = masterSheet.Cells(3, j).Value
        sourceAddress = Split(cellAddress, "!")
            
        If UBound(sourceAddress) > 0 Then
        
            Set sourceSheet = sourceWorkbook.Worksheets(sourceAddress(0))
            ' 찾은 데이터를 마스터 파일의 적절한 위치에 복사
            masterSheet.Cells(fileRow, j).Value = sourceSheet.Range(sourceAddress(1)).Value
        End If
    Next j
    
End Sub

