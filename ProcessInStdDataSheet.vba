''' Module Name : DataFinder '''
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
        fileNameCell = masterSheet.Range("B" & i).value
        
        
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
        cellAddress = masterSheet.Cells(3, j).value
        sourceAddress = Split(cellAddress, "!")
            
        If UBound(sourceAddress) > 0 Then
        
            Set sourceSheet = sourceWorkbook.Worksheets(sourceAddress(0))
            ' 찾은 데이터를 마스터 파일의 적절한 위치에 복사
            masterSheet.Cells(fileRow, j).value = sourceSheet.Range(sourceAddress(1)).value
        End If
    Next j
    
End Sub

''' Module Name : MainProcess '''
Sub RunMasterFolderProcessing()

    Dim sourceFolderPath As String
    Dim masterFilesFolderPath As String
    Dim masterFileName As String
    Dim masterFilePath As String
    Dim sourceSubFolderPath As String
    Dim fd As FileDialog
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 소스 폴더 선택
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "소스 폴더 선택"
        If .Show = -1 Then
            sourceFolderPath = .SelectedItems(1)
        Else
            MsgBox "선택 취소"
            Exit Sub
        End If
    End With
    
    ' 마스터 파일 폴더 선택
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "마스터 파일 폴더 선택"
        If .Show = -1 Then
            masterFilesFolderPath = .SelectedItems(1)
        Else
            MsgBox "선택 취소"
            Exit Sub
        End If
    End With
    
    ' 마스터 파일 폴더 내의 각 파일 처리
    Dim masterFile As String
    masterFile = Dir(masterFilesFolderPath & "\*.xls*")
    Do While masterFile <> ""
        masterFilePath = masterFilesFolderPath & "\" & masterFile
        masterFileName = Left(masterFile, InStrRev(masterFile, ".") - 1)
        sourceSubFolderPath = sourceFolderPath & "\" & masterFileName
        
        ' 소스 서브폴더에서 파일 처리
        If Len(Dir(sourceSubFolderPath, vbDirectory)) <> 0 Then
            ProcessMasterFileAndSubfolders sourceSubFolderPath, masterFilePath
        End If
        
        masterFile = Dir() ' 다음 마스터 파일로
    Loop
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "모든 처리가 완료되었습니다."
End Sub

''' Module Name : MasterFileProcessor'''
Public Sub ProcessMasterFileAndSubfolders(sourceSubFolderPath As String, masterFilePath As String)
    
    Dim myFile As String
    Dim myWorkbook As Workbook
    Dim masterWorkbook As Workbook
    Dim fileCount As Integer
    Dim processedCount As Integer

    Application.ScreenUpdating = False ' 화면 업데이트 끄기
    Application.DisplayAlerts = False ' 경고 표시 안함

    
    ' 마스터 워크북 열기
    Set masterWorkbook = Workbooks.Open(masterFilePath)

    ' 파일 경로 검증
    If Right(sourceSubFolderPath, 1) <> "\" Then sourceSubFolderPath = sourceSubFolderPath & "\"

    ' 파일 카운트 초기화
    myFile = Dir(sourceSubFolderPath & "*.xls*")
    While myFile <> ""
        fileCount = fileCount + 1
        myFile = Dir
    Wend

    myFile = Dir(sourceSubFolderPath & "*.xls*")
    Do While myFile <> ""
        processedCount = processedCount + 1
        Application.StatusBar = "Processing file " & processedCount & " of " & fileCount ' 진행 상태 업데이트
        
        Set myWorkbook = Workbooks.Open(Filename:=sourceSubFolderPath & myFile, ReadOnly:=True) ' 읽기 전용으로 열기
        
        FindValuesAndMove myWorkbook, masterWorkbook

        myWorkbook.Close False ' 변경 사항 없으므로 저장하지 않고 닫음
        myFile = Dir ' 다음 파일로 이동
    Loop

    masterWorkbook.Save
    'masterWorkbook.Close
    MsgBox "완료"

    Application.StatusBar = False ' 상태 바 초기화
    Application.ScreenUpdating = True ' 화면 업데이트
    Application.DisplayAlerts = True
End Sub

''' Module Name : WorkbookProcessor '''
Public Sub ProcessWorkbookFiles()
    
    Dim myPath As String
    Dim myFile As String
    Dim masterFilePath As String
    Dim myWorkbook As Workbook
    Dim masterWorkbook As Workbook
    Dim fd As FileDialog
    Dim fileCount As Integer
    Dim processedCount As Integer

    Application.ScreenUpdating = False ' 화면 업데이트 끄기
    Application.DisplayAlerts = False ' 경고 표시 안함

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

    ' 파일 경로 검증
    If Right(myPath, 1) <> "\" Then myPath = myPath & "\"

    ' 파일 카운트 초기화
    myFile = Dir(myPath & "*.xls*")
    While myFile <> ""
        fileCount = fileCount + 1
        myFile = Dir
    Wend

    myFile = Dir(myPath & "*.xls*")
    Do While myFile <> ""
        processedCount = processedCount + 1
        Application.StatusBar = "Processing file " & processedCount & " of " & fileCount ' 진행 상태 업데이트
        
        Set myWorkbook = Workbooks.Open(Filename:=myPath & myFile, ReadOnly:=True) ' 읽기 전용으로 열기
        
        FindValuesAndMove myWorkbook, masterWorkbook

        myWorkbook.Close False ' 변경 사항 없으므로 저장하지 않고 닫음
        myFile = Dir ' 다음 파일로 이동
    Loop

    masterWorkbook.Save
    'masterWorkbook.Close
    MsgBox "완료"

    Application.StatusBar = False ' 상태 바 초기화
    Application.ScreenUpdating = True ' 화면 업데이트
    Application.DisplayAlerts = True
End Sub
