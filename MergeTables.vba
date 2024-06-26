Sub MergeTables()

    Application.ScreenUpdating = False ' 화면 업데이트 끄기
    Application.DisplayAlerts = False ' 경고 표시 안함

    Dim wsKey As Worksheet, wsSecondary As Worksheet, wsResult As Worksheet
    Dim lastRowKey As Long, lastRowSec As Long, lastColKey As Long, lastColSec As Long
    Dim rKey As Long, rSec As Long, c As Long
    Dim dict As Object
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    Set wsKey = Sheets(1)
    Set wsSecondary = Sheets(2)
    Set wsResult = Sheets(3)
    
    wsResult.Cells.Clear
    
    ' 해더 복사
    lastColKey = wsKey.Cells(1, Columns.Count).End(xlToLeft).Column
    wsKey.Rows(1).Copy
    wsResult.Rows(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    
    Application.CutCopyMode = False
    
    ' 키와 속성 값 범위 탐색
    lastRowKey = wsKey.Cells(Rows.Count, 1).End(xlUp).Row
    lastRowSec = wsSecondary.Cells(Rows.Count, 1).End(xlUp).Row
    lastColSec = wsSecondary.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' 기준 시트에서 데이터 복사 (값 + 서식)
    For rKey = 2 To lastRowKey
        For c = 1 To lastColKey
            wsKey.Cells(rKey, c).Copy
            wsResult.Cells(rKey, c).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
        Next c
        dict(wsKey.Cells(rKey, 1).value) = rKey
    Next rKey
    
    Application.CutCopyMode = False
    
    ' 보조 시트 데이터 병합 (값 + 서식)
    For rSec = 2 To lastRowSec
        If dict.Exists(wsSecondary.Cells(rSec, 1).value) Then
            For c = 1 To lastColSec
                If wsResult.Cells(dict(wsSecondary.Cells(rSec, 1).value), c).value = "" Then
                    wsSecondary.Cells(rSec, c).Copy
                    wsResult.Cells(dict(wsSecondary.Cells(rSec, 1).value), c).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
                End If
            Next c
        Else
            rKey = rKey + 1
            For c = 1 To lastColSec
                wsSecondary.Cells(rSec, c).Copy
                wsResult.Cells(rKey, c).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
            Next c
        End If
    Next rSec
    
    Application.CutCopyMode = False
    
    Application.StatusBar = False ' 상태 바 초기화
    Application.ScreenUpdating = True ' 화면 업데이트
    
    
    ' 결과 시트 정리
    wsResult.Columns.AutoFit
    MsgBox "병합완료", vbInformation
    
End Sub

