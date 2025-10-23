Option Explicit

Dim today, weekdayNum
today = Date
weekdayNum = Weekday(today, vbMonday) ' 将周一作为第一天，周一=1, 周日=7

'-------------------------------------------
' 自定义上班日与休息日
'-------------------------------------------
' 默认上班日（周一到周日，1到7）
Dim defaultWorkDays
defaultWorkDays = Array(1, 2, 3, 4, 5)

' 额外上班日（如补班日，可用 yyyy-mm-dd 格式）
Dim extraWorkDays
extraWorkDays = Array("2025-10-26", "2025-11-08")

' 额外休息日（如节假日，可用 yyyy-mm-dd 格式）
Dim extraOffDays
extraOffDays = Array("2025-10-01", "2025-10-02", "2025-10-03")

'-------------------------------------------
' 判断逻辑
'-------------------------------------------
Dim isWorkDay
isWorkDay = False

If IsInArray(FormatDate(today), extraOffDays) Then
    isWorkDay = False
ElseIf IsInArray(FormatDate(today), extraWorkDays) Then
    isWorkDay = True
ElseIf IsInArray(weekdayNum, defaultWorkDays) Then
    isWorkDay = True
End If

'-------------------------------------------
' 弹出提醒
'-------------------------------------------
If isWorkDay Then
    MsgBox "tip-waimai"
End If

'-------------------------------------------
' 辅助函数
'-------------------------------------------
Function IsInArray(val, arr)
    Dim i
    For i = 0 To UBound(arr)
        If arr(i) = val Then
            IsInArray = True
            Exit Function
        End If
    Next
    IsInArray = False
End Function

Function FormatDate(d)
    FormatDate = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
End Function
