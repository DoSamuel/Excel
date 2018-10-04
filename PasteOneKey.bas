
    Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long

 
Private Sub CE()

    Dim lngRet As Long
     
    lngRet = OpenClipboard(Application.hwnd)
     
    If lngRet Then
        EmptyClipboard
        CloseClipboard
    End If
End Sub

Private Sub 一键粘贴()
Range("A29").Select

    Range("A29:AS999").Select
    Selection.NumberFormatLocal = "G/通用格式"
    Range("A29").Select

分别粘贴

Dim u8, u9, t8, re As Integer

    u8 = Range("P4").Value
    u9 = Range("P5").Value
    t8 = Range("O4").Value

If t8 = 1 Then
    For re = 1 To 14 Step 1
    循环
    Next re
Else
    If u8 = 2 Then
        If u9 = 2 Then
            For re = 1 To 14 Step 1
            循环
            Next re
        Else
            For re = 1 To 14 Step 1
            循环
            Next re
            ActiveCell.Offset(0, 2).Select
            For re = 1 To 3 Step 1
            循环
            Next re
        End If
    Else
        If u9 = 2 Then
            For re = 1 To 15 Step 1
            循环
            Next re
            循环2
        Else
            For re = 1 To 18 Step 1
            循环
            Next re
            循环2
        End If
    End If
End If

粘贴总数据

Range("A29").Select

分别粘贴

Range("B2").Select
 
    MsgBox "粘贴完成"

End Sub
Private Sub 粘贴总数据()
'
startt:

' 粘贴总数据 宏
'

On Error Resume Next


Range("AM29").Select
'
Dim b As Integer

    ActiveCell.Range("A1:F30").Select
    Selection.ClearContents
    ActiveCell.Select
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
AppActivate "Google Chrome"
s6
SendKeys "^t"
SendKeys "^v"
SendKeys "{enter}"
s4
SendKeys "^a"
SendKeys "^c"
s6
SendKeys "{NUMLOCK}"

AppActivate "Excel"

ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
s6
    
CE

    ActiveCell.Offset(30, 0).Select
    
        ActiveCell.Range("A1:F30").Select
    Selection.ClearContents
    ActiveCell.Select

AppActivate "Google Chrome"
s6

For b = 1 To 21 Step 1

SendKeys "{tab}"
Next b
SendKeys "{enter}"
s2
SendKeys "^a"
SendKeys "^c"

s6
AppActivate "Excel"
    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
        
CE
        
    ActiveCell.Offset(30, 0).Select
    
            ActiveCell.Range("A1:F30").Select
    Selection.ClearContents
    ActiveCell.Select
    
AppActivate "Google Chrome"
s6

For b = 1 To 3 Step 1

SendKeys "{tab}"
Next b
SendKeys "{enter}"
s2
SendKeys "^a"
SendKeys "^c"
s6
SendKeys "^w"
AppActivate "Excel"
    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
        
CE
        
    ActiveCell.Offset(30, 0).Select
    
    Range("A1").Select
    
    SendKeys "{NUMLOCK}"

If Err.Number > 0 Or Range("B4") = "#N/A" Then
Err.Clear
GoTo startt
End If

    
End Sub

Private Sub 复制总数据()
'
' 复制总数据 宏
'

'
muban

ActiveWorkbook.Save

    Range("B2:J22").Select
    Selection.Copy
    Range("M2").Select
    s6
AppActivate "商家服务数据时段播报 - Excel"

End Sub

Private Sub 循环()
line2:
On Error Resume Next

ActiveCell.Offset(1, 0).Range("A1:B971").Select
    Selection.ClearContents
    ActiveCell.Offset(-1, 0).Range("A1").Select
    
        ActiveCell.Offset(-1, 0).Range("A1").Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select

AppActivate "Google Chrome"
s6
SendKeys "^t"
SendKeys "^v"
SendKeys "{enter}"
s4
SendKeys "^a"
SendKeys "^c"
s6
SendKeys "^w"
SendKeys "{NUMLOCK}"

AppActivate "Excel"

 ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True

CE
If Err.Number > 0 Then
Err.Clear
GoTo line2
End If
    s6
        ActiveCell.Offset(0, 2).Select

End Sub
Private Sub 分别粘贴()

Dim col As Integer
Dim row As Integer
col = Selection.Column
row = Selection.row

If row = 29 Then
line1:
    ActiveCell.Offset(1, 0).Range("A1:B971").Select
    Selection.ClearContents
    ActiveCell.Offset(-1, 0).Range("A1").Select
    
        ActiveCell.Offset(-1, 0).Range("A1").Select

    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select

On Error Resume Next

AppActivate "Google Chrome"
s6
SendKeys "^t"
SendKeys "^v"
SendKeys "{enter}"
s4
SendKeys "^a"
SendKeys "^c"
s6
SendKeys "^w"
AppActivate "Excel"
s6
    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
CE
If Err.Number > 0 Then
Err.Clear
GoTo line1
End If

    ActiveCell.Offset(0, 2).Select

SendKeys "{NUMLOCK}"
Else
MsgBox "请移动到绿色区域"
End If


End Sub
Private Sub s4() '暂停4秒，期间可以进行其他操作
    '前面的代码
    t = Timer
    While Timer < t + 3.5
        DoEvents
    Wend
    '后面的代码
End Sub


Private Sub s2() '暂停2秒，期间可以进行其他操作
    '前面的代码
    t = Timer
    While Timer < t + 2
        DoEvents
    Wend
    '后面的代码
End Sub
Private Sub s6() '暂停0.3秒，期间可以进行其他操作
    '前面的代码
    t = Timer
    While Timer < t + 0.3
        DoEvents
    Wend
    '后面的代码
End Sub
Private Sub 循环2()
first:
On Error Resume Next

Range("AT29").Select
ActiveCell.Offset(1, 0).Range("A1:J599").Select
Selection.NumberFormatLocal = "G/通用格式"
    Selection.ClearContents
    ActiveCell.Offset(-1, 0).Range("A1").Select
    
        ActiveCell.Offset(-1, 0).Range("A1").Select

    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
AppActivate "Google Chrome"
s6
SendKeys "^t"
SendKeys "^v"
SendKeys "{enter}"
s2

Dim ae
For ae = 1 To 6 Step 1

SendKeys "{tab}"
Next ae
s6
SendKeys "{Enter}"
s6
SendKeys "{Down}"
s6
SendKeys "{Enter}"
s6
s6
SendKeys "^a"
SendKeys "^c"
s6
SendKeys "^w"
SendKeys "{NUMLOCK}"

AppActivate "Excel"

 ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
        
CE
        
    s6
Range("A1").Select
        
If Err.Number > 0 Or Range("AT41").Value = 0 Then
Err.Clear
GoTo first
End If



End Sub

Private Sub muban()
'
' muban 宏
'

'
    Range("B1:B22").Select
    Selection.Copy
    ActiveSheet.Previous.Select
    Range("A264").Select
    Selection.End(xlToRight).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.End(xlUp).Select
        Dim c1
    c1 = Selection.row
        If c1 = 1 Then

    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Else
        
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        End If
        ActiveSheet.Next.Select
End Sub




