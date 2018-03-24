# Excel
Excel test


Private Sub 一键粘贴()
Range("A29").Select

    Range("A29:AK628").Select
    Selection.NumberFormatLocal = "G/通用格式"
    Range("A29").Select

分别粘贴
Dim we
we = Range("I20").Value
If we = 1 Then
Dim awe
For awe = 1 To 6 Step 1
循环
Next awe
ActiveCell.Offset(0, 4).Select
循环
Else
Dim ti
ti = Range("I21").Value
If ti = 1 Then
Dim ati
For ati = 1 To 6 Step 1
循环
Next ati
ActiveCell.Offset(0, 4).Select
循环
Else

Dim a
For a = 1 To 9 Step 1

循环

Next a
ActiveCell.Offset(0, 7).Select
Dim ab
For ab = 1 To 5 Step 1

循环2

Next ab
End If
End If

Range("U29").Select
s6
Selection.Copy
s6
粘贴总数据
 
    MsgBox "粘贴完成"
    
    
End Sub
Private Sub 粘贴总数据()
'
' 粘贴总数据 宏
'
Range("U29").Select
'
Dim b As Integer

    ActiveCell.Range("A1:F30").Select
    Selection.ClearContents
    ActiveCell.Select
            ActiveCell.Offset(-1, 0).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
AppActivate "Google Chrome"
s6
SendKeys "{F5}"
s6
SendKeys "{tab}"
SendKeys "+{tab}"
s6
SendKeys "^a"
SendKeys "^v"
SendKeys "{enter}"
s2
SendKeys "^a"
SendKeys "^c"
s6
SendKeys "{NUMLOCK}"

AppActivate "Excel"

 ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
    s6
    ActiveCell.Offset(30, 0).Select
    
        ActiveCell.Range("A1:F30").Select
    Selection.ClearContents
    ActiveCell.Select
            ActiveCell.Offset(0, 6).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(0, -6).Range("A1").Select
    ActiveWindow.ScrollColumn = 1
    
AppActivate "Google Chrome"
s6

For b = 1 To 21 Step 1

SendKeys "{tab}"
Next b
SendKeys "{enter}"
s2
SendKeys "^a"
SendKeys "^c"
SendKeys "{F5}"
s6
AppActivate "Excel"
    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
    ActiveCell.Offset(30, 0).Select
    
            ActiveCell.Range("A1:F30").Select
    Selection.ClearContents
    ActiveCell.Select
            ActiveCell.Offset(0, 6).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(0, -6).Range("A1").Select
    ActiveWindow.ScrollColumn = 1
    
AppActivate "Google Chrome"
s6

For b = 1 To 22 Step 1

SendKeys "{tab}"
Next b
SendKeys "{enter}"
s2
SendKeys "^a"
SendKeys "^c"
SendKeys "{F5}"
s6
AppActivate "Excel"
    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
    ActiveCell.Offset(30, 0).Select
    
    ActiveWindow.ScrollRow = 1
    
    SendKeys "{NUMLOCK}"

    
End Sub

Private Sub 复制总数据()
'
' 复制总数据 宏
'

'
    Range("B2:H21").Select
    Selection.Copy
    Range("K2").Select
    ActiveSheet.Previous.Select
    ActiveSheet.Previous.Select

End Sub

Private Sub 循环()

ActiveCell.Offset(1, 0).Range("A1:B599").Select
    Selection.ClearContents
    ActiveCell.Offset(-1, 0).Range("A1").Select
    
        ActiveCell.Offset(-1, 0).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
AppActivate "Google Chrome"
s6
SendKeys "{tab}"
SendKeys "+{tab}"
s6
SendKeys "^a"
SendKeys "^v"
SendKeys "{enter}"
s2
SendKeys "^a"
SendKeys "^c"
s6
SendKeys "{NUMLOCK}"

AppActivate "Excel"

 ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
    s6
        ActiveCell.Offset(0, 2).Select
        

End Sub
Private Sub 分别粘贴()

Dim col As Integer
Dim row As Integer
col = Selection.Column
row = Selection.row

If row = 29 Then
If col = 1 Or col = 3 Or col = 5 Or col = 7 Or col = 9 Or col = 11 Or col = 13 Or col = 15 Or col = 17 Or col = 19 Or col = 28 Or col = 30 Or col = 32 Or col = 34 Or col = 36 Then

    ActiveCell.Offset(1, 0).Range("A1:B599").Select
    Selection.ClearContents
    ActiveCell.Offset(-1, 0).Range("A1").Select
    
        ActiveCell.Offset(-1, 0).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    
AppActivate "Google Chrome"
s6
SendKeys "{F5}"
s2
SendKeys "{tab}"
SendKeys "+{tab}"
s6
SendKeys "^a"
SendKeys "^v"
SendKeys "{enter}"
s2
SendKeys "^a"
s6
SendKeys "^c"
SendKeys "{F5}"
s6
AppActivate "Excel"
s6
    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
    ActiveCell.Offset(0, 2).Select

SendKeys "{NUMLOCK}"
Else
MsgBox "请移动到绿色区域"
End If
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
    While Timer < t + 1.5
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

ActiveCell.Offset(1, 0).Range("A1:B599").Select
    Selection.ClearContents
    ActiveCell.Offset(-1, 0).Range("A1").Select
    
        ActiveCell.Offset(-1, 0).Range("A1").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
AppActivate "Google Chrome"
s6
SendKeys "{tab}"
SendKeys "+{tab}"
s6
SendKeys "^a"
SendKeys "^v"
SendKeys "{enter}"
s2
SendKeys "^a"
SendKeys "^c"
s6
SendKeys "{NUMLOCK}"

AppActivate "Excel"

 ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
    s6
        ActiveCell.Offset(0, 2).Select
        

End Sub
