Sub 截图()
'
' 截图 宏
'
' 快捷键: Ctrl+j
'

    Dim f As String
    
    f = ActiveCell.Value
    
    If f = "LOB" Then
    
    Dim c As Integer
        c = Range("H331").Value
    If c = 0 Then
    
        If ActiveCell.Offset(22, 0).Range("A1").Value = "" Then
        ActiveCell.Offset(0, 0).Range("A1:N22").Select
        Application.CutCopyMode = False
        Selection.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        ActiveWindow.SmallScroll Down:=9
        ActiveCell.Offset(25, 0).Range("A1").Select
        ActiveWorkbook.Save
       Else
       ActiveCell.Offset(0, 0).Range("A1:N23").Select
        Application.CutCopyMode = False
        Selection.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        ActiveWindow.SmallScroll Down:=9
        ActiveCell.Offset(25, 0).Range("A1").Select
        ActiveWorkbook.Save
        End If
    Else
    If ActiveCell.Offset(22, 0).Range("A1").Value = "" Then
        ActiveCell.Offset(0, 0).Range("A1:M22").Select
        Application.CutCopyMode = False
        Selection.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        ActiveWindow.SmallScroll Down:=6
        ActiveCell.Offset(25, 0).Range("A1").Select
        ActiveWorkbook.Save
      Else
      ActiveCell.Offset(0, 0).Range("A1:M23").Select
        Application.CutCopyMode = False
        Selection.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        ActiveWindow.SmallScroll Down:=6
        ActiveCell.Offset(25, 0).Range("A1").Select
        ActiveWorkbook.Save
        End If
    End If
    
        MsgBox "成功复制截图到剪切板"
        
       Dim s
       
        s = Shell("C:\Program Files (x86)\DingDing\main\current\DingTalk.exe", 1)
    
    Set s = Nothing
Else
        MsgBox "请移动到 LOB 单元格"
End If

End Sub
