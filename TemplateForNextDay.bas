Private Sub 模板()
'
' 模板 宏
'
' 快捷键: Ctrl+m

            Sheets.Add After:=ActiveSheet   ' 添加sheet
            ActiveSheet.Previous.Select
            Range("A327:H335").Select
            Selection.Copy
            Range("D322").Select
            ActiveSheet.Next.Select
            Range("A327").Select
            ActiveSheet.Paste   ' 复制日期以及按钮
            
        Dim a As String
    
            a = Range("H328").Value
            ActiveSheet.Select
            ActiveSheet.Name = a        ' 以日期命名sheet
            
            Range("A1").Select
            Rows("1:25").Select
            Selection.EntireRow.Hidden = True   ' 隐藏1-25行
            Range("A26").Select
    
        Dim b As Integer
    
            b = Range("H329").Value
            
    If b = 1 Then   ' 如果当天不是周日
    
                 Dim j As Integer
                j = Range("H330").Value
            If j = 1 Then        ' 如果当天是周五周六
               Sheets("模板").Select
            
            Range("BN26:CD48").Select
            Selection.Copy
            Range("J27").Select
            ActiveSheet.Previous.Select
            Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False     ' 粘贴为原始宽度
            ActiveSheet.Paste
            
            Dim x As Integer
            For x = 51 To 301 Step 25
                Range("A" & x).Select
                ActiveSheet.Paste
            Next x      ' 粘贴剩余的模板

        
        Dim y As Integer
        
        For y = 26 To 301 Step 25
            
            Range("A" & y & ":" & "C" & (y + 20)).Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
        Next y      ' 去公式
        
            Range("B26").Select
            
   Else
            Sheets("模板").Select
            
            Range("AA26:AQ48").Select
            Selection.Copy
            Range("J27").Select
            ActiveSheet.Previous.Select
            Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False     ' 粘贴为原始宽度
            ActiveSheet.Paste
            
            Dim k As Integer
            For k = 51 To 301 Step 25
                Range("A" & k).Select
                ActiveSheet.Paste
            Next k      ' 粘贴剩余的模板

        
        Dim l As Integer
        
        For l = 26 To 301 Step 25
            
            Range("A" & l & ":" & "C" & (l + 20)).Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
        Next l      ' 去公式
        
    
    
    Dim o As String                 '六点之后合并单元格，填写下班信息

o = "AE/快递 下班"
    Range("M268:N271").Select
    Selection.ClearContents
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("M268").Value = o
    
   
    Range("M293:N296").Select
    Selection.ClearContents
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
Range("M293").Value = o

    Range("M318:N321").Select
    Selection.ClearContents
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("M318").Value = o
        
            Range("B26").Select
    End If
    Else    ' 如果当天是周日
    
            Sheets("模板").Select
            
            Range("AS26:BL48").Select
            Selection.Copy
            
            Range("J27").Select
            ActiveSheet.Previous.Select
            Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            ActiveSheet.Paste
            
        
        Dim z As Integer
        
        For z = 51 To 301 Step 25
            Range("A" & z).Select
            ActiveSheet.Paste
            
        Next z      ' 粘贴剩余的模板
        
        
        Dim w As Integer
        
        For w = 26 To 301 Step 25
            
            Range("A" & w & ":" & "D" & (w + 20)).Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
        Next w      ' 去公式
        
        Dim p As String

p = "AE/快递 下班"
    Range("O268:P271").Select
    Selection.ClearContents
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("O268").Value = p
    
   
    Range("O293:P296").Select
    Selection.ClearContents
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
Range("O293").Value = p

    Range("O318:P321").Select
    Selection.ClearContents
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("O318").Value = p
        
            Range("C26").Select
            
            
    End If
            
End Sub


