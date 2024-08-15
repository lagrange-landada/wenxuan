' 异文格式可参考胡克家《文选考异》
' 格式一：XXXX：YYYY。=>那么就会在XXXX处，插入脚注：XXXX：YYYY。
' 格式二：注「XXXX」：YYYY。=>那么就会在XXXX处，插入脚注：注「XXXX」：YYYY。
' 格式三：注「XXXX」下至「YYYY」：ZZZZ。=>那么就会在YYYY处，插入脚注：注「XXXX」下至「YYYY」：ZZZZ。
' 注意事项：
' 1、设置异文路径；
' 1、每条异文至少应包含一个冒号；
' 2、最后一行不要空；
' 3、虽然插入脚注的时候，是按照向下的顺序依次插入的，但是尽量保证每条异文是唯一性的，比如：
' 		注「説文曰」：XXXX。=>待插入的文档中，可能会存在多条符合插入条件的位置。如果是这样，建议分批插入；
' 4、对于没有插入成功的异文，会生成一个”未匹配.txt“文档，用来手动添加。
Sub 批量添加脚注()
    Dim rng As Range
    Dim footnoteText As String
    Dim regex As Object
    Dim matches As Object
    Dim i As Integer
    Dim filePath As String
    Dim stream As Object
    Dim fileContent As String
    Dim footnotesArray() As String
    Dim regexPatterns() As String
    Dim pattern As String
    Dim docRange As Range
    Dim startPos As Long
    Dim addedFootnotes As Object ' Dictionary 对象
    Dim processedFootnotes As Object ' 处理过的脚注列表
    Dim unmatchedFilePath As String
    Dim unmatchedStream As Object
    Dim unmatchedFootnotes As Collection
    Dim unmatchedText As Variant ' 这里将变量定义为 Variant 类型
    
    ' 创建正则表达式对象
    Set regex = CreateObject("VBScript.RegExp")
    
    ' 设置正则表达式属性
    regex.IgnoreCase = True
    regex.Global = False ' 每个正则表达式只匹配一次
    
    ' 设置文件路径
    filePath = "E:\A书籍\文学\国内文学\文选\新建文件夹\资源.txt" ' 请替换为实际的文件路径
    unmatchedFilePath = "E:\A书籍\文学\国内文学\文选\新建文件夹\未匹配.txt" ' 未匹配文件的路径
    
    ' 创建 ADODB.Stream 对象
    Set stream = CreateObject("ADODB.Stream")
    
    On Error GoTo ErrorHandler
    
    ' 配置 Stream 对象
    With stream
        .Type = 2 ' 指定为文本类型
        .Charset = "UTF-8" ' 根据实际文件编码设置
        .Open
        .LoadFromFile filePath
        fileContent = .ReadText
        .Close
    End With
    
    ' 将文件内容按行分割
    footnotesArray = Split(fileContent, vbCrLf)
    
    ' 创建正则表达式和脚注内容数组
    ReDim regexPatterns(UBound(footnotesArray))
    
    ' 提取正则表达式和脚注内容
    For i = 0 To UBound(footnotesArray)
        ' 按第一个“：”分割
        Dim parts() As String
        parts = Split(footnotesArray(i), "：", 2)
        If UBound(parts) = 1 Then
            ' 处理多个「」符号的情况
            Dim patternText As String
            patternText = parts(0)
            If InStr(patternText, "「") > 0 Then
                Dim lastBracketStart As Long
                Dim lastBracketEnd As Long
                lastBracketStart = InStrRev(patternText, "「")
                lastBracketEnd = InStrRev(patternText, "」")
                If lastBracketStart > 0 And lastBracketEnd > lastBracketStart Then
                    patternText = Mid(patternText, lastBracketStart + 1, lastBracketEnd - lastBracketStart - 1)
                End If
            End If
            
            ' 生成正则表达式
            patternText = InsertWildcards(patternText)
            
            regexPatterns(i) = patternText
            footnotesArray(i) = footnotesArray(i) ' 获取整个行作为脚注内容
        End If
    Next i
    
    ' 获取整个文档内容
    Set docRange = ActiveDocument.Content
    
    ' 设置初始起始位置
    startPos = 0
    
    ' 创建 Dictionary 对象
    Set addedFootnotes = CreateObject("Scripting.Dictionary")
    Set processedFootnotes = CreateObject("Scripting.Dictionary")
    Set unmatchedFootnotes = New Collection
    
    ' 遍历所有正则表达式并添加脚注
    For i = 0 To UBound(regexPatterns)
        pattern = regexPatterns(i)
        footnoteText = footnotesArray(i)
        
        ' 设置正则表达式模式
        regex.pattern = pattern
        
        ' 查找范围
        docRange.Start = startPos
        docRange.End = ActiveDocument.Content.End
        If regex.Test(docRange.text) Then
            Set matches = regex.Execute(docRange.text)
            If matches.Count > 0 Then
                ' 找到匹配位置
                Dim matchPos As Long
                matchPos = matches(0).FirstIndex + 1 + startPos
                
                ' 检查脚注内容是否已处理过
                If Not processedFootnotes.Exists(footnoteText) Then
                    ' 添加脚注
                    ActiveDocument.Footnotes.Add Range:=ActiveDocument.Range(Start:=matchPos, End:=matchPos + Len(matches(0).Value) - 1), text:=footnoteText
                    
                    ' 标记脚注内容为已处理
                    processedFootnotes.Add footnoteText, True
                    
                    ' 移动光标到匹配位置
                    ActiveDocument.Range(Start:=matchPos, End:=matchPos).Select
                End If
                
                ' 更新起始位置
                startPos = matchPos + Len(matches(0).Value)
            End If
        Else
            ' 未匹配的脚注内容加入未匹配列表
            unmatchedFootnotes.Add footnotesArray(i)
        End If
    Next i
    
    ' 将未匹配的脚注写入文件
    Set unmatchedStream = CreateObject("ADODB.Stream")
    With unmatchedStream
        .Type = 2 ' 指定为文本类型
        .Charset = "UTF-8" ' 根据实际文件编码设置
        .Open
        For Each unmatchedText In unmatchedFootnotes
            .WriteText unmatchedText & vbCrLf
        Next
        .SaveToFile unmatchedFilePath, 2 ' 2 = adSaveCreateOverWrite
        .Close
    End With
    
    MsgBox "脚注添加完成。未匹配的脚注已写入到“未匹配.txt”。"

    Exit Sub

ErrorHandler:
    MsgBox "读取文件时出错。请检查文件路径和文件编码。"
    If stream.State = 1 Then stream.Close
    If unmatchedStream.State = 1 Then unmatchedStream.Close
End Sub

' 将每个字符之间插入 ".*"
Function InsertWildcards(text As String) As String
    Dim result As String
    Dim i As Integer
    
    result = ""
    
    For i = 1 To Len(text)
        result = result & Mid(text, i, 1) & "([，。、：])?"
    Next i
    
    ' 移除最后一个 ".*"
    If Len(result) > 0 Then
        result = Left(result, Len(result) - 9)
    End If
    
    InsertWildcards = result
End Function