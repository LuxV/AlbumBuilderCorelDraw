Attribute VB_Name = "modParser"
Option Explicit

Function ExpandOrnament(abbr As String) As String
    Static dict As Object
    If dict Is Nothing Then
        Set dict = CreateObject("Scripting.Dictionary")
        dict.Add "н.о.", "ногтевым орнаментом"
        dict.Add "т.о.", "тычковым орнаментом"
        dict.Add "л.о.", "линейным орнаментом"
        dict.Add "в.о.", "волнистым орнаментом"
        dict.Add "ш.о.", "штампованным орнаментом"
        dict.Add "орн.", "орнаментом"
        dict.Add "к.н.", "косыми насечками"
    End If

    If dict.Exists(abbr) Then
        ExpandOrnament = dict(abbr)
    Else
        ExpandOrnament = abbr
    End If
End Function


Sub BuildParagraphTextFromFile(ByVal fPath As String)
    gTextCursorYmm = 10

    Dim lines() As String
    lines = ReadAllLines(fPath)

    Dim i As Long
    Dim NumBlockCounter As Integer
    
    i = 0
    NumBlockCounter = 1
    
    AddGuidesToPage ActivePage
        
    Do While i <= UBound(lines)
        
       
        If IsBlockTitle(lines(i)) Then
            
            i = ProcessBlock(lines, i, NumBlockCounter)
                        
            If NumBlockCounter Mod 2 = 0 Then
                EnsureNewPage ActiveDocument, ""
            End If
            
            NumBlockCounter = NumBlockCounter + 1
            
        Else
            i = i + 1
        End If
 
    Loop
End Sub


Function ProcessBlock(lines() As String, startIdx As Long, NumBlockCounter As Integer) As Long
    Dim title As String
    title = ExtractTitle(lines(startIdx))
    
    Dim i As Long: i = startIdx + 2   ' первая строка таблицы
    Dim resultText As String
    Dim maxNum As Long
    Dim blocks As New Collection
    
    resultText = "Илл. 00. ВСТАВИТЬ НАЗВАНИЕ ОБЪЕКТА " & LCase(title) & _
                 ". Массовый материал. Керамика, фрагменты сосудов: "
    
    ' --- строки таблицы ---
    Do While i <= UBound(lines)
        If IsSeparatorLine(lines(i)) Then Exit Do
        If IsBlockTitle(lines(i)) Then Exit Do
        
        ProcessRow lines(i), resultText, maxNum, blocks
        i = i + 1
    Loop
    
    ' --- создаём текст параграфа и блоки чисел со свойствами ---
    CreateParagraphText resultText, NumBlockCounter
    CreateNumberBlocks blocks, NumBlockCounter
    
    ' --- пропуск разделителей ---
    Do While i <= UBound(lines)
        If Not IsSeparatorLine(lines(i)) Then Exit Do
        i = i + 1
    Loop
    
    ProcessBlock = i
End Function

Function ReadAllLines(path As String) As String()
    Dim f As Integer: f = FreeFile
    Open path For Input As #f

    Dim buf() As String
    Dim n As Long: n = -1
    Dim s As String

    Do Until EOF(f)
        Line Input #f, s
        n = n + 1
        ReDim Preserve buf(n)
        buf(n) = Trim(s)
    Loop

    Close #f
    ReadAllLines = buf
End Function


Sub ProcessRow(line As String, ByRef outText As String, ByRef maxNum As Long, ByRef blocks As Collection)
    Dim parts() As String
    parts = Split(line, ";")
    
    outText = outText & Trim(parts(0)) & ": "
    
    AppendCell parts(1), "венчики", outText, maxNum, blocks
    AppendCell parts(2), "стенки", outText, maxNum, blocks
    AppendCell parts(3), "донца", outText, maxNum, blocks
    AppendCell parts(4), "", outText, maxNum, blocks
    
    ' Чтобы не обрезать строку при пустых колонках:
    'If Right(outText, 2) = ", " Then
    '    outText = Left(outText, Len(outText) - 2)
    'End If
    '
    'outText = outText & "; "
    
    outText = Left(outText, Len(outText) - 2) & "; "
    
End Sub


Sub AppendCell(cellText As String, partName As String, _
               ByRef outText As String, ByRef maxNum As Long, _
               ByRef blocks As Collection)

    cellText = Trim(cellText)
    If cellText = "" Then Exit Sub

    Dim items() As String
    items = Split(cellText, ",")

    Dim numbers As String
    Dim totalNumbers As Long
    Dim textMarkers As New Collection

    Dim it As Variant
    Dim n As Long

    For Each it In items

        Dim s As String
        s = Trim(it)

        Dim brackets As Collection
        Set brackets = SplitBrackets(s)

        Dim baseNumber As String
        baseNumber = Trim(RemoveBracketText(s))

        Dim n1 As Long, n2 As Long
        Dim a() As String

        If InStr(baseNumber, "-") > 0 Then
            a = Split(baseNumber, "-")
            n1 = CLng(a(0))
            n2 = CLng(a(1))
        Else
            n1 = CLng(baseNumber)
            n2 = n1
        End If

        Dim props As New Collection
        Dim prop As Variant
        
        ' маркеры добавляем один раз
        For Each prop In brackets
        
            Dim specialPartName As String
            
            If InStr(CStr(prop), "=") > 0 Then
                
                props.Add CStr(prop)
            
            ElseIf InStr(CStr(prop), ".") > 0 Then
                
                textMarkers.Add ExpandOrnament(CStr(prop))
            
            Else
                
                specialPartName = CStr(prop)
            
            End If
        
        Next
        
        
        For n = n1 To n2
        
            numbers = numbers & n & ", "
            totalNumbers = totalNumbers + 1
        
            If n > maxNum Then maxNum = n
        
            Dim nb As clsNumberBlock
            Set nb = New clsNumberBlock
            nb.NumberValue = n
        
            Dim p As Variant
            For Each p In props
                nb.Properties.Add p
            Next
        
            blocks.Add nb

        Next

    Next

    numbers = Left(numbers, Len(numbers) - 2)

    ' --- определение формы слова ---
    Dim isSingle As Boolean
    isSingle = (totalNumbers = 1)

    ' --- формирование текста ---
    outText = outText & numbers

    If partName <> "" Then
        
        outText = outText & " – " & PartLabel(partName, isSingle) & JoinMarkers(textMarkers)
    
    ElseIf specialPartName <> "" Then
        
        outText = outText & " – " & specialPartName & JoinMarkers(textMarkers)
    
    End If

    outText = outText & ", "

End Sub


Function PartLabel(ByVal base As String, ByVal isSingle As Boolean) As String
    Select Case base
        Case "венчики": PartLabel = IIf(isSingle, "венчик", "венчики")
        Case "стенки":  PartLabel = IIf(isSingle, "стенка", "стенки")
        Case "донца":   PartLabel = IIf(isSingle, "донце", "донца")
        Case Else:      PartLabel = base
    End Select
End Function

Function IsSingleNumber(s As String) As Boolean
    IsSingleNumber = (InStr(s, ",") = 0 And InStr(s, "-") = 0)
End Function

Function JoinMarkers(markers As Collection) As String
    Dim i As Long
    Dim result As String

    If markers.Count = 0 Then Exit Function

    If markers.Count = 1 Then
        JoinMarkers = " с " & markers(1)
        Exit Function
    End If

    result = " с "
    For i = 1 To markers.Count
        Select Case i
            Case markers.Count
                result = result & "и " & markers(i)
            Case Else
                result = result & markers(i) & " "
        End Select
    Next

    JoinMarkers = Trim(result)
End Function
