Attribute VB_Name = "modHelpers"
Option Explicit

Function IsBlockTitle(s As String) As Boolean
    Dim parts() As String
    parts = Split(s, ";")

    ' первая колонка заполнена
    If Trim(parts(0)) = "" Then Exit Function

    ' все остальные колонки пустые
    Dim i As Long
    For i = 1 To UBound(parts)
        If Trim(parts(i)) <> "" Then Exit Function
    Next

    IsBlockTitle = True
End Function

Function ExtractTitle(s As String) As String
    ExtractTitle = Trim(Replace(s, ";;;;;", ""))
End Function

Function ExtractBracketText(s As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(s, "(")
    p2 = InStr(s, ")")
    If p1 > 0 And p2 > p1 Then
        ExtractBracketText = Mid(s, p1 + 1, p2 - p1 - 1)
    End If
End Function

Function RemoveBracketText(s As String) As String
    Dim p1 As Long
    p1 = InStr(s, "(")
    If p1 > 0 Then
        RemoveBracketText = Left(s, p1 - 1)
    Else
        RemoveBracketText = s
    End If
End Function

Function IsSeparatorLine(s As String) As Boolean
    IsSeparatorLine = Trim(Replace(s, ";", "")) = ""
End Function

Function SplitBrackets(s As String) As Collection
    ' Возвращает коллекцию всех скобок в строке
    Dim c As New Collection
    Dim startPos As Long, endPos As Long
    startPos = InStr(s, "(")
    Do While startPos > 0
        endPos = InStr(startPos, s, ")")
        If endPos = 0 Then Exit Do
        c.Add Mid(s, startPos + 1, endPos - startPos - 1)
        startPos = InStr(endPos + 1, s, "(")
    Loop
    Set SplitBrackets = c
End Function
