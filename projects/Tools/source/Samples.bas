Attribute VB_Name = "Samples"
Option Explicit

Const ILL_PATTERN As String = "Илл\.[^0-9]*(\d*)\."
Private Const V_GUIDE_LEFT As Double = 25
Private Const V_GUIDE_RIGHT As Double = 195
Private Const H_GUIDE_BOTTOM As Double = 40
Private Const H_GUIDE_MID As Double = 150
Private Const H_GUIDE_TOP As Double = 280

Private Const IMAGE_TARGET_WIDTH As Double = 170

Sub ReplaceIllWithNumber_TopDownSorted()
    Dim pg As Page, sh As Shape
    Dim regEx As Object
    Dim items() As Variant
    Dim counter As Long, idx As Long
    Dim startN As Integer
    Dim txt As String, newText As String, s As String
    Dim fixedCount As Long, alreadyOkCount As Long
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = ILL_PATTERN
    regEx.IgnoreCase = True
    regEx.Global = False
    
    counter = 0
    fixedCount = 0
    alreadyOkCount = 0
    
    s = InputBox("С какого числа начать нумерацию?", "Нумерация Илл.", "1")
    If s = "" Then Exit Sub ' отмена
    If Not IsNumeric(s) Then
        MsgBox "Введите число!", vbExclamation
        Exit Sub
    End If
    counter = CLng(s)
    If counter <= 0 Then
        MsgBox "Количество должно быть больше 0!", vbExclamation
        Exit Sub
    End If
     
    counter = counter - 1
    startN = counter
    
    Debug.Print "=== Начало обработки документа ==="
    
    For Each pg In ActiveDocument.Pages
        Debug.Print "Страница " & pg.Index
        
        ' --- собираем кандидатов ---
        idx = -1
        For Each sh In pg.shapes
            If sh.Type = cdrTextShape Then
                If regEx.Test(sh.Text.Story.Text) Then
                    idx = idx + 1
                    ReDim Preserve items(2, idx)
                    items(0, idx) = sh.StaticID
                    items(1, idx) = sh.PositionY
                End If
            End If
        Next sh
        
        ' --- сортировка по Y (сверху вниз) ---
        If idx >= 1 Then Call SortItemsByY(items, idx)
        
        ' --- нумерация ---
        Dim k As Long
        For k = 0 To idx
            counter = counter + 1
            Set sh = pg.FindShape(StaticID:=items(0, k))
            txt = sh.Text.Story.Text
            
            newText = ReplaceIllTextIfNeeded(txt, counter)
            If newText <> txt Then
                sh.Text.Story.Text = newText
                fixedCount = fixedCount + 1
                Debug.Print " > Shape ID=" & sh.StaticID & " заменён на '" & newText & "'"
            Else
                alreadyOkCount = alreadyOkCount + 1
                Debug.Print " > Shape ID=" & sh.StaticID & " уже верный"
            End If
        Next k
    Next pg
    
    Debug.Print "=== Обработка завершена. Всего перенумеровано: " & counter & " объектов ==="
    Debug.Print " > Исправлено: " & fixedCount
    Debug.Print " > Уже верные: " & alreadyOkCount
    
    MsgBox "Нумерация проверена." & vbCrLf & _
           "Всего найдено: " & (counter - startN) & vbCrLf & _
           "Исправлено: " & fixedCount & vbCrLf & _
           "Уже верные: " & alreadyOkCount, vbInformation
End Sub

' --- проверка/замена номера ---
Private Function ReplaceIllTextIfNeeded(ByVal inputText As String, ByVal expectedNumber As Long) As String
    Dim re As Object, matches As Object
    Dim foundNumStr As String, currentNumber As Long
    
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = ILL_PATTERN
    re.IgnoreCase = True
    re.Global = False
    
    Set matches = re.Execute(inputText)
    If matches.Count = 0 Then
        ReplaceIllTextIfNeeded = inputText
        Exit Function
    End If
    
    If matches(0).SubMatches.Count > 0 Then
        foundNumStr = CStr(matches(0).SubMatches(0))
    Else
        foundNumStr = ExtractDigits(matches(0).Value)
    End If
    
    If foundNumStr = "" Then
        foundNumStr = 0
    End If
    
    If Len(foundNumStr) = 0 Or Not IsNumeric(foundNumStr) Then
        ReplaceIllTextIfNeeded = inputText
        Exit Function
    End If
    
    currentNumber = CLng(foundNumStr)
    
    If currentNumber = expectedNumber Then
        ReplaceIllTextIfNeeded = inputText
    Else
        ReplaceIllTextIfNeeded = re.Replace(inputText, "Илл. " & expectedNumber & ".")
    End If
End Function

' --- извлекаем первую группу цифр ---
Private Function ExtractDigits(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9]" Then
            out = out & ch
        ElseIf Len(out) > 0 Then
            Exit For
        End If
    Next i
    ExtractDigits = out
End Function

' --- пузырьковая сортировка массива Variant ---

Private Sub SortItemsByY(arr() As Variant, ByVal lastIndex As Long)
    Dim i As Long, j As Long
    Dim tmpID As Variant, tmpY As Variant
    
    For i = 0 To lastIndex - 1
        For j = i + 1 To lastIndex
            If arr(1, i) < arr(1, j) Then
                ' меняем местами
                tmpID = arr(0, i)
                tmpY = arr(1, i)
                
                arr(0, i) = arr(0, j)
                arr(1, i) = arr(1, j)
                
                arr(0, j) = tmpID
                arr(1, j) = tmpY
            End If
        Next j
    Next i
End Sub


Sub ChangeFont()

  Const TargetFont = "Times New Roman"
  Const TargetSize = 8.5
  Dim PageCounter As Integer
  PageCounter = 0
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Dim Page As Page
  Dim Shape As Shape
  For Each Page In ActiveDocument.Pages
    PageCounter = PageCounter + 1
    If (PageCounter > 1) Then
        For Each Shape In Page.shapes
            If (Shape.Type = cdrTextShape) Then
                Shape.Text.Story.Font = TargetFont
                Shape.Text.Story.Size = TargetSize
            End If
        Next Shape
    End If
  Next Page
  
  MsgBox "Text replacement font completed."
  
End Sub

' Перевод миллиметров в дюймы (единицы CorelDRAW API)
Function MMtoDocUnits(ByVal mm As Double) As Double
    MMtoDocUnits = mm / 25.4
End Function

' Перевод дюймов (API) в миллиметры
Function DocUnitsToMM(ByVal val As Double) As Double
    DocUnitsToMM = val * 25.4
End Function


Sub AddGuides_MM()
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "Нет открытого документа.", vbExclamation
        Exit Sub
    End If

    Dim pg As Page
    Set pg = doc.ActivePage

    ' Получаем размеры страницы в дюймах
    Dim wInch As Double, hInch As Double
    pg.GetSize wInch, hInch

    ' Для справки: размеры в мм
    Dim wMM As Double, hMM As Double
    wMM = DocUnitsToMM(wInch)
    hMM = DocUnitsToMM(hInch)
    Debug.Print "Размер страницы: "; wMM & " X " & hMM & " мм"

    ' Находим слой направляющих
    Dim gLayer As Layer
    On Error Resume Next
    Set gLayer = pg.Layers("Guides")
    On Error GoTo 0
    If gLayer Is Nothing Then Set gLayer = pg.CreateLayer("Guides")

    ' Список позиций (в мм)
    Dim posMM_V As Variant, posMM_G As Variant
    posMM_V = Array(50, 100, 150)
    posMM_G = Array(50, 100, 150)
    Dim i As Long, posInch As Double

    ' Начало координат - нижний левый угол
    
    ' Вертикальные
    For i = LBound(posMM_V) To UBound(posMM_V)
        posInch = MMtoDocUnits(posMM_V(i))
        gLayer.CreateGuide posInch, 0, posInch, hInch
    Next i

    ' Горизонтальные
    For i = LBound(posMM_G) To UBound(posMM_G)
        posInch = MMtoDocUnits(posMM_G(i))
        gLayer.CreateGuide 0, posInch, wInch, posInch
    Next i

    MsgBox "Направляющие добавлены (координаты заданы в мм).", vbInformation
End Sub

' --- Получение или создание слоя по имени ---
Function GetOrCreateLayer(ByVal pg As Page, ByVal layerName As String) As Layer
    Dim lyr As Layer
    On Error Resume Next
    Set lyr = pg.Layers(layerName)
    On Error GoTo 0
    
    If lyr Is Nothing Then
        Set lyr = pg.CreateLayer(layerName)
    End If
    
    Set GetOrCreateLayer = lyr
End Function

Function PlaceParagraphTextMM(pg As Page, ByVal txt As String, _
                              ByVal xMM As Double, ByVal yMM As Double, _
                              ByVal wMM As Double, ByVal hMM As Double) As Shape
    
    ' Переводим координаты в дюймы
    Dim x1Inch As Double, y1Inch As Double, x2Inch As Double, y2Inch As Double
    x1Inch = MMtoDocUnits(xMM)
    y1Inch = MMtoDocUnits(yMM)
    x2Inch = MMtoDocUnits(xMM + wMM)
    y2Inch = MMtoDocUnits(yMM + hMM)

    ' Слой для текста
    Dim lyr As Layer
    Set lyr = GetOrCreateLayer(pg, "Text")

    ' Создаём параграфный текст
    Set PlaceParagraphTextMM = lyr.CreateParagraphText(x1Inch, y1Inch, x2Inch, y2Inch, txt)
End Function


Sub Demo_AddArcheoImage()
    Dim pg As Page
    Set pg = ActiveDocument.ActivePage

    ' Вставляем схему раскопа шириной 170 мм внизу страницы
    Dim img As Shape
    Set img = PlaceImageMM(pg, "D:\tmp\sample.jpg", V_GUIDE_LEFT, H_GUIDE_BOTTOM)
    
    ' Подпись под изображением
    PlaceParagraphTextMM pg, _
        "Объект: Поселение у р. Вепря" & vbCrLf & "Работа: Фотофиксация, шурф 1", _
        V_GUIDE_LEFT, H_GUIDE_BOTTOM - 25, IMAGE_TARGET_WIDTH, 20
End Sub

Function PlaceImageMM(pg As Page, ByVal filePath As String, ByVal xMM As Double, ByVal yMM As Double) As Shape
    On Error GoTo ErrHandler

    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then Exit Function

    ' Получаем/создаём слой Images
    Dim lyr As Layer
    Set lyr = GetOrCreateLayer(pg, "Images")

    ' 1) Импорт (метод Import не возвращает объект — он просто импортирует и выделяет результат)
    lyr.Import filePath    ' <- без Set, просто вызываем процедуру. Импортированные формы будут выделены.

    ' 2) Берём выделение (SelectionRange) — это ShapeRange с импортированными формами
    Dim sr As ShapeRange
    Set sr = doc.SelectionRange   ' или ActiveDocument.SelectionRange — возвращает ShapeRange.

    If sr Is Nothing Or sr.Count = 0 Then
        MsgBox "Не удалось получить импортированный объект: " & filePath, vbExclamation
        Exit Function
    End If

    ' Возьмём первый объект из импортированного набора
    Dim s As Shape
    Set s = sr(1)

    ' 3) Масштабируем пропорционально: ширина = IMAGE_TARGET_WIDTH
    Dim targetW As Double, targetH As Double
    targetW = MMtoDocUnits(IMAGE_TARGET_WIDTH)
    targetH = s.SizeHeight * (targetW / s.SizeWidth)
    
    s.SetSize targetW, targetH


    ' 4) Устанавливаем позицию. Чтобы позиция интерпретировалась как левый-нижний угол, временно ставим ReferencePoint = cdrBottomLeft
    Dim prevRef As Long
    prevRef = doc.ReferencePoint
    doc.ReferencePoint = cdrBottomLeft

    s.SetPosition MMtoDocUnits(xMM), MMtoDocUnits(yMM)

    ' Восстанавливаем прежнюю опору
    doc.ReferencePoint = prevRef

    ' Возвращаем вставленный shape
    Set PlaceImageMM = s
    Exit Function

ErrHandler:
    MsgBox "Ошибка при вставке изображения: " & Err.Number & " — " & Err.Description, vbExclamation
End Function


' --- Добавление рамки (обрезного прямоугольника) ---
Sub AddPageBorder(pg As Page, ByVal marginMM As Double)
    Dim lyr As Layer
    Set lyr = pg.ActiveLayer
    
    Dim wInch As Double, hInch As Double
    pg.GetSize wInch, hInch
    
    Dim marginInch As Double
    marginInch = MMtoDocUnits(marginMM)
    
    ' Создаем прямоугольник с отступом от края
    lyr.CreateRectangle marginInch, marginInch, _
                        wInch - marginInch, hInch - marginInch
End Sub

' --- Выравнивание объекта по центру страницы ---
Sub AlignShapeToCenter(pg As Page, s As Shape)
    s.AlignToPage cdrAlignHCenter + cdrAlignVCenter
End Sub


Function InsertCharAt(ByVal txt As String, ByVal ch As String, ByVal pos As Long) As String
    ' Если позиция меньше 1 — вставляем в начало
    If pos < 1 Then pos = 1
    ' Если позиция больше длины строки + 1 — вставляем в конец
    If pos > Len(txt) + 1 Then pos = Len(txt) + 1
    
    InsertCharAt = Left(txt, pos - 1) & ch & Mid(txt, pos)
End Function


Function RemoveCharAt(ByVal txt As String, ByVal pos As Long) As String
    ' Проверяем, что позиция в пределах строки
    If pos < 1 Or pos > Len(txt) Then
        RemoveCharAt = txt
    Else
        RemoveCharAt = Left(txt, pos - 1) & Mid(txt, pos + 1)
    End If
End Function

Sub CreateNumberedTextObjects()
    Dim doc As Document
    Dim pg As Page
    Dim s As String
    Dim i As Long, n As Long
    Dim x As Double, y As Double

    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "Нет открытого документа!", vbExclamation
        Exit Sub
    End If
    
    Set pg = doc.ActivePage
    
    ' запрос количества объектов
    s = InputBox("Введите количество создаваемых объектов (n):", "Создание текстовых объектов", "10")
    If s = "" Then Exit Sub ' отмена
    If Not IsNumeric(s) Then
        MsgBox "Введите число!", vbExclamation
        Exit Sub
    End If
    n = CLng(s)
    If n <= 0 Then
        MsgBox "Количество должно быть больше 0!", vbExclamation
        Exit Sub
    End If
    
    x = 10   ' начальная координата X (мм)
    y = 10   ' начальная координата Y (мм)
    
    For i = 1 To n
         
         ' Пример: текстовый блок 60x20 мм внизу страницы
         PlaceParagraphTextMM pg, _
             "Шурф №" & i & vbCrLf & ". Археологическое исследование", _
             x, y, 60, 20
         
            y = y + 25 ' шаг по вертикали
    Next i
    
    MsgBox "Создано " & n & " объектов с числами от 1 до " & n
End Sub