Attribute VB_Name = "AlbumGeneratorWorkingInCurent2"
Option Explicit

' ================== Параметры ==================
Private Const V_GUIDE_LEFT As Double = 25      ' мм
Private Const V_GUIDE_RIGHT As Double = 195    ' мм (лево+170)
Private Const H_GUIDE_TOP As Double = 280       ' мм (верхняя направляющая)
Private Const H_GUIDE_BOTTOM As Double = 35   ' мм (нижняя направляющая)

Private Const IMAGE_TARGET_WIDTH As Double = 170   ' мм по ширине
Private Const CAPTION_FONT As String = "Times New Roman"
Private Const CAPTION_SIZE As Double = 11

' глобальные
Public g_IllNumber As Long
Public g_ObjectName As String
Public g_DocGuidesSupported As Boolean

' ================== Утилиты ==================
Function MMtoDocUnits(ByVal mm As Double) As Double
    MMtoDocUnits = mm / 25.4
End Function

Function PickFolderFromShell(ByVal prompt As String) As String
    Dim sh As Object, fldr As Object
    On Error Resume Next
    Set sh = CreateObject("Shell.Application")
    Set fldr = sh.BrowseForFolder(0, prompt, 0, 0)
    On Error GoTo 0
    If Not fldr Is Nothing Then PickFolderFromShell = fldr.Self.Path Else PickFolderFromShell = ""
End Function

Function GetOrCreateLayer(ByVal pg As Page, ByVal layerName As String) As Layer
    Dim lyr As Layer
    On Error Resume Next
    Set lyr = pg.Layers(layerName)
    On Error GoTo 0
    If lyr Is Nothing Then Set lyr = pg.CreateLayer(layerName)
    Set GetOrCreateLayer = lyr
End Function

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
    'If Left(out, 1) = "0" Then
    '  out = Right(out, 1)
    'End If
    ExtractDigits = out
End Function


' ================== Guides ==================
' Создаёт направляющие на странице
Sub AddGuidesToPage(ByVal pg As Page)
    Dim gLayer As Layer
    'Set gLayer = GetOrCreateLayer(pg, "Guides")
    
    Set gLayer = pg.GuidesLayer
       
    Dim wInch As Double, hInch As Double
    pg.GetSize wInch, hInch
    
    Dim xLeft As Double, xRight As Double
    Dim yTopInch As Double, yBottomInch As Double
    xLeft = MMtoDocUnits(V_GUIDE_LEFT)
    xRight = MMtoDocUnits(V_GUIDE_RIGHT)
    yTopInch = MMtoDocUnits(H_GUIDE_TOP)
    yBottomInch = MMtoDocUnits(H_GUIDE_BOTTOM)
    
    ' Создаём вертикальные направляющие
    gLayer.CreateGuide xLeft, 0, xLeft, hInch
    gLayer.CreateGuide xRight, 0, xRight, hInch
    
    ' Горизонтальные
    gLayer.CreateGuide 0, yTopInch, wInch, yTopInch
    gLayer.CreateGuide 0, yBottomInch, wInch, yBottomInch
 
End Sub


' ================== Новая страница ==================
Function EnsureNewPage(doc As Document) As Page
    doc.AddPages 1
    Set EnsureNewPage = doc.Pages.Last
    EnsureNewPage.Activate
End Function

' ================== Вставка изображения + подпись (устойчивая версия) ==================
Sub PlaceImageWithCaption(ByVal pg As Page, ByVal filePath As String, ByVal caption As String, ByVal isTop As Boolean)
    On Error GoTo ErrHandler
    Dim doc As Document: Set doc = ActiveDocument
    Dim imgLayer As Layer
    Set imgLayer = GetOrCreateLayer(pg, "Images")
    
    ' Попытка подавить всплывающие окна при импорте
    On Error Resume Next
    Application.Optimization = True
    Application.EventsEnabled = False
    On Error GoTo ErrHandler
    
    ' Import — процедура, поэтому НЕ используем Set
    imgLayer.Import filePath
    
    ' Восстановим события (но делаем это после того, как получили Selection)
    On Error Resume Next
    Application.Optimization = False
    Application.EventsEnabled = True
    On Error GoTo ErrHandler
    
    ' Берём импортированный объект через SelectionRange
    Dim sr As ShapeRange
    Set sr = doc.SelectionRange
    If sr Is Nothing Or sr.Count = 0 Then
        Err.Raise vbObjectError + 1000, , "Не удалось получить импортированный объект: " & filePath
    End If
    Dim s As Shape
    Set s = sr(1)
    
    ' Масштаб по ширине (IMAGE_TARGET_WIDTH мм)
    Dim targetW As Double, targetH As Double
    targetW = MMtoDocUnits(IMAGE_TARGET_WIDTH)
    targetH = s.SizeHeight * (targetW / s.SizeWidth)
    s.SetSize targetW, targetH
    
    ' Параметры направляющих (в единицах документа)
    Dim xLeftInch As Double, xRightInch As Double, yTopInch As Double, yBottomInch As Double
    xLeftInch = MMtoDocUnits(V_GUIDE_LEFT)
    xRightInch = MMtoDocUnits(V_GUIDE_RIGHT)
    yTopInch = MMtoDocUnits(H_GUIDE_TOP)
    yBottomInch = MMtoDocUnits(H_GUIDE_BOTTOM)
    
    Dim zoneW As Double
    zoneW = xRightInch - xLeftInch
    
    ' Вычисляем левый X (чтобы центрировать картинку между вертикальями)
    Dim leftX As Double
    leftX = xLeftInch + (zoneW - s.SizeWidth) / 2
    
    ' Позиционирование по Y: верхняя — верх = yTopInch; нижняя — низ = yBottomInch
    Dim bottomY As Double
    Dim prevRef As Long
    prevRef = doc.ReferencePoint
    doc.ReferencePoint = cdrBottomLeft
    If isTop Then
        bottomY = yTopInch - s.SizeHeight   ' верх = yTopInch
    Else
        bottomY = yBottomInch               ' низ = yBottomInch
    End If
    s.SetPosition leftX, bottomY
    doc.ReferencePoint = prevRef
    
    ' Подпись всегда ПОД картинкой (рамка чуть ниже bottomY)
    Dim captionTop As Double, captionBottom As Double
    captionTop = bottomY '- MMtoDocUnits(2)         ' 2 мм зазор
    captionBottom = captionTop - MMtoDocUnits(20)  ' высота подписи ~20 мм
    If captionBottom < 0 Then captionBottom = 0
    
    Dim txtLayer As Layer
    Set txtLayer = GetOrCreateLayer(pg, "Text")
    Dim fullCaption As String
    fullCaption = "Илл. " & g_IllNumber & ". Археологические разведки на земельном участке, отведенном для расположения объекта: «" & g_ObjectName & "». " & caption
    
    Dim txt As Shape
    Set txt = txtLayer.CreateParagraphText(xLeftInch, captionBottom, xRightInch, captionTop, fullCaption)
    
    ' Настройка шрифта — пробуем безопасно разные доступы
    On Error Resume Next
    txt.Text.Story.Font = CAPTION_FONT
    txt.Text.Story.Size = CAPTION_SIZE
    txt.Text.Font = CAPTION_FONT
    txt.Text.Size = CAPTION_SIZE
    Err.Clear
    On Error GoTo ErrHandler
    
    ' Очистим выделение (чтобы следующий import не спутал selection)
    On Error Resume Next
    doc.ClearSelection
    On Error GoTo ErrHandler
    
    g_IllNumber = g_IllNumber + 1
    Exit Sub

ErrHandler:
    ' постараемся восстановить события
    On Error Resume Next
    Application.Optimization = False
    Application.EventsEnabled = True
    MsgBox "Ошибка при вставке изображения: " & Err.Number & " — " & Err.Description, vbExclamation
End Sub


' ================== Сортировка файлов в папке по имени (asc, регистронезависимо) ==================
Function GetSortedFiles(ByVal folderPath As String) As Collection
    Dim fso As Object, fld As Object, fil As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(folderPath)
    
    Dim tmp() As Variant
    Dim cnt As Long: cnt = 0
    For Each fil In fld.files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(fil.Path))
        If ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "tif" Or ext = "tiff" Then
            ReDim Preserve tmp(0 To cnt)
            ' tmp: [0] = lowercased file name, [1] = full path
            tmp(cnt) = Array(LCase(fil.Name), fil.Path)
            cnt = cnt + 1
        End If
    Next fil
    
    Dim col As New Collection
    If cnt = 0 Then
        Set GetSortedFiles = col
        Exit Function
    End If
    
    ' сортировка вставкой по tmp(...)(0) — имени файла (asc)
    Dim i As Long, j As Long
    For i = 1 To UBound(tmp)
        Dim keyName As String, keyPath As String
        keyName = tmp(i)(0): keyPath = tmp(i)(1)
        j = i - 1
        Do While j >= 0
            If tmp(j)(0) > keyName Then
                tmp(j + 1) = tmp(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        tmp(j + 1) = Array(keyName, keyPath)
    Next i
    
    For i = 0 To UBound(tmp)
        col.Add tmp(i)(1)   ' добавляем полный путь, как и раньше
    Next i
    Set GetSortedFiles = col
End Function



' ================== Сортировка файлов в папке по дате создания (asc) ==================
Function GetSortedFiles_date(ByVal folderPath As String) As Collection
    Dim fso As Object, fld As Object, fil As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(folderPath)
    
    Dim tmp() As Variant
    Dim cnt As Long: cnt = 0
    For Each fil In fld.files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(fil.Path))
        If ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "tif" Or ext = "tiff" Then
            ReDim Preserve tmp(0 To cnt)
            tmp(cnt) = Array(fil.Path, fil.DateCreated)
            cnt = cnt + 1
        End If
    Next fil
    
    Dim col As New Collection
    If cnt = 0 Then
        Set GetSortedFiles = col
        Exit Function
    End If
    
    ' сортировка вставкой (asc)
    Dim i As Long, j As Long
    For i = 1 To UBound(tmp)
        Dim keyPath As String, keyDate As Date
        keyPath = tmp(i)(0): keyDate = tmp(i)(1)
        j = i - 1
        Do While j >= 0
        
            If tmp(j)(1) > keyDate Then
                tmp(j + 1) = tmp(j)
                j = j - 1
            Else
                Exit Do
            End If
        
        Loop
        tmp(j + 1) = Array(keyPath, keyDate)
    Next i
    
    For i = 0 To UBound(tmp)
        col.Add tmp(i)(0)
    Next i
    Set GetSortedFiles = col
End Function


' ================== Подписи для папки ==================
Function GetCaptionsForFolder(ByVal folderName As String, ByVal fileCount As Long) As Collection
    Dim caps As New Collection
    Dim baseName As String: baseName = folderName
    Dim num As String: num = ""
    'InStr(1, folderName, "Точка фотофиксации", vbTextCompare) > 0
    If InStr(1, folderName, "тфф", vbTextCompare) > 0 Then
        num = ExtractDigits(baseName)
        num = "Точка фотофиксации №" + num
        caps.Add num & ". Вид с Ю."
        caps.Add num & ". Вид с З."
        caps.Add num & ". Вид с С."
        caps.Add num & ". Вид с В."
    'InStr(1, folderName, "Шурф", vbTextCompare) > 0
    ElseIf InStr(1, folderName, "ш", vbTextCompare) > 0 Then
        num = ExtractDigits(baseName)
        If fileCount = 5 Then
            caps.Add "Разметка шурфа №" & num & ". Вид с Ю."
            caps.Add "Общий вид шурфа №" & num & ". Вид с Ю."
            caps.Add "Материк шурфа №" & num & ". Вид с Ю."
            caps.Add "Контрольный прокоп шурфа №" & num & ". Вид с Ю."
            caps.Add "Рекультивация шурфа №" & num & ". Вид с Ю."
        ElseIf fileCount = 4 Then
            caps.Add "Разметка шурфа №" & num & ". Вид с Ю."
            caps.Add "Материк шурфа №" & num & ". Вид с Ю."
            caps.Add "Контрольный прокоп шурфа №" & num & ". Вид с Ю."
            caps.Add "Рекультивация шурфа №" & num & ". Вид с Ю."
        End If
    Else
        Dim i As Integer
                
        For i = 1 To fileCount
            caps.Add baseName & " - файл " & CStr(i)
        Next
    End If
    Set GetCaptionsForFolder = caps
End Function

' ================== Основная процедура: BuildAlbum ==================
Sub BuildAlbum()
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "Откройте документ CorelDRAW перед запуском.", vbExclamation
        Exit Sub
    End If
    
    g_ObjectName = InputBox("Введите название объекта:", "Название объекта")
    If Len(Trim(g_ObjectName)) = 0 Then Exit Sub
    g_IllNumber = 1
    

    
    Dim rootPath As String
    rootPath = PickFolderFromShell("Выберите корневую папку с подпапками")
    If rootPath = "" Then Exit Sub
    
    Dim fso As Object, rootFld As Object, subFld As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFld = fso.GetFolder(rootPath)
    
    Dim pg As Page
    ' стартовая страница: создаём новую (чтобы макрос работал на пустом документе)
    Set pg = EnsureNewPage(doc)
    AddGuidesToPage pg
    
    ' формируем страницы с подписями до страниц с тф и шурмами
    Dim capsStart As Collection
    Dim i As Integer
    Dim xLeftInch As Double, xRightInch As Double, captionTop As Double, captionBottom As Double
    
    Set capsStart = New Collection

    capsStart.Add "Карта Пензенской области с обозначением участка исследования."
    capsStart.Add "Карта Пензенского района с обозначением участка исследования. Выкопировка из топосновы."
    capsStart.Add "Карта Пензенского района с обозначением участка исследования. Снимок со спутника."
    capsStart.Add "Карта памятников археологии в районе участка исследования."
    capsStart.Add "Обозначение участка исследования на «Старой карте 1»"
    capsStart.Add "Обозначение участка исследования на «Старой карте 2»"
    capsStart.Add "Обозначение участка исследования на «Старой карте 3»"
    capsStart.Add "Ситуационный план расположения шурфов и точек фотофиксации. Выкопировка из топосновы."
    capsStart.Add "Ситуационный план расположения шурфов и точек фотофиксации. Снимок со спутника."
    
    xLeftInch = MMtoDocUnits(V_GUIDE_LEFT)
    xRightInch = MMtoDocUnits(V_GUIDE_RIGHT)
    captionTop = MMtoDocUnits(H_GUIDE_BOTTOM)
    captionBottom = captionTop - MMtoDocUnits(20)  ' высота подписи ~20 мм
    If captionBottom < 0 Then captionBottom = 0
    For i = 1 To capsStart.Count

        Dim txtLayer As Layer
        Set txtLayer = GetOrCreateLayer(pg, "Text")
        Dim fullCaption As String
        fullCaption = "Илл. " & g_IllNumber & ". Археологические разведки на земельном участке, отведенном для расположения объекта: «" & g_ObjectName & "». " & capsStart(i)
        
        Dim txt As Shape
        Set txt = txtLayer.CreateParagraphText(xLeftInch, captionBottom, xRightInch, captionTop, fullCaption)
       
        On Error GoTo ErrHandler
        ' Настройка шрифта — пробуем безопасно разные доступы
        txt.Text.Story.Font = CAPTION_FONT
        txt.Text.Story.Size = CAPTION_SIZE
        
        ' Очистим выделение (чтобы следующий import не спутал selection)
        doc.ClearSelection
        
        g_IllNumber = g_IllNumber + 1
        Set pg = EnsureNewPage(doc)
        
        On Error GoTo 0      ' выключаем обработку ошибок
ContinueLoop:               ' метка должна быть внутри цикла, но до Next
    Next i
    
    Dim placeTop As Boolean: placeTop = True
    Dim files As Collection, captions As Collection
    For Each subFld In rootFld.SubFolders
        
        Set files = GetSortedFiles(subFld.Path)
        If files.Count <> 0 Then
            Set captions = GetCaptionsForFolder(subFld.Name, files.Count)
            If captions.Count < files.Count Then
                For i = captions.Count + 1 To files.Count
                    captions.Add subFld.Name & " - файл " & CStr(i)
                Next
            End If
            
            For i = 1 To files.Count
                PlaceImageWithCaption pg, files(i), captions(i), placeTop
                If placeTop Then
                    placeTop = False
                Else
                    Set pg = EnsureNewPage(doc)
                    placeTop = True
                End If
            Next
        End If
            
    Next subFld
    
    MsgBox "Альбом создан. Всего иллюстраций: " & (g_IllNumber - 1), vbInformation
    
    Exit Sub                ' <-- не забываем выйти, чтобы не провалиться в ErrHandler

' ---- обработчик ошибок ----
ErrHandler:
    Debug.Print "Ошибка на i=" & i & " (" & Err.Number & "): " & Err.Description
    Err.Clear
    Resume ContinueLoop     ' вернуться в цикл и продолжить

End Sub
