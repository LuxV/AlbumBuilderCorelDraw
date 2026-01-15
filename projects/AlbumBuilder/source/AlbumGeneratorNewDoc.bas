Attribute VB_Name = "AlbumGeneratorNewDoc"
' ================== Константы ==================
Private Const V_GUIDE_LEFT As Double = 25
Private Const V_GUIDE_RIGHT As Double = 195
Private Const H_GUIDE_TOP As Double = 280
Private Const H_GUIDE_BOTTOM As Double = 40
Private Const IMAGE_TARGET_WIDTH As Double = 170

Private Const CAPTION_FONT As String = "Times New Roman"
Private Const CAPTION_SIZE As Double = 11

' ================== Глобальные переменные ==================
Public g_IllNumber As Long
Public g_ObjectName As String

' ================== Вспомогательные функции ==================
Function MMtoDocUnits(mm As Double) As Double
    MMtoDocUnits = mm / 25.4
End Function

Function DocUnitsToMM(units As Double) As Double
    DocUnitsToMM = units * 25.4
End Function

Function GetOrCreateLayer(pg As Page, layerName As String) As Layer
    Dim lyr As Layer
    On Error Resume Next
    Set lyr = pg.Layers(layerName)
    On Error GoTo 0
    If lyr Is Nothing Then
        Set lyr = pg.CreateLayer(layerName)
    End If
    Set GetOrCreateLayer = lyr
End Function

' ================== Добавление направляющих ==================
Sub AddGuidesToPage(ByVal pg As Page)
    Dim gLayer As Layer
    Set gLayer = GetOrCreateLayer(pg, "Guides")
    
    Dim wInch As Double, hInch As Double
    pg.GetSize wInch, hInch
    
    Dim xLeft As Double, xRight As Double
    Dim yTopInch As Double, yBottomInch As Double
    xLeft = MMtoDocUnits(V_GUIDE_LEFT)
    xRight = MMtoDocUnits(V_GUIDE_RIGHT)
    yTopInch = MMtoDocUnits(H_GUIDE_TOP)
    yBottomInch = MMtoDocUnits(H_GUIDE_BOTTOM)
    
    ' Вертикальные
    gLayer.CreateGuide xLeft, 0, xLeft, hInch
    gLayer.CreateGuide xRight, 0, xRight, hInch
    
    ' Горизонтальные
    gLayer.CreateGuide 0, yTopInch, wInch, yTopInch
    gLayer.CreateGuide 0, yBottomInch, wInch, yBottomInch
End Sub

' ================== Вставка изображения + подпись ==================
Sub PlaceImageWithCaption(ByVal pg As Page, ByVal filePath As String, ByVal caption As String, ByVal isTop As Boolean)
    On Error GoTo ErrHandler
    Dim doc As Document: Set doc = ActiveDocument
    Dim imgLayer As Layer
    Set imgLayer = GetOrCreateLayer(pg, "Images")
    
    Application.Optimization = True
    Application.EventsEnabled = False
    
    ' Импорт через ImportEx
    Dim s As Shape
    Set s = imgLayer.ImportEx(filePath, cdrImportFull, cdrAutoSense)
    
    Application.Optimization = False
    Application.EventsEnabled = True
    
    If s Is Nothing Then
        Err.Raise vbObjectError + 1000, , "Не удалось импортировать объект: " & filePath
    End If
    'Dim s As Shape
    'Set s = sr(1)
    
    ' Масштабирование по ширине
    Dim targetW As Double, targetH As Double
    targetW = MMtoDocUnits(IMAGE_TARGET_WIDTH)
    targetH = s.SizeHeight * (targetW / s.SizeWidth)
    s.SetSize targetW, targetH
    
    ' Вычисление позиции
    Dim xLeftInch As Double, xRightInch As Double, yTopInch As Double, yBottomInch As Double
    xLeftInch = MMtoDocUnits(V_GUIDE_LEFT)
    xRightInch = MMtoDocUnits(V_GUIDE_RIGHT)
    yTopInch = MMtoDocUnits(H_GUIDE_TOP)
    yBottomInch = MMtoDocUnits(H_GUIDE_BOTTOM)
    
    Dim zoneW As Double
    zoneW = xRightInch - xLeftInch
    
    Dim leftX As Double
    leftX = xLeftInch + (zoneW - s.SizeWidth) / 2
    
    Dim bottomY As Double
    Dim prevRef As Long
    prevRef = doc.ReferencePoint
    doc.ReferencePoint = cdrBottomLeft
    If isTop Then
        bottomY = yTopInch - s.SizeHeight
    Else
        bottomY = yBottomInch
    End If
    s.SetPosition leftX, bottomY
    doc.ReferencePoint = prevRef
    
    ' Подпись
    Dim captionTop As Double, captionBottom As Double
    captionTop = bottomY - MMtoDocUnits(2)
    captionBottom = captionTop - MMtoDocUnits(18)
    If captionBottom < 0 Then captionBottom = 0
    
    Dim txtLayer As Layer
    Set txtLayer = GetOrCreateLayer(pg, "Text")
    Dim fullCaption As String
    fullCaption = "Илл. №" & g_IllNumber & ". Археологические разведки на земельном участке, отведенном для расположения объекта: «" & g_ObjectName & "»." & vbCrLf & caption
    
    Dim txt As Shape
    Set txt = txtLayer.CreateParagraphText(xLeftInch, captionBottom, xRightInch, captionTop, fullCaption)
    txt.Text.Story.Font = CAPTION_FONT
    txt.Text.Story.Size = CAPTION_SIZE
    
    doc.ClearSelection
    
    g_IllNumber = g_IllNumber + 1
    Exit Sub

ErrHandler:
    Application.Optimization = False
    Application.EventsEnabled = True
    MsgBox "Ошибка при вставке изображения: " & Err.number & " — " & Err.Description, vbExclamation
End Sub

' ================== Формирование подписей ==================
Function GetCaptionsForFolder(folderName As String, fileCount As Long) As Collection
    Dim captions As New Collection
    Dim baseName As String
    baseName = folderName
    
    If InStr(1, folderName, "Точка фотофиксации", vbTextCompare) > 0 Or InStr(1, folderName, "тфф", vbTextCompare) > 0  Then
        captions.Add baseName & ". Вид с Ю."
        captions.Add baseName & ". Вид с З."
        captions.Add baseName & ". Вид с С."
        captions.Add baseName & ". Вид с В."
    ElseIf InStr(1, folderName, "Шурф", vbTextCompare) > 0 Or InStr(1, folderName, "ш", vbTextCompare) > 0 Then
        Dim n As String
        n = Replace(folderName, "Шурф ", "")
        If fileCount = 5 Then
            captions.Add "Разметка шурфа №" & n & ". Вид с Ю."
            captions.Add "Общий вид шурфа №" & n & ". Вид с Ю."
            captions.Add "Материк шурфа №" & n & ". Вид с Ю."
            captions.Add "Контрольный прокоп шурфа №" & n & ". Вид с Ю."
            captions.Add "Рекультивация шурфа №" & n & ". Вид с Ю."
        Else
            captions.Add "Разметка шурфа №" & n & ". Вид с Ю."
            captions.Add "Материк шурфа №" & n & ". Вид с Ю."
            captions.Add "Контрольный прокоп шурфа №" & n & ". Вид с Ю."
            captions.Add "Рекультивация шурфа №" & n & ". Вид с Ю."
        End If
    End If
    
    Set GetCaptionsForFolder = captions
End Function

' ================== Основной макрос ==================
Sub ImportArchaeologyFolder()
    Dim rootPath As String
    Dim sh As Object, folder As Object
    Dim fso As Object, rootFolder As Object, subFolder As Object, fileObj As Object
    Dim doc As Document, pg As Page
    Dim files As New Collection, captions As Collection
    Dim i As Long
    
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "Нет открытого документа.", vbExclamation
        Exit Sub
    End If
    
    g_ObjectName = InputBox("Введите название объекта:", "Название объекта")
    If Trim(g_ObjectName) = "" Then
        MsgBox "Название объекта не указано. Работа прервана.", vbExclamation
        Exit Sub
    End If
    
    ' Выбор папки через Shell
    Set sh = CreateObject("Shell.Application")
    Set folder = sh.BrowseForFolder(0, "Выберите корневую папку", 0, 0)
    If folder Is Nothing Then Exit Sub
    rootPath = folder.items().Item().Path
    
    ' Работа с FSO
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(rootPath)
    
    g_IllNumber = 1
    
    ' Перебираем подпапки
    For Each subFolder In rootFolder.SubFolders
        Set files = New Collection
        
        For Each fileObj In subFolder.files
            If LCase(fso.GetExtensionName(fileObj.Name)) = "jpg" Or LCase(fso.GetExtensionName(fileObj.Name)) = "jpeg" Then
                files.Add fileObj
            End If
        Next fileObj
        
        If files.Count = 0 Then GoTo NextFolder
        
        Call SortFilesByDateColl(files)
        Set captions = GetCaptionsForFolder(subFolder.Name, files.Count)
        
        For i = 1 To files.Count
            If (g_IllNumber = 1) Or (g_IllNumber Mod 2 = 1) Then
                doc.AddPages 1
                Set pg = doc.Pages.Last
                AddGuidesToPage pg
            End If
            
            Dim isTop As Boolean
            isTop = ((g_IllNumber Mod 2) = 1)
            
            If i <= captions.Count Then
                PlaceImageWithCaption pg, files(i).Path, captions(i), isTop
            Else
                PlaceImageWithCaption pg, files(i).Path, subFolder.Name, isTop
            End If
        Next i
        
NextFolder:
    Next subFolder
    
    MsgBox "Импорт завершён.", vbInformation
End Sub

' ================== Сортировка файлов по дате ==================
Sub SortFilesByDateColl(ByRef files As Collection)
    Dim i As Long, j As Long
    Dim tmp As Object
    For i = 1 To files.Count - 1
        For j = i + 1 To files.Count
            If files(i).DateCreated > files(j).DateCreated Then
                Set tmp = files(i)
                files.Remove i
                files.Add tmp, , j - 1
            End If
        Next j
    Next i
End Sub

