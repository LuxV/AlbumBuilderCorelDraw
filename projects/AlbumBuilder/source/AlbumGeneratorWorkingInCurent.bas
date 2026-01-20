Attribute VB_Name = "AlbumGeneratorWorkingInCurent"
Option Explicit

' ================== Настройки ==================
Private Const V_GUIDE_LEFT As Double = 25      ' мм
Private Const V_GUIDE_RIGHT As Double = 195    ' мм (лево+170)
Private Const H_GUIDE_TOP As Double = 280       ' мм (верхняя направляющая)
Private Const H_GUIDE_BOTTOM As Double = 35   ' мм (нижняя направляющая)

Private Const IMAGE_TARGET_WIDTH As Double = 170   ' мм по ширине
Private Const CAPTION_FONT As String = "Times New Roman"
Private Const CAPTION_SIZE As Double = 11

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
    Dim i As Long, ch As String
    Dim out As String
    
    out = ""
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        
        If ch Like "[0-9]" Then
            out = out & ch
        ElseIf Len(out) > 0 Then
            Exit For
        End If
    Next i
    
    ' Удаляем ведущие нули, если есть
    Do While Len(out) > 1 And Left$(out, 1) = "0"
        out = Mid$(out, 2)
    Loop
    
    ExtractDigits = out
End Function


Function DetectFolderStructure(ByVal rootPath As String) As String
    Dim fso As Object, rootFld As Object, fld As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFld = fso.GetFolder(rootPath)

    For Each fld In rootFld.SubFolders
        If InStr(1, LCase$(fld.Name), "кв", vbTextCompare) > 0 Then
            DetectFolderStructure = "KV"
            Exit Function
        End If
    Next

    DetectFolderStructure = "FLAT"
End Function
Function CollectFlatStructure(ByVal rootPath As String) As Collection
    Dim fso As Object, rootFld As Object, subFld As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFld = fso.GetFolder(rootPath)

    Dim col As New Collection

    For Each subFld In rootFld.SubFolders
        col.Add subFld
    Next

    Set CollectFlatStructure = col
End Function

' ================== Guides ==================
Sub AddGuidesToPage(ByVal pg As Page)
    Dim gLayer As Layer
    Set gLayer = pg.GuidesLayer
    Dim wInch As Double, hInch As Double
    pg.GetSize wInch, hInch
    Dim xLeft As Double, xRight As Double
    Dim yTopInch As Double, yBottomInch As Double
    xLeft = MMtoDocUnits(V_GUIDE_LEFT)
    xRight = MMtoDocUnits(V_GUIDE_RIGHT)
    yTopInch = MMtoDocUnits(H_GUIDE_TOP)
    yBottomInch = MMtoDocUnits(H_GUIDE_BOTTOM)
    On Error Resume Next
    gLayer.CreateGuide xLeft, 0, xLeft, hInch
    gLayer.CreateGuide xRight, 0, xRight, hInch
    gLayer.CreateGuide 0, yTopInch, wInch, yTopInch
    gLayer.CreateGuide 0, yBottomInch, wInch, yBottomInch
    On Error GoTo 0
End Sub

' ================== Новая страница ==================
Function EnsureNewPage(doc As Document) As Page
    doc.AddPages 1
    Set EnsureNewPage = doc.Pages.Last
    EnsureNewPage.Activate
End Function

' ================== Вставка изображения + подпись ==================
Sub PlaceImageWithCaption(  ByVal pg As Page,_ 
                            ByVal filePath As String,_
                            ByVal caption As String,_
                            ByVal isTop As Boolean,_
                            ByRef illNumber As Long,_
                            ByVal objectName As String)
    On Error GoTo ErrHandler
    Dim doc As Document: Set doc = ActiveDocument
    Dim imgLayer As Layer
    Set imgLayer = GetOrCreateLayer(pg, "Images")

    ' Попытка подавить всплывающие окна при импорте
    On Error Resume Next
    Application.Optimization = True
    Application.EventsEnabled = False
    On Error GoTo ErrHandler

    imgLayer.Import filePath

    On Error Resume Next
    Application.Optimization = False
    Application.EventsEnabled = True
    On Error GoTo ErrHandler

    Dim sr As ShapeRange
    Set sr = doc.SelectionRange
    If sr Is Nothing Or sr.Count = 0 Then
        Err.Raise vbObjectError + 1000, , "Не удалось получить импортированный объект: " & filePath
    End If
    Dim s As Shape
    Set s = sr(1)

    Dim targetW As Double, targetH As Double
    targetW = MMtoDocUnits(IMAGE_TARGET_WIDTH)
    targetH = s.SizeHeight * (targetW / s.SizeWidth)
    s.SetSize targetW, targetH

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
    captionTop = bottomY
    captionBottom = captionTop - MMtoDocUnits(20)
    If captionBottom < 0 Then captionBottom = 0

    Dim txtLayer As Layer
    Set txtLayer = GetOrCreateLayer(pg, "Text")
    Dim fullCaption As String
    fullCaption = "Илл. " & illNumber & ". Археологические разведки на земельном участке, отведенном для расположения объекта: «" & objectName & "». " & caption

    Dim txt As Shape
    Set txt = txtLayer.CreateParagraphText(xLeftInch, captionBottom, xRightInch, captionTop, fullCaption)

    On Error Resume Next
    txt.Text.Story.Font = CAPTION_FONT
    txt.Text.Story.Size = CAPTION_SIZE
    txt.Text.Font = CAPTION_FONT
    txt.Text.Size = CAPTION_SIZE
    Err.Clear
    On Error GoTo ErrHandler

    On Error Resume Next
    doc.ClearSelection
    On Error GoTo ErrHandler

    illNumber = illNumber + 1
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.Optimization = False
    Application.EventsEnabled = True
    MsgBox "Ошибка при вставке изображения: " & Err.Number & " — " & Err.Description, vbExclamation
End Sub

Private Function DirOrder(ByVal fileName As String) As Long
    Dim ch As String
    ch = LCase(Left(fileName, 1))
    
    Select Case ch
        Case "ю": DirOrder = 1
        Case "з": DirOrder = 2
        Case "с": DirOrder = 3
        Case "в": DirOrder = 4
        Case Else: DirOrder = 99   ' всё остальное — в конец
    End Select
End Function

' ================== Сортировка файлов — упрощённая и безопасная ==================
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
            ' [0] = order, [1] = имя, [2] = путь
            tmp(cnt) = Array( _
                DirOrder(fil.Name), _
                LCase(fil.Name), _
                fil.Path _
            )
            cnt = cnt + 1
        End If
    Next fil
    
    Dim col As New Collection
    If cnt = 0 Then
        Set GetSortedFiles = col
        Exit Function
    End If
    
    ' сортировка вставкой
    Dim i As Long, j As Long
    For i = 1 To UBound(tmp)
        Dim key
        key = tmp(i)
        j = i - 1
        
        Do While j >= 0
            If tmp(j)(0) > key(0) _
               Or (tmp(j)(0) = key(0) And tmp(j)(1) > key(1)) Then
                tmp(j + 1) = tmp(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        tmp(j + 1) = key
    Next i
    
    For i = 0 To UBound(tmp)
        col.Add tmp(i)(2)
    Next i
    
    Set GetSortedFiles = col
End Function


' ================== Новая логика формирования подписей ==================
' Сохраняет старую логику под именем GetCaptionsForFolder_Legacy
Function GetCaptionsForFolder_Legacy(ByVal folderName As String, ByVal fileCount As Long) As Collection
    Dim caps As New Collection
    Dim baseName As String: baseName = folderName
    Dim num As String: num = ""
    If InStr(1, folderName, "тфф", vbTextCompare) > 0 Then
        num = ExtractDigits(baseName)
        num = "Точка фотофиксации №" + num
        caps.Add num & ". Вид с Ю."
        caps.Add num & ". Вид с З."
        caps.Add num & ". Вид с С."
        caps.Add num & ". Вид с В."
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
    Set GetCaptionsForFolder_Legacy = caps
End Function

' Новая, расширяемая логика — разбирает имя папки на токены и генерирует подписи по шаблонам
Function GetCaptionsForFolder_Enhanced(ByVal folderName As String, ByVal files As Collection, ByVal kvNumber As String) As Collection
    Dim caps As New Collection
    Dim norm As String
    norm = LCase$(folderName)

    Dim isPlast As Boolean: isPlast = (InStr(1, norm, "пласт", vbTextCompare) > 0)
    Dim isMaterik As Boolean: isMaterik = (InStr(1, norm, "материк", vbTextCompare) > 0)
    Dim isProfil As Boolean: isProfil = (InStr(1, norm, "профил", vbTextCompare) > 0) Or (InStr(1, norm, "профиля", vbTextCompare) > 0)

    Dim plastNum As String
    plastNum = ExtractDigits(folderName)

    Dim i As Long
    For i = 1 To files.Count
        Dim fPath As String: fPath = files(i)
        Dim fName As String
        fName = Mid$(fPath, InStrRev(fPath, "\") + 1)
        Dim p As Long: p = InStrRev(fName, ".")
        If p > 0 Then fName = Left$(fName, p - 1)
        fName = Trim$(LCase$(fName))

        Dim dirToken As String: dirToken = "Х"
        If Len(fName) > 0 Then
            Dim ch As String: ch = Left$(fName, 1)
            Select Case ch
                Case "с", "c"
                    dirToken = "С"
                Case "ю", "y"
                    dirToken = "Ю"
                Case "з", "z"
                    dirToken = "З"
                Case "в", "v"
                    dirToken = "В"
                Case Else
                    dirToken = "Х"
            End Select
        End If

        Dim profileDir As String: profileDir = "УТОЧНИТЬ"
        Select Case LCase$(Left$(fName, 1))
            Case "ю"
                profileDir = "Северный"
            Case "с"
                profileDir = "Южный"
            Case "в"
                profileDir = "Западный"
            Case "з"
                profileDir = "Восточный"
            Case Else
                profileDir = "УТОЧНИТЬ"
        End Select

        Dim caption As String
        If isPlast Then
            Dim plastLabel As String
            If plastNum <> "" Then plastLabel = plastNum Else plastLabel = "#"
            caption = "Пласт " & plastLabel & ", кв. " & kvNumber & ". Вид с " & dirToken & "."
        ElseIf isMaterik Then
            caption = "Материк, кв. " & kvNumber & ". Вид с " & dirToken & "."
        ElseIf isProfil Then
            caption = profileDir & " профиль, кв. " & kvNumber & ". Вид с " & dirToken & "."
        Else
            caption = folderName & " - файл " & CStr(i)
        End If

        caps.Add caption
    Next i

    Set GetCaptionsForFolder_Enhanced = caps
End Function


' ================== Новая функция: сбор по структуре "кв -> (пласты, материк, профиля)" ==================
Function CollectFoldersForKvStructure(ByVal rootPath As String) As Collection
    Dim fso As Object, rootFld As Object, kvFld As Object, subFld As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFld = fso.GetFolder(rootPath)
    Dim col As New Collection

    For Each kvFld In rootFld.SubFolders
        If InStr(1, LCase$(kvFld.Name), "кв", vbTextCompare) > 0 Then
            Dim kvNum As String: kvNum = ExtractDigits(kvFld.Name)
            ' 1) Пласты — все подпапки содержащие "пласт", отсортированные по номеру
            Dim plastArr() As Variant
            Dim pCnt As Long: pCnt = 0
            For Each subFld In kvFld.SubFolders
                If InStr(1, LCase$(subFld.Name), "пласт", vbTextCompare) > 0 Then
                    ReDim Preserve plastArr(0 To pCnt)
                    plastArr(pCnt) = Array(subFld.Path, ExtractDigits(subFld.Name))
                    pCnt = pCnt + 1
                End If
            Next subFld
            If pCnt > 1 Then
                Dim i As Long, j As Long, tmp As Variant
                For i = 0 To pCnt - 2
                    For j = i + 1 To pCnt - 1
                        Dim ni As Long, nj As Long
                        ni = Val(plastArr(i)(1)): nj = Val(plastArr(j)(1))
                        If ni > nj Then
                            tmp = plastArr(i): plastArr(i) = plastArr(j): plastArr(j) = tmp
                        End If
                    Next j
                Next i
            End If
            For i = 0 To pCnt - 1
                col.Add Array(plastArr(i)(0), kvNum)
            Next i

            ' 2) Материк — добавляем все подпапки содержащие "материк" (или "матер")
            For Each subFld In kvFld.SubFolders
                If InStr(1, LCase$(subFld.Name), "матер", vbTextCompare) > 0 Then
                    col.Add Array(subFld.Path, kvNum)
                End If
            Next subFld

            ' 3) Профиля — все подпапки содержащие "профил" или "проф" (в порядке обнаружения)
            For Each subFld In kvFld.SubFolders
                If InStr(1, LCase$(subFld.Name), "профил", vbTextCompare) > 0 Or InStr(1, LCase$(subFld.Name), "проф", vbTextCompare) > 0 Then
                    col.Add Array(subFld.Path, kvNum)
                End If
            Next subFld
        End If
    Next kvFld

    Set CollectFoldersForKvStructure = col
End Function

' ================== Основная процедура: BuildAlbum ==================
Public Sub BuildAlbum(  ByVal rootPath As String, _
                        ByVal objectName As String, _
                        ByVal onlyPhotos As Boolean, _
                        ByVal startIndex As Integer)
    
    If Dir(rootPath, vbDirectory) = "" Then
        MsgBox "Указанная папка не существует."
        Exit Sub
    End If

    Dim doc As Document
    Set doc = ActiveDocument
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If doc Is Nothing Then
        MsgBox "Откройте документ CorelDRAW перед запуском.", vbExclamation
        Exit Sub
    End If

    Dim folders As Collection
    Dim structureType As String
    
    structureType = DetectFolderStructure(rootPath)
    
    If structureType = "KV" Then
        Set folders = CollectFoldersForKvStructure(rootPath)
    Else
        Set folders = CollectFlatStructure(rootPath)
    End If

    Dim ui As New frmProgress
    ui.Init folders.Count
    Dim folderIndex As Integer: folderIndex = 0

    Dim pg As Page
    Set pg = EnsureNewPage(doc)
    AddGuidesToPage pg
    
    If Not onlyPhotos Then

        ' Вставляем стартовые подписи
        Dim capsStart As Collection: Set capsStart = New Collection
        capsStart.Add "Карта Пензенской области с обозначением участка исследования."
        capsStart.Add "Карта Пензенского района с обозначением участка исследования. Выкопировка из топосновы."
        capsStart.Add "Карта Пензенского района с обозначением участка исследования. Снимок со спутника."
        capsStart.Add "Карта памятников археологии в районе участка исследования."
        capsStart.Add "Обозначение участка исследования на «Старой карте 1»"
        capsStart.Add "Обозначение участка исследования на «Старой карте 2»"
        capsStart.Add "Обозначение участка исследования на «Старой карте 3»"
        capsStart.Add "Ситуационный план расположения шурфов и точек фотофиксации. Выкопировка из топосновы."
        capsStart.Add "Ситуационный план расположения шурфов и точек фотофиксации. Снимок со спутника."

        Dim xLeftInch As Double, xRightInch As Double, captionTop As Double, captionBottom As Double
        xLeftInch = MMtoDocUnits(V_GUIDE_LEFT)
        xRightInch = MMtoDocUnits(V_GUIDE_RIGHT)
        captionTop = MMtoDocUnits(H_GUIDE_BOTTOM)
        captionBottom = captionTop - MMtoDocUnits(20)
        If captionBottom < 0 Then captionBottom = 0

        Dim i As Integer
        For i = 1 To capsStart.Count
            Dim txtLayer As Layer
            Set txtLayer = GetOrCreateLayer(pg, "Text")
            Dim fullCaption As String
            fullCaption = "Илл. " & startIndex & ". Археологические разведки на земельном участке, отведенном для расположения объекта: «" & ObjectName & "». " & capsStart(i)
            Dim txt As Shape
            Set txt = txtLayer.CreateParagraphText(xLeftInch, captionBottom, xRightInch, captionTop, fullCaption)
            On Error Resume Next
            txt.Text.Story.Font = CAPTION_FONT
            txt.Text.Story.Size = CAPTION_SIZE
            On Error GoTo 0
            doc.ClearSelection
            startIndex = startIndex + 1
            Set pg = EnsureNewPage(doc)
        Next i
    End If

    Dim placeTop As Boolean: placeTop = True
    Dim files As Collection, captions As Collection
    Dim fld As Object

    Dim fldInfo As Variant
    Dim fldPath As String
    Dim kvNum As String
    Dim folderDisplayName As String
  
    For Each fldInfo In folders
    
        If ui.Cancelled Then Exit For
    
        ' --- определить тип элемента: Folder object, Array(path, kv) или просто строка ---
        If IsObject(fldInfo) Then
            ' старый режим: элемент — объект Folder
            fldPath = fldInfo.Path
            folderDisplayName = fldInfo.Name
            kvNum = ExtractDigits(folderDisplayName)    ' если есть номер квадрата в имени — возьмём
        ElseIf IsArray(fldInfo) Then
            ' элемент — Array(path, kvNumber)
            fldPath = fldInfo(0)
            kvNum = fldInfo(1)
            On Error Resume Next
            folderDisplayName = fso.GetFolder(fldPath).Name
            If Err.Number <> 0 Then folderDisplayName = fldPath
            On Error GoTo 0
        Else
            ' на всякий случай: строка с путём
            fldPath = CStr(fldInfo)
            On Error Resume Next
            folderDisplayName = fso.GetFolder(fldPath).Name
            If Err.Number <> 0 Then folderDisplayName = fldPath
            On Error GoTo 0
            kvNum = ExtractDigits(folderDisplayName)
        End If
    
        folderIndex = folderIndex + 1
        ui.UpdateFolder folderIndex, folders.Count
        ui.SetStatus "Папка: " & folderDisplayName
    
        ' --- получить файлы в папке ---
        Set files = GetSortedFiles(fldPath)
        If files.Count = 0 Then
            ' нет изображений — пропустить
            GoTo NextFolder
        End If
    
        ' --- выбрать функцию формирования подписей ---
        If structureType = "KV" Then
            Set captions = GetCaptionsForFolder_Enhanced(folderDisplayName, files, kvNum)
        Else
            Set captions = GetCaptionsForFolder_Legacy(folderDisplayName, files.Count)
        End If
    
        ' --- если подписей меньше, чем файлов — дополним по умолчанию ---
        If captions.Count < files.Count Then
            For i = captions.Count + 1 To files.Count
                captions.Add folderDisplayName & " - файл " & CStr(i)
            Next i
        End If
    
        ui.UpdateFiles 0, files.Count
        Dim fileIndex As Integer: fileIndex = 0
        
            
        ' --- вставка файлов с подписями (максимум files.Count) ---
        For i = 1 To files.Count
        
            If ui.Cancelled Then Exit For
        
            PlaceImageWithCaption pg, files(i), captions(i), placeTop startIndex objectName
            
            fileIndex = fileIndex + 1
            ui.UpdateFiles fileIndex, files.Count
            
            If placeTop Then
                placeTop = False
            Else
                Set pg = EnsureNewPage(doc)
                placeTop = True
            End If
        Next i
    
NextFolder:
    Next fldInfo


    MsgBox "Альбом создан. Всего иллюстраций: " & (startIndex - 1), vbInformation
    If ui.Cancelled Then
        MsgBox "Операция отменена пользователем."
    End If
    Unload ui
End Sub

Public Sub StartAlbumBuilder()
    frmGenAlbum.Show
End Sub
