Attribute VB_Name = "modCorelText"
Option Explicit

Private Const V_GUIDE_LEFT As Double = 25
Private Const V_GUIDE_RIGHT As Double = 195
Private Const H_GUIDE_TOP As Double = 280
Private Const H_GUIDE_MID As Double = 167
Private Const H_GUIDE_BOTTOM As Double = 35


Private Const TEXT_WIDTH As Double = 170
Private Const CAPTION_FONT As String = "Times New Roman"
Private Const CAPTION_SIZE As Double = 11

Public gTextCursorYmm As Double
Public gTextCursorXmm As Double

Function MMtoDocUnits(ByVal mm As Double) As Double
    MMtoDocUnits = mm / 25.4
End Function

Function DocUnitsToMM(ByVal val As Double) As Double
    DocUnitsToMM = val * 25.4
End Function


Private Sub CreateVGuide(ByVal gLayer As Layer, ByVal x As Double, ByVal pageHeight As Double)
    gLayer.CreateGuide x, 0, x, pageHeight
End Sub

Private Sub CreateHGuide(ByVal gLayer As Layer, ByVal y As Double, ByVal pageWidth As Double)
    gLayer.CreateGuide 0, y, pageWidth, y
End Sub


Sub AddGuidesToPage(ByVal pg As Page)
    
    Dim gLayer As Layer
    Set gLayer = pg.GuidesLayer

    Dim w As Double, h As Double
    pg.GetSize w, h

    CreateVGuide gLayer, MMtoDocUnits(V_GUIDE_LEFT), h
    CreateVGuide gLayer, MMtoDocUnits(V_GUIDE_RIGHT), h

    CreateHGuide gLayer, MMtoDocUnits(H_GUIDE_TOP), w
    CreateHGuide gLayer, MMtoDocUnits(H_GUIDE_MID), w
    CreateHGuide gLayer, MMtoDocUnits(H_GUIDE_BOTTOM), w

End Sub


Function EnsureNewPage(doc As Document, name As String) As Page
    
    doc.AddPages 1
    Set EnsureNewPage = doc.Pages.Last
    EnsureNewPage.Activate
    If name <> "" Then
        EnsureNewPage.name = "name"
    End If
         
End Function


Function GetOrCreateLayer(pg As Page, layerName As String) As Layer
    Dim lyr As Layer
    For Each lyr In pg.Layers
        If lyr.name = layerName Then
            Set GetOrCreateLayer = lyr
            Exit Function
        End If
    Next

    Set GetOrCreateLayer = pg.CreateLayer(layerName)
End Function


Function PlaceParagraphTextMM(pg As Page, ByVal txt As String, ByVal IsParagraph As Boolean, _
                              ByVal xMM As Double, ByVal yMM As Double, _
                              ByVal wMM As Double, ByVal hMM As Double) As Shape

    ' Переводим координаты в дюймы
    Dim x1Inch As Double, y1Inch As Double, x2Inch As Double, y2Inch As Double
    x1Inch = MMtoDocUnits(xMM)
    y1Inch = MMtoDocUnits(yMM)
    x2Inch = MMtoDocUnits(xMM + wMM)
    y2Inch = MMtoDocUnits(yMM + hMM)

    Dim lyr As Layer
    Set lyr = GetOrCreateLayer(pg, "Text")

    If IsParagraph Then
        Set PlaceParagraphTextMM = lyr.CreateParagraphText( _
            x1Inch, y1Inch, x2Inch, y2Inch, txt, , , CAPTION_FONT, CAPTION_SIZE, , , , cdrFullJustifyAlignment)
    Else
        Set PlaceParagraphTextMM = lyr.CreateArtisticText( _
            x1Inch, y1Inch, txt, , , CAPTION_FONT, CAPTION_SIZE)
    End If
  
End Function


Sub CreateParagraphText(s As String, NumberBlock As Integer)

    Dim CursorYmm As Double
        
    If NumberBlock Mod 2 = 0 Then
        CursorYmm = H_GUIDE_BOTTOM
    Else
        CursorYmm = H_GUIDE_MID
    End If

    PlaceParagraphTextMM ActivePage, s, True, V_GUIDE_LEFT, CursorYmm - 20, TEXT_WIDTH, 20

End Sub


Sub CreateNumberBlocks(blocks As Collection, NumberBlock As Integer)

    Dim CursorYmm As Double
        
    If NumberBlock Mod 2 = 0 Then
        CursorYmm = 5
    Else
        CursorYmm = 285
    End If


    Dim nb As clsNumberBlock
    Dim prop As Variant

    For Each nb In blocks
    
        ' создаём основной текст числа
        PlaceParagraphTextMM ActivePage, CStr(nb.NumberValue), False, gTextCursorXmm, CursorYmm, 0, 0
        gTextCursorXmm = gTextCursorXmm + 10

        ' создаём текстовые блоки для свойств
        For Each prop In nb.Properties
            PlaceParagraphTextMM ActivePage, CStr(prop), False, gTextCursorXmm, CursorYmm, 0, 0
            gTextCursorXmm = gTextCursorXmm + 10
        Next

    Next

    gTextCursorXmm = 0

End Sub
