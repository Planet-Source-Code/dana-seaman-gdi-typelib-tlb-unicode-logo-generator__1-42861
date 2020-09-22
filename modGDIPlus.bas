Attribute VB_Name = "modGDIPlus"
Option Explicit

Public Function AppPath() As String
   AppPath = App.path
   'Root ends in "\" so check first
   If Right$(AppPath, 1) <> "\" Then
      AppPath = AppPath & "\"
   End If
End Function

Public Function ColorSetAlpha(ByVal lColor As Long, _
   ByVal Alpha As Byte) As Long

   Dim bytestruct       As COLORBYTES
   Dim result           As COLORLONG

   result.longval = lColor
   LSet bytestruct = result

   bytestruct.AlphaByte = Alpha

   LSet result = bytestruct
   ColorSetAlpha = result.longval
End Function

Public Function Gdi2VbColor(ByVal GdiColor As Long) As Long
   Dim Red              As Long
   Dim Green            As Long
   Dim Blue             As Long

   GdiColor = GdiColor And &HFFFFFF 'Strip alpha(if any)
   'Get components. Rebuild with Red and Blue swapped.
   Red = (GdiColor \ 65536) And &HFF
   Green = (GdiColor \ 256) And &HFF
   Blue = GdiColor And &HFF

   Gdi2VbColor = (Blue * 65536) + (Green * 256) + Red
End Function

Public Sub InflateRectF(rct As RECTF, ByVal x As Variant, ByVal y As Variant)
   rct.Left = rct.Left - x
   rct.Top = rct.Top - y
   rct.Width = rct.Width + x + x
   rct.Height = rct.Height + y + y
End Sub

Public Sub DrawGdipLogo(ByVal lhdc As Long, _
   ByRef rct As RECTF, _
   ByVal IsEllipse As Boolean, _
   ByRef penWidth() As Integer, _
   ByRef BrushType() As BrushType, _
   ByRef GradientMode() As LinearGradientMode, _
   ByRef HatchStyle() As HatchStyle, _
   ByVal sText As String, _
   ByVal AntiAliased As Boolean, _
   ByRef obj As Object, _
   ByVal Alignment As Integer, _
   ByRef Gamma() As Boolean, _
   ByRef StartColor() As Long, _
   ByRef EndColor() As Long, _
   ByRef Texture() As String, _
   ByVal IsUnicode As Boolean, _
   ByVal LocaleID As Long)

   Dim Idx              As Integer
   Dim graphics         As Long
   Dim img              As Long
   Dim pen              As Long
   Dim brush            As Long
   Dim fontFam          As Long
   Dim curFont          As Long
   Dim strFormat        As Long
   Dim box              As RECTF
   Dim rct2             As RECTF
   Dim FS               As Long
   Dim IsAvailable      As Long
   Dim path             As Long
   Dim penColorStart    As Long
   Dim penColorEnd      As Long
   Dim borderWidth      As Single
   Dim borderCount      As Long
   
   ' Initializations
   GdipCreateFromHDC lhdc, graphics   ' Initialize the graphics class - required for all drawing
   
   If BrushType(0) = BrushTypeLinearGradient Then
      borderCount = 2
   End If
   For Idx = 0 To borderCount
      GdipCreatePath FillModeWinding, path
      Select Case Idx
         Case 0
            penColorStart = StartColor(0)
            penColorEnd = EndColor(0)
            borderWidth = penWidth(0)
         Case 1 'Swap colors for 3D effect
            InflateRectF rct, -penWidth(0), -penWidth(0)
            penColorStart = EndColor(0)
            penColorEnd = StartColor(0)
            borderWidth = penWidth(0)
         Case 2 'Innermost border is half Outer/Mid
            borderWidth = 0.5 * penWidth(0)
            InflateRectF rct, 1.5 * -borderWidth, 1.5 * -borderWidth
            penColorStart = StartColor(0)
            penColorEnd = EndColor(0)
      End Select
      If IsEllipse Then
         GdipAddPathEllipse path, _
            rct.Left, _
            rct.Top, _
            rct.Width, _
            rct.Height
      Else
         GdipAddPathRectangles path, rct, 1
      End If

      If AntiAliased Then
         GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
      End If
      If Idx = borderCount Then
         'Background Fill
         Select Case BrushType(1)
            Case BrushTypeSolidColor
               GdipCreateSolidFill StartColor(1), brush
            Case BrushTypeLinearGradient
               rct2 = SetGradientRectF(rct, GradientMode(1), 2, True)
               GdipCreateLineBrushFromRect rct2, _
                  StartColor(1), _
                  EndColor(1), _
                  GradientMode(1), _
                  WrapModeTileFlipX, brush
               GdipSetLineGammaCorrection brush, Gamma(1)
            Case BrushTypePathGradient
               GdipCreatePathGradientFromPath path, brush
               GdipSetPathGradientCenterColor brush, StartColor(1)
               GdipSetPathGradientSurroundColorsWithCount brush, EndColor(1), 1
            Case BrushTypeTextureFill
               GdipLoadImageFromFile Texture(1), img
               GdipCreateTexture img, WrapModeTile, brush
               GdipDisposeImage img
            Case BrushTypeHatchFill
               GdipCreateHatchBrush HatchStyle(1), StartColor(1), EndColor(1), brush
         End Select

         If brush Then
            GdipFillPath graphics, brush, path
            GdipDeleteBrush brush
         End If
      End If
      
      'Border Pen
      Select Case BrushType(0)
         Case PenTypeUnknown
         Case PenTypeSolidColor
            GdipCreatePen1 StartColor(0), penWidth(0), UnitPixel, pen
         Case PenTypeLinearGradient
            rct2 = SetGradientRectF(rct, GradientMode(0), 2, True)
            GdipCreateLineBrushFromRect rct2, _
               penColorStart, _
               penColorEnd, _
               GradientMode(0), _
               WrapModeTileFlipX, brush
            GdipSetLineGammaCorrection brush, Gamma(0)
            GdipCreatePen2 brush, borderWidth, UnitPixel, pen
            GdipDeleteBrush brush
         Case PenTypePathGradient
            GdipCreatePathGradientFromPath path, brush
            GdipSetPathGradientCenterColor brush, StartColor(0)
            GdipSetPathGradientSurroundColorsWithCount brush, EndColor(0), 1
            GdipSetPathGradientFocusScales brush, _
               1 - (2 * borderWidth / rct.Width), 1 - (2 * borderWidth / rct.Height)
            GdipSetLineGammaCorrection brush, Gamma(0)
            GdipCreatePen2 brush, 2 * penWidth(0), UnitPixel, pen
         Case PenTypeTextureFill
            GdipLoadImageFromFile Texture(0), img
            GdipCreateTexture img, WrapModeTile, brush
            GdipDisposeImage img
            GdipCreatePen2 brush, penWidth(0), UnitPixel, pen
            GdipDeleteBrush brush
         Case PenTypeHatchFill
            GdipCreateHatchBrush HatchStyle(0), StartColor(0), EndColor(0), brush
            GdipCreatePen2 brush, penWidth(0), UnitPixel, pen
            GdipDeleteBrush brush
      End Select

      If pen Then
         GdipDrawPath graphics, pen, path
         GdipDeletePen pen
      End If
      'Free the path for next
      GdipDeletePath path
   Next
   'Text stuff follows

   ' Set the Text Rendering Quality
   'GdipSetTextRenderingHint graphics, TextRenderingHint

   ' Create a font family object to allow us to create a font
   ' We have no font collection here, so pass a NULL for that parameter

   GdipCreateFontFamilyFromName obj.FontName, 0, fontFam
   GdipIsStyleAvailable fontFam, FS, IsAvailable
   If IsAvailable = 0 Then
      Dim Msg              As String
      Msg = "Font family " & _
         obj.FontName & _
         " NOT available under GDI+"
      MsgBox Msg
      Exit Sub
   End If
   ' Create the font from the specified font family name
   ' >> Note that we have changed the drawing Unit from pixels to points!!
   FS = FontStyleRegular 'not really needed since it's zero
   If obj.FontBold Then FS = FS + FontStyleBold
   If obj.FontItalic Then FS = FS + FontStyleItalic
   If obj.FontStrikethru Then FS = FS + FontStyleStrikeout
   If obj.FontUnderline Then FS = FS + FontStyleUnderline

   GdipCreateFont fontFam, obj.FontSize, FS, UnitPoint, curFont
   ' Create the StringFormat object
   ' We can pass NULL for the flags and language id if we want
   GdipCreateStringFormat 0, LocaleID, strFormat

   ' Justify each line of text
   GdipSetStringFormatAlign strFormat, Alignment Mod 3

   ' Justify the block of text (top to bottom) in the rectangle.
   GdipSetStringFormatLineAlign strFormat, Int(Alignment \ 3)

   GdipCreatePath FillModeWinding, path

   GdipAddPathString path, _
      sText, -1, _
      fontFam, _
      FS, _
      obj.FontSize, _
      rct, _
      strFormat

   'GdipMeasureString graphics, sText, -1, curFont, _
   rct, strFormat, box, 0, 0
   LSet box = rct

   ' Create brushe for text fill
   Select Case BrushType(3)
      Case BrushTypeSolidColor
         GdipCreateSolidFill StartColor(3), brush
      Case BrushTypeLinearGradient
         rct2 = SetGradientRectF(box, GradientMode(3), 2, True)
         GdipCreateLineBrushFromRect rct2, _
            StartColor(3), _
            EndColor(3), _
            GradientMode(3), _
            WrapModeTileFlipX, brush
         GdipSetLineGammaCorrection brush, Gamma(3)
      Case BrushTypePathGradient
         GdipCreatePathGradientFromPath path, brush
         'GdipSetPathGradientLinearBlend brush, 50, 50
         GdipSetPathGradientCenterColor brush, StartColor(3)
         GdipSetPathGradientSurroundColorsWithCount brush, EndColor(3), 1
      Case BrushTypeTextureFill
         GdipLoadImageFromFile Texture(3), img
         GdipCreateTexture img, WrapModeTile, brush
         GdipDisposeImage img
      Case BrushTypeHatchFill
         GdipCreateHatchBrush HatchStyle(3), StartColor(3), EndColor(3), brush
   End Select

   If brush Then
      GdipFillPath graphics, brush, path
      GdipDeleteBrush brush
   End If

   'Create text pen
   Select Case BrushType(2)
      Case PenTypeUnknown
      Case PenTypeSolidColor
         GdipCreatePen1 StartColor(2), penWidth(2), UnitPixel, pen
      Case PenTypeLinearGradient
         rct2 = SetGradientRectF(box, GradientMode(2), 2, True)
         GdipCreateLineBrushFromRect rct2, _
            StartColor(2), _
            EndColor(2), _
            GradientMode(2), _
            WrapModeTileFlipX, brush
         GdipSetLineGammaCorrection brush, Gamma(2)
         GdipCreatePen2 brush, penWidth(2), UnitPixel, pen
         GdipDeleteBrush brush
      Case PenTypePathGradient
         GdipCreatePathGradientFromPath path, brush
         GdipSetPathGradientCenterColor brush, StartColor(2)
         GdipSetPathGradientSurroundColorsWithCount brush, EndColor(2), 1
         GdipCreatePen2 brush, penWidth(2), UnitPixel, pen
      Case PenTypeTextureFill
         GdipLoadImageFromFile Texture(2), img
         GdipCreateTexture img, WrapModeTile, brush
         GdipCreatePen2 brush, penWidth(2), UnitPixel, pen
         GdipDeleteBrush brush
      Case PenTypeHatchFill
         GdipCreateHatchBrush HatchStyle(2), StartColor(2), EndColor(2), brush
         GdipCreatePen2 brush, penWidth(2), UnitPixel, pen
         GdipDeleteBrush brush
   End Select

   If pen Then
      GdipDrawPath graphics, pen, path
      GdipDeletePen pen
   End If

   ' Cleanup
   GdipDeleteStringFormat strFormat
   GdipDeleteFont curFont
   GdipDeleteFontFamily fontFam
   GdipDeletePath path
   GdipDeleteGraphics graphics

End Sub

Private Function SetGradientRectF(ByRef TR As RECTF, _
   ByVal GradientMode As LinearGradientMode, _
   ByVal sngScale As Single, _
   ByVal MaintainAspectRatio As Boolean) As RECTF

   Dim sngW             As Single
   Dim sngH             As Single

   LSet SetGradientRectF = TR
   sngW = SetGradientRectF.Width / sngScale
   sngH = SetGradientRectF.Height / sngScale

   Select Case GradientMode
      Case LinearGradientModeHorizontal
         SetGradientRectF.Width = sngW
      Case LinearGradientModeVertical
         SetGradientRectF.Height = sngH
      Case LinearGradientModeForwardDiagonal, LinearGradientModeBackwardDiagonal
         If MaintainAspectRatio Then
            SetGradientRectF.Width = sngW
            SetGradientRectF.Height = sngH
         Else
            If SetGradientRectF.Width < SetGradientRectF.Height Then
               SetGradientRectF.Width = sngW
            Else
               SetGradientRectF.Width = sngH
            End If
            SetGradientRectF.Height = SetGradientRectF.Width
         End If
   End Select

End Function

