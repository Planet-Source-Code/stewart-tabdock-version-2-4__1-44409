Attribute VB_Name = "modCaption"
Option Explicit
DefInt A-Z

Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Public Style As Long
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 32
End Type

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type

Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_CENTER = &H1
Private Const DT_BOTTOM = &H8


Private Const TRANSPARENT = 1
Private Const OPAQUE = 2

Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const SM_CYSMCAPTION = 51


Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


Dim captionFont As LOGFONT




Public Sub gradateColors(Colors() As Long, ByVal color1 As Long, ByVal Color2 As Long)

'Alright, I admit -- this routine was
'taken from a VBPJ issue a few months back.

Dim i As Integer
Dim dblR As Double, dblG As Double, dblB As Double
Dim addR As Double, addG As Double, addB As Double
Dim bckR As Double, bckG As Double, bckB As Double

   dblR = CDbl(color1 And &HFF)
   dblG = CDbl(color1 And &HFF00&) / 255
   dblB = CDbl(color1 And &HFF0000) / &HFF00&
   bckR = CDbl(Color2 And &HFF&)
   bckG = CDbl(Color2 And &HFF00&) / 255
   bckB = CDbl(Color2 And &HFF0000) / &HFF00&
   
   addR = (bckR - dblR) / UBound(Colors)
   addG = (bckG - dblG) / UBound(Colors)
   addB = (bckB - dblB) / UBound(Colors)
   
   For i = 0 To UBound(Colors)
      dblR = dblR + addR
      dblG = dblG + addG
      dblB = dblB + addB
      If dblR > 255 Then dblR = 255
      If dblG > 255 Then dblG = 255
      If dblB > 255 Then dblB = 255
      If dblR < 0 Then dblR = 0
      If dblG < 0 Then dblG = 0
      If dblG < 0 Then dblB = 0
      Colors(i) = RGB(dblR, dblG, dblB)
   Next
End Sub

Public Sub drawGradient(captionRect As RECT, hDC As Long, captionText As String, bActive As Boolean, gradient As Boolean, Optional captionOrientation As Integer, Optional captionForm As Form)

    Dim hBr As Long
    Dim drawDC As Long
    Dim bar As Long
    Dim width As Long
    Dim pixelStep As Long
    Dim storedCaptionRect As RECT
    Dim tmpGradFont As Long
    Dim oldFont As Long
    Dim hDCTemp As Long
    
    hDCTemp = hDC
    
    'Debug.Print captionText, captionOrientation, hDC, hDCTemp
    
    storedCaptionRect = captionRect
    
    If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
        width = captionRect.Right - captionRect.Left
    Else
        width = captionRect.Bottom - captionRect.Top
    End If
    
    pixelStep = width / 4
    
    ReDim Colors(pixelStep) As Long
    
    ' determine colors of gradient fill also determine if a gradient fill is required
    If bActive Then
        If gradient Then
            gradateColors Colors(), GradClr1, GradClr2
        Else
            gradateColors Colors(), TranslateColor(vbActiveTitleBar), TranslateColor(vbActiveTitleBar)
        End If
    Else
        If gradient Then
            gradateColors Colors(), TranslateColor(vbInactiveTitleBar), TranslateColor(vbButtonFace)
        Else
            gradateColors Colors(), TranslateColor(vbInactiveTitleBar), TranslateColor(vbInactiveTitleBar)
        End If
    End If
    
    For bar = 1 To pixelStep - 1
        hBr = CreateSolidBrush(Colors(bar))
        
        FillRect hDCTemp, captionRect, hBr
        
        If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
            captionRect.Left = captionRect.Left + 4
        Else
            captionRect.Bottom = captionRect.Bottom - 4
        End If
        
        DeleteObject hBr
    Next bar
  
    'draw caption text
    'Use a white caption, since the background is black
    'on the left side
    
    'get caption font information
    getCapsFont
    
    'If getting the caption font failed, use the font
    'from the gradient caption form.
    tmpGradFont = 0
    
    If captionText = "Form6" Then
      '  Beep
    End If
    
    If tmpGradFont = 0 Then
    
        'tmpGradFont = CreateFontIndirect(captionFont)
        
        If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
            captionFont.lfEscapement = 900
            
            'hDCTemp = captionForm.hDC
            'debug.Print "gradient font hdc set"
        End If
        
        tmpGradFont = CreateFontIndirect(captionFont)
        oldFont = SelectObject(hDCTemp, tmpGradFont)
    End If
    
    SetBkMode hDCTemp, TRANSPARENT
    
    If (bActive) Then
       SetTextColor hDCTemp, TranslateColor(vbActiveTitleBarText)
    Else
       SetTextColor hDCTemp, TranslateColor(vbInactiveTitleBarText)
    End If
    
    'move text a wee bit to the right
    If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
        'captionForm.CurrentX = 50
        'captionForm.CurrentY = captionForm.ScaleHeight - 100
        'captionForm.Print captionText
        'Debug.Print "caption text drawn", captionForm.CurrentX
        storedCaptionRect.Right = storedCaptionRect.Bottom - 40
        storedCaptionRect.Bottom = 8 + (captionForm.Height / Screen.TwipsPerPixelY)
        'Debug.Print "pixel height = "; captionForm.height / Screen.TwipsPerPixelY
        
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_BOTTOM
    Else
        storedCaptionRect.Left = storedCaptionRect.Left + 2
        storedCaptionRect.Right = storedCaptionRect.Right - 40
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS 'Or DT_HCENTER
    End If
    
    SelectObject hDCTemp, oldFont
    DeleteObject tmpGradFont
    tmpGradFont = 0

End Sub
Public Sub drawOfficeXP(captionRect As RECT, hDC As Long, captionText As String, bActive As Boolean, gradient As Boolean, Optional captionOrientation As Integer, Optional captionForm As Form)

    Dim hBr As Long
    Dim drawDC As Long
    Dim bar As Long
    Dim width As Long
    Dim pixelStep As Long
    Dim storedCaptionRect As RECT
    Dim tmpGradFont As Long
    Dim oldFont As Long
    Dim hDCTemp As Long
    Dim colorOutline As Long
    Dim colorInline As Long
    
    hDCTemp = hDC
    
    'Debug.Print captionText, captionOrientation, hDC, hDCTemp
    
    storedCaptionRect = captionRect
    
    If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
        width = captionRect.Right - captionRect.Left
    Else
        width = captionRect.Bottom - captionRect.Top
    End If
        
    ' determine colors of gradient fill also determine if a gradient fill is required
    If bActive Then
        colorOutline = TranslateColor(vbActiveTitleBar)
        colorInline = TranslateColor(vbActiveTitleBar)
    Else
        colorOutline = TranslateColor(vbInactiveTitleBar)
        colorInline = TranslateColor(vbButtonFace)
    End If
    

    hBr = CreateSolidBrush(colorOutline)
        
    FillRect hDCTemp, captionRect, hBr
    
    With captionRect
        .Top = .Top + 1
        .Left = .Left + 1
        .Right = .Right - 1
        .Bottom = .Bottom - 1
    End With
        
    hBr = CreateSolidBrush(colorInline)
        
    FillRect hDCTemp, captionRect, hBr
    
    DeleteObject hBr
  
    'draw caption text
    'Use a white caption, since the background is black
    'on the left side
    
    'get caption font information
    getCapsFont
    
    'If getting the caption font failed, use the font
    'from the gradient caption form.
    tmpGradFont = 0
    
    If captionText = "Form6" Then
      '  Beep
    End If
    
    If tmpGradFont = 0 Then
    
        'tmpGradFont = CreateFontIndirect(captionFont)
        
        If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
            captionFont.lfEscapement = 900
            
            'hDCTemp = captionForm.hDC
            'debug.Print "gradient font hdc set"
        End If
        
        tmpGradFont = CreateFontIndirect(captionFont)
        oldFont = SelectObject(hDCTemp, tmpGradFont)
    End If
    
    SetBkMode hDCTemp, TRANSPARENT
    
    If (bActive) Then
       SetTextColor hDCTemp, TranslateColor(vbActiveTitleBarText)
    Else
       SetTextColor hDCTemp, TranslateColor(vbInactiveTitleBarText)
    End If
    
    'move text a wee bit to the right
    If captionOrientation = tdAlignTop Or captionOrientation = tdAlignBottom Then
        'captionForm.CurrentX = 50
        'captionForm.CurrentY = captionForm.ScaleHeight - 100
        'captionForm.Print captionText
        'Debug.Print "caption text drawn", captionForm.CurrentX
        storedCaptionRect.Right = storedCaptionRect.Bottom - 40
        storedCaptionRect.Bottom = 8 + (captionForm.Height / Screen.TwipsPerPixelY)
        'Debug.Print "pixel height = "; captionForm.height / Screen.TwipsPerPixelY
        
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_BOTTOM
    Else
        storedCaptionRect.Left = storedCaptionRect.Left + 2
        storedCaptionRect.Right = storedCaptionRect.Right - 40
        DrawText hDCTemp, captionText, Len(captionText), storedCaptionRect, DT_SINGLELINE Or DT_END_ELLIPSIS 'Or DT_HCENTER
    End If
    
    SelectObject hDCTemp, oldFont
    DeleteObject tmpGradFont
    tmpGradFont = 0

End Sub

Sub drawGripper(captionRect As RECT, hDC As Long, gripStyle As Long, gripSides As Long, oneBar As Boolean, captionHeight As Long, Optional captionOrientation As Integer, Optional maximiseButton As Boolean)
    
    Dim numOfButtons As Integer
    
    If maximiseButton Then
        numOfButtons = 2
    Else
        numOfButtons = 1
    End If
    
    If oneBar Then
    
        If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
            With captionRect
                .Top = .Top + ((captionHeight - 11) / 2)
                .Left = .Left + 1
                .Right = .Right - (captionHeight * numOfButtons) + 5
                .Bottom = .Top + 4
            End With
        Else
            With captionRect
                .Top = .Top + (captionHeight * numOfButtons) - 4
                .Left = .Left + ((captionHeight - 14) / 2)
                .Right = .Left + 4
                .Bottom = .Bottom - 2
            End With
        End If
        
        DrawEdge hDC, captionRect, gripStyle, gripSides
    
    Else
    
        If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
            With captionRect
                .Top = .Top + ((captionHeight - 16) / 2)
                .Left = .Left + 1
                .Right = .Right - (captionHeight * numOfButtons) + 5
                .Bottom = .Top + 4
            End With
        Else
            With captionRect
                .Top = .Top + (captionHeight * numOfButtons) - 4
                .Left = .Left + ((captionHeight - 20) / 2) + 1
                .Right = .Left + 4
                .Bottom = .Bottom - 2
            End With
        End If
        
        DrawEdge hDC, captionRect, gripStyle, gripSides
        
        If captionOrientation <> tdAlignTop And captionOrientation <> tdAlignBottom Then
            With captionRect
                .Top = .Bottom + 1
                .Bottom = .Bottom + 5
            End With
        Else
            With captionRect
                .Left = .Right + 1
                .Right = .Left + 4
            End With
        End If
        
        DrawEdge hDC, captionRect, gripStyle, gripSides
        
    End If

End Sub

Private Sub getCapsFont()

    Dim NCM As NONCLIENTMETRICS
    Dim lfNew As LOGFONT

    NCM.cbSize = Len(NCM)
    Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)
    
    If NCM.iCaptionHeight = 0 Then
       captionFont.lfHeight = 0
    Else
       captionFont = NCM.lfSMCaptionFont
       'If captionFont.lfHeight < 10 Then
       ' captionFont.lfHeight = 14
       'End If
    End If
    
End Sub

Function getCaptionButtonHeight() As Long
    
    Dim NCM As NONCLIENTMETRICS
    Dim lfNew As LOGFONT

    NCM.cbSize = Len(NCM)
    Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)
    
    If NCM.iCaptionHeight = 0 Then
       'captionFont.lfHeight = 0
       getCaptionButtonHeight = 14
    Else
       'captionFont = NCM.lfSMCaptionFont
       getCaptionButtonHeight = NCM.iSMCaptionHeight
    End If
    
End Function
Function getCaptionHeight() As Long
        
    getCaptionHeight = GetSystemMetrics(SM_CYSMCAPTION)
    
    'If getCaptionHeight < 20 Then getCaptionHeight = 15
    
End Function

