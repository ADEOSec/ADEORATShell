VERSION 5.00
Begin VB.UserControl XP_ProgressBar 
   BackColor       =   &H0000C000&
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ScaleHeight     =   990
   ScaleWidth      =   3000
End
Attribute VB_Name = "XP_ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Mario Flores Cool Xp ProgressBar
'Emulating The Windows XP Progress Bar
'Open Source
'6 May 2004

'CD JUAREZ CHIHUAHUA MEXICO

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal HDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal HDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal HDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal HDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal HDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal HDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal HDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long


Const RGN_DIFF        As Long = 4
Const DT_SINGLELINE   As Long = &H20


'=====================================================
'THE RECT STRUCTURE
Private Type RECT
    Left      As Long     'The RECT structure defines the coordinates of the upper-left and lower-right corners of a rectangle
    Top       As Long
    Right     As Long
    Bottom    As Long
End Type

'=====================================================
'THE TRIVERTEX STRUCTURE
Private Type TRIVERTEX
    X         As Long
    Y         As Long
    Red       As Integer     'The TRIVERTEX structure contains color information and position information.
    Green     As Integer
    Blue      As Integer
    Alpha     As Integer
End Type
'=====================================================

'=====================================================
'THE GRADIENT_RECT STRUCTURE
Private Type GRADIENT_RECT
    UPPERLEFT  As Long       'The GRADIENT_RECT structure specifies the index of two vertices in the pVertex array.
    LOWERRIGHT As Long       'These two vertices form the upper-left and lower-right boundaries of a rectangle.
End Type
'=====================================================

'=====================================================
'THE RGB STRUCTURE
Private Type RGB
    R         As Integer
    G         As Integer     'Selects a red, green, blue (RGB) color based on the arguments supplied
    B         As Integer
End Type
'=====================================================


Public Enum cScrolling
    ccScrollingStandard = 0
    ccScrollingSmooth = 1
    ccScrollingSearch = 2
End Enum

Public Enum cOrientation
    ccOrientationHorizontal = 0
    ccOrientationVertical = 1
End Enum

Private m_Scrolling   As cScrolling
Private m_Orientation As cOrientation

'----------------------------------------------------
Private m_Color      As OLE_COLOR
Private m_hDC        As Long
Private m_hWnd       As Long        'PROPERTIES VARIABLES
Private m_Max        As Long
Private m_Min        As Long
Private m_Value      As Long
Private m_ShowText   As Boolean
Private m_ShowInTask As Boolean
'----------------------------------------------------


Private m_MemDC    As Boolean
Private m_ThDC     As Long
Private m_hBmp     As Long
Private m_hBmpOld  As Long
Private iFnt       As IFont
Private m_fnt      As IFont
Private hFntOld    As Long
Private m_lWidth   As Long
Private m_lHeight  As Long
Private fPercent   As Double
Private TR         As RECT
Private TBR        As RECT
Private TSR        As RECT
Private lSegmentWidth   As Long
Private lSegmentSpacing As Long



'==========================================================
'/---Draw ALL ProgressXP Bar  !!!!PUBLIC CALL!!!
'==========================================================

Public Sub DrawProgressBar()

  GetClientRect m_hWnd, TR                '//--- Reference = Control Client Area
  
            
            DrawFillRectangle TR, vbWhite, m_hDC
            
            CalcBarSize                   '//--- Calculate Progress and Percent Values
  
            PBarDraw                      '//--- Draw Scolling Bar (Inside Bar)
          
            If m_Scrolling = 0 Then DrawDivisions  '//--- Draw SegmentSpacing (This Will Generate the Blocks Effect)
            
            DrawTexto
  
            pDrawBorder                  '//--- Draw The XP Look Border

    If m_MemDC Then
        With UserControl
        pDraw .HDC, 0, 0, .ScaleWidth, .ScaleHeight, .ScaleLeft, .ScaleTop
        End With
    End If

End Sub


'==========================================================
'/---Calculate Division Bars & Percent Values
'==========================================================

Private Sub CalcBarSize()

   lSegmentWidth = 8   '/-- Windows Default
   lSegmentSpacing = 2 '/-- Windows Default
         
   LSet TBR = TR

   fPercent = (m_Value - m_Min) / (m_Max - m_Min)
   If fPercent > 1# Then fPercent = 1#              '/--  0 < Percent < 100
   If fPercent < 0# Then fPercent = 0#
   
      If m_Orientation = 0 Then
      
      '=======================================================================================
      '                                 Calc Horizontal ProgressBar
      '---------------------------------------------------------------------------------------
         TBR.Right = TR.Left + (TR.Right - TR.Left) * fPercent
         TBR.Right = TBR.Right - ((TBR.Right - TBR.Left) Mod (lSegmentWidth + lSegmentSpacing))
         If TBR.Right < TR.Left Then
            TBR.Right = TR.Left
         End If
         If TBR.Right < TR.Left Then TBR.Right = TR.Left
         
      Else
      
      '=======================================================================================
      '                                 Calc Vertical ProgressBar
      '---------------------------------------------------------------------------------------
         fPercent = 1# - fPercent - 0.02
         TBR.Top = TR.Top + (TR.Bottom - TR.Top) * fPercent
         TBR.Top = TBR.Top - ((TBR.Top - TBR.Bottom) Mod (lSegmentWidth + lSegmentSpacing))
         If TBR.Top > TR.Bottom Then TBR.Top = TR.Bottom
    
         
      
      End If

End Sub

'==========================================================
'/---Draw Division Bars
'==========================================================

Private Sub DrawDivisions()
 Dim i As Long
 Dim hBR As Long
  
  hBR = CreateSolidBrush(vbWhite)
  
      LSet TSR = TR
      
      If m_Orientation = 0 Then
      
      '=======================================================================================
      '                                 Draw Horizontal ProgressBar
      '---------------------------------------------------------------------------------------
         For i = TBR.Left + lSegmentWidth To TBR.Right Step lSegmentWidth + lSegmentSpacing
            TSR.Left = i + 2
            TSR.Right = i + 2 + lSegmentSpacing
            FillRect m_hDC, TSR, hBR
         Next i
      '---------------------------------------------------------------------------------------
      
      Else
      
      '=======================================================================================
      '                                  Draw Vertical ProgressBar
      '---------------------------------------------------------------------------------------
         For i = TBR.Bottom To TBR.Top + lSegmentWidth Step -(lSegmentWidth + lSegmentSpacing)
            TSR.Top = i - 2
            TSR.Bottom = i - 2 + lSegmentSpacing
            FillRect m_hDC, TSR, hBR
         Next i
       '---------------------------------------------------------------------------------------
      
      End If
      
      DeleteObject hBR
     
End Sub


'==========================================================
'/---Draw The ProgressXP Bar Border  ;)
'==========================================================

Private Sub pDrawBorder()
Dim RTemp As RECT
 
 Let RTemp = TR
  
 RTemp.Left = TR.Left + 1: RTemp.Top = TR.Top + 1
 DrawRectangle RTemp, GetLngColor(&HBEBEBE), m_hDC
 RTemp.Left = TR.Left + 1: RTemp.Top = TR.Top + 2: RTemp.Right = TR.Right - 1: RTemp.Bottom = TR.Bottom - 1
 DrawRectangle RTemp, GetLngColor(&HEFEFEF), m_hDC
 DrawRectangle TR, GetLngColor(&H686868), m_hDC

 Call SetPixelV(m_hDC, 1, 1, GetLngColor(&H686868))
 Call SetPixelV(m_hDC, TR.Right - 2, 1, GetLngColor(&H686868))
 Call SetPixelV(m_hDC, 1, TR.Bottom - 2, GetLngColor(&H686868))
 Call SetPixelV(m_hDC, TR.Right - 2, TR.Bottom - 2, GetLngColor(&H686868))  '//--Clean Up Corners

End Sub


'==========================================================
'/---Draw The ProgressXP Bar ;)
'==========================================================

Private Sub PBarDraw()
Dim TempRect As RECT
Dim ITemp    As Long
 
If m_Orientation = 0 Then

    TempRect.Left = TBR.Right
    TempRect.Right = 2
    TempRect.Top = 8
    TempRect.Bottom = TR.Bottom - 6


    '=======================================================================================
    '                                 Draw Horizontal ProgressBar
    '---------------------------------------------------------------------------------------
     
     If m_Scrolling = ccScrollingSearch Then
         GoSub HorizontalSearch
     Else
         DrawGradient m_hDC, 2, 3, TBR.Right - 2, 6, GetRGBColors(ShiftColorXP(m_Color, 150)), GetRGBColors(m_Color)
         DrawFillRectangle TempRect, m_Color, m_hDC
         DrawGradient m_hDC, 2, TempRect.Bottom - 2, TBR.Right - 2, 6, GetRGBColors(m_Color), GetRGBColors(ShiftColorXP(m_Color, 150))
     End If
     
Else
    
    TempRect.Left = 7
    TempRect.Right = TR.Right - 8
    TempRect.Top = TBR.Top
    TempRect.Bottom = TR.Bottom
    
    '=======================================================================================
    '                                 Draw Vertical ProgressBar
    '---------------------------------------------------------------------------------------
   
    If m_Scrolling = ccScrollingSearch Then
         GoSub VerticalSearch
    Else
         DrawGradient m_hDC, 2, TBR.Top, 6, TR.Bottom, GetRGBColors(ShiftColorXP(m_Color, 150)), GetRGBColors(m_Color), 0
         DrawFillRectangle TempRect, m_Color, m_hDC
         DrawGradient m_hDC, TR.Right - 8, TBR.Top, 6, TR.Bottom, GetRGBColors(m_Color), GetRGBColors(ShiftColorXP(m_Color, 150)), 0
    End If
   
    '--------------------   <-------- Gradient Color From (- to +)
    '||||||||||||||||||||   <-------- Fill Color
    '--------------------   <-------- Gradient Color From (+ to -)

End If

Exit Sub

HorizontalSearch:
    
    
    For ITemp = 0 To 2
    
        With TempRect
          .Left = TBR.Right + ((lSegmentSpacing + 10) * ITemp)
          .Right = .Left + 10
          .Top = 8
          .Bottom = TR.Bottom - 6
          DrawGradient m_hDC, .Left, 3, 10, 6, GetRGBColors(ShiftColorXP(m_Color, 220 - (40 * ITemp))), GetRGBColors(ShiftColorXP(m_Color, 200 - (40 * ITemp)))
          DrawFillRectangle TempRect, ShiftColorXP(m_Color, 200 - (40 * ITemp)), m_hDC
          DrawGradient m_hDC, .Left, .Bottom - 2, 10, 6, GetRGBColors(ShiftColorXP(m_Color, 200 - (40 * ITemp))), GetRGBColors(ShiftColorXP(m_Color, 220 - (40 * ITemp)))
        End With
        
    Next ITemp

Return

VerticalSearch:
    
     
    For ITemp = 0 To 2
    
        With TempRect
          .Left = 8
          .Right = TR.Right - 8
          .Top = TBR.Top + ((lSegmentSpacing + 10) * ITemp)
          .Bottom = .Top + 10
          DrawGradient m_hDC, 2, .Top, 6, 10, GetRGBColors(ShiftColorXP(m_Color, 220 - (40 * ITemp))), GetRGBColors(ShiftColorXP(m_Color, 200 - (40 * ITemp)))
          DrawFillRectangle TempRect, ShiftColorXP(m_Color, 200 - (40 * ITemp)), m_hDC
          DrawGradient m_hDC, .Right, .Top, 6, 10, GetRGBColors(ShiftColorXP(m_Color, 200 - (40 * ITemp))), GetRGBColors(ShiftColorXP(m_Color, 220 - (40 * ITemp)))
        End With
        
    Next ITemp

Return



End Sub

'======================================================================
'DRAWS THE PERCENT TEXT ON PROGRESS BAR
Private Function DrawTexto()
Dim ThisText As String

 If m_Scrolling = ccScrollingSearch Then
    ThisText = "Searching.."
 Else
    ThisText = (m_Max * m_Value) / 100 & " %"
 End If
 
  If (m_ShowText) Then
      Set iFnt = Font
      hFntOld = SelectObject(m_hDC, iFnt.hFont)
      SetBkMode m_hDC, 1
      SetTextColor m_hDC, vbBlack
      DrawText m_hDC, ThisText, -1, TR, DT_SINGLELINE Or 1 Or 4
      SelectObject m_hDC, hFntOld
   End If

End Function
'======================================================================

'======================================================================
'CONVERTION FUNCTION
Private Function GetLngColor(Color As Long) As Long
    
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function
'======================================================================

'======================================================================
'CONVERTION FUNCTION
Private Function GetRGBColors(Color As Long) As RGB

Dim HexColor As String
        
    HexColor = String(6 - Len(Hex(Color)), "0") & Hex(Color)
    GetRGBColors.R = "&H" & Mid(HexColor, 5, 2) & "00"
    GetRGBColors.G = "&H" & Mid(HexColor, 3, 2) & "00"
    GetRGBColors.B = "&H" & Mid(HexColor, 1, 2) & "00"
End Function
'======================================================================

'======================================================================
'DRAWS A BORDER RECTANGLE AREA OF AN SPECIFIED COLOR
Private Sub DrawRectangle(ByRef BRect As RECT, ByVal Color As Long, ByVal HDC As Long)

Dim hBrush As Long
    
    hBrush = CreateSolidBrush(Color)
    FrameRect HDC, BRect, hBrush
    DeleteObject hBrush

End Sub
'======================================================================

'======================================================================
'BLENDS AN SPECIFIED COLOR TO GET XP COLOR LOOK
Private Function ShiftColorXP(ByVal MyColor As Long, ByVal Base As Long) As Long

    Dim R As Long, G As Long, B As Long, Delta As Long

    R = (MyColor And &HFF)
    G = ((MyColor \ &H100) Mod &H100)
    B = ((MyColor \ &H10000) Mod &H100)
    
    Delta = &HFF - Base

    B = Base + B * Delta \ &HFF
    G = Base + G * Delta \ &HFF
    R = Base + R * Delta \ &HFF

    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If B > 255 Then B = 255

    ShiftColorXP = R + 256& * G + 65536 * B

End Function
'======================================================================

'======================================================================
'DRAWS A 2 COLOR GRADIENT AREA WITH A PREDEFINED DIRECTION
Private Sub DrawGradient( _
           ByVal cHdc As Long, _
           ByVal X As Long, _
           ByVal Y As Long, _
           ByVal X2 As Long, _
           ByVal Y2 As Long, _
           ByRef Color1 As RGB, _
           ByRef Color2 As RGB, _
           Optional Direction = 1)

    Dim Vert(1) As TRIVERTEX
    Dim gRect   As GRADIENT_RECT
   
    With Vert(0)
        .X = X
        .Y = Y
        .Red = Color1.R
        .Green = Color1.G
        .Blue = Color1.B
        .Alpha = 0&
    End With

    With Vert(1)
        .X = Vert(0).X + X2
        .Y = Vert(0).Y + Y2
        .Red = Color2.R
        .Green = Color2.G
        .Blue = Color2.B
        .Alpha = 0&
    End With

    gRect.UPPERLEFT = 1
    gRect.LOWERRIGHT = 0

    GradientFillRect cHdc, Vert(0), 2, gRect, 1, Direction

End Sub
'======================================================================

'======================================================================
'DRAWS A FILL RECTANGLE AREA OF AN SPECIFIED COLOR
Private Sub DrawFillRectangle(ByRef hRect As RECT, ByVal Color As Long, ByVal MyHdc As Long)

Dim hBrush As Long
 
   hBrush = CreateSolidBrush(GetLngColor(Color))
   FillRect MyHdc, hRect, hBrush
   DeleteObject hBrush

End Sub
'======================================================================

'======================================================================
'ROUNDS THE SELECTED WINDOW CORNERS
Private Sub RoundCorners(ByRef RcItem As RECT, ByVal m_hWnd As Long)

Dim rgn1 As Long, rgn2 As Long, rgnNorm As Long
    
    rgnNorm = CreateRectRgn(0, 0, RcItem.Right, RcItem.Bottom)
    rgn2 = CreateRectRgn(0, 0, 0, 0)

        rgn1 = CreateRectRgn(0, 0, 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, RcItem.Bottom, 2, RcItem.Bottom - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(RcItem.Right, 0, RcItem.Right - 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(RcItem.Right, RcItem.Bottom, RcItem.Right - 2, RcItem.Bottom - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, 1, 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, RcItem.Bottom - 1, 1, RcItem.Bottom - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(RcItem.Right, 1, RcItem.Right - 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(RcItem.Right, RcItem.Bottom - 1, RcItem.Right - 1, RcItem.Bottom - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        
        DeleteObject rgn1
        DeleteObject rgn2
        SetWindowRgn m_hWnd, rgnNorm, True
        DeleteObject rgnNorm
End Sub
'======================================================================

'======================================================================
'CHECKS-CREATES CORRECT DIMENSIONS OF THE TEMP DC
Private Function ThDC(Width As Long, Height As Long) As Long
   If m_ThDC = 0 Then
      If (Width > 0) And (Height > 0) Then
         pCreate Width, Height
      End If
   Else
      If Width > m_lWidth Or Height > m_lHeight Then
         pCreate Width, Height
      End If
   End If
   ThDC = m_ThDC
End Function
'======================================================================

'======================================================================
'CREATES THE TEMP DC
Private Sub pCreate(ByVal Width As Long, ByVal Height As Long)
Dim lhDCC As Long
   pDestroy
   lhDCC = CreateDC("DISPLAY", "", "", ByVal 0&)
   If Not (lhDCC = 0) Then
      m_ThDC = CreateCompatibleDC(lhDCC)
      If Not (m_ThDC = 0) Then
         m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
         If Not (m_hBmp = 0) Then
            m_hBmpOld = SelectObject(m_ThDC, m_hBmp)
            If Not (m_hBmpOld = 0) Then
               m_lWidth = Width
               m_lHeight = Height
               DeleteDC lhDCC
               Exit Sub
            End If
         End If
      End If
      DeleteDC lhDCC
      pDestroy
   End If
End Sub
'======================================================================

'======================================================================
'DRAWS THE TEMP DC
Public Sub pDraw( _
      ByVal HDC As Long, _
      Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
      Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
      Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0 _
   )
   If WidthSrc <= 0 Then WidthSrc = m_lWidth
   If HeightSrc <= 0 Then HeightSrc = m_lHeight
   BitBlt HDC, xDst, yDst, WidthSrc, HeightSrc, m_ThDC, xSrc, ySrc, vbSrcCopy
   
   On Error Resume Next
   
   
  
   
End Sub
'======================================================================

'======================================================================
'DESTROYS THE TEMP DC
Private Sub pDestroy()
   If Not m_hBmpOld = 0 Then
      SelectObject m_ThDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If Not m_hBmp = 0 Then
      DeleteObject m_hBmp
      m_hBmp = 0
   End If
   If Not m_ThDC = 0 Then
      DeleteDC m_ThDC
      m_ThDC = 0
   End If
   m_lWidth = 0
   m_lHeight = 0
End Sub
'======================================================================


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================================================
'USER CONTROL EVENTS
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================================================

Private Sub UserControl_Initialize()
    
 
     Dim fnt As New StdFont
         fnt.Name = "Tahoma"
         fnt.Size = 8
         Set Font = fnt
    
     With UserControl
        .BackColor = vbWhite
        .ScaleMode = vbPixels
     End With
     
     '----------------------------------------------------------
     'Default Values
     HDC = UserControl.HDC
     hwnd = UserControl.hwnd
     Max = 100
     Min = 0
     Value = 0
     Orientation = 0
     Scrolling = 0
     Color = &HC000&
     DrawProgressBar
     '----------------------------------------------------------

End Sub

Private Sub UserControl_Paint()

Dim cRect As RECT

 DrawProgressBar
 
 '-----------------------------------------------------------------------
 With UserControl
     GetClientRect .hwnd, cRect     'Round the Corners of the ProgressBar
     RoundCorners cRect, .hwnd
 End With
 '-----------------------------------------------------------------------
  
End Sub

Private Sub UserControl_Resize()
HDC = UserControl.HDC
End Sub

Private Sub UserControl_Terminate()
 pDestroy 'Destroy Temp DC
End Sub


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================================================
'USER CONTROL PROPERTIES
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================================================

Public Property Get Color() As OLE_COLOR
   Color = m_Color
End Property

Public Property Let Color(ByVal lColor As OLE_COLOR)
   m_Color = GetLngColor(lColor)
End Property

Public Property Get Font() As IFont
   Set Font = m_fnt
End Property

Public Property Set Font(ByRef fnt As IFont)
   Set m_fnt = fnt
End Property

Public Property Let Font(ByRef fnt As IFont)
   Set m_fnt = fnt
End Property

Public Property Get hwnd() As Long
   hwnd = m_hWnd
End Property

Public Property Let hwnd(ByVal chWnd As Long)
   m_hWnd = chWnd
End Property

Public Property Get HDC() As Long
   HDC = m_hDC
End Property

Public Property Let HDC(ByVal cHdc As Long)
   '=============================================
   'AntiFlick...Cleaner HDC
   m_hDC = ThDC(UserControl.ScaleWidth, UserControl.ScaleHeight)
   
   If m_hDC = 0 Then
      m_hDC = UserControl.HDC   'On Fail...Do it Normally
   Else
      m_MemDC = True
   End If
   '=============================================
End Property

Public Property Get Min() As Long
   Min = m_Min
End Property

Public Property Let Min(ByVal cMin As Long)
   m_Min = cMin
End Property

Public Property Get Max() As Long
   Max = m_Max
End Property

Public Property Let Max(ByVal cMax As Long)
   m_Max = cMax
End Property

Public Property Get Orientation() As cOrientation
   Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal cOrientation As cOrientation)
   m_Orientation = cOrientation
End Property

Public Property Get Scrolling() As cScrolling
   Scrolling = m_Scrolling
End Property

Public Property Let Scrolling(ByVal lScrolling As cScrolling)
   m_Scrolling = lScrolling
End Property

Public Property Get ShowText() As Boolean
   ShowText = m_ShowText
End Property

Public Property Let ShowText(ByVal bShowText As Boolean)
   m_ShowText = bShowText
   DrawProgressBar
End Property

Public Property Get ShowInTask() As Boolean
   ShowInTask = m_ShowInTask
End Property

Public Property Let ShowInTask(ByVal bShowInTask As Boolean)

End Property

Public Property Get Value() As Long
   Value = m_Value
End Property

Public Property Let Value(ByVal cValue As Long)
    m_Value = cValue
    DrawProgressBar
End Property
