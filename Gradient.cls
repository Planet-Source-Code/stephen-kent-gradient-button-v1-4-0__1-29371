VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Gradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'API Declarations
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RectAPI) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As PointAPI) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long

'API Data Structures
Private Type PointAPI   'API Point structure
    X   As Long
    Y   As Long
End Type

Private Type PointSng   'Internal Point structure
    X   As Double       'Uses Doubles for more precision and speed.
    Y   As Double
End Type

Private Type RectAPI    'API Rect structure
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'Drawing Constants
Private Const PS_SOLID As Long = 0  'Solid Pen Style (Used for CreatePen())

'Conversion Constants
Private Const PI    As Double = 3.14159265358979
Private Const RADS  As Double = PI / 180    '<Degrees> * RADS = radians

'HSL Types & constants
Const RGBMAX    As Long = 255       'Note since RGMMAX=HSLMAX I'll simplify the math in the HSL routines
Const HSLMAX    As Long = RGBMAX 'with Windows HSLMAX is 240 to make it dividable by 6
                                 'and still fit in a byte; we're using floating point
                                 'arithmetic so there's no need for that
Public Enum GradBlendMode
    gbmRGB = 0
    gbmHSL = 1
End Enum

Public Enum GradType
    gtNormal = 0
    gtElliptical = 1
    gtRectangular = 2
End Enum

'Property Storage Variables
Private mlColor1        As Long
Private mlColor2        As Long
Private mfAngle         As Double
Private mfRepetitions   As Double
Private m_GradientType  As GradType
Private m_BlendMode     As GradBlendMode

'Property Default Constants
Private Const mlDefColor1       As Long = &HFFFFFF  'White
Private Const mlDefColor2       As Long = &H0       'Black
Private Const mfDefAngle        As Double = 0       'Right to Left
Private Const mfDefRepetitions  As Double = 1
Private Const m_DefGradientType As Long = gtNormal
Private Const m_DefBlendMode    As Long = gbmRGB

Public Function Draw(picObj As Object) As Boolean
'Note: This class uses API functions to draw. If the
'      destination object is in AutoRedraw mode, the
'      Refresh method for that object must be invoked.

'picObj can be a Form or PictureBox.

    Dim lRet    As Long
    Dim lIdx    As Long
'    Dim lTime   As Long
    Dim uRect   As RectAPI

'    lTime = GetTickCount()

    On Error GoTo LocalError

    'Stop the window from updating until we're finished.
    lRet = LockWindowUpdate(picObj.hWnd)

    'Get the client rect in pixels
    lRet = GetClientRect(picObj.hWnd, uRect)

    'Test for possible errors (GetClientRect failure or Rect < 2 pixels)
    If lRet <> 0 Then
        If uRect.Right > 1 And uRect.Bottom > 1 Then
            lIdx = DrawGradient(picObj.hDc, uRect.Right, uRect.Bottom)
            Draw = (lIdx > 0)
        End If
    End If

    'My P3-500 took 99 millisecs (.099 secs) to create and draw 2554 diagonal
    'lines at 315 degrees. That was frmDemo maximized on a 1280 x 1024 screen.
    'At this speed I can redraw an entire 1280px. screen over 10 times per second.

    'Same size rect at a 0 degree angle took 48 millisecs (.048 secs) to create and
    'draw 1278 lines. This speed can redraw a 1280px. screen 20 times per second.

    'Uncomment the two lines below and the lTime line at the top
    'of this function to test the times on your PC.

'    lTime = GetTickCount() - lTime
'    Debug.Print CStr(lIdx / 2) & " lines drawn in " & CStr(lTime) & " milliseconds"

NormalExit:
    'Unlock the window to allow it to update now.
    lRet = LockWindowUpdate(0)
    Exit Function

LocalError:
    Err.Raise Err.Number
    Resume NormalExit
End Function

Private Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long, ByVal lSteps As Long, fRepetitions As Double, laRetColors() As Long) As Long
'Creates an array of colors blending from
'Color1 to Color2 in lSteps number of steps.
'Returns the count and fills the laRetColors() array.

    Dim lIdx    As Long
    Dim lIdx2   As Long
    Dim lRed    As Long
    Dim lGrn    As Long
    Dim lBlu    As Long
    Dim fRedStp As Double
    Dim fGrnStp As Double
    Dim fBluStp As Double
    Dim sHSL1   As HSLModel
    Dim sHSL2   As HSLModel
    Dim vHSL    As HSLModel

    'Stop possible error
    If lSteps < 2 Then lSteps = 2

    'Create the colors
    ReDim laRetColors(lSteps * 2)
    Select Case m_BlendMode
        Case gbmRGB
            'Extract Red, Blue and Green values from the start and end colors.
            lRed = (lColor1 And &HFF&)
            lGrn = (lColor1 And &HFF00&) / &H100
            lBlu = (lColor1 And &HFF0000) / &H10000

            'Find the amount of change for each color element per color change.
            fRedStp = Div((lColor2 And &HFF&) - lRed, lSteps / fRepetitions)
            fGrnStp = Div(((lColor2 And &HFF00&) / &H100&) - lGrn, lSteps / fRepetitions)
            fBluStp = Div(((lColor2 And &HFF0000) / &H10000) - lBlu, lSteps / fRepetitions)

            laRetColors(0) = lColor1    'First Color
            laRetColors(Int(lSteps / fRepetitions)) = lColor2        'Last Color
            laRetColors(Int(lSteps / fRepetitions) + 1) = lColor2    'Last Color
            For lIdx = 1 To Int(lSteps / fRepetitions) - 1           'All Colors between
                laRetColors(lIdx) = CLng(lRed + (fRedStp * lIdx)) + _
                    (CLng(lGrn + (fGrnStp * lIdx)) * &H100&) + _
                    (CLng(lBlu + (fBluStp * lIdx)) * &H10000)
            Next lIdx
        Case gbmHSL
            Set sHSL1 = New HSLModel
            Set sHSL2 = New HSLModel
            Set vHSL = New HSLModel
            ' A HUE gradient
            sHSL1.Color = lColor1
            sHSL2.Color = lColor2
            lIdx = sHSL2.Hue - sHSL1.Hue
            If Abs(lIdx) > HSLMAX / 2 Then
                lIdx = (HSLMAX - Abs(lIdx)) * -Sgn(lIdx)
            End If
            vHSL.Hue = (lIdx) / Int(lSteps / fRepetitions)
            vHSL.Saturation = (sHSL2.Saturation - sHSL1.Saturation) / Int(lSteps / fRepetitions)
            vHSL.Luminance = (sHSL2.Luminance - sHSL1.Luminance) / Int(lSteps / fRepetitions)
            With sHSL1
                For lIdx = 0 To Int(lSteps / fRepetitions) + 1
                    .Hue = .Hue + vHSL.Hue
                    Select Case .Hue
                    Case Is < 0
                        .Hue = .Hue + HSLMAX
                    Case Is > HSLMAX
                        .Hue = .Hue - HSLMAX
                    End Select
                    .Saturation = .Saturation + vHSL.Saturation
                    .Luminance = .Luminance + vHSL.Luminance
                    laRetColors(lIdx) = .Color
                Next lIdx
            End With
    End Select

    ' first the complete color repetitions
    If Int(fRepetitions) >= 1 Then
      For lIdx2 = 1 To Int(fRepetitions) + 1
        If lIdx2 / 2 = Int(lIdx2 / 2) Then
          ' reverse gradient
          For lIdx = 0 To Int(lSteps / fRepetitions)
            laRetColors(((lIdx2 - 1) * Int(lSteps / fRepetitions)) + lIdx) = laRetColors((lSteps / fRepetitions) - lIdx)
          Next lIdx
        Else
          ' normal gradient
          For lIdx = 0 To Int(lSteps / fRepetitions)
            laRetColors(((lIdx2 - 1) * Int(lSteps / fRepetitions)) + lIdx) = laRetColors(lIdx)
          Next lIdx
        End If
      Next lIdx2
    End If

    'Return number of colors in array
    BlendColors = lSteps
End Function

Private Function DrawGradient(ByVal hDc As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
    Dim bDone       As Boolean
    Dim iIncX       As Integer
    Dim iIncY       As Integer
    Dim lIdx        As Long
    Dim lRet        As Long
    Dim hPen        As Long
    Dim hOldPen     As Long
    Dim hBrush      As Long
    Dim hOldBrush   As Long
    Dim lPointCnt   As Long
    Dim laColors()  As Long
    Dim fMovX       As Double
    Dim fMovY       As Double
    Dim fDist       As Double
    Dim fAngle      As Double
    Dim fLongSide   As Double
    Dim uTmpPt      As PointAPI
    Dim uaPts()     As PointAPI
    Dim uaTmpPts()  As PointSng
    Dim MinDist     As Long

    On Error GoTo LocalError

    Select Case m_GradientType
        Case gtNormal
            'Start with center of rect
            ReDim uaTmpPts(2)
            uaTmpPts(2).X = Int(lWidth / 2)
            uaTmpPts(2).Y = Int(lHeight / 2)
        
            'Calc distance to furthest edge as if rect were square
            fLongSide = IIf(lWidth > lHeight, lWidth, lHeight)
            fDist = (Sqr((fLongSide ^ 2) + (fLongSide ^ 2)) + 2) / 2
        
            'Create points to the left and the right at a 0º angle (horizontal)
            uaTmpPts(0).X = uaTmpPts(2).X - fDist
            uaTmpPts(0).Y = uaTmpPts(2).Y
            uaTmpPts(1).X = uaTmpPts(2).X + fDist
            uaTmpPts(1).Y = uaTmpPts(2).Y
        
            'Lines will be drawn perpendicular to mfAngle so
            'add 90º and correct for 360º wrap
            fAngle = (mfAngle + 90) Mod 360
        
            'Rotate second and third points to fAngle
            Call RotatePoint(uaTmpPts(2), uaTmpPts(0), fAngle)
            Call RotatePoint(uaTmpPts(2), uaTmpPts(1), fAngle)
        
            'We now have a line that crosses the center and
            'two sides of the rect at the correct angle.
        
            'Calc the starting quadrant, direction of and amount of first move
            '(fMovX, fMovY moves line from center to starting edge)
            'and direction of each incremental move (iIncX, iIncY).
            Select Case mfAngle
                Case 0 To 90
                    'Left Bottom
                    If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                        'Move line to left edge; Draw left to right
                        fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, -uaTmpPts(0).X, -uaTmpPts(1).X)
                        fMovY = 0
                        iIncX = 1
                        iIncY = 0
                    Else
                        'Move line to bottom edge; Draw bottom to top
                        fMovX = 0
                        fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, lHeight - uaTmpPts(1).Y, lHeight - uaTmpPts(0).Y)
                        iIncX = 0
                        iIncY = -1
                    End If
                Case 90 To 180
                    'Right Bottom
                    If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                        'Move line to right edge; Draw right to left
                        fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, lWidth - uaTmpPts(1).X, lWidth - uaTmpPts(0).X)
                        fMovY = 0
                        iIncX = -1
                        iIncY = 0
                    Else
                        'Move line to bottom edge; Draw bottom to top
                        fMovX = 0
                        fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, lHeight - uaTmpPts(1).Y, lHeight - uaTmpPts(0).Y)
                        iIncX = 0
                        iIncY = -1
                    End If
                Case 180 To 270
                    'Right Top
                    If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                        'Move line to right edge; Draw right to left
                        fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, lWidth - uaTmpPts(1).X, lWidth - uaTmpPts(0).X)
                        fMovY = 0
                        iIncX = -1
                        iIncY = 0
                    Else
                        'Move line to top edge; Draw top to bottom
                        fMovX = 0
                        fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, -uaTmpPts(0).Y, -uaTmpPts(1).Y)
                        iIncX = 0
                        iIncY = 1
                    End If
                Case Else   '(270 to 360)
                    'Left Top
                    If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                        'Move line to left edge; Draw left to right
                        fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, -uaTmpPts(0).X, -uaTmpPts(1).X)
                        fMovY = 0
                        iIncX = 1
                        iIncY = 0
                    Else
                        'Move line to top edge; Draw top to bottom
                        fMovX = 0
                        fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, -uaTmpPts(0).Y, -uaTmpPts(1).Y)
                        iIncX = 0
                        iIncY = 1
                    End If
            End Select
        
            'At this point we could calculate where the lines will cross the rect edges, but
            'this would slow things down. The picObj clipping region will take care of this.
        
            'Start with 1000 points and add more if needed. This increases
            'speed by not re-dimming the array in each loop.
            ReDim uaPts(999)
        
            'Set the first two points in the array
            uaPts(0).X = uaTmpPts(0).X + fMovX
            uaPts(0).Y = uaTmpPts(0).Y + fMovY
            uaPts(1).X = uaTmpPts(1).X + fMovX
            uaPts(1).Y = uaTmpPts(1).Y + fMovY
        
            lIdx = 2
            'Create the rest of the points by incrementing both points
            'on each line iIncX, iIncY from the previous line's points.
            'Where we stop depends on the direction of travel.
            'We'll continue until both points in a set reach the end.
            While Not bDone
                uaPts(lIdx).X = uaPts(lIdx - 2).X + iIncX
                uaPts(lIdx).Y = uaPts(lIdx - 2).Y + iIncY
                lIdx = lIdx + 1
                Select Case True
                    Case iIncX > 0  'Moving Left to Right
                        bDone = uaPts(lIdx - 1).X > lWidth And uaPts(lIdx - 2).X > lWidth
                    Case iIncX < 0  'Moving Right to Left
                        bDone = uaPts(lIdx - 1).X < 0 And uaPts(lIdx - 2).X < 0
                    Case iIncY > 0  'Moving Top to Bottom
                        bDone = uaPts(lIdx - 1).Y > lHeight And uaPts(lIdx - 2).Y > lHeight
                    Case iIncY < 0  'Moving Bottom to Top
                        bDone = uaPts(lIdx - 1).Y < 0 And uaPts(lIdx - 2).Y < 0
                End Select
                If (lIdx Mod 1000) = 0 Then
                    ReDim Preserve uaPts(UBound(uaPts) + 1000)
                End If
            Wend
        
            'Free excess memory (may have 1001 points dimmed to 2000)
            ReDim Preserve uaPts(lIdx - 1)
        
            'Create the array of colors blending from mlColor1 to mlColor2
            lRet = BlendColors(mlColor1, mlColor2, lIdx / 2, mfRepetitions, laColors)

            'Now draw each line in it's own color
            For lIdx = 0 To UBound(uaPts) - 1 Step 2
                'Move to next point
                lRet = MoveToEx(hDc, uaPts(lIdx).X, uaPts(lIdx).Y, uTmpPt)
                'Create the colored pen and select it into the DC
                hPen = CreatePen(PS_SOLID, 1, laColors(Int(lIdx / 2)))
                hOldPen = SelectObject(hDc, hPen)
                'Draw the line
                lRet = LineTo(hDc, uaPts(lIdx + 1).X, uaPts(lIdx + 1).Y)
                'Get the pen back out of the DC and destroy it
                lRet = SelectObject(hDc, hOldPen)
                lRet = DeleteObject(hPen)
            Next lIdx
            DrawGradient = lIdx
        Case gtElliptical
            iIncX = Int(lWidth / 2) - Int((lWidth / -lHeight) * (lHeight / 2))
            iIncY = Int(lHeight / 2) - Int((-lHeight / lWidth) * (lWidth / 2))
            lPointCnt = IIf((iIncX < iIncY), iIncX, iIncY)
            fMovX = IIf((iIncX < iIncY), 1, (lWidth / 2) / lPointCnt)
            fMovY = IIf((iIncX < iIncY), (lHeight / 2) / lPointCnt, 1)
            lRet = BlendColors(mlColor1, mlColor2, lPointCnt, mfRepetitions, laColors)
            For lIdx = 0 To lPointCnt
                'Create the colored pen and select it into the DC
                hPen = CreatePen(PS_SOLID, 1, laColors(lIdx))
                hOldPen = SelectObject(hDc, hPen)
                'Create the colored brush and select it into the DC
                hBrush = CreateSolidBrush(laColors(lIdx))
                hOldBrush = SelectObject(hDc, hBrush)
                'Draw the ellipse
                lRet = Ellipse(hDc, Int(lWidth / 2) - (iIncX - (lIdx * fMovX)), Int(lHeight / 2) - (iIncY - (lIdx * fMovY)), Int(lWidth / 2) + (iIncX - (lIdx * fMovX)), Int(lHeight / 2) + (iIncY - (lIdx * fMovY)))
                'Get the brush back out of the DC and destroy it
                lRet = SelectObject(hDc, hOldBrush)
                lRet = DeleteObject(hBrush)
                'Get the pen back out of the DC and destroy it
                lRet = SelectObject(hDc, hOldPen)
                lRet = DeleteObject(hPen)
            Next lIdx
        Case gtRectangular
            MinDist = IIf((lWidth < lHeight), lWidth, lHeight)
            lRet = BlendColors(mlColor1, mlColor2, Int(MinDist / 2), mfRepetitions, laColors)
            For lIdx = 0 To (MinDist / 2)
                'Create the colored pen and select it into the DC
                hPen = CreatePen(PS_SOLID, 1, laColors(lIdx))
                hOldPen = SelectObject(hDc, hPen)
                'Draw the rectangle
                lRet = Rectangle(hDc, lIdx, lIdx, lWidth - lIdx, lHeight - lIdx)
                'Get the pen back out of the DC and destroy it
                lRet = SelectObject(hDc, hOldPen)
                lRet = DeleteObject(hPen)
            Next lIdx
    End Select

NormalExit:
    'Free the memory
    Erase laColors
    Erase uaPts
    Erase uaTmpPts
    Exit Function

LocalError:
    'MsgBox Err.Description, vbExclamation, "GradientRect.cls"
    DrawGradient = 0
    Resume 'NormalExit
End Function

Private Sub RotatePoint(uAxisPt As PointSng, uRotatePt As PointSng, fDegrees As Double)
    Dim fDX         As Double
    Dim fDY         As Double
    Dim fRadians    As Double

    fRadians = fDegrees * RADS
    fDX = uRotatePt.X - uAxisPt.X
    fDY = uRotatePt.Y - uAxisPt.Y
    uRotatePt.X = uAxisPt.X + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
    uRotatePt.Y = uAxisPt.Y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))
End Sub

Private Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double
'Divides dNumer by dDenom if dDenom <> 0
'Auto-Handles 'Division By Zero' error.

    On Error Resume Next
    Div = dNumer / dDenom
End Function

Public Property Let Color1(ByVal lData As Long)
    Dim lIdx As Long

    mlColor1 = lData
    If mlColor1 < 0 Then
        lIdx = (mlColor1 And Not &H80000000)
        If lIdx >= 0 And lIdx <= 24 Then
            mlColor1 = GetSysColor(lIdx)
        End If
    End If
End Property

Public Property Get Color1() As Long
    Color1 = mlColor1
End Property

Public Property Let Color2(ByVal lData As Long)
    Dim lIdx As Long

    mlColor2 = lData
    If mlColor2 < 0 Then
        lIdx = (mlColor2 And Not &H80000000)
        If lIdx >= 0 And lIdx <= 24 Then
            mlColor2 = GetSysColor(lIdx)
        End If
    End If
End Property

Public Property Get Color2() As Long
    Color2 = mlColor2
End Property

Public Property Let Angle(ByVal fData As Double)
'Angles are counter-clockwise and may be
'any Single value from 0 to 359.999999999.

' 135  90 45
'    \ | /
'180 --o-- 0
'    / | \
' 235 270 315

    'Correct angle to ensure between 0 and 359.999999999
    mfAngle = fData Mod 360
End Property

Public Property Get Angle() As Double
    Angle = mfAngle
End Property

Public Property Let Repetitions(ByVal fData As Double)
    mfRepetitions = Abs(fData)
    If mfRepetitions = 0 Then mfRepetitions = 1
End Property

Public Property Get Repetitions() As Double
    Repetitions = mfRepetitions
End Property

Public Property Let GradientType(ByVal eData As GradType)
    m_GradientType = eData
End Property

Public Property Get GradientType() As GradType
    GradientType = m_GradientType
End Property

Public Property Let BlendMode(ByVal eData As GradBlendMode)
    m_BlendMode = eData
End Property

Public Property Get BlendMode() As GradBlendMode
    BlendMode = m_BlendMode
End Property

Private Sub Class_Initialize()
    mlColor1 = mlDefColor1
    mlColor2 = mlDefColor2
    mfAngle = mfDefAngle
    mfRepetitions = mfDefRepetitions
    m_GradientType = m_DefGradientType
    m_BlendMode = m_DefBlendMode
End Sub
