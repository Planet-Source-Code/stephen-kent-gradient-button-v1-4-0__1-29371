VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HSLModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'HSL Types & constants
Const RGBMAX    As Long = 255       'Note since RGMMAX=HSLMAX I'll simplify the math in the HSL routines
Const HSLMAX    As Long = RGBMAX 'with Windows HSLMAX is 240 to make it dividable by 6
                                 'and still fit in a byte; we're using floating point
                                 'arithmetic so there's no need for that

Private m_dHue As Double
Private m_dLuminance As Double
Private m_dSaturation As Double

Public Property Get Hue() As Double
    Hue = m_dHue
End Property

Public Property Let Hue(ByVal dNewHue As Double)
    m_dHue = dNewHue
End Property

Public Property Get Luminance() As Double
    Luminance = m_dLuminance
End Property

Public Property Let Luminance(ByVal dNewLuminance As Double)
    m_dLuminance = dNewLuminance
End Property

Public Property Get Saturation() As Double
    Saturation = m_dSaturation
End Property

Public Property Let Saturation(ByVal dNewSaturation As Double)
    m_dSaturation = dNewSaturation
End Property

Public Property Get Color() As OLE_COLOR
    Color = lHSLtoRGB(m_dHue, m_dSaturation, m_dLuminance)
End Property

Public Property Let Color(ByVal ocNewColor As OLE_COLOR)
    If (ocNewColor = (ocNewColor Or &H80000000)) Then
        ocNewColor = GetSysColor(ocNewColor Xor &H80000000)
    End If
    vRGBtoHSL ocNewColor
End Property

Private Function Max(i1 As Variant, i2 As Variant) As Variant
    Max = IIf(i1 > i2, i1, i2)
End Function

Private Function Min(i1 As Variant, i2 As Variant) As Variant
    Min = IIf(i1 < i2, i1, i2)
End Function

Private Function vRGBtoHSL(Color As Long)
    Dim lRed      As Long
    Dim lGreen    As Long
    Dim lBlue     As Long
    Dim ColorMax  As Long
    Dim ColorMin  As Long
    Dim MmM       As Long
    Dim MpM       As Long
    Dim RD        As Double
    Dim GD        As Double
    Dim BD        As Double

'   untangle colors
    lRed = (Color And &HFF&)
    lGreen = (Color And &HFF00&) / &H100
    lBlue = (Color And &HFF0000) / &H10000

    ColorMax = Max(Max(lRed, lGreen), lBlue)
    ColorMin = Min(Min(lRed, lGreen), lBlue)
    MmM = ColorMax - ColorMin
    MpM = ColorMax + ColorMin
    m_dLuminance = MpM / 2          'MpM * HSLMAX / (RGBMAX + RGBMAX) -> MpM / 2
    If ColorMax = ColorMin Then     'all color components equal -> gray
        m_dSaturation = 0           'gray -> no saturation
        m_dHue = HSLMAX * 2 / 3     'default blue for achromatic (same as Windows)
    Else
        If m_dLuminance <= HSLMAX / 2 Then
            m_dSaturation = MmM * HSLMAX / MpM
        Else
            m_dSaturation = MmM * HSLMAX / (RGBMAX + RGBMAX - MpM)
        End If
        RD = (ColorMax - lRed) * HSLMAX / MmM / 6
        GD = (ColorMax - lGreen) * HSLMAX / MmM / 6
        BD = (ColorMax - lBlue) * HSLMAX / MmM / 6
        Select Case ColorMax
            Case lRed
                m_dHue = BD - GD
            Case lGreen
                m_dHue = HSLMAX / 3 + RD - BD
            Case lBlue
                m_dHue = HSLMAX * 2 / 3 + GD - RD
        End Select
        If m_dHue < 0 Then
            m_dHue = m_dHue + HSLMAX
        End If
    End If
End Function

Private Function lHSLtoRGB(ByVal dHue As Double, ByVal dSaturation As Double, ByVal dLuminance As Double) As Long
    Dim FactA     As Double
    Dim FactB     As Double
    Dim lRed      As Long
    Dim lGreen    As Long
    Dim lBlue     As Long

    dHue = Max(0, Min(dHue, HSLMAX))
    dSaturation = Max(0, Min(dSaturation, HSLMAX))
    dLuminance = Max(0, Min(dLuminance, HSLMAX))
    If dSaturation = 0 Then                     'no saturation -> gray
        lRed = dLuminance '* RGBMAX / HSLMAX -> * 1  'luminance to all color components
        lGreen = lRed
        lBlue = lRed
    Else
        If dLuminance <= HSLMAX / 2 Then
            FactA = dLuminance * (HSLMAX + dSaturation) / HSLMAX
        Else
            FactA = dLuminance + dSaturation - dLuminance * dSaturation / HSLMAX
        End If
        FactB = dLuminance + dLuminance - FactA
        lRed = lHUEtoRGB(FactA, FactB, dHue + HSLMAX / 3) '* RGBMAX / HSLMAX -> * 1
        lGreen = lHUEtoRGB(FactA, FactB, dHue) '* RGBMAX / HSLMAX -> * 1
        lBlue = lHUEtoRGB(FactA, FactB, dHue - HSLMAX / 3) '* RGBMAX / HSLMAX -> * 1
    End If
    lHSLtoRGB = RGB(lRed, lGreen, lBlue)
End Function

Private Function lHUEtoRGB(FA As Double, FB As Double, ByVal H As Double) As Long
    Select Case H
        Case Is < 0
            H = H + HSLMAX
        Case Is > HSLMAX
            H = H - HSLMAX
    End Select
    Select Case True
        Case H < HSLMAX / 6
            lHUEtoRGB = FB + 6 * (FA - FB) * H / HSLMAX
        Case H < HSLMAX / 2
            lHUEtoRGB = FA
        Case H < HSLMAX * 2 / 3
            lHUEtoRGB = FB + 6 * (FA - FB) * (HSLMAX * 2 / 3 - H) / HSLMAX
        Case Else
            lHUEtoRGB = FB
    End Select
End Function
