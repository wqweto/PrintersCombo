Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoW" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Type LOGFONTW
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName          As String * 32 ' LF_FACESIZE
End Type

Public Property Get SystemIconFont() As StdFont
    Const SPI_GETICONTITLELOGFONT       As Long = 31
    Const FW_NORMAL                     As Long = 400
    Const LOGPIXELSY                    As Long = 90
    Dim uFont           As LOGFONTW
    Dim hTempDC         As Long
    Dim lLogPixels      As Long
    
    Call SystemParametersInfo(SPI_GETICONTITLELOGFONT, LenB(uFont), VarPtr(uFont), 0)
    Set SystemIconFont = New StdFont
    With SystemIconFont
        .Name = uFont.lfFaceName
        .Bold = (uFont.lfWeight > FW_NORMAL)
        .Charset = uFont.lfCharSet
        .Italic = (uFont.lfItalic <> 0)
        .Strikethrough = (uFont.lfStrikeOut <> 0)
        .Underline = (uFont.lfUnderline <> 0)
        .Weight = uFont.lfWeight
        hTempDC = GetDC(0)
        lLogPixels = GetDeviceCaps(hTempDC, LOGPIXELSY)
        Call ReleaseDC(0, hTempDC)
        If lLogPixels <> 0 Then
            .Size = -(uFont.lfHeight * 72#) / lLogPixels
        End If
    End With
End Property

Public Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Public Sub DebugLog(sModule As String, sFunction As String, sText As String, Optional ByVal eType As LogEventTypeConstants = vbLogEventTypeInformation)
    Debug.Print Format$(TimerEx, "0.000") & " " & Switch( _
        eType = vbLogEventTypeError, "[ERROR]", _
        eType = vbLogEventTypeWarning, "[WARN]", _
        True, "[INFO]") & " " & sText & " [" & sModule & "." & sFunction & "]"
End Sub
