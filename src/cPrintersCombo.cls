VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrintersCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' Printers ComboBox (c) 2024 by wqweto@gmail.com
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cPrintersCombo"

'=========================================================================
' API
'=========================================================================

Private Const S_OK                      As Long = 0
Private Const MAX_PATH                  As Long = 260
'--- Windows Msgs
Private Const WM_CTLCOLORBTN            As Long = &H135
'--- Window Styles
Private Const WS_DISABLED               As Long = &H8000000
'--- for SHGetSpecialFolderLocation
Private Const CSIDL_PRINTERS            As Long = &H4
'--- for SHGetFileInfo
Private Const SHGFI_LARGEICON           As Long = &H0
Private Const SHGFI_SMALLICON           As Long = &H1
Private Const SHGFI_SHELLICONSIZE       As Long = &H4
Private Const SHGFI_PIDL                As Long = &H8
Private Const SHGFI_DISPLAYNAME         As Long = &H200
Private Const SHGFI_SYSICONINDEX        As Long = &H4000
'--- for IShellFolder.EnumObjects
Private Const SHCONTF_NONFOLDERS        As Long = &H40
'--- for ImageList_Draw
Private Const ILD_TRANSPARENT           As Long = &H1
Private Const ILD_BLEND25               As Long = &H2
Private Const ILD_BLEND50               As Long = &H4
'--- for IQueryInfo.GetInfoTip
Private Const QITIPF_SINGLELINE         As Long = &H10
'--- for SHChangeNotifyRegister events
Private Const SHCNE_CREATE              As Long = &H2
Private Const SHCNE_DELETE              As Long = &H4
Private Const SHCNE_UPDATEITEM          As Long = &H2000
'--- for RegisterClipboardFormat
Private Const CFSTR_PRINTERGROUP        As String = "PrinterFriendlyName"
'--- for SetBkMode
Private Const BS_TRANSPARENT            As Long = 1
'--- for DrawText
Private Const DT_LEFT                   As Long = &H0
Private Const DT_TOP                    As Long = &H0
Private Const DT_VCENTER                As Long = &H4
Private Const DT_BOTTOM                 As Long = &H8
Private Const DT_SINGLELINE             As Long = &H20
Private Const DT_NOPREFIX               As Long = &H800

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpStr As Long) As Long
'--- user32
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
'--- gdi32
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As Long, ByVal lpMetrics As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As SIZEAPI) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'--- ole32
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Sub ReleaseStgMedium Lib "ole32" (uStgMedium As STGMEDIUM)
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
'--- shell32
Private Declare Function SHGetDesktopFolder Lib "shell32" (ppshf As stdole.IUnknown) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoW" (pszPath As Any, ByVal dwFileAttributes As Long, ByVal psfi As Long, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
'--- comctl32
Private Declare Function ImageList_Draw Lib "comctl32" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_DrawEx Lib "comctl32" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal himl As Long, cx As Long, cy As Long) As Long

Private Type SIZEAPI
    cx                  As Long
    cy                  As Long
End Type

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type SHFILEINFO
    hIcon               As Long
    iIcon               As Long
    dwAttributes        As Long
    szDisplayName       As String * MAX_PATH
    szTypeName          As String * 80
End Type

Private Type POINTAPI
    X                   As Long
    Y                   As Long
End Type

Private Type DROPFILES
    pFiles              As Long
    pt                  As POINTAPI
    fNC                 As Long
    fWide               As Long
End Type

Private Type TEXTMETRICW
    tmHeight            As Long
    tmAscent            As Long
    tmDescent           As Long
    tmInternalLeading   As Long
    tmExternalLeading   As Long
    tmAveCharWidth      As Long
    tmMaxCharWidth      As Long
    tmWeight            As Long
    tmOverhang          As Long
    tmDigitizedAspectX  As Long
    tmDigitizedAspectY  As Long
    tmFirstChar         As Integer
    tmLastChar          As Integer
    tmDefaultChar       As Integer
    tmBreakChar         As Integer
    tmItalic            As Byte
    tmUnderlined        As Byte
    tmStruckOut         As Byte
    tmPitchAndFamily    As Byte
    tmCharSet           As Byte
End Type

Private Type FORMATETC
    cfFormat            As Long
    ptd                 As Long
    dwAspect            As Long
    lIndex              As Long
    tymed               As Long
End Type

Private Type STGMEDIUM
    tymed               As Long
    pData               As Long
    pUnkForRelease      As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private IID_IShellFolder(0 To 3)    As Long
Private IID_IDataObject(0 To 3)     As Long
Private IID_IQueryInfo(0 To 3)      As Long

Private WithEvents m_oCtl       As ctxOwnerDrawCombo
Attribute m_oCtl.VB_VarHelpID = -1
Private m_cPrinters             As Collection
Private m_bEnabled              As Boolean
Private m_pPrintersFolder       As stdole.IUnknown
Private m_pidlPrinters()        As Byte
'--- UI elements
Private m_pNormalFont           As IFont
Private m_pBoldFont             As IFont
Private m_pItalicFont           As IFont
Private m_lIndent               As Long
Private m_cxLarge               As Long
Private m_cyLarge               As Long
Private m_cxSmall               As Long
Private m_cySmall               As Long
Private m_lOrigItemHeight       As Long
Private m_eTimerAction          As UcsFormTimerAction
Private m_bInUpdatePrinterInfo  As Boolean

Private Enum UcsInfoIndexes
    ucsIdxDeviceName
    ucsIdxDisplayName
    ucsIdxToolTip
    ucsIdxLargeImageList
    ucsIdxLargeIconIndex
    ucsIdxSmallImageList
    ucsIdxSmallIconIndex
End Enum

Private Enum UcsFormTimerAction '--- bitmask
    ucsTmrInitStatus = 2 ^ 0
    ucsTmrUpdateStatus = 2 ^ 1
    ucsTmrDelayUpdateStatus = 2 ^ 2
    ucsTmrAll = 2 ^ 3 - 1
End Enum

Private Enum UcsVtblIndexes
    '--- IShellFolder
    IShellFolder_EnumObjects = 4
    IShellFolder_BindToObject = 5
    IShellFolder_GetUIObjectOf = 10
    '--- IEnumIDList
    IEnumIDList_Next = 3
    '--- IDataObject
    IDataObject_GetData = 3
    '--- IQueryInfo
    IQueryInfo_GetInfoTip = 3
End Enum

'=========================================================================
' Error management
'=========================================================================

Private Function PrintError(sFunction As String) As VbMsgBoxResult
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Function

'=========================================================================
' Properties
'=========================================================================

Public Property Get Printers() As Object
    Set Printers = frPrinters
End Property

Friend Property Get frPrinters() As Collection
    Dim vInfo           As Variant
    Dim vElem           As Variant
    
    If m_cPrinters Is Nothing Then
        '--- create printers collection from Printers virtual folder
        Set m_cPrinters = New Collection
        vInfo = Array(vbNullString, vbNullString, vbNullString, 0, 0, 0, 0)
        For Each vElem In EnumPrinters
            vInfo(ucsIdxDeviceName) = vElem
            vInfo(ucsIdxDisplayName) = vElem
            m_cPrinters.Add vInfo, vInfo(ucsIdxDeviceName)
        Next
    End If
    Set frPrinters = m_cPrinters
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal bValue As Boolean)
    Const FUNC_NAME     As String = "Enabled [let]"
    Dim lItemHeight     As Long
    
    On Error GoTo EH
    m_bEnabled = bValue
    If bValue Then
        m_oCtl.ListWidth = pvCalcMaxWidth(m_cPrinters)
    Else
        m_oCtl.ListWidth = -1
    End If
    m_oCtl_MeasureItem 0, lItemHeight
    m_oCtl.ItemHeight(0) = lItemHeight
    Exit Property
EH:
    PrintError FUNC_NAME
End Property

Private Property Get pvTimerAction(ByVal eType As UcsFormTimerAction) As Boolean
    pvTimerAction = ((m_eTimerAction And eType) <> 0)
End Property

Private Property Let pvTimerAction(ByVal eType As UcsFormTimerAction, ByVal bValue As Boolean)
    Const FUNC_NAME     As String = "pvTimerAction [let]"

    On Error GoTo EH
    If bValue Then
        m_eTimerAction = (m_eTimerAction Or eType) And ucsTmrAll
        If m_eTimerAction <> 0 Then
            m_oCtl.RegisterFireOnceTimer IIf(eType = ucsTmrDelayUpdateStatus, 1000, 0)
        End If
    Else
        m_eTimerAction = (m_eTimerAction And Not eType) And ucsTmrAll
        If m_eTimerAction = 0 Then
            m_oCtl.RegisterFireOnceTimer -1
        End If
    End If

    Exit Property
EH:
    PrintError FUNC_NAME
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oCtl As Object, sDeviceName As String) As Boolean
    Const FUNC_NAME     As String = "frInit"
    Dim vElem           As Variant
    
    On Error GoTo EH
    '--- init member vars
    Set m_oCtl = oCtl
    Set m_pNormalFont = m_oCtl.Font
    Set m_pBoldFont = CloneFont(m_pNormalFont)
    m_pBoldFont.Bold = True
    Set m_pItalicFont = CloneFont(m_pNormalFont)
    m_pItalicFont.Italic = True
    m_lOrigItemHeight = GetFontHeight(m_oCtl.Font)
    m_bEnabled = True
    m_lIndent = 480 / Screen.TwipsPerPixelX
    m_cxSmall = 16
    m_cySmall = 16
    Set m_cPrinters = Printers
    '--- setup notifications
    m_oCtl.RegisterShellChange VarPtr(m_pidlPrinters(0))
    '--- fill combo
    m_oCtl.Clear
    For Each vElem In m_cPrinters
        m_oCtl.AddItem C_Str(vElem(ucsIdxDeviceName))
        If LCase$(vElem(ucsIdxDeviceName)) = LCase$(sDeviceName) Then
            m_oCtl.ListIndex = m_oCtl.ListCount - 1
        End If
    Next
    pvTimerAction(ucsTmrInitStatus) = True
    pvTimerProc
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function EnumPrinters() As Collection
    Const FUNC_NAME     As String = "EnumPrinters"
    Dim lPtr            As Long
    Dim pEnumIds        As stdole.IUnknown
    Dim pDataObject     As stdole.IUnknown
    Dim pidlCurrent()   As Byte
    Dim sDeviceName     As String
    Dim vElem           As Variant
    Dim lCount          As Long
    
    On Error GoTo EH
    Set EnumPrinters = New Collection
    If m_pPrintersFolder Is Nothing Then
        m_pidlPrinters = pvGetPrintersFolder(m_pPrintersFolder)
    End If
    If Not m_pPrintersFolder Is Nothing Then
        DispCallByVtbl m_pPrintersFolder, IShellFolder_EnumObjects, 0&, SHCONTF_NONFOLDERS, VarPtr(pEnumIds)
    End If
    If Not pEnumIds Is Nothing Then
        Do While DispCallByVtbl(pEnumIds, IEnumIDList_Next, 1&, VarPtr(lPtr), VarPtr(lCount)) = S_OK And lCount = 1
            pidlCurrent = pvToPidl(lPtr)
            '--- get DeviceName
            Set pDataObject = Nothing
            DispCallByVtbl m_pPrintersFolder, IShellFolder_GetUIObjectOf, 0&, 1&, VarPtr(VarPtr(pidlCurrent(0))), VarPtr(IID_IDataObject(0)), 0&, VarPtr(pDataObject)
            sDeviceName = pvGetDeviceName(pDataObject)
            If LenB(sDeviceName) <> 0 And sDeviceName <> "WinUtils_NewObject" And Not SearchCollection(EnumPrinters, sDeviceName) Then
                EnumPrinters.Add sDeviceName, sDeviceName
            End If
        Loop
    Else
        For Each vElem In VB.Printers
            sDeviceName = vElem.DeviceName
            If Not SearchCollection(EnumPrinters, sDeviceName) Then
                EnumPrinters.Add sDeviceName, sDeviceName
            End If
        Next
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

'= private ===============================================================

Private Function pvTimerProc() As Long
    Const FUNC_NAME     As String = "pvTimerProc"
    
    On Error GoTo EH
    '--- delay update
    If m_bInUpdatePrinterInfo Then
        pvTimerAction(m_eTimerAction) = True
        Exit Function
    End If
    If pvTimerAction(ucsTmrInitStatus) Then
        pvTimerAction(ucsTmrInitStatus) = False
        If pvUpdatePrinterInfo(False) Then
            m_oCtl.ListWidth = pvCalcMaxWidth(m_cPrinters)
            m_oCtl.Refresh
        End If
    ElseIf pvTimerAction(ucsTmrUpdateStatus Or ucsTmrDelayUpdateStatus) Then
        pvTimerAction(ucsTmrUpdateStatus Or ucsTmrDelayUpdateStatus) = False
        If pvUpdatePrinterInfo(True) Then
            m_oCtl.ListWidth = pvCalcMaxWidth(m_cPrinters)
        End If
        If Not m_oCtl Is Nothing Then
            If m_oCtl.DropDown Then
                pvTimerAction(ucsTmrDelayUpdateStatus) = True
            End If
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvGetPrintersFolder(Optional pPrintersFolder As stdole.IUnknown, Optional ByVal Flush As Boolean) As Byte()
    Const FUNC_NAME     As String = "pvGetPrintersFolder"
    Static g_pidlPrinters() As Byte
    Static g_pPrintersFolder As stdole.IUnknown
    Dim pDesktopFolder  As stdole.IUnknown
    Dim lPtr            As Long
    
    On Error GoTo EH
    If Flush Then
        Erase g_pidlPrinters
        Set g_pPrintersFolder = Nothing
    ElseIf g_pPrintersFolder Is Nothing Then
        '--- init shell folders
        Call SHGetDesktopFolder(pDesktopFolder)
        Call SHGetSpecialFolderLocation(0, CSIDL_PRINTERS, lPtr)
        g_pidlPrinters = pvToPidl(lPtr)
        DispCallByVtbl pDesktopFolder, IShellFolder_BindToObject, VarPtr(g_pidlPrinters(0)), 0&, VarPtr(IID_IShellFolder(0)), VarPtr(g_pPrintersFolder)
    End If
    Set pPrintersFolder = g_pPrintersFolder
    pvGetPrintersFolder = g_pidlPrinters
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvGetDeviceName(pDataObject As stdole.IUnknown) As String
    Const FUNC_NAME     As String = "pvGetDeviceName"
    Const ucsDvaContent As Long = 1
    Const ucsTmdHGlobal As Long = 1
    Dim uFmte           As FORMATETC
    Dim uStorage        As STGMEDIUM
    Dim uDrop           As DROPFILES
    Dim sTemp           As String
    Dim lPtr            As Long
    
    On Error GoTo EH
    If Not pDataObject Is Nothing Then
        With uFmte
            .cfFormat = RegisterClipboardFormat(CFSTR_PRINTERGROUP)
            .dwAspect = ucsDvaContent
            .lIndex = -1
            .tymed = ucsTmdHGlobal
        End With
        If DispCallByVtbl(pDataObject, IDataObject_GetData, VarPtr(uFmte), VarPtr(uStorage)) = S_OK Then
            lPtr = GlobalLock(uStorage.pData)
            If lPtr <> 0 Then
                Call CopyMemory(uDrop, ByVal lPtr, LenB(uDrop))
                If uDrop.fWide <> 0 Then
                    pvGetDeviceName = pvToString(UnsignedAdd(lPtr, uDrop.pFiles))
                ElseIf lstrlenA(UnsignedAdd(lPtr, uDrop.pFiles)) = 1 Then   '--- bug w NT4: fWide = 0
                    pvGetDeviceName = pvToString(UnsignedAdd(lPtr, uDrop.pFiles))
                Else
                    sTemp = String$(lstrlenA(UnsignedAdd(lPtr, uDrop.pFiles)), 0)
                    Call CopyMemory(ByVal sTemp, ByVal UnsignedAdd(lPtr, uDrop.pFiles), Len(sTemp))
                    pvGetDeviceName = sTemp
                End If
                Call GlobalUnlock(uStorage.pData)
            End If
            Call ReleaseStgMedium(uStorage)
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvUpdateCollection(cPrinters As Collection, vInfo As Variant) As Boolean
    Const FUNC_NAME     As String = "pvUpdateCollection"
    Dim lIdx            As Long
    Dim vElem           As Variant
    
    On Error GoTo EH
    If Not SearchCollection(cPrinters, vInfo(ucsIdxDeviceName)) Then
        Exit Function
    End If
    For Each vElem In cPrinters
        lIdx = lIdx + 1
        If LCase$(vElem(ucsIdxDeviceName)) = LCase$(vInfo(ucsIdxDeviceName)) Then
            If vElem(ucsIdxDisplayName) <> vInfo(ucsIdxDisplayName) Or _
                    vElem(ucsIdxToolTip) <> vInfo(ucsIdxToolTip) Or _
                    vElem(ucsIdxLargeImageList) <> vInfo(ucsIdxLargeImageList) Or _
                    vElem(ucsIdxLargeIconIndex) <> vInfo(ucsIdxLargeIconIndex) Or _
                    vElem(ucsIdxSmallImageList) <> vInfo(ucsIdxSmallImageList) Or _
                    vElem(ucsIdxSmallIconIndex) <> vInfo(ucsIdxSmallIconIndex) Then
                cPrinters.Remove lIdx
                If lIdx < cPrinters.Count Then
                    cPrinters.Add vInfo, vInfo(ucsIdxDeviceName), lIdx
                Else
                    cPrinters.Add vInfo, vInfo(ucsIdxDeviceName)
                End If
                pvUpdateCollection = True
            End If
            Exit Function
        End If
    Next
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvUpdatePrinterInfo(ByVal bTooltip As Boolean) As Boolean
    Const FUNC_NAME     As String = "pvUpdatePrinterInfo"
    Dim pEnumIds        As stdole.IUnknown
    Dim pDataObject     As stdole.IUnknown
    Dim pidlCurrent()   As Byte
    Dim pidlAbsolute()  As Byte
    Dim lPtr            As Long
    Dim uInfo           As SHFILEINFO
    Dim vInfo           As Variant
    Dim lItemHeight     As Long
    Dim pQueryInfo      As stdole.IUnknown
    Dim rc              As RECT
    Dim lCount          As Long
    
    On Error GoTo EH
    m_bInUpdatePrinterInfo = True
    '--- retrieve enumerator of Printers virtual folder
    If Not m_pPrintersFolder Is Nothing Then
        DispCallByVtbl m_pPrintersFolder, IShellFolder_EnumObjects, 0&, SHCONTF_NONFOLDERS, VarPtr(pEnumIds)
    End If
    If Not pEnumIds Is Nothing Then
        '--- loop printers
        Do While DispCallByVtbl(pEnumIds, IEnumIDList_Next, 1&, VarPtr(lPtr), VarPtr(lCount)) = S_OK And lCount = 1
            pidlCurrent = pvToPidl(lPtr)
            vInfo = Array(vbNullString, vbNullString, vbNullString, 0, 0, 0, 0)
            '--- get DeviceName
            Set pDataObject = Nothing
            DispCallByVtbl m_pPrintersFolder, IShellFolder_GetUIObjectOf, 0&, 1&, VarPtr(VarPtr(pidlCurrent(0))), VarPtr(IID_IDataObject(0)), 0&, VarPtr(pDataObject)
            vInfo(ucsIdxDeviceName) = pvGetDeviceName(pDataObject)
            If SearchCollection(m_cPrinters, vInfo(ucsIdxDeviceName)) Then
            If LenB(vInfo(ucsIdxDeviceName)) <> 0 Then
                '--- get tooltip
                If bTooltip Then
                    Set pQueryInfo = Nothing
                    '--- note: win98 does not support IID_IQueryInfo
                    DispCallByVtbl m_pPrintersFolder, IShellFolder_GetUIObjectOf, 0&, 1&, VarPtr(VarPtr(pidlCurrent(0))), VarPtr(IID_IQueryInfo(0)), 0&, VarPtr(pQueryInfo)
                    If Not pQueryInfo Is Nothing Then
                        lPtr = 0
                        DispCallByVtbl pQueryInfo, IQueryInfo_GetInfoTip, QITIPF_SINGLELINE, VarPtr(lPtr)
                        If lPtr <> 0 Then
                            vInfo(ucsIdxToolTip) = Replace(Replace(Replace(Replace(Replace(pvToString(lPtr), vbCrLf, ", "), vbLf, ", "), vbTab, vbNullString), "  ", " "), "  ", " ")
                            Call CoTaskMemFree(lPtr)
                        End If
                    End If
                End If
                '--- combine pidls: Printers + Current
                ReDim pidlAbsolute(0 To UBound(m_pidlPrinters) + UBound(pidlCurrent)) As Byte
                Call CopyMemory(pidlAbsolute(0), m_pidlPrinters(0), UBound(m_pidlPrinters) - 1)
                Call CopyMemory(pidlAbsolute(UBound(m_pidlPrinters) - 1), pidlCurrent(0), UBound(pidlCurrent) - 1)
                '--- retrieve info
                vInfo(ucsIdxLargeImageList) = SHGetFileInfo(pidlAbsolute(0), 0, VarPtr(uInfo), LenB(uInfo), SHGFI_PIDL Or SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX Or SHGFI_LARGEICON Or SHGFI_SHELLICONSIZE)
                vInfo(ucsIdxDisplayName) = Left$(uInfo.szDisplayName, InStr(uInfo.szDisplayName, vbNullChar) - 1)
                vInfo(ucsIdxLargeIconIndex) = uInfo.iIcon
                vInfo(ucsIdxSmallImageList) = SHGetFileInfo(pidlAbsolute(0), 0, VarPtr(uInfo), LenB(uInfo), SHGFI_PIDL Or SHGFI_SYSICONINDEX Or SHGFI_SMALLICON Or SHGFI_SHELLICONSIZE)
                vInfo(ucsIdxSmallIconIndex) = uInfo.iIcon
                If pvUpdateCollection(m_cPrinters, vInfo) Then
                    pvUpdatePrinterInfo = True
                    '--- repaint dropdown
                    If bTooltip And Not m_oCtl Is Nothing Then
                        If m_oCtl.DropDown Then
                            If m_cxLarge = 0 Then
                                Call ImageList_GetIconSize(vInfo(ucsIdxLargeImageList), m_cxLarge, m_cyLarge)
                                Call ImageList_GetIconSize(vInfo(ucsIdxSmallImageList), m_cxSmall, m_cySmall)
                            End If
                            Call GetClientRect(m_oCtl.hDropdown, rc)
                            Call InvalidateRect(m_oCtl.hDropdown, rc, 0)
                            DoEvents
                            If m_oCtl Is Nothing Then
                                pvUpdatePrinterInfo = False
                                GoTo QH
                            End If
                        End If
                    End If
                End If
            End If
            End If
        Loop
        If m_cPrinters.Count > 0 Then
            Call ImageList_GetIconSize(m_cPrinters.Item(1)(ucsIdxLargeImageList), m_cxLarge, m_cyLarge)
            Call ImageList_GetIconSize(m_cPrinters.Item(1)(ucsIdxSmallImageList), m_cxSmall, m_cySmall)
        End If
    End If
    '--- min size: 32px in small fonts, 40px in large-fonts
    m_lIndent = 480 / Screen.TwipsPerPixelX
    If m_cxLarge > m_lIndent Then
        m_lIndent = m_cxLarge
    End If
    m_oCtl_MeasureItem 0, lItemHeight
    '--- one time, otherwise combo's vertical scrollbar goes haywire
    If m_oCtl.ItemHeight(0) < lItemHeight Then
        m_oCtl.ItemHeight(0) = lItemHeight
        m_oCtl.DropdownRows = Screen.Height / Screen.TwipsPerPixelX / lItemHeight \ 2
    End If
QH:
    m_bInUpdatePrinterInfo = False
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvCalcMaxWidth(cPrinters As Collection) As Long
    Const FUNC_NAME     As String = "pvCalcMaxWidth"
    Dim vElem           As Variant
    Dim hDC             As Long
    Dim hPrevFont       As Long
    Dim uSize           As SIZEAPI
    Dim lWidth          As Long
    
    On Error GoTo EH
    pvCalcMaxWidth = -1
    If cPrinters.Count = 0 Then
        GoTo QH
    End If
    hDC = GetDC(0)
    hPrevFont = SelectObject(hDC, m_pBoldFont.hFont)
    For Each vElem In cPrinters
        If hDC <> 0 Then
            Call SelectObject(hDC, m_pBoldFont.hFont)
            Call GetTextExtentPoint32(hDC, StrPtr(vElem(ucsIdxDisplayName)), Len(vElem(ucsIdxDisplayName)), uSize)
            lWidth = uSize.cx + 8 + m_lIndent + 8
        Else
            lWidth = Len(vElem(ucsIdxDisplayName)) * m_pBoldFont.Size / 2 + 8 + m_lIndent + 8
        End If
        If lWidth > pvCalcMaxWidth Then
            pvCalcMaxWidth = lWidth
        End If
        If hDC <> 0 Then
            Call SelectObject(hDC, m_pItalicFont.hFont)
            Call GetTextExtentPoint32(hDC, StrPtr(vElem(ucsIdxToolTip)), Len(vElem(ucsIdxToolTip)), uSize)
            lWidth = 10 + uSize.cx + 8 + m_lIndent + 8
        Else
            lWidth = 10 + Len(vElem(ucsIdxToolTip)) * m_pItalicFont.Size / 2 + 8 + m_lIndent + 8
        End If
        If lWidth > pvCalcMaxWidth Then
            pvCalcMaxWidth = lWidth
        End If
    Next
    pvCalcMaxWidth = pvCalcMaxWidth + 8
QH:
    If hDC <> 0 Then
        Call SelectObject(hDC, hPrevFont)
        Call ReleaseDC(0, hDC)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvToPidl(ByVal lPtr As Long, Optional ByVal Release As Boolean = True) As Byte()
    Dim lTotal      As Long
    Dim nSize       As Integer
    Dim baRetVal()  As Byte
    
    If lPtr <> 0 Then
        Do
            Call CopyMemory(nSize, ByVal UnsignedAdd(lPtr, lTotal), 2)
            lTotal = lTotal + nSize
        Loop While nSize <> 0
        ReDim baRetVal(0 To lTotal + 1) As Byte
        Call CopyMemory(baRetVal(0), ByVal lPtr, lTotal + 2)
        If Release Then
            Call CoTaskMemFree(lPtr)
        End If
    Else
        ReDim baRetVal(0 To 1) As Byte
    End If
    pvToPidl = baRetVal
End Function

Private Function pvToString(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToString = String$(lstrlen(lPtr), 0)
        Call CopyMemory(ByVal StrPtr(pvToString), ByVal lPtr, LenB(pvToString))
    End If
End Function

Private Sub pvFillRect(ByVal hDC As Long, _
            Optional ByVal LeftX As Long, _
            Optional ByVal TopY As Long, _
            Optional ByVal RightX As Long, _
            Optional ByVal BottomY As Long, _
            Optional ByVal clrFill As Long = -1, _
            Optional ByVal hbrFill As Long)
    Dim uRect           As RECT
    
    If clrFill <> -1 Then
        hbrFill = CreateSolidBrush(TranslateColor(clrFill))
    End If
    With uRect
        .Left = LeftX
        .Top = TopY
        .Right = RightX
        .Bottom = BottomY
    End With
    Call FillRect(hDC, uRect, hbrFill)
    If hbrFill <> 0 Then
        Call DeleteObject(hbrFill)
    End If
End Sub

Private Sub pvDrawText(ByVal hDC As Long, _
            Text As String, _
            ByVal LeftX As Long, _
            ByVal TopY As Long, _
            ByVal RightX As Long, _
            ByVal BottomY As Long, _
            Optional ByVal wFormat As Long)
    Dim uRect           As RECT
    
    With uRect
        .Left = LeftX
        .Top = TopY
        .Right = RightX
        .Bottom = BottomY
    End With
    Call DrawText(hDC, StrPtr(Text), Len(Text), uRect, wFormat)
End Sub

'= shared ================================================================

Private Function CloneFont(pFont As IFont) As StdFont
    If Not pFont Is Nothing Then
        pFont.Clone CloneFont
    Else
        Set CloneFont = New StdFont
    End If
End Function

Private Function GetFontHeight(pFont As IFont) As Single
    Const FUNC_NAME     As String = "GetFontHeight"
    Dim hDC             As Long
    Dim hPrevFont       As Long
    Dim uMetric         As TEXTMETRICW
    
    On Error GoTo EH
    hDC = GetDC(0)
    hPrevFont = SelectObject(hDC, pFont.hFont)
    Call GetTextMetrics(hDC, VarPtr(uMetric))
    Call SelectObject(hDC, hPrevFont)
    Call ReleaseDC(0, hDC)
    GetFontHeight = uMetric.tmHeight
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
    On Error GoTo QH
    RetVal = pCol.Item(Index)
    '--- success
    SearchCollection = True
QH:
End Function

Private Function RemoveCollection(ByVal pCol As Object, Index As Variant)
    On Error Resume Next
    pCol.Remove Index
End Function

Private Function UnsignedAdd(ByVal lStart As Long, ByVal lIncr As Long) As Long
    UnsignedAdd = ((lStart Xor &H80000000) + lIncr) Xor &H80000000
End Function

Private Property Get HwndStyle(ByVal hWnd As Long, ByVal eStyle As Long, Optional ByVal Index As Long = -16) As Boolean
    HwndStyle = (GetWindowLong(hWnd, Index) And eStyle) = eStyle
End Property

Private Function C_Str(v As Variant) As String
    On Error Resume Next
    C_Str = CStr(v)
End Function

Private Function DispCallByVtbl(pUnk As stdole.IUnknown, ByVal lIndex As Long, ParamArray A() As Variant) As Variant
    Const CC_STDCALL    As Long = 4
    Const PTR_SIZE      As Long = 4
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    Dim hResult         As Long
    
    vParam = A
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(ObjPtr(pUnk), lIndex * PTR_SIZE, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc"
    End If
End Function

Private Function TranslateColor(ByVal clrColor As OLE_COLOR) As Long
    If clrColor = -1 Then
        TranslateColor = -1
    Else
        Call OleTranslateColor(clrColor, 0, TranslateColor)
    End If
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub m_oCtl_Click()
    Const FUNC_NAME     As String = "m_oCtl_Click"
    
    On Error GoTo EH
    m_oCtl_Change
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oCtl_Change()
    Const FUNC_NAME     As String = "m_oCtl_Change"
    Dim rc              As RECT
    
    On Error GoTo EH
    If Not m_bEnabled Then
        Exit Sub
    End If
    If m_oCtl.DropDown Then
        Call GetClientRect(m_oCtl.hDropdown, rc)
        Call InvalidateRect(m_oCtl.hDropdown, rc, 0)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oCtl_DropDown()
    Const FUNC_NAME     As String = "m_oCtl_DropDown"
    
    On Error GoTo EH
    If m_bEnabled Then
        pvTimerAction(ucsTmrUpdateStatus) = True
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oCtl_CloseUp()
    Const FUNC_NAME     As String = "m_oCtl_CloseUp"
    
    On Error GoTo EH
    If m_bEnabled Then
        pvTimerAction(ucsTmrUpdateStatus Or ucsTmrDelayUpdateStatus) = False
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oCtl_FireOnceTimer()
    Const FUNC_NAME     As String = "m_oCtl_FireOnceTimer"
    
    On Error GoTo EH
    pvTimerProc
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oCtl_MeasureItem(ByVal ItemID As Long, ItemHeight As Long)
    Const FUNC_NAME     As String = "m_oCtl_MeasureItem"
    
    On Error GoTo EH
    If Not m_bEnabled Then
        If ItemID >= 0 Then
            ItemHeight = m_lOrigItemHeight
        End If
    Else
        If ItemID < 0 Then
            If ItemHeight < m_cySmall Then
                ItemHeight = m_cySmall
            End If
        Else
            If m_lIndent + 8 > m_lOrigItemHeight * 2 Then
                ItemHeight = m_lIndent + 8
            Else
                ItemHeight = m_lOrigItemHeight * 2
            End If
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oCtl_DrawItem(ByVal ItemID As Long, ByVal hDC As Long, Left As Long, Top As Long, Right As Long, Bottom As Long)
    Const FUNC_NAME     As String = "m_oCtl_DrawItem"
    Dim vElem           As Variant
    Dim hPrevFont       As Long
    Dim rcFocus         As RECT
    Dim bSelected       As Boolean
    Dim bEnabled        As Boolean
    
    On Error GoTo EH
    Call SetBkMode(hDC, BS_TRANSPARENT)
    If Not m_bEnabled Then
        pvFillRect hDC, Left, Top, Right, Bottom, IIf(m_oCtl.SelectIndex = ItemID, vbHighlight, m_oCtl.BackColor)
        Call SetTextColor(hDC, TranslateColor(IIf(m_oCtl.SelectIndex = ItemID, vbHighlightText, m_oCtl.ForeColor)))
        pvDrawText hDC, m_oCtl.List(ItemID), Left + 4, Top, Right, Bottom, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER Or DT_NOPREFIX
        GoTo QH
    End If
    If Not SearchCollection(m_cPrinters, IIf(ItemID < 0, m_oCtl.Text, m_oCtl.List(ItemID)), RetVal:=vElem) Then
        GoTo QH
    End If
    rcFocus.Top = Top
    rcFocus.Right = Right - 1
    rcFocus.Bottom = Bottom
    bEnabled = Not HwndStyle(m_oCtl.hWnd, WS_DISABLED)
    If Bottom - Top < m_lIndent + 8 Then
        If Not bEnabled Then
            pvFillRect hDC, Left, Top, Right, Bottom, SendMessage(GetParent(m_oCtl.hWnd), WM_CTLCOLORBTN, hDC, ByVal m_oCtl.hWnd)
        Else
            pvFillRect hDC, Left, Top, Left + m_lIndent + 6, Bottom, IIf(bEnabled, m_oCtl.BackColor, vbButtonFace)
            If Not m_oCtl.DropDown And GetFocus() = m_oCtl.hWndCombo Then
                bSelected = True
                pvFillRect hDC, Left + m_lIndent + 6, Top, Right, Bottom, vbHighlight
                rcFocus.Left = Left + m_lIndent + 6
                rcFocus.Right = Right
                Call DrawFocusRect(hDC, rcFocus)
            Else
                pvFillRect hDC, Left + m_lIndent + 6, Top, Right, Bottom, IIf(bEnabled, m_oCtl.BackColor, vbButtonFace)
            End If
        End If
        hPrevFont = SelectObject(hDC, m_pNormalFont.hFont)
        Call SetTextColor(hDC, TranslateColor(IIf(bEnabled, IIf(bSelected, vbHighlightText, m_oCtl.ForeColor), vbButtonShadow)))
        pvDrawText hDC, C_Str(vElem(ucsIdxDisplayName)), Left + m_lIndent + 6 + 4, Top, Right, Bottom, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER Or DT_NOPREFIX
        If bEnabled Then
            Call ImageList_Draw(vElem(ucsIdxSmallImageList), vElem(ucsIdxSmallIconIndex), hDC, Left + (m_lIndent + 6 - m_cxSmall) \ 2, Top + (Bottom - Top - m_cySmall) \ 2, ILD_TRANSPARENT Or -bSelected * ILD_BLEND25)
        Else
            Call ImageList_DrawEx(vElem(ucsIdxSmallImageList), vElem(ucsIdxSmallIconIndex), hDC, Left + (m_lIndent + 6 - m_cxSmall) \ 2, Top + (Bottom - Top - m_cySmall) \ 2, 0, 0, -1, TranslateColor(vbButtonFace), ILD_TRANSPARENT Or ILD_BLEND50)
        End If
    Else
        pvFillRect hDC, Left, Top, Left + m_lIndent + 8, Bottom, m_oCtl.BackColor
        If m_oCtl.SelectIndex = ItemID Then
            bSelected = True
            pvFillRect hDC, Left + m_lIndent + 8, Top, Right, Bottom, vbHighlight
            rcFocus.Left = Left + m_lIndent + 8
            Call DrawFocusRect(hDC, rcFocus)
        Else
            pvFillRect hDC, Left + m_lIndent + 8, Top, Right, Bottom, m_oCtl.BackColor
        End If
        hPrevFont = SelectObject(hDC, IIf(m_oCtl.ListIndex = ItemID, m_pBoldFont.hFont, m_pNormalFont.hFont))
        Call SetTextColor(hDC, TranslateColor(IIf(bSelected, vbHighlightText, m_oCtl.ForeColor)))
        If LenB(vElem(ucsIdxToolTip)) = 0 Then
            pvDrawText hDC, C_Str(vElem(ucsIdxDisplayName)), Left + m_lIndent + 8 + 4, Top, Right, Bottom, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER Or DT_NOPREFIX
        Else
            pvDrawText hDC, C_Str(vElem(ucsIdxDisplayName)), Left + m_lIndent + 8 + 4, Top, Right, Top + (Bottom - Top) \ 2, DT_SINGLELINE Or DT_LEFT Or DT_BOTTOM Or DT_NOPREFIX
            Call SelectObject(hDC, m_pItalicFont.hFont)
            Call SetTextColor(hDC, TranslateColor(IIf(bSelected, vbHighlightText, IIf(m_oCtl.ListIndex = ItemID, m_oCtl.ForeColor, vbButtonShadow))))
            pvDrawText hDC, C_Str(vElem(ucsIdxToolTip)), Left + m_lIndent + 8 + 4 + 10, Top + (Bottom - Top) \ 2 + 1, Right, Bottom, DT_SINGLELINE Or DT_LEFT Or DT_TOP Or DT_NOPREFIX
        End If
        Call ImageList_Draw(vElem(ucsIdxLargeImageList), vElem(ucsIdxLargeIconIndex), hDC, Left + (m_lIndent + 8 - m_cxLarge) \ 2, Top + (Bottom - Top - m_cyLarge) \ 2, ILD_TRANSPARENT Or -bSelected * ILD_BLEND25)
    End If
    Call SelectObject(hDC, hPrevFont)
    Right = Left
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oCtl_ShellChange(ByVal lEvent As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)
    Const FUNC_NAME     As String = "m_oCtl_ShellChange"
    Dim pDataObject     As stdole.IUnknown
    Dim sDeviceName     As String
    Dim vElem           As Variant
    Dim pidlCurrent()   As Byte
    Dim lIdx            As Long
    Dim bDropped        As Boolean
    Dim vInfo           As Variant

    On Error GoTo EH
    Select Case lEvent
    Case SHCNE_CREATE, SHCNE_DELETE, SHCNE_UPDATEITEM
        pidlCurrent = pvToPidl(dwItem1, Release:=False)
        DispCallByVtbl m_pPrintersFolder, IShellFolder_GetUIObjectOf, 0&, 1&, VarPtr(VarPtr(pidlCurrent(UBound(m_pidlPrinters) - 1))), VarPtr(IID_IDataObject(0)), 0&, VarPtr(pDataObject)
        sDeviceName = pvGetDeviceName(pDataObject)
    End Select
    If LenB(sDeviceName) <> 0 Then
        Select Case lEvent
        Case SHCNE_DELETE
            RemoveCollection m_cPrinters, sDeviceName
        Case SHCNE_CREATE
            If Not SearchCollection(m_cPrinters, sDeviceName) Then
                vInfo = Array(vbNullString, vbNullString, vbNullString, 0, 0, 0, 0)
                vInfo(ucsIdxDeviceName) = sDeviceName
                vInfo(ucsIdxDisplayName) = sDeviceName
                m_cPrinters.Add vInfo, sDeviceName
            End If
        Case SHCNE_UPDATEITEM
            pvTimerAction(ucsTmrUpdateStatus) = True
        End Select
        If m_oCtl.ListCount < 50 Then
            bDropped = m_oCtl.DropDown
            For Each vElem In m_cPrinters
                If lIdx < m_oCtl.ListCount Then
                    If m_oCtl.List(lIdx) <> vElem(ucsIdxDeviceName) Then
                        m_oCtl.List(lIdx) = vElem(ucsIdxDeviceName)
                    End If
                Else
                    m_oCtl.AddItem C_Str(vElem(ucsIdxDeviceName))
                End If
                lIdx = lIdx + 1
            Next
            Do While lIdx < m_oCtl.ListCount
                m_oCtl.RemoveItem lIdx
            Loop
            If bDropped And m_bEnabled Then
                pvTimerAction(ucsTmrUpdateStatus) = True
            End If
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    IID_IShellFolder(0) = &H214E6   '--- {000214E6-0000-0000-C000-000000000046}
    IID_IShellFolder(2) = &HC0
    IID_IShellFolder(3) = &H46000000
    IID_IQueryInfo(0) = &H21500     '--- {00021500-0000-0000-c000-000000000046}
    IID_IQueryInfo(2) = &HC0
    IID_IQueryInfo(3) = &H46000000
    IID_IDataObject(0) = &H10E      '--- 0000010e-0000-0000-C000-000000000046
    IID_IDataObject(2) = &HC0
    IID_IDataObject(3) = &H46000000
End Sub
