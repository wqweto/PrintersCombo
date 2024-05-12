VERSION 5.00
Begin VB.UserControl ctxOwnerDrawCombo 
   ClientHeight    =   648
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3336
   ScaleHeight     =   648
   ScaleWidth      =   3336
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   2112
   End
End
Attribute VB_Name = "ctxOwnerDrawCombo"
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
Private Const MODULE_NAME As String = "ctxOwnerDrawCombo"

#Const ImplSelfContained = True

'=========================================================================
' Events
'=========================================================================

Event Click()
Event Change()
Event DropDown()
Event CloseUp()
Event MeasureItem(ByVal ItemID As Long, ItemHeight As Long)
Event DrawItem(ByVal ItemID As Long, ByVal hDC As Long, Left As Long, Top As Long, Right As Long, Bottom As Long)
Event FireOnceTimer()
Event ShellChange(ByVal lEvent As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)

'=========================================================================
' API declares
'=========================================================================

'--- Windows Messages
Private Const WM_DRAWITEM                   As Long = &H2B
Private Const WM_MEASUREITEM                As Long = &H2C
Private Const WM_COMMAND                    As Long = &H111
Private Const CB_GETCURSEL                  As Long = &H147
Private Const CB_SETCURSEL                  As Long = &H14E
Private Const CB_SHOWDROPDOWN               As Long = &H14F
Private Const CB_SETITEMHEIGHT              As Long = &H153
Private Const CB_GETITEMHEIGHT              As Long = &H154
Private Const CB_SETDROPPEDWIDTH            As Long = &H160
Private Const WM_MYNOTIFY                   As Long = &H1000 + 100
'--- ComboBox Notifications
Private Const CBN_SELCHANGE                 As Long = 1
Private Const CBN_DROPDOWN                  As Long = 7
Private Const CBN_CLOSEUP                   As Long = 8
'--- Window Styles
Private Const GWL_STYLE                     As Long = -16
Private Const CBS_OWNERDRAWFIXED            As Long = &H10
Private Const CBS_HASSTRINGS                As Long = &H200&
Private Const ODS_SELECTED                  As Long = &H1
'--- for Windows Hooks
Private Const WH_CBT                        As Long = 5
Private Const HCBT_CREATEWND                As Long = 3
'--- for GetSystemMetrics
Private Const SM_CXVSCROLL                  As Long = 2
'--- for SetWindowPos
Private Const SWP_NOZORDER                  As Long = &H4
Private Const SWP_NOACTIVATE                As Long = &H10
'--- for SHChangeNotifyRegister
Private Const SHCNE_ALLEVENTS               As Long = &H7FFFFFFF
Private Const SHCNRF_ShellLevel             As Long = 2
'--- for MST
Private Const SIGN_BIT                      As Long = &H80000000
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameW" (ByVal hWnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SHChangeNotifyRegister Lib "shell32" Alias "#2" (ByVal hWnd As Long, ByVal uFlags As Long, ByVal dwEventID As Long, ByVal uMsg As Long, ByVal cItems As Long, lpps As Any) As Long
'--- for MST
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
#End If
#If ImplSelfContained Then
    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
#End If

Private Type MEASUREITEMSTRUCT
    CtlType             As Long
    CtlID               As Long
    ItemID              As Long
    ItemWidth           As Long
    ItemHeight          As Long
    ItemData            As Long
End Type

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType             As Long
    CtlID               As Long
    ItemID              As Long
    ItemAction          As Long
    ItemState           As Long
    hwndItem            As Long
    hDC                 As Long
    rcItem              As RECT
    ItemData            As Long
End Type

Private Type SHCHANGENOTIFYENTRY
    pidl                As Long
    bWatchSubFolders    As Long
End Type

Private Type SHNOTIFYSTRUCT
    dwItem1             As Long
    dwItem2             As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_pHook                 As stdole.IUnknown
Private m_pSubclass             As stdole.IUnknown
Private m_pTimer                As stdole.IUnknown
Private m_pCleanupNotify        As stdole.IUnknown
Private WithEvents m_oCombo     As VB.ComboBox
Attribute m_oCombo.VB_VarHelpID = -1
Private m_hDropdown             As Long
Private WithEvents m_oFont      As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_lListIndex            As Long
Private m_lSelectIndex          As Long
Private m_bDropped              As Boolean
Private m_lDropdownRows         As Long
Private m_lListWidth            As Long
Private m_oExt                  As Object

'=========================================================================
' Error handling
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

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_oCombo.BackColor
    PropertyChanged
End Property

Public Property Let BackColor(ByVal clrValue As OLE_COLOR)
    m_oCombo.BackColor = clrValue
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_oCombo.ForeColor
End Property

Public Property Let ForeColor(ByVal clrValue As OLE_COLOR)
    m_oCombo.ForeColor = clrValue
    PropertyChanged
End Property

Public Property Get Font() As StdFont
    If m_oFont Is Nothing Then
        Set m_oFont = New StdFont
    End If
    Set Font = m_oFont
End Property

Public Property Let Font(oValue As StdFont)
    Set m_oFont = CloneFont(oValue)
    UserControl_Resize
    PropertyChanged
End Property

Public Property Set Font(oValue As StdFont)
    Set m_oFont = oValue
    If m_oFont Is Nothing Then
        Set m_oFont = New StdFont
    End If
    Set m_oCombo.Font = m_oFont
    UserControl_Resize
    PropertyChanged
End Property

'= run-time ==============================================================

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get hWndCombo() As Long
    hWndCombo = m_oCombo.hWnd
End Property

Public Property Get hDropdown() As Long
    hDropdown = m_hDropdown
End Property

Public Property Get Text() As String
    Text = m_oCombo.Text
End Property

Property Get List(ByVal Index As Long) As String
    List = m_oCombo.List(Index)
End Property

Property Let List(ByVal Index As Long, sValue As String)
    m_oCombo.List(Index) = sValue
End Property

Public Property Get ListCount() As Long
    ListCount = m_oCombo.ListCount
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = m_lListIndex
End Property

Public Property Let ListIndex(ByVal lValue As Long)
    Call SendMessage(m_oCombo.hWnd, CB_SETCURSEL, lValue, ByVal 0)
    m_lListIndex = SendMessage(m_oCombo.hWnd, CB_GETCURSEL, 0, ByVal 0)
End Property

Public Property Get SelectIndex() As Long
    SelectIndex = m_lSelectIndex
End Property

Public Property Get DropDown() As Boolean
Attribute DropDown.VB_MemberFlags = "400"
    DropDown = m_bDropped
End Property

Public Property Let DropDown(ByVal bValue As Boolean)
    Call SendMessage(m_oCombo.hWnd, CB_SHOWDROPDOWN, IIf(bValue, 1, 0), ByVal 0)
    m_bDropped = bValue
End Property

Public Property Get DropdownRows() As Long
Attribute DropdownRows.VB_MemberFlags = "400"
    DropdownRows = m_lDropdownRows
End Property

Public Property Let DropdownRows(ByVal lValue As Long)
    Const FUNC_NAME     As String = "DropdownRows [let]"
    
    On Error GoTo EH
    If lValue > 0 Then
        Call SetWindowPos(m_oCombo.hWnd, 0, 0, 0, ScaleWidth \ Screen.TwipsPerPixelX, _
            ScaleHeight \ Screen.TwipsPerPixelY + lValue * ItemHeight(0), SWP_NOZORDER Or SWP_NOACTIVATE)
    End If
    m_lDropdownRows = lValue
    Exit Property
EH:
    PrintError FUNC_NAME
    Resume Next
End Property

Public Property Get ItemHeight(ByVal lIdx As Long) As Long
    ItemHeight = SendMessage(m_oCombo.hWnd, CB_GETITEMHEIGHT, lIdx, ByVal 0)
End Property

Public Property Let ItemHeight(ByVal lIdx As Long, ByVal lValue As Long)
    Call SendMessage(m_oCombo.hWnd, CB_SETITEMHEIGHT, lIdx, ByVal lValue)
    DropdownRows = m_lDropdownRows
End Property

Public Property Get ListWidth() As Long
    ListWidth = m_lListWidth
End Property

Public Property Let ListWidth(ByVal lValue As Long)
    Const FUNC_NAME     As String = "ListWidth [let]"
    Dim rc              As RECT
    
    On Error GoTo EH
    If lValue > 0 Then
        If m_oCombo.ListCount > DropdownRows Then
            lValue = lValue + GetSystemMetrics(SM_CXVSCROLL)
        End If
    End If
    If m_bDropped Then
        Call GetWindowRect(m_hDropdown, rc)
        If rc.Right < rc.Left + lValue Then
            rc.Right = rc.Left + lValue
        End If
        Call SetWindowPos(m_hDropdown, 0, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, SWP_NOZORDER Or SWP_NOACTIVATE)
    Else
        Call SendMessage(m_oCombo.hWnd, CB_SETDROPPEDWIDTH, lValue, ByVal 0)
    End If
    m_lListWidth = lValue
    Exit Property
EH:
    PrintError FUNC_NAME
    Resume Next
End Property

Public Property Get Extension() As Object
    Set Extension = m_oExt
End Property

'= private ===============================================================

Private Property Get pvAddressOfHookProc() As ctxOwnerDrawCombo
    Set pvAddressOfHookProc = InitAddressOfMethod(Me, 4)
End Property

Private Property Get pvAddressOfSubclassProc() As ctxOwnerDrawCombo
    Set pvAddressOfSubclassProc = InitAddressOfMethod(Me, 5)
End Property

Private Property Get pvAddressOfTimerProc() As ctxOwnerDrawCombo
    Set pvAddressOfTimerProc = InitAddressOfMethod(Me, 0)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Clear()
    m_oCombo.Clear
End Sub

Public Sub AddItem(Item As String)
    m_oCombo.AddItem Item
End Sub

Public Sub RemoveItem(ByVal Index As Long)
    m_oCombo.RemoveItem Index
End Sub

Public Sub Refresh()
    m_oCombo.Refresh
End Sub

Public Sub RegisterExtension(oExt As Object)
    Set m_oExt = oExt
    m_oExt.Init Me, Printer.DeviceName
End Sub

Public Sub RegisterFireOnceTimer(Optional ByVal Delay As Long)
    Const FUNC_NAME     As String = "RegisterFireOnceTimer"
    
    On Error GoTo EH
    If Delay >= 0 Then
        Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc, Delay)
    Else
        Set m_pTimer = Nothing
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Public Sub RegisterShellChange(ByVal lPidl As Long)
    Const FUNC_NAME     As String = "RegisterShellChange"
    Dim uNotify         As SHCHANGENOTIFYENTRY
    Dim hShellNotify    As Long
    
    On Error GoTo EH
    uNotify.pidl = lPidl
    uNotify.bWatchSubFolders = 1
    hShellNotify = SHChangeNotifyRegister(m_oCombo.hWnd, SHCNRF_ShellLevel, SHCNE_ALLEVENTS, WM_MYNOTIFY, 1, uNotify)
    Set m_pCleanupNotify = InitCleanupThunk(hShellNotify, "shell32", "#4") ' SHChangeNotifyDeregister
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

'= private ===============================================================

Private Sub pvInit()
    Const FUNC_NAME     As String = "pvInit"
    
    On Error GoTo EH
    Set m_pHook = Nothing
    Set m_pSubclass = InitSubclassingThunk(hWnd, Me, pvAddressOfSubclassProc.SubclassProc(0, 0, 0, 0, 0))
    Set m_oCombo = Combo1
    m_oCombo.Move 0, 0, ScaleWidth
    m_oCombo.Visible = True
    Extender.Height = m_oCombo.Height
    m_lListIndex = -1
    m_lSelectIndex = -1
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvMeasureItem(uMeasure As MEASUREITEMSTRUCT)
    Const FUNC_NAME     As String = "pvMeasureItem"
    Dim lNewItemHeight  As Long
    
    On Error GoTo EH
    lNewItemHeight = uMeasure.ItemHeight
    RaiseEvent MeasureItem(uMeasure.ItemID, lNewItemHeight)
    If lNewItemHeight <> uMeasure.ItemHeight Then
        ItemHeight(IIf(uMeasure.ItemID < 0, -1, 0)) = lNewItemHeight
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvDrawItem(uDraw As DRAWITEMSTRUCT)
    Const FUNC_NAME     As String = "pvDrawItem"
    
    On Error GoTo EH
    If (uDraw.ItemState And ODS_SELECTED) <> 0 Then
        m_lSelectIndex = uDraw.ItemID
    ElseIf m_lSelectIndex = uDraw.ItemID Then
        m_lSelectIndex = -1
    End If
    RaiseEvent DrawItem(uDraw.ItemID, uDraw.hDC, uDraw.rcItem.Left, uDraw.rcItem.Top, uDraw.rcItem.Right, uDraw.rcItem.Bottom)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Public Function CbtHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Attribute CbtHookProc.VB_MemberFlags = "40"
    Const FUNC_NAME     As String = "CbtHookProc"
    
    #If lParam And Handled Then '--- touch args
    #End If
    On Error GoTo EH
    If nCode = HCBT_CREATEWND Then
        Select Case VBGetClassName(wParam)
        Case "ThunderComboBox", "ThunderRT6ComboBox"
            Call SetWindowLong(wParam, GWL_STYLE, GetWindowLong(wParam, GWL_STYLE) Or CBS_OWNERDRAWFIXED Or CBS_HASSTRINGS)
        Case "ComboLBox"
            m_hDropdown = wParam
        End Select
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Attribute SubclassProc.VB_MemberFlags = "40"
    Const FUNC_NAME     As String = "SubclassProc"
    Dim uMeasure        As MEASUREITEMSTRUCT
    Dim uDraw           As DRAWITEMSTRUCT
    Dim uNotify         As SHNOTIFYSTRUCT
    
    On Error GoTo EH
    #If hWnd And wParam And lParam And Handled Then '--- touch args
    #End If
    Select Case wMsg
    Case WM_MEASUREITEM
        Call CopyMemory(uMeasure, ByVal lParam, LenB(uMeasure))
        pvMeasureItem uMeasure
        SubclassProc = 1
        Handled = True
    Case WM_DRAWITEM
        Call CopyMemory(uDraw, ByVal lParam, LenB(uDraw))
        pvDrawItem uDraw
        SubclassProc = 1
        Handled = True
    Case WM_COMMAND
        Select Case wParam \ &H10000
        Case CBN_SELCHANGE
            m_lListIndex = SendMessage(m_oCombo.hWnd, CB_GETCURSEL, 0, ByVal 0)
        Case CBN_DROPDOWN
            m_lSelectIndex = -1
            m_bDropped = True
            RaiseEvent DropDown
        Case CBN_CLOSEUP
            m_bDropped = False
            If m_lListWidth > 0 Then
                Call SendMessage(m_oCombo.hWnd, CB_SETDROPPEDWIDTH, m_lListWidth, ByVal 0)
            End If
            RaiseEvent CloseUp
        End Select
    Case WM_MYNOTIFY
        Call CopyMemory(uNotify, ByVal wParam, LenB(uNotify))
        RaiseEvent ShellChange(lParam, uNotify.dwItem1, uNotify.dwItem2)
    End Select
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function TimerProc() As Long
Attribute TimerProc.VB_MemberFlags = "40"
    RaiseEvent FireOnceTimer
End Function

'= shared ================================================================

Private Function CloneFont(pFont As IFont) As StdFont
    If Not pFont Is Nothing Then
        pFont.Clone CloneFont
    Else
        Set CloneFont = New StdFont
    End If
End Function

Private Function FontToString(oFont As StdFont) As String
    FontToString = oFont.Name & oFont.Size & oFont.Bold & oFont.Italic & oFont.Underline & oFont.Strikethrough & oFont.Weight & oFont.Charset
End Function

Private Function VBGetClassName(ByVal hWnd As Long) As String
    If hWnd <> 0 Then
        VBGetClassName = String$(1000, 0)
        Call GetClassName(hWnd, StrPtr(VBGetClassName), Len(VBGetClassName) - 1)
        VBGetClassName = Left$(VBGetClassName, InStr(VBGetClassName, vbNullChar) - 1)
    End If
End Function

'= MST ===================================================================

Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As Object
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
    If hThunk = 0 Then
        Exit Function
    End If
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitSubclassingThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgepwEBAAV1aLdCQUg8YIgz4AdC+L+oHHKBIQAIvCBQwREACri8IFSBEQAKuLwgVYERAAq4vCBYAREACruQkAAADzpYHCKBIQAFJqHP9SEFqL+IvCq7gBAAAAqzPAq4tEJAyri3QkFKWlM8Crg+8cagBX/3IM/3cM/1IYi0QkGIk4Xl+4XBIQAC1wEBAAwhAADx8Ai0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1GIsKUv9xDP9yDP9RHItUJASLClL/URQzwMIEAJBVi+yLVRj/QgT/QhiLQhg7QgR0b4tCEIXAdGiLCotBLIXAdDdS/9BaiUIIg/gBd1OFwHUJgX0MAwIAAHRGiwpS/1EwWoXAdTuLClJq8P9xJP9RKFqpAAAACHUoUjPAUFCNRCQEUI1EJARQ/3UU/3UQ/3UM/3UI/3IQ/1IUWVhahcl1E1KLCv91FP91EP91DP91CP9RIFr/ShhQUug4////WF3CGAAPHwA=" ' 9.6.2020 13:56:03
    Const THUNK_SIZE    As Long = 492
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitSubclassingThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcByOrdinal(GetModuleHandle("comctl32"), 410)             '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcByOrdinal(GetModuleHandle("comctl32"), 412)             '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcByOrdinal(GetModuleHandle("comctl32"), 413)             '--- 413 = DefSubclassProc ordinal
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitSubclassingThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitHookingThunk(ByVal idHook As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgepwEN4AV1aLdCQUg8YIgz4AdCqL+oHHTBLeAIvCBVQR3gCri8IFkBHeAKuLwgWgEd4AqzPAq7kJAAAA86WBwkwS3gBSahj/UhBai/iLwqu4AQAAAKszwKuri3QkFKWlg+8Yi0oM/0IMgWIM/wAAAI0Eyo0EyI1MiDTHAf80JLiJeQTHQQiJRCQEi8ItTBLeAAXEEd4AUMHgCAW4AAAAiUEMWMHoGAUA/+CQiUEQ/3QkEGoAUf90JBiLD/9RGIlHDItEJBiJOF5fuIAS3gAtcBDeAAUAFAAAwhAAi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1FIsK/3IM/1Eci1QkBIsKUv9RFDPAwgQAkFWL7ItVCP9CBItCEIXAdFiLCotBLIXAdCpS/9BaiUIIg/gBd0OLClL/UTBahcB1OIsKUmrw/3Ek/1EoWqkAAAAIdSVSM8BQUI1EJARQjUQkBFD/dRT/dRD/dQz/chD/UhRZWFqFyXUTUosK/3UU/3UQ/3UM/3IM/1EgWlBS6Fr///9YXcIQAJA=" ' 13.5.2020 18:24:28
    Const THUNK_SIZE    As Long = 5648
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitHookingThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetWindowsHookExA")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "UnhookWindowsHookEx")
        aParams(6) = GetProcAddress(GetModuleHandle("user32"), "CallNextHookEx")
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitHookingThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, idHook, App.ThreadID, VarPtr(aParams(0)), VarPtr(InitHookingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitFireOnceTimerThunk(pObj As Object, ByVal pfnCallback As Long, Optional Delay As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgeogEQUAV1aLdCQUg8YIgz4AdCqL+oHHDBMFAIvCBSgSBQCri8IFZBIFAKuLwgV0EgUAqzPAq7kIAAAA86WBwgwTBQBSahj/UhBai/iLwqu4AQAAAKszwKuri3QkFKWlg+8Yi0IMSCX/AAAAUItKDDsMJHULWIsPV/9RFDP/62P/QgyBYgz/AAAAjQTKjQTIjUyIMIB5EwB101jHAf80JLiJeQTHQQiJRCQEi8ItDBMFAAWgEgUAUMHgCAW4AAAAiUEMWMHoGAUA/+CQiUEQiU8MUf90JBRqAGoAiw//URiJRwiLRCQYiTheX7g8EwUALSARBQAFABQAAMIQAGaQi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1HYtCDMZAEwCLCv9yCGoA/1Eci1QkBIsKUv9RFDPAwgQAi1QkBItCEIXAdFuLCotBKIXAdCdS/9Bag/gBd0mLClL/USxahcB1PosKUmrw/3Eg/1EkWqkAAAAIdSuLClL/cghqAP9RHFr/QgQzwFBU/3IQ/1IUi1QkCMdCCAAAAABS6G////9YwhQADx8AjURAAQ==" ' 13.5.2020 18:59:12
    Const THUNK_SIZE    As Long = 5660
    Static hThunk       As Long
    Dim aParams(0 To 9) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitFireOnceTimerThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetTimer")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "KillTimer")
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(6))
        If aParams(6) <> 0 Then
            aParams(7) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(8) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitFireOnceTimerThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, 0, Delay, VarPtr(aParams(0)), VarPtr(InitFireOnceTimerThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitCleanupThunk(ByVal hHandle As Long, sModuleName As String, sProcName As String) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgepQEDwBV1aLdCQUgz4AdCeL+oHHPBE8AYvCBcwQPAGri8IFCBE8AauLwgUYETwBq7kCAAAA86WBwjwRPAFSahD/Ugxai/iLwqu4AQAAAKuLRCQMq4tEJBCrg+8Qi0QkGIk4Xl+4UBE8AS1QEDwBwhAAkItEJAiDOAB1KoN4BAB1JIF4CMAAAAB1G4F4DAAAAEZ1EotUJAT/QgSLRCQMiRAzwMIMALgCQACAwgwAkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEdRL/cgj/UgyLVCQEiwpS/1EQM8DCBAAPHwA=" ' 25.3.2019 14:03:56
    Const THUNK_SIZE    As Long = 256
    Static hThunk       As Long
    Dim aParams(0 To 1) As Long
    Dim pfnCleanup      As Long
    Dim lSize           As Long
    
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitCleanupThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(0) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(1) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        #If ImplSelfContained Then
            pvThunkGlobalData("InitCleanupThunk") = hThunk
        #End If
    End If
    If Left$(sProcName, 1) = "#" Then
        pfnCleanup = GetProcByOrdinal(GetModuleHandle(sModuleName), Mid$(sProcName, 2))
    Else
        pfnCleanup = GetProcAddress(GetModuleHandle(sModuleName), sProcName)
    End If
    If pfnCleanup <> 0 Then
        lSize = CallWindowProc(hThunk, hHandle, pfnCleanup, VarPtr(aParams(0)), VarPtr(InitCleanupThunk))
        Debug.Assert lSize = THUNK_SIZE
    End If
End Function

Private Function pvThunkIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvThunkIdeOwner = True
End Function

Private Function pvThunkAllocate(sText As String, Optional ByVal Size As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    pvThunkAllocate = VirtualAlloc(0, IIf(Size > 0, Size, (Len(sText) \ 4) * 3), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pvThunkAllocate = 0 Then
        Exit Function
    End If
    '--- init decoding maps
    If Map(65) = 0 Then
        baInput = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
        For lIdx = 0 To UBound(baInput)
            lChar = baInput(lIdx)
            Map(&H0 + lChar) = lIdx * (2 ^ 2)
            Map(&H100 + lChar) = (lIdx And &H30) \ (2 ^ 4) Or (lIdx And &HF) * (2 ^ 12)
            Map(&H200 + lChar) = (lIdx And &H3) * (2 ^ 22) Or (lIdx And &H3C) * (2 ^ 6)
            Map(&H300 + lChar) = lIdx * (2 ^ 16)
        Next
    End If
    '--- base64 decode loop
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    lPtr = pvThunkAllocate
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, lChar, 3)
        lPtr = (lPtr Xor SIGN_BIT) + 3 Xor SIGN_BIT
    Next
End Function

#If ImplSelfContained Then
Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property
#End If

'=========================================================================
' Control Events
'=========================================================================

Private Sub m_oCombo_Click()
    RaiseEvent Click
End Sub

Private Sub m_oCombo_Change()
    RaiseEvent Change
End Sub

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    Set m_oCombo.Font = m_oFont
    UserControl_Resize
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub UserControl_Resize()
    If Not m_oCombo Is Nothing Then
        m_oCombo.Width = ScaleWidth
        Extender.Height = m_oCombo.Height
    End If
End Sub

Private Sub UserControl_InitProperties()
    Const FUNC_NAME     As String = "UserControl_InitProperties"
    
    On Error GoTo EH
    pvInit
    BackColor = vbWindowBackground
    ForeColor = vbWindowText
    Set Font = Ambient.Font
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"
    
    On Error GoTo EH
    pvInit
    BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_WriteProperties"
    
    On Error GoTo EH
    PropBag.WriteProperty "BackColor", BackColor, vbWindowBackground
    PropBag.WriteProperty "ForeColor", ForeColor, vbWindowText
    If FontToString(Font) <> FontToString(Ambient.Font) Then
        PropBag.WriteProperty "Font", Font, Ambient.Font
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_Initialize()
    Const FUNC_NAME     As String = "UserControl_Initialize"
    
    On Error GoTo EH
    Set m_pHook = InitHookingThunk(WH_CBT, Me, pvAddressOfHookProc.CbtHookProc(0, 0, 0, 0))
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub
