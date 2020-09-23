VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ToolTip.cls

Option Explicit

Private Declare Function InitCommonControls Lib "comctl32" () As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
   ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, _
   ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, _
   ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, typRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
   ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
   ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long


Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left     As Long
    Top      As Long
    Right    As Long
    bottom   As Long
End Type

Private Type TOOLINFO
    lSize   As Long
    lFlags  As Long
    lHwnd   As Long
    lId     As Long
    typRect  As RECT
    hInst   As Long
    lpStr   As Long
    lParam  As Long
End Type

Public Enum ToolTipStyleEnum
    [TT_Normal] = 0
    [TT_Balloon] = 1
End Enum

Public Enum ToolTipTypeEnum
    [Type_None] = 0
    [Type_Info] = 1
    [Type_Warning] = 2
    [Type_Error] = 3
End Enum

Private Type LARGE_INTEGER
   LowPart   As Long
   HighPart  As Long
End Type

Private Type CurrencyType
   CurrValue As Currency
End Type


Private Const CW_USEDEFAULT         As Long = &H80000000
Private Const TOOLTIPS_CLASSA       As String = "tooltips_class32"

Private Const HWND_TOPMOST          As Long = -&H1

Private Const SWP_NOSIZE            As Long = &H1
Private Const SWP_NOMOVE            As Long = &H2
Private Const SWP_NOACTIVATE        As Long = &H10

Private Const TTF_IDISHWND          As Long = &H1
Private Const TTF_SUBCLASS          As Long = &H10

Private Const TTS_ALWAYSTIP         As Long = &H1
Private Const TTS_NOPREFIX          As Long = &H2
Private Const TTS_BALLOON           As Long = &H40

Private Const WM_USER               As Long = &H400
Private Const TTM_SETMAXTIPWIDTH    As Long = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR     As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR   As Long = (WM_USER + 20)
Private Const TTM_SETTITLE          As Long = (WM_USER + 32)
Private Const TTM_SETTITLEW         As Long = (WM_USER + 33)
Private Const TTM_ADDTOOLW          As Long = (WM_USER + 50)
Private Const TTM_DELTOOLW          As Long = (WM_USER + 51)

Dim m_ToolTipInfo As TOOLINFO
Dim m_ToolTipHWND As Long
Dim m_ToolTipStyle As Long
Dim m_ToolTipType As Long
Dim m_ToolTipForeColor As Long
Dim m_ToolTipBackColor As Long
Dim m_pt As POINTAPI
Dim m_HighResCounterOK As Boolean
Dim m_LargeInt As LARGE_INTEGER
Dim m_TickPace As Double



Private Sub Class_Initialize()
    If App.LogMode = 1 Then
        Call InitCommonControls
    End If
    m_ToolTipStyle = TT_Balloon
    m_ToolTipType = Type_None
    m_ToolTipForeColor = -1               ' Unless changed, use system default
    m_ToolTipBackColor = -1
    
    m_TickPace = GetTickPace()
    m_HighResCounterOK = (m_TickPace > 0)
End Sub



Public Sub ShowToolTip(ByVal inHWND As Long, _
                        ByVal inToolTip As String, _
                        Optional ByVal inHeader As String = "")
    Dim typRect As RECT
    Dim pt As POINTAPI
    Dim mStyle As Long
    Dim mStartTime As Double
    Dim mCurrTime As Double
    
    DoEvents
    
    GetCursorPos pt
    If m_pt.x = pt.x And m_pt.y = pt.y Then
         Exit Sub
    End If
     
      '==================================
      ' Since slipped through, process it
      '----------------------------------
      ' Remove previous if any
    ClearUp
    
    If Len(inToolTip) = 0 Then Exit Sub
    
    mStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX              ' Allow non-active and ampersand
    If m_ToolTipStyle = TT_Balloon Then mStyle = mStyle Or TTS_BALLOON
    
      ' Create tooltip window
    m_ToolTipHWND = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, mStyle, _
         CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, inHWND, 0&, App.hInstance, 0&)
                
      ' Prevent tooltip window going underneath
    SetWindowPos m_ToolTipHWND, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOMOVE
    
      ' Get the rect of the parent control
    GetClientRect inHWND, typRect
    
    With m_ToolTipInfo
        .lSize = Len(m_ToolTipInfo)
        .lFlags = TTF_SUBCLASS
        .lHwnd = inHWND
        .lId = 0
        .hInst = App.hInstance
        .lpStr = StrPtr(inToolTip)
        .typRect = typRect
    End With
    
      ' Ensure multiline in case header ""
    SendMessage m_ToolTipHWND, TTM_SETMAXTIPWIDTH, 0&, 0
    
      ' If user change default colors
    If m_ToolTipForeColor >= 0 Then SendMessage m_ToolTipHWND, TTM_SETTIPTEXTCOLOR, m_ToolTipForeColor, 0&
    If m_ToolTipBackColor >= 0 Then SendMessage m_ToolTipHWND, TTM_SETTIPBKCOLOR, m_ToolTipBackColor, 0&
    
      ' Add icon and header if any
    If m_ToolTipType > 0 Then
        If Len(inHeader) = 0 Then inHeader = " "        ' Enable icon display when header "".
    End If
    If Len(inHeader) > 0 Then SendMessage m_ToolTipHWND, TTM_SETTITLEW, m_ToolTipType, ByVal StrPtr(inHeader)
         
      ' Show it
    SendMessage m_ToolTipHWND, TTM_ADDTOOLW, 0&, m_ToolTipInfo
    
      '============================================================
      ' Tooltip is now in display, let user have due time to see it
      '------------------------------------------------------------
      ' Note down time before loop
    If m_HighResCounterOK Then
         mStartTime = GetHighResCounter
    Else
         mStartTime = CDbl(timeGetTime)
    End If
    m_pt.x = pt.x:  m_pt.y = pt.y
    Do While m_pt.x = pt.x And m_pt.y = pt.y
         ' System-displayed tooltip might be gone after a little while, to ensure that
         ' we are not being tied up here, limit the staying time in the loop.
         ' To avoid removing tooptip prematurely, set about the same or slightly
         ' longer than tooltip staying time.
        If m_HighResCounterOK Then
             If GetHighResCounter > mStartTime + 6000 Then
                  Exit Do
             End If
        Else
             mCurrTime = timeGetTime
             If mCurrTime < mStartTime Then
                  mCurrTime = mCurrTime + 86400000                 ' i.e. 24 * 60 * 60 * 60 * 1000
             End If
             If mCurrTime > mStartTime + 16000 Then
                  Exit Do
             End If
        End If
        DoEvents
        GetCursorPos pt
    Loop
    m_pt.x = pt.x:  m_pt.y = pt.y
    ClearUp
    Exit Sub
    
errHandler:
    ClearUp
End Sub



Public Property Let ToolTipStyle(ByVal inStyle As ToolTipStyleEnum)
    m_ToolTipStyle = inStyle
End Property



' Relevant only when header not ""
Public Property Let ToolTipType(ByVal inType As ToolTipTypeEnum)
    m_ToolTipType = inType
End Property



Public Property Let ToolTipForeColor(ByVal inColor As Long)
    m_ToolTipForeColor = inColor
End Property



Public Property Let ToolTipBackColor(ByVal inColor As Long)
    m_ToolTipBackColor = inColor
End Property



Private Function GetTickPace() As Double
    Dim q As Long
    Dim Freq As Long
    q = QueryPerformanceFrequency(m_LargeInt)
     ' Zero indicates that the installed hardware does not support a high-resolution performance counter.
    If q = 0 Then
         GetTickPace = 0
         Exit Function
    End If
    m_HighResCounterOK = True
    Freq = m_LargeInt.LowPart
    GetTickPace = 1000 / Freq
End Function



Private Function GetHighResCounter() As Double
    Dim q As Long
    Dim typCurr As CurrencyType
    q = QueryPerformanceCounter(m_LargeInt)
    LSet typCurr = m_LargeInt
    GetHighResCounter = typCurr.CurrValue * 10000 * m_TickPace
End Function



Public Sub ClearUp()
   If m_ToolTipHWND <> 0 Then
      Call SendMessage(m_ToolTipInfo.lHwnd, TTM_DELTOOLW, 0, m_ToolTipInfo)
      DestroyWindow m_ToolTipHWND
      m_ToolTipHWND = 0
   End If
End Sub



Private Sub Class_Terminate()
    ClearUp
End Sub



