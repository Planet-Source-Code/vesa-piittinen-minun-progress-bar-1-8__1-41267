Attribute VB_Name = "MinunPBarModule"
'SmartEdge www.vbsmart.com
'Modified for use with the Minun Progress Bar
Option Explicit
Private Const SED_OLDPROC = "SED_OLDPROC"
Private Const SED_OLDGWLSTYLE = "SED_OLDGWLSTYLE"
Private Const SED_OLDGWLEXSTYLE = "SED_OLDGWLEXSTYLE"
Private Const SED_BORDERS = "SED_BORDERS"
'API declarations:
Private Const WM_NCPAINT = &H85
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40
Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_THICKFRAME = &H40000
Private Const WS_BORDER = &H800000
Private Const WS_EX_WINDOWEDGE = &H100&
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Function pWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_NCPAINT
            pDrawBorder hwnd, wParam, GetProp(hwnd, SED_BORDERS)
        Case Else
            pWindowProc = CallWindowProc(GetProp(hwnd, SED_OLDPROC), hwnd, uMsg, wParam, lParam)
    End Select
End Function
Private Sub pDrawBorder(ByVal hwnd As Long, ByVal wParam As Long, ByVal lBorderType As munBorderStyle)
    Dim lRet As Long
    Dim lMode As Long
    Dim hdc As Long
    Dim Rec As RECT
    If lBorderType = munNone Then Exit Sub
    hdc = GetWindowDC(hwnd)
    lRet = GetWindowRect(hwnd, Rec)
    Rec.Right = Rec.Right - Rec.Left
    Rec.Bottom = Rec.Bottom - Rec.Top
    Rec.Left = 0
    Rec.Top = 0
    lMode = 0
    Select Case lBorderType
        Case munRaised
            lMode = BDR_RAISED
        Case munRaisedInner
            lMode = BDR_RAISEDINNER
        Case munSunken
            lMode = BDR_SUNKEN
        Case munSunkenOuter
            lMode = BDR_SUNKENOUTER
        Case munEtched
            lMode = BDR_SUNKENOUTER Or BDR_RAISEDINNER
        Case munBump
            lMode = BDR_SUNKENINNER Or BDR_RAISEDOUTER
    End Select
    lRet = DrawEdge(hdc, Rec, lMode, BF_RECT)
    lRet = ReleaseDC(hwnd, hdc)
End Sub
Public Function EdgeSubClass(ByVal hwnd As Long, ByVal eBorderStyle As munBorderStyle) As munBorderWidth
    Dim lRet As Long
    lRet = GetProp(hwnd, SED_OLDPROC)
    If lRet <> 0 Then
        SetWindowLong hwnd, GWL_WNDPROC, lRet
    Else
        SetProp hwnd, SED_OLDGWLSTYLE, GetWindowLong(hwnd, GWL_STYLE)
        SetProp hwnd, SED_OLDGWLEXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE)
    End If
    EdgeSubClass = pSetBorder(hwnd, eBorderStyle)
    lRet = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf pWindowProc)
    SetProp hwnd, SED_OLDPROC, lRet
    SetProp hwnd, SED_BORDERS, CLng(eBorderStyle)
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Function
Public Function EdgeUnSubClass(ByVal hwnd As Long) As Boolean
    Dim lRet As Long
    lRet = GetProp(hwnd, SED_OLDPROC)
    If lRet <> 0 Then
        lRet = SetWindowLong(hwnd, GWL_WNDPROC, lRet)
        SetWindowLong hwnd, GWL_STYLE, GetProp(hwnd, SED_OLDGWLSTYLE)
        SetWindowLong hwnd, GWL_EXSTYLE, GetProp(hwnd, SED_OLDGWLEXSTYLE)
        SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
        RemoveProp hwnd, SED_OLDPROC
        RemoveProp hwnd, SED_OLDGWLSTYLE
        RemoveProp hwnd, SED_OLDGWLEXSTYLE
        RemoveProp hwnd, SED_BORDERS
    End If
    EdgeUnSubClass = (lRet <> 0)
End Function
Private Function pSetBorder(ByVal hwnd As Long, ByVal eBorderStyle As munBorderStyle) As munBorderWidth
    Dim pWidth As munBorderWidth
    Select Case eBorderStyle
        Case munNone
            pWidth = sbwNone
        Case munRaised
            pWidth = sbwDouble
        Case munRaisedInner
            pWidth = sbwSingle
        Case munSunken
            pWidth = sbwDouble
        Case munSunkenOuter
            pWidth = sbwSingle
        Case munEtched
            pWidth = sbwDouble
        Case munBump
            pWidth = sbwDouble
    End Select
    Select Case pWidth
        Case sbwNone
            pWinStyleNeg hwnd, GWL_STYLE, WS_BORDER Or WS_THICKFRAME
            pWinStyleNeg hwnd, GWL_EXSTYLE, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
        Case sbwSingle
            pWinStyleNeg hwnd, GWL_STYLE, WS_BORDER Or WS_THICKFRAME
            pWinStyleNeg hwnd, GWL_EXSTYLE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
            pWinStyleAdd hwnd, GWL_EXSTYLE, WS_EX_STATICEDGE
        Case sbwDouble
            pWinStyleNeg hwnd, GWL_STYLE, WS_BORDER Or WS_THICKFRAME
            pWinStyleNeg hwnd, GWL_EXSTYLE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE
            pWinStyleAdd hwnd, GWL_EXSTYLE, WS_EX_CLIENTEDGE
    End Select
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    pSetBorder = pWidth
End Function
Private Sub pWinStyleAdd(ByVal hwnd As Long, ByVal lStyle As Long, ByVal lFlags As Long)
    SetWindowLong hwnd, lStyle, GetWindowLong(hwnd, lStyle) Or lFlags
End Sub
Private Sub pWinStyleNeg(ByVal hwnd As Long, ByVal lStyle As Long, ByVal lFlags As Long)
    SetWindowLong hwnd, lStyle, GetWindowLong(hwnd, lStyle) And Not lFlags
End Sub
