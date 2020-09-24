VERSION 5.00
Begin VB.UserControl MinunPBar 
   Alignable       =   -1  'True
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   ScaleHeight     =   270
   ScaleWidth      =   4710
   ToolboxBitmap   =   "MinunPBar.ctx":0000
   Begin VB.Timer Scroller 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox DG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   311
      TabIndex        =   2
      Top             =   360
      Width           =   4695
      Visible         =   0   'False
   End
   Begin VB.PictureBox BG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox FG 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   15
      End
   End
End
Attribute VB_Name = "MinunPBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Minun Progress Bar version 1.8 - My First ActiveX, made in a week

'If you make changes, change UserControl name, project name...everything!
'Just to make sure your modified control doesn't mess up with the original :)
'It'll be best for both. If you find a "critical" bug, please mail me.

'The control is freeware, but it'd be nice to know if you spread your own
'version that is based on this :) It'd be also nice to know if you use the
'control in your project(s).

'I didn't comment the code...sorry about that.

'Last update on 3rd of December 2002 - Merri, merry@mbnet.fi
'(project started in 14th November 2002)
Option Explicit
Public Enum munBorderStyle
    munNone
    munSunken
    munSunkenOuter
    munRaised
    munRaisedInner
    munBump
    munEtched
End Enum
Public Enum munBorderWidth
    sbwNone
    sbwSingle
    sbwDouble
End Enum
Public Enum munFadeStyle
    munStill
    munMoving
End Enum
Public Enum munPercentAlign
    munCenter
    munBarOut
    munBarIn
    munCenterOut
    munCenterIn
    munLeft
    munRight
    munScrollRight
    munScrollLeft
    munScrollDown
    munScrollUp
    munBounceHorizontal
    munBounceHorizontalIn
    munBounceVertical
    munBounceVerticalIn
End Enum
'Constant Declarations:
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
'Default Property Values:
Const m_def_BackColor = &H8000000E
Const m_def_BarStyle = 0
Const m_def_BorderStyle = 2
Const m_def_Custom = 0
Const m_def_CustomText = "Done!"
Const m_def_Fade = False
Const m_def_FadeBG1 = &H8000000E
Const m_def_FadeBG2 = &H8000000E
Const m_def_FadeFG1 = &H8000000E
Const m_def_FadeFG2 = &H8000000D
Const m_def_FadeStyle = 0
Const m_def_ForeColor = &H8000000D
Const m_def_Interval = 1
Const m_def_ManualRefresh = False
Const m_def_Max = 10000
Const m_def_Min = 0
Const m_def_NoPercent = False
Const m_def_Percent = 0
Const m_def_PercentAfter = "%"
Const m_def_PercentAlign = 0
Const m_def_PercentBefore = ""
Const m_def_Reverse = False
Const m_def_ScaleMode = vbTwips
Const m_def_Vertical = False
'Property Variables:
Dim m_BackColor As Long
Dim m_BarStyle As munBorderStyle
Dim m_BorderStyle As munBorderStyle
Dim m_Custom As Boolean
Dim m_CustomText As String
Dim m_Fade As Boolean
Dim m_FadeBG1 As Long
Dim m_FadeBG2 As Long
Dim m_FadeFG1 As Long
Dim m_FadeFG2 As Long
Dim m_FadeStyle As Byte
Dim m_Font As Font
Dim m_ForeColor As Long
Dim m_Interval As Integer
Dim m_ManualRefresh As Boolean
Dim m_Max As Currency
Dim m_Min As Currency
Dim m_NoPercent As Boolean
Dim m_Percent As Byte
Dim m_PercentAfter As String
Dim m_PercentAlign As Byte
Dim m_PercentBefore As String
Dim m_Reverse As Boolean
Dim m_ScaleMode As Integer
Dim m_Value As Currency
Dim m_Vertical As Boolean
'Internal Variables:
Dim m_Down As Byte
Dim m_BGborder As munBorderStyle
Dim m_FGborder As munBorderStyle
Dim m_Scroll As Integer
Dim m_OldPercent As Byte
Dim Text As String
Dim Temp As Currency, TempMax As Currency
'API Declarations:
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Event Declarations:
Event BarResize()
Attribute BarResize.VB_Description = "Occurs when the progress bar of a control has resized."
Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses the mouse button."
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Event Resize()
Attribute Resize.VB_Description = "Occurs when the size of a control has changed."
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets object's background color."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    FG.ForeColor = New_BackColor
    BG.BackColor = New_BackColor
    PropertyChanged "BackColor"
    Draw
End Property
Public Property Get BarStyle() As munBorderStyle
Attribute BarStyle.VB_Description = "Sets object's barstyle."
Attribute BarStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarStyle = m_BarStyle
End Property
Public Property Let BarStyle(ByVal New_BarStyle2 As munBorderStyle)
    Dim New_BarStyle As munBorderStyle
    If New_BarStyle2 > 6 Then
        New_BarStyle = 6
    Else
        New_BarStyle = New_BarStyle2
    End If
    m_BarStyle = New_BarStyle
    EdgeUnSubClass FG.hwnd
    m_FGborder = EdgeSubClass(FG.hwnd, New_BarStyle)
    PropertyChanged "BarStyle"
    Draw
    If Not m_ManualRefresh Then Refresh
End Property
Public Property Get BorderStyle() As munBorderStyle
Attribute BorderStyle.VB_Description = "Sets object's borderstyle."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle2 As munBorderStyle)
    Dim New_BorderStyle As munBorderStyle
    If New_BorderStyle2 > 6 Then
        New_BorderStyle = 6
    Else
        New_BorderStyle = New_BorderStyle2
    End If
    m_BorderStyle = New_BorderStyle
    EdgeUnSubClass BG.hwnd
    m_BGborder = EdgeSubClass(BG.hwnd, New_BorderStyle)
    PropertyChanged "BorderStyle"
    m_Down = Abs(m_BorderStyle = munNone Or m_BorderStyle = munSunken Or m_BorderStyle = munSunkenOuter)
    Draw
    If Not m_ManualRefresh Then Refresh
End Property
Public Property Get Custom() As Boolean
Attribute Custom.VB_Description = "View custom ""Done!"" text?"
Attribute Custom.VB_ProcData.VB_Invoke_Property = "Settings"
    Custom = m_Custom
End Property
Public Property Let Custom(ByVal New_Custom As Boolean)
    m_Custom = New_Custom
    PropertyChanged "Custom"
    If Not m_ManualRefresh Then DrawFade
End Property
Public Property Get CustomText() As String
Attribute CustomText.VB_Description = "Sets object's custom ""Done!"" text."
Attribute CustomText.VB_ProcData.VB_Invoke_Property = "Settings"
    CustomText = m_CustomText
End Property
Public Property Let CustomText(ByVal New_CustomText As String)
    m_CustomText = New_CustomText
    PropertyChanged "CustomText"
    If Not m_ManualRefresh Then DrawFade
End Property
Private Sub Draw()
    Dim A As Integer, Col(5) As Single, Cols(2) As Single, Inc(2) As Single
    Dim Color1 As Long, Color2 As Long, Color3 As Long, Color4 As Long
    BG.Cls
    FG.Cls
    If m_FadeBG1 < 0 Or m_FadeBG1 > RGB(255, 255, 255) Then
        DG.BackColor = m_FadeBG1
        Color1 = GetPixel(DG.hdc, 1, 1)
    Else
        Color1 = m_FadeBG1
    End If
    If m_FadeBG2 < 0 Or m_FadeBG2 > RGB(255, 255, 255) Then
        DG.BackColor = m_FadeBG2
        Color2 = GetPixel(DG.hdc, 1, 1)
    Else
        Color2 = m_FadeBG2
    End If
    If m_FadeFG1 < 0 Or m_FadeFG1 > RGB(255, 255, 255) Then
        DG.BackColor = m_FadeFG1
        Color3 = GetPixel(DG.hdc, 1, 1)
    Else
        Color3 = m_FadeFG1
    End If
    If m_FadeFG2 < 0 Or m_FadeFG2 > RGB(255, 255, 255) Then
        DG.BackColor = m_FadeFG2
        Color4 = GetPixel(DG.hdc, 1, 1)
    Else
        Color4 = m_FadeFG2
    End If
    If m_Fade And Not m_Vertical Then
        'background color
        Col(0) = Color1 \ 65536
        Col(1) = (Color1 Mod 65536) \ 256
        Col(2) = Color1 Mod 256
        Col(3) = Color2 \ 65536
        Col(4) = (Color2 Mod 65536) \ 256
        Col(5) = Color2 Mod 256
        Cols(0) = Col(0)
        Cols(1) = Col(1)
        Cols(2) = Col(2)
        Inc(0) = (Col(3) - Col(0)) / BG.ScaleWidth
        Inc(1) = (Col(4) - Col(1)) / BG.ScaleWidth
        Inc(2) = (Col(5) - Col(2)) / BG.ScaleWidth
        For A = 0 To BG.ScaleWidth
            Call SetPixel(DG.hdc, A, 0, RGB(CByte(Cols(2)), CByte(Cols(1)), CByte(Cols(0))))
            Cols(0) = Cols(0) + Inc(0)
            Cols(1) = Cols(1) + Inc(1)
            Cols(2) = Cols(2) + Inc(2)
            If Cols(0) < 0 Then Cols(0) = 0
            If Cols(1) < 0 Then Cols(1) = 0
            If Cols(2) < 0 Then Cols(2) = 0
            If Cols(0) > 255 Then Cols(0) = 255
            If Cols(1) > 255 Then Cols(1) = 255
            If Cols(2) > 255 Then Cols(2) = 255
        Next A
        StretchBlt DG.hdc, 0, 0, BG.ScaleWidth, BG.ScaleHeight, DG.hdc, 0, 0, BG.ScaleWidth, 1, vbSrcCopy
        'foreground color
        Col(0) = Color3 \ 65536
        Col(1) = (Color3 Mod 65536) \ 256
        Col(2) = Color3 Mod 256
        Col(3) = Color4 \ 65536
        Col(4) = (Color4 Mod 65536) \ 256
        Col(5) = Color4 Mod 256
        Cols(0) = Col(0)
        Cols(1) = Col(1)
        Cols(2) = Col(2)
        Inc(0) = (Col(3) - Col(0)) / (BG.ScaleWidth + 1)
        Inc(1) = (Col(4) - Col(1)) / (BG.ScaleWidth + 1)
        Inc(2) = (Col(5) - Col(2)) / (BG.ScaleWidth + 1)
        For A = 0 To (BG.ScaleWidth + 1)
            Call SetPixel(DG.hdc, A, BG.ScaleHeight + 1, RGB(CByte(Cols(2)), CByte(Cols(1)), CByte(Cols(0))))
            Cols(0) = Cols(0) + Inc(0)
            Cols(1) = Cols(1) + Inc(1)
            Cols(2) = Cols(2) + Inc(2)
            If Cols(0) < 0 Then Cols(0) = 0
            If Cols(1) < 0 Then Cols(1) = 0
            If Cols(2) < 0 Then Cols(2) = 0
            If Cols(0) > 255 Then Cols(0) = 255
            If Cols(1) > 255 Then Cols(1) = 255
            If Cols(2) > 255 Then Cols(2) = 255
        Next A
        StretchBlt DG.hdc, 0, BG.ScaleHeight + 1, BG.ScaleWidth + 1, BG.ScaleHeight, DG.hdc, 0, BG.ScaleHeight + 1, BG.ScaleWidth + 1, 1, vbSrcCopy
    ElseIf m_Fade And m_Vertical Then
        'background color
        Col(0) = Color1 \ 65536
        Col(1) = (Color1 Mod 65536) \ 256
        Col(2) = Color1 Mod 256
        Col(3) = Color2 \ 65536
        Col(4) = (Color2 Mod 65536) \ 256
        Col(5) = Color2 Mod 256
        Cols(0) = Col(0)
        Cols(1) = Col(1)
        Cols(2) = Col(2)
        Inc(0) = (Col(3) - Col(0)) / BG.ScaleHeight
        Inc(1) = (Col(4) - Col(1)) / BG.ScaleHeight
        Inc(2) = (Col(5) - Col(2)) / BG.ScaleHeight
        For A = 0 To BG.ScaleHeight
            Call SetPixel(DG.hdc, 0, A, RGB(CByte(Cols(2)), CByte(Cols(1)), CByte(Cols(0))))
            Cols(0) = Cols(0) + Inc(0)
            Cols(1) = Cols(1) + Inc(1)
            Cols(2) = Cols(2) + Inc(2)
            If Cols(0) < 0 Then Cols(0) = 0
            If Cols(1) < 0 Then Cols(1) = 0
            If Cols(2) < 0 Then Cols(2) = 0
            If Cols(0) > 255 Then Cols(0) = 255
            If Cols(1) > 255 Then Cols(1) = 255
            If Cols(2) > 255 Then Cols(2) = 255
        Next A
        StretchBlt DG.hdc, 0, 0, BG.ScaleWidth, BG.ScaleHeight, DG.hdc, 0, 0, 1, BG.ScaleHeight, vbSrcCopy
        'foreground color
        Col(0) = Color3 \ 65536
        Col(1) = (Color3 Mod 65536) \ 256
        Col(2) = Color3 Mod 256
        Col(3) = Color4 \ 65536
        Col(4) = (Color4 Mod 65536) \ 256
        Col(5) = Color4 Mod 256
        Cols(0) = Col(0)
        Cols(1) = Col(1)
        Cols(2) = Col(2)
        Inc(0) = (Col(3) - Col(0)) / (BG.ScaleHeight + 1)
        Inc(1) = (Col(4) - Col(1)) / (BG.ScaleHeight + 1)
        Inc(2) = (Col(5) - Col(2)) / (BG.ScaleHeight + 1)
        For A = 0 To (BG.ScaleHeight + 1)
            Call SetPixel(DG.hdc, 0, BG.ScaleHeight + 1 + A, RGB(CByte(Cols(2)), CByte(Cols(1)), CByte(Cols(0))))
            Cols(0) = Cols(0) + Inc(0)
            Cols(1) = Cols(1) + Inc(1)
            Cols(2) = Cols(2) + Inc(2)
            If Cols(0) < 0 Then Cols(0) = 0
            If Cols(1) < 0 Then Cols(1) = 0
            If Cols(2) < 0 Then Cols(2) = 0
            If Cols(0) > 255 Then Cols(0) = 255
            If Cols(1) > 255 Then Cols(1) = 255
            If Cols(2) > 255 Then Cols(2) = 255
            StretchBlt DG.hdc, 0, BG.ScaleHeight + 1, BG.ScaleWidth, BG.ScaleHeight + 1, DG.hdc, 0, BG.ScaleHeight + 1, 1, BG.ScaleHeight + 1, vbSrcCopy
        Next A
    End If
    If Not m_ManualRefresh Then DrawFade
End Sub
Private Sub DrawFade()
    Dim A As Integer
    TextRefresh
    If m_Fade And Not m_Vertical Then
        A = FG.Width + FG.Left
        Select Case m_FadeStyle
            Case 1
                BitBlt BG.hdc, A, 0, BG.ScaleWidth - A, BG.ScaleHeight, DG.hdc, 0, 0, vbSrcCopy
                BitBlt FG.hdc, 0, 0, FG.ScaleWidth, BG.ScaleHeight, DG.hdc, BG.ScaleWidth - FG.ScaleWidth, BG.ScaleHeight + 1, vbSrcCopy
            Case Else
                BitBlt BG.hdc, A, 0, BG.ScaleWidth - A, BG.ScaleHeight, DG.hdc, A, 0, vbSrcCopy
                BitBlt FG.hdc, 0, 0, FG.ScaleWidth, BG.ScaleHeight, DG.hdc, 0, BG.ScaleHeight + 1, vbSrcCopy
        End Select
    ElseIf m_Fade And m_Vertical Then
        A = FG.Height + FG.Top
        Select Case m_FadeStyle
            Case 1
                BitBlt BG.hdc, 0, A, BG.ScaleWidth, BG.ScaleHeight - A, DG.hdc, 0, 0, vbSrcCopy
                BitBlt FG.hdc, 0, 0, BG.ScaleWidth, FG.ScaleHeight, DG.hdc, 0, BG.ScaleHeight * 2 - FG.ScaleHeight + 1, vbSrcCopy
            Case Else
                BitBlt BG.hdc, 0, A, BG.ScaleWidth, BG.ScaleHeight - A, DG.hdc, 0, A, vbSrcCopy
                BitBlt FG.hdc, 0, 0, BG.ScaleWidth, FG.ScaleHeight, DG.hdc, 0, BG.ScaleHeight + 1, vbSrcCopy
        End Select
    End If
    If Text <> "" Then
        If Not m_Fade Then BG.Cls: FG.Cls
        If m_PercentAlign < 7 Then
            If Not m_Vertical Then
                BG.CurrentY = (BG.ScaleHeight - BG.TextHeight(Text)) / 2 + m_Down
                FG.CurrentY = BG.CurrentY - m_FGborder - FG.Top
                Select Case m_PercentAlign
                    Case 1
                        BG.CurrentX = FG.Width + m_Down
                        BG.Print Text
                    Case 2
                        FG.CurrentX = FG.ScaleWidth - FG.TextWidth(Text) + m_Down
                        FG.Print Text
                    Case 3
                        BG.CurrentX = FG.Width + ((BG.ScaleWidth - FG.Width) - BG.TextWidth(Text)) / 2 + m_Down
                        BG.Print Text
                    Case 4
                        FG.CurrentX = (FG.ScaleWidth - FG.TextWidth(Text)) / 2 + m_Down
                        FG.Print Text
                    Case 5
                        BG.CurrentX = m_Down
                        FG.CurrentX = BG.CurrentX - m_FGborder - FG.Left
                        BG.Print Text
                        FG.Print Text
                    Case 6
                        BG.CurrentX = BG.ScaleWidth - BG.TextWidth(Text) + m_Down
                        FG.CurrentX = BG.CurrentX - m_FGborder - FG.Left
                        BG.Print Text
                        FG.Print Text
                    Case Else
                        BG.CurrentX = (BG.ScaleWidth - BG.TextWidth(Text)) / 2 + m_Down
                        FG.CurrentX = BG.CurrentX - m_FGborder - FG.Left
                        BG.Print Text
                        FG.Print Text
                End Select
            Else
                BG.CurrentX = (BG.ScaleWidth - BG.TextWidth(Text)) / 2 + m_Down
                FG.CurrentX = BG.CurrentX - m_FGborder - FG.Left
                Select Case m_PercentAlign
                    Case 1
                        BG.CurrentY = FG.Height + m_Down
                        BG.Print Text
                    Case 2
                        FG.CurrentY = FG.ScaleHeight - FG.TextHeight(Text) + m_Down
                        FG.Print Text
                    Case 3
                        BG.CurrentY = FG.Height + ((BG.ScaleHeight - FG.Height) - BG.TextHeight(Text)) / 2 + m_Down
                        BG.Print Text
                    Case 4
                        FG.CurrentY = (FG.ScaleHeight - FG.TextHeight(Text)) / 2 + m_Down
                        FG.Print Text
                    Case 5
                        BG.CurrentY = m_Down
                        FG.CurrentY = BG.CurrentY - m_FGborder - FG.Top
                        BG.Print Text
                        FG.Print Text
                    Case 6
                        BG.CurrentY = BG.ScaleHeight - BG.TextHeight(Text) + m_Down
                        FG.CurrentY = BG.CurrentY - m_FGborder - FG.Top
                        BG.Print Text
                        FG.Print Text
                    Case Else
                        BG.CurrentY = (BG.ScaleHeight - BG.TextHeight(Text)) / 2 + m_Down
                        FG.CurrentY = BG.CurrentY - m_FGborder - FG.Top
                        BG.Print Text
                        FG.Print Text
                End Select
            End If
        Else
            Select Case PercentAlign
                Case 7, 8, 11, 12
                    BG.CurrentY = (BG.ScaleHeight - BG.TextHeight(Text)) / 2 + m_Down
                    FG.CurrentY = BG.CurrentY - m_FGborder - FG.Top
                    BG.CurrentX = m_Scroll + m_Down
                    FG.CurrentX = BG.CurrentX - m_FGborder
                    BG.Print Text
                    FG.Print Text
                Case 9, 10, 13, 14
                    BG.CurrentX = (BG.ScaleWidth - BG.TextWidth(Text)) / 2 + m_Down
                    FG.CurrentX = BG.CurrentX - m_FGborder - FG.Left
                    BG.CurrentY = m_Scroll + m_Down
                    FG.CurrentY = BG.CurrentY - m_FGborder
                    BG.Print Text
                    FG.Print Text
            End Select
        End If
    End If
    BG.Refresh
    FG.Refresh
End Sub
Public Property Get Fade() As Boolean
Attribute Fade.VB_Description = "Show colorfade?"
Attribute Fade.VB_ProcData.VB_Invoke_Property = "Settings"
    Fade = m_Fade
End Property
Public Property Let Fade(ByVal New_Fade As Boolean)
    m_Fade = New_Fade
    PropertyChanged "Fade"
    Draw
End Property
Public Property Get FadeBG1() As OLE_COLOR
Attribute FadeBG1.VB_Description = "Sets object's fade color one for background."
Attribute FadeBG1.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeBG1 = m_FadeBG1
End Property
Public Property Let FadeBG1(ByVal New_FadeBG1 As OLE_COLOR)
    m_FadeBG1 = New_FadeBG1
    PropertyChanged "FadeBG1"
    Draw
End Property
Public Property Get FadeBG2() As OLE_COLOR
Attribute FadeBG2.VB_Description = "Sets object's fade color two for background."
Attribute FadeBG2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeBG2 = m_FadeBG2
End Property
Public Property Let FadeBG2(ByVal New_FadeBG2 As OLE_COLOR)
    m_FadeBG2 = New_FadeBG2
    PropertyChanged "FadeBG2"
    Draw
End Property
Public Property Get FadeFG1() As OLE_COLOR
Attribute FadeFG1.VB_Description = "Sets object's fade color one for foreground."
Attribute FadeFG1.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeFG1 = m_FadeFG1
End Property
Public Property Let FadeFG1(ByVal New_FadeFG1 As OLE_COLOR)
    m_FadeFG1 = New_FadeFG1
    PropertyChanged "FadeFG1"
    Draw
End Property
Public Property Get FadeFG2() As OLE_COLOR
Attribute FadeFG2.VB_Description = "Sets object's fade color two for foreground."
Attribute FadeFG2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeFG2 = m_FadeFG2
End Property
Public Property Let FadeFG2(ByVal New_FadeFG2 As OLE_COLOR)
    m_FadeFG2 = New_FadeFG2
    PropertyChanged "FadeFG2"
    Draw
End Property
Public Property Get FadeStyle() As munFadeStyle
Attribute FadeStyle.VB_Description = "Sets object's color fading style."
Attribute FadeStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FadeStyle = m_FadeStyle
End Property
Public Property Let FadeStyle(ByVal New_FadeStyle As munFadeStyle)
    If New_FadeStyle > 1 Then New_FadeStyle = 1
    m_FadeStyle = New_FadeStyle
    PropertyChanged "FadeStyle"
    Draw
End Property
Public Property Get Font() As Font
Attribute Font.VB_Description = "Sets object's font."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    Set BG.Font = New_Font
    Set FG.Font = New_Font
    PropertyChanged "Font"
    If Not m_ManualRefresh Then DrawFade
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Sets object's foreground color."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    BG.ForeColor = New_ForeColor
    FG.BackColor = New_ForeColor
    PropertyChanged "ForeColor"
    Draw
End Property
Private Sub InitPercentAlign()
    TextRefresh
    Select Case m_PercentAlign
        Case 7, 11
            m_Scroll = -BG.TextWidth(Text)
        Case 8
            m_Scroll = BG.ScaleWidth
        Case 9, 13
            m_Scroll = -BG.TextHeight(Text)
        Case 11
            m_Scroll = BG.ScaleHeight
        Case 12, 14
            m_Scroll = 0
    End Select
End Sub
Public Property Get Interval() As Integer
    Interval = m_Interval
End Property
Public Property Let Interval(ByVal New_Interval As Integer)
    m_Interval = New_Interval
    Scroller.Interval = New_Interval
    PropertyChanged "Interval"
End Property
Public Property Get ManualRefresh() As Boolean
Attribute ManualRefresh.VB_Description = "Refresh must be invoked manually through code?"
Attribute ManualRefresh.VB_ProcData.VB_Invoke_Property = "Settings"
    ManualRefresh = m_ManualRefresh
End Property
Public Property Let ManualRefresh(ByVal New_ManualRefresh As Boolean)
    m_ManualRefresh = New_ManualRefresh
    PropertyChanged "ManualRefresh"
End Property
Public Property Get Max() As Currency
Attribute Max.VB_Description = "Sets object's max value."
Attribute Max.VB_ProcData.VB_Invoke_Property = "Settings"
    Max = m_Max
End Property
Public Property Let Max(ByVal New_Max As Currency)
    If New_Max < m_Min + 1 Then New_Max = m_Min + 1
    m_Max = New_Max
    PropertyChanged "Max"
    If m_Max < m_Value Then m_Value = m_Max
    If Not m_ManualRefresh Then Refresh: DrawFade
End Property
Public Property Get Min() As Currency
Attribute Min.VB_Description = "Sets object's minimal value."
Attribute Min.VB_ProcData.VB_Invoke_Property = "Settings"
    Min = m_Min
End Property
Public Property Let Min(ByVal New_Min As Currency)
    If New_Min > m_Max - 1 Then New_Min = m_Max - 1
    m_Min = New_Min
    PropertyChanged "Min"
    If m_Min > m_Value Then m_Value = m_Min
    If Not m_ManualRefresh Then Refresh: DrawFade
End Property
Public Property Get NoPercent() As Boolean
Attribute NoPercent.VB_Description = "Don't show percentage?"
Attribute NoPercent.VB_ProcData.VB_Invoke_Property = "Settings"
    NoPercent = m_NoPercent
End Property
Public Property Let NoPercent(ByVal New_NoPercent As Boolean)
    m_NoPercent = New_NoPercent
    PropertyChanged "NoPercent"
    If Not m_ManualRefresh Then DrawFade
End Property
Public Property Get Percent() As Byte
Attribute Percent.VB_Description = "Returns percent value."
Attribute Percent.VB_MemberFlags = "400"
    Percent = m_Percent
End Property
Public Property Let Percent(ByVal New_Percent As Byte)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
End Property
Public Property Get PercentAfter() As String
Attribute PercentAfter.VB_Description = "Text after percentage."
Attribute PercentAfter.VB_ProcData.VB_Invoke_Property = "Settings"
    PercentAfter = m_PercentAfter
End Property
Public Property Let PercentAfter(ByVal New_PercentAfter As String)
    m_PercentAfter = New_PercentAfter
    PropertyChanged "PercentAfter"
    If Not m_ManualRefresh Then DrawFade
End Property
Public Property Get PercentBefore() As String
Attribute PercentBefore.VB_Description = "Text before percentage."
Attribute PercentBefore.VB_ProcData.VB_Invoke_Property = "Settings"
    PercentBefore = m_PercentBefore
End Property
Public Property Let PercentBefore(ByVal New_PercentBefore As String)
    m_PercentBefore = New_PercentBefore
    PropertyChanged "PercentBefore"
    If Not m_ManualRefresh Then DrawFade
End Property
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresh object."
    Dim A As Integer
    On Error Resume Next
    If Not m_Vertical Then
        A = (BG.ScaleWidth + 1) * Temp / TempMax
        FG.Move -1, 0, A, BG.ScaleHeight
    Else
        A = (BG.ScaleHeight + 1) * Temp / TempMax
        FG.Move 0, -1, BG.ScaleWidth, A
    End If
    If m_ManualRefresh Then DrawFade
End Sub
Public Property Get Reverse() As Boolean
Attribute Reverse.VB_Description = "Reverse percentage (100% = 0%)."
Attribute Reverse.VB_ProcData.VB_Invoke_Property = "Settings"
    Reverse = m_Reverse
End Property
Public Property Let Reverse(ByVal New_Reverse As Boolean)
    Dim A As Integer
    On Error Resume Next
    m_Reverse = New_Reverse
    PropertyChanged "Reverse"
    A = Int(Temp / TempMax * 100)
    If Not m_Reverse Then
        m_Percent = A
    Else
        m_Percent = 100 - A
    End If
    If Not m_ManualRefresh Then DrawFade
End Property
Public Property Get ScaleMode() As ScaleModeConstants
Attribute ScaleMode.VB_Description = "Returns/sets object's scalemode."
Attribute ScaleMode.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleMode = m_ScaleMode
End Property
Public Property Let ScaleMode(ByVal New_ScaleMode As ScaleModeConstants)
    m_ScaleMode = New_ScaleMode
    UserControl.ScaleMode = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property
Public Sub SetParent(ByVal hwnd As Long)
Attribute SetParent.VB_Description = "Sets object's parent."
    SetWindowParent UserControl.hwnd, hwnd
End Sub
Public Property Get PercentAlign() As munPercentAlign
Attribute PercentAlign.VB_Description = "Returns/sets object's text alignment."
Attribute PercentAlign.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PercentAlign = m_PercentAlign
End Property
Public Property Let PercentAlign(ByVal New_PercentAlign As munPercentAlign)
    m_PercentAlign = New_PercentAlign
    PropertyChanged "PercentAlign"
    If New_PercentAlign > 6 Then InitPercentAlign
    If Not m_ManualRefresh Then DrawFade
End Property
Private Sub TextRefresh()
    If Not m_NoPercent Then
        If m_Custom And ((m_Value = m_Max And Not m_Reverse) Or (m_Value = m_Min And m_Reverse)) Then
            Text = m_CustomText
        Else
            Text = m_PercentBefore & Chr(32) & m_Percent & Chr(32) & m_PercentAfter
        End If
    ElseIf m_PercentBefore <> "" Then
        Text = m_PercentBefore
    End If
End Sub
Public Property Get Value() As Currency
Attribute Value.VB_Description = "Sets object's position value."
Attribute Value.VB_ProcData.VB_Invoke_Property = "Settings"
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property
Public Property Let Value(ByVal New_Value As Currency)
    Dim A As Byte
    On Error Resume Next
    If New_Value < m_Min Then New_Value = m_Min
    If New_Value > m_Max Then New_Value = m_Max
    If m_Value <> New_Value Then
        m_Value = New_Value
        Temp = m_Value - m_Min
        TempMax = m_Max - m_Min
        A = Fix(Temp / TempMax * 100)
        If Not m_Reverse Then
            m_Percent = A
        Else
            m_Percent = 100 - A
        End If
        PropertyChanged "Value"
        If Not m_ManualRefresh Then Refresh
        If m_Percent <> m_OldPercent Then
            m_OldPercent = m_Percent
            If Not m_ManualRefresh And Not NoPercent Then DrawFade
            RaiseEvent Change
        End If
        If m_Value = m_Max Or m_Value = m_Min Then
            If Not m_ManualRefresh Then DrawFade
            RaiseEvent Change
        End If
    Else
        PropertyChanged "Value"
    End If
End Property
Public Property Get Vertical() As Boolean
Attribute Vertical.VB_Description = "View bar vertically?"
Attribute Vertical.VB_ProcData.VB_Invoke_Property = "Settings"
    Vertical = m_Vertical
End Property
Public Property Let Vertical(ByVal New_Vertical As Boolean)
    m_Vertical = New_Vertical
    PropertyChanged "Vertical"
    If Not m_ManualRefresh Then Refresh
    Draw
End Property
Private Sub BG_Resize()
    If Not m_ManualRefresh Then Refresh
End Sub
Private Sub FG_Resize()
    If Not m_ManualRefresh Then DrawFade
    RaiseEvent BarResize
End Sub
Private Sub Scroller_Timer()
    Static m_Direction As Boolean, B As Integer
    If m_PercentAlign < 7 Then Exit Sub
    B = -1
    Select Case m_PercentAlign
        Case 7
            m_Scroll = m_Scroll + 1
            If m_Scroll > BG.ScaleWidth Then
                TextRefresh
                m_Scroll = -BG.TextWidth(Text)
            End If
        Case 8
            m_Scroll = m_Scroll - 1
            TextRefresh
            If m_Scroll < -BG.TextWidth(Text) Then m_Scroll = BG.ScaleWidth
        Case 9
            m_Scroll = m_Scroll + 1
            If m_Scroll > BG.ScaleHeight Then
                TextRefresh
                m_Scroll = -BG.TextHeight(Text)
            End If
        Case 10
            m_Scroll = m_Scroll - 1
            TextRefresh
            If m_Scroll < -BG.TextHeight(Text) Then m_Scroll = BG.ScaleHeight
        Case 11
            If Not m_Direction Then B = 1
            m_Scroll = m_Scroll + B
            TextRefresh
            If m_Direction = False And m_Scroll > BG.ScaleWidth Then
                m_Direction = True
            ElseIf m_Direction And m_Scroll < -BG.TextWidth(Text) Then
                m_Direction = False
            End If
        Case 12
            If Not m_Direction Then B = 1
            m_Scroll = m_Scroll + B
            TextRefresh
            If m_Direction = False And m_Scroll > BG.ScaleWidth - BG.TextWidth(Text) Then
                m_Direction = True
            ElseIf m_Direction And m_Scroll < 0 Then
                m_Direction = False
            End If
        Case 13
            If Not m_Direction Then B = 1
            m_Scroll = m_Scroll + B
            TextRefresh
            If m_Direction = False And m_Scroll > BG.ScaleHeight Then
                m_Direction = True
            ElseIf m_Direction And m_Scroll < -BG.TextHeight(Text) Then
                m_Direction = False
            End If
        Case 14
            If Not m_Direction Then B = 1
            m_Scroll = m_Scroll + B
            TextRefresh
            If m_Direction = False And m_Scroll > BG.ScaleHeight - BG.TextHeight(Text) Then
                m_Direction = True
            ElseIf m_Direction And m_Scroll < 0 Then
                m_Direction = False
            End If
    End Select
    If Not ManualRefresh Then DrawFade
End Sub
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_BarStyle = m_def_BarStyle
    m_BorderStyle = m_def_BorderStyle
    m_Custom = m_def_Custom
    m_CustomText = m_def_CustomText
    m_Fade = m_def_Fade
    m_FadeBG1 = m_def_FadeBG1
    m_FadeBG2 = m_def_FadeBG2
    m_FadeFG1 = m_def_FadeFG1
    m_FadeFG2 = m_def_FadeFG2
    m_FadeStyle = m_def_FadeStyle
    Set m_Font = Ambient.Font
    Set BG.Font = Ambient.Font
    Set FG.Font = Ambient.Font
    m_ForeColor = m_def_ForeColor
    m_Interval = m_def_Interval
    m_ManualRefresh = m_def_ManualRefresh
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_NoPercent = m_def_NoPercent
    m_Percent = m_def_Percent
    m_PercentAfter = m_def_PercentAfter
    m_PercentBefore = m_def_PercentBefore
    m_Reverse = m_def_Reverse
    m_ScaleMode = m_def_ScaleMode
    m_PercentAlign = m_def_PercentAlign
    m_Vertical = m_def_Vertical
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_BarStyle = PropBag.ReadProperty("BarStyle", m_def_BarStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Custom = PropBag.ReadProperty("Custom", m_def_Custom)
    m_CustomText = PropBag.ReadProperty("CustomText", m_def_CustomText)
    m_Fade = PropBag.ReadProperty("Fade", m_def_Fade)
    m_FadeBG1 = PropBag.ReadProperty("FadeBG1", m_def_FadeBG1)
    m_FadeBG2 = PropBag.ReadProperty("FadeBG2", m_def_FadeBG2)
    m_FadeFG1 = PropBag.ReadProperty("FadeFG1", m_def_FadeFG1)
    m_FadeFG2 = PropBag.ReadProperty("FadeFG2", m_def_FadeFG2)
    m_FadeStyle = PropBag.ReadProperty("FadeStyle", m_def_FadeStyle)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Interval = PropBag.ReadProperty("Interval", m_def_Interval)
    m_ManualRefresh = PropBag.ReadProperty("ManualRefresh", m_def_ManualRefresh)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_NoPercent = PropBag.ReadProperty("NoPercent", m_def_NoPercent)
    m_Percent = PropBag.ReadProperty("Percent", m_def_Percent)
    m_PercentAfter = PropBag.ReadProperty("PercentAfter", m_def_PercentAfter)
    m_PercentAlign = PropBag.ReadProperty("PercentAlign", m_def_PercentAlign)
    m_PercentBefore = PropBag.ReadProperty("PercentBefore", m_def_PercentBefore)
    m_Reverse = PropBag.ReadProperty("Reverse", m_def_Reverse)
    m_ScaleMode = PropBag.ReadProperty("ScaleMode", m_def_ScaleMode)
    m_Value = PropBag.ReadProperty("Value", 0)
    m_Vertical = PropBag.ReadProperty("Vertical", m_def_Vertical)
    m_BGborder = EdgeSubClass(BG.hwnd, m_BorderStyle)
    m_FGborder = EdgeSubClass(FG.hwnd, m_BarStyle)
    BG.BackColor = m_BackColor
    FG.BackColor = m_ForeColor
    BG.ForeColor = m_ForeColor
    FG.ForeColor = m_BackColor
    Set BG.Font = m_Font
    Set FG.Font = m_Font
    Temp = m_Value - m_Min
    TempMax = m_Max - m_Min
    InitPercentAlign
    Scroller.Interval = m_Interval
    Refresh
    Draw
End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    BG.Move 0, 0, ScaleWidth, ScaleHeight
    DG.Move 0, 0, ScaleWidth + 1, ScaleHeight * 2 + 1
    Draw
    RaiseEvent Resize
End Sub
Private Sub UserControl_Terminate()
    EdgeUnSubClass BG.hwnd
    EdgeUnSubClass FG.hwnd
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("BarStyle", m_BarStyle, m_def_BarStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Custom", m_Custom, m_def_Custom)
    Call PropBag.WriteProperty("CustomText", m_CustomText, m_def_CustomText)
    Call PropBag.WriteProperty("Fade", m_Fade, m_def_Fade)
    Call PropBag.WriteProperty("FadeBG1", m_FadeBG1, m_def_FadeBG1)
    Call PropBag.WriteProperty("FadeBG2", m_FadeBG2, m_def_FadeBG2)
    Call PropBag.WriteProperty("FadeFG1", m_FadeFG1, m_def_FadeFG1)
    Call PropBag.WriteProperty("FadeFG2", m_FadeFG2, m_def_FadeFG2)
    Call PropBag.WriteProperty("FadeStyle", m_FadeStyle, m_def_FadeStyle)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Interval", m_Interval, m_def_Interval)
    Call PropBag.WriteProperty("ManualRefresh", m_ManualRefresh, m_def_ManualRefresh)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("NoPercent", m_NoPercent, m_def_NoPercent)
    Call PropBag.WriteProperty("Percent", m_Percent, m_def_Percent)
    Call PropBag.WriteProperty("PercentAfter", m_PercentAfter, m_def_PercentAfter)
    Call PropBag.WriteProperty("PercentAlign", m_PercentAlign, m_def_PercentAlign)
    Call PropBag.WriteProperty("PercentBefore", m_PercentBefore, m_def_PercentBefore)
    Call PropBag.WriteProperty("Reverse", m_Reverse, m_def_Reverse)
    Call PropBag.WriteProperty("ScaleMode", m_ScaleMode, m_def_ScaleMode)
    Call PropBag.WriteProperty("Value", m_Value, 0)
    Call PropBag.WriteProperty("Vertical", m_Vertical, m_def_Vertical)
End Sub
