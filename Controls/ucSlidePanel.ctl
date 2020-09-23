VERSION 5.00
Begin VB.UserControl ucSlidePanel 
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2550
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   2550
   ToolboxBitmap   =   "ucSlidePanel.ctx":0000
   Begin Project1.isButton PanelButton 
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   50
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   661
      Style           =   8
      Caption         =   "isButton"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2040
      Top             =   0
   End
   Begin VB.Shape PanelFrame 
      BorderColor     =   &H8000000C&
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   1995
   End
End
Attribute VB_Name = "ucSlidePanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+  File Description:
'       ucSlidePanel - Sliding (Rollup) Panel (Container) Control
'
'   Product Name:
'       ucSlidePanel.ctl
'
'   Compatability:
'       Windows: 95, 98, ME, NT, 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Based on the following On-Line Articles
'       (isButton - Fred.cpp)
'       URL: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56053&lngWId=1
'
'   Legal Copyright & Trademarks:
'       Copyright © 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this  software. This software is owned by Paul R. Territo, Ph.D and is
'       sold for use as a license in accordance with the terms of the License
'       Agreement in the accompanying the documentation.
'
'   Contact Information:
'       For Technical Assistance:
'       Email: pwterrito@insightbb.com
'
'-  Modification(s) History:
'       04Jun05 - Initial test harness and usercontrol finished
'
'   Force Declarations
Option Explicit

'   Private API Declarations & Constants
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_BTNTEXT             As Long = 18

'   Public Events
Public Event ButtonClick()
Public Event ButtonMouseEnter()
Public Event ButtonMouseLeave()
Public Event ButtonMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ButtonMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event HitTest(X As Single, Y As Single, HitResult As Integer)
Public Event PanelClick()
Public Event PanelMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PanelMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PanelResize()

'   Public Enumerations
Public Enum SpeedEnum                   'Speed here actually refers to the step
    [spVerySlow] = 0                    'size for each iteration. Since the rate
    [spSlow] = 1                        'of change here is Stepsize/5msec, the
    [spMedium] = 2                      'smaller the stepsize the slower the panel
    [spFast] = 3                        'will close...we could have adjusted the
    [spVeryFast] = 4                    'timer interval, but this method was much
    [spInstantaneous] = -1              'more consistent.
End Enum

Public Enum BackStyleEnum
    [spTransparent] = 0
    [spOpaque] = 1
End Enum

Public Enum ClipBehaviorEnum
    [None] = 0
    [UseRegion] = 1
End Enum

'   Local Variables - Not Handled by Components
Dim m_ButtonPad             As Long
Dim m_BackColor             As OLE_COLOR
Dim m_ControlsBackColor     As OLE_COLOR
Dim m_Enabled               As Boolean
Dim m_EnabledControls       As Boolean
Dim m_ExpandedHeight        As Long
Dim m_FrameColor            As OLE_COLOR
Dim m_FrameVisible          As Boolean
Dim m_HitBehavior           As Integer
Dim m_PanelMoving           As Boolean
Dim m_PanelOpen             As Boolean
Dim m_Speed                 As SpeedEnum
Dim m_Version               As String

Public Property Get AutoRedraw() As Boolean
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.PanelFrame.BackColor
End Property

Public Property Let BackColor(Value As OLE_COLOR)
    '   Set the Frames back color
    UserControl.PanelFrame.FillColor = Value
    UserControl.PanelFrame.FillStyle = 0
    m_BackColor = Value
    '   Make sure the rest of the usercontrol looks like the parent
    UserControl.BackColor = UserControl.Parent.BackColor
    Call ControlsTransparent(spOpaque)
    PropertyChanged "BackColor"
End Property

Public Property Get BackStyle() As BackStyleEnum
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BackStyleEnum)
    UserControl.BackStyle = New_BackStyle
    If New_BackStyle = spOpaque Then
        UserControl.PanelFrame.FillStyle = 0
        UserControl.PanelFrame.FillColor = m_BackColor
    Else
        UserControl.PanelFrame.FillStyle = 1
        If UserControl.PanelFrame.FillColor = 0 Then
            m_BackColor = UserControl.BackColor
        Else
            m_BackColor = UserControl.PanelFrame.FillColor
        End If
    End If
    PropertyChanged "BackStyle"
End Property

Public Property Get ButtonHeight() As Long
    ButtonHeight = PanelButton.Height
End Property

Public Property Let ButtonHeight(Value As Long)
    PanelButton.Height = Value
    If Value > UserControl.Height Then
        UserControl.Height = Value + (m_ButtonPad * 2)
    End If
    '   Fire a resize event to repaint/position the control
    UserControl_Resize
    PropertyChanged "ButtonHeight"
End Property

Public Property Get ButtonPad() As Long
    ButtonPad = m_ButtonPad
End Property

Public Property Let ButtonPad(Value As Long)
    m_ButtonPad = Value
    If (PanelButton.Height + (Value * 2)) > UserControl.Height Then
        UserControl.Height = PanelButton.Height + (Value * 2)
    End If
    '   Fire a resize event to repaint/position the control
    UserControl_Resize
    PropertyChanged "ButtonPad"
End Property

Public Property Get ButtonStyle() As isbStyle
    ButtonStyle = PanelButton.Style
End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As isbStyle)
    PanelButton.Style = New_ButtonStyle
    PropertyChanged "ButtonStyle"
End Property

Public Property Get CaptionAlign() As isbAlign
    CaptionAlign = PanelButton.CaptionAlign
End Property

Public Property Let CaptionAlign(ByVal Value As isbAlign)
    PanelButton.CaptionAlign = Value
    PropertyChanged "CaptionAlign"
End Property

Public Property Get Caption() As String
    Caption = PanelButton.Caption
End Property

Public Property Let Caption(Value As String)
    PanelButton.Caption = Value
    PropertyChanged "Caption"
End Property

Public Property Get ClipBehavior() As ClipBehaviorEnum
    ClipBehavior = UserControl.ClipBehavior
End Property

Public Property Let ClipBehavior(ByVal Value As ClipBehaviorEnum)
    UserControl.ClipBehavior = Value
    PropertyChanged "ClipBehavior"
End Property

Public Property Get ClipControls() As Boolean
    ClipControls = UserControl.ClipControls
End Property

Public Property Let ClipControls(Value As Boolean)
    UserControl.ClipControls = Value
    If Value = True Then
        Call ControlsTransparent(spTransparent)
    Else
        Call ControlsTransparent(spOpaque)
    End If
    UserControl.Refresh
    PropertyChanged "ClipControls"
End Property

Public Property Get ClosedHeight() As Long
    ClosedHeight = ButtonHeight + (m_ButtonPad * 2)
End Property

Public Sub Cls()
    UserControl.Cls
End Sub

Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Private Sub ControlsEnabled(Value As Boolean, All As Boolean)
    Dim Ctl     As Control
    
    On Error Resume Next
    '   Enable/Disable all of the contained controls
    For Each Ctl In UserControl.ContainedControls
        Ctl.Enabled = Value
    Next Ctl
    If All Then
        '   Now Enable/Disable our own controls
        For Each Ctl In UserControl.Controls
            Ctl.Enabled = Value
        Next Ctl
    End If
    On Error GoTo 0
End Sub

Private Sub ControlsTransparent(Value As BackStyleEnum)
    Dim Ctl     As Control
    
    On Error Resume Next
    With UserControl
        For Each Ctl In .ContainedControls
            Ctl.BackStyle = Value
            If (TypeOf Ctl Is CheckBox) Or (TypeOf Ctl Is OptionButton) Or (TypeOf Ctl Is Label) Then
                If ClipControls = False Then
                    Ctl.BackColor = m_ControlsBackColor
                Else
                    m_ControlsBackColor = Ctl.BackColor
                    Ctl.BackColor = m_BackColor
                End If
            End If
        Next Ctl
    End With
    On Error GoTo 0
End Sub

Private Sub ControlsVisible(Value As Boolean)
    Dim Ctl     As Control
    With UserControl
        For Each Ctl In .ContainedControls
            Ctl.Visible = Value
        Next Ctl
    End With
End Sub

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(Value As Boolean)
    m_Enabled = Value
    Call ControlsEnabled(m_Enabled, True)
    PropertyChanged "Enabled"
End Property

Public Property Let EnabledControls(Value As Boolean)
    m_EnabledControls = Value
    Call ControlsEnabled(m_EnabledControls, False)
    PropertyChanged "EnabledControls"
End Property

Public Property Get EnabledControls() As Boolean
    EnabledControls = m_EnabledControls
End Property

Public Property Get FontColor() As OLE_COLOR
    FontColor = PanelButton.FontColor
End Property

Public Property Let FontColor(lFontColor As OLE_COLOR)
    PanelButton.FontColor = lFontColor
    PropertyChanged "FontColor"
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.PanelButton.Font
End Property

Public Property Set Font(New_Font As StdFont)
    '   Make sure any new fonts are universal in the control
    Set UserControl.Font = New_Font
    Set UserControl.PanelButton.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get FontHighlightColor() As OLE_COLOR
    FontHighlightColor = PanelButton.FontHighlightColor
End Property

Public Property Let FontHighlightColor(lFontHighlightColor As OLE_COLOR)
    PanelButton.FontHighlightColor = lFontHighlightColor
    PropertyChanged "FontHighlightColor"
End Property

Public Property Get FrameColor() As OLE_COLOR
    FrameColor = PanelFrame.BorderColor
End Property

Public Property Let FrameColor(Value As OLE_COLOR)
    m_FrameColor = Value
    PanelFrame.BorderColor = Value
    PropertyChanged "FrameColor"
End Property

Public Property Get FrameShape() As ShapeConstants
    FrameShape = PanelFrame.Shape
End Property

Public Property Let FrameShape(Value As ShapeConstants)
    PanelFrame.Shape = Value
    PropertyChanged "FrameShape"
End Property

Public Property Get FrameThickness() As Long
    FrameThickness = PanelFrame.BorderWidth
End Property

Public Property Get FrameVisible() As Boolean
    FrameVisible = m_FrameVisible
End Property

Public Property Let FrameVisible(Value As Boolean)
    '   Create the illusion of hiding the frame...
    '   Setting the PanelFrame.Visible property causes
    '   odd effects on ALL frames, not just the local one...?
    If Value = True Then
        m_FrameVisible = True
        PanelFrame.BorderColor = m_FrameColor
    Else
        m_FrameVisible = False
        PanelFrame.BorderColor = BackColor
    End If
    PropertyChanged "FrameVisible"
End Property

Public Property Let FrameThickness(Value As Long)
    PanelFrame.BorderWidth = Value
    PropertyChanged "FrameThickness"
End Property

Public Property Get HasDC() As Boolean
    HasDC = UserControl.HasDC
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get HitBehavior() As Integer
    HitBehavior = m_HitBehavior
End Property

Public Property Let HitBehavior(ByVal New_HitBehavior As Integer)
    m_HitBehavior = New_HitBehavior
    UserControl.HitBehavior() = New_HitBehavior
    PropertyChanged "HitBehavior"
End Property

Public Property Get IconAlign() As isbAlign
    IconAlign = PanelButton.IconAlign
End Property

Public Property Let IconAlign(ByVal NewIconAlign As isbAlign)
    PanelButton.IconAlign = NewIconAlign
    PropertyChanged "IconAlign"
End Property

Public Property Get Icon() As StdPicture
    Set Icon = PanelButton.Icon
End Property

Public Property Set Icon(NewIcon As StdPicture)
    Set PanelButton.Icon = NewIcon
    PropertyChanged "Icon"
End Property

Public Property Get IconSize() As Integer
    IconSize = PanelButton.IconSize
End Property

Public Property Let IconSize(ByVal NewIconSize As Integer)
    PanelButton.IconSize = NewIconSize
    PropertyChanged "IconSize"
End Property

Public Property Get NonThemeStyle() As isbStyle
    NonThemeStyle = PanelButton.NonThemeStyle
End Property

Public Property Let NonThemeStyle(ByVal NewNonThemeStyle As isbStyle)
    PanelButton.NonThemeStyle = NewNonThemeStyle
    PropertyChanged "NonThemeStyle"
End Property

Private Sub PanelButton_Click()
    With UserControl
        If m_PanelOpen = False Then
            With .PanelFrame
                .Height = UserControl.PanelButton.Height + (m_ButtonPad * 2)
                UserControl.Height = UserControl.PanelButton.Height + (m_ButtonPad * 2)
                .Visible = True
            End With
        End If
        .Timer1.Interval = 1
        .Timer1.Enabled = True
    End With
    RaiseEvent ButtonClick
End Sub

Private Property Get PanelMoving() As Boolean
    PanelMoving = m_PanelMoving
End Property

Private Property Let PanelMoving(Value As Boolean)
    m_PanelMoving = Value
End Property

Public Property Get PanelOpen() As Boolean
    PanelOpen = m_PanelOpen
End Property

Public Property Let PanelOpen(ByVal New_PanelOpen As Boolean)
    m_PanelOpen = New_PanelOpen
    '   Panel was open, so force it closed...
    With UserControl
        If New_PanelOpen = False Then
            PanelMoving = True
            Reset
        End If
    End With
    PropertyChanged "PanelOpen"
End Property

Public Property Get ScaleHeight() As Single
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

Public Property Get ScaleLeft() As Single
    ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    UserControl.ScaleLeft = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

Public Property Get ScaleMode() As ScaleModeConstants
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As ScaleModeConstants)
    UserControl.ScaleMode = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

Public Property Get ScaleTop() As Single
    ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    UserControl.ScaleTop = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

Public Property Get ScaleWidth() As Single
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

Public Function ScaleX(ByVal Width As Single, ByVal FromScale As Variant, ByVal ToScale As Variant) As Single
    ScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
End Function

Public Function ScaleY(ByVal Height As Single, ByVal FromScale As Variant, ByVal ToScale As Variant) As Single
    ScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
End Function

Public Property Get Speed() As SpeedEnum
    Speed = m_Speed
End Property

Public Property Let Speed(Value As SpeedEnum)
    m_Speed = Value
    PropertyChanged "Speed"
End Property

Private Sub PanelButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent PanelMouseDown(Button, Shift, X, Y)
End Sub

Private Sub PanelButton_MouseEnter()
    RaiseEvent ButtonMouseEnter
End Sub

Private Sub PanelButton_MouseLeave()
    RaiseEvent ButtonMouseLeave
End Sub

Private Sub PanelButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent PanelMouseDown(Button, Shift, X, Y)
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub

Public Sub Reset()
    With UserControl
        If m_PanelOpen Then
        'Note: This code can be used instead of the timed closure method.
        '      The net effect is an instanatnous closure of the control...
'            .PanelFrame.Height = .PanelButton.Height + (m_ButtonPad * 2)
'            .Height = .PanelButton.Height + (m_ButtonPad * 2)
'            .Timer1.Interval = 0
'            .Timer1.Enabled = False
'            .PanelFrame.Visible = False
'            m_PanelOpen = False
            If PanelMoving = False Then
                .Timer1.Interval = 5
                .Timer1.Enabled = True
            End If
        End If
    End With
End Sub

Private Sub Timer1_Timer()
    Dim StepSize    As Double
    
    Select Case Speed
        Case 0: StepSize = 0.00625      'Very Slow  = 0.63% of Height/Iteration
        Case 1: StepSize = 0.0125       'Slow       = 1.25% of Height/Iteration
        Case 2: StepSize = 0.025        'Medium     = 2.50% of Height/Iteration
        Case 3: StepSize = 0.0375       'Fast       = 3.75% of Height/Iteration
        Case 4: StepSize = 0.05         'Very Fast  = 5.00% of Height/Iteration
        Case Else: StepSize = 0.0375    'Other      = 3.75% of Height/Iteration
    End Select
    With UserControl
        If m_PanelOpen = False Then
            '   We are closed so open the frame
            If m_Speed = spInstantaneous Then
                .Height = m_ExpandedHeight - m_ButtonPad
                .PanelFrame.Height = m_ExpandedHeight - m_ButtonPad
                .Timer1.Interval = 0
                .Timer1.Enabled = False
                Call ControlsVisible(True)
                m_PanelOpen = True
                PanelMoving = False
            Else
                If .PanelFrame.Height >= m_ExpandedHeight - m_ButtonPad Then
                    .Timer1.Interval = 0
                    .Timer1.Enabled = False
                    Call ControlsVisible(True)
                    m_PanelOpen = True
                Else
'                    .PanelFrame.Height = .PanelFrame.Height + m_Speed
'                    .Height = .PanelFrame.Height + m_Speed
                    .PanelFrame.Height = .PanelFrame.Height + (.PanelFrame.Height * StepSize)
                    .Height = .PanelFrame.Height + (.PanelFrame.Height * StepSize)
                End If
            End If
        Else
            '   Opened, so close the frame
            If m_Speed = spInstantaneous Then
                .Height = .PanelButton.Height + (m_ButtonPad * 2)
                .PanelFrame.Height = .PanelButton.Height + (m_ButtonPad * 2)
                .Timer1.Interval = 0
                .Timer1.Enabled = False
                Call ControlsVisible(False)
                .PanelFrame.Visible = False
                m_PanelOpen = False
                PanelMoving = False
            Else
                '   Check to see if the panel is smaller than the button + pad for this
                '   iteration, and also check to see if next iteration will be too small...
'                If (.PanelFrame.Height <= .PanelButton.Height + (m_ButtonPad * 2)) Or _
'                    (.PanelFrame.Height - m_Speed <= .PanelButton.Height + (m_ButtonPad * 2)) Then
                If (.PanelFrame.Height <= .PanelButton.Height + (m_ButtonPad * 2)) Or _
                    (.PanelFrame.Height - (.PanelFrame.Height * StepSize) <= .PanelButton.Height + (m_ButtonPad * 2)) Then
                    '   Make sure the control is the size we started with...
                    '.PanelFrame.Height = .PanelButton.Height + (m_ButtonPad * 2)
                    .Timer1.Interval = 0
                    .Timer1.Enabled = False
                    .PanelFrame.Visible = False
                    m_PanelOpen = False
                Else
                    '   Make sure the contained controls are still hidden and proceed
                    '   with making the panel smaller...
                    Call ControlsVisible(False)
'                    .PanelFrame.Height = .PanelFrame.Height - m_Speed
'                    .Height = .PanelFrame.Height - m_Speed
                    .PanelFrame.Height = .PanelFrame.Height - (.PanelFrame.Height * StepSize)
                    .Height = .PanelFrame.Height - (.PanelFrame.Height * StepSize)
                End If
            End If
        End If
        RaiseEvent PanelResize
    End With
End Sub

Private Sub UserControl_InitProperties()
    AutoRedraw = False
    BackStyle = 1 'spOpaque
    BackColor = &H8000000F
    m_BackColor = &H8000000F
    ButtonHeight = 375
    ButtonPad = 50
    Caption = UserControl.Extender.Name
    CaptionAlign = isbLeft
    ClipBehavior = UseRegion
    ClipControls = True
    Enabled = True
    EnabledControls = True
    FontColor = GetSysColor(COLOR_BTNTEXT)
    FontHighlightColor = GetSysColor(COLOR_BTNTEXT)
    FrameColor = &H8000000C
    FrameShape = vbShapeRectangle
    FrameThickness = 1
    FrameVisible = True
    HitBehavior = 1
    IconAlign = 0
    IconSize = 16
    NonThemeStyle = [isbWindowsXP]
    Speed = spVeryFast
    ButtonStyle = isbGalaxy
    m_Version = "1.0.0"
    Set Icon = Nothing
    Set UserControl.PanelButton.Font = UserControl.PanelButton.Font
    UserControl.PanelFrame.Visible = True
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    RaiseEvent HitTest(X, Y, HitResult)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent PanelMouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent PanelMouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    '   Hide all of the contained controls until we
    '   open the Panel Frame
    If PanelOpen = False Then
        Call ControlsVisible(False)
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_ButtonPad = .ReadProperty("ButtonPad", 50)
        m_Enabled = .ReadProperty("Enabled", True)
        m_EnabledControls = .ReadProperty("EnabledControls", True)
        m_HitBehavior = .ReadProperty("HitBehavior", 1)
        m_Speed = .ReadProperty("Speed", spVeryFast)
        Set UserControl.PanelButton.Icon = PropBag.ReadProperty("Icon", Nothing)
        Set UserControl.PanelButton.Font = .ReadProperty("Font", UserControl.PanelButton.Font)
        UserControl.AutoRedraw = .ReadProperty("AutoRedraw", False)
        UserControl.PanelFrame.BackColor = .ReadProperty("BackColor", &H8000000F)
'        UserControl.PanelFrame.FillStyle = .ReadProperty("FillStyle", 0)
        UserControl.BackStyle = .ReadProperty("BackStyle", 1)
        UserControl.PanelButton.Caption = .ReadProperty("Caption", "ucPanelButton")
        UserControl.PanelButton.CaptionAlign = .ReadProperty("CaptionAlign", isbLeft)
        UserControl.ClipBehavior = .ReadProperty("ClipBehavior", UseRegion)
        UserControl.ClipControls = .ReadProperty("ClipControls", True)
        UserControl.PanelButton.FontColor = PropBag.ReadProperty("FontColor", GetSysColor(COLOR_BTNTEXT))
        UserControl.PanelButton.FontHighlightColor = PropBag.ReadProperty("FontHighlightColor", GetSysColor(COLOR_BTNTEXT))
        UserControl.PanelButton.Height = .ReadProperty("ButtonHeight", 375)
        UserControl.PanelButton.IconAlign = .ReadProperty("IconAlign", 0)
        UserControl.PanelButton.IconSize = .ReadProperty("IconSize", 16)
        UserControl.PanelButton.NonThemeStyle = PropBag.ReadProperty("NonThemeStyle", [isbWindowsXP])
        UserControl.PanelButton.Style = PropBag.ReadProperty("ButtonStyle", 8)
        UserControl.PanelFrame.BorderColor = .ReadProperty("FrameColor", &H8000000C)
        m_FrameColor = FrameColor
        UserControl.PanelFrame.Shape = .ReadProperty("FrameShape", vbShapeRectangle)
        UserControl.PanelFrame.BorderWidth = .ReadProperty("FrameThickness", 1)
        m_FrameVisible = .ReadProperty("FrameVisible", True)
        UserControl.ScaleHeight = .ReadProperty("ScaleHeight", 1725)
        UserControl.ScaleLeft = .ReadProperty("ScaleLeft", 0)
        UserControl.ScaleMode = .ReadProperty("ScaleMode", vbTwips)
        UserControl.ScaleTop = .ReadProperty("ScaleTop", 0)
        UserControl.ScaleWidth = .ReadProperty("ScaleWidth", 2550)
        m_ControlsBackColor = UserControl.Parent.BackColor
        m_BackColor = UserControl.Parent.BackColor
        
        With UserControl
            m_ExpandedHeight = UserControl.Height
            If Ambient.UserMode = False Then
                .PanelFrame.Visible = True
            Else
                UserControl.PanelFrame.Visible = False
                .PanelFrame.Height = .PanelButton.Height + (m_ButtonPad * 2)
                .Height = .PanelButton.Height + (m_ButtonPad * 2)
                UserControl.Refresh
            End If
        End With
    End With
End Sub

Private Sub UserControl_Resize()
    With UserControl
        .PanelFrame.Height = .Height
        .PanelFrame.Width = .Width
        '   Make sure the button is smaller then the control
        .PanelButton.Width = .Width - (m_ButtonPad * 2)
        '   Now center the button and pad the top
        .PanelButton.Top = m_ButtonPad
        .PanelButton.Left = m_ButtonPad
    End With
End Sub

Private Sub UserControl_Show()
    With UserControl
        If Ambient.UserMode = False Then
            .PanelFrame.Visible = True
        Else
            UserControl.PanelFrame.Visible = False
            .PanelFrame.Height = .PanelButton.Height + (m_ButtonPad * 2)
            .Height = .PanelButton.Height + (m_ButtonPad * 2)
        End If
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
        Call .WriteProperty("BackColor", UserControl.PanelFrame.BackColor, &H8000000F)
'        Call .WriteProperty("FillStyle", UserControl.PanelFrame.FillStyle, 0)
        Call .WriteProperty("BackStyle", UserControl.BackStyle, 1)
        Call .WriteProperty("ButtonHeight", UserControl.PanelButton.Height, 375)
        Call .WriteProperty("ButtonPad", m_ButtonPad, 50)
        Call .WriteProperty("Caption", UserControl.PanelButton.Caption, "ucPanelButton")
        Call .WriteProperty("CaptionAlign", UserControl.PanelButton.CaptionAlign, isbLeft)
        Call .WriteProperty("ClipBehavior", UserControl.ClipBehavior, UseRegion)
        Call .WriteProperty("ClipControls", UserControl.ClipControls, True)
        Call .WriteProperty("Enabled", m_Enabled, True)
        Call .WriteProperty("EnabledControls", m_EnabledControls, True)
        Call .WriteProperty("Font", UserControl.PanelButton.Font)
        Call .WriteProperty("FontColor", UserControl.PanelButton.FontColor, GetSysColor(COLOR_BTNTEXT))
        Call .WriteProperty("FontHighlightColor", UserControl.PanelButton.FontHighlightColor, GetSysColor(COLOR_BTNTEXT))
        Call .WriteProperty("FrameColor", UserControl.PanelFrame.BorderColor, &H8000000C)
        Call .WriteProperty("FrameShape", UserControl.PanelFrame.Shape, vbShapeRectangle)
        Call .WriteProperty("FrameThickness", UserControl.PanelFrame.BorderWidth, 1)
        Call .WriteProperty("FrameVisible", m_FrameVisible, True)
        Call .WriteProperty("HitBehavior", m_HitBehavior, 1)
        Call .WriteProperty("Icon", UserControl.PanelButton.Icon)
        Call .WriteProperty("IconAlign", UserControl.PanelButton.IconAlign, 0)
        Call .WriteProperty("IconSize", UserControl.PanelButton.IconSize, 16)
        Call .WriteProperty("NonThemeStyle", UserControl.PanelButton.NonThemeStyle, isbWindowsXP)
        Call .WriteProperty("Speed", m_Speed, spVeryFast)
        Call .WriteProperty("ButtonStyle", UserControl.PanelButton.Style, 8)
        Call .WriteProperty("ScaleHeight", UserControl.ScaleHeight, 1725)
        Call .WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
        Call .WriteProperty("ScaleMode", UserControl.ScaleMode, vbTwips)
        Call .WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
        Call .WriteProperty("ScaleWidth", UserControl.ScaleWidth, 2550)
    End With
End Sub

Public Function Version() As String
    Version = m_Version
End Function
