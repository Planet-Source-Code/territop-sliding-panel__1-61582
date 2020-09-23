VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ucSlidePanel ucSlidePanel4 
      Height          =   2415
      Left            =   4320
      TabIndex        =   34
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4260
      BackColor       =   -2147483643
      Caption         =   "ucSlidePanel Example"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   2415
      ScaleMode       =   0
      ScaleWidth      =   2295
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   1200
         TabIndex        =   38
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   36
         Text            =   "Combo1"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   975
         Left            =   1680
         TabIndex        =   35
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   255
         Left            =   1200
         TabIndex        =   41
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin Project1.ucSlidePanel ucSlidePanel3 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3625
      Caption         =   "Frame Options"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Speed           =   -1
      ScaleHeight     =   2055
      ScaleMode       =   0
      ScaleWidth      =   3615
      Begin VB.CheckBox FrameVisible 
         Caption         =   "Frame Visible"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   1560
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox FrameColorValue 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "-2147483636 "
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton GetFrameColors 
         Caption         =   "..."
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Thickness 
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Text            =   "1"
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox FrameShape 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "FrameColor"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Frame Thickness"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Frame Style"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin Project1.ucSlidePanel ucSlidePanel2 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3625
      Caption         =   "Button Options"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Speed           =   -1
      ScaleHeight     =   2055
      ScaleMode       =   0
      ScaleWidth      =   3615
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   3120
         TabIndex        =   33
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox ButtonStyle 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox ButtonHeight 
         Height          =   315
         Left            =   2160
         TabIndex        =   27
         Text            =   "375"
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox ButtonCaption 
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Text            =   "ucSlidePanel Example"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox ButtonPad 
         Height          =   315
         Left            =   2160
         TabIndex        =   25
         Text            =   "50"
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Button Style"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Button Height"
         Height          =   255
         Left            =   2160
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Button Caption"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Button Pad"
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin Project1.ucSlidePanel ucSlidePanel1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3625
      Caption         =   "Panel Options"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Speed           =   -1
      ScaleHeight     =   2055
      ScaleMode       =   0
      ScaleWidth      =   3615
      Begin VB.CheckBox EnablePanel 
         Caption         =   "Enable Panel"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   24
         Top             =   480
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton GetColors 
         Caption         =   "..."
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   1605
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox PanelRate 
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox EnablePanel 
         Caption         =   "Enable Controls"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   18
         Top             =   720
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton BackStyle 
         Caption         =   "Opaque"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton BackStyle 
         Caption         =   "Transpartent"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox ColorValue 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "-2147483633"
         Top             =   1605
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox ClipControl 
         Caption         =   "Clip Controls"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   1080
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Open / Close Rate"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "BackStyle"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "BackColor"
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   1365
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Label Label8 
      Caption         =   "If you can see this label then the controls Transparency = True"
      Height          =   735
      Left            =   4680
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Changing the properties in the control panels to see the effect on the ucSlidePanel Below."
      Height          =   615
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim PrevValue       As Integer
Dim PrevValueBH     As Long
Dim PrevValueBP     As Integer

Private Sub AlignPanels(Index As Integer)
    With Me
        If (.ucSlidePanel1.Speed = spInstantaneous) And (.ucSlidePanel2.Speed = spInstantaneous) And (.ucSlidePanel3.Speed = spInstantaneous) Then
            .ucSlidePanel1.Top = 120
            .ucSlidePanel1.Left = 120
            .ucSlidePanel2.Top = .ucSlidePanel1.Top + .ucSlidePanel1.Height + 20
            .ucSlidePanel2.Left = 120
            .ucSlidePanel3.Top = .ucSlidePanel2.Top + .ucSlidePanel2.Height + 20
            .ucSlidePanel3.Left = 120
        Else
            '   Only move the panels that need it, or the effect is
            '   the panels bounce up and down...by using the following routines
            '   the net effect is smooth scrolling of the panels during transitions
            If (.ucSlidePanel1.PanelOpen = False) And _
               (.ucSlidePanel2.PanelOpen = False) And _
               (.ucSlidePanel3.PanelOpen = False) Then
                .ucSlidePanel1.Top = 120
                .ucSlidePanel1.Left = 120
                .ucSlidePanel2.Top = .ucSlidePanel1.Top + .ucSlidePanel1.Height + 20
                .ucSlidePanel2.Left = 120
                .ucSlidePanel3.Top = .ucSlidePanel2.Top + .ucSlidePanel2.Height + 20
                .ucSlidePanel3.Left = 120
            ElseIf (.ucSlidePanel1.PanelOpen = True) And _
               (.ucSlidePanel2.PanelOpen = False) And _
               (.ucSlidePanel3.PanelOpen = False) Then
                .ucSlidePanel1.Top = 120
                .ucSlidePanel1.Left = 120
                .ucSlidePanel2.Top = .ucSlidePanel1.Top + .ucSlidePanel1.Height + 20
                .ucSlidePanel2.Left = 120
                '   Make sure we only move this when we have to...
                If ((Index <> 1) And (Index = 3)) Then
                    .ucSlidePanel3.Top = .ucSlidePanel2.Top + .ucSlidePanel2.Height + 20
                    .ucSlidePanel3.Left = 120
                End If
            ElseIf (.ucSlidePanel1.PanelOpen = False) And _
               (.ucSlidePanel2.PanelOpen = True) And _
               (.ucSlidePanel3.PanelOpen = False) Then
                .ucSlidePanel2.Top = .ucSlidePanel1.Top + .ucSlidePanel1.Height + 20
                .ucSlidePanel2.Left = 120
                '   Make sure we only move this when we have to...
                If ((Index <> 2) And (Index = 3)) Then
                    .ucSlidePanel3.Top = .ucSlidePanel2.Top + .ucSlidePanel2.Height + 20
                    .ucSlidePanel3.Left = 120
                End If
            ElseIf (.ucSlidePanel1.PanelOpen = False) And _
               (.ucSlidePanel2.PanelOpen = False) And _
               (.ucSlidePanel3.PanelOpen = True) Then
               '   Make sure we only move this when we have to...
                If (Index = 1) And (Index <> 3) Then
                    .ucSlidePanel2.Top = .ucSlidePanel1.Top + .ucSlidePanel1.Height + 20
                    .ucSlidePanel2.Left = 120
                End If
                .ucSlidePanel3.Top = .ucSlidePanel2.Top + .ucSlidePanel2.Height + 20
                .ucSlidePanel3.Left = 120
            End If
        End If
    End With
End Sub

Private Sub BackStyle_Click(Index As Integer)
    With Me
        Select Case Index
            Case 0:
                .ucSlidePanel4.BackStyle = 1
                If .ucSlidePanel4.PanelOpen Then .Label8.Visible = False
            Case 1:
                .ucSlidePanel4.BackStyle = 0
                If .ucSlidePanel4.PanelOpen Then .Label8.Visible = True
        End Select
    End With
End Sub

Private Sub ButtonCaption_Change()
    With Me
        .ucSlidePanel4.Caption = .ButtonCaption.Text
    End With
End Sub

Private Sub ButtonHeight_Change()
    With Me
        If IsNumeric(.ButtonHeight.Text) Then
            If ButtonHeight.Text <= 1000 And .ButtonHeight.Text >= 1 Then
                .ucSlidePanel4.ButtonHeight = .ButtonHeight.Text
                PrevValueBH = .ButtonHeight.Text
            Else
                MsgBox "Please enter a number between 1 and 1000!", vbExclamation
                .ButtonHeight.Text = PrevValueBH
            End If
        Else
            MsgBox "Please enter a number!", vbExclamation
            .ButtonHeight.Text = PrevValueBH
        End If
    End With

End Sub

Private Sub ButtonPad_Change()
    With Me
        If IsNumeric(.ButtonPad.Text) Then
            If ButtonPad.Text <= 199 And .ButtonPad.Text >= 1 Then
                .ucSlidePanel4.ButtonPad = .ButtonPad.Text
                PrevValueBP = .ButtonPad.Text
            Else
                MsgBox "Please enter a number between 1 and 200!", vbExclamation
                .ButtonPad.Text = PrevValueBP
            End If
        Else
            MsgBox "Please enter a number!", vbExclamation
            .ButtonPad.Text = PrevValueBP
        End If
    End With
End Sub

Private Sub ButtonStyle_Click()
    With Me
        .ucSlidePanel4.ButtonStyle = .ButtonStyle.ListIndex
    End With
End Sub

Private Sub ClipControl_Click()
    With Me
        .ucSlidePanel4.ClipControls = .ClipControl.Value
    End With
End Sub

Private Sub ColorValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Me.ucSlidePanel4.BackColor = Me.ColorValue.Text
    End If
End Sub

Private Sub EnablePanel_Click(Index As Integer)
    With Me
        Select Case Index
            Case 0: .ucSlidePanel4.Enabled = Me.EnablePanel(0).Value
            Case 1: .ucSlidePanel4.EnabledControls = Me.EnablePanel(1).Value
        End Select
    End With
End Sub

Private Sub FrameShape_Click()
    With Me
        .ucSlidePanel4.FrameShape = .FrameShape.ListIndex
    End With
End Sub

Private Sub FrameVisible_Click()
    With Me
        .ucSlidePanel4.FrameVisible = Abs(.FrameVisible.Value)
    End With
End Sub

Private Sub GetColors_Click()
    Dim ColorVal        As SelectedColor
    
    With Me
        ColorVal = ShowColor(Me.hwnd, True)
        If ColorVal.bCanceled = False Then
            .ColorValue.Text = ColorVal.oSelectedColor
            .ucSlidePanel4.BackColor = CLng(.ColorValue.Text)
        End If
    End With
End Sub

Private Sub GetFrameColors_Click()
    Dim ColorVal        As SelectedColor
    
    With Me
        ColorVal = ShowColor(Me.hwnd, True)
        If ColorVal.bCanceled = False Then
            .FrameColorValue.Text = ColorVal.oSelectedColor
            .ucSlidePanel4.FrameColor = CLng(.FrameColorValue.Text)
        End If
    End With
End Sub

Private Sub PanelRate_Click()
    With Me
        Select Case .PanelRate.ListIndex
            Case 0: .ucSlidePanel4.Speed = spVerySlow
            Case 1: .ucSlidePanel4.Speed = spSlow
            Case 2: .ucSlidePanel4.Speed = spMedium
            Case 3: .ucSlidePanel4.Speed = spFast
            Case 4: .ucSlidePanel4.Speed = spVeryFast
            Case 5: .ucSlidePanel4.Speed = spInstantaneous
        End Select
        .ucSlidePanel1.Speed = .ucSlidePanel4.Speed
        .ucSlidePanel2.Speed = .ucSlidePanel4.Speed
        .ucSlidePanel3.Speed = .ucSlidePanel4.Speed
    End With
End Sub

Private Sub Form_Load()
    With Me
        With .PanelRate
            .AddItem "Very Slow"
            .AddItem "Slow"
            .AddItem "Medium"
            .AddItem "Fast"
            .AddItem "Very Fast"
            .AddItem "Instantaneous"
            .ListIndex = 4
        End With
        .Label8.Visible = False
        PrevValue = 1
        PrevValueBH = 375
        PrevValueBP = 50
        With .ButtonStyle
            .AddItem "isbNormal"
            .AddItem "isbSoft"
            .AddItem "isbFlat"
            .AddItem "isbJava"
            .AddItem "isbOfficeXP"
            .AddItem "isbWindowsXP"
            .AddItem "isbWindowsTheme"
            .AddItem "isbPlastik"
            .AddItem "isbGalaxy"
            .AddItem "isbKeramik"
            .AddItem "isbMacOSX"
            .ListIndex = 8
        End With
        With .FrameShape
            .AddItem "Rectangle"
            .AddItem "Square"
            .AddItem "Oval"
            .AddItem "Circle"
            .AddItem "RoundedRectangle"
            .AddItem "RoundedSquare"
            .ListIndex = 0
        End With
    End With
End Sub

Private Sub Thickness_Change()
    With Me
        If IsNumeric(.Thickness.Text) Then
            If Thickness.Text <= 20 And .Thickness.Text >= 1 Then
                .ucSlidePanel4.FrameThickness = .Thickness.Text
                PrevValue = .Thickness.Text
            Else
                MsgBox "Please enter a number between 1 and 20!", vbExclamation
                .Thickness.Text = PrevValue
            End If
        Else
            MsgBox "Please enter a number!", vbExclamation
            .Thickness.Text = PrevValue
        End If
    End With
End Sub

Private Sub ucSlidePanel1_ButtonClick()
    '   Make sure the panels are aligned to start...
    Call AlignPanels(1)
    With Me
        '   Close all opened panels using their default methods
        With .ucSlidePanel2
            If .PanelOpen Then .Reset
        End With
        With .ucSlidePanel3
            If .PanelOpen Then .Reset
        End With
    End With
End Sub

Private Sub ucSlidePanel1_PanelResize()
    '   Make sure the panels are aligned to start...
    Call AlignPanels(1)
End Sub

Private Sub ucSlidePanel2_ButtonClick()
    '   Make sure the panels are aligned to start...
    Call AlignPanels(2)
    With Me
        '   Close all opened panels using their default methods
        With .ucSlidePanel1
            If .PanelOpen Then .Reset
        End With
        With .ucSlidePanel3
            If .PanelOpen Then .Reset
        End With
    End With
End Sub

Private Sub ucSlidePanel2_PanelResize()
    Call AlignPanels(2)
End Sub

Private Sub ucSlidePanel3_ButtonClick()
    '   Make sure the panels are aligned to start...
    Call AlignPanels(3)
    With Me
        '   Close all opened panels using their default methods
        With .ucSlidePanel1
            If .PanelOpen Then .Reset
        End With
        With .ucSlidePanel2
            If .PanelOpen Then .Reset
        End With
    End With
End Sub

Private Sub ucSlidePanel3_PanelResize()
    Call AlignPanels(3)
End Sub

Private Sub UpDown1_DownClick()
    With Me
        If IsNumeric(.Thickness.Text) Then
            If Thickness.Text <= 20 And .Thickness.Text > 1 Then
                Thickness.Text = CInt(Thickness.Text) - 1
            End If
        Else
            MsgBox "Please enter a number!", vbExclamation
        End If
    End With
End Sub

Private Sub UpDown1_UpClick()
    With Me
        If IsNumeric(.Thickness.Text) Then
            If Thickness.Text <= 19 And .Thickness.Text > 0 Then
                Thickness.Text = CInt(Thickness.Text) + 1
            End If
        Else
            MsgBox "Please enter a number!", vbExclamation
        End If
    End With
End Sub

Private Sub UpDown2_Change()
    With Me
        If IsNumeric(.ButtonPad.Text) Then
            If ButtonPad.Text <= 200 And .ButtonPad.Text > 1 Then
                ButtonPad.Text = CInt(ButtonPad.Text) - 1
            End If
        Else
            MsgBox "Please enter a number!", vbExclamation
        End If
    End With
End Sub

Private Sub UpDown2_UpClick()
    With Me
        If IsNumeric(.ButtonPad.Text) Then
            If ButtonPad.Text <= 199 And .ButtonPad.Text > 0 Then
                ButtonPad.Text = CInt(ButtonPad.Text) + 1
            End If
        Else
            MsgBox "Please enter a number!", vbExclamation
        End If
    End With
End Sub
