VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "xFrame"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin xFrameProject.xFrame xFrame1 
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2566
      BorderColor     =   13619151
      Button          =   -1  'True
      ButtonColor     =   0
      ButtonHighlightColor=   11513775
      ButtonPin       =   -1  'True
      ColorScheme     =   0
      Caption         =   "Will Collapse on Left Click"
      DisplayPicture  =   -1  'True
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      ForeColor       =   0
      FontItalic      =   0   'False
      FontSize        =   8,25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      GradientBottom  =   15724527
      HeaderGradientBottom=   13619151
      HeaderGradientTop=   16316664
      Picture         =   "frmMain.frx":0000
      Begin VB.TextBox Text 
         Height          =   285
         Left            =   720
         TabIndex        =   26
         Text            =   "Text"
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   975
         Left            =   120
         TabIndex        =   20
         Text            =   "Text3"
         Top             =   360
         Width           =   3495
      End
   End
   Begin xFrameProject.xFrame xFrame1 
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2566
      Button          =   -1  'True
      Caption         =   "xFrame"
      DisplayPicture  =   -1  'True
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontSize        =   8,25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeaderGradientBottom=   12611136
      Picture         =   "frmMain.frx":059A
      Begin VB.TextBox Text4 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Text            =   "Text4"
         Top             =   360
         Width           =   3495
      End
   End
   Begin xFrameProject.xFrame xFrame1 
      Height          =   1455
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2566
      BorderColor     =   7645851
      Button          =   -1  'True
      ButtonColor     =   4487268
      ButtonHighlightColor=   7645851
      ColorScheme     =   2
      Caption         =   "xFrame"
      DisplayPicture  =   -1  'True
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      ForeColor       =   4487268
      FontItalic      =   0   'False
      FontSize        =   8,25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      GradientBottom  =   14938092
      HeaderGradientBottom=   7975330
      HeaderGradientTop=   14938092
      Picture         =   "frmMain.frx":0B34
   End
   Begin xFrameProject.xFrame xFrame1 
      Height          =   1695
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2990
      BackColor       =   12648447
      BorderColor     =   12298664
      Button          =   -1  'True
      ButtonColor     =   8283750
      ButtonHighlightColor=   12298664
      ColorScheme     =   3
      Caption         =   "xFrame"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      ForeColor       =   8283750
      FontItalic      =   0   'False
      FontSize        =   8,25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      GradientBottom  =   8421631
      HeaderGradientBottom=   14140358
      HeaderGradientTop=   16118000
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Display Picture"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enabled"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enable Gradient"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cboColourScheme 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         TabIndex        =   10
         Text            =   "xpBlue"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Show Button"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Show Button Pin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Frame Pinned"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
   End
   Begin xFrameProject.xFrame xFrame1 
      Height          =   1455
      Index           =   4
      Left            =   3960
      TabIndex        =   14
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2566
      BorderColor     =   13619151
      Button          =   -1  'True
      ButtonColor     =   0
      ButtonHighlightColor=   11513775
      ColorScheme     =   0
      Caption         =   "xFrame"
      Enabled         =   -1  'True
      EnableGradient  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      ForeColor       =   0
      FontItalic      =   0   'False
      FontSize        =   8,25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      FramePinned     =   -1  'True
      GradientBottom  =   15724527
      HeaderGradientBottom=   13619151
      HeaderGradientTop=   16316664
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Text            =   "Text5"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Test"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Test"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
   End
   Begin xFrameProject.xFrame xFrame1 
      Height          =   1455
      Index           =   5
      Left            =   3960
      TabIndex        =   15
      Top             =   1680
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2566
      Button          =   -1  'True
      ButtonPin       =   -1  'True
      Caption         =   "xFrame"
      DisplayPicture  =   -1  'True
      Enabled         =   -1  'True
      EnableGradient  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontSize        =   8,25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeaderGradientBottom=   12611136
      Picture         =   "frmMain.frx":0ECE
   End
   Begin xFrameProject.xFrame xFrame1 
      Height          =   1455
      Index           =   6
      Left            =   3960
      TabIndex        =   16
      Top             =   3240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2566
      BorderColor     =   7645851
      ButtonColor     =   4487268
      ButtonHighlightColor=   7645851
      ButtonPin       =   -1  'True
      ColorScheme     =   2
      Caption         =   "xFrame"
      Enabled         =   -1  'True
      EnableGradient  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      ForeColor       =   4487268
      FontItalic      =   -1  'True
      FontSize        =   8,25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      GradientBottom  =   14938092
      HeaderGradientBottom=   7975330
      HeaderGradientTop=   14938092
   End
   Begin xFrameProject.xFrame xFrame1 
      Height          =   4815
      Index           =   7
      Left            =   3960
      TabIndex        =   17
      Top             =   4800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8493
      BorderColor     =   12298664
      Button          =   -1  'True
      ButtonColor     =   8283750
      ButtonHighlightColor=   12298664
      ColorScheme     =   3
      Caption         =   "xFrame"
      DisplayPicture  =   -1  'True
      Enabled         =   -1  'True
      EnableGradient  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      ForeColor       =   8283750
      FontItalic      =   0   'False
      FontSize        =   8,25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      GradientBottom  =   15920108
      HeaderGradientBottom=   14140358
      HeaderGradientTop=   16118000
      Picture         =   "frmMain.frx":1468
   End
   Begin xFrameProject.xFrame xFrame1 
      Height          =   3015
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5318
      Button          =   -1  'True
      Caption         =   "xFrame Disabled"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontSize        =   8,25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeaderGradientBottom=   12611136
      HeaderGradientTop=   12611136
      Picture         =   "frmMain.frx":1802
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set Borders Ctrls"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   27
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2760
         TabIndex        =   0
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   $"frmMain.frx":1B9C
         Height          =   1335
         Index           =   1
         Left            =   1800
         TabIndex        =   29
         Top             =   1320
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click(Index As Integer)
          Me.xFrame1(8).FormatInnerControls FLAT_CUSTOM_COLOR, "160,160,160", "99,146, 206", "CommandButton|Textbox"
End Sub

Private Sub xFrame1_Click(Index As Integer)
'    MsgBox xFrame1(Index).Expanded
If Index = 8 Then
       xFrame1(Index).Resize_xFrames_Down xFrame1(Index), xFrame1(3)
       xFrame1_ReSize (Index)
       xFrame1_ReSize (3)
Else
       xFrame1(Index).Resize_xFrames_Up xFrame1(Index), xFrame1(Index + 1)
'       xFrame1_ReSize (Index)
       xFrame1_ReSize (Index + 1)
End If
End Sub

Private Sub xFrame1_ClickHeaderBar(Index As Integer)
If Index = 8 Then
'       xFrame1(Index).Resize_xFrames_Down xFrame1(Index), xFrame1(3)
'       xFrame1_ReSize (3)
End If
End Sub

Private Sub xFrame1_ClickImgLeft(Index As Integer)
'    MsgBox "Img Left clicked"
    If Index = 0 Then
        xFrame1(Index).Resize_xFrames_Left xFrame1(0), xFrame1(4)
        xFrame1(4).Expanded = False
        xFrame1(4).Expanded = True
    ElseIf Index = 4 Then
        xFrame1(Index).Resize_xFrames_Right xFrame1(4), xFrame1(0)
         xFrame1(4).Expanded = False
        xFrame1(4).Expanded = True
    End If

Me.Refresh
End Sub

Private Sub xFrame1_ReSize(Index As Integer)
    If Index = 0 Then
        xFrame1(Index).ResizeControlInFrame xFrame1(Index), Text3
    ElseIf Index = 1 Then
        xFrame1(Index).ResizeControlInFrame xFrame1(Index), Text4(0)
    ElseIf Index = 4 Then
        xFrame1(Index).ResizeControlInFrame xFrame1(Index), Text2, , True
        xFrame1(Index).ResizeControlInFrame xFrame1(Index), Text5, , True
    End If
End Sub




Private Sub Form_Resize()
On Error Resume Next

    xFrame1(4).Width = Me.ScaleWidth - Me.xFrame1(4).Left - 100
End Sub
Private Sub Form_Load()
    With cboColourScheme
        .AddItem "xpDefault"
        .AddItem "xpBlue"
        .AddItem "xpOliveGreen"
        .AddItem "xpSilver"
    End With
    Me.xFrame1(0).Collapse_On_Left_Click ' .CollapseOnLeftClick = True
    
    Me.xFrame1(1).Collapse_On_Left_Click
    
    Me.xFrame1(4).Collapse_On_Left_Click

End Sub
Private Sub cboColourScheme_Change()
    xFrame1(8).ColorScheme = cboColourScheme.ListIndex
End Sub

Private Sub cboColourScheme_Click()
    Call cboColourScheme_Change
End Sub

Private Sub chkOptions_Click(Index As Integer)
    Select Case Index
        Case 0
            xFrame1(8).DisplayPicture = chkOptions(0).Value
        Case 1
            xFrame1(8).Enabled = chkOptions(1).Value
        Case 2
            xFrame1(8).EnableGradient = chkOptions(2).Value
        Case 3
            xFrame1(8).Button = chkOptions(3).Value
        Case 4
            xFrame1(8).ButtonPin = chkOptions(4).Value
        Case 5
            xFrame1(8).FramePinned = chkOptions(5).Value
    End Select
End Sub


