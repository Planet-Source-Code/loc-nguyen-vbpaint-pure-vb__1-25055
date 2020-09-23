VERSION 5.00
Begin VB.Form FrmPaint 
   Caption         =   "VBPaint"
   ClientHeight    =   4935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPaint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmEraser 
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
      Begin VB.PictureBox PicEraser 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   60
         Index           =   0
         Left            =   240
         ScaleHeight     =   30
         ScaleWidth      =   30
         TabIndex        =   24
         Top             =   465
         Width           =   60
      End
      Begin VB.PictureBox PicEraser 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   90
         Index           =   1
         Left            =   585
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   23
         Top             =   450
         Width           =   90
      End
      Begin VB.PictureBox PicEraser 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   2
         Left            =   930
         ScaleHeight     =   90
         ScaleWidth      =   90
         TabIndex        =   22
         Top             =   435
         Width           =   120
      End
      Begin VB.PictureBox PicEraser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   3
         Left            =   1275
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   21
         Top             =   420
         Width           =   150
      End
   End
   Begin VB.Timer TmrAirbrush 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame FrmColors 
      Height          =   855
      Left            =   3240
      TabIndex        =   10
      Top             =   -45
      Width           =   2175
      Begin VB.PictureBox PicColors 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   690
         Picture         =   "FrmPaint.frx":030A
         ScaleHeight     =   675
         ScaleWidth      =   1440
         TabIndex        =   12
         Top             =   135
         Width           =   1440
      End
      Begin VB.PictureBox PicFore 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   105
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   11
         Top             =   225
         Width           =   375
      End
      Begin VB.PictureBox PicBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame FrmTools 
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   -45
      Width           =   1455
      Begin VB.ComboBox CmbTools 
         Height          =   315
         ItemData        =   "FrmPaint.frx":35EC
         Left            =   120
         List            =   "FrmPaint.frx":360E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label LblTools 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tools:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.Frame FrmOptions 
      Height          =   855
      Left            =   1500
      TabIndex        =   3
      Top             =   -45
      Width           =   1695
   End
   Begin VB.PictureBox PicWS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   3765
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   840
      Width           =   5415
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   3510
         Width           =   5100
      End
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   3510
         Left            =   5100
         TabIndex        =   17
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox PicDither 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2400
         Picture         =   "FrmPaint.frx":366B
         ScaleHeight     =   240
         ScaleWidth      =   2640
         TabIndex        =   16
         Top             =   3120
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.PictureBox PicHanSE 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   4455
         MousePointer    =   8  'Size NW SE
         Picture         =   "FrmPaint.frx":57AD
         ScaleHeight     =   120
         ScaleWidth      =   105
         TabIndex        =   6
         Top             =   3375
         Width           =   105
      End
      Begin VB.PictureBox PicHanS 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   2040
         MousePointer    =   7  'Size N S
         Picture         =   "FrmPaint.frx":58AF
         ScaleHeight     =   120
         ScaleWidth      =   105
         TabIndex        =   5
         Top             =   3375
         Width           =   105
      End
      Begin VB.PictureBox PicHanE 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   4455
         MousePointer    =   9  'Size W E
         Picture         =   "FrmPaint.frx":59B1
         ScaleHeight     =   120
         ScaleWidth      =   105
         TabIndex        =   4
         Top             =   1560
         Width           =   105
      End
      Begin VB.PictureBox PicJoinScr 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5085
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   19
         Top             =   3495
         Width           =   285
      End
      Begin VB.PictureBox PicPaint 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   3375
         ScaleWidth      =   4455
         TabIndex        =   1
         Top             =   0
         Width           =   4455
         Begin VB.PictureBox PicTemp 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2520
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   25
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox TxtText 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1560
            MousePointer    =   3  'I-Beam
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Image ImgEraser 
            Height          =   480
            Index           =   3
            Left            =   1920
            Picture         =   "FrmPaint.frx":5AB3
            Top             =   600
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image ImgEraser 
            Height          =   480
            Index           =   2
            Left            =   1320
            Picture         =   "FrmPaint.frx":5DBD
            Top             =   600
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image ImgEraser 
            Height          =   480
            Index           =   1
            Left            =   720
            Picture         =   "FrmPaint.frx":60C7
            Top             =   600
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image ImgEraser 
            Height          =   480
            Index           =   0
            Left            =   120
            Picture         =   "FrmPaint.frx":63D1
            Top             =   600
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Shape ShpOval 
            Height          =   375
            Left            =   120
            Shape           =   2  'Oval
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Shape ShpRect 
            Height          =   375
            Left            =   600
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Line ShpLine 
            Visible         =   0   'False
            X1              =   1080
            X2              =   1440
            Y1              =   120
            Y2              =   480
         End
         Begin VB.Shape ShpBorder 
            BorderStyle     =   3  'Dot
            Height          =   375
            Left            =   2040
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
      End
   End
   Begin VB.Label LblStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2730
      TabIndex        =   15
      Top             =   4680
      Width           =   2685
   End
   Begin VB.Label LblStatus 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   4680
      Width           =   2685
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu MnuFileSep01 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuFileSep02 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuImage 
      Caption         =   "&Image"
      Begin VB.Menu MnuImageInvert 
         Caption         =   "&Invert Colors"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuImageAttributes 
         Caption         =   "&Attributes..."
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuImageSep00 
         Caption         =   "-"
      End
      Begin VB.Menu MnuImageBW 
         Caption         =   "&Black and White (1 bpp)"
         Shortcut        =   ^B
      End
      Begin VB.Menu MnuImageSep01 
         Caption         =   "-"
      End
      Begin VB.Menu MnuImageClear 
         Caption         =   "&Clear Image"
      End
   End
   Begin VB.Menu MnuColors 
      Caption         =   "&Colors"
      Begin VB.Menu MnuColorsEdit 
         Caption         =   "&Edit Colors..."
      End
   End
   Begin VB.Menu MnuFilter 
      Caption         =   "&Filter"
      Begin VB.Menu MnuFilterBlur 
         Caption         =   "&Blur"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuHelpAbout 
         Caption         =   "&About VBPaint"
      End
   End
End
Attribute VB_Name = "FrmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private oX As Single, oY As Single
Private RadX As Single, RadY As Single
Private oPenX As Single, oPenY As Single
Private DoAirbrush As Boolean
Private AirbrushX As Single, AirbrushY As Single
Private AirbrushRad As Integer
Private oEraseX As Single, oEraseY As Single
Private InitX As Integer, InitY As Integer
Private OnText As Boolean
Private EraserIndex As Integer
Public ImgChanged As Boolean
Public CurColor As Long
Public CurFile As String, CurFullFile As String

Private Sub CmbTools_Click()
    If CmbTools.ListIndex = 1 Then
        PicPaint.MouseIcon = ImgEraser(EraserIndex).Picture
        PicPaint.MousePointer = 99
        FrmEraser.Visible = True
    Else
        FrmEraser.Visible = False
    End If
End Sub

Private Sub Form_Load()
    CmbTools.ListIndex = 0
    CurColor = RGB(0, 0, 0)
    AirbrushRad = 5
    EraserIndex = 3
    PicHanS.Left = PicPaint.Width / 2 - PicHanS.Width / 2
    PicHanE.Top = PicPaint.Height / 2 - PicHanE.Height / 2
    PicHanSE.Top = PicPaint.Height
    PicHanSE.Left = PicPaint.Width
    CurFile = "Untitled"
    Me.Caption = App.Title & " [" & CurFile & "]"
    FrmEraser.Left = FrmOptions.Left
    FrmEraser.Top = FrmOptions.Top
    PicPaint_Change
End Sub

Public Sub Form_Resize()
    On Error GoTo ErrorHandler
    PicWS.Width = Me.Width - 120
    PicWS.Height = Me.Height - 1800
    LblStatus(0).Top = Me.Height - 945
    LblStatus(1).Top = Me.Height - 945
    LblStatus(0).Width = (Me.Width - 165) / 2
    LblStatus(1).Width = Me.Width - (LblStatus(0).Width + 165)
    LblStatus(1).Left = LblStatus(0).Width + 45
    HScroll1.Width = PicWS.Width - 315
    VScroll1.Height = PicWS.Height - 315
    HScroll1.Top = PicWS.Height - 315
    VScroll1.Left = PicWS.Width - 315
    PicJoinScr.Left = HScroll1.Width
    PicJoinScr.Top = VScroll1.Height
    If PicPaint.Width + 105 >= PicWS.Width - 315 Then
        HScroll1.Enabled = True
    Else
        HScroll1.Enabled = False
        PicPaint.Left = 0
    End If
    If PicPaint.Height + 120 >= PicWS.Height - 315 Then
        VScroll1.Enabled = True
    Else
        VScroll1.Enabled = False
        PicPaint.Top = 0
    End If
    HScroll1.Max = -(PicWS.ScaleWidth - (PicPaint.ScaleWidth + 360)) / 15
    VScroll1.Max = -(PicWS.ScaleHeight - (PicPaint.ScaleHeight + 360)) / 15
    HScroll1.LargeChange = HScroll1.Max * (PicWS.Width / PicPaint.Width)
    VScroll1.LargeChange = VScroll1.Max * (PicWS.Height / PicPaint.Height)
    HScroll1.SmallChange = PicPaint.Width / HScroll1.LargeChange
    VScroll1.SmallChange = PicPaint.Height / VScroll1.LargeChange
ErrorHandler:
    If Me.WindowState <> 1 Then
        If Me.Width < 5535 Then Me.Width = 5535
        If Me.Height < 5625 Then Me.Height = 5625
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MnuFileExit_Click
    Cancel = 1
End Sub

Private Sub MnuColorsEdit_Click()
    FrmEditColors.Show 1
End Sub

Private Sub MnuFileExit_Click()
    If ImgChanged = True Then
        MsgRes = MsgBox("Save changes to " & Chr(34) & CurFile & Chr(34) & "?", vbExclamation + vbYesNoCancel, "VBPaint")
        If MsgRes = vbYes Then
            MnuFileSave_Click
            End
        ElseIf MsgRes = vbNo Then
            End
        End If
    Else
        End
    End If
End Sub

Private Sub MnuFileNew_Click()
    If ImgChanged = True Then
        MsgRes = MsgBox("Save changes to " & Chr(34) & CurFile & Chr(34) & "?", vbExclamation + vbYesNoCancel, "VBPaint")
        If MsgRes = vbYes Then
            MnuFileSave_Click
            CurFile = "Untitled"
            Me.Caption = App.Title & " [" & CurFile & "]"
            PicPaint.Cls
            PicPaint.Picture = Nothing
            ImgChanged = False
        ElseIf MsgRes = vbNo Then
            CurFile = "Untitled"
            Me.Caption = App.Title & " [" & CurFile & "]"
            PicPaint.Cls
            PicPaint.Picture = Nothing
            ImgChanged = False
        End If
    Else
        CurFile = "Untitled"
        Me.Caption = App.Title & " [" & CurFile & "]"
        PicPaint.Cls
        PicPaint.Picture = Nothing
        ImgChanged = False
    End If
End Sub

Private Sub MnuFileOpen_Click()
    If ImgChanged = True Then
        MsgRes = MsgBox("Save changes to " & Chr(34) & CurFile & Chr(34) & "?", vbExclamation + vbYesNoCancel, "VBPaint")
        If MsgRes = vbYes Then
            MnuFileSave_Click
            LoadDialog "Open Image...", True
        ElseIf MsgRes = vbNo Then
            LoadDialog "Open Image...", True
        End If
    Else
        LoadDialog "Open Image...", True
    End If
End Sub

Private Sub MnuFilePrint_Click()
    oScale = PicPaint.ScaleMode
    PicPaint.ScaleMode = vbCentimeters
    ImageLeft = (21 - PicPaint.ScaleWidth) / 2
    ImageTop = (29.7 - PicPaint.ScaleHeight) / 2
    Printer.ScaleMode = vbCentimeters
    Printer.PaintPicture PicPaint.Image, ImageLeft, ImageTop
    Printer.EndDoc
    PicPaint.ScaleMode = oScale
End Sub

Private Sub MnuFileSave_Click()
    If CurFile = "Untitled" Then
        MnuFileSaveAs_Click
    Else
        SavePicture PicPaint.Image, CurFullFile
        ImgChanged = False
    End If
End Sub

Private Sub MnuFileSaveAs_Click()
    LoadDialog "Save Image As...", False
End Sub

Private Sub MnuFilterBlur_Click()
    On Error Resume Next
    PicTemp.Width = PicPaint.Width
    PicTemp.Height = PicPaint.Height
    Dim PixCol(0 To 3) As Long
    Dim ActCol(0 To 2) As Long
    For k = 0 To PicTemp.Height - 15 Step 30
        For j = 0 To PicTemp.Width - 15 Step 30
            PixCol(0) = PicPaint.Point(j, k)
            PixCol(1) = PicPaint.Point(j, k + 15)
            PixCol(2) = PicPaint.Point(j + 15, k)
            PixCol(3) = PicPaint.Point(j + 15, k + 15)
            ActCol(0) = (RGBCon(PixCol(0), 1) + RGBCon(PixCol(1), 1) + RGBCon(PixCol(2), 1) + RGBCon(PixCol(3), 1)) / 4
            ActCol(1) = (RGBCon(PixCol(0), 2) + RGBCon(PixCol(1), 2) + RGBCon(PixCol(2), 2) + RGBCon(PixCol(3), 2)) / 4
            ActCol(2) = (RGBCon(PixCol(0), 3) + RGBCon(PixCol(1), 3) + RGBCon(PixCol(2), 3) + RGBCon(PixCol(3), 3)) / 4
            PicTemp.PSet (j, k), RGB(ActCol(0), ActCol(1), ActCol(2))
            PicTemp.PSet (j, k + 15), RGB(ActCol(0), ActCol(1), ActCol(2))
            PicTemp.PSet (j + 15, k), RGB(ActCol(0), ActCol(1), ActCol(2))
            PicTemp.PSet (j + 15, k + 15), RGB(ActCol(0), ActCol(1), ActCol(2))
        Next j
        LblStatus(0).Caption = "Blurring: " & Format(50 * (k + j / (PicPaint.Width - 15)) / (PicPaint.Height - 15), "0.00") & "%"
        DoEvents
    Next k
    Dim ActCol1(0 To 2) As Long
    Dim ActCol2(0 To 2) As Long
    Dim DrawCol(0 To 3) As Long
    Dim CurCol As Long
    For k = 0 To PicTemp.Height - 15 Step 30
        For j = 0 To PicTemp.Width - 15 Step 30
            PixCol(0) = PicTemp.Point(j - 15, k - 15)
            PixCol(1) = PicTemp.Point(j + 30, k - 15)
            PixCol(2) = PicTemp.Point(j - 15, k + 30)
            PixCol(3) = PicTemp.Point(j + 30, k + 30)
            CurCol = PicTemp.Point(j, k)
            ActCol1(0) = RGBCon(CurCol, 1)
            ActCol1(1) = RGBCon(CurCol, 2)
            ActCol1(2) = RGBCon(CurCol, 3)
            For l = 0 To 3
                ActCol2(0) = RGBCon(PixCol(l), 1)
                ActCol2(1) = RGBCon(PixCol(l), 2)
                ActCol2(2) = RGBCon(PixCol(l), 3)
                DrawCol(l) = RGB(ActCol2(0) + (3 * (ActCol1(0) - ActCol2(0))) / 4, ActCol2(1) + (3 * (ActCol1(1) - ActCol2(1))) / 4, ActCol2(2) + (3 * (ActCol1(2) - ActCol2(2))) / 4)
            Next l
            PicPaint.PSet (j, k), DrawCol(0)
            PicPaint.PSet (j + 15, k), DrawCol(1)
            PicPaint.PSet (j, k + 15), DrawCol(2)
            PicPaint.PSet (j + 15, k + 15), DrawCol(3)
        Next j
        LblStatus(0).Caption = "Blurring: " & Format(50 + 50 * (k + j / (PicPaint.Width - 15)) / (PicPaint.Height - 15), "0.00") & "%"
        DoEvents
    Next k
    LblStatus(0).Caption = ""
    ImgChanged = True
End Sub

Private Sub MnuHelpAbout_Click()
    FrmAboutPaint.Show 1
End Sub

Private Sub MnuImageAttributes_Click()
    FrmAttributes.Show 1
End Sub

Private Sub MnuImageBW_Click()
    GradInt = (PicDither.Width / 15) / 11
    For v = 0 To PicPaint.Height / 15 - 1
        For h = 0 To PicPaint.Width / 15 - 1
            CurRGB = (RGBCon(PicPaint.Point(h * 15, v * 15), 1) + RGBCon(PicPaint.Point(h * 15, v * 15), 2) + RGBCon(PicPaint.Point(h * 15, v * 15), 3)) / 3
            PalLoc = 0
            Do Until GradInt * PalLoc > (CurRGB / 255) * (PicDither.Width / 15)
                PalLoc = PalLoc + 1
            Loop
            PalLoc = PalLoc - 1
            Sclh = h
            Sclv = v
            If h > 15 Then Sclh = h Mod 16
            If v > 15 Then Sclv = v Mod 16
            PicPaint.PSet (h * 15, v * 15), PicDither.Point(GradInt * 15 * PalLoc + Sclh * 15, Sclv * 15)
        Next h
        LblStatus(0).Caption = "Converting to 2 colors: " & Format(100 * (v + h / (PicPaint.Width / 15 - 1)) / (PicPaint.Height / 15 - 1), "0.00") & "%"
        DoEvents
    Next v
    LblStatus(0).Caption = ""
    ImgChanged = True
End Sub

Private Sub MnuImageClear_Click()
    PicPaint.Cls
    PicPaint.Picture = Nothing
    ImgChanged = True
End Sub

Private Sub MnuImageInvert_Click()
    For v = 0 To PicPaint.Height / 15 - 1
        For h = 0 To PicPaint.Width / 15 - 1
            CurR = Abs(RGBCon(PicPaint.Point(h * 15, v * 15), 1) - 255)
            CurG = Abs(RGBCon(PicPaint.Point(h * 15, v * 15), 2) - 255)
            CurB = Abs(RGBCon(PicPaint.Point(h * 15, v * 15), 3) - 255)
            PicPaint.PSet (h * 15, v * 15), RGB(CurR, CurG, CurB)
        Next h
        LblStatus(0).Caption = "Inverting colors: " & Format(100 * (v + h / (PicPaint.Width / 15 - 1)) / (PicPaint.Height / 15 - 1), "0.00") & "%"
        DoEvents
    Next v
    LblStatus(0).Caption = ""
    ImgChanged = True
End Sub

Private Sub PicColors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        PicFore.BackColor = PicColors.Point(X, Y)
        CurColor = PicColors.Point(X, Y)
    ElseIf Button = 2 Then
        PicBack.BackColor = PicColors.Point(X, Y)
        PicPaint.BackColor = PicColors.Point(X, Y)
    End If
End Sub

Private Sub PicEraser_Click(Index As Integer)
    For k = 0 To 3
        PicEraser(k).BackColor = RGB(0, 0, 0)
    Next k
    PicEraser(Index).BackColor = RGB(255, 255, 255)
    EraserIndex = Index
    CmbTools_Click
End Sub

Private Sub PicHanE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        InitX = X
        InitY = Y
    End If
End Sub

Public Sub PicHanE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandler
    If Button = 1 Then
        PicHanE.Left = PicHanE.Left + X - InitX
        PicPaint.Width = PicHanE.Left - 15
        HanEMove
        Form_Resize
    End If
    Exit Sub
ErrorHandler:
    PicHanE.Left = PicPaint.Width
End Sub

Public Sub PicHanS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandler
    If Button = 1 Then
        PicHanS.Top = PicHanS.Top + Y - InitY
        PicPaint.Height = PicHanS.Top - 15
        HanSMove
        Form_Resize
    End If
    Exit Sub
ErrorHandler:
    PicHanS.Top = PicPaint.Height
End Sub

Public Sub PicHanSE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandler
    If Button = 1 Then
        PicHanSE.Left = PicHanSE.Left + X - InitX
        PicHanSE.Top = PicHanSE.Top + Y - InitY
        PicPaint.Width = PicHanSE.Left
        PicPaint.Height = PicHanSE.Top
        HanSEMove
        Form_Resize
    End If
    Exit Sub
ErrorHandler:
    PicHanS.Top = PicPaint.Height
    PicHanE.Left = PicPaint.Width
    PicHanSE.Top = PicPaint.Height
    PicHanSE.Left = PicPaint.Width
End Sub

Private Sub PicPaint_Change()
    Form_Resize
End Sub

Private Sub PicPaint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        oX = X
        oY = Y
        Select Case CmbTools.ListIndex
            Case 0
                'Selection
                ShpBorder.Width = 15
                ShpBorder.Height = 15
                ShpBorder.BorderColor = RGB(Abs(RGBCon(PicPaint.BackColor, 1) - 255), Abs(RGBCon(PicPaint.BackColor, 2) - 255), Abs(RGBCon(PicPaint.BackColor, 3) - 255))
                ShpBorder.Visible = True
            Case 2
                'Fill
                Fill X, Y, PicPaint.Point(X, Y), CurColor, PicPaint
            Case 3
                'Color Picker
                PicFore.BackColor = PicPaint.Point(X, Y)
                CurColor = PicPaint.Point(X, Y)
            Case 5
                'Airbrush
                DoAirbrush = True
            Case 6
                'Text
                ShpBorder.Width = 15
                ShpBorder.Height = 15
                ShpBorder.BorderColor = RGB(Abs(RGBCon(PicPaint.BackColor, 1) - 255), Abs(RGBCon(PicPaint.BackColor, 2) - 255), Abs(RGBCon(PicPaint.BackColor, 3) - 255))
                TxtText.ForeColor = CurColor
                ShpBorder.Visible = True
            Case 7
                'Line
                ShpLine.X2 = ShpLine.X1 + 15
                ShpLine.Y2 = ShpLine.Y1 + 15
                ShpLine.BorderColor = CurColor
                ShpLine.Visible = True
            Case 8
                'Rectangle
                ShpRect.Width = 15
                ShpRect.Height = 15
                ShpRect.BorderColor = CurColor
                ShpRect.Visible = True
            Case 9
                'Ellipse
                ShpOval.Width = 15
                ShpOval.Height = 15
                ShpOval.BorderColor = CurColor
                ShpOval.Visible = True
        End Select
        ImgChanged = True
    ElseIf Button = 2 Then
        Select Case CmbTools.ListIndex
            Case 3
                'Color Picker
                PicBack.BackColor = PicPaint.Point(X, Y)
                PicPaint.BackColor = PicPaint.Point(X, Y)
        End Select
    End If
End Sub

Private Sub PicPaint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RadX = Abs(X - oX)
    RadY = Abs(Y - oY)
    Select Case CmbTools.ListIndex
        Case 0
            'Selection
            PicPaint.MousePointer = 2
            If X < oX Then ShpBorder.Left = X Else ShpBorder.Left = oX
            If Y < oY Then ShpBorder.Top = Y Else ShpBorder.Top = oY
            ShpBorder.Width = RadX + 15
            ShpBorder.Height = RadY + 15
        Case 1
            'Eraser
            If Button = 1 Then
                For v = Y To Y + (2 * (EraserIndex + 1)) * 15 Step 15
                    For h = X To X + (2 * (EraserIndex + 1)) * 15 Step 15
                        PicPaint.Line (h - X + oEraseX, v - Y + oEraseY)-(h, v), PicPaint.BackColor
                    Next h
                Next v
            End If
            oEraseX = X
            oEraseY = Y
        Case 2
            'Fill
            PicPaint.MousePointer = 0
        Case 3
            'Color Picker
            PicPaint.MousePointer = 0
        Case 4
            'Pencil
            PicPaint.MousePointer = 0
            If Button = 1 Then PicPaint.Line (oPenX, oPenY)-(X, Y), CurColor
            oPenX = X
            oPenY = Y
        Case 5
            'Airbrush
            PicPaint.MousePointer = 0
            If Button = 1 Then DoAirbrush = True
            AirbrushX = X
            AirbrushY = Y
        Case 6
            'Text
            PicPaint.MousePointer = 2
            If OnText = False Then
                If X < oX Then ShpBorder.Left = X Else ShpBorder.Left = oX
                If Y < oY Then ShpBorder.Top = Y Else ShpBorder.Top = oY
                ShpBorder.Width = RadX + 15
                ShpBorder.Height = RadY + 15
            End If
        Case 7
            'Line
            PicPaint.MousePointer = 2
            ShpLine.X1 = oX
            ShpLine.Y1 = oY
            If X < oX Then XOffset = -1 Else XOffset = 1
            If Y < oY Then YOffset = -1 Else YOffset = 1
            ShpLine.X2 = ShpLine.X1 + XOffset * RadX
            ShpLine.Y2 = ShpLine.Y1 + YOffset * RadY
        Case 8
            'Rectangle
            PicPaint.MousePointer = 2
            If X < oX Then ShpRect.Left = X Else ShpRect.Left = oX
            If Y < oY Then ShpRect.Top = Y Else ShpRect.Top = oY
            ShpRect.Width = RadX + 15
            ShpRect.Height = RadY + 15
        Case 9
            'Ellipse
            PicPaint.MousePointer = 2
            ShpOval.Left = oX - RadX
            ShpOval.Top = oY - RadY
            ShpOval.Width = RadX * 2 + 15
            ShpOval.Height = RadY * 2 + 15
    End Select
    If Button = 1 Then
        LblStatus(1).Caption = RadX / 15 + 1 & "x" & RadY / 15 + 1
    Else
        LblStatus(1).Caption = X / 15 & ", " & Y / 15
    End If
End Sub

Private Sub PicPaint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        Select Case CmbTools.ListIndex
            Case 0
                'Selection
                ShpBorder.Visible = False
            Case 1
                'Eraser
                For v = Y To Y + (2 * (EraserIndex + 1)) * 15 Step 15
                    PicPaint.Line (X, v)-(X + (2 * (EraserIndex + 1)) * 15, v), PicPaint.BackColor
                Next v
            Case 4
                'Pencil
                PicPaint.PSet (X, Y), CurColor
            Case 5
                'Airbrush
                DoAirbrush = False
            Case 6
                'Text
                If oX <> X And oY <> Y Then
                    If X < oX Then oX = X
                    If Y < oY Then oY = Y
                    ShpBorder.Left = ShpBorder.Left - 15
                    ShpBorder.Width = ShpBorder.Width + 15
                    ShpBorder.Top = ShpBorder.Top - 15
                    ShpBorder.Height = ShpBorder.Height + 15
                    TxtText.Left = oX
                    TxtText.Top = oY
                    TxtText.Width = RadX
                    TxtText.Height = RadY
                    TxtText.Visible = True
                    TxtText.SetFocus
                    OnText = True
                End If
            Case 7
                'Line
                ShpLine.Visible = False
                PicPaint.Line (oX, oY)-(X, Y), CurColor
                PicPaint.PSet (X, Y), CurColor
            Case 8
                'Rectangle
                ShpRect.Visible = False
                If oX <> X And oY <> Y Then
                    If X < oX Then oX = X
                    If Y < oY Then oY = Y
                    PicPaint.Line (oX, oY)-(oX + RadX, oY), CurColor
                    PicPaint.Line (oX + RadX, oY)-(oX + RadX, oY + RadY), CurColor
                    PicPaint.Line (oX + RadX, oY + RadY)-(oX, oY + RadY), CurColor
                    PicPaint.Line (oX, oY + RadY)-(oX, oY), CurColor
                End If
            Case 9
                'Ellipse
                ShpOval.Visible = False
                If oX <> X And oY <> Y Then
                    For CurX = -RadX To RadX Step 15
                        CurY = RadY * Sqr(1 - (CurX / RadX) ^ 2)
                        vbX = CurX + oX
                        vbY = (oY * 2 - CurY) - oY
                        PicPaint.PSet (vbX, vbY), CurColor
                        vbY = (oY * 2 + CurY) - oY
                        PicPaint.PSet (vbX, vbY), CurColor
                    Next CurX
                    For CurY = -RadY To RadY Step 15
                        CurX = RadX * Sqr(1 - (CurY / RadY) ^ 2)
                        vbY = CurY + oY
                        vbX = (oX * 2 - CurX) - oX
                        PicPaint.PSet (vbX, vbY), CurColor
                        vbX = (oX * 2 + CurX) - oX
                        PicPaint.PSet (vbX, vbY), CurColor
                    Next CurY
                End If
        End Select
    End If
End Sub

Private Sub Fill(XFill As Single, YFill As Single, ColorFill As Long, ColorToFill As Long, PicFill As PictureBox)
    On Error GoTo ErrorHandler
    Dim nColorFill As Long
    PicFill.PSet (XFill, YFill), ColorToFill
    nColorFill = PicFill.Point(XFill - 15, YFill)
    If nColorFill = ColorFill Then Fill XFill - 15, YFill, nColorFill, ColorToFill, PicFill
    nColorFill = PicFill.Point(XFill + 15, YFill)
    If nColorFill = ColorFill Then Fill XFill + 15, YFill, nColorFill, ColorToFill, PicFill
    nColorFill = PicFill.Point(XFill, YFill - 15)
    If nColorFill = ColorFill Then Fill XFill, YFill - 15, nColorFill, ColorToFill, PicFill
    nColorFill = PicFill.Point(XFill, YFill + 15)
    If nColorFill = ColorFill Then Fill XFill, YFill + 15, nColorFill, ColorToFill, PicFill
    Exit Sub
ErrorHandler:
    DoEvents
End Sub

Private Sub PicPaint_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicWS_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub TmrAirbrush_Timer()
    If DoAirbrush = True Then
        Randomize Timer
        AirLen = Int(Rnd * (AirbrushRad + 1)) * 15
        AirVect = Rnd * (2 * 3.14159265358979 + 1)
        AirX = Cos(AirVect) * AirLen
        AirY = Sin(AirVect) * AirLen
        AirX = AirX + AirbrushX
        AirY = (AirbrushY * 2 - AirY) - AirbrushY
        PicPaint.PSet (AirX, AirY), CurColor
    End If
End Sub

Private Sub TxtText_LostFocus()
    TxtText.Visible = False
    ShpBorder.Visible = False
    OnText = False
    PicPaint.Font = TxtText.Font
    PicPaint.FontBold = TxtText.FontBold
    PicPaint.FontItalic = TxtText.FontItalic
    PicPaint.FontName = TxtText.FontName
    PicPaint.FontSize = TxtText.FontSize
    oForeColor = PicPaint.ForeColor
    PicPaint.ForeColor = CurColor
    PicPaint.CurrentX = TxtText.Left
    PicPaint.CurrentY = TxtText.Top
    'Parse MultiLine text
    oStart = 1
    If InStr(1, TxtText.Text, Chr(13)) = 0 Then
        PicPaint.Print TxtText.Text
    Else
        Do Until InStr(oStart, TxtText.Text, Chr(13)) = 0
            PicPaint.Print Mid(TxtText.Text, oStart, InStr(oStart, TxtText.Text, Chr(13)) - oStart)
            oStart = InStr(oStart, TxtText.Text, Chr(13)) + 2
            PicPaint.CurrentX = TxtText.Left
        Loop
        PicPaint.Print Mid(TxtText.Text, oStart, Len(TxtText.Text) - oStart + 1)
    End If
    TxtText.Text = ""
    PicPaint.ForeColor = oForeColor
End Sub

Private Function RGBCon(RGBColor, CType As Integer)
    'Convert RGB integer to R (CType = 1), G (CType = 2), or B (CType = 3) integer
    If CType = 1 Then
        CType = 3
    ElseIf CType = 3 Then
        CType = 1
    End If
    HRGB = Left("000000", 6 - Len(Hex(RGBColor))) & Hex(RGBColor)
    RGBCon = Val("&H" & Mid(HRGB, 1 + 2 * (CType - 1), 2))
End Function

Private Sub LoadDialog(DlgTitle As String, Optional DlgOpen As Boolean)
    FrmFilePaint.Caption = DlgTitle
    FrmFilePaint.IsOpen = CBool(DlgOpen)
    FrmFilePaint.Show 1
End Sub

Private Sub VScroll1_Change()
    PicPaint.Top = -VScroll1.Value * 15
    PicHanS.Top = PicPaint.Height - VScroll1.Value * 15
    PicHanE.Top = PicPaint.Height / 2 - PicHanE.Height / 2 - VScroll1.Value * 15
    PicHanSE.Top = PicPaint.Height - VScroll1.Value * 15
End Sub

Private Sub VScroll1_Scroll()
    PicPaint.Top = -VScroll1.Value * 15
    PicHanS.Top = PicPaint.Height - VScroll1.Value * 15
    PicHanE.Top = PicPaint.Height / 2 - PicHanE.Height / 2 - VScroll1.Value * 15
    PicHanSE.Top = PicPaint.Height - VScroll1.Value * 15
End Sub

Private Sub HScroll1_Scroll()
    PicPaint.Left = -HScroll1.Value * 15
    PicHanS.Left = PicPaint.Width / 2 - PicHanS.Width / 2 - HScroll1.Value * 15
    PicHanE.Left = PicPaint.Width - HScroll1.Value * 15
    PicHanSE.Left = PicPaint.Width - HScroll1.Value * 15
End Sub

Private Sub HScroll1_Change()
    PicPaint.Left = -HScroll1.Value * 15
    PicHanS.Left = PicPaint.Width / 2 - PicHanS.Width / 2 - HScroll1.Value * 15
    PicHanE.Left = PicPaint.Width - HScroll1.Value * 15
    PicHanSE.Left = PicPaint.Width - HScroll1.Value * 15
End Sub

Public Sub HanEMove()
    PicHanS.Left = PicPaint.Width / 2 - PicHanS.Width / 2
    PicHanSE.Top = PicPaint.Height
    PicHanSE.Left = PicPaint.Width
End Sub

Public Sub HanSMove()
    PicHanE.Top = PicPaint.Height / 2 - PicHanE.Height / 2
    PicHanSE.Top = PicPaint.Height
    PicHanSE.Left = PicPaint.Width
End Sub

Public Sub HanSEMove()
    PicHanS.Left = PicPaint.Width / 2 - PicHanS.Width / 2
    PicHanE.Top = PicPaint.Height / 2 - PicHanE.Height / 2
    PicHanS.Top = PicPaint.Height
    PicHanE.Left = PicPaint.Width
End Sub

Public Sub DoResize(ResWidth, ResHeight)
    Unload FrmAttributes
    PicTemp.Picture = PicPaint.Image
    PicTemp.Width = ResWidth * 15
    PicTemp.Height = ResHeight * 15
    For v = 0 To PicTemp.Height - 15 Step 15
        For h = 0 To PicTemp.Width - 15 Step 15
            PicTemp.PSet (h, v), PicPaint.Point((PicPaint.Width / PicTemp.Width) * h, (PicPaint.Height / PicTemp.Height) * v)
        Next h
        LblStatus(0).Caption = "Resizing: " & Format(100 * (v + h / (PicTemp.Width - 15)) / (PicTemp.Height - 15), "0.00") & "%"
        DoEvents
    Next v
    PicPaint.Picture = PicTemp.Image
    LblStatus(0).Caption = ""
    ImgChanged = True
End Sub

Private Sub PicWS_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TmpFile = Data.Files(1)
    Do Until InStr(TmpFile, "\") = 0
        TmpFile = Right(TmpFile, Len(TmpFile) - InStr(TmpFile, "\"))
    Loop
    If ImgChanged = True Then
        MsgRes = MsgBox("Save changes to " & Chr(34) & CurFile & Chr(34) & "?", vbExclamation + vbYesNoCancel, "VBPaint")
        If MsgRes = vbYes Then
            MnuFileSave_Click
            IsLoaded = True
        ElseIf MsgRes = vbNo Then
            IsLoaded = True
        End If
    Else
        IsLoaded = True
    End If
    If IsLoaded = True Then
        FrmPaint.PicPaint.Picture = LoadPicture(Data.Files(1))
        FrmPaint.CurFile = TmpFile
        FrmPaint.CurFullFile = Data.Files(1)
        FrmPaint.Caption = App.Title & " [" & FrmPaint.CurFile & "]"
        FrmPaint.ImgChanged = False
    End If
    FrmPaint.HanEMove
    FrmPaint.HanSMove
    FrmPaint.HanSEMove
End Sub
