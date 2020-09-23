VERSION 5.00
Begin VB.Form FrmAttributes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Attributes"
   ClientHeight    =   2235
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAttributes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrmResize 
      Caption         =   "Type"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   60
      Width           =   3255
      Begin VB.OptionButton OptResize 
         Caption         =   "&Resize"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptDim 
         Caption         =   "&Crop"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox TxtRatio 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Text            =   "0"
      Top             =   1140
      Width           =   735
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3600
      TabIndex        =   8
      Top             =   540
      Width           =   1215
   End
   Begin VB.Frame FrmUnits 
      Caption         =   "Units"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1500
      Width           =   3255
      Begin VB.OptionButton OptPixels 
         Caption         =   "&Pixels"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptCm 
         Caption         =   "C&m"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptInches 
         Caption         =   "&Inches"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox TxtHeight 
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Text            =   "0"
      Top             =   780
      Width           =   735
   End
   Begin VB.TextBox TxtWidth 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "0"
      Top             =   780
      Width           =   735
   End
   Begin VB.Label LblRatio 
      BackStyle       =   0  'Transparent
      Caption         =   "&Aspect Ratio:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1140
      Width           =   1335
   End
   Begin VB.Label LblHeight 
      BackStyle       =   0  'Transparent
      Caption         =   "&Height:"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   780
      Width           =   615
   End
   Begin VB.Label LblWidth 
      BackStyle       =   0  'Transparent
      Caption         =   "&Width:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   780
      Width           =   615
   End
End
Attribute VB_Name = "FrmAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private CurUnit As Integer
Private CurType As Integer
Private ActWidth As Double, ActHeight As Double
Private PixPerIn As Double
Private TxtWidthSel As Boolean

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim oWidth As Double, oHeight As Double
    Dim oUnit As Integer
    Dim PixRes As String
    oWidth = ActWidth
    oHeight = ActHeight
    oUnit = CurUnit
    OptPixels_Click
    If ActWidth < 1 Then PixRes = "Width is not a valid value."
    If ActHeight < 1 Then PixRes = "Height is not a valid value."
    If ActWidth < 1 And ActHeight < 1 Then PixRes = "Width and Height are not a valid values."
    If PixRes = "" Then
        If CurType = 1 Then
            FrmPaint.PicPaint.Width = ActWidth * 15
            FrmPaint.PicPaint.Height = ActHeight * 15
            FrmPaint.ImgChanged = True
            Unload Me
        ElseIf CurType = 2 Then
            FrmPaint.DoResize ActWidth, ActHeight
        End If
        FrmPaint.HanEMove
        FrmPaint.HanSMove
        FrmPaint.HanSEMove
        FrmPaint.Form_Resize
    Else
        ActWidth = oWidth
        ActHeight = oHeight
        CurUnit = oUnit
        TxtWidth.Text = oWidth
        TxtHeight.Text = oHeight
        MsgBox PixRes, vbExclamation, Me.Caption
    End If
End Sub

Private Sub Form_Load()
    PixPerIn = 96
    ActWidth = FrmPaint.PicPaint.Width / 15
    ActHeight = FrmPaint.PicPaint.Height / 15
    TxtRatio.Text = ActWidth / ActHeight
    TxtWidth.Text = Format(ActWidth, "0.0#")
    TxtHeight.Text = Format(ActHeight, "0.0#")
    CurUnit = 3
    CurType = 1
End Sub

Private Function Cm(IntPixels As Integer) As Double
    Cm = (IntPixels * 29.7) / (PixPerIn * 12)
End Function

Private Function Inches(IntPixels As Integer) As Double
    Inches = IntPixels / PixPerIn
End Function

Private Function Pixels(IntUnits As Double, Optional IsCm As Boolean) As Integer
    If IsCm = True Then
        Pixels = Int((IntUnits * (PixPerIn * 12)) / 29.7)
    Else
        Pixels = Int(IntUnits * PixPerIn)
    End If
End Function

Private Sub OptDim_Click()
    LblRatio.Enabled = False
    TxtRatio.Enabled = False
    CurType = 1
End Sub

Private Sub OptInches_Click()
    If CurUnit = 3 Then
        ActWidth = Inches(Int(ActWidth))
        ActHeight = Inches(Int(ActHeight))
    ElseIf CurUnit = 2 Then
        ActWidth = Inches(Pixels(ActWidth, True))
        ActHeight = Inches(Pixels(ActHeight, True))
    End If
    TxtWidth.Text = Format(ActWidth, "0.0#")
    TxtHeight.Text = Format(ActHeight, "0.0#")
    CurUnit = 1
End Sub

Private Sub OptCm_Click()
    If CurUnit = 3 Then
        ActWidth = Cm(Int(ActWidth))
        ActHeight = Cm(Int(ActHeight))
    ElseIf CurUnit = 1 Then
        ActWidth = Cm(Pixels(ActWidth, False))
        ActHeight = Cm(Pixels(ActHeight, False))
    End If
    TxtWidth.Text = Format(ActWidth, "0.0#")
    TxtHeight.Text = Format(ActHeight, "0.0#")
    CurUnit = 2
End Sub

Private Sub OptPixels_Click()
    If CurUnit = 2 Then
        ActWidth = Pixels(ActWidth, True)
        ActHeight = Pixels(ActHeight, True)
    ElseIf CurUnit = 1 Then
        ActWidth = Pixels(ActWidth, False)
        ActHeight = Pixels(ActHeight, False)
    End If
    TxtWidth.Text = Format(ActWidth, "0.0#")
    TxtHeight.Text = Format(ActHeight, "0.0#")
    CurUnit = 3
End Sub

Private Sub OptResize_Click()
    LblRatio.Enabled = True
    TxtRatio.Enabled = True
    CurType = 2
End Sub

Private Sub TxtHeight_GotFocus()
    TxtHeight.SelStart = 0
    TxtHeight.SelLength = Len(TxtHeight.Text)
    TxtWidthSel = False
End Sub

Private Sub TxtHeight_KeyUp(KeyCode As Integer, Shift As Integer)
    ActHeight = Val(TxtHeight.Text)
    On Error GoTo ErrorHandler
    If TxtWidthSel = False And CurType = 2 Then
        ActWidth = Val(TxtRatio.Text) * ActHeight
        TxtWidth.Text = Format(ActWidth, "0.0#")
    End If
ErrorHandler:
End Sub

Private Sub TxtHeight_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitTextInput(KeyAscii, "0123456789.")
End Sub

Private Sub TxtRatio_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitTextInput(KeyAscii, "0123456789.")
End Sub

Private Sub TxtRatio_KeyUp(KeyCode As Integer, Shift As Integer)
    ActWidth = Val(TxtRatio.Text) * ActHeight
    TxtWidth.Text = Format(ActWidth, "0.0#")
End Sub

Private Sub TxtWidth_GotFocus()
    TxtWidth.SelStart = 0
    TxtWidth.SelLength = Len(TxtWidth.Text)
    TxtWidthSel = True
End Sub

Private Sub TxtWidth_KeyUp(KeyCode As Integer, Shift As Integer)
    ActWidth = Val(TxtWidth.Text)
    On Error GoTo ErrorHandler
    If TxtWidthSel = True And CurType = 2 Then
        ActHeight = (1 / Val(TxtRatio.Text)) * ActWidth
        TxtHeight.Text = Format(ActHeight, "0.0#")
    End If
ErrorHandler:
End Sub

Private Sub TxtWidth_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitTextInput(KeyAscii, "0123456789.")
End Sub

Private Function LimitTextInput(Source, Allow As String) As String
    If Source <> 8 Then
        If InStr(Allow, Chr(Source)) = 0 Then
            LimitTextInput = 0
            Exit Function
        End If
    End If
    LimitTextInput = Source
End Function
