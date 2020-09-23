VERSION 5.00
Begin VB.Form FrmEditColors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Colors"
   ClientHeight    =   1980
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2895
   Icon            =   "FrmEditColors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   9
      Top             =   1500
      Width           =   1215
   End
   Begin VB.TextBox TxtDec 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "255"
      Top             =   900
      Width           =   375
   End
   Begin VB.TextBox TxtHex 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2400
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "FF"
      Top             =   900
      Width           =   375
   End
   Begin VB.TextBox TxtRGB 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   525
      MaxLength       =   6
      TabIndex        =   5
      Text            =   "FFFFFF"
      Top             =   900
      Width           =   735
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   8
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Frame FrmRGB 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2655
      Begin VB.HScrollBar SldRGB 
         Height          =   135
         Left            =   120
         Max             =   255
         TabIndex        =   4
         Top             =   480
         Value           =   255
         Width           =   2415
      End
      Begin VB.OptionButton OptBlue 
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   180
         Width           =   735
      End
      Begin VB.OptionButton OptGreen 
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   180
         Width           =   735
      End
      Begin VB.OptionButton OptRed 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Line HR1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   2745
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line HR2 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   2760
      Y1              =   1335
      Y2              =   1320
   End
   Begin VB.Shape ShpColor 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      Height          =   285
      Left            =   120
      Top             =   900
      Width           =   285
   End
End
Attribute VB_Name = "FrmEditColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Red As Integer, Green As Integer, Blue As Integer

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    FrmPaint.PicFore.BackColor = RGB(Red, Green, Blue)
    FrmPaint.CurColor = RGB(Red, Green, Blue)
    Unload Me
End Sub

Private Sub Form_Load()
    Red = 255
    Green = 255
    Blue = 255
End Sub

Private Sub OptBlue_Click()
    SldRGB.Value = Blue
    TxtDec.Text = Blue
    TxtHex.Text = Hex(Blue)
    Red = Val("&H" & Mid(TxtRGB.Text, 1, 2))
    Green = Val("&H" & Mid(TxtRGB.Text, 3, 2))
End Sub

Private Sub OptGreen_Click()
    SldRGB.Value = Green
    TxtDec.Text = Green
    TxtHex.Text = Hex(Green)
    Red = Val("&H" & Mid(TxtRGB.Text, 1, 2))
    Blue = Val("&H" & Mid(TxtRGB.Text, 5, 2))
End Sub

Private Sub OptRed_Click()
    SldRGB.Value = Red
    TxtDec.Text = Red
    TxtHex.Text = Hex(Red)
    Green = Val("&H" & Mid(TxtRGB.Text, 3, 2))
    Blue = Val("&H" & Mid(TxtRGB.Text, 5, 2))
End Sub

Private Sub SldRGB_Change()
    SldRGB_Scroll
End Sub

Private Sub SldRGB_Scroll()
    TxtDec.Text = SldRGB.Value
    TxtHex.Text = Hex(SldRGB.Value)
    If OptRed.Value = True Then Red = TxtDec.Text Else: If OptGreen.Value = True Then Green = TxtDec.Text Else: If OptBlue.Value = True Then Blue = TxtDec.Text
    fillr = Left("00", 2 - Len(Hex(Red)))
    fillg = Left("00", 2 - Len(Hex(Green)))
    fillb = Left("00", 2 - Len(Hex(Blue)))
    TxtRGB.Text = fillr & Hex(Red) & fillg & Hex(Green) & fillb & Hex(Blue)
    ShpColor.BackColor = RGB(Red, Green, Blue)
End Sub

Private Sub TxtDec_Change()
    If TxtDec.Text = "" Then TxtDec.Text = 0
    If TxtDec.Text < 0 Then TxtDec.Text = 0
    If TxtDec.Text > 255 Then TxtDec.Text = 255
    TxtHex = Hex(TxtDec.Text)
    SldRGB.Value = TxtDec.Text
    If OptRed.Value = True Then Red = TxtDec.Text Else: If OptGreen.Value = True Then Green = TxtDec.Text Else: If OptBlue.Value = True Then Blue = TxtDec.Text
    fillr = Left("00", 2 - Len(Hex(Red)))
    fillg = Left("00", 2 - Len(Hex(Green)))
    fillb = Left("00", 2 - Len(Hex(Blue)))
    TxtRGB.Text = fillr & Hex(Red) & fillg & Hex(Green) & fillb & Hex(Blue)
    ShpColor.BackColor = RGB(Red, Green, Blue)
End Sub

Private Sub TxtDec_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitTextInput(KeyAscii, "0123456789")
End Sub

Private Sub TxtHex_Change()
    If Val("&H" & TxtHex.Text) < 0 Then TxtHex.Text = Hex(0)
    If Val("&H" & TxtHex.Text) > 255 Then TxtHex.Text = Hex(255)
    TxtDec = Val("&H" & TxtHex.Text)
    SldRGB.Value = Val("&H" & TxtHex.Text)
    If OptRed.Value = True Then Red = TxtDec.Text Else: If OptGreen.Value = True Then Green = TxtDec.Text Else: If OptBlue.Value = True Then Blue = TxtDec.Text
    fillr = Left("00", 2 - Len(Hex(Red)))
    fillg = Left("00", 2 - Len(Hex(Green)))
    fillb = Left("00", 2 - Len(Hex(Blue)))
    TxtRGB.Text = fillr & Hex(Red) & fillg & Hex(Green) & fillb & Hex(Blue)
    ShpColor.BackColor = RGB(Red, Green, Blue)
End Sub

Private Sub TxtHex_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitTextInput(KeyAscii, "0123456789ABCDEFabcdef")
End Sub

Private Sub TxtRGB_Change()
    Red = Val("&H" & Mid(TxtRGB.Text, 1, 2))
    Green = Val("&H" & Mid(TxtRGB.Text, 3, 2))
    Blue = Val("&H" & Mid(TxtRGB.Text, 5, 2))
    ShpColor.BackColor = RGB(Red, Green, Blue)
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

Private Sub TxtRGB_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitTextInput(KeyAscii, "0123456789ABCDEFabcdef")
End Sub
