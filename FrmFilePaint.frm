VERSION 5.00
Begin VB.Form FrmFilePaint 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3060
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFilePaint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4440
      TabIndex        =   8
      Top             =   2610
      Width           =   1095
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   4440
      TabIndex        =   7
      Top             =   2265
      Width           =   1095
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      ItemData        =   "FrmFilePaint.frx":000C
      Left            =   1080
      List            =   "FrmFilePaint.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2625
      Width           =   3255
   End
   Begin VB.TextBox TxtFile 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   2265
      Width           =   3255
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   2880
      Pattern         =   "*.bmp"
      TabIndex        =   2
      Top             =   105
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   465
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label LblType 
      BackStyle       =   0  'Transparent
      Caption         =   "File &Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2625
      Width           =   855
   End
   Begin VB.Label LblName 
      BackStyle       =   0  'Transparent
      Caption         =   "File &Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2265
      Width           =   855
   End
End
Attribute VB_Name = "FrmFilePaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public IsOpen As Boolean

Private Sub CmbType_Click()
    If CmbType.ListIndex = 0 Then File1.Pattern = "*.bmp" Else File1.Pattern = "*.*"
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    On Error GoTo ErrorHandler
    BlnSlash = CBool(Right(Dir1.Path, 1) = "\")
    AddSlash = String(Abs(BlnSlash - BlnSlash ^ BlnSlash), "\")
    If CmbType.ListIndex = 0 And LCase(Right(TxtFile.Text, 4)) <> ".bmp" Then AddExt = ".bmp" Else AddExt = ""
    ActFile = Dir1.Path & AddSlash & TxtFile.Text & AddExt
    If IsOpen = True Then
        FrmPaint.PicPaint.Picture = LoadPicture(ActFile)
    Else
        If FileExist(ActFile) Then
            MsgRes = MsgBox(Chr(34) & ActFile & Chr(34) & " already exists." & vbCrLf & "Do you want to replace it?", vbExclamation + vbYesNo, "Save As")
            If MsgRes = vbYes Then
                SavePicture FrmPaint.PicPaint.Image, ActFile
            ElseIf MsgRes = vbNo Then
                Exit Sub
            End If
        Else
            SavePicture FrmPaint.PicPaint.Image, ActFile
        End If
    End If
    FrmPaint.CurFile = TxtFile.Text
    FrmPaint.CurFullFile = ActFile
    FrmPaint.Caption = App.Title & " [" & FrmPaint.CurFile & "]"
    FrmPaint.ImgChanged = False
ErrorHandler:
    FrmPaint.HanEMove
    FrmPaint.HanSMove
    FrmPaint.HanSEMove
    Unload Me
End Sub

Private Sub Drive1_Change()
    On Error GoTo ErrorHandler
    oDrive = LCase(Left(Dir1.Path, 2))
    Dir1.Path = Drive1.Drive
    Exit Sub
ErrorHandler:
    Drive1.Drive = oDrive
End Sub

Private Sub Dir1_Change()
    On Error Resume Next
    File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
    TxtFile.Text = File1.FileName
End Sub

Private Sub File1_DblClick()
    CmdOK_Click
End Sub

Private Sub Form_Load()
    CmbType.ListIndex = 0
    If IsOpen = False Then TxtFile.Text = FrmPaint.CurFile & ".bmp"
End Sub

Private Function FileExist(Path) As Boolean
    On Error Resume Next
    Dim X
    X = FreeFile
    Open Path For Input As X
        If Err = 0 Then FileExist = True Else FileExist = False
    Close X
End Function
