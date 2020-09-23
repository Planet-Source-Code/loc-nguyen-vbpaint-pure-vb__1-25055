VERSION 5.00
Begin VB.Form FrmAboutPaint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About VBPaint"
   ClientHeight    =   2115
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5925
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAboutPaint.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   4680
      TabIndex        =   3
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Line HR1 
      BorderColor     =   &H80000010&
      X1              =   1560
      X2              =   5745
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Line HR2 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   1560
      X2              =   5760
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label LblLicense 
      BackStyle       =   0  'Transparent
      Caption         =   "This  product is licensed as FREEWARE."
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label LblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2001 Loc Nguyen"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label LblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "VBPaint"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image ImgIcon 
      Height          =   495
      Left            =   480
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "FrmAboutPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ImgIcon.Picture = FrmPaint.Icon
End Sub
