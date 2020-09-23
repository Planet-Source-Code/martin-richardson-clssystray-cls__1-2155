VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.PictureBox pichook 
      Height          =   855
      Left            =   2760
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0000
            Key             =   "newicon"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":031A
            Key             =   "secondicon"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopupMenu 
      Caption         =   "&Popup Menu"
      Begin VB.Menu mnuChangeIcon 
         Caption         =   "&Change Icon"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1

Private Sub Command1_Click()
    gSysTray.ToolTip = Text1
End Sub

Private Sub Form_Load()
    Set gSysTray = New clsSysTray
    Set gSysTray.SourceWindow = Me
    gSysTray.ChangeIcon ImageList1.ListImages("newicon").Picture
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        gSysTray.MinToSysTray
    End If
End Sub


Private Sub gSysTray_RButtonUP()
    PopupMenu Me.mnuPopupMenu
End Sub

Private Sub mnuChangeIcon_Click()
    If gSysTray.Icon = ImageList1.ListImages("newicon").Picture Then
        gSysTray.Icon = ImageList1.ListImages("secondicon").Picture
    Else
        gSysTray.Icon = ImageList1.ListImages("newicon").Picture
    End If
End Sub

Private Sub mnuQuit_Click()
    gSysTray.RemoveFromSysTray
    End
End Sub
