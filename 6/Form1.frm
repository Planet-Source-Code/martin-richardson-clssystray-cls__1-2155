VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3825
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "newicon"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0452
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

Private Sub Form_Load()
    Set gSysTray = New clsSysTray
    Set gSysTray.SourceWindow = Me
    gSysTray.ChangeIcon ImageList1.ListImages("newicon").Picture
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    gSysTray.RemoveFromSysTray
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
    End
End Sub
