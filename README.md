﻿<div align="center">

## clsSysTray\.cls


</div>

### Description

A simple class to very easily minimize an application to, or to just create an icon in, the system tray. UPDATED & FIXED 6/25/1999 - Added new properties and fixed various things. If you have gotten this code before, please get it again (yes, it actually works now!). Read the comments for instructions.
 
### More Info
 
This code must go in a CLASS module. The Class module name must be "clsSysTray" for the sample code provided to work. This class instantiates an instance of the PictureBox control on the fly. VB 5 does not support this, so you have to add a picturebox to the form to minimize and call it "pichook" for this to work, as well as set the VB_VERSION constant in the code to be 5.


<span>             |<span>
---                |---
**Submitted On**   |1999-06-28 11:55:10
**By**             |[Martin Richardson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/martin-richardson.md)
**Level**          |Intermediate
**User Rating**    |4.5 (76 globes from 17 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD241512201999\.zip](https://github.com/Planet-Source-Code/martin-richardson-clssystray-cls__1-2155/archive/master.zip)





### Source Code

```
'*:********************************************************************************
'*: Class. . . . . . . . . . : clsSysTray.cls
'*: Description. . . . . . . : When the application is minimized, it minimizes to
'*:              be an icon in the system tray.
'*: Author . . . . . . . . . : Martin Richardson
'*: Acknowledgements . . . . : Mark Hunter
'*: Copyright. . . . . . . . : This class is freeware
'*: VB Versions:
'*:
'*: 5.0 - Change the following constant definition to:
'*:    Private Const VB_VERSION = 5
#Const VB_VERSION = 6
'*:   - Add a picturebox control to your form, turn visible for it off, and
'*:    call it "pichook"
'*:
'*: 6.0 - Make sure the VB_VERSION constant is set to value of 6
'*:********************************************************************************
'*: Code to set up in the main form:
'Private WithEvents gSysTray As clsSysTray
'Private Sub Form_Load()
'  Set gSysTray = New clsSysTray
'  Set gSysTray.SourceWindow = Me
'End Sub
'Private Sub Form_Resize()
'  If Me.WindowState = vbMinimized Then
'    gSysTray.MinToSysTray
'  End If
'End Sub
'*: For VB5.0, add an invisible picture box to the form and call it "pichook"
'*: Properties
'*:
'*: Icon
'*:   Icon displayed in the taskbar. Use this property to set the icon, or return
'*:   it.
'*: ToolTip
'*:   Tooltip text displayed when the mouse is over the icon in the system tray. Use
'*:   this property to assign text to the tooltip, or to return the value of it.
'*: SourceWindow As Form
'*:   Reference to the form which will minimize to the system tray.
'*: DefaultDblClk As Boolean
'*:   Set to True to fire the DEFAULT (defined below) for the mouse double click event
'*:   which will show the application and remove the icon from the tray. (default)
'*:   Set to FALSE to override the below default event.
'*:
'*: Methods:
'*:
'*: MinToSysTray
'*:   Minimize the application, have it appear as an icon in the system tray.
'*:   The applicion disappears from the task bar and only appears in the
'*:   system tray.
'*: IconInSysTray
'*:   Create an icon for the application in the system tray, but leave the icon
'*:   visible and on the task bar.
'*: RemoveFromSysTray
'*:   Remove the icon from the system tray.
'*:
'*: These methods are available, but the same actions can be accomplished by
'*: setting the ICON and TOOLTIP properties.
'*:
'*: ChangeToolTip( sNewToolTip As String )
'*:   Set/change the tooltip displayed when the mouse is over the tray icon.
'*:   ex: gSysTray.ChangeToolTip "Processing..."
'*: ChangeIcon( pNewIcon As Picture )
'*:   Set/change the icon which appears in the system tray. The default icon
'*:   is the icon of the form.
'*:   ex: gSysTray.ChangeIcon ImageList1.ListImages("busyicon").picture
'*: Events:
'*: LButtonDblClk
'*:   Fires when double clicking the left mouse button over the tray icon. This event
'*:   has default code which will show the form and remove the icon from the
'*:   system tray when it fires. Set the property DefaultDblClk to False to
'*:   bypass this code.
'*: LButtonDown
'*:   Fires when the left mouse button goes down over the tray icon.
'*: LButtonUp
'*:   Fires when the left mouse button comes up over the tray icon.
'*: RButtonDblClk
'*:   Fires when double clicking the right mouse button over the tray icon.
'*: RButtonDown
'*:   Fires when the right mouse button goes down over the tray icon.
'*: RButtonUp
'*:   Fires when the right mouse button comes up over the tray icon.
'*:   Best place for calling a popup menu.
'*:
'*: Example of utilizing a popup menu with the RButtonUp event:
'*: 1. Create a menu on the form being minimized, or on it's own seperate form.
'*:   Let's say the form with the menu is called frmMenuForm.
'*: 2. Set the name of the root menu item to be mnuRightClickMenu
'*: 3. Assuming the name of the global SysTray object is still gSysTray, use this code
'*:   in the main form:
'*:
'Private Sub gSysTray_RButtonUP()
'  PopUpMenu frmMenuForm.mnuRightClickMenu
'End Sub
'*:
'*:********************************************************************************
Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uId As Long
  uFlags As Long
  ucallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private t As NOTIFYICONDATA
Private WithEvents pichook As PictureBox
Private mvarToolTip As String
Public Event LButtonDblClk()
Public Event LButtonDown()
Public Event LButtonUp()
Public Event RButtonDblClk()
Public Event RButtonDown()
Public Event RButtonUp()
'local variable(s) to hold property value(s)
Private mvarSourceWindow As Form 'local copy
Private mvarDefaultDblClk As Boolean 'local copy
Public Property Let ToolTip(ByVal vData As String)
  ChangeToolTip vData
End Property
Public Property Get ToolTip() As String
  ToolTip = mvarToolTip
End Property
Public Property Let Icon(ByVal vData As Variant)
  ChangeIcon vData
End Property
Public Property Get Icon() As Variant
  Icon = t.hIcon   'pichook.Picture
End Property
Public Property Let DefaultDblClk(ByVal vData As Boolean)
  mvarDefaultDblClk = vData
End Property
Public Property Get DefaultDblClk() As Boolean
  DefaultDblClk = mvarDefaultDblClk
End Property
Public Property Set SourceWindow(ByVal vData As Form)
  Set mvarSourceWindow = vData
  SetPicHook
End Property
Public Property Get SourceWindow() As Form
  Set SourceWindow = mvarSourceWindow
End Property
Public Sub ChangeToolTip(ByVal cNewTip As String)
  mvarToolTip = cNewTip
  t.szTip = cNewTip & Chr$(0)
  Shell_NotifyIcon NIM_MODIFY, t
  If mvarSourceWindow.WindowState = vbMinimized Then
    mvarSourceWindow.Caption = cNewTip
  End If
End Sub
Private Sub Class_Initialize()
  mvarDefaultDblClk = True
  t.cbSize = Len(t)
  t.uId = 1&
  t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  t.ucallbackMessage = WM_MOUSEMOVE
  t.hIcon = Me.Icon
  t.szTip = Chr$(0)    'Default to no tooltip
End Sub
Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static rec As Boolean, msg As Long, oldmsg As Long
  oldmsg = msg
  msg = X / Screen.TwipsPerPixelX
  If rec = False Then
    rec = True
    Select Case msg
      Case WM_LBUTTONDBLCLK:
        LButtonDblClk
      Case WM_LBUTTONDOWN:
        RaiseEvent LButtonDown
      Case WM_LBUTTONUP:
        RaiseEvent LButtonUp
      Case WM_RBUTTONDBLCLK:
        RaiseEvent RButtonDblClk
      Case WM_RBUTTONDOWN:
        RaiseEvent RButtonDown
      Case WM_RBUTTONUP:
        RaiseEvent RButtonUp
    End Select
    rec = False
  End If
End Sub
' Since VB doesn't really have inheretance (&^$%#&*!!) we have to fake it by using a
' variable to override default events...
Private Sub LButtonDblClk()
  If mvarDefaultDblClk Then
    mvarSourceWindow.WindowState = vbNormal
    mvarSourceWindow.Show
    App.TaskVisible = True
    RemoveFromSysTray
  End If
  RaiseEvent LButtonDblClk
End Sub
Public Sub RemoveFromSysTray()
  t.cbSize = Len(t)
  t.hwnd = pichook.hwnd
  t.uId = 1&
  Shell_NotifyIcon NIM_DELETE, t
End Sub
Public Sub IconInSysTray()
  Shell_NotifyIcon NIM_ADD, t
End Sub
Public Sub MinToSysTray()
  Me.IconInSysTray
  mvarSourceWindow.Hide
  App.TaskVisible = False
End Sub
Private Sub SetPicHook()
On Error GoTo AlreadyAdded
#If VB_VERSION = 6 Then
  Set pichook = mvarSourceWindow.Controls.Add("VB.PictureBox", "pichook")
#Else
  Set pichook = mvarSourceWindow.pichook
#End If
  pichook.Visible = False
  pichook.Picture = mvarSourceWindow.Icon
  t.hwnd = pichook.hwnd
  Exit Sub
AlreadyAdded:
  If Err.Number <> 727 Then ' pichook has already been added
    MsgBox "Run-time error '" & Err.Number & "':" & vbCrLf & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error"
    Stop
    Resume
  End If
End Sub
Public Sub ChangeIcon(toNewIcon)
  Set pichook.Picture = toNewIcon
  t.hIcon = pichook.Picture
  Shell_NotifyIcon NIM_MODIFY, t
End Sub
```

