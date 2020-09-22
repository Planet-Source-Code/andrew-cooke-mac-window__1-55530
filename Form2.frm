VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "TrayIcon"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu Menu 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuHeader 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuRow0 
         Caption         =   "Row0"
      End
      Begin VB.Menu mnuRow1 
         Caption         =   "Row1"
      End
      Begin VB.Menu mnuRow2 
         Caption         =   "Row2"
      End
      Begin VB.Menu mnuRow3 
         Caption         =   "Row3"
      End
      Begin VB.Menu mnuRow4 
         Caption         =   "Row4"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRow5 
         Caption         =   "Row5"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Adding an icon to the system tray.
'by Black Bird http://kickme.to/pipiscrew
'Black Bird VB useful program at:
'http://users.otenet.gr/~pipiscr/FALister/index.htm
'Thanks: Peh Tee Howe for the SysTray icon

Private Sub RefreshTray()
   mnuRow0.Visible = False
   mnuRow1.Visible = False
   mnuRow2.Visible = False
   mnuRow3.Visible = False
   mnuRow4.Visible = False
   mnuRow5.Visible = False
End Sub


Private Sub Form_Load()
 
   'Set the individual values of the NOTIFYICONDATA data type.
   nid.cbSize = Len(nid)
   nid.hwnd = Form2.hwnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = Form2.Icon
   nid.szTip = "Taskbar Status Area Sample Program" & vbNullChar

   'Call the Shell_NotifyIcon function to add the icon to the taskbar
   'status area.
   Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lngMsg As Long

' get the WM Message passed via X since X is by default mes. in Twips,
' devide it by the number of twips / pixel


lngMsg = X / Screen.TwipsPerPixelX
            
Select Case lngMsg
            Case WM_RBUTTONUP ' right button
                        SetForegroundWindow Me.hwnd
                        RefreshTray
                        PopupMenu Menu, , , , mnuHeader 'Bold the Header -word SHOW-
            Case WM_LBUTTONDBLCLK
                        If Form2.WindowState = 1 Then
                            Form2.WindowState = 0
                            Exit Sub
                        End If
                        
                        If Form2.WindowState = 0 Then Form2.WindowState = 1
            End Select



End Sub

Private Sub Form_Terminate()
   'Delete the added icon from the taskbar status area when the
   'program ends.
   Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuExit_Click()
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub mnuHeader_Click()
Form1.Visible = True
Unload Me
End Sub
