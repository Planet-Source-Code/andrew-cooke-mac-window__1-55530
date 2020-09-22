VERSION 5.00
Begin VB.UserControl MacApp 
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
   ScaleHeight     =   1890
   ScaleWidth      =   7305
   Begin VB.PictureBox picTray 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6930
      Left            =   5490
      ScaleHeight     =   6930
      ScaleWidth      =   8445
      TabIndex        =   1
      Top             =   2070
      Width           =   8445
   End
   Begin VB.Image imgMax 
      Height          =   195
      Index           =   2
      Left            =   4050
      Picture         =   "MacApp.ctx":0000
      Top             =   1215
      Width           =   195
   End
   Begin VB.Image imgMax 
      Height          =   195
      Index           =   1
      Left            =   4050
      Picture         =   "MacApp.ctx":03F3
      Top             =   1035
      Width           =   195
   End
   Begin VB.Image imgMax 
      Height          =   195
      Index           =   0
      Left            =   4050
      Picture         =   "MacApp.ctx":079C
      ToolTipText     =   "maximize"
      Top             =   855
      Width           =   195
   End
   Begin VB.Image imgMin 
      Height          =   195
      Index           =   2
      Left            =   3510
      Picture         =   "MacApp.ctx":0B45
      Top             =   1170
      Width           =   195
   End
   Begin VB.Image imgMin 
      Height          =   195
      Index           =   1
      Left            =   3510
      Picture         =   "MacApp.ctx":0F15
      Top             =   990
      Width           =   195
   End
   Begin VB.Image imgMin 
      Height          =   195
      Index           =   0
      Left            =   3510
      Picture         =   "MacApp.ctx":12AA
      ToolTipText     =   "minimize"
      Top             =   855
      Width           =   195
   End
   Begin VB.Image imgTray 
      Height          =   195
      Index           =   2
      Left            =   3285
      Picture         =   "MacApp.ctx":163F
      Top             =   1215
      Width           =   195
   End
   Begin VB.Image imgTray 
      Height          =   195
      Index           =   1
      Left            =   3285
      Picture         =   "MacApp.ctx":1A30
      Top             =   1035
      Width           =   195
   End
   Begin VB.Image imgTray 
      Height          =   195
      Index           =   0
      Left            =   3285
      Picture         =   "MacApp.ctx":1DCB
      ToolTipText     =   "send to system tray"
      Top             =   855
      Width           =   195
   End
   Begin VB.Image imgRollup 
      Height          =   195
      Index           =   2
      Left            =   3060
      Picture         =   "MacApp.ctx":2166
      Top             =   1215
      Width           =   195
   End
   Begin VB.Image imgRollup 
      Height          =   195
      Index           =   1
      Left            =   3060
      Picture         =   "MacApp.ctx":254F
      Top             =   1035
      Width           =   195
   End
   Begin VB.Image imgRollup 
      Height          =   195
      Index           =   0
      Left            =   3060
      Picture         =   "MacApp.ctx":28E9
      ToolTipText     =   "roll up/down form"
      Top             =   855
      Width           =   195
   End
   Begin VB.Image imgClose 
      Height          =   195
      Index           =   2
      Left            =   3735
      Picture         =   "MacApp.ctx":2C83
      Top             =   1215
      Width           =   195
   End
   Begin VB.Image imgClose 
      Height          =   195
      Index           =   1
      Left            =   3735
      Picture         =   "MacApp.ctx":3090
      Top             =   1035
      Width           =   195
   End
   Begin VB.Image imgClose 
      Height          =   195
      Index           =   0
      Left            =   3735
      Picture         =   "MacApp.ctx":3431
      ToolTipText     =   "close"
      Top             =   855
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Xeons Mad MAC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1710
      TabIndex        =   0
      Top             =   90
      Width           =   1365
   End
End
Attribute VB_Name = "MacApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'[API]
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'[LOCAL VARS]
Private TrayIco As NOTIFYICONDATA
Private InTray As Boolean
Private m_frmHeight As Integer

'[OBJECT VARS]
Private WithEvents F As Form
Attribute F.VB_VarHelpID = -1
 

'[EVENTS]
Event CloseClick()
Event MinimizeClick()
Event MaximizeClick(bFormMaximized As Boolean)
Event TrayClick()
Event RollupClick(bRolledUp As Boolean)
Event TrayRightButton()
Event TrayLeftButton()
Event AddedToTray()
Event RemovedFromTray()

'[TYPES]
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'[ENUMS]
Public Enum enBtns
    CloseOnly
    CloseTray
    CloseMinimize
    CloseMinimizeTray
    CloseMaximize
    CloseMaximixeMinimizeTray
End Enum


'[CONTANTS]
Private Const WM_SYSCOMMAND As Long = &H112
Private Const NIM_ADD = &H0  'Add to Tray
Private Const NIM_MODIFY = &H1 'Modify Details
Private Const NIM_DELETE = &H2 'Remove From Tray
Private Const NIF_MESSAGE = &H1 'Message
Private Const NIF_ICON = &H2 'Icon
Private Const NIF_TIP = &H4 'TooTipText
Private Const WM_MOUSEMOVE = &H200 'On Mousemove
Private Const WM_LBUTTONDBLCLK = &H203 'Left Double Click
Private Const WM_RBUTTONDOWN = &H204 'Right Button Down
Private Const WM_RBUTTONUP = &H205 'Right Button Up
Private Const WM_RBUTTONDBLCLK = &H206 'Right Double Click

 
 

'Property Variables:
Dim m_TrayToolTip As String
Dim m_bFormFloat As Boolean
Dim m_bAllowFormMove As Boolean
Dim m_Buttons As Variant
Dim m_ToolTipText As String

'Default Property Values:
Const m_def_TrayToolTip = ""
Const m_def_bFormFloat = 0      'form not on top default
Const m_def_bAllowFormMove = 0  'can only move form by titlebar default
Const m_def_Buttons = 5         'show all buttons default




''  _     .    .          .  '    .    .     .  ' _     .  '    .      .
' / _| ___  _ _  _ __ __     ' ___  ___ __  __  '(_) _ _'  ' ___  _ __
'| |_ / _ \| '_\| '_ ` _ \   '/ _ \/ __|\ \/ /  '| |/ __|  '/ _ \| '_ `
'|  _| (_) | |  | | | | | |  ' (_)  (__  >  <   '| |\__ \  ' (_) | | | |
'|_|  \___/|_|  |_| |_| |_|  '\___/\___|/_/\_\  '|_||___/  '\___/|_| |_|

 
Private Sub F_Load()
  If bFormFloat = True Then
      SetWindowPos F.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
  End If
End Sub

Private Sub F_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If bAllowFormMove = True Then
     If Button = 1 Then
       Call mod_Move(F.hwnd)
     End If
   End If
End Sub

Private Sub F_Paint()
   Call PaintFormMac
End Sub

Private Sub F_Resize()
  Call UserControl_Resize
End Sub








'     .    .     .     '
'_____  _ _  __ _ _ _
'_   _|| '_\/ _` | |_| |
' | |  | |   (_| |\__, |
' |_|  |_|  \__,_||___/


Private Sub AddToTray()
On Error GoTo ERR_HANDLER:
'-------------------------------------------------
' ltray icon will be the forms icon unless
' specified otherwise
'-------------------------------------------------
'VARIABLES:
 
'CODE:
   InTray = True

   'initialize tray info
   With TrayIco
            .cbSize = Len(TrayIco)
            .hwnd = picTray.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = picTray.Picture
            .szTip = m_TrayToolTip & vbNullChar  'the tray tooltip
   End With
   'add this to tray
   Shell_NotifyIcon NIM_ADD, TrayIco
   RaiseEvent AddedToTray
'END CODE:
Exit Sub
ERR_HANDLER:
  Debug.Print Err.Description
End Sub

Sub RemoveFromTray()
On Error GoTo ERR_HANDLER:
'-------------------------------------------------
' remove the icon from tray..either because showing
' form or ending app
'-------------------------------------------------
'VARIABLES:

'CODE:
  'remove the tray icon
   Shell_NotifyIcon NIM_DELETE, TrayIco
   InTray = False
   RaiseEvent RemovedFromTray
'END CODE:
Exit Sub
ERR_HANDLER:
  Debug.Print Err.Description
End Sub

Sub ModifyTray(Optional sNewToolTip As String, Optional lNewIcon As Long)
On Error GoTo ERR_HANDLER:
'-------------------------------------------------
' change either the tooltip of the icon associated
' with tray icon
'-------------------------------------------------
'VARIABLES:

'CODE:
 With TrayIco
   If lNewIcon <> 0 Then .hIcon = lNewIcon
   If Len(Trim(sNewToolTip)) > 0 Then .szTip = sNewToolTip & vbNullChar
 End With
 'update tray icon with new values
 Shell_NotifyIcon NIM_MODIFY, TrayIco
'END CODE:
Exit Sub
ERR_HANDLER:
  Debug.Print Err.Description
End Sub
 
Private Sub PicTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_HANDLER:
'-------------------------------------------------
'this sub is called when the mouse moves over the tray
'icon because trays callback msg is wm_mousemove
'-=[thanks to LCSBSSRHXXX for the much shortened tray code]=-
'-------------------------------------------------
'VARIABLES:

'CODE:
    Select Case InTray
        Case True 'if were in tray and theres a r click show the menu
            If Button = 1 Then
                RaiseEvent TrayLeftButton
            ElseIf Button = 2 Then
                RaiseEvent TrayRightButton
            End If
        Case False
            Exit Sub
    End Select
'END CODE:
Exit Sub
ERR_HANDLER:
  Debug.Print Err.Description
End Sub
 
 
 
 
 
 
 

 ' _   .     .     .     .    .      .    .
'| |__ _   _ _____ _____  ___  _ __   _ _'
'|  _ \ | | |_   _|_   _|/ _ \| '_ ` / __|
'| |_)  |_| | | |   | |   (_) | | | |\__ \
'|_.__/\__,_| |_|   |_|  \___/|_| |_||___/


Private Sub imgClose_Click(Index As Integer)
    If Index = 0 Then
       RaiseEvent CloseClick
       Unload F
    End If
End Sub

Private Sub imgMax_Click(Index As Integer)
  If Index = 0 Then
     On Error Resume Next

     If F.WindowState <> vbMaximized Then
        RaiseEvent MaximizeClick(True)
        F.WindowState = 2 'maximized
     Else
        RaiseEvent MaximizeClick(False)
        F.WindowState = 0 'normal
     End If
     '
     F.Cls
    'repaint the lines
     Call PaintFormMac
  End If
End Sub

Private Sub imgMin_Click(Index As Integer)
   If Index = 0 Then
      On Error Resume Next
      RaiseEvent MinimizeClick
      F.WindowState = 1 'minimize
   End If
End Sub

Private Sub imgRollup_Click(Index As Integer)
    If Index = 0 Then
       
       If F.Height > 350 Then 'roll form up
          RaiseEvent RollupClick(True)
          m_frmHeight = F.Height
          F.Height = 320
       Else                   'roll form down
          RaiseEvent RollupClick(False)
          F.Height = m_frmHeight
       End If
    End If
End Sub
Private Sub imgTray_Click(Index As Integer)
   If Index = 0 Then
    Load Form2
    Form1.Visible = False
    F.Visible = False
   End If
End Sub
   
Private Sub imgMax_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgMax(0).Picture = imgMax(2).Picture
End Sub
Private Sub imgClose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgClose(0).Picture = imgClose(2).Picture
End Sub
Private Sub imgMin_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgMin(0).Picture = imgMin(2).Picture
End Sub
Private Sub imgRollup_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   imgRollup(0).Picture = imgRollup(2).Picture
End Sub
 
Private Sub imgTray_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   imgTray(0).Picture = imgTray(2).Picture
End Sub

Private Sub UserControl_DblClick()
'---------------------------------
'dbl clicking on titlebar maximizes/normalizes form
'just like a normal windows titlebar
 Call imgMax_Click(0)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'revert buttn pics to non mouseovers
 imgClose(0).Picture = imgClose(1).Picture
 imgRollup(0).Picture = imgRollup(1).Picture
 imgTray(0).Picture = imgTray(1).Picture
 imgMin(0).Picture = imgMin(1).Picture
 imgMax(0).Picture = imgMax(1).Picture
 
 If Ambient.UserMode = True Then
   If Button = 1 Then
      Call mod_Move(F.hwnd)
   End If
 End If
End Sub










' _ __      . _       .     . _       . __ _   '    .    .    .     .     .    .    _.
'| '_ \ __ _ (_) _ __  _____ (_) _ __  / _` |  ' _ _  ___  _    __ _ _____  ___  __| |
'| |_) / _` || || '_ ` _   _|| || '_ `  (_| |  '| '_\/ _ \| |  / _` |_   _|/ _ \/ _  |
'| .__/ (_| || || | | | | |  | || | | |\__, |  '| |    __/| |_  (_| | | |    __/ (_| |
'|_|   \__,_||_||_| |_| |_|  |_||_| |_||___/   '|_|  \___||___|\__,_| |_|  \___|\__,_|


Private Sub UserControl_Paint()
    'position the label that is the caption
    With Label1
       .Top = 80
       .Height = (Height - 130)
    End With
    'makes sure the titlebar is in the correct placement
    Call UserControl_Resize
    'paint the graphics effects (all the lines and colors)
    Call PaintMacTitleBar
End Sub

 
Private Sub UserControl_Resize()
 On Error Resume Next
 '----------------------------
 'position and size the mac titlebar
 '----------------------------
    With UserControl
        .Extender.Left = 0
        .Extender.Top = 0
        .Extender.Width = UserControl.Parent.Width
        .Extender.Height = 320
    End With
    
    Call PositionButtons
    'reset caption(label1)
    Call CalCaptPlacement
End Sub

Private Sub PositionButtons(Optional bCloseMinSelOnly As Boolean = False)
'----------------------------
'position buttons on titlebar
'----------------------------
    With imgMin(0)      'min button pos
       If Buttons = CloseMinimize Or Buttons = CloseMinimizeTray Then
          .Left = (UserControl.Width - 550)
       Else
          .Left = (UserControl.Width - 800)
       End If
       .Top = 75
    End With
    
    'this boolean value is set to true if
    'CloseMin was the buttons selected
    'all we want to do here is slide the
    'min button to the right in the position
    'the max button would normally be IF it
    'were visible
    If bCloseMinSelOnly = True Then
       Exit Sub
    End If
    
    With imgClose(0)    'close button pos
       .Top = 75
       .Left = (UserControl.Width - 300)
    End With
    With imgRollup(0)   'rollup button pos
       .Left = 100
       .Top = 75
    End With
    With imgTray(0)     'tray button pos
       .Left = 350
       .Top = 75
    End With
    With imgMax(0)      'max button pos
       .Left = (UserControl.Width - 550)
       .Top = 75
    End With
End Sub


Private Sub PaintFormMac()
On Error GoTo ERR_HANDLER:
'-------------------------------------------------
'this paints the borders that create 3d illusion
'creating a form that looks like a mac form
'-------------------------------------------------
'CODE:
  'F represents the form this control is on
  F.DrawWidth = 5
  F.BackColor = RGB(222, 223, 222)
 
  F.Line (0, 320)-(F.Width, 320), RGB(250, 250, 250)             'top hilite
  F.Line (0, 0)-(0, F.Height), RGB(250, 250, 250)                'left hilite
  F.Line (50, F.Height)-(F.Width, F.Height), RGB(130, 130, 140)  'bottom shadow
  F.Line (F.Width, 350)-(F.Width, F.Height), RGB(130, 130, 140)  'right shadow
'END CODE:
Exit Sub
ERR_HANDLER:
  Debug.Print Err.Description
End Sub

Private Sub PaintMacTitleBar()
On Error GoTo ERR_HANDLER:
'-------------------------------------------------
'
'-------------------------------------------------
'VARIABLES:
  Dim i%, lPoint%, rPoint%
'CODE:
  With UserControl
     .Cls
     .DrawWidth = 4
     .BackColor = RGB(206, 207, 206)
     Label1.BackColor = .BackColor
     'outside border
     UserControl.Line (0, 0)-(.Width, 0), RGB(250, 250, 250)              'top hilite
     UserControl.Line (0, 0)-(0, .Height), RGB(250, 250, 250)             'left hilite
     UserControl.Line (30, .Height)-(.Width, .Height), RGB(130, 130, 140) 'bottom shadow
     UserControl.Line (.Width, 50)-(.Width, .Height), RGB(130, 130, 140)  'right shadow
     
     'start and end points of the lines in title
     'bar (gives title bar the classic mac look)
     'are determined by which buttons are showing
     If Buttons = CloseMaximixeMinimizeTray Then
        lPoint = 600
        rPoint = (.Width - 900)
     ElseIf Buttons = CloseMaximize Or Buttons = CloseMinimize Then
        lPoint = 400
        rPoint = (.Width - 700)
     ElseIf Buttons = CloseMinimizeTray Then
        lPoint = 600
        rPoint = (.Width - 700)
     ElseIf Buttons = CloseOnly Then
        lPoint = 400
        rPoint = (.Width - 400)
     ElseIf Buttons = CloseTray Then
        lPoint = 600
        rPoint = (.Width - 400)
     End If
     
     'inside horizonatal mac lines
     .DrawWidth = 5
     UserControl.Line ((lPoint + 40), 140)-(rPoint - 30, 200), RGB(115, 117, 119), B
     
     .DrawWidth = 1
     
      For i = 90 To 210 Step 30
         UserControl.Line (lPoint, i)-(rPoint, i), RGB(250, 250, 250)
      Next i
      '
      CalCaptPlacement
  End With
'END CODE:
Exit Sub
ERR_HANDLER:
  Debug.Print Err.Description
End Sub

Private Sub CalCaptPlacement()
 On Error Resume Next
'VARIABLES
  Dim wid%
'CODE
  'create our own autosize of sorts
  'that adds a little extra at the edges
  wid = Len(MacCaption)
  Label1.Width = (wid * 110)
  Label1.Left = (UserControl.Width / 2) - (Label1.Width / 2)
'END CODE
End Sub
 








'    .     . _   .    .   _ '  _      .      .    .     . _     .      .    .
' _ _'_   _ | |__  _ _'  / / / _|_   _  _ __   ___ _____ (_) ___  _ __   _ _'
'/ __| | | ||  _ \/ __| / / | |_  | | || '_ ` / __|_   _|| |/ _ \| '_ ` / __|
'\__ \ |_| || |_) \__ \/_/  |  _| |_| || | | | (__  | |  | | (_) | | | |\__ \
'|___/\__,_||_.__/|___/    '|_|  \__,_||_| |_|\___| |_|  |_|\___/|_| |_||___/

'----------------------------------------------------------------------
' COMMENTS: | holds the code related to moving controls
'----------------------------------------------------------------------
Private Sub mod_Move(lhWnd&)
'CODE:
  ReleaseCapture
  SendMessage lhWnd, WM_SYSCOMMAND, &HF012&, 0&
'END CODE:
End Sub
'----------------------------------------------------------------------
'   INPUTS: | Handle to the window, boolean value OnTop
' COMMENTS: | this will place the specified window on top or remove it from top
'----------------------------------------------------------------------
Private Sub WindowOnTop(lhWnd&, Optional bOnTop As Boolean = True)
'CODE:
      Call SetWindowPos( _
                     lhWnd, CLng(bOnTop), 0, 0, 0, 0, _
                     &H1 Or &H2 _
                      )
'END CODE:
End Sub

 
 
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'--------------------------------
'move form when mouse down of MacTitlebar
'--------------------------------
 If Ambient.UserMode = True Then
   If Button = 1 Then
      Call mod_Move(F.hwnd)
   End If
 End If
End Sub
 
 
 
 
 
 
 
 
 
 
 
 
 
 ' _ __     .    . _ __     .    .     . _     .    .
'| '_ \ _ _  ___ | '_ \ ___  _ _ _____ (_) ___  _ _'
'| |_) | '_\/ _ \| |_) / _ \| '_\_   _|| |/ _ \/ __|
'| .__/| |   (_) | .__/  __/| |   | |  | |  __/\__ \
'|_|   |_|  \___/|_|   \___||_|   |_|  |_|\___||___/


'MOVE THE FORM WITH MOUSEMOVE ON FORM
Public Property Get bAllowFormMove() As Boolean
        bAllowFormMove = m_bAllowFormMove
End Property

Public Property Let bAllowFormMove(ByVal New_bAllowFormMove As Boolean)
        m_bAllowFormMove = New_bAllowFormMove
        PropertyChanged "bAllowFormMove"
End Property

' FLOAT (PARENT FORM ON TOP)
Public Property Get bFormFloat() As Boolean
        bFormFloat = m_bFormFloat
End Property

Public Property Let bFormFloat(ByVal New_bFormFloat As Boolean)
 Dim lPos&
        m_bFormFloat = New_bFormFloat
        PropertyChanged "bFormFloat"
        
        If Ambient.UserMode = False Then
             Exit Property
        End If
        
        If m_bFormFloat = True Then
           lPos = -1
        Else
           lPos = 1
        End If
        
        SetWindowPos F.hwnd, lPos, 0, 0, 0, 0, &H1 Or &H2
End Property

'TITLEBAR BUTTONS
Public Property Get Buttons() As enBtns
        Buttons = m_Buttons
End Property

Public Property Let Buttons(ByVal New_Buttons As enBtns)
        m_Buttons = New_Buttons
        PropertyChanged "Buttons"
        Call PickButtons
End Property

Private Sub PickButtons()
        If Buttons = CloseOnly Then
           imgMax(0).Visible = False
           imgMin(0).Visible = False
           imgTray(0).Visible = False
           imgClose(0).Visible = True
           
        ElseIf Buttons = CloseMaximize Then
           imgMax(0).Visible = True
           imgMin(0).Visible = False
           imgTray(0).Visible = False
           imgClose(0).Visible = True
        
        ElseIf Buttons = CloseTray Then
           imgMax(0).Visible = False
           imgMin(0).Visible = False
           imgTray(0).Visible = True
           imgClose(0).Visible = True
           
        ElseIf Buttons = CloseMinimize Then
           imgMax(0).Visible = False
           imgMin(0).Visible = True
           imgTray(0).Visible = False
           imgClose(0).Visible = True
           Call PositionButtons(True)
           
        ElseIf Buttons = CloseMinimizeTray Then
           imgMax(0).Visible = False
           imgMin(0).Visible = True
           imgTray(0).Visible = True
           imgClose(0).Visible = True
           Call PositionButtons
           
        ElseIf Buttons = CloseMaximixeMinimizeTray Then
           imgMax(0).Visible = True
           imgMin(0).Visible = True
           imgTray(0).Visible = True
           imgClose(0).Visible = True
           Call PositionButtons
           
        End If
        
        'always show the rollup button
        'imgRollup(0).Visible = True
        Call PaintMacTitleBar
End Sub
 
 'TOOL TIP TEXT
Public Property Get ToolTipText() As String
        ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
        m_ToolTipText = New_ToolTipText
        PropertyChanged "ToolTipText"
End Property
 
 'CAPTION
Public Property Get MacCaption() As String
        MacCaption = Label1.Caption
End Property

Public Property Let MacCaption(ByVal New_MacCaption As String)
        Label1.Caption() = New_MacCaption
        PropertyChanged "MacCaption"
        Call CalCaptPlacement
End Property

  'CAPTION COLOR
Public Property Get CaptionColor() As OLE_COLOR
        CaptionColor = Label1.ForeColor
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
        Label1.ForeColor() = New_CaptionColor
        PropertyChanged "CaptionColor"
End Property

 'TRAY ICON
Public Property Get TrayIcon() As Picture
        Set TrayIcon = picTray.Picture
End Property

Public Property Set TrayIcon(ByVal New_TrayIcon As Picture)
        Set picTray.Picture = New_TrayIcon
        PropertyChanged "TrayIcon"
End Property

'[TRAY TOOL TIP]
Public Property Get TrayToolTip() As String
        TrayToolTip = m_TrayToolTip
End Property

Public Property Let TrayToolTip(ByVal New_TrayToolTip As String)
        m_TrayToolTip = New_TrayToolTip
        PropertyChanged "TrayToolTip"
End Property

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error Resume Next
   Call PropBag.WriteProperty("MacCaption", Label1.Caption, "MAC caption")
   Call PropBag.WriteProperty("CaptionColor", Label1.ForeColor, &HFF0000)
   Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, "")
   Call PropBag.WriteProperty("Buttons", m_Buttons, m_def_Buttons)
   Call PropBag.WriteProperty("bAllowFormMove", m_bAllowFormMove, m_def_bAllowFormMove)
   Call PropBag.WriteProperty("bFormFloat", m_bFormFloat, m_def_bFormFloat)
   Call PropBag.WriteProperty("TrayIcon", picTray.Picture, Nothing)
   Call PropBag.WriteProperty("TrayToolTip", m_TrayToolTip, m_def_TrayToolTip)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 On Error Resume Next
  Label1.Caption = PropBag.ReadProperty("MacCaption", "MAC caption")
  Label1.ForeColor = PropBag.ReadProperty("CaptionColor", &H0)
  m_ToolTipText = PropBag.ReadProperty("ToolTipText", "")
  m_Buttons = PropBag.ReadProperty("Buttons", m_def_Buttons)
  m_bAllowFormMove = PropBag.ReadProperty("bAllowFormMove", m_def_bAllowFormMove)
  m_bFormFloat = PropBag.ReadProperty("bFormFloat", m_def_bFormFloat)
  Set picTray.Picture = PropBag.ReadProperty("TrayIcon", Nothing)
  m_TrayToolTip = PropBag.ReadProperty("TrayToolTip", m_def_TrayToolTip)
  
  If UserControl.Ambient.UserMode = True Then 'this means the app is running
     Set F = UserControl.Parent
     Label1.FontBold = True
     Call PickButtons
  Else 'this means the app is in design mode (your programming)
     Set F = Nothing
  End If
 
End Sub
 
 

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
        m_ToolTipText = "MAC titlebars can even have tooltips"
        m_Buttons = m_def_Buttons
        m_bAllowFormMove = m_def_bAllowFormMove
        m_bFormFloat = m_def_bFormFloat
        Label1.FontBold = True
        m_TrayToolTip = m_def_TrayToolTip
End Sub
 
 

 

