VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   2565
   ClientTop       =   2745
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   6570
   Begin MacAppOcxExample.MacApp MacApp1 
      Height          =   320
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6570
      _ExtentX        =   11086
      _ExtentY        =   556
      MacCaption      =   "Mac Caption"
      CaptionColor    =   0
      Object.ToolTipText     =   "MAC titlebars can even have tooltips"
      TrayIcon        =   "Form1.frx":0000
      TrayToolTip     =   "tray tool tip"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Unload Form2
  With MacApp1
     .MacCaption = "Example1" 'This will change the text on the Window
     .ToolTipText = "Example2" 'This will allow you to change the text that appears when you hover over the titlebar
  End With
End Sub

'=[THESE ARE THE THE EVENTS FOR THE MacApp control]

Private Sub MacApp1_AddedToTray()
End Sub

Private Sub MacApp1_TrayLeftButton()
   MacApp1.RemoveFromTray
   Visible = True
End Sub
