VERSION 5.00
Begin VB.Form PromptForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Title"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3540
   Icon            =   "提示框.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3540
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton confirm 
      Caption         =   "确认"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label msg_display 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "hehe"
      Height          =   900
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3225
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "PromptForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




'//置顶函数声明
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Const a& = -1
Private Const b& = &H1
Private Const c& = &H2

'通用声音声明
Private Declare Function MessageBeep Lib "User32" (ByVal wType As Long) As Long

Private Const MB_ICONHAND = &H10&
Private Const MB_ICONQUESTION = &H20&
Private Const MB_ICONEXCLAMATION = &H30&
Private Const MB_ICONASTERISK = &H40&

Private Sub Form_Load()
    SetWindowPos Me.hWnd, a, 0, 0, 0, 0, b Or c    '//置顶
    MessageBeep MB_ICONASTERISK                '//"叮"的一声

'//出现位置form1中正中央
    Me.Left = MainForm.Left + (MainForm.Width - Me.Width) / 2
    Me.Top = MainForm.Top + (MainForm.Height - Me.Height) / 2

    '//如果拖到边缘要优化下

    '//移动到屏幕最左边
    If Me.Left <= 0 Then Me.Left = 0

    '//移动到屏幕最右边
    If Me.Left + Me.Width - Screen.Width >= 0 Then Me.Left = Screen.Width - Me.Width

    '//移动到屏幕最上边
    If Me.Top <= 0 Then Me.Top = 0

    '//移动到屏幕最下边
    If Me.Top + Me.Height - Screen.Height >= 0 Then Me.Top = Screen.Height - Me.Height


End Sub

Private Sub confirm_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainForm.Enabled = True
End Sub
