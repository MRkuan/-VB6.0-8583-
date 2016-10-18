VERSION 5.00
Begin VB.Form Form_About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5055
   Icon            =   "关于框.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5055
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      Begin VB.TextBox Text_About 
         BackColor       =   &H8000000F&
         Height          =   1575
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   120
         Picture         =   "关于框.frx":0ECA
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.CommandButton confirm 
      Caption         =   "确认"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Copyright (C) 2014-2015"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   2895
   End
End
Attribute VB_Name = "Form_About"
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

    '    i = MsgBox("8583辅助解析工具V1.2                      " _
         '               & vbCrLf & "" _
         '               & vbCrLf & "作者：高建宽" _
         '               & vbCrLf & "版本：V1.2" _
         '               & vbCrLf & "QQ：1062220953" _
         '               & vbCrLf & "日期：2015年01月10日" _
         '               & vbCrLf & "TIPS：在按钮或框架悬停查看帮助" _
         '               , vbOKOnly, "关于")

    Text_About.Text = "8583辅助解析工具V1.3.1                     " _
                      & vbCrLf & "" _
                      & vbCrLf & "作者：高建宽" _
                      & vbCrLf & "版本：V1.3.1" _
                      & vbCrLf & "QQ  ：1062220953" _
                      & vbCrLf & "日期：2015年06月30日" _
                      & vbCrLf & "TIPS：在按钮或框架悬停查看帮助" _



End Sub

Private Sub confirm_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainForm.Enabled = True
End Sub


