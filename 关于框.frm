VERSION 5.00
Begin VB.Form Form_About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5055
   Icon            =   "���ڿ�.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5055
   StartUpPosition =   3  '����ȱʡ
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
         Picture         =   "���ڿ�.frx":0ECA
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.CommandButton confirm 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Copyright (C) 2014-2015"
      BeginProperty Font 
         Name            =   "����"
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




'//�ö���������
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Const a& = -1
Private Const b& = &H1
Private Const c& = &H2

'ͨ����������
Private Declare Function MessageBeep Lib "User32" (ByVal wType As Long) As Long

Private Const MB_ICONHAND = &H10&
Private Const MB_ICONQUESTION = &H20&
Private Const MB_ICONEXCLAMATION = &H30&
Private Const MB_ICONASTERISK = &H40&

Private Sub Form_Load()
    SetWindowPos Me.hWnd, a, 0, 0, 0, 0, b Or c    '//�ö�
    MessageBeep MB_ICONASTERISK                '//"��"��һ��

'//����λ��form1��������
    Me.Left = MainForm.Left + (MainForm.Width - Me.Width) / 2
    Me.Top = MainForm.Top + (MainForm.Height - Me.Height) / 2

    '//����ϵ���ԵҪ�Ż���

    '//�ƶ�����Ļ�����
    If Me.Left <= 0 Then Me.Left = 0

    '//�ƶ�����Ļ���ұ�
    If Me.Left + Me.Width - Screen.Width >= 0 Then Me.Left = Screen.Width - Me.Width

    '//�ƶ�����Ļ���ϱ�
    If Me.Top <= 0 Then Me.Top = 0

    '//�ƶ�����Ļ���±�
    If Me.Top + Me.Height - Screen.Height >= 0 Then Me.Top = Screen.Height - Me.Height

    '    i = MsgBox("8583������������V1.2                      " _
         '               & vbCrLf & "" _
         '               & vbCrLf & "���ߣ��߽���" _
         '               & vbCrLf & "�汾��V1.2" _
         '               & vbCrLf & "QQ��1062220953" _
         '               & vbCrLf & "���ڣ�2015��01��10��" _
         '               & vbCrLf & "TIPS���ڰ�ť������ͣ�鿴����" _
         '               , vbOKOnly, "����")

    Text_About.Text = "8583������������V1.3.1                     " _
                      & vbCrLf & "" _
                      & vbCrLf & "���ߣ��߽���" _
                      & vbCrLf & "�汾��V1.3.1" _
                      & vbCrLf & "QQ  ��1062220953" _
                      & vbCrLf & "���ڣ�2015��06��30��" _
                      & vbCrLf & "TIPS���ڰ�ť������ͣ�鿴����" _



End Sub

Private Sub confirm_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainForm.Enabled = True
End Sub


