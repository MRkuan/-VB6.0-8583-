VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "8583������������V1.3.1"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13590
   Icon            =   "8583��������С����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   13590
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton analyse_level_3 
      Caption         =   "ȫ�����"
      Enabled         =   0   'False
      Height          =   495
      Left            =   12240
      TabIndex        =   95
      ToolTipText     =   "ר��ģʽ��BT��"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame10 
      Caption         =   "��ˮ/���κ�"
      Height          =   1215
      Left            =   12120
      TabIndex        =   93
      ToolTipText     =   "������ˮ/���κ�"
      Top             =   7320
      Width           =   1335
      Begin VB.Label batch_Number_60 
         Caption         =   "N/A"
         Height          =   200
         Left            =   120
         TabIndex        =   99
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "���κţ�"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   720
         Width           =   735
      End
      Begin VB.Label trace_no_11 
         Caption         =   "N/A"
         Height          =   195
         Left            =   120
         TabIndex        =   97
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "��ˮ�ţ�"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton analyse_level_1 
      Caption         =   "��ͨ����"
      Height          =   495
      Left            =   12240
      TabIndex        =   1
      ToolTipText     =   "��ͨ����"
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame9 
      Caption         =   "��Ӧ����ʾ"
      Height          =   2175
      Left            =   12120
      TabIndex        =   90
      ToolTipText     =   "POS 2010�淶����Ӧ����ʾ"
      Top             =   5040
      Width           =   1335
      Begin VB.Label Response_code_view 
         Caption         =   "N/A"
         Height          =   1335
         Left            =   120
         TabIndex        =   92
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Response_code 
         Alignment       =   2  'Center
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   91
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton bit_map_set 
      Caption         =   "λͼ"
      Height          =   360
      Left            =   2800
      TabIndex        =   7
      ToolTipText     =   "ֻӰ��λͼ��λͼ��ʾ"
      Top             =   6765
      Width           =   615
   End
   Begin VB.CommandButton bit_map_clear 
      Caption         =   "���"
      Height          =   360
      Left            =   3430
      TabIndex        =   8
      ToolTipText     =   "ֻ���λͼ��λͼ��ʾ"
      Top             =   6765
      Width           =   615
   End
   Begin VB.Frame Frame7 
      Caption         =   "λͼ��ʾ"
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      ToolTipText     =   "���λͼ��ʾ���ֻ��Ӱ��λͼ"
      Top             =   7320
      Width           =   11895
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "64"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   63
         Left            =   11280
         TabIndex        =   89
         ToolTipText     =   "��64"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "63"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   62
         Left            =   10920
         TabIndex        =   88
         ToolTipText     =   "��63"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "62"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   61
         Left            =   10560
         TabIndex        =   87
         ToolTipText     =   "��62"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "61"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   60
         Left            =   10200
         TabIndex        =   86
         ToolTipText     =   "��61"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "60"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   59
         Left            =   9840
         TabIndex        =   85
         ToolTipText     =   "��60"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "59"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   58
         Left            =   9480
         TabIndex        =   84
         ToolTipText     =   "��59"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "58"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   57
         Left            =   9120
         TabIndex        =   83
         ToolTipText     =   "��58"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "57"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   56
         Left            =   8760
         TabIndex        =   82
         ToolTipText     =   "��57"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "56"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   55
         Left            =   8400
         TabIndex        =   81
         ToolTipText     =   "��56"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "55"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   54
         Left            =   8040
         TabIndex        =   80
         ToolTipText     =   "��55"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "54"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   53
         Left            =   7680
         TabIndex        =   79
         ToolTipText     =   "��54"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "53"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   52
         Left            =   7320
         TabIndex        =   78
         ToolTipText     =   "53"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "52"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   51
         Left            =   6960
         TabIndex        =   77
         ToolTipText     =   "��52"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "51"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   50
         Left            =   6600
         TabIndex        =   76
         ToolTipText     =   "��51"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   49
         Left            =   6240
         TabIndex        =   75
         ToolTipText     =   "��50"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "49"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   48
         Left            =   5880
         TabIndex        =   74
         ToolTipText     =   "��49"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "48"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   47
         Left            =   5520
         TabIndex        =   73
         ToolTipText     =   "��48"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "47"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   46
         Left            =   5160
         TabIndex        =   72
         ToolTipText     =   "��47"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "46"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   45
         Left            =   4800
         TabIndex        =   71
         ToolTipText     =   "��46"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "45"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   44
         Left            =   4440
         TabIndex        =   70
         ToolTipText     =   "��45"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "44"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   43
         Left            =   4080
         TabIndex        =   69
         ToolTipText     =   "��44"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "43"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   42
         Left            =   3720
         TabIndex        =   68
         ToolTipText     =   "��43"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "42"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   41
         Left            =   3360
         TabIndex        =   67
         ToolTipText     =   "��42"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "41"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   40
         Left            =   3000
         TabIndex        =   66
         ToolTipText     =   "��41"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   39
         Left            =   2640
         TabIndex        =   65
         ToolTipText     =   "��40"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "39"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   38
         Left            =   2280
         TabIndex        =   64
         ToolTipText     =   "��39"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   37
         Left            =   1920
         TabIndex        =   63
         ToolTipText     =   "��38"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "37"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   36
         Left            =   1560
         TabIndex        =   62
         ToolTipText     =   "��37"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   35
         Left            =   1200
         TabIndex        =   61
         ToolTipText     =   "��36"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "35"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   34
         Left            =   840
         TabIndex        =   60
         ToolTipText     =   "��35"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   33
         Left            =   480
         TabIndex        =   59
         ToolTipText     =   "��34"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "33"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   32
         Left            =   120
         TabIndex        =   58
         ToolTipText     =   "��33"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   31
         Left            =   11280
         TabIndex        =   57
         ToolTipText     =   "��32"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   10920
         TabIndex        =   56
         ToolTipText     =   "��31"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   10560
         TabIndex        =   55
         ToolTipText     =   "��30"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   10200
         TabIndex        =   54
         ToolTipText     =   "��29"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   9840
         TabIndex        =   53
         ToolTipText     =   "��28"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   9480
         TabIndex        =   52
         ToolTipText     =   "��27"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   9120
         TabIndex        =   51
         ToolTipText     =   "��26"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   8760
         TabIndex        =   50
         ToolTipText     =   "��25"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   8400
         TabIndex        =   49
         ToolTipText     =   "��24"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   8040
         TabIndex        =   48
         ToolTipText     =   "��23"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   7680
         TabIndex        =   47
         ToolTipText     =   "��22"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   7320
         TabIndex        =   46
         ToolTipText     =   "��21"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   6960
         TabIndex        =   45
         ToolTipText     =   "��20"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   6600
         TabIndex        =   44
         ToolTipText     =   "��19"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   6240
         TabIndex        =   43
         ToolTipText     =   "��18"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   5880
         TabIndex        =   42
         ToolTipText     =   "��17"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   5520
         TabIndex        =   41
         ToolTipText     =   "��16"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   5160
         TabIndex        =   40
         ToolTipText     =   "��15"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   4800
         TabIndex        =   39
         ToolTipText     =   "��14"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   4440
         TabIndex        =   38
         ToolTipText     =   "��13"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   4080
         TabIndex        =   37
         ToolTipText     =   "��12"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   3720
         TabIndex        =   36
         ToolTipText     =   "��11"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   3360
         TabIndex        =   35
         ToolTipText     =   "��10"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   3000
         TabIndex        =   34
         ToolTipText     =   "��9"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2640
         TabIndex        =   33
         ToolTipText     =   "��8"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2280
         TabIndex        =   32
         ToolTipText     =   "��7"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   31
         ToolTipText     =   "��6"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   30
         ToolTipText     =   "��5"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   29
         ToolTipText     =   "��4"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   28
         ToolTipText     =   "��3"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   27
         ToolTipText     =   "��2"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "��1"
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame_bit_map 
      Caption         =   "λͼ[������ܸ�������]"
      Height          =   615
      Left            =   120
      TabIndex        =   23
      ToolTipText     =   "����λͼ"
      Top             =   6600
      Width           =   3975
      Begin VB.TextBox bit_map 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CheckBox Totallen_check 
      Caption         =   "Totallen"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4800
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Frame Frame6 
      Caption         =   "��������[��ѡ��������ɳ������]"
      Height          =   735
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "ȥ��һЩ���ܲ���Ҫ��������Ϣ(��ѡ��������ɳ������)"
      Top             =   4440
      Width           =   3975
      Begin VB.CheckBox MessageHeader_check 
         Caption         =   "MessageHeader"
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox TPDU_check 
         Caption         =   "TPDU"
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "�����ж�[����������жϣ������ο�]"
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "����������жϣ������ο�"
      Top             =   5280
      Width           =   3975
      Begin VB.Label judge_mode 
         Caption         =   "N/A"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "ģʽ��"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label trans_type 
         Caption         =   "N/A"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   960
         TabIndex        =   16
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "�������ͣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton help 
      Caption         =   "˵��"
      Height          =   495
      Left            =   12240
      TabIndex        =   6
      ToolTipText     =   "˵��"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton about 
      Caption         =   "����"
      Height          =   495
      Left            =   12240
      TabIndex        =   5
      ToolTipText     =   "����"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton clear 
      Caption         =   "����"
      Height          =   495
      Left            =   12240
      TabIndex        =   4
      ToolTipText     =   "�����Ļ����"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton analyse_level_2 
      Caption         =   "ר�ҽ���"
      Height          =   495
      Left            =   12240
      TabIndex        =   2
      ToolTipText     =   "ר�ҽ���"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox analyse_after_data 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   4320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   360
      Width           =   7575
   End
   Begin VB.TextBox analyse_before_data 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Frame Frame_analyse_before_data 
      Caption         =   "����ǰ����[������ܸ�������]"
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "���ǽ���ǰ����"
      Top             =   120
      Width           =   3975
   End
   Begin VB.Frame Frame_analyse_after_data 
      Caption         =   "����������[������ܸ�������]"
      Height          =   7095
      Left            =   4200
      TabIndex        =   10
      ToolTipText     =   "���ǽ���������"
      Top             =   120
      Width           =   7815
   End
   Begin VB.Frame Frame3 
      Caption         =   "����ѡ��"
      Height          =   2655
      Left            =   12120
      TabIndex        =   11
      ToolTipText     =   "��������ѡ��"
      Top             =   2280
      Width           =   1335
      Begin VB.CommandButton END 
         Caption         =   "�˳�"
         Height          =   495
         Left            =   120
         TabIndex        =   94
         ToolTipText     =   "�˳�"
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ģʽѡ��"
      Height          =   2055
      Left            =   12120
      TabIndex        =   12
      ToolTipText     =   "����ģʽѡ��"
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/****************************************************************************
'  Copyright (c)   ����
'  File Name:      8583����С����
'  Author:         �߽���
'  Version:        V1.3
'  Date:           2014��12��26��
'  Description:    ��8583�ַ�����
'  Function List():
'
'History: V1.0���ӻ���������λͼ��ʾ�����ã���ͣ����밴ť���Բ鿴�������
'         V1.1��������ͨ������1.0�����������Ϊר�ҽ�����������Ӧ����ʾ
'         V1.3����55�����
'
'
'Author:
'Modification:
'
'**************************************************************************/





Option Explicit

Dim bit_map_view_Click_flag(1 To 64) As Boolean    '// λͼ���±�־
Dim change_begin As Integer    '//ת����ʼ��־

Dim help_flag As Integer    '//���������־
Dim help_count As Long     '//���尴�°�������

Dim analyse_mode As Integer    '//�������ģʽ 1 ��ͨ���� 2 ר�ҽ��� 3 ȫ�����


Private Type PtrStruct
    ptr As String
    Ptrlen As Integer
End Type


'//����8583POS�еĽṹ��
Private Type POS_Sturct_TYPE
    messagetype As String
    bitmap As String
    procode_3 As String
    consume_amount_4 As String
    trace_no_11 As String
    trade_time_12 As String
    trade_date_13 As String
    exp_date_14 As String
    settlement_date_15 As String
    entry_mode_22 As String
    card_serial_number_23 As String
    service_conditon_25 As String
    service_conditon_pin_26 As String
    reference_number_37 As String
    authorization_code_38 As String
    Response_code_39 As String
    terminal_no_41 As String
    merchant_no_42 As String
    merchant_name_43 As PtrStruct
    currency_code_49 As String
    pri_pin_52 As String
    safety_53 As String
    mac_64 As String
    macCheckFlag As String
    pan_2 As PtrStruct
    api_code_32 As PtrStruct
    track2_35 As PtrStruct
    track3_36 As PtrStruct
    rsp_code_44 As PtrStruct
    pay_signature_46 As PtrStruct
    settleAccounts_48 As PtrStruct
    attachment_amount_54 As PtrStruct
    icData_55 As PtrStruct
    private_data_56 As PtrStruct
    private_data_57 As PtrStruct
    private_data_59 As PtrStruct
    private_data_60 As PtrStruct
    private_data_61 As PtrStruct
    private_data_62 As PtrStruct
    private_data_63 As PtrStruct
    private_data_58 As PtrStruct

    private_data_21 As PtrStruct    '//2015��4��23��14:28:28���
    private_data_47 As PtrStruct    '//2015��6��18��10:51:05���
End Type


'//������
Private Sub analyse_level_1_Click()
   On Error GoTo ErrHandle    '//�������
    Dim tex2STR, tempStr, Totallen, TPDU, MessageHeader As String
    Dim tempCount As Long
    Dim POS_Sturct As POS_Sturct_TYPE
    Dim i As Integer

    PROCESS_Analyze_data    '//����ǰ����

    If change_begin = 1 Then

        analyse_after_data.SetFocus
        '//   analyse_after_data.IMEMode = 3 '//���ı���õ�����ʱ�������뷨
        analyse_mode = 1    '//�������ģʽ 1 ��ͨ����
        trace_no_11.Caption = "N/A"
        batch_Number_60.Caption = "N/A"

        help_flag = 0
        help_count = 0
        help.Caption = "˵��"
        tempCount = 1    '//����

        tempStr = analyse_after_data.Text
        '//       tex2STR = tex2STR & "/************8583������������ǻ����ķָ��߿�ͷ��***************" & vbCrLf & vbCrLf


        If Totallen_check.Value = 1 Then
            '//�ܳ�����
            Totallen = Mid(tempStr, tempCount, 4)
            Totallen = ins_space(Totallen)
            tex2STR = tex2STR & "[Totallen]" & "     " & Totallen & vbCrLf
            tempCount = tempCount + 4
        End If

        If TPDU_check.Value = 1 Then
            '//TPDU����
            TPDU = Mid(tempStr, tempCount, 10)
            TPDU = ins_space(TPDU)
            tex2STR = tex2STR & "[TPDU]" & "         " & TPDU & vbCrLf
            tempCount = tempCount + 10
        End If

        If MessageHeader_check.Value = 1 Then
            '//MessageHeader����
            MessageHeader = Mid(tempStr, tempCount, 12)
            MessageHeader = ins_space(MessageHeader)
            tex2STR = tex2STR & "[MessageHeader]" & "" & MessageHeader & vbCrLf
            tempCount = tempCount + 12
        End If


        '//messagetype����
        POS_Sturct.messagetype = Mid(tempStr, tempCount, 4)
        POS_Sturct.messagetype = ins_space(POS_Sturct.messagetype)
        tex2STR = tex2STR & "[messagetype]" & "  " & POS_Sturct.messagetype & vbCrLf
        tempCount = tempCount + 4

        '//bitmap����
        POS_Sturct.bitmap = Mid(tempStr, tempCount, 16)
        POS_Sturct.bitmap = ins_space(POS_Sturct.bitmap)
        tex2STR = tex2STR & "[bitmap]" & "       " & POS_Sturct.bitmap & vbCrLf
        tempCount = tempCount + 16
        bit_map.Text = POS_Sturct.bitmap
        bit_map_change_function

        Dim temp_bcd_flag_str As String
        Dim tempLen As Integer


        '//bitmap����
        temp_bcd_flag_str = HEX_to_BIN(delete_space(POS_Sturct.bitmap))

        Dim tempPan As String
        Dim temp_Pan_len As String
        '//��2������ �����˺š�
        If Mid(temp_bcd_flag_str, 2, 1) Then
            temp_Pan_len = Mid(tempStr, tempCount, 2)
            POS_Sturct.pan_2.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.pan_2.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.pan_2.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            tempPan = Mid(POS_Sturct.pan_2.ptr, 1, POS_Sturct.pan_2.Ptrlen)
            POS_Sturct.pan_2.ptr = ins_space(POS_Sturct.pan_2.ptr)
            tex2STR = tex2STR & "[field2]" & "       " & temp_Pan_len & " " & POS_Sturct.pan_2.ptr & vbCrLf
        End If

        '//��3������ �����״����롿
        If Mid(temp_bcd_flag_str, 3, 1) Then
            POS_Sturct.procode_3 = Mid(tempStr, tempCount, 6)
            POS_Sturct.procode_3 = ins_space(POS_Sturct.procode_3)
            tex2STR = tex2STR & "[field3]" & "       " & POS_Sturct.procode_3 & vbCrLf
            tempCount = tempCount + 6
        End If
        '//�����ж�
        i = Type_judge(POS_Sturct.messagetype, Mid(POS_Sturct.procode_3, 1, 2), Mid(temp_bcd_flag_str, 61, 1))

        Dim temp_consume_amount_4 As String
        '//��4������ �����׽�
        If Mid(temp_bcd_flag_str, 4, 1) Then
            POS_Sturct.consume_amount_4 = Mid(tempStr, tempCount, 12)
            temp_consume_amount_4 = Val(POS_Sturct.consume_amount_4) / 100
            POS_Sturct.consume_amount_4 = ins_space(POS_Sturct.consume_amount_4)
            tex2STR = tex2STR & "[field4]" & "       " & POS_Sturct.consume_amount_4 & vbCrLf
            tempCount = tempCount + 12
        End If

        Dim temp_trace_no_11 As String
        '//��11������ ���ܿ���ϵͳ���ٺš�
        If Mid(temp_bcd_flag_str, 11, 1) Then
            POS_Sturct.trace_no_11 = Mid(tempStr, tempCount, 6)
            temp_trace_no_11 = POS_Sturct.trace_no_11
            POS_Sturct.trace_no_11 = ins_space(POS_Sturct.trace_no_11)
            tex2STR = tex2STR & "[field11]" & "      " & POS_Sturct.trace_no_11 & vbCrLf
            trace_no_11.Caption = temp_trace_no_11
            tempCount = tempCount + 6
        End If


        '//��12������ ���ܿ������ڵ�ʱ�䡿
        If Mid(temp_bcd_flag_str, 12, 1) Then
            POS_Sturct.trade_time_12 = Mid(tempStr, tempCount, 6)
            POS_Sturct.trade_time_12 = ins_space(POS_Sturct.trade_time_12)
            tex2STR = tex2STR & "[field12]" & "      " & POS_Sturct.trade_time_12 & vbCrLf
            tempCount = tempCount + 6
        End If

        '//��13������ ���ܿ������ڵ����ڡ�
        If Mid(temp_bcd_flag_str, 13, 1) Then
            POS_Sturct.trade_date_13 = Mid(tempStr, tempCount, 4)
            POS_Sturct.trade_date_13 = ins_space(POS_Sturct.trade_date_13)
            tex2STR = tex2STR & "[field13]" & "      " & POS_Sturct.trade_date_13 & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��14������ ������Ч�ڡ�
        If Mid(temp_bcd_flag_str, 14, 1) Then
            POS_Sturct.exp_date_14 = Mid(tempStr, tempCount, 4)
            POS_Sturct.exp_date_14 = ins_space(POS_Sturct.exp_date_14)
            tex2STR = tex2STR & "[field14]" & "      " & POS_Sturct.exp_date_14 & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��15������ ���������ڡ�
        If Mid(temp_bcd_flag_str, 15, 1) Then
            POS_Sturct.settlement_date_15 = Mid(tempStr, tempCount, 4)
            POS_Sturct.settlement_date_15 = ins_space(POS_Sturct.settlement_date_15)
            tex2STR = tex2STR & "[field15]" & "      " & POS_Sturct.settlement_date_15 & vbCrLf
            tempCount = tempCount + 4
        End If

        Dim temp_21_len_str As String
        '//��21������ ��private_data_21��
        If Mid(temp_bcd_flag_str, 21, 1) Then
            temp_21_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_21.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_21.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_21.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_21.ptr = ins_space(POS_Sturct.private_data_21.ptr)
            temp_21_len_str = ins_space(temp_21_len_str)
            tex2STR = tex2STR & "[field21]" & "      " & temp_21_len_str & " " & POS_Sturct.private_data_21.ptr & vbCrLf

        End If

        '//��22������ ����������뷽ʽ�롿
        If Mid(temp_bcd_flag_str, 22, 1) Then
            POS_Sturct.entry_mode_22 = Mid(tempStr, tempCount, 4)
            POS_Sturct.entry_mode_22 = ins_space(POS_Sturct.entry_mode_22)
            tex2STR = tex2STR & "[field22]" & "      " & POS_Sturct.entry_mode_22 & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��23������ �������кš�
        If Mid(temp_bcd_flag_str, 23, 1) Then
            POS_Sturct.card_serial_number_23 = Mid(tempStr, tempCount, 4)
            POS_Sturct.card_serial_number_23 = ins_space(POS_Sturct.card_serial_number_23)
            tex2STR = tex2STR & "[field23]" & "      " & POS_Sturct.card_serial_number_23 & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��25������ ������������롿
        If Mid(temp_bcd_flag_str, 25, 1) Then
            POS_Sturct.service_conditon_25 = Mid(tempStr, tempCount, 2)
            POS_Sturct.service_conditon_25 = ins_space(POS_Sturct.service_conditon_25)
            tex2STR = tex2STR & "[field25]" & "      " & POS_Sturct.service_conditon_25 & vbCrLf
            tempCount = tempCount + 2
        End If

        '//��26������ �������PIN��ȡ�롿
        If Mid(temp_bcd_flag_str, 26, 1) Then
            POS_Sturct.service_conditon_pin_26 = Mid(tempStr, tempCount, 2)
            POS_Sturct.service_conditon_pin_26 = ins_space(POS_Sturct.service_conditon_pin_26)
            tex2STR = tex2STR & "[field26]" & "      " & POS_Sturct.service_conditon_pin_26 & vbCrLf
            tempCount = tempCount + 2
        End If

        Dim temp_32_len_str As String
        '//��32������ ��������ʶ�롿
        If Mid(temp_bcd_flag_str, 32, 1) Then
            temp_32_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.api_code_32.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.api_code_32.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.api_code_32.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.api_code_32.ptr = ins_space(POS_Sturct.api_code_32.ptr)

            tex2STR = tex2STR & "[field32]" & "      " & temp_32_len_str & " " & POS_Sturct.api_code_32.ptr & vbCrLf

        End If

        Dim temp_35_len_str As String
        '//��35������ ��2�ŵ����ݡ�
        If Mid(temp_bcd_flag_str, 35, 1) Then
            temp_35_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.track2_35.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.track2_35.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.track2_35.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.track2_35.ptr = ins_space(POS_Sturct.track2_35.ptr)

            tex2STR = tex2STR & "[field35]" & "      " & temp_35_len_str & " " & POS_Sturct.track2_35.ptr & vbCrLf

        End If

        Dim temp_36_len_str As String
        '//��36������ ��3�ŵ����ݡ�
        If Mid(temp_bcd_flag_str, 36, 1) Then
            temp_36_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.track3_36.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.track3_36.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.track3_36.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.track3_36.ptr = ins_space(POS_Sturct.track3_36.ptr)
            temp_36_len_str = ins_space(temp_36_len_str)
            tex2STR = tex2STR & "[field36]" & "      " & temp_36_len_str & " " & POS_Sturct.track3_36.ptr & vbCrLf

        End If

        Dim temp_37str As String
        '//��37������ �������ο��š�
        If Mid(temp_bcd_flag_str, 37, 1) Then
            POS_Sturct.reference_number_37 = Mid(tempStr, tempCount, 24)
            POS_Sturct.reference_number_37 = ins_space(POS_Sturct.reference_number_37)
            temp_37str = ASCchange(POS_Sturct.reference_number_37)
            tex2STR = tex2STR & "[field37]" & "      " & POS_Sturct.reference_number_37 & vbCrLf
            tempCount = tempCount + 24
        End If

        Dim temp_38str As String
        '//��38������ ����Ȩ��ʶӦ���롿
        If Mid(temp_bcd_flag_str, 38, 1) Then
            POS_Sturct.authorization_code_38 = Mid(tempStr, tempCount, 12)
            POS_Sturct.authorization_code_38 = ins_space(POS_Sturct.authorization_code_38)
            temp_38str = ASCchange(POS_Sturct.authorization_code_38)
            tex2STR = tex2STR & "[field38]" & "      " & POS_Sturct.authorization_code_38 & vbCrLf
            tempCount = tempCount + 12
        End If

        Dim temp_39str As String
        '//��39������ ��Ӧ���롿
        If Mid(temp_bcd_flag_str, 39, 1) Then
            POS_Sturct.Response_code_39 = Mid(tempStr, tempCount, 4)
            POS_Sturct.Response_code_39 = ins_space(POS_Sturct.Response_code_39)
            temp_39str = ASCchange(POS_Sturct.Response_code_39)
            tex2STR = tex2STR & "[field39]" & "      " & POS_Sturct.Response_code_39 & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��Ӧ���ж���ʾ
        i = Response_code_39_Type_judge(ASCchange(POS_Sturct.Response_code_39))

        Dim temp_41str As String
        '//��41������ ���ܿ����ն˱�ʶ�롿
        If Mid(temp_bcd_flag_str, 41, 1) Then
            POS_Sturct.terminal_no_41 = Mid(tempStr, tempCount, 16)
            POS_Sturct.terminal_no_41 = ins_space(POS_Sturct.terminal_no_41)
            temp_41str = ASCchange(POS_Sturct.terminal_no_41)
            tex2STR = tex2STR & "[field41]" & "      " & POS_Sturct.terminal_no_41 & vbCrLf
            tempCount = tempCount + 16
        End If

        Dim temp_42str As String
        '//��42������ ���ܿ�����ʶ�롿
        If Mid(temp_bcd_flag_str, 42, 1) Then
            POS_Sturct.merchant_no_42 = Mid(tempStr, tempCount, 30)
            POS_Sturct.merchant_no_42 = ins_space(POS_Sturct.merchant_no_42)
            temp_42str = ASCchange(POS_Sturct.merchant_no_42)
            tex2STR = tex2STR & "[field42]" & "      " & POS_Sturct.merchant_no_42 & vbCrLf
            tempCount = tempCount + 30
        End If


        Dim temp_43_len_str As String
        '//��43������ ��merchant_name_43��
        If Mid(temp_bcd_flag_str, 43, 1) Then
            temp_43_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.merchant_name_43.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.merchant_name_43.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.merchant_name_43.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.merchant_name_43.ptr = ins_space(POS_Sturct.merchant_name_43.ptr)
            temp_43_len_str = ins_space(temp_43_len_str)
            tex2STR = tex2STR & "[field43]" & "      " & temp_43_len_str & " " & POS_Sturct.merchant_name_43.ptr & vbCrLf

        End If


        Dim temp_44_len_str As String
        Dim issuing_bank As String   '//������
        Dim Acquiring_bank As String    '//�յ���
        '//��44������ ��������Ӧ���ݡ�
        If Mid(temp_bcd_flag_str, 44, 1) Then
            temp_44_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.rsp_code_44.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = POS_Sturct.rsp_code_44.Ptrlen * 2
            tempCount = tempCount + 2

            POS_Sturct.rsp_code_44.ptr = Mid(tempStr, tempCount, tempLen)
            issuing_bank = ASCchange(Mid(POS_Sturct.rsp_code_44.ptr, 1, 22))
            Acquiring_bank = ASCchange(Mid(POS_Sturct.rsp_code_44.ptr, 23, 22))


            tempCount = tempCount + tempLen
            POS_Sturct.rsp_code_44.ptr = ins_space(POS_Sturct.rsp_code_44.ptr)
            tex2STR = tex2STR & "[field44]" & "      " & temp_44_len_str & " " & POS_Sturct.rsp_code_44.ptr & vbCrLf

        End If

        Dim temp_46_len_str As String
        '//��46������ ��pay_signature_46��
        If Mid(temp_bcd_flag_str, 46, 1) Then
            temp_46_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.pay_signature_46.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.pay_signature_46.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.pay_signature_46.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.pay_signature_46.ptr = ins_space(POS_Sturct.pay_signature_46.ptr)
            temp_46_len_str = ins_space(temp_46_len_str)
            tex2STR = tex2STR & "[field46]" & "      " & temp_46_len_str & " " & POS_Sturct.pay_signature_46.ptr & vbCrLf

        End If

        Dim temp_47_len_str As String
        '//��47������ ��private_data_47��
        If Mid(temp_bcd_flag_str, 47, 1) Then
            temp_47_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_47.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_47.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_47.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_47.ptr = ins_space(POS_Sturct.private_data_47.ptr)
            temp_47_len_str = ins_space(temp_47_len_str)
            tex2STR = tex2STR & "[field47]" & "      " & temp_47_len_str & " " & POS_Sturct.private_data_47.ptr & vbCrLf

        End If




        Dim temp_48_len_str As String
        '//��48������ ���������� - ˽�С�
        If Mid(temp_bcd_flag_str, 48, 1) Then
            temp_48_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.settleAccounts_48.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.settleAccounts_48.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.settleAccounts_48.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.settleAccounts_48.ptr = ins_space(POS_Sturct.settleAccounts_48.ptr)
            temp_48_len_str = ins_space(temp_48_len_str)
            tex2STR = tex2STR & "[field48]" & "      " & temp_48_len_str & " " & POS_Sturct.settleAccounts_48.ptr & vbCrLf
        End If

        Dim temp_49str As String
        '//��49������ �����׻��Ҵ��롿
        If Mid(temp_bcd_flag_str, 49, 1) Then
            POS_Sturct.currency_code_49 = Mid(tempStr, tempCount, 6)
            POS_Sturct.currency_code_49 = ins_space(POS_Sturct.currency_code_49)
            tex2STR = tex2STR & "[field49]" & "      " & POS_Sturct.currency_code_49 & vbCrLf
            tempCount = tempCount + 6
        End If

        '//��52������ �����˱�ʶ�롿
        If Mid(temp_bcd_flag_str, 52, 1) Then
            POS_Sturct.pri_pin_52 = Mid(tempStr, tempCount, 16)
            POS_Sturct.pri_pin_52 = ins_space(POS_Sturct.pri_pin_52)
            tex2STR = tex2STR & "[field52]" & "      " & POS_Sturct.pri_pin_52 & vbCrLf

            tempCount = tempCount + 16
        End If

        '//��53������ ����ȫ������Ϣ��
        If Mid(temp_bcd_flag_str, 53, 1) Then
            POS_Sturct.safety_53 = Mid(tempStr, tempCount, 16)
            POS_Sturct.safety_53 = ins_space(POS_Sturct.safety_53)
            tex2STR = tex2STR & "[field53]" & "      " & POS_Sturct.safety_53 & vbCrLf

            tempCount = tempCount + 16
        End If

        Dim temp_54str As String
        Dim temp_54_len_str As String
        Dim temp_consume_amount_54 As String
        '//��54������ ����
        If Mid(temp_bcd_flag_str, 54, 1) Then
            temp_54_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.attachment_amount_54.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.attachment_amount_54.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.attachment_amount_54.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen

            POS_Sturct.attachment_amount_54.ptr = ins_space(POS_Sturct.attachment_amount_54.ptr)
            temp_54_len_str = ins_space(temp_54_len_str)

            temp_54str = ASCchange(POS_Sturct.attachment_amount_54.ptr)
            temp_consume_amount_54 = Val(Mid(temp_54str, 9, 12)) / 100
            tex2STR = tex2STR & "[field54]" & "      " & temp_54_len_str & " " & POS_Sturct.attachment_amount_54.ptr & vbCrLf
        End If

        Dim temp_55_len_str As String

        Dim POS_Sturct_icData_55_temp_ptr As String       '//55����ʱ����
        Dim temp_55_str_ptr As Integer                     '//�����ʼλ��
        Dim temp_55_str_ptr_len As String                  '//���򳤶�
        '//��55������ ��IC��������
        If Mid(temp_bcd_flag_str, 55, 1) Then
            temp_55_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.icData_55.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.icData_55.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.icData_55.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen


            '            POS_Sturct.icData_55.ptr = ins_space(POS_Sturct.icData_55.ptr)
            '//���55��


            '//������Ϣ�����б�
            '�γ����ָ�ʽ ->[9F 26] [08] 5E 14 AA 9F 20 46 A9 21   HEX_to_DEC
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F26")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                '            temp_55_str_ptr_len = temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)
                '
                '            Dim tempStr1 As String
                '            Dim tempStr2 As String
                '
                '            tempStr1 = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1)
                '            tempStr2 = Right(POS_Sturct.icData_55.ptr, Val(temp_55_str_ptr_len))
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F27")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F10")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F37")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F36")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "95")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9C")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F02")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "5F2A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "82")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))   '//ȥ������������Դ
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F1A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F03")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F33")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F74")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
                '//��ѡ��Ϣ�����б�
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F34")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F35")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F1E")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "84")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F09")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf



                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F41")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "91")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "71")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "72")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF31")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F63")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            '//�ѻ�����ר�������б�
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "8A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            '// �ֻ�оƬ����ר�������б�
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF32")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF33")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf



                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF34")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            POS_Sturct.icData_55.ptr = POS_Sturct_icData_55_temp_ptr
            temp_55_len_str = ins_space(temp_55_len_str)
            tex2STR = tex2STR & "[field55]" & "      " & temp_55_len_str & vbCrLf & POS_Sturct.icData_55.ptr
            'Debug.Print tex2STR
        End If

        Dim temp_56_len_str As String
        '//��56������ ��private_data_56��
        If Mid(temp_bcd_flag_str, 56, 1) Then
            temp_56_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_56.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_56.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_56.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_56.ptr = ins_space(POS_Sturct.private_data_56.ptr)
            temp_56_len_str = ins_space(temp_56_len_str)
            tex2STR = tex2STR & "[field56]" & "      " & temp_56_len_str & " " & POS_Sturct.private_data_56.ptr & vbCrLf

        End If

        Dim temp_57_len_str As String
        '//��57������ ��private_data_57��
        If Mid(temp_bcd_flag_str, 57, 1) Then
            temp_57_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_57.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_57.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_57.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_57.ptr = ins_space(POS_Sturct.private_data_57.ptr)
            temp_57_len_str = ins_space(temp_57_len_str)
            tex2STR = tex2STR & "[field57]" & "      " & temp_57_len_str & " " & POS_Sturct.private_data_57.ptr & vbCrLf

        End If


        Dim temp_58_len_str As String
        '//��58������ ��PBOC����Ǯ����׼�Ľ�����Ϣ��
        If Mid(temp_bcd_flag_str, 58, 1) Then
            temp_58_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_58.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_58.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_58.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_58.ptr = ins_space(POS_Sturct.private_data_58.ptr)
            temp_58_len_str = ins_space(temp_58_len_str)
            tex2STR = tex2STR & "[field58]" & "      " & temp_58_len_str & " " & POS_Sturct.private_data_58.ptr & vbCrLf

        End If

        Dim temp_59_len_str As String
        '//��59������ ��private_data_59��
        If Mid(temp_bcd_flag_str, 59, 1) Then
            temp_59_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_59.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_59.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_59.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_59.ptr = ins_space(POS_Sturct.private_data_59.ptr)
            temp_59_len_str = ins_space(temp_59_len_str)
            tex2STR = tex2STR & "[field59]" & "      " & temp_59_len_str & " " & POS_Sturct.private_data_59.ptr & vbCrLf

        End If

        Dim temp_60_len_str As String
        Dim temp_batch_Number_60 As String
        '//��60������ ��private_data_60��
        If Mid(temp_bcd_flag_str, 60, 1) Then
            temp_60_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_60.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.private_data_60.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_60.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen


            temp_batch_Number_60 = Mid(POS_Sturct.private_data_60.ptr, 3, 6)
            POS_Sturct.private_data_60.ptr = ins_space(POS_Sturct.private_data_60.ptr)

            temp_60_len_str = ins_space(temp_60_len_str)
            tex2STR = tex2STR & "[field60]" & "      " & temp_60_len_str & " " & POS_Sturct.private_data_60.ptr & vbCrLf
            batch_Number_60.Caption = temp_batch_Number_60

        End If

        Dim temp_61_len_str As String
        Dim Original_batch_Number_61 As String    '//ԭʼ�������κ�
        Dim Original_trace_no_61 As String    '//ԭʼ����POS��ˮ��

        '//��61������ ��ԭʼ��Ϣ��
        If Mid(temp_bcd_flag_str, 61, 1) Then
            temp_61_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_61.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.private_data_61.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_61.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            Original_batch_Number_61 = Mid(POS_Sturct.private_data_61.ptr, 1, 6)
            Original_trace_no_61 = Mid(POS_Sturct.private_data_61.ptr, 7, 6)

            POS_Sturct.private_data_61.ptr = ins_space(POS_Sturct.private_data_61.ptr)

            temp_61_len_str = ins_space(temp_61_len_str)
            tex2STR = tex2STR & "[field61]" & "      " & temp_61_len_str & " " & POS_Sturct.private_data_61.ptr & vbCrLf

        End If

        Dim temp_62_len_str As String
        '//��62������ ��private_data_62��
        If Mid(temp_bcd_flag_str, 62, 1) Then
            temp_62_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_62.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_62.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_62.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_62.ptr = ins_space(POS_Sturct.private_data_62.ptr)
            temp_62_len_str = ins_space(temp_62_len_str)
            tex2STR = tex2STR & "[field62]" & "      " & temp_62_len_str & " " & POS_Sturct.private_data_62.ptr & vbCrLf

        End If

        Dim temp_63_len_str As String
        '//��63������ ��private_data_63��
        If Mid(temp_bcd_flag_str, 63, 1) Then
            temp_63_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_63.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_63.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_63.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_63.ptr = ins_space(POS_Sturct.private_data_63.ptr)
            temp_63_len_str = ins_space(temp_63_len_str)
            tex2STR = tex2STR & "[field63]" & "      " & temp_63_len_str & " " & POS_Sturct.private_data_63.ptr & vbCrLf

        End If

        '//��64������ �����ļ����롿
        If Mid(temp_bcd_flag_str, 64, 1) Then
            POS_Sturct.mac_64 = Mid(tempStr, tempCount, 16)
            POS_Sturct.mac_64 = ins_space(POS_Sturct.mac_64)
            tex2STR = tex2STR & "[field64]" & "      " & POS_Sturct.mac_64 & vbCrLf
            tempCount = tempCount + 16
        End If



        '//  tex2STR = tex2STR & "************8583������������ǻ����ķָ��߽�β��***************/" & vbCrLf
        analyse_after_data.Text = tex2STR
    End If

    Exit Sub  'һ��Ҫд
ErrHandle:
    'MSG_BOX Err.Number & Err.Source, "����"
    ' Err.clear

    analyse_after_data.Text = ""
    bit_map_clear_Click
    trans_type.Caption = "N/A"
    judge_mode.Caption = "N/A"
    judge_mode.ForeColor = &H0&    '//�ָ��ɺ�ɫ

    Response_code.Caption = "N/A"
    Response_code_view.Caption = "N/A"

    trace_no_11.Caption = "N/A"
    batch_Number_60.Caption = "N/A"

    help_flag = 0
    help.Caption = "˵��"
    analyse_before_data.SetFocus


    MSG_BOX _
            "[Err.Source      ]:" & Err.Source & vbCrLf & _
                                  "[Err.Number      ]:" & Err.Number & vbCrLf & _
                                  "[Err.Description ]:" & Err.Description & vbCrLf, "����"

    '"[Err.HelpContext ]:" & Err.HelpContext & vbCrLf & _
     '"[Err.HelpFile    ]:" & Err.HelpFile & vbCrLf & _
     '"[Err.LastDllError]:" & Err.LastDllError & vbCrLf & _

     Err.clear

End Sub


'//������
Private Sub analyse_level_2_Click()
     'On Error GoTo ErrHandle    '//�������
    Dim tex2STR, tempStr, Totallen, TPDU, MessageHeader As String
    Dim tempCount As Long
    Dim POS_Sturct As POS_Sturct_TYPE
    Dim i As Integer

    PROCESS_Analyze_data    '//����ǰ����

    If change_begin = 1 Then

        analyse_after_data.SetFocus
        analyse_mode = 2    '//�������ģʽ 2 ר�ҽ���
        trace_no_11.Caption = "N/A"
        batch_Number_60.Caption = "N/A"
        help_flag = 0
        help_count = 0
        help.Caption = "˵��"
        tempCount = 1    '//����


        tempStr = analyse_after_data.Text
        '//       tex2STR = tex2STR & "/************8583������������ǻ����ķָ��߿�ͷ��***************" & vbCrLf & vbCrLf

        If Totallen_check.Value = 1 Then
            '//�ܳ�����

            Totallen = Mid(tempStr, tempCount, 4)
            Totallen = ins_space(Totallen)
            tex2STR = tex2STR & "[Totallen:�ܳ���]" & vbCrLf & "" & Totallen & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        If TPDU_check.Value = 1 Then
            '//TPDU����
            TPDU = Mid(tempStr, tempCount, 10)
            TPDU = ins_space(TPDU)
            tex2STR = tex2STR & "[TPDU:��ַ]" & vbCrLf & "" & TPDU & vbCrLf & vbCrLf
            tempCount = tempCount + 10
        End If

        If MessageHeader_check.Value = 1 Then
            '//MessageHeader����
            MessageHeader = Mid(tempStr, tempCount, 12)
            MessageHeader = ins_space(MessageHeader)
            tex2STR = tex2STR & "[MessageHeader:����ͷ]" & vbCrLf & "" & MessageHeader & vbCrLf & vbCrLf
            tempCount = tempCount + 12
        End If


        '//messagetype����
        POS_Sturct.messagetype = Mid(tempStr, tempCount, 4)
        POS_Sturct.messagetype = ins_space(POS_Sturct.messagetype)
        tex2STR = tex2STR & "[messagetype:��Ϣ����]" & vbCrLf & "" & POS_Sturct.messagetype & vbCrLf & vbCrLf
        tempCount = tempCount + 4

        '//bitmap����
        POS_Sturct.bitmap = Mid(tempStr, tempCount, 16)
        POS_Sturct.bitmap = ins_space(POS_Sturct.bitmap)
        tex2STR = tex2STR & "[bitmap:λԪ��]" & vbCrLf & "" & POS_Sturct.bitmap & vbCrLf & vbCrLf
        tempCount = tempCount + 16
        bit_map.Text = POS_Sturct.bitmap
        bit_map_change_function

        Dim temp_bcd_flag_str As String
        Dim tempLen As Integer


        '//bitmap����
        temp_bcd_flag_str = HEX_to_BIN(delete_space(POS_Sturct.bitmap))

        Dim tempPan As String
        Dim temp_Pan_len As String
        '//��2������ �����˺š�
        If Mid(temp_bcd_flag_str, 2, 1) Then
            temp_Pan_len = Mid(tempStr, tempCount, 2)
            POS_Sturct.pan_2.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.pan_2.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.pan_2.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            tempPan = Mid(POS_Sturct.pan_2.ptr, 1, POS_Sturct.pan_2.Ptrlen)
            POS_Sturct.pan_2.ptr = ins_space(POS_Sturct.pan_2.ptr)
            tex2STR = tex2STR & "[field2:���˺�(Primary Account Number)]" & vbCrLf & "" & "[" & temp_Pan_len & "]" & " " & POS_Sturct.pan_2.ptr _
                      & vbCrLf & "->[����] " & tempPan & vbCrLf & vbCrLf
        End If

        '//��3������ �����״����롿
        If Mid(temp_bcd_flag_str, 3, 1) Then
            POS_Sturct.procode_3 = Mid(tempStr, tempCount, 6)
            POS_Sturct.procode_3 = ins_space(POS_Sturct.procode_3)
            tex2STR = tex2STR & "[field3:���״�����(Transaction Processing Code)]" & vbCrLf & "" & POS_Sturct.procode_3 & vbCrLf & vbCrLf
            tempCount = tempCount + 6
        End If
        '//�����ж�
        i = Type_judge(POS_Sturct.messagetype, Mid(POS_Sturct.procode_3, 1, 2), Mid(temp_bcd_flag_str, 61, 1))

        Dim temp_consume_amount_4 As String
        '//��4������ �����׽�
        If Mid(temp_bcd_flag_str, 4, 1) Then
            POS_Sturct.consume_amount_4 = Mid(tempStr, tempCount, 12)
            temp_consume_amount_4 = Val(POS_Sturct.consume_amount_4) / 100

            If Val(temp_consume_amount_4) < 1 And Val(temp_consume_amount_4) > 0 Then
                temp_consume_amount_4 = 0 & temp_consume_amount_4
            End If


            POS_Sturct.consume_amount_4 = ins_space(POS_Sturct.consume_amount_4)
            tex2STR = tex2STR & "[field4:���׽��(Amount Of Transactions)]" & vbCrLf & "" & POS_Sturct.consume_amount_4 & _
                      vbCrLf & "->[�����] " & temp_consume_amount_4 & "Ԫ" & vbCrLf & vbCrLf
            tempCount = tempCount + 12
        End If

        Dim temp_trace_no_11 As String
        '//��11������ ���ܿ���ϵͳ���ٺš�
        If Mid(temp_bcd_flag_str, 11, 1) Then
            POS_Sturct.trace_no_11 = Mid(tempStr, tempCount, 6)
            temp_trace_no_11 = POS_Sturct.trace_no_11
            POS_Sturct.trace_no_11 = ins_space(POS_Sturct.trace_no_11)
            tex2STR = tex2STR & "[field11:�ܿ���ϵͳ���ٺ�(System Trace Audit Number)]" & vbCrLf & "" & POS_Sturct.trace_no_11 & _
                      vbCrLf & "->[��ˮ��] " & temp_trace_no_11 & vbCrLf & vbCrLf

            trace_no_11.Caption = temp_trace_no_11

            tempCount = tempCount + 6
        End If


        '//��12������ ���ܿ������ڵ�ʱ�䡿
        If Mid(temp_bcd_flag_str, 12, 1) Then
            POS_Sturct.trade_time_12 = Mid(tempStr, tempCount, 6)
            POS_Sturct.trade_time_12 = ins_space(POS_Sturct.trade_time_12)
            tex2STR = tex2STR & "[field12:�ܿ������ڵ�ʱ��(Local Time Of Transaction)]" & vbCrLf & "" & POS_Sturct.trade_time_12 & _
                      vbCrLf & "->[ʱ��] " & Mid(POS_Sturct.trade_time_12, 1, 2) & "ʱ" & Mid(POS_Sturct.trade_time_12, 4, 2) & "��" _
                      & Mid(POS_Sturct.trade_time_12, 7, 2) & "��" & vbCrLf & vbCrLf
            tempCount = tempCount + 6
        End If

        '//��13������ ���ܿ������ڵ����ڡ�
        If Mid(temp_bcd_flag_str, 13, 1) Then
            POS_Sturct.trade_date_13 = Mid(tempStr, tempCount, 4)
            POS_Sturct.trade_date_13 = ins_space(POS_Sturct.trade_date_13)
            tex2STR = tex2STR & "[field13:�ܿ������ڵ�����(Local Date Of Transaction)]" & vbCrLf & "" & POS_Sturct.trade_date_13 & "   " & _
                      vbCrLf & "->[����] " & Mid(POS_Sturct.trade_date_13, 1, 2) & "��" & Mid(POS_Sturct.trade_date_13, 4, 2) & "��" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��14������ ������Ч�ڡ�
        If Mid(temp_bcd_flag_str, 14, 1) Then
            POS_Sturct.exp_date_14 = Mid(tempStr, tempCount, 4)
            POS_Sturct.exp_date_14 = ins_space(POS_Sturct.exp_date_14)
            tex2STR = tex2STR & "[field14:����Ч��(Date Of Expired)]" & vbCrLf & "" & POS_Sturct.exp_date_14 & "   " & _
                      vbCrLf & "->[��Ч��] " & "20" & Mid(POS_Sturct.exp_date_14, 1, 2) & "��" & Mid(POS_Sturct.exp_date_14, 4, 2) & "��" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��15������ ���������ڡ�
        If Mid(temp_bcd_flag_str, 15, 1) Then
            POS_Sturct.settlement_date_15 = Mid(tempStr, tempCount, 4)
            POS_Sturct.settlement_date_15 = ins_space(POS_Sturct.settlement_date_15)
            tex2STR = tex2STR & "[field15:��������(Date Of Settlement)]" & vbCrLf & "" & POS_Sturct.settlement_date_15 & "   " & _
                      vbCrLf & "->[��������] " & Mid(POS_Sturct.settlement_date_15, 1, 2) & "��" & Mid(POS_Sturct.settlement_date_15, 4, 2) & "��" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If
        Dim temp_21_len_str As String
        '//��21������ ��private_data_21��
        If Mid(temp_bcd_flag_str, 21, 1) Then
            temp_21_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_21.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_21.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_21.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_21.ptr = ins_space(POS_Sturct.private_data_21.ptr)
            temp_21_len_str = ins_space(temp_21_len_str)
            tex2STR = tex2STR & "[field21:�Զ�����(private_data_21)]" & vbCrLf & "" & "[" & temp_21_len_str & "]" & "  " & POS_Sturct.private_data_21.ptr & vbCrLf & vbCrLf

        End If

        '//��22������ ����������뷽ʽ�롿
        If Mid(temp_bcd_flag_str, 22, 1) Then
            POS_Sturct.entry_mode_22 = Mid(tempStr, tempCount, 4)
            POS_Sturct.entry_mode_22 = ins_space(POS_Sturct.entry_mode_22)
            tex2STR = tex2STR & "[field22:��������뷽ʽ��(Point Of Service Entry Mode)]" & vbCrLf & "" & POS_Sturct.entry_mode_22 & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '        '//��23������ �������кš�
        '        If Mid(temp_bcd_flag_str, 23, 1) Then
        '            POS_Sturct.card_serial_number_23 = Mid(tempStr, tempCount, 4)
        '            POS_Sturct.card_serial_number_23 = ins_space(POS_Sturct.card_serial_number_23)
        '            tex2STR = tex2STR & "[field23:�����к�(Card Sequence Number)]" & vbCrLf & "" & POS_Sturct.card_serial_number_23 & vbCrLf & vbCrLf
        '            tempCount = tempCount + 4
        '        End If









        Dim Right_23_str As String    '//��Ƭ���к��ҿ�
        Dim left_23_str As String    '//��Ƭ���к���
        '//��23������ �������кš�
        If Mid(temp_bcd_flag_str, 23, 1) Then
            POS_Sturct.card_serial_number_23 = Mid(tempStr, tempCount, 4)
            Right_23_str = Mid(POS_Sturct.card_serial_number_23, 2, 3)
            left_23_str = Mid(POS_Sturct.card_serial_number_23, 1, 3)
            POS_Sturct.card_serial_number_23 = ins_space(POS_Sturct.card_serial_number_23)
            tex2STR = tex2STR & "[field23:�����к�(Card Sequence Number)]" & vbCrLf & "" & POS_Sturct.card_serial_number_23 & vbCrLf _
                      & "->[��Ƭ���к���] " & left_23_str & vbCrLf & "->[��Ƭ���к��ҿ�] " & Right_23_str & " [һ������ѡ���ҿ�]" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If


        '//��25������ ������������롿
        If Mid(temp_bcd_flag_str, 25, 1) Then
            POS_Sturct.service_conditon_25 = Mid(tempStr, tempCount, 2)
            POS_Sturct.service_conditon_25 = ins_space(POS_Sturct.service_conditon_25)
            tex2STR = tex2STR & "[field25:�����������(Point Of Service Condition Mode)]" & vbCrLf & "" & POS_Sturct.service_conditon_25 & vbCrLf & vbCrLf
            tempCount = tempCount + 2
        End If

        '//��26������ �������PIN��ȡ�롿
        If Mid(temp_bcd_flag_str, 26, 1) Then
            POS_Sturct.service_conditon_pin_26 = Mid(tempStr, tempCount, 2)
            POS_Sturct.service_conditon_pin_26 = ins_space(POS_Sturct.service_conditon_pin_26)
            tex2STR = tex2STR & "[field26:�����PIN��ȡ��(Point Of Service PIN Capture Code)]" & vbCrLf & "" & POS_Sturct.service_conditon_pin_26 & vbCrLf & vbCrLf
            tempCount = tempCount + 2
        End If

        Dim temp_32_len_str As String
        '//��32������ ��������ʶ�롿
        If Mid(temp_bcd_flag_str, 32, 1) Then
            temp_32_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.api_code_32.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.api_code_32.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.api_code_32.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.api_code_32.ptr = ins_space(POS_Sturct.api_code_32.ptr)

            tex2STR = tex2STR & "[field32:������ʶ��(Acquiring Institution Id Code)]" & vbCrLf & "" & "[" & temp_32_len_str & "]" & " " & POS_Sturct.api_code_32.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_35_len_str As String
        '//��35������ ��2�ŵ����ݡ�
        If Mid(temp_bcd_flag_str, 35, 1) Then
            temp_35_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.track2_35.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.track2_35.Ptrlen + 1) \ 2) * 2
            '// tempLen = POS_Sturct.track2_35.Ptrlen * 2
            tempCount = tempCount + 2

            POS_Sturct.track2_35.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.track2_35.ptr = ins_space(POS_Sturct.track2_35.ptr)

            tex2STR = tex2STR & "[field35:2�ŵ�����(Track 2 Data)]" & vbCrLf & "" & "[" & temp_35_len_str & "]" & "  " & POS_Sturct.track2_35.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_36_len_str As String
        '//��36������ ��3�ŵ����ݡ�
        If Mid(temp_bcd_flag_str, 36, 1) Then
            temp_36_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.track3_36.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.track3_36.Ptrlen + 1) \ 2) * 2
            '//tempLen = POS_Sturct.track3_36.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.track3_36.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.track3_36.ptr = ins_space(POS_Sturct.track3_36.ptr)
            temp_36_len_str = ins_space(temp_36_len_str)
            tex2STR = tex2STR & "[field36:3�ŵ�����(Track 3 Data)]" & vbCrLf & "" & "[" & temp_36_len_str & "]" & "  " & POS_Sturct.track3_36.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_37str As String
        '//��37������ �������ο��š�
        If Mid(temp_bcd_flag_str, 37, 1) Then
            POS_Sturct.reference_number_37 = Mid(tempStr, tempCount, 24)
            POS_Sturct.reference_number_37 = ins_space(POS_Sturct.reference_number_37)
            temp_37str = ASCchange(POS_Sturct.reference_number_37)
            tex2STR = tex2STR & "[field37:�����ο���(Retrieval Reference Number)]" & vbCrLf & "" & POS_Sturct.reference_number_37 & vbCrLf & "->[�ο���] " & temp_37str & vbCrLf & vbCrLf
            tempCount = tempCount + 24
        End If

        Dim temp_38str As String
        '//��38������ ����Ȩ��ʶӦ���롿
        If Mid(temp_bcd_flag_str, 38, 1) Then
            POS_Sturct.authorization_code_38 = Mid(tempStr, tempCount, 12)
            POS_Sturct.authorization_code_38 = ins_space(POS_Sturct.authorization_code_38)
            temp_38str = ASCchange(POS_Sturct.authorization_code_38)
            tex2STR = tex2STR & "[field38:��Ȩ��ʶӦ����(Authorization Id Response Code)]" & vbCrLf & "" & POS_Sturct.authorization_code_38 & vbCrLf & "->" & "[��Ȩ��] " & temp_38str & vbCrLf & vbCrLf
            tempCount = tempCount + 12
        End If

        Dim temp_39str As String
        '//��39������ ��Ӧ���롿
        If Mid(temp_bcd_flag_str, 39, 1) Then
            POS_Sturct.Response_code_39 = Mid(tempStr, tempCount, 4)
            POS_Sturct.Response_code_39 = ins_space(POS_Sturct.Response_code_39)
            temp_39str = ASCchange(POS_Sturct.Response_code_39)
            tex2STR = tex2STR & "[field39:Ӧ����(Response Code)]" & vbCrLf & "" & POS_Sturct.Response_code_39 & vbCrLf & "->" & "[Ӧ����:] " & temp_39str & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��Ӧ���ж���ʾ
        i = Response_code_39_Type_judge(ASCchange(POS_Sturct.Response_code_39))

        Dim temp_41str As String
        '//��41������ ���ܿ����ն˱�ʶ�롿
        If Mid(temp_bcd_flag_str, 41, 1) Then
            POS_Sturct.terminal_no_41 = Mid(tempStr, tempCount, 16)
            POS_Sturct.terminal_no_41 = ins_space(POS_Sturct.terminal_no_41)
            temp_41str = ASCchange(POS_Sturct.terminal_no_41)
            tex2STR = tex2STR & "[field41:�ܿ����ն˱�ʶ��(Card Acceptor Terminal Id)]" & vbCrLf & "" & POS_Sturct.terminal_no_41 & vbCrLf & "->[�ն˺�] " & temp_41str & vbCrLf & vbCrLf
            tempCount = tempCount + 16
        End If

        Dim temp_42str As String
        '//��42������ ���ܿ�����ʶ�롿
        If Mid(temp_bcd_flag_str, 42, 1) Then
            POS_Sturct.merchant_no_42 = Mid(tempStr, tempCount, 30)
            POS_Sturct.merchant_no_42 = ins_space(POS_Sturct.merchant_no_42)
            temp_42str = ASCchange(POS_Sturct.merchant_no_42)
            tex2STR = tex2STR & "[field42:�ܿ�����ʶ��(Card Acceptor Id Code)]" & vbCrLf & "" & POS_Sturct.merchant_no_42 & vbCrLf & "->[�̻���] " & temp_42str & vbCrLf & vbCrLf
            tempCount = tempCount + 30
        End If


        Dim temp_43_len_str As String
        '//��43������ ��merchant_name_43��
        If Mid(temp_bcd_flag_str, 43, 1) Then
            temp_43_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.merchant_name_43.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.merchant_name_43.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.merchant_name_43.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.merchant_name_43.ptr = ins_space(POS_Sturct.merchant_name_43.ptr)
            temp_43_len_str = ins_space(temp_43_len_str)
            tex2STR = tex2STR & "[field43:�Զ�����(merchant_name_43)]" & vbCrLf & "" & "[" & temp_43_len_str & "]" & "  " & POS_Sturct.merchant_name_43.ptr & vbCrLf & vbCrLf

        End If


        Dim temp_44_len_str As String
        Dim issuing_bank As String   '//������
        Dim Acquiring_bank As String    '//�յ���
        '//��44������ ��������Ӧ���ݡ�
        If Mid(temp_bcd_flag_str, 44, 1) Then
            temp_44_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.rsp_code_44.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = POS_Sturct.rsp_code_44.Ptrlen * 2
            tempCount = tempCount + 2

            POS_Sturct.rsp_code_44.ptr = Mid(tempStr, tempCount, tempLen)
            issuing_bank = ASCchange(Mid(POS_Sturct.rsp_code_44.ptr, 1, 22))
            Acquiring_bank = ASCchange(Mid(POS_Sturct.rsp_code_44.ptr, 23, 22))


            tempCount = tempCount + tempLen
            POS_Sturct.rsp_code_44.ptr = ins_space(POS_Sturct.rsp_code_44.ptr)
            If tempLen = 44 Then
                tex2STR = tex2STR & "[field44:������Ӧ����(Additional Response Data)]" & vbCrLf & "" & "[" & temp_44_len_str & "]" & "  " & POS_Sturct.rsp_code_44.ptr _
                          & vbCrLf & "->[������] " & issuing_bank & vbCrLf & "->[�յ���] " & Acquiring_bank & vbCrLf & vbCrLf
            Else
                tex2STR = tex2STR & "[field44:������Ӧ����(Additional Response Data)]" & vbCrLf & "" & "[" & temp_44_len_str & "]" & "  " & POS_Sturct.rsp_code_44.ptr & vbCrLf & vbCrLf
            End If
        End If

        Dim temp_46_len_str As String
        '//��46������ ��pay_signature_46��
        If Mid(temp_bcd_flag_str, 46, 1) Then
            temp_46_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.pay_signature_46.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.pay_signature_46.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.pay_signature_46.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.pay_signature_46.ptr = ins_space(POS_Sturct.pay_signature_46.ptr)
            temp_46_len_str = ins_space(temp_46_len_str)
            tex2STR = tex2STR & "[field46:�Զ�����(pay_signature_46)]" & vbCrLf & "" & "[" & temp_46_len_str & "]" & "  " & POS_Sturct.pay_signature_46.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_47_len_str As String
        '//��47������ ��private_data_47��
        If Mid(temp_bcd_flag_str, 47, 1) Then
            temp_47_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_47.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_47.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_47.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_47.ptr = ins_space(POS_Sturct.private_data_47.ptr)
            temp_47_len_str = ins_space(temp_47_len_str)
            tex2STR = tex2STR & "[field47:�Զ�����(private_data_47)]" & vbCrLf & "" & "[" & temp_47_len_str & "]" & "  " & POS_Sturct.private_data_47.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_48_len_str As String
        '//��48������ ���������� - ˽�С�
        If Mid(temp_bcd_flag_str, 48, 1) Then
            temp_48_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.settleAccounts_48.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.settleAccounts_48.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.settleAccounts_48.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.settleAccounts_48.ptr = ins_space(POS_Sturct.settleAccounts_48.ptr)
            temp_48_len_str = ins_space(temp_48_len_str)
            tex2STR = tex2STR & "[field48:�������� - ˽��(Additional Data - Private)]" & vbCrLf & "" & "[" & temp_48_len_str & "]" & "  " & POS_Sturct.settleAccounts_48.ptr & vbCrLf & vbCrLf
        End If

        Dim temp_49str As String
        '//��49������ �����׻��Ҵ��롿
        If Mid(temp_bcd_flag_str, 49, 1) Then
            POS_Sturct.currency_code_49 = Mid(tempStr, tempCount, 6)
            POS_Sturct.currency_code_49 = ins_space(POS_Sturct.currency_code_49)
            temp_49str = ASCchange(POS_Sturct.currency_code_49)

            tex2STR = tex2STR & "[field49:���׻��Ҵ���(Currency Code Of Transaction)]" & vbCrLf & "" & POS_Sturct.currency_code_49 & vbCrLf & vbCrLf

            tempCount = tempCount + 6
        End If

        '//��52������ �����˱�ʶ�롿
        If Mid(temp_bcd_flag_str, 52, 1) Then
            POS_Sturct.pri_pin_52 = Mid(tempStr, tempCount, 16)
            POS_Sturct.pri_pin_52 = ins_space(POS_Sturct.pri_pin_52)
            tex2STR = tex2STR & "[field52:���˱�ʶ������(PIN Data)]" & vbCrLf & "" & POS_Sturct.pri_pin_52 & vbCrLf & vbCrLf

            tempCount = tempCount + 16
        End If

        '//��53������ ����ȫ������Ϣ��
        If Mid(temp_bcd_flag_str, 53, 1) Then
            POS_Sturct.safety_53 = Mid(tempStr, tempCount, 16)
            POS_Sturct.safety_53 = ins_space(POS_Sturct.safety_53)
            tex2STR = tex2STR & "[field53:��ȫ������Ϣ(Security Related Control Information )]" & vbCrLf & "" & POS_Sturct.safety_53 & vbCrLf & vbCrLf

            tempCount = tempCount + 16
        End If

        Dim temp_54str As String
        Dim temp_54_len_str As String
        Dim temp_consume_amount_54 As String
        '//��54������ ����
        If Mid(temp_bcd_flag_str, 54, 1) Then
            temp_54_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.attachment_amount_54.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.attachment_amount_54.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.attachment_amount_54.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen

            POS_Sturct.attachment_amount_54.ptr = ins_space(POS_Sturct.attachment_amount_54.ptr)
            temp_54_len_str = ins_space(temp_54_len_str)

            temp_54str = ASCchange(POS_Sturct.attachment_amount_54.ptr)
            temp_consume_amount_54 = Val(Mid(temp_54str, 9, 12)) / 100


            If Val(temp_consume_amount_54) < 1 And Val(temp_consume_amount_54) > 0 Then
                temp_consume_amount_54 = 0 & temp_consume_amount_54
            End If
            tex2STR = tex2STR & "[field54:���(Balanc Amount)]" & vbCrLf & "" & "[" & temp_54_len_str & "]" & "  " & POS_Sturct.attachment_amount_54.ptr & vbCrLf & "->[ASCIIת��] " _
                      & temp_54str & vbCrLf & "->[���]      " & temp_consume_amount_54 & "Ԫ" & vbCrLf & vbCrLf
        End If

        Dim temp_55_len_str As String
        Dim POS_Sturct_icData_55_temp_ptr As String       '//55����ʱ����
        Dim temp_55_str_ptr As Integer                     '//�����ʼλ��
        Dim temp_55_str_ptr_len As String                  '//���򳤶�
        '//��55������ ��IC��������
        If Mid(temp_bcd_flag_str, 55, 1) Then
            temp_55_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.icData_55.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.icData_55.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.icData_55.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            '            POS_Sturct.icData_55.ptr = ins_space(POS_Sturct.icData_55.ptr)
            '//���55��


            '//������Ϣ�����б�
            '�γ����ָ�ʽ ->[9F 26][Ӧ������]
            '               [08] 5E 14 AA 9F 20 46 A9 21
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F26")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[Ӧ������]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F27")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[������Ϣ����]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F10")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[������Ӧ������]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F37")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[����Ԥ֪��]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                      POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F36")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[Ӧ�ý��׼�����]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "95")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [�ն���֤���]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [��������]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9C")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [��������]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F02")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[��Ȩ���]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "5F2A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[���׻��Ҵ���]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "82")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [Ӧ�ý�������]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F1A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[�ն˹��Ҵ���]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F03")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[�������]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F33")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[�ն�����]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F74")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[�����ֽ𷢿�����Ȩ��]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If


            '//��ѡ��Ϣ�����б�

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F34")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[�ֿ�����֤�������]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F35")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[�ն�����]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F1E")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[�ӿ��豸���к�]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "84")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [ר���ļ�����]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F09")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[����汾��]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F41")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[�������м�����]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "91")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [��������֤����]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "71")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [�����нű� 1]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "72")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [�����нű� 2]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF31")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[�������ű����]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F63")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[����Ʒ��ʶ��Ϣ]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            '//�ѻ�����ר�������б�

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "8A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [��Ȩ��Ӧ��]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            '// �ֻ�оƬ����ר�������б�

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF32")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[оƬ���к�]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF33")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[������Կ����]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF34")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// ��Ϊ�գ���Ϊ��������ֹ�ҵ������������ ���� 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[�ŵ���ȡʱ��]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//ȥ������������Դ
            End If



            POS_Sturct.icData_55.ptr = POS_Sturct_icData_55_temp_ptr
            temp_55_len_str = ins_space(temp_55_len_str)
            tex2STR = tex2STR & "[field55:IC��������(IC Card System Related Data)]" & vbCrLf & "" & "[" & temp_55_len_str & "]" & vbCrLf & POS_Sturct.icData_55.ptr & vbCrLf

        End If

        Dim temp_56_len_str As String
        '//��56������ ��private_data_56��
        If Mid(temp_bcd_flag_str, 56, 1) Then
            temp_56_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_56.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_56.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_56.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_56.ptr = ins_space(POS_Sturct.private_data_56.ptr)
            temp_56_len_str = ins_space(temp_56_len_str)
            tex2STR = tex2STR & "[field56:�Զ�����(private_data_56)]" & vbCrLf & "" & "[" & temp_56_len_str & "]" & "  " & POS_Sturct.private_data_56.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_57_len_str As String
        '//��57������ ��private_data_57��
        If Mid(temp_bcd_flag_str, 57, 1) Then
            temp_57_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_57.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_57.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_57.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_57.ptr = ins_space(POS_Sturct.private_data_57.ptr)
            temp_57_len_str = ins_space(temp_57_len_str)
            tex2STR = tex2STR & "[field57:�Զ�����(private_data_57)]" & vbCrLf & "" & "[" & temp_57_len_str & "]" & "  " & POS_Sturct.private_data_57.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_58_len_str As String
        '//��58������ ��PBOC����Ǯ����׼�Ľ�����Ϣ��
        If Mid(temp_bcd_flag_str, 58, 1) Then
            temp_58_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_58.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_58.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_58.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_58.ptr = ins_space(POS_Sturct.private_data_58.ptr)
            temp_58_len_str = ins_space(temp_58_len_str)
            tex2STR = tex2STR & "[field58:PBOC����Ǯ����׼�Ľ�����Ϣ(PBOC_ELECTRONIC_DATA)]" & vbCrLf & "" & "[" & temp_58_len_str & "]" & "  " & POS_Sturct.private_data_58.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_59_len_str As String
        '//��59������ ��private_data_59��
        If Mid(temp_bcd_flag_str, 59, 1) Then
            temp_59_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_59.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_59.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_59.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_59.ptr = ins_space(POS_Sturct.private_data_59.ptr)
            temp_59_len_str = ins_space(temp_59_len_str)
            tex2STR = tex2STR & "[field59:�Զ�����(private_data_59)]" & vbCrLf & "" & "[" & temp_59_len_str & "]" & "  " & POS_Sturct.private_data_59.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_60_len_str As String
        Dim temp_batch_Number_60 As String
        '//��60������ ��private_data_60��
        If Mid(temp_bcd_flag_str, 60, 1) Then
            temp_60_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_60.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.private_data_60.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_60.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen


            temp_batch_Number_60 = Mid(POS_Sturct.private_data_60.ptr, 3, 6)
            POS_Sturct.private_data_60.ptr = ins_space(POS_Sturct.private_data_60.ptr)

            temp_60_len_str = ins_space(temp_60_len_str)
            tex2STR = tex2STR & "[field60:�Զ�����(private_data_60)]" & vbCrLf & "" & "[" & temp_60_len_str & "]" & "  " & POS_Sturct.private_data_60.ptr _
                      & vbCrLf & "->[���κ�] " & temp_batch_Number_60 & vbCrLf & vbCrLf

            batch_Number_60.Caption = temp_batch_Number_60

        End If

        Dim temp_61_len_str As String
        Dim Original_batch_Number_61 As String    '//ԭʼ�������κ�
        Dim Original_trace_no_61 As String    '//ԭʼ����POS��ˮ��

        '//��61������ ��ԭʼ��Ϣ��
        If Mid(temp_bcd_flag_str, 61, 1) Then
            temp_61_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_61.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.private_data_61.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_61.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            Original_batch_Number_61 = Mid(POS_Sturct.private_data_61.ptr, 1, 6)
            Original_trace_no_61 = Mid(POS_Sturct.private_data_61.ptr, 7, 6)

            POS_Sturct.private_data_61.ptr = ins_space(POS_Sturct.private_data_61.ptr)

            temp_61_len_str = ins_space(temp_61_len_str)
            tex2STR = tex2STR & "[field61:ԭʼ��Ϣ��(Original Message)]" & vbCrLf & "" & "[" & temp_61_len_str & "]" & "  " & POS_Sturct.private_data_61.ptr _
                      & vbCrLf & "->[ԭʼ�������κ�]    " & Original_batch_Number_61 & vbCrLf & "->[ԭʼ����POS��ˮ��] " & Original_trace_no_61 & vbCrLf & vbCrLf

        End If

        Dim temp_62_len_str As String
        '//��62������ ��private_data_62��
        If Mid(temp_bcd_flag_str, 62, 1) Then
            temp_62_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_62.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_62.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_62.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_62.ptr = ins_space(POS_Sturct.private_data_62.ptr)
            temp_62_len_str = ins_space(temp_62_len_str)
            tex2STR = tex2STR & "[field62:�Զ�����(private_data_62)]" & vbCrLf & "" & "[" & temp_62_len_str & "]" & "  " & POS_Sturct.private_data_62.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_63_len_str As String
        '//��63������ ��private_data_63��
        If Mid(temp_bcd_flag_str, 63, 1) Then
            temp_63_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_63.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_63.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_63.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_63.ptr = ins_space(POS_Sturct.private_data_63.ptr)
            temp_63_len_str = ins_space(temp_63_len_str)
            tex2STR = tex2STR & "[field63:�Զ�����(private_data_63)]" & vbCrLf & "" & "[" & temp_63_len_str & "]" & "  " & POS_Sturct.private_data_63.ptr & vbCrLf & vbCrLf

        End If

        '//��64������ �����ļ����롿
        If Mid(temp_bcd_flag_str, 64, 1) Then
            POS_Sturct.mac_64 = Mid(tempStr, tempCount, 16)
            POS_Sturct.mac_64 = ins_space(POS_Sturct.mac_64)
            tex2STR = tex2STR & "[field64:���ļ�����(Message Authentication Code)]" & vbCrLf & "" & POS_Sturct.mac_64 & vbCrLf & vbCrLf
            tempCount = tempCount + 16
        End If



        '//  tex2STR = tex2STR & "************8583������������ǻ����ķָ��߽�β��***************/" & vbCrLf
        analyse_after_data.Text = tex2STR
    End If
    Exit Sub  'һ��Ҫд
ErrHandle:
    'MSG_BOX Err.Number & Err.Source, "����"
    ' Err.clear

    analyse_after_data.Text = ""
    bit_map_clear_Click
    trans_type.Caption = "N/A"
    judge_mode.Caption = "N/A"
    judge_mode.ForeColor = &H0&    '//�ָ��ɺ�ɫ

    Response_code.Caption = "N/A"
    Response_code_view.Caption = "N/A"

    trace_no_11.Caption = "N/A"
    batch_Number_60.Caption = "N/A"

    help_flag = 0
    help.Caption = "˵��"
    analyse_before_data.SetFocus


    MSG_BOX _
            "[Err.Source      ]:" & Err.Source & vbCrLf & _
                                  "[Err.Number      ]:" & Err.Number & vbCrLf & _
                                  "[Err.Description ]:" & Err.Description & vbCrLf, "����"

    '"[Err.HelpContext ]:" & Err.HelpContext & vbCrLf & _
     '"[Err.HelpFile    ]:" & Err.HelpFile & vbCrLf & _
     '"[Err.LastDllError]:" & Err.LastDllError & vbCrLf & _

     Err.clear

End Sub

'//������
Private Sub analyse_level_3_Click()

    Dim tex2STR, tempStr, Totallen, TPDU, MessageHeader As String
    Dim tempCount As Long
    Dim POS_Sturct As POS_Sturct_TYPE
    Dim i As Integer

    PROCESS_Analyze_data    '//����ǰ����

    If change_begin = 1 Then
        analyse_after_data.SetFocus
        analyse_mode = 3    '//�������ģʽ 3 ȫ�����
        trace_no_11.Caption = "N/A"
        batch_Number_60.Caption = "N/A"
        help_flag = 0
        help_count = 0
        help.Caption = "˵��"
        tempCount = 1    '//����


        tempStr = analyse_after_data.Text
        '//       tex2STR = tex2STR & "/************8583������������ǻ����ķָ��߿�ͷ��***************" & vbCrLf & vbCrLf

        If Totallen_check.Value = 1 Then
            '//�ܳ�����
            Dim temp_TotallenStr As String

            Totallen = Mid(tempStr, tempCount, 4)
            temp_TotallenStr = HEX_to_DEC(Mid(Totallen, 1, 2)) * 16 * 16 + HEX_to_DEC(Mid(Totallen, 3, 4))

            Totallen = ins_space(Totallen)
            tex2STR = tex2STR & "[Totallen:�ܳ���]" & vbCrLf & "" & Totallen & "         " & vbCrLf & _
                      "->[ʮ�����ܳ���]" & " " & temp_TotallenStr & "+2=�� " & temp_TotallenStr + 2 & " �ֽ�" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        If TPDU_check.Value = 1 Then
            '//TPDU����
            Dim TPDU_ID, TPDU_DEST_ADR, TPDU_SRC_ADR As String


            TPDU = Mid(tempStr, tempCount, 10)
            TPDU_ID = ins_space(Mid(TPDU, 1, 2))
            TPDU_DEST_ADR = ins_space(Mid(TPDU, 3, 4))
            TPDU_SRC_ADR = ins_space(Mid(TPDU, 7, 4))


            TPDU = ins_space(TPDU)
            tex2STR = tex2STR & "[TPDU:��ַ]" & vbCrLf & TPDU & vbCrLf _
                      & "------------" & vbCrLf & "->[ID]" & vbCrLf & TPDU_ID & vbCrLf & "------------" & vbCrLf & "->[Ŀ�ĵ�ַ]" & vbCrLf & TPDU_DEST_ADR & vbCrLf & "------------" & vbCrLf & "->[Դ��ַ]" & vbCrLf & TPDU_SRC_ADR & vbCrLf & "------------" & vbCrLf & vbCrLf
            tempCount = tempCount + 10
        End If

        If MessageHeader_check.Value = 1 Then

            Dim MessageHeader_App_Type As String    '// Ӧ�������
            Dim MessageHeader_App_Type_Str As String    '// Ӧ���������ʾ

            Dim MessageHeader_Software_Total_Ver_Num As String    '// ����ܰ汾��
            Dim MessageHeader_Software_Total_Ver_Num_Str As String    '// ����ܰ汾����ʾ

            Dim MessageHeader_Terminal_State As String   '// �ն�״̬
            Dim MessageHeader_Terminal_State_Str As String   '// �ն�״̬��ʾ

            Dim MessageHeader_Process_Require As String    '// ����Ҫ��
            Dim MessageHeader_Process_Require_Str As String   '// �ն�״̬��ʾ

            Dim MessageHeader_Software_Part_Ver_Num As String    '// ����ְ汾��
            Dim MessageHeader_Software_Part_Ver_Num_Str As String    '// ����ְ汾����ʾ


            '//MessageHeader����
            MessageHeader = Mid(tempStr, tempCount, 12)

            '//Ӧ����������
            MessageHeader_App_Type = Mid(MessageHeader, 1, 2)

            Select Case MessageHeader_App_Type
            Case 60
                MessageHeader_App_Type_Str = "����������֧����Ӧ��"
            Case 61
                MessageHeader_App_Type_Str = "IC������֧����Ӧ��"
            Case 62
                MessageHeader_App_Type_Str = "��������ֵҵ����֧��"
            Case 63
                MessageHeader_App_Type_Str = "IC����ֵҵ����֧��"
            Case Else
                MessageHeader_App_Type_Str = "N/A"
            End Select

            '//����ܰ汾�Ž���
            MessageHeader_Software_Total_Ver_Num = Mid(MessageHeader, 3, 2)

            Select Case MessageHeader_Software_Total_Ver_Num
            Case 10
                MessageHeader_Software_Total_Ver_Num_Str = "2001����������POS�淶֮ǰ�汾"
            Case 11
                MessageHeader_Software_Total_Ver_Num_Str = "2001����������POS�淶�汾"
            Case 21
                MessageHeader_Software_Total_Ver_Num_Str = "2002������POS�淶�汾"
            Case 22
                MessageHeader_Software_Total_Ver_Num_Str = "2004������POS�淶�汾"
            Case 30
                MessageHeader_Software_Total_Ver_Num_Str = "2009������POS�淶�汾"
            Case 31
                MessageHeader_Software_Total_Ver_Num_Str = "2010������POS�淶�汾"
            Case Else
                MessageHeader_Software_Total_Ver_Num_Str = "N/A"
            End Select

            '//�ն�״̬����
            MessageHeader_Terminal_State = Mid(MessageHeader, 5, 1)

            Select Case MessageHeader_Terminal_State
            Case 0
                MessageHeader_Terminal_State_Str = "��������״̬"
            Case Else
                MessageHeader_Terminal_State_Str = "N/A"
            End Select

            '//����Ҫ�����
            MessageHeader_Process_Require = Mid(MessageHeader, 6, 1)

            Select Case MessageHeader_Process_Require
            Case 0
                MessageHeader_Process_Require_Str = "�޴���Ҫ��"
            Case 1
                MessageHeader_Process_Require_Str = "�´��ն˴���������"
            Case 2
                MessageHeader_Process_Require_Str = "�ϴ��ն˴�����״̬��Ϣ"
            Case 3
                MessageHeader_Process_Require_Str = "����ǩ��"
            Case 4
                MessageHeader_Process_Require_Str = "֪ͨ�ն˷�����¹�Կ��Ϣ����"
            Case 5
                MessageHeader_Process_Require_Str = "�����ն�IC������"
            Case 6
                MessageHeader_Process_Require_Str = "TMS��������"
            Case 7
                MessageHeader_Process_Require_Str = "��BIN ����������"
            Case 8
                MessageHeader_Process_Require_Str = "���ֻ������أ����ھ���ʹ�ã�/��Ūȡ�������ѱ������أ����ھ���ʹ�ã�"

            Case Else
                MessageHeader_Process_Require_Str = "N/A"
            End Select

            '// ����ְ汾��
            MessageHeader_Software_Part_Ver_Num = ins_space(Mid(MessageHeader, 7, 6))
            MessageHeader_Software_Part_Ver_Num_Str = "ǰ���ֽ�ͬ����汾�ţ������ֽ��ɳ������ж���"

            MessageHeader = ins_space(MessageHeader)
            tex2STR = tex2STR & "[MessageHeader:����ͷ]" & vbCrLf & "" & MessageHeader & vbCrLf _
                      & "----------------" & vbCrLf & "->[Ӧ�������]" & vbCrLf & MessageHeader_App_Type & vbCrLf & MessageHeader_App_Type_Str & vbCrLf _
                      & "----------------" & vbCrLf & "->[����ܰ汾��]" & vbCrLf & MessageHeader_Software_Total_Ver_Num & vbCrLf & MessageHeader_Software_Total_Ver_Num_Str & vbCrLf _
                      & "----------------" & vbCrLf & "->[�ն�״̬]" & vbCrLf & MessageHeader_Terminal_State & vbCrLf & MessageHeader_Terminal_State_Str & vbCrLf _
                      & "----------------" & vbCrLf & "->[����Ҫ��]" & vbCrLf & MessageHeader_Process_Require & vbCrLf & MessageHeader_Process_Require_Str & vbCrLf _
                      & "----------------" & vbCrLf & "->[����ְ汾��]" & vbCrLf & MessageHeader_Software_Part_Ver_Num & vbCrLf & MessageHeader_Software_Part_Ver_Num_Str & vbCrLf & "----------------" & vbCrLf & vbCrLf
            tempCount = tempCount + 12
        End If



        Dim temp_messagetype_str

        '//messagetype����
        POS_Sturct.messagetype = Mid(tempStr, tempCount, 4)

        Select Case POS_Sturct.messagetype
        Case "0100"
            temp_messagetype_str = "���� 0100 ��Ȩ��������Ϣ��" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ����" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��������" & vbCrLf & _
                                   "-> �������ֽ��ֵ�˻���֤����" & vbCrLf
        Case "0110"
            temp_messagetype_str = "���� 0110 ��Ȩ��Ӧ����Ϣ��" & vbCrLf & _
                                   "-> POS Ԥ��ȨӦ��" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ����Ӧ��" & vbCrLf & _
                                   "-> �������ֽ��ֵ�˻���֤Ӧ��" & vbCrLf
        Case "0200"
            temp_messagetype_str = "���� 0200 ������������Ϣ��" & vbCrLf & _
                                   "-> POS ��ѯ����" & vbCrLf & _
                                   "-> POS ��������" & vbCrLf & _
                                   "-> POS ���ѳ�������" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��ɣ���������" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��ɳ�������" & vbCrLf & _
                                   "-> �����ֽ��ѻ���������" & vbCrLf & _
                                   "-> ���ڸ�����������" & vbCrLf & _
                                   "-> ���ڸ������ѳ�������" & vbCrLf & _
                                   "-> ���� PBOC ����Ǯ��/�����ֽ�� IC Ȧ���ཻ������" & vbCrLf & _
                                   "-> �������ֽ��ֵ����" & vbCrLf & _
                                   "-> �������ʻ���ֵ����" & vbCrLf
        Case "0210"
            temp_messagetype_str = "���� 0210 ������Ӧ����Ϣ��" & vbCrLf & _
                                   "-> POS ��ѯӦ��" & vbCrLf & _
                                   "-> POS ����Ӧ��" & vbCrLf & _
                                   "-> POS ���ѳ���Ӧ��" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��ɣ�����Ӧ��" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��ɳ���Ӧ��" & vbCrLf & _
                                   "-> �����ֽ��ѻ�����Ӧ��" & vbCrLf & _
                                   "-> ���ڸ�������Ӧ��" & vbCrLf & _
                                   "-> ���ڸ������ѳ���Ӧ��" & vbCrLf & _
                                   "-> ���� PBOC ����Ǯ��/�����ֽ�� IC Ȧ���ཻ��Ӧ��" & vbCrLf & _
                                   "-> �������ֽ��ֵӦ��" & vbCrLf & _
                                   "-> �������ʻ���ֵӦ��" & vbCrLf
        Case "0220"
            temp_messagetype_str = "���� 0220 ����֪ͨ����Ϣ��" & vbCrLf & _
                                   "-> POS �˻�֪ͨ" & vbCrLf & _
                                   "-> POS ���߽���֪ͨ" & vbCrLf & _
                                   "-> POS �������֪ͨ" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��ɣ�֪ͨ��֪ͨ" & vbCrLf & _
                                   "-> �������ֽ��ֵȷ��֪ͨ" & vbCrLf
        Case "0230"
            temp_messagetype_str = "���� 0230 ����֪ͨ��Ӧ����Ϣ��" & vbCrLf & _
                                   "-> POS �˻�Ӧ��" & vbCrLf & _
                                   "-> POS ���߽���Ӧ��" & vbCrLf & _
                                   "-> POS �������Ӧ��" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��ɣ�֪ͨ��Ӧ��" & vbCrLf & _
                                   "-> �������ֽ��ֵȷ��Ӧ��" & vbCrLf
        Case "0320"
            temp_messagetype_str = "���� 0320 ��������Ϣ��" & vbCrLf & _
                                   "-> POS �ն�������" & vbCrLf
        Case "0330"
            temp_messagetype_str = "���� 0330 ������Ӧ����Ϣ��" & vbCrLf & _
                                   "-> POS �ն�������Ӧ��" & vbCrLf

        Case "0400"
            temp_messagetype_str = "���� 0400 ��������Ϣ��" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ����" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��������" & vbCrLf & _
                                   "-> POS ���ѳ���" & vbCrLf & _
                                   "-> POS ���ѳ�������" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��ɣ����󣩳���" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��ɳ�������" & vbCrLf & _
                                   "-> ���� PBOC ����Ǯ��/�����ֽ�� IC Ȧ���ཻ�׳���" & vbCrLf
        Case "0410"
            temp_messagetype_str = "���� 0410 ������Ӧ����Ϣ��" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ����Ӧ��" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��������Ӧ��" & vbCrLf & _
                                   "-> POS ���ѳ���Ӧ��" & vbCrLf & _
                                   "-> POS ���ѳ�������Ӧ��" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��ɣ����󣩳���Ӧ��" & vbCrLf & _
                                   "-> POS Ԥ��Ȩ��ɳ�������Ӧ��" & vbCrLf & _
                                   "-> ���� PBOC ����Ǯ��/�����ֽ�� IC Ȧ���ཻ�׳���Ӧ��" & vbCrLf

        Case "0500"
            temp_messagetype_str = "���� 0500 ��������Ϣ��" & vbCrLf & _
                                   "-> POS �ն�����������" & vbCrLf
        Case "0510"
            temp_messagetype_str = "���� 0510 ������Ӧ����Ϣ��" & vbCrLf & _
                                   "-> POS �ն�������Ӧ��" & vbCrLf

        Case "0620"
            temp_messagetype_str = "���� 0620 ���� PBOC ��/���ǿ���׼�� IC ���ű�������֪ͨ��Ϣ��" & vbCrLf & _
                                   "-> ���� PBOC ��/���ǿ���׼�� IC ���ű�������֪ͨ" & vbCrLf
        Case "0630"
            temp_messagetype_str = "���� 0630 ���� PBOC ��/���ǿ���׼�� IC ���ű�������֪ͨӦ��" & vbCrLf & _
                                   "-> ���� PBOC ��/���ǿ���׼�� IC ���ű�������֪ͨӦ��" & vbCrLf

        Case "0800"
            temp_messagetype_str = "���� 0800 ����ҵ���������Ϣ��" & vbCrLf & _
                                   "-> POS �ն�ǩ������" & vbCrLf & _
                                   "-> POS �ն˲�����������" & vbCrLf
        Case "0810"
            temp_messagetype_str = "���� 0810 ����ҵ�������Ӧ����Ϣ��" & vbCrLf & _
                                   "-> POS �ն�ǩ��Ӧ��" & vbCrLf & _
                                   "-> POS �ն˲�������Ӧ��" & vbCrLf

        Case "0820"
            temp_messagetype_str = "���� 0820 ����ҵ���������Ϣ��" & vbCrLf & _
                                   "-> POS �ն�ǩ������" & vbCrLf & _
                                   "-> POS �ն˻����������" & vbCrLf & _
                                   "-> POS �ն�״̬����" & vbCrLf
        Case "0830"
            temp_messagetype_str = "���� 0830 ����ҵ�������Ӧ����Ϣ��" & vbCrLf & _
                                   "-> POS �ն�ǩ��Ӧ��" & vbCrLf & _
                                   "-> POS �ն˻������Ӧ��" & vbCrLf & _
                                   "-> POS �ն�״̬����Ӧ��" & vbCrLf


        Case Else
            temp_messagetype_str = "N/A"
        End Select



        POS_Sturct.messagetype = ins_space(POS_Sturct.messagetype)
        tex2STR = tex2STR & "[messagetype:��Ϣ����]" & vbCrLf & "" & POS_Sturct.messagetype & vbCrLf & temp_messagetype_str & vbCrLf
        tempCount = tempCount + 4


        Dim temp_bcd_flag_display_str As String
        Dim temp_bcd_flag_str As String
        Dim temp_bcd_flag_str_count As Integer
        Dim tempLen As Integer

        '//bitmap����
        POS_Sturct.bitmap = Mid(tempStr, tempCount, 16)
        POS_Sturct.bitmap = ins_space(POS_Sturct.bitmap)

        '//bitmap����
        temp_bcd_flag_str = HEX_to_BIN(delete_space(POS_Sturct.bitmap))


        '//bitmap������ʾ����
        For temp_bcd_flag_str_count = 1 To Len(temp_bcd_flag_str)
            If Mid(temp_bcd_flag_str, temp_bcd_flag_str_count, 1) = 1 Then
                If temp_bcd_flag_str_count <> Len(temp_bcd_flag_str) Then
                    temp_bcd_flag_display_str = temp_bcd_flag_display_str & temp_bcd_flag_str_count & "�� "
                Else
                    temp_bcd_flag_display_str = temp_bcd_flag_display_str & temp_bcd_flag_str_count & "��"
                End If
            End If
        Next temp_bcd_flag_str_count


        tex2STR = tex2STR & "[bitmap:λԪ��]" & vbCrLf & "" & POS_Sturct.bitmap & vbCrLf & "->[����������] " & temp_bcd_flag_display_str & vbCrLf & vbCrLf
        tempCount = tempCount + 16
        bit_map.Text = POS_Sturct.bitmap
        bit_map_change_function





        Dim tempPan As String
        Dim temp_Pan_len As String
        '//��2������ �����˺š�
        If Mid(temp_bcd_flag_str, 2, 1) Then
            temp_Pan_len = Mid(tempStr, tempCount, 2)
            POS_Sturct.pan_2.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.pan_2.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.pan_2.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            tempPan = Mid(POS_Sturct.pan_2.ptr, 1, POS_Sturct.pan_2.Ptrlen)
            POS_Sturct.pan_2.ptr = ins_space(POS_Sturct.pan_2.ptr)
            tex2STR = tex2STR & "[field2:���˺�(Primary Account Number)]" & vbCrLf & "" & "[" & temp_Pan_len & "]" & " " & POS_Sturct.pan_2.ptr _
                      & vbCrLf & "->[����] " & tempPan & vbCrLf & vbCrLf
        End If

        '//��3������ �����״����롿
        If Mid(temp_bcd_flag_str, 3, 1) Then
            POS_Sturct.procode_3 = Mid(tempStr, tempCount, 6)
            POS_Sturct.procode_3 = ins_space(POS_Sturct.procode_3)
            tex2STR = tex2STR & "[field3:���״�����(Transaction Processing Code)]" & vbCrLf & "" & POS_Sturct.procode_3 & vbCrLf & vbCrLf
            tempCount = tempCount + 6
        End If
        '//�����ж�
        i = Type_judge(POS_Sturct.messagetype, Mid(POS_Sturct.procode_3, 1, 2), Mid(temp_bcd_flag_str, 61, 1))

        Dim temp_consume_amount_4 As String
        '//��4������ �����׽�
        If Mid(temp_bcd_flag_str, 4, 1) Then
            POS_Sturct.consume_amount_4 = Mid(tempStr, tempCount, 12)
            temp_consume_amount_4 = Val(POS_Sturct.consume_amount_4) / 100

            If Val(temp_consume_amount_4) < 1 And Val(temp_consume_amount_4) > 0 Then
                temp_consume_amount_4 = 0 & temp_consume_amount_4
            End If


            POS_Sturct.consume_amount_4 = ins_space(POS_Sturct.consume_amount_4)
            tex2STR = tex2STR & "[field4:���׽��(Amount Of Transactions)]" & vbCrLf & "" & POS_Sturct.consume_amount_4 & _
                      vbCrLf & "->[�����] " & temp_consume_amount_4 & "Ԫ" & vbCrLf & vbCrLf
            tempCount = tempCount + 12
        End If

        Dim temp_trace_no_11 As String
        '//��11������ ���ܿ���ϵͳ���ٺš�
        If Mid(temp_bcd_flag_str, 11, 1) Then
            POS_Sturct.trace_no_11 = Mid(tempStr, tempCount, 6)
            temp_trace_no_11 = POS_Sturct.trace_no_11
            POS_Sturct.trace_no_11 = ins_space(POS_Sturct.trace_no_11)
            tex2STR = tex2STR & "[field11:�ܿ���ϵͳ���ٺ�(System Trace Audit Number)]" & vbCrLf & "" & POS_Sturct.trace_no_11 & _
                      vbCrLf & "->[��ˮ��:] " & temp_trace_no_11 & vbCrLf & vbCrLf

            trace_no_11.Caption = temp_trace_no_11

            tempCount = tempCount + 6
        End If


        '//��12������ ���ܿ������ڵ�ʱ�䡿
        If Mid(temp_bcd_flag_str, 12, 1) Then
            POS_Sturct.trade_time_12 = Mid(tempStr, tempCount, 6)
            POS_Sturct.trade_time_12 = ins_space(POS_Sturct.trade_time_12)
            tex2STR = tex2STR & "[field12:�ܿ������ڵ�ʱ��(Local Time Of Transaction)]" & vbCrLf & "" & POS_Sturct.trade_time_12 & _
                      vbCrLf & "->[ʱ��] " & Mid(POS_Sturct.trade_time_12, 1, 2) & "ʱ" & Mid(POS_Sturct.trade_time_12, 4, 2) & "��" _
                      & Mid(POS_Sturct.trade_time_12, 7, 2) & "��" & vbCrLf & vbCrLf
            tempCount = tempCount + 6
        End If

        '//��13������ ���ܿ������ڵ����ڡ�
        If Mid(temp_bcd_flag_str, 13, 1) Then
            POS_Sturct.trade_date_13 = Mid(tempStr, tempCount, 4)
            POS_Sturct.trade_date_13 = ins_space(POS_Sturct.trade_date_13)
            tex2STR = tex2STR & "[field13:�ܿ������ڵ�����(Local Date Of Transaction)]" & vbCrLf & "" & POS_Sturct.trade_date_13 & "   " & _
                      vbCrLf & "->[����] " & Mid(POS_Sturct.trade_date_13, 1, 2) & "��" & Mid(POS_Sturct.trade_date_13, 4, 2) & "��" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��14������ ������Ч�ڡ�
        If Mid(temp_bcd_flag_str, 14, 1) Then
            POS_Sturct.exp_date_14 = Mid(tempStr, tempCount, 4)
            POS_Sturct.exp_date_14 = ins_space(POS_Sturct.exp_date_14)
            tex2STR = tex2STR & "[field14:����Ч��(Date Of Expired)]" & vbCrLf & "" & POS_Sturct.exp_date_14 & "   " & _
                      vbCrLf & "->[��Ч��] " & "20" & Mid(POS_Sturct.exp_date_14, 1, 2) & "��" & Mid(POS_Sturct.exp_date_14, 4, 2) & "��" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��15������ ���������ڡ�
        If Mid(temp_bcd_flag_str, 15, 1) Then
            POS_Sturct.settlement_date_15 = Mid(tempStr, tempCount, 4)
            POS_Sturct.settlement_date_15 = ins_space(POS_Sturct.settlement_date_15)
            tex2STR = tex2STR & "[field15:��������(Date Of Settlement)]" & vbCrLf & "" & POS_Sturct.settlement_date_15 & "   " & _
                      vbCrLf & "->[��������] " & Mid(POS_Sturct.settlement_date_15, 1, 2) & "��" & Mid(POS_Sturct.settlement_date_15, 4, 2) & "��" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��22������ ����������뷽ʽ�롿
        If Mid(temp_bcd_flag_str, 22, 1) Then
            POS_Sturct.entry_mode_22 = Mid(tempStr, tempCount, 4)
            POS_Sturct.entry_mode_22 = ins_space(POS_Sturct.entry_mode_22)
            tex2STR = tex2STR & "[field22:��������뷽ʽ��(Point Of Service Entry Mode)]" & vbCrLf & "" & POS_Sturct.entry_mode_22 & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If




        Dim Right_23_str As String    '//��Ƭ���к��ҿ�
        Dim left_23_str As String    '//��Ƭ���к���
        '//��23������ �������кš�
        If Mid(temp_bcd_flag_str, 23, 1) Then
            POS_Sturct.card_serial_number_23 = Mid(tempStr, tempCount, 4)
            Right_23_str = Mid(POS_Sturct.card_serial_number_23, 2, 3)
            left_23_str = Mid(POS_Sturct.card_serial_number_23, 1, 3)
            POS_Sturct.card_serial_number_23 = ins_space(POS_Sturct.card_serial_number_23)
            tex2STR = tex2STR & "[field23:�����к�(Card Sequence Number)]" & vbCrLf & "" & POS_Sturct.card_serial_number_23 & vbCrLf _
                      & "->[��Ƭ���к���] " & left_23_str & vbCrLf & "->[��Ƭ���к��ҿ�] " & Right_23_str & " [һ������ѡ���ҿ�]" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��25������ ������������롿
        If Mid(temp_bcd_flag_str, 25, 1) Then
            POS_Sturct.service_conditon_25 = Mid(tempStr, tempCount, 2)
            POS_Sturct.service_conditon_25 = ins_space(POS_Sturct.service_conditon_25)
            tex2STR = tex2STR & "[field25:�����������(Point Of Service Condition Mode)]" & vbCrLf & "" & POS_Sturct.service_conditon_25 & vbCrLf & vbCrLf
            tempCount = tempCount + 2
        End If

        '//��26������ �������PIN��ȡ�롿
        If Mid(temp_bcd_flag_str, 26, 1) Then
            POS_Sturct.service_conditon_pin_26 = Mid(tempStr, tempCount, 2)
            POS_Sturct.service_conditon_pin_26 = ins_space(POS_Sturct.service_conditon_pin_26)
            tex2STR = tex2STR & "[field26:�����PIN��ȡ��(Point Of Service PIN Capture Code)]" & vbCrLf & "" & POS_Sturct.service_conditon_pin_26 & vbCrLf & vbCrLf
            tempCount = tempCount + 2
        End If

        Dim temp_32_len_str As String
        '//��32������ ��������ʶ�롿
        If Mid(temp_bcd_flag_str, 32, 1) Then
            temp_32_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.api_code_32.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.api_code_32.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.api_code_32.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.api_code_32.ptr = ins_space(POS_Sturct.api_code_32.ptr)

            tex2STR = tex2STR & "[field32:������ʶ��(Acquiring Institution Id Code)]" & vbCrLf & "" & "[" & temp_32_len_str & "]" & " " & POS_Sturct.api_code_32.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_35_len_str As String
        '//��35������ ��2�ŵ����ݡ�
        If Mid(temp_bcd_flag_str, 35, 1) Then
            temp_35_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.track2_35.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.track2_35.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.track2_35.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.track2_35.ptr = ins_space(POS_Sturct.track2_35.ptr)

            tex2STR = tex2STR & "[field35:2�ŵ�����(Track 2 Data)]" & vbCrLf & "" & "[" & temp_35_len_str & "]" & "  " & POS_Sturct.track2_35.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_36_len_str As String
        '//��36������ ��3�ŵ����ݡ�
        If Mid(temp_bcd_flag_str, 36, 1) Then
            temp_36_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.track3_36.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.track3_36.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.track3_36.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.track3_36.ptr = ins_space(POS_Sturct.track3_36.ptr)
            temp_36_len_str = ins_space(temp_36_len_str)
            tex2STR = tex2STR & "[field36:3�ŵ�����(Track 3 Data)]" & vbCrLf & "" & "[" & temp_36_len_str & "]" & "  " & POS_Sturct.track3_36.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_37str As String
        '//��37������ �������ο��š�
        If Mid(temp_bcd_flag_str, 37, 1) Then
            POS_Sturct.reference_number_37 = Mid(tempStr, tempCount, 24)
            POS_Sturct.reference_number_37 = ins_space(POS_Sturct.reference_number_37)
            temp_37str = ASCchange(POS_Sturct.reference_number_37)
            tex2STR = tex2STR & "[field37:�����ο���(Retrieval Reference Number)]" & vbCrLf & "" & POS_Sturct.reference_number_37 & vbCrLf & "->[�ο���] " & temp_37str & vbCrLf & vbCrLf
            tempCount = tempCount + 24
        End If

        Dim temp_38str As String
        '//��38������ ����Ȩ��ʶӦ���롿
        If Mid(temp_bcd_flag_str, 38, 1) Then
            POS_Sturct.authorization_code_38 = Mid(tempStr, tempCount, 12)
            POS_Sturct.authorization_code_38 = ins_space(POS_Sturct.authorization_code_38)
            temp_38str = ASCchange(POS_Sturct.authorization_code_38)
            tex2STR = tex2STR & "[field38:��Ȩ��ʶӦ����(Authorization Id Response Code)]" & vbCrLf & "" & POS_Sturct.authorization_code_38 & vbCrLf & "->[��Ȩ��] " & temp_38str & vbCrLf & vbCrLf
            tempCount = tempCount + 12
        End If

        Dim temp_39str As String
        '//��39������ ��Ӧ���롿
        If Mid(temp_bcd_flag_str, 39, 1) Then
            POS_Sturct.Response_code_39 = Mid(tempStr, tempCount, 4)
            POS_Sturct.Response_code_39 = ins_space(POS_Sturct.Response_code_39)
            temp_39str = ASCchange(POS_Sturct.Response_code_39)
            tex2STR = tex2STR & "[field39:Ӧ����(Response Code)]" & vbCrLf & "" & POS_Sturct.Response_code_39 & vbCrLf & "->[Ӧ����:] " & temp_39str & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//��Ӧ���ж���ʾ
        i = Response_code_39_Type_judge(ASCchange(POS_Sturct.Response_code_39))

        Dim temp_41str As String
        '//��41������ ���ܿ����ն˱�ʶ�롿
        If Mid(temp_bcd_flag_str, 41, 1) Then
            POS_Sturct.terminal_no_41 = Mid(tempStr, tempCount, 16)
            POS_Sturct.terminal_no_41 = ins_space(POS_Sturct.terminal_no_41)
            temp_41str = ASCchange(POS_Sturct.terminal_no_41)
            tex2STR = tex2STR & "[field41:�ܿ����ն˱�ʶ��(Card Acceptor Terminal Id)]" & vbCrLf & "" & POS_Sturct.terminal_no_41 & vbCrLf & "->[�ն˺�] " & temp_41str & vbCrLf & vbCrLf
            tempCount = tempCount + 16
        End If

        Dim temp_42str As String
        '//��42������ ���ܿ�����ʶ�롿
        If Mid(temp_bcd_flag_str, 42, 1) Then
            POS_Sturct.merchant_no_42 = Mid(tempStr, tempCount, 30)
            POS_Sturct.merchant_no_42 = ins_space(POS_Sturct.merchant_no_42)
            temp_42str = ASCchange(POS_Sturct.merchant_no_42)
            tex2STR = tex2STR & "[field42:�ܿ�����ʶ��(Card Acceptor Id Code)]" & vbCrLf & "" & POS_Sturct.merchant_no_42 & vbCrLf & "->[�̻���] " & temp_42str & vbCrLf & vbCrLf
            tempCount = tempCount + 30
        End If


        Dim temp_43_len_str As String
        '//��43������ ��merchant_name_43��
        If Mid(temp_bcd_flag_str, 43, 1) Then
            temp_43_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.merchant_name_43.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.merchant_name_43.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.merchant_name_43.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.merchant_name_43.ptr = ins_space(POS_Sturct.merchant_name_43.ptr)
            temp_43_len_str = ins_space(temp_43_len_str)
            tex2STR = tex2STR & "[field43:�Զ�����(merchant_name_43)]" & vbCrLf & "" & "[" & temp_43_len_str & "]" & "  " & POS_Sturct.merchant_name_43.ptr & vbCrLf & vbCrLf

        End If


        Dim temp_44_len_str As String
        Dim issuing_bank As String   '//������
        Dim Acquiring_bank As String    '//�յ���
        '//��44������ ��������Ӧ���ݡ�
        If Mid(temp_bcd_flag_str, 44, 1) Then
            temp_44_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.rsp_code_44.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = POS_Sturct.rsp_code_44.Ptrlen * 2
            tempCount = tempCount + 2

            POS_Sturct.rsp_code_44.ptr = Mid(tempStr, tempCount, tempLen)
            issuing_bank = ASCchange(Mid(POS_Sturct.rsp_code_44.ptr, 1, 22))
            Acquiring_bank = ASCchange(Mid(POS_Sturct.rsp_code_44.ptr, 23, 22))


            tempCount = tempCount + tempLen
            POS_Sturct.rsp_code_44.ptr = ins_space(POS_Sturct.rsp_code_44.ptr)
            If tempLen = 44 Then
                tex2STR = tex2STR & "[field44:������Ӧ����(Additional Response Data)]" & vbCrLf & "" & "[" & temp_44_len_str & "]" & "  " & POS_Sturct.rsp_code_44.ptr _
                          & vbCrLf & "->[������] " & issuing_bank & vbCrLf & "->[�յ���] " & Acquiring_bank & vbCrLf & vbCrLf
            Else
                tex2STR = tex2STR & "[field44:������Ӧ����(Additional Response Data)]" & vbCrLf & "" & "[" & temp_44_len_str & "]" & "  " & POS_Sturct.rsp_code_44.ptr & vbCrLf & vbCrLf
            End If
        End If

        Dim temp_46_len_str As String
        '//��46������ ��pay_signature_46��
        If Mid(temp_bcd_flag_str, 46, 1) Then
            temp_46_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.pay_signature_46.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.pay_signature_46.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.pay_signature_46.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.pay_signature_46.ptr = ins_space(POS_Sturct.pay_signature_46.ptr)
            temp_46_len_str = ins_space(temp_46_len_str)
            tex2STR = tex2STR & "[field46:�Զ�����(pay_signature_46)]" & vbCrLf & "" & "[" & temp_46_len_str & "]" & "  " & POS_Sturct.pay_signature_46.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_48_len_str As String
        '//��48������ ���������� - ˽�С�
        If Mid(temp_bcd_flag_str, 48, 1) Then
            temp_48_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.settleAccounts_48.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.settleAccounts_48.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.settleAccounts_48.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.settleAccounts_48.ptr = ins_space(POS_Sturct.settleAccounts_48.ptr)
            temp_48_len_str = ins_space(temp_48_len_str)
            tex2STR = tex2STR & "[field48:�������� - ˽��(Additional Data - Private)]" & vbCrLf & "" & "[" & temp_48_len_str & "]" & "  " & POS_Sturct.settleAccounts_48.ptr & vbCrLf & vbCrLf
        End If

        Dim temp_49str As String
        '//��49������ �����׻��Ҵ��롿
        If Mid(temp_bcd_flag_str, 49, 1) Then
            POS_Sturct.currency_code_49 = Mid(tempStr, tempCount, 6)
            POS_Sturct.currency_code_49 = ins_space(POS_Sturct.currency_code_49)
            temp_49str = ASCchange(POS_Sturct.currency_code_49)

            tex2STR = tex2STR & "[field49:���׻��Ҵ���(Currency Code Of Transaction)]" & vbCrLf & "" & POS_Sturct.currency_code_49 & vbCrLf & vbCrLf

            tempCount = tempCount + 6
        End If

        '//��52������ �����˱�ʶ�롿
        If Mid(temp_bcd_flag_str, 52, 1) Then
            POS_Sturct.pri_pin_52 = Mid(tempStr, tempCount, 16)
            POS_Sturct.pri_pin_52 = ins_space(POS_Sturct.pri_pin_52)
            tex2STR = tex2STR & "[field52:���˱�ʶ������(PIN Data)]" & vbCrLf & "" & POS_Sturct.pri_pin_52 & vbCrLf & vbCrLf

            tempCount = tempCount + 16
        End If

        '//��53������ ����ȫ������Ϣ��
        If Mid(temp_bcd_flag_str, 53, 1) Then
            POS_Sturct.safety_53 = Mid(tempStr, tempCount, 16)
            POS_Sturct.safety_53 = ins_space(POS_Sturct.safety_53)
            tex2STR = tex2STR & "[field53:��ȫ������Ϣ(Security Related Control Information )]" & vbCrLf & "" & POS_Sturct.safety_53 & vbCrLf & vbCrLf

            tempCount = tempCount + 16
        End If

        Dim temp_54str As String
        Dim temp_54_len_str As String
        Dim temp_consume_amount_54 As String
        '//��54������ ����
        If Mid(temp_bcd_flag_str, 54, 1) Then
            temp_54_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.attachment_amount_54.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.attachment_amount_54.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.attachment_amount_54.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen

            POS_Sturct.attachment_amount_54.ptr = ins_space(POS_Sturct.attachment_amount_54.ptr)
            temp_54_len_str = ins_space(temp_54_len_str)

            temp_54str = ASCchange(POS_Sturct.attachment_amount_54.ptr)
            temp_consume_amount_54 = Val(Mid(temp_54str, 9, 12)) / 100


            If Val(temp_consume_amount_54) < 1 And Val(temp_consume_amount_54) > 0 Then
                temp_consume_amount_54 = 0 & temp_consume_amount_54
            End If
            tex2STR = tex2STR & "[field54:���(Balanc Amount)]" & vbCrLf & "" & "[" & temp_54_len_str & "]" & "  " & POS_Sturct.attachment_amount_54.ptr & vbCrLf & "->[ASCIIת��] " _
                      & temp_54str & vbCrLf & "->[���]      " & temp_consume_amount_54 & "Ԫ" & vbCrLf & vbCrLf
        End If



        Dim Application_Cryptogram$    '//Ӧ������  9F26   8�ֽ�



        Dim temp_55_len_str As String
        '//��55������ ��IC��������
        If Mid(temp_bcd_flag_str, 55, 1) Then
            temp_55_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.icData_55.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.icData_55.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.icData_55.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.icData_55.ptr = ins_space(POS_Sturct.icData_55.ptr)
            temp_55_len_str = ins_space(temp_55_len_str)
            tex2STR = tex2STR & "[field55:IC��������(IC Card System Related Data)]" & vbCrLf & "" & "[" & temp_55_len_str & "]" & "  " & POS_Sturct.icData_55.ptr & vbCrLf & vbCrLf

        End If

















        Dim temp_56_len_str As String
        '//��56������ ��private_data_56��
        If Mid(temp_bcd_flag_str, 56, 1) Then
            temp_56_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_56.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_56.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_56.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_56.ptr = ins_space(POS_Sturct.private_data_56.ptr)
            temp_56_len_str = ins_space(temp_56_len_str)
            tex2STR = tex2STR & "[field56:�Զ�����(private_data_56)]" & vbCrLf & "" & "[" & temp_56_len_str & "]" & "  " & POS_Sturct.private_data_56.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_57_len_str As String
        '//��57������ ��private_data_57��
        If Mid(temp_bcd_flag_str, 57, 1) Then
            temp_57_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_57.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_57.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_57.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_57.ptr = ins_space(POS_Sturct.private_data_57.ptr)
            temp_57_len_str = ins_space(temp_57_len_str)
            tex2STR = tex2STR & "[field57:�Զ�����(private_data_57)]" & vbCrLf & "" & "[" & temp_57_len_str & "]" & "  " & POS_Sturct.private_data_57.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_58_len_str As String
        '//��58������ ��PBOC����Ǯ����׼�Ľ�����Ϣ��
        If Mid(temp_bcd_flag_str, 58, 1) Then
            temp_58_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_58.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_58.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_58.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_58.ptr = ins_space(POS_Sturct.private_data_58.ptr)
            temp_58_len_str = ins_space(temp_58_len_str)
            tex2STR = tex2STR & "[field58:PBOC����Ǯ����׼�Ľ�����Ϣ(PBOC_ELECTRONIC_DATA)]" & vbCrLf & "" & "[" & temp_58_len_str & "]" & "  " & POS_Sturct.private_data_58.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_59_len_str As String
        '//��59������ ��private_data_59��
        If Mid(temp_bcd_flag_str, 59, 1) Then
            temp_59_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_59.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_59.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_59.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_59.ptr = ins_space(POS_Sturct.private_data_59.ptr)
            temp_59_len_str = ins_space(temp_59_len_str)
            tex2STR = tex2STR & "[field59:�Զ�����(private_data_59)]" & vbCrLf & "" & "[" & temp_59_len_str & "]" & "  " & POS_Sturct.private_data_59.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_60_len_str As String

        Dim temp_trans_Type_60 As String          '//60.1 ��Ϣ������
        Dim temp_batch_Number_60 As String        '//60.2 ���κ�
        Dim temp_network_60 As String             '//60.3 ���������Ϣ��
        Dim temp_readingAbility_60 As String      '//60.4 �ն˶�ȡ����
        Dim temp_conditionCode_60 As String       '//60.5 ���� PBOC ��/���Ǳ�׼�� IC ����������
        Dim temp_supportSome_60 As String         '//60.6 ֧�ֲ��ֿۿ�ͷ�������־
        Dim temp_account_type_60 As String        '//60.7 �ʻ�����

        Dim temp_trans_Type_60_str As String      '//60.1 ��Ϣ��������ʾ
        Dim temp_batch_Number_60_str As String    '//60.2 ���κ���ʾ
        Dim temp_network_60_str As String         '//60.3 ���������Ϣ����ʾ
        Dim temp_readingAbility_60_str As String  '//60.4 �ն˶�ȡ������ʾ
        Dim temp_conditionCode_60_str As String   '//60.5 ���� PBOC ��/���Ǳ�׼�� IC ������������ʾ
        Dim temp_supportSome_60_str As String     '//60.6 ֧�ֲ��ֿۿ�ͷ�������־��ʾ
        Dim temp_account_type_60_str As String    '//60.7 �ʻ�������ʾ

        '//��60������ ��private_data_60��
        If Mid(temp_bcd_flag_str, 60, 1) Then
            temp_60_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_60.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.private_data_60.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_60.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen


            temp_trans_Type_60 = Mid(POS_Sturct.private_data_60.ptr, 1, 2)    '//60.1 ��Ϣ������
            Select Case temp_trans_Type_60
            Case "00"
                temp_trans_Type_60_str = "�����ཻ�ף��ű�֪ͨ����"
            Case "01"
                temp_trans_Type_60_str = "��ѯ"
            Case "03"
                temp_trans_Type_60_str = "���ֲ�ѯ"
            Case "10"
                temp_trans_Type_60_str = "Ԥ��Ȩ/����"
            Case "11"
                temp_trans_Type_60_str = "Ԥ��Ȩ����/����"
            Case "20"
                temp_trans_Type_60_str = "Ԥ��Ȩ��ɣ����� /����"
            Case "21"
                temp_trans_Type_60_str = "Ԥ��Ȩ��ɳ���/����"
            Case "22"
                temp_trans_Type_60_str = "����/����"
            Case "23"
                temp_trans_Type_60_str = "���ѳ���/����"
            Case "24"
                temp_trans_Type_60_str = "Ԥ��Ȩ��ɣ�֪ͨ��"
            Case "25"
                temp_trans_Type_60_str = "�˻����������˻����˻���"
            Case "27"
                temp_trans_Type_60_str = "IC ���ѻ������˻�"
            Case "30"
                temp_trans_Type_60_str = "���߽���"
            Case "32"
                temp_trans_Type_60_str = "�������"
            Case "34"
                temp_trans_Type_60_str = "�������(׷��С��)"
            Case "36"
                temp_trans_Type_60_str = "�ѻ�����"
            Case "40"
                temp_trans_Type_60_str = "����Ǯ���� IC ��ָ���˻�Ȧ��/����"
            Case "41"
                temp_trans_Type_60_str = "����Ǯ���� IC ���ֽ��ֵ/����"
            Case "42"
                temp_trans_Type_60_str = "����Ǯ���� IC ����ָ���˻�ת��Ȧ��/����"
            Case "45"
                temp_trans_Type_60_str = "�����ֽ�ָ���˻�Ȧ��/����"
            Case "46"
                temp_trans_Type_60_str = "�����ֽ��ֽ��ֵ�������� /����"
            Case "47"
                temp_trans_Type_60_str = "�����ֽ��ָ���˻�ת��Ȧ��/����"
            Case "48"
                temp_trans_Type_60_str = "�������ֽ��ֵ/ȷ��"
            Case "49"
                temp_trans_Type_60_str = "�������ʻ���ֵ"
            Case "51"
                temp_trans_Type_60_str = "�����ֽ��ֽ��ֵ����/����"
            Case "53"
                temp_trans_Type_60_str = "ԤԼ���ѳ���/����"
            Case "54"
                temp_trans_Type_60_str = "ԤԼ����/����"
            Case Else
                temp_trans_Type_60_str = "N/A"
            End Select

            temp_batch_Number_60 = Mid(POS_Sturct.private_data_60.ptr, 3, 6)    '//60.2 ���κ�

            temp_network_60 = Mid(POS_Sturct.private_data_60.ptr, 9, 3)    '//60.3 ���������Ϣ��
            Select Case temp_network_60
            Case "001"
                temp_network_60_str = "POS �ն�ǩ������������Կ�㷨��"
            Case "002"
                temp_network_60_str = "POS �ն�ǩ��"
            Case "003"
                temp_network_60_str = "POS �ն�ǩ����˫������Կ�㷨)"
            Case "004"
                temp_network_60_str = "POS �ն�ǩ����˫������Կ�㷨�����ŵ���Կ��"
            Case "201"
                temp_network_60_str = "POS �ն�������"
            Case "201"
                temp_network_60_str = "POS �ն�������"
            Case "202"
                temp_network_60_str = "���˲�ƽ��ʱ�� POS �ն������ͽ���"
            Case "203"
                temp_network_60_str = "����ƽ��ʱ�� POS �ն����ͳɹ��� IC ��������"
            Case "204"
                temp_network_60_str = "����ƽ��ʱ�� POS �ն����� IC ��֪ͨ��Ϣ"
            Case "205"
                temp_network_60_str = "���˲�ƽ��ʱ�� POS �ն����ͳɹ��� IC ������"
            Case "206"
                temp_network_60_str = "���˲�ƽ��ʱ�� POS �ն����� IC ��֪ͨ��Ϣ"
            Case "207"
                temp_network_60_str = "����ƽ��ʱ�� POS �ն������ͽ���"
            Case "208"
                temp_network_60_str = "����ƽ��ʱ�� POS �ն�����Ȧ�潻��Ȧ��ȷ����"
            Case "209"
                temp_network_60_str = "���˲�ƽ��ʱ�� POS �ն�����Ȧ�潻��Ȧ��ȷ��"
            Case "301"
                temp_network_60_str = "�������"
            Case "401"
                temp_network_60_str = "����Աǩ��"
            Case "362"
                temp_network_60_str = "POS �ն�״̬���"
            Case "360"
                temp_network_60_str = "POS �ն˴�������������"
            Case "361"
                temp_network_60_str = "POS �ն˴������������ؽ���"
            Case "364"
                temp_network_60_str = "POS �ն� TMS ��������"
            Case "365"
                temp_network_60_str = "POS �ն� TMS �������ؽ���"
            Case "370"
                temp_network_60_str = "POS �ն� IC ����Կ����"
            Case "371"
                temp_network_60_str = "POS �ն� IC ����Կ���ؽ���"
            Case "372"
                temp_network_60_str = "POS �ն� IC ����Կ��Ϣ��ѯ"
            Case "380"
                temp_network_60_str = "POS �ն� IC ����������"
            Case "381"
                temp_network_60_str = "POS �ն� IC ���������ؽ���"
            Case "382"
                temp_network_60_str = "POS �ն� IC ��������Ϣ��ѯ"
            Case "384"
                temp_network_60_str = "POS �ն˱��ֻ������أ����ھ���ʹ�ã�"
            Case "385"
                temp_network_60_str = "POS �ն˱��ֻ������ؽ��������ھ���ʹ�ã�"
            Case "390"
                temp_network_60_str = "POS �ն˿� BIN ����������"
            Case "391"
                temp_network_60_str = "POS �ն˿� BIN ���������ؽ���"
            Case "392"
                temp_network_60_str = "POS �ն�С��ȡ�ֵ����������أ�Ԥ����"
            Case "393"
                temp_network_60_str = "POS �ն�С��ȡ�ֵ����������ؽ�����Ԥ����"
            Case "951"
                temp_network_60_str = "���� PBOC ��/���Ǳ�׼ IC ���ű�������֪ͨ"
            Case Else
                temp_network_60_str = "N/A"
            End Select

            temp_readingAbility_60 = Mid(POS_Sturct.private_data_60.ptr, 12, 1)     '//60.4 �ն˶�ȡ����
            Select Case temp_readingAbility_60
            Case "0"
                temp_readingAbility_60_str = "�ն˶�ȡ��������֪"
            Case "2"
                temp_readingAbility_60_str = "�ɶ�ȡ������"
            Case "5"
                temp_readingAbility_60_str = "�ɽӴ�ʽ�����ȡ IC �������ڵ���Ǯ���ķǽӴ������ȡ������Ҳ�� 5"
            Case "6"
                temp_readingAbility_60_str = "�ɷǽӴ�ʽ�����ȡ IC ���������ɶ�ȡ CUPMobile �ƶ�֧�������зǽӴ�ʽ�նˣ���" & _
                                             "��22 ��ǰ��λȡֵ 07�� 91�� 96 �� 98 ʱ����������� 6�������ڵ���Ǯ���ķǽӴ������ȡ��������Ȼ�� 5"
            Case Else
                temp_readingAbility_60_str = "N/A"
            End Select

            temp_conditionCode_60 = Mid(POS_Sturct.private_data_60.ptr, 13, 1)       '//60.5 ���� PBOC ��/���Ǳ�׼�� IC ����������
            Select Case temp_conditionCode_60
            Case "0"
                temp_conditionCode_60_str = "δʹ�û����������ڣ����ֻ�оƬ����"
            Case "1"
                temp_conditionCode_60_str = "��һ�ʽ��ײ��� IC �����׻���һ�ʳɹ��� IC ������"
            Case "2"
                temp_conditionCode_60_str = "��һ�ʽ������� IC �����׵�ʧ��"
            Case Else
                temp_conditionCode_60_str = "N/A"
            End Select

            temp_supportSome_60 = Mid(POS_Sturct.private_data_60.ptr, 14, 1)         '//60.6 ֧�ֲ��ֿۿ�ͷ�������־
            Select Case temp_conditionCode_60
            Case "0"
                temp_supportSome_60_str = "֧�ֲ��ֿۿ�ͷ�������־"
            Case "1"
                temp_supportSome_60_str = "��֧�ֲ��ֿۿ�ͷ�������־"
            Case Else
                temp_supportSome_60_str = "N/A"
            End Select

            temp_account_type_60 = Mid(POS_Sturct.private_data_60.ptr, 15, 3)           '//60.7 �ʻ�����
            Select Case temp_account_type_60
            Case "0"
                temp_account_type_60_str = "�����л��֣���ʾ����0��ASCII��"
            Case "1"
                temp_account_type_60_str = "�������˻��֣���ʾ��ĸA��ASCII��"
            Case Else
                temp_account_type_60_str = "N/A"
            End Select

            POS_Sturct.private_data_60.ptr = ins_space(POS_Sturct.private_data_60.ptr)

            temp_60_len_str = ins_space(temp_60_len_str)
            tex2STR = tex2STR & "[field60:�Զ�����(private_data_60)]" & vbCrLf & "" & "[" & temp_60_len_str & "]" & "  " & POS_Sturct.private_data_60.ptr & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[��Ϣ������]" & vbCrLf & temp_trans_Type_60 & vbCrLf & temp_trans_Type_60_str & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[���κ�]" & vbCrLf & temp_batch_Number_60 & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[���������Ϣ��]" & vbCrLf & temp_network_60 & vbCrLf & temp_network_60_str & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[�ն˶�ȡ����]" & vbCrLf & temp_readingAbility_60 & vbCrLf & temp_readingAbility_60_str & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[����PBOC ��/���Ǳ�׼��IC����������]" & vbCrLf & temp_conditionCode_60 & vbCrLf & temp_conditionCode_60_str & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[֧�ֲ��ֿۿ�ͷ�������־]" & vbCrLf & temp_supportSome_60 & vbCrLf & temp_supportSome_60_str & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[�ʻ�����]" & vbCrLf & temp_account_type_60 & vbCrLf & temp_account_type_60_str & vbCrLf & "-------------------------------------" & vbCrLf & vbCrLf

            batch_Number_60.Caption = temp_batch_Number_60

        End If

        Dim temp_61_len_str As String
        Dim Original_batch_Number_61 As String                             '//ԭʼ�������κ�
        Dim Original_trace_no_61 As String                                 '//ԭʼ����POS��ˮ��
        Dim Original_trans_date_61 As String                               '//ԭʼ��������
        Dim Original_trans_authorization_61 As String                      '// ԭ������Ȩ��ʽ
        Dim Original_trans_authorization_institution_code_61 As String     '//ԭ������Ȩ��������

        '//��61������ ��ԭʼ��Ϣ��
        If Mid(temp_bcd_flag_str, 61, 1) Then
            temp_61_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_61.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.private_data_61.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_61.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen

            Original_batch_Number_61 = Mid(POS_Sturct.private_data_61.ptr, 1, 6)
            Original_trace_no_61 = Mid(POS_Sturct.private_data_61.ptr, 7, 6)
            Original_trans_date_61 = ins_space(Mid(POS_Sturct.private_data_61.ptr, 13, 4))
            Original_trans_authorization_61 = Mid(POS_Sturct.private_data_61.ptr, 17, 2)
            Original_trans_authorization_institution_code_61 = Mid(POS_Sturct.private_data_61.ptr, 19, 11)

            POS_Sturct.private_data_61.ptr = ins_space(POS_Sturct.private_data_61.ptr)

            temp_61_len_str = ins_space(temp_61_len_str)
            tex2STR = tex2STR & "[field61:ԭʼ��Ϣ��(Original Message)]" & vbCrLf & "" & "[" & temp_61_len_str & "]" & "  " & POS_Sturct.private_data_61.ptr _
                      & vbCrLf & "---------------------" & vbCrLf & "->[ԭʼ�������κ�]" & vbCrLf & Original_batch_Number_61 _
                      & vbCrLf & "---------------------" & vbCrLf & "->[ԭʼ����POS��ˮ��]" & vbCrLf & Original_trace_no_61 _
                      & vbCrLf & "---------------------" & vbCrLf & "->[ԭʼ��������]" & vbCrLf & Original_trans_date_61 _
                      & "-> " & Mid(Original_trans_date_61, 1, 2) & "��" & Mid(Original_trans_date_61, 4, 2) & "��" _
                      & vbCrLf & "---------------------" & vbCrLf & "->[ԭ������Ȩ��ʽ]" & vbCrLf & Original_trans_authorization_61 _
                      & vbCrLf & "---------------------" & vbCrLf & "->[ԭ������Ȩ��������]" & vbCrLf & Original_trans_authorization_institution_code_61 _
                      & vbCrLf & "---------------------" & vbCrLf & vbCrLf

        End If


        Dim temp_62str As String
        Dim temp_62_len_str As String
        '//��62������ ��private_data_62��
        If Mid(temp_bcd_flag_str, 62, 1) Then
            temp_62_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_62.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_62.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_62.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_62.ptr = ins_space(POS_Sturct.private_data_62.ptr)
            temp_62str = ASCchange(POS_Sturct.private_data_62.ptr)
            temp_62_len_str = ins_space(temp_62_len_str)
            tex2STR = tex2STR & "[field62:�Զ�����(private_data_62)]" & vbCrLf & "" & "[" & temp_62_len_str & "]" & "  " & POS_Sturct.private_data_62.ptr & vbCrLf _
                      & "->[ASCIIת��]" & vbCrLf & temp_62str & vbCrLf & vbCrLf

        End If


        Dim temp_63str As String
        Dim temp_63_len_str As String
        '//��63������ ��private_data_63��
        If Mid(temp_bcd_flag_str, 63, 1) Then
            temp_63_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_63.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_63.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_63.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_63.ptr = ins_space(POS_Sturct.private_data_63.ptr)
            temp_63str = ASCchange(POS_Sturct.private_data_63.ptr)
            temp_63_len_str = ins_space(temp_63_len_str)
            tex2STR = tex2STR & "[field63:�Զ�����(private_data_63)]" & vbCrLf & "" & "[" & temp_63_len_str & "]" & "  " & POS_Sturct.private_data_63.ptr & vbCrLf _
                      & "->[ASCIIת��]" & vbCrLf & temp_63str & vbCrLf & vbCrLf

        End If

        '//��64������ �����ļ����롿
        If Mid(temp_bcd_flag_str, 64, 1) Then
            POS_Sturct.mac_64 = Mid(tempStr, tempCount, 16)
            POS_Sturct.mac_64 = ins_space(POS_Sturct.mac_64)
            tex2STR = tex2STR & "[field64:���ļ�����(Message Authentication Code)]" & vbCrLf & "" & POS_Sturct.mac_64 & vbCrLf & vbCrLf
            tempCount = tempCount + 16
        End If



        '//  tex2STR = tex2STR & "************8583������������ǻ����ķָ��߽�β��***************/" & vbCrLf
        analyse_after_data.Text = tex2STR
    End If

End Sub






'//����������

Private Sub bit_map_clear_Click()
    Dim i As Integer
    For i = 1 To 64
        bit_map_view_Click_flag(i) = False
        bit_map_view(i - 1).BackColor = &H8000000F    '//�ָ�ԭ��
    Next
    bit_map.Text = ""
End Sub

Private Sub bit_map_set_Click()
    bit_map.SetFocus
    Dim i As Integer
    If bit_map.Text <> "" Then
        bit_map_change_function

    Else
        '        i = MsgBox("λͼ��Ϊ�գ�����ȷ����", vbCritical, "��ʾ")
        MSG_BOX "λͼ��Ϊ��" & vbCrLf & "����ȷ����", "��ʾ"

        MainForm.Enabled = True
        bit_map.SetFocus
        MainForm.Enabled = False

        '        Debug.Print PromptForm.Enabled
        '        PromptForm.Enabled = True
        PromptForm.Show
    End If
End Sub

Private Function bit_map_view_Click_function(Index As Integer)
    Dim temp_bit_map_str As String
    Dim temp_bcd_str, i As Integer

    Index = Index + 1
    If bit_map_view_Click_flag(Index) = False Then
        bit_map_view_Click_flag(Index) = True
        bit_map_view(Index - 1).BackColor = &HFF8080    '//����ȥǳ��ɫ
    Else
        bit_map_view_Click_flag(Index) = False
        bit_map_view(Index - 1).BackColor = &H8000000F    '//�ָ�ԭ��
    End If

    For i = 1 To 64

        If bit_map_view_Click_flag(i) = False Then

            temp_bcd_str = temp_bcd_str & "0"
        ElseIf bit_map_view_Click_flag(i) = True Then

            temp_bcd_str = temp_bcd_str & "1"
        End If

    Next
    temp_bit_map_str = BIN_to_HEX(temp_bcd_str)
    bit_map.Text = ins_space(temp_bit_map_str)
End Function


Private Sub bit_map_view_Click(Index As Integer)
    bit_map.SetFocus
    bit_map_view_Click_function (Index)
End Sub

Private Sub bit_map_view_DblClick(Index As Integer)
    bit_map.SetFocus
    bit_map_view_Click_function (Index)
End Sub


Private Function bit_map_change_function()

    Dim temp_bit_map_str As String
    Dim temp_bcd_str, i As Integer

    For i = 1 To 64
        bit_map_view_Click_flag(i) = False
        bit_map_view(i - 1).BackColor = &H8000000F    '//�ָ�ԭ��
    Next   '//���λͼ

    temp_bit_map_str = delete_space(bit_map.Text)


    temp_bcd_str = HEX_to_BIN(temp_bit_map_str)

    If Len(temp_bcd_str) > 64 Then

        temp_bcd_str = Mid(temp_bcd_str, 1, 64)

    End If

    Do While Len(temp_bcd_str) < 64
        temp_bcd_str = temp_bcd_str & "0"
    Loop

    For i = 1 To 64
        If Mid(temp_bcd_str, i, 1) Then

            bit_map_view_Click_flag(i) = True
            bit_map_view(i - 1).BackColor = &HFF8080  '//����ȥǳ��ɫ
        End If
    Next
    temp_bit_map_str = BIN_to_HEX(temp_bcd_str)
    bit_map.Text = ins_space(temp_bit_map_str)

End Function


Private Sub clear_Click()    '//�����¼�
    analyse_before_data.Text = ""
    analyse_after_data.Text = ""
    bit_map_clear_Click
    trans_type.Caption = "N/A"
    judge_mode.Caption = "N/A"
    judge_mode.ForeColor = &H0&    '//�ָ��ɺ�ɫ

    Response_code.Caption = "N/A"
    Response_code_view.Caption = "N/A"

    trace_no_11.Caption = "N/A"
    batch_Number_60.Caption = "N/A"

    help_flag = 0
    help.Caption = "˵��"


    analyse_before_data.SetFocus


End Sub

Private Sub about_Click()    '//�����¼�
    Dim i As Integer

    '    i = MsgBox("8583������������V1.2                      " _
         '               & vbCrLf & "" _
         '               & vbCrLf & "���ߣ��߽���" _
         '               & vbCrLf & "�汾��V1.2" _
         '               & vbCrLf & "QQ��1062220953" _
         '               & vbCrLf & "���ڣ�2015��01��10��" _
         '               & vbCrLf & "TIPS���ڰ�ť������ͣ�鿴����" _
         '               , vbOKOnly, "����")
    About_BOX
End Sub




Private Sub END_Click()    '//�˳��¼�
    End
End Sub



Private Sub Frame_analyse_after_data_Click()    '//���ƽ���������
    Clipboard.clear
    Clipboard.SetText analyse_after_data.Text    '//�Զ����ƽ��
    MSG_BOX "���������ݸ��Ƴɹ�", "��ʾ"
End Sub



Private Sub Frame_analyse_before_data_Click()   '//���ƽ���ǰ����
    Clipboard.clear
    Clipboard.SetText analyse_before_data.Text    '//�Զ����ƽ��
    MSG_BOX "����ǰ���ݸ��Ƴɹ�", "��ʾ"
End Sub

Private Sub Frame_bit_map_Click()  '//����λͼ����
    Clipboard.clear
    Clipboard.SetText bit_map.Text    '//�Զ����ƽ��
    MSG_BOX "λͼ���Ƴɹ�", "��ʾ"
End Sub


Private Sub help_Click()    '//�����¼�

    Static temp_analyse_after_data_str As String
    If help_count = 0 Then
        temp_analyse_after_data_str = analyse_after_data.Text
    End If
    help_count = help_count + 1
    If help_flag = 0 Then
        analyse_after_data.SetFocus
        analyse_after_data.Text = "                   ��8583������������V1.3.1 ˵����                     " & vbCrLf & vbCrLf & _
                                  "���ա����۵��նˣ�POS��Ӧ�ù淶(QCUP 009.1-2010)�����н���" & vbCrLf & vbCrLf & _
                                  "V1.0���ӻ���������λͼ��ʾ�����ã���ͣ����밴ť���Բ鿴�������" & vbCrLf & vbCrLf & _
                                  "V1.1��������ͨ������1.0�����������Ϊר�ҽ�����������Ӧ����ʾ��TIPS:��֧�����۵��նˣ�POS��Ӧ�ù淶(QCUP 009.1-2010)Э��Ľ���������������Ǳ�׼Э�鱨�Ľ��������ܻ�������Ͳ�ƥ��ľ��档" & vbCrLf & vbCrLf & _
                                  "V1.2�Ľ�����" & vbCrLf & _
                                  "->1.�����µ�ͼ��ͽ���" & vbCrLf & _
                                  "->2.�Ż���ר�ҽ�����������ȫ�����" & vbCrLf & _
                                  "->3.��������ˮ��/���κ���ʾ" & vbCrLf & _
                                  "->4.�����˳���ť" & vbCrLf & vbCrLf & _
                                  "V1.3�Ľ�����" & vbCrLf & _
                                  "->1.������55�����" & vbCrLf & _
                                  "->2.����Ӧ��39���ڱ�׼���Ҳ�������ʾ����ʧ�ܡ�����77��Ӧ��" & vbCrLf & _
                                  "->3.������ȫ�����" & vbCrLf & _
                                  "->4.���������Ͳ�ƥ��Ľ���ǰ���ݣ�����ȷ�Ϻ��˳�����" & vbCrLf & _
                                  "->5.�������ڴ���" & vbCrLf & _
                                  "->6.���ӿ����ʾ���ƿ�������ݹ���" & vbCrLf & vbCrLf & _
                                  "V1.3.1�Ľ�����" & vbCrLf & _
                                  "->1.�޸�55���������TAGֵ�����������" & vbCrLf

        help_flag = 1
        help.Caption = "���˵��"
    Else
        analyse_after_data.SetFocus
        help_flag = 0
        help.Caption = "˵��"
        analyse_after_data.Text = temp_analyse_after_data_str

    End If

End Sub


Private Sub MessageHeader_check_Click()
    If analyse_before_data.Text <> "" Then

        Select Case analyse_mode
        Case 1
            analyse_level_1_Click
        Case 2
            analyse_level_2_Click
        Case 3
            analyse_level_3_Click

        Case Else

        End Select
    End If
End Sub


Private Sub Totallen_check_Click()
    If analyse_before_data.Text <> "" Then
        Select Case analyse_mode
        Case 1
            analyse_level_1_Click
        Case 2
            analyse_level_2_Click
        Case 3
            analyse_level_3_Click

        Case Else

        End Select
    End If
End Sub

Private Sub TPDU_check_Click()
    If analyse_before_data.Text <> "" Then
        Select Case analyse_mode
        Case 1
            analyse_level_1_Click
        Case 2
            analyse_level_2_Click
        Case 3
            analyse_level_3_Click

        Case Else

        End Select
    End If
End Sub






























'//������


'/***********************************************************************************
'��������:PROCESS_Analyze_data
'��������:�������ǰ�����ݣ�������ת���ɿ��Է���������,ȥ�������еĿո񣬻س������з������ڽ���
'�䡡��:
'�䡡��:
'��  ע:
'���ӣ�12 34 56 78 90 ������->1234567890
'
'***********************************************************************************/

Private Sub PROCESS_Analyze_data()
    Dim Text1_tempStr, Text2_tempStr As String, count, totalNum, i As Long

    Text1_tempStr = analyse_before_data.Text

    If Len(Text1_tempStr) = 0 Then

        '        i = MsgBox("����ǰ����Ϊ�գ�����ȷ����", vbCritical, "��ʾ")
        MSG_BOX "����ǰ����Ϊ��" & vbCrLf & "����ȷ����", "��ʾ"
        MainForm.Enabled = True
        analyse_before_data.SetFocus
        MainForm.Enabled = False
        '        PromptForm.Enabled = True
        PromptForm.Show
        change_begin = 0

    Else
        change_begin = 1
        For count = 1 To Len(Text1_tempStr)
            If Mid(Text1_tempStr, count, 1) <> Chr(32) Then _
               If Mid(Text1_tempStr, count, 1) <> Chr(13) Then _
               If Mid(Text1_tempStr, count, 1) <> Chr(10) Then _
               Text2_tempStr = Text2_tempStr & Mid(Text1_tempStr, count, 1): totalNum = totalNum + 1
        Next

        analyse_before_data.Text = ins_space(Text2_tempStr)
        analyse_after_data.Text = Text2_tempStr
    End If
End Sub

'/*********************************************************************************************************
'** ��������: ins_space
'** ��������: ����ո�
'** �䡡��:
'** �䡡��:s
'** ��  ע:����:1234-->12 34
'********************************************************************************************************/


Private Function ins_space(ByVal src As String) As String    '//����ո�����:1234-->12 34

    Dim tempStr As String, count, even_num, i As Long

    even_num = 1

    If Len(src) = 0 Then


    Else

        For count = 1 To Len(src)
            If Mid(src, count, 1) <> Chr(32) Then _
               If Mid(src, count, 1) <> Chr(13) Then _
               If Mid(src, count, 1) <> Chr(10) Then _
               tempStr = tempStr & Mid(src, count, 1): even_num = even_num + 1

            If Mid(src, count, 1) <> Chr(32) Then _
               If Mid(src, count, 1) <> Chr(13) Then _
               If Mid(src, count, 1) <> Chr(10) Then _
               If even_num Mod 2 = 1 Then _
               If count <> Len(src) Then tempStr = tempStr & " "

        Next

    End If
    ins_space = tempStr
End Function

'/*********************************************************************************************************
'** ��������: delete_space
'** ��������: ɾ���ո�
'** �䡡��:
'** �䡡��:
'** ��  ע:����:12 34-->1234
'********************************************************************************************************/


Private Function delete_space(ByVal src As String) As String    '//����ո�����:1234-->12 34

    Dim tempStr As String, count, even_num, i As Long


    even_num = 1

    If Len(src) = 0 Then


    Else

        For count = 1 To Len(src)
            If Mid(src, count, 1) <> Chr(32) Then _
               If Mid(src, count, 1) <> Chr(13) Then _
               If Mid(src, count, 1) <> Chr(10) Then _
               tempStr = tempStr & Mid(src, count, 1): even_num = even_num + 1
        Next

    End If
    delete_space = tempStr
End Function

' ��;����ʮ������ת��Ϊ������
' ���룺Hex(ʮ��������)
' �����������ͣ�String
' �����HEX_to_BIN(��������)
' ����������ͣ�String
' ����������Ϊ2147483647���ַ�
Public Function HEX_to_BIN(ByVal HEX As String) As String
    Dim i As Long
    Dim b As String

    HEX = UCase(HEX)
    For i = 1 To Len(HEX)
        Select Case Mid(HEX, i, 1)
        Case "0": b = b & "0000"
        Case "1": b = b & "0001"
        Case "2": b = b & "0010"
        Case "3": b = b & "0011"
        Case "4": b = b & "0100"
        Case "5": b = b & "0101"
        Case "6": b = b & "0110"
        Case "7": b = b & "0111"
        Case "8": b = b & "1000"
        Case "9": b = b & "1001"
        Case "A": b = b & "1010"
        Case "B": b = b & "1011"
        Case "C": b = b & "1100"
        Case "D": b = b & "1101"
        Case "E": b = b & "1110"
        Case "F": b = b & "1111"
        End Select
    Next i
    HEX_to_BIN = b
End Function


' ��;����������ת��Ϊʮ������
' ���룺Bin(��������)
' �����������ͣ�String
' �����BIN_to_HEX(ʮ��������)
' ����������ͣ�String
' ����������Ϊ2147483647���ַ�
Public Function BIN_to_HEX(ByVal Bin As String) As String
    Dim i As Long
    Dim H As String
    If Len(Bin) Mod 4 <> 0 Then
        Bin = String(4 - Len(Bin) Mod 4, "0") & Bin
    End If

    For i = 1 To Len(Bin) Step 4
        Select Case Mid(Bin, i, 4)
        Case "0000": H = H & "0"
        Case "0001": H = H & "1"
        Case "0010": H = H & "2"
        Case "0011": H = H & "3"
        Case "0100": H = H & "4"
        Case "0101": H = H & "5"
        Case "0110": H = H & "6"
        Case "0111": H = H & "7"
        Case "1000": H = H & "8"
        Case "1001": H = H & "9"
        Case "1010": H = H & "A"
        Case "1011": H = H & "B"
        Case "1100": H = H & "C"
        Case "1101": H = H & "D"
        Case "1110": H = H & "E"
        Case "1111": H = H & "F"
        End Select
    Next i
    BIN_to_HEX = H
End Function





' ��;����ʮ������ת��Ϊʮ���� 31->48(֧�������λʮ������ת����ʮ����)
Public Function HEX_to_DEC(HEX As String) As Integer
    Dim shiwei, baiwei As String
    Dim i As Integer
    Dim b As Integer
    HEX = UCase(HEX)

    For i = 1 To 2
        Select Case Mid(HEX, i, 1)
        Case "0": b = b + 16 ^ (2 - i) * 0
        Case "1": b = b + 16 ^ (2 - i) * 1
        Case "2": b = b + 16 ^ (2 - i) * 2
        Case "3": b = b + 16 ^ (2 - i) * 3
        Case "4": b = b + 16 ^ (2 - i) * 4
        Case "5": b = b + 16 ^ (2 - i) * 5
        Case "6": b = b + 16 ^ (2 - i) * 6
        Case "7": b = b + 16 ^ (2 - i) * 7
        Case "8": b = b + 16 ^ (2 - i) * 8
        Case "9": b = b + 16 ^ (2 - i) * 9
        Case "A": b = b + 16 ^ (2 - i) * 10
        Case "B": b = b + 16 ^ (2 - i) * 11
        Case "C": b = b + 16 ^ (2 - i) * 12
        Case "D": b = b + 16 ^ (2 - i) * 13
        Case "E": b = b + 16 ^ (2 - i) * 14
        Case "F": b = b + 16 ^ (2 - i) * 15
        End Select
    Next i

    HEX_to_DEC = b
End Function

'/***********************************************************************************
'��������:ACSchange
'��������:ACSCII��ת��
'�䡡��:
'�䡡��:
'��  ע:���ӣ�31 32 33 34 35 36 37 38 39 30->1234567890
'***********************************************************************************/
Private Function ASCchange(ByVal ASC_scr As String) As String

    Dim tempStr As String, count, even_num, i As Long

    even_num = 1

    If Len(ASC_scr) = 0 Then


    Else
        ASC_scr = delete_space(ASC_scr)
        For count = 1 To Len(ASC_scr)

            tempStr = tempStr & Mid(ASC_scr, count, 1)
            even_num = even_num + 1
            If even_num Mod 2 = 1 Then
                ASCchange = ASCchange & Chr(HEX_to_DEC(tempStr))
                tempStr = ""
            End If
        Next
    End If
End Function

'/************************************************************************************************
'* ����:     MSG_BOX
'* ����˵��: ��ʾ��
'* ����:
'* ����:     Prompt����ѡ���ַ������ʽ����ʾ�ڶԻ����е���Ϣ
'            title:  ��ѡ���ַ������ʽ���ڶԻ������������ʾ������
'* ����ֵ:
'* ��ע:
'************************************************************************************************/
Private Sub MSG_BOX(Prompt As String, title As String)
    PromptForm.Show
    MainForm.Enabled = False

    PromptForm.msg_display.Caption = Prompt
    PromptForm.Caption = title
End Sub


'/************************************************************************************************
'* ����:     About_BOX
'* ����˵��: ���ڿ�
'* ����:
'* ����:     Prompt����ѡ���ַ������ʽ����ʾ�ڶԻ����е���Ϣ
'            title:  ��ѡ���ַ������ʽ���ڶԻ������������ʾ������
'* ����ֵ:
'* ��ע:
'************************************************************************************************/
Private Sub About_BOX()
    Form_About.Show
    MainForm.Enabled = False

    '    PromptForm.msg_display.Caption = Prompt
    '    PromptForm.Caption = title
End Sub



























'//������
'/***********************************************************************************
'��������:Type_judge
'��������:�����ж�
'�䡡��:
'�䡡��:
'��  ע:�������۵��նˣ�POS��Ӧ�ù淶(QCUP 009.1-2010)�淶
'��������ģʽ�к�ɫ����������ɫ������Ӧ
'����field61_flag��־λ���� ���Ѻ�Ԥ��Ȩ��� 2014��12��24��09:51:16
'������PBOCϵ�С��͡�������ϵ�С��͡���������ѻ��ࡿ�͡�����Ա����ǩ������û���ж�
'����������жϣ��Դ˲�����
'***********************************************************************************/
Private Function Type_judge(ByVal messagetype As String, ByVal procode As String, ByVal field61_flag As String)

'//������
    If messagetype = "02 00" And procode = "31" Then
        trans_type.Caption = "����ѯ"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "02 10" And procode = "31" Then
        trans_type.Caption = "����ѯ"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "02 00" And procode = "00" And field61_flag = 0 Then
        trans_type.Caption = "����"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "02 10" And procode = "00" And field61_flag = 0 Then
        trans_type.Caption = "����"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "04 00" And procode = "00" Then
        trans_type.Caption = "���ѳ���"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "04 10" And procode = "00" Then
        trans_type.Caption = "���ѳ���"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "02 00" And procode = "20" Then
        trans_type.Caption = "���ѳ���"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "02 10" And procode = "20" Then
        trans_type.Caption = "���ѳ���"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "04 00" And procode = "20" Then
        trans_type.Caption = "���ѳ�������"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "04 10" And procode = "20" Then
        trans_type.Caption = "���ѳ�������"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "02 20" And procode = "20" Then
        trans_type.Caption = "�˻�"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "02 30" And procode = "20" Then
        trans_type.Caption = "�˻�"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "01 00" And procode = "03" Then
        trans_type.Caption = "Ԥ��Ȩ"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "01 10" And procode = "03" Then
        trans_type.Caption = "Ԥ��Ȩ"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "04 00" And procode = "03" Then
        trans_type.Caption = "Ԥ��Ȩ����"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "04 10" And procode = "03" Then
        trans_type.Caption = "Ԥ��Ȩ����"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "01 00" And procode = "20" Then
        trans_type.Caption = "Ԥ��Ȩ����"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "01 10" And procode = "20" Then
        trans_type.Caption = "Ԥ��Ȩ����"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "04 00" And procode = "20" Then
        trans_type.Caption = "Ԥ��Ȩ��������"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "04 10" And procode = "20" Then
        trans_type.Caption = "Ԥ��Ȩ��������"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "02 00" And procode = "00" And field61_flag = 1 Then
        trans_type.Caption = "Ԥ��Ȩ���(����)"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "02 10" And procode = "00" And field61_flag = 1 Then
        trans_type.Caption = "Ԥ��Ȩ���(����)"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "02 20" And procode = "00" And field61_flag = 1 Then
        trans_type.Caption = "Ԥ��Ȩ���(֪ͨ)"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "02 30" And procode = "00" And field61_flag = 1 Then
        trans_type.Caption = "Ԥ��Ȩ���(֪ͨ)"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "04 00" And procode = "00" And field61_flag = 1 Then
        trans_type.Caption = "Ԥ��Ȩ���(����)����"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "04 10" And procode = "00" And field61_flag = 1 Then
        trans_type.Caption = "Ԥ��Ȩ���(����)����"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "02 00" And procode = "20" And field61_flag = 1 Then
        trans_type.Caption = "Ԥ��Ȩ��ɳ���"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "02 10" And procode = "20" And field61_flag = 1 Then
        trans_type.Caption = "Ԥ��Ȩ��ɳ���"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "04 00" And procode = "20" And field61_flag = 1 Then
        trans_type.Caption = "Ԥ��Ȩ��ɳ�������"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "04 10" And procode = "20" And field61_flag = 1 Then
        trans_type.Caption = "Ԥ��Ȩ��ɳ�������"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

        '//������
    ElseIf messagetype = "08 00" Then
        trans_type.Caption = "ǩ��"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "08 10" Then
        trans_type.Caption = "ǩ��"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "08 20" Then
        trans_type.Caption = "ǩ��"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "08 30" Then
        trans_type.Caption = "ǩ��"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "05 00" Then
        trans_type.Caption = "������"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "05 10" Then
        trans_type.Caption = "������"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ

    ElseIf messagetype = "03 20" Then
        trans_type.Caption = "�����ͽ��ڽ���/�����ͽ���"
        judge_mode.Caption = "����"
        judge_mode.ForeColor = &HFF&    '//��ɫ
    ElseIf messagetype = "03 30" Then
        trans_type.Caption = "�����ͽ��ڽ���/�����ͽ���"
        judge_mode.Caption = "��Ӧ"
        judge_mode.ForeColor = &HFF0000    '//��ɫ
    Else
        trans_type.Caption = "N/A"
        judge_mode.Caption = "N/A"
        judge_mode.ForeColor = &H0&    '//�ָ��ɺ�ɫ
    End If

End Function



'/***********************************************************************************
'��������:Response_code_39_Type_judge
'��������:�����ж�
'�䡡��:
'�䡡��:
'��  ע:�������۵��նˣ�POS��Ӧ�ù淶(QCUP 009.1-2010)�淶
'����ultra edit ��д
'***********************************************************************************/
Private Function Response_code_39_Type_judge(ByVal src As String)

    If src = "" Then
        Response_code_view.Caption = "N/A": Response_code.Caption = "N/A"
    ElseIf src = "00" Then
        Response_code_view.Caption = "���׳ɹ�": Response_code.Caption = src
    ElseIf src = "01" Then
        Response_code_view.Caption = "��ֿ����뷢��������ϵ": Response_code.Caption = src
    ElseIf src = "03" Then
        Response_code_view.Caption = "��Ч�̻�": Response_code.Caption = src
    ElseIf src = "04" Then
        Response_code_view.Caption = "�˿���û��": Response_code.Caption = src
    ElseIf src = "05" Then
        Response_code_view.Caption = "�ֿ�����֤ʧ��": Response_code.Caption = src
    ElseIf src = "10" Then
        Response_code_view.Caption = "��ʾ������׼����ʾ����Ա": Response_code.Caption = src
    ElseIf src = "11" Then
        Response_code_view.Caption = "�ɹ���VIP�ͻ�": Response_code.Caption = src
    ElseIf src = "12" Then
        Response_code_view.Caption = "��Ч����": Response_code.Caption = src
    ElseIf src = "13" Then
        Response_code_view.Caption = "��Ч���": Response_code.Caption = src
    ElseIf src = "14" Then
        Response_code_view.Caption = "��Ч����": Response_code.Caption = src
    ElseIf src = "15" Then
        Response_code_view.Caption = "�˿��޶�Ӧ������": Response_code.Caption = src
    ElseIf src = "21" Then
        Response_code_view.Caption = "�ÿ�δ��ʼ����˯�߿�": Response_code.Caption = src
    ElseIf src = "22" Then
        Response_code_view.Caption = "�������󣬻򳬳�������������": Response_code.Caption = src
    ElseIf src = "25" Then
        Response_code_view.Caption = "û��ԭʼ���ף�����ϵ������": Response_code.Caption = src
    ElseIf src = "30" Then
        Response_code_view.Caption = "������": Response_code.Caption = src
    ElseIf src = "34" Then
        Response_code_view.Caption = "���׿�,�׿�": Response_code.Caption = src
    ElseIf src = "38" Then
        Response_code_view.Caption = "�������������ޣ����뷢������ϵ": Response_code.Caption = src
    ElseIf src = "40" Then
        Response_code_view.Caption = "��������֧�ֵĽ�������": Response_code.Caption = src
    ElseIf src = "41" Then
        Response_code_view.Caption = "��ʧ������û�գ�POS��": Response_code.Caption = src
    ElseIf src = "43" Then
        Response_code_view.Caption = "���Կ�����û��": Response_code.Caption = src
    ElseIf src = "51" Then
        Response_code_view.Caption = "��������": Response_code.Caption = src
    ElseIf src = "54" Then
        Response_code_view.Caption = "�ÿ��ѹ���": Response_code.Caption = src
    ElseIf src = "55" Then
        Response_code_view.Caption = "�����": Response_code.Caption = src
    ElseIf src = "57" Then
        Response_code_view.Caption = "������˿�����": Response_code.Caption = src
    ElseIf src = "58" Then
        Response_code_view.Caption = "������������ÿ��ڱ��ն˽��д˽���": Response_code.Caption = src
    ElseIf src = "59" Then
        Response_code_view.Caption = "��ƬУ���": Response_code.Caption = src
    ElseIf src = "61" Then
        Response_code_view.Caption = "���׽���": Response_code.Caption = src
    ElseIf src = "62" Then
        Response_code_view.Caption = "�����ƵĿ�": Response_code.Caption = src
    ElseIf src = "64" Then
        Response_code_view.Caption = "���׽����ԭ���ײ�ƥ��": Response_code.Caption = src
    ElseIf src = "65" Then
        Response_code_view.Caption = "�������Ѵ�������": Response_code.Caption = src
    ElseIf src = "68" Then
        Response_code_view.Caption = "���׳�ʱ��������": Response_code.Caption = src
    ElseIf src = "75" Then
        Response_code_view.Caption = "��������������": Response_code.Caption = src
    ElseIf src = "90" Then
        Response_code_view.Caption = "ϵͳ���У����Ժ�����": Response_code.Caption = src
    ElseIf src = "91" Then
        Response_code_view.Caption = "������״̬�����������Ժ�����": Response_code.Caption = src
    ElseIf src = "92" Then
        Response_code_view.Caption = "��������·�쳣�����Ժ�����": Response_code.Caption = src
    ElseIf src = "94" Then
        Response_code_view.Caption = "�ܾ����ظ����ף����Ժ�����": Response_code.Caption = src
    ElseIf src = "96" Then
        Response_code_view.Caption = "�ܾ������������쳣�����Ժ�����": Response_code.Caption = src
    ElseIf src = "97" Then
        Response_code_view.Caption = "�ն�δ�Ǽ�": Response_code.Caption = src
    ElseIf src = "98" Then Response_code_view.Caption = "��������ʱ": Response_code.Caption = src
    ElseIf src = "99" Then
        Response_code_view.Caption = "PIN��ʽ��������ǩ��": Response_code.Caption = src
    ElseIf src = "A0" Then
        Response_code_view.Caption = "MACУ���������ǩ��": Response_code.Caption = src
    ElseIf src = "A1" Then
        Response_code_view.Caption = "ת�˻��Ҳ�һ��": Response_code.Caption = src
    ElseIf src = "A2" Then
        Response_code_view.Caption = "���׳ɹ������򷢿���ȷ��": Response_code.Caption = src
    ElseIf src = "A3" Then
        Response_code_view.Caption = "�˻�����ȷ": Response_code.Caption = src
    ElseIf src = "A4" Then
        Response_code_view.Caption = "���׳ɹ������򷢿���ȷ��": Response_code.Caption = src
    ElseIf src = "A5" Then
        Response_code_view.Caption = "���׳ɹ������򷢿���ȷ��": Response_code.Caption = src
    ElseIf src = "A6" Then
        Response_code_view.Caption = "���׳ɹ������򷢿���ȷ��": Response_code.Caption = src
    ElseIf src = "A7" Then
        Response_code_view.Caption = "�ܾ������������쳣�����Ժ�����": Response_code.Caption = src
    ElseIf src = "77" Then
        Response_code_view.Caption = "����Ա����ǩ������������ ": Response_code.Caption = src
    Else
        Response_code_view.Caption = "����ʧ��": Response_code.Caption = src   '//�������ֱ�ӽ���ʧ��
    End If

End Function





















