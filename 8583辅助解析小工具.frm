VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "8583辅助解析工具V1.3.1"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13590
   Icon            =   "8583辅助解析小工具.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   13590
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton analyse_level_3 
      Caption         =   "全面解析"
      Enabled         =   0   'False
      Height          =   495
      Left            =   12240
      TabIndex        =   95
      ToolTipText     =   "专家模式的BT版"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame10 
      Caption         =   "流水/批次号"
      Height          =   1215
      Left            =   12120
      TabIndex        =   93
      ToolTipText     =   "我是流水/批次号"
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
         Caption         =   "批次号："
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
         Caption         =   "流水号："
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton analyse_level_1 
      Caption         =   "普通解析"
      Height          =   495
      Left            =   12240
      TabIndex        =   1
      ToolTipText     =   "普通解析"
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame9 
      Caption         =   "响应码提示"
      Height          =   2175
      Left            =   12120
      TabIndex        =   90
      ToolTipText     =   "POS 2010规范中响应码提示"
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
            Name            =   "宋体"
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
      Caption         =   "位图"
      Height          =   360
      Left            =   2800
      TabIndex        =   7
      ToolTipText     =   "只影响位图和位图显示"
      Top             =   6765
      Width           =   615
   End
   Begin VB.CommandButton bit_map_clear 
      Caption         =   "清空"
      Height          =   360
      Left            =   3430
      TabIndex        =   8
      ToolTipText     =   "只清空位图和位图显示"
      Top             =   6765
      Width           =   615
   End
   Begin VB.Frame Frame7 
      Caption         =   "位图显示"
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      ToolTipText     =   "点击位图显示结果只能影响位图"
      Top             =   7320
      Width           =   11895
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "64"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域64"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "63"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域63"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "62"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域62"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "61"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域61"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "60"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域60"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "59"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域59"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "58"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域58"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "57"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域57"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "56"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域56"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "55"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域55"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "54"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域54"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "53"
         BeginProperty Font 
            Name            =   "宋体"
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
            Name            =   "宋体"
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
         ToolTipText     =   "域52"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "51"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域51"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域50"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "49"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域49"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "48"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域48"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "47"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域47"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "46"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域46"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "45"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域45"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "44"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域44"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "43"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域43"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "42"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域42"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "41"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域41"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域40"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "39"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域39"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域38"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "37"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域37"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域36"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "35"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域35"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域34"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "33"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域33"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域32"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域31"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域30"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域29"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域28"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域27"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域26"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域25"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域24"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域23"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域22"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域21"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域20"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域19"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域18"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域17"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域16"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域15"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域14"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域13"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域12"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域11"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域10"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域9"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域8"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域7"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域6"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域5"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域4"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域3"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域2"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label bit_map_view 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
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
         ToolTipText     =   "域1"
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame_bit_map 
      Caption         =   "位图[单击框架复制数据]"
      Height          =   615
      Left            =   120
      TabIndex        =   23
      ToolTipText     =   "我是位图"
      Top             =   6600
      Width           =   3975
      Begin VB.TextBox bit_map 
         BeginProperty Font 
            Name            =   "宋体"
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
      Caption         =   "解析设置[慎选，可能造成程序崩溃]"
      Height          =   735
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "去掉一些可能不需要解析的信息(慎选，可能造成程序崩溃)"
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
      Caption         =   "类型判断[可能造成误判断，仅供参考]"
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "可能造成误判断，仅作参考"
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
         Caption         =   "模式："
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
         Caption         =   "交易类型："
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
      Caption         =   "说明"
      Height          =   495
      Left            =   12240
      TabIndex        =   6
      ToolTipText     =   "说明"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton about 
      Caption         =   "关于"
      Height          =   495
      Left            =   12240
      TabIndex        =   5
      ToolTipText     =   "关于"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton clear 
      Caption         =   "清屏"
      Height          =   495
      Left            =   12240
      TabIndex        =   4
      ToolTipText     =   "清除屏幕内容"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton analyse_level_2 
      Caption         =   "专家解析"
      Height          =   495
      Left            =   12240
      TabIndex        =   2
      ToolTipText     =   "专家解析"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox analyse_after_data 
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "解析前数据[单击框架复制数据]"
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "我是解析前数据"
      Top             =   120
      Width           =   3975
   End
   Begin VB.Frame Frame_analyse_after_data 
      Caption         =   "解析后数据[单击框架复制数据]"
      Height          =   7095
      Left            =   4200
      TabIndex        =   10
      ToolTipText     =   "我是解析后数据"
      Top             =   120
      Width           =   7815
   End
   Begin VB.Frame Frame3 
      Caption         =   "其他选择"
      Height          =   2655
      Left            =   12120
      TabIndex        =   11
      ToolTipText     =   "我是其他选择"
      Top             =   2280
      Width           =   1335
      Begin VB.CommandButton END 
         Caption         =   "退出"
         Height          =   495
         Left            =   120
         TabIndex        =   94
         ToolTipText     =   "退出"
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "模式选择"
      Height          =   2055
      Left            =   12120
      TabIndex        =   12
      ToolTipText     =   "我是模式选择"
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
'  Copyright (c)   阿宽
'  File Name:      8583解析小工具
'  Author:         高建宽
'  Version:        V1.3
'  Date:           2014年12月26日
'  Description:    对8583字符处理
'  Function List():
'
'History: V1.0增加基本解析，位图显示与设置，悬停框架与按钮可以查看具体帮助
'         V1.1增加了普通解析，1.0版基本解析变为专家解析，增加响应码提示
'         V1.3增加55域详解
'
'
'Author:
'Modification:
'
'**************************************************************************/





Option Explicit

Dim bit_map_view_Click_flag(1 To 64) As Boolean    '// 位图按下标志
Dim change_begin As Integer    '//转换开始标志

Dim help_flag As Integer    '//定义帮助标志
Dim help_count As Long     '//定义按下帮助次数

Dim analyse_mode As Integer    '//定义解析模式 1 普通解析 2 专家解析 3 全面解析


Private Type PtrStruct
    ptr As String
    Ptrlen As Integer
End Type


'//定义8583POS中的结构体
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

    private_data_21 As PtrStruct    '//2015年4月23日14:28:28添加
    private_data_47 As PtrStruct    '//2015年6月18日10:51:05添加
End Type


'//解析区
Private Sub analyse_level_1_Click()
   On Error GoTo ErrHandle    '//捕获错误
    Dim tex2STR, tempStr, Totallen, TPDU, MessageHeader As String
    Dim tempCount As Long
    Dim POS_Sturct As POS_Sturct_TYPE
    Dim i As Integer

    PROCESS_Analyze_data    '//解析前处理

    If change_begin = 1 Then

        analyse_after_data.SetFocus
        '//   analyse_after_data.IMEMode = 3 '//当文本框得到焦点时禁用输入法
        analyse_mode = 1    '//定义解析模式 1 普通解析
        trace_no_11.Caption = "N/A"
        batch_Number_60.Caption = "N/A"

        help_flag = 0
        help_count = 0
        help.Caption = "说明"
        tempCount = 1    '//计数

        tempStr = analyse_after_data.Text
        '//       tex2STR = tex2STR & "/************8583解析结果（我是华丽的分割线开头）***************" & vbCrLf & vbCrLf


        If Totallen_check.Value = 1 Then
            '//总长解析
            Totallen = Mid(tempStr, tempCount, 4)
            Totallen = ins_space(Totallen)
            tex2STR = tex2STR & "[Totallen]" & "     " & Totallen & vbCrLf
            tempCount = tempCount + 4
        End If

        If TPDU_check.Value = 1 Then
            '//TPDU解析
            TPDU = Mid(tempStr, tempCount, 10)
            TPDU = ins_space(TPDU)
            tex2STR = tex2STR & "[TPDU]" & "         " & TPDU & vbCrLf
            tempCount = tempCount + 10
        End If

        If MessageHeader_check.Value = 1 Then
            '//MessageHeader解析
            MessageHeader = Mid(tempStr, tempCount, 12)
            MessageHeader = ins_space(MessageHeader)
            tex2STR = tex2STR & "[MessageHeader]" & "" & MessageHeader & vbCrLf
            tempCount = tempCount + 12
        End If


        '//messagetype解析
        POS_Sturct.messagetype = Mid(tempStr, tempCount, 4)
        POS_Sturct.messagetype = ins_space(POS_Sturct.messagetype)
        tex2STR = tex2STR & "[messagetype]" & "  " & POS_Sturct.messagetype & vbCrLf
        tempCount = tempCount + 4

        '//bitmap解析
        POS_Sturct.bitmap = Mid(tempStr, tempCount, 16)
        POS_Sturct.bitmap = ins_space(POS_Sturct.bitmap)
        tex2STR = tex2STR & "[bitmap]" & "       " & POS_Sturct.bitmap & vbCrLf
        tempCount = tempCount + 16
        bit_map.Text = POS_Sturct.bitmap
        bit_map_change_function

        Dim temp_bcd_flag_str As String
        Dim tempLen As Integer


        '//bitmap分离
        temp_bcd_flag_str = HEX_to_BIN(delete_space(POS_Sturct.bitmap))

        Dim tempPan As String
        Dim temp_Pan_len As String
        '//第2域数据 【主账号】
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

        '//第3域数据 【交易处理码】
        If Mid(temp_bcd_flag_str, 3, 1) Then
            POS_Sturct.procode_3 = Mid(tempStr, tempCount, 6)
            POS_Sturct.procode_3 = ins_space(POS_Sturct.procode_3)
            tex2STR = tex2STR & "[field3]" & "       " & POS_Sturct.procode_3 & vbCrLf
            tempCount = tempCount + 6
        End If
        '//类型判断
        i = Type_judge(POS_Sturct.messagetype, Mid(POS_Sturct.procode_3, 1, 2), Mid(temp_bcd_flag_str, 61, 1))

        Dim temp_consume_amount_4 As String
        '//第4域数据 【交易金额】
        If Mid(temp_bcd_flag_str, 4, 1) Then
            POS_Sturct.consume_amount_4 = Mid(tempStr, tempCount, 12)
            temp_consume_amount_4 = Val(POS_Sturct.consume_amount_4) / 100
            POS_Sturct.consume_amount_4 = ins_space(POS_Sturct.consume_amount_4)
            tex2STR = tex2STR & "[field4]" & "       " & POS_Sturct.consume_amount_4 & vbCrLf
            tempCount = tempCount + 12
        End If

        Dim temp_trace_no_11 As String
        '//第11域数据 【受卡方系统跟踪号】
        If Mid(temp_bcd_flag_str, 11, 1) Then
            POS_Sturct.trace_no_11 = Mid(tempStr, tempCount, 6)
            temp_trace_no_11 = POS_Sturct.trace_no_11
            POS_Sturct.trace_no_11 = ins_space(POS_Sturct.trace_no_11)
            tex2STR = tex2STR & "[field11]" & "      " & POS_Sturct.trace_no_11 & vbCrLf
            trace_no_11.Caption = temp_trace_no_11
            tempCount = tempCount + 6
        End If


        '//第12域数据 【受卡方所在地时间】
        If Mid(temp_bcd_flag_str, 12, 1) Then
            POS_Sturct.trade_time_12 = Mid(tempStr, tempCount, 6)
            POS_Sturct.trade_time_12 = ins_space(POS_Sturct.trade_time_12)
            tex2STR = tex2STR & "[field12]" & "      " & POS_Sturct.trade_time_12 & vbCrLf
            tempCount = tempCount + 6
        End If

        '//第13域数据 【受卡方所在地日期】
        If Mid(temp_bcd_flag_str, 13, 1) Then
            POS_Sturct.trade_date_13 = Mid(tempStr, tempCount, 4)
            POS_Sturct.trade_date_13 = ins_space(POS_Sturct.trade_date_13)
            tex2STR = tex2STR & "[field13]" & "      " & POS_Sturct.trade_date_13 & vbCrLf
            tempCount = tempCount + 4
        End If

        '//第14域数据 【卡有效期】
        If Mid(temp_bcd_flag_str, 14, 1) Then
            POS_Sturct.exp_date_14 = Mid(tempStr, tempCount, 4)
            POS_Sturct.exp_date_14 = ins_space(POS_Sturct.exp_date_14)
            tex2STR = tex2STR & "[field14]" & "      " & POS_Sturct.exp_date_14 & vbCrLf
            tempCount = tempCount + 4
        End If

        '//第15域数据 【清算日期】
        If Mid(temp_bcd_flag_str, 15, 1) Then
            POS_Sturct.settlement_date_15 = Mid(tempStr, tempCount, 4)
            POS_Sturct.settlement_date_15 = ins_space(POS_Sturct.settlement_date_15)
            tex2STR = tex2STR & "[field15]" & "      " & POS_Sturct.settlement_date_15 & vbCrLf
            tempCount = tempCount + 4
        End If

        Dim temp_21_len_str As String
        '//第21域数据 【private_data_21】
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

        '//第22域数据 【服务点输入方式码】
        If Mid(temp_bcd_flag_str, 22, 1) Then
            POS_Sturct.entry_mode_22 = Mid(tempStr, tempCount, 4)
            POS_Sturct.entry_mode_22 = ins_space(POS_Sturct.entry_mode_22)
            tex2STR = tex2STR & "[field22]" & "      " & POS_Sturct.entry_mode_22 & vbCrLf
            tempCount = tempCount + 4
        End If

        '//第23域数据 【卡序列号】
        If Mid(temp_bcd_flag_str, 23, 1) Then
            POS_Sturct.card_serial_number_23 = Mid(tempStr, tempCount, 4)
            POS_Sturct.card_serial_number_23 = ins_space(POS_Sturct.card_serial_number_23)
            tex2STR = tex2STR & "[field23]" & "      " & POS_Sturct.card_serial_number_23 & vbCrLf
            tempCount = tempCount + 4
        End If

        '//第25域数据 【服务点条件码】
        If Mid(temp_bcd_flag_str, 25, 1) Then
            POS_Sturct.service_conditon_25 = Mid(tempStr, tempCount, 2)
            POS_Sturct.service_conditon_25 = ins_space(POS_Sturct.service_conditon_25)
            tex2STR = tex2STR & "[field25]" & "      " & POS_Sturct.service_conditon_25 & vbCrLf
            tempCount = tempCount + 2
        End If

        '//第26域数据 【服务点PIN获取码】
        If Mid(temp_bcd_flag_str, 26, 1) Then
            POS_Sturct.service_conditon_pin_26 = Mid(tempStr, tempCount, 2)
            POS_Sturct.service_conditon_pin_26 = ins_space(POS_Sturct.service_conditon_pin_26)
            tex2STR = tex2STR & "[field26]" & "      " & POS_Sturct.service_conditon_pin_26 & vbCrLf
            tempCount = tempCount + 2
        End If

        Dim temp_32_len_str As String
        '//第32域数据 【受理方标识码】
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
        '//第35域数据 【2磁道数据】
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
        '//第36域数据 【3磁道数据】
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
        '//第37域数据 【检索参考号】
        If Mid(temp_bcd_flag_str, 37, 1) Then
            POS_Sturct.reference_number_37 = Mid(tempStr, tempCount, 24)
            POS_Sturct.reference_number_37 = ins_space(POS_Sturct.reference_number_37)
            temp_37str = ASCchange(POS_Sturct.reference_number_37)
            tex2STR = tex2STR & "[field37]" & "      " & POS_Sturct.reference_number_37 & vbCrLf
            tempCount = tempCount + 24
        End If

        Dim temp_38str As String
        '//第38域数据 【授权标识应答码】
        If Mid(temp_bcd_flag_str, 38, 1) Then
            POS_Sturct.authorization_code_38 = Mid(tempStr, tempCount, 12)
            POS_Sturct.authorization_code_38 = ins_space(POS_Sturct.authorization_code_38)
            temp_38str = ASCchange(POS_Sturct.authorization_code_38)
            tex2STR = tex2STR & "[field38]" & "      " & POS_Sturct.authorization_code_38 & vbCrLf
            tempCount = tempCount + 12
        End If

        Dim temp_39str As String
        '//第39域数据 【应答码】
        If Mid(temp_bcd_flag_str, 39, 1) Then
            POS_Sturct.Response_code_39 = Mid(tempStr, tempCount, 4)
            POS_Sturct.Response_code_39 = ins_space(POS_Sturct.Response_code_39)
            temp_39str = ASCchange(POS_Sturct.Response_code_39)
            tex2STR = tex2STR & "[field39]" & "      " & POS_Sturct.Response_code_39 & vbCrLf
            tempCount = tempCount + 4
        End If

        '//响应码判断显示
        i = Response_code_39_Type_judge(ASCchange(POS_Sturct.Response_code_39))

        Dim temp_41str As String
        '//第41域数据 【受卡机终端标识码】
        If Mid(temp_bcd_flag_str, 41, 1) Then
            POS_Sturct.terminal_no_41 = Mid(tempStr, tempCount, 16)
            POS_Sturct.terminal_no_41 = ins_space(POS_Sturct.terminal_no_41)
            temp_41str = ASCchange(POS_Sturct.terminal_no_41)
            tex2STR = tex2STR & "[field41]" & "      " & POS_Sturct.terminal_no_41 & vbCrLf
            tempCount = tempCount + 16
        End If

        Dim temp_42str As String
        '//第42域数据 【受卡方标识码】
        If Mid(temp_bcd_flag_str, 42, 1) Then
            POS_Sturct.merchant_no_42 = Mid(tempStr, tempCount, 30)
            POS_Sturct.merchant_no_42 = ins_space(POS_Sturct.merchant_no_42)
            temp_42str = ASCchange(POS_Sturct.merchant_no_42)
            tex2STR = tex2STR & "[field42]" & "      " & POS_Sturct.merchant_no_42 & vbCrLf
            tempCount = tempCount + 30
        End If


        Dim temp_43_len_str As String
        '//第43域数据 【merchant_name_43】
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
        Dim issuing_bank As String   '//发卡行
        Dim Acquiring_bank As String    '//收单行
        '//第44域数据 【附加响应数据】
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
        '//第46域数据 【pay_signature_46】
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
        '//第47域数据 【private_data_47】
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
        '//第48域数据 【附加数据 - 私有】
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
        '//第49域数据 【交易货币代码】
        If Mid(temp_bcd_flag_str, 49, 1) Then
            POS_Sturct.currency_code_49 = Mid(tempStr, tempCount, 6)
            POS_Sturct.currency_code_49 = ins_space(POS_Sturct.currency_code_49)
            tex2STR = tex2STR & "[field49]" & "      " & POS_Sturct.currency_code_49 & vbCrLf
            tempCount = tempCount + 6
        End If

        '//第52域数据 【个人标识码】
        If Mid(temp_bcd_flag_str, 52, 1) Then
            POS_Sturct.pri_pin_52 = Mid(tempStr, tempCount, 16)
            POS_Sturct.pri_pin_52 = ins_space(POS_Sturct.pri_pin_52)
            tex2STR = tex2STR & "[field52]" & "      " & POS_Sturct.pri_pin_52 & vbCrLf

            tempCount = tempCount + 16
        End If

        '//第53域数据 【安全控制信息】
        If Mid(temp_bcd_flag_str, 53, 1) Then
            POS_Sturct.safety_53 = Mid(tempStr, tempCount, 16)
            POS_Sturct.safety_53 = ins_space(POS_Sturct.safety_53)
            tex2STR = tex2STR & "[field53]" & "      " & POS_Sturct.safety_53 & vbCrLf

            tempCount = tempCount + 16
        End If

        Dim temp_54str As String
        Dim temp_54_len_str As String
        Dim temp_consume_amount_54 As String
        '//第54域数据 【余额】
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

        Dim POS_Sturct_icData_55_temp_ptr As String       '//55域临时数据
        Dim temp_55_str_ptr As Integer                     '//子域初始位置
        Dim temp_55_str_ptr_len As String                  '//子域长度
        '//第55域数据 【IC卡数据域】
        If Mid(temp_bcd_flag_str, 55, 1) Then
            temp_55_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.icData_55.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.icData_55.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.icData_55.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen


            '            POS_Sturct.icData_55.ptr = ins_space(POS_Sturct.icData_55.ptr)
            '//详解55域


            '//基本信息子域列表
            '形成这种格式 ->[9F 26] [08] 5E 14 AA 9F 20 46 A9 21   HEX_to_DEC
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F26")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
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
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F27")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F10")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F37")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F36")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "95")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9C")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F02")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "5F2A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "82")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))   '//去除搜索到的资源
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F1A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F03")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F33")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F74")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
                '//可选信息子域列表
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F34")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf

                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F35")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F1E")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "84")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F09")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf



                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F41")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "91")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "71")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "72")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF31")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F63")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            '//脱机交易专用子域列表
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "8A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "]    " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            '// 手机芯片交易专用子域列表
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF32")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF33")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf



                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF34")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf


                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            POS_Sturct.icData_55.ptr = POS_Sturct_icData_55_temp_ptr
            temp_55_len_str = ins_space(temp_55_len_str)
            tex2STR = tex2STR & "[field55]" & "      " & temp_55_len_str & vbCrLf & POS_Sturct.icData_55.ptr
            'Debug.Print tex2STR
        End If

        Dim temp_56_len_str As String
        '//第56域数据 【private_data_56】
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
        '//第57域数据 【private_data_57】
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
        '//第58域数据 【PBOC电子钱包标准的交易信息】
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
        '//第59域数据 【private_data_59】
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
        '//第60域数据 【private_data_60】
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
        Dim Original_batch_Number_61 As String    '//原始交易批次号
        Dim Original_trace_no_61 As String    '//原始交易POS流水号

        '//第61域数据 【原始信息域】
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
        '//第62域数据 【private_data_62】
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
        '//第63域数据 【private_data_63】
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

        '//第64域数据 【报文鉴别码】
        If Mid(temp_bcd_flag_str, 64, 1) Then
            POS_Sturct.mac_64 = Mid(tempStr, tempCount, 16)
            POS_Sturct.mac_64 = ins_space(POS_Sturct.mac_64)
            tex2STR = tex2STR & "[field64]" & "      " & POS_Sturct.mac_64 & vbCrLf
            tempCount = tempCount + 16
        End If



        '//  tex2STR = tex2STR & "************8583解析结果（我是华丽的分割线结尾）***************/" & vbCrLf
        analyse_after_data.Text = tex2STR
    End If

    Exit Sub  '一定要写
ErrHandle:
    'MSG_BOX Err.Number & Err.Source, "警告"
    ' Err.clear

    analyse_after_data.Text = ""
    bit_map_clear_Click
    trans_type.Caption = "N/A"
    judge_mode.Caption = "N/A"
    judge_mode.ForeColor = &H0&    '//恢复成黑色

    Response_code.Caption = "N/A"
    Response_code_view.Caption = "N/A"

    trace_no_11.Caption = "N/A"
    batch_Number_60.Caption = "N/A"

    help_flag = 0
    help.Caption = "说明"
    analyse_before_data.SetFocus


    MSG_BOX _
            "[Err.Source      ]:" & Err.Source & vbCrLf & _
                                  "[Err.Number      ]:" & Err.Number & vbCrLf & _
                                  "[Err.Description ]:" & Err.Description & vbCrLf, "警告"

    '"[Err.HelpContext ]:" & Err.HelpContext & vbCrLf & _
     '"[Err.HelpFile    ]:" & Err.HelpFile & vbCrLf & _
     '"[Err.LastDllError]:" & Err.LastDllError & vbCrLf & _

     Err.clear

End Sub


'//解析区
Private Sub analyse_level_2_Click()
     'On Error GoTo ErrHandle    '//捕获错误
    Dim tex2STR, tempStr, Totallen, TPDU, MessageHeader As String
    Dim tempCount As Long
    Dim POS_Sturct As POS_Sturct_TYPE
    Dim i As Integer

    PROCESS_Analyze_data    '//解析前处理

    If change_begin = 1 Then

        analyse_after_data.SetFocus
        analyse_mode = 2    '//定义解析模式 2 专家解析
        trace_no_11.Caption = "N/A"
        batch_Number_60.Caption = "N/A"
        help_flag = 0
        help_count = 0
        help.Caption = "说明"
        tempCount = 1    '//计数


        tempStr = analyse_after_data.Text
        '//       tex2STR = tex2STR & "/************8583解析结果（我是华丽的分割线开头）***************" & vbCrLf & vbCrLf

        If Totallen_check.Value = 1 Then
            '//总长解析

            Totallen = Mid(tempStr, tempCount, 4)
            Totallen = ins_space(Totallen)
            tex2STR = tex2STR & "[Totallen:总长度]" & vbCrLf & "" & Totallen & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        If TPDU_check.Value = 1 Then
            '//TPDU解析
            TPDU = Mid(tempStr, tempCount, 10)
            TPDU = ins_space(TPDU)
            tex2STR = tex2STR & "[TPDU:地址]" & vbCrLf & "" & TPDU & vbCrLf & vbCrLf
            tempCount = tempCount + 10
        End If

        If MessageHeader_check.Value = 1 Then
            '//MessageHeader解析
            MessageHeader = Mid(tempStr, tempCount, 12)
            MessageHeader = ins_space(MessageHeader)
            tex2STR = tex2STR & "[MessageHeader:报文头]" & vbCrLf & "" & MessageHeader & vbCrLf & vbCrLf
            tempCount = tempCount + 12
        End If


        '//messagetype解析
        POS_Sturct.messagetype = Mid(tempStr, tempCount, 4)
        POS_Sturct.messagetype = ins_space(POS_Sturct.messagetype)
        tex2STR = tex2STR & "[messagetype:消息类型]" & vbCrLf & "" & POS_Sturct.messagetype & vbCrLf & vbCrLf
        tempCount = tempCount + 4

        '//bitmap解析
        POS_Sturct.bitmap = Mid(tempStr, tempCount, 16)
        POS_Sturct.bitmap = ins_space(POS_Sturct.bitmap)
        tex2STR = tex2STR & "[bitmap:位元表]" & vbCrLf & "" & POS_Sturct.bitmap & vbCrLf & vbCrLf
        tempCount = tempCount + 16
        bit_map.Text = POS_Sturct.bitmap
        bit_map_change_function

        Dim temp_bcd_flag_str As String
        Dim tempLen As Integer


        '//bitmap分离
        temp_bcd_flag_str = HEX_to_BIN(delete_space(POS_Sturct.bitmap))

        Dim tempPan As String
        Dim temp_Pan_len As String
        '//第2域数据 【主账号】
        If Mid(temp_bcd_flag_str, 2, 1) Then
            temp_Pan_len = Mid(tempStr, tempCount, 2)
            POS_Sturct.pan_2.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.pan_2.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.pan_2.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            tempPan = Mid(POS_Sturct.pan_2.ptr, 1, POS_Sturct.pan_2.Ptrlen)
            POS_Sturct.pan_2.ptr = ins_space(POS_Sturct.pan_2.ptr)
            tex2STR = tex2STR & "[field2:主账号(Primary Account Number)]" & vbCrLf & "" & "[" & temp_Pan_len & "]" & " " & POS_Sturct.pan_2.ptr _
                      & vbCrLf & "->[卡号] " & tempPan & vbCrLf & vbCrLf
        End If

        '//第3域数据 【交易处理码】
        If Mid(temp_bcd_flag_str, 3, 1) Then
            POS_Sturct.procode_3 = Mid(tempStr, tempCount, 6)
            POS_Sturct.procode_3 = ins_space(POS_Sturct.procode_3)
            tex2STR = tex2STR & "[field3:交易处理码(Transaction Processing Code)]" & vbCrLf & "" & POS_Sturct.procode_3 & vbCrLf & vbCrLf
            tempCount = tempCount + 6
        End If
        '//类型判断
        i = Type_judge(POS_Sturct.messagetype, Mid(POS_Sturct.procode_3, 1, 2), Mid(temp_bcd_flag_str, 61, 1))

        Dim temp_consume_amount_4 As String
        '//第4域数据 【交易金额】
        If Mid(temp_bcd_flag_str, 4, 1) Then
            POS_Sturct.consume_amount_4 = Mid(tempStr, tempCount, 12)
            temp_consume_amount_4 = Val(POS_Sturct.consume_amount_4) / 100

            If Val(temp_consume_amount_4) < 1 And Val(temp_consume_amount_4) > 0 Then
                temp_consume_amount_4 = 0 & temp_consume_amount_4
            End If


            POS_Sturct.consume_amount_4 = ins_space(POS_Sturct.consume_amount_4)
            tex2STR = tex2STR & "[field4:交易金额(Amount Of Transactions)]" & vbCrLf & "" & POS_Sturct.consume_amount_4 & _
                      vbCrLf & "->[￥金额] " & temp_consume_amount_4 & "元" & vbCrLf & vbCrLf
            tempCount = tempCount + 12
        End If

        Dim temp_trace_no_11 As String
        '//第11域数据 【受卡方系统跟踪号】
        If Mid(temp_bcd_flag_str, 11, 1) Then
            POS_Sturct.trace_no_11 = Mid(tempStr, tempCount, 6)
            temp_trace_no_11 = POS_Sturct.trace_no_11
            POS_Sturct.trace_no_11 = ins_space(POS_Sturct.trace_no_11)
            tex2STR = tex2STR & "[field11:受卡方系统跟踪号(System Trace Audit Number)]" & vbCrLf & "" & POS_Sturct.trace_no_11 & _
                      vbCrLf & "->[流水号] " & temp_trace_no_11 & vbCrLf & vbCrLf

            trace_no_11.Caption = temp_trace_no_11

            tempCount = tempCount + 6
        End If


        '//第12域数据 【受卡方所在地时间】
        If Mid(temp_bcd_flag_str, 12, 1) Then
            POS_Sturct.trade_time_12 = Mid(tempStr, tempCount, 6)
            POS_Sturct.trade_time_12 = ins_space(POS_Sturct.trade_time_12)
            tex2STR = tex2STR & "[field12:受卡方所在地时间(Local Time Of Transaction)]" & vbCrLf & "" & POS_Sturct.trade_time_12 & _
                      vbCrLf & "->[时间] " & Mid(POS_Sturct.trade_time_12, 1, 2) & "时" & Mid(POS_Sturct.trade_time_12, 4, 2) & "分" _
                      & Mid(POS_Sturct.trade_time_12, 7, 2) & "秒" & vbCrLf & vbCrLf
            tempCount = tempCount + 6
        End If

        '//第13域数据 【受卡方所在地日期】
        If Mid(temp_bcd_flag_str, 13, 1) Then
            POS_Sturct.trade_date_13 = Mid(tempStr, tempCount, 4)
            POS_Sturct.trade_date_13 = ins_space(POS_Sturct.trade_date_13)
            tex2STR = tex2STR & "[field13:受卡方所在地日期(Local Date Of Transaction)]" & vbCrLf & "" & POS_Sturct.trade_date_13 & "   " & _
                      vbCrLf & "->[日期] " & Mid(POS_Sturct.trade_date_13, 1, 2) & "月" & Mid(POS_Sturct.trade_date_13, 4, 2) & "日" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//第14域数据 【卡有效期】
        If Mid(temp_bcd_flag_str, 14, 1) Then
            POS_Sturct.exp_date_14 = Mid(tempStr, tempCount, 4)
            POS_Sturct.exp_date_14 = ins_space(POS_Sturct.exp_date_14)
            tex2STR = tex2STR & "[field14:卡有效期(Date Of Expired)]" & vbCrLf & "" & POS_Sturct.exp_date_14 & "   " & _
                      vbCrLf & "->[有效期] " & "20" & Mid(POS_Sturct.exp_date_14, 1, 2) & "年" & Mid(POS_Sturct.exp_date_14, 4, 2) & "月" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//第15域数据 【清算日期】
        If Mid(temp_bcd_flag_str, 15, 1) Then
            POS_Sturct.settlement_date_15 = Mid(tempStr, tempCount, 4)
            POS_Sturct.settlement_date_15 = ins_space(POS_Sturct.settlement_date_15)
            tex2STR = tex2STR & "[field15:清算日期(Date Of Settlement)]" & vbCrLf & "" & POS_Sturct.settlement_date_15 & "   " & _
                      vbCrLf & "->[清算日期] " & Mid(POS_Sturct.settlement_date_15, 1, 2) & "月" & Mid(POS_Sturct.settlement_date_15, 4, 2) & "日" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If
        Dim temp_21_len_str As String
        '//第21域数据 【private_data_21】
        If Mid(temp_bcd_flag_str, 21, 1) Then
            temp_21_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_21.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_21.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_21.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_21.ptr = ins_space(POS_Sturct.private_data_21.ptr)
            temp_21_len_str = ins_space(temp_21_len_str)
            tex2STR = tex2STR & "[field21:自定义域(private_data_21)]" & vbCrLf & "" & "[" & temp_21_len_str & "]" & "  " & POS_Sturct.private_data_21.ptr & vbCrLf & vbCrLf

        End If

        '//第22域数据 【服务点输入方式码】
        If Mid(temp_bcd_flag_str, 22, 1) Then
            POS_Sturct.entry_mode_22 = Mid(tempStr, tempCount, 4)
            POS_Sturct.entry_mode_22 = ins_space(POS_Sturct.entry_mode_22)
            tex2STR = tex2STR & "[field22:服务点输入方式码(Point Of Service Entry Mode)]" & vbCrLf & "" & POS_Sturct.entry_mode_22 & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '        '//第23域数据 【卡序列号】
        '        If Mid(temp_bcd_flag_str, 23, 1) Then
        '            POS_Sturct.card_serial_number_23 = Mid(tempStr, tempCount, 4)
        '            POS_Sturct.card_serial_number_23 = ins_space(POS_Sturct.card_serial_number_23)
        '            tex2STR = tex2STR & "[field23:卡序列号(Card Sequence Number)]" & vbCrLf & "" & POS_Sturct.card_serial_number_23 & vbCrLf & vbCrLf
        '            tempCount = tempCount + 4
        '        End If









        Dim Right_23_str As String    '//卡片序列号右靠
        Dim left_23_str As String    '//卡片序列号左靠
        '//第23域数据 【卡序列号】
        If Mid(temp_bcd_flag_str, 23, 1) Then
            POS_Sturct.card_serial_number_23 = Mid(tempStr, tempCount, 4)
            Right_23_str = Mid(POS_Sturct.card_serial_number_23, 2, 3)
            left_23_str = Mid(POS_Sturct.card_serial_number_23, 1, 3)
            POS_Sturct.card_serial_number_23 = ins_space(POS_Sturct.card_serial_number_23)
            tex2STR = tex2STR & "[field23:卡序列号(Card Sequence Number)]" & vbCrLf & "" & POS_Sturct.card_serial_number_23 & vbCrLf _
                      & "->[卡片序列号左靠] " & left_23_str & vbCrLf & "->[卡片序列号右靠] " & Right_23_str & " [一般优先选择右靠]" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If


        '//第25域数据 【服务点条件码】
        If Mid(temp_bcd_flag_str, 25, 1) Then
            POS_Sturct.service_conditon_25 = Mid(tempStr, tempCount, 2)
            POS_Sturct.service_conditon_25 = ins_space(POS_Sturct.service_conditon_25)
            tex2STR = tex2STR & "[field25:服务点条件码(Point Of Service Condition Mode)]" & vbCrLf & "" & POS_Sturct.service_conditon_25 & vbCrLf & vbCrLf
            tempCount = tempCount + 2
        End If

        '//第26域数据 【服务点PIN获取码】
        If Mid(temp_bcd_flag_str, 26, 1) Then
            POS_Sturct.service_conditon_pin_26 = Mid(tempStr, tempCount, 2)
            POS_Sturct.service_conditon_pin_26 = ins_space(POS_Sturct.service_conditon_pin_26)
            tex2STR = tex2STR & "[field26:服务点PIN获取码(Point Of Service PIN Capture Code)]" & vbCrLf & "" & POS_Sturct.service_conditon_pin_26 & vbCrLf & vbCrLf
            tempCount = tempCount + 2
        End If

        Dim temp_32_len_str As String
        '//第32域数据 【受理方标识码】
        If Mid(temp_bcd_flag_str, 32, 1) Then
            temp_32_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.api_code_32.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.api_code_32.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.api_code_32.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.api_code_32.ptr = ins_space(POS_Sturct.api_code_32.ptr)

            tex2STR = tex2STR & "[field32:受理方标识码(Acquiring Institution Id Code)]" & vbCrLf & "" & "[" & temp_32_len_str & "]" & " " & POS_Sturct.api_code_32.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_35_len_str As String
        '//第35域数据 【2磁道数据】
        If Mid(temp_bcd_flag_str, 35, 1) Then
            temp_35_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.track2_35.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.track2_35.Ptrlen + 1) \ 2) * 2
            '// tempLen = POS_Sturct.track2_35.Ptrlen * 2
            tempCount = tempCount + 2

            POS_Sturct.track2_35.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.track2_35.ptr = ins_space(POS_Sturct.track2_35.ptr)

            tex2STR = tex2STR & "[field35:2磁道数据(Track 2 Data)]" & vbCrLf & "" & "[" & temp_35_len_str & "]" & "  " & POS_Sturct.track2_35.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_36_len_str As String
        '//第36域数据 【3磁道数据】
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
            tex2STR = tex2STR & "[field36:3磁道数据(Track 3 Data)]" & vbCrLf & "" & "[" & temp_36_len_str & "]" & "  " & POS_Sturct.track3_36.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_37str As String
        '//第37域数据 【检索参考号】
        If Mid(temp_bcd_flag_str, 37, 1) Then
            POS_Sturct.reference_number_37 = Mid(tempStr, tempCount, 24)
            POS_Sturct.reference_number_37 = ins_space(POS_Sturct.reference_number_37)
            temp_37str = ASCchange(POS_Sturct.reference_number_37)
            tex2STR = tex2STR & "[field37:检索参考号(Retrieval Reference Number)]" & vbCrLf & "" & POS_Sturct.reference_number_37 & vbCrLf & "->[参考号] " & temp_37str & vbCrLf & vbCrLf
            tempCount = tempCount + 24
        End If

        Dim temp_38str As String
        '//第38域数据 【授权标识应答码】
        If Mid(temp_bcd_flag_str, 38, 1) Then
            POS_Sturct.authorization_code_38 = Mid(tempStr, tempCount, 12)
            POS_Sturct.authorization_code_38 = ins_space(POS_Sturct.authorization_code_38)
            temp_38str = ASCchange(POS_Sturct.authorization_code_38)
            tex2STR = tex2STR & "[field38:授权标识应答码(Authorization Id Response Code)]" & vbCrLf & "" & POS_Sturct.authorization_code_38 & vbCrLf & "->" & "[授权码] " & temp_38str & vbCrLf & vbCrLf
            tempCount = tempCount + 12
        End If

        Dim temp_39str As String
        '//第39域数据 【应答码】
        If Mid(temp_bcd_flag_str, 39, 1) Then
            POS_Sturct.Response_code_39 = Mid(tempStr, tempCount, 4)
            POS_Sturct.Response_code_39 = ins_space(POS_Sturct.Response_code_39)
            temp_39str = ASCchange(POS_Sturct.Response_code_39)
            tex2STR = tex2STR & "[field39:应答码(Response Code)]" & vbCrLf & "" & POS_Sturct.Response_code_39 & vbCrLf & "->" & "[应答码:] " & temp_39str & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//响应码判断显示
        i = Response_code_39_Type_judge(ASCchange(POS_Sturct.Response_code_39))

        Dim temp_41str As String
        '//第41域数据 【受卡机终端标识码】
        If Mid(temp_bcd_flag_str, 41, 1) Then
            POS_Sturct.terminal_no_41 = Mid(tempStr, tempCount, 16)
            POS_Sturct.terminal_no_41 = ins_space(POS_Sturct.terminal_no_41)
            temp_41str = ASCchange(POS_Sturct.terminal_no_41)
            tex2STR = tex2STR & "[field41:受卡机终端标识码(Card Acceptor Terminal Id)]" & vbCrLf & "" & POS_Sturct.terminal_no_41 & vbCrLf & "->[终端号] " & temp_41str & vbCrLf & vbCrLf
            tempCount = tempCount + 16
        End If

        Dim temp_42str As String
        '//第42域数据 【受卡方标识码】
        If Mid(temp_bcd_flag_str, 42, 1) Then
            POS_Sturct.merchant_no_42 = Mid(tempStr, tempCount, 30)
            POS_Sturct.merchant_no_42 = ins_space(POS_Sturct.merchant_no_42)
            temp_42str = ASCchange(POS_Sturct.merchant_no_42)
            tex2STR = tex2STR & "[field42:受卡方标识码(Card Acceptor Id Code)]" & vbCrLf & "" & POS_Sturct.merchant_no_42 & vbCrLf & "->[商户号] " & temp_42str & vbCrLf & vbCrLf
            tempCount = tempCount + 30
        End If


        Dim temp_43_len_str As String
        '//第43域数据 【merchant_name_43】
        If Mid(temp_bcd_flag_str, 43, 1) Then
            temp_43_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.merchant_name_43.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.merchant_name_43.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.merchant_name_43.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.merchant_name_43.ptr = ins_space(POS_Sturct.merchant_name_43.ptr)
            temp_43_len_str = ins_space(temp_43_len_str)
            tex2STR = tex2STR & "[field43:自定义域(merchant_name_43)]" & vbCrLf & "" & "[" & temp_43_len_str & "]" & "  " & POS_Sturct.merchant_name_43.ptr & vbCrLf & vbCrLf

        End If


        Dim temp_44_len_str As String
        Dim issuing_bank As String   '//发卡行
        Dim Acquiring_bank As String    '//收单行
        '//第44域数据 【附加响应数据】
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
                tex2STR = tex2STR & "[field44:附加响应数据(Additional Response Data)]" & vbCrLf & "" & "[" & temp_44_len_str & "]" & "  " & POS_Sturct.rsp_code_44.ptr _
                          & vbCrLf & "->[发卡行] " & issuing_bank & vbCrLf & "->[收单行] " & Acquiring_bank & vbCrLf & vbCrLf
            Else
                tex2STR = tex2STR & "[field44:附加响应数据(Additional Response Data)]" & vbCrLf & "" & "[" & temp_44_len_str & "]" & "  " & POS_Sturct.rsp_code_44.ptr & vbCrLf & vbCrLf
            End If
        End If

        Dim temp_46_len_str As String
        '//第46域数据 【pay_signature_46】
        If Mid(temp_bcd_flag_str, 46, 1) Then
            temp_46_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.pay_signature_46.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.pay_signature_46.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.pay_signature_46.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.pay_signature_46.ptr = ins_space(POS_Sturct.pay_signature_46.ptr)
            temp_46_len_str = ins_space(temp_46_len_str)
            tex2STR = tex2STR & "[field46:自定义域(pay_signature_46)]" & vbCrLf & "" & "[" & temp_46_len_str & "]" & "  " & POS_Sturct.pay_signature_46.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_47_len_str As String
        '//第47域数据 【private_data_47】
        If Mid(temp_bcd_flag_str, 47, 1) Then
            temp_47_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_47.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_47.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_47.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_47.ptr = ins_space(POS_Sturct.private_data_47.ptr)
            temp_47_len_str = ins_space(temp_47_len_str)
            tex2STR = tex2STR & "[field47:自定义域(private_data_47)]" & vbCrLf & "" & "[" & temp_47_len_str & "]" & "  " & POS_Sturct.private_data_47.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_48_len_str As String
        '//第48域数据 【附加数据 - 私有】
        If Mid(temp_bcd_flag_str, 48, 1) Then
            temp_48_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.settleAccounts_48.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.settleAccounts_48.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.settleAccounts_48.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.settleAccounts_48.ptr = ins_space(POS_Sturct.settleAccounts_48.ptr)
            temp_48_len_str = ins_space(temp_48_len_str)
            tex2STR = tex2STR & "[field48:附加数据 - 私有(Additional Data - Private)]" & vbCrLf & "" & "[" & temp_48_len_str & "]" & "  " & POS_Sturct.settleAccounts_48.ptr & vbCrLf & vbCrLf
        End If

        Dim temp_49str As String
        '//第49域数据 【交易货币代码】
        If Mid(temp_bcd_flag_str, 49, 1) Then
            POS_Sturct.currency_code_49 = Mid(tempStr, tempCount, 6)
            POS_Sturct.currency_code_49 = ins_space(POS_Sturct.currency_code_49)
            temp_49str = ASCchange(POS_Sturct.currency_code_49)

            tex2STR = tex2STR & "[field49:交易货币代码(Currency Code Of Transaction)]" & vbCrLf & "" & POS_Sturct.currency_code_49 & vbCrLf & vbCrLf

            tempCount = tempCount + 6
        End If

        '//第52域数据 【个人标识码】
        If Mid(temp_bcd_flag_str, 52, 1) Then
            POS_Sturct.pri_pin_52 = Mid(tempStr, tempCount, 16)
            POS_Sturct.pri_pin_52 = ins_space(POS_Sturct.pri_pin_52)
            tex2STR = tex2STR & "[field52:个人标识码数据(PIN Data)]" & vbCrLf & "" & POS_Sturct.pri_pin_52 & vbCrLf & vbCrLf

            tempCount = tempCount + 16
        End If

        '//第53域数据 【安全控制信息】
        If Mid(temp_bcd_flag_str, 53, 1) Then
            POS_Sturct.safety_53 = Mid(tempStr, tempCount, 16)
            POS_Sturct.safety_53 = ins_space(POS_Sturct.safety_53)
            tex2STR = tex2STR & "[field53:安全控制信息(Security Related Control Information )]" & vbCrLf & "" & POS_Sturct.safety_53 & vbCrLf & vbCrLf

            tempCount = tempCount + 16
        End If

        Dim temp_54str As String
        Dim temp_54_len_str As String
        Dim temp_consume_amount_54 As String
        '//第54域数据 【余额】
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
            tex2STR = tex2STR & "[field54:余额(Balanc Amount)]" & vbCrLf & "" & "[" & temp_54_len_str & "]" & "  " & POS_Sturct.attachment_amount_54.ptr & vbCrLf & "->[ASCII转换] " _
                      & temp_54str & vbCrLf & "->[金额]      " & temp_consume_amount_54 & "元" & vbCrLf & vbCrLf
        End If

        Dim temp_55_len_str As String
        Dim POS_Sturct_icData_55_temp_ptr As String       '//55域临时数据
        Dim temp_55_str_ptr As Integer                     '//子域初始位置
        Dim temp_55_str_ptr_len As String                  '//子域长度
        '//第55域数据 【IC卡数据域】
        If Mid(temp_bcd_flag_str, 55, 1) Then
            temp_55_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.icData_55.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.icData_55.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.icData_55.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            '            POS_Sturct.icData_55.ptr = ins_space(POS_Sturct.icData_55.ptr)
            '//详解55域


            '//基本信息子域列表
            '形成这种格式 ->[9F 26][应用密文]
            '               [08] 5E 14 AA 9F 20 46 A9 21
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F26")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[应用密文]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F27")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[密文信息数据]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F10")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[发卡行应用数据]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F37")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[不可预知数]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                      POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F36")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[应用交易计数器]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "95")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [终端验证结果]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [交易日期]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9C")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [交易类型]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F02")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[授权金额]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "5F2A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[交易货币代码]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "82")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [应用交互特征]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If


            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F1A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[终端国家代码]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F03")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[其它金额]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F33")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[终端性能]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F74")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[电子现金发卡行授权码]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If


            '//可选信息子域列表

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F34")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[持卡人验证方法结果]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F35")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[终端类型]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F1E")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[接口设备序列号]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "84")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [专用文件名称]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F09")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[软件版本号]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F41")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[交易序列计数器]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "91")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [发卡行认证数据]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "71")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [发卡行脚本 1]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "72")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [发卡行脚本 2]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF31")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[发卡方脚本结果]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "9F63")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[卡产品标识信息]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            '//脱机交易专用子域列表

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "8A")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 2, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 2)) & "] " & "   [授权响应码]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 2 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            '// 手机芯片交易专用子域列表

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF32")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[芯片序列号]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If

            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF33")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[过程密钥数据]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If
            temp_55_str_ptr = InStr(1, POS_Sturct.icData_55.ptr, "DF34")
            If (temp_55_str_ptr <> 0 And temp_55_str_ptr Mod 2 = 1) Then  '// 不为空，且为奇数【防止找到像例子情况】 例子 08 A0
                temp_55_str_ptr_len = Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 4, 2)
                POS_Sturct_icData_55_temp_ptr = POS_Sturct_icData_55_temp_ptr & "->[" & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr, 4)) & "] " & "[磁道读取时间]" & vbCrLf _
                                                & "[" & temp_55_str_ptr_len & "] " & ins_space(Mid(POS_Sturct.icData_55.ptr, temp_55_str_ptr + 6, 2 * HEX_to_DEC(temp_55_str_ptr_len))) & vbCrLf
                POS_Sturct.icData_55.ptr = Left(POS_Sturct.icData_55.ptr, temp_55_str_ptr - 1) & Right(POS_Sturct.icData_55.ptr, 1 + Len(POS_Sturct.icData_55.ptr) - (temp_55_str_ptr + 4 + 2 + 2 * HEX_to_DEC(temp_55_str_ptr_len)))    '//去除搜索到的资源
            End If



            POS_Sturct.icData_55.ptr = POS_Sturct_icData_55_temp_ptr
            temp_55_len_str = ins_space(temp_55_len_str)
            tex2STR = tex2STR & "[field55:IC卡数据域(IC Card System Related Data)]" & vbCrLf & "" & "[" & temp_55_len_str & "]" & vbCrLf & POS_Sturct.icData_55.ptr & vbCrLf

        End If

        Dim temp_56_len_str As String
        '//第56域数据 【private_data_56】
        If Mid(temp_bcd_flag_str, 56, 1) Then
            temp_56_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_56.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_56.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_56.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_56.ptr = ins_space(POS_Sturct.private_data_56.ptr)
            temp_56_len_str = ins_space(temp_56_len_str)
            tex2STR = tex2STR & "[field56:自定义域(private_data_56)]" & vbCrLf & "" & "[" & temp_56_len_str & "]" & "  " & POS_Sturct.private_data_56.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_57_len_str As String
        '//第57域数据 【private_data_57】
        If Mid(temp_bcd_flag_str, 57, 1) Then
            temp_57_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_57.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_57.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_57.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_57.ptr = ins_space(POS_Sturct.private_data_57.ptr)
            temp_57_len_str = ins_space(temp_57_len_str)
            tex2STR = tex2STR & "[field57:自定义域(private_data_57)]" & vbCrLf & "" & "[" & temp_57_len_str & "]" & "  " & POS_Sturct.private_data_57.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_58_len_str As String
        '//第58域数据 【PBOC电子钱包标准的交易信息】
        If Mid(temp_bcd_flag_str, 58, 1) Then
            temp_58_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_58.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_58.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_58.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_58.ptr = ins_space(POS_Sturct.private_data_58.ptr)
            temp_58_len_str = ins_space(temp_58_len_str)
            tex2STR = tex2STR & "[field58:PBOC电子钱包标准的交易信息(PBOC_ELECTRONIC_DATA)]" & vbCrLf & "" & "[" & temp_58_len_str & "]" & "  " & POS_Sturct.private_data_58.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_59_len_str As String
        '//第59域数据 【private_data_59】
        If Mid(temp_bcd_flag_str, 59, 1) Then
            temp_59_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_59.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_59.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_59.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_59.ptr = ins_space(POS_Sturct.private_data_59.ptr)
            temp_59_len_str = ins_space(temp_59_len_str)
            tex2STR = tex2STR & "[field59:自定义域(private_data_59)]" & vbCrLf & "" & "[" & temp_59_len_str & "]" & "  " & POS_Sturct.private_data_59.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_60_len_str As String
        Dim temp_batch_Number_60 As String
        '//第60域数据 【private_data_60】
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
            tex2STR = tex2STR & "[field60:自定义域(private_data_60)]" & vbCrLf & "" & "[" & temp_60_len_str & "]" & "  " & POS_Sturct.private_data_60.ptr _
                      & vbCrLf & "->[批次号] " & temp_batch_Number_60 & vbCrLf & vbCrLf

            batch_Number_60.Caption = temp_batch_Number_60

        End If

        Dim temp_61_len_str As String
        Dim Original_batch_Number_61 As String    '//原始交易批次号
        Dim Original_trace_no_61 As String    '//原始交易POS流水号

        '//第61域数据 【原始信息域】
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
            tex2STR = tex2STR & "[field61:原始信息域(Original Message)]" & vbCrLf & "" & "[" & temp_61_len_str & "]" & "  " & POS_Sturct.private_data_61.ptr _
                      & vbCrLf & "->[原始交易批次号]    " & Original_batch_Number_61 & vbCrLf & "->[原始交易POS流水号] " & Original_trace_no_61 & vbCrLf & vbCrLf

        End If

        Dim temp_62_len_str As String
        '//第62域数据 【private_data_62】
        If Mid(temp_bcd_flag_str, 62, 1) Then
            temp_62_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_62.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_62.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_62.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_62.ptr = ins_space(POS_Sturct.private_data_62.ptr)
            temp_62_len_str = ins_space(temp_62_len_str)
            tex2STR = tex2STR & "[field62:自定义域(private_data_62)]" & vbCrLf & "" & "[" & temp_62_len_str & "]" & "  " & POS_Sturct.private_data_62.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_63_len_str As String
        '//第63域数据 【private_data_63】
        If Mid(temp_bcd_flag_str, 63, 1) Then
            temp_63_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_63.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_63.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_63.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_63.ptr = ins_space(POS_Sturct.private_data_63.ptr)
            temp_63_len_str = ins_space(temp_63_len_str)
            tex2STR = tex2STR & "[field63:自定义域(private_data_63)]" & vbCrLf & "" & "[" & temp_63_len_str & "]" & "  " & POS_Sturct.private_data_63.ptr & vbCrLf & vbCrLf

        End If

        '//第64域数据 【报文鉴别码】
        If Mid(temp_bcd_flag_str, 64, 1) Then
            POS_Sturct.mac_64 = Mid(tempStr, tempCount, 16)
            POS_Sturct.mac_64 = ins_space(POS_Sturct.mac_64)
            tex2STR = tex2STR & "[field64:报文鉴别码(Message Authentication Code)]" & vbCrLf & "" & POS_Sturct.mac_64 & vbCrLf & vbCrLf
            tempCount = tempCount + 16
        End If



        '//  tex2STR = tex2STR & "************8583解析结果（我是华丽的分割线结尾）***************/" & vbCrLf
        analyse_after_data.Text = tex2STR
    End If
    Exit Sub  '一定要写
ErrHandle:
    'MSG_BOX Err.Number & Err.Source, "警告"
    ' Err.clear

    analyse_after_data.Text = ""
    bit_map_clear_Click
    trans_type.Caption = "N/A"
    judge_mode.Caption = "N/A"
    judge_mode.ForeColor = &H0&    '//恢复成黑色

    Response_code.Caption = "N/A"
    Response_code_view.Caption = "N/A"

    trace_no_11.Caption = "N/A"
    batch_Number_60.Caption = "N/A"

    help_flag = 0
    help.Caption = "说明"
    analyse_before_data.SetFocus


    MSG_BOX _
            "[Err.Source      ]:" & Err.Source & vbCrLf & _
                                  "[Err.Number      ]:" & Err.Number & vbCrLf & _
                                  "[Err.Description ]:" & Err.Description & vbCrLf, "警告"

    '"[Err.HelpContext ]:" & Err.HelpContext & vbCrLf & _
     '"[Err.HelpFile    ]:" & Err.HelpFile & vbCrLf & _
     '"[Err.LastDllError]:" & Err.LastDllError & vbCrLf & _

     Err.clear

End Sub

'//解析区
Private Sub analyse_level_3_Click()

    Dim tex2STR, tempStr, Totallen, TPDU, MessageHeader As String
    Dim tempCount As Long
    Dim POS_Sturct As POS_Sturct_TYPE
    Dim i As Integer

    PROCESS_Analyze_data    '//解析前处理

    If change_begin = 1 Then
        analyse_after_data.SetFocus
        analyse_mode = 3    '//定义解析模式 3 全面解析
        trace_no_11.Caption = "N/A"
        batch_Number_60.Caption = "N/A"
        help_flag = 0
        help_count = 0
        help.Caption = "说明"
        tempCount = 1    '//计数


        tempStr = analyse_after_data.Text
        '//       tex2STR = tex2STR & "/************8583解析结果（我是华丽的分割线开头）***************" & vbCrLf & vbCrLf

        If Totallen_check.Value = 1 Then
            '//总长解析
            Dim temp_TotallenStr As String

            Totallen = Mid(tempStr, tempCount, 4)
            temp_TotallenStr = HEX_to_DEC(Mid(Totallen, 1, 2)) * 16 * 16 + HEX_to_DEC(Mid(Totallen, 3, 4))

            Totallen = ins_space(Totallen)
            tex2STR = tex2STR & "[Totallen:总长度]" & vbCrLf & "" & Totallen & "         " & vbCrLf & _
                      "->[十进制总长度]" & " " & temp_TotallenStr & "+2=共 " & temp_TotallenStr + 2 & " 字节" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        If TPDU_check.Value = 1 Then
            '//TPDU解析
            Dim TPDU_ID, TPDU_DEST_ADR, TPDU_SRC_ADR As String


            TPDU = Mid(tempStr, tempCount, 10)
            TPDU_ID = ins_space(Mid(TPDU, 1, 2))
            TPDU_DEST_ADR = ins_space(Mid(TPDU, 3, 4))
            TPDU_SRC_ADR = ins_space(Mid(TPDU, 7, 4))


            TPDU = ins_space(TPDU)
            tex2STR = tex2STR & "[TPDU:地址]" & vbCrLf & TPDU & vbCrLf _
                      & "------------" & vbCrLf & "->[ID]" & vbCrLf & TPDU_ID & vbCrLf & "------------" & vbCrLf & "->[目的地址]" & vbCrLf & TPDU_DEST_ADR & vbCrLf & "------------" & vbCrLf & "->[源地址]" & vbCrLf & TPDU_SRC_ADR & vbCrLf & "------------" & vbCrLf & vbCrLf
            tempCount = tempCount + 10
        End If

        If MessageHeader_check.Value = 1 Then

            Dim MessageHeader_App_Type As String    '// 应用类别定义
            Dim MessageHeader_App_Type_Str As String    '// 应用类别定义显示

            Dim MessageHeader_Software_Total_Ver_Num As String    '// 软件总版本号
            Dim MessageHeader_Software_Total_Ver_Num_Str As String    '// 软件总版本号显示

            Dim MessageHeader_Terminal_State As String   '// 终端状态
            Dim MessageHeader_Terminal_State_Str As String   '// 终端状态显示

            Dim MessageHeader_Process_Require As String    '// 处理要求
            Dim MessageHeader_Process_Require_Str As String   '// 终端状态显示

            Dim MessageHeader_Software_Part_Ver_Num As String    '// 软件分版本号
            Dim MessageHeader_Software_Part_Ver_Num_Str As String    '// 软件分版本号显示


            '//MessageHeader解析
            MessageHeader = Mid(tempStr, tempCount, 12)

            '//应用类别定义解析
            MessageHeader_App_Type = Mid(MessageHeader, 1, 2)

            Select Case MessageHeader_App_Type
            Case 60
                MessageHeader_App_Type_Str = "磁条卡金融支付类应用"
            Case 61
                MessageHeader_App_Type_Str = "IC卡金融支付类应用"
            Case 62
                MessageHeader_App_Type_Str = "磁条卡增值业务类支付"
            Case 63
                MessageHeader_App_Type_Str = "IC卡增值业务类支付"
            Case Else
                MessageHeader_App_Type_Str = "N/A"
            End Select

            '//软件总版本号解析
            MessageHeader_Software_Total_Ver_Num = Mid(MessageHeader, 3, 2)

            Select Case MessageHeader_Software_Total_Ver_Num
            Case 10
                MessageHeader_Software_Total_Ver_Num_Str = "2001年人民银行POS规范之前版本"
            Case 11
                MessageHeader_Software_Total_Ver_Num_Str = "2001年人民银行POS规范版本"
            Case 21
                MessageHeader_Software_Total_Ver_Num_Str = "2002年银联POS规范版本"
            Case 22
                MessageHeader_Software_Total_Ver_Num_Str = "2004年银联POS规范版本"
            Case 30
                MessageHeader_Software_Total_Ver_Num_Str = "2009年银联POS规范版本"
            Case 31
                MessageHeader_Software_Total_Ver_Num_Str = "2010年银联POS规范版本"
            Case Else
                MessageHeader_Software_Total_Ver_Num_Str = "N/A"
            End Select

            '//终端状态解析
            MessageHeader_Terminal_State = Mid(MessageHeader, 5, 1)

            Select Case MessageHeader_Terminal_State
            Case 0
                MessageHeader_Terminal_State_Str = "正常交易状态"
            Case Else
                MessageHeader_Terminal_State_Str = "N/A"
            End Select

            '//处理要求解析
            MessageHeader_Process_Require = Mid(MessageHeader, 6, 1)

            Select Case MessageHeader_Process_Require
            Case 0
                MessageHeader_Process_Require_Str = "无处理要求"
            Case 1
                MessageHeader_Process_Require_Str = "下传终端磁条卡参数"
            Case 2
                MessageHeader_Process_Require_Str = "上传终端磁条卡状态信息"
            Case 3
                MessageHeader_Process_Require_Str = "重新签到"
            Case 4
                MessageHeader_Process_Require_Str = "通知终端发起更新公钥信息操作"
            Case 5
                MessageHeader_Process_Require_Str = "下载终端IC卡参数"
            Case 6
                MessageHeader_Process_Require_Str = "TMS参数下载"
            Case 7
                MessageHeader_Process_Require_Str = "卡BIN 黑名单下载"
            Case 8
                MessageHeader_Process_Require_Str = "币种汇率下载（仅在境外使用）/助弄取款手续费比率下载（仅在境内使用）"

            Case Else
                MessageHeader_Process_Require_Str = "N/A"
            End Select

            '// 软件分版本号
            MessageHeader_Software_Part_Ver_Num = ins_space(Mid(MessageHeader, 7, 6))
            MessageHeader_Software_Part_Ver_Num_Str = "前两字节同软件版本号，后四字节由厂商自行定义"

            MessageHeader = ins_space(MessageHeader)
            tex2STR = tex2STR & "[MessageHeader:报文头]" & vbCrLf & "" & MessageHeader & vbCrLf _
                      & "----------------" & vbCrLf & "->[应用类别定义]" & vbCrLf & MessageHeader_App_Type & vbCrLf & MessageHeader_App_Type_Str & vbCrLf _
                      & "----------------" & vbCrLf & "->[软件总版本号]" & vbCrLf & MessageHeader_Software_Total_Ver_Num & vbCrLf & MessageHeader_Software_Total_Ver_Num_Str & vbCrLf _
                      & "----------------" & vbCrLf & "->[终端状态]" & vbCrLf & MessageHeader_Terminal_State & vbCrLf & MessageHeader_Terminal_State_Str & vbCrLf _
                      & "----------------" & vbCrLf & "->[处理要求]" & vbCrLf & MessageHeader_Process_Require & vbCrLf & MessageHeader_Process_Require_Str & vbCrLf _
                      & "----------------" & vbCrLf & "->[软件分版本号]" & vbCrLf & MessageHeader_Software_Part_Ver_Num & vbCrLf & MessageHeader_Software_Part_Ver_Num_Str & vbCrLf & "----------------" & vbCrLf & vbCrLf
            tempCount = tempCount + 12
        End If



        Dim temp_messagetype_str

        '//messagetype解析
        POS_Sturct.messagetype = Mid(tempStr, tempCount, 4)

        Select Case POS_Sturct.messagetype
        Case "0100"
            temp_messagetype_str = "—— 0100 授权类请求消息：" & vbCrLf & _
                                   "-> POS 预授权请求" & vbCrLf & _
                                   "-> POS 预授权撤销请求" & vbCrLf & _
                                   "-> 磁条卡现金充值账户验证请求" & vbCrLf
        Case "0110"
            temp_messagetype_str = "—— 0110 授权类应答消息：" & vbCrLf & _
                                   "-> POS 预授权应答" & vbCrLf & _
                                   "-> POS 预授权撤销应答" & vbCrLf & _
                                   "-> 磁条卡现金充值账户验证应答" & vbCrLf
        Case "0200"
            temp_messagetype_str = "—— 0200 金融类请求消息：" & vbCrLf & _
                                   "-> POS 查询请求" & vbCrLf & _
                                   "-> POS 消费请求" & vbCrLf & _
                                   "-> POS 消费撤销请求" & vbCrLf & _
                                   "-> POS 预授权完成（请求）请求" & vbCrLf & _
                                   "-> POS 预授权完成撤销请求" & vbCrLf & _
                                   "-> 电子现金脱机消费请求" & vbCrLf & _
                                   "-> 分期付款消费请求" & vbCrLf & _
                                   "-> 分期付款消费撤销请求" & vbCrLf & _
                                   "-> 基于 PBOC 电子钱包/电子现金的 IC 圈存类交易请求" & vbCrLf & _
                                   "-> 磁条卡现金充值请求" & vbCrLf & _
                                   "-> 磁条卡帐户充值请求" & vbCrLf
        Case "0210"
            temp_messagetype_str = "—— 0210 金融类应答消息：" & vbCrLf & _
                                   "-> POS 查询应答" & vbCrLf & _
                                   "-> POS 消费应答" & vbCrLf & _
                                   "-> POS 消费撤销应答" & vbCrLf & _
                                   "-> POS 预授权完成（请求）应答" & vbCrLf & _
                                   "-> POS 预授权完成撤销应答" & vbCrLf & _
                                   "-> 电子现金脱机消费应答" & vbCrLf & _
                                   "-> 分期付款消费应答" & vbCrLf & _
                                   "-> 分期付款消费撤销应答" & vbCrLf & _
                                   "-> 基于 PBOC 电子钱包/电子现金的 IC 圈存类交易应答" & vbCrLf & _
                                   "-> 磁条卡现金充值应答" & vbCrLf & _
                                   "-> 磁条卡帐户充值应答" & vbCrLf
        Case "0220"
            temp_messagetype_str = "—— 0220 金融通知类消息：" & vbCrLf & _
                                   "-> POS 退货通知" & vbCrLf & _
                                   "-> POS 离线结算通知" & vbCrLf & _
                                   "-> POS 结算调整通知" & vbCrLf & _
                                   "-> POS 预授权完成（通知）通知" & vbCrLf & _
                                   "-> 磁条卡现金充值确认通知" & vbCrLf
        Case "0230"
            temp_messagetype_str = "—— 0230 金融通知类应答消息：" & vbCrLf & _
                                   "-> POS 退货应答" & vbCrLf & _
                                   "-> POS 离线结算应答" & vbCrLf & _
                                   "-> POS 结算调整应答" & vbCrLf & _
                                   "-> POS 预授权完成（通知）应答" & vbCrLf & _
                                   "-> 磁条卡现金充值确认应答" & vbCrLf
        Case "0320"
            temp_messagetype_str = "—— 0320 批上送消息：" & vbCrLf & _
                                   "-> POS 终端批上送" & vbCrLf
        Case "0330"
            temp_messagetype_str = "—— 0330 批上送应答消息：" & vbCrLf & _
                                   "-> POS 终端批上送应答" & vbCrLf

        Case "0400"
            temp_messagetype_str = "—— 0400 冲正类消息：" & vbCrLf & _
                                   "-> POS 预授权冲正" & vbCrLf & _
                                   "-> POS 预授权撤销冲正" & vbCrLf & _
                                   "-> POS 消费冲正" & vbCrLf & _
                                   "-> POS 消费撤销冲正" & vbCrLf & _
                                   "-> POS 预授权完成（请求）冲正" & vbCrLf & _
                                   "-> POS 预授权完成撤销冲正" & vbCrLf & _
                                   "-> 基于 PBOC 电子钱包/电子现金的 IC 圈存类交易冲正" & vbCrLf
        Case "0410"
            temp_messagetype_str = "—— 0410 冲正类应答消息：" & vbCrLf & _
                                   "-> POS 预授权冲正应答" & vbCrLf & _
                                   "-> POS 预授权撤销冲正应答" & vbCrLf & _
                                   "-> POS 消费冲正应答" & vbCrLf & _
                                   "-> POS 消费撤销冲正应答" & vbCrLf & _
                                   "-> POS 预授权完成（请求）冲正应答" & vbCrLf & _
                                   "-> POS 预授权完成撤销冲正应答" & vbCrLf & _
                                   "-> 基于 PBOC 电子钱包/电子现金的 IC 圈存类交易冲正应答" & vbCrLf

        Case "0500"
            temp_messagetype_str = "—— 0500 对账类消息：" & vbCrLf & _
                                   "-> POS 终端批结算请求" & vbCrLf
        Case "0510"
            temp_messagetype_str = "—— 0510 对账类应答消息：" & vbCrLf & _
                                   "-> POS 终端批结算应答" & vbCrLf

        Case "0620"
            temp_messagetype_str = "—— 0620 基于 PBOC 借/贷记卡标准的 IC 卡脚本处理结果通知消息：" & vbCrLf & _
                                   "-> 基于 PBOC 借/贷记卡标准的 IC 卡脚本处理结果通知" & vbCrLf
        Case "0630"
            temp_messagetype_str = "—— 0630 基于 PBOC 借/贷记卡标准的 IC 卡脚本处理结果通知应答：" & vbCrLf & _
                                   "-> 基于 PBOC 借/贷记卡标准的 IC 卡脚本处理结果通知应答" & vbCrLf

        Case "0800"
            temp_messagetype_str = "—— 0800 网络业务管理类消息：" & vbCrLf & _
                                   "-> POS 终端签到请求" & vbCrLf & _
                                   "-> POS 终端参数传递请求" & vbCrLf
        Case "0810"
            temp_messagetype_str = "—— 0810 网络业务管理类应答消息：" & vbCrLf & _
                                   "-> POS 终端签到应答" & vbCrLf & _
                                   "-> POS 终端参数传递应答" & vbCrLf

        Case "0820"
            temp_messagetype_str = "—— 0820 网络业务管理类消息：" & vbCrLf & _
                                   "-> POS 终端签退请求" & vbCrLf & _
                                   "-> POS 终端回响测试请求" & vbCrLf & _
                                   "-> POS 终端状态上送" & vbCrLf
        Case "0830"
            temp_messagetype_str = "—— 0830 网络业务管理类应答消息：" & vbCrLf & _
                                   "-> POS 终端签退应答" & vbCrLf & _
                                   "-> POS 终端回响测试应答" & vbCrLf & _
                                   "-> POS 终端状态上送应答" & vbCrLf


        Case Else
            temp_messagetype_str = "N/A"
        End Select



        POS_Sturct.messagetype = ins_space(POS_Sturct.messagetype)
        tex2STR = tex2STR & "[messagetype:消息类型]" & vbCrLf & "" & POS_Sturct.messagetype & vbCrLf & temp_messagetype_str & vbCrLf
        tempCount = tempCount + 4


        Dim temp_bcd_flag_display_str As String
        Dim temp_bcd_flag_str As String
        Dim temp_bcd_flag_str_count As Integer
        Dim tempLen As Integer

        '//bitmap解析
        POS_Sturct.bitmap = Mid(tempStr, tempCount, 16)
        POS_Sturct.bitmap = ins_space(POS_Sturct.bitmap)

        '//bitmap分离
        temp_bcd_flag_str = HEX_to_BIN(delete_space(POS_Sturct.bitmap))


        '//bitmap分离显示处理
        For temp_bcd_flag_str_count = 1 To Len(temp_bcd_flag_str)
            If Mid(temp_bcd_flag_str, temp_bcd_flag_str_count, 1) = 1 Then
                If temp_bcd_flag_str_count <> Len(temp_bcd_flag_str) Then
                    temp_bcd_flag_display_str = temp_bcd_flag_display_str & temp_bcd_flag_str_count & "域 "
                Else
                    temp_bcd_flag_display_str = temp_bcd_flag_display_str & temp_bcd_flag_str_count & "域"
                End If
            End If
        Next temp_bcd_flag_str_count


        tex2STR = tex2STR & "[bitmap:位元表]" & vbCrLf & "" & POS_Sturct.bitmap & vbCrLf & "->[存在以下域] " & temp_bcd_flag_display_str & vbCrLf & vbCrLf
        tempCount = tempCount + 16
        bit_map.Text = POS_Sturct.bitmap
        bit_map_change_function





        Dim tempPan As String
        Dim temp_Pan_len As String
        '//第2域数据 【主账号】
        If Mid(temp_bcd_flag_str, 2, 1) Then
            temp_Pan_len = Mid(tempStr, tempCount, 2)
            POS_Sturct.pan_2.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.pan_2.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.pan_2.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            tempPan = Mid(POS_Sturct.pan_2.ptr, 1, POS_Sturct.pan_2.Ptrlen)
            POS_Sturct.pan_2.ptr = ins_space(POS_Sturct.pan_2.ptr)
            tex2STR = tex2STR & "[field2:主账号(Primary Account Number)]" & vbCrLf & "" & "[" & temp_Pan_len & "]" & " " & POS_Sturct.pan_2.ptr _
                      & vbCrLf & "->[卡号] " & tempPan & vbCrLf & vbCrLf
        End If

        '//第3域数据 【交易处理码】
        If Mid(temp_bcd_flag_str, 3, 1) Then
            POS_Sturct.procode_3 = Mid(tempStr, tempCount, 6)
            POS_Sturct.procode_3 = ins_space(POS_Sturct.procode_3)
            tex2STR = tex2STR & "[field3:交易处理码(Transaction Processing Code)]" & vbCrLf & "" & POS_Sturct.procode_3 & vbCrLf & vbCrLf
            tempCount = tempCount + 6
        End If
        '//类型判断
        i = Type_judge(POS_Sturct.messagetype, Mid(POS_Sturct.procode_3, 1, 2), Mid(temp_bcd_flag_str, 61, 1))

        Dim temp_consume_amount_4 As String
        '//第4域数据 【交易金额】
        If Mid(temp_bcd_flag_str, 4, 1) Then
            POS_Sturct.consume_amount_4 = Mid(tempStr, tempCount, 12)
            temp_consume_amount_4 = Val(POS_Sturct.consume_amount_4) / 100

            If Val(temp_consume_amount_4) < 1 And Val(temp_consume_amount_4) > 0 Then
                temp_consume_amount_4 = 0 & temp_consume_amount_4
            End If


            POS_Sturct.consume_amount_4 = ins_space(POS_Sturct.consume_amount_4)
            tex2STR = tex2STR & "[field4:交易金额(Amount Of Transactions)]" & vbCrLf & "" & POS_Sturct.consume_amount_4 & _
                      vbCrLf & "->[￥金额] " & temp_consume_amount_4 & "元" & vbCrLf & vbCrLf
            tempCount = tempCount + 12
        End If

        Dim temp_trace_no_11 As String
        '//第11域数据 【受卡方系统跟踪号】
        If Mid(temp_bcd_flag_str, 11, 1) Then
            POS_Sturct.trace_no_11 = Mid(tempStr, tempCount, 6)
            temp_trace_no_11 = POS_Sturct.trace_no_11
            POS_Sturct.trace_no_11 = ins_space(POS_Sturct.trace_no_11)
            tex2STR = tex2STR & "[field11:受卡方系统跟踪号(System Trace Audit Number)]" & vbCrLf & "" & POS_Sturct.trace_no_11 & _
                      vbCrLf & "->[流水号:] " & temp_trace_no_11 & vbCrLf & vbCrLf

            trace_no_11.Caption = temp_trace_no_11

            tempCount = tempCount + 6
        End If


        '//第12域数据 【受卡方所在地时间】
        If Mid(temp_bcd_flag_str, 12, 1) Then
            POS_Sturct.trade_time_12 = Mid(tempStr, tempCount, 6)
            POS_Sturct.trade_time_12 = ins_space(POS_Sturct.trade_time_12)
            tex2STR = tex2STR & "[field12:受卡方所在地时间(Local Time Of Transaction)]" & vbCrLf & "" & POS_Sturct.trade_time_12 & _
                      vbCrLf & "->[时间] " & Mid(POS_Sturct.trade_time_12, 1, 2) & "时" & Mid(POS_Sturct.trade_time_12, 4, 2) & "分" _
                      & Mid(POS_Sturct.trade_time_12, 7, 2) & "秒" & vbCrLf & vbCrLf
            tempCount = tempCount + 6
        End If

        '//第13域数据 【受卡方所在地日期】
        If Mid(temp_bcd_flag_str, 13, 1) Then
            POS_Sturct.trade_date_13 = Mid(tempStr, tempCount, 4)
            POS_Sturct.trade_date_13 = ins_space(POS_Sturct.trade_date_13)
            tex2STR = tex2STR & "[field13:受卡方所在地日期(Local Date Of Transaction)]" & vbCrLf & "" & POS_Sturct.trade_date_13 & "   " & _
                      vbCrLf & "->[日期] " & Mid(POS_Sturct.trade_date_13, 1, 2) & "月" & Mid(POS_Sturct.trade_date_13, 4, 2) & "日" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//第14域数据 【卡有效期】
        If Mid(temp_bcd_flag_str, 14, 1) Then
            POS_Sturct.exp_date_14 = Mid(tempStr, tempCount, 4)
            POS_Sturct.exp_date_14 = ins_space(POS_Sturct.exp_date_14)
            tex2STR = tex2STR & "[field14:卡有效期(Date Of Expired)]" & vbCrLf & "" & POS_Sturct.exp_date_14 & "   " & _
                      vbCrLf & "->[有效期] " & "20" & Mid(POS_Sturct.exp_date_14, 1, 2) & "年" & Mid(POS_Sturct.exp_date_14, 4, 2) & "月" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//第15域数据 【清算日期】
        If Mid(temp_bcd_flag_str, 15, 1) Then
            POS_Sturct.settlement_date_15 = Mid(tempStr, tempCount, 4)
            POS_Sturct.settlement_date_15 = ins_space(POS_Sturct.settlement_date_15)
            tex2STR = tex2STR & "[field15:清算日期(Date Of Settlement)]" & vbCrLf & "" & POS_Sturct.settlement_date_15 & "   " & _
                      vbCrLf & "->[清算日期] " & Mid(POS_Sturct.settlement_date_15, 1, 2) & "月" & Mid(POS_Sturct.settlement_date_15, 4, 2) & "日" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//第22域数据 【服务点输入方式码】
        If Mid(temp_bcd_flag_str, 22, 1) Then
            POS_Sturct.entry_mode_22 = Mid(tempStr, tempCount, 4)
            POS_Sturct.entry_mode_22 = ins_space(POS_Sturct.entry_mode_22)
            tex2STR = tex2STR & "[field22:服务点输入方式码(Point Of Service Entry Mode)]" & vbCrLf & "" & POS_Sturct.entry_mode_22 & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If




        Dim Right_23_str As String    '//卡片序列号右靠
        Dim left_23_str As String    '//卡片序列号左靠
        '//第23域数据 【卡序列号】
        If Mid(temp_bcd_flag_str, 23, 1) Then
            POS_Sturct.card_serial_number_23 = Mid(tempStr, tempCount, 4)
            Right_23_str = Mid(POS_Sturct.card_serial_number_23, 2, 3)
            left_23_str = Mid(POS_Sturct.card_serial_number_23, 1, 3)
            POS_Sturct.card_serial_number_23 = ins_space(POS_Sturct.card_serial_number_23)
            tex2STR = tex2STR & "[field23:卡序列号(Card Sequence Number)]" & vbCrLf & "" & POS_Sturct.card_serial_number_23 & vbCrLf _
                      & "->[卡片序列号左靠] " & left_23_str & vbCrLf & "->[卡片序列号右靠] " & Right_23_str & " [一般优先选择右靠]" & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//第25域数据 【服务点条件码】
        If Mid(temp_bcd_flag_str, 25, 1) Then
            POS_Sturct.service_conditon_25 = Mid(tempStr, tempCount, 2)
            POS_Sturct.service_conditon_25 = ins_space(POS_Sturct.service_conditon_25)
            tex2STR = tex2STR & "[field25:服务点条件码(Point Of Service Condition Mode)]" & vbCrLf & "" & POS_Sturct.service_conditon_25 & vbCrLf & vbCrLf
            tempCount = tempCount + 2
        End If

        '//第26域数据 【服务点PIN获取码】
        If Mid(temp_bcd_flag_str, 26, 1) Then
            POS_Sturct.service_conditon_pin_26 = Mid(tempStr, tempCount, 2)
            POS_Sturct.service_conditon_pin_26 = ins_space(POS_Sturct.service_conditon_pin_26)
            tex2STR = tex2STR & "[field26:服务点PIN获取码(Point Of Service PIN Capture Code)]" & vbCrLf & "" & POS_Sturct.service_conditon_pin_26 & vbCrLf & vbCrLf
            tempCount = tempCount + 2
        End If

        Dim temp_32_len_str As String
        '//第32域数据 【受理方标识码】
        If Mid(temp_bcd_flag_str, 32, 1) Then
            temp_32_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.api_code_32.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.api_code_32.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.api_code_32.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.api_code_32.ptr = ins_space(POS_Sturct.api_code_32.ptr)

            tex2STR = tex2STR & "[field32:受理方标识码(Acquiring Institution Id Code)]" & vbCrLf & "" & "[" & temp_32_len_str & "]" & " " & POS_Sturct.api_code_32.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_35_len_str As String
        '//第35域数据 【2磁道数据】
        If Mid(temp_bcd_flag_str, 35, 1) Then
            temp_35_len_str = Mid(tempStr, tempCount, 2)
            POS_Sturct.track2_35.Ptrlen = Mid(tempStr, tempCount, 2)
            tempLen = ((POS_Sturct.track2_35.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 2

            POS_Sturct.track2_35.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.track2_35.ptr = ins_space(POS_Sturct.track2_35.ptr)

            tex2STR = tex2STR & "[field35:2磁道数据(Track 2 Data)]" & vbCrLf & "" & "[" & temp_35_len_str & "]" & "  " & POS_Sturct.track2_35.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_36_len_str As String
        '//第36域数据 【3磁道数据】
        If Mid(temp_bcd_flag_str, 36, 1) Then
            temp_36_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.track3_36.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.track3_36.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.track3_36.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.track3_36.ptr = ins_space(POS_Sturct.track3_36.ptr)
            temp_36_len_str = ins_space(temp_36_len_str)
            tex2STR = tex2STR & "[field36:3磁道数据(Track 3 Data)]" & vbCrLf & "" & "[" & temp_36_len_str & "]" & "  " & POS_Sturct.track3_36.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_37str As String
        '//第37域数据 【检索参考号】
        If Mid(temp_bcd_flag_str, 37, 1) Then
            POS_Sturct.reference_number_37 = Mid(tempStr, tempCount, 24)
            POS_Sturct.reference_number_37 = ins_space(POS_Sturct.reference_number_37)
            temp_37str = ASCchange(POS_Sturct.reference_number_37)
            tex2STR = tex2STR & "[field37:检索参考号(Retrieval Reference Number)]" & vbCrLf & "" & POS_Sturct.reference_number_37 & vbCrLf & "->[参考号] " & temp_37str & vbCrLf & vbCrLf
            tempCount = tempCount + 24
        End If

        Dim temp_38str As String
        '//第38域数据 【授权标识应答码】
        If Mid(temp_bcd_flag_str, 38, 1) Then
            POS_Sturct.authorization_code_38 = Mid(tempStr, tempCount, 12)
            POS_Sturct.authorization_code_38 = ins_space(POS_Sturct.authorization_code_38)
            temp_38str = ASCchange(POS_Sturct.authorization_code_38)
            tex2STR = tex2STR & "[field38:授权标识应答码(Authorization Id Response Code)]" & vbCrLf & "" & POS_Sturct.authorization_code_38 & vbCrLf & "->[授权码] " & temp_38str & vbCrLf & vbCrLf
            tempCount = tempCount + 12
        End If

        Dim temp_39str As String
        '//第39域数据 【应答码】
        If Mid(temp_bcd_flag_str, 39, 1) Then
            POS_Sturct.Response_code_39 = Mid(tempStr, tempCount, 4)
            POS_Sturct.Response_code_39 = ins_space(POS_Sturct.Response_code_39)
            temp_39str = ASCchange(POS_Sturct.Response_code_39)
            tex2STR = tex2STR & "[field39:应答码(Response Code)]" & vbCrLf & "" & POS_Sturct.Response_code_39 & vbCrLf & "->[应答码:] " & temp_39str & vbCrLf & vbCrLf
            tempCount = tempCount + 4
        End If

        '//响应码判断显示
        i = Response_code_39_Type_judge(ASCchange(POS_Sturct.Response_code_39))

        Dim temp_41str As String
        '//第41域数据 【受卡机终端标识码】
        If Mid(temp_bcd_flag_str, 41, 1) Then
            POS_Sturct.terminal_no_41 = Mid(tempStr, tempCount, 16)
            POS_Sturct.terminal_no_41 = ins_space(POS_Sturct.terminal_no_41)
            temp_41str = ASCchange(POS_Sturct.terminal_no_41)
            tex2STR = tex2STR & "[field41:受卡机终端标识码(Card Acceptor Terminal Id)]" & vbCrLf & "" & POS_Sturct.terminal_no_41 & vbCrLf & "->[终端号] " & temp_41str & vbCrLf & vbCrLf
            tempCount = tempCount + 16
        End If

        Dim temp_42str As String
        '//第42域数据 【受卡方标识码】
        If Mid(temp_bcd_flag_str, 42, 1) Then
            POS_Sturct.merchant_no_42 = Mid(tempStr, tempCount, 30)
            POS_Sturct.merchant_no_42 = ins_space(POS_Sturct.merchant_no_42)
            temp_42str = ASCchange(POS_Sturct.merchant_no_42)
            tex2STR = tex2STR & "[field42:受卡方标识码(Card Acceptor Id Code)]" & vbCrLf & "" & POS_Sturct.merchant_no_42 & vbCrLf & "->[商户号] " & temp_42str & vbCrLf & vbCrLf
            tempCount = tempCount + 30
        End If


        Dim temp_43_len_str As String
        '//第43域数据 【merchant_name_43】
        If Mid(temp_bcd_flag_str, 43, 1) Then
            temp_43_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.merchant_name_43.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.merchant_name_43.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.merchant_name_43.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.merchant_name_43.ptr = ins_space(POS_Sturct.merchant_name_43.ptr)
            temp_43_len_str = ins_space(temp_43_len_str)
            tex2STR = tex2STR & "[field43:自定义域(merchant_name_43)]" & vbCrLf & "" & "[" & temp_43_len_str & "]" & "  " & POS_Sturct.merchant_name_43.ptr & vbCrLf & vbCrLf

        End If


        Dim temp_44_len_str As String
        Dim issuing_bank As String   '//发卡行
        Dim Acquiring_bank As String    '//收单行
        '//第44域数据 【附加响应数据】
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
                tex2STR = tex2STR & "[field44:附加响应数据(Additional Response Data)]" & vbCrLf & "" & "[" & temp_44_len_str & "]" & "  " & POS_Sturct.rsp_code_44.ptr _
                          & vbCrLf & "->[发卡行] " & issuing_bank & vbCrLf & "->[收单行] " & Acquiring_bank & vbCrLf & vbCrLf
            Else
                tex2STR = tex2STR & "[field44:附加响应数据(Additional Response Data)]" & vbCrLf & "" & "[" & temp_44_len_str & "]" & "  " & POS_Sturct.rsp_code_44.ptr & vbCrLf & vbCrLf
            End If
        End If

        Dim temp_46_len_str As String
        '//第46域数据 【pay_signature_46】
        If Mid(temp_bcd_flag_str, 46, 1) Then
            temp_46_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.pay_signature_46.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.pay_signature_46.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.pay_signature_46.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.pay_signature_46.ptr = ins_space(POS_Sturct.pay_signature_46.ptr)
            temp_46_len_str = ins_space(temp_46_len_str)
            tex2STR = tex2STR & "[field46:自定义域(pay_signature_46)]" & vbCrLf & "" & "[" & temp_46_len_str & "]" & "  " & POS_Sturct.pay_signature_46.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_48_len_str As String
        '//第48域数据 【附加数据 - 私有】
        If Mid(temp_bcd_flag_str, 48, 1) Then
            temp_48_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.settleAccounts_48.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.settleAccounts_48.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.settleAccounts_48.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.settleAccounts_48.ptr = ins_space(POS_Sturct.settleAccounts_48.ptr)
            temp_48_len_str = ins_space(temp_48_len_str)
            tex2STR = tex2STR & "[field48:附加数据 - 私有(Additional Data - Private)]" & vbCrLf & "" & "[" & temp_48_len_str & "]" & "  " & POS_Sturct.settleAccounts_48.ptr & vbCrLf & vbCrLf
        End If

        Dim temp_49str As String
        '//第49域数据 【交易货币代码】
        If Mid(temp_bcd_flag_str, 49, 1) Then
            POS_Sturct.currency_code_49 = Mid(tempStr, tempCount, 6)
            POS_Sturct.currency_code_49 = ins_space(POS_Sturct.currency_code_49)
            temp_49str = ASCchange(POS_Sturct.currency_code_49)

            tex2STR = tex2STR & "[field49:交易货币代码(Currency Code Of Transaction)]" & vbCrLf & "" & POS_Sturct.currency_code_49 & vbCrLf & vbCrLf

            tempCount = tempCount + 6
        End If

        '//第52域数据 【个人标识码】
        If Mid(temp_bcd_flag_str, 52, 1) Then
            POS_Sturct.pri_pin_52 = Mid(tempStr, tempCount, 16)
            POS_Sturct.pri_pin_52 = ins_space(POS_Sturct.pri_pin_52)
            tex2STR = tex2STR & "[field52:个人标识码数据(PIN Data)]" & vbCrLf & "" & POS_Sturct.pri_pin_52 & vbCrLf & vbCrLf

            tempCount = tempCount + 16
        End If

        '//第53域数据 【安全控制信息】
        If Mid(temp_bcd_flag_str, 53, 1) Then
            POS_Sturct.safety_53 = Mid(tempStr, tempCount, 16)
            POS_Sturct.safety_53 = ins_space(POS_Sturct.safety_53)
            tex2STR = tex2STR & "[field53:安全控制信息(Security Related Control Information )]" & vbCrLf & "" & POS_Sturct.safety_53 & vbCrLf & vbCrLf

            tempCount = tempCount + 16
        End If

        Dim temp_54str As String
        Dim temp_54_len_str As String
        Dim temp_consume_amount_54 As String
        '//第54域数据 【余额】
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
            tex2STR = tex2STR & "[field54:余额(Balanc Amount)]" & vbCrLf & "" & "[" & temp_54_len_str & "]" & "  " & POS_Sturct.attachment_amount_54.ptr & vbCrLf & "->[ASCII转换] " _
                      & temp_54str & vbCrLf & "->[金额]      " & temp_consume_amount_54 & "元" & vbCrLf & vbCrLf
        End If



        Dim Application_Cryptogram$    '//应用密文  9F26   8字节



        Dim temp_55_len_str As String
        '//第55域数据 【IC卡数据域】
        If Mid(temp_bcd_flag_str, 55, 1) Then
            temp_55_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.icData_55.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.icData_55.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.icData_55.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.icData_55.ptr = ins_space(POS_Sturct.icData_55.ptr)
            temp_55_len_str = ins_space(temp_55_len_str)
            tex2STR = tex2STR & "[field55:IC卡数据域(IC Card System Related Data)]" & vbCrLf & "" & "[" & temp_55_len_str & "]" & "  " & POS_Sturct.icData_55.ptr & vbCrLf & vbCrLf

        End If

















        Dim temp_56_len_str As String
        '//第56域数据 【private_data_56】
        If Mid(temp_bcd_flag_str, 56, 1) Then
            temp_56_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_56.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_56.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_56.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_56.ptr = ins_space(POS_Sturct.private_data_56.ptr)
            temp_56_len_str = ins_space(temp_56_len_str)
            tex2STR = tex2STR & "[field56:自定义域(private_data_56)]" & vbCrLf & "" & "[" & temp_56_len_str & "]" & "  " & POS_Sturct.private_data_56.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_57_len_str As String
        '//第57域数据 【private_data_57】
        If Mid(temp_bcd_flag_str, 57, 1) Then
            temp_57_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_57.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_57.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_57.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_57.ptr = ins_space(POS_Sturct.private_data_57.ptr)
            temp_57_len_str = ins_space(temp_57_len_str)
            tex2STR = tex2STR & "[field57:自定义域(private_data_57)]" & vbCrLf & "" & "[" & temp_57_len_str & "]" & "  " & POS_Sturct.private_data_57.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_58_len_str As String
        '//第58域数据 【PBOC电子钱包标准的交易信息】
        If Mid(temp_bcd_flag_str, 58, 1) Then
            temp_58_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_58.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_58.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_58.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_58.ptr = ins_space(POS_Sturct.private_data_58.ptr)
            temp_58_len_str = ins_space(temp_58_len_str)
            tex2STR = tex2STR & "[field58:PBOC电子钱包标准的交易信息(PBOC_ELECTRONIC_DATA)]" & vbCrLf & "" & "[" & temp_58_len_str & "]" & "  " & POS_Sturct.private_data_58.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_59_len_str As String
        '//第59域数据 【private_data_59】
        If Mid(temp_bcd_flag_str, 59, 1) Then
            temp_59_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_59.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = POS_Sturct.private_data_59.Ptrlen * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_59.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen
            POS_Sturct.private_data_59.ptr = ins_space(POS_Sturct.private_data_59.ptr)
            temp_59_len_str = ins_space(temp_59_len_str)
            tex2STR = tex2STR & "[field59:自定义域(private_data_59)]" & vbCrLf & "" & "[" & temp_59_len_str & "]" & "  " & POS_Sturct.private_data_59.ptr & vbCrLf & vbCrLf

        End If

        Dim temp_60_len_str As String

        Dim temp_trans_Type_60 As String          '//60.1 消息类型码
        Dim temp_batch_Number_60 As String        '//60.2 批次号
        Dim temp_network_60 As String             '//60.3 网络管理信息码
        Dim temp_readingAbility_60 As String      '//60.4 终端读取能力
        Dim temp_conditionCode_60 As String       '//60.5 基于 PBOC 借/贷记标准的 IC 卡条件代码
        Dim temp_supportSome_60 As String         '//60.6 支持部分扣款和返回余额标志
        Dim temp_account_type_60 As String        '//60.7 帐户类型

        Dim temp_trans_Type_60_str As String      '//60.1 消息类型码显示
        Dim temp_batch_Number_60_str As String    '//60.2 批次号显示
        Dim temp_network_60_str As String         '//60.3 网络管理信息码显示
        Dim temp_readingAbility_60_str As String  '//60.4 终端读取能力显示
        Dim temp_conditionCode_60_str As String   '//60.5 基于 PBOC 借/贷记标准的 IC 卡条件代码显示
        Dim temp_supportSome_60_str As String     '//60.6 支持部分扣款和返回余额标志显示
        Dim temp_account_type_60_str As String    '//60.7 帐户类型显示

        '//第60域数据 【private_data_60】
        If Mid(temp_bcd_flag_str, 60, 1) Then
            temp_60_len_str = Mid(tempStr, tempCount, 4)
            POS_Sturct.private_data_60.Ptrlen = Mid(tempStr, tempCount, 4)
            tempLen = ((POS_Sturct.private_data_60.Ptrlen + 1) \ 2) * 2
            tempCount = tempCount + 4

            POS_Sturct.private_data_60.ptr = Mid(tempStr, tempCount, tempLen)
            tempCount = tempCount + tempLen


            temp_trans_Type_60 = Mid(POS_Sturct.private_data_60.ptr, 1, 2)    '//60.1 消息类型码
            Select Case temp_trans_Type_60
            Case "00"
                temp_trans_Type_60_str = "管理类交易，脚本通知交易"
            Case "01"
                temp_trans_Type_60_str = "查询"
            Case "03"
                temp_trans_Type_60_str = "积分查询"
            Case "10"
                temp_trans_Type_60_str = "预授权/冲正"
            Case "11"
                temp_trans_Type_60_str = "预授权撤销/冲正"
            Case "20"
                temp_trans_Type_60_str = "预授权完成（请求） /冲正"
            Case "21"
                temp_trans_Type_60_str = "预授权完成撤销/冲正"
            Case "22"
                temp_trans_Type_60_str = "消费/冲正"
            Case "23"
                temp_trans_Type_60_str = "消费撤销/冲正"
            Case "24"
                temp_trans_Type_60_str = "预授权完成（通知）"
            Case "25"
                temp_trans_Type_60_str = "退货（包含联盟积分退货）"
            Case "27"
                temp_trans_Type_60_str = "IC 卡脱机交易退货"
            Case "30"
                temp_trans_Type_60_str = "离线结算"
            Case "32"
                temp_trans_Type_60_str = "结算调整"
            Case "34"
                temp_trans_Type_60_str = "结算调整(追加小费)"
            Case "36"
                temp_trans_Type_60_str = "脱机消费"
            Case "40"
                temp_trans_Type_60_str = "电子钱包的 IC 卡指定账户圈存/冲正"
            Case "41"
                temp_trans_Type_60_str = "电子钱包的 IC 卡现金充值/冲正"
            Case "42"
                temp_trans_Type_60_str = "电子钱包的 IC 卡非指定账户转账圈存/冲正"
            Case "45"
                temp_trans_Type_60_str = "电子现金指定账户圈存/冲正"
            Case "46"
                temp_trans_Type_60_str = "电子现金现金充值（撤销） /冲正"
            Case "47"
                temp_trans_Type_60_str = "电子现金非指定账户转账圈存/冲正"
            Case "48"
                temp_trans_Type_60_str = "磁条卡现金充值/确认"
            Case "49"
                temp_trans_Type_60_str = "磁条卡帐户充值"
            Case "51"
                temp_trans_Type_60_str = "电子现金现金充值撤销/冲正"
            Case "53"
                temp_trans_Type_60_str = "预约消费撤销/冲正"
            Case "54"
                temp_trans_Type_60_str = "预约消费/冲正"
            Case Else
                temp_trans_Type_60_str = "N/A"
            End Select

            temp_batch_Number_60 = Mid(POS_Sturct.private_data_60.ptr, 3, 6)    '//60.2 批次号

            temp_network_60 = Mid(POS_Sturct.private_data_60.ptr, 9, 3)    '//60.3 网络管理信息码
            Select Case temp_network_60
            Case "001"
                temp_network_60_str = "POS 终端签到（单倍长密钥算法）"
            Case "002"
                temp_network_60_str = "POS 终端签退"
            Case "003"
                temp_network_60_str = "POS 终端签到（双倍长密钥算法)"
            Case "004"
                temp_network_60_str = "POS 终端签到（双倍长密钥算法，含磁道密钥）"
            Case "201"
                temp_network_60_str = "POS 终端批结算"
            Case "201"
                temp_network_60_str = "POS 终端批上送"
            Case "202"
                temp_network_60_str = "对账不平衡时， POS 终端批上送结束"
            Case "203"
                temp_network_60_str = "对账平衡时， POS 终端上送成功的 IC 卡联机交"
            Case "204"
                temp_network_60_str = "对账平衡时， POS 终端上送 IC 卡通知信息"
            Case "205"
                temp_network_60_str = "对账不平衡时， POS 终端上送成功的 IC 卡联机"
            Case "206"
                temp_network_60_str = "对账不平衡时， POS 终端上送 IC 卡通知信息"
            Case "207"
                temp_network_60_str = "对账平衡时， POS 终端批上送结束"
            Case "208"
                temp_network_60_str = "对账平衡时， POS 终端上送圈存交易圈存确认明"
            Case "209"
                temp_network_60_str = "对账不平衡时， POS 终端上送圈存交易圈存确认"
            Case "301"
                temp_network_60_str = "回响测试"
            Case "401"
                temp_network_60_str = "收银员签到"
            Case "362"
                temp_network_60_str = "POS 终端状态监控"
            Case "360"
                temp_network_60_str = "POS 终端磁条卡参数下载"
            Case "361"
                temp_network_60_str = "POS 终端磁条卡参数下载结束"
            Case "364"
                temp_network_60_str = "POS 终端 TMS 参数下载"
            Case "365"
                temp_network_60_str = "POS 终端 TMS 参数下载结束"
            Case "370"
                temp_network_60_str = "POS 终端 IC 卡公钥下载"
            Case "371"
                temp_network_60_str = "POS 终端 IC 卡公钥下载结束"
            Case "372"
                temp_network_60_str = "POS 终端 IC 卡公钥信息查询"
            Case "380"
                temp_network_60_str = "POS 终端 IC 卡参数下载"
            Case "381"
                temp_network_60_str = "POS 终端 IC 卡参数下载结束"
            Case "382"
                temp_network_60_str = "POS 终端 IC 卡参数信息查询"
            Case "384"
                temp_network_60_str = "POS 终端币种汇率下载（仅在境外使用）"
            Case "385"
                temp_network_60_str = "POS 终端币种汇率下载结束（仅在境外使用）"
            Case "390"
                temp_network_60_str = "POS 终端卡 BIN 黑名单下载"
            Case "391"
                temp_network_60_str = "POS 终端卡 BIN 黑名单下载结束"
            Case "392"
                temp_network_60_str = "POS 终端小额取现的手续费下载（预留）"
            Case "393"
                temp_network_60_str = "POS 终端小额取现的手续费下载结束（预留）"
            Case "951"
                temp_network_60_str = "基于 PBOC 借/贷记标准 IC 卡脚本处理结果通知"
            Case Else
                temp_network_60_str = "N/A"
            End Select

            temp_readingAbility_60 = Mid(POS_Sturct.private_data_60.ptr, 12, 1)     '//60.4 终端读取能力
            Select Case temp_readingAbility_60
            Case "0"
                temp_readingAbility_60_str = "终端读取能力不可知"
            Case "2"
                temp_readingAbility_60_str = "可读取磁条卡"
            Case "5"
                temp_readingAbility_60_str = "可接触式界面读取 IC 卡。对于电子钱包的非接触界面读取，该域也填 5"
            Case "6"
                temp_readingAbility_60_str = "可非接触式界面读取 IC 卡（包括可读取 CUPMobile 移动支付方案中非接触式终端）。" & _
                                             "当22 域前两位取值 07、 91、 96 或 98 时，该域必须填 6。但对于电子钱包的非接触界面读取，该域仍然填 5"
            Case Else
                temp_readingAbility_60_str = "N/A"
            End Select

            temp_conditionCode_60 = Mid(POS_Sturct.private_data_60.ptr, 13, 1)       '//60.5 基于 PBOC 借/贷记标准的 IC 卡条件代码
            Select Case temp_conditionCode_60
            Case "0"
                temp_conditionCode_60_str = "未使用或后续子域存在，或手机芯片交易"
            Case "1"
                temp_conditionCode_60_str = "上一笔交易不是 IC 卡交易或是一笔成功的 IC 卡交易"
            Case "2"
                temp_conditionCode_60_str = "上一笔交易虽是 IC 卡交易但失败"
            Case Else
                temp_conditionCode_60_str = "N/A"
            End Select

            temp_supportSome_60 = Mid(POS_Sturct.private_data_60.ptr, 14, 1)         '//60.6 支持部分扣款和返回余额标志
            Select Case temp_conditionCode_60
            Case "0"
                temp_supportSome_60_str = "支持部分扣款和返回余额标志"
            Case "1"
                temp_supportSome_60_str = "不支持部分扣款和返回余额标志"
            Case Else
                temp_supportSome_60_str = "N/A"
            End Select

            temp_account_type_60 = Mid(POS_Sturct.private_data_60.ptr, 15, 3)           '//60.7 帐户类型
            Select Case temp_account_type_60
            Case "0"
                temp_account_type_60_str = "发卡行积分，表示数字0的ASCII码"
            Case "1"
                temp_account_type_60_str = "银联联盟积分，表示字母A的ASCII码"
            Case Else
                temp_account_type_60_str = "N/A"
            End Select

            POS_Sturct.private_data_60.ptr = ins_space(POS_Sturct.private_data_60.ptr)

            temp_60_len_str = ins_space(temp_60_len_str)
            tex2STR = tex2STR & "[field60:自定义域(private_data_60)]" & vbCrLf & "" & "[" & temp_60_len_str & "]" & "  " & POS_Sturct.private_data_60.ptr & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[消息类型码]" & vbCrLf & temp_trans_Type_60 & vbCrLf & temp_trans_Type_60_str & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[批次号]" & vbCrLf & temp_batch_Number_60 & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[网络管理信息码]" & vbCrLf & temp_network_60 & vbCrLf & temp_network_60_str & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[终端读取能力]" & vbCrLf & temp_readingAbility_60 & vbCrLf & temp_readingAbility_60_str & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[基于PBOC 借/贷记标准的IC卡条件代码]" & vbCrLf & temp_conditionCode_60 & vbCrLf & temp_conditionCode_60_str & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[支持部分扣款和返回余额标志]" & vbCrLf & temp_supportSome_60 & vbCrLf & temp_supportSome_60_str & vbCrLf _
                      & "-------------------------------------" & vbCrLf & "->[帐户类型]" & vbCrLf & temp_account_type_60 & vbCrLf & temp_account_type_60_str & vbCrLf & "-------------------------------------" & vbCrLf & vbCrLf

            batch_Number_60.Caption = temp_batch_Number_60

        End If

        Dim temp_61_len_str As String
        Dim Original_batch_Number_61 As String                             '//原始交易批次号
        Dim Original_trace_no_61 As String                                 '//原始交易POS流水号
        Dim Original_trans_date_61 As String                               '//原始交易日期
        Dim Original_trans_authorization_61 As String                      '// 原交易授权方式
        Dim Original_trans_authorization_institution_code_61 As String     '//原交易授权机构代码

        '//第61域数据 【原始信息域】
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
            tex2STR = tex2STR & "[field61:原始信息域(Original Message)]" & vbCrLf & "" & "[" & temp_61_len_str & "]" & "  " & POS_Sturct.private_data_61.ptr _
                      & vbCrLf & "---------------------" & vbCrLf & "->[原始交易批次号]" & vbCrLf & Original_batch_Number_61 _
                      & vbCrLf & "---------------------" & vbCrLf & "->[原始交易POS流水号]" & vbCrLf & Original_trace_no_61 _
                      & vbCrLf & "---------------------" & vbCrLf & "->[原始交易日期]" & vbCrLf & Original_trans_date_61 _
                      & "-> " & Mid(Original_trans_date_61, 1, 2) & "月" & Mid(Original_trans_date_61, 4, 2) & "日" _
                      & vbCrLf & "---------------------" & vbCrLf & "->[原交易授权方式]" & vbCrLf & Original_trans_authorization_61 _
                      & vbCrLf & "---------------------" & vbCrLf & "->[原交易授权机构代码]" & vbCrLf & Original_trans_authorization_institution_code_61 _
                      & vbCrLf & "---------------------" & vbCrLf & vbCrLf

        End If


        Dim temp_62str As String
        Dim temp_62_len_str As String
        '//第62域数据 【private_data_62】
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
            tex2STR = tex2STR & "[field62:自定义域(private_data_62)]" & vbCrLf & "" & "[" & temp_62_len_str & "]" & "  " & POS_Sturct.private_data_62.ptr & vbCrLf _
                      & "->[ASCII转换]" & vbCrLf & temp_62str & vbCrLf & vbCrLf

        End If


        Dim temp_63str As String
        Dim temp_63_len_str As String
        '//第63域数据 【private_data_63】
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
            tex2STR = tex2STR & "[field63:自定义域(private_data_63)]" & vbCrLf & "" & "[" & temp_63_len_str & "]" & "  " & POS_Sturct.private_data_63.ptr & vbCrLf _
                      & "->[ASCII转换]" & vbCrLf & temp_63str & vbCrLf & vbCrLf

        End If

        '//第64域数据 【报文鉴别码】
        If Mid(temp_bcd_flag_str, 64, 1) Then
            POS_Sturct.mac_64 = Mid(tempStr, tempCount, 16)
            POS_Sturct.mac_64 = ins_space(POS_Sturct.mac_64)
            tex2STR = tex2STR & "[field64:报文鉴别码(Message Authentication Code)]" & vbCrLf & "" & POS_Sturct.mac_64 & vbCrLf & vbCrLf
            tempCount = tempCount + 16
        End If



        '//  tex2STR = tex2STR & "************8583解析结果（我是华丽的分割线结尾）***************/" & vbCrLf
        analyse_after_data.Text = tex2STR
    End If

End Sub






'//其他功能区

Private Sub bit_map_clear_Click()
    Dim i As Integer
    For i = 1 To 64
        bit_map_view_Click_flag(i) = False
        bit_map_view(i - 1).BackColor = &H8000000F    '//恢复原样
    Next
    bit_map.Text = ""
End Sub

Private Sub bit_map_set_Click()
    bit_map.SetFocus
    Dim i As Integer
    If bit_map.Text <> "" Then
        bit_map_change_function

    Else
        '        i = MsgBox("位图串为空，请正确输入", vbCritical, "提示")
        MSG_BOX "位图串为空" & vbCrLf & "请正确输入", "提示"

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
        bit_map_view(Index - 1).BackColor = &HFF8080    '//按下去浅蓝色
    Else
        bit_map_view_Click_flag(Index) = False
        bit_map_view(Index - 1).BackColor = &H8000000F    '//恢复原样
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
        bit_map_view(i - 1).BackColor = &H8000000F    '//恢复原样
    Next   '//清除位图

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
            bit_map_view(i - 1).BackColor = &HFF8080  '//按下去浅蓝色
        End If
    Next
    temp_bit_map_str = BIN_to_HEX(temp_bcd_str)
    bit_map.Text = ins_space(temp_bit_map_str)

End Function


Private Sub clear_Click()    '//清屏事件
    analyse_before_data.Text = ""
    analyse_after_data.Text = ""
    bit_map_clear_Click
    trans_type.Caption = "N/A"
    judge_mode.Caption = "N/A"
    judge_mode.ForeColor = &H0&    '//恢复成黑色

    Response_code.Caption = "N/A"
    Response_code_view.Caption = "N/A"

    trace_no_11.Caption = "N/A"
    batch_Number_60.Caption = "N/A"

    help_flag = 0
    help.Caption = "说明"


    analyse_before_data.SetFocus


End Sub

Private Sub about_Click()    '//关于事件
    Dim i As Integer

    '    i = MsgBox("8583辅助解析工具V1.2                      " _
         '               & vbCrLf & "" _
         '               & vbCrLf & "作者：高建宽" _
         '               & vbCrLf & "版本：V1.2" _
         '               & vbCrLf & "QQ：1062220953" _
         '               & vbCrLf & "日期：2015年01月10日" _
         '               & vbCrLf & "TIPS：在按钮或框架悬停查看帮助" _
         '               , vbOKOnly, "关于")
    About_BOX
End Sub




Private Sub END_Click()    '//退出事件
    End
End Sub



Private Sub Frame_analyse_after_data_Click()    '//复制解析后数据
    Clipboard.clear
    Clipboard.SetText analyse_after_data.Text    '//自动复制结果
    MSG_BOX "解析后数据复制成功", "提示"
End Sub



Private Sub Frame_analyse_before_data_Click()   '//复制解析前数据
    Clipboard.clear
    Clipboard.SetText analyse_before_data.Text    '//自动复制结果
    MSG_BOX "解析前数据复制成功", "提示"
End Sub

Private Sub Frame_bit_map_Click()  '//复制位图数据
    Clipboard.clear
    Clipboard.SetText bit_map.Text    '//自动复制结果
    MSG_BOX "位图复制成功", "提示"
End Sub


Private Sub help_Click()    '//帮助事件

    Static temp_analyse_after_data_str As String
    If help_count = 0 Then
        temp_analyse_after_data_str = analyse_after_data.Text
    End If
    help_count = help_count + 1
    If help_flag = 0 Then
        analyse_after_data.SetFocus
        analyse_after_data.Text = "                   【8583辅助解析工具V1.3.1 说明】                     " & vbCrLf & vbCrLf & _
                                  "按照【销售点终端（POS）应用规范(QCUP 009.1-2010)】进行解析" & vbCrLf & vbCrLf & _
                                  "V1.0增加基本解析，位图显示与设置，悬停框架与按钮可以查看具体帮助" & vbCrLf & vbCrLf & _
                                  "V1.1增加了普通解析，1.0版基本解析变为专家解析，增加响应码提示。TIPS:暂支持销售点终端（POS）应用规范(QCUP 009.1-2010)协议的解析，如果用其他非标准协议报文解析，可能会出现类型不匹配的警告。" & vbCrLf & vbCrLf & _
                                  "V1.2改进如下" & vbCrLf & _
                                  "->1.换了新的图标和界面" & vbCrLf & _
                                  "->2.优化了专家解析，增加了全面解析" & vbCrLf & _
                                  "->3.增加了流水号/批次号提示" & vbCrLf & _
                                  "->4.增加退出按钮" & vbCrLf & vbCrLf & _
                                  "V1.3改进如下" & vbCrLf & _
                                  "->1.增加了55域解析" & vbCrLf & _
                                  "->2.当响应码39域在标准中找不到，提示交易失败。增加77响应码" & vbCrLf & _
                                  "->3.屏蔽了全面解析" & vbCrLf & _
                                  "->4.当出现类型不匹配的解析前数据，不会确认后退出程序" & vbCrLf & _
                                  "->5.新增关于窗体" & vbCrLf & _
                                  "->6.增加框架提示复制框架内数据功能" & vbCrLf & vbCrLf & _
                                  "V1.3.1改进如下" & vbCrLf & _
                                  "->1.修复55域解析其中TAG值分析不对情况" & vbCrLf

        help_flag = 1
        help.Caption = "清除说明"
    Else
        analyse_after_data.SetFocus
        help_flag = 0
        help.Caption = "说明"
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






























'//函数区


'/***********************************************************************************
'函数名称:PROCESS_Analyze_data
'功能描述:处理解析前的数据，把数据转换成可以分析的数据,去除数据中的空格，回车，换行符，便于解析
'输　入:
'输　出:
'备  注:
'例子：12 34 56 78 90 。。。->1234567890
'
'***********************************************************************************/

Private Sub PROCESS_Analyze_data()
    Dim Text1_tempStr, Text2_tempStr As String, count, totalNum, i As Long

    Text1_tempStr = analyse_before_data.Text

    If Len(Text1_tempStr) = 0 Then

        '        i = MsgBox("解析前数据为空，请正确输入", vbCritical, "提示")
        MSG_BOX "解析前数据为空" & vbCrLf & "请正确输入", "提示"
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
'** 函数名称: ins_space
'** 功能描述: 插入空格
'** 输　入:
'** 输　出:s
'** 备  注:例子:1234-->12 34
'********************************************************************************************************/


Private Function ins_space(ByVal src As String) As String    '//插入空格例子:1234-->12 34

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
'** 函数名称: delete_space
'** 功能描述: 删除空格
'** 输　入:
'** 输　出:
'** 备  注:例子:12 34-->1234
'********************************************************************************************************/


Private Function delete_space(ByVal src As String) As String    '//插入空格例子:1234-->12 34

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

' 用途：将十六进制转化为二进制
' 输入：Hex(十六进制数)
' 输入数据类型：String
' 输出：HEX_to_BIN(二进制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
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


' 用途：将二进制转化为十六进制
' 输入：Bin(二进制数)
' 输入数据类型：String
' 输出：BIN_to_HEX(十六进制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
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





' 用途：将十六进制转化为十进制 31->48(支持最多两位十六进制转换成十进制)
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
'函数名称:ACSchange
'功能描述:ACSCII码转换
'输　入:
'输　出:
'备  注:例子：31 32 33 34 35 36 37 38 39 30->1234567890
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
'* 名称:     MSG_BOX
'* 功能说明: 提示框
'* 调用:
'* 输入:     Prompt：必选。字符串表达式，显示在对话框中的消息
'            title:  必选。字符串表达式，在对话框标题栏中显示的内容
'* 返回值:
'* 备注:
'************************************************************************************************/
Private Sub MSG_BOX(Prompt As String, title As String)
    PromptForm.Show
    MainForm.Enabled = False

    PromptForm.msg_display.Caption = Prompt
    PromptForm.Caption = title
End Sub


'/************************************************************************************************
'* 名称:     About_BOX
'* 功能说明: 关于框
'* 调用:
'* 输入:     Prompt：必选。字符串表达式，显示在对话框中的消息
'            title:  必选。字符串表达式，在对话框标题栏中显示的内容
'* 返回值:
'* 备注:
'************************************************************************************************/
Private Sub About_BOX()
    Form_About.Show
    MainForm.Enabled = False

    '    PromptForm.msg_display.Caption = Prompt
    '    PromptForm.Caption = title
End Sub



























'//常数区
'/***********************************************************************************
'函数名称:Type_judge
'功能描述:类型判断
'输　入:
'输　出:
'备  注:按照销售点终端（POS）应用规范(QCUP 009.1-2010)规范
'交易类型模式中红色代表请求，蓝色代表响应
'增加field61_flag标志位区分 消费和预授权完成 2014年12月24日09:51:16
'【基于PBOC系列】和【磁条卡系列】和【离线类和脱机类】和【收银员积分签到】的没做判断
'可能造成误判断，对此不负责
'***********************************************************************************/
Private Function Type_judge(ByVal messagetype As String, ByVal procode As String, ByVal field61_flag As String)

'//交易类
    If messagetype = "02 00" And procode = "31" Then
        trans_type.Caption = "余额查询"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "02 10" And procode = "31" Then
        trans_type.Caption = "余额查询"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "02 00" And procode = "00" And field61_flag = 0 Then
        trans_type.Caption = "消费"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "02 10" And procode = "00" And field61_flag = 0 Then
        trans_type.Caption = "消费"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "04 00" And procode = "00" Then
        trans_type.Caption = "消费冲正"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "04 10" And procode = "00" Then
        trans_type.Caption = "消费冲正"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "02 00" And procode = "20" Then
        trans_type.Caption = "消费撤销"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "02 10" And procode = "20" Then
        trans_type.Caption = "消费撤销"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "04 00" And procode = "20" Then
        trans_type.Caption = "消费撤销冲正"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "04 10" And procode = "20" Then
        trans_type.Caption = "消费撤销冲正"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "02 20" And procode = "20" Then
        trans_type.Caption = "退货"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "02 30" And procode = "20" Then
        trans_type.Caption = "退货"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "01 00" And procode = "03" Then
        trans_type.Caption = "预授权"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "01 10" And procode = "03" Then
        trans_type.Caption = "预授权"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "04 00" And procode = "03" Then
        trans_type.Caption = "预授权冲正"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "04 10" And procode = "03" Then
        trans_type.Caption = "预授权冲正"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "01 00" And procode = "20" Then
        trans_type.Caption = "预授权撤销"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "01 10" And procode = "20" Then
        trans_type.Caption = "预授权撤销"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "04 00" And procode = "20" Then
        trans_type.Caption = "预授权撤销冲正"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "04 10" And procode = "20" Then
        trans_type.Caption = "预授权撤销冲正"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "02 00" And procode = "00" And field61_flag = 1 Then
        trans_type.Caption = "预授权完成(请求)"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "02 10" And procode = "00" And field61_flag = 1 Then
        trans_type.Caption = "预授权完成(请求)"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "02 20" And procode = "00" And field61_flag = 1 Then
        trans_type.Caption = "预授权完成(通知)"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "02 30" And procode = "00" And field61_flag = 1 Then
        trans_type.Caption = "预授权完成(通知)"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "04 00" And procode = "00" And field61_flag = 1 Then
        trans_type.Caption = "预授权完成(请求)冲正"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "04 10" And procode = "00" And field61_flag = 1 Then
        trans_type.Caption = "预授权完成(请求)冲正"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "02 00" And procode = "20" And field61_flag = 1 Then
        trans_type.Caption = "预授权完成撤销"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "02 10" And procode = "20" And field61_flag = 1 Then
        trans_type.Caption = "预授权完成撤销"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "04 00" And procode = "20" And field61_flag = 1 Then
        trans_type.Caption = "预授权完成撤销冲正"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "04 10" And procode = "20" And field61_flag = 1 Then
        trans_type.Caption = "预授权完成撤销冲正"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

        '//管理类
    ElseIf messagetype = "08 00" Then
        trans_type.Caption = "签到"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "08 10" Then
        trans_type.Caption = "签到"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "08 20" Then
        trans_type.Caption = "签退"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "08 30" Then
        trans_type.Caption = "签退"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "05 00" Then
        trans_type.Caption = "批结算"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "05 10" Then
        trans_type.Caption = "批结算"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色

    ElseIf messagetype = "03 20" Then
        trans_type.Caption = "批上送金融交易/批上送结束"
        judge_mode.Caption = "请求"
        judge_mode.ForeColor = &HFF&    '//红色
    ElseIf messagetype = "03 30" Then
        trans_type.Caption = "批上送金融交易/批上送结束"
        judge_mode.Caption = "响应"
        judge_mode.ForeColor = &HFF0000    '//蓝色
    Else
        trans_type.Caption = "N/A"
        judge_mode.Caption = "N/A"
        judge_mode.ForeColor = &H0&    '//恢复成黑色
    End If

End Function



'/***********************************************************************************
'函数名称:Response_code_39_Type_judge
'功能描述:类型判断
'输　入:
'输　出:
'备  注:按照销售点终端（POS）应用规范(QCUP 009.1-2010)规范
'用了ultra edit 编写
'***********************************************************************************/
Private Function Response_code_39_Type_judge(ByVal src As String)

    If src = "" Then
        Response_code_view.Caption = "N/A": Response_code.Caption = "N/A"
    ElseIf src = "00" Then
        Response_code_view.Caption = "交易成功": Response_code.Caption = src
    ElseIf src = "01" Then
        Response_code_view.Caption = "请持卡人与发卡银行联系": Response_code.Caption = src
    ElseIf src = "03" Then
        Response_code_view.Caption = "无效商户": Response_code.Caption = src
    ElseIf src = "04" Then
        Response_code_view.Caption = "此卡被没收": Response_code.Caption = src
    ElseIf src = "05" Then
        Response_code_view.Caption = "持卡人认证失败": Response_code.Caption = src
    ElseIf src = "10" Then
        Response_code_view.Caption = "显示部分批准金额，提示操作员": Response_code.Caption = src
    ElseIf src = "11" Then
        Response_code_view.Caption = "成功，VIP客户": Response_code.Caption = src
    ElseIf src = "12" Then
        Response_code_view.Caption = "无效交易": Response_code.Caption = src
    ElseIf src = "13" Then
        Response_code_view.Caption = "无效金额": Response_code.Caption = src
    ElseIf src = "14" Then
        Response_code_view.Caption = "无效卡号": Response_code.Caption = src
    ElseIf src = "15" Then
        Response_code_view.Caption = "此卡无对应发卡方": Response_code.Caption = src
    ElseIf src = "21" Then
        Response_code_view.Caption = "该卡未初始化或睡眠卡": Response_code.Caption = src
    ElseIf src = "22" Then
        Response_code_view.Caption = "操作有误，或超出交易允许天数": Response_code.Caption = src
    ElseIf src = "25" Then
        Response_code_view.Caption = "没有原始交易，请联系发卡方": Response_code.Caption = src
    ElseIf src = "30" Then
        Response_code_view.Caption = "请重试": Response_code.Caption = src
    ElseIf src = "34" Then
        Response_code_view.Caption = "作弊卡,呑卡": Response_code.Caption = src
    ElseIf src = "38" Then
        Response_code_view.Caption = "密码错误次数超限，请与发卡方联系": Response_code.Caption = src
    ElseIf src = "40" Then
        Response_code_view.Caption = "发卡方不支持的交易类型": Response_code.Caption = src
    ElseIf src = "41" Then
        Response_code_view.Caption = "挂失卡，请没收（POS）": Response_code.Caption = src
    ElseIf src = "43" Then
        Response_code_view.Caption = "被窃卡，请没收": Response_code.Caption = src
    ElseIf src = "51" Then
        Response_code_view.Caption = "可用余额不足": Response_code.Caption = src
    ElseIf src = "54" Then
        Response_code_view.Caption = "该卡已过期": Response_code.Caption = src
    ElseIf src = "55" Then
        Response_code_view.Caption = "密码错": Response_code.Caption = src
    ElseIf src = "57" Then
        Response_code_view.Caption = "不允许此卡交易": Response_code.Caption = src
    ElseIf src = "58" Then
        Response_code_view.Caption = "发卡方不允许该卡在本终端进行此交易": Response_code.Caption = src
    ElseIf src = "59" Then
        Response_code_view.Caption = "卡片校验错": Response_code.Caption = src
    ElseIf src = "61" Then
        Response_code_view.Caption = "交易金额超限": Response_code.Caption = src
    ElseIf src = "62" Then
        Response_code_view.Caption = "受限制的卡": Response_code.Caption = src
    ElseIf src = "64" Then
        Response_code_view.Caption = "交易金额与原交易不匹配": Response_code.Caption = src
    ElseIf src = "65" Then
        Response_code_view.Caption = "超出消费次数限制": Response_code.Caption = src
    ElseIf src = "68" Then
        Response_code_view.Caption = "交易超时，请重试": Response_code.Caption = src
    ElseIf src = "75" Then
        Response_code_view.Caption = "密码错误次数超限": Response_code.Caption = src
    ElseIf src = "90" Then
        Response_code_view.Caption = "系统日切，请稍后重试": Response_code.Caption = src
    ElseIf src = "91" Then
        Response_code_view.Caption = "发卡方状态不正常，请稍后重试": Response_code.Caption = src
    ElseIf src = "92" Then
        Response_code_view.Caption = "发卡方线路异常，请稍后重试": Response_code.Caption = src
    ElseIf src = "94" Then
        Response_code_view.Caption = "拒绝，重复交易，请稍后重试": Response_code.Caption = src
    ElseIf src = "96" Then
        Response_code_view.Caption = "拒绝，交换中心异常，请稍后重试": Response_code.Caption = src
    ElseIf src = "97" Then
        Response_code_view.Caption = "终端未登记": Response_code.Caption = src
    ElseIf src = "98" Then Response_code_view.Caption = "发卡方超时": Response_code.Caption = src
    ElseIf src = "99" Then
        Response_code_view.Caption = "PIN格式错，请重新签到": Response_code.Caption = src
    ElseIf src = "A0" Then
        Response_code_view.Caption = "MAC校验错，请重新签到": Response_code.Caption = src
    ElseIf src = "A1" Then
        Response_code_view.Caption = "转账货币不一致": Response_code.Caption = src
    ElseIf src = "A2" Then
        Response_code_view.Caption = "交易成功，请向发卡行确认": Response_code.Caption = src
    ElseIf src = "A3" Then
        Response_code_view.Caption = "账户不正确": Response_code.Caption = src
    ElseIf src = "A4" Then
        Response_code_view.Caption = "交易成功，请向发卡行确认": Response_code.Caption = src
    ElseIf src = "A5" Then
        Response_code_view.Caption = "交易成功，请向发卡行确认": Response_code.Caption = src
    ElseIf src = "A6" Then
        Response_code_view.Caption = "交易成功，请向发卡行确认": Response_code.Caption = src
    ElseIf src = "A7" Then
        Response_code_view.Caption = "拒绝，交换中心异常，请稍后重试": Response_code.Caption = src
    ElseIf src = "77" Then
        Response_code_view.Caption = "操作员重新签到，再作交易 ": Response_code.Caption = src
    Else
        Response_code_view.Caption = "交易失败": Response_code.Caption = src   '//其他情况直接交易失败
    End If

End Function





















