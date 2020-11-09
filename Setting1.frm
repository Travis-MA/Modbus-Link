VERSION 5.00
Begin VB.Form Setting1 
   Caption         =   "数据库设置"
   ClientHeight    =   3876
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3876
   ScaleWidth      =   8880
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "数据源"
      Height          =   3492
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8652
      Begin VB.TextBox Text6 
         Height          =   264
         Left            =   1560
         TabIndex        =   13
         Text            =   "生产数据"
         Top             =   1560
         Width           =   2892
      End
      Begin VB.TextBox Text5 
         Height          =   264
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   11
         Text            =   "Jdljnd1470"
         Top             =   1200
         Width           =   2892
      End
      Begin VB.TextBox Text4 
         Height          =   264
         Left            =   1560
         TabIndex        =   10
         Text            =   "sa"
         Top             =   840
         Width           =   2892
      End
      Begin VB.TextBox Text3 
         Height          =   264
         Left            =   1560
         TabIndex        =   7
         Text            =   "10"
         Top             =   2040
         Width           =   732
      End
      Begin VB.CommandButton Command1 
         Caption         =   "测试连接"
         Height          =   492
         Left            =   6960
         TabIndex        =   5
         Top             =   480
         Width           =   1452
      End
      Begin VB.TextBox Text2 
         Height          =   372
         Left            =   1560
         ScrollBars      =   1  'Horizontal
         TabIndex        =   4
         Text            =   $"Setting1.frx":0000
         Top             =   2520
         Width           =   5892
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Left            =   1560
         TabIndex        =   3
         Text            =   "f705e3a93053f184.natapp.cc,21433"
         Top             =   480
         Width           =   5172
      End
      Begin VB.Label Label9 
         Caption         =   "Microsoft"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   7080
         TabIndex        =   15
         Top             =   1680
         Width           =   1452
      End
      Begin VB.Label Label8 
         Caption         =   "数据库：SQL Server 2008"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   13.8
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   4680
         TabIndex        =   14
         Top             =   1320
         Width           =   3732
      End
      Begin VB.Label Label6 
         Caption         =   "对象名："
         Height          =   372
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1452
      End
      Begin VB.Label Label5 
         Caption         =   "密码"
         Height          =   372
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1452
      End
      Begin VB.Label Label4 
         Caption         =   "登陆用户名："
         Height          =   252
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1452
      End
      Begin VB.Label Label3 
         Caption         =   "最大重连次数："
         Height          =   372
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1572
      End
      Begin VB.Label Label2 
         Caption         =   "数据源名称："
         Height          =   252
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "连接字符串："
         Height          =   252
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   1092
      End
   End
End
Attribute VB_Name = "Setting1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error GoTo EH
    conn = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Text4.Text & ";Password =" & Text5.Text & "; Initial Catalog=" & Text6.Text & ";Data Source=" & Text1.Text
    CONNECT_LOOP_MAX = Val(Text3.Text)
    Text2.Text = conn
    DB_Connect
        
    MsgBox "连接成功"
    
    Exit Sub
EH:
    MsgBox "连接失败"
End Sub

Private Sub Form_Load()
    Text2.Text = conn
    Text3.Text = CONNECT_LOOP_MAX
    
End Sub

