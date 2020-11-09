VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "服务控制面板"
   ClientHeight    =   7620
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5592
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9132.584
   ScaleMode       =   0  'User
   ScaleWidth      =   5391.91
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "远程连接和安全"
      Height          =   2892
      Left            =   240
      TabIndex        =   8
      Top             =   4440
      Width           =   5052
      Begin VB.CommandButton Command10 
         Caption         =   "暂停服务"
         Height          =   372
         Left            =   2520
         TabIndex        =   11
         Top             =   960
         Width           =   2172
      End
      Begin VB.CommandButton Command9 
         Caption         =   "安全性"
         Height          =   372
         Left            =   2520
         TabIndex        =   10
         Top             =   480
         Width           =   2172
      End
      Begin VB.CommandButton Command8 
         Caption         =   "数据库连接"
         Height          =   372
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   2052
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "厂内通信配置"
      Height          =   3972
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5052
      Begin VB.CommandButton Command55 
         Caption         =   "连接"
         Height          =   372
         Left            =   3960
         TabIndex        =   15
         Top             =   2400
         Width           =   972
      End
      Begin VB.CommandButton Command44 
         Caption         =   "刷新"
         Height          =   372
         Left            =   2880
         TabIndex        =   14
         Top             =   2400
         Width           =   972
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   1680
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   2400
         Width           =   972
      End
      Begin VB.CommandButton Command7 
         Caption         =   "7号釜"
         Height          =   372
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   2052
      End
      Begin VB.CommandButton Command6 
         Caption         =   "6号釜"
         Height          =   372
         Left            =   2520
         TabIndex        =   6
         Top             =   1440
         Width           =   2172
      End
      Begin VB.CommandButton Command5 
         Caption         =   "5号釜"
         Height          =   372
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   2052
      End
      Begin VB.CommandButton Command4 
         Caption         =   "4号釜"
         Height          =   372
         Left            =   2520
         TabIndex        =   4
         Top             =   960
         Width           =   2172
      End
      Begin VB.CommandButton Command3 
         Caption         =   "3号釜"
         Height          =   372
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   2052
      End
      Begin VB.CommandButton Command2 
         Caption         =   "2号釜"
         Height          =   372
         Left            =   2520
         TabIndex        =   2
         Top             =   480
         Width           =   2172
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1号釜"
         Height          =   372
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   2052
      End
      Begin VB.Label Label2 
         Caption         =   "指定总线串口："
         Height          =   252
         Left            =   360
         TabIndex        =   12
         Top             =   2400
         Width           =   1332
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ZYF.FuNumber = 1
    Set ZYF.MSComm1 = MSComm1
    ZYF.Hide
    ZYF.Show
    ZYF.Caption = "1号釜传感器控制"
End Sub

Private Sub Command2_Click()
    ZYF.FuNumber = 2
    Set ZYF.MSComm1 = MSComm1
    ZYF.Hide
    ZYF.Show
    ZYF.Caption = "2号釜传感器控制"
End Sub

Private Sub Command3_Click()
    ZYF.FuNumber = 3
    ZYF.Hide
    ZYF.Show
    ZYF.Caption = "3号釜传感器控制"
End Sub

Private Sub Command4_Click()
    ZYF.FuNumber = 4
    ZYF.Hide
    ZYF.Show
    ZYF.Caption = "4号釜传感器控制"
End Sub

Private Sub Command44_Click()
    SerialPortCheck
End Sub
Private Sub SerialPortCheck()

    '检测当前有效串口
    On Error Resume Next

    For i = 1 To 16 Step 1
        MSComm1.CommPort = i
        MSComm1.PortOpen = True
        If Err.Number = 0 Then
            Combo1.AddItem "COM" & i
        End If
        MSComm1.PortOpen = False
        Err.Clear
    Next
        
    Combo1.Text = Combo1.List(0)
   
End Sub
Private Sub Command5_Click()
    ZYF.FuNumber = 5
    ZYF.Hide
    ZYF.Show
    ZYF.Caption = "5号釜传感器控制"
End Sub

Private Sub Command55_Click()
    On Error Resume Next
    If MSComm1.PortOpen = False Then
        MSComm1.CommPort = Right(Combo1.Text, Len(Combo1.Text) - 3)
        MSComm1.PortOpen = True
        If Err.Number <> 0 Then
            MsgBox "无法打开串口，请检查串口是否被占用！", vbOKOnly, "串口打开错误"
        Else
            Command1.Caption = "关闭串口"
            nowSerial = MSComm1.CommPort
            MsgBox ("服务器串口COM" & nowSerial & "已打开" & vbCrLf & "错误代码：" & Err.Number)
        End If
    Else
        MSComm1.PortOpen = False
        Command1.Caption = "打开串口"
        MsgBox "串口已关闭"
    End If



End Sub

Private Sub Command6_Click()
    ZYF.FuNumber = 6
    ZYF.Hide
    ZYF.Show
    ZYF.Caption = "6号釜传感器控制"
End Sub

Private Sub Command7_Click()
    ZYF.FuNumber = 7
    ZYF.Hide
    ZYF.Show
    ZYF.Caption = "7号釜传感器控制"
End Sub



Private Sub Form_load()
    SerialPortCheck
    
End Sub
