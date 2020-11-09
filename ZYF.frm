VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form ZYF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "传感器管理测试工具"
   ClientHeight    =   7716
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   9900
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7716
   ScaleWidth      =   9900
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Caption         =   "数据库设置"
      Height          =   372
      Left            =   7440
      TabIndex        =   57
      Top             =   6480
      Width           =   1572
   End
   Begin VB.Frame Frame1 
      Caption         =   "蒸压釜RS485数据采集器"
      Height          =   7212
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   9132
      Begin VB.CheckBox Check3 
         Caption         =   "轮询刷新"
         Height          =   252
         Left            =   240
         TabIndex        =   62
         Top             =   2040
         Width           =   2172
      End
      Begin VB.CommandButton Command7 
         Caption         =   "一键启动"
         Height          =   372
         Left            =   7320
         TabIndex        =   61
         Top             =   6720
         Width           =   1572
      End
      Begin VB.CheckBox Check2 
         Caption         =   "开启数据库记录"
         Height          =   252
         Left            =   2400
         TabIndex        =   60
         Top             =   6240
         Width           =   1692
      End
      Begin VB.TextBox Text28 
         Height          =   264
         Left            =   5520
         TabIndex        =   59
         Text            =   "10"
         Top             =   6240
         Width           =   1332
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Left            =   7320
         TabIndex        =   56
         Text            =   "Combo2"
         Top             =   600
         Width           =   1572
      End
      Begin VB.TextBox Text27 
         Height          =   264
         Left            =   5520
         TabIndex        =   54
         Text            =   "****"
         Top             =   5760
         Width           =   3372
      End
      Begin VB.TextBox Text26 
         Height          =   264
         Left            =   5520
         TabIndex        =   53
         Text            =   "****"
         Top             =   5400
         Width           =   3372
      End
      Begin VB.TextBox Text25 
         Height          =   264
         Left            =   5520
         TabIndex        =   52
         Text            =   "****"
         Top             =   5040
         Width           =   3372
      End
      Begin VB.TextBox Text24 
         Height          =   264
         Left            =   5520
         TabIndex        =   51
         Text            =   "****"
         Top             =   4680
         Width           =   3372
      End
      Begin VB.TextBox Text23 
         Height          =   264
         Left            =   5520
         TabIndex        =   50
         Text            =   "****"
         Top             =   4320
         Width           =   3372
      End
      Begin VB.TextBox Text22 
         Height          =   264
         Left            =   5520
         TabIndex        =   49
         Text            =   "****"
         Top             =   3960
         Width           =   3372
      End
      Begin VB.TextBox Text21 
         Height          =   264
         Left            =   1560
         TabIndex        =   46
         Text            =   "********"
         Top             =   5760
         Width           =   3852
      End
      Begin VB.TextBox Text20 
         Height          =   264
         Left            =   1560
         TabIndex        =   45
         Text            =   "********"
         Top             =   5400
         Width           =   3852
      End
      Begin VB.TextBox Text19 
         Height          =   264
         Left            =   1560
         TabIndex        =   44
         Text            =   "********"
         Top             =   5040
         Width           =   3852
      End
      Begin VB.TextBox Text18 
         Height          =   264
         Left            =   1560
         TabIndex        =   43
         Text            =   "********"
         Top             =   4680
         Width           =   3852
      End
      Begin VB.CheckBox Check1 
         Caption         =   "仅显示数据"
         Height          =   180
         Left            =   2400
         TabIndex        =   36
         Top             =   3720
         Width           =   1212
      End
      Begin VB.CommandButton Command5 
         Caption         =   "测试发送"
         Height          =   372
         Left            =   7320
         TabIndex        =   35
         Top             =   3120
         Width           =   1572
      End
      Begin VB.TextBox Text17 
         Height          =   264
         Left            =   1560
         TabIndex        =   34
         Text            =   "********"
         Top             =   4320
         Width           =   3852
      End
      Begin VB.TextBox Text16 
         Height          =   264
         Left            =   1560
         TabIndex        =   33
         Text            =   "********"
         Top             =   3960
         Width           =   3852
      End
      Begin VB.TextBox Text15 
         Height          =   264
         Left            =   1560
         TabIndex        =   32
         Text            =   "点击文本框显示对应请求报文"
         Top             =   3120
         Width           =   5292
      End
      Begin VB.CommandButton Command4 
         Caption         =   "刷新"
         Height          =   372
         Left            =   480
         TabIndex        =   30
         Top             =   1440
         Width           =   972
      End
      Begin VB.CommandButton Command3 
         Caption         =   "取消"
         Height          =   372
         Left            =   7320
         TabIndex        =   29
         Top             =   2280
         Width           =   1572
      End
      Begin VB.CommandButton Command2 
         Caption         =   "确定"
         Height          =   372
         Left            =   7320
         TabIndex        =   28
         Top             =   1800
         Width           =   1572
      End
      Begin VB.TextBox Text14 
         Height          =   264
         Left            =   1560
         TabIndex        =   27
         Text            =   "1000"
         Top             =   2400
         Width           =   972
      End
      Begin VB.TextBox Text13 
         Height          =   264
         Left            =   5520
         TabIndex        =   25
         Text            =   "Text13"
         Top             =   2400
         Width           =   1332
      End
      Begin VB.TextBox Text12 
         Height          =   264
         Left            =   5520
         TabIndex        =   24
         Text            =   "Text12"
         Top             =   2040
         Width           =   1332
      End
      Begin VB.TextBox Text11 
         Height          =   264
         Left            =   5520
         TabIndex        =   23
         Text            =   "Text11"
         Top             =   1680
         Width           =   1332
      End
      Begin VB.TextBox Text10 
         Height          =   264
         Left            =   5520
         TabIndex        =   22
         Text            =   "Text10"
         Top             =   1320
         Width           =   1332
      End
      Begin VB.TextBox Text9 
         Height          =   264
         Left            =   5520
         TabIndex        =   21
         Text            =   "Text9"
         Top             =   960
         Width           =   1332
      End
      Begin VB.TextBox Text8 
         Height          =   264
         Left            =   5520
         TabIndex        =   20
         Text            =   "0"
         Top             =   600
         Width           =   1332
      End
      Begin VB.TextBox Text7 
         Height          =   264
         Left            =   3960
         TabIndex        =   17
         Text            =   "40006"
         Top             =   2400
         Width           =   1452
      End
      Begin VB.TextBox Text6 
         Height          =   264
         Left            =   3960
         TabIndex        =   16
         Text            =   "40005"
         Top             =   2040
         Width           =   1452
      End
      Begin VB.TextBox Text5 
         Height          =   264
         Left            =   3960
         TabIndex        =   15
         Text            =   "40004"
         Top             =   1680
         Width           =   1452
      End
      Begin VB.TextBox Text4 
         Height          =   264
         Left            =   3960
         TabIndex        =   14
         Text            =   "40003"
         Top             =   1320
         Width           =   1452
      End
      Begin VB.TextBox Text3 
         Height          =   264
         Left            =   3960
         TabIndex        =   13
         Text            =   "40002"
         Top             =   960
         Width           =   1452
      End
      Begin VB.TextBox Text2 
         Height          =   264
         Left            =   3960
         TabIndex        =   12
         Text            =   "40001"
         Top             =   600
         Width           =   1452
      End
      Begin VB.CommandButton Command1 
         Caption         =   "连接"
         Height          =   372
         Left            =   1560
         TabIndex        =   5
         Top             =   1440
         Width           =   972
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   1560
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1080
         Width           =   972
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Left            =   1560
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   600
         Width           =   612
      End
      Begin VB.Label Label23 
         Caption         =   "上传周期(sec):"
         Height          =   372
         Left            =   4200
         TabIndex        =   58
         Top             =   6260
         Width           =   1692
      End
      Begin VB.Label Label22 
         Caption         =   "蒸压釜编号"
         Height          =   252
         Left            =   7320
         TabIndex        =   55
         Top             =   360
         Width           =   1212
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   8880
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label20 
         Caption         =   "解析值"
         Height          =   252
         Left            =   5520
         TabIndex        =   48
         Top             =   3720
         Width           =   1332
      End
      Begin VB.Label Label19 
         Caption         =   "原始值"
         Height          =   252
         Left            =   1560
         TabIndex        =   47
         Top             =   3720
         Width           =   1332
      End
      Begin VB.Label Label18 
         Caption         =   "釜表温度4："
         Height          =   252
         Left            =   480
         TabIndex        =   42
         Top             =   5760
         Width           =   1332
      End
      Begin VB.Label Label17 
         Caption         =   "釜表温度3："
         Height          =   372
         Left            =   480
         TabIndex        =   41
         Top             =   5400
         Width           =   1452
      End
      Begin VB.Label Label16 
         Caption         =   "釜表温度2："
         Height          =   372
         Left            =   480
         TabIndex        =   40
         Top             =   5040
         Width           =   1212
      End
      Begin VB.Label Label15 
         Caption         =   "釜表温度1："
         Height          =   372
         Left            =   480
         TabIndex        =   39
         Top             =   4680
         Width           =   1092
      End
      Begin VB.Label Label14 
         Caption         =   "釜内温度："
         Height          =   252
         Left            =   480
         TabIndex        =   38
         Top             =   4320
         Width           =   972
      End
      Begin VB.Label Label13 
         Caption         =   "釜内压力："
         Height          =   252
         Left            =   480
         TabIndex        =   37
         Top             =   3960
         Width           =   1092
      End
      Begin VB.Label Label12 
         Caption         =   "请求报文:"
         Height          =   252
         Left            =   480
         TabIndex        =   31
         Top             =   3120
         Width           =   1332
      End
      Begin VB.Label Label11 
         Caption         =   "轮询间隔(ms):"
         Height          =   252
         Left            =   240
         TabIndex        =   26
         Top             =   2400
         Width           =   1212
      End
      Begin VB.Label Label10 
         Caption         =   "计算公式"
         Height          =   252
         Left            =   5520
         TabIndex        =   19
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label9 
         Caption         =   "寄存器地址"
         Height          =   252
         Left            =   3960
         TabIndex        =   18
         Top             =   360
         Width           =   1452
      End
      Begin VB.Label Label8 
         Caption         =   "釜表温度4："
         Height          =   252
         Left            =   2880
         TabIndex        =   11
         Top             =   2400
         Width           =   1932
      End
      Begin VB.Label Label7 
         Caption         =   "釜表温度3："
         Height          =   372
         Left            =   2880
         TabIndex        =   10
         Top             =   2040
         Width           =   1932
      End
      Begin VB.Label Label6 
         Caption         =   "釜表温度2："
         Height          =   372
         Left            =   2880
         TabIndex        =   9
         Top             =   1680
         Width           =   1932
      End
      Begin VB.Label Label5 
         Caption         =   "釜表温度1："
         Height          =   372
         Left            =   2880
         TabIndex        =   8
         Top             =   1320
         Width           =   1812
      End
      Begin VB.Label Label4 
         Caption         =   "釜内温度："
         Height          =   372
         Left            =   2880
         TabIndex        =   7
         Top             =   960
         Width           =   1812
      End
      Begin VB.Label Label3 
         Caption         =   "釜内压力："
         Height          =   252
         Left            =   2880
         TabIndex        =   6
         Top             =   600
         Width           =   1932
      End
      Begin VB.Label Label2 
         Caption         =   "指定总线串口："
         Height          =   252
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "指定采集器ID："
         Height          =   252
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1332
      End
   End
   Begin VB.Timer Timer1 
      Left            =   9120
      Top             =   1320
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8280
      Top             =   960
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer2 
      Left            =   9120
      Top             =   2040
   End
End
Attribute VB_Name = "ZYF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const VarNum = 6
Private Const FuNum = 7
Private FuNumber As Integer

Private YaLiReg As String
Private NeiWenReg As String
Private WaiWen1Reg As String
Private WaiWen2Reg As String
Private WaiWen3Reg As String
Private WaiWen4Reg As String

Private AYaLi As String
Private ANeiWen As String
Private AWaiWen1 As String
Private AWaiWen2 As String
Private AWaiWen3 As String
Private AWaiWen4 As String

Private SlaveID As String
Private Interval As String

Private saved As Integer

Private KandC(VarNum) As String
Private req(VarNum) As String
Private resp(VarNum) As String
Private dec(VarNum) As Integer
Private para(VarNum) As Double

Private paraList(FuNum, VarNum) As Double



Private index As Integer
Private FuLoc(9) As Integer













Private Sub Check1_Click()
    If Check1 = 1 Then
        Text16.Text = dec(0) & "/4095 = " & CDbl(dec(0)) / 4095
        Text17.Text = dec(1) & "/4095 = " & CDbl(dec(1)) / 4095
        Text18.Text = dec(2) & "/4095 = " & CDbl(dec(2)) / 4095
        Text19.Text = dec(3) & "/4095 = " & CDbl(dec(3)) / 4095
        Text20.Text = dec(4) & "/4095 = " & CDbl(dec(4)) / 4095
        Text21.Text = dec(5) & "/4095 = " & CDbl(dec(5)) / 4095
    Else
        Text16.Text = resp(0)
        Text17.Text = resp(1)
        Text18.Text = resp(2)
        Text19.Text = resp(3)
        Text20.Text = resp(4)
        Text21.Text = resp(5)
    End If
End Sub

Private Sub Check2_Click()

If Check2 = 1 Then
On Error GoTo EH
    If MSComm1.PortOpen = True Then
        If Check2 = 1 Then
       
                DB_Connect
            
                Timer1.Enabled = False
                Timer1.Interval = 1000 * Val(Text28.Text)
                Timer1.Enabled = True
                MsgBox "测试连接成功，记录已打开"
        Else
            Timer1.Enabled = False
            MsgBox "记录已关闭"
        End If
    Else
        MsgBox "请打开串口"
        Check2 = 0
    End If
    
End If

    Exit Sub
EH:
        MsgBox ("数据库配置，时钟或串口错误，打开失败")
        Check2 = 0

End Sub

Private Sub Check3_Click()

On Error GoTo EH
    If MSComm1.PortOpen = True Then
        If Check3 = 1 Then
       
                Timer2.Enabled = False
                Timer2.Interval = Val(Text14.Text)
                Timer2.Enabled = True
                MsgBox "轮询刷新已打开"
        Else
            Timer2.Enabled = False
            MsgBox "轮询刷新已关闭"
        End If
    Else
        MsgBox "请打开串口"
        Check3 = 0
    End If

    Exit Sub
EH:
        MsgBox ("时钟或串口错误")
        Check3 = 0

End Sub

Private Sub Combo2_Click()
    FuRefresh
End Sub

Private Sub Command1_Click()
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

Private Sub Command2_Click()
    Save
    refreshText
End Sub

Private Sub Command3_Click()
    LoadOld
    
    refreshText
End Sub

Private Sub Command4_Click()
    
    If MSComm1.PortOpen = False Then
        SerialPortCheck
    Else
        MsgBox "请先断开串口连接"
    End If
End Sub

Private Sub Command5_Click()
    StartSending
    
End Sub


Private Sub Fini()
    dec(0) = Val("&H" & Mid(resp(0), 7, 4))
    dec(1) = Val("&H" & Mid(resp(1), 7, 4))
    dec(2) = Val("&H" & Mid(resp(2), 7, 4))
    dec(3) = Val("&H" & Mid(resp(3), 7, 4))
    dec(4) = Val("&H" & Mid(resp(4), 7, 4))
    dec(5) = Val("&H" & Mid(resp(5), 7, 4))
    
    Dim t As Integer

    
    On Error Resume Next
    For t = 0 To VarNum - 1 Step 1
        
        a = Split(KandC(t), "x+")
        para(t) = CDbl(dec(t)) / 4095 * CDbl(a(0)) + CDbl(a(1))
        'MsgBox (a(0) & ":" & a(1))
    Next
    
    
    Text22.Text = para(0)
    Text23.Text = para(1)
    Text24.Text = para(2)
    Text25.Text = para(3)
    Text26.Text = para(4)
    Text27.Text = para(5)
    
    
    If Check1 = 1 Then
        Text16.Text = dec(0) & "/4095 = " & CDbl(dec(0)) / 4095
        Text17.Text = dec(1) & "/4095 = " & CDbl(dec(1)) / 4095
        Text18.Text = dec(2) & "/4095 = " & CDbl(dec(2)) / 4095
        Text19.Text = dec(3) & "/4095 = " & CDbl(dec(3)) / 4095
        Text20.Text = dec(4) & "/4095 = " & CDbl(dec(4)) / 4095
        Text21.Text = dec(5) & "/4095 = " & CDbl(dec(5)) / 4095
    Else
        Text16.Text = resp(0)
        Text17.Text = resp(1)
        Text18.Text = resp(2)
        Text19.Text = resp(3)
        Text20.Text = resp(4)
        Text21.Text = resp(5)
    End If
    

End Sub


Private Sub Command6_Click()
    Setting1.Show
End Sub

Private Sub Form_Load()
    
    conn = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Text4.Text & ";Password =" & Text5.Text & "; Initial Catalog=" & Text6.Text & ";Data Source=" & Text1.Text
    CONNECT_LOOP_MAX = 10
    
    '7条釜
     
    Combo2.AddItem ("1")
    Combo2.AddItem ("2")
    Combo2.AddItem ("3")
    Combo2.AddItem ("4")
    Combo2.AddItem ("5")
    Combo2.AddItem ("6")
    Combo2.AddItem ("7")
    
    Combo2.Text = "1"
    
    SerialPortCheck
    
    FuRefresh
    

End Sub


Private Sub FuRefresh()

    FuNumber = Val(Combo2.Text)
    
    
    index = 7
    
    saved = 0

    LoadOld
    
    refreshText
    
    FuLoc(SlaverID) = FuNumber

    MSComm1.RThreshold = 1
    
    MSComm1.InputMode = comInputModeBinary '二进制接收
    
    For i = 0 To VarNum Step 1
        
        para(i) = 0
        dec(i) = 0
    Next

    Text16.Text = 0
    Text17.Text = 0
    Text18.Text = 0
    Text19.Text = 0
    Text20.Text = 0
    Text21.Text = 0
    Text22.Text = 0
    Text23.Text = 0
    Text24.Text = 0
    Text25.Text = 0
    Text26.Text = 0
    Text27.Text = 0
    

End Sub



Private Sub Form_Unload(Cancel As Integer)

    myexit = MsgBox("请再次确认是否退出服务! 可能会导致数据记录中断!", vbExclamation + vbYesNo + vbDefaultButton2, "退出确认")
    If myexit = vbNo Then
        Cancel = True
    Else
    
        If saved = 0 Then
            myexit = MsgBox("是否保存已修改参数", vbExclamation + vbYesNo + vbDefaultButton2, "保存确认")
            
            '不保存
            If myexit = vbNo Then
                
            End If
            
            If myexit = vbYes Then
                Save
            End If
        
        End If
    End If
    
End Sub


Private Sub Save()

    On Error Resume Next
    
    saved = 1
    YaLiReg = Text2.Text
    NeiWenReg = Text3.Text
    WaiWen1Reg = Text4.Text
    WaiWen2Reg = Text5.Text
    WaiWen3Reg = Text6.Text
    WaiWen4Reg = Text7.Text
    KandC(0) = Text8.Text
    KandC(1) = Text9.Text
    KandC(2) = Text10.Text
    KandC(3) = Text11.Text
    KandC(4) = Text12.Text
    KandC(5) = Text13.Text
    
    SlaveID = Text1.Text
    Interval = 1000

    
    Open App.Path & "\config\data" & FuNumber & ".txt" For Output As #1
        Write #1, YaLiReg, NeiWenReg, WaiWen1Reg, WaiWen2Reg, WaiWen3Reg, WaiWen4Reg, KandC(0), KandC(1), KandC(2), KandC(3), KandC(4), KandC(5), SlaveID, Interval

    Close #1
    MsgBox "参数已保存！"
End Sub

Private Sub refreshText()
    
    Text2.Text = YaLiReg
    Text3.Text = NeiWenReg
    Text4.Text = WaiWen1Reg
    Text5.Text = WaiWen2Reg
    Text6.Text = WaiWen3Reg
    Text7.Text = WaiWen4Reg
    Text8.Text = KandC(0)
    Text9.Text = KandC(1)
    Text10.Text = KandC(2)
    Text11.Text = KandC(3)
    Text12.Text = KandC(4)
    Text13.Text = KandC(5)
    
    Text1.Text = SlaveID

End Sub

Private Sub SerialPortCheck()

    
    '检测当前有效串口
    On Error Resume Next

    Combo1.Clear
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

Private Sub LoadOld()
    '读取历史设置值
    On Error Resume Next
    
    Open App.Path & "\config\data" & FuNumber & ".txt" For Input As #1
        Input #1, YaLiReg, NeiWenReg, WaiWen1Reg, WaiWen2Reg, WaiWen3Reg, WaiWen4Reg, KandC(0), KandC(1), KandC(2), KandC(3), KandC(4), KandC(5), SlaveID, Interval
    Close #1

End Sub



Private Function Send()
    
    On Error Resume Next
    
    Dim ByteChar(8) As Byte
   
    ByteChar(0) = "&H" & Mid(req(index), 1, 2)
    ByteChar(1) = "&H" & Mid(req(index), 4, 2)
    ByteChar(2) = "&H" & Mid(req(index), 7, 2)
    ByteChar(3) = "&H" & Mid(req(index), 10, 2)
    ByteChar(4) = "&H" & Mid(req(index), 13, 2)
    ByteChar(5) = "&H" & Mid(req(index), 16, 2)
    ByteChar(6) = "&H" & Mid(req(index), 19, 2)
    ByteChar(7) = "&H" & Mid(req(index), 22, 2)
    
    '01 03 00 00 00 01 84 0A
    'ByteChar(0) = &H1
    'ByteChar(1) = &H3
    'ByteChar(2) = &H0
    'ByteChar(3) = &H0
    'ByteChar(4) = &H0
    'ByteChar(5) = &H1
    'ByteChar(6) = &H84
    'ByteChar(7) = &HA
    
    
    MSComm1.InBufferCount = 0       '"清除发送缓冲区数据
    MSComm1.OutBufferCount = 0       '"清除接收缓冲区数据

    Text15.Text = req(index)
    MSComm1.Output = ByteChar
       
    
End Function

Private Sub Label21_Click()

End Sub

Private Sub MSComm1_OnComm()

    On Error Resume Next
    
    Dim strData As String
    Dim bytInput() As Byte
    Dim intInputLen As Integer

    
    

    MSComm1.InputMode = comInputModeBinary '二进制接收
    intInputLen = MSComm1.InBufferCount
    
    Select Case Me.MSComm1.CommEvent
    Case comEvReceive
    
        ReDim bytInput(intInputLen)
        bytInput = MSComm1.Input
        Dim i As Integer
        For i = 0 To UBound(bytInput)
            If Len(Hex(bytInput(i))) = 1 Then
                strData = strData & "0" & Hex(bytInput(i))
            Else
                strData = strData & Hex(bytInput(i))
            End If
        Next
        
        If Len(strData) = 14 Then
        
            If index < VarNum Then
                resp(index) = strData
                If index < VarNum - 1 Then
                    index = index + 1
                    'MsgBox (index & " - " & strData)
                    Send
                End If
                
                If index = VarNum - 1 Then
                    Fini
                End If
                
            End If
            
        Else
            Send
        End If
             
    End Select
End Sub




Private Sub Text16_Click()
    Text15.Text = req(0)
End Sub
Private Sub Text17_Click()
    Text15.Text = req(1)
End Sub
Private Sub Text18_Click()
    Text15.Text = req(2)
End Sub
Private Sub Text19_Click()
    Text15.Text = req(3)
End Sub
Private Sub Text20_Click()
    Text15.Text = req(4)
End Sub
Private Sub Text21_Click()
    Text15.Text = req(5)
End Sub

Private Sub Timer1_Timer()
    'MsgBox "timer"
    
    
    
    
End Sub

Private Sub StartSending()
    If MSComm1.PortOpen = False Then
        MsgBox "串口未开启"
    Else
        req(0) = GenerateReq(SlaveID, YaLiReg)
        req(1) = GenerateReq(SlaveID, NeiWenReg)
        req(2) = GenerateReq(SlaveID, WaiWen1Reg)
        req(3) = GenerateReq(SlaveID, WaiWen2Reg)
        req(4) = GenerateReq(SlaveID, WaiWen3Reg)
        req(5) = GenerateReq(SlaveID, WaiWen4Reg)
        
        index = 0
        Send
        
    End If
End Sub

Private Sub Timer2_Timer()
    If FuNumber >= FuNum Then
        FuNumber = 1
        
    Else
        FuNumber = FuNumber + 1
    End If
    'MsgBox (FuNumber)
    Combo2 = FuNumber
    FuRefresh
    StartSending
        
End Sub
