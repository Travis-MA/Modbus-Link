VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����������"
   ClientHeight    =   7620
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5592
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9132.584
   ScaleMode       =   0  'User
   ScaleWidth      =   5391.91
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame2 
      Caption         =   "Զ�����ӺͰ�ȫ"
      Height          =   2892
      Left            =   240
      TabIndex        =   8
      Top             =   4440
      Width           =   5052
      Begin VB.CommandButton Command10 
         Caption         =   "��ͣ����"
         Height          =   372
         Left            =   2520
         TabIndex        =   11
         Top             =   960
         Width           =   2172
      End
      Begin VB.CommandButton Command9 
         Caption         =   "��ȫ��"
         Height          =   372
         Left            =   2520
         TabIndex        =   10
         Top             =   480
         Width           =   2172
      End
      Begin VB.CommandButton Command8 
         Caption         =   "���ݿ�����"
         Height          =   372
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   2052
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����ͨ������"
      Height          =   3972
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5052
      Begin VB.CommandButton Command55 
         Caption         =   "����"
         Height          =   372
         Left            =   3960
         TabIndex        =   15
         Top             =   2400
         Width           =   972
      End
      Begin VB.CommandButton Command44 
         Caption         =   "ˢ��"
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
         Caption         =   "7�Ÿ�"
         Height          =   372
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   2052
      End
      Begin VB.CommandButton Command6 
         Caption         =   "6�Ÿ�"
         Height          =   372
         Left            =   2520
         TabIndex        =   6
         Top             =   1440
         Width           =   2172
      End
      Begin VB.CommandButton Command5 
         Caption         =   "5�Ÿ�"
         Height          =   372
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   2052
      End
      Begin VB.CommandButton Command4 
         Caption         =   "4�Ÿ�"
         Height          =   372
         Left            =   2520
         TabIndex        =   4
         Top             =   960
         Width           =   2172
      End
      Begin VB.CommandButton Command3 
         Caption         =   "3�Ÿ�"
         Height          =   372
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   2052
      End
      Begin VB.CommandButton Command2 
         Caption         =   "2�Ÿ�"
         Height          =   372
         Left            =   2520
         TabIndex        =   2
         Top             =   480
         Width           =   2172
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1�Ÿ�"
         Height          =   372
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   2052
      End
      Begin VB.Label Label2 
         Caption         =   "ָ�����ߴ��ڣ�"
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
    ZYF.Caption = "1�Ÿ�����������"
End Sub

Private Sub Command2_Click()
    ZYF.FuNumber = 2
    Set ZYF.MSComm1 = MSComm1
    ZYF.Hide
    ZYF.Show
    ZYF.Caption = "2�Ÿ�����������"
End Sub

Private Sub Command3_Click()
    ZYF.FuNumber = 3
    ZYF.Hide
    ZYF.Show
    ZYF.Caption = "3�Ÿ�����������"
End Sub

Private Sub Command4_Click()
    ZYF.FuNumber = 4
    ZYF.Hide
    ZYF.Show
    ZYF.Caption = "4�Ÿ�����������"
End Sub

Private Sub Command44_Click()
    SerialPortCheck
End Sub
Private Sub SerialPortCheck()

    '��⵱ǰ��Ч����
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
    ZYF.Caption = "5�Ÿ�����������"
End Sub

Private Sub Command55_Click()
    On Error Resume Next
    If MSComm1.PortOpen = False Then
        MSComm1.CommPort = Right(Combo1.Text, Len(Combo1.Text) - 3)
        MSComm1.PortOpen = True
        If Err.Number <> 0 Then
            MsgBox "�޷��򿪴��ڣ����鴮���Ƿ�ռ�ã�", vbOKOnly, "���ڴ򿪴���"
        Else
            Command1.Caption = "�رմ���"
            nowSerial = MSComm1.CommPort
            MsgBox ("����������COM" & nowSerial & "�Ѵ�" & vbCrLf & "������룺" & Err.Number)
        End If
    Else
        MSComm1.PortOpen = False
        Command1.Caption = "�򿪴���"
        MsgBox "�����ѹر�"
    End If



End Sub

Private Sub Command6_Click()
    ZYF.FuNumber = 6
    ZYF.Hide
    ZYF.Show
    ZYF.Caption = "6�Ÿ�����������"
End Sub

Private Sub Command7_Click()
    ZYF.FuNumber = 7
    ZYF.Hide
    ZYF.Show
    ZYF.Caption = "7�Ÿ�����������"
End Sub



Private Sub Form_load()
    SerialPortCheck
    
End Sub
