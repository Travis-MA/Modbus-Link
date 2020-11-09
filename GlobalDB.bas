Attribute VB_Name = "GlobalDB"

Public conn As String
Public CONNECT_LOOP_MAX

Private IsConnect As Boolean  '标记数据库是否连接

Private Connect_Num As Integer  '标记执行Connect()函数后访问数据的次数

Private cnn As ADODB.Connection '连接数据库的Connect对象
 
Private re As ADODB.Recordset  '保存结果集的Recordset对象

'连接数据库
 Private Sub Connect()
 '如果连接标记为真，则返回。
  If IsConnect = True Then
     Exit Sub
  End If

  Set cnn = New ADODB.Connection '关键new用于创建新对象cnn
  
  cnn.ConnectionString = conn
  
  cnn.Open
  '判断连接的状态
  If cnn.State <> adStateOpen Then
     MsgBox "数据库连接失败"
     End
  End If

  '设置连接标识，表示已经连接到数据库
  IsConnect = True
End Sub


'断开与数据库的连接
Private Sub DisConnect()
  Dim rc As Long

  If IsConnect = False Then
     Exit Sub
  End If
  '关闭连接
  cnn.Close
  '释放cnn
  Set cnn = Nothing
  IsConnect = False
End Sub

'使用Connect_Num控制数据连接
Public Sub DB_Connect()
   Connect_Num = Connect_Num + 1
   Connect
End Sub

'使用Connect_Num控制数据断开
Public Sub DB_Disconnect()
If Connect_Num >= CONNECT_LOOP_MAX Then
   Connect_Num = 0
   DisConnect
 End If
 End Sub

'强制关闭api方式访问俄的数据库，计数器复位
Public Sub DBapi_Disconnect()
   Connect_Num = 0
   DisConnect
End Sub

'执行数据库操作语言
'byval 就是按参数的值传递，再传递过程中，参数不会发生变化(也就是将参数值而不是将地址传递给过程的方式，这就使过程访问发哦变量的副本，过程不可改变变量的值)；
'与之对应的是byref,指按参数的地址传值，byref可以省略

Public Sub SQLExt(ByVal TmpSQLstmt As String)

    Dim cmd As New ADODB.Command '创建Command对象cmd
    
    DB_Connect  '连接数据库
    
    Set cmd.ActiveConnection = cnn '设置cmd的ActiveConnect属性，指定与其关联的数据库连接

    cmd.CommandText = TmpSQLstmt '设置要执行的命令文本

    'MsgBox TmpSQLstmt

    cmd.Execute  '执行命令

    Set cmd = Nothing

    DB_Disconnect  '断开与数据库的连接

End Sub

'执行数据库查询语句
Public Function QueryExt(ByVal TmpSQLstmt As String) As ADODB.Recordset
    
    Dim rst As New ADODB.Recordset  '创建Rescordset对象rst

    DB_Connect  '连接数据库

    Set rst.ActiveConnection = cnn  '设置rst的ActiveConnection属性,指定与其相关的数据库的连接

    rst.CursorType = adOpenDynamic  '设置游标类型

    rst.LockType = adLockOptimistic  '设置锁定类型

    rst.Open TmpSQLstmt '打开记录集

    Set QueryExt = rst '返回记录集

    End Function

