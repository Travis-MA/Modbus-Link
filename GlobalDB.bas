Attribute VB_Name = "GlobalDB"

Public conn As String
Public CONNECT_LOOP_MAX

Private IsConnect As Boolean  '������ݿ��Ƿ�����

Private Connect_Num As Integer  '���ִ��Connect()������������ݵĴ���

Private cnn As ADODB.Connection '�������ݿ��Connect����
 
Private re As ADODB.Recordset  '����������Recordset����

'�������ݿ�
 Private Sub Connect()
 '������ӱ��Ϊ�棬�򷵻ء�
  If IsConnect = True Then
     Exit Sub
  End If

  Set cnn = New ADODB.Connection '�ؼ�new���ڴ����¶���cnn
  
  cnn.ConnectionString = conn
  
  cnn.Open
  '�ж����ӵ�״̬
  If cnn.State <> adStateOpen Then
     MsgBox "���ݿ�����ʧ��"
     End
  End If

  '�������ӱ�ʶ����ʾ�Ѿ����ӵ����ݿ�
  IsConnect = True
End Sub


'�Ͽ������ݿ������
Private Sub DisConnect()
  Dim rc As Long

  If IsConnect = False Then
     Exit Sub
  End If
  '�ر�����
  cnn.Close
  '�ͷ�cnn
  Set cnn = Nothing
  IsConnect = False
End Sub

'ʹ��Connect_Num������������
Public Sub DB_Connect()
   Connect_Num = Connect_Num + 1
   Connect
End Sub

'ʹ��Connect_Num�������ݶϿ�
Public Sub DB_Disconnect()
If Connect_Num >= CONNECT_LOOP_MAX Then
   Connect_Num = 0
   DisConnect
 End If
 End Sub

'ǿ�ƹر�api��ʽ���ʶ�����ݿ⣬��������λ
Public Sub DBapi_Disconnect()
   Connect_Num = 0
   DisConnect
End Sub

'ִ�����ݿ��������
'byval ���ǰ�������ֵ���ݣ��ٴ��ݹ����У��������ᷢ���仯(Ҳ���ǽ�����ֵ�����ǽ���ַ���ݸ����̵ķ�ʽ�����ʹ���̷��ʷ�Ŷ�����ĸ��������̲��ɸı������ֵ)��
'��֮��Ӧ����byref,ָ�������ĵ�ַ��ֵ��byref����ʡ��

Public Sub SQLExt(ByVal TmpSQLstmt As String)

    Dim cmd As New ADODB.Command '����Command����cmd
    
    DB_Connect  '�������ݿ�
    
    Set cmd.ActiveConnection = cnn '����cmd��ActiveConnect���ԣ�ָ��������������ݿ�����

    cmd.CommandText = TmpSQLstmt '����Ҫִ�е������ı�

    'MsgBox TmpSQLstmt

    cmd.Execute  'ִ������

    Set cmd = Nothing

    DB_Disconnect  '�Ͽ������ݿ������

End Sub

'ִ�����ݿ��ѯ���
Public Function QueryExt(ByVal TmpSQLstmt As String) As ADODB.Recordset
    
    Dim rst As New ADODB.Recordset  '����Rescordset����rst

    DB_Connect  '�������ݿ�

    Set rst.ActiveConnection = cnn  '����rst��ActiveConnection����,ָ��������ص����ݿ������

    rst.CursorType = adOpenDynamic  '�����α�����

    rst.LockType = adLockOptimistic  '������������

    rst.Open TmpSQLstmt '�򿪼�¼��

    Set QueryExt = rst '���ؼ�¼��

    End Function

