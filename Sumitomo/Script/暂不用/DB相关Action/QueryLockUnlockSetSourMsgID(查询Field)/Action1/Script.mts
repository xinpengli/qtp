Dim adoConn '定义ADO连接对象
Dim ConnectionStr '定义数据库连接字符串
'获取数据库连接字符串
ConnectionStr="Driver={SQL SERVER};SERVER=192.168.30.172\qctest;UID=sa;PWD=TYKJ66tykj;DATABASE=Sumitomo;PORT="
'获取数据库查询语句
Dim sqlQuery
sqlQuery="select MAX(msgid) as MsgID  from SendToTerminal.dbo.CMPPSendSubAll where DestinationAddress='"&right(Datatable("SIMCardNo","Global"),11)&"'"
'创建数据库连接对象
Set adoConn=CreateObject("adodb.Connection")
'打开数据库
adoConn.Open ConnectionStr
'执行sql返回对应的结果集
Set adoRst=adoConn.Execute(sqlQuery)
'如果查询结果集为空即BOF或EOF为真，则不进行如下update操作
if(Not(adoRst.BOF or adoRst.EOF))then
	'获得结果集中ID\时间设置字段的值
	Datatable("LockUnlockSetSourMsgID","Global")=adoRst.Fields.Item("MsgID").Value
end if
'关闭数据库
adoConn.Close
'释放数据库对象
Set adoConn=nothing
