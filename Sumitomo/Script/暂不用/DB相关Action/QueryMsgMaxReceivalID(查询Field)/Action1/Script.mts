On error resume next
Dim adoConn '定义ADO连接对象
Dim ConnectionStr '定义数据库连接字符串
'获取数据库连接字符串
ConnectionStr="Driver={SQL SERVER};SERVER=192.168.30.172\qctest;UID=sa;PWD=TYKJ66tykj;DATABASE=Sumitomo;PORT="
'获取数据库查询语句
Dim sqlQuery
sqlQuery="select MAX(ReceivalID) as ReceivalID  from cmppSum.dbo.CMPPReceivalNew  where OriginalAddress='"&Datatable("SIMCardNo","Global")&"'"
'创建数据库连接对象
Set adoConn=CreateObject("adodb.Connection")
'打开数据库
adoConn.Open ConnectionStr
'执行sql返回对应的结果集
Set adoRst=adoConn.Execute(sqlQuery)
'如果查询结果集为空即BOF或EOF为真，则不进行如下update操作
if(Not(adoRst.BOF or adoRst.EOF))then
	'获得结果集中ID\时间设置字段的值
	Datatable("ReceivalID","Global")=adoRst.Fields.Item("ReceivalID").Value
end if
'关闭数据库
adoConn.Close
'释放数据库对象
Set adoConn=nothing
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If