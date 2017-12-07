On error resume next 
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'定义源码中实际的参数列表段和期望的参数列表段,如果定义在查询结果集后,有时程序走不到该代码块,此action会略去检查
Dim ParaListSour,ExpParaListSour 
'获取设置的指定日期或循环日期的16进制格式
Dim  AppoDate
'========数据库操作相关,及结果集处理
Dim ConnectionStr  '定义数据库连接字符串 
ConnectionStr="Driver={SQL SERVER};SERVER=192.168.30.172\qctest;UID=sa;PWD=TYKJ66tykj;DATABASE=Sumitomo;PORT="   '获取数据库连接字符串
'获取数据库查询语句1
Dim sqlQuery1
sqlQuery1="SELECT TOP 1 MsgMIS_CmdSign  FROM  Sumitomo.dbo.Msg_MatnInfo_Set   where MsgMIS_Vcl_ID =(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')  order by MsgMIS_ID desc"
Dim adoConn  '定义ADO连接对象
Set adoConn=CreateObject("adodb.Connection")  '创建数据库连接对象
adoConn.Open ConnectionStr  '打开数据库
'执行sql1返回对应的结果集
Dim adoRst
Set adoRst=adoConn.Execute(sqlQuery1)
'判断保养时间设置结果集是否为空，及对照码是否回写
While  adoRst.EOF or adoRst.BOF or  IsNull(adoRst.Fields("MsgMIS_CmdSign"))
	wait 2
	Set adoRst=adoConn.Execute(sqlQuery1)
Wend
Datatable("ControlCode1","Global")=adoRst.Fields.Item("MsgMIS_CmdSign").Value
'获取数据库查询语句2
Dim sqlQuery2
sqlQuery2="SELECT TOP 1 UserData  FROM  SendToTerminal.dbo.CMPPSendSubAll   where UserData like  '%"&Datatable("ControlCode1","Global")&"%'  order by MsgID desc"
'执行sql2返回对应的结果集
Set adoRst=adoConn.Execute(sqlQuery2)
'根据对照码查询源码，不确定对照码是否和源码其它字节重复，如行不通，再考虑根据SIM卡号查询最大信息ID
While adoRst.BOF or adoRst.EOF or  IsNull(adoRst.Fields("UserData"))
	wait 2
	Set adoRst=adoConn.Execute(sqlQuery2)
Wend
'如果查询结果集为空即BOF或EOF为真，则不进行如下update操作
if(Not(adoRst.BOF or adoRst.EOF))then
	'获得结果集中源码字段的值
	Dim UserData
	UserData=adoRst.Fields.Item("UserData").Value
	'截取UserData中锁解车参数列表段, 其中随时值，日期等字节需要后续处理下
	'14表示源码中参数列表的前10位(即消息帧长度、信息类型，对照码)和后4位(校验码和结束符)除去，取中间的参数段，故从11位开始截取
	ParaListSour=mid(UserData,11,len(UserData)-14)
end if
'关闭数据库
adoConn.Close
'释放数据库对象
Set adoConn=nothing
'========检查参数列表源码段=====
ExpParaListSour=Datatable("OtherContent1","Global")
if("00"&ParaListSour=ExpParaListSour)then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养剩余时间设置源码中参数列表检查通过","期望值："& ExpParaListSour &" 实际值：00"& ParaListSour
else
reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养剩余时间设置源码中参数列表检查失败","期望值："& ExpParaListSour &" 实际值：00"& ParaListSour
end if
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If
