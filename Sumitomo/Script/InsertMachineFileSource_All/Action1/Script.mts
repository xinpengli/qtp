On error resume next 
'执行DB交互的函数文件
executefile  "..\..\Sumitomo\Func&VBS\DBFunc.txt"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'========加载测试数据---测试用
'datatable.ImportSheet "..\..\Sumitomo\TestData\InsertLockReportSource.xls",1,"Global"
'源码中校验码的计算
Function CreatCheckCode(message)
	 if (message = null)then
	 CreatCheckCode=""
	end if
	Dim sum
	sum= 0
	For i=1 to len(message) step 2
		sum=sum+clng("&H" & mid(message,i,2) )   'VBS16进制转化为10进制
	Next
	Dim hexSum 
	hexSum=Hex(sum)     'VBS10进制转化为16进制
	CreatCheckCode=right(hexSum,2)
End Function
'========获取源码插入次数,1-初期设定指令x30; 2-初期设定完成信息指令x32; 3-全复位通知信息x1A指令; 4-机器Touch通知信息xA2指令
Dim i
i=Parameter("i")
'========定义并获取源码消息帧长度、信息类型、版本号
Dim MsgHeadPart 
MsgHeadPart=Datatable("MsgHeadPart"&i,"Global")
'========定义并获取信息生成时间
Dim InfoGeneTime 
InfoGeneTime=Cstr(now)  '获取当前时间
Datatable("InfoGeneTime"&i,"Global")=InfoGeneTime
InfoGeneTimeHex=right("0"& hex(right(year(InfoGeneTime),2)),2) & right("0"& hex(month(InfoGeneTime)),2) & right("0"& hex(day(InfoGeneTime)),2) _
&right("0"& hex(hour(InfoGeneTime)),2) & right("0"& hex(minute(InfoGeneTime)),2) & right("0"& hex(second(InfoGeneTime)),2)   '当前时间十六进制形式
if(i=3)then
Datatable("VehicleInfoTreeView_item_55","Action1")="复位通知信息                "&Datatable("InfoGeneTime3","Global") & " AllReset"
end if
if(i=4)then
'框架页机器档案中的点检实绩时间预期值
Datatable("VehicleInfoTreeView_item_56","Action1")="点检实绩                        "& Datatable("InfoGeneTime4","Global") &"  河北省石家庄市裕华区 黄河大道附近"  
 '点检实绩页面的点检时间预期值
Datatable("InspeRes_CheckTime","Action1")=Datatable("InfoGeneTime4","Global")   
end if
'========获取对照码
if(i=2)then
'	'========定义并获取对照码
'	Dim adoConn_t            '定义ADO连接对象
'	Dim ConnectionStr_t '定义数据库连接字符串
'	ConnectionStr_t="Driver={SQL SERVER};SERVER=192.168.30.172\qctest;UID=sa;PWD=TYKJ66tykj;DATABASE=Sumitomo;PORT="  '获取数据库连接字符串
'	Dim sqlQuery_t         '获取数据库查询语句
'	sqlQuery_t="SELECT top 1 UserData  FROM SendToTerminal.dbo.CMPPSendSubAll where destinationAddress='"&right(Datatable("SIMCardNo","Global"),11)&"'  order by MsgID desc"
'	Set adoConn_t=CreateObject("adodb.Connection")      '创建数据库连接对象
'	adoConn_t.Open ConnectionStr_t                                                 '打开数据库
'	wait 4 '叫数据编码入库慢
'	Set adoRst_t=adoConn_t.Execute(sqlQuery_t)                        '执行sql返回对应的结果集
'	Datatable("ControlCode"&i,"Global")=mid(adoRst_t.Fields.Item("UserData").Value,5,6)        '获得结果集中源码串中的对照码	
'	adoConn_t.Close    '关闭数据库
'	Set adoConn_t=nothing     '释放数据库对象
	Dim sqlQuery_t         '获取数据库查询语句
	sqlQuery_t="SELECT top 1 UserData  FROM SendToTerminal.dbo.CMPPSendSubAll where destinationAddress='"&right(Datatable("SIMCardNo","Global"),11)&"'  order by MsgID desc"
    Datatable("ControlCode"&i,"Global")=mid(QueryDBColumn(sqlQuery_t,"UserData"),5,6)
end if
'========定义并获取源码其它内容
Dim OtherContent
OtherContent=Datatable("OtherContent"&i,"Global")
if(i=1)then
'如果是第1组源码,即初期设定指令x30,需要将其中的发动机编号替换为实际的测试发动机编号,和中心管理中车辆信息的发动机编号一致,否则可能会造成拆车
OtherContent=replace(OtherContent,mid(OtherContent,51,12),Datatable("EgnNo","Global")) 
end if
'========定义并获取检验码,当插入第2组源码x32时，需要拼接对照码，其它组源码不需要
Dim TempStr
if(i<>2)then
TempStr=MsgHeadPart  & InfoGeneTimeHex & OtherContent
else
TempStr=MsgHeadPart  & InfoGeneTimeHex & Datatable("ControlCode"&i,"Global") & OtherContent
end if
'定义和获取 校验码和结束符
Dim CheckCodeAndOver
CheckCodeAndOver=CreatCheckCode(TempStr)&"00"
'========拼接源码
Dim StrSource
StrSource=TempStr & CheckCodeAndOver
'========插入数据库源码，定义并获取insert语句
Dim sqlInsert
sqlInsert="INSERT INTO [cmppSum].[dbo].[CMPPReceivalNew] VALUES((select isnull(MAX(ReceivalID),0) from [cmppSum].[dbo].[CMPPReceivalNew]) + 1,'70001',2,1,'10657509110066',2,1,'"+Datatable("SIMCardNo","Global")+"',0,0,0,0,245,'',24,'"+StrSource+"',0,'',null,null,'',GETDATE(),null,null)"
'执行sql并返回结果
Dim RetuVal
RetuVal=ExecDB(sqlInsert) 
'根据执行结果写日志
if(RetuVal>=0)then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"insert"& mid(MsgHeadPart,3,2) &"源码成功","insert"& mid(MsgHeadPart,3,2) &"源码成功"
else
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"insert"& mid(MsgHeadPart,3,2) &"源码失败","insert"& mid(MsgHeadPart,3,2) &"源码失败"
end if
wait  4 '等待解析时间
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If
