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
'========获取入参，即源码操作次数
Dim i
i=Parameter("i")
'========定义并获取源码消息帧长度、信息类型、版本号
Dim MsgHeadPart 
MsgHeadPart=Datatable("MsgHeadPart"&i,"Global")
'========定义并获取信息生成时间
Dim InfoGeneTime
'信息生成时间与统计程序执行日期相同,因为跑的是日志和压力分布信息
if(i<3)then
InfoGeneTime=Cstr(now-1)  '获取信息生成时间为当天-1
else
	if(i<5)then
	InfoGeneTime=Cstr(now)  '获取信息生成时间为当天
	else
	InfoGeneTime=Cstr(now+1)  '获取信息生成时间为当天加1
	end if
end if
Datatable("InfoGeneTime"&i,"Global")=InfoGeneTime '存储当前时间的十进制形式,用于后续页面相关时间检查
InfoGeneTimeHex=right("0"& hex(right(year(InfoGeneTime),2)),2) & right("0"& hex(month(InfoGeneTime)),2) & right("0"& hex(day(InfoGeneTime)),2) _
&right("0"& hex(hour(InfoGeneTime)),2) & right("0"& hex(minute(InfoGeneTime)),2) & right("0"& hex(second(InfoGeneTime)),2)   '当前时间十六进制形式
'========定义并获取源码其它内容
Dim OtherContent
OtherContent=Datatable("OtherContent"&i,"Global")
'替换日志指令和压力分布指令的内容串中的日期字节
InfoGeneTime=Cstr(Cdate(InfoGeneTime)-1)  '源码中的日期需要早于信息生成时间1天
Dim SourDate
SourDate=right("0"&Day(InfoGeneTime),2) & right("0"&Month(InfoGeneTime),2) & right(year(InfoGeneTime),2)
OtherContent=replace(OtherContent,left(OtherContent,6),SourDate)
'========定义并获取检验码
Dim TempStr
TempStr=MsgHeadPart & InfoGeneTimeHex & OtherContent
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
wait 2 '等待解析时间
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If