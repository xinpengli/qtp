On error resume next 
'========加载测试数据---测试用
'datatable.ImportSheet "..\..\Sumitomo\TestData\InsertLockReportSource.xls",1,"Global"
'执行InsertDB函数
executefile  "..\..\Sumitomo\Func&VBS\DBFunc.txt"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
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
'========定义并获取源码消息帧长度、信息类型、版本号
Dim MsgHeadPart 
MsgHeadPart=Datatable("MsgHeadPart","Global")
Dim MsgFlag '定义并获取源码信息头参数中的信息类型
MsgFlag=mid(Datatable("MsgHeadPart","Global"),3,2)
'========定义并获取信息生成时间
Dim InfoGeneTime 
InfoGeneTime=Cstr(now)  '获取当前时间
Datatable("InfoGeneTime","Global")=InfoGeneTime '存储当前时间的十进制形式,用于后续页面相关时间检查
InfoGeneTimeHex=right("0"& hex(right(year(InfoGeneTime),2)),2) & right("0"& hex(month(InfoGeneTime)),2) & right("0"& hex(day(InfoGeneTime)),2) _
&right("0"& hex(hour(InfoGeneTime)),2) & right("0"& hex(minute(InfoGeneTime)),2) & right("0"& hex(second(InfoGeneTime)),2)   '当前时间十六进制形式
'如果用例是Case4_1_VehiMsg_MsgQuery_Timing,需要存储十六进制串，因故障需要用到和定时一致的信息生成时间
Dim RelaPath '定义相对路径
Dim fso '定义FSO对象
if(Environment("TestName")="Case4_1_VehiMsg_MsgQuery_Timing")then 
	'创建FSO对象
	Set fso=createobject("Scripting.filesystemobject")
	RelaPath = PathFinder.Locate("Sumitomo") &"\DownFiles\"
	'如果文件已存在，则删除重新创建并写入
	if(fso.FileExists(RelaPath &"TimingInfoGeneTime.txt"))then
		fso.DeleteFile(RelaPath &"TimingInfoGeneTime.txt")
	end if
	'重新创建文件
	Set txtFile=fso.CreateTextFile(RelaPath &"TimingInfoGeneTime.txt",true)
	'记录定时用例中源码的信息生成时间
	txtFile.Write InfoGeneTimeHex&InfoGeneTime
	'关闭txtFile对象
	txtFile.Close
	'释放对象
	Set txtFile=nothing
	Set fso=nothing
end if
'========定义并获取源码其它内容
Dim OtherContent
OtherContent=Datatable("OtherContent","Global")
'========定义并获取检验码
Dim DayMonYear
DayMonYear=right("0"&Day(Date-1),2) & right("0"&Month(Date),2)  & right(Year(Date),2)
Datatable("DayMonYear","Global")=Year(Date) &"-"&  right("0"&Month(Date),2)  &"-"&  right("0"&Day(Date-1),2) '用于页面保养所属日期的检查
Dim TempStr
'除了日志和保养信息源码需要拼接日期外，其它的不需要
if(MsgFlag="20" or MsgFlag="26")then
TempStr=MsgHeadPart & InfoGeneTimeHex & DayMonYear & OtherContent
else
     '其它类型源码拼接
     '首先关于10故障报警信息需要信息生成时间和定时一致，因页面有些字段取自定时，故直接取定时的信息生成时间来用
	 if(MsgFlag="10")then
		 '先获取定时源码中的信息生成时间
		 Set fso1=createobject("Scripting.filesystemobject")
		 RelaPath =PathFinder.Locate("Sumitomo") &"\DownFiles\" '获取相对路径
		 Set txtfile1=fso1.OpenTextFile(RelaPath&"TimingInfoGeneTime.txt",1,true)  '1表示读模式
		 TempStr=txtfile1.ReadAll
		 InfoGeneTimeHex=left(TempStr,12)   '关于故障重新存储16进制的信息生成时间
		 Datatable("InfoGeneTime","Global")=right(TempStr,len(TempStr)-12)  '关于故障重新存储10进制的信息生成时间
		 txtfile1.Close
		 Set fso1=nothing
		 Set txtfile1=nothing
	end if
	'其它类源码正常拼接 
	TempStr=MsgHeadPart & InfoGeneTimeHex  & OtherContent
end if
'定义和获取 校验码和结束符
Dim CheckCodeAndOver
CheckCodeAndOver=CreatCheckCode(TempStr)&"00"
'========拼接源码
Dim StrSource
StrSource=TempStr & CheckCodeAndOver
'========插入数据库源码
Dim sqlInsert
sqlInsert="INSERT INTO [cmppSum].[dbo].[CMPPReceivalNew] VALUES((select isnull(MAX(ReceivalID),0) from [cmppSum].[dbo].[CMPPReceivalNew]) + 1,'70001',2,1,'10657509110066',2,1,'"+Datatable("SIMCardNo","Global")+"',0,0,0,0,245,'',24,'"+StrSource+"',0,'',null,null,'',GETDATE(),null,null)"
'执行sql并返回结果
Dim RetuVal
RetuVal=ExecDB(sqlInsert) 
'根据执行结果写日志
if(RetuVal>=0)then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"insert"& Datatable("InputMsgType","Global") &"源码成功","insert"& Datatable("InputMsgType","Global") &"源码成功"
else
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"insert"& Datatable("InputMsgType","Global") &"源码失败","insert"& Datatable("InputMsgType","Global") &"源码失败"
end if
wait 1 '等待解析时间
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If
