On error resume next 
'执行DB交互的函数文件
executefile  "..\..\Sumitomo\Func&VBS\DBFunc.txt"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
''加载测试数据--调试用
''datatable.ImportSheet "..\..\Sumitomo\TestData\InsertUnLockSetReplySource.xls",1,"Global"
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
'''''定义并获取源码消息帧长度、信息类型、版本号
Dim UnLockSetRepSour_MsgHeadPart 
UnLockSetRepSour_MsgHeadPart=Datatable("UnLockSetRepSour_MsgHeadPart","Global")
'''''定义并获取信息生成时间
Dim UnLockSetRepSour_InfoGeneTime 
UnLockSetRepSour_InfoGeneTime=Cstr(now)  '获取当前时间
Datatable("UnLockSetRepSour_InfoGeneTime","Global")=UnLockSetRepSour_InfoGeneTime '存储当前时间的十进制形式,用于后续页面相关时间检查
UnLockSetRepSour_InfoGeneTimeHex=right("0"& hex(right(year(UnLockSetRepSour_InfoGeneTime),2)),2) & right("0"& hex(month(UnLockSetRepSour_InfoGeneTime)),2) & right("0"& hex(day(UnLockSetRepSour_InfoGeneTime)),2) _
&right("0"& hex(hour(UnLockSetRepSour_InfoGeneTime)),2) & right("0"& hex(minute(UnLockSetRepSour_InfoGeneTime)),2) & right("0"& hex(second(UnLockSetRepSour_InfoGeneTime)),2)   '当前时间十六进制形式
'''''定义并获取对照码
Datatable("UnLockSetRepSour_ControlCode","Global")=Datatable("ControlCode","Global")                      '存储对照码
'''''定义并获取源码其它内容
Dim UnLockSetRepSour_OtherContent
UnLockSetRepSour_OtherContent=Datatable("UnLockSetRepSour_OtherContent","Global")
Dim RepDate '定义替换的日期，即页面设置的循环密码锁日期
Select Case Datatable("UnlockType","Global")
    Case "总工作时间锁":
		'解车不需处理
	Case "指定日期锁":
		'解车不需处理
	Case "指定位置锁":
		'解车不需处理
	Case "循环日期锁":
		RepDate= right("0"& hex(right(Datatable("CircDateLock_Y","Global"),2)),2) & right("0"& hex(Datatable("CircDateLock_M","Global")),2)
		'将源码其它内容串中循环密码日期替换为设置的日期，测试数据源码中是写死的日期
		UnLockSetRepSour_OtherContent=replace(UnLockSetRepSour_OtherContent,mid(UnLockSetRepSour_OtherContent,5,4),RepDate)'
	Case "立即锁":
		'解车不需处理
	Case "总工作时间锁/指定日期锁/指定位置锁/循环日期锁/立即锁":
'		'无需替换源码中的循环日期锁日期==源码后边跟4个F
'		UnLockSetRepSour_OtherContent=replace(UnLockSetRepSour_OtherContent,mid(UnLockSetRepSour_OtherContent,13,4),RepDate)
	Case "":
		'全解车场景，默认即可
End Select
'''''定义并获取检验码
Dim CheckCodeAndOver
Dim TempStr
TempStr=UnLockSetRepSour_MsgHeadPart&UnLockSetRepSour_InfoGeneTimeHex& Datatable("UnLockSetRepSour_ControlCode","Global") &UnLockSetRepSour_OtherContent
CheckCodeAndOver=CreatCheckCode(TempStr)&"00"
'''''拼接源码
Dim StrSource
StrSource=TempStr & CheckCodeAndOver
'''''插入数据库源码
Dim sqlInsert   '定义并获取insert语句
sqlInsert="INSERT INTO [cmppSum].[dbo].[CMPPReceivalNew] VALUES((select isnull(MAX(ReceivalID),0) from [cmppSum].[dbo].[CMPPReceivalNew]) + 1,'70001',2,1,'10657509110066',2,1,'"+Datatable("SIMCardNo","Global")+"',0,0,0,0,245,'',24,'"+StrSource+"',0,'',null,null,'',GETDATE(),null,null)"
'执行sql并返回结果
Dim RetuVal
RetuVal=ExecDB(sqlInsert) 
'根据执行结果写日志
if(RetuVal>=0)then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"insert"& mid(UnLockSetRepSour_MsgHeadPart,3,2) &"源码成功","insert"& mid(UnLockSetRepSour_MsgHeadPart,3,2) &"源码成功"
else
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"insert"& mid(UnLockSetRepSour_MsgHeadPart,3,2) &"源码失败","insert"& mid(UnLockSetRepSour_MsgHeadPart,3,2) &"源码失败"
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
