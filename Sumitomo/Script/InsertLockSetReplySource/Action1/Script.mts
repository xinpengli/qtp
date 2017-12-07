On error resume next 
'执行DB交互的函数文件
executefile  "..\..\Sumitomo\Func&VBS\DBFunc.txt"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'========加载测试数据---测试用
'datatable.ImportSheet "..\..\Sumitomo\TestData\InsertLockSetReplySource.xls",1,"Global"
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
Dim LockSetRepSour_MsgHeadPart 
LockSetRepSour_MsgHeadPart=Datatable("LockSetRepSour_MsgHeadPart","Global")
'========定义并获取信息生成时间
Dim LockSetRepSour_InfoGeneTime 
LockSetRepSour_InfoGeneTime=Cstr(now)  '获取当前时间
Datatable("LockSetRepSour_InfoGeneTime","Global")=LockSetRepSour_InfoGeneTime '存储当前时间的十进制形式,用于后续页面相关时间检查
LockSetRepSour_InfoGeneTimeHex=right("0"& hex(right(year(LockSetRepSour_InfoGeneTime),2)),2) & right("0"& hex(month(LockSetRepSour_InfoGeneTime)),2) & right("0"& hex(day(LockSetRepSour_InfoGeneTime)),2) _
&right("0"& hex(hour(LockSetRepSour_InfoGeneTime)),2) & right("0"& hex(minute(LockSetRepSour_InfoGeneTime)),2) & right("0"& hex(second(LockSetRepSour_InfoGeneTime)),2)   '当前时间十六进制形式
'========定义并获取对照码
Datatable("LockSetRepSour_ControlCode","Global")=Datatable("ControlCode","Global")                        '存储对照码
'========定义并获取源码其它内容
Dim LockSetRepSour_OtherContent
LockSetRepSour_OtherContent=Datatable("LockSetRepSour_OtherContent","Global")
Dim RepDate '定义替换的日期，即页面设置的循环密码锁日期
Dim CirDate '定义循环日期串
Select Case Datatable("LockType","Global")
    Case "总工作时间锁":
		'暂不需处理
	Case "指定日期锁":
		LockSetRepSour_OtherContent=replace(LockSetRepSour_OtherContent,mid(LockSetRepSour_OtherContent,7,6),left(LockSetRepSour_InfoGeneTimeHex,6))'将源码其它内容串中指定日期替换为信息生成日期，测试数据源码中是写死的日期
	Case "指定位置锁":
		'暂不需处理
	Case "循环日期锁":
        RepDate= right("0"& hex(right(Datatable("CircDateLock_Y","Global"),2)),2) & right("0"& hex(Datatable("CircDateLock_M","Global")),2) &  right("0"& hex(Datatable("CircDateLock_D","Global")),2)
		'将源码其它内容串中循环密码日期替换为设置的日期，测试数据源码中是写死的日期
		LockSetRepSour_OtherContent=replace(LockSetRepSour_OtherContent,mid(LockSetRepSour_OtherContent,7,6),RepDate)
		'将源码中锁车日48个字节,替换为实际的锁车日
		For j=1  to 48
			CirDate=CirDate & right("0"& hex(Datatable("CircDateLock_D","Global")),2)
		Next
		LockSetRepSour_OtherContent=replace(LockSetRepSour_OtherContent,right(LockSetRepSour_OtherContent,96),CirDate)
	Case "立即锁":
		'暂不需处理
	Case "总工作时间锁/指定日期锁/指定位置锁/循环日期锁/立即锁":
		'先替换源码中的指定日期锁日期：将源码其它内容串中指定日期替换为信息生成日期，测试数据源码中是写死的日期
		LockSetRepSour_OtherContent=replace(LockSetRepSour_OtherContent,mid(LockSetRepSour_OtherContent,47,6),left(LockSetRepSour_InfoGeneTimeHex,6))
		'再替换源码中的循环日期锁日期：将源码其它内容串中循环密码日期替换为设置的日期，测试数据源码中是写死的日期
        RepDate= right("0"& hex(right(Datatable("CircDateLock_Y","Global"),2)),2) & right("0"& hex(Datatable("CircDateLock_M","Global")),2) &  right("0"& hex(Datatable("CircDateLock_D","Global")),2)
		LockSetRepSour_OtherContent=replace(LockSetRepSour_OtherContent,mid(LockSetRepSour_OtherContent,63,6),RepDate)
		'将源码中锁车日48个字节,替换为实际的锁车日
		For j=1  to 48
			CirDate=CirDate & right("0"& hex(Datatable("CircDateLock_D","Global")),2)
		Next
		LockSetRepSour_OtherContent=replace(LockSetRepSour_OtherContent,right(LockSetRepSour_OtherContent,96),CirDate)
End Select
'========定义并获取检验码
Dim CheckCodeAndOver
Dim TempStr
TempStr=LockSetRepSour_MsgHeadPart & LockSetRepSour_InfoGeneTimeHex & Datatable("LockSetRepSour_ControlCode","Global") & LockSetRepSour_OtherContent
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
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"insert"& mid(LockSetRepSour_MsgHeadPart,3,2) &"源码成功","insert"& mid(LockSetRepSour_MsgHeadPart,3,2) &"源码成功"
else
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"insert"& mid(LockSetRepSour_MsgHeadPart,3,2) &"源码失败","insert"& mid(LockSetRepSour_MsgHeadPart,3,2) &"源码失败"
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
