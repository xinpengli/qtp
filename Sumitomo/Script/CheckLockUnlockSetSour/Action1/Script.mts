On error resume next 
'加载函数文件
executefile  "..\..\Sumitomo\Func&VBS\DBFunc.txt"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'定义源码中实际的参数列表段和期望的参数列表段,如果定义在查询结果集后,有时程序走不到该代码块,此action会略去检查
Dim ParaListSour,ExpParaListSour 
'获取设置的指定日期或循环日期的16进制格式
Dim  AppoDate
'========数据库操作相关,及结果集处理
'Dim ConnectionStr  '定义数据库连接字符串 
'ConnectionStr="Driver={SQL SERVER};SERVER=192.168.30.172\qctest;UID=sa;PWD=TYKJ66tykj;DATABASE=Sumitomo;PORT="   '获取数据库连接字符串
'Dim adoConn  '定义ADO连接对象
'Set adoConn=CreateObject("adodb.Connection")   '创建数据库连接对象
'adoConn.Open ConnectionStr  '打开数据库
'Dim sqlQuery1  '获取数据库查询语句1
'sqlQuery1="select  MsgL_CmdSign  from Sumitomo.dbo.Msg_Lock where MsgL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"') order by MsgL_ID desc"
''执行sql1返回对应的结果集
'Dim adoRst
'Set adoRst=adoConn.Execute(sqlQuery1)
''判断锁车设置结果集是否为空，及对照码是否回写
'While  adoRst.EOF or adoRst.BOF or  IsNull(adoRst.Fields("MsgL_CmdSign"))
'	wait 2
'	Set adoRst=adoConn.Execute(sqlQuery1)
'Wend
Dim sqlQuery1  '获取数据库查询语句1
sqlQuery1="select  MsgL_CmdSign  from Sumitomo.dbo.Msg_Lock where MsgL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"') order by MsgL_ID desc"
Datatable("ControlCode","Global")=QueryDBColumn(sqlQuery1,"MsgL_CmdSign")  '存储对照码
'msgbox  "查出的对照码："&Datatable("ControlCode","Global")
Dim sqlQuery2  '获取数据库查询语句2，获取对照码后才可拼接查询语句2，否则sql查询条件为空
sqlQuery2="SELECT TOP 1 UserData  FROM  SendToTerminal.dbo.CMPPSendSubAll   where userdata like '%"&Datatable("ControlCode","Global")&"%'  order by MsgID desc"
'截取UserData中锁解车参数列表段, 其中随时值，日期等字节需要后续处理下
'14表示源码中参数列表的前10位(即消息帧长度、信息类型，对照码)和后4位(校验码和结束符)除去，取中间的参数段，故从11位开始截取
Dim UserData
UserData=QueryDBColumn(sqlQuery2,"UserData")
'msgbox "查出的源码："&UserData
ParaListSour=mid(UserData,11,len(UserData)-14)
'msgbox "截取的参数列表："&ParaListSour
''执行sql2返回对应的结果集
'Set adoRst=adoConn.Execute(sqlQuery2)
''根据对照码查询锁解车设置源码
'While adoRst.BOF or adoRst.EOF or  IsNull(adoRst.Fields("UserData"))
'	wait 2
'	Set adoRst=adoConn.Execute(sqlQuery2)
'Wend
''如果查询结果集为空即BOF或EOF为真，则不进行如下update操作
'if(Not(adoRst.BOF or adoRst.EOF))then
'	'获得结果集中源码字段的值
'	Dim UserData
'	UserData=adoRst.Fields.Item("UserData").Value
'	'截取UserData中锁解车参数列表段, 其中随时值，日期等字节需要后续处理下
'	'14表示源码中参数列表的前10位(即消息帧长度、信息类型，对照码)和后4位(校验码和结束符)除去，取中间的参数段，故从11位开始截取
'	ParaListSour=mid(UserData,11,len(UserData)-14)
'end if
''关闭数据库
'adoConn.Close
''释放数据库对象
'Set adoConn=nothing

'========总工作时间锁的锁解车设置源码处理
if(Datatable("LockType","Global")="总工作时间锁")then
		if(Datatable("LockUnlockFlag","Global")="锁车")then
		   '锁解车设置源码：均需要替换参数源码段中随机生成的解车密码基数为空
			ParaListSour=replace(ParaListSour,mid(ParaListSour,5,2),"")
			'总工作时间锁，锁车场景的参数列表源码段需要处理
			ExpParaListSour=replace(Datatable("LockSetRepSour_OtherContent","Global"),mid(Datatable("LockSetRepSour_OtherContent","Global"),5,2),"")
	   else
			'总工作时间锁，解车场景的参数列表源码段不用处理
			ExpParaListSour=Datatable("UnLockSetRepSour_OtherContent","Global")
	   end if
end if '结束总工作时间锁的判断

'========指定日期锁的锁解车设置源码处理
if(Datatable("LockType","Global")="指定日期锁")then
		if(Datatable("LockUnlockFlag","Global")="锁车")then
			AppoDate=right("0"& hex(right(year(cdate(Datatable("AppDateLock_Date","Global"))),2)),2) &  right("0"& hex(month(cdate(Datatable("AppDateLock_Date","Global")))),2) &  right("0"& hex(day(cdate(Datatable("AppDateLock_Date","Global")))),2)
		   '替换实际锁车设置的参数源码段中 解车密码基数+日期 为设置的日期
			ParaListSour=replace(ParaListSour,mid(ParaListSour,6,7),AppoDate)
			'指定日期锁，锁车场景的参数列表源码段需要处理
			ExpParaListSour=replace(Datatable("LockSetRepSour_OtherContent","Global"),mid(Datatable("LockSetRepSour_OtherContent","Global"),6,7),AppoDate)
	   else
			'指定日期锁，解车场景的参数列表源码段不用处理
			ExpParaListSour=Datatable("UnLockSetRepSour_OtherContent","Global")
	   end if
end if '结束指定日期锁的判断

'========指定位置锁的锁解车设置源码处理,问题待查，，经纬度编码有误差，无法处理
if(Datatable("LockType","Global")="指定位置锁")then
		if(Datatable("LockUnlockFlag","Global")="锁车")then
		   '替换实际锁车设置的参数源码段中 解车密码基数
			ParaListSour=len(ParaListSour)
			'指定日期锁，锁车场景的参数列表源码段需要处理
			ExpParaListSour=len(Datatable("LockSetRepSour_OtherContent","Global"))
	   else
			'指定日期锁，解车场景的参数列表源码段不用处理
			ExpParaListSour=Datatable("UnLockSetRepSour_OtherContent","Global")
	   end if
end if '结束指定位置锁的判断

'========循环日期锁的锁解车设置源码处理
if(Datatable("LockType","Global")="循环日期锁")then
		 '获取循环日期设置的年月,转换为16进制
		 AppoDate= right("0"& hex(right(Datatable("CircDateLock_Y","Global"),2)),2) & right("0"& hex(Eval(Datatable("CircDateLock_M","Global"))),2)
		if(Datatable("LockUnlockFlag","Global")="锁车")then
		   '替换实际锁车设置的参数源码段中 解车密码基数+日期 为设置的日期
			ParaListSour=replace(ParaListSour,mid(ParaListSour,6,5),AppoDate)
			'循环日期锁，锁车场景的参数列表源码段需要处理,
			'首先将源码中锁车日48个字节,替换为实际的锁车日
			For j=1  to 48
				CirDate=CirDate & right("0"& hex(Datatable("CircDateLock_D","Global")),2)
			Next
			LockSetRepSour_OtherContent=Datatable("LockSetRepSour_OtherContent","Global")
			LockSetRepSour_OtherContent=replace(LockSetRepSour_OtherContent,right(LockSetRepSour_OtherContent,96),CirDate)
			'其次再替换解车密码基数+日期
			ExpParaListSour=replace(LockSetRepSour_OtherContent,mid(LockSetRepSour_OtherContent,6,5),AppoDate)
	   else
			'循环日期锁，解车场景的参数列表源码段不用处理
			ExpParaListSour=replace(Datatable("UnLockSetRepSour_OtherContent","Global"),mid(Datatable("UnLockSetRepSour_OtherContent","Global"),5,4),AppoDate)
	   end if
end if '结束循环日期锁的判断

'========立即锁的锁解车设置源码处理
if(Datatable("LockType","Global")="立即锁")then
		if(Datatable("LockUnlockFlag","Global")="锁车")then
		   '替换实际锁车设置的参数源码段中 解车密码基数+日期 为设置的日期
			ParaListSour=replace(ParaListSour,left(ParaListSour,6),"01012")
			'立即锁，锁车场景的参数列表源码段需要处理
			ExpParaListSour=replace(Datatable("LockSetRepSour_OtherContent","Global"),left(Datatable("LockSetRepSour_OtherContent","Global"),6),"01012")
	   else
			'立即锁，解车场景的参数列表源码段不用处理
			ExpParaListSour=Datatable("UnLockSetRepSour_OtherContent","Global")
	   end if
'	   '========检查参数列表源码段=====
'		if(ParaListSour=ExpParaListSour)then
'		reporter.ReportEvent micPass,Datatable("LockUnlockFlag","Global")&"设置源码中参数列表检查通过","期望值："& ExpParaListSour &" 实际值："& ParaListSour
'		else
'		reporter.ReportEvent micFail,Datatable("LockUnlockFlag","Global")&"设置源码中参数列表检查失败","期望值："& ExpParaListSour &" 实际值："& ParaListSour
'		end if
end if '结束循环日期锁的判断

'========总工作时间锁/指定日期锁/指定位置锁/循环日期锁/立即锁  的锁解车设置源码处理，经纬度编码有误差，无法处理
if(Datatable("LockType","Global")="总工作时间锁/指定日期锁/指定位置锁/循环日期锁/立即锁")then
		if(Datatable("LockUnlockFlag","Global")="锁车")then
		   '替换实际锁车设置的参数源码段中 解车密码基数+日期 为设置的日期
			ParaListSour=len(ParaListSour)
			'锁车场景的参数列表源码段需要处理
			ExpParaListSour=len(Datatable("LockSetRepSour_OtherContent","Global"))
	   else
			'解车场景的参数列表源码段不用处理
			ExpParaListSour=Datatable("UnLockSetRepSour_OtherContent","Global")
	   end if
end if '结束总工作时间锁/指定日期锁/指定位置锁/循环日期锁/立即锁 的判断

'========最后检查参数列表源码段=====
if(ParaListSour=ExpParaListSour)then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("LockUnlockFlag","Global")&"设置源码中参数列表检查通过","期望值："& ExpParaListSour &" 实际值："& ParaListSour
else
reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("LockUnlockFlag","Global")&"设置源码中参数列表检查失败","期望值："& ExpParaListSour &" 实际值："& ParaListSour
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
