On error resume next
'========加载测试数据--调试用
'datatable.ImportSheet "..\..\Sumitomo\TestData\CheckLockSet.xls",1,"Global"
'datatable.ImportSheet "..\..\Sumitomo\TestData\CheckLockSet.xls",2,"Action1"
'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'========点击车辆信息页-最新提交
if(Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Exist)then
Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Click
end if
'========点击车辆信息-最新提交-查看锁车设置/回复信息
if(Browser("住友").Page("主页_车辆信息").Link("最新提交_查看锁车设置/回复信息").Exist)then
Browser("住友").Page("主页_车辆信息").Link("最新提交_查看锁车设置/回复信息").Click
end if
'========检查是否正常进入“查看锁车设置/回复信息”页
Dim PosiLockSetOrRepMsgPage
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebElement("位置>>车辆信息>>查看锁车设置/回复信息").Exist)then
	PosiLockSetOrRepMsgPage=Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebElement("位置>>车辆信息>>查看锁车设置/回复信息").GetROProperty("innertext")
	if(trim(PosiLockSetOrRepMsgPage)=Datatable("PosiLockSetOrRepMsgPage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入最新提交-查看锁车设置/回复信息页成功","期望值："&Datatable("PosiLockSetOrRepMsgPage","Global")&" 实际值："&PosiLockSetOrRepMsgPage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入最新提交-查看锁车设置/回复信息页失败","期望值："&Datatable("PosiLockSetOrRepMsgPage","Global")&" 实际值："&PosiLockSetOrRepMsgPage
	end if
end if
'========输入时间段
Dim currDay '获取当天日期
currDay=Cstr(Year(Date)&"-"&right("0"&Month(Date),2)&"-"&right("0"&Day(Date),2))
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebEdit("开始时间").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebEdit("开始时间").Object.value=currDay
end if
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebEdit("结束时间").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebEdit("结束时间").Object.value=currDay
end if
'========选择设置类型
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebList("设置/回复").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebList("设置/回复").Select  "设置"
end if
'========查询信息
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("查询").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("查询").Click
end if
'========检查查询结果列表
datatable.GetSheet("Action1").SetCurrentRow(datatable.GetSheet("Global").GetCurrentRow)  '设置Action1与global行数对应，避免检查串行
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebTable("锁车设置/回复信息列表").Exist)then
	Set wt=Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebTable("锁车设置/回复信息列表")
	'等待数据加载
	if(Datatable("LockUnlockFlag","Global")="锁车")then
		While trim(wt.GetCellData(2,2))<>Datatable("ExpLockSet_CommAccount","Action1")
			wait 2
			Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("查询").Click
		Wend
	else
		While trim(wt.GetCellData(2,2))<>Datatable("ExpUnlockSet_CommAccount","Action1")
			wait 2
			Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("查询").Click
		Wend
	end if
	'开始检查结果列表，因列表倒序排列，故只检查第一行数据即可
	Dim ColName,ColVal '定义列名列值变量
	For i=1 to wt.ColumnCount(1)
		'循环获取列名及对应列值
		ColName=trim(wt.GetCellData(1,i))
		ColVal=trim(wt.GetCellData(2,i))
		'检查锁车设置查询列表
		if(Datatable("LockUnlockFlag","Global")="锁车")then
			Select Case ColName
				Case "提交时间":
					'因提交时间为提交时刻服务器记录时间，故不检查				
				Case "提交帐户":				
					if(ColVal=Datatable("ExpLockSet_CommAccount","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CommAccount","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CommAccount","Action1")&" 实际值："&ColVal
					end if
				Case "立即锁":
					if(ColVal=Datatable("ExpLockSet_ImmeLock","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_ImmeLock","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_ImmeLock","Action1")&" 实际值："&ColVal
					end if
				Case "总工作时间锁":
					if(ColVal=Datatable("ExpLockSet_WorkTimeLock","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_WorkTimeLock","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_WorkTimeLock","Action1")&" 实际值："&ColVal
					end if
				Case "指定总工作时间":
					if(ColVal=Datatable("ExpLockSet_WorkTimeLock_Hour","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_WorkTimeLock_Hour","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_WorkTimeLock_Hour","Action1")&" 实际值："&ColVal
					end if
				Case "指定日期锁":  
					if(ColVal=Datatable("ExpLockSet_AppDateLock","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppDateLock","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppDateLock","Action1")&" 实际值："&ColVal
					end if
				Case "指定时间":
					'检查时使用设置的Global值AppDateLock_Date，不使用Action1中的ExpLockSet_AppDateLock_Date
					if(ColVal=Datatable("AppDateLock_Date","Global"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("AppDateLock_Date","Global")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("AppDateLock_Date","Global")&" 实际值："&ColVal
					end if
				Case "位置锁":
					if(ColVal=Datatable("ExpLockSet_AppPosiLock","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock","Action1")&" 实际值："&ColVal
					end if
				Case "指定经度":
					if(ColVal=Datatable("ExpLockSet_AppPosiLock_Long","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock_Long","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock_Long","Action1")&" 实际值："&ColVal
					end if
				Case "指定纬度":
					if(ColVal=Datatable("ExpLockSet_AppPosiLock_Lati","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock_Lati","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock_Lati","Action1")&" 实际值："&ColVal
					end if
				Case "半径":
					if(ColVal=Datatable("ExpLockSet_AppPosiLock_Radi","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock_Radi","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock_Radi","Action1")&" 实际值："&ColVal
					end if
				Case "循环密码锁":
					if(ColVal=Datatable("ExpLockSet_CircDateLock","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CircDateLock","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CircDateLock","Action1")&" 实际值："&ColVal
					end if
				Case "循环密码锁时间":
					'锁车设置查询页面展示循环密码锁时间内容，不能直接跟Action1中预期数据作对比，需如下处理后再进行检查
					if(Datatable("LockType","Global")="循环日期锁" or Datatable("LockType","Global")="总工作时间锁/指定日期锁/指定位置锁/循环日期锁/立即锁")then
					Datatable("ExpLockSet_CircDateLock_Date","Action1")=Datatable("CircDateLock_Y","Global")&"-"&right("0" & Datatable("CircDateLock_M","Global"),2)&"-"& right("0"&Datatable("CircDateLock_D","Global"),2)
					end if
					if(ColVal=Datatable("ExpLockSet_CircDateLock_Date","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CircDateLock_Date","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CircDateLock_Date","Action1")&" 实际值："&ColVal
					end if
				Case "期数":
					Datatable("ExpLockSet_CircDateLock_LockM","Action1")=Datatable("CircDateLock_LockM","Global")
					if(ColVal=Datatable("ExpLockSet_CircDateLock_LockM","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CircDateLock_LockM","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CircDateLock_LockM","Action1")&" 实际值："&ColVal
					end if
			End Select
		'检查解车全解车的设置查询列表
		else 
			Select Case ColName
				Case "提交时间":
					'因提交时间为提交时刻服务器记录时间，故不做检查
				Case "提交帐户":				
					if(ColVal=Datatable("ExpUnlockSet_CommAccount","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_CommAccount","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_CommAccount","Action1")&" 实际值："&ColVal
					end if
				Case "立即锁":
					if(ColVal=Datatable("ExpUnlockSet_ImmeLock","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_ImmeLock","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_ImmeLock","Action1")&" 实际值："&ColVal
					end if
				Case "总工作时间锁":
					if(ColVal=Datatable("ExpUnlockSet_WorkTimeLock","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_WorkTimeLock","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_WorkTimeLock","Action1")&" 实际值："&ColVal
					end if
				Case "指定总工作时间":
					if(ColVal=Datatable("ExpUnlockSet_WorkTimeLock_Hour","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_WorkTimeLock_Hour","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_WorkTimeLock_Hour","Action1")&" 实际值："&ColVal
					end if
				Case "指定日期锁":
					if(ColVal=Datatable("ExpUnlockSet_AppDateLock","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_AppDateLock","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_AppDateLock","Action1")&" 实际值："&ColVal
					end if
				Case "指定时间":
					if(ColVal=Datatable("ExpUnlockSet_AppDateLock_Date","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_AppDateLock_Date","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_AppDateLock_Date","Action1")&" 实际值："&ColVal
					end if
				Case "位置锁":
					if(ColVal=Datatable("ExpUnlockSet_AppPosiLock","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_AppPosiLock","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_AppPosiLock","Action1")&" 实际值："&ColVal
					end if
				Case "指定经度":
					if(ColVal=Datatable("ExpUnlockSet_AppPosiLock_Long","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_AppPosiLock_Long","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_AppPosiLock_Long","Action1")&" 实际值："&ColVal
					end if
				Case "指定纬度":
					if(ColVal=Datatable("ExpUnlockSet_AppPosiLock_Lati","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_AppPosiLock_Lati","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_AppPosiLock_Lati","Action1")&" 实际值："&ColVal
					end if
				Case "半径":
					if(ColVal=Datatable("ExpUnlockSet_AppPosiLock_Radi","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_AppPosiLock_Radi","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_AppPosiLock_Radi","Action1")&" 实际值："&ColVal
					end if
				Case "循环密码锁":
					if(ColVal=Datatable("ExpUnlockSet_CircDateLock","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_CircDateLock","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_CircDateLock","Action1")&" 实际值："&ColVal
					end if
				Case "循环密码锁时间":
					'解车设置查询页面展示循环密码锁时间年-月，Action1中预期数据需如下处理后再进行检查
					if(Datatable("UnlockType","Global")="循环日期锁")then
					Datatable("ExpUnlockSet_CircDateLock_Date","Action1")=Datatable("CircDateLock_Y","Global") &"-"& right("0"&Datatable("CircDateLock_M","Global"),2)
					end if
					if(ColVal=Datatable("ExpUnlockSet_CircDateLock_Date","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_CircDateLock_Date","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_CircDateLock_Date","Action1")&" 实际值："&ColVal
					end if
				Case "期数":
					'解车页面无此字段设置，故期望值设置为空
					if(ColVal=Datatable("ExpUnlockSet_CircDateLock_LockM","Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_CircDateLock_LockM","Action1")&" 实际值："&ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSet_CircDateLock_LockM","Action1")&" 实际值："&ColVal
					end if
			End Select
		end if
	Next
	'下载功能
	if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("下载").Exist)then
	Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("下载").Click
	end if
	wait 2
	RunAction "Action1 [DownFile]", oneIteration
end if   '锁解车设置/回复列表检查完毕
'返回车辆信息页
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").Link("车辆信息").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").Link("车辆信息").Click
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
