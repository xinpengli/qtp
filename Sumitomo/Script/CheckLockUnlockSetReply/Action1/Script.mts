On error resume next
'加载测试数据--调试用
'datatable.ImportSheet "..\..\Sumitomo\TestData\CheckLockSetReply.xls",1,"Global"
'datatable.ImportSheet "..\..\Sumitomo\TestData\CheckLockSetReply.xls",2,"Action1"
'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'点击车辆信息页-最新提交
if(Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Exist)then
Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Click
end if
'点击车辆信息-最新提交-查看锁车设置/回复信息
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
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebList("设置/回复").Select  "回复"
end if
'========查询信息
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("查询").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("查询").Click
end if
Browser("住友").Page("主页_车辆信息").Sync '执行到下面while语句时偶尔会提示无法识别webtable
'========检查查询结果列表
datatable.GetSheet("Action1").SetCurrentRow(datatable.GetSheet("Global").GetCurrentRow)  '设置Action1与global行数对应，避免检查串行
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebTable("锁车设置/回复信息列表").Exist)then
	Set wt=Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebTable("锁车设置/回复信息列表")
	wait 3 
	if(Datatable("LockUnlockFlag","Global")="锁车")then
		'等待数据加载-锁车回复场景
		While trim(wt.GetCellData(2,1))<>Datatable("LockSetRepSour_InfoGeneTime","Global")
			Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("查询").Click
			wait 2
		Wend
	else
		'等待数据加载-解车回复场景
		While trim(wt.GetCellData(2,1))<>Datatable("UnLockSetRepSour_InfoGeneTime","Global")
			Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("查询").Click
			wait 2
		Wend
	end if
	'开始检查结果列表，因列表倒序排列，故只检查第一行数据即可
	Dim ColName,ColVal '定义列名列值变量
	For i=1 to wt.ColumnCount(1)
		'循环获取列名及对应列值
		ColName=trim(wt.GetCellData(1,i))
		ColVal=trim(wt.GetCellData(2,i))		
		if(Datatable("LockUnlockFlag","Global")="锁车")then  '锁车设置回复检查
				Select Case ColName
					Case "信息发送时间":'信息发送时间ExpLockRep_MsgSendTime 等同 源码中信息生成时间LockSetSour_InfoGeneTime，此处只是对应页面列出
						if(ColVal=Datatable("LockSetRepSour_InfoGeneTime","Global"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("LockSetRepSour_InfoGeneTime","Global")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("LockSetRepSour_InfoGeneTime","Global")&" 实际值："&ColVal
						end if				
					Case "立即锁":
						if(ColVal=Datatable("ExpLockSetRep_ImmeLock","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_ImmeLock","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_ImmeLock","Action1")&" 实际值："&ColVal
						end if
					Case "总工作时间锁":
						if(ColVal=Datatable("ExpLockSetRep_WorkTimeLock","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_WorkTimeLock","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_WorkTimeLock","Action1")&" 实际值："&ColVal
						end if
					Case "工作时间":
						if(ColVal=Datatable("ExpLockSetRep_WorkTimeLock_Hour","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_WorkTimeLock_Hour","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_WorkTimeLock_Hour","Action1")&" 实际值："&ColVal
						end if
					Case "指定日期锁": 
						if(ColVal=Datatable("ExpLockSetRep_AppDateLock","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_AppDateLock","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_AppDateLock","Action1")&" 实际值："&ColVal
						end if
					Case "指定日期":'检查不使用Action1中的ExpLockSetRep_AppDateLock，使用Global中的AppDateLock_Date
						if(ColVal=Datatable("AppDateLock_Date","Global"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("AppDateLock_Date","Global")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("AppDateLock_Date","Global")&" 实际值："&ColVal
						end if
					Case "位置锁":
						if(ColVal=Datatable("ExpLockSetRep_AppPosiLock","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_AppPosiLock","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_AppPosiLock","Action1")&" 实际值："&ColVal
						end if
					Case "经度":
						if(Datatable("CheckLoFlag","Global")="")then
							if(ColVal=Datatable("ExpLockSetRep_AppPosiLock_Long","Action1"))then
							reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_AppPosiLock_Long","Action1")&" 实际值："&ColVal
							else
							reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_AppPosiLock_Long","Action1")&" 实际值："&ColVal
							end if							
							Datatable("CheckLoFlag","Global")="Y"   '标识位，因此字段重复，下次检查根据标识取期望值ExpLockSetRep_Long5
						else
							if(ColVal=Datatable("ExpLockSetRep_Long5","Action1"))then
							reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"5检查","期望值："&Datatable("ExpLockSetRep_Long5","Action1")&" 实际值："&ColVal
							else
							reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"5检查","期望值："&Datatable("ExpLockSetRep_Long5","Action1")&" 实际值："&ColVal
							end if							 
							Datatable("CheckLoFlag","Global")=""  '标识位，因此字段重复，下次检查根据标识取期望值ExpLockSetRep_Long
						end if
					Case "纬度":
						if(Datatable("CheckLaFlag","Global")="")then
							if(ColVal=Datatable("ExpLockSetRep_AppPosiLock_Lati","Action1"))then
							reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_AppPosiLock_Lati","Action1")&" 实际值："&ColVal
							else
							reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_AppPosiLock_Lati","Action1")&" 实际值："&ColVal
							end if							 
							Datatable("CheckLaFlag","Global")="Y"  '标识位，因此字段重复，下次检查根据标识取期望值ExpLockSetRep_lati5
						else
							if(ColVal=Datatable("ExpLockSetRep_lati5","Action1"))then
							reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"5检查","期望值："&Datatable("ExpLockSetRep_lati5","Action1")&" 实际值："&ColVal
							else
							reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"5检查","期望值："&Datatable("ExpLockSetRep_lati5","Action1")&" 实际值："&ColVal
							end if							 
							Datatable("CheckLaFlag","Global")=""  '标识位，因此字段重复，下次检查根据标识取其它期望值ExpLockSetRep_lati
						end if
					Case "半径":
						if(ColVal=Datatable("ExpLockSetRep_AppPosiLock_Radi","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_AppPosiLock_Radi","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_AppPosiLock_Radi","Action1")&" 实际值："&ColVal
						end if
					Case "循环密码锁":
						if(ColVal=Datatable("ExpLockSetRep_CircDateLock","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_CircDateLock","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_CircDateLock","Action1")&" 实际值："&ColVal
						end if
					Case "循环密码锁时间":
						'锁车设置回复页面展示循环密码锁时间内容，不能直接跟Action1中预期数据作对比，需如下处理后再进行检查
						if(Datatable("LockType","Global")="循环日期锁" or Datatable("LockType","Global")="总工作时间锁/指定日期锁/指定位置锁/循环日期锁/立即锁")then
						Datatable("ExpLockSetRep_CircDateLock_Date","Action1")=Datatable("CircDateLock_Y","Global")&"-"&right("0" & Datatable("CircDateLock_M","Global"),2)&"-"& right("0"&Datatable("CircDateLock_D","Global"),2)
						end if
						if(ColVal=Datatable("ExpLockSetRep_CircDateLock_Date","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_CircDateLock_Date","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_CircDateLock_Date","Action1")&" 实际值："&ColVal
						end if
					Case "总工作时间":
						if(ColVal=Datatable("ExpLockSetRep_ToalWorkHour5","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_ToalWorkHour5","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_ToalWorkHour5","Action1")&" 实际值："&ColVal
						end if
					Case "位置":
						if(ColVal=Datatable("ExpLockSetRep_Position5","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_Position5","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpLockSetRep_Position5","Action1")&" 实际值："&ColVal
						end if
				End Select		
		else  '解车设置回复检查
				Select Case ColName
					Case "信息发送时间":'信息发送时间ExpLockRep_MsgSendTime 等同 源码中信息生成时间LockSetRepSour_InfoGeneTime
						if(ColVal=Datatable("UnLockSetRepSour_InfoGeneTime","Global"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("UnLockSetRepSour_InfoGeneTime","Global")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("UnLockSetRepSour_InfoGeneTime","Global")&" 实际值："&ColVal
						end if				
					Case "立即锁":
						if(ColVal=Datatable("ExpUnlockSetRep_ImmeLock","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_ImmeLock","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_ImmeLock","Action1")&" 实际值："&ColVal
						end if
					Case "总工作时间锁":
						if(ColVal=Datatable("ExpUnlockSetRep_WorkTimeLock","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_WorkTimeLock","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_WorkTimeLock","Action1")&" 实际值："&ColVal
						end if
					Case "工作时间":
						if(ColVal=Datatable("ExpUnlockSetRep_WorkTimeLock_Hour","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_WorkTimeLock_Hour","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_WorkTimeLock_Hour","Action1")&" 实际值："&ColVal
						end if
					Case "指定日期锁":
						if(ColVal=Datatable("ExpUnlockSetRep_AppDateLock","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_AppDateLock","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_AppDateLock","Action1")&" 实际值："&ColVal
						end if
					Case "指定日期":
						if(ColVal=Datatable("ExpUnlockSetRep_AppDateLock_Date","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_AppDateLock_Date","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_AppDateLock_Date","Action1")&" 实际值："&ColVal
						end if
					Case "位置锁":
						if(ColVal=Datatable("ExpUnlockSetRep_AppPosiLock","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_AppPosiLock","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_AppPosiLock","Action1")&" 实际值："&ColVal
						end if
					Case "半径":
						if(ColVal=Datatable("ExpUnlockSetRep_AppPosiLock_Radi","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_AppPosiLock_Radi","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_AppPosiLock_Radi","Action1")&" 实际值："&ColVal
						end if
					Case "循环密码锁":
						if(ColVal=Datatable("ExpUnlockSetRep_CircDateLock","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_CircDateLock","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_CircDateLock","Action1")&" 实际值："&ColVal
						end if
					Case "循环密码锁时间":
						if((Datatable("UnlockType","Global")="循环日期锁" or Datatable("UnlockType","Global")="总工作时间锁/指定日期锁/指定位置锁/循环日期锁/立即锁")and Datatable("CircDateUnlock_All","Global")<>"全部")then
						Datatable("ExpUnlockSetRep_CircDateLock_Date","Action1")=Datatable("CircDateLock_Y","Global") &"-"& right("0"&Datatable("CircDateLock_M","Global"),2)
						end if
						if(ColVal=Datatable("ExpUnlockSetRep_CircDateLock_Date","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_CircDateLock_Date","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_CircDateLock_Date","Action1")&" 实际值："&ColVal
						end if
					Case "总工作时间":
						if(ColVal=Datatable("ExpUnlockSetRep_ToalWorkHour5","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_ToalWorkHour5","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_ToalWorkHour5","Action1")&" 实际值："&ColVal
						end if
					Case "经度":
						if(ColVal=Datatable("ExpUnlockSetRep_Long5","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_Long5","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_Long5","Action1")&" 实际值："&ColVal
						end if
					Case "纬度":
						if(ColVal=Datatable("ExpUnlockSetRep_lati5","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_lati5","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_lati5","Action1")&" 实际值："&ColVal
						end if
					Case "位置":
						if(ColVal=Datatable("ExpUnlockSetRep_Position5","Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_Position5","Action1")&" 实际值："&ColVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置回复列表-"&ColName&"检查","期望值："&Datatable("ExpUnlockSetRep_Position5","Action1")&" 实际值："&ColVal
						end if
				End Select
		end if
	Next
	'下载Excel
	Datatable("LockUnlockReplyFlag","Global")="设置回复" '暂时用于锁解车回复信息下载文件命名
	if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("下载").Exist)then
	Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("下载").Click
	end if
	wait 2
	RunAction "Action1 [DownFile]", oneIteration
	Datatable("LockUnlockReplyFlag","Global")="" '置空，否则会影响解车设置的下载文件命名
end if   '锁解车设置/回复列表检查完毕
'========返回车辆信息页
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
