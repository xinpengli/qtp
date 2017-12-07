On error resume next
'========加载测试数据--调试用
'datatable.ImportSheet "..\..\Sumitomo\TestData\LockReport.xls",1,"Global"
'datatable.ImportSheet "..\..\Sumitomo\TestData\LockReport.xls",2,"Action1"
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
if(Browser("住友").Page("主页_车辆信息").Link("最新提交_查看锁/解车报告").Exist)then
Browser("住友").Page("主页_车辆信息").Link("最新提交_查看锁/解车报告").Click
end if
'========检查是否正常进入“查看锁车设置/回复信息”页
Dim PosiLockSetOrRepMsgPage
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").WebElement("您的位置>>车辆信息>>查看锁/解车报告").Exist)then
	PosiLockUnlockReportPage=Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").WebElement("您的位置>>车辆信息>>查看锁/解车报告").GetROProperty("innertext")
	if(trim(PosiLockUnlockReportPage)=Datatable("PosiLockUnlockReportPage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入最新提交-查看锁/解车报告页成功","期望值："&Datatable("PosiLockUnlockReportPage","Global")&" 实际值："& PosiLockUnlockReportPage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入最新提交-查看锁/解车报告页失败","期望值："&Datatable("PosiLockUnlockReportPage","Global")&" 实际值："& PosiLockUnlockReportPage
	end if
end if
'========输入时间段
Dim currDay '获取当天日期
currDay=Cstr(Year(Date)&"-"&right("0"&Month(Date),2)&"-"&right("0"&Day(Date),2))
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").WebEdit("开始时间").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").WebEdit("开始时间").Object.value=currDay
end if
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").WebEdit("结束时间").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").WebEdit("结束时间").Object.value=currDay
end if
'========查询信息
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").WebButton("查询").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").WebButton("查询").Click
end if
'========检查查询结果列表
datatable.GetSheet("Action1").SetCurrentRow(datatable.GetSheet("Global").GetCurrentRow)  '设置Action1与global行数对应，避免检查串行
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").WebTable("锁解车列表").Exist)then
	Set wt=Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").WebTable("锁解车列表")
	'等待数据加载
	While trim(wt.GetCellData(2,1))<>Datatable("ExpLockUnlockRpt_VclNo","Action1")
		wait 2
		Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").WebButton("查询").Click
	Wend
	'开始检查结果列表，因列表倒序排列，故只检查第一行数据即可
	Dim ColName,ColVal '定义列名列值变量
	For i=1 to wt.ColumnCount(1)
		'循环获取列名及对应列值
		ColName=trim(wt.GetCellData(1,i))
		ColVal=trim(wt.GetCellData(2,i))
		Select Case ColName
			Case "机号":
				if(ColVal=Datatable("ExpLockUnlockRpt_VclNo","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockUnlockRpt_VclNo","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockUnlockRpt_VclNo","Action1")&" 实际值："&ColVal
				end if
			Case "信息发送时间":				
				if(ColVal=Datatable("LockReportSour_InfoGeneTime","Global"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("LockReportSour_InfoGeneTime","Global")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("LockReportSour_InfoGeneTime","Global")&" 实际值："&ColVal
				end if
			Case "经度":
				if(ColVal=Datatable("ExpLockUnlockRpt_Long","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockUnlockRpt_Long","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockUnlockRpt_Long","Action1")&" 实际值："&ColVal
				end if
			Case "纬度":
				if(ColVal=Datatable("ExpLockUnlockRpt_Lati","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockUnlockRpt_Lati","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockUnlockRpt_Lati","Action1")&" 实际值："&ColVal
				end if
			Case "位置":
				if(ColVal=Datatable("ExpLockUnlockRpt_Posi","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockUnlockRpt_Posi","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockUnlockRpt_Posi","Action1")&" 实际值："&ColVal
				end if
			Case "发动机工作时间":
				if(ColVal=Datatable("ExpLockUnlockRpt_RunWorkTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockUnlockRpt_RunWorkTime","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockUnlockRpt_RunWorkTime","Action1")&" 实际值："&ColVal
				end if
			Case "锁车项目":
				if(ColVal=Datatable("ExpLockUnlockRpt_LockItem","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockUnlockRpt_LockItem","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁解车报告结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockUnlockRpt_LockItem","Action1")&" 实际值："&ColVal
				end if
		End Select
	Next
end if   '锁解车报告列表检查完毕
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").Link("车辆信息").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_锁/解车报告").Link("车辆信息").Click
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
