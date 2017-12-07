'========加载测试数据--调试用
'datatable.ImportSheet "..\..\Sumitomo\TestData\CheckLockSet.xls",1,"Global"
'datatable.ImportSheet "..\..\Sumitomo\TestData\CheckLockSet.xls",2,"Action1"
'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'========点击车辆信息页-最新提交
Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Click
'========点击车辆信息-最新提交-查看锁车设置/回复信息
if(Browser("住友").Page("主页_车辆信息").Link("最新提交_查看锁车设置/回复信息").Exist)then
Browser("住友").Page("主页_车辆信息").Link("最新提交_查看锁车设置/回复信息").Click
end if
'========检查是否正常进入“查看锁车设置/回复信息”页
Dim PosiLockSetOrRepMsgPage
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebElement("位置>>车辆信息>>查看锁车设置/回复信息").Exist)then
	PosiLockSetOrRepMsgPage=Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebElement("位置>>车辆信息>>查看锁车设置/回复信息").GetROProperty("innertext")
	if(trim(PosiLockSetOrRepMsgPage)=Datatable("PosiLockSetOrRepMsgPage","Global"))then
	reporter.ReportEvent micPass,"进入最新提交-查看锁车设置/回复信息页成功","期望值："&Datatable("PosiLockSetOrRepMsgPage","Global")&" 实际值："&PosiLockSetOrRepMsgPage
	else
	reporter.ReportEvent micFail,"进入最新提交-查看锁车设置/回复信息页失败","期望值："&Datatable("PosiLockSetOrRepMsgPage","Global")&" 实际值："&PosiLockSetOrRepMsgPage
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
	While trim(wt.GetCellData(2,2))<>Datatable("ExpLockSet_CommAccount","Action1")
		wait 2
		Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("查询").Click
	Wend
	'开始检查结果列表，因列表倒序排列，故只检查第一行数据即可
	Dim ColName,ColVal '定义列名列值变量
	For i=1 to wt.ColumnCount(1)
		'循环获取列名及对应列值
		ColName=trim(wt.GetCellData(1,i))
		ColVal=trim(wt.GetCellData(2,i))
		Select Case ColName
			Case "提交时间":
				'因提交时间为提交时刻服务器记录时间，故不检查				
			Case "提交帐户":				
				if(ColVal=Datatable("ExpLockSet_CommAccount","Action1"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CommAccount","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CommAccount","Action1")&" 实际值："&ColVal
				end if
			Case "立即锁":
				if(ColVal=Datatable("ExpLockSet_ImmeLock","Action1"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_ImmeLock","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_ImmeLock","Action1")&" 实际值："&ColVal
				end if
			Case "总工作时间锁":
				if(ColVal=Datatable("ExpLockSet_WorkTimeLock","Action1"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_WorkTimeLock","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_WorkTimeLock","Action1")&" 实际值："&ColVal
				end if
			Case "指定总工作时间":
				if(ColVal=Datatable("ExpLockSet_WorkTimeLock_Hour","Action1"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_WorkTimeLock_Hour","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_WorkTimeLock_Hour","Action1")&" 实际值："&ColVal
				end if
			Case "指定日期锁":  
				if(ColVal=Datatable("ExpLockSet_AppDateLock","Action1"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppDateLock","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppDateLock","Action1")&" 实际值："&ColVal
				end if
			Case "指定时间":
				'检查时使用设置的Global值AppDateLock_Date，不使用Action1中的ExpLockSet_AppDateLock_Date
				if(ColVal=Datatable("AppDateLock_Date","Global"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("AppDateLock_Date","Global")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("AppDateLock_Date","Global")&" 实际值："&ColVal
				end if
			Case "位置锁":
				if(ColVal=Datatable("ExpLockSet_AppPosiLock","Action1"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock","Action1")&" 实际值："&ColVal
				end if
			Case "指定经度":
				if(ColVal=Datatable("ExpLockSet_AppPosiLock_Long","Action1"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock_Long","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock_Long","Action1")&" 实际值："&ColVal
				end if
			Case "指定纬度":
				if(ColVal=Datatable("ExpLockSet_AppPosiLock_Lati","Action1"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock_Lati","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock_Lati","Action1")&" 实际值："&ColVal
				end if
			Case "半径":
				if(ColVal=Datatable("ExpLockSet_AppPosiLock_Radi","Action1"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock_Radi","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_AppPosiLock_Radi","Action1")&" 实际值："&ColVal
				end if
			Case "循环密码锁":
				if(ColVal=Datatable("ExpLockSet_CircDateLock","Action1"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CircDateLock","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CircDateLock","Action1")&" 实际值："&ColVal
				end if
			Case "循环密码锁时间":
				'锁车设置查询页面展示循环密码锁时间内容，不能直接跟Action1中预期数据作对比，需如下处理后再进行检查
				if(Datatable("LockType","Global")="循环日期锁" or Datatable("LockType","Global")="总工作时间锁/指定日期锁/指定位置锁/循环日期锁/立即锁")then
				Datatable("ExpLockSet_CircDateLock_Date","Action1")=Datatable("CircDateLock_Y","Global")&"-"&right("0" & Datatable("CircDateLock_M","Global"),2)&"-"& right("0"&Datatable("CircDateLock_D","Global"),2)
				end if
				if(ColVal=Datatable("ExpLockSet_CircDateLock_Date","Action1"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CircDateLock_Date","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CircDateLock_Date","Action1")&" 实际值："&ColVal
				end if
			Case "期数":
				Datatable("ExpLockSet_CircDateLock_LockM","Action1")=Datatable("CircDateLock_LockM","Global")
				if(ColVal=Datatable("ExpLockSet_CircDateLock_LockM","Action1"))then
				reporter.ReportEvent micPass,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CircDateLock_LockM","Action1")&" 实际值："&ColVal
				else
				reporter.ReportEvent micFail,"锁车设置结果列表-"&ColName&"检查","期望值："&Datatable("ExpLockSet_CircDateLock_LockM","Action1")&" 实际值："&ColVal
				end if
		End Select
	Next
	'========下载功能
	if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("下载").Exist)then
	Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebButton("下载").Click
	end if
	wait 2
	'执行下载操作
	if(Dialog("文件下载").WinButton("保存(S)").Exist)then
	Dialog("文件下载").WinButton("保存(S)").Click
	end if
	'输入文件名
	if(Dialog("已完成安装-进度").Dialog("另存为").WinEdit("文件名(N)").Exist)then
	Datatable("ExcelAddr","Global")="D:\"&Environment("TestName")&"_"&Environment("ActionName")&"_"&right("0"& Hour(now),2)&right("0"& minute(now),2)&right("0"&second(now),2)&".xls"
	Dialog("已完成安装-进度").Dialog("另存为").WinEdit("文件名(N)").Set  Datatable("ExcelAddr","Global")
	end if
	'保存excel
	if(Dialog("已完成安装-进度").Dialog("另存为").WinButton("保存(S)").Exist)then
	Dialog("已完成安装-进度").Dialog("另存为").WinButton("保存(S)").Click
	end if
	'========检查下载的excel是否存在，并进行检查
	Set fso=createobject("scripting.filesystemobject")
	wait 1
	if(fso.FileExists(Datatable("ExcelAddr","Global")))then
	reporter.ReportEvent micPass,"excel下载成功","excel下载成功"
	else
	reporter.ReportEvent micPass,"excel下载失败","excel下载失败"
	end if
	Set fso=nothing
	' 创建Excel应用程序对象        
	Set oExcel = CreateObject("Excel.Application")              
	' 打开Excel文件        
	oExcel.Workbooks.Open(Datatable("ExcelAddr","Global"))        
	' 获取表格的使用范围列数
	Dim ColCount
	ColCount=oExcel.Worksheets(1).UsedRange.columns.count
	For i=1 to ColCount
		'比较列名
		if(trim(wt.GetCellData(1,i))=oExcel.Worksheets(1).cells(1,i))then
			'比较列值
			if(trim(wt.GetCellData(2,i))=oExcel.Worksheets(1).cells(2,i))then
			reporter.ReportEvent micPass,"Excel列-"& oExcel.Worksheets(1).cells(1,i)&"-检查通过","期望值："&wt.GetCellData(2,i)&" 实际值："& oExcel.Worksheets(1).cells(2,i)
			else
			reporter.ReportEvent micFail,"Excel列-"& oExcel.Worksheets(1).cells(1,i)&"-检查通过","期望值："&wt.GetCellData(2,i)&" 实际值："& oExcel.Worksheets(1).cells(2,i)
			end if
		end if
	Next
	' 关闭工作簿        
	oExcel.WorkBooks.Item(1).Close        
	' 退出Excel        
	oExcel.Quit        
	Set oExcel = Nothing     
end if   '锁解车设置/回复列表检查完毕
'返回车辆信息页
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").Link("车辆信息").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").Link("车辆信息").Click
end if
