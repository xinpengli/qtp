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
'========点击车辆信息-机器档案
if(Browser("住友").Page("主页_车辆信息").WebElement("机器档案").Exist)then
Browser("住友").Page("主页_车辆信息").WebElement("机器档案").Click
end if
Browser("住友").Page("主页_车辆信息").Sync   '等待页面加载
'进入所有权状态页，设置所有权，返回车辆信息页
if(Browser("住友").Page("主页_车辆信息").Frame("机器档案").Link("所有权").Exist)then
	Browser("住友").Page("主页_车辆信息").Frame("机器档案").Link("所有权").Click
	Browser("住友").Page("主页_车辆信息").Sync
	RunAction "Action1 [CheckVclMsgOwnershipStatusPage]", oneIteration
end if
'========重新进入机器档案
if(Browser("住友").Page("主页_车辆信息").WebElement("机器档案").Exist)then
Browser("住友").Page("主页_车辆信息").WebElement("机器档案").Click
end if
Browser("住友").Page("主页_车辆信息").Sync   '等待页面加载
'========检查机器档案页面
'先获取安装日期的值再进行遍历检查,该值取的是初期设定完成的信息生成时间
Datatable("VehicleInfoTreeView_item_49","Action1")="安装日期                        "&Datatable("InfoGeneTime2","Global")
'获取Frame对象
set oFrame=Browser("住友").Page("主页_车辆信息").Frame("机器档案").Object
'循环i值为web元素对应的ID中数值，自增值，目前脚本和web元素对应，之后有变更，对应修改测试数据即可
For i=42 to 65 
	'通过元素ID获取检查项
	set wt= oFrame.getElementById("VehicleInfoTreeView_item_"&i)
	'获取检查项的值，位置固定，且每个检查项均为1行的table
	ActVal=trim(wt.rows(0).cells(4).innertext)
	ExpVal=trim(Datatable("VehicleInfoTreeView_item_"&i,"Action1"))
	'检查实际值和期望值
	if(ActVal=ExpVal)then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案VehicleInfoTreeView_item_"&i&"项检查通过","期望值："& ExpVal &"实际值："&ActVal
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案VehicleInfoTreeView_item_"&i&"项检查失败","期望值："& ExpVal &"实际值："&ActVal
	end if
Next
'检查信息登录链接页，点击框架页信息登录链接进入该页，检查完毕后关闭该页
if(Browser("住友").Page("主页_车辆信息").Frame("机器档案").Link("信息登录").Exist)then
	Browser("住友").Page("主页_车辆信息").Frame("机器档案").Link("信息登录").Click
	Browser("住友").Page("主页_车辆信息").Sync
	RunAction "Action1 [CheckVclMsg_MachFile_MsgLoginPage]", oneIteration
end if
'检查点检实绩链接页,点击框架页点检实绩链接进入该页，检查完毕后返回车辆信息页
if(Browser("住友").Page("主页_车辆信息").Frame("机器档案").Link("点检实绩位置").Exist)then
	Browser("住友").Page("主页_车辆信息").Frame("机器档案").Link("点检实绩位置").Click
	Browser("住友").Page("主页_车辆信息").Sync
	RunAction "Action1 [CheckVclMsgInspectionResultsPage]", oneIteration
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
