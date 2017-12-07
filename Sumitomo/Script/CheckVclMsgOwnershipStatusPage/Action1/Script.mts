On error resume next
'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'========检查是否正常进入“所有权设定”页
Dim InspectionResultsPage
if(Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebElement("位置").Exist)then
	InspectionResultsPage=Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebElement("位置").GetROProperty("innertext")
	if(trim(InspectionResultsPage)=Datatable("OwnerShipStatusPage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息-所有权设定页成功","期望值："&Datatable("OwnerShipStatusPage","Global")&" 实际值："& OwnerShipStatusPage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息-所有权设定页失败","期望值："&Datatable("OwnerShipStatusPage","Global")&" 实际值："& OwnerShipStatusPage
	end if
end if
'进入所有权设定页,设定"有所有权"
if(Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebButton("所有权设定").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebButton("所有权设定").Click
end if
if(Browser("住友_车辆信息_机器档案链接页").Page("所有权设定").WebList("设定状态").Exist)then
Browser("住友_车辆信息_机器档案链接页").Page("所有权设定").WebList("设定状态").Select "有所有权"  '之后可将此参数化,暂时写成固定值
end if
if(Browser("住友_车辆信息_机器档案链接页").Page("所有权设定").WebButton("设定变更").Exist)then
Browser("住友_车辆信息_机器档案链接页").Page("所有权设定").WebButton("设定变更").Click
end if
if(Browser("住友_车辆信息_机器档案链接页").Dialog("来自网页的消息").WinButton("确定").Exist)then
	if(Browser("住友_车辆信息_机器档案链接页").Dialog("来自网页的消息").static("text:=所有权设定成功！").Exist)then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"所有权设定页设定成功","所有权设定页设定成功"
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"所有权设定页设定失败","所有权设定页设定失败"
	end if
	Browser("住友_车辆信息_机器档案链接页").Dialog("来自网页的消息").WinButton("确定").Click
end if
Browser("住友").Page("主页_车辆信息").Sync
'所有权设定回复
RunAction "Action1 [InsertSetTerminalPara2ReplySource]", oneIteration
'查询所有权设定表格
if(Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebEdit("开始日期").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebEdit("开始日期").Set Cstr(date)
end if
if(Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebEdit("结束日期").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebEdit("结束日期").Set Cstr(date)
end if
if(Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebButton("查询").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebButton("查询").Click
end if
Browser("住友").Page("主页_车辆信息").Sync
'检查所有权设定表格
if(Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebTable("所有权设定").Exist)then
	Set wt=Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebTable("所有权设定")
	wait 2
	For j=1 to wt.ColumnCount(1)
		ColName=trim(wt.GetCellData(1,j))
		ColValue=trim(wt.GetCellData(2,j))
		Select Case ColName
			Case "机号":
				if(ColValue=Datatable("JiHao","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"所有权设定信息-机号检查通过","期望值: "&Datatable("JiHao","Action1")&"实际值: "&ColValue
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"所有权设定信息-机号检查失败","期望值: "&Datatable("JiHao","Action1")&"实际值: "&ColValue
				end if
			Case "设定时间":
				'设定时间为提交时间,不作检查
			Case "设定状态":
				'有所有权或没有所有权
				if(ColValue=Datatable("OwnerShip_SetStatus","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"所有权设定信息-设定状态检查通过","期望值: "&Datatable("OwnerShip_SetStatus","Action1")&"实际值: "& ColValue
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"所有权设定信息-设定状态检查失败","期望值: "&Datatable("OwnerShip_SetStatus","Action1")&"实际值: "& ColValue
				end if
			Case "设定人员":
				'设定人员为登录帐户
				Datatable("OwnerShip_SetPerson","Action1")=Datatable("Account","Global")
				if(ColValue=Datatable("OwnerShip_SetPerson","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"所有权设定信息-设定人员检查通过","期望值: "&Datatable("OwnerShip_SetPerson","Action1")&"实际值: "& ColValue
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"所有权设定信息-设定人员检查失败","期望值: "&Datatable("OwnerShip_SetPerson","Action1")&"实际值: "& ColValue
				end if
			Case "发送状态":
				if(ColValue=Datatable("OwnerShip_SendStatus","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"所有权设定信息-发送状态检查通过","期望值: "&Datatable("OwnerShip_SendStatus","Action1")&"实际值: "& ColValue
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"所有权设定信息-发送状态检查失败","期望值: "&Datatable("OwnerShip_SendStatus","Action1")&"实际值: "& ColValue
				end if
			Case "回复状态":
				if(ColValue=Datatable("OwnerShip_ReplyStatus","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"所有权设定信息-回复状态检查通过","期望值: "&Datatable("OwnerShip_ReplyStatus","Action1")&"实际值: "& ColValue
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"所有权设定信息-回复状态检查失败","期望值: "&Datatable("OwnerShip_ReplyStatus","Action1")&"实际值: "& ColValue
				end if
		End Select
	Next
end if
'下载文件 
if(Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebButton("下载").Exist)then
	Browser("住友").Page("主页_车辆信息").Frame("所有权设定信息").WebButton("下载").Click
	Browser("住友").Page("主页_车辆信息").Sync
	RunAction "Action1 [DownFile]", oneIteration
end if
'返回车辆信息页
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").Link("车辆信息").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").Link("车辆信息").Click
end if
Browser("住友").Page("主页_车辆信息").Sync
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If
