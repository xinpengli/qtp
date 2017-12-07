On error resume next
'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'========检查是否正常进入“点检实绩”页
Dim InspectionResultsPage
if(Browser("住友").Page("主页_车辆信息").Frame("点检实绩").WebElement("位置").Exist)then
	InspectionResultsPage=Browser("住友").Page("主页_车辆信息").Frame("点检实绩").WebElement("位置").GetROProperty("innertext")
	if(trim(InspectionResultsPage)=Datatable("InspectionResultsPage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息-点检实绩页成功","期望值："&Datatable("InspectionResultsPage","Global")&" 实际值："& InspectionResultsPage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息-点检实绩页失败","期望值："&Datatable("InspectionResultsPage","Global")&" 实际值："& InspectionResultsPage
	end if
end if
'检查点检实绩表格
if(Browser("住友").Page("主页_车辆信息").Frame("点检实绩").WebTable("点检实绩").Exist)then
	Set wt=Browser("住友").Page("主页_车辆信息").Frame("点检实绩").WebTable("点检实绩")
	wait 2
	For j=1 to wt.ColumnCount(1)
		ColName=trim(wt.GetCellData(1,j))
		ColValue=trim(wt.GetCellData(2,j))
		Select Case ColName
			Case "机号":
				if(ColValue=Datatable("JiHao","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机号检查通过","期望值: "&Datatable("JiHao","Action1")&"实际值: "&ColValue
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机号检查失败","期望值: "&Datatable("JiHao","Action1")&"实际值: "&ColValue
				end if
			Case "发动机工作小时":
				if(ColValue=Datatable("InspeRes_EgnWorkHour","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"发动机工作小时检查通过","期望值: "&Datatable("InspeRes_EgnWorkHour","Action1")&"实际值: "&ColValue
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"发动机工作小时检查失败","期望值: "&Datatable("InspeRes_EgnWorkHour","Action1")&"实际值: "&ColValue
				end if
			Case "点检时间":
				'Datatable("CheckTime","Action1")在插入源码的action中有赋值
				if(ColValue=Datatable("InspeRes_CheckTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"点检时间检查通过","期望值: "&Datatable("InspeRes_CheckTime","Action1")&"实际值: "&ColValue
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"点检时间检查失败","期望值: "&Datatable("InspeRes_CheckTime","Action1")&"实际值: "&ColValue
				end if
			Case "点检地点":
				if(ColValue=Datatable("InspeRes_CheckPoint","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"点检地点检查通过","期望值: "&Datatable("InspeRes_CheckPoint","Action1")&"实际值: "&ColValue
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"点检地点检查失败","期望值: "&Datatable("InspeRes_CheckPoint","Action1")&"实际值: "&ColValue
				end if
			Case "修改":
				if(ColValue=Datatable("InspeRes_Update","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"修改列检查通过","期望值: "&Datatable("InspeRes_Update","Action1")&"实际值: "&ColValue
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"修改列检查失败","期望值: "&Datatable("InspeRes_Update","Action1")&"实际值: "&ColValue
				end if
				'获取修改链接,进入修改页面
				Set UpdLink=wt.ChildItem(2,5,"Link",0)
				wait 1
				UpdLink.click
				'输入点检服务工程师
				if(Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").WebEdit("点检服务工程师").Exist)then
				Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").WebEdit("点检服务工程师").Set "test"
				end if
				'输入点检类型
				if(Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").WebEdit("点检类型").Exist)then
				Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").WebEdit("点检类型").Set "test"
				end if
				'输入点检内容
				if(Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").WebEdit("点检内容").Exist)then
				Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").WebEdit("点检内容").Set "test"
				end if
				'输入备注
				if(Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").WebEdit("备注").Exist)then
				Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").WebEdit("备注").Set "test"
				end if
				'点击修改
				if(Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").WebButton("修 改").Exist)then
				Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").WebButton("修 改").Click
				end if
				'确认
				if(Browser("住友_车辆信息_机器档案链接页").Dialog("来自网页的消息").WinButton("确定").Exist)then
					if(Browser("住友_车辆信息_机器档案链接页").Dialog("来自网页的消息").static("text:=修改车辆点检实绩时间成功！").Exist)then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"点检实绩页修改成功","点检实绩页修改成功"
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"点检实绩页修改失败","点检实绩页修改失败"
					end if
					Browser("住友_车辆信息_机器档案链接页").Dialog("来自网页的消息").WinButton("确定").Click
				end if	
		End Select
	Next
end if
'下载文件 
if(Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").WebButton("下 载").Exist)then
	Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").WebButton("下 载").Click
	Browser("住友_车辆信息_机器档案链接页").Page("车辆点检实绩时间").Sync
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
