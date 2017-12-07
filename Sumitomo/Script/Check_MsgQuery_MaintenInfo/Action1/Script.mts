On error resume next
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
if(Browser("住友").Page("主页_车辆信息").Link("信息查询_").Exist)then
Browser("住友").Page("主页_车辆信息").Link("信息查询_").Click
end if
'========检查是否正常进入“查看锁车设置/回复信息”页
Dim PosiVehiMsgPage
if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebElement("您的位置>>车辆信息>>信息查询").Exist)then
	PosiMsgQueryPage=Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebElement("您的位置>>车辆信息>>信息查询").GetROProperty("innertext")
	if(trim(PosiMsgQueryPage)=Datatable("PosiMsgQueryPage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息-信息查询页成功","期望值："&Datatable("PosiMsgQueryPage","Global")&" 实际值："& PosiMsgQueryPage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息-信息查询页失败","期望值："&Datatable("PosiMsgQueryPage","Global")&" 实际值："& PosiMsgQueryPage
	end if
end if
'========输入时间段
Dim currDay '获取当天日期
currDay=Cstr(Year(Date)&"-"&right("0"&Month(Date),2)&"-"&right("0"&Day(Date),2))
if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebEdit("开始时间").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebEdit("开始时间").Object.value=currDay
end if
if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebEdit("结束时间").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebEdit("结束时间").Object.value=currDay
end if
if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebList("信息类型").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebList("信息类型").Select  Datatable("InputMsgType","Global")
end if
'========查询信息
if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebButton("查询").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebButton("查询").Click
end if
Browser("住友").Page("主页_车辆信息").Sync '等待加载
'设置预期值存放的action与当前执行的Global行数对应
Datatable.GetSheet("Action1").SetCurrentRow(Datatable.GetSheet("Global").GetCurrentRow)
'判断webtable是否存在
if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebTable("信息查询列表").Exist)then
	Set wt=Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebTable("信息查询列表")
	wait 2
	Dim ColVal,ExpColVal
	For i=1 to wt.ColumnCount(1)
		ColName=trim(wt.GetCellData(1,i)) '获取页面列名
		ColVal=trim(wt.GetCellData(2,i))       '获取页面列值
		Select Case ColName
			Case "机号":
				if(ColVal=Datatable("JiHao","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("JiHao","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("JiHao","Action1") &"实际值: "& ColVal
				end if
			Case "保养所属日期":
				Datatable("BaoYangDate","Action1")=Datatable("DayMonYear","Global")
				if(ColVal=Datatable("BaoYangDate","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("BaoYangDate","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("BaoYangDate","Action1") &"实际值: "& ColVal
				end if
			Case "柴油粗滤":
				if(ColVal=Datatable("DieselRoughFilter","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("DieselRoughFilter","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("DieselRoughFilter","Action1") &"实际值: "& ColVal
				end if
			Case "柴油细滤":
				if(ColVal=Datatable("DieselFineFilter","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("DieselFineFilter","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("DieselFineFilter","Action1") &"实际值: "& ColVal
				end if
			Case "机油":
				if(ColVal=Datatable("Oil","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("Oil","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("Oil","Action1") &"实际值: "& ColVal
				end if
			Case "机油滤芯":
				if(ColVal=Datatable("OilFilter","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("OilFilter","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("OilFilter","Action1") &"实际值: "& ColVal
				end if
			Case "空滤":
				if(ColVal=Datatable("AirFilter","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("AirFilter","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("AirFilter","Action1") &"实际值: "& ColVal
				end if
			Case "透气阀滤芯":
				if(ColVal=Datatable("VentFilter","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("VentFilter","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("VentFilter","Action1") &"实际值: "& ColVal
				end if
			Case "先导油滤芯":
				if(ColVal=Datatable("PilotOilFilter","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("PilotOilFilter","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("PilotOilFilter","Action1") &"实际值: "& ColVal
				end if
			Case "回油滤芯":
				if(ColVal=Datatable("BackOilFilter","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("BackOilFilter","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("BackOilFilter","Action1") &"实际值: "& ColVal
				end if
			Case "吸油滤芯":
				if(ColVal=Datatable("InhaleOilFilter","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("InhaleOilFilter","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("InhaleOilFilter","Action1") &"实际值: "& ColVal
				end if
			Case "液压油":
				if(ColVal=Datatable("HydraulicOil","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("HydraulicOil","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("HydraulicOil","Action1") &"实际值: "& ColVal
				end if
			Case "回转减速机齿轮油":
				if(ColVal=Datatable("RotaryGearOil","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("RotaryGearOil","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("RotaryGearOil","Action1") &"实际值: "& ColVal
				end if
			Case "行走减速机齿轮油":
				if(ColVal=Datatable("WalkGearOil","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WalkGearOil","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WalkGearOil","Action1") &"实际值: "& ColVal
				end if
			Case "空调滤芯":
				if(ColVal=Datatable("AirFilter2","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("AirFilter2","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("AirFilter2","Action1") &"实际值: "& ColVal
				end if
		End Select
	Next
	'下载文件 
	if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebButton("下载").Exist)then
	Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebButton("下载").Click
	end if
	Browser("住友").Page("主页_车辆信息").Sync
	RunAction "Action1 [DownFile]", oneIteration
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
