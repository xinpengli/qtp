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
if(Browser("住友").Page("主页_车辆信息").Link("信息年表").Exist)then
Browser("住友").Page("主页_车辆信息").Link("信息年表").Click
end if
'========检查是否正常进入“车辆信息/信息年表”页
if(Browser("住友").Page("主页_车辆信息").Frame("信息年表").WebElement("位置").Exist)then
	Dim PosiInfoChronologyPage
	PosiInfoChronologyPage=Browser("住友").Page("主页_车辆信息").Frame("信息年表").WebElement("位置").GetROProperty("innertext")
	if(trim(PosiInfoChronologyPage)=Datatable("PosiInfoChronologyPage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息-信息年表页成功","期望值："&Datatable("PosiInfoChronologyPage","Global")&" 实际值："& PosiInfoChronologyPage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息-信息年表页失败","期望值："&Datatable("PosiInfoChronologyPage","Global")&" 实际值："& PosiInfoChronologyPage
	end if
end if
'========检查设备号信息
if(Browser("住友").Page("主页_车辆信息").Frame("信息年表").WebElement("x设备工作信息年表，请选择查询项").Exist)then
	Dim Vcl_No_Msg
	Vcl_No_Msg=Browser("住友").Page("主页_车辆信息").Frame("信息年表").WebElement("x设备工作信息年表，请选择查询项").GetROProperty("innertext")
	if(trim(Vcl_No_Msg)="【"&Datatable("Vcl_No","Global")&"】工作信息年表，请选择查询项:")then
		reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-信息年表页-设备信息检查成功","期望值："&"【"&Datatable("Vcl_No","Global")&"】工作信息年表，请选择查询项:"&" 实际值："& Vcl_No_Msg
	else
		reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-信息年表页-设备信息检查失败","期望值："&"【"&Datatable("Vcl_No","Global")&"】工作信息年表，请选择查询项:"&" 实际值："& Vcl_No_Msg
	end if
end if
'========选择信息年表查询项
'料余量通知信息-item value-6; 维修通知信息-item value-5; 防盗动作通知信息-item value-4; 故障信息-item value-3; 信号强度信息-item value-8; Dim i'定义入参-源码插入的条数
'根据前边action插入的源码选择对应的信息年表查询项
if(Browser("住友").Page("主页_车辆信息").Frame("信息年表").WebRadioGroup("查询项按钮组").Exist)then
	Browser("住友").Page("主页_车辆信息").Frame("信息年表").WebRadioGroup("查询项按钮组").Select  Datatable("InfoChroType","Global")
end if
'========选择信息年表查看的年份
if(Browser("住友").Page("主页_车辆信息").Frame("信息年表").WebList("年份").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("信息年表").WebList("年份").Select  Cstr(year(date))'获取当前年,如需要实现任意年可在Datatable中设置参数存放
end if
'========点击信息年表查询按钮
if(Browser("住友").Page("主页_车辆信息").Frame("信息年表").WebButton("选好了,立即查看").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("信息年表").WebButton("选好了,立即查看").Click
end if
Browser("住友").Page("主页_车辆信息").Sync
'========检查查询列表
Datatable.GetSheet("Action1").SetCurrentRow(Datatable.GetSheet("Global").GetCurrentRow) '设置Action1中执行行和Global中执行行一致,避免预期值取错
Set InfoChro=Browser("住友").Page("主页_车辆信息").Frame("信息年表").WebTable("年表_数值")
wait 2
Dim mm,dd
mm=month(Datatable("InfoGeneTime","Global"))
dd=day(Datatable("InfoGeneTime","Global"))
Dim ActVal,ExpVal
'当信息年表检查的是日志中信息强度时,获取的天为信息生成时间的前一天,否则即与信息生成时间的日期一致
'获取实际值
if(mid(Datatable("MsgHeadPart","Global"),3,2)<>"20")then
	ActVal=trim(InfoChro.GetCellData(mm,dd))
	else
	dd=day(CDate(Datatable("InfoGeneTime","Global"))-1)
	ActVal=trim(InfoChro.GetCellData(mm,dd))
end if
'获取期望值
ExpVal=Datatable("ExpInfoChroVal"&Parameter("i"),"Action1")
if(ActVal=ExpVal)then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"信息年表-"&mid(Datatable("MsgHeadPart","Global"),3,2)&"指令"&Datatable("StatTaskExeDate","Global")&"统计后检查通过","期望值: "& ExpVal &" 实际值:"& ActVal
else
reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"信息年表-"&mid(Datatable("MsgHeadPart","Global"),3,2)&"指令"&Datatable("StatTaskExeDate","Global")&"统计后检查失败","期望值: "& ExpVal &" 实际值:"& ActVal
end if
'========返回车辆信息页
if(Browser("住友").Page("主页_车辆信息").Frame("信息年表").Link("车辆信息").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("信息年表").Link("车辆信息").Click
Browser("住友").Page("主页_车辆信息").Sync
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
