On error resume next
'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'当插入新源码后需要刷新当前页面
if(Parameter("j")>0)then
    RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
end if
'检查发动机工作小时
Dim EgnWorkHour
if(Browser("住友").Page("主页_车辆信息").Frame("工作情况").WebElement("发动机工作小时").Exist)then
	EgnWorkHour=Browser("住友").Page("主页_车辆信息").Frame("工作情况").WebElement("发动机工作小时").GetROProperty("innertext")
	if(EgnWorkHour=Datatable("ExpEgnWorkHour"&Parameter("j"),"Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-工作情况-发动机工作小时检查通过","期望值："&Datatable("ExpEgnWorkHour"&Parameter("j"),"Global")&" 实际值："& EgnWorkHour
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-工作情况-发动机工作小时检查失败","期望值："&Datatable("ExpEgnWorkHour"&Parameter("j"),"Global")&" 实际值："& EgnWorkHour
	end if
end if
'检查累计耗油量
Dim SumOil
if(Browser("住友").Page("主页_车辆信息").Frame("工作情况").WebElement("累计耗油量").Exist)then
	SumOil=Browser("住友").Page("主页_车辆信息").Frame("工作情况").WebElement("累计耗油量").GetROProperty("innertext")
	if(SumOil=Datatable("ExpSumOil"&Parameter("j"),"Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-工作情况-累计耗油量检查通过","期望值："&Datatable("ExpSumOil"&Parameter("j"),"Global")&" 实际值："& SumOil
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-工作情况-累计耗油量检查失败","期望值："&Datatable("ExpSumOil"&Parameter("j"),"Global")&" 实际值："& SumOil
	end if
end if
'检查燃油位
Dim OilLevel
if(Browser("住友").Page("主页_车辆信息").Frame("工作情况").WebElement("燃油位").Exist)then
	OilLevel=Browser("住友").Page("主页_车辆信息").Frame("工作情况").Image("燃油位格数").Object.title
	wait 1
'	'通过如下事件实现飘浮文字检查比较好，但目前飘浮的对象抓不下来
'	Browser("住友").Page("主页_车辆信息").Frame("工作情况").Image("燃油位格数").FireEvent "OnMouse"
	if(OilLevel=Datatable("ExpOilLevel"&Parameter("j"),"Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-工作情况-燃油位检查通过","期望值："&Datatable("ExpOilLevel"&Parameter("j"),"Global")&" 实际值："& OilLevel
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-工作情况-燃油位检查失败","期望值："&Datatable("ExpOilLevel"&Parameter("j"),"Global")&" 实际值："& OilLevel
	end if
end if
'检查电瓶被拆报警
Dim Alarm
if(Browser("住友").Page("主页_车辆信息").Frame("工作情况").WebElement("电瓶被拆报警").Exist)then
	Alarm=Browser("住友").Page("主页_车辆信息").Frame("工作情况").WebElement("电瓶被拆报警").GetROProperty("innertext")
	if(Parameter("j")>0)then
	Datatable("ExpAlarm"&Parameter("j"),"Global")="电瓶被拆报警       "&Datatable("InfoGeneTime2","Global")&"报警中"
	end if
	if(Alarm=Datatable("ExpAlarm"&Parameter("j"),"Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-工作情况-电瓶被拆报警检查通过","期望值："&Datatable("ExpAlarm"&Parameter("j"),"Global")&" 实际值："& Alarm
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-工作情况-电瓶被拆报警检查失败","期望值："&Datatable("ExpAlarm"&Parameter("j"),"Global")&" 实际值："& Alarm
	end if
end if
'检查机械电瓶电压
Dim MechBattVolt
if(Browser("住友").Page("主页_车辆信息").Frame("工作情况").WebElement("机械电瓶电压").Exist)then
	MechBattVolt=Browser("住友").Page("主页_车辆信息").Frame("工作情况").WebElement("机械电瓶电压").GetROProperty("innertext")
	if(MechBattVolt=Datatable("ExpMechBattVolt"&Parameter("j"),"Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-工作情况-机械电瓶电压检查通过","期望值："&Datatable("ExpMechBattVolt"&Parameter("j"),"Global")&" 实际值："& MechBattVolt
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-工作情况-机械电瓶电压检查失败","期望值："&Datatable("ExpMechBattVolt"&Parameter("j"),"Global")&" 实际值："& MechBattVolt
	end if
end if
'检查发电机电压
Dim GeneVolt
if(Browser("住友").Page("主页_车辆信息").Frame("工作情况").WebElement("发电机电压").Exist)then
	GeneVolt=Browser("住友").Page("主页_车辆信息").Frame("工作情况").WebElement("发电机电压").GetROProperty("innertext")
	if(GeneVolt=Datatable("ExpGeneVolt"&Parameter("j"),"Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-工作情况-机械电瓶电压检查通过","期望值："&Datatable("ExpGeneVolt"&Parameter("j"),"Global")&" 实际值："& GeneVolt
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"车辆信息-工作情况-机械电瓶电压检查失败","期望值："&Datatable("ExpGeneVolt"&Parameter("j"),"Global")&" 实际值："& GeneVolt
	end if
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
