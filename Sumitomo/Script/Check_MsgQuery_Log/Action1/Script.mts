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
Dim currDay '获取昨日日期，日志只能看昨天的
currDay=Cstr(Year(Date-1)&"-"&right("0"&Month(Date-1),2)&"-"&right("0"&Day(Date-1),2))
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
			Case "信息发送时间":
				Datatable("MsgSendTime","Action1")=Datatable("InfoGeneTime","Global")
				if(ColVal=Datatable("MsgSendTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("MsgSendTime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"间-检查通过","期望值: "& Datatable("MsgSendTime","Action1") &"实际值: "& ColVal
				end if
			Case "日志日期":
				Datatable("LogDate","Action1")=year(date-1) &"-"& right("0"&month(date-1),2) & "-" & right("0"&day(date-1),2)
				if(ColVal=Datatable("LogDate","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LogDate","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LogDate","Action1") &"实际值: "& ColVal
				end if
			Case "经度":
				if(ColVal=Datatable("Long","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("Long","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("Long","Action1") &"实际值: "& ColVal
				end if
			Case "纬度":
				if(ColVal=Datatable("Lati","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("Lati","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("Lati","Action1") &"实际值: "& ColVal
				end if
			Case "位置":
				if(ColVal=Datatable("Posi","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("Posi","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("Posi","Action1") &"实际值: "& ColVal
				end if
			Case "累计耗油量":
				if(ColVal=Datatable("tFuelCons","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tFuelCons","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tFuelCons","Action1") &"实际值: "& ColVal
				end if
			Case "机器操作时间":
				if(ColVal=Datatable("tOperationTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tOperationTime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tOperationTime","Action1") &"实际值: "& ColVal
				end if
			Case "上物操作时间":
				if(ColVal=Datatable("tSlingTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tSlingTime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tSlingTime","Action1") &"实际值: "& ColVal
				end if
			Case "回转操作时间":
				if(ColVal=Datatable("tTurningTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tTurningTime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tTurningTime","Action1") &"实际值: "& ColVal
				end if
			Case "行走操作时间":
				if(ColVal=Datatable("tWalkTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tWalkTime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tWalkTime","Action1") &"实际值: "& ColVal
				end if
			Case "SP模式操作时间":
				if(ColVal=Datatable("tSPTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tSPTime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tSPTime","Action1") &"实际值: "& ColVal
				end if
			Case "H模式操作时间":
				if(ColVal=Datatable("tHTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tHTime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tHTime","Action1") &"实际值: "& ColVal
				end if
			Case "A模式操作时间":
				if(ColVal=Datatable("tATime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tATime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tATime","Action1") &"实际值: "& ColVal
				end if
			Case "B模式操作时间":
				if(ColVal=Datatable("tBTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tBTime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tBTime","Action1") &"实际值: "& ColVal
				end if
			Case "C模式操作时间":
				if(ColVal=Datatable("tCTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tCTime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tCTime","Action1") &"实际值: "& ColVal
				end if
			Case "冷却水温最高温度":
				if(ColVal=Datatable("maxWaterTmprt","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxWaterTmprt","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxWaterTmprt","Action1") &"实际值: "& ColVal
				end if
			Case "燃料温最高温度":
				if(ColVal=Datatable("maxFuelTmprt","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxFuelTmprt","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxFuelTmprt","Action1") &"实际值: "& ColVal
				end if
			Case "吸气最高温度":
				if(ColVal=Datatable("maxInhaleTmprt","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxInhaleTmprt","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxInhaleTmprt","Action1") &"实际值: "& ColVal
				end if
			Case "加压最高温度":
				if(ColVal=Datatable("maxPressurizeTmprt","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxPressurizeTmprt","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxPressurizeTmprt","Action1") &"实际值: "& ColVal
				end if
			Case "液压油温最高温度":
				if(ColVal=Datatable("maxFluidTmprt","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxFluidTmprt","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxFluidTmprt","Action1") &"实际值: "& ColVal
				end if
			Case "冷却水温最低温度":
				if(ColVal=Datatable("minWaterTmprt","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("minWaterTmprt","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("minWaterTmprt","Action1") &"实际值: "& ColVal
				end if
			Case "大气压最低压力":
				if(ColVal=Datatable("minBarometric","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("minBarometric","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("minBarometric","Action1") &"实际值: "& ColVal
				end if
			Case "T＜77℃(水温分布)":
				if(ColVal=Datatable("WaterTmprtSgmt1","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt1","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt1","Action1") &"实际值: "& ColVal
				end if
			Case "77℃≦T＜82℃":
				if(ColVal=Datatable("WaterTmprtSgmt2","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt2","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt2","Action1") &"实际值: "& ColVal
				end if
			Case "82℃≦T＜97℃":
				if(ColVal=Datatable("WaterTmprtSgmt3","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt3","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt3","Action1") &"实际值: "& ColVal
				end if
			Case "97℃≦T＜100℃":
				if(ColVal=Datatable("WaterTmprtSgmt4","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt4","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt4","Action1") &"实际值: "& ColVal
				end if
			Case "100℃≦T＜103℃":
				if(ColVal=Datatable("WaterTmprtSgmt5","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt5","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt5","Action1") &"实际值: "& ColVal
				end if
			Case "103℃≦T＜105℃":
				if(ColVal=Datatable("WaterTmprtSgmt6","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt6","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt6","Action1") &"实际值: "& ColVal
				end if
			Case "105℃≦T":
				if(ColVal=Datatable("WaterTmprtSgmt7","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt7","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtSgmt7","Action1") &"实际值: "& ColVal
				end if
			Case "R＜30％(负荷率分布)":
				if(ColVal=Datatable("LoadRateSgmt1","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt1","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt1","Action1") &"实际值: "& ColVal
				end if
			Case "30％≦R＜40％":
				if(ColVal=Datatable("LoadRateSgmt2","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt2","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt2","Action1") &"实际值: "& ColVal
				end if
			Case "40％≦R＜50％":
				if(ColVal=Datatable("LoadRateSgmt3","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt3","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt3","Action1") &"实际值: "& ColVal
				end if
			Case "50％≦R＜60％":
				if(ColVal=Datatable("LoadRateSgmt4","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt4","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt4","Action1") &"实际值: "& ColVal
				end if
			Case "60％≦R＜70％":
				if(ColVal=Datatable("LoadRateSgmt5","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt5","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt5","Action1") &"实际值: "& ColVal
				end if
			Case "70％≦R＜80％":
				if(ColVal=Datatable("LoadRateSgmt6","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt6","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt6","Action1") &"实际值: "& ColVal
				end if
			Case "80％≦R":
				if(ColVal=Datatable("LoadRateSgmt7","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt7","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("LoadRateSgmt7","Action1") &"实际值: "& ColVal
				end if
			Case "油压最高值":
				if(ColVal=Datatable("maxOilPress","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxOilPress","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxOilPress","Action1") &"实际值: "& ColVal
				end if
			Case "增压后进气压力最高值":
				if(ColVal=Datatable("maxIntakePress","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxIntakePress","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("maxIntakePress","Action1") &"实际值: "& ColVal
				end if
			Case "燃料温最低温度":
				if(ColVal=Datatable("minFuelTmprt","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("minFuelTmprt","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("minFuelTmprt","Action1") &"实际值: "& ColVal
				end if
			Case "油压最低值":
				if(ColVal=Datatable("minOilPress","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("minOilPress","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("minOilPress","Action1") &"实际值: "& ColVal
				end if
			Case "发动机工作小时":
				if(ColVal=Datatable("tWorkTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tWorkTime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("tWorkTime","Action1") &"实际值: "& ColVal
				end if
			Case "当天发动机工作小时":
				if(ColVal=Datatable("iWorkTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("iWorkTime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("iWorkTime","Action1") &"实际值: "& ColVal
				end if
			Case "当天耗油量":
				if(ColVal=Datatable("iFuelCons","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("iFuelCons","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("iFuelCons","Action1") &"实际值: "& ColVal
				end if
			Case "吸气最低温度":
				if(ColVal=Datatable("SniffMinTmprt","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("SniffMinTmprt","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("SniffMinTmprt","Action1") &"实际值: "& ColVal
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
