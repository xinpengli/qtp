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
			Case "信息生成时间":
				Datatable("MsgGeneTime","Action1")=Datatable("InfoGeneTime","Global")
				if(ColVal=Datatable("MsgGeneTime","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("MsgGeneTime","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("MsgGeneTime","Action1") &"实际值: "& ColVal
				end if
			Case "发动机扭矩平均值":
				if(ColVal=Datatable("EgnTorqueAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("EgnTorqueAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("EgnTorqueAvg","Action1") &"实际值: "& ColVal
				end if
			Case "发动机扭矩标准偏差平方":
				if(ColVal=Datatable("EgnTorqueStdDev","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("EgnTorqueStdDev","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("EgnTorqueStdDev","Action1") &"实际值: "& ColVal
				end if
			Case "发动机实际转数平均值":
				if(ColVal=Datatable("EgnSpeedAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("EgnSpeedAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("EgnSpeedAvg","Action1") &"实际值: "& ColVal
				end if
			Case "发动机实际转数标准偏差平方":
				if(ColVal=Datatable("EgnSpeedStdDev","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("EgnSpeedStdDev","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("EgnSpeedStdDev","Action1") &"实际值: "& ColVal
				end if
			Case "冷却水温平均值":
				if(ColVal=Datatable("WaterTmprtAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtAvg","Action1") &"实际值: "& ColVal
				end if
			Case "冷却水温偏差":
				if(ColVal=Datatable("WaterTmprtDev","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtDev","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("WaterTmprtDev","Action1") &"实际值: "& ColVal
				end if
			Case "燃料温度平均值":
				if(ColVal=Datatable("FuelTmprtAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("FuelTmprtAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("FuelTmprtAvg","Action1") &"实际值: "& ColVal
				end if
			Case "吸气温度平均值":
				if(ColVal=Datatable("InhaleTmprtAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("InhaleTmprtAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("InhaleTmprtAvg","Action1") &"实际值: "& ColVal
				end if
			Case "大气压平均值":
				if(ColVal=Datatable("BarometricAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("BarometricAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("BarometricAvg","Action1") &"实际值: "& ColVal
				end if
			Case "增压后进气压力平均値":
				if(ColVal=Datatable("IntakePressAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("IntakePressAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("IntakePressAvg","Action1") &"实际值: "& ColVal
				end if
			Case "增压后进气压力标准偏差平方":
				if(ColVal=Datatable("IntakePressStdDev","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("IntakePressStdDev","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("IntakePressStdDev","Action1") &"实际值: "& ColVal
				end if
			Case "增压后进气压力范围":
				if(ColVal=Datatable("IntakePressRange","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("IntakePressRange","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("IntakePressRange","Action1") &"实际值: "& ColVal
				end if
			Case "增压后进气压力偏差":
				if(ColVal=Datatable("IntakePressDev","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("IntakePressDev","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("IntakePressDev","Action1") &"实际值: "& ColVal
				end if
			Case "增压后进气温度平均值":
				if(ColVal=Datatable("IntakeTmprtAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("IntakeTmprtAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("IntakeTmprtAvg","Action1") &"实际值: "& ColVal
				end if
			Case "P1圧力平均值":
				if(ColVal=Datatable("P1PressAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("P1PressAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("P1PressAvg","Action1") &"实际值: "& ColVal
				end if
			Case "P1圧力标准偏差平方":
				if(ColVal=Datatable("P1PressStdDev","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("P1PressStdDev","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("P1PressStdDev","Action1") &"实际值: "& ColVal
				end if
			Case "P2圧力平均值":
				if(ColVal=Datatable("P2PressAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("P2PressAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("P2PressAvg","Action1") &"实际值: "& ColVal
				end if
			Case "P2圧力标准偏差平方":
				if(ColVal=Datatable("P2PressStdDev","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("P2PressStdDev","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("P2PressStdDev","Action1") &"实际值: "& ColVal
				end if
			Case "泵电流值平均值":
				if(ColVal=Datatable("PumpElectricAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("PumpElectricAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("PumpElectricAvg","Action1") &"实际值: "& ColVal
				end if
			Case "液压油温平均值":
				if(ColVal=Datatable("FluidTmprtAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("FluidTmprtAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("FluidTmprtAvg","Action1") &"实际值: "& ColVal
				end if
			Case "共轨压力平均值":
				if(ColVal=Datatable("RailPressAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("RailPressAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("RailPressAvg","Action1") &"实际值: "& ColVal
				end if
			Case "共轨压力标准偏差平方":
				if(ColVal=Datatable("RailPressStdDev","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("RailPressStdDev","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("RailPressStdDev","Action1") &"实际值: "& ColVal
				end if
			Case "机油压力平均值":
				if(ColVal=Datatable("OilPressAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("OilPressAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("OilPressAvg","Action1") &"实际值: "& ColVal
				end if
			Case "机油压力标准偏差平方":
				if(ColVal=Datatable("OilPressStdDev","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("OilPressStdDev","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("OilPressStdDev","Action1") &"实际值: "& ColVal
				end if
			Case "加压泵入口压平均值":
				if(ColVal=Datatable("PumpInletPressAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("PumpInletPressAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("PumpInletPressAvg","Action1") &"实际值: "& ColVal
				end if
			Case "EGR开度平均值":
				if(ColVal=Datatable("EGRAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("EGRAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("EGRAvg","Action1") &"实际值: "& ColVal
				end if
			Case "共轨差圧平均值":
				if(ColVal=Datatable("RailPressDiffAvg","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("RailPressDiffAvg","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("RailPressDiffAvg","Action1") &"实际值: "& ColVal
				end if
			Case "共轨差圧标准偏差平方":
				if(ColVal=Datatable("RailPressDiffStdDev","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("RailPressDiffStdDev","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("RailPressDiffStdDev","Action1") &"实际值: "& ColVal
				end if
			Case "吹风风量":
				if(ColVal=Datatable("BlowValue","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("BlowValue","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("BlowValue","Action1") &"实际值: "& ColVal
				end if
			Case "设定温度":
				if(ColVal=Datatable("SetTmprt","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("SetTmprt","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("SetTmprt","Action1") &"实际值: "& ColVal
				end if
			Case "A/C设定":
				if(ColVal=Datatable("ACSet","Action1"))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("ACSet","Action1") &"实际值: "& ColVal
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"-"&ColName&"-检查通过","期望值: "& Datatable("ACSet","Action1") &"实际值: "& ColVal
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
