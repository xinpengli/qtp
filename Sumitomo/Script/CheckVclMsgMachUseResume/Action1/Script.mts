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
if(Browser("住友").Page("主页_车辆信息").Link("机器使用履历").Exist)then
Browser("住友").Page("主页_车辆信息").Link("机器使用履历").Click
Browser("住友").Page("主页_车辆信息").Sync
end if
''========检查是否正常进入“车辆信息/机器使用履历”页
if(Browser("住友").Page("主页_车辆信息").Frame("机器使用履历").WebElement("位置").Exist)then
	Dim PosiMachUseResuPage
	PosiMachUseResuPage=Browser("住友").Page("主页_车辆信息").Frame("机器使用履历").WebElement("位置").GetROProperty("innertext")
	if(trim(PosiMachUseResuPage)=Datatable("PosiMachUseResuPage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息-机器使用履历页成功","期望值："&Datatable("PosiMachUseResuPage","Global")&" 实际值："& PosiMachUseResuPage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息-机器使用履历页失败","期望值："&Datatable("PosiMachUseResuPage","Global")&" 实际值："& PosiMachUseResuPage
	end if
end if
'========输入机器使用履历查询条件：天
if(Browser("住友").Page("主页_车辆信息").Frame("机器使用履历").WebRadioGroup("查询条件按钮组").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("机器使用履历").WebRadioGroup("查询条件按钮组").Select "rad_select_time"
end if
if(Browser("住友").Page("主页_车辆信息").Frame("机器使用履历").WebEdit("按天查询_开始").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("机器使用履历").WebEdit("按天查询_开始").Object.value=date-2 '日期同待查的日志时间一致
end if
if(Browser("住友").Page("主页_车辆信息").Frame("机器使用履历").WebEdit("按天查询_结束").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("机器使用履历").WebEdit("按天查询_结束").Object.value= date '日期同待查的日志时间一致
end if
if(Browser("住友").Page("主页_车辆信息").Frame("机器使用履历").WebButton("查询").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("机器使用履历").WebButton("查询").Click
end if
Browser("住友").Page("主页_车辆信息").Sync
'========检查机器使用履历列表
if(Browser("住友").Page("主页_车辆信息").Frame("机器使用履历").WebTable("机器使用履历").Exist)then
	Set MachUseResuList=Browser("住友").Page("主页_车辆信息").Frame("机器使用履历").WebTable("机器使用履历")
	wait 2
	Dim i,j,ActVal
	Dim k
	k=0
	For i=1 to MachUseResuList.RowCount
			'======1行和6行(末行)的检查
			If(i=1 or i=MachUseResuList.RowCount)Then
				For j=1 to MachUseResuList.ColumnCount(i)
					ActVal=MachUseResuList.GetCellData(i,j)
					If(ActVal=Datatable("ObjMacUseResu_"&i&"_"&j,"Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器使用履历表"&i&"行"&j&"列检查通过","期望值: "&Datatable("ObjMacUseResu_"&i&"_"&j,"Action1")&" 实际值: "& ActVal
					else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器使用履历表"&i&"行"&j&"列检查失败","期望值: "&Datatable("ObjMacUseResu_"&i&"_"&j,"Action1")&" 实际值: "& ActVal
					end if
				Next
			End If
			'======2行的检查,此行预期值是固定的,未在Action1中体现,直接脚本控制
			if(i=2)then
				For j=1 to MachUseResuList.ColumnCount(i)
					ActVal=MachUseResuList.GetCellData(i,j)
					'固定的24列检查
					if(ActVal=right("0"&k,2))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器使用履历表"&i&"行"&j&"列检查通过","期望值: "&right("0"&k,2)&" 实际值: "& ActVal
					else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器使用履历表"&i&"行"&j&"列检查失败","期望值: "&right("0"&k,2)&" 实际值: "& ActVal
					end if
					k=k+1		
				Next
			end if  '2行的检查结束
			'======3-5行的检查
			if(i=3 or i=4 or i=5)then
				'设置第一列的日期值,同源码中的日志日期
				Datatable("ObjMacUseResu_3_1","Action1")=right("0"&month(Date),2) &"-"& right("0"&day(Date),2)
				Datatable("ObjMacUseResu_4_1","Action1")=right("0"&month(Date-1),2) &"-"& right("0"&day(Date-1),2)
				Datatable("ObjMacUseResu_5_1","Action1")=right("0"&month(Date-2),2) &"-"& right("0"&day(Date-2),2)
				'遍历检查列
				Dim ExpValIndex  '定义期望的机器履历背景色源码串下标,并赋值为0
				ExpValIndex=1
				For j=1 to MachUseResuList.ColumnCount(i)					
					'当列是1,2和99,100,101时取Datatable中预期值
					if(j=1 or j=2 or j=99 or j=100 or j=101)then
						ActVal=trim(MachUseResuList.GetCellData(i,j))
						if(ActVal=Datatable("ObjMacUseResu_"&i&"_"&j,"Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器使用履历表"&i&"行"&j&"列检查通过","期望值: "&Datatable("ObjMacUseResu_"&i&"_"&j,"Action1")&" 实际值: "& ActVal
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器使用履历表"&i&"行"&j&"列检查失败","期望值: "&Datatable("ObjMacUseResu_"&i&"_"&j,"Action1")&" 实际值: "& ActVal
						end if
					end if
					'当列数>2 & <99时,检查单元格属性是否有背景色
					if(j>2 and j<99)then
					   '获取背景色实际值******通过object取值时，下标从0开始*******
						BgColorVal=MachUseResuList.object.rows(i-1).cells(j-1).bgColor 
						'获取背景色期望值源码:逐个取源码中的每一位
						Select Case i
							Case "3":
								ExpBgColorVal=mid(Datatable("ExpBgColorStr3","Action1"),ExpValIndex,1)
							Case "4":
								ExpBgColorVal=mid(Datatable("ExpBgColorStr2","Action1"),ExpValIndex,1)
							Case "5":
								ExpBgColorVal=mid(Datatable("ExpBgColorStr1","Action1"),ExpValIndex,1)
						End Select
						'机器履历背景色期望值源码串下标值自增
						ExpValIndex=ExpValIndex+1 
						'检查实际值背景色和期望值源码,表示一致:#0000ff表示蓝色,源码中为1; 为空表示白色,源码中为0
						if((BgColorVal=Datatable("ExpBgColor","Action1")  and ExpBgColorVal="1") or (BgColorVal=Datatable("ExpBgColor_blank","Action1")  and ExpBgColorVal="0"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器使用履历表"&i&"行"&j&"列背景色检查通过","期望值：源码串中为"& ExpBgColorVal &"，实际值：页面展示背景色串为"&Chr(34)&BgColorVal&Chr(34)
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器使用履历表"&i&"行"&j&"列背景色检查失败","期望值：源码串中为"& ExpBgColorVal &"，实际值：页面展示背景色为"&Chr(34)&BgColorVal&Chr(34)
						end if
					end if  '列数>2 & <99判断结束
				Next '3-5行的列遍历检查结束
			end if '3-5行的检查结束
	Next '行遍历结束
end if  'webtable对象判断结束
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If
