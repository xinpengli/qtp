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
'========点击车辆信息页-最新提交
if(Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Exist)then
Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Click
end if
'========点击车辆信息-最新提交-查看保养剩余时间设置/回复信息
if(Browser("住友").Page("主页_车辆信息").Link("最新提交_查看保养剩余时间设置/回复信息").Exist)then
Browser("住友").Page("主页_车辆信息").Link("最新提交_查看保养剩余时间设置/回复信息").Click
end if
Browser("住友").Page("主页_车辆信息").Sync   '等待页面加载
'========检查是否正常进入“查看保养剩余时间设置/回复信息”页
Dim PosiBaoYangTimeSetAndReplyPage
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebElement("位置>>车辆信息>>查看保养剩余时间设置/回复信息").Exist)then
	PosiBaoYangTimeSetAndReplyPage=Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebElement("位置>>车辆信息>>查看保养剩余时间设置/回复信息").GetROProperty("innertext")
	if(trim(PosiBaoYangTimeSetAndReplyPage)=Datatable("PosiBaoYangTimeSetAndReplyPage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入最新提交-查看保养剩余时间设置/回复信息页成功","期望值："&Datatable("PosiBaoYangTimeSetAndReplyPage","Global")&" 实际值："& PosiBaoYangTimeSetAndReplyPage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入最新提交-查看保养剩余时间设置/回复信息页失败","期望值："&Datatable("PosiBaoYangTimeSetAndReplyPage","Global")&" 实际值："& PosiBaoYangTimeSetAndReplyPage
	end if
end if
'========输入时间段
Dim currDay '获取当天日期
currDay=Cstr(Year(Date)&"-"&right("0"&Month(Date),2)&"-"&right("0"&Day(Date),2))
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebEdit("开始时间").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebEdit("开始时间").Object.value=currDay
end if
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebEdit("结束时间").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebEdit("结束时间").Object.value=currDay
end if
'========选择设置类型
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebList("设置/回复").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebList("设置/回复").Select  Datatable("BaoYangSetReplyFlag","Global")
end if
'========点击查询
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebButton("查询").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebButton("查询").Click
end if
Browser("住友").Page("主页_车辆信息").Sync '等待加载
'========检查webtable是否存在，检查配件名及设置的剩余时间
'保养这块检查的实现基本还是按字段的固定顺序来的,一旦字段顺序打乱,得调测试数据; 因字段命名太麻烦,否则可以按字段名查询检查,暂时先按当前方式走
Dim ColName,ExpColName,ColVal,ExpColVal
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebTable("保养剩余时间设置/回复信息列表").Exist)then
	Set wt=Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebTable("保养剩余时间设置/回复信息列表")
	wait 2
	'检查设置列表
	if(Datatable("BaoYangSetReplyFlag","Global")="设置")then
		For j=2  to wt.ColumnCount(1)
			'提交时间列不做检查，此时间为提交时记录的服务器时间
			'获取列名、列值
			ColName=trim(wt.GetCellData(1,j))
			ExpColName=Datatable("ExpPartsName"&j,"Action1")
			ColVal=trim(wt.GetCellData(2,j))
			ExpColVal=Datatable("ExpPartsTime"&j,"Action1")
			Select Case ColName			
				Case ExpColName:
					if(ColVal=ExpColVal)then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养剩余时间设置查询-"&ExpColName&"检查通过","期望值："& ExpColVal &" 实际值："& ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养剩余时间设置查询-"&ExpColName&"检查失败","期望值："& ExpColVal &" 实际值："& ColVal
					end if
			End Select
		Next
	end if
	'检查回复列表
	if(Datatable("BaoYangSetReplyFlag","Global")="回复")then
	'获取实际的信息生成时间,即信息提交时间
	Datatable("ExpRepPartsName1","Action1")=Datatable("InfoGeneTime","Global")
		For j=1  to wt.ColumnCount(1)
			'获取列名、列值
			ColName=trim(wt.GetCellData(1,j))
			ExpColName=Datatable("ExpRepPartsName"&j,"Action1")
			ColVal=trim(wt.GetCellData(2,j))
			ExpColVal=Datatable("ExpRepPartsTime"&j,"Action1")
			Select Case ColName			
				Case ExpColName:
					if(ColVal=ExpColVal)then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养剩余时间设置回复查询-"&ExpColName&"检查通过","期望值："& ExpColVal &" 实际值："& ColVal
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养剩余时间设置回复查询-"&ExpColName&"检查失败","期望值："& ExpColVal &" 实际值："& ColVal
					end if
			End Select
		Next
	end if
end if
'========下载功能
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebButton("下载").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").WebButton("下载").Click
end if
wait 2
RunAction "Action1 [DownFile]", oneIteration
'========回到最新提交页
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").Link("车辆信息").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看保养剩余时间设置/回复信息").Link("车辆信息").Click
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
