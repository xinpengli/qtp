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
'========点击车辆信息-最新提交-设置保养剩余时间
if(Browser("住友").Page("主页_车辆信息").Link("最新提交_设置保养剩余时间").Exist)then
Browser("住友").Page("主页_车辆信息").Link("最新提交_设置保养剩余时间").Click
end if
Browser("住友").Page("主页_车辆信息").Sync   '等待页面加载
'========检查是否正常进入“设置保养剩余时间”页
Dim PosiBaoYangTimePage
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_设置保养剩余时间").WebElement("位置>>车辆信息>>设置保养剩余时间").Exist)then
	PosiBaoYangTimePage=Browser("住友").Page("主页_车辆信息").Frame("最新提交_设置保养剩余时间").WebElement("位置>>车辆信息>>设置保养剩余时间").GetROProperty("innertext")
	if(trim(PosiBaoYangTimePage)=Datatable("PosiBaoYangTimePage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入最新提交-设置保养剩余时间页成功","期望值："&Datatable("PosiBaoYangTimePage","Global")&" 实际值："& PosiBaoYangTimePage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入最新提交-设置保养剩余时间页失败","期望值："&Datatable("PosiBaoYangTimePage","Global")&" 实际值："& PosiBaoYangTimePage
	end if
end if
'========检查webtable是否存在，检查配件名，并输入保养剩余时间
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_设置保养剩余时间").WebTable("设置保养剩余时间").Exist)then
	Set wt=Browser("住友").Page("主页_车辆信息").Frame("最新提交_设置保养剩余时间").WebTable("设置保养剩余时间")
	wait 2
	Dim PartsName '定义配件名
	For i=2  to wt.RowCount
		'判断到(最后一行-1时)跳出For循环,因最后一行为按钮所在行,另外还隐藏了一行参数"预滤更管（480专用）"
		if(i=wt.RowCount-1)then
		Exit for
		end if
		'获取配件名,并检查配件名是否与预期一致	
		PartsName=trim(wt.GetCellData(i,2))
		if(PartsName=Datatable("ExpPartsName"& i,"Action1"))then
		reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"配件名检查通过","期望值: "&Datatable("ExpPartsName"& i,"Action1")&" 实际值: "& PartsName
		else
		reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"配件名检查失败","期望值: "&Datatable("ExpPartsName"& i,"Action1")&" 实际值: "& PartsName
		end if
		'设置对应的配件名后边的值
		set WE=wt.ChildItem(i,3,"WebEdit",0)
		WE.set Datatable("ExpPartsTime"& i,"Action1")
	Next
end if
'========点击设置
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_设置保养剩余时间").WebButton("设 置").Exist)then
	Browser("住友").Page("主页_车辆信息").Frame("最新提交_设置保养剩余时间").WebButton("设 置").Click
	'判断弹出框提示语是否设置成功
	if(Browser("住友").Dialog("来自网页的消息").Static("弹出框文本").GetROProperty("text")="设置间隔成功。")then
	reporter.ReportEvent micPass ,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"设置保养剩余时间页-间隔设定成功","设置保养剩余时间页-间隔设定成功"
	else
	reporter.ReportEvent micPass ,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"设置保养剩余时间页-间隔设定失败","设置保养剩余时间页-间隔设定失败"
	end if
	'点击确定按钮
	if(Browser("住友").Dialog("来自网页的消息").WinButton("确定").Exist)then
	Browser("住友").Dialog("来自网页的消息").WinButton("确定").Click
	end if
end if
'置标识位,便于保养剩余时间设置\回复查询页使用
Datatable("BaoYangSetReplyFlag","Global")="设置"
'回到最新提交页
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_设置保养剩余时间").Link("车辆信息").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("最新提交_设置保养剩余时间").Link("车辆信息").Click
end if
Browser("住友").Page("主页_车辆信息").Sync
'========保养的邮件通过数据库查询邮件记录，此页面邮件维护功能已转到其它页面，后续再说
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If
