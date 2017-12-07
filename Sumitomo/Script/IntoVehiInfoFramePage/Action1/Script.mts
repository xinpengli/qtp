On error resume next
'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'点击导航栏"查询中心"
if(Browser("住友").Page("主页").Frame("左导航栏").WebElement("查询中心").Exist)then
Browser("住友").Page("主页").Frame("左导航栏").WebElement("查询中心").Click
end if
'判断是否进入查询中心页
Dim PosiQueryCenterPage
if(Browser("住友").Page("主页_查询中心").Frame("查询中心").WebElement("位置:查询中心").Exist)then
	PosiQueryCenterPage=trim(Browser("住友").Page("主页_查询中心").Frame("查询中心").WebElement("位置:查询中心").GetROProperty("innertext"))
	if(PosiQueryCenterPage=Datatable("PosiQueryCenterPage"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入查询中心页成功","期望位置:"&Datatable("PosiQueryCenterPage")&"  实际位置:"&PosiQueryCenterPage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入查询中心页失败","期望位置:"&Datatable("PosiQueryCenterPage")&"  实际位置:"&PosiQueryCenterPage
	end if
end if
'输入机号,进行唯一查询
if(Browser("住友").Page("主页_查询中心").Frame("查询中心").WebEdit("机号").Exist)then
Browser("住友").Page("主页_查询中心").Frame("查询中心").WebEdit("机号").Set Datatable("Vcl_No","Global")
end if
if(Browser("住友").Page("主页_查询中心").Frame("查询中心").WebButton("唯一查询").Exist)then
Browser("住友").Page("主页_查询中心").Frame("查询中心").WebButton("唯一查询").Click
end if
'判断是否进入车辆信息页
Dim PosiVehiMsgPage
if(Browser("住友").Page("主页_车辆信息").WebElement("您的位置>>车辆信息").Exist)then
	PosiVehiMsgPage=trim(Browser("住友").Page("主页_车辆信息").WebElement("您的位置>>车辆信息").GetROProperty("innertext"))
	if(PosiVehiMsgPage=Datatable("PosiVehiMsgPage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息页成功","期望位置:"&Datatable("PosiVehiMsgPage")&"  实际位置:"&PosiVehiMsgPage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息页失败","期望位置:"&Datatable("PosiVehiMsgPage")&"  实际位置:"&PosiVehiMsgPage
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
