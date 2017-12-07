On error resume next
'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'当插入新源码后需要刷新当前页面
if(Parameter("j")>0)then
'	Browser("住友").Refresh
'	Browser("住友").Page("主页_车辆信息").Sync
    RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
end if
'检查配对状态
Dim MatchStatus
if(Browser("住友").Page("主页_车辆信息").WebElement("配对状态").Exist)then
	MatchStatus=Browser("住友").Page("主页_车辆信息").WebElement("配对状态").GetROProperty("innertext")
	if(MatchStatus=Datatable("ExpMatchStatus"&Parameter("j"),"Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"匹配状态检查通过","期望值："&Datatable("ExpMatchStatus"&Parameter("j"),"Global")&" 实际值："&MatchStatus
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"匹配状态检查失败","期望值："&Datatable("ExpMatchStatus"&Parameter("j"),"Global")&" 实际值："&MatchStatus
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