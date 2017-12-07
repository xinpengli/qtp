On error resume next
'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'========检查是否正常进入"机器档案"--“信息登录”页
Dim InfoLoginPage
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebElement("位置").Exist)then
	InfoLoginPage=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebElement("位置").GetROProperty("innertext")
	if(trim(InfoLoginPage)=Datatable("InfoLoginPage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息-机器档案-信息登录页成功","期望值："&Datatable("InfoLoginPage","Global")&" 实际值："& InfoLoginPage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入车辆信息-机器档案-信息登录页失败","期望值："&Datatable("InfoLoginPage","Global")&" 实际值："& InfoLoginPage
	end if
end if
'========检查信息登录页各项
'信息登录--机号检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("机号").Exist)then
	InfoLogin_JiHao=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("机号").Object.value
	if(trim(InfoLogin_JiHao)=Datatable("InfoLogin_JiHao","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-机号检查通过","期望值: "&Datatable("InfoLogin_JiHao","Action1")&" 实际值："&InfoLogin_JiHao
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-机号检查失败","期望值: "&Datatable("InfoLogin_JiHao","Action1")&" 实际值："&InfoLogin_JiHao
	end if
end if
'信息登录--机器类别检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebList("机器类别").Exist)then
InfoLogin_MachCate=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebList("机器类别").GetROProperty("Value")
	if(trim(InfoLogin_MachCate)=Datatable("InfoLogin_MachCate","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-机器类别检查通过","期望值: "&Datatable("InfoLogin_MachCate","Action1")&" 实际值："& InfoLogin_MachCate
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-机器类别检查失败","期望值: "&Datatable("InfoLogin_MachCate","Action1")&" 实际值："& InfoLogin_MachCate
	end if
end if
'信息登录-机器型号检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebList("机器型号").Exist)then
InfoLogin_MachType=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebList("机器型号").GetROProperty("Value")
	if(trim(InfoLogin_MachType)=Datatable("InfoLogin_MachType","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-机器型号检查通过","期望值: "&Datatable("InfoLogin_MachType","Action1")&" 实际值："& InfoLogin_MachType
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-机器型号检查失败","期望值: "&Datatable("InfoLogin_MachType","Action1")&" 实际值："& InfoLogin_MachType
	end if
end if
'信息登录-生产部门检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("生产部门").Exist)then
InfoLogin_ProdDep=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("生产部门").Object.value
	if(trim(InfoLogin_ProdDep)=Datatable("InfoLogin_ProdDep","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-生产部门检查通过","期望值: "&Datatable("InfoLogin_ProdDep","Action1")&" 实际值："&InfoLogin_ProdDep
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-生产部门检查失败","期望值: "&Datatable("InfoLogin_ProdDep","Action1")&" 实际值："&InfoLogin_ProdDep
	end if
end if
'信息登录-销售公司检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("销售公司").Exist)then
InfoLogin_SaleComp=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("销售公司").Object.value
	if(trim(InfoLogin_SaleComp)=Datatable("InfoLogin_SaleComp","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-销售公司检查通过","期望值: "&Datatable("InfoLogin_SaleComp","Action1")&" 实际值："&InfoLogin_SaleComp
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-销售公司检查失败","期望值: "&Datatable("InfoLogin_SaleComp","Action1")&" 实际值："&InfoLogin_SaleComp
	end if
end if
'信息登录-经销商检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("经销商").Exist)then
InfoLogin_Agency=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("经销商").Object.value
	if(trim(InfoLogin_Agency)=Datatable("InfoLogin_Agency","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-经销商检查通过","期望值: "&Datatable("InfoLogin_Agency","Action1")&" 实际值："&InfoLogin_Agency
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-经销商检查失败","期望值: "&Datatable("InfoLogin_Agency","Action1")&" 实际值："&InfoLogin_Agency
	end if
end if
'信息登录-销售人员检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("销售人员").Exist)then
InfoLogin_Sale=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("销售人员").Object.value
	if(trim(InfoLogin_Sale)=Datatable("InfoLogin_Sale","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-销售人员检查通过","期望值: "&Datatable("InfoLogin_Sale","Action1")&" 实际值："&InfoLogin_Sale
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-销售人员检查失败","期望值: "&Datatable("InfoLogin_Sale","Action1")&" 实际值："&InfoLogin_Sale
	end if
end if
'信息登录-客户名称检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("客户名称").Exist)then
InfoLogin_GuestName=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("客户名称").Object.value
	if(trim(InfoLogin_GuestName)=Datatable("InfoLogin_GuestName","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-客户名称检查通过","期望值: "&Datatable("InfoLogin_GuestName","Action1")&" 实际值："&InfoLogin_GuestName
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-客户名称检查失败","期望值: "&Datatable("InfoLogin_GuestName","Action1")&" 实际值："&InfoLogin_GuestName
	end if
end if
'信息登录-购机时间检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("购机时间").Exist)then
InfoLogin_BuyDate=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("购机时间").Object.value
	if(trim(InfoLogin_BuyDate)=Datatable("InfoLogin_BuyDate","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-购机时间检查通过","期望值: "&Datatable("InfoLogin_BuyDate","Action1")&" 实际值："&InfoLogin_BuyDate
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-购机时间检查失败","期望值: "&Datatable("InfoLogin_BuyDate","Action1")&" 实际值："&InfoLogin_BuyDate
	end if
end if
'信息登录-购机方式检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebList("购机方式").Exist)then
InfoLogin_BuyMethod=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebList("购机方式").GetROProperty("Value")
	if(trim(InfoLogin_BuyMethod)=Datatable("InfoLogin_BuyMethod","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-购机方式检查通过","期望值: "&Datatable("InfoLogin_BuyMethod","Action1")&" 实际值："& InfoLogin_BuyMethod
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-购机方式检查失败","期望值: "&Datatable("InfoLogin_BuyMethod","Action1")&" 实际值："& InfoLogin_BuyMethod
	end if
end if
'信息登录-客户行业检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebList("客户行业").Exist)then
InfoLogin_CustIndustry=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebList("客户行业").GetROProperty("Value")
	if(trim(InfoLogin_CustIndustry)=Datatable("InfoLogin_CustIndustry","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-客户行业检查通过","期望值: "&Datatable("InfoLogin_CustIndustry","Action1")&" 实际值："& InfoLogin_CustIndustry
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-客户行业检查失败","期望值: "&Datatable("InfoLogin_CustIndustry","Action1")&" 实际值："& InfoLogin_CustIndustry
	end if
end if
'信息登录-交机日期检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("交机日期").Exist)then
InfoLogin_DeliDate=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("交机日期").Object.value
	if(trim(InfoLogin_DeliDate)=Datatable("InfoLogin_DeliDate","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-交机日期检查通过","期望值: "&Datatable("InfoLogin_DeliDate","Action1")&" 实际值："&InfoLogin_DeliDate
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-交机日期检查失败","期望值: "&Datatable("InfoLogin_DeliDate","Action1")&" 实际值："& InfoLogin_DeliDate
	end if
end if
'信息登录-出保日期检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("出保日期").Exist)then
InfoLogin_OutOfDate=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("出保日期").Object.value
	if(trim(InfoLogin_OutOfDate)=Datatable("InfoLogin_OutOfDate","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-出保日期检查通过","期望值: "&Datatable("InfoLogin_OutOfDate","Action1")&" 实际值："& InfoLogin_OutOfDate
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-出保日期检查失败","期望值: "&Datatable("InfoLogin_OutOfDate","Action1")&" 实际值："& InfoLogin_OutOfDate
	end if
end if
'信息登录-保修类别检查
if(Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("保修类别").Exist)then
InfoLogin_WarrCate=Browser("住友_车辆信息_机器档案链接页").Page("信息登录").WebEdit("保修类别").Object.value
	if(trim(InfoLogin_WarrCate)=Datatable("InfoLogin_WarrCate","Action1"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-保修类别检查通过","期望值: "&Datatable("InfoLogin_WarrCate","Action1")&" 实际值："& InfoLogin_WarrCate
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"机器档案-信息登录-保修类别检查失败","期望值: "&Datatable("InfoLogin_WarrCate","Action1")&" 实际值："& InfoLogin_WarrCate
	end if
end if
'========关闭信息登录页
Browser("住友_车辆信息_机器档案链接页").Close
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If
