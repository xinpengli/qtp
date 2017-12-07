On error resume next
'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'获取检查次数
Dim CheckCount
CheckCount=Parameter("count")
'在第4次和5次调用此action时,因有定时信息了,故需要将开机状态和位置中的时间替换为定时指令的信息生成时间
if(CheckCount =4 or CheckCount=5)then
Datatable("VehicleInfoTreeView_item_2"&"_"& CheckCount,"Action1")="在"&Datatable("InfoGeneTime","Global")&"处于关机状态" 
Datatable("VehicleInfoTreeView_item_3"&"_"& CheckCount,"Action1")="最新位置 — "&Datatable("InfoGeneTime","Global")&"河北省石家庄市裕华区 黄河大道附近"
end if
'========检查工作状态/位置页
'获取Frame对象
set oFrame=Browser("住友").Page("主页_车辆信息").Frame("工作状态/位置").Object
'循环i值为web元素对应的ID中数值，自增值，目前脚本和web元素对应，之后有变更，对应修改测试数据即可（此处未用到工作状态/位置 下抓取的对象）
For i=2 to 5 
	'通过元素ID获取检查项
	set wt= oFrame.getElementById("VehicleInfoTreeView_item_"&i)
	'获取检查项的值，位置固定，且每个检查项均为1行的table
	ActVal=trim(wt.rows(0).cells(4).innertext)
	ExpVal=trim(Datatable("VehicleInfoTreeView_item_"&i &"_"& CheckCount,"Action1"))
'	msgbox  ActVal
'	msgbox  ExpVal
	'检查实际值和期望值
	if(ActVal=ExpVal)then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"工作状态/位置VehicleInfoTreeView_item_"&i &"  "& CheckCount &"次检查通过","期望值："& ExpVal &"实际值："&ActVal
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"工作状态/位置VehicleInfoTreeView_item_"&i &"  "&CheckCount &"次检查失败","期望值："& ExpVal &"实际值："&ActVal
	end if
Next
'========回到车辆信息页
if(Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").Link("车辆信息").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").Link("车辆信息").Click
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
