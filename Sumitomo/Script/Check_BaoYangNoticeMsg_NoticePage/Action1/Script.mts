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
'========点击车辆信息-保养通知信息
if(Browser("住友").Page("主页_车辆信息").Link("保养通知信息").Exist)then
Browser("住友").Page("主页_车辆信息").Link("保养通知信息").Click
end if
Browser("住友").Page("主页_车辆信息").Sync   '等待页面加载
'机器对象表格在9_1用例中有覆盖检查,此处不再重复检查
'选择保养履历页面
if(Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebRadioGroup("保养页面选择").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebRadioGroup("保养页面选择").Select "3"
end if
Browser("住友").Page("主页_车辆信息").Sync   '等待页面加载
''刷新页面,避免源码解释慢,页面变更的数据没出来
'if(Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebButton("刷新").Exist)then
'Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebButton("刷新").Click
'Browser("住友").Page("主页_车辆信息").Sync
'end if
'wait 2
'获取源码操作次数
k=Parameter("k")
'保养履历页面的对应检查
if(Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebTable("交换部件").Exist)then
	Set wt=Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebTable("交换部件")
	wait 2
	For j=1 to wt.ColumnCount(1)
		ColName=wt.GetCellData(1,j)
		Select Case ColName
		Case "通知类型":
			if(trim(wt.GetCellData(2,j))=Datatable("Parts_2_"&j&"_"&k,"Action1"))then
			reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查通过","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			else
			reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查失败","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			end if
		Case "配件名":
			if(trim(wt.GetCellData(2,j))=Datatable("Parts_2_"&j&"_"&k,"Action1"))then
			reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查通过","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			else
			reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查失败","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			end if
		Case "设置间隔":
			if(trim(wt.GetCellData(2,j))=Datatable("Parts_2_"&j&"_"&k,"Action1"))then
			reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查通过","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			else
			reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查失败","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			end if
		Case "剩余时间":
			if(trim(wt.GetCellData(2,j))=Datatable("Parts_2_"&j&"_"&k,"Action1"))then
			reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查通过","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			else
			reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查失败","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			end if
		Case "信息生成时间":
			Datatable("Parts_2_"&j&"_"&k,"Action1")=Datatable("InfoGeneTime"&k,"Global")
			if(trim(wt.GetCellData(2,j))=Datatable("Parts_2_"&j&"_"&k,"Action1"))then
			reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查通过","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			else
			reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查失败","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			end if
		Case "信息接收时间":
			'不做检查，为源码插入数据库时间
		Case "备注":
			if(trim(wt.GetCellData(2,j))=Datatable("Parts_2_"&j&"_"&k,"Action1"))then
			reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查通过","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			else
			reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查失败","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			end if
		Case "编辑":
			if(trim(wt.GetCellData(2,j))=Datatable("Parts_2_"&j&"_"&k,"Action1"))then
			reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查通过","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			else
			reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&ColName&"检查失败","期望值："&Datatable("Parts_2_"&j&"_"&k,"Action1")&"实际值："&wt.GetCellData(2,j)
			end if
			'点击编辑
			set ObjLink=wt.ChildItem(2,j,"Link",0)
			wait 1
			ObjLink.click
			Browser("住友").Page("主页_车辆信息").Sync
			'获取备注，添加备注
'			Set ObjEdit=wt.ChildItem(2,(j-1),"WebEdit",0)
			set ObjEdit=Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebTable("交换部件").ChildItem(2,j-1,"WebEdit",0)
			ObjEdit.set "remark"&k
			'点击更新链接
			set ObjLink=wt.ChildItem(2,j,"Link",0)
			wait 1
			ObjLink.click
			Browser("住友").Page("主页_车辆信息").Sync
		End Select
	Next
end if  'webtable对象判断结束
'回到最新提交页
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
