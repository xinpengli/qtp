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
Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebRadioGroup("保养页面选择").Select "2"
end if
Browser("住友").Page("主页_车辆信息").Sync   '等待页面加载
''刷新页面,避免源码解释慢,页面变更的数据没出来
'if(Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebButton("刷新").Exist)then
'Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebButton("刷新").Click
'Browser("住友").Page("主页_车辆信息").Sync
'end if
'wait 2
'保养履历页面的检查
if(Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebTable("交换部件").Exist)then
	Set wt=Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebTable("交换部件")
	wait 2
	For i=1 to wt.RowCount
		'i为第1、2、3行时，列检查逻辑一致
		if(i<4)then  		
				For j=1 to  wt.ColumnCount(wt.RowCount)
						'j非第8列时，正常取测试数据参数名下标值
						if(j<>8)then
								if(trim(wt.GetCellData(i,j))=Datatable("ObjParts_"&i&"_"&j,"Action1"))then
								reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养履历页面-"&i&"行"&j&"列检查通过","期望值："&Datatable("ObjParts_"&i&"_"&j,"Action1")&"实际值："&  wt.GetCellData(i,j)
								else
								reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养履历页面-"&i&"行"&j&"列检查失败","期望值："&Datatable("ObjParts_"&i&"_"&j,"Action1")&"实际值："&  wt.GetCellData(i,j)
								end if
						'j等于第8列时，测试数据参数名下标值需要特殊区分
						else										
								'如源码插入为第1次则取下标值_行_列_1
								if(Parameter("k")="1")then
										if(trim(wt.GetCellData(i,j))=Datatable("ObjParts_"&i&"_"&j&"_1","Action1"))then
										reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养履历页面-"&i&"行"&j&"列检查通过","期望值："&Datatable("ObjParts_"&i&"_"&j&"_1","Action1")&"实际值："&  wt.GetCellData(i,j)
										else
										reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养履历页面-"&i&"行"&j&"列检查失败","期望值："&Datatable("ObjParts_"&i&"_"&j&"_1","Action1")&"实际值："&  wt.GetCellData(i,j)
										end if
								'如果源码插入为第2次，则取下标值_行_列_2
								else
										if(trim(wt.GetCellData(i,j))=Datatable("ObjParts_"&i&"_"&j&"_2","Action1"))then
										reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养履历页面-"&i&"行"&j&"列检查通过","期望值："&Datatable("ObjParts_"&i&"_"&j&"_2","Action1")&"实际值："&  wt.GetCellData(i,j)
										else
										reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养履历页面-"&i&"行"&j&"列检查失败","期望值："&Datatable("ObjParts_"&i&"_"&j&"_2","Action1")&"实际值："&  wt.GetCellData(i,j)
										end if
								end if
						end if
				Next
		'i=4、5行时，列检查逻辑区别于1、2、3行的列检查，测试数据存放区别
		else          
				For j=1 to  wt.ColumnCount(wt.RowCount)
						if(trim(wt.GetCellData(i,j))=Datatable("ObjParts_"&i&"_"&j,"Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养履历页面-"&i&"行"&j&"列检查通过","期望值："&Datatable("ObjParts_"&i&"_"&j,"Action1")&"实际值："&  wt.GetCellData(i,j)
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"保养履历页面-"&i&"行"&j&"列检查失败","期望值："&Datatable("ObjParts_"&i&"_"&j,"Action1")&"实际值："&  wt.GetCellData(i,j)
						end if
				Next '列检查结束
		end if
	Next '行检查结束
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
