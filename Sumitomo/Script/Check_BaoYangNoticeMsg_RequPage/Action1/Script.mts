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
'========检查对象机器的两表格内容
if(Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebTable("对象机器1").Exist)then
Set  ObjMac1=Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebTable("对象机器1")
wait 1
For i=1 to ObjMac1.RowCount
	For j=1 to ObjMac1.ColumnCount(1)
			if(trim(ObjMac1.GetCellData(i,j))=Datatable("ExpObjMac1_"&i&"_"&j,"Action1"))then
			reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"对象机器1-"&i&"行"&j&"列检查通过","期望值："&Datatable("ExpObjMac1_"&i&"_"&j,"Action1")&" 实际值："& ObjMac1.GetCellData(i,j)
			else
			reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"对象机器1-"&i&"行"&j&"列检查失败","期望值："&Datatable("ExpObjMac1_"&i&"_"&j,"Action1")&" 实际值："& ObjMac1.GetCellData(i,j)
			end if
	Next
Next
end if
if(Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebTable("对象机器2").Exist)then
Set ObjMac2=Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebTable("对象机器2")
wait 1
'Datatable("ExpObjMac2_2_2","Action1")="个月"  '设置购买时长预期值--待完善
For m=1 to ObjMac2.RowCount
	For n=1 to ObjMac2.ColumnCount(1)
			if(trim(ObjMac2.GetCellData(m,n))=Datatable("ExpObjMac2_"&m&"_"&n,"Action1"))then
			reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"对象机器2-"&m&"行"&n&"列检查通过","期望值："&Datatable("ExpObjMac2_"&m&"_"&n,"Action1")&" 实际值："& ObjMac2.GetCellData(m,n)
			else
			reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"对象机器2-"&m&"行"&n&"列检查失败","期望值："&Datatable("ExpObjMac2_"&m&"_"&n,"Action1")&" 实际值："& ObjMac2.GetCellData(m,n)
			end if
	Next
Next
end if
'刷新页面,避免源码解释慢,页面变更的数据没出来
if(Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebButton("刷新").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebButton("刷新").Click
Browser("住友").Page("主页_车辆信息").Sync
end if
wait 2
'========检查机器信息的两表格内容
'机器工作小时数检查 放到履历或通知页面检查，此页面不作检查了
'交换部件表格内容检查
if(Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebTable("交换部件").Exist)then
	Set ObjParts=Browser("住友").Page("主页_车辆信息").Frame("保养通知信息").WebTable("交换部件")
	wait 1
	Dim p,q '定义循环下标
	For p=2 to ObjParts.RowCount  'webtable行遍历开始
		'因为第8行webtable数据是源码关联的变化行，测试数据存储不同于其它行，故分开检查
		if(p=8)then
				For q=1 to ObjParts.ColumnCount(1)
					if(trim(ObjParts.GetCellData(p,q))=Datatable("ObjParts_"&p&"_"&q&"_"&Parameter("i"),"Action1"))then
					reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"交换部件表格"&p&"行"&q&"列"&Parameter("i")&"次检查通过","期望值："&Datatable("ObjParts_"&p&"_"&q&"_"&Parameter("i"),"Action1")&" 实际值："& ObjParts.GetCellData(p,q)
					else
					reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"交换部件表格"&p&"行"&q&"列"&Parameter("i")&"检查失败","期望值："&Datatable("ObjParts_"&p&"_"&q&"_"&Parameter("i"),"Action1")&" 实际值："& ObjParts.GetCellData(p,q)
					end if
				Next
		'非第8行的webtable数据检查
		else
		        '如果插入的源码非第1次，则正常检查
		        if(Parameter("i")>=1)then
					For q=1 to ObjParts.ColumnCount(1)
						if(trim(ObjParts.GetCellData(p,q))=Datatable("ObjParts_"&p&"_"&q,"Action1"))then
						reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"交换部件表格"&p&"行"&q&"列检查通过","期望值："&Datatable("ObjParts_"&p&"_"&q,"Action1")&" 实际值："& ObjParts.GetCellData(p,q)
						else
						reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"交换部件表格"&p&"行"&q&"列检查失败","期望值："&Datatable("ObjParts_"&p&"_"&q,"Action1")&" 实际值："& ObjParts.GetCellData(p,q)
						end if
					Next
			    '如果插入的源码是第0次，则第4列检查要取默认值，前3列还按行列标取值
				else				  
					if(Parameter("i")=0)then
						For q=1 to ObjParts.ColumnCount(1)
							if(q=ObjParts.ColumnCount(1))then
									if(trim(ObjParts.GetCellData(p,q))=Datatable("ObjParts_"&p&"_"&q&"_"&"default","Action1"))then
									reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"交换部件表格"&p&"行"&q&"列默认值检查通过","期望值："&Datatable("ObjParts_"&p&"_"&q&"_"&"default","Action1")&" 实际值："& ObjParts.GetCellData(p,q)
									else
									reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"交换部件表格"&p&"行"&q&"列默认值检查失败","期望值："&Datatable("ObjParts_"&p&"_"&q&"_"&"default","Action1")&" 实际值："& ObjParts.GetCellData(p,q)
									end if
							else					       
									if(trim(ObjParts.GetCellData(p,q))=Datatable("ObjParts_"&p&"_"&q,"Action1"))then
									reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"交换部件表格"&p&"行"&q&"列检查通过","期望值："&Datatable("ObjParts_"&p&"_"&q,"Action1")&" 实际值："& ObjParts.GetCellData(p,q)
									else
									reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"交换部件表格"&p&"行"&q&"列检查失败","期望值："&Datatable("ObjParts_"&p&"_"&q,"Action1")&" 实际值："& ObjParts.GetCellData(p,q)
									end if
							end if
						Next
					end if
			   end if  
		end if  '非第8行的webtable数据检查结束
	Next   'webtable行遍历结束
end if '交换部件表格判断结束
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
