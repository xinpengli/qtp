On error resume next
'========加载测试数据---测试用
'datatable.ImportSheet "..\..\Sumitomo\TestData\InsertLockReportSource.xls",1,"Global"
'加载函数文件
executefile  "..\..\Sumitomo\Func&VBS\DBFunc.txt"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'定义并获取delete sql语句
Dim SqlDel
SqlDel="delete Sumitomo.dbo.Msg_MatnInfo_Set where MsgMIS_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_MatnInfo_Set_Reply where MsgMISR_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
'执行sql并返回结果
Dim RetuVal
RetuVal=ExecDB(SqlDel) 
'根据执行结果写日志
if(RetuVal>=0)then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"对应设备"&Datatable("Vcl_No","Global")&"历史数据(保养时间设定和回复)清除完毕","对应设备"&Datatable("Vcl_No","Global")&"历史数据(保养时间设定和回复)清除完毕"
else
reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"对应设备"&Datatable("Vcl_No","Global")&"历史数据(保养时间设定和回复)清除完毕","对应设备"&Datatable("Vcl_No","Global")&"历史数据(保养时间设定和回复)清除完毕"
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
