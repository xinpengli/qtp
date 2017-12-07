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
Dim TempDate '定义并获取日期，用于表名拼接
TempDate=year(date)&right("0"&month(date),2)&right("0"&day(date),2)
'如不清除定时的历史数据,因后续测试源码含全复位指令,会导致在点检实绩中展示:点检实绩中工作小时数+历史定时last中工作小时数(本用例中不涉及定时的插入,这样会造成预期结果数据不固定)
SqlDel="update [Sumitomo].[dbo].[VclInfo] set Vcl_ECMNo='',Vcl_CtrlrANo='',Vcl_CtrlrTNo='',Vcl_CtrlrTSN='',Vcl_CtrlrBNo=null,Vcl_FirmwareVersion=null,Vcl_FlashVersion=null where Vcl_No='"&Datatable("Vcl_No","Global")&"';delete Sumitomo.dbo.Msg_Tmnl_Set2 where MsgTS2_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_Tmnl_Set2_Reply where MsgTS2R_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_Daily_Revise where MsgDR_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_Touch where Msgth_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_Touch_Last where Msgthl_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_Timing_"&TempDate&"  where msgt_vcl_id=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_Timing_Last where MsgTL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
'执行sql并返回结果
Dim RetuVal
RetuVal=ExecDB(SqlDel) 
'根据执行结果写日志
if(RetuVal>=0)then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"对应设备"&Datatable("Vcl_No","Global")&"历史数据清除完毕","对应设备"&Datatable("Vcl_No","Global")&"历史数据清除完毕"
else
reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"对应设备"&Datatable("Vcl_No","Global")&"历史数据清除完毕","对应设备"&Datatable("Vcl_No","Global")&"历史数据清除完毕"
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