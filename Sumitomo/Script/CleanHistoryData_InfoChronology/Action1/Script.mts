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
Dim SqlDel,Sql1,Sql2,Sql3,Sql4,Sql5,Sql6,Sql7
'统计表
Sql1="delete sumitomo.dbo.StatData_Daily  where SDD_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');"
'燃料余量数据清除
Sql2="delete sumitomo.dbo.Msg_ResidualFuel where MsgRF_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete sumitomo.dbo.Msg_ResidualFuel_Last where MsgRFL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');"
'维修通知清除
Sql3="delete sumitomo.dbo.Msg_Maintain where MsgM_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete sumitomo.dbo.Msg_Maintain_Last where MsgML_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');"
'防盗动作通知清除
Sql4="delete sumitomo.dbo.Msg_AntiTheft where MsgAT_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete sumitomo.dbo.Msg_AntiTheft_Last where MsgATL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');"
'故障信息清除
Sql5="delete sumitomo.dbo.Msg_Fault where MsgF_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete sumitomo.dbo.Msg_Fault_Content where MsgFC_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete sumitomo.dbo.Msg_Fault_Last where MsgFL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');"
'日志信息清除(含信号强度)
Sql6="delete sumitomo.dbo.Msg_Daily where MsgD_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete sumitomo.dbo.Msg_Daily_Last where MsgDL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');"
'日志校正表信息清除,避免解释程序报错
Sql7="delete sumitomo.dbo.Msg_Daily_Revise where MsgDR_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
SqlDel=Sql1 & Sql2 & Sql3 & Sql4 & Sql5 & Sql6 & Sql7
'执行sql并返回结果
Dim RetuVal
RetuVal=ExecDB(SqlDel) 
'根据执行结果写日志
if(RetuVal>=0)then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"对应设备"&Datatable("Vcl_No","Global")&"信息年表历史数据清除完毕","对应设备"&Datatable("Vcl_No","Global")&"信息年表历史数据清除完毕"
else
reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"对应设备"&Datatable("Vcl_No","Global")&"信息年表历史数据清除完毕","对应设备"&Datatable("Vcl_No","Global")&"信息年表历史数据清除完毕"
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