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
Select Case Datatable("MsgType","Global")
	 Case "定时":
		 SqlDel="delete  Sumitomo.dbo.Msg_Timing_"& TempDate &" where msgt_vcl_id=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"') ;delete  Sumitomo.dbo.Msg_Timing_Last where MsgTL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_Daily_Revise where MsgDR_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
	 Case "日志":
		 SqlDel="delete  Sumitomo.dbo.Msg_Daily where MsgD_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete  Sumitomo.dbo.Msg_Daily_Last where MsgDL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_Daily_Revise where MsgDR_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
	 Case "故障":
		 SqlDel="delete Sumitomo.dbo.Msg_Fault where MsgF_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_Fault_Content where MsgFC_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_Fault_Last where MsgFL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_Daily_Revise where MsgDR_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
	 Case "防盗动作通知":
		 SqlDel="delete Sumitomo.dbo.Msg_AntiTheft where MsgAT_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_AntiTheft_Last where MsgATL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
	 Case "维修通知":
		 SqlDel="delete Sumitomo.dbo.Msg_Maintain where MsgM_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete  Sumitomo.dbo.Msg_Maintain_Last where MsgML_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
	 Case "燃料余量通知":
		 SqlDel="delete Sumitomo.dbo.Msg_ResidualFuel where MsgRF_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_ResidualFuel_Last where MsgRFL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
	 Case "复位通知":
		 SqlDel="delete Sumitomo.dbo.Msg_Rest where MsgR_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete  Sumitomo.dbo.Msg_Rest_Last where MsgRL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
	 Case "控制器配对异常":
		 SqlDel="delete Sumitomo.dbo.Msg_PairFailed where MsgPF_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete  Sumitomo.dbo.Msg_PairFailed_Last where MsgPFL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
	 Case "保养信息":
		 SqlDel="delete Sumitomo.dbo.Msg_MatnInfo where MsgMI_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_MatnInfo_Last where MsgMIL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
	 Case "取样数据信息":
		 SqlDel="delete Sumitomo.dbo.Msg_Timing_Para_"&TempDate&"  where MsgTP_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete  Sumitomo.dbo.Msg_Timing_Para_last where MsgTPL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
	 Case Other
		 msgbox "无匹配的源码信息类型数据要清除"
End Select
'执行sql并返回结果
Dim RetuVal
RetuVal=ExecDB(SqlDel) 
'根据执行结果写日志
if(RetuVal>=0)then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"对应设备"&Datatable("Vcl_No","Global")&"历史数据清除完毕",Datatable("MsgType","Global")&"对应设备"&Datatable("Vcl_No","Global")&"历史数据清除完毕"
else
reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&Datatable("MsgType","Global")&"对应设备"&Datatable("Vcl_No","Global")&"历史数据清除完毕",Datatable("MsgType","Global")&"对应设备"&Datatable("Vcl_No","Global")&"历史数据清除完毕"
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