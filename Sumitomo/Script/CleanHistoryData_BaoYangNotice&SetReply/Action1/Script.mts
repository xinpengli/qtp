﻿On error resume next
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
SqlDel="delete Sumitomo.dbo.Msg_MatnInfo_Set where msgmis_vcl_id=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_MatnInfo_Set_Reply where msgmisr_vcl_id=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_MatnNotice where msgmn_vcl_id=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_MatnNotice_Last where msgmnl_vcl_id=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"');delete Sumitomo.dbo.Msg_MatnNotice_Parts_Last where msgmnpl_vcl_id=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"')"
'执行sql并返回结果
Dim RetuVal
RetuVal=ExecDB(SqlDel) 
'根据执行结果写日志
if(RetuVal>=0)then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"设备"&Datatable("Vcl_No","Global")&"历史数据清除完毕","设备"&Datatable("Vcl_No","Global")&"历史数据清除完毕"
else
reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"设备"&Datatable("Vcl_No","Global")&"历史数据清除完毕","设备"&Datatable("Vcl_No","Global")&"历史数据清除完毕"
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
