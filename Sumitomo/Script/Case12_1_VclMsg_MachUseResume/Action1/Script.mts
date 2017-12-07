'加载DB交互的函数文件
executefile  "..\..\Sumitomo\Func&VBS\DBFunc.txt"
''清除历史数据(机器使用履历)
RunAction "Action1 [CleanHistoryData_MachUseResume]", oneIteration
'插入机器使用履历相关源码(日志x20和压力分布x24为1组,共3组源码)
Dim sqlQuery
For i=1  to 6
	RunAction "Action1 [InsertMachUseResuSource]", oneIteration,i
	'源码共3组,1\2  3\4  5\6 故在2\4\6时分别传入参1\2\3
	if( i mod 2="0")then
		'执行exe统计程序
		RunAction "Action1 [ExeStatTask_Daily_Single]", oneIteration,i/2
		'查询机器使用履历页面使用履历对应的压力分布源码串
		sqlQuery="select  MsgP_EngOnSwitch  from sumitomo.dbo.Msg_Pressure where MsgP_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"') and MsgP_MsgTime='"&Datatable("InfoGeneTime"&i,"Global")&"'"
        Datatable("ExpBgColorStr"&(i/2),"Action1")=QueryDBColumn(sqlQuery,"MsgP_EngOnSwitch")
	end if
Next
'登录系统
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'检查机器使用履历页面
RunAction "Action1 [CheckVclMsgMachUseResume]", oneIteration
'退出系统
RunAction "Action1 [Logout]", oneIteration