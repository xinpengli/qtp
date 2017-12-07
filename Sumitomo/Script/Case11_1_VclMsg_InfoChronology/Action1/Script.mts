'运行两个迭代,,第二迭代验证第一迭代基础上的累计情况
if(Datatable.GetSheet("Global").GetCurrentRow=1)then
	'清除历史数据(信息年表统计及相关源码)
	RunAction "Action1 [CleanHistoryData_InfoChronology]", oneIteration
	'循环插入信息年表源码All (信息生成时间为上月末1天：含燃料余量通知信息、维修通知信息、防盗动作通知信息、故障信息、日志信息等指令)
	For i=1  to 5
		RunAction "Action1 [InsertInfoChroSource_All]", oneIteration,i
	Next
	'系统登录
	RunAction "Action1 [Login]", oneIteration
	'进入车辆信息页
	RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
	'执行统计任务 (设置执行时间为当前月1号,即修改当前机器时间)
	RunAction "Action1 [ExeStatTask_Daily_All]", oneIteration
	'检查信息年表页(根据不同的源码循环做检查)
	For j=1 to 5
		RunAction "Action1 [CheckVclMsgInfoChronology_All]", oneIteration,j
	Next
	'系统退出
	RunAction "Action1 [Logout]", oneIteration
end if
if(Datatable.GetSheet("Global").GetCurrentRow=2)then
	'*****不清除历史数据,验证累计的情况*****
	'循环插入信息年表源码All (信息生成时间为上月末1天：含燃料余量通知信息、维修通知信息、防盗动作通知信息、故障信息、日志信息等指令)
	'*****不再插入日志信息,按实际情况,日志信息一天一次******
	For i=1  to 4
		RunAction "Action1 [InsertInfoChroSource_All]", oneIteration,i
	Next
	'系统登录
	RunAction "Action1 [Login]", oneIteration
	'进入车辆信息页
	RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
	'执行统计任务 (设置执行时间为当前月1号,即修改当前机器时间)
	RunAction "Action1 [ExeStatTask_Daily_All]", oneIteration
	'检查信息年表页(根据不同的源码循环做检查)
	For j=1 to 4
		RunAction "Action1 [CheckVclMsgInfoChronology_All]", oneIteration,j
	Next
	'系统退出
	RunAction "Action1 [Logout]", oneIteration
end if