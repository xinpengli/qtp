'清除历史数据(信息年表统计及相关源码)
RunAction "Action1 [CleanHistoryData_InfoChronology]", oneIteration
'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
For i=1 to 3
	'插入信息年表源码Single
	RunAction "Action1 [InsertInfoChroSource_Single]", oneIteration,i
	'执行统计任务Single
	RunAction "Action1 [ExeStatTask_Daily_Single]", oneIteration,i
	'检查信息年表页Single
	RunAction "Action1 [CheckVclMsgInfoChronology_Single]", oneIteration,i
Next
'系统退出
'RunAction "Action1 [Logout]", oneIteration