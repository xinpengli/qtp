'清除历史数据(定时和报警)
RunAction "Action1 [CleanHistoryData_Timing&Alarm]", oneIteration
'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'检查车辆信息的工作情况-0表示清除数据后的情况
RunAction "Action1 [CheckVclMsgWorkCondition]", oneIteration,0
'插入源码(定时源码、电瓶被拆报警)
For i=1 to 2
	RunAction "Action1 [InsertWorkConditionSource]", oneIteration,i
Next
'检查车辆信息的工作情况-1表示插入源码数据后的情况
RunAction "Action1 [CheckVclMsgWorkCondition]", oneIteration,1
'系统退出
RunAction "Action1 [Logout]", oneIteration