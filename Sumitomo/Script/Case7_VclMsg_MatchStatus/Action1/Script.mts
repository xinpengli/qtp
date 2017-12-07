'清除车辆信息页的匹配状态信息(即定时信息和匹配未通过信息)
RunAction "Action1 [CleanVclMsgMatchStatus]", oneIteration
'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'检查车辆信息页的匹配状态(匹配成功)
RunAction "Action1 [CheckVclMsgMatchStatus]", oneIteration,0
For i=1 to 2
	'插入源码(匹配未通过)
	RunAction "Action1 [InsertMatchStatusSource]", oneIteration,i
	'检查车辆信息页的匹配状态
	RunAction "Action1 [CheckVclMsgMatchStatus]", oneIteration,i
Next
'系统退出
RunAction "Action1 [Logout]", oneIteration