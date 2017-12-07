'清除机器档案相关历史数据
RunAction "Action1 [CleanHistoryData_MachineFile]", oneIteration
'插入机器档案相关源码(初期设定指令x30\初期设定完成信息指令x32\全复位通知信息x1A指令\机器Touch通知信息xA2指令)
For i=1  to 4
	'插入机器档案相关指令(初期设定指令x30,初期设定完成信息指令x32,全复位通知信息x1A,机器Touch通知信息xA2)
	RunAction "Action1 [InsertMachineFileSource_All]", oneIteration,i
Next
'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'检查机器档案页面
RunAction "Action1 [Check_MachineFilePage]", oneIteration
'退出系统
RunAction "Action1 [Logout]", oneIteration