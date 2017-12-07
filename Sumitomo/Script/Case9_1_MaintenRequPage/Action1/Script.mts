'清除保养的相关历史数据(保养通知和保养时间设定回复)
RunAction "Action1 [CleanHistoryData_BaoYangNotice&SetReply]", oneIteration
'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'检查保养要求页面(通知类型=无，剩余时间的默认值检查)
RunAction "Action1 [Check_BaoYangNoticeMsg_RequPage]", oneIteration,0
'保养剩余时间设置
RunAction "Action1 [SetBaoYangTime]", oneIteration
'检查保养时间设置源码(完全同CheckBaoYangTimeSetSour,只是涉及的Datatable参数命名有区别而已)
RunAction "Action1 [CheckBaoYangTimeSetSour_BaoYangNotice]", oneIteration
For k=1  to 5
	'插入源码(保养剩余时间设置回复指令)---1
	'插入源码(保养通知-预告)---2
	'插入源码(保养通知-警报)---3
	'插入源码(保养通知指令警报或保养指令，其中到下次更换为止的剩余时间为负)---4
	'插入保养通知信息(通知类型=复位)---5
	RunAction "Action1 [InsertBaoYangNotice&TimeSetReplySource]", oneIteration,k
	'检查保养要求页面(只更新剩余时间)---1
	'检查保养要求页面(通知类型=预告)---2
	'检查保养要求页面(通知类型=预警)---3
	'检查保养要求页面(通知类型=经过)---4
	'检查保养要求页面(通知类型为空)---5
	RunAction "Action1 [Check_BaoYangNoticeMsg_RequPage]", oneIteration,k
Next
'系统退出
RunAction "Action1 [Logout]", oneIteration