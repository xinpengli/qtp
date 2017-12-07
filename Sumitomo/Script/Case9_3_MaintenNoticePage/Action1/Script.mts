'清除保养的相关历史数据(保养通知和保养时间设定回复)
RunAction "Action1 [CleanHistoryData_BaoYangNotice&SetReply]", oneIteration
'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'插入保养通知信息(预告、预警、经过、复位)
For i=1 to 4
	'插入保养通知信息(预告、预警、经过、复位)
	RunAction "Action1 [InsertBaoYangNoticeSource_All]", oneIteration,i
	'检查保养通知页面
	RunAction "Action1 [Check_BaoYangNoticeMsg_NoticePage]", oneIteration,i
Next
'系统退出
RunAction "Action1 [Logout]", oneIteration