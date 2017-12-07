'清除历史数据(保养时间设定和回复)
RunAction "Action1 [CleanHistoryData_BaoYangTimeSet&Reply]", oneIteration
'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'设置保养剩余时间(保养时间设定)
RunAction "Action1 [SetBaoYangTime]", oneIteration
'检查保养剩余时间设置的源码
RunAction "Action1 [CheckBaoYangTimeSetSour]", oneIteration
'查看保养剩余时间设置/回复(保养时间设定)
RunAction "Action1 [CheckBaoYangTimeSetAndReply]", oneIteration
'插入保养时间设定回复源码
RunAction "Action1 [InsertBaoYangTimeSetReplySource]", oneIteration
'查看保养剩余时间设置/回复(保养时间设定回复)
RunAction "Action1 [CheckBaoYangTimeSetAndReply]", oneIteration
'系统退出
RunAction "Action1 [Logout]", oneIteration
