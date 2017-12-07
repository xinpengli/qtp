'清除历史数据(锁解车及报告)
RunAction "Action1 [CleanHistoryData_LockUnlock]", oneIteration
'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration

'锁车设置
RunAction "Action1 [LockSet]", oneIteration
'检查锁解车设置源码
RunAction "Action1 [CheckLockUnlockSetSour]", oneIteration
'查看锁解车设置信息(锁车设置信息)
RunAction "Action1 [CheckLockUnlockSet]", oneIteration

'插入锁车设置回复源码
RunAction "Action1 [InsertLockSetReplySource]", oneIteration
'查看锁解车设置回复信息(锁车设置回复信息)
RunAction "Action1 [CheckLockUnlockSetReply]", oneIteration

'下载解车密码
RunAction "Action1 [DownUnlockPwd]", oneIteration
'系统退出
RunAction "Action1 [Logout]", oneIteration