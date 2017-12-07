'清除历史数据
RunAction "Action1 [CleanHistoryData_LockUnlock]", oneIteration

'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration


'锁车设置
RunAction "Action1 [LockSet]", oneIteration
'锁车源码检查(只检查参数列表)
RunAction "Action1 [CheckLockUnlockSetSour]", oneIteration
'查看锁解车设置信息(锁车设置信息)
RunAction "Action1 [CheckLockUnlockSet]", oneIteration

'插入锁车设置回复源码
RunAction "Action1 [InsertLockSetReplySource]", oneIteration
'查看锁解车设置回复信息(锁车设置回复信息)
RunAction "Action1 [CheckLockUnlockSetReply]", oneIteration

'插入锁解车报告源码(锁车报告)
RunAction "Action1 [InsertLockReportSource]", oneIteration
'查看锁解车报告(锁车报告检查)
RunAction "Action1 [CheckLockUnlockReport]", oneIteration

'解车设置
RunAction "Action1 [UnlockSet]", oneIteration
'解车源码检查(只检查参数列表)
RunAction "Action1 [CheckLockUnlockSetSour]", oneIteration
'查看锁解车设置信息(解车设置信息)
RunAction "Action1 [CheckLockUnlockSet]", oneIteration

'插入解车设置回复源码
RunAction "Action1 [InsertUnLockSetReplySource]", oneIteration
'查看锁解车设置回复信息(解车设置回复信息)
RunAction "Action1 [CheckLockUnlockSetReply]", oneIteration

'系统退出
RunAction "Action1 [Logout]", oneIteration
