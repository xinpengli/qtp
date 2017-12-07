'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'查询锁解车设置类x81源码表的最大MsgID
RunAction "Action1 [QueryLockUnlockSetSourMsgID]", oneIteration
'锁车设置
RunAction "Action1 [LockSet]", oneIteration
'锁车设置源码检查，只检查参数段
RunAction "Action1 [CheckLockUnlockSetSour]", oneIteration
'查看锁车设置信息
RunAction "Action1 [CheckLockUnlockSet]", oneIteration
'插入锁车设置回复源码
RunAction "Action1 [InsertLockSetReplySource]", oneIteration
'查看锁车设置回复信息
RunAction "Action1 [CheckLockUnlockSetReply]", oneIteration
'插入锁解车报告源码(锁车报告)
RunAction "Action1 [InsertLockReportSource]", oneIteration
'查看锁解车报告(锁车报告检查)
RunAction "Action1 [CheckLockUnlockReport]", oneIteration
'查询锁解车设置类x81源码表的最大MsgID
RunAction "Action1 [QueryLockUnlockSetSourMsgID]", oneIteration
'解车设置
RunAction "Action1 [UnlockSet]", oneIteration
'解车设置源码检查，只检查参数段
RunAction "Action1 [CheckLockUnlockSetSour]", oneIteration
'查看解车设置信息
RunAction "Action1 [CheckLockUnlockSet]", oneIteration
'插入解车设置回复源码
RunAction "Action1 [InsertUnLockSetReplySource]", oneIteration
'查看解车设置回复信息
RunAction "Action1 [CheckLockUnlockSetReply]", oneIteration
'系统退出
RunAction "Action1 [Logout]", oneIteration