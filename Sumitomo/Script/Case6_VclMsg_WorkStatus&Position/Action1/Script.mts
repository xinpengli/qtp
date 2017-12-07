'========注：Datatable Action1中的测试数据存放如VehicleInfoTreeView_item_2_3
'========VehicleInfoTreeView_item_2对应的是web页面元素中的节点名
'========_3对应的传递到Action CheckVclMsgWorkStatus&Position中的参数

'清除历史数据(定时及位置)
RunAction "Action1 [CleanHistoryData_Timing&Position]", oneIteration
'清除历史数据(锁解车及报告)
RunAction "Action1 [CleanHistoryData_LockUnlock]", oneIteration
'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'检查车辆信息的工作状态/位置-1
RunAction "Action1 [CheckVclMsgWorkStatus&Position]", oneIteration,1

'锁车设置
RunAction "Action1 [LockSet]", oneIteration
'检查锁解车设置源码
RunAction "Action1 [CheckLockUnlockSetSour]", oneIteration
'插入锁车设置回复源码(不更新锁车状态，更新锁车设置情况)
RunAction "Action1 [InsertLockSetReplySource]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'检查车辆信息的工作状态/位置-2
RunAction "Action1 [CheckVclMsgWorkStatus&Position]", oneIteration,2

'插入锁解车报告源码(锁车报告，更新锁解车状态，且锁车设置情况清空)
RunAction "Action1 [InsertLockReportSource]", oneIteration
'查看锁解车报告(锁车报告检查)
RunAction "Action1 [CheckLockUnlockReport]", oneIteration
'进入车辆信息页--目的是刷新框架页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'检查车辆信息的工作状态/位置-3
RunAction "Action1 [CheckVclMsgWorkStatus&Position]", oneIteration,3
		
'insert定时源码，源码中锁解车状态为：按“GPS天线防拆锁”锁车，不会更新锁车设置情况
RunAction "Action1 [InsertPublicSource]", oneIteration
'进入车辆信息页--目的是刷新框架页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'检查车辆信息的工作状态/位置-4
RunAction "Action1 [CheckVclMsgWorkStatus&Position]", oneIteration,4

'解车设置
RunAction "Action1 [UnlockSet]", oneIteration
'检查锁解车设置源码
RunAction "Action1 [CheckLockUnlockSetSour]", oneIteration
'插入解车设置回复源码
RunAction "Action1 [InsertUnLockSetReplySource]", oneIteration
'进入车辆信息页--目的是刷新框架页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'检查车辆信息的工作状态/位置-5
RunAction "Action1 [CheckVclMsgWorkStatus&Position]", oneIteration,5
'系统退出
'RunAction "Action1 [Logout]", oneIteration