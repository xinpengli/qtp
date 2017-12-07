'========基于定时信息的用例Case4_1_VehiMsg_MsgQuery_Timing之后，需要使用跟定时相同的信息生成时间========
'清除历史数据
RunAction "Action1 [CleanHistoryData]", oneIteration
'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'插入信息查询类源码
RunAction "Action1 [InsertMsgQuerySource]", oneIteration
'信息查询检查
RunAction "Action1 [Check_MsgQuery_BattIsRemoAlarm]", oneIteration
'系统退出
RunAction "Action1 [Logout]", oneIteration