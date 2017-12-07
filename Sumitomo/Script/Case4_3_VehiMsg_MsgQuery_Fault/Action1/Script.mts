'======因故障的信息生成时间取的是定时的信息生成时间,故在执行4_3之前先执行下4_1用例,以便存储到本地的信息生成时间是最新的
'清除历史数据
RunAction "Action1 [CleanHistoryData]", oneIteration
'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'插入信息查询类源码
RunAction "Action1 [InsertMsgQuerySource]", oneIteration
'信息查询检查
RunAction "Action1 [Check_MsgQuery_Fault]", oneIteration
'系统退出
RunAction "Action1 [Logout]", oneIteration