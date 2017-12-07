'系统登录
RunAction "Action1 [Login]", oneIteration
'进入车辆信息页
RunAction "Action1 [IntoVehiInfoFramePage]", oneIteration
'查询信息类最大ReceivalID(终端上发中心)
RunAction "Action1 [QueryMsgMaxReceivalID]", oneIteration
'插入信息查询类源码
RunAction "Action1 [InsertMsgQuerySource]", oneIteration
'信息查询
'RunAction "Action1 [MsgQuery]", oneIteration
''系统退出
'RunAction "Action1 [Logout]", oneIteration