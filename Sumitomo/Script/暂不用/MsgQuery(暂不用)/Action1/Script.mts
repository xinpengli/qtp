'动态加载对象库,关注相对路径的问题
'RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'========点击车辆信息页-最新提交
if(Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Exist)then
Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Click
end if
'========点击车辆信息-最新提交-查看锁车设置/回复信息
if(Browser("住友").Page("主页_车辆信息").Link("信息查询_").Exist)then
Browser("住友").Page("主页_车辆信息").Link("信息查询_").Click
end if
'========检查是否正常进入“查看锁车设置/回复信息”页
Dim PosiVehiMsgPage
if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebElement("您的位置>>车辆信息>>信息查询").Exist)then
	PosiMsgQueryPage=Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebElement("您的位置>>车辆信息>>信息查询").GetROProperty("innertext")
	if(trim(PosiMsgQueryPage)=Datatable("PosiMsgQueryPage","Global"))then
	reporter.ReportEvent micPass,"进入车辆信息-信息查询页成功","期望值："&Datatable("PosiMsgQueryPage","Global")&" 实际值："& PosiMsgQueryPage
	else
	reporter.ReportEvent micFail,"进入车辆信息-信息查询页失败","期望值："&Datatable("PosiMsgQueryPage","Global")&" 实际值："& PosiMsgQueryPage
	end if
end if
'========输入时间段
Dim currDay '获取当天日期
currDay=Cstr(Year(Date)&"-"&right("0"&Month(Date),2)&"-"&right("0"&Day(Date),2))
if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebEdit("开始时间").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebEdit("开始时间").Object.value=currDay
end if
if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebEdit("结束时间").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebEdit("结束时间").Object.value=currDay
end if
'========查询信息
if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebButton("查询").Exist)then
Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebButton("查询").Click
end if
Browser("住友").Page("主页_车辆信息").Sync '等待加载
'设置预期值存放的action与当前执行的Global行数对应
Datatable.GetSheet("Action1").SetCurrentRow(Datatable.GetSheet("Global").GetCurrentRow)
'判断webtable是否存在
if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebTable("信息查询列表").Exist)then
	Set wt=Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebTable("信息查询列表")
	wait 2
	Dim ColVal,ExpColVal
	For i=1 to wt.ColumnCount(1)
		ColVal=trim(wt.GetCellData(2,i))
		ExpColVal=Datatable.GetSheet("Action1").GetParameter(i).Value
		if(ColVal=ExpColVal)then
		reporter.ReportEvent micPass,Datatable("MsgType","Global")&i&"列检查通过","期望值: "& ExpColVal &"实际值: "& ColVal
		else
		end if
	Next
	'下载文件 
	if(Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebButton("下载").Exist)then
	Browser("住友").Page("主页_车辆信息").Frame("信息查询").WebButton("下载").Click
	end if

RunAction "Action1 [DownFile]", oneIteration
end if
