On error resume next
'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'获取相对路径，用于拼excel下载地址，且用于检查excel文件是否存在
RelaPath=pathFinder.Locate("DownFiles") &"\"
'========下载功能
'执行下载操作
if(Dialog("文件下载").WinButton("保存(S)").Exist)then
Dialog("文件下载").WinButton("保存(S)").Click
end if
'输入文件名
if(Dialog("已完成安装-进度").Dialog("另存为").WinEdit("文件名(N)").Exist)then
    TempTime=right("0"& hour(time),2)&right("0"&minute(time),2)&right("0"&second(time),2)
	'锁解车类
	if(Environment("TestName")="Case1_VehiMsg_LastSubm_LockAndUnlock" or Environment("TestName")="Case2_VehiMsg_LastSubm_DownUnlockPwd")then
		if(Datatable("LockType","Global")="总工作时间锁/指定日期锁/指定位置锁/循环日期锁/立即锁")then
		Datatable("ExcelAddr","Global")=RelaPath &Environment("TestName")&"_"&"多种锁车"&"_"&Datatable("LockUnlockFlag","Global")&Datatable("LockUnlockReplyFlag","Global") &TempTime&".xls"
		else
		Datatable("ExcelAddr","Global")=RelaPath &Environment("TestName")&"_"&Datatable("LockType","Global")&"_"&Datatable("LockUnlockFlag","Global")&Datatable("LockUnlockReplyFlag","Global") & TempTime &".xls"
		end if
	end if
	'强制锁解车类
	if(left(Environment("TestName"),6)="Case5_")then
	Datatable("ExcelAddr","Global")=RelaPath &Environment("TestName")&"_"&Datatable("LockUnlockFlag","Global")& TempTime &".xls"
	end if
	'保养类
	if(Environment("TestName")="Case3_VehiMsg_LastSubm_BaoYangTimeSetAndReply")then
	Datatable("ExcelAddr","Global")=RelaPath &Environment("TestName")&"_"&Datatable("BaoYangSetReplyFlag","Global")&TempTime &".xls"
	end if
	'信息查询类
	if(mid(Environment("TestName"),8,18)="_VehiMsg_MsgQuery_" or mid(Environment("TestName"),9,18)="_VehiMsg_MsgQuery_")then
	Datatable("ExcelAddr","Global")=RelaPath &Environment("TestName")&"_"&Datatable("InputMsgType","Global")&TempTime &".xls"
	end if
	'机器档案--点检实绩&所有权设定
	if(Environment("TestName")="Case10_VclMsg_MachineFile")then
	Datatable("ExcelAddr","Global")=RelaPath &Environment("TestName")&"_"&TempTime &".xls"
	end if
	Dialog("已完成安装-进度").Dialog("另存为").WinEdit("文件名(N)").Set  Datatable("ExcelAddr","Global")
end if
'保存excel
if(Dialog("已完成安装-进度").Dialog("另存为").WinButton("保存(S)").Exist)then
Dialog("已完成安装-进度").Dialog("另存为").WinButton("保存(S)").Click
end if
'检查下载的excel是否存在，即是否下载成功
wait 1
'fso定义用于后续判断excel文件是否存在
Set fso=createobject("Scripting.filesystemobject")
if(fso.FileExists(Datatable("ExcelAddr","Global")))then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"excel下载成功","excel下载成功"
else
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"excel下载失败","excel下载失败"
end if
'释放FSO对象
Set fso=nothing
'========检查excel内容和页面datatable数据是否一致
Dim wt '定义webtable对象
Dim ColCount '定义表格的使用范围列数
' 创建Excel应用程序对象        
Set oExcel = CreateObject("Excel.Application")              
' 打开Excel文件        
oExcel.Workbooks.Open(Datatable("ExcelAddr","Global"))        
' 获取表格的使用范围列数
ColCount=oExcel.Worksheets(1).UsedRange.columns.count
'Case1 车辆信息->最新提交->锁解车用例，检查Excel下载内容
'Case2 车辆信息->最新提交->下载解车密码，不检查Excel下载内容
'Case3  车辆信息->最新提交->保养时间设定和回复用例，检查Excel下载内容
'Case4 车辆信息->信息查询 相关用例，检查Excel下载内容
'Case5  车辆信息->最新提交->锁解车用例(区分上下级权限)，不检查Excel下载内容
if(Environment("TestName")="Case1_VehiMsg_LastSubm_LockAndUnlock" or Environment("TestName")="Case3_VehiMsg_LastSubm_BaoYangTimeSetAndReply" or  mid(Environment("TestName"),8,18)="_VehiMsg_MsgQuery_" or mid(Environment("TestName"),9,18)="_VehiMsg_MsgQuery_")then
	'获取页面datatable数据
	if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebTable("锁车设置/回复信息列表").Exist)then
		Set wt=Browser("住友").Page("主页_车辆信息").Frame("最新提交_查看锁车设置/回复信息").WebTable("锁车设置/回复信息列表")
	end if
	'页面数据和excel数据作比较，根据页面上的每一列列名去excel中循环匹配，如匹配上则进行检查，除经纬度屏蔽外，excel中是否还缺失列还是需要人工检查
	For i=1 to wt.ColumnCount(1)  'webtable列循环
		For j=1 to ColCount   'excel列循环
			if(trim(oExcel.Worksheets(1).cells(1,j))=trim(wt.GetCellData(1,i)))then
				'比较列值
				if(trim(oExcel.Worksheets(1).cells(2,j))=trim(wt.GetCellData(2,i)))then
				reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"Excel列-"& oExcel.Worksheets(1).cells(1,j)&"-检查通过","期望值："&wt.GetCellData(2,i)&" 实际值："& trim(oExcel.Worksheets(1).cells(2,j))
				else
				reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"Excel列-"& oExcel.Worksheets(1).cells(1,j)&"-检查失败","期望值："&wt.GetCellData(2,i)&" 实际值："& trim(oExcel.Worksheets(1).cells(2,j))
				end if
				Exit For '如匹配上列，则比较后退出当前excel列循环
			end if
		Next
	Next
end if
' 关闭工作簿        
oExcel.WorkBooks.Item(1).Close        
' 退出Excel        
oExcel.Quit        
Set oExcel = Nothing
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If







