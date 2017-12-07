On error resume next
'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'定义FSO对象，获取相对路径
Set fso=createobject("Scripting.filesystemobject")
RelaPath = PathFinder.Locate("Sumitomo") &"\DownFiles\"
'点击车辆信息页-最新提交
if(Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Exist)then
Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Click
end if
'点击车辆信息-最新提交-查看锁车设置/回复信息
if(Browser("住友").Page("主页_车辆信息").Link("最新提交_下载解车密码").Exist)then
Browser("住友").Page("主页_车辆信息").Link("最新提交_下载解车密码").Click
end if
'========检查是否正常进入“下载解车密码”页
Dim PosiDownUnlockPwdPage
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebElement("您的位置：>>车辆信息>>下载解车密码").Exist)then
	PosiDownUnlockPwdPage=Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebElement("您的位置：>>车辆信息>>下载解车密码").GetROProperty("innertext")
	if(trim(PosiDownUnlockPwdPage)=Datatable("PosiDownUnlockPwdPage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入下载解车密码页成功","期望值："&Datatable("PosiDownUnlockPwdPage","Global")&" 实际值："& PosiDownUnlockPwdPage
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入下载解车密码页失败","期望值："&Datatable("PosiDownUnlockPwdPage","Global")&" 实际值："& PosiDownUnlockPwdPage
	end if
end if
'========检查是否是对应的车
if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebElement("【设备号】下载解车密码").Exist)then
	Dim VclMsg,ExpVclMsg
	VclMsg=trim(Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebElement("【设备号】下载解车密码").GetROProperty("outertext"))
	ExpVclMsg="【"&Datatable("Vcl_No","Global")&"】下载解车密码"
	if(VclMsg=ExpVclMsg)then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"下载解车密码-设备号检查","期望值: "& ExpVclMsg&" 实际值: "& VclMsg
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"下载解车密码-设备号检查","期望值: "& ExpVclMsg&" 实际值: "& VclMsg
	end if
end if
'========不同的迭代测试，在多种锁的情况下，需要拆分锁类型进行处理
Dim arr
arr=split(Datatable("LockType","Global"),"/")
For i=0 to ubound(arr)
	'========选择相应的锁车类型
	Select Case arr(i)
		Case "总工作时间锁":
			Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebRadioGroup("解车密码按钮组").Select  "1"
		Case "指定日期锁":
			Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebRadioGroup("解车密码按钮组").Select  "2"
		Case "指定位置锁":
			Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebRadioGroup("解车密码按钮组").Select  "3"
		Case "循环日期锁":
			Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebRadioGroup("解车密码按钮组").Select  "4"
			if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebList("循环日期锁_年").Exist)then
			Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebList("循环日期锁_年").Select  Datatable("CircDateLock_Y","Global")
			end if
			if(Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebList("循环日期锁_月").Exist)then
			Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebList("循环日期锁_月").Select  Datatable("CircDateLock_M","Global")
			end if
		Case "立即锁":
			Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebRadioGroup("解车密码按钮组").Select  "0"
	End Select
	'========点击生成预置解车密码按钮
	'========生成密码跟源码中的解车密码基数字节有关，该字节每次是随机生成，目前拼接的锁车回复的源码中此字节是写死的，因为无论生成的解车密码正确与否，无法实现自动化验证
	if( Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebButton("生成预置解车密码").Exist)then
	Browser("住友").Page("主页_车辆信息").Frame("最新提交_下载解车密码").WebButton("生成预置解车密码").Click
	end if
	wait 2 '等待下载框
	'下载txt密码文档
	if(Dialog("文件下载").WinButton("保存(S)").Exist)then
	Dialog("文件下载").WinButton("保存(S)").Click
	end if
	'输入文件名
	if(Dialog("已完成安装-进度").Dialog("另存为").WinEdit("文件名(N)").Exist)then
	TempTime=right("0"& hour(time),2)&right("0"&minute(time),2)&right("0"&second(time),2)
	Datatable("ExcelAddr","Global")=RelaPath &Environment("TestName")&"_"& TempTime &"_"&Dialog("已完成安装-进度").Dialog("另存为").WinEdit("文件名(N)").GetROProperty("text")
	Dialog("已完成安装-进度").Dialog("另存为").WinEdit("文件名(N)").Set  Datatable("ExcelAddr","Global")
	end if
	'保存excel
	if(Dialog("已完成安装-进度").Dialog("另存为").WinButton("保存(S)").Exist)then
	Dialog("已完成安装-进度").Dialog("另存为").WinButton("保存(S)").Click
	end if
	'========检查下载的excel是否存在，即是否下载成功
	wait 1
	if(fso.FileExists(Datatable("ExcelAddr","Global")))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车密码文件下载成功","解车密码文件下载成功"
	else
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车密码文件下载失败","解车密码文件下载失败"
	end if
Next
'释放FSO对象
Set fso=nothing
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If
