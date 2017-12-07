On error resume next
'加载测试数据--调试用
'datatable.ImportSheet "..\..\Sumitomo\TestData\UnlockSet.xls",1,"Global"
'========动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'========执行DB交互的函数文件
executefile  "..\..\Sumitomo\Func&VBS\DBFunc.txt"
'========执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'========更新上一条锁车记录的设置时间，避免系统提示：该机器在5分钟内已有过操作请求，请在5分钟后提交您的操作请求!
''''定义字典对象，存储查询函数返回值
'''Set oDict = CreateObject("Scripting.Dictionary") 
''''定义sql，查询最近的一条锁车记录，取出MsgL_ID,MsgL_SetTime
'''Dim sqlQuery
'''sqlQuery="SELECT TOP 1 MsgL_ID,MsgL_SetTime  FROM Sumitomo.dbo.Msg_Lock where MsgL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"+Datatable("Vcl_No","Global")+"') order by MsgL_ID desc"
''''执行查询函数
'''oDict=QueryDBColumn_Dict(sqlQuery,"MsgL_ID","MsgL_SetTime")
'''wait 1
''' '将时间设置字段减5分钟存储，在锁解车设置/回复查询页，展示为提交时间
'''Datatable("SetTime","Global")=dateadd("N",-5,oDict("MsgL_SetTime"))
'定义并执行数据库更新语句
 Dim sqlUpd
 'sqlUpd="update [Sumitomo].dbo.[Msg_Lock] set msgl_settime='"+Cstr(Datatable("SetTime","Global"))+"' where MsgL_ID='"+Cstr(oDict("MsgL_ID"))+"'"
 '转成如下sql可避免调用QueryDBColumn_Dict,但是sql比较繁琐
 sqlUpd="update Sumitomo.dbo.Msg_Lock set MsgL_SetTime=DATEADD(N,-5,MsgL_SetTime) where MsgL_ID=(SELECT TOP 1 MsgL_ID FROM Sumitomo.dbo.Msg_Lock where MsgL_Vcl_ID=(select Vcl_ID from Sumitomo.dbo.VclInfo where Vcl_No='"&Datatable("Vcl_No","Global")&"') order by MsgL_ID desc)"
 '执行sql并返回结果
Dim RetuVal
RetuVal=ExecDB(sqlUpd) 
'根据执行结果写日志
if(RetuVal>=0)then
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"更新锁解车设置时间成功","更新锁解车设置时间成功"
else
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"更新锁解车设置时间失败","更新锁解车设置时间失败"
end if
 wait 1
'========点击车辆信息-最新提交
if(Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Exist)then
Browser("住友").Page("主页_车辆信息").WebElement("最新提交_").Click
end if
'========点击车辆信息-锁解车设置
if(Browser("住友").Page("主页_车辆信息").Link("最新提交_锁/解车设置").Exist)then
Browser("住友").Page("主页_车辆信息").Link("最新提交_锁/解车设置").Click
end if
'========判断进入"锁解车设置"页
if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebElement("您的位置：>> 车辆信息>>锁/解车设置").Exist)then
	PosiLockUnlockPage=trim(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebElement("您的位置：>> 车辆信息>>锁/解车设置").GetROProperty("innertext"))
	if(PosiLockUnlockPage=Datatable("PosiLockUnlockPage","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入锁解车设置页成功","期望值: "&Datatable("PosiLockUnlockPage","Global")&" 实际值: "&PosiLockUnlockPage
	else
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"进入锁解车设置页失败","期望值: "&Datatable("PosiLockUnlockPage","Global")&" 实际值: "&PosiLockUnlockPage
	end if
end if
'========选择锁车或解车
Dim SeleUnlock
SeleUnlock=Datatable("SeleUnlock","Global")
Select Case SeleUnlock
	Case "解车":
	Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebRadioGroup("锁解车按钮").Select "4"
	Case "全解车":
	Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebRadioGroup("锁解车按钮").Select "5"
End Select
'========选择锁车或解车类型--适用多种锁车类型，即拆分为数组循环执行
Dim UnlockType
UnlockType=Datatable("UnlockType","Global")
Dim arr
arr=Split(UnlockType,"/")
For i=0 to Ubound(arr)
	Select Case arr(i)
		Case "总工作时间锁":
			'勾选总工作时间锁
			if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("总工作时间锁").Exist)then
			Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("总工作时间锁").Set "on"
			end if
		Case "指定日期锁":
			'勾选指定日期锁
			if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("指定日期锁").Exist)then
			Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("指定日期锁").Set "on"
			end if
		'位置锁目前有bug,待开发修改程序???
		Case "指定位置锁":
			'勾选指定位置锁
			if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("指定位置锁").Exist)then
			Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("指定位置锁").Set "on"
			end if
		Case "循环日期锁":
			'勾选循环日期锁
			if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("循环日期锁").Exist)then
			Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("循环日期锁").Set "on"
			end if
			'如果是解车，需要设置是全部解车还是某个锁车类型解车
			if(Datatable("SeleUnlock","Global")="解车")then
				if(Datatable("CircDateUnlock_All","Global")="全部")then
					if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebRadioGroup("循环日期解锁_类型").Exist)then
					Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebRadioGroup("循环日期解锁_类型").Select  "rbULAllPwd"
					end if
				else
					if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebRadioGroup("循环日期解锁_类型").Exist)then
					Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebRadioGroup("循环日期解锁_类型").Select  "rbULDatePwd"
					end if
					if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebList("循环日期解锁_年").Exist)then
					Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebList("循环日期解锁_年").Select Datatable("CircDateLock_Y","Global")
					end if
					if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebList("循环日期解锁_月").Exist)then
					Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebList("循环日期解锁_月").Select Datatable("CircDateLock_M","Global")
					end if
				end if
			end if
		Case "立即锁":
			'勾选立即锁
			if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("立即锁").Exist)then
			Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("立即锁").Set "on"
			end if
	End Select
Next
'========输入锁解车用户名密码,即当前登陆用户名密码,点击提交
if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("账号").Exist)then
	if(Datatable("LockUnlockFlag","Global")="锁车")then
	Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("账号").Set Datatable("Account_2","Global")
	else
	Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("账号").Set Datatable("Account","Global")
	end if
end if
if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("密码").Exist)then
Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("密码").Set Datatable("Pwd","Global")
end if
if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebButton("提交").Exist)then
Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebButton("提交").Click
end if
'========点击确认框
if(Datatable("SeleUnlock","Global")="解车" or Datatable("SeleUnlock","Global")="全解车")then
    '如果选择解车或全解车，均提示“解车设置成功”
	if(Browser("住友_锁解车设置").Dialog("来自网页的消息").static("text:=解车设置成功！").Exist(5))then
	    reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置成功","解车设置成功"
		Browser("住友_锁解车设置").Dialog("来自网页的消息").WinButton("确定").Click
	else
	    '目前未见过解车设置失败的提示语
		if(Browser("住友_锁解车设置").Dialog("来自网页的消息").static("text:=解车设置失败！").Exist(5))then
			reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"解车设置失败","解车设置失败"
			Browser("住友_锁解车设置").Dialog("来自网页的消息").WinButton("确定").Click
		else
			If(Browser("住友_锁解车设置").Dialog("来自网页的消息").static("text:=您无权设置，"&Datatable("Account","Global")&"设置过"&Datatable("LockType","Global")&" ，请确认后再设置。").Exist(5))then
			reporter.ReportEvent micWarning,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"无权设置解车","您无权设置，"&Datatable("Account","Global")&"设置过"&Datatable("LockType","Global")&" ，请确认后再设置。"
			Browser("住友_锁解车设置").Dialog("来自网页的消息").WinButton("确定").Click
			end if
		end if
	end if
end if
'========关闭叫数据页面
Browser("住友_锁解车设置").Close
'========设置锁解车标志位，用于解车设置页面的检查
Datatable("LockUnlockFlag","Global")=Datatable("SeleUnlock","Global")
'========记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If

