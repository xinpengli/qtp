On error resume next
'加载测试数据--调试用
'datatable.ImportSheet "..\..\Sumitomo\TestData\LockSet.xls",1,"Global"
'========动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'========执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
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
'========选择锁车
if(Datatable("SeleLock","Global")="锁车")then
	Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebRadioGroup("锁解车按钮").Select "1"
end if
'========选择锁车或解车类型--适用多种锁车类型，即拆分为数组循环执行
Dim LockType
LockType=Datatable("LockType","Global")
Dim arr
arr=Split(LockType,"/")
For i=0 to Ubound(arr)
	Select Case arr(i)
		Case "总工作时间锁":
			'勾选总工作时间锁
			if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("总工作时间锁").Exist)then
			Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("总工作时间锁").Set "on"
			end if
			'如果是锁车，需要设置总工作小时数
			if(Datatable("SeleLock","Global")="锁车")then
				'设置总工作时间小时数
				if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("总工作时间锁_小时").Exist)then
				Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("总工作时间锁_小时").Set Datatable("WorkTimeLock_Hour","Global")
				end if
			end if
		Case "指定日期锁":
			'勾选指定日期锁
			if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("指定日期锁").Exist)then
			Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("指定日期锁").Set "on"
			end if
			'如果是锁车，需要设置指定日期
			if(Datatable("SeleLock","Global")="锁车")then
				'设置指定日期
				if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("指定日期锁_锁定日期").Exist)then		
'				CurrDate=DateAdd("D",1,Date) '当前日期+1天
				CurrDate=Date
				CurrDate=Cstr(Year(CurrDate) &"-"& right("0"&Month(CurrDate),2) &"-"& right("0"&Day(CurrDate),2)) '转换日期格式
				Datatable("AppDateLock_Date","Global")=CurrDate
				Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("指定日期锁_锁定日期").Object.value=Datatable("AppDateLock_Date","Global")
				end if
			end if
		'位置锁目前有bug,待开发修改程序???
		Case "指定位置锁":
			'勾选指定位置锁
			if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("指定位置锁").Exist)then
			Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("指定位置锁").Set "on"
			end if
			'如果是锁车，需要设置指定位置
			if(Datatable("SeleLock","Global")="锁车")then
				'输入位置经纬度及半径
				if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("指定位置锁_经度").Exist)then
				Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("指定位置锁_经度").Set Datatable("AppPosiLock_Long","Global")
				end if
				if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("指定位置锁_纬度").Exist)then
				Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("指定位置锁_纬度").Set Datatable("AppPosiLock_Lati","Global")
				end if
				if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("指定位置锁_半径").Exist)then
				Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("指定位置锁_半径").Set Datatable("AppPosiLock_Radi","Global")
				end if
			end if
		Case "循环日期锁":
			'勾选循环日期锁
			if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("循环日期锁").Exist)then
			Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebCheckBox("循环日期锁").Set "on"
			end if
			'如果是锁车，需要设置循环日期密码锁及锁车月份
			if(Datatable("SeleLock","Global")="锁车")then
				'设置循环日期年\月\日\锁车月份
				if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebList("循环日期锁_年").Exist)then
				Datatable("CircDateLock_Y","Global")=Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebList("循环日期锁_年").Object.value
				end if
				if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebList("循环日期锁_月").Exist)then
				Datatable("CircDateLock_M","Global")=Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebList("循环日期锁_月").Object.value
				end if
				if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebList("循环日期锁_日").Exist)then
				Datatable("CircDateLock_D","Global")=Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebList("循环日期锁_日").Object.value
				end if
				if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebList("循环日期锁_锁车月份").Exist)then
				Datatable("CircDateLock_LockM","Global")=Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebList("循环日期锁_锁车月份").Object.value 
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
Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("账号").Set Datatable("Account","Global")
end if
if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("密码").Exist)then
Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebEdit("密码").Set Datatable("Pwd","Global")
end if
if(Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebButton("提交").Exist)then
Browser("住友_锁解车设置").Page("最新提交_锁解车设置").WebButton("提交").Click
end if
'========点击确认框
if(Datatable("SeleLock","Global")="锁车")then
	if(Browser("住友_锁解车设置").Dialog("来自网页的消息").static("text:=锁车设置成功！").Exist)then
	    reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置成功","锁车设置成功"
		Browser("住友_锁解车设置").Dialog("来自网页的消息").WinButton("确定").Click
	else
	    '目前未见过锁车设置失败的提示语
		if(Browser("住友_锁解车设置").Dialog("来自网页的消息").static("text:=锁车设置失败！").Exist)then
			reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"锁车设置失败","锁车设置失败"
			Browser("住友_锁解车设置").Dialog("来自网页的消息").WinButton("确定").Click
		end if
	end if
end if
''========关闭叫数据页面
Browser("住友_锁解车设置").Close
'========设置锁解车标志位，用于锁车设置页面的检查
Datatable("LockUnlockFlag","Global")=Datatable("SeleLock","Global")
