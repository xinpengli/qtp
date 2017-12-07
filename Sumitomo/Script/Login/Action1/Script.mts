'动态加载对象库,关注相对路径的问题
RepositoriesCollection.Add "..\..\Sumitomo\ObjectRepository\Sumitomo.tsr"
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'清除原qtp打开的页面
systemutil.CloseDescendentProcesses
'打开住友Url
systemutil.Run  "iexplore.exe",Datatable("Url","Global")
'输入用户名密码,登录
if(Browser("住友").Page("登录页").WebEdit("帐户").Exist)then
    '如果设置过锁车了,则表示二次登陆,使用帐户2登陆,否则使用原帐户登陆
	if(Datatable("LockUnlockFlag","Global")="锁车")then
	Browser("住友").Page("登录页").WebEdit("帐户").Set Datatable("Account_2","Global")
	else
	Browser("住友").Page("登录页").WebEdit("帐户").Set Datatable("Account","Global")
	end if
end if
if(Browser("住友").Page("登录页").WebEdit("密码").Exist)then
Browser("住友").Page("登录页").WebEdit("密码").Set Datatable("Pwd","Global")
end if
if(Browser("住友").Page("登录页").WebButton("登录").Exist)then
Browser("住友").Page("登录页").WebButton("登录").Click
end if
'等待主页加载
Browser("住友").Page("主页").Sync
'判断主页跳转是否成功,通过判断主页展示的用户登陆名是否正确
Dim Account
Account=Browser("住友").Page("主页").Frame("HeadFrame").WebElement("登录用户名").GetROProperty("innertext")
'如果设置过锁车了,则表示二次登陆,检查帐户2,否则检查原帐户
if(Datatable("LockUnlockFlag","Global")="锁车")then
	if(trim(Account)=Datatable("Account_2","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"二次登录成功","期望登录用户名: "&Datatable("Account_2","Global")&"实际登录用户名是: "&Account
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"二次登录失败","期望登录用户名: "&Datatable("Account_2","Global")&" 实际登录用户名是: "&Account
	end if 
else
	if(trim(Account)=Datatable("Account","Global"))then
	reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"登录成功","期望登录用户名: "&Datatable("Account","Global")&"实际登录用户名是: "&Account
	else
	reporter.ReportEvent micFail,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"登录失败","期望登录用户名: "&Datatable("Account","Global")&" 实际登录用户名是: "&Account
	end if 
end if
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If
