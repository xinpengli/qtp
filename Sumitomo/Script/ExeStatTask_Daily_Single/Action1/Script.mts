On error resume next
'执行重写Reporter的vbs,重新实例化Reporter
executefile  "..\..\Sumitomo\Func&VBS\Reporter.vbs"
Dim Reporter
Set Reporter= GetReporter()
'========定义并获取Wscript对象
Dim WshShell
Set WshShell = CreateObject("wscript.Shell")
'========存储本机日期
Datatable("StanSysDate","Global")=date
'========修改本机日期,即exe的机器执行时间
'========因定时表默认创建到date+1,如机器执行时间为date+2,则exe执行时报错: 找不到总工作时间,休眠....
Select Case Parameter("i")
	Case "1":
		Datatable("StatTaskExeDate","Global")=Cstr(date-1)
	Case "2":
		Datatable("StatTaskExeDate","Global")=Cstr(date)
	Case "3":
		Datatable("StatTaskExeDate","Global")=Cstr(date+1)
End Select
WshShell.Run "cmd.exe /c date "&Datatable("StatTaskExeDate","Global")&"", 0
wait 2  '执行完修改时间操作后加等待,避免脚本走的快,取到的非最新时间
'========获取exe程序所在路径,通过FSO或WScript获取的相对路径问题变化,不知何因
Dim TestPath
TestPath=Environment("TestDir")
TestPath=Replace(TestPath,right(TestPath,len(TestPath)-instrRev(TestPath,"\")+7), "StatProgram\StatData_Daily_zuixin") '+7等价于截取的是倒数第二个斜杠位置之后的内容
'========执行exe程序，且等待3秒后再执行恢复本机日期的操作
TestPath="cmd /c cd "&TestPath&" && StatData_Daily.exe"   '可成功执行
WshShell.Run TestPath,1,true
wait 3 
'========判断Exe是否执行完毕，如完毕则恢复本机日期
set rd=getobject("winmgmts:\\.")    '  ":\\" 选择计算机地址      "."指本地计算机
set sysProcess=rd.instancesof("win32_process")      ' "instancesof("win32_process")"系统进程
Dim flag 'empty表示要找的进程不存在，true表示要找的进程存在
flag=empty
For each r in sysProcess	
	If r.name="cmd.exe" then     '  ".name"单个进程的名称
'		msgbox r.name
		flag=true
	else
	end if
next
if(flag)then
WshShell.Run "cmd.exe /c date "&Datatable("StanSysDate","Global")&"", 0
wait 2  '执行完修改时间操作后加等待,避免脚本走的快,取到的非最新时间，比如写报告时取的非最系统标准时间
reporter.ReportEvent micWarning,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"cmd.exe未正常执行完毕,恢复系统时间","cmd.exe未正常执行完毕,恢复系统时间"
else
WshShell.Run "cmd.exe /c date "&Datatable("StanSysDate","Global")&"", 0
wait 2  '执行完修改时间操作后加等待,避免脚本走的快,取到的非最新时间，比如写报告时取的非最系统标准时间
reporter.ReportEvent micPass,Environment("TestName")&"?"&Datatable.GetSheet("Global").GetCurrentRow&"?"&"cmd.exe正常执行完毕,恢复系统时间成功","cmd.exe正常执行完毕,恢复系统时间成功"
end if
'========释放Wsh对象,关闭cmd进程(通过exit退出cmd后进程仍存在,故关进程 )
Set WshShell=nothing
'关闭cmd相关进程
systemutil.CloseProcessByName("cmd.exe")
'记录err
If err.number<>0 Then
	   testName=environment("TestName")
	   versionNo=datatable("VersionNo","Global")
	   actionName=environment("ActionName")
	   currRow=cstr(datatable.GetSheet("Global").GetCurrentRow)
	   rowCount=cstr(datatable.GetSheet("Global").GetRowCount)
       Reporter.XmlDomDoc_ErrLog testName,versionNo,actionName,currRow,rowCount,Cstr(err.number),err.description,err.source,cstr(now())
End If
