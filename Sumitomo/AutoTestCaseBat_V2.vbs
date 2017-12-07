'=====================VBS创建=======================
'创建QTP应用
Dim qtApp
Set qtApp=CreateObject("QuickTest.Application")
'启动QTP应用
qtApp.Launch
'定义qtp显示运行
qtApp.Visible = True
'获取当前VBS的相对路径，即当前执行的vbs文件所在的文件夹的路径，即项目名
Dim VbsPath
Set fso=CreateObject("Scripting.FileSystemObject")
VbsPath=fso.GetFolder(".")
'定义ErrLogRead.vbs返回对象
Dim dictXML
'定义ErrLogxxxx.xml路径
Dim xmlPath
xmlPath=VbsPath &"\Result\ErrLog"&Cstr(date())&".xml"
'定义用例路径
Dim casePath_prefix
casePath_prefix=VbsPath&"\Script\"
'定义记录用例检查点状态的txt文件路径
Dim StatusTxtPath
StatusTxtPath=VbsPath & "\Result\CaseStatus"&Cstr(date())&".txt"
'=======================用例的执行=========================
'定义用例执行excel中获取参数后的字典对象
Dim objDic
set objDic=GetExcelPara()
Dim Res  '定义用例的检查点结果
'根据字典值个数去循环执行用例,字典值数量除2是因为每两个key的下标值一样
For k=1 to objDic.count/2
  '清除记录用例检查点状态的txt文件
   if(fso.FileExists(StatusTxtPath))then
     fso.DeleteFile(StatusTxtPath)
   end if
  'case执行
  qtApp.Open  casePath_prefix & objDic("CaseName"&k)
  'qtApp.Test.Settings.Run.IterationMode="rngAll"  '*******要运行指定行需要注释此脚本，且QTP的setting要设置为"只运行第一行"******
  qtApp.Test.Settings.Run.StartIteration=objDic("RowNum"&k)
  qtApp.Test.Settings.Run.EndIteration=objDic("RowNum"&k)
  qtApp.Options.Run.RunMode = "Fast"
  qtApp.Test.Run
  '判断ErrLog.xml是否存在，是否重跑Case
  Call ErrLogIsExist(casePath_prefix, objDic("CaseName"&k),objDic("RowNum"&k), xmlPath)
  '获取记录的测试结果，写回用例excel结果列!!!!!!用例excel写入前一定要签出到本地，否则为只读，无法写入结果!!!!!!!!
  Res= ReadCaseStatus
  Call WriteExcelResults(objDic("CaseName"&k),objDic("RowNum"&k),Res)  
Next

'/*****************************获取用例excel中的参数：用例名，迭代数据行号****************************************/
Function GetExcelPara()
  Dim oDict
  '定义字典，存储此函数返回的多值
  Set oDict=CreateObject("Scripting.Dictionary")                        
  '定义excel应用
  Set excelApp=createobject("excel.application")
  '打开workbook
  excelApp.Workbooks.Open(VbsPath&"\AutoTestCase.xlsx") '打开用例文档AutoTestCase.xlsx
  '获取sheet
  set oSheet=excelApp.Sheets(1)   '获取第X个Sheet
  Dim RowCount
  RowCount=oSheet.UsedRange.Rows.Count '获取excel总行数
  Dim i 'excel行标
  Dim j 'excel列标
  j=1
  For i=1  to RowCount
    if(oSheet.cells(i,5)="Y")then  '目前用例是否执行的标志是固定的第6列，如不固定，需要改为动态遍历
	oDict("CaseName"& j)=oSheet.cells(i,1)
	oDict("RowNum"& j)=oSheet.cells(i,2)
	j=j+1
    end if
  Next
  '返回字典
  Set GetExcelPara=oDict
  '关闭当前工作薄
  excelApp.ActiveWorkbook.Close
  '关闭excel
  excelApp.Quit
  '释放excel对象
  Set excelApp=nothing
End Function
'/*****************************读取用例当前迭代的检查点状态记录****************************************/
Function ReadCaseStatus
   Const ForReading = 1
	if(fso.FileExists(StatusTxtPath))then
		Set txtFile=fso.OpenTextFile(StatusTxtPath,ForReading,true)
		ReadCaseStatus=txtFile.ReadAll
		txtFile.close
		Set txtFile=nothing
	else
		msgbox "用例当前迭代的检查点状态记录文件不存在"
	end if
End Function

'/*****************************根据用例名，迭代数据行号对应将运行结果保存至用例excel结果列****************************************/
Sub WriteExcelResults(currCaseName,currRowNum,LastRes)        
   if(LastRes="0")then
   LastRes="Success"
   end if
   if(LastRes="1")then
   LastRes="Failed"
   end if
  '定义excel应用
  Set excApp=createobject("excel.application")
  '打开workbook
  excApp.Workbooks.Open(VbsPath & "\AutoTestCase.xlsx") '用例汇总的EXCEL文档地址AutoTestCase.xlsx
  '获取sheet
  set oSheet=excApp.Sheets(1)     '获取第X个Sheet
  Dim RowCount
  RowCount=oSheet.UsedRange.Rows.Count
  For t=1  to RowCount            'i为excel数据所在行号
    if(oSheet.cells(t,1)=currCaseName and oSheet.cells(t,2)=currRowNum)then  '目前用例是否执行的标志是固定的第6列，如不固定，需要改为动态遍历
       oSheet.cells(t,4)=LastRes  '用例excel中结果列为第5列
       Exit For
    end if
  Next
  '保存当前工作薄
  excApp.ActiveWorkbook.Save
  '关闭当前工作薄
  excApp.ActiveWorkbook.Close
  '关闭excel
  excApp.Quit
  '释放excel对象
  Set excApp=nothing
End Sub
'/*****************************判断ErrLog.xml是否存在重新运行case****************************************/
Sub ErrLogIsExist(casePath_prefix, caseName,currRowNum,xmlPath)
    '定义oFSO对象
    Dim oFSO
    set oFSO=createObject("Scripting.FileSystemObject")
    if(oFSO.FileExists(xmlPath))then
      set dictXML=GetXml(xmlPath,caseName)
      '如果读取xml返回的对象不为空，则读取带回的CurrentRow，RowCount
      if(dictXML.count>0)then
	if(dictXML("CurrentRow")>0)then
          Call IsRunCaseAgain(casePath_prefix, caseName,currRowNum)
	  '释放资源
	  set dictXML=nothing
	end if
      end if
    end if
    '释放资源
    Set oFSO=nothing
End Sub
'/*****************************判断当前迭代是否存在检查点Fail，需要重新运行一次当前迭代****************************************/
Sub IsRunCaseAgain(casePath_prefix, caseName,currRowNum)
  '清除记录用例检查点状态的txt文件
   if(fso.FileExists(StatusTxtPath))then
     fso.DeleteFile(StatusTxtPath)
   end if
   '重跑case
   qtApp.Open casePath_prefix & caseName
   qtApp.Test.Settings.Run.StartIteration=currRowNum
   qtApp.Test.Settings.Run.EndIteration=currRowNum
   qtApp.Options.Run.RunMode = "Fast"
   qtApp.Test.Run
End Sub
'/*********************************读取ErrLog.xml********************************************/
Function GetXml(strXmlFilePath,caseNameVal)
  Dim oDict
  Set oDict=CreateObject("Scripting.Dictionary")        '定义字典，存储此函数返回的多值
  Dim xmlDoc,xmlRoot
  Set xmlDoc = CreateObject("MSXML2.DOMDocument")    	'创建XML DOM对象
  xmlDoc.async = False                                  '控制加载模式为同步模式（xml树加载完毕后再执行后续代码）                     
  xmlDoc.load strXmlFilePath                            '载入xml文件       
  If xmlDoc.parseError.errorCode <> 0 Then
     MsgBox "XML文件格式不对，原因是：" & Chr(13) & xmlDoc.parseError.reason
     Exit Function
  End If
  '获取根结点
  Set xmlRoot = xmlDoc.documentElement
  '遍历xml
  For i=0 to  xmlRoot.childNodes.length-1
      if(xmlRoot.childNodes(i).attributes(0).text=caseNameVal)then
	  For j=0 to xmlRoot.childNodes(i).childNodes.length-1
	     For p=0 to xmlRoot.childNodes(i).childNodes(j).childNodes.length-1
		 if(xmlRoot.childNodes(i).childNodes(j).selectSingleNode("ErrCode").text<>0)then
		    '将当前行号和所有行数两个节点值存储到字典里
		    oDict("CurrentRow")=xmlRoot.childNodes(i).childNodes(j).selectSingleNode("CurrentRow").text
		    oDict("RowCount")=xmlRoot.childNodes(i).childNodes(j).selectSingleNode("RowCount").text
		    set GetXml = oDict
		    Exit for  '找到相同版本号后退出三级节点循环，停止遍历
		 end if 
	     Next
	     Exit For   '直接退出二级节点循环，返回值
	  next
	  Exit for  '找到相同用例名后则退出外层循环，停止遍历
     else
	  Set GetXml=oDict
     end if
  Next
End Function
'释放fso对象
set fso=nothing
'退出QTP
qtApp.Quit
'释放qtp应用对象
Set qtApp=Nothing
