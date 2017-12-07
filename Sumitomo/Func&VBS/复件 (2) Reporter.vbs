'重写自定义类 
Dim Reporter 
Set Reporter = New clsReporter 
'定义可以被ACTION调用的FUNCTION 
Public Function GetReporter 
    Set GetReporter = Reporter 
End Function 

'类重定义 
Class clsReporter 
    Dim oExcel  '定义excel应用对象
    Dim arr        '定义数据，存储入参sStepName拆分后的内容
    Dim fso       '定义FSO对象
    Dim CurrRow 'txt存储的excel报告所写入的行数
'***************************************************************************************************************************
    '写检查点到excel报告中
    Public Sub ReportEvent(iStatus, sStepName, sDetails) 
		'获取到脚本名目录或bin目录 fso.GetFolder(".")；获取项目名或HP目录 fso.GetFolder(".").ParentFolder.ParentFolder
		Dim ExcelPath,TxtPath,StatusTxtPath
msgbox "Reporter.vbs中相对路径："&fso.GetFolder(".")
		if(fso.GetFolder(".").ParentFolder.ParentFolder="C:\Program Files\HP\QuickTest Professional\Tests\Sumitomo")then
		   ExcelPath=fso.GetFolder(".").ParentFolder.ParentFolder & "\Result\Reporter"&Cstr(date())&".xls"
		   TxtPath=fso.GetFolder(".").ParentFolder.ParentFolder & "\Result\ExcelCurrRow"&Cstr(date())&".txt"
           StatusTxtPath=fso.GetFolder(".").ParentFolder.ParentFolder & "\Result\CaseStatus"&Cstr(date())&".txt"
		else
		   if(fso.GetFolder(".")="C:\Program Files\HP\QuickTest Professional\bin")then
	         ExcelPath=fso.GetFolder(".").ParentFolder & "\Tests\Sumitomo\Result\Reporter"&Cstr(date())&".xls"
		     TxtPath=fso.GetFolder(".").ParentFolder & "\Tests\Sumitomo\Result\ExcelCurrRow"&Cstr(date())&".txt"
		     StatusTxtPath=fso.GetFolder(".").ParentFolder & "\Tests\Sumitomo\Result\CaseStatus"&Cstr(date())&".txt"
		   end if
		end if	 
	   '拆分sStepName参数，目前含用例名和步骤名，后续添加版本号后，需要调整数组下标
		arr=split(sStepName,"?") 
		'写入excel检查点,  如果Reporter.xls存在，则直接记录检查点；否则创建excel后再写检查点
		if(fso.FileExists(ExcelPath))then
				'msgbox "excel存在"
				'读取txt记录的excel已记录的行数
				if(fso.FileExists(TxtPath))then
					Set txtFile=fso.OpenTextFile(TxtPath,1,true)
					CurrRow=txtFile.ReadAll
					Set txtFile=nothing
				else
					'txt文件读取失败，正常不出现此情况，暂不考虑
				end if
				' 打开Excel文件        
				oExcel.Workbooks.Open(ExcelPath)
				'写excel
				oExcel.Worksheets(1).cells(CurrRow+1,1)=arr(0)  '用例名
				oExcel.Worksheets(1).cells(CurrRow+1,2)=arr(1)  '迭代行号
				oExcel.Worksheets(1).cells(CurrRow+1,3)=arr(2)  '检查点
				oExcel.Worksheets(1).cells(CurrRow+1,4)=sDetails '检查结果
				oExcel.Worksheets(1).cells(CurrRow+1,5)=iStatus  '执行结果状态
				oExcel.Worksheets(1).cells(CurrRow+1,6)=Cstr(now)      '检查点记录时间
				'txt文件存在时,判断检查点状态,再存储
				'msgbox "二次记录检查点状态: "& iStatus
				if(fso.FileExists(StatusTxtPath))then
					'msgbox "文件存在: "& iStatus
					if(iStatus="1")then
                                           'msgbox "1是Failed"
					   Set txtFile=fso.OpenTextFile(StatusTxtPath,2,true)
					   txtFile.write(iStatus)
					   txtFile.close
					   Set txtFile=nothing
					else
					   '如果状态是Passed则不再更新
					end if
				'txt文件不存在时，先创建再存储用例的检查点结果
				else
					'msgbox "文件新创建: "& iStatus
					Set txtFile=fso.CreateTextFile(StatusTxtPath,true)        
					txtFile.write(iStatus)
					txtFile.close
					Set txtFile=nothing
				end if 
				' 存储excel最新记录的行数,便于下次追加
				CurrRow=oExcel.Worksheets(1).UsedRange.Rows.Count
			   '保存工作薄
				oExcel.DisplayAlerts=false   '此句可屏蔽保存前Resume.xlw文件是否替换的提示
				oExcel.Workbooks(1).Save
				'关闭工作薄
				oExcel.WorkBooks.Item(1).close 		
		else '创建excel再写检查点,即第一次写,不需要去读取txt文件记录的行数，且txt文件也不存在
				'msgbox "excel不存在"
				set oWorkbook=oExcel.Workbooks.Add
				Set oWorksheet=oWorkbook.WorkSheets.add
				oWorksheet.name="Reporter"
				'写excel，如不创建新的sheet通过index找sheet有问题，导致保存的内容为空
				oWorksheet.cells(1,1)="用例名"
				oWorksheet.cells(1,2)="迭代行号"
				oWorksheet.cells(1,3)="检查步骤"
				oWorksheet.cells(1,4)="检查结果"
				oWorksheet.cells(1,5)="是否通过"
				oWorksheet.cells(1,6)="执行时间"
				oWorksheet.cells(2,1)=arr(0)  '用例名
				oWorksheet.cells(2,2)=arr(1)  '迭代行号
				oWorksheet.cells(2,3)=arr(2)  '检查步骤
				oWorksheet.cells(2,4)=sDetails  '检查结果
				oWorksheet.cells(2,5)=iStatus  '是否通过
				oWorksheet.cells(2,6)=Cstr(now)  '执行时间
				'txt文件存在时,判断检查点状态,再存储
				'msgbox "一次记录检查点状态: "& iStatus
				if(fso.FileExists(StatusTxtPath))then
					'msgbox "文件存在: "& iStatus
					Set txtFile=fso.OpenTextFile(StatusTxtPath,2,true)
					txtFile.write(iStatus)
					txtFile.close
					Set txtFile=nothing
				'txt文件不存在时，先创建再存储用例的检查点结果
				else
					'msgbox "文件新创建: "& iStatus
					Set txtFile=fso.CreateTextFile(StatusTxtPath,true)        
					txtFile.write(iStatus)
					txtFile.close
					Set txtFile=nothing
				end if 
				' 存储excel最新记录的行数,便于下次追加
				CurrRow=oExcel.Worksheets("Reporter").UsedRange.Rows.Count
				oWorkbook.SaveAs ExcelPath
			        oWorkbook.Close
		end if         
		'txt文件存在时，存储excel记录的最大行号
		if(fso.FileExists(TxtPath))then
			Set txtFile=fso.OpenTextFile(TxtPath,2,true)
			txtFile.write(CurrRow)
			txtFile.close
			Set txtFile=nothing
		'txt文件不存在时，先创建再存储excel记录的最大行号，即首次存储的场景
		else
			Set txtFile=fso.CreateTextFile(TxtPath,true)        
			txtFile.write(CurrRow)
			txtFile.close
			Set txtFile=nothing
		end if  
    End Sub
    '初始化Reporter类 
    Private Sub Class_Initialize
        ' 创建Excel应用程序对象        
        Set oExcel = CreateObject("Excel.Application") 
        ' 创建fso对象
        Set fso=CreateObject("scripting.FileSystemObject") 
        'msgbox "初始化完毕"        
    End Sub 
    Private Sub Class_Terminate
	'释放fso对象
	Set fso=nothing
	' 退出Excel        
	oExcel.Quit        
	Set oExcel = Nothing
        'msgbox "Reporter对象销毁"
    End Sub
'/*****************************根据用例名，迭代数据行号对应将运行结果保存至用例excel结果列****************************************/
Sub WriteExcelResults(currCaseName,currRowNum,LastRes)         
  '定义excel应用
  Set excApp=createobject("excel.application")
  '打开workbook
  excApp.Workbooks.Open("C:\Program Files\HP\QuickTest Professional\Tests\IEMS\AutoTestCase.xlsx") '用例汇总的EXCEL文档地址AutoTestCase.xlsx
  '获取sheet
  set oSheet=excApp.Sheets(1)     '获取第X个Sheet
  Dim RowCount
  RowCount=oSheet.UsedRange.Rows.Count
  For t=1  to RowCount            'i为excel数据所在行号
    if(oSheet.cells(t,1)=currCaseName and oSheet.cells(t,2)=currRowNum)then  '目前用例是否执行的标志是固定的第6列，如不固定，需要改为动态遍历
       oSheet.cells(t,5)=LastRes  '用例excel中结果列为第5列
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
    '***************************************************************************************************************************
    'XmlDomDoc_ErrLog功能: 新建或修改记录ErrLog的Xml文件
    '***************************************************************************************************************************
    Sub  XmlDomDoc_ErrLog(caseName,versionNo,actionName,currRow,rowCount,errCode,errDesc,errSour,RecTime)  
        '定义FSO对象，用于判断xml是否存在
        Dim oFSO         
	    set oFSO = CreateObject ("Scripting.FileSystemObject") 	
        'xml文档路径，获取到脚本名目录或bin目录 fso.GetFolder(".")；获取项目名或HP目录 fso.GetFolder(".").ParentFolder.ParentFolder
        Dim xmPath
        if(fso.GetFolder(".").ParentFolder.ParentFolder="C:\Program Files\HP\QuickTest Professional\Tests\Sumitomo")then
          xmlPath=fso.GetFolder(".").ParentFolder.ParentFolder & "\Result\ErrLog"&Cstr(date())&".xml"
        else
			  if(fso.GetFolder(".")="C:\Program Files\HP\QuickTest Professional\bin")then
			  xmlPath=fso.GetFolder(".").ParentFolder & "\Tests\Sumitomo\Result\ErrLog"&Cstr(date())&".xml"
			  end if
        end if	 
        '定义xml文档对象
		Dim xmlDoc     
		Set xmlDoc = CreateObject("MSXML2.DOMDocument") 	
		Dim rootE1                                                                              '定义根节点
		Dim ChildE1,ChildE1Attribute1,ChildE1Attribute2    	'定义一级节点及属性	
		Dim ChildE2_VersionNo,ChildE2Attribute1                '定义二级节点及属性
		Dim ChildE3_ActionName,ChildE3_CurrRow,ChildE3_ErrCode,ChildE3_ErrDesc,ChildE3_ErrSour,ChildE3_RecTime   '定义三级节点及属性
		'判断xml  Log文件是否存在，存在则获取根节点，不存在则添加根节点
		BoolVal=oFSO.FileExists(xmlPath)
		if(BoolVal)then
				'加载xml文件
				xmlDoc.load xmlPath
				 '获取根节点
				Set	rootE1=xmlDoc.documentElement 
				'***************ml文件存在,则根据一级节点(用例名)、二级节点(版本号)修改三级节点的值
				'判断根下是否有结点
				If  rootE1.hasChildNodes then
					   Dim flag    '是否创建新节点
						flag=true  '默认值是true，即创建节点
					   Dim flag_first      '是否创建一级节点
					   flag_first=true     '默认值是true，即原用例下追加版本号节点
					   '遍历根下的一级节点
						For i=0 to rootE1.childNodes.length-1											
								'一级节点的CaseName属性值是否为入参用例名
								if(rootE1.childNodes(i).attributes(0).nodeName="CaseName" and rootE1.childNodes(i).attributes(0).text=caseName)then
										For j=0 to rootE1.childNodes(i).childNodes.length-1
												'二级节点的VersionNo属性值是否为入参版本号
												if(rootE1.childNodes(i).childNodes(j).attributes(0).nodeName="VersionNo" and rootE1.childNodes(i).childNodes(j).attributes(0).text=versionNo)then
														'一级节点(用例名)和二级节点(版本号)都 相同，则三级节点的节点值更新
														 rootE1.childNodes(i).childNodes(j).childNodes(0).text=actionName
														 rootE1.childNodes(i).childNodes(j).childNodes(1).text=currRow
														 rootE1.childNodes(i).childNodes(j).childNodes(2).text=rowCount
														 rootE1.childNodes(i).childNodes(j).childNodes(3).text=errCode
														 rootE1.childNodes(i).childNodes(j).childNodes(4).text=errDesc
														 rootE1.childNodes(i).childNodes(j).childNodes(5).text=errSour
														 rootE1.childNodes(i).childNodes(j).childNodes(6).text=RecTime												 
														 flag=false         '当更新相同用例名和版本号的节点数据后，置flag=false，即不再创建新节点 
														 Exit For             '退出当前循环，即二级节点的遍历
												end if                   '判断二级节点的VersionNo属性值是否为入参版本号
										 Next                      '二级节点遍历结束
										 flag_first=false  '用例名相同，版本号不同，直接在其下创建版本号节点
										 Exit For               '退出当前循环，即一级节点的遍历
								end if   '一级节点判断结束
						Next  '遍历根下的一级节点结束
						'*******************当标记值flag=true时新创建一级节点(即用例)或在原一级节点追加二级节点(即追加版本号)***********************
						if(flag)then
							 if(flag_first)then
								'创建一级节点，设置属性
								Set ChildE1=xmlDoc.createElement("TestCase")
								'设置一级节点的属性
								Set ChildE1Attribute1=xmlDoc.createAttribute("CaseName")
								ChildE1Attribute1.text=caseName
								ChildE1.setAttributeNode ChildE1Attribute1
								Set ChildE1Attribute2=xmlDoc.createAttribute("Description")
								ChildE1Attribute2.text="ErrLog"
								ChildE1.setAttributeNode ChildE1Attribute2
							else
							   '获取一级节点
								For m=0 to rootE1.childNodes.length-1
									if(rootE1.childNodes(m).attributes(0).text=caseName)then
										 Set ChildE1=rootE1.childNodes(m)
									end if
								Next
							end if
							'创建二级节点	，设置属性	
							Set ChildE2_VersionNo=xmlDoc.createElement("VersionNo")
							'设置二级节点的属性
							Set ChildE2Attribute1=xmlDoc.createAttribute("VersionNo")
							ChildE2Attribute1.text=versionNo
							ChildE2_VersionNo.setAttributeNode ChildE2Attribute1
							ChildE1.appendChild  ChildE2_VersionNo
							'创建三级节点1		
							Set ChildE3_ActionName=xmlDoc.createElement("ActionName")
							ChildE3_ActionName.text=actionName
							ChildE2_VersionNo.appendChild ChildE3_ActionName
							'创建三级节点2		
							Set ChildE3_CurrRow=xmlDoc.createElement("CurrentRow")
							ChildE3_CurrRow.text=currRow
							ChildE2_VersionNo.appendChild ChildE3_CurrRow
							'创建三级节点3		
							Set ChildE3_RowCount=xmlDoc.createElement("RowCount")
							ChildE3_RowCount.text=rowCount
							ChildE2_VersionNo.appendChild ChildE3_RowCount
							'创建三级节点4		
							Set ChildE3_ErrCode=xmlDoc.createElement("ErrCode")
							ChildE3_ErrCode.text=errCode 
							ChildE2_VersionNo.appendChild ChildE3_ErrCode
							'创建三级节点5
							Set ChildE3_ErrDesc=xmlDoc.createElement("ErrDescription")
							ChildE3_ErrDesc.text=errDesc 
							ChildE2_VersionNo.appendChild ChildE3_ErrDesc
							'创建三级节点6	
							Set ChildE3_ErrSour=xmlDoc.createElement("ErrSource")
							ChildE3_ErrSour.text=errSour
							ChildE2_VersionNo.appendChild ChildE3_ErrSour
							'创建三级节点7
							Set ChildE3_RecTime=xmlDoc.createElement("RecordTime")
							ChildE3_RecTime.text=RecTime
							ChildE2_VersionNo.appendChild ChildE3_RecTime
							'添加一级节点到根节点	
							rootE1.appendChild ChildE1
						end if     '新创建用例节点及子节点结束
				end if        '判断根下是否有结点结束	
		
		 else	 'xml不存在，如下新创建xml文档      
				'创建 XML processing instruction，把它加到根元素之前
				Set p=xmlDoc.createProcessingInstruction("xml","version='1.0' encoding='GB2312'")
				xmlDoc.insertBefore p,xmlDoc.childNodes(0)	
				'创建根元素并将之加入文档
				Set rootE1=xmlDoc.createElement("ErrMsg")
				xmlDoc.appendChild rootE1
				 '*******************创建一级节点二级节点***********************
				'创建一级节点，设置属性
				Set ChildE1=xmlDoc.createElement("TestCase")
				'设置一级节点的属性
				Set ChildE1Attribute1=xmlDoc.createAttribute("CaseName")
				ChildE1Attribute1.text=caseName
				ChildE1.setAttributeNode ChildE1Attribute1
				Set ChildE1Attribute2=xmlDoc.createAttribute("Description")
				ChildE1Attribute2.text="ErrLog"
				ChildE1.setAttributeNode ChildE1Attribute2
				'创建二级节点	，设置属性	
				Set ChildE2_VersionNo=xmlDoc.createElement("VersionNo")
				'设置二级节点的属性
				Set ChildE2Attribute1=xmlDoc.createAttribute("VersionNo")
				ChildE2Attribute1.text=versionNo
				ChildE2_VersionNo.setAttributeNode ChildE2Attribute1
				ChildE1.appendChild  ChildE2_VersionNo
				'创建三级节点1		
				Set ChildE3_ActionName=xmlDoc.createElement("ActionName")
				ChildE3_ActionName.text=actionName
				ChildE2_VersionNo.appendChild ChildE3_ActionName
				'创建三级节点2		
				Set ChildE3_CurrRow=xmlDoc.createElement("CurrentRow")
				ChildE3_CurrRow.text=currRow
				ChildE2_VersionNo.appendChild ChildE3_CurrRow
				'创建三级节点3
				Set ChildE3_RowCount=xmlDoc.createElement("RowCount")
				ChildE3_RowCount.text=rowCount
				ChildE2_VersionNo.appendChild ChildE3_RowCount
				'创建三级节点4		
				Set ChildE3_ErrCode=xmlDoc.createElement("ErrCode")
				ChildE3_ErrCode.text=errCode 
				ChildE2_VersionNo.appendChild ChildE3_ErrCode
				'创建三级节点5
				Set ChildE3_ErrDesc=xmlDoc.createElement("ErrDescription")
				ChildE3_ErrDesc.text=errDesc 
				ChildE2_VersionNo.appendChild ChildE3_ErrDesc
				'创建三级节点6	
				Set ChildE3_ErrSour=xmlDoc.createElement("ErrSource")
				ChildE3_ErrSour.text=errSour
				ChildE2_VersionNo.appendChild ChildE3_ErrSour
				'创建三级节点7
				Set ChildE3_RecTime=xmlDoc.createElement("RecordTime")
				ChildE3_RecTime.text=RecTime
				ChildE2_VersionNo.appendChild ChildE3_RecTime
				'添加一级节点到根节点	
				rootE1.appendChild ChildE1
		end if	
		'文件保存
		xmlDoc.Save xmlPath
    End Sub
End Class