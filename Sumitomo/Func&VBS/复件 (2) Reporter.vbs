'��д�Զ����� 
Dim Reporter 
Set Reporter = New clsReporter 
'������Ա�ACTION���õ�FUNCTION 
Public Function GetReporter 
    Set GetReporter = Reporter 
End Function 

'���ض��� 
Class clsReporter 
    Dim oExcel  '����excelӦ�ö���
    Dim arr        '�������ݣ��洢���sStepName��ֺ������
    Dim fso       '����FSO����
    Dim CurrRow 'txt�洢��excel������д�������
'***************************************************************************************************************************
    'д���㵽excel������
    Public Sub ReportEvent(iStatus, sStepName, sDetails) 
		'��ȡ���ű���Ŀ¼��binĿ¼ fso.GetFolder(".")����ȡ��Ŀ����HPĿ¼ fso.GetFolder(".").ParentFolder.ParentFolder
		Dim ExcelPath,TxtPath,StatusTxtPath
msgbox "Reporter.vbs�����·����"&fso.GetFolder(".")
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
	   '���sStepName������Ŀǰ���������Ͳ�������������Ӱ汾�ź���Ҫ���������±�
		arr=split(sStepName,"?") 
		'д��excel����,  ���Reporter.xls���ڣ���ֱ�Ӽ�¼���㣻���򴴽�excel����д����
		if(fso.FileExists(ExcelPath))then
				'msgbox "excel����"
				'��ȡtxt��¼��excel�Ѽ�¼������
				if(fso.FileExists(TxtPath))then
					Set txtFile=fso.OpenTextFile(TxtPath,1,true)
					CurrRow=txtFile.ReadAll
					Set txtFile=nothing
				else
					'txt�ļ���ȡʧ�ܣ����������ִ�������ݲ�����
				end if
				' ��Excel�ļ�        
				oExcel.Workbooks.Open(ExcelPath)
				'дexcel
				oExcel.Worksheets(1).cells(CurrRow+1,1)=arr(0)  '������
				oExcel.Worksheets(1).cells(CurrRow+1,2)=arr(1)  '�����к�
				oExcel.Worksheets(1).cells(CurrRow+1,3)=arr(2)  '����
				oExcel.Worksheets(1).cells(CurrRow+1,4)=sDetails '�����
				oExcel.Worksheets(1).cells(CurrRow+1,5)=iStatus  'ִ�н��״̬
				oExcel.Worksheets(1).cells(CurrRow+1,6)=Cstr(now)      '�����¼ʱ��
				'txt�ļ�����ʱ,�жϼ���״̬,�ٴ洢
				'msgbox "���μ�¼����״̬: "& iStatus
				if(fso.FileExists(StatusTxtPath))then
					'msgbox "�ļ�����: "& iStatus
					if(iStatus="1")then
                                           'msgbox "1��Failed"
					   Set txtFile=fso.OpenTextFile(StatusTxtPath,2,true)
					   txtFile.write(iStatus)
					   txtFile.close
					   Set txtFile=nothing
					else
					   '���״̬��Passed���ٸ���
					end if
				'txt�ļ�������ʱ���ȴ����ٴ洢�����ļ�����
				else
					'msgbox "�ļ��´���: "& iStatus
					Set txtFile=fso.CreateTextFile(StatusTxtPath,true)        
					txtFile.write(iStatus)
					txtFile.close
					Set txtFile=nothing
				end if 
				' �洢excel���¼�¼������,�����´�׷��
				CurrRow=oExcel.Worksheets(1).UsedRange.Rows.Count
			   '���湤����
				oExcel.DisplayAlerts=false   '�˾�����α���ǰResume.xlw�ļ��Ƿ��滻����ʾ
				oExcel.Workbooks(1).Save
				'�رչ�����
				oExcel.WorkBooks.Item(1).close 		
		else '����excel��д����,����һ��д,����Ҫȥ��ȡtxt�ļ���¼����������txt�ļ�Ҳ������
				'msgbox "excel������"
				set oWorkbook=oExcel.Workbooks.Add
				Set oWorksheet=oWorkbook.WorkSheets.add
				oWorksheet.name="Reporter"
				'дexcel���粻�����µ�sheetͨ��index��sheet�����⣬���±��������Ϊ��
				oWorksheet.cells(1,1)="������"
				oWorksheet.cells(1,2)="�����к�"
				oWorksheet.cells(1,3)="��鲽��"
				oWorksheet.cells(1,4)="�����"
				oWorksheet.cells(1,5)="�Ƿ�ͨ��"
				oWorksheet.cells(1,6)="ִ��ʱ��"
				oWorksheet.cells(2,1)=arr(0)  '������
				oWorksheet.cells(2,2)=arr(1)  '�����к�
				oWorksheet.cells(2,3)=arr(2)  '��鲽��
				oWorksheet.cells(2,4)=sDetails  '�����
				oWorksheet.cells(2,5)=iStatus  '�Ƿ�ͨ��
				oWorksheet.cells(2,6)=Cstr(now)  'ִ��ʱ��
				'txt�ļ�����ʱ,�жϼ���״̬,�ٴ洢
				'msgbox "һ�μ�¼����״̬: "& iStatus
				if(fso.FileExists(StatusTxtPath))then
					'msgbox "�ļ�����: "& iStatus
					Set txtFile=fso.OpenTextFile(StatusTxtPath,2,true)
					txtFile.write(iStatus)
					txtFile.close
					Set txtFile=nothing
				'txt�ļ�������ʱ���ȴ����ٴ洢�����ļ�����
				else
					'msgbox "�ļ��´���: "& iStatus
					Set txtFile=fso.CreateTextFile(StatusTxtPath,true)        
					txtFile.write(iStatus)
					txtFile.close
					Set txtFile=nothing
				end if 
				' �洢excel���¼�¼������,�����´�׷��
				CurrRow=oExcel.Worksheets("Reporter").UsedRange.Rows.Count
				oWorkbook.SaveAs ExcelPath
			        oWorkbook.Close
		end if         
		'txt�ļ�����ʱ���洢excel��¼������к�
		if(fso.FileExists(TxtPath))then
			Set txtFile=fso.OpenTextFile(TxtPath,2,true)
			txtFile.write(CurrRow)
			txtFile.close
			Set txtFile=nothing
		'txt�ļ�������ʱ���ȴ����ٴ洢excel��¼������кţ����״δ洢�ĳ���
		else
			Set txtFile=fso.CreateTextFile(TxtPath,true)        
			txtFile.write(CurrRow)
			txtFile.close
			Set txtFile=nothing
		end if  
    End Sub
    '��ʼ��Reporter�� 
    Private Sub Class_Initialize
        ' ����ExcelӦ�ó������        
        Set oExcel = CreateObject("Excel.Application") 
        ' ����fso����
        Set fso=CreateObject("scripting.FileSystemObject") 
        'msgbox "��ʼ�����"        
    End Sub 
    Private Sub Class_Terminate
	'�ͷ�fso����
	Set fso=nothing
	' �˳�Excel        
	oExcel.Quit        
	Set oExcel = Nothing
        'msgbox "Reporter��������"
    End Sub
'/*****************************���������������������кŶ�Ӧ�����н������������excel�����****************************************/
Sub WriteExcelResults(currCaseName,currRowNum,LastRes)         
  '����excelӦ��
  Set excApp=createobject("excel.application")
  '��workbook
  excApp.Workbooks.Open("C:\Program Files\HP\QuickTest Professional\Tests\IEMS\AutoTestCase.xlsx") '�������ܵ�EXCEL�ĵ���ַAutoTestCase.xlsx
  '��ȡsheet
  set oSheet=excApp.Sheets(1)     '��ȡ��X��Sheet
  Dim RowCount
  RowCount=oSheet.UsedRange.Rows.Count
  For t=1  to RowCount            'iΪexcel���������к�
    if(oSheet.cells(t,1)=currCaseName and oSheet.cells(t,2)=currRowNum)then  'Ŀǰ�����Ƿ�ִ�еı�־�ǹ̶��ĵ�6�У��粻�̶�����Ҫ��Ϊ��̬����
       oSheet.cells(t,5)=LastRes  '����excel�н����Ϊ��5��
       Exit For
    end if
  Next
  '���浱ǰ������
  excApp.ActiveWorkbook.Save
  '�رյ�ǰ������
  excApp.ActiveWorkbook.Close
  '�ر�excel
  excApp.Quit
  '�ͷ�excel����
  Set excApp=nothing
End Sub
    '***************************************************************************************************************************
    'XmlDomDoc_ErrLog����: �½����޸ļ�¼ErrLog��Xml�ļ�
    '***************************************************************************************************************************
    Sub  XmlDomDoc_ErrLog(caseName,versionNo,actionName,currRow,rowCount,errCode,errDesc,errSour,RecTime)  
        '����FSO���������ж�xml�Ƿ����
        Dim oFSO         
	    set oFSO = CreateObject ("Scripting.FileSystemObject") 	
        'xml�ĵ�·������ȡ���ű���Ŀ¼��binĿ¼ fso.GetFolder(".")����ȡ��Ŀ����HPĿ¼ fso.GetFolder(".").ParentFolder.ParentFolder
        Dim xmPath
        if(fso.GetFolder(".").ParentFolder.ParentFolder="C:\Program Files\HP\QuickTest Professional\Tests\Sumitomo")then
          xmlPath=fso.GetFolder(".").ParentFolder.ParentFolder & "\Result\ErrLog"&Cstr(date())&".xml"
        else
			  if(fso.GetFolder(".")="C:\Program Files\HP\QuickTest Professional\bin")then
			  xmlPath=fso.GetFolder(".").ParentFolder & "\Tests\Sumitomo\Result\ErrLog"&Cstr(date())&".xml"
			  end if
        end if	 
        '����xml�ĵ�����
		Dim xmlDoc     
		Set xmlDoc = CreateObject("MSXML2.DOMDocument") 	
		Dim rootE1                                                                              '������ڵ�
		Dim ChildE1,ChildE1Attribute1,ChildE1Attribute2    	'����һ���ڵ㼰����	
		Dim ChildE2_VersionNo,ChildE2Attribute1                '��������ڵ㼰����
		Dim ChildE3_ActionName,ChildE3_CurrRow,ChildE3_ErrCode,ChildE3_ErrDesc,ChildE3_ErrSour,ChildE3_RecTime   '���������ڵ㼰����
		'�ж�xml  Log�ļ��Ƿ���ڣ��������ȡ���ڵ㣬����������Ӹ��ڵ�
		BoolVal=oFSO.FileExists(xmlPath)
		if(BoolVal)then
				'����xml�ļ�
				xmlDoc.load xmlPath
				 '��ȡ���ڵ�
				Set	rootE1=xmlDoc.documentElement 
				'***************ml�ļ�����,�����һ���ڵ�(������)�������ڵ�(�汾��)�޸������ڵ��ֵ
				'�жϸ����Ƿ��н��
				If  rootE1.hasChildNodes then
					   Dim flag    '�Ƿ񴴽��½ڵ�
						flag=true  'Ĭ��ֵ��true���������ڵ�
					   Dim flag_first      '�Ƿ񴴽�һ���ڵ�
					   flag_first=true     'Ĭ��ֵ��true����ԭ������׷�Ӱ汾�Žڵ�
					   '�������µ�һ���ڵ�
						For i=0 to rootE1.childNodes.length-1											
								'һ���ڵ��CaseName����ֵ�Ƿ�Ϊ���������
								if(rootE1.childNodes(i).attributes(0).nodeName="CaseName" and rootE1.childNodes(i).attributes(0).text=caseName)then
										For j=0 to rootE1.childNodes(i).childNodes.length-1
												'�����ڵ��VersionNo����ֵ�Ƿ�Ϊ��ΰ汾��
												if(rootE1.childNodes(i).childNodes(j).attributes(0).nodeName="VersionNo" and rootE1.childNodes(i).childNodes(j).attributes(0).text=versionNo)then
														'һ���ڵ�(������)�Ͷ����ڵ�(�汾��)�� ��ͬ���������ڵ�Ľڵ�ֵ����
														 rootE1.childNodes(i).childNodes(j).childNodes(0).text=actionName
														 rootE1.childNodes(i).childNodes(j).childNodes(1).text=currRow
														 rootE1.childNodes(i).childNodes(j).childNodes(2).text=rowCount
														 rootE1.childNodes(i).childNodes(j).childNodes(3).text=errCode
														 rootE1.childNodes(i).childNodes(j).childNodes(4).text=errDesc
														 rootE1.childNodes(i).childNodes(j).childNodes(5).text=errSour
														 rootE1.childNodes(i).childNodes(j).childNodes(6).text=RecTime												 
														 flag=false         '��������ͬ�������Ͱ汾�ŵĽڵ����ݺ���flag=false�������ٴ����½ڵ� 
														 Exit For             '�˳���ǰѭ�����������ڵ�ı���
												end if                   '�ж϶����ڵ��VersionNo����ֵ�Ƿ�Ϊ��ΰ汾��
										 Next                      '�����ڵ��������
										 flag_first=false  '��������ͬ���汾�Ų�ͬ��ֱ�������´����汾�Žڵ�
										 Exit For               '�˳���ǰѭ������һ���ڵ�ı���
								end if   'һ���ڵ��жϽ���
						Next  '�������µ�һ���ڵ����
						'*******************�����ֵflag=trueʱ�´���һ���ڵ�(������)����ԭһ���ڵ�׷�Ӷ����ڵ�(��׷�Ӱ汾��)***********************
						if(flag)then
							 if(flag_first)then
								'����һ���ڵ㣬��������
								Set ChildE1=xmlDoc.createElement("TestCase")
								'����һ���ڵ������
								Set ChildE1Attribute1=xmlDoc.createAttribute("CaseName")
								ChildE1Attribute1.text=caseName
								ChildE1.setAttributeNode ChildE1Attribute1
								Set ChildE1Attribute2=xmlDoc.createAttribute("Description")
								ChildE1Attribute2.text="ErrLog"
								ChildE1.setAttributeNode ChildE1Attribute2
							else
							   '��ȡһ���ڵ�
								For m=0 to rootE1.childNodes.length-1
									if(rootE1.childNodes(m).attributes(0).text=caseName)then
										 Set ChildE1=rootE1.childNodes(m)
									end if
								Next
							end if
							'���������ڵ�	����������	
							Set ChildE2_VersionNo=xmlDoc.createElement("VersionNo")
							'���ö����ڵ������
							Set ChildE2Attribute1=xmlDoc.createAttribute("VersionNo")
							ChildE2Attribute1.text=versionNo
							ChildE2_VersionNo.setAttributeNode ChildE2Attribute1
							ChildE1.appendChild  ChildE2_VersionNo
							'���������ڵ�1		
							Set ChildE3_ActionName=xmlDoc.createElement("ActionName")
							ChildE3_ActionName.text=actionName
							ChildE2_VersionNo.appendChild ChildE3_ActionName
							'���������ڵ�2		
							Set ChildE3_CurrRow=xmlDoc.createElement("CurrentRow")
							ChildE3_CurrRow.text=currRow
							ChildE2_VersionNo.appendChild ChildE3_CurrRow
							'���������ڵ�3		
							Set ChildE3_RowCount=xmlDoc.createElement("RowCount")
							ChildE3_RowCount.text=rowCount
							ChildE2_VersionNo.appendChild ChildE3_RowCount
							'���������ڵ�4		
							Set ChildE3_ErrCode=xmlDoc.createElement("ErrCode")
							ChildE3_ErrCode.text=errCode 
							ChildE2_VersionNo.appendChild ChildE3_ErrCode
							'���������ڵ�5
							Set ChildE3_ErrDesc=xmlDoc.createElement("ErrDescription")
							ChildE3_ErrDesc.text=errDesc 
							ChildE2_VersionNo.appendChild ChildE3_ErrDesc
							'���������ڵ�6	
							Set ChildE3_ErrSour=xmlDoc.createElement("ErrSource")
							ChildE3_ErrSour.text=errSour
							ChildE2_VersionNo.appendChild ChildE3_ErrSour
							'���������ڵ�7
							Set ChildE3_RecTime=xmlDoc.createElement("RecordTime")
							ChildE3_RecTime.text=RecTime
							ChildE2_VersionNo.appendChild ChildE3_RecTime
							'���һ���ڵ㵽���ڵ�	
							rootE1.appendChild ChildE1
						end if     '�´��������ڵ㼰�ӽڵ����
				end if        '�жϸ����Ƿ��н�����	
		
		 else	 'xml�����ڣ������´���xml�ĵ�      
				'���� XML processing instruction�������ӵ���Ԫ��֮ǰ
				Set p=xmlDoc.createProcessingInstruction("xml","version='1.0' encoding='GB2312'")
				xmlDoc.insertBefore p,xmlDoc.childNodes(0)	
				'������Ԫ�ز���֮�����ĵ�
				Set rootE1=xmlDoc.createElement("ErrMsg")
				xmlDoc.appendChild rootE1
				 '*******************����һ���ڵ�����ڵ�***********************
				'����һ���ڵ㣬��������
				Set ChildE1=xmlDoc.createElement("TestCase")
				'����һ���ڵ������
				Set ChildE1Attribute1=xmlDoc.createAttribute("CaseName")
				ChildE1Attribute1.text=caseName
				ChildE1.setAttributeNode ChildE1Attribute1
				Set ChildE1Attribute2=xmlDoc.createAttribute("Description")
				ChildE1Attribute2.text="ErrLog"
				ChildE1.setAttributeNode ChildE1Attribute2
				'���������ڵ�	����������	
				Set ChildE2_VersionNo=xmlDoc.createElement("VersionNo")
				'���ö����ڵ������
				Set ChildE2Attribute1=xmlDoc.createAttribute("VersionNo")
				ChildE2Attribute1.text=versionNo
				ChildE2_VersionNo.setAttributeNode ChildE2Attribute1
				ChildE1.appendChild  ChildE2_VersionNo
				'���������ڵ�1		
				Set ChildE3_ActionName=xmlDoc.createElement("ActionName")
				ChildE3_ActionName.text=actionName
				ChildE2_VersionNo.appendChild ChildE3_ActionName
				'���������ڵ�2		
				Set ChildE3_CurrRow=xmlDoc.createElement("CurrentRow")
				ChildE3_CurrRow.text=currRow
				ChildE2_VersionNo.appendChild ChildE3_CurrRow
				'���������ڵ�3
				Set ChildE3_RowCount=xmlDoc.createElement("RowCount")
				ChildE3_RowCount.text=rowCount
				ChildE2_VersionNo.appendChild ChildE3_RowCount
				'���������ڵ�4		
				Set ChildE3_ErrCode=xmlDoc.createElement("ErrCode")
				ChildE3_ErrCode.text=errCode 
				ChildE2_VersionNo.appendChild ChildE3_ErrCode
				'���������ڵ�5
				Set ChildE3_ErrDesc=xmlDoc.createElement("ErrDescription")
				ChildE3_ErrDesc.text=errDesc 
				ChildE2_VersionNo.appendChild ChildE3_ErrDesc
				'���������ڵ�6	
				Set ChildE3_ErrSour=xmlDoc.createElement("ErrSource")
				ChildE3_ErrSour.text=errSour
				ChildE2_VersionNo.appendChild ChildE3_ErrSour
				'���������ڵ�7
				Set ChildE3_RecTime=xmlDoc.createElement("RecordTime")
				ChildE3_RecTime.text=RecTime
				ChildE2_VersionNo.appendChild ChildE3_RecTime
				'���һ���ڵ㵽���ڵ�	
				rootE1.appendChild ChildE1
		end if	
		'�ļ�����
		xmlDoc.Save xmlPath
    End Sub
End Class