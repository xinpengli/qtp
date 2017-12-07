'=====================VBS����=======================
'����QTPӦ��
Dim qtApp
Set qtApp=CreateObject("QuickTest.Application")
'����QTPӦ��
qtApp.Launch
'����qtp��ʾ����
qtApp.Visible = True
'��ȡ��ǰVBS�����·��������ǰִ�е�vbs�ļ����ڵ��ļ��е�·��������Ŀ��
Dim VbsPath
Set fso=CreateObject("Scripting.FileSystemObject")
VbsPath=fso.GetFolder(".")
'����ErrLogRead.vbs���ض���
Dim dictXML
'����ErrLogxxxx.xml·��
Dim xmlPath
xmlPath=VbsPath &"\Result\ErrLog"&Cstr(date())&".xml"
'��������·��
Dim casePath_prefix
casePath_prefix=VbsPath&"\Script\"
'�����¼��������״̬��txt�ļ�·��
Dim StatusTxtPath
StatusTxtPath=VbsPath & "\Result\CaseStatus"&Cstr(date())&".txt"
'=======================������ִ��=========================
'��������ִ��excel�л�ȡ��������ֵ����
Dim objDic
set objDic=GetExcelPara()
Dim Res  '���������ļ�����
'�����ֵ�ֵ����ȥѭ��ִ������,�ֵ�ֵ������2����Ϊÿ����key���±�ֵһ��
For k=1 to objDic.count/2
  '�����¼��������״̬��txt�ļ�
   if(fso.FileExists(StatusTxtPath))then
     fso.DeleteFile(StatusTxtPath)
   end if
  'caseִ��
  qtApp.Open  casePath_prefix & objDic("CaseName"&k)
  'qtApp.Test.Settings.Run.IterationMode="rngAll"  '*******Ҫ����ָ������Ҫע�ʹ˽ű�����QTP��settingҪ����Ϊ"ֻ���е�һ��"******
  qtApp.Test.Settings.Run.StartIteration=objDic("RowNum"&k)
  qtApp.Test.Settings.Run.EndIteration=objDic("RowNum"&k)
  qtApp.Options.Run.RunMode = "Fast"
  qtApp.Test.Run
  '�ж�ErrLog.xml�Ƿ���ڣ��Ƿ�����Case
  Call ErrLogIsExist(casePath_prefix, objDic("CaseName"&k),objDic("RowNum"&k), xmlPath)
  '��ȡ��¼�Ĳ��Խ����д������excel�����!!!!!!����excelд��ǰһ��Ҫǩ�������أ�����Ϊֻ�����޷�д����!!!!!!!!
  Res= ReadCaseStatus
  Call WriteExcelResults(objDic("CaseName"&k),objDic("RowNum"&k),Res)  
Next

'/*****************************��ȡ����excel�еĲ����������������������к�****************************************/
Function GetExcelPara()
  Dim oDict
  '�����ֵ䣬�洢�˺������صĶ�ֵ
  Set oDict=CreateObject("Scripting.Dictionary")                        
  '����excelӦ��
  Set excelApp=createobject("excel.application")
  '��workbook
  excelApp.Workbooks.Open(VbsPath&"\AutoTestCase.xlsx") '�������ĵ�AutoTestCase.xlsx
  '��ȡsheet
  set oSheet=excelApp.Sheets(1)   '��ȡ��X��Sheet
  Dim RowCount
  RowCount=oSheet.UsedRange.Rows.Count '��ȡexcel������
  Dim i 'excel�б�
  Dim j 'excel�б�
  j=1
  For i=1  to RowCount
    if(oSheet.cells(i,5)="Y")then  'Ŀǰ�����Ƿ�ִ�еı�־�ǹ̶��ĵ�6�У��粻�̶�����Ҫ��Ϊ��̬����
	oDict("CaseName"& j)=oSheet.cells(i,1)
	oDict("RowNum"& j)=oSheet.cells(i,2)
	j=j+1
    end if
  Next
  '�����ֵ�
  Set GetExcelPara=oDict
  '�رյ�ǰ������
  excelApp.ActiveWorkbook.Close
  '�ر�excel
  excelApp.Quit
  '�ͷ�excel����
  Set excelApp=nothing
End Function
'/*****************************��ȡ������ǰ�����ļ���״̬��¼****************************************/
Function ReadCaseStatus
   Const ForReading = 1
	if(fso.FileExists(StatusTxtPath))then
		Set txtFile=fso.OpenTextFile(StatusTxtPath,ForReading,true)
		ReadCaseStatus=txtFile.ReadAll
		txtFile.close
		Set txtFile=nothing
	else
		msgbox "������ǰ�����ļ���״̬��¼�ļ�������"
	end if
End Function

'/*****************************���������������������кŶ�Ӧ�����н������������excel�����****************************************/
Sub WriteExcelResults(currCaseName,currRowNum,LastRes)        
   if(LastRes="0")then
   LastRes="Success"
   end if
   if(LastRes="1")then
   LastRes="Failed"
   end if
  '����excelӦ��
  Set excApp=createobject("excel.application")
  '��workbook
  excApp.Workbooks.Open(VbsPath & "\AutoTestCase.xlsx") '�������ܵ�EXCEL�ĵ���ַAutoTestCase.xlsx
  '��ȡsheet
  set oSheet=excApp.Sheets(1)     '��ȡ��X��Sheet
  Dim RowCount
  RowCount=oSheet.UsedRange.Rows.Count
  For t=1  to RowCount            'iΪexcel���������к�
    if(oSheet.cells(t,1)=currCaseName and oSheet.cells(t,2)=currRowNum)then  'Ŀǰ�����Ƿ�ִ�еı�־�ǹ̶��ĵ�6�У��粻�̶�����Ҫ��Ϊ��̬����
       oSheet.cells(t,4)=LastRes  '����excel�н����Ϊ��5��
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
'/*****************************�ж�ErrLog.xml�Ƿ������������case****************************************/
Sub ErrLogIsExist(casePath_prefix, caseName,currRowNum,xmlPath)
    '����oFSO����
    Dim oFSO
    set oFSO=createObject("Scripting.FileSystemObject")
    if(oFSO.FileExists(xmlPath))then
      set dictXML=GetXml(xmlPath,caseName)
      '�����ȡxml���صĶ���Ϊ�գ����ȡ���ص�CurrentRow��RowCount
      if(dictXML.count>0)then
	if(dictXML("CurrentRow")>0)then
          Call IsRunCaseAgain(casePath_prefix, caseName,currRowNum)
	  '�ͷ���Դ
	  set dictXML=nothing
	end if
      end if
    end if
    '�ͷ���Դ
    Set oFSO=nothing
End Sub
'/*****************************�жϵ�ǰ�����Ƿ���ڼ���Fail����Ҫ��������һ�ε�ǰ����****************************************/
Sub IsRunCaseAgain(casePath_prefix, caseName,currRowNum)
  '�����¼��������״̬��txt�ļ�
   if(fso.FileExists(StatusTxtPath))then
     fso.DeleteFile(StatusTxtPath)
   end if
   '����case
   qtApp.Open casePath_prefix & caseName
   qtApp.Test.Settings.Run.StartIteration=currRowNum
   qtApp.Test.Settings.Run.EndIteration=currRowNum
   qtApp.Options.Run.RunMode = "Fast"
   qtApp.Test.Run
End Sub
'/*********************************��ȡErrLog.xml********************************************/
Function GetXml(strXmlFilePath,caseNameVal)
  Dim oDict
  Set oDict=CreateObject("Scripting.Dictionary")        '�����ֵ䣬�洢�˺������صĶ�ֵ
  Dim xmlDoc,xmlRoot
  Set xmlDoc = CreateObject("MSXML2.DOMDocument")    	'����XML DOM����
  xmlDoc.async = False                                  '���Ƽ���ģʽΪͬ��ģʽ��xml��������Ϻ���ִ�к������룩                     
  xmlDoc.load strXmlFilePath                            '����xml�ļ�       
  If xmlDoc.parseError.errorCode <> 0 Then
     MsgBox "XML�ļ���ʽ���ԣ�ԭ���ǣ�" & Chr(13) & xmlDoc.parseError.reason
     Exit Function
  End If
  '��ȡ�����
  Set xmlRoot = xmlDoc.documentElement
  '����xml
  For i=0 to  xmlRoot.childNodes.length-1
      if(xmlRoot.childNodes(i).attributes(0).text=caseNameVal)then
	  For j=0 to xmlRoot.childNodes(i).childNodes.length-1
	     For p=0 to xmlRoot.childNodes(i).childNodes(j).childNodes.length-1
		 if(xmlRoot.childNodes(i).childNodes(j).selectSingleNode("ErrCode").text<>0)then
		    '����ǰ�кź��������������ڵ�ֵ�洢���ֵ���
		    oDict("CurrentRow")=xmlRoot.childNodes(i).childNodes(j).selectSingleNode("CurrentRow").text
		    oDict("RowCount")=xmlRoot.childNodes(i).childNodes(j).selectSingleNode("RowCount").text
		    set GetXml = oDict
		    Exit for  '�ҵ���ͬ�汾�ź��˳������ڵ�ѭ����ֹͣ����
		 end if 
	     Next
	     Exit For   'ֱ���˳������ڵ�ѭ��������ֵ
	  next
	  Exit for  '�ҵ���ͬ�����������˳����ѭ����ֹͣ����
     else
	  Set GetXml=oDict
     end if
  Next
End Function
'�ͷ�fso����
set fso=nothing
'�˳�QTP
qtApp.Quit
'�ͷ�qtpӦ�ö���
Set qtApp=Nothing
