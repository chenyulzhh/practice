'----------------------------------------------------
' ��excel������ͼ, 
'  ʹ�÷�������ͼƬ�ϵ����ű���
'                   Code by C.Y 2014-08-10
'----------------------------------------------------

Dim Img 'As ImageFile
Dim ImgFileName
Dim IP 'As ImageProcess
Dim v 'As Vector

'ͼƬ�ĳ����Լ�����ʱʹ�õĳ��Ϳ�
Dim Width,Height,Rw,Rh 

'ָ��Ҫ���Ƶ�����ͼ�������(�����ء�������)��
Rw = 200    '<== Ŀǰ�Ĵ����ֵ���ܴ��� 26 * 26 = 676

Dim ExcelApp  'Excel Application
Dim ExcelBook 'Excel workbook
Dim ExcelSheet 'Excel sheet
Dim Fso
Dim currentPath '��ǰĿ¼·�������ڱ������ɵ�excel�ļ�

Set Img = CreateObject("WIA.ImageFile")
Set IP = CreateObject("WIA.ImageProcess")

'�������в����л�ȡ������ͼƬ��
If WScript.Arguments.Count>0 Then
   ImgFileName =  WScript.Arguments(0)
Else 
   wscript.echo "�뽫Ҫת����ͼƬ�ϵ��������ͼ����"
   wscript.quit
End If

' WIA��� ����ͼƬ
Img.LoadFile ImgFileName
Width = Img.Width
Height = Img.Height
'wscript.echo "ͼƬ��С" & Width & "px X " & Height & "px"
'wscript.echo "ͼƬ����" & Img.FileExtension

'���Ƹ߶� Rh ����ʵ��ͼƬ�ݺ�Ƚ��м���,����ʹ������
Rh = Rw * Height \ Width 

If Width > Rw Then
    'ͼƬʵ�ʳߴ���ڻ��Ƴߴ�ʱ������ͼƬ
    IP.Filters.Add IP.FilterInfos("Scale").FilterID
    IP.Filters(1).Properties("MaximumWidth") = Rw
    IP.Filters(1).Properties("MaximumHeight") = Rh
    Set Img = IP.Apply(Img)
'����һ������������ʱ�����õĳ�����һ���������ź���ʵ�ĳ������� Rw Rh ��Ҫ���´����ź��Img�����ȡ
'MsgBox "���Ƴߴ�: " & Rw & "px X " & Rh & "px" & vbcrlf & img.width & " x " & img.height
    Rw = Img.Width
    Rh = Img.Height
Else
    ' С��Ԥ���ߴ�ģ���ʵ�ʴ�С����
    Rw = Width
    Rh = Height
End If


' ��ȡͼƬ��������ֵ,ÿ��Ԫ����һ��aRgbֵ
Set v = Img.ARGBData
'MsgBox "������:" & v.Count
Wscript.Echo "���ȷ����ʼ��̨��ͼ,��Ҫ����"& v.Count &"����Ԫ��" & vbCrLf & "����ʱ����ͼƬ��С�йأ���ͼ�����ᵯ����ʾ"

Set ExcelApp = CreateObject("Excel.Application")
Set Fso = CreateObject("Scripting.FileSystemObject")
currentPath = Fso.GetFile(Wscript.ScriptFullName).ParentFolder.Path & "\"

'����ʾ��ʾ��Ϣ,���������ʱ��Ͳ�����ʾ�Ƿ�Ҫ����ԭ�ļ�
ExcelApp.DisplayAlerts=FALSE
'����EXCEL�ļ���ʱ����ʾ
ExcelApp.visible=FALSE

Set ExcelBook = ExcelApp.workbooks.Add '�½�һ��excel workbookbook
Set ExcelSheet = ExcelApp.Sheets.Item(1)
'Msgbox Rw & " " & Rw/26 & " " & Chr(Asc("a")+(Rw\26)-1)

'�иߣ��п���ֵ����ȡ�ı�������Ϊ5px�����֣�excel�иߺͿ��õĵ�λ��ͬ��
'���ܲ����������л������д��о�
ExcelSheet.range("a:"& Chr(Asc("a")+(Rw\26)-1) &  Chr(Asc("a")+(Rw Mod 26)-1)).ColumnWidth = Round(0.77/2 ,2)
ExcelSheet.range("a1:a"& Rh).RowHeight =Round( 7.50/2 ,2) 

'������ص����
For i = 1 To v.Count
   aRGBstr= Hex(v(i))
   cr =  Mid(aRGBstr,3,2)
   cg =  Mid(aRGBstr,5,2)
   cb =  Mid(aRGBstr,7,2)
   ExcelSheet.Cells( ((i-1)\Rw)+1, ((i-1) Mod Rw)+ 1).Interior.color = RGB("&h"&cr, "&h"&cg,"&h"&cb)
Next

'�ļ����Ϊ
ExcelBook.SaveAs(currentPath & Fso.GetBaseName(ImgFileName) & ".xlsx") 
ExcelBook.Close
ExcelApp.quit

Set ExcelSheet = Nothing
Set ExcelBook = Nothing
Set ExcelApp = Nothing
Set Fso = Nothing
Set v = Nothing
Set IP = Nothing
Set Img = Nothing

Wscript.Echo "�������"