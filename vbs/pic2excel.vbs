'----------------------------------------------------
' 用excel画像素图, 
'  使用方法，将图片拖到本脚本上
'                   Code by C.Y 2014-08-10
'----------------------------------------------------

Dim Img 'As ImageFile
Dim ImgFileName
Dim IP 'As ImageProcess
Dim v 'As Vector

'图片的长宽以及绘制时使用的长和宽
Dim Width,Height,Rw,Rh 

'指定要绘制的像素图的最大宽度(“像素”格子数)，
Rw = 200    '<== 目前的代码该值不能大于 26 * 26 = 676

Dim ExcelApp  'Excel Application
Dim ExcelBook 'Excel workbook
Dim ExcelSheet 'Excel sheet
Dim Fso
Dim currentPath '当前目录路径，用于保存生成的excel文件

Set Img = CreateObject("WIA.ImageFile")
Set IP = CreateObject("WIA.ImageProcess")

'从命令行参数中获取待处理图片名
If WScript.Arguments.Count>0 Then
   ImgFileName =  WScript.Arguments(0)
Else 
   wscript.echo "请将要转换的图片拖到本程序的图标上"
   wscript.quit
End If

' WIA组件 载入图片
Img.LoadFile ImgFileName
Width = Img.Width
Height = Img.Height
'wscript.echo "图片大小" & Width & "px X " & Height & "px"
'wscript.echo "图片类型" & Img.FileExtension

'绘制高度 Rh 根据实际图片纵横比进行计算,这里使用整除
Rh = Rw * Height \ Width 

If Width > Rw Then
    '图片实际尺寸大于绘制尺寸时，缩放图片
    IP.Filters.Add IP.FilterInfos("Scale").FilterID
    IP.Filters(1).Properties("MaximumWidth") = Rw
    IP.Filters(1).Properties("MaximumHeight") = Rh
    Set Img = IP.Apply(Img)
'下面一句结果表明缩放时候设置的长宽并不一定就是缩放后真实的长宽，所以 Rw Rh 需要重新从缩放后的Img对象获取
'MsgBox "绘制尺寸: " & Rw & "px X " & Rh & "px" & vbcrlf & img.width & " x " & img.height
    Rw = Img.Width
    Rh = Img.Height
Else
    ' 小于预定尺寸的，按实际大小画出
    Rw = Width
    Rh = Height
End If


' 获取图片所有像素值,每个元素是一个aRgb值
Set v = Img.ARGBData
'MsgBox "像素数:" & v.Count
Wscript.Echo "点击确定后开始后台绘图,需要绘制"& v.Count &"个单元格" & vbCrLf & "绘制时间与图片大小有关，绘图结束会弹出提示"

Set ExcelApp = CreateObject("Excel.Application")
Set Fso = CreateObject("Scripting.FileSystemObject")
currentPath = Fso.GetFile(Wscript.ScriptFullName).ParentFolder.Path & "\"

'不显示提示信息,这样保存的时候就不会提示是否要覆盖原文件
ExcelApp.DisplayAlerts=FALSE
'调用EXCEL文件的时候不显示
ExcelApp.visible=FALSE

Set ExcelBook = ExcelApp.workbooks.Add '新建一个excel workbookbook
Set ExcelSheet = ExcelApp.Sheets.Item(1)
'Msgbox Rw & " " & Rw/26 & " " & Chr(Asc("a")+(Rw\26)-1)

'行高，列宽数值都是取的本机设置为5px的数字，excel中高和宽用的单位不同，
'可能不适用与所有机器，有待研究
ExcelSheet.range("a:"& Chr(Asc("a")+(Rw\26)-1) &  Chr(Asc("a")+(Rw Mod 26)-1)).ColumnWidth = Round(0.77/2 ,2)
ExcelSheet.range("a1:a"& Rh).RowHeight =Round( 7.50/2 ,2) 

'逐个像素点绘制
For i = 1 To v.Count
   aRGBstr= Hex(v(i))
   cr =  Mid(aRGBstr,3,2)
   cg =  Mid(aRGBstr,5,2)
   cb =  Mid(aRGBstr,7,2)
   ExcelSheet.Cells( ((i-1)\Rw)+1, ((i-1) Mod Rw)+ 1).Interior.color = RGB("&h"&cr, "&h"&cg,"&h"&cb)
Next

'文件另存为
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

Wscript.Echo "处理完成"