
PPTpath = "C:\Users\Lufeng\Desktop\PPT模板\1-200"
PPTname = "31页宽屏清新商务PPT图表.ppt"
//fileHandle = Plugin.File.OpenFile(PPTpath & "\" & PPTname)

//Call RunApp("C:\Users\Lufeng\Desktop\PPT模板\1-200\31页宽屏清新商务PPT图表.ppt")

//Call RunApp(PPTpath & "\" & PPTname)


Call 遍历文件夹内文件(PPTpath)

Function 遍历文件夹内文件(Ppath)


//注意：返回的是数组变量，存储着每一个文件名。
数组 = lib.文件.遍历指定目录下所有文件名(Ppath)
For i=0 to UBound(数组)-1
TracePrint "目前文件是" & 数组(i)
//取文件扩展名

s = Split(数组(i), ".")
扩展名 = s(1)

TracePrint "扩展名" & 扩展名
TracePrint "判断" & StrComp(扩展名,"ppt",1)

If 	Instr(扩展名,"ppt")>0 or Instr(扩展名,"pptx")>0 Then 
Delay 1000
WholePath = Ppath & "\" & 数组(i)
Call 打开文件(WholePath)
Call 操作()
Call 关闭文件()
Call 移动文件(Ppath,"C:\Users\Lufeng\Desktop\PPT模板\1-200\删完",数组(i))
//Call 导出视频()


End If    


Next

End Function




Function 打开文件(WholePath)

Call RunApp(WholePath)

TracePrint "打开文件"
Delay 3000
// window max(press alt + space then press X)
KeyDown 18, 1
KeyPress "Space", 1
KeyUp 18, 1
Delay 1000
KeyPress "X", 1
Delay 1000
TracePrint "最大化"

End Function

Function 操作 ()
//判断文件是否打开
X = - 2 
Y = - 2 


Do While (X<0 and Y<0)
TracePrint "X|Y="&X&"|"&Y
XY=Plugin.Color.FindMutiColor(7,168,26,191,"4D68D6","1|6|4D68D6",1)
dim MyArray
MyArray = Split(XY, "|")
X = CInt(MyArray(0)): Y = CInt(MyArray(1))

If X > 0 and Y > 0 Then 
MsgBox "文件已经打开，坐标位置：" & X & "," & Y

Exit Do	
End If

Loop


Delay 1000
KeyPress "End", 1
Delay 1000
delay 1000
delay 1000

//判断最后一页是不是YP,如果是，就删除
AB=Plugin.Color.FindMutiColor(34,810,224,957,"F0B000","-3|22|50D092,-3|22|50D092,-24|16|289461,72|10|A87C00",1)
dim MyArray1
MyArray1 = Split(AB, "|")
A = CInt(MyArray1(0)): B = CInt(MyArray1(1))
	
	If A > 0 and B > 0 Then 
	MsgBox "是YP页面,坐标是" & A & "," & B
	Delay 1000
	

	MoveTo 34,810
	LeftClick 1
	delay 500
	LeftClick 1
	
	delay 1000
	KeyPress "Delete", 1
	
	End If





End Function

Function 关闭文件()
'//Press  Ctrl+S
KeyDown 17, 1
KeyPress 83, 1
KeyUp 17, 1
Delay 1000
'//Press  Alt + F4
KeyDown 18, 1
KeyPress 115, 1
KeyUp 18, 1
Delay 1000
TracePrint "关闭了文件"
End Function

Function 移动文件(Opath,TargetPath,Filename)
delay 1000
Call Plugin.File.MoveFile(Opath & "\" & Filename, Targetpath & "\" & Filename)
TracePrint "移动了文件"
End Function
