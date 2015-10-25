
<%
'当表单里既有文本域又有文件域的时候,我们必须把表单的编码类型设置成"multipart/form-data"类型
'这时候上传上来的编码文件并不能直接取出文本域的值和文件域的二进制数据,这就需要拆分表单域
'在上传上来的数据流中在每个表单域间都有一个随机的分隔符,这个分隔符是在同一个流中不变的,不同的流分隔符不变,
'这个分隔符在流的最开头,并且以一个chrb(13) + chrb(10)结束,知道这个后我们就可以用这个分隔符来遍历拆分表单域了.
'对于文件域,我们要解析字段名,文件名,文件类型和文件内容,域名是以"name="为前导,并包含在一对双引号中,文件名的值是以"filename="为前导,也包含在双引号里,其中包含文件的全路径和文件名,紧跟着后面又是一对回车换行府(chrb(13) +chrb(10)),字符串"content-type:"和两对回车换行之间的内容为文件类型字符串,两对回车换行后到一对回车换行之间的数据为文件内容
'对于文本域,我们只要解析他的值就可以了,域的名称是以"name="之后,用双引号包着,两对回车换行后到以一对回车换行开始的域分隔符之间为该文本域的值
'当然上传上来的流是二进制格式,在操作的时候需要用一些操作二进制的函数,而不是平时用的操作字符串的函数,比如说leftB,midB,instrB等,下面就是算法的实现
Class GetPostImg
	private BdataStr,SeparationStr,wawa_stream		'提交的信息,表单域间分隔字符
	'类初始化
	Private Sub Class_Initialize
	   set wawa_stream=CreateObject("Adodb.Stream")	'创建全局流
	   wawa_stream.mode=3	'读写模式
	   wawa_stream.type=1	'二进制读取模式 
	   wawa_stream.open		'打开流
	   BdataStr=Request.BinaryRead(Request.TotalBytes)'获取上传的所有数据
	   wawa_stream.write BdataStr	'读取数据   
	   SeparationStr=LeftB(BdataStr,CokeShow.CokeClng(inStrb(BdataStr,ChrB(13) + ChrB(10)))-1)	'分隔字符串
	End Sub
	'类的析构函数,卸载全局流对象
	Private Sub Class_Terminate          
	   wawa_stream.close
		 set wawa_5xSoft_Stream=nothing
	End Sub
	'返回file型表单域的值(二进制)
	Public Function GetFile (FieldName)
	   Dim L1,DataStart,DataLng
	   L1 = InStrB(BdataStr,GetBinary("name=" + Chr(34) +FieldName +Chr(34)))
	   DataStart = InStrB(L1,BdataStr,ChrB(13) + ChrB(10) + ChrB(13) + ChrB(10)) +4
	   DataLng = InStrB(DataStart,BdataStr,SeparationStr) - DataStart -2
	   GetFile =MidB(BdataStr,DataStart,DataLng)
	End Function
	'返回文件的类型
	Public Function GetFileType (FieldName)
	   Dim L1,DataStart,DataLng
	   L1 = InStrB(BdataStr,GetBinary("name=" + Chr(34) +FieldName +Chr(34)))
	   DataStart = InStrB(L1,BdataStr,GetBinary("Content-Type:")) + 13
	   DataLng = InStrB(DataStart,BdataStr,ChrB(13) + ChrB(10) + ChrB(13) + ChrB(10)) - DataStart
	   GetFileType =GetText(MidB(BdataStr,DataStart,DataLng))
	End Function
	'返回文件的原始路径
	Public Function GetFilePath (FieldName)
	   Dim L1,DataStart,DataLng
	   L1 = InStrB(BdataStr,GetBinary("name=" + Chr(34) +FieldName +Chr(34)))
	   DataStart = InStrB(L1,BdataStr,GetBinary("filename=")) + 9
	   DataLng = InStrB(DataStart,BdataStr,ChrB(13) + ChrB(10)) - DataStart
	   GetFilePath = GetText(MidB(BdataStr,DataStart+1,DataLng-2))		'去掉最左边和最右边的双引号,不知道为什么右边的双引号要减去2
	End Function
	'返回原始文件的后缀名
	Function GetExtendName(FieldName)
		
	   FileName = GetFilePath(FieldName)
	   If isNull(FileName) or FileName="" Then
		GetExtendName=""
		Exit Function
	   End If
	   GetExtendName = Mid(FileName,InStrRev(FileName, "."))
	End Function
	'返回file型表单域的值(二进制)
	Public Function GetFileSize (FieldName)
	   Dim L1,DataStart,DataLng
	   L1 = InStrB(BdataStr,GetBinary("name=" + Chr(34) +FieldName +Chr(34)))
	   DataStart = InStrB(L1,BdataStr,ChrB(13) + ChrB(10) + ChrB(13) + ChrB(10)) +4
	   DataLng = InStrB(DataStart,BdataStr,SeparationStr) - DataStart -2
	   GetFileSize = DataLng
	End Function
	'从二进制字符串里取出表单域的值(字符串)
	Public Function RetFieldText (FieldName)
	   Dim L1,DataStart,DataLng
	   L1 = InStrB(BdataStr,GetBinary("name=" + Chr(34) +FieldName +Chr(34)))
	   DataStart = InStrB(L1,BdataStr,ChrB(13) + ChrB(10) + ChrB(13) + ChrB(10)) +4
	   DataLng = InStrB(DataStart,BdataStr,SeparationStr) - DataStart -2
	   RetFieldText =GetText(MidB(BdataStr,DataStart,DataLng))
	End Function
	'返回一个时间和随机数连接后的字符串,用于构建文件名
	Function getrandStr()
	   Dim RanNum
	   Randomize
	   RanNum = Int(90000*rnd)+10000
	   getrandStr = Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&RanNum
	End Function
	
	'将二进制外码系列转换成vb字符串
	Private Function GetText (Str1r)
	   Dim s,t,t1,i
	   s = "":t="":t1=""
	   For i =1 To LenB(str1r)
		t= AscB(MidB(Str1r,i,1))	'按字节取出外码
		if not(t > 127) Then		'字节高位为0,表示英文字符
		 s = s + Chr(t)
		Else
		 i = i +1					'当为汉字时,取第二个字节
		 t1 = AscB(MidB(Str1r,i,1)) 
		 s = s + Chr(t * 256 + t1)	'将汉字两字节外码组合成ANSI码
		End If
	   Next
	   GetText = s
	End Function
	'将字符串转换为二进制系列
	Private Function GetBinary(str1)
		Dim i
	   Dim T2,t1
	   For i = 1 To Len(Str1)
		t1 = CStr(Hex(Asc(Mid(Str1,i,1))))
		If Len(t1)=2 Then
		 T2 = T2 + ChrB(CokeShow.CokeClng("&h" + Trim(t1)))
		Else
		 T2 = T2 + ChrB(CokeShow.CokeClng("&H") + Mid(Trim(t1),1,2))
		 T2 = T2 + ChrB(CokeShow.CokeClng("&H") + Mid(Trim(t1),3,2))
		End If
	   Next
	   GetBinary = T2
	End Function
	'将上传的文件保存在服务器的硬盘上
	Public Function SaveToFile (FieldName,fullpath)
	   
	   dim dr		'定义创建一个流
	   SaveToFile=""
	   if trim(fullpath)="" or FileName="" then exit function	'检测参数是否有真实数据
	   if right(fullpath,1)="/" then exit function				'检测路径的正确性
	   set dr=CreateObject("Adodb.Stream")
	   dr.Mode=3	'读写模式
	   dr.Type=1	'二进制模式
	   dr.Open		'打开
	   Dim L1,DataStart,DataLng
	   L1 = InStrB(BdataStr,GetBinary("name=" + Chr(34) +FieldName +Chr(34)))			'获取file域的位置
	   DataStart = InStrB(L1,BdataStr,ChrB(13) + ChrB(10) + ChrB(13) + ChrB(10)) +4		'实体数据的开始位置
	   DataLng = InStrB(DataStart,BdataStr,ChrB(13) + ChrB(10) + ChrB(13) + ChrB(10)) - DataStart	'实体数据的大小
	   wawa_stream.position=DataStart-1		'设置全局流的游标,因为全局流和全局数据BdataStr对应的
	   wawa_stream.copyto dr,DataLng		'从全局流里获取数据
	   dr.SaveToFile FullPath,2				'保存在指定位置
	   dr.Close			'关闭流
	   set dr=nothing	'析构流
	   SaveToFile=Mid(FileName,InStrRev(FileName, "\")+1)	'返回上传文件的文件名
	End Function 
End Class


%>