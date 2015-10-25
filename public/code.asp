<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'Option Explicit
'Response.Buffer = True	'缓冲数据，才向客户输出；
Response.Charset="utf-8"
Session.CodePage = 65001
%>
<%
'------------------
'图片验证码
'netdust
'http://www.oyj.cn
'05-8-29 发布
'以前的图片验证码很容易被破解，所以在其基础上进行改进，即使知道源代码
'也很难破解。
'干扰次数每加一  ->  穷举法的代价提高约100*100倍
'适当设定杂点几率可以有效降低软件识别正确率
'05-12-20 修正了一个错误
'------------------

'Option Explicit

Call Com_CreatValidCode("GetCode")

Rem 生成验证码图片
Sub Com_CreatValidCode(pSN)
	Const codeLen = 4						'验证码位数
	Const cOdds = 3 						'杂点出现的机率
	Const dbtTimes = 1					'干扰次数（安全考虑，最好不要小于3）

	Const cAmount = 10 					'字库数量
	Const cCode = "0123456789"	'字库对应的字符
	Const UnitWidth = 16				'字宽(要为4的倍数)
	Const UnitHeight = 13				'字高
	Const DotsLimit = 5				'每次删除有效点的上限(避免无法人为识别)
	Const tryCount = 5					'避免删除有效点超过上限的尝试次数限制

	'-----------

	Randomize
	Dim i, ii, iii

	' 禁止缓存
	Response.Expires = -9999
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "cache-ctrol","no-cache"
	Response.ContentType = "Image/BMP"

	' 颜色的数据(字符，背景)
	Dim vColorData(1)
	vColorData(0) = ChrB(0) & ChrB(0) & ChrB(0)  ' 蓝0，绿0，红0（黑色）
	vColorData(1) = ChrB(255) & ChrB(255) & ChrB(255) ' 蓝250，绿236，红211（浅蓝色）

	' 字符的数据(可以自己修改，如果修改了尺寸，记得把前面的设定也改了)
	Dim vNumberData(9)
	vNumberData(0) = "1111000000001111111000000000011111100111111001111110011111100111111001111110011111100111111001111110011111100111111001111110011111100111111001111110011111100111111001111110011111100000000001111111000000001111"
	vNumberData(1) = "1111110001111111111100000111111111100000011111111100110001111111111111000111111111111100011111111111110001111111111111000111111111111100011111111111110001111111111111000111111111100000000011111110000000001111"
	vNumberData(2) = "1111110000011111111110000000111111110001110011111110001111001111111111111001111111111111001111111111111001111111111111001111111111111001111111111111001111001111111001111100111111100000000011111110000000001111"
	vNumberData(3) = "1111100000011111111100000000111111100111111001111110011111001111111111111001111111111110001111111111111000111111111111111001111111111111110011111110011111100111111001111110011111110000000011111111100000011111"
	vNumberData(4) = "1111111100111111111110110011111111110011001111111111001100111111111001110011111111001111001111111000000000000011100000000000001111111111001111111111111100111111111111110011111111111111001111111111111100111111"
	vNumberData(5) = "1110000000000111110011111111111111001111111111111100111111111111110011111111111111001100000011111100000111100111111111111110011111111111111001111111111111100111110011111110011111001111111001111110000000001111"
	vNumberData(6) = "1111110000011111111110000000111111110011111001111110011111111111111001111111111111100100000111111110000000001111111000111110011111100111111001111110011111100111111001111110011111110000000011111111100000011111"
	vNumberData(7) = "1110000000000111111000000000011111100111111001111110011111100111111111111100111111111111110011111111111110011111111111110011111111111111001111111111111100111111111111110011111111111111001111111111111100111111"
	vNumberData(8) = "1111100000011111111100000000111111100111111001111110011111100111111001111110011111110000000011111111000000001111111100111100111111100111111001111110011111100111111001111110011111110000000011111111100000011111"
	vNumberData(9) = "1111100000011111111100000000111111100111111001111110011111100111111001111110011111110000000001111111000000100111111111111110011111111111111001111111111111100111111001111100111111110000000011111111100000011111"

	' 随机产生字符
	Dim vCodes
	ReDim vCode(codeLen-1)
	For i = 0 To codeLen-1
	  vCode(i) = Int(Rnd * cAmount)
	  vCodes = vCodes & Mid(cCode, vCode(i) + 1, 1)
	  vCode(i) = pcd_doubter(vNumberData(vCode(i)),UnitWidth,UnitHeight,DotsLimit,tryCount,dbtTimes)
	Next

	Session(pSN) = vCodes  '记录入Session

	' 输出图像文件头
	Response.BinaryWrite ChrB(66) & ChrB(77) & Num2ChrB(54+UnitWidth*UnitHeight*CodeLen*3,4) & ChrB(0) & ChrB(0) &_
	  ChrB(0) & ChrB(0) & ChrB(54) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(40) & ChrB(0) &_
	  ChrB(0) & ChrB(0) & Num2ChrB(UnitWidth*CodeLen,4) & Num2ChrB(UnitHeight,4) &_
	  ChrB(1) & ChrB(0)

	' 输出图像信息头
	Response.BinaryWrite ChrB(24) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & Num2ChrB(UnitWidth*UnitHeight*CodeLen*3,4) &_
	  ChrB(18) & ChrB(11) & ChrB(0) & ChrB(0) & ChrB(18) & ChrB(11) &_
	  ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) &_
	  ChrB(0) & ChrB(0)

	For i = UnitHeight-1 To 0 Step -1  ' 历经所有行
	  For ii = 0 To codeLen-1  ' 历经所有字
	   For iii = 1 To UnitWidth ' 历经所有像素
	    If Rnd * 99 + 1 >= cOdds Then	' 逐行、逐字、逐像素地输出图像数据
	     Response.BinaryWrite vColorData(Mid(vCode(ii), i * UnitWidth + iii, 1))
	    Else ' 随机生成杂点
	     Response.BinaryWrite vColorData(1-Cint(Mid(vCode(ii), i * UnitWidth + iii, 1)))
	    End If
	   Next
	  Next
	Next
End Sub

Rem 对单个字的点阵进行干扰
Rem 干扰思想：在点阵范围内随机产生2个端点，进行连线，以位移较大的一方做横轴，先将连线上的点删除，再将被删除点的纵轴方向上方或下方的点（随机确定）移向被删除点，移动后的空白用背景色补充
Function pcd_doubter(ByVal str,UnitWidth,UnitHeight,DotsLimit,tryCount,dbtTimes)
	Randomize
	Dim x1,x2,y1,y2,xOffSet,yOffSet,direction,flag,rows,step,yu,yuStr,i,ii,iii,f1,f2
	For f1=1 To dbtTimes	'干扰次数
		For f2=1 To tryCount	'避免删除有效点超过上限的尝试次数限制
			'随机确定2个端点
			x1 = int(Rnd*UnitWidth)
			x2 = int(Rnd*UnitWidth)
			y1 = int(Rnd*UnitHeight)
			y2 = int(Rnd*UnitHeight)
			'x,y位移量
			xOffSet = Abs(x2-x1)
			yOffSet = Abs(y2-y1)
			If xOffSet >= yOffSet Then	'以位移量较大方做横轴
				direction = "x"
				ReDim ary(xOffSet)	'用来记录连线各点y值
				'x2,y2存储x值较大的点
				If x2 < x1 Then
					i = x1
					x1 = x2
					x2 = i
					i = y1
					y1 = y2
					y2 = i
				End If
				'判断从x1->x2在纵轴方向上是增是减
				If y2 >= y1 Then
					flag = 1
				Else
					flag = -1
				End If
				'下面计算连线上点的分布（先是平均分配各行的点，然后随机分配剩余的点到各行）
				rows = yOffSet + 1	'所占行数
				step = (xOffSet+1) \ rows	'各行平均分配的点
				yu = (xOffSet+1) Mod rows	'剩余的点数
				ReDim ary2(rows-1)	'用来记录剩余点的随机分配
				While yu > 0
					i = int(Rnd*rows)
					ary2(i) = ary2(i)&"."		'被分配到的行则加一个字符"."
					yu = yu - 1
				WEnd
				iii = 0
				'将连线的点信息记录到数组
				For i=0 To rows-1
					For ii=1 To step+Len(ary2(i))
						ary(iii) = y1 + i*flag
						iii = iii + 1
					Next
				Next
				ii = 0
				'统计连线上有效点的数量
				For i=0 To xOffSet
					If pcd_getDot(x1+i,ary(i),str,UnitWidth) = "0" Then ii = ii + 1
				Next
			Else
				'这里是以y为横轴，原理与x时相同
				direction = "y"
				ReDim ary(yOffSet)
				If y2 < y1 Then
					i = x1
					x1 = x2
					x2 = i
					i = y1
					y1 = y2
					y2 = i
				End If
				If x2 >= x1 Then
					flag = 1
				Else
					flag = -1
				End If
				rows = xOffSet + 1
				step = (yOffSet+1) \ rows
				yu = (yOffSet+1) Mod rows
				ReDim ary2(rows-1)
				While yu > 0
					i = int(Rnd*10)
					If i < rows Then
						ary2(i) = ary2(i)&"."
						yu = yu - 1
					End If
				WEnd
				iii = 0
				For i=0 To rows-1
					For ii=1 To step+Len(ary2(i))
						ary(iii) = x1 + i*flag
						iii = iii + 1
					Next
				Next
				ii = 0
				For i=0 To yOffSet
					If pcd_getDot(ary(i),y1+i,str,UnitWidth) = "0" Then ii = ii + 1
				Next
			End If
			'如未超过有效点上限则跳出循环，执行干扰
			If ii <= DotsLimit Then Exit For
		Next

		If direction = "x" Then
			'随机确定在纵轴方向上或下进行移动
			If int(Rnd*10) > 4 Then
				'变量连线上的点
				For i=0 To xOffSet
					'遍历移动
					For ii=ary(i) To 1 Step -1
						Call pcd_setDot(x1+i,ii,str,pcd_getDot(x1+i,ii-1,str,UnitWidth),UnitWidth)
					Next
					'添补空白
					Call pcd_setDot(x1+i,0,str,"1",UnitWidth)
				Next
			Else
				For i=0 To xOffSet
					For ii=ary(i) To UnitHeight-2
						Call pcd_setDot(x1+i,ii,str,pcd_getDot(x1+i,ii+1,str,UnitWidth),UnitWidth)
					Next
					Call pcd_setDot(x1+i,UnitHeight-1,str,"1",UnitWidth)
				Next
			End If
		Else
			If int(Rnd*10) > 4 Then
				For i=0 To yOffSet
					For ii=ary(i) To 1 Step -1
						Call pcd_setDot(ii,y1+i,str,pcd_getDot(ii-1,y1+i,str,UnitWidth),UnitWidth)
					Next
					Call pcd_setDot(0,y1+i,str,"1",UnitWidth)
				Next
			Else
				For i=0 To yOffSet
					For ii=ary(i) To UnitWidth-2
						Call pcd_setDot(ii,y1+i,str,pcd_getDot(ii+1,y1+i,str,UnitWidth),UnitWidth)
					Next
					Call pcd_setDot(UnitWidth-1,y1+i,str,"1",UnitWidth)
				Next
			End If
		End If
	Next
	pcd_doubter = str
End Function

Rem 得到某点的字符
Function pcd_getDot(x,y,str,UnitWidth)
	pcd_getDot = Mid(str,x+1+y*UnitWidth,1)
End Function

Rem 设置某点的字符
Sub pcd_setDot(x,y,ByRef str,newDot,UnitWidth)
	str = Left(str,x+y*UnitWidth) & newDot & Right(str,Len(str)-x-y*UnitWidth-1)
End Sub

Rem 将数字转为bmp需要的格式 lens是目标字节长度
Function Num2ChrB(num,lens)
	Dim ret,i
	ret = ""
	While (num>0)
		ret = ret & ChrB(num mod 256)
		num = num \ 256
	WEnd
	For i=Lenb(ret) To lens-1
		ret = ret & chrB(0)
	Next
	Num2ChrB = ret
End Function
%>