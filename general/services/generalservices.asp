<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True	'缓冲数据，才向客户输出；
Response.Charset="utf-8"
Session.CodePage = 65001
%>
<%
'模块说明：后台操作公共服务.
'日期说明：2009-12-2
'版权说明：www.cokeshow.com.cn	可乐秀(北京)技术有限公司产品
'友情提示：本产品已经申请专利，请勿用于商业用途，如果需要帮助请联系可乐秀技术有限公司。
'联系电话：(010)-67659219
'电子邮件：cokeshow@qq.com
%>
<!--#include file="../inc/_public.asp"-->




<%
'变量定义区.
'(用来存储对象的变量，用全大写!)
Const maxPerPage=15							'当前模块分页设置.
Dim CurrentPageNow,TitleName,UnitName
CurrentPageNow 	= "generalservices.asp"			'当前页.
TitleName 		= "后台操作公共服务"				'此模块管理页的名字.
UnitName 		= "记录"					'此模块涉及记录的元素名.
'自定义设置.
'本地设置.
Dim CurrentTableName
CurrentTableName 	= "[CXBG_product]"		'此模块涉及的[表]名.
%>



<%
Dim totalPut,totalPages,currentPage			'分页用的控制变量.
Dim RS, sql									'查询列表记录用的变量.
Dim FoundErr,ErrMsg							'控制错误流程用的控制变量.
Dim strFileName								'构建查询字符串用的控制变量.
Dim ExecuteSearch,Keyword,TypeSearch,Action	'构建查询字符串以及流程控制用的控制变量.
Dim strGuide		'导航文字.

currentPage		=CokeShow.filtRequest(Request("Page"))
ExecuteSearch	=CokeShow.filtRequest(Request("ExecuteSearch"))
Keyword			=CokeShow.filtRequest(Request("Keyword"))
TypeSearch		=CokeShow.filtRequest(Request("TypeSearch"))
Action			=CokeShow.filtRequest(Request("Action"))

'处理查询执行 控制变量
If ExecuteSearch="" Then
	ExecuteSearch=0
Else
	If isNumeric(ExecuteSearch) Then ExecuteSearch=CokeShow.CokeClng(ExecuteSearch) Else ExecuteSearch=0
End If
'构造查询字符串
strFileName=CurrentPageNow &"?ExecuteSearch="& ExecuteSearch
If Keyword<>"" Then
	strFileName=strFileName&"&Keyword="& Keyword
End If
If TypeSearch<>"" Then
	strFileName=strFileName&"&TypeSearch="& TypeSearch
End If
'处理当前页码的控制变量，通过获取到的传值获取，默认为第一页1.
If currentPage<>"" Then
    If isNumeric(currentPage) Then currentPage=CokeShow.CokeClng(currentPage) Else currentPage=1
Else
	currentPage=1
End If

%>
		<%
		If Action="DeleteProductPhotos" Then
			Call DeleteProductPhotos()
		ElseIf Action="SaveAdd" Then
			Call SaveAdd()
		ElseIf Action="Modify" Then
			Call Modify()
		ElseIf Action="SaveModify" Then
			Call SaveModify()
		
		Else
			Call Main()
		End If
		
		
		If FoundErr=True Then
			Response.Write "{valid: false, message: '"& ErrMsg &"'}"
			'CokeShow.AlertErrMsg_general( ErrMsg )
		End If
		%>
<%
Sub Main()
	
		
End Sub

Sub showMain()
	
	Response.Write "{valid: false, message: '无任何操作~ <a href=http://www.CokeShow.com.cn/ target=_blank>www.CokeShow.com.cn</a>'}"
	
End Sub

'供Ajax调用，删除菜品相册中的一个图片服务.
'参数:
'1.id	删除的相册图片在数据库表的记录中的id号.
'要操作的目标表为:[CXBG_product__photos]
Sub DeleteProductPhotos()
	Dim intID
	intID	=CokeShow.filtRequest(Request("id"))
	
	If isNumeric(intID) Then
		'报错或者成功消息，均由Delete函数负责输出通知.
		Call Delete(intID,"[CXBG_product__photos]")
		
	Else
		'报错通知.
		Response.Write "{valid: false, message: '传入图片id不是数字！(出现异常)'}"
	End If
	
End Sub



Sub Modify()
	
	
End Sub


Sub SaveAdd()
	Dim ProductName,ProductNo,product_class_id,product_class_id_extend,product_brand_id,description
	Dim UnitPrice,UnitPrice_Market,isSales,UnitPrice_Sales,isSales_StartDate,isSales_StopDate,jifen,product_keywords,isOnsale
	Dim product_businessUSE_id,product_businessUSE_id_extend,product_activityUSE_id,product_activityUSE_id_extend,product_welfareUSE_id,product_welfareUSE_id_extend
	
	'获取其它参数
	ProductName					=CokeShow.filtRequest(Request("ProductName"))
	ProductNo					=CokeShow.filtRequest(Request("ProductNo"))
	product_class_id			=CokeShow.filtRequest(Request("product_class_id"))
	product_class_id_extend		=CokeShow.filtRequest(Request("product_class_id_extend"))
	'BEGIN
	product_businessUSE_id		=CokeShow.filtRequest(Request("product_businessUSE_id"))
	product_businessUSE_id_extend=CokeShow.filtRequest(Request("product_businessUSE_id_extend"))
	product_activityUSE_id		=CokeShow.filtRequest(Request("product_activityUSE_id"))
	product_activityUSE_id_extend=CokeShow.filtRequest(Request("product_activityUSE_id_extend"))
	product_welfareUSE_id		=CokeShow.filtRequest(Request("product_welfareUSE_id"))
	product_welfareUSE_id_extend=CokeShow.filtRequest(Request("product_welfareUSE_id_extend"))
	'END
	product_brand_id			=CokeShow.filtRequest(Request("product_brand_id"))
	description					=CokeShow.filtRequest(Request("description"))
	
	UnitPrice					=CokeShow.filtRequest(Request("UnitPrice"))
	UnitPrice_Market			=CokeShow.filtRequest(Request("UnitPrice_Market"))
	isSales						=CokeShow.filtRequest(Request("isSales"))
	UnitPrice_Sales				=CokeShow.filtRequest(Request("UnitPrice_Sales"))
	isSales_StartDate			=CokeShow.filtRequest(Request("isSales_StartDate"))
	isSales_StopDate			=CokeShow.filtRequest(Request("isSales_StopDate"))
	jifen						=CokeShow.filtRequest(Request("jifen"))
	product_keywords			=CokeShow.filtRequest(Request("product_keywords"))
	isOnsale					=CokeShow.filtRequest(Request("isOnsale"))
	
	
	'验证
	If ProductName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>菜品名称不能为空！</li>"
	Else
		If CokeShow.strLength(ProductName)>50 Or CokeShow.strLength(ProductName)<2 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品名称长度不能大于50个字符，也不能小于2个字符！</li>"
		Else
			ProductName=ProductName
		End If
	End If
	
	If ProductNo="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>菜品编号不能为空！</li>"
	Else
		If CokeShow.strLength(ProductNo)>20 Or CokeShow.strLength(ProductNo)<4 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品编号长度不能大于20个字符，也不能小于4个字符！</li>"
		Else
			ProductNo=ProductNo
		End If
	End If
	
'	If cnname<>"" Then
'		If Len(cnname)>20 Then
'			FoundErr=True
'			ErrMsg=ErrMsg &"<br><li>中文名只能20位字符之内！此项也可以不填。</li>"
'		Else
'			cnname=cnname
'		End If
'	End If
	
	If product_class_id="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请选择菜品分类！</li>"
	Else
		If isNumeric(product_class_id) Then
			product_class_id=CokeShow.CokeClng(product_class_id)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前菜品分类数字不是数字(出现异常)！</li>"
		End If
	End If
	
	If product_brand_id="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请选择菜品品牌！</li>"
	Else
		If isNumeric(product_brand_id) Then
			product_brand_id=CokeShow.CokeClng(product_brand_id)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前菜品品牌数字不是数字(出现异常)！</li>"
		End If
	End If
	
	If product_class_id_extend="" Or isNull(product_class_id_extend) Or isEmpty(product_class_id_extend) Then
'		FoundErr=True
'		ErrMsg=ErrMsg &"<br><li>菜品扩展分类不能为空！</li>"
		product_class_id_extend=""
	Else
		If CokeShow.strLength(product_class_id_extend)>255 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品扩展分类长度不能大于255个字符！</li>"
		Else
			product_class_id_extend=product_class_id_extend
		End If
	End If
	
	If description="" Or isNull(description) Or isEmpty(description) Then
'		FoundErr=True
'		ErrMsg=ErrMsg &"<br><li>菜品扩展分类不能为空！</li>"
		description=""
	Else
		If CokeShow.strLength(description)>255 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品描述长度不能大于255个字符！</li>"
		Else
			description=description
		End If
	End If
	
'	If MSN<>"" Then
'		If CokeShow.IsValidEmail(MSN)=false Then
'			FoundErr=True
'			ErrMsg=ErrMsg & "<br><li>你的MSN格式不正确！</li>"
'		End If
'	End If
	
	
	If UnitPrice="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请填写持卡会员价格！</li>"
	Else
		If isNumeric(UnitPrice) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前持卡会员价格不是数字(非法输入)！</li>"
		End If
	End If
	
	If UnitPrice_Market="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请填写价格！</li>"
	Else
		If isNumeric(UnitPrice_Market) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前价格不是数字(非法输入)！</li>"
		End If
	End If
	
	If isSales="" Then
		isSales=0
		
	Else
		If isNumeric(isSales) Then
			isSales=CokeShow.CokeClng(isSales)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前促销选项的参数不是数字(非法输入)！</li>"
		End If
	End If
	
	If UnitPrice_Sales="" Then
		
		
	Else
		If isNumeric(UnitPrice_Sales) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前促销价不是数字(非法输入)！</li>"
		End If
	End If
	
	If isSales_StartDate="" Then
		
		
	Else
		If isDate(isSales_StartDate) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前起始促销日期格式不对！</li>"
		End If
	End If
	
	If isSales_StopDate="" Then
		
		
	Else
		If isDate(isSales_StopDate) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前结束促销日期格式不对！</li>"
		End If
	End If
	
	'jifen,product_keywords,isOnsale
	If jifen="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请填写积分，如果无积分请填写为0！</li>"
	Else
		If isNumeric(jifen) Then
			jifen=CokeShow.CokeClng(jifen)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前积分不是数字(非法输入)！</li>"
		End If
	End If
	
	If product_keywords="" Or isNull(product_keywords) Or isEmpty(product_keywords) Then
		product_keywords=""
		'关键.
	Else
		If CokeShow.strLength(product_keywords)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前菜品关键词超过了50字符！</li>"
		Else
			product_keywords=""
			
		End If
	End If
	
	If isOnsale="" Then
		isOnsale=0
		
	Else
		If isNumeric(isOnsale) Then
			isOnsale=CokeShow.CokeClng(isOnsale)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前上架选项的参数不是数字(非法输入)！</li>"
		End If
	End If
	
	
	'拦截错误，不然错误往下进行！
	If FoundErr=True Then Exit Sub
	Dim newID
	
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT * FROM "& CurrentTableName
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	RS.AddNew
		
		RS("ProductName")					=ProductName	'必填项
		RS("ProductNo")						=ProductNo		'非必填项
		RS("product_class_id")				=product_class_id
		RS("product_class_id_extend")		=toProcessRequest(product_class_id_extend)
		'BEGIN
		RS("product_businessUSE_id")			=product_businessUSE_id
		RS("product_businessUSE_id_extend")		=toProcessRequest(product_businessUSE_id_extend)
		RS("product_activityUSE_id")			=product_activityUSE_id
		RS("product_activityUSE_id_extend")		=toProcessRequest(product_activityUSE_id_extend)
		RS("product_welfareUSE_id")				=product_welfareUSE_id
		RS("product_welfareUSE_id_extend")		=toProcessRequest(product_welfareUSE_id_extend)
		'END
		RS("product_brand_id")				=product_brand_id
		RS("description")					=description
		
		RS("UnitPrice")					=UnitPrice
		RS("UnitPrice_Market")			=UnitPrice_Market
		RS("isSales")					=isSales
		RS("UnitPrice_Sales")			=UnitPrice_Sales
		If isSales_StartDate<>"" Then RS("isSales_StartDate")	=isSales_StartDate
		If isSales_StopDate<>"" Then RS("isSales_StopDate")		=isSales_StopDate
		RS("jifen")						=jifen
		RS("product_keywords")			=product_keywords
		RS("isOnsale")					=isOnsale
	
	RS.Update
	RS.MoveLast
	newID = RS("id")
	
	RS.Close
	Set RS=Nothing
	
	CokeShow.ShowOK "添加"& UnitName &"成功!",CurrentPageNow &"?Action=Modify&id="& newID
End Sub


Sub SaveModify()
	Dim ProductName,ProductNo,product_class_id,product_class_id_extend,product_brand_id,description
	Dim UnitPrice,UnitPrice_Market,isSales,UnitPrice_Sales,isSales_StartDate,isSales_StopDate,jifen,product_keywords,isOnsale
	Dim product_businessUSE_id,product_businessUSE_id_extend,product_activityUSE_id,product_activityUSE_id_extend,product_welfareUSE_id,product_welfareUSE_id_extend
	
	Dim intID
	intID	=CokeShow.filtRequest(Request("id"))
	'检测id参数.
	If intID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		Exit Sub
	Else
		intID=CokeShow.CokeClng(intID)
	End If
	
	
	
	'获取其它参数
	ProductName					=CokeShow.filtRequest(Request("ProductName"))
	ProductNo					=CokeShow.filtRequest(Request("ProductNo"))
	product_class_id			=CokeShow.filtRequest(Request("product_class_id"))
	product_class_id_extend		=CokeShow.filtRequest(Request("product_class_id_extend"))
	'BEGIN
	product_businessUSE_id		=CokeShow.filtRequest(Request("product_businessUSE_id"))
	product_businessUSE_id_extend=CokeShow.filtRequest(Request("product_businessUSE_id_extend"))
	product_activityUSE_id		=CokeShow.filtRequest(Request("product_activityUSE_id"))
	product_activityUSE_id_extend=CokeShow.filtRequest(Request("product_activityUSE_id_extend"))
	product_welfareUSE_id		=CokeShow.filtRequest(Request("product_welfareUSE_id"))
	product_welfareUSE_id_extend=CokeShow.filtRequest(Request("product_welfareUSE_id_extend"))
	'END
	product_brand_id			=CokeShow.filtRequest(Request("product_brand_id"))
	description					=CokeShow.filtRequest(Request("description"))
	'获取相册参数
	Dim photos_id,photos_src,photos_orderid
	photos_id					=CokeShow.filtRequest(Request("photos_id"))
	photos_src					=CokeShow.filtRequest(Request("photos_src"))
	photos_orderid				=CokeShow.filtRequest(Request("photos_orderid"))
	
	UnitPrice					=CokeShow.filtRequest(Request("UnitPrice"))
	UnitPrice_Market			=CokeShow.filtRequest(Request("UnitPrice_Market"))
	isSales						=CokeShow.filtRequest(Request("isSales"))
	UnitPrice_Sales				=CokeShow.filtRequest(Request("UnitPrice_Sales"))
	isSales_StartDate			=CokeShow.filtRequest(Request("isSales_StartDate"))
	isSales_StopDate			=CokeShow.filtRequest(Request("isSales_StopDate"))
	jifen						=CokeShow.filtRequest(Request("jifen"))
	product_keywords			=CokeShow.filtRequest(Request("product_keywords"))
	isOnsale					=CokeShow.filtRequest(Request("isOnsale"))
	
	
	'验证
	If ProductName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>菜品名称不能为空！</li>"
	Else
		If CokeShow.strLength(ProductName)>50 Or CokeShow.strLength(ProductName)<2 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品名称长度不能大于50个字符，也不能小于2个字符！</li>"
		Else
			ProductName=ProductName
		End If
	End If
	
	If ProductNo="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>菜品编号不能为空！</li>"
	Else
		If CokeShow.strLength(ProductNo)>20 Or CokeShow.strLength(ProductNo)<4 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品编号长度不能大于20个字符，也不能小于4个字符！</li>"
		Else
			ProductNo=ProductNo
		End If
	End If
	
'	If cnname<>"" Then
'		If Len(cnname)>20 Then
'			FoundErr=True
'			ErrMsg=ErrMsg &"<br><li>中文名只能20位字符之内！此项也可以不填。</li>"
'		Else
'			cnname=cnname
'		End If
'	End If
	
	If product_class_id="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请选择菜品分类！</li>"
	Else
		If isNumeric(product_class_id) Then
			product_class_id=CokeShow.CokeClng(product_class_id)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前菜品分类数字不是数字(出现异常)！</li>"
		End If
	End If
	
	If product_brand_id="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请选择菜品品牌！</li>"
	Else
		If isNumeric(product_brand_id) Then
			product_brand_id=CokeShow.CokeClng(product_brand_id)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前菜品品牌数字不是数字(出现异常)！</li>"
		End If
	End If
	
	If product_class_id_extend="" Or isNull(product_class_id_extend) Or isEmpty(product_class_id_extend) Then
'		FoundErr=True
'		ErrMsg=ErrMsg &"<br><li>菜品扩展分类不能为空！</li>"
		product_class_id_extend=""
	Else
		If CokeShow.strLength(product_class_id_extend)>255 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品扩展分类长度不能大于255个字符！</li>"
		Else
			product_class_id_extend=product_class_id_extend
		End If
	End If
	
	If description="" Or isNull(description) Or isEmpty(description) Then
'		FoundErr=True
'		ErrMsg=ErrMsg &"<br><li>菜品扩展分类不能为空！</li>"
		description=""
	Else
		If CokeShow.strLength(description)>255 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品描述长度不能大于255个字符！</li>"
		Else
			description=description
		End If
	End If
		
'	If MSN<>"" Then
'		If CokeShow.IsValidEmail(MSN)=false Then
'			FoundErr=True
'			ErrMsg=ErrMsg & "<br><li>你的MSN格式不正确！</li>"
'		End If
'	End If
	
	
	If UnitPrice="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请填写持卡会员价格！</li>"
	Else
		If isNumeric(UnitPrice) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前持卡会员价格不是数字(非法输入)！</li>"
		End If
	End If
	
	If UnitPrice_Market="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请填写价格！</li>"
	Else
		If isNumeric(UnitPrice_Market) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前价格不是数字(非法输入)！</li>"
		End If
	End If
	
	If isSales="" Then
		isSales=0
		
	Else
		If isNumeric(isSales) Then
			isSales=CokeShow.CokeClng(isSales)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前促销选项的参数不是数字(非法输入)！</li>"
		End If
	End If
	
	If UnitPrice_Sales="" Then
		
		
	Else
		If isNumeric(UnitPrice_Sales) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前促销价不是数字(非法输入)！</li>"
		End If
	End If
	
	If isSales_StartDate="" Then
		
		
	Else
		If isDate(isSales_StartDate) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前起始促销日期格式不对！</li>"
		End If
	End If
	
	If isSales_StopDate="" Then
		
		
	Else
		If isDate(isSales_StopDate) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前结束促销日期格式不对！</li>"
		End If
	End If
	
	'jifen,product_keywords,isOnsale
	If jifen="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请填写积分，如果无积分请填写为0！</li>"
	Else
		If isNumeric(jifen) Then
			jifen=CokeShow.CokeClng(jifen)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前积分不是数字(非法输入)！</li>"
		End If
	End If
	
	If product_keywords="" Or isNull(product_keywords) Or isEmpty(product_keywords) Then
		product_keywords=""
		'关键.
	Else
		If CokeShow.strLength(product_keywords)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前菜品关键词超过了50字符！</li>"
		Else
			product_keywords=""
		End If
	End If
	
	If isOnsale="" Then
		isOnsale=0
		
	Else
		If isNumeric(isOnsale) Then
			isOnsale=CokeShow.CokeClng(isOnsale)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前上架选项的参数不是数字(非法输入)！</li>"
		End If
	End If
	
	
	'拦截错误，不然错误往下进行！
	If FoundErr=True Then Exit Sub
	
	
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT * FROM "& CurrentTableName &" WHERE id="& intID
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,1,3
	
	'拦截此记录的异常情况.
	If RS.Bof And RS.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的"& UnitName &"！</li>"
		Exit Sub
	End If
	
		RS("ProductName")					=ProductName	'必填项
		RS("ProductNo")						=ProductNo		'非必填项
		RS("product_class_id")				=product_class_id
		RS("product_class_id_extend")		=toProcessRequest(product_class_id_extend)
		'BEGIN
		RS("product_businessUSE_id")			=product_businessUSE_id
		RS("product_businessUSE_id_extend")		=toProcessRequest(product_businessUSE_id_extend)
		RS("product_activityUSE_id")			=product_activityUSE_id
		RS("product_activityUSE_id_extend")		=toProcessRequest(product_activityUSE_id_extend)
		RS("product_welfareUSE_id")				=product_welfareUSE_id
		RS("product_welfareUSE_id_extend")		=toProcessRequest(product_welfareUSE_id_extend)
		'END
		RS("product_brand_id")				=product_brand_id
		RS("description")					=description
		
		RS("modifydate")					=Now()
		
		
		RS("UnitPrice")					=UnitPrice
		RS("UnitPrice_Market")			=UnitPrice_Market
		RS("isSales")					=isSales
		RS("UnitPrice_Sales")			=UnitPrice_Sales
		If isSales_StartDate<>"" Then RS("isSales_StartDate")	=isSales_StartDate
		If isSales_StopDate<>"" Then RS("isSales_StopDate")		=isSales_StopDate
		RS("jifen")						=jifen
		RS("product_keywords")			=product_keywords
		RS("isOnsale")					=isOnsale
	
	RS.Update
	RS.Close
	Set RS=Nothing
	
	'更新菜品相册数据
	Call SaveModify__photos( photos_id, photos_src, photos_orderid, intID )
	
	CokeShow.ShowOK "修改"& UnitName &"成功!",CurrentPageNow
End Sub


Sub Delete(theID,theTableName)
	Dim strID,i
	strID=CokeShow.filtRequest(theID)
	If strID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要删除的"& UnitName &"</li>"
		Exit Sub
	End If
	If Instr(strID,",")>0 Then
		strID=Split(strID,",")
		For i=0 To Ubound(strID)
			Call DeleteOne(strID(i),theTableName)
		Next
	Else
		Call DeleteOne(strID,theTableName)
	End If
	
	'拦截错误，不然错误往下进行！
	If FoundErr=True Then Exit Sub
	
	'CokeShow.ShowOK "删除操作成功!",CurrentPageNow
	Response.Write "{valid: true, message: '您的删除操作成功!'}"
	
End Sub

Sub DeleteOne(strID,theTableName)
	strID=CokeShow.CokeClng(strID)
	'response.Write theTableName
	If Not IsObject(CONN) Then link_database
	Set RS=CONN.Execute("SELECT * FROM "& theTableName &" WHERE id="& strID)
	
	If Not RS.Eof Then
		
		CONN.Execute("DELETE FROM "& theTableName &" WHERE id="& strID)
		'//CokeShow.Execute("UPDATE "& CurrentTableName &" SET deleted=1 WHERE username='"& username &"'")
		'CONN.Execute("UPDATE "& CurrentTableName &" SET deleted=1 WHERE id="& strID)
		
	Else
		'找不着记录，则
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>记录id为"& strID &"的"& UnitName &"删除操作没成功，此记录有可能早已丢失</li>"
		Exit Sub
		
	End If
	
End Sub


'设置会员是否通过我站点审核...通过便可以在站点显示，成为公开会员用户.
'示例函数.
Sub change_isValidate()
	Dim strState		'本过程的状态，字符串.
	strState = "审核"
	
	Dim strID
	Dim isValidate
	
	strID=CokeShow.filtRequest(Request("id"))
	isValidate= CokeShow.CokeClng(CokeShow.filtRequest(Request("isValidate")))
	
	If strID="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>参数不足！</li>"
		Exit Sub
	Else
		strID=CokeShow.CokeClng(strID)
	End If
	

	If FoundErr=True Then
		Exit Sub
	End If

	If isValidate= 1 Then
		CONN.Execute("UPDATE "& CurrentTableName &" SET isValidate=0 WHERE id="& id )
	ElseIf isValidate= 0 Then
		CONN.Execute("UPDATE "& CurrentTableName &" SET isValidate=1 WHERE id="& id )
	End If

	CokeShow.ShowOK "更新"& UnitName & "的"& strState &"状态成功！",CurrentPageNow
	
End Sub


%>