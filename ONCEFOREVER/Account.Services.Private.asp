<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True	'缓冲数据，才向客户输出；
Response.Charset="utf-8"
Session.CodePage = 65001
Response.ContentType="text/html"
%>﻿
<%
'
'Ajax远程调用asp文件.只供已登录帐号的私有调用服务，所以调用的前提必须是在检测通过了登录的情况下使用.（可以放置调用敏感信息的函数）
'
%>﻿
<!--#include virtual="/system/system_conn.asp"-->
<!--#include virtual="/system/system_class.asp"-->
<!--#include virtual="/system/md5.asp"-->

<!--#include virtual="/public/JSON.asp"--> 
<!--#include virtual="/public/JSON_UTIL.asp"--> 
<%
'类实例化
Dim CokeShow
Set CokeShow = New SystemClass
CokeShow.Start		'调用类的Start方法，初始化类里的ReloadSetup()函数，并得到二维数组Setup
Call CokeShow.SQLWarningSys()	'预警.【上边一行原本是被注销的，WL为未来可能发生的问题提示！】
%>
<%
'公共变量定义区.
Dim ServicesAction,FoundErr,ErrMsg
Dim strServiceFeedback			'本服务的完全反馈信息字符.WL
Dim ResultJSON					'JSON system
%>
<%
'公共变量赋值区.
ServicesAction	=CokeShow.filtRequest(Request("ServicesAction"))
FoundErr		=False
ErrMsg			=""
%>
<%
'预处理区.

%>
<%
'安全处理区.
'要保证来源来自站内链接，并且有会员帐号的登录身份.
'如果已经登录，并且来源来自站内链接，则允许继续通行.
If CokeShow.CheckUserLogined()=True And isNumeric(Session("id")) And Len(Session("username"))>=10 And CokeShow.ChkPost=True Then
	'通行，会员帐号通过.
Else
	'CokeShow.AlertErrMsg_foreground("您的网络处于不安全级别II.已写入管理日志"):Response.Redirect "/ONCEFOREVER/LogOn.Welcome":Response.End()	'被JSON系统替换掉了.
	FoundErr=True
	ErrMsg=ErrMsg &"登录状态已过期，请您先完成登录/注册操作！<br />当您登录/注册完后将自动回到当前页面哦。<br /><br /><center>立刻 <a class=button_img77 style=display:inline-block; color:red; href=/ONCEFOREVER/LogOn.Welcome?fromurl="& CokeShow.EncodeURL( Request.ServerVariables("HTTP_REFERER"),"" ) &">登录/注册</a></center>"
	'Exit Function
End If
%>


<%
'逻辑判断区
SELECT CASE ServicesAction
	Case "addFavorite"
		If Not FoundErr=True Then strServiceFeedback = addFavorite()				'所有函数均不用para参数，每个function独立的直接用request获取想要的参数！
	Case "addOutOfStoreToRregister"
		If Not FoundErr=True Then strServiceFeedback = addOutOfStoreToRregister()	'所有函数均不用para参数，每个function独立的直接用request获取想要的参数！
	Case "addAccount_RemarkOn"
		If Not FoundErr=True Then strServiceFeedback = addAccount_RemarkOn()		'所有函数均不用para参数，每个function独立的直接用request获取想要的参数！
	Case "addAccount_TuijianPengyou_CheckReady"
		If Not FoundErr=True Then strServiceFeedback = addAccount_TuijianPengyou_CheckReady()	'先GET检测当前Ajax操作环境是否就绪，然后才进行以下Ajax的Form提交表单！
	Case "addAccount_TuijianPengyou"
		If Not FoundErr=True Then strServiceFeedback = addAccount_TuijianPengyou()				'最终推荐朋友操作的Ajax的Form提交表单处理函数！
	Case "addAccount_GiftCertificated"
		If Not FoundErr=True Then strServiceFeedback = addAccount_GiftCertificated()			'会员领取兑换券的Ajax的Form提交表单处理函数！
	Case "addAccount_SendMessage"
		If Not FoundErr=True Then strServiceFeedback = addAccount_SendMessage()			'站内信息的Ajax的Form提交表单处理函数！
	Case "addAccount_BindingMyVIPCard"
		If Not FoundErr=True Then strServiceFeedback = addAccount_BindingMyVIPCard()			'会员申请VIP卡绑定Ajax的Form提交表单处理函数！
	Case "addAccount_PasswordOnChange"
		If Not FoundErr=True Then strServiceFeedback = addAccount_PasswordOnChange()			'会员修改密码Ajax的Form提交表单处理函数！
	Case "addAccount_PersonalInformation"
		If Not FoundErr=True Then strServiceFeedback = addAccount_PersonalInformation()			'会员修改个人资料Ajax的Form提交表单处理函数！
	
	Case Else
		'no
		
End Select



If FoundErr=True Then
'	strServiceFeedback = ErrMsg
	
	'-------------------------------------------------JSON system
	'JSON system
	'创建JSON之ASP对象.
	Set ResultJSON = jsObject()
	'JSON system
	
	'JSON system
	'构建结果集反馈信息.
	ResultJSON("isAjaxSuccessful")		="true"
	ResultJSON("theResult_true_false")	="false"
	ResultJSON("theAllInformation")		=ErrMsg
	'strShow = ResultJSON.FlushNow		'输出形式2.
	ResultJSON.Flush					'输出形式1.
	'JSON system
	
	'JSON system
	'销毁JSON之ASP对象.
	Set ResultJSON = Nothing
	'JSON system
	'-------------------------------------------------JSON system
	
	Response.End()
End If
'If FoundErr=True Then
'	strServiceFeedback = ErrMsg
'End If



'如果没有什么错误，那么就输出特定函数的输出字符串(JSON描述)
Response.Write strServiceFeedback	'被JSON系统替换掉了.
Response.End()
'结束Services.
%>

<%
'具体服务区.

'服务名称：		为会员添加收藏菜品服务.
'服务描述：		向系统请求加入一个收藏菜品？请回答.
'输出用的方式：	JSON.
'被调用的方式：	URL Get Ajax.
'将要接收的参数：	1.有username，为会员帐号信息.
'				2.有id，为菜品ID号.
'返回数据：		JSON数据对象
'				第一个对象属性，名为valid，值为false操作失败 或 true操作成功。
'				第二个对象属性，名为message，值为反馈信息。
Public Function addFavorite()
	
	'定义内部新变量进行内部操作.
	Dim strShow
	Dim userName,intID
	Dim sql,RS
	
	'初始化赋值.
	addFavorite	=False	'默认为False操作失败.
	userName	=CokeShow.filtPass(Session("username"))
	intID		=CokeShow.filtRequest(Request("id"))
	strShow		=""
	
	'判断有各种效性.
	'intID
	If isNumeric(intID) Then
		If CokeShow.CokeCint(intID)=0 Then
			'参数不正确，退出.操作失败.
			FoundErr=True
			ErrMsg=ErrMsg &"菜品ID参数不正确！"
			Exit Function
		End If
		intID=CokeShow.CokeCint(intID)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addFavorite的参数ID不正确，无法为会员添加收藏菜品服务的操作！"
		Exit Function
	End If
	'userName
	If isNull(userName) Or isEmpty(userName) Or userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addFavorite的参数userName不正确，无法为会员添加收藏菜品服务的操作！"
		Exit Function
	End If
	If userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"帐号不能为空！"
		Exit Function
	Else
		If CokeShow.strLength(userName)>50 Or CokeShow.strLength(userName)<10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的Email帐号长度不应大于50个字符，也不应小于10个字符的！"
			Exit Function
		Else
			If CokeShow.IsValidEmail(userName)=False Then
				FoundErr=True
				ErrMsg=ErrMsg & "您的帐号(即Email)格式不正确！"
				Exit Function
			Else
				userName=userName
			End If
		End If
	End If
	
	'验证此帐号中，是否已经存在当前菜品的收藏了.
	Dim rs2,sql2
	sql2="SELECT TOP 1 * FROM [CXBG_account_Favorite] WHERE deleted=0 AND Account_LoginID='"& userName &"' AND product_id="& intID
	Set rs2=CONN.Execute(sql2)
	'如果已存在收藏，则报错退出函数.
	If Not rs2.Eof Then
		FoundErr=True
		'ErrMsg=ErrMsg &"{valid: ""false"", message:""您已收藏过此菜品了.""}"	'被JSON系统替换掉了.
		ErrMsg=ErrMsg &"您已经收藏过此菜品了哦."
		Exit Function
	End If
	rs2.Close
	Set rs2=Nothing
	
	'验证此菜品ID是否存在其相应的菜品记录.
	sql2="SELECT TOP 1 * FROM [CXBG_product] WHERE deleted=0 AND id="& intID
	Set rs2=CONN.Execute(sql2)
	'如果根本没有此菜，则报错退出函数.
	If rs2.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg &"您当前所浏览的菜品可能已经被删除，或者不存在."
		Exit Function
	End If
	rs2.Close
	Set rs2=Nothing
	
	
	'-----------------Go Begin
	'新增收藏菜品.
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT TOP 1 * FROM [CXBG_account_Favorite] WHERE deleted=0"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	
	RS.AddNew
		
		RS("Account_LoginID")	=userName
		RS("product_id")		=intID
		
		'-------------------------------------------------JSON system
		'JSON system
		'创建JSON之ASP对象.
		Set ResultJSON = jsObject()
		'JSON system
		
		'JSON system
		'构建结果集反馈信息.
		ResultJSON("isAjaxSuccessful")		="true"
		ResultJSON("theResult_true_false")	="true"
		ResultJSON("theAllInformation")		="恭喜您，菜品已经成功加入了您收藏夹中!"
		strShow = ResultJSON.FlushNow		'输出形式2.
		'JSON system
		
		'JSON system
		'销毁JSON之ASP对象.
		Set ResultJSON = Nothing
		'JSON system
		'-------------------------------------------------JSON system
		
		'strShow = "{valid: ""true"", message:""恭喜您，菜品成功加入了您收藏夹中!""}"	'被JSON系统替换掉了.
		
		
	RS.Update
	RS.MoveLast
	Dim newID
	newID = RS("id")
	
	RS.Close
	Set RS=Nothing
	'-----------------Go End
	
	
	
	'积分运算.Begin
		'获得该菜品所规定的积分.（如果是绑定会员卡级别会员，那么积分翻倍并且文字要说明清楚双倍情况）
		Dim intTmp101
		'intTmp101=CokeShow.otherField("[CXBG_product]",intID,"id","jifen",True,0)
		'If isNumeric(intTmp101) Then intTmp101=CokeShow.CokeCint(intTmp101) Else intTmp101=0
		intTmp101=1		'收藏菜品统一为1积分奖励.
		'如果绑定了会员卡.给双倍积分！
		If Cstr(Session("isBindingVIPCardNumber"))="1" And Session("BindingVIPCardNumber")<>"" And Len(Session("BindingVIPCardNumber"))>=4 Then
			If CokeShow.JifenSystemExecute(6,userName,intTmp101*2,"，获得了网站主办的鼓励小活动-收藏菜品的 <span style=color:red;>双倍积分奖励</span>"& intTmp101*2 &"网站积分.",newID)=True Then Response.Write "" Else Response.Write ""
		'给单倍积分！
		Else
			If CokeShow.JifenSystemExecute(6,userName,intTmp101,"，获得了网站主办的鼓励小活动-收藏菜品的积分奖励"& intTmp101 &"网站积分.",newID)=True Then Response.Write "" Else Response.Write ""
		End If
	'积分运算.End
	
	
	
	addFavorite = strShow
	
End Function

'服务名称：		为会员添加缺货登记服务.
'服务描述：		向系统请求加入一个菜品缺货登记？请回答.
'输出用的方式：	JSON.
'被调用的方式：	URL Get Ajax.
'将要接收的参数：	1.有username，为会员帐号信息.
'				2.有id，为菜品ID号.
'返回数据：		JSON数据对象
'				第一个对象属性，名为valid，值为false操作失败 或 true操作成功。
'				第二个对象属性，名为message，值为反馈信息。
Public Function addOutOfStoreToRregister()
	
	'定义内部新变量进行内部操作.
	Dim strShow
	Dim userName,intID
	Dim sql,RS
	
	'初始化赋值.
	addOutOfStoreToRregister	=False	'默认为False操作失败.
	userName	=CokeShow.filtPass(Session("username"))
	intID		=CokeShow.filtRequest(Request("id"))
	strShow		=""
	
	'判断有各种效性.
	'intID
	If isNumeric(intID) Then
		If CokeShow.CokeCint(intID)=0 Then
			'参数不正确，退出.操作失败.
			FoundErr=True
			ErrMsg=ErrMsg &"菜品ID参数不正确！"
			Exit Function
		End If
		intID=CokeShow.CokeClng(intID)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addOutOfStoreToRregister的参数ID不正确，无法为会员添加菜品缺货登记服务的操作！<br />"
		Exit Function
	End If
	'userName
	If isNull(userName) Or isEmpty(userName) Or userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addOutOfStoreToRregister的参数userName不正确，无法为会员添加菜品缺货登记服务的操作！<br />"
		Exit Function
	End If
	If userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"帐号不能为空！"
		Exit Function
	Else
		If CokeShow.strLength(userName)>50 Or CokeShow.strLength(userName)<10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的Email帐号长度不应大于50个字符，也不应小于10个字符的！"
			Exit Function
		Else
			If CokeShow.IsValidEmail(userName)=False Then
				FoundErr=True
				ErrMsg=ErrMsg & "您的帐号(即Email)格式不正确！"
				Exit Function
			Else
				userName=userName
			End If
		End If
	End If
	'验证此帐号中，是否已经存在当前菜品的缺货登记了.
	Dim rs2,sql2
	sql2="SELECT TOP 1 * FROM [CXBG_account_OutOfStoreToRregister] WHERE deleted=0 AND Account_LoginID='"& userName &"' AND product_id="& intID
	Set rs2=CONN.Execute(sql2)
	'如果已存在缺货登记，则报错退出函数.
	If Not rs2.Eof Then
		FoundErr=True
		'ErrMsg=ErrMsg &"{valid: ""false"", message:""您已登记过此菜品的缺货信息了，请耐心等待痴心不改餐厅为您处理并且与您取得联系.<br />""}"	'被JSON系统替换掉了.
		ErrMsg=ErrMsg &"您已登记过此菜品的缺货信息了，请耐心等待痴心不改餐厅为您处理并且与您取得联系.<br />"
		Exit Function
	End If
	rs2.Close
	Set rs2=Nothing
	
	
	'-----------------Go Begin
	'新增菜品缺货登记.
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT TOP 1 * FROM [CXBG_account_OutOfStoreToRregister] WHERE deleted=0"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	
	RS.AddNew
		
		RS("Account_LoginID")	=userName
		RS("product_id")		=intID
		
		'-------------------------------------------------JSON system
		'JSON system
		'创建JSON之ASP对象.
		Set ResultJSON = jsObject()
		'JSON system
		
		'JSON system
		'构建结果集反馈信息.
		ResultJSON("isAjaxSuccessful")		="true"
		ResultJSON("theResult_true_false")	="true"
		ResultJSON("theAllInformation")		="恭喜您，菜品成功加入了缺货登记系统中!请耐心等待痴心不改餐厅为您处理并且与您取得联系.<br />"
		strShow = ResultJSON.FlushNow		'输出形式2.
		'JSON system
		
		'JSON system
		'销毁JSON之ASP对象.
		Set ResultJSON = Nothing
		'JSON system
		'-------------------------------------------------JSON system
		
		'strShow = "{valid: ""true"", message:""恭喜您，菜品成功加入了缺货登记系统中!请耐心等待痴心不改餐厅为您处理并且与您取得联系.<br />""}"	'被JSON系统替换掉了.
		
		
	RS.Update
	
	RS.Close
	Set RS=Nothing
	'-----------------Go End
	
	addOutOfStoreToRregister = strShow
	
End Function


'服务名称：		为会员添加点评菜品的点评内容的服务.
'服务描述：		向系统发送本会员的点评内容，成功了吗？请回答.
'输出用的方式：	JSON.
'被调用的方式：	Form Post Ajax.
'将要接收的参数：	1.被点评菜品的ID号.				id
'				2.口味分数.						ChineseDish_Taste
'				3.环境分数.						ChineseDish_DiningArea
'				4.服务分数.						ChineseDish_Service
'				5.人均消费.						ChineseDish_ConsumePerPerson
'				6.点评详细内容.					logtext
'				7.点评星级(只对绑定了卡的会员有效).	theStarRatingForChineseDishInformation
'返回数据：		JSON数据对象
'				第一个对象属性，名为valid，值为false操作失败 或 true操作成功。
'				第二个对象属性，名为message，值为反馈信息。
Public Function addAccount_RemarkOn()
	
	'定义内部新变量进行内部操作.
	Dim strShow
	Dim userName,intID
	Dim sql,RS
	
	'初始化赋值.
	addAccount_RemarkOn	=False	'默认为False操作失败.
	userName	=CokeShow.filtPass(Session("username"))
	intID		=CokeShow.filtRequest(Request("id"))
	strShow		=""
	'其它初始化赋值.
	Dim ChineseDish_Taste,ChineseDish_DiningArea,ChineseDish_Service,ChineseDish_ConsumePerPerson,logtext,theStarRatingForChineseDishInformation
	ChineseDish_Taste					=CokeShow.filtRequest(Request("ChineseDish_Taste"))
	ChineseDish_DiningArea				=CokeShow.filtRequest(Request("ChineseDish_DiningArea"))
	ChineseDish_Service					=CokeShow.filtRequest(Request("ChineseDish_Service"))
	ChineseDish_ConsumePerPerson		=CokeShow.filtRequest(Request("ChineseDish_ConsumePerPerson"))
	logtext								=CokeShow.filtRequest(Request("logtext"))
	theStarRatingForChineseDishInformation	=CokeShow.filtRequest(Request("theStarRatingForChineseDishInformation"))
	
	'判断有各种效性.
	'intID
	If isNumeric(intID) Then
		If CokeShow.CokeCint(intID)=0 Then
			'参数不正确，退出.操作失败.
			FoundErr=True
			ErrMsg=ErrMsg &"菜品ID参数不正确！"
			Exit Function
		End If
		intID=CokeShow.CokeClng(intID)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_RemarkOn的参数ID不正确，为会员添加点评菜品的点评内容的服务！"
		Exit Function
	End If
	'userName
	If isNull(userName) Or isEmpty(userName) Or userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_RemarkOn的参数userName不正确，为会员添加点评菜品的点评内容的服务！"
		Exit Function
	End If
	If userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"帐号不能为空！"
		Exit Function
	Else
		If CokeShow.strLength(userName)>50 Or CokeShow.strLength(userName)<10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的Email帐号长度不应大于50个字符，也不应小于10个字符的！"
			Exit Function
		Else
			If CokeShow.IsValidEmail(userName)=False Then
				FoundErr=True
				ErrMsg=ErrMsg & "您的帐号(即Email)格式不正确！"
				Exit Function
			Else
				userName=userName
			End If
		End If
	End If
	'ChineseDish_Taste
	If isNumeric(ChineseDish_Taste) Then
		If CokeShow.CokeCint(ChineseDish_Taste)=0 Then
			'参数不正确，退出.操作失败.
			FoundErr=True
			ErrMsg=ErrMsg &"请选择口味分数哦！"
			Exit Function
		End If
		ChineseDish_Taste=CokeShow.CokeCint(ChineseDish_Taste)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"参数错误，口味分数不是数字！"
		Exit Function
	End If
	'ChineseDish_DiningArea
	If isNumeric(ChineseDish_DiningArea) Then
		If CokeShow.CokeCint(ChineseDish_DiningArea)=0 Then
			'参数不正确，退出.操作失败.
			FoundErr=True
			ErrMsg=ErrMsg &"请选择环境分数哦！"
			Exit Function
		End If
		ChineseDish_DiningArea=CokeShow.CokeCint(ChineseDish_DiningArea)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"参数错误，环境分数不是数字！"
		Exit Function
	End If
	'ChineseDish_Service
	If isNumeric(ChineseDish_Service) Then
		If CokeShow.CokeCint(ChineseDish_Service)=0 Then
			'参数不正确，退出.操作失败.
			FoundErr=True
			ErrMsg=ErrMsg &"请选择服务分数哦！"
			Exit Function
		End If
		ChineseDish_Service=CokeShow.CokeCint(ChineseDish_Service)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"参数错误，服务分数不是数字！"
		Exit Function
	End If
	'ChineseDish_ConsumePerPerson
	If isNumeric(ChineseDish_ConsumePerPerson) Then
		ChineseDish_ConsumePerPerson=FormatCurrency(ChineseDish_ConsumePerPerson, 2)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"请正确填写人均消费的价格哦！"
		Exit Function
	End If
	
	'logtext
	If logtext="" Or isNull(logtext) Or isEmpty(logtext) Then
		FoundErr=True
		ErrMsg=ErrMsg &"点评内容不能为空！</li>"
		Exit Function
	Else
		If CokeShow.strLength(logtext)<3 Or CokeShow.strLength(logtext)>250 Then
			FoundErr=True
			ErrMsg=ErrMsg &"点评内容文字长度不能太少，也不能超过规定的长度哦"
			Exit Function
		Else
			logtext=logtext
		End If
	End If
	'theStarRatingForChineseDishInformation
	If isNumeric(theStarRatingForChineseDishInformation) Then
'		If CokeShow.CokeCint(theStarRatingForChineseDishInformation)=0 Then
'			'参数不正确，退出.操作失败.
'			FoundErr=True
'			ErrMsg=ErrMsg &"请选择星级，以全面为菜品赋予总体星级评价！"
'			Exit Function
'		End If
		theStarRatingForChineseDishInformation=CokeShow.CokeCint(theStarRatingForChineseDishInformation)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"参数错误，星级选择不是数字！"
		Exit Function
	End If
	
	
	'验证码
	If Not CokeShow.CodePassII("CodeStr") Then
		FoundErr=True
		ErrMsg=ErrMsg &"验证码不正确哦，如果仍然不正确，请点击重新刷新验证码。"
		Exit Function
	End If
	
	'验证此帐号中，是否已经存在当前菜品的评论了（重复评论：在同一个菜品前提下，不能在10分钟内再次评论，需要等待10分钟后才能再次进行评论！并且提示:"您刚才已经成功的提交了评论，需要等待十分钟后才能再次评论此菜品"）.
	Dim rs2,sql2
	sql2="SELECT TOP 1 * FROM [CXBG_account_RemarkOn] WHERE deleted=0 AND Account_LoginID='"& userName &"' AND product_id="& intID &" AND DateDiff(mi, adddate, GETDATE())<10 "	'GETDATE()
'Response.Write sql2
	Set rs2=CONN.Execute(sql2)
	'如果已存在10分钟之内对此菜品做过点评操作，则报错退出函数，并提示用户需要10分钟后才能再点评.
	If Not rs2.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg &"您刚才已经成功的提交过评论了，需要等待十分钟后才能再次点评此菜品."
		Exit Function
	End If
	rs2.Close
	Set rs2=Nothing
	
	
	'-----------------Go Begin
	'新增点评.
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT TOP 1 * FROM [CXBG_account_RemarkOn] WHERE deleted=0"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	
	RS.AddNew
		
		RS("Account_LoginID")	=userName
		RS("product_id")		=intID
		
		RS("Account_LoginCNNAME")		=Session("cnname")
		RS("Account_LoginCLIENT_NAME")	=Session("client_name")
		RS("Account_LoginSEX")			=Session("Sex")
		
		RS("ChineseDish_Taste")				=ChineseDish_Taste
		RS("ChineseDish_DiningArea")		=ChineseDish_DiningArea
		RS("ChineseDish_Service")			=ChineseDish_Service
		RS("ChineseDish_ConsumePerPerson")	=ChineseDish_ConsumePerPerson
		'如果绑定了会员卡.
		'If Cstr(Session("isBindingVIPCardNumber"))="1" And Session("BindingVIPCardNumber")<>"" And Len(Session("BindingVIPCardNumber"))>=4 Then
			'绑定会员卡的会员所提交的星级评.
			RS("theStarRatingForChineseDishInformation")	=theStarRatingForChineseDishInformation		'星级，默认是0星级.
		'End If
		RS("logtext")						=logtext
		

		
		
	RS.Update
	RS.MoveLast
	Dim newID
	newID = RS("id")
	
	RS.Close
	Set RS=Nothing
	'-----------------Go End
	
	'积分到底得了多少.
	Dim intJifenFollowUp,strJifenFollowUp
	intJifenFollowUp=CokeShow.otherField("[CXBG_product]",intID,"id","jifen",True,0)
	strJifenFollowUp=""
	
	'积分运算.Begin
		'验证此帐号中，是否已经存在当前菜品的评论了（如果点评过了，就是得到过积分了，那么就不能再得积分了！）.
		Dim rs3,sql3
		sql3="SELECT TOP 1 * FROM [CXBG_account_RemarkOn] WHERE deleted=0 AND Account_LoginID='"& userName &"' AND product_id="& intID &" AND DateDiff(mi, adddate, GETDATE())>10 "		'AND DateDiff(mi, adddate, GETDATE())>10为了排除刚刚添加完的点评记录也在内，所以限制了一下，只看十分钟以前有没有相应点评存在。
'Response.Write sql3
		Set rs3=CONN.Execute(sql3)
		'如果不存在对此菜品做过点评记录，则给积分！
		''''''''''''''''''wl餐厅商量的积分改进策略——点评过了也得积分！'''''''''''''''''''If rs3.Eof Then
			'获得该菜品所规定的积分.（如果是绑定会员卡级别会员，那么积分翻倍并且文字要说明清楚双倍情况）
			If isNumeric(intJifenFollowUp) Then intJifenFollowUp=CokeShow.CokeCint(intJifenFollowUp) Else intJifenFollowUp=0
			'如果绑定了会员卡.给双倍积分！
			If Cstr(Session("isBindingVIPCardNumber"))="1" And Session("BindingVIPCardNumber")<>"" And Len(Session("BindingVIPCardNumber"))>=4 Then
				intJifenFollowUp=intJifenFollowUp*2
				If CokeShow.JifenSystemExecute(2,userName,intJifenFollowUp,"，获得了菜品点评的 <span style=color:red;>双倍积分奖励</span>"& intJifenFollowUp &"网站积分.",newID)=True Then Response.Write "" Else Response.Write ""
				strJifenFollowUp="<span style=color:red;>双倍积分奖励</span>"
			'给单倍积分！
			Else
				intJifenFollowUp=intJifenFollowUp*1
				If CokeShow.JifenSystemExecute(2,userName,intJifenFollowUp,"，获得了菜品点评积分奖励"& intJifenFollowUp &"网站积分.",newID)=True Then Response.Write "" Else Response.Write ""
				strJifenFollowUp="<span style=color:red;></span>"
			End If
		''''''''''''''''''wl餐厅商量的积分改进策略——点评过了也得积分！'''''''''''''''''''Else
		''''''''''''''''''wl餐厅商量的积分改进策略——点评过了也得积分！'''''''''''''''''''	intJifenFollowUp=0
		''''''''''''''''''wl餐厅商量的积分改进策略——点评过了也得积分！'''''''''''''''''''End If
		rs3.Close
		Set rs3=Nothing
	'积分运算.End
	
	'-------------------------------------------------JSON system
	'JSON输出.
	'JSON system
	'创建JSON之ASP对象.
	Set ResultJSON = jsObject()
	'JSON system
	
	'JSON system
	'构建结果集反馈信息.
	ResultJSON("isAjaxSuccessful")		="true"
	ResultJSON("theResult_true_false")	="true"
	ResultJSON("theAllInformation")		="恭喜您，菜品点评成功!&nbsp;<a href='/ChineseDish/ChineseDishInformation.Welcome?CokeMark="& CokeShow.AddCode_Num(intID) &"'>确认</a>"
	ResultJSON("intJifenFollowUp")		=""& intJifenFollowUp &""
	ResultJSON("strJifenFollowUp")		=strJifenFollowUp
	strShow = ResultJSON.FlushNow		'输出形式2.
	'JSON system
	
	'JSON system
	'销毁JSON之ASP对象.
	Set ResultJSON = Nothing
	'JSON system
	'-------------------------------------------------JSON system
	
	addAccount_RemarkOn = strShow
	
End Function

'服务名称：		会员推荐朋友之前的检测操作环境是否就绪的检测服务.
'服务描述：		向系统发送本会员的推荐朋友之环境检测，我登录好并且能够提交推荐朋友了吗？请回答.
'输出用的方式：	JSON.
'被调用的方式：	URL Get Ajax.
'将要接收的参数：	1.被点评菜品的ID号.				id
'返回数据：		JSON数据对象
'				第一个对象属性，名为valid，值为false操作失败 或 true操作成功。
'				第二个对象属性，名为message，值为反馈信息。
Public Function addAccount_TuijianPengyou_CheckReady()
	'定义内部新变量进行内部操作.
	Dim strShow
	Dim userName,intID
	Dim sql,RS
	
	'初始化赋值.
	addAccount_TuijianPengyou_CheckReady	=False	'默认为False操作失败.
	userName	=CokeShow.filtPass(Session("username"))
	intID		=CokeShow.filtRequest(Request("id"))
	strShow		=""
	
	'判断有各种效性.
	'intID
	If isNumeric(intID) Then
		If CokeShow.CokeCint(intID)=0 Then
			'参数不正确，退出.操作失败.
			FoundErr=True
			ErrMsg=ErrMsg &"菜品ID参数不正确！"
			Exit Function
		End If
		intID=CokeShow.CokeClng(intID)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_TuijianPengyou_CheckReady的参数ID不正确，为会员添加点评菜品的点评内容的服务！"
		Exit Function
	End If
	'userName
	If isNull(userName) Or isEmpty(userName) Or userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_TuijianPengyou_CheckReady的参数userName不正确，为会员添加点评菜品的点评内容的服务！"
		Exit Function
	End If
	If userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"帐号不能为空！"
		Exit Function
	Else
		If CokeShow.strLength(userName)>50 Or CokeShow.strLength(userName)<10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的Email帐号长度不应大于50个字符，也不应小于10个字符的！"
			Exit Function
		Else
			If CokeShow.IsValidEmail(userName)=False Then
				FoundErr=True
				ErrMsg=ErrMsg & "您的帐号(即Email)格式不正确！"
				Exit Function
			Else
				userName=userName
			End If
		End If
	End If
	
	'-------------------------------------------------JSON system
	'JSON输出.
	'JSON system
	'创建JSON之ASP对象.
	Set ResultJSON = jsObject()
	'JSON system
	
	'JSON system
	'构建结果集反馈信息.
	ResultJSON("isAjaxSuccessful")		="true"
	ResultJSON("theResult_true_false")	="true"
	ResultJSON("theAllInformation")		="开始推荐"
	strShow = ResultJSON.FlushNow		'输出形式2.
	'JSON system
	
	'JSON system
	'销毁JSON之ASP对象.
	Set ResultJSON = Nothing
	'JSON system
	'-------------------------------------------------JSON system
	
	addAccount_TuijianPengyou_CheckReady = strShow
	
End Function

'服务名称：		会员推荐菜品给朋友的Ajax表单处理服务.
'服务描述：		向系统发送本会员的点评内容，成功了吗？请回答.
'输出用的方式：	JSON.
'被调用的方式：	Form Post Ajax.
'将要接收的参数：	1.被点评菜品的ID号.				id
'				2.会员的朋友的称呼.				FName
'				3.会员的朋友的Email.				FEmail
'				4.菜品的菜名.						PName
'				5.菜品的价格.						PPrice
'				6.菜品的持卡会员价格.				HPrice
'返回数据：		JSON数据对象
'				第一个对象属性，名为valid，值为false操作失败 或 true操作成功。
'				第二个对象属性，名为message，值为反馈信息。
Public Function addAccount_TuijianPengyou()
	
	'定义内部新变量进行内部操作.
	Dim strShow
	Dim userName,intID
	Dim sql,RS
	
	'初始化赋值.
	addAccount_TuijianPengyou	=False	'默认为False操作失败.
	userName	=CokeShow.filtPass(Session("username"))
	intID		=CokeShow.filtRequest(Request("id"))
	strShow		=""
	'其它初始化赋值.
	Dim FName,FEmail,PName,PPrice,HPrice
	FName				=CokeShow.filtRequest(Request("FName"))
	FEmail				=CokeShow.filtRequest(Request("FEmail"))
	PName				=CokeShow.filtRequest(Request("PName"))
	PPrice				=CokeShow.filtRequest(Request("PPrice"))
	HPrice				=CokeShow.filtRequest(Request("HPrice"))
		
	'判断有各种效性.
	'intID
	If isNumeric(intID) Then
		If CokeShow.CokeCint(intID)=0 Then
			'参数不正确，退出.操作失败.
			FoundErr=True
			ErrMsg=ErrMsg &"菜品ID参数不正确！"
			Exit Function
		End If
		intID=CokeShow.CokeClng(intID)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_TuijianPengyou的参数ID不正确，为会员添加点评菜品的点评内容的服务！"
		Exit Function
	End If
	'userName
	If isNull(userName) Or isEmpty(userName) Or userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_TuijianPengyou的参数userName不正确，为会员添加点评菜品的点评内容的服务！"
		Exit Function
	End If
	If userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"帐号不能为空！"
		Exit Function
	Else
		If CokeShow.strLength(userName)>50 Or CokeShow.strLength(userName)<10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的Email帐号长度不应大于50个字符，也不应小于10个字符的！"
			Exit Function
		Else
			If CokeShow.IsValidEmail(userName)=False Then
				FoundErr=True
				ErrMsg=ErrMsg & "您的帐号(即Email)格式不正确！"
				Exit Function
			Else
				userName=userName
			End If
		End If
	End If
	'FName
	If FName="" Or isNull(FName) Or isEmpty(FName) Then
		FoundErr=True
		ErrMsg=ErrMsg &"您推荐的朋友的名字称呼不能为空哦！</li>"
		Exit Function
	Else
		If CokeShow.strLength(FName)<2 Or CokeShow.strLength(FName)>20 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您推荐的朋友的名字称呼的长度不能少于2个字或超过20个字哦！"
			Exit Function
		Else
			FName=FName
		End If
	End If
	'FEmail
	If FEmail="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"您推荐的朋友的Email不能为空！"
		Exit Function
	Else
		If CokeShow.strLength(FEmail)>50 Or CokeShow.strLength(FEmail)<10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您推荐的朋友的Email长度不应大于50个字符，也不应小于10个字符哦！"
			Exit Function
		Else
			If CokeShow.IsValidEmail(FEmail)=False Then
				FoundErr=True
				ErrMsg=ErrMsg & "您推荐的朋友的Email格式不正确！"
				Exit Function
			Else
				FEmail=FEmail
			End If
		End If
	End If
	'PName
	If PName="" Or isNull(PName) Or isEmpty(PName) Then
		FoundErr=True
		ErrMsg=ErrMsg &"参数出错，请重试！菜品名丢失~</li>"
		Exit Function
	Else
		If CokeShow.strLength(PName)<1 Or CokeShow.strLength(PName)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"参数出错，菜品名的长度不能少于1个字或超过50个字！"
			Exit Function
		Else
			PName=PName
		End If
	End If
	'PPrice
	If isNumeric(PPrice) Then
		PPrice=FormatCurrency(PPrice, 2)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"参数出错，菜品价格不是数字！"
		Exit Function
	End If
	
	'HPrice
	If isNumeric(HPrice) Then
		HPrice=FormatCurrency(HPrice, 2)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"参数出错，菜品持卡会员价格不是数字！"
		Exit Function
	End If
	
	
	'验证码
	If Not CokeShow.CodePassII("CodeStr") Then
		FoundErr=True
		ErrMsg=ErrMsg &"验证码不正确哦，如果仍然不正确，请点击重新刷新验证码。"
		Exit Function
	End If
	
	'验证此帐号是否已经在今天发送过当前Email了.如果发送过，阻止.
	Dim rs2,sql2
	sql2="SELECT TOP 1 * FROM [CXBG_account_TuijianPengyou] WHERE deleted=0 AND Account_LoginID='"& userName &"' AND product_id="& intID &" AND FEmail='"& FEmail &"' AND DateDiff(day, adddate, GETDATE())=0 "	'GETDATE()
'Response.Write sql2
	Set rs2=CONN.Execute(sql2)
	'如果在同一天里，有对此菜品推荐过给同一个朋友的Email，则报错退出函数，并提示用户今天已经推荐过给当前朋友email了，每日推荐，请次日再推荐.
	If Not rs2.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg &"您今天已经成功的为您的朋友<strong>"& FName &"("& FEmail &")</strong>推荐过菜品:"& PName &"了，为了不打扰您的好朋友，我们可以明天继续给他/她推荐当前菜品，或者您现在可以为他/她推荐其它的菜品哦.<br /><br /><img src=/images/ico/emotion_wink.png />——痴心不改每日推荐好友活动，积分回馈、乐趣共享"
		Exit Function
	End If
	rs2.Close
	Set rs2=Nothing
	
	
	'-----------------Go Begin
	'新增记录.
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT TOP 1 * FROM [CXBG_account_TuijianPengyou] WHERE deleted=0"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	
	RS.AddNew
		
		RS("Account_LoginID")	=userName
		RS("product_id")		=intID
		
		RS("Account_LoginCNNAME")		=Session("cnname")
		RS("Account_LoginCLIENT_NAME")	=Session("client_name")
		RS("Account_LoginSEX")			=Session("Sex")
		
		RS("FName")				=FName
		RS("FEmail")			=FEmail
		RS("PName")				=PName
		RS("PPrice")			=PPrice
		RS("HPrice")			=HPrice
'		'如果绑定了会员卡.
'		If Cstr(Session("isBindingVIPCardNumber"))="1" And Session("BindingVIPCardNumber")<>"" And Len(Session("BindingVIPCardNumber"))>=4 Then
'			'绑定会员卡的会员所提交的星级评.
'			RS("theStarRatingForChineseDishInformation")	=theStarRatingForChineseDishInformation		'星级，默认是0星级.
'		End If
		
	RS.Update
	RS.MoveLast
	Dim newID
	newID=RS("id")
	
	RS.Close
	Set RS=Nothing
	'-----------------Go End
	
	'积分运算.Begin
		'获得推荐朋友所规定的3积分.（如果是绑定会员卡级别会员，那么积分翻倍并且文字要说明清楚双倍情况）
		Dim intTmp101
		Dim intJifenFollowUp,strJifenFollowUp
		intTmp101=3		'推荐菜品给朋友，统一为3的积分奖励.
		'如果绑定了会员卡.给双倍积分！
		If Cstr(Session("isBindingVIPCardNumber"))="1" And Session("BindingVIPCardNumber")<>"" And Len(Session("BindingVIPCardNumber"))>=4 Then
			intJifenFollowUp=intTmp101*2
			strJifenFollowUp="<br /><img src=/images/ico/small/coins_add.png /> 您已经获得了推荐朋友的特别积分奖励之 <span style=color:red;>双倍积分奖励</span>"& intJifenFollowUp &"积分哦."
			If CokeShow.JifenSystemExecute(10,userName,intJifenFollowUp,"，您获得了推荐朋友的特别积分奖励的 <span style=color:red;>双倍积分奖励</span>"& intJifenFollowUp &"网站积分哦.",newID)=True Then Response.Write "" Else Response.Write ""
		'给单倍积分！
		Else
			intJifenFollowUp=intTmp101
			strJifenFollowUp="<br /><img src=/images/ico/small/coins_add.png /> 您已经获得了推荐朋友的特别积分奖励 "& intJifenFollowUp &"积分哦."
			If CokeShow.JifenSystemExecute(10,userName,intJifenFollowUp,"，您获得了推荐朋友的特别积分奖励"& intJifenFollowUp &"网站积分哦.",newID)=True Then Response.Write "" Else Response.Write ""
		End If
	'积分运算.End
	
	'-------------------------------------------------JSON system
	'JSON输出.
	'JSON system
	'创建JSON之ASP对象.
	Set ResultJSON = jsObject()
	'JSON system
	
	'JSON system
	'构建结果集反馈信息.
	ResultJSON("isAjaxSuccessful")		="true"
	ResultJSON("theResult_true_false")	="true"
	ResultJSON("theAllInformation")		="<img src=/images/ico/emotion_happy.png /> 恭喜您，推荐信成功发送，您的朋友已经收到您的邮件推荐!<img src=/images/ico/small/accept.png />"
	ResultJSON("intJifenFollowUp")		=""& intJifenFollowUp &""
	ResultJSON("strJifenFollowUp")		=strJifenFollowUp
	strShow = ResultJSON.FlushNow		'输出形式2.
	'JSON system
	
	'JSON system
	'销毁JSON之ASP对象.
	Set ResultJSON = Nothing
	'JSON system
	'-------------------------------------------------JSON system
	
	'发送电子邮件！！最后================================Jmail.Begin
	'发送邮件.
	'定义标题和内容.
	Dim Topic,LogText
	'标题.
	Topic	="您的朋友在痴心不改餐厅的网站上向您推荐了一道美味菜品，您同时可以查看到餐厅为您推荐的最新菜品哦——痴心不改餐厅(北京)"
	'内容.
	Dim strPhotoSrc
	strPhotoSrc=CokeShow.otherField("[CXBG_product]",intID,"id","Photo",True,0)
	
	If Len(strPhotoSrc)>8 Then strPhotoSrc=strPhotoSrc Else strPhotoSrc="/images/NoPic.png"
	LogText	="您的朋友<span color=orange;>"& Session("client_name") &"("& userName &")</span>在痴心不改餐厅的网站上向您推荐了一道地道的美味菜品哦。推荐菜品的详情，请查阅痴心不改餐厅网站链接&nbsp;<img src="& system_user_domain &""& strPhotoSrc &" width=""110"" height=""110"" />&nbsp;<a href="""& system_user_domain &"/ChineseDish/ChineseDishInformation.Welcome?CokeMark="& CokeShow.AddCode_Num(intID) &""" target=""_blank""><span color=orange;>"& PName &"&nbsp;"& FormatCurrency(PPrice,2) &"</span></a>，如需了解更多详情，我们邀请您您一起登录痴心不改餐厅官方网站 <a href="""& system_user_domain &""" target=""_blank"">"& system_user_domain &"</a>"
	
	'构造模板 b
	Dim strLogText
	strLogText=strLogText &"<style>A:visited {	TEXT-DECORATION: none	}"
	strLogText=strLogText &"A:active  {	TEXT-DECORATION: none	}"
	strLogText=strLogText &"A:hover   {	TEXT-DECORATION: underline	}"
	strLogText=strLogText &"A:link 	  {	text-decoration: none;}"
	strLogText=strLogText &"BODY   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}"
	strLogText=strLogText &"TD	   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt	}</style>"
	strLogText=strLogText &"<TABLE border=0 width='95%' align=center><TBODY><TR><TD>"
	
	strLogText=strLogText &"尊敬的"& FName &"您好，"
	strLogText=strLogText & LogText 
	strLogText=strLogText &"<br /><br /><br /><br /><br /><br />"
	
	strLogText=strLogText &"<p>痴心不改餐厅，时尚菜系、经典美味！ <a href="""& system_user_domain &""" target=""_blank"">痴心不改，钟情一生</a></p>"
	'strLogText=strLogText &"*************************************************************************************<br />"
	strLogText=strLogText &"可回复邮箱地址：&nbsp;&nbsp;&nbsp;<a href="& system_ReplyEmailAddress &">"& system_ReplyEmailAddress &"</a><br />"
	strLogText=strLogText &"餐厅地址："& CokeShow.Setup(31,0) &"<br />"
	strLogText=strLogText &"餐厅名称：痴心不改<br />"
	'strLogText=strLogText &"*************************************************************************************"
	strLogText=strLogText &"</TD></TR></TBODY></TABLE>"
	'构造模板 e
	
	'如果发送成功，则输出一些成功消息.
	If CokeShow.SendMail(FEmail,"痴心不改餐厅",system_ReplyEmailAddress,Topic,strLogText,"gb2312","text/html",system_JMailFrom,system_JMailSMTP,system_JMailMailServerUserName,system_JMailMailServerPassWord)=True Then
		'ErrMsg="\r\n发送成功！您的朋友将收到您推荐的邮件，谢谢您一如既往的关注\r\n"
		'Response.Write( "<script type=""text/javascript"">alert('"& ErrMsg &"');< /script>" )
	Else
	'发送失败时.
		'ErrMsg="\r\n发送失败！"
		'Response.Write( "<script type=""text/javascript"">alert('"& ErrMsg &"');</ script>" )
	End If
	'发送电子邮件！！最后================================Jmail.End
	
	addAccount_TuijianPengyou = strShow
	
	
	'更新推荐朋友的邮件标题字段和邮件内容字段.Begin
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT TOP 1 * FROM [CXBG_account_TuijianPengyou] WHERE id="& newID
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,1,3
	
	'拦截此记录的异常情况.
	If RS.Bof And RS.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg &"可能是由于系统原因，邮件没有发送出去！但是您的积分已经送出了哦，谢谢您的热情支持！"
		Exit Function
	End If
	
		RS("Topic")					=Topic
		RS("LogText")				=LogText
	
	RS.Update
	
	RS.Close
	Set RS=Nothing
	'更新推荐朋友的邮件标题字段和邮件内容字段.End
	
End Function

'服务名称：		为会员兑换礼品券.兑换礼品券并添加扣除积分.
'服务描述：		向系统发送本会员的兑换礼品券请求，我登录好并且能够提交兑换礼品券了吗？请回答.
'输出用的方式：	JSON.
'被调用的方式：	URL Get Ajax.
'将要接收的参数：	1.要兑换礼品券的ID号.				id
'返回数据：		JSON数据对象
'				第一个对象属性，名为valid，值为false操作失败 或 true操作成功。
'				第二个对象属性，名为message，值为反馈信息。
Public Function addAccount_GiftCertificated()
	'定义内部新变量进行内部操作.
	Dim strShow
	Dim userName,intID
	Dim sql,RS
	
	'初始化赋值.
	addAccount_GiftCertificated	=False	'默认为False操作失败.
	userName	=CokeShow.filtPass(Session("username"))
	intID		=CokeShow.filtRequest(Request("id"))
	strShow		=""
	
	'判断有各种效性.
	'intID
	If isNumeric(intID) Then
		If CokeShow.CokeCint(intID)=0 Then
			'参数不正确，退出.操作失败.
			FoundErr=True
			ErrMsg=ErrMsg &"菜品ID参数不正确！"
			Exit Function
		End If
		intID=CokeShow.CokeClng(intID)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_GiftCertificated的参数ID不正确，无法为会员兑换礼品券！"
		Exit Function
	End If
	'userName
	If isNull(userName) Or isEmpty(userName) Or userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_GiftCertificated的参数userName不正确，无法为会员兑换礼品券！"
		Exit Function
	End If
	If userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"帐号不能为空！"
		Exit Function
	Else
		If CokeShow.strLength(userName)>50 Or CokeShow.strLength(userName)<10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的Email帐号长度不应大于50个字符，也不应小于10个字符的！"
			Exit Function
		Else
			If CokeShow.IsValidEmail(userName)=False Then
				FoundErr=True
				ErrMsg=ErrMsg & "您的帐号(即Email)格式不正确！"
				Exit Function
			Else
				userName=userName
			End If
		End If
	End If
	
	'验证此帐号中，是否已经存在当前礼品券了（已经曾经对换过的礼品券就不用重复兑换了，以免浪费会员积分）.
	Dim rs2,sql2
	sql2="SELECT TOP 1 * FROM [CXBG_account_GiftCertificated] WHERE deleted=0 AND Account_LoginID='"& userName &"' AND GiftCertificated_id="& intID &""	'AND DateDiff(mi, adddate, GETDATE())<10 
'Response.Write sql2
	Set rs2=CONN.Execute(sql2)
	'如果已经曾经对换过的礼品券就不用重复兑换了.
	If Not rs2.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg &"您已经对换过此礼品券了，请您进入您的帐号管理中心，并进入我兑换的礼品券，查找您所需要的礼品券！."
		Exit Function
	End If
	rs2.Close
	Set rs2=Nothing
	
	'验证此帐号中，总共可用积分是否足够支付当前积分消费（）.
	Dim intTmp101
		intTmp101=CokeShow.otherField("[CXBG_GiftCertificated]",intID,"id","jifen",True,0)
		If isNumeric(intTmp101) Then intTmp101=CokeShow.CokeCint(intTmp101) Else intTmp101=0
		
	'检测当前积分总数是否足够支付礼品券！！！
	'当前积分为
	Dim JifenNow
	JifenNow=CokeShow.ChkAccountUserNameAllJifen(userName)
	If (JifenNow-intTmp101)>=0 Then
		'积分充足，通过.
	Else
		FoundErr=True
		ErrMsg=ErrMsg &"您的当前积分为："& JifenNow &"，已经不足够支付当前的礼品券了！您可以按以下若干方法去获取更多积分哦：）<br /><br />1. 每日首次登录获取积分<br />2. 点评菜品获积分<br />3. 在餐厅办理VIP卡，绑定后可获双倍积分<br />4. 在某个菜品详情页面中，可以推荐给您的好朋友获取积分<br />"
		Exit Function
	End If
	
	
	'-----------------Go Begin
	'新增会员兑换的礼品券.
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT TOP 1 * FROM [CXBG_account_GiftCertificated] WHERE deleted=0"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	
	RS.AddNew
		
		RS("Account_LoginID")			=userName
		RS("GiftCertificated_id")		=intID
		
		
		
	RS.Update
	RS.MoveLast
	Dim newID
	newID = RS("id")
	
	RS.Close
	Set RS=Nothing
	'-----------------Go End
	
	
	'积分运算.Begin
		'获得该菜品所规定的积分.（如果是绑定会员卡级别会员，那么积分翻倍并且文字要说明清楚双倍情况）
		'//Dim intTmp101
'//		intTmp101=CokeShow.otherField("[CXBG_GiftCertificated]",intID,"id","jifen",True,0)
'//		If isNumeric(intTmp101) Then intTmp101=CokeShow.CokeCint(intTmp101) Else intTmp101=0
		If CokeShow.JifenSystemExecute(4,userName,intTmp101,"，成功领取了痴心不改餐厅的礼品券，并扣除了您消费的"& intTmp101 &"网站积分.",newID)=True Then Response.Write "" Else Response.Write ""
	'积分运算.End
	
	'-------------------------------------------------JSON system
	'JSON输出.
	'JSON system
	'创建JSON之ASP对象.
	Set ResultJSON = jsObject()
	'JSON system
	
	'JSON system
	'构建结果集反馈信息.
	ResultJSON("isAjaxSuccessful")		="true"
	ResultJSON("theResult_true_false")	="true"
	ResultJSON("theAllInformation")		="恭喜您，礼品券已经加入到您的帐号中！请您进入帐号管理中心，点击进入‘我兑换的礼品券’管理项，然后查收刚刚所获得的礼品券哦！<br /><br /><img src=/images/ico/small/printer.png /> 最后下载到桌面并自行在家打印即可到来餐厅享受优惠了！<br /><br /><img src=/images/ico/small/coins_delete.png /> 痴心不改提示您，您已经成功消费了"& intTmp101 &"积分，如需查看请进入帐号管理中心，查看积分历史."
	strShow = ResultJSON.FlushNow		'输出形式2.
	'JSON system
	
	'JSON system
	'销毁JSON之ASP对象.
	Set ResultJSON = Nothing
	'JSON system
	'-------------------------------------------------JSON system
	
	addAccount_GiftCertificated = strShow
	
End Function

'服务名称：		会员发送新站内留言（即发问题），以网站大后台回复会员留言.(主要为ReplyID为0的新留言 或者 为某条留言记录的id号的回复信息)
'服务描述：		向系统发送本会员或者大后台的发信和回信请求，我登录好并且能够提交发信或回信了吗？请回答.
'输出用的方式：	JSON.
'被调用的方式：	Form Post Ajax.
'将要接收的参数：	1.要回复信息的ID号.				ReplyID
'返回数据：		JSON数据对象
'				第一个对象属性，名为valid，值为false操作失败 或 true操作成功。
'				第二个对象属性，名为message，值为反馈信息。
Public Function addAccount_SendMessage()
	'定义内部新变量进行内部操作.
	Dim strShow
	Dim userName,intID
	Dim sql,RS
	Dim title,content,toWho,toWho_CNNAME
	
	'初始化赋值.
	addAccount_SendMessage	=False	'默认为False操作失败.
	userName	=CokeShow.filtPass(Session("username"))
	intID		=CokeShow.filtRequest(Request("ReplyID"))
	strShow		=""
	
	title		=CokeShow.filtRequest(Request("title"))
	content		=CokeShow.filtRequest(Request("content"))
	'MessageType	=CokeShow.filtRequest(Request("MessageType"))
	toWho		=CokeShow.filtRequest(Request("toWho"))
	toWho_CNNAME=CokeShow.filtRequest(Request("toWho_CNNAME"))
	
	'判断有各种效性.
	'intID
	If isNumeric(intID) Then
		If CokeShow.CokeCint(intID)=0 Then
			'为新留言请求.
			intID=0
		End If
		intID=CokeShow.CokeClng(intID)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_SendMessage的参数ID不正确，无法发送站内留言！"
		Exit Function
	End If
	'userName
	If intID=0 Then		'是会员在发新留言请求.
		userName	=CokeShow.filtPass(Session("username"))
	Else				'是大后台在发回复请求.
		userName	="Coke"
	End If
	If isNull(userName) Or isEmpty(userName) Or userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_SendMessage的参数userName不正确，无法发送站内留言！"
		Exit Function
	End If
	If userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"帐号不能为空！"
		Exit Function
	Else
		If intID=0 Then		'是会员在发新留言请求.
			'BBB
			If CokeShow.strLength(userName)>50 Or CokeShow.strLength(userName)<10 Then
				FoundErr=True
				ErrMsg=ErrMsg &"您的Email帐号长度不应大于50个字符，也不应小于10个字符的！"
				Exit Function
			Else
				If CokeShow.IsValidEmail(userName)=False Then
					FoundErr=True
					ErrMsg=ErrMsg & "您的帐号(即Email)格式不正确！"
					Exit Function
				Else
					userName=userName
				End If
			End If
			'BBB
			
		Else				'是大后台在发回复请求.
			'AAA
			If CokeShow.strLength(userName)>6 Or CokeShow.strLength(userName)<4 Then
				FoundErr=True
				ErrMsg=ErrMsg &"您的大后台帐号长度不应大于6个字符，也不应小于4个字符的！"
				Exit Function
			Else
				userName=userName
			End If
			'AAA
			
		End If
	End If
	
	'验证此帐号中，是否已经存在当前礼品券了（已经曾经对换过的礼品券就不用重复兑换了，以免浪费会员积分）.
	If title<>"" Then
		If CokeShow.strLength(title)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"留言标题只能50位字符之内！"
			Exit Function
		Else
			title=title
		End If
	Else
		FoundErr=True
		ErrMsg=ErrMsg &"留言标题不能为空！"
		Exit Function
	End If
	
	If content<>"" Then
		If CokeShow.strLength(content)<6 Or CokeShow.strLength(content)>188 Then
			FoundErr=True
			ErrMsg=ErrMsg &"留言内容只能188位字符之内，同时填写内容字数不能太少，具体信息能够让我们更好的为您解决问题！"
			Exit Function
		Else
			content=content
		End If
	Else
		FoundErr=True
		ErrMsg=ErrMsg &"留言内容不能为空！"
		Exit Function
	End If
	
	'验证码
	If Not CokeShow.CodePassII("CodeStr") Then
		FoundErr=True
		ErrMsg=ErrMsg &"验证码不正确哦，如果仍然不正确，请点击重新刷新验证码。"
		Exit Function
	End If
	
'	If MessageType<>"" Then
'		If CokeShow.strLength(MessageType)>10 Then
'			FoundErr=True
'			ErrMsg=ErrMsg &"<br><li>留言类型只能10位字符之内！</li>"
'		Else
'			MessageType=MessageType
'		End If
'	Else
'		MessageType=""
'	End If
	'toWho.
	If isNull(toWho) Or isEmpty(toWho) Or toWho="" Then
		toWho=""
	End If
	If toWho="" Then
		toWho=""
	Else
		If CokeShow.strLength(toWho)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"对方的Email帐号长度不应大于50个字符！"
			Exit Function
		Else
			If CokeShow.IsValidEmail(toWho)=False Then
				FoundErr=True
				ErrMsg=ErrMsg & "对方帐号(即Email)格式不正确！"
				Exit Function
			Else
				toWho=toWho
			End If
		End If
	End If
	'toWho_CNNAME.
	If toWho_CNNAME<>"" Then
		If CokeShow.strLength(toWho_CNNAME)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"对方中文称呼只能50位字符之内！"
			Exit Function
		Else
			toWho_CNNAME=toWho_CNNAME
		End If
	Else
		toWho_CNNAME=""
	End If
	
	
	'-----------------Go Begin
	'新增会员兑换的礼品券.
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT TOP 1 * FROM [CXBG_account_Message] WHERE deleted=0"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	
	RS.AddNew
		
		RS("title")			=title
		RS("content")		=content
		'RS("MessageType")	=MessageType
		
		If intID=0 Then		'是会员在发新留言请求.
			RS("Account_LoginID")			=userName
			RS("Account_LoginCNNAME")		=Session("cnname")
			RS("toWho")						="Coke"
			RS("toWho_CNNAME")				=CokeShow.Setup(1,0)
		Else				'是大后台在发回复请求.
			RS("Account_LoginID")			=userName
			RS("Account_LoginCNNAME")		=CokeShow.Setup(1,0)
			RS("toWho")						=toWho
			RS("toWho_CNNAME")				=toWho_CNNAME
		End If
		
		
	RS.Update
	
	RS.Close
	Set RS=Nothing
	'-----------------Go End
	
	
	'-------------------------------------------------JSON system
	'JSON输出.
	'JSON system
	'创建JSON之ASP对象.
	Set ResultJSON = jsObject()
	'JSON system
	
	'JSON system
	'构建结果集反馈信息.
	ResultJSON("isAjaxSuccessful")		="true"
	ResultJSON("theResult_true_false")	="true"
	ResultJSON("theAllInformation")		="恭喜您，您的留言发送已成功！<img src=/images/ico/small/accept.png />"
	strShow = ResultJSON.FlushNow		'输出形式2.
	'JSON system
	
	'JSON system
	'销毁JSON之ASP对象.
	Set ResultJSON = Nothing
	'JSON system
	'-------------------------------------------------JSON system
	
		
	addAccount_SendMessage = strShow
	
End Function


'服务名称：		会员申请VIP卡绑定.
'服务描述：		向系统发送本会员的申请VIP卡绑定请求，我登录好并且能够提交申请VIP卡绑定了吗？请回答.
'输出用的方式：	JSON.
'被调用的方式：	Form Post Ajax.
'将要接收的参数：	1.无效参数.				id
'返回数据：		JSON数据对象
'				第一个对象属性，名为valid，值为false操作失败 或 true操作成功。
'				第二个对象属性，名为message，值为反馈信息。
Public Function addAccount_BindingMyVIPCard()
	'定义内部新变量进行内部操作.
	Dim strShow
	Dim userName,intID
	Dim sql,RS
	Dim VIPCARDNUMBER
	
	'初始化赋值.
	addAccount_BindingMyVIPCard	=False	'默认为False操作失败.
	userName	=CokeShow.filtPass(Session("username"))
	intID		=CokeShow.filtRequest(Request("id"))
	strShow		=""
	VIPCARDNUMBER=CokeShow.filtRequest(Request("VIPCARDNUMBER"))
	
	'判断有各种效性.
	'intID
	If isNumeric(intID) Then
		If CokeShow.CokeCint(intID)=0 Then
			'参数不正确，退出.操作失败.
			FoundErr=True
			ErrMsg=ErrMsg &"ID参数不正确！"
			Exit Function
		End If
		intID=CokeShow.CokeClng(intID)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_BindingMyVIPCard的参数ID不正确，无法为会员兑换礼品券！"
		Exit Function
	End If
	'userName
	If isNull(userName) Or isEmpty(userName) Or userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_BindingMyVIPCard的参数userName不正确，无法为会员兑换礼品券！"
		Exit Function
	End If
	If userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"帐号不能为空！"
		Exit Function
	Else
		If CokeShow.strLength(userName)>50 Or CokeShow.strLength(userName)<10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的Email帐号长度不应大于50个字符，也不应小于10个字符的！"
			Exit Function
		Else
			If CokeShow.IsValidEmail(userName)=False Then
				FoundErr=True
				ErrMsg=ErrMsg & "您的帐号(即Email)格式不正确！"
				Exit Function
			Else
				userName=userName
			End If
		End If
	End If
	
	'VIPCARDNUMBER
	If VIPCARDNUMBER="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"VIP卡号不能为空！"
		Exit Function
	Else
		If CokeShow.strLength(VIPCARDNUMBER)>20 Or CokeShow.strLength(VIPCARDNUMBER)<4 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的VIP卡号长度不应大于20个字符，也不应小于4个字符的！"
			Exit Function
		Else
			VIPCARDNUMBER=VIPCARDNUMBER
		End If
	End If
	
	'验证码
	If Not CokeShow.CodePassII("CodeStr") Then
		FoundErr=True
		ErrMsg=ErrMsg &"验证码不正确哦，如果仍然不正确，请点击重新刷新验证码。"
		Exit Function
	End If
	
	'验证此帐号中，是否已经绑定过卡号了（已经绑定过就不用再此绑定卡号了）.
	Dim rs2,sql2
	sql2="SELECT TOP 1 * FROM [View_isBoundVIPcard_AccountInformation_Records] WHERE deleted=0 AND username='"& userName &"'"			'AND DateDiff(mi, adddate, GETDATE())<10 
'Response.Write sql2
	Set rs2=CONN.Execute(sql2)
	'如果已经绑定过卡号了.
	If Not rs2.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg &"您已经成功绑定过VIP卡号了，不用再次绑定！."
		Exit Function
	End If
	rs2.Close
	Set rs2=Nothing
	
	'此卡号是否可以绑定，是否为餐厅录入的卡号？
	Dim rs3,sql3
	sql3="SELECT TOP 1 * FROM [CXBG_VIPcard] WHERE isOnpublic=1 AND classname='"& VIPCARDNUMBER &"'"			'AND DateDiff(mi, adddate, GETDATE())<10 
'Response.Write sql3
	Set rs3=CONN.Execute(sql3)
	'如果已经绑定过卡号了.
	If rs3.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg &"此VIP卡号暂不在痴心不改餐厅会员卡卡号的列表中，您确定此卡号是正确的餐厅会员卡卡号吗？如果是，请及时和我们取得联系，我们将会尽快为您处理解并决问题哦！"
		Exit Function
	End If
	rs3.Close
	Set rs3=Nothing
	
	'验证是否已经有其他人绑定此卡号啦！？
	Dim rs4,sql4
	sql4="SELECT TOP 1 * FROM [CXBG_account] WHERE deleted=0 AND username<>'"& userName &"' AND BindingVIPCardNumber='"& VIPCARDNUMBER &"'"			'AND DateDiff(mi, adddate, GETDATE())<10 
'Response.Write sql4
	Set rs4=CONN.Execute(sql4)
	'如果已经绑定过卡号了.
	If Not rs4.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg &"已经有其它会员绑定了此VIP卡号了，您确定此卡号是正确的餐厅会员卡卡号吗？如有问题，请及时和我们取得联系，我们将会尽快为您查实并处理！"
		Exit Function
	End If
	rs4.Close
	Set rs4=Nothing
	
	
	'-----------------Go Begin
	'修改会员表，为其绑定VIP卡号.
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT TOP 1 * FROM [CXBG_account] WHERE deleted=0 AND username='"& userName &"'"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	
	
		
		RS("isBindingVIPCardNumber")			=1
		RS("BindingVIPCardNumber")				=VIPCARDNUMBER
		
		
		
	RS.Update
	RS.MoveLast
	Dim newID
	newID = RS("id")
	
	RS.Close
	Set RS=Nothing
	'-----------------Go End
	
	
	'更新对应的VIP卡列表中的标记，标记已经有会员认领了此卡号.
	CONN.Execute( "UPDATE [CXBG_VIPcard] SET modifydate=GETDATE() WHERE classname='"& VIPCARDNUMBER &"'" )
	'不用了，可以通过视窗【View_isBoundVIPcard_AccountInformation_Records】判断出列表中的卡是否有人认领！通过classname关联卡号number。
	
	
	
	'-------------------------------------------------JSON system
	'JSON输出.
	'JSON system
	'创建JSON之ASP对象.
	Set ResultJSON = jsObject()
	'JSON system
	
	'JSON system
	'构建结果集反馈信息.
	ResultJSON("isAjaxSuccessful")		="true"
	ResultJSON("theResult_true_false")	="true"
	ResultJSON("theAllInformation")		="恭喜您，VIP会员卡绑定帐号操作成功！<img src=/images/ico/small/accept.png />"
	strShow = ResultJSON.FlushNow		'输出形式2.
	'JSON system
	
	'JSON system
	'销毁JSON之ASP对象.
	Set ResultJSON = Nothing
	'JSON system
	'-------------------------------------------------JSON system
	
	addAccount_BindingMyVIPCard = strShow
	
End Function


'服务名称：		会员修改密码.
'服务描述：		向系统发送本会员的修改密码请求，我登录好并且能够提交修改密码了吗？请回答.
'输出用的方式：	JSON.
'被调用的方式：	Form Post Ajax.
'将要接收的参数：	1.无效的ID号.				id
'返回数据：		JSON数据对象
'				第一个对象属性，名为valid，值为false操作失败 或 true操作成功。
'				第二个对象属性，名为message，值为反馈信息。
Public Function addAccount_PasswordOnChange()
	'定义内部新变量进行内部操作.
	Dim strShow
	Dim userName,intID
	Dim sql,RS
	Dim oldpassword,password,repassword
	
	'初始化赋值.
	addAccount_PasswordOnChange	=False	'默认为False操作失败.
	userName	=CokeShow.filtPass(Session("username"))
	intID		=CokeShow.filtRequest("1")
	strShow		=""
	
	oldpassword	=CokeShow.filtPass(Request("oldpassword"))
	password	=CokeShow.filtPass(Request("password"))
	repassword	=CokeShow.filtPass(Request("repassword"))
	
	'判断有各种效性.
	'intID
	If isNumeric(intID) Then
		If CokeShow.CokeCint(intID)=0 Then
			'参数不正确，退出.操作失败.
			FoundErr=True
			ErrMsg=ErrMsg &"ID参数不正确！"
			Exit Function
		End If
		intID=CokeShow.CokeClng(intID)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_PasswordOnChange的参数ID不正确，无法修改密码！"
		Exit Function
	End If
	'userName
	If isNull(userName) Or isEmpty(userName) Or userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_PasswordOnChange的参数userName不正确，无法修改密码！"
		Exit Function
	End If
	If userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"帐号不能为空！"
		Exit Function
	Else
		If CokeShow.strLength(userName)>50 Or CokeShow.strLength(userName)<10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的Email帐号长度不应大于50个字符，也不应小于10个字符的！"
			Exit Function
		Else
			If CokeShow.IsValidEmail(userName)=False Then
				FoundErr=True
				ErrMsg=ErrMsg & "您的帐号(即Email)格式不正确！"
				Exit Function
			Else
				userName=userName
			End If
		End If
	End If
	
	'验证有没有留空未填的，再验证有没有超过字数的.
	If oldpassword="" Or password="" Or repassword="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "新旧密码与确认密码均不能为空！"
		Exit Function
	Else
		If (CokeShow.strLength(oldpassword)>20 Or CokeShow.strLength(oldpassword)<6) Or (CokeShow.strLength(password)>20 Or CokeShow.strLength(password)<6) Then
			FoundErr=True
			ErrMsg=ErrMsg & "密码长度均不能大于20个字符，也不能小于6个字符！"
			Exit Function
		Else
			password=password
		End If
	End If
	
	'验证有没有新密码和确认密码不符合的.
	If password<>repassword Then
		FoundErr=True
		ErrMsg=ErrMsg & "新密码和确认密码不一致！"
		Exit Function
	End If
	
	'检测帐号旧密码是否正确.
	If CokeShow.CheckUserPassword( userName, oldpassword )=False Then
		FoundErr=True
		ErrMsg=ErrMsg & "您的旧密码不正确，请您重新输入！"
		Exit Function
	End If
	
	'验证码
	If Not CokeShow.CodePassII("CodeStr") Then
		FoundErr=True
		ErrMsg=ErrMsg &"验证码不正确哦，如果仍然不正确，请点击重新刷新验证码。"
		Exit Function
	End If
	
	
	'-----------------Go Begin
	'新增会员兑换的礼品券.
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT * FROM [CXBG_account] WHERE deleted=0 And username='"& userName &"'"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	
	If RS.Bof And RS.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "系统出现异常了！进入自我保护状态."
		Exit Function
		
	Else	
		
		
		RS("password")			=md5(password)
		
		
		
	RS.Update
	
	End If
	
	RS.Close
	Set RS=Nothing
	'-----------------Go End
	
	
	'-------------------------------------------------JSON system
	'JSON输出.
	'JSON system
	'创建JSON之ASP对象.
	Set ResultJSON = jsObject()
	'JSON system
	
	'JSON system
	'构建结果集反馈信息.
	ResultJSON("isAjaxSuccessful")		="true"
	ResultJSON("theResult_true_false")	="true"
	ResultJSON("theAllInformation")		="恭喜您，密码修改成功！<img src=/images/ico/small/accept.png />"
	strShow = ResultJSON.FlushNow		'输出形式2.
	'JSON system
	
	'JSON system
	'销毁JSON之ASP对象.
	Set ResultJSON = Nothing
	'JSON system
	'-------------------------------------------------JSON system
	
	addAccount_PasswordOnChange = strShow
	
End Function


'服务名称：		会员修改个人资料.
'服务描述：		向系统发送本会员的修改个人资料请求，我登录好并且能够提交修改个人资料了吗？请回答.
'输出用的方式：	JSON.
'被调用的方式：	Form Post Ajax.
'将要接收的参数：	1.无效的ID号.				id
'返回数据：		JSON数据对象
'				第一个对象属性，名为valid，值为false操作失败 或 true操作成功。
'				第二个对象属性，名为message，值为反馈信息。
Public Function addAccount_PersonalInformation()
	'定义内部新变量进行内部操作.
	Dim strShow
	Dim userName,intID
	Dim sql,RS
	Dim cnname,Birthday,Sex,province,city
	Dim client_name,client_telephone,client_schooling,client_memberoffamily,client_befondof,client_MonthlyIncome,client_work
	
	'初始化赋值.
	addAccount_PersonalInformation	=False	'默认为False操作失败.
	userName	=CokeShow.filtPass(Session("username"))
	intID		=CokeShow.filtRequest("1")
	strShow		=""
	
	cnname		=CokeShow.filtRequest(Request("cnname"))
	Birthday	=CokeShow.filtRequest(Request("selectyCoke")) &"-"& CokeShow.filtRequest(Request("selectmCoke")) &"-"& CokeShow.filtRequest(Request("selectdCoke"))
	province	=CokeShow.filtRequest(Request("province"))
	city		=CokeShow.filtRequest(Request("city"))
	Sex			=CokeShow.filtRequest(Request("Sex"))
	
	client_name				=CokeShow.filtRequest(Request("client_name"))
	client_telephone		=CokeShow.filtRequest(Request("client_telephone"))
	client_schooling		=CokeShow.filtRequest(Request("client_schooling"))
	client_memberoffamily	=CokeShow.filtRequest(Request("client_memberoffamily"))
	client_befondof			=CokeShow.filtRequest(Request("client_befondof"))
	client_MonthlyIncome	=CokeShow.filtRequest(Request("client_MonthlyIncome"))
	client_work				=CokeShow.filtRequest(Request("client_work"))
	
	'判断有各种效性.
	'intID
	If isNumeric(intID) Then
		If CokeShow.CokeCint(intID)=0 Then
			'参数不正确，退出.操作失败.
			FoundErr=True
			ErrMsg=ErrMsg &"ID参数不正确！"
			Exit Function
		End If
		intID=CokeShow.CokeClng(intID)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_PersonalInformation的参数ID不正确，无法修改密码！"
		Exit Function
	End If
	'userName
	If isNull(userName) Or isEmpty(userName) Or userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services,服务addAccount_PersonalInformation的参数userName不正确，无法修改密码！"
		Exit Function
	End If
	If userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"帐号不能为空！"
		Exit Function
	Else
		If CokeShow.strLength(userName)>50 Or CokeShow.strLength(userName)<10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的Email帐号长度不应大于50个字符，也不应小于10个字符的！"
			Exit Function
		Else
			If CokeShow.IsValidEmail(userName)=False Then
				FoundErr=True
				ErrMsg=ErrMsg & "您的帐号(即Email)格式不正确！"
				Exit Function
			Else
				userName=userName
			End If
		End If
	End If
	
	'验证.
	If cnname<>"" Then
		If CokeShow.strLength(cnname)>8 Then
			FoundErr=True
			ErrMsg=ErrMsg &"昵称只能8个字之内！"
			Exit Function
		Else
			cnname=cnname
		End If
	Else
		'不填写时的默认值.
		cnname=""
	End If
	
	If Replace(Birthday,"-","")<>"" Then
		If CokeShow.strLength(Birthday)>10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"生日日期错误！"
			Exit Function
		Else
			If isDate(Birthday)=False Then
				FoundErr=True
				ErrMsg=ErrMsg &"您的生日日期的格式不正确！"
				Exit Function
			Else
				Birthday=Birthday
			End If
		End If
	Else
		'不填写时的默认值.
		Birthday=""
	End If
	
	If Sex<>"" Then
		If CokeShow.strLength(Sex)>1 Then
			FoundErr=True
			ErrMsg=ErrMsg &"性别的参数错误！"
			Exit Function
		Else
			If isNumeric(Sex)=False Then
				FoundErr=True
				ErrMsg=ErrMsg &"您的性别的参数格式不正确！"
				Exit Function
			Else
				Sex=CokeShow.CokeCint(Sex)
			End If
		End If
	Else
		'不填写时的默认值.
		Sex=0
	End If
	
	If province<>"" Then
		If CokeShow.strLength(province)>10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"省份应该在10位字符之内！此项也可以不填！"
			Exit Function
		Else
			province=province
		End If
	Else
		'不填写时的默认值.
		province=""
	End If
	If city<>"" Then
		If CokeShow.strLength(city)>10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"城市应该在10位字符之内！此项也可以不填。"
			Exit Function
		Else
			city=city
		End If
	Else
		'不填写时的默认值.
		city=""
	End If
	
	If client_name<>"" Then
		If CokeShow.strLength(client_name)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的姓名只能50位字符之内！此项也可以不填。"
			Exit Function
		Else
			client_name=client_name
		End If
	Else
		'不填写时的默认值.
		client_name=""
	End If
	If client_telephone<>"" Then
		If CokeShow.strLength(client_telephone)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"电话应该在50位字符之内！此项也可以不填写。"
			Exit Function
		Else
			client_telephone=client_telephone
		End If
	Else
		'不填写时的默认值.
		client_telephone=""
	End If
	
	If client_schooling<>"" Then
		If isNumeric(client_schooling)=False Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的学历的参数格式不正确！"
			Exit Function
		Else
			client_schooling=CokeShow.CokeClng(client_schooling)
		End If
	Else
		'不填写时的默认值.
		client_schooling=0
	End If
	
	If client_memberoffamily<>"" Then
		If CokeShow.strLength(client_memberoffamily)>10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"家庭成员应该在10位字符之内！此项也可以不选。"
			Exit Function
		Else
			client_memberoffamily=client_memberoffamily
		End If
	Else
		'不填写时的默认值.
		client_memberoffamily=""
	End If
	
	If client_befondof<>"" Then
		If CokeShow.strLength(client_befondof)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"爱好应该在50位字符之内！此项也可以不填写。"
			Exit Function
		Else
			client_befondof=client_befondof
		End If
	Else
		'不填写时的默认值.
		client_befondof=""
	End If
	
	If client_MonthlyIncome<>"" Then
		If isNumeric(client_MonthlyIncome)=False Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的月收入的参数格式不正确！"
			Exit Function
		Else
			client_MonthlyIncome=CokeShow.CokeClng(client_MonthlyIncome)
		End If
	Else
		'不填写时的默认值.
		client_MonthlyIncome=0
	End If
	
	If client_work<>"" Then
		If isNumeric(client_work)=False Then
			FoundErr=True
			ErrMsg=ErrMsg &"您的职业的参数格式不正确！"
			Exit Function
		Else
			client_work=CokeShow.CokeClng(client_work)
		End If
	Else
		'不填写时的默认值.
		client_work=0
	End If
	
	'验证码
	If Not CokeShow.CodePassII("CodeStr") Then
		FoundErr=True
		ErrMsg=ErrMsg &"验证码不正确哦，如果仍然不正确，请点击重新刷新验证码。"
		Exit Function
	End If
	
	
	'-----------------Go Begin
	'新增会员兑换的礼品券.
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT * FROM [CXBG_account] WHERE deleted=0 And username='"& userName &"'"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	
	If RS.Bof And RS.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "系统出现异常了！进入自我保护状态."
		Exit Function
		
	Else	
		
		
		RS("cnname")		=cnname
		If isDate(Birthday) Then RS("Birthday")=Birthday
		RS("Sex")			=Sex
		RS("province")		=province
		RS("city")			=city
		
		RS("client_name")				=client_name
		RS("client_telephone")			=client_telephone
		RS("client_schooling")			=client_schooling
		RS("client_memberoffamily")		=client_memberoffamily
		RS("client_befondof")			=client_befondof
		RS("client_MonthlyIncome")		=client_MonthlyIncome
		RS("client_work")				=client_work
		
		
		
	RS.Update
	
	End If
	
	RS.Close
	Set RS=Nothing
	'-----------------Go End
	
	
	
	'积分运算.Begin
		'选填项-所在地区.
		If province<>"" And city<>"" And addAccount_PersonalInformation___CheckJifenIsUsed(userName,"进行了选填项-所在地区的填写")=False Then
			If CokeShow.JifenSystemExecute(1,userName,6,"进行了选填项-所在地区的填写,您已成功获得了餐厅送出的6网站积分.",0)=True Then Response.Write "" Else Response.Write "积分处理失败！<br />"
		End If
		'选填项-您的姓名.
		If client_name<>"" And addAccount_PersonalInformation___CheckJifenIsUsed(userName,"进行了选填项-您的姓名的填写")=False Then
			If CokeShow.JifenSystemExecute(1,userName,6,"进行了选填项-您的姓名的填写,您已成功获得了餐厅送出的6网站积分.",0)=True Then Response.Write "" Else Response.Write "积分处理失败！<br />"
		End If
		'选填项-电话.
		If client_telephone<>"" And addAccount_PersonalInformation___CheckJifenIsUsed(userName,"进行了选填项-电话的填写")=False Then
			If CokeShow.JifenSystemExecute(1,userName,6,"进行了选填项-电话的填写,您已成功获得了餐厅送出的6网站积分.",0)=True Then Response.Write "" Else Response.Write "积分处理失败！<br />"
		End If
		'选填项-学历.
		If client_schooling>0 And addAccount_PersonalInformation___CheckJifenIsUsed(userName,"进行了选填项-学历的填写")=False Then
			If CokeShow.JifenSystemExecute(1,userName,6,"进行了选填项-学历的填写,您已成功获得了餐厅送出的6网站积分.",0)=True Then Response.Write "" Else Response.Write "积分处理失败！<br />"
		End If
		'选填项-家庭成员.
		If client_memberoffamily<>"" And addAccount_PersonalInformation___CheckJifenIsUsed(userName,"进行了选填项-家庭成员的填写")=False Then
			If CokeShow.JifenSystemExecute(1,userName,6,"进行了选填项-家庭成员的填写,您已成功获得了餐厅送出的6网站积分.",0)=True Then Response.Write "" Else Response.Write "积分处理失败！<br />"
		End If
		'选填项-爱好.
		If client_befondof<>"" And addAccount_PersonalInformation___CheckJifenIsUsed(userName,"进行了选填项-爱好的填写")=False Then
			If CokeShow.JifenSystemExecute(1,userName,6,"进行了选填项-爱好的填写,您已成功获得了餐厅送出的6网站积分.",0)=True Then Response.Write "" Else Response.Write "积分处理失败！<br />"
		End If
		'选填项-月收入.
		If client_MonthlyIncome>0 And addAccount_PersonalInformation___CheckJifenIsUsed(userName,"进行了选填项-月收入的填写")=False Then
			If CokeShow.JifenSystemExecute(1,userName,6,"进行了选填项-月收入的填写,您已成功获得了餐厅送出的6网站积分.",0)=True Then Response.Write "" Else Response.Write "积分处理失败！<br />"
		End If
		'选填项-职业.
		If client_work>0 And addAccount_PersonalInformation___CheckJifenIsUsed(userName,"进行了选填项-职业的填写")=False Then
			If CokeShow.JifenSystemExecute(1,userName,6,"进行了选填项-职业的填写,您已成功获得了餐厅送出的6网站积分.",0)=True Then Response.Write "" Else Response.Write "积分处理失败！<br />"
		End If
	'积分运算.End
	
	
	
	'-------------------------------------------------JSON system
	'JSON输出.
	'JSON system
	'创建JSON之ASP对象.
	Set ResultJSON = jsObject()
	'JSON system
	
	'JSON system
	'构建结果集反馈信息.
	ResultJSON("isAjaxSuccessful")		="true"
	ResultJSON("theResult_true_false")	="true"
	ResultJSON("theAllInformation")		="恭喜您，您的个人资料修改成功！<img src=/images/ico/small/accept.png />"
	strShow = ResultJSON.FlushNow		'输出形式2.
	'JSON system
	
	'JSON system
	'销毁JSON之ASP对象.
	Set ResultJSON = Nothing
	'JSON system
	'-------------------------------------------------JSON system
	
	addAccount_PersonalInformation = strShow
	
End Function
'检测是否已经获得过某个积分.
'addAccount_PersonalInformation___CheckJifenIsUsed(userName,"进行了选填项-月收入的填写")=True/False
Public Function addAccount_PersonalInformation___CheckJifenIsUsed(paraUserName,paraStrJifenDescription)
	addAccount_PersonalInformation___checkJifenIsUsed=False		'默认没有获得过积分.
	
	Dim rsCheckJifenIsUsed,sqlCheckJifenIsUsed
	sqlCheckJifenIsUsed="SELECT * FROM [CXBG_account_JifenSystem] WHERE deleted=0 AND Account_LoginID='"& paraUserName &"' AND JifenDescription LIKE '%"& paraStrJifenDescription &"%'"
	Set rsCheckJifenIsUsed=CONN.Execute(sqlCheckJifenIsUsed)
	
	If (rsCheckJifenIsUsed.Bof And rsCheckJifenIsUsed.Eof)=False Then
		'已经获得过积分.	
		addAccount_PersonalInformation___CheckJifenIsUsed=True	'已经获得过积分!
		
		rsCheckJifenIsUsed.Close
		Set rsCheckJifenIsUsed=Nothing
		
		Exit Function
	End If
	
	rsCheckJifenIsUsed.Close
	Set rsCheckJifenIsUsed=Nothing
End Function
%>