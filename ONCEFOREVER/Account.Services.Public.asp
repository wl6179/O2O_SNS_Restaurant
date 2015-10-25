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
'Ajax远程调用asp文件.供所有程序包括客户端在内的所有人调用，是公有调用服务.（所以不要放置能查询出敏感信息的功能！）
'
%>﻿
<!--#include virtual="/system/system_conn.asp"-->
<!--#include virtual="/system/system_class.asp"-->

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
'逻辑判断区
SELECT CASE ServicesAction
	Case "CheckAccountName"
		strServiceFeedback = CheckAccountName()	'所有函数均不用para参数，每个function独立的直接用request获取想要的参数！
	Case "SubmitQuestionnaires"
		strServiceFeedback = SubmitQuestionnaires()	'所有函数均不用para参数，每个function独立的直接用request获取想要的参数！
		
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



'如果没有什么错误，那么就输出特定函数的输出字符串(JSON描述)
Response.Write strServiceFeedback	'被JSON系统替换掉了.
Response.End()
'结束Services.
%>

<%
'具体服务区.

'服务名称：		检测会员帐号同名情况.true表示通过，false表示失败！
'服务描述：		询问系统是否有同名的会员帐号？请回答.
'输出用的方式：	JSON.
'被调用的方式：	URL Get Ajax.
'将要接收的参数：	1.有username，为会员帐号信息.
'返回数据：		JSON数据对象
'				第一个对象属性，名为valid，值为false或true。
Public Function CheckAccountName()
	
	'定义内部新变量进行内部操作.
	Dim strShow
	Dim userName
	Dim sql,RS
	
	'初始化赋值.
	CheckAccountName	=False	'默认为False.
	userName			=CokeShow.filtPass(Request("username"))
	
	'判断有各种效性.
'	If isNumeric(CurrentClassID) Then
'		CurrentClassID=CokeShow.CokeClng(CurrentClassID)
'	Else
'		'参数不正确，退出.操作失败.
'		FoundErr=True
'		ErrMsg=ErrMsg &"Account.Services.Public,服务CheckAccountName的参数ID不正确，无法获取检测会员帐号同名情况的操作！"
'		Exit Function
'	End If
	If isNull(userName) Or isEmpty(userName) Or userName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services.Public,服务CheckAccountName的参数userName不正确，无法获取检测会员帐号同名情况的操作！"
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
	
	'-----------------Go Begin
	'检测帐号.
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT TOP 1 * FROM [CXBG_account] WHERE username='"& userName &"'"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN
	
	'JSON system
	'创建JSON之ASP对象.
	Set ResultJSON = jsObject()
	'JSON system
	
	
	
	If RS.Eof Then				'用户名合法，没有同名！
		'strShow = "{theResult_true_false: ""false""}"		'theResult_true_false[同名是否为真]	'被JSON系统替换掉了.
		'JSON system
		'构建结果集反馈信息.
		ResultJSON("isAjaxSuccessful")		="true"
		ResultJSON("theResult_true_false")	="true"
		ResultJSON("theAllInformation")		="<img src='/images/yes.gif' /> 恭喜您，此会员帐号可以使用<img src='/images/ico/small/emotion_happy.png' />"
		strShow = ResultJSON.FlushNow	'输出形式2.
		'JSON system
	Else						'用户名不合法.有同名！
		'strShow = "{theResult_true_false: ""true""}"	'被JSON系统替换掉了.
		'JSON system
		'构建结果集反馈信息.
		ResultJSON("isAjaxSuccessful")		="true"
		ResultJSON("theResult_true_false")	="false"
		ResultJSON("theAllInformation")		="<img src='/images/del.gif' /> 此会员帐号已存在<img src='/images/ico/small/emotion_suprised.png' />，请用其它的Email地址！"
		strShow = ResultJSON.FlushNow	'输出形式2.
		'JSON system
	End If
	
	'JSON system
	'销毁JSON之ASP对象.
	Set ResultJSON = Nothing
	'JSON system
	
	RS.Close
	Set RS=Nothing
	'-----------------Go End
	
	CheckAccountName = strShow
	
End Function


'服务名称：		提交调查问卷.
'服务描述：		询问系统是否成功提交了我的问卷？请回答.
'输出用的方式：	JSON.
'被调用的方式：	Form POST Ajax.
'将要接收的参数：	1.有QuestionnairesID，为填写调查表问卷的选择选项id号.
'返回数据：		JSON数据对象
Public Function SubmitQuestionnaires()
	
	'定义内部新变量进行内部操作.
	Dim strShow
	Dim QuestionnairesID,id
	Dim sql,RS
	
	'初始化赋值.
	SubmitQuestionnaires	=False	'默认为False.
	QuestionnairesID		=CokeShow.filtRequest(Request("QuestionnairesID"))
	id						=CokeShow.filtRequest(Request("id"))
	
	'判断有各种效性.
	If isNumeric(QuestionnairesID) Then
		QuestionnairesID=CokeShow.CokeClng(QuestionnairesID)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"请您选择一个选择项！"
		Exit Function
	End If
	If isNumeric(id) Then
		id=CokeShow.CokeClng(id)
	Else
		'参数不正确，退出.操作失败.
		FoundErr=True
		ErrMsg=ErrMsg &"Account.Services.Public,服务SubmitQuestionnaires的参数ID不正确，无法获取提交调查问卷的操作！"
		Exit Function
	End If
	
	'限制重复提交.
	'如果没有访问过的痕迹，则允许通过.
	If isEmpty(Session("visitor_services_SubmitQuestionnaires_datetime_" )) Or isNull(Session("visitor_services_SubmitQuestionnaires_datetime_" )) Or Session("visitor_services_SubmitQuestionnaires_datetime_" )="" Then
		'通过~
	'如果有访问过的迹象，则不允许通过!
	Else
		'阻挡.
		FoundErr=True
		ErrMsg=ErrMsg &"您刚刚已经完成调查问卷的提交了！谢谢您的大力支持！"
		Exit Function
	End If
	'记下被当前访客访问的痕迹！
	Session("visitor_services_SubmitQuestionnaires_datetime_" )="gogogo"
	
	
	'-----------------Go Begin
	'新增收藏菜品.
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT TOP 1 * FROM [CXBG_Questionnaire_Result] WHERE deleted=0"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	
	RS.AddNew
		
		If CokeShow.CheckUserLogined()=True And isNumeric(Session("id")) And Len(Session("username"))>=10 Then
			RS("Account_LoginID")				=Session("username")
		End If
		RS("byask_QuestionnaireID")			=id
		RS("byselected_QuestionnaireID")	=QuestionnairesID
		RS("IPaddress")				=Request.ServerVariables("REMOTE_ADDR")
		RS("HTTP_REFERER")			=Request.ServerVariables("HTTP_REFERER")
		RS("HTTP_GetAllUrlII")		=CokeShow.GetAllUrlII
	
		'-------------------------------------------------JSON system
		'JSON system
		'创建JSON之ASP对象.
		Set ResultJSON = jsObject()
		'JSON system
		
		'JSON system
		'构建结果集反馈信息.
		ResultJSON("isAjaxSuccessful")		="true"
		ResultJSON("theResult_true_false")	="true"
		ResultJSON("theAllInformation")		="恭喜您，调查问卷提交成功，感谢您的细心参与!"
		strShow = ResultJSON.FlushNow		'输出形式2.
		'JSON system
		
		'JSON system
		'销毁JSON之ASP对象.
		Set ResultJSON = Nothing
		'JSON system
		'-------------------------------------------------JSON system
		
	RS.Update
	
	RS.Close
	Set RS=Nothing
	'-----------------Go End
	
	SubmitQuestionnaires = strShow
	
End Function

%>