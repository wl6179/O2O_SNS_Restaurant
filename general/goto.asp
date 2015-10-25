<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True	'缓冲数据，才向客户输出；
Response.Charset="utf-8"
Session.CodePage = 65001
%>
<%
'模块说明：后台管理人员帐号管理模块.
'日期说明：2009-7-7
'版权说明：www.cokeshow.com.cn	可乐秀(北京)技术有限公司产品
'友情提示：本产品已经申请专利，请勿用于商业用途，如果需要帮助请联系可乐秀技术有限公司。
'联系电话：(010)-67659219
'电子邮件：cokeshow@qq.com
%>
<!--#include file="inc/_public.asp"-->
<!--#include file="_works.asp"-->

<%
'变量定义区.
'(用来存储对象的变量，用全大写!)
Const maxPerPage=15							'当前模块分页设置.
Dim CurrentPageNow,TitleName,UnitName
CurrentPageNow 	= "goto.asp"			'当前页.
TitleName 		= "进入会员管理后台中介模块"				'此模块管理页的名字.
UnitName 		= "进入会员管理后台中介模块"					'此模块涉及记录的元素名.
'自定义设置.
'本地设置.
Dim CurrentTableName
CurrentTableName 	= "[CXBG_account]"		'此模块涉及的[表]名.
%>



<%
Dim totalPut,totalPages,currentPage			'分页用的控制变量.
Dim RS, sql									'查询列表记录用的变量.
Dim FoundErr,ErrMsg							'控制错误流程用的控制变量.
Dim strFileName								'构建查询字符串用的控制变量.
Dim ExecuteSearch,Keyword,TypeSearch,Action	'构建查询字符串以及流程控制用的控制变量.
Dim strGuide		'导航文字.
DIM username

currentPage		=CokeShow.filtRequest(Request("Page"))
ExecuteSearch	=CokeShow.filtRequest(Request("ExecuteSearch"))
Keyword			=CokeShow.filtRequest(Request("Keyword"))
TypeSearch		=CokeShow.filtRequest(Request("TypeSearch"))
Action			=CokeShow.filtRequest(Request("Action"))


Dim intID
intID=CokeShow.filtRequest(Request("id"))
'处理id传值
If intID="" Then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
Else
	intID=CokeShow.CokeClng(intID)
End If


username		=CokeShow.otherField("[CXBG_account]",intID,"id","username",True,999)





		If Action="Add" Then
			Call Add()
		ElseIf Action="SaveAdd" Then
			Call SaveAdd()
		ElseIf Action="Modify" Then
			Call Modify()
		ElseIf Action="SaveModify" Then
			Call SaveModify()
		ElseIf Action="Delete" Then
			Call Delete()
		ElseIf Action="Lock" Then
			Call Lock()
		ElseIf Action="UnLock" Then
			Call UnLock()
		
		
		
		ElseIf Action="GoTo" Then
			Call GoToEnter()
		
		Else
			'Call Main()
		End If
		
		
		If FoundErr=True Then
			CokeShow.AlertErrMsg_general( ErrMsg )
		End If



'进入会员后台的中介桥梁函数.

Sub GoToEnter()

'记入日志.
Call CokeShow.AddLog("成功进入了帐号为"& username &"的会员用户中心后台！", "")

	'先发放一个临时会话变量，凭此可以免费登录！
	Session("YouCanLoginID_Temp")="CokeShow"& CokeShow.GetRandomizeCode
	
	Response.Redirect "/ONCEFOREVER/?Action=RegisterSuccessUI&username="& username &"&ID_Temp="& Replace(CokeShow.PassEncode( Ucase(Md5( Session("YouCanLoginID_Temp") & username )) ),"%","")
	'免费登录算法：CokeShow.AddCode_Num( Ucase(Md5( Session("YouCanLoginID_Temp") & CokeShow.filtPass(Request("username")) )) )
	'1.捆绑大于6字符的Session("YouCanLoginID_Temp")值+帐号值在一起； 2.对其进行MD5加密； 3.对其全大写化； 4.最后对其进行PassEncode加密+去除百分号%。 5.对照一下谁能够传递过来这样的匹配字符串，就允许其登录相应的会员帐号！

	
	
End Sub


%>