<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True	'缓冲数据，才向客户输出；
Response.Charset="utf-8"
Session.CodePage = 65001
%>
<%
'模块说明：错误处理页面.
'日期说明：2010-4-x
'版权说明：www.cokeshow.com.cn	可乐秀(北京)技术有限公司产品
'友情提示：本产品已经申请专利，请勿用于商业用途，如果需要帮助请联系可乐秀。
'联系电话：(010)-67659219
'电子邮件：cokeshow@qq.com
%>
<!--#include file="system/system_conn.asp"-->
<!--#include file="system/system_class.asp"-->

<!--#include file="system/foreground_class.asp"-->

<%
'类实例化
Dim CokeShow
Set CokeShow = New SystemClass
CokeShow.Start		'调用类的Start方法，初始化类里的ReloadSetup()函数，并得到二维数组Setup
Call CokeShow.SQLWarningSys()	'预警.

'前台类实例化
Dim Coke
Set Coke = New ForegroundClass
Coke.Start
%>
<%
'发送邮件，获取404的详细资料.
Dim strVisitURL,strRefererURL
'用户试图访问的页面地址.
'strVisitURL		=Right( Request.QueryString, Len(Request.QueryString)-4 )	'原装是(404;http://localhost/niceforms-default.css)。
strVisitURL		=Request.ServerVariables("SCRIPT_NAME")
'用户试图访问页面之前的 点击来源页面地址.
strRefererURL	=Request.ServerVariables("HTTP_REFERER")


'将错误报告给支持小组.
'定义标题和内容.
Dim Topic,LogText
'标题.
Topic	="可乐秀COKESHOW系统邮件通知：客户CXBG发生了404错误报告 On "& Now() &" From: "& strVisitURL
'内容.
LogText	="在"& Now() &"的时候，一个404错误发生在CXBG客户的网站上。当顾客访客正在访问页面(<a href="""& strVisitURL &""" target=""_blank"">"&  strVisitURL &"</a>)时发生了此次错误。其中引用页来源页是(<a href="""& strRefererURL &""" target=""_blank"">"&  strRefererURL &"</a>)。如下是更为详尽的信息报告：<br />"
'循环写下详尽参数信息.
Dim strName
For Each strName In Request.ServerVariables
	LogText=LogText &"<br />"& strName &"---"& Request.ServerVariables(strName) &"<br />"& vbCrLf
Next
LogText=Replace(LogText,"User-Agent","<b>User-Agent</b>")
LogText=Replace(LogText,"HTTP_USER_AGENT","<b>HTTP_USER_AGENT</b>")
LogText=Replace(LogText,"SCRIPT_NAME","<b>SCRIPT_NAME</b>")
LogText=Replace(LogText,"REQUEST_METHOD","<b>REQUEST_METHOD</b>")
LogText=Replace(LogText,"REMOTE_ADDR","<b>REMOTE_ADDR</b>")
LogText=Replace(LogText,"REMOTE_HOST","<b>REMOTE_HOST</b>")
LogText=Replace(LogText,"LOCAL_ADDR","<b>LOCAL_ADDR</b>")
LogText=Replace(LogText,"HTTP_REFERER","<b>HTTP_REFERER</b>")
LogText=Replace(LogText,"cokeshow.com.cn","<b>cokeshow.com.cn</b>")
LogText=Replace(LogText,"HTTP_CONNECTION","<b>HTTP_CONNECTION</b>")
LogText=Replace(LogText,"HTTP_COOKIE","<b>HTTP_COOKIE</b>")


'构造模板 b
Dim strLogText
strLogText=strLogText &"<style>A:visited {	TEXT-DECORATION: none	}"
strLogText=strLogText &"A:active  {	TEXT-DECORATION: none	}"
strLogText=strLogText &"A:hover   {	TEXT-DECORATION: underline	}"
strLogText=strLogText &"A:link 	  {	text-decoration: none;}"
strLogText=strLogText &"BODY   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}"
strLogText=strLogText &"TD	   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt	}</style>"
strLogText=strLogText &"<TABLE border=0 width='95%' align=center><TBODY><TR><TD>"

strLogText=strLogText &"尊敬的可乐秀COKESHOW负责人，"
strLogText=strLogText & LogText 
strLogText=strLogText &"<br /><br /><br /><br /><br /><br />"

strLogText=strLogText &"<p>可乐秀COKESHOW-互联网品牌升级服务&nbsp;&nbsp;&nbsp;<a href="""& "http://www.cokeshow.com.cn/" &""" target=""_blank"">"& "http://www.cokeshow.com.cn/" &"</a></p>"
'strLogText=strLogText &"*************************************************************************************<br />"
'strLogText=strLogText &"此邮件来自  <a href=http://www.cokeshow.com.cn target=_blank>可乐秀CokeShow (http://www.cokeshow.com.cn)</a><br />"
'strLogText=strLogText &"如果您需要回复，请发邮件至 cokeshow@qq.com<br />"
'strLogText=strLogText &"*************************************************************************************"

strLogText=strLogText &"</TD></TR></TBODY></TABLE>"
'构造模板 e

'检查404报错系统是否开启.
If is_cokeshow_404ErrorAlert_system=1 Then
	'发送邮件！
	Dim isSendOK
	isSendOK=""
	'针对IE8的浏览器泄漏js问题而打的补丁.
	If Not Instr(strVisitURL, "/[object]")>0 Then	'只要不含有“[object]”字样的IE8漏洞错误频繁报告，都可以发送邮件给可乐秀管理员。
		If CokeShow.SendMail("44533122@qq.com","痴心不改餐厅",system_ReplyEmailAddress,Topic,strLogText,"gb2312","text/html",system_JMailFrom,system_JMailSMTP,system_JMailMailServerUserName,system_JMailMailServerPassWord)=True Then
			'Response.Write "<br />发送成功！谢谢您一如既往的关注.<br />"
			isSendOK="100%"
		End If
	End If
End If
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>您找的页面不存在---<% =CokeShow.otherField("[CXBG_controller]",1,"ID","site_title",True,0) %><% =isSendOK %></title>
    
    <style type="text/css">
        <!--
        BODY {	PADDING-RIGHT: 0px; PADDING-LEFT: 35px; BACKGROUND: url() repeat-x left top; PADDING-BOTTOM: 0px; MARGIN: 0px; FONT: 12px Arial, Helvetica, sans-serif; COLOR: #333; PADDING-TOP: 35px}
        A {	COLOR: #007ab7; TEXT-DECORATION: none}
        A:hover {COLOR: #007ab7; TEXT-DECORATION: none}
        A:hover {COLOR: #de1d6a}
        .hidehr {DISPLAY: none}
        .show12 {PADDING-RIGHT: 0px; DISPLAY: block; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 5px 0px; PADDING-TOP: 0px}
        .show13 {PADDING-RIGHT: 0px; DISPLAY: block; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 5px 0px; PADDING-TOP: 0px}
        .show12 A {	BORDER-RIGHT: #bfdeed 1px solid; PADDING-RIGHT: 6px; BORDER-TOP: #bfdeed 1px solid; DISPLAY: inline-block; PADDING-LEFT: 6px; BACKGROUND: #d8ebf4; PADDING-BOTTOM: 2px; OVERFLOW: hidden; BORDER-LEFT: #bfdeed 1px solid; LINE-HEIGHT: 17px; PADDING-TOP: 2px; BORDER-BOTTOM: #bfdeed 1px solid; HEIGHT: 16px}
        .show13 A {	BORDER-RIGHT: #bfdeed 1px solid; PADDING-RIGHT: 6px; BORDER-TOP: #bfdeed 1px solid; DISPLAY: inline-block; PADDING-LEFT: 6px; BACKGROUND: #d8ebf4; PADDING-BOTTOM: 2px; OVERFLOW: hidden; BORDER-LEFT: #bfdeed 1px solid; LINE-HEIGHT: 17px; PADDING-TOP: 2px; BORDER-BOTTOM: #bfdeed 1px solid; HEIGHT: 16px}
        .show12 A:hover {	BORDER-RIGHT: #ea5e96 1px solid; BORDER-TOP: #ea5e96 1px solid; BACKGROUND: #fce8f0; BORDER-LEFT: #ea5e96 1px solid; COLOR: #de1d6a; BORDER-BOTTOM: #ea5e96 1px solid; TEXT-DECORATION: none}
        .show13 A:hover {	BORDER-RIGHT: #ea5e96 1px solid; BORDER-TOP: #ea5e96 1px solid; BACKGROUND: #fce8f0; BORDER-LEFT: #ea5e96 1px solid; COLOR: #de1d6a; BORDER-BOTTOM: #ea5e96 1px solid; TEXT-DECORATION: none}
        .show12 A {	FONT-WEIGHT: normal; FONT-SIZE: 12px}
        .show13 A {	BORDER-RIGHT: #ed268c 1px solid; BORDER-TOP: #ed268c 1px solid; BACKGROUND: #dd137b; BORDER-LEFT: #ed268c 1px solid; COLOR: #fff; BORDER-BOTTOM: #ed268c 1px solid}
        .img404 {PADDING-RIGHT: 0px; PADDING-LEFT: 0px; BACKGROUND: url(/images/404.gif) no-repeat left top; FLOAT: left; PADDING-BOTTOM: 0px; MARGIN: 0px; WIDTH: 80px; PADDING-TOP: 0px; POSITION: relative; HEIGHT: 90px}
        H2 {PADDING-RIGHT: 0px; PADDING-LEFT: 0px; FONT-SIZE: 16px; FLOAT: left; PADDING-BOTTOM: 25px; MARGIN: 0px; WIDTH: 80%; LINE-HEIGHT: 0; PADDING-TOP: 25px; BORDER-BOTTOM: #ccc 1px solid; POSITION: relative}
        H3.wearesorry {	PADDING-RIGHT: 0px; PADDING-LEFT: 0px; FONT-WEIGHT: normal; FONT-SIZE: 10px; LEFT: 117px; PADDING-BOTTOM: 0px; MARGIN: 0px; COLOR: #ccc; LINE-HEIGHT: 10px; PADDING-TOP: 0px; POSITION: absolute; TOP: 70px}
        .content {	CLEAR: both; PADDING-RIGHT: 0px; PADDING-LEFT: 0px; FONT-SIZE: 13px; LEFT: 80px; FLOAT: left; PADDING-BOTTOM: 0px; MARGIN: 0px; WIDTH: 80%; LINE-HEIGHT: 19px; PADDING-TOP: 0px; POSITION: relative; TOP: -30px}
        .content UL {PADDING-RIGHT: 35px; PADDING-LEFT: 35px; PADDING-BOTTOM: 20px; MARGIN: 0px; PADDING-TOP: 10px}
        .show12 UL {PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 0px 0px 0px 20px; PADDING-TOP: 0px; LIST-STYLE-TYPE: none}
        .show14 UL LI {	MARGIN-BOTTOM: 5px}
        
        -->
    </style>
</head>
<noscript><br />Come From MyHomestay-Developer Team<br /></noscript>
<noscript><b>由 MyHomestay 原班创业团队开发设计制作，欢迎与CokeShow.com.cn联系.</b><br /></noscript>
<noscript>BeiJing.China e-mail:cokeshow@qq.com</noscript>
<body>
    

    <DIV class=img404>　</DIV>
    <H2>十分抱歉，痴心不改餐厅找不到您要的页面……</H2>
    <H3 class=wearesorry>We're sorry but the page your are looking for is Not 
    Found...</H3>
    <DIV class=content>我们已经为您仔细的找过啦，没有发现您要找的页面。最可能的原因是： 
      <UL>
      <LI>在地址中可能存在键入错误。 
      <LI>当你点击某个链接时，它可能已过期。 
    </LI></UL>
    <STRONG>给您造成不便，我们深感抱歉。此时，您何不看看我们餐厅网站的这些内容呢：）</STRONG>（<a href="/ChineseDish/ChineseDish.Welcome?StarRating=True">星级菜推荐</a>）： 
    <DIV class=show14>
    <UL>
      <LI><A title=返回餐厅首页 href="/">返回餐厅首页</A> 
      <LI><A title=返回上一个页面 href="javascript:history.back(-1)">返回上一页</A></LI>
    </UL></DIV>
    要不，我们去痴心不改餐厅的好伙伴：<A href="http://www.cokeshow.com.cn/" target=_blank><u>可乐秀CokeShow-商务网站设计顾问</u></A>官方网站看看吧~~ 将您发现的技术问题告诉他们听！可以帮助餐厅快速的解决您发现的问题哦。</DIV>
    <SPAN 
    style="VISIBILITY: hidden"></SPAN>
	
</body>
</html>