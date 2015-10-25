<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Charset="utf-8"
Session.CodePage = 65001
%>
<!--#include file="../system/system_conn.asp"-->
<!--#include file="../system/system_class.asp"-->
<!--#include file="../system/md5.asp"-->
<%
'类实例化
Dim CokeShow
Set CokeShow = New SystemClass
CokeShow.Start		'调用类的Start方法，初始化类里的ReloadSetup()函数，并得到二维数组Setup
Call CokeShow.SQLWarningSys()	'预警.
%>


<%
'变量定义.
Dim enterName,enterPassword
Dim FoundErr,ErrMsg
%>


<%
'处理退出.
If Request("Action")="Logout" Then
	Session("enterName")=""				'销毁登录标记.
	Session("enterId")=""				'销毁登录标记.
	Session("enterPassword")=""			'销毁登录标记.
	Session("enterCnName")=""			'销毁登录标记.
	Session("enterLevel")=""			'销毁登录标记.
	Session("isHaveWork_supervisor")=""	'销毁登录标记.
	
	Response.Redirect "_main.asp"	'验证不过，自动转向登录页面.
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>大后台登录</title>
	
	<link type="text/css" rel="stylesheet" href="<% =filename_dj_MainCss %>" />
	<link type="text/css" rel="stylesheet" href="<% =filename_dj_ThemesCss %>" />
    <!-- CSS -->
	<style type="text/css">
		@import "/style/UpdateStyle.css";
	</style>
    
    <link href="/css/other_yhzx.css" type="text/css" rel="stylesheet" />
    <link href="/css/dlzc.css" type="text/css" rel="stylesheet" />
    <link href="/css/cxbg.css" type="text/css" rel="stylesheet" />
	
	<link type="text/css" rel="stylesheet" href="../style/general_style.css" />
	
	<script type="text/javascript" src="../script/public.js"></script>
	
	<script type="text/javascript" src="<% =filename_dj %>" djConfig="parseOnLoad:<% =parseOnLoad_dj %>, isDebug:<% =isDebug_dj %>, debugAtAllCosts:<% =isDebug_dj %>"></script>
	<script type="text/javascript" src="<% =filenameWidgetsCompress_dj %>"></script>
	<script type="text/javascript">
		dojo.require("dojo.parser");
		dojo.require("dijit.form.ValidationTextBox");
		dojo.require("dijit.form.Button");
		dojo.require("dijit.form.Form");
		dojo.addOnLoad(function() {
			dojo.byId("enterName").focus();
		});
		
	</script>
	
	
	<script type="text/javascript" src="../script/imgpngPrecc.js"></script>
	
	
	
	<script type="text/javascript">
		function submitLogin() {
			dijit.byId("Login").submit();
			
			//dijit.byId("enterName").displayMessage(errorMessage);
			//dijit.byId("enterPassword").displayMessage(errorMessage);
			//dijit.byId("CodeStr").displayMessage(errorMessage);
		}
		
		function replaceGetCode() {
			dojo.byId("GetCode").src = "/public/code.asp?c=" + Math.random();
		}
	</script>

</head>

<noscript><br />Come From MyHomestay-Developer Team<br /></noscript>
<noscript><b>由 MyHomestay 原班创业团队开发设计制作，欢迎与CokeShow.com.cn联系.</b><br /></noscript>
<noscript>BeiJing.China e-mail:cokeshow@qq.com</noscript>

<body class="<% =classname_dj_ThemesCss_foreground %> cxbgbody">

<div id="cxbgbg_img">
 <div id="dlzc_imgmid">
  <div id="dlzc_headbg" style="background:url(/images/dlzc_img01A.jpg) repeat-y;">
	<span class="fontred" style="font-size:10px;">CokeShow<span class="font25"></span></span>
  </div>
<!--middle start-->
  <div class="dlzcmidbg">
    <div class="dlzccentbg">
      <div class="hydlzcbt">大后台管理登录：</div>
	  <div class="hyzctable" style="height:auto;">
	  
<!--主页面-->
<%
'判断处理
If Request("Action")<>"Login" Then
	Call LoginUI()
Else
	Call LoginNow()
End If


If FoundErr=True Then
	CokeShow.AlertErrMsg_general( ErrMsg )
End If


%>
<!--主页面-->
      
      
	</div>
	</div>
    <div class="clear"></div>
  </div>
  <div class="zcdlfooter"></div>
<!--middle end-->
</div>
</div>
</body>
</html>




<%
'登录界面.
Sub LoginUI()
%>
		<div class="mainInfo" style="text-align: center; border: none;">
		<form action="enter.asp" method="post" name="Login" id="Login" target="_parent"
		dojoType="dijit.form.Form"
		execute="processForm('Login')"
		>
			
		  <table width="338" id="listGo" cellpadding="0" cellspacing="0" style="margin: 0 auto; font-size:12px; border:0px #eee solid;">
			<tr> 
			  
			  <td colspan="2" style="border:0px;">
			  
			  <table  border="0" >
				  <tr> 
					<td height="38" colspan="2"><font style="font-size:14px;"><strong>CokeShow System 2010&nbsp;</strong></font> </td>
				  </tr>
				  <tr> 
					<td width="35%" align="right">管理帐号：</td>
					<td width="65%">
					
					<input type="text" id="enterName" name="enterName"
						dojoType="dijit.form.ValidationTextBox"
						required="true"
						promptMessage="欢迎光临"
						invalidMessage="帐号长度必须在6-30之内，例如：wangliang_6179"
						trim="true"
						lowercase="true"
						regExp="[a-zA-Z0-9\_\-\.\@]{4,30}"
						class="input_tell"
						/>
					</td>
				  </tr>
				  <tr> 
					<td align="right">帐号密码：</td>
					<td>
					
					<input type="password" id="enterPassword" name="enterPassword" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="true"
						propercase="false"
						invalidMessage="密码不能为空！"
						trim="true"
						 value=""
                         class="input_tell"
						/>
					</td>
				  </tr>
				  <tr> 
					<td align="right">验 证 码：</td>
					<td>
					
					<input type="text" id="CodeStr" name="CodeStr" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="true"
						propercase="false"
						invalidMessage="请填写4位数字！"
						trim="true"
						regExp="\d{4}"
						maxlength="4"
						 value=""
						  style="width: 80px;"
                          class="input_tell"
						/>
					&nbsp;
					<% =CokeShow.GetCode %></td>
				  </tr>
				  <tr> 
					<td colspan="2" align="center">
						
						<button type="submit" id="submitbtn" 
						  dojoType="dijit.form.Button"
                          class="button"
						  >
						  &nbsp;登录&nbsp;
						  </button>
						
						&nbsp;
						
						<button type="reset" id="resetbtn" 
						  dojoType="dijit.form.Button"
                          class="button"
						  >
						  &nbsp;清除填写&nbsp;
						  </button>
						
						<br /><br />
						<img src="/images/ico/shield.png" /> 痴心不改餐厅
						<!--<img src="/images/chixinbugailogonlogo.jpg" width="300" />-->
						
					</td>
				  </tr>
				</table>
			  
			  </td>
			</tr>
			<tr>
			  <td height="3"></td>
			</tr>
		  </table>
		  
		  	<input type="hidden" name="Action"
			  value="Login"
			  />
		  </form>
		  
		</div>  

<%
End Sub


'登录处理.
Sub LoginNow()
	
	If CokeShow.ChkPost=False Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>记入日志！</li>"
	End If
	
	If Not CokeShow.CodePass Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>验证码错误！</li>"
	End If
	
	'获取登录帐号.
	enterName		=CokeShow.filtPass(Request("enterName"))
	enterPassword	=CokeShow.filtPass(Request("enterPassword"))
	
	'验证
	If enterName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>帐号不能为空！</li>"
	Else
		If CokeShow.strLength(enterName)>30 Or CokeShow.strLength(enterName)<4 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>帐号长度不能大于50个字符，也不能小于10个字符！</li>"
		Else
			enterName=enterName
		End If
	End If
	
	If enterPassword="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>密码不能为空！</li>"
	Else
		If CokeShow.strLength(enterPassword)>18 Or CokeShow.strLength(enterPassword)<6 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>密码长度不能大于16个字符，也不能小于6个字符（至少六位）！</li>"
		Else
			enterPassword=enterPassword
		End If
	End If
	
	
	'拦截错误，不然错误往下进行！
	If FoundErr=True Then Exit Sub
	
	
	enterPassword=md5(enterPassword)
	'检测是否有此帐号.
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT * FROM [CXBG_supervisor] WHERE username='"& enterName &"'"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,1,3
	
	'没有此帐号.
	If RS.Bof And RS.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>没有此帐号！</li>"
	Else
	'存在此帐号时，开始检测密码.
		If enterPassword<>RS("password") Then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>帐号或密码错误！</li>"
		Else
		'通过验证.
			'登录成功赋值.
			Session.Timeout=100
			Session("enterId")			=RS("id")
			Session("enterName")		=RS("username")
			Session("enterPassword")	=RS("password")
			Session("enterCnName")		=RS("cnname")
			Session("enterLevel")		=RS("admin_level")
			
			Session("isHaveWork_supervisor")	=RS("isHaveWork_supervisor")
			
			'更新此次登录信息.
			RS("LastLoginIP")	=Request.ServerVariables("REMOTE_ADDR")
			RS("LastLoginTime")	=Now()
			RS("LoginTimes")	=RS("LoginTimes")+1
			RS.Update
			
'记入日志.
Call CokeShow.AddLog("登录操作：成功登录大后台", sql)
			
			'根据用户级别跳转到相应页面.
			Select Case Session("enterLevel")
				Case 0,1,2
					Response.Redirect "_main.asp"
				Case 3
					Response.Redirect "A_index.asp"
				Case 4
					Response.Redirect "B_index.asp"
				Case Else
					Response.Write "请设置正确的管理人员级别！"
				
			End Select
		End If
		
	End If
	
End Sub

%>
