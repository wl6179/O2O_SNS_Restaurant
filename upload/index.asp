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

<%
'可乐秀.中国CokeShow保护盾2010
'记录访问者IP也有点必要RS("LastLoginIP")	=Request.ServerVariables("REMOTE_ADDR")
If Session("enterLevel")="" Or isNull(Session("enterLevel")) Or isEmpty(Session("enterLevel")) Then
	'阻止执行
	Response.Write "<script type=""text/javascript"">alert('"& "未授权情况下使用编辑器！" &"');</script>"
	Response.End()
	Response.Redirect "/"
Else
	If Not isNumeric(Session("enterLevel")) Then
		'阻止执行
		Response.Write "<script type=""text/javascript"">alert('"& "未授权情况下使用编辑器！" &"');</script>"
		Response.End()
		Response.Redirect "/"
	End If
End If


Public Function CheckPostSafe()
	Dim server_v1, server_v2
	CheckPostSafe = False
	
	server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))	'http://localhost:45233/test.asp
	server_v2 = CStr(Request.ServerVariables("SERVER_NAME"))	'localhost	(www.myhomestay.com.cn)
	If Mid(server_v1, 8, Len(server_v2))=server_v2 Then CheckPostSafe=True	'截取字符数Len(server_v2)；
	CheckPostSafe = True	'''WL;强制避免限制；
End Function


If CheckPostSafe()=False Then
	Response.Write "<br />对不起，为了系统安全，本操作已经记入日志。"
	Response.End
	Response.Redirect "/"
End If

%>

<!--#include file="../system/system_conn.asp"-->
<!--#include file="../system/system_class.asp"-->
<!--GetPost类 -->
<!--#include file="_upload.asp"-->
<%
'类实例化
Dim CokeShow
Set CokeShow = New SystemClass
CokeShow.Start		'调用类的Start方法，初始化类里的ReloadSetup()函数，并得到二维数组Setup
Call CokeShow.SQLWarningSys()	'预警.


Dim CurrentTableName,UnitName		'图片分类表名、分类名称
CurrentTableName 	= "[CXBG_upfiles_class]"
UnitName			= "图片分类"
%>


<%
'接收参数
Dim Action,fileid
Dim controlStr
Action 	= CokeShow.filtRequest(Request("Action"))
fileid 	= CokeShow.filtRequest(Request("fileid"))	'当是修改操作时，才有.

controlStr= CokeShow.filtRequest(Request("controlStr"))
If controlStr="" Or isNull(controlStr) Or isEmpty(controlStr) Then controlStr="null"

%>


<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>上传模板</title>
	
	<link type="text/css" rel="stylesheet" href="<% =filename_dj_MainCss %>" />
	<link type="text/css" rel="stylesheet" href="<% =filename_dj_ThemesCss %>" />
	
	<link type="text/css" rel="stylesheet" href="../style/general_style.css" />
	
	<script type="text/javascript" src="<% =filename_dj %>" djConfig="parseOnLoad:<% =parseOnLoad_dj %>, isDebug:<% =isDebug_dj %>, debugAtAllCosts:<% =isDebug_dj %>"></script>
	<script type="text/javascript" src="<% =filenameWidgetsCompress_dj %>"></script>
	<script type="text/javascript">
		dojo.require("dojo.parser");
//		dojo.require("dijit.form.TextBox");
		dojo.require("dijit.form.ValidationTextBox");
//		dojo.require("dijit.form.CheckBox");
		dojo.require("dijit.form.Button");
		dojo.require("dijit.form.Form");
		
		dojo.require("dojo.io.iframe");
	</script>
	
	

	
</head>
<body class="tundra">


<!--main-->
	<!-- Begin mainleft-->
	
	<!-- End mainleft-->
	
	
	
	<!-- Begin mainright-->	
	<div class="mainContainer" style="font-size: 12px; margin:0px; padding:0px;">
		
		<!--rightInfo-->
		
		<!--rightInfo-->
		
		
		<!--rightInfo-->
		<!--rightInfo-->
		
		
		<!--mainInfo-->
		<!--mainInfo1-->
		
		<!--mainInfo-->
		<!--mainInfo1-->
		<div class="mainInfo" style="font-size: 12px; margin:0px; padding:0px;">
		
			<!--<h2 style="font-size: 12px;">&#187;上传图片</h2>-->
				
			<p style="font-size: 12px;">
			<!--所有上传图片 &#187; 共找到 <font color=red>5</font> 个上传图片-->
				<script type="text/javascript">
					function upload() {
					dojo.io.iframe.send({
						form : "foo",
						handleAs : "html",
						url : "/upload/upfile.asp",
						load : function(response, ioArgs) {
							console.log(response, ioArgs);
							return response;
						},
						error : function(response, ioArgs) {
							console.log("Error");
							console.log(response, ioArgs);
							return response;
						}
						});
						};
						
					
					
				</script>
				
				
				<form action="upfile.asp" method="post" name="foo" id="foo" enctype="multipart/form-data">
					<%
					Dim ParentID
					ParentID=0
					%>
					<!--<select name="classid">
						<%' Call CokeShow.ClassOption_classid(CurrentTableName,"",0,ParentID) %>
						<option value="0"></option>
					</select>--><input type="hidden" name="classid" value="0" />
					<input type="text" name="cnname" value="" style="display:none;" /><!--图片分类
					<br />-->
					<!--enname:--><input type="text" name="enname" value="" style="display:none;" />
					<!--<br />-->
					<!--<br />-->
					<br />
					请选择图片:<br /><input type="file" name="file1" value="" />
					<br />
					<!--file2:--><input type="file" name="file2" style="display:none;" />
					
					<br />
					<!--<input type="submit" name="sbm" value="普通上传" />-->
					
					
					
					
					
					<input type="hidden" name="Action" value="<% =Action %>" />
					<%
					'如果是修改操作.带上id.
					If Action="Modify" Then
					%>
						<input type="hidden" name="fileid" value="<% =fileid %>" />
					<%
					End If
					%>
					
				</form>
				
				
				
				
				<!--<button type="button" onClick="javascript:upload();">上传图片</button>-->
				<button type="button"
				  dojoType="dijit.form.Button"
				  onclick="javascript:Upload(<% If controlStr<>"null" Then Response.Write "'"& controlStr &"'" Else Response.Write(controlStr) %>);"
				  >
				  &nbsp;开始上传图片&nbsp;
				</button>
				
				
				<style type="text/css">
					#uploading {
					color: #FF6600;
					}
				</style>
				<span id="uploading">
					
				</span>
				
			</p>
					
			
		
		</div>
		<!--mainInfo1-->
		<!--mainInfo-->

		<!--mainInfo1-->
		<!--mainInfo-->
		
			
	</div>
	<!-- End mainright-->
<!--main-->


</body>
</html>
