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
'Ajax远程调用asp文件.检测管理人帐号是否有重复的处理模块.
'
%>﻿
<!--#include file="../system/system_conn.asp"-->
<!--#include file="../system/system_class.asp"-->
<%
'类实例化
Dim CokeShow
Set CokeShow = New SystemClass
CokeShow.Start		'调用类的Start方法，初始化类里的ReloadSetup()函数，并得到二维数组Setup
Call CokeShow.SQLWarningSys()	'预警.【上边一行原本是被注销的，WL为未来可能发生的问题提示！】
%>


<%
Dim userName
userName=CokeShow.filtRequest(Request("username"))

'检测帐号.
Dim sql,RS
Set RS=Server.CreateObject("Adodb.RecordSet")
sql="SELECT * FROM [CXBG_supervisor] WHERE username='"& userName &"'"
If Not IsObject(CONN) Then link_database
RS.Open sql,CONN

If RS.Bof And RS.Eof Then	'用户名合法
	Response.Write "{valid: true}"
Else						'用户名不合法
	Response.Write "{valid: false}"
End If

RS.Close
Set RS=Nothing
%>