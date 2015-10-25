<!--#include file="system_config.asp"-->

<%
'检测系统当前状态.
'Call SystemState

Dim connString,CONN

Sub link_database()
	'SQL数据库连接参数：数据库名、用户密码、用户名、连接名（本地用local，外地用IP）
	Dim sql_databasename,sql_password,sql_username,sql_localname
	
	sql_localname		= "127.0.0.1"
	sql_databasename	= "sq_cxbg"
	sql_username		= "sa"
	sql_password		= "pddas2KKo3000oooa3480988daJLdK3AAAAAAAAAAADDAAKOIERUICMkjd999"
	
	connString = "Provider=Sqloledb; User ID="& sql_username &"; Password="& sql_password &"; Initial Catalog="& sql_databasename &"; Data Source="& sql_localname &";"
	''connString = "Provider=SQLNCLI;Password="& sql_password &";Persist Security Info=True;User ID="& sql_username &";Initial Catalog="& sql_databasename &";Data Source="& sql_localname &""
	'connString = "Provider=SQLNCLI;PWD="& sql_password &";UID="& sql_username &";DATABASE="& sql_databasename &";Data Source="& sql_localname &""
	
	'On Error Resume Next
		Set CONN = Server.CreateObject("Adodb.Connection")
		CONN.Open connString
	
	If Err Then
		Err.Clear	'清除错误……继续脚本的执行.
		Set CONN = Nothing
		Response.Write "数据库连接异常，请重新尝试操作..."
		Response.End
	End If
End Sub
%>
