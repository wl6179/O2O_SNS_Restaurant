<!--#include file="../../system/system_conn.asp"-->
<!--#include file="../../system/system_class.asp"-->
<!--#include file="../../system/md5.asp"-->
<%
'类实例化
Dim CokeShow
Set CokeShow = New SystemClass
CokeShow.Start		'调用类的Start方法，初始化类里的ReloadSetup()函数，并得到二维数组Setup
Call CokeShow.SQLWarningSys()	'预警.
%>


<%
'变量定义区.
'(用来存储对象的变量，用全大写!)
Dim enterName
Dim enterPassword

%>


<%
'验证登录情况.
Call ChkEnter
If Not Request("open")="True" Then
	If CheckSafePath(0)=False Then CokeShow.AlertErrMsg_general("您的网络处于不安全级别.已写入管理日志"):Response.End():Response.Redirect "enter.asp"
End If


%>


<%
'验证过程.
Sub ChkEnter()
	
	enterName		=CokeShow.filtPass(Session("enterName"))
	enterPassword	=CokeShow.filtPass(Session("enterPassword"))
	If enterName="" Then
		'Response.Write "session为空！"
		'Response.End()
		Response.Redirect "enter.asp"
		Exit Sub
	End If
	
	sql="SELECT id FROM [CXBG_supervisor] WHERE username='"& enterName &"' AND password='"& enterPassword &"'"
	If Not IsObject(CONN) Then link_database
	Set RS=CONN.Execute(sql)
	
	If RS.Bof And RS.Eof Then
		'Response.Write "查不到用户名！"
		'Response.End()
		Response.Redirect "enter.asp"
		Exit Sub
	End If
	
End Sub

'检测请求网络来源是否为本地网络.
Function CheckSafePath(byVal strMode)
	Dim strPathFrom,strPathSelf,arrFrom,arrSelf,i
	CheckSafePath=False
	If CokeShow.ChkPost=False Then Exit Function
	
	strPathFrom = Replace(LCase(CStr(Request.ServerVariables("HTTP_REFERER"))),"http://","")'来源http://localhost:45233/general/supervisor.asp
    strPathSelf = Replace(LCase(CStr(Request.ServerVariables("URL"))),"http://","")			'当前/general/supervisor.asp
    
	If strPathFrom="" Then Exit Function
    If strPathSelf="" Then Exit Function
	
    arrFrom=Split(strPathFrom,"/")
    arrSelf=Split(strPathSelf,"/")
	
    For i=0 To UBound(arrFrom)
    	'Response.Write "arrFrom("&i&")="& arrFrom(i) & "<BR/>"
	Next
	For i=0 To UBound(arrSelf)
    	'Response.Write "arrSelf("&i&")="& arrSelf(i) & "<BR/>"
	Next
	
	'Response.Write LCase(Request.ServerVariables("SERVER_NAME"))
	
    Select Case strMode
    	Case "0"
			For i = 1 To (UBound(arrSelf)-1)
            	If arrFrom(i)=arrSelf(i) And Instr( arrFrom(0),LCase(Request.ServerVariables("SERVER_NAME")) )>0 Then CheckSafePath=True
			Next
	End Select
	
End Function

%>
