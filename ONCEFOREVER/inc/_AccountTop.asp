<!--#include virtual="/system/system_conn.asp"-->
<!--#include virtual="/system/system_class.asp"-->
<!--#include virtual="/system/md5.asp"-->

<!--#include virtual="/system/foreground_class.asp"-->
<% '!--#include virtual="/CokeCart/Shoppingcart.Class.asp"--%>
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

''购物车类实例化
'Dim CokeCart
'Set CokeCart = New ShoppingcartClass
%>


<%
'变量定义区.
'(用来存储对象的变量，用全大写!)
Dim enterName
Dim enterPassword

Dim RequestMessage,strMessage
RequestMessage=Request("Message")
'如果有传过来的系统通知信息，则记下，并且继续传递.
If RequestMessage<>"" And Len(RequestMessage)>6 Then
	strMessage="&Message="& RequestMessage
Else
	strMessage=""
End If
%>


<%

'验证登录情况.
'response.Write CheckSafePath("0")
'response.End()
'If Not Request("open")="True" Then
	'非常必要安全性，防止直接在地址栏输入参数，要求必须有来路，如：某个站内的链接。WL
	If CheckSafePath("0")=False Then CokeShow.AlertErrMsg_foreground("您的网络处于不安全级别.已写入管理日志"):Response.Redirect "/ONCEFOREVER/LogOn.Welcome":Response.End()
'End If

'Check
Call ChkEnter

%>


<%
'验证过程.
'免费登录算法：CokeShow.PassEncode( Ucase(Md5( Session("YouCanLoginID_Temp") & CokeShow.filtPass(Request("username")) )) )
'1.捆绑大于6字符的Session("YouCanLoginID_Temp")值+帐号值在一起； 2.对其进行MD5加密； 3.对其全大写化； 4.最后对其进行PassEncode加密+去除百分号%。 5.对照一下谁能够传递过来这样的匹配字符串，就允许其登录相应的会员帐号！
Sub ChkEnter()
	'如果有临时通行证ID_Temp+对应的帐号username，那么就可以直接登录相应帐号.
	If Trim(Request("ID_Temp"))=Replace(CokeShow.PassEncode( Ucase(Md5( Session("YouCanLoginID_Temp") & CokeShow.filtPass(Request("username")) )) ),"%","")  AND Len(Session("YouCanLoginID_Temp"))>10 Then
		'检测是否有此帐号.
		Set RS=Server.CreateObject("Adodb.RecordSet")
		sql="SELECT * FROM [CXBG_account] WHERE deleted=0 AND username='"& CokeShow.filtPass(Request("username")) &"'"
		If Not IsObject(CONN) Then link_database
		RS.Open sql,CONN,1,3
		
		'没有此帐号.
		If RS.Bof And RS.Eof Then
			FoundErr=True
			ErrMsg=ErrMsg & "<br /><li>没有此帐号！</li>"
		Else
		'存在此帐号时，开始检测密码.
			'If enterPassword<>RS("password") Then
'				FoundErr=True
'				ErrMsg=ErrMsg & "<br /><li>帐号或密码错误！</li>"
'			Else
			'通过验证.
				'登录成功赋值.
				Session.Timeout=120
				Session("id")			=RS("id")
				Session("username")		=RS("username")
				Session("password")		=RS("password")
				Session("lastloginip")	=RS("lastloginip")
				Session("lastlogintime")=RS("lastlogintime")
				Session("logintimes")	=RS("logintimes")
				Session("account_level")=RS("account_level")
				Session("cnname")		=RS("cnname")
				Session("deleted")		=RS("deleted")
	'			Session("isHaveWork_account")=RS("isHaveWork_account")
	'			Session("myjifen")		=RS("myjifen")
	'			Session("money_WriteIn")		=RS("money_WriteIn")
				
				Session("Birthday")		=RS("Birthday")
				Session("Sex")			=RS("Sex")
				Session("province")		=RS("province")
				Session("city")			=RS("city")
				Session("adddate")		=RS("adddate")
				Session("client_name")			=RS("client_name")
				Session("client_telephone")		=RS("client_telephone")
				Session("client_schooling")		=RS("client_schooling")
				Session("client_memberoffamily")=RS("client_memberoffamily")
				Session("client_befondof")		=RS("client_befondof")
				Session("client_MonthlyIncome")	=RS("client_MonthlyIncome")
				Session("client_work")			=RS("client_work")
				Session("isBindingVIPCardNumber")	=RS("isBindingVIPCardNumber")
				Session("BindingVIPCardNumber")	=RS("BindingVIPCardNumber")
				
				'更新此次登录信息.
				RS("lastloginip")	=Request.ServerVariables("REMOTE_ADDR")
				RS("lastlogintime")	=Now()
				'RS("logintimes")	=RS("logintimes") + 1	'为了不影响会员的登录次数，固此处登录次数不加一.
				
				RS.Update
			'End If
		End If
		
		
	Else
		
		enterName		=CokeShow.filtPass(Session("username"))
		enterPassword	=CokeShow.filtPass(Session("password"))	'此Session中记住的仍然是老密码,如果刚刚更新了新密码，当前的session自然就对不上数据库中刚刚改好的新密码了，所以如果改了新密码得重新登录.此处实时只认新数据库的状态.
		If enterName="" Then
			'Response.Write "session为空！"
			'Response.End()
			Response.Redirect "/ONCEFOREVER/LogOn.Welcome?fromurl="& CokeShow.EncodeURL( CokeShow.GetAllUrlII,"" ) & strMessage '&"&message=请您重新登录！"
			Exit Sub
		End If
		
		sql="SELECT id FROM [CXBG_account] WHERE deleted=0 AND username='"& enterName &"' AND password='"& enterPassword &"'"
'response.Write sql
'response.Write strMessage
'response.End()
		If Not IsObject(CONN) Then link_database
		Set RS=CONN.Execute(sql)
		
		If RS.Bof And RS.Eof Then
			'Response.Write "查不到用户名！"
			'Response.End()
			Response.Redirect "/ONCEFOREVER/LogOn.Welcome?fromurl="& CokeShow.EncodeURL( CokeShow.GetAllUrlII,"" ) & strMessage '&"&message=您的会话结束了，会话数据中查询不到帐号名！"
			Exit Sub
		End If
	
	End If
End Sub

'检测请求网络来源是否为本地网络.
'这是验证大家都在同一个文件夹结构下，例如，都在/general/内链接和跳转，则为安全；万一链接来自根目录，而当前为/general/文件夹，那么会被判断为不安全。WL
'Function CheckSafePath(byVal strMode)
'	Dim strPathFrom,strPathSelf,arrFrom,arrSelf,i
'	CheckSafePath=False
'	If CokeShow.ChkPost=False Then Exit Function
'	
'	strPathFrom = Replace(LCase(CStr(Request.ServerVariables("HTTP_REFERER"))),"http://","")'来源http://localhost:45233/general/supervisor.asp
'    strPathSelf = Replace(LCase(CStr(Request.ServerVariables("URL"))),"http://","")			'当前/general/supervisor.asp
'    'response.Write "strPathFrom:"& strPathFrom &"<br />"
'	'response.Write "strPathSelf:"& strPathSelf &"<br />"
'	If strPathFrom="" Then Exit Function
'    If strPathSelf="" Then Exit Function	
'	
'    arrFrom=Split(strPathFrom,"/")
'    arrSelf=Split(strPathSelf,"/")
'	
'    For i=0 To UBound(arrFrom)
'    	'Response.Write "arrFrom("&i&")="& arrFrom(i) & "<BR/>"
'	Next
'	For i=0 To UBound(arrSelf)
'    	'Response.Write "arrSelf("&i&")="& arrSelf(i) & "<BR/>"
'	Next
'	
'	'Response.Write LCase(Request.ServerVariables("SERVER_NAME"))
'	
'    Select Case strMode
'    	Case "0"
'			For i = 1 To (UBound(arrSelf)-1)
'            	If arrFrom(i)=arrSelf(i) And Instr( arrFrom(0),LCase(Request.ServerVariables("SERVER_NAME")) )>0 Then CheckSafePath=True
'			Next
'	End Select
'	
'End Function

Function CheckSafePath(byVal strMode)
	Dim strPathFrom,strPathSelf,arrFrom,arrSelf,i
	CheckSafePath=False
	If CokeShow.ChkPost=True Then CheckSafePath=True
	
End Function

%>
