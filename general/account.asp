<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True	'缓冲数据，才向客户输出；
Response.Charset="utf-8"
Session.CodePage = 65001
%>
<%
'模块说明：后台会员员帐号管理模块.
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
CurrentPageNow 	= "account.asp"			'当前页.
TitleName 		= "注册会员帐号管理"				'此模块管理页的名字.
UnitName 		= "注册会员帐号"					'此模块涉及记录的元素名.
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

currentPage		=CokeShow.filtRequest(Request("Page"))
ExecuteSearch	=CokeShow.filtRequest(Request("ExecuteSearch"))
Keyword			=CokeShow.filtRequest(Request("Keyword"))
TypeSearch		=CokeShow.filtRequest(Request("TypeSearch"))
Action			=CokeShow.filtRequest(Request("Action"))

'处理查询执行 控制变量
If ExecuteSearch="" Then
	ExecuteSearch=0
Else
	If isNumeric(ExecuteSearch) Then ExecuteSearch=CokeShow.CokeClng(ExecuteSearch) Else ExecuteSearch=0
End If
'构造查询字符串
strFileName=CurrentPageNow &"?ExecuteSearch="& ExecuteSearch
If Keyword<>"" Then
	strFileName=strFileName&"&Keyword="& Keyword
End If
If TypeSearch<>"" Then
	strFileName=strFileName&"&TypeSearch="& TypeSearch
End If
'处理当前页码的控制变量，通过获取到的传值获取，默认为第一页1.
If currentPage<>"" Then
    If isNumeric(currentPage) Then currentPage=CokeShow.CokeClng(currentPage) Else currentPage=1
Else
	currentPage=1
End If

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title><% = TitleName %></title>
	
	<link type="text/css" rel="stylesheet" href="<% =filename_dj_MainCss %>" />
	<link type="text/css" rel="stylesheet" href="<% =filename_dj_ThemesCss %>" />
	
	<link type="text/css" rel="stylesheet" href="../style/general_style.css" />
	
	<script type="text/javascript" src="../script/public.js"></script>
	
	<script type="text/javascript" src="<% =filename_dj %>" djConfig="parseOnLoad:<% =parseOnLoad_dj %>, isDebug:<% =isDebug_dj %>, debugAtAllCosts:<% =isDebug_dj %>"></script>
	<script type="text/javascript" src="<% =filenameWidgetsCompress_dj %>"></script>
	<script type="text/javascript">
		dojo.require("dojo.parser");
		dojo.require("dijit.form.ValidationTextBox");
		dojo.require("dijit.form.CurrencyTextBox");
		dojo.require("dijit.form.Button");
		dojo.require("dijit.form.Form");
		dojo.require("dijit.form.NumberTextBox");
		dojo.require("dijit.form.FilteringSelect");
		
		
	</script>
	
	<script type="text/javascript">
		//Table偶数行变色函数.
		dojo.addOnLoad(function() {
			stripeTables("listGo");
		});
	</script>
	
	<script type="text/javascript">
	//确认删除函数，仅针对Table的批量删除操作
	function deleteLot() {
		var conf = confirm("确定要执行删除操作吗？");
		if(conf == true){
			dojo.byId("deleteAction").submit();
		}
	}
	</script>
	
	<script type="text/javascript">
	//dojo全选复选框操作函数
	function checkAll(elementIdName) {
		var checkbox_input_name = "id";		//设置需要控制的选择框的id.
		if (dojo.byId(elementIdName).checked) {
			dojo.forEach(dojo.query("input[name='" + checkbox_input_name + "']"), function(x) {
				x.setAttribute('checked', true);
			});
		}
		else {
			dojo.forEach(dojo.query("input[name='" + checkbox_input_name + "']"), function(x) {
				x.setAttribute('checked', false);
			});
		}
	}
	</script>
	
	<!--#include file="_screen_JS.asp"-->
	

	
</head>
<body class="<% =classname_dj_ThemesCss %>">


<!--main-->
	<!-- Begin mainleft-->
	<div class="newsContainer">
		
		<div class="news1">
			
			<h3>相关管理项目</h3>
				
				<p>
					<!--#include file="menu02_1.asp"-->
				</p>
				
			<h3>当前操作</h3>
				
				<ul>
					<!--<li><a href="http://www.iw3c2.org/">&#187;IW<sup>3</sup>C<sup>2</sup></a></li>-->
					<li><a href="/ONCEFOREVER/AccedeToRegiste.Welcome" target="_blank">&#187;+新增<% =UnitName %></a></li>
					<li><a href="<% = CurrentPageNow %>">&#187;返回列表</a></li>
					<li><a href="#" onClick="deleteLot();">&#187;删除操作</a></li>
				</ul>
				
		</div><!-- End news1-->
		
		<div class="news2">
		
			<h3>查询操作</h3>
			<form action="<% =CurrentPageNow %>" method="GET" name="custForm" id="custForm"
			dojoType="dijit.form.Form"
			>
			<p>
					
					<select name="TypeSearch" id="TypeSearch">
					    
					    <option value="id" selected>按ID查询</option>
					    <option value="cnname" >按中文姓名查询</option>
						<option value="username" >按帐号查询</option>
						
				    </select>
					<br /><br />
					
					
					关键字:
					 <input type="text" id="Keyword" name="Keyword" size="20"
					dojoType="dijit.form.ValidationTextBox"
					required="true"
					propercase="false"
					promptMessage="输入您要查询的关键字..."
					invalidMessage="keyword is required."
					trim="true"
					 style="width: 80px;"
					/>
					<br /><br />
					
					<button type="submit" id="sub" 
					dojoType="dijit.form.Button"
					>
					&nbsp;查询&nbsp;
					</button>
					
					&nbsp;&nbsp;
					
					<button type="button" id="back"
					dojoType="dijit.form.Button"
					onclick="history.back(-1);"
					>&nbsp;返回&nbsp;
					</button>
					<br />
					
					
					<input type="hidden" name="ExecuteSearch" id="ExecuteSearch" value="10" /> 
			</p>
			</form>
		</div><!-- End news2-->
	
	</div>
	<!-- End mainleft-->
	
	
	
	<!-- Begin mainright-->	
	<div class="mainContainer">
		
		<!--rightInfo-->
		<%
		'
		'自动检测显示屏宽度并处理：
		'当屏幕太小小于等于1024、还有新提示都完成时，自动处理消除此最右栏.
		'
		%>
		<!--#include file="_news.asp"-->
		<!--rightInfo-->
		
		
		<!--mainInfo-->
		<!--mainInfo1-->
		<%
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
		
		Else
			Call Main()
		End If
		
		
		If FoundErr=True Then
			CokeShow.AlertErrMsg_general( ErrMsg )
		End If
		%>
		<!--mainInfo1-->
		<!--mainInfo-->
		
			
	</div>
	<!-- End mainright-->
<!--main-->


</body>
</html>
<%
Sub Main()
	
	Select Case ExecuteSearch
		Case 0
			sql="SELECT TOP 500 * FROM "& CurrentTableName &" WHERE deleted=0 ORDER BY id DESC"
			strGuide=strGuide & "所有"& UnitName
		Case 1
			sql="SELECT TOP 500 * FROM "& CurrentTableName &" WHERE deleted=0 ORDER BY logintimes DESC"
			strGuide=strGuide & "登录次数最多的前500个"& UnitName
		
		Case 10
			If Keyword="" Then
				sql="SELECT TOP 500 * FROM "& CurrentTableName &" WHERE deleted=0 ORDER BY id DESC"
				strGuide=strGuide & "所有"& UnitName
			Else
				Select Case TypeSearch
					Case "id"
						If IsNumeric(Keyword)=False Then
								FoundErr=True
							ErrMsg=ErrMsg &"<br><li>"& UnitName &"ID必须是整数！</li>"
						Else
							sql="select * from "& CurrentTableName &" where deleted=0 and id="& CokeShow.CokeClng(Keyword)
							strGuide=strGuide & UnitName &"ID等于<font color=red> " & CokeShow.CokeClng(Keyword) & " </font>的"& UnitName
						End If
					Case "cnname"
						sql="select * from "& CurrentTableName &" where deleted=0 and cnname like '%"& Keyword &"%' order by id desc"
						strGuide=strGuide & "中文姓名中含有“ <font color=red>" & Keyword & "</font> ”的"& UnitName
					Case "username"
						sql="select * from "& CurrentTableName &" where deleted=0 and username like '%"& Keyword &"%' order by id desc"
						strGuide=strGuide & "帐号中含有“ <font color=red>" & Keyword & "</font> ”的"& UnitName
					
				End Select
				
			End If
			
			
		Case Else
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>错误的参数！</li>"
		
	End Select
	
	'拦截错误.
	If FoundErr=True Then Exit Sub
	
	If Not IsObject(CONN) Then link_database
	Set RS=Server.CreateObject("Adodb.RecordSet")
'	Response.Write sql
'	Response.End 
	RS.Open sql,CONN,1,1
	
  	If RS.Eof And RS.Bof Then
		strGuide=strGuide & " &#187; 共找到 <font color=red>0</font> 个"& UnitName
		Call showMain
	Else
    	totalPut=RS.RecordCount		'记录总数.
		strGuide=strGuide & " &#187; 共找到 <font color=red>" & totalPut & "</font> 个"& UnitName
		
		
		'处理页码
		If currentPage<1 Then
       		currentPage=1
    	End If
		'如果传递过来的Page当前页值很大，超过了应有的页数时，进行处理.
    	If (currentPage-1) * maxPerPage > totalPut Then
	   		If (totalPut Mod maxPerPage)=0 Then
	     		'如果整好够页数，赋予当前页最大页.
				currentPage= totalPut \ maxPerPage
		  	Else
		      	'如果不整好，最有一页只有零散几条记录（不丰满的多余页），赋予当前页最大页.（不能整除情况下计算）
				currentPage= totalPut \ maxPerPage + 1
	   		End If

    	End If
	    If currentPage=1 Then
			
        	Call showMain
			
   	 	Else
   	     	'如果传递过来的Page当前页值不大，在应有的页数范围之内时，理应(currentPage-1) * maxPerPage < totalPut，此时进行一些处理.
			if (currentPage-1) * maxPerPage < totalPut then
         	   	'指针指到(currentPage-1)页（前一页）的最后一个记录处.
				RS.Move  (currentPage-1) * maxPerPage
				'RS.BookMark？
         		Dim bookMark
           		bookMark = RS.BookMark
				
            	Call showMain
				
        	else
			'如果传递过来的Page当前页值很大，超过了应有的页数时.打开第一页.
	        	currentPage=1
				
           		Call showMain
				
	    	end if
		End If
	End If
	
End Sub

Sub showMain()
   	Dim i
    i=0
%>
		<!--mainInfo-->
		<!--mainInfo1-->
		<div class="mainInfo" >
		<%
		'
		'自动检测显示屏宽度并处理：
		'1024*768: 	必须消除最右栏.
		'			消除最右栏，  style="width: 750px（auto）; margin-right: 0px;" .
		'1280*960: 	没有style，  不消除最右栏.
		'			消除最右栏，  style="width: 1000px（auto）; margin-right: 0px;".
		'1440*900: 	没有style，  不消除最右栏.
		'			消除最右栏，  style="width: 1120px（auto）; margin-right: 0px;".
		'>1440*900: 没有style，  不消除最右栏.
		'
		%>
			<h2><% =MenuName %> &#187; <% =TitleName %></h2>
				
			<p>
			<%
			Response.Write strGuide
			%>
				<form action="<% = CurrentPageNow %>" method="get" name="deleteAction" id="deleteAction">
					
					<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
					  <thead>
					  <tr>
						<th>
						  <input
						  type="checkbox"
						  name="checkAll1"
						  id="checkAll1"
						  onclick="checkAll('checkAll1');"
						  ></input>
						</th>
						<th>ID</th>
						<th>帐号</th>
						
						<th>最后IP</th>
						<th>最后时间</th>
						<th>登录次数</th>
						
						<th>等级</th>
						<th>中文姓名</th>
						
						<th>访问量</th>
						<th>所在地</th>
                        <th>生日</th>
						
						<th>操作</th>
					  </tr>
					  </thead>
					  <tbody>
					  
					  <%
					  If RS.EOF Then
					  %>
					  <tr>
						<td colspan="10" style="color:red;">对不起，没有记录...</td>
					  </tr>
					  <%
					  End If
					  %>
					  
					  
					  <%
					  Do While Not RS.EOF
					  %>
					  <tr>
						<td height="20">
						  <input
						  type="checkbox"
						  name="id"
						  value="<%=RS("id")%>"
						  ></input>
						</td>
						<td><%=RS("id")%>&nbsp;</td>
						<td><%=RS("username")%>&nbsp;</td>
						
						<td><%=RS("lastloginip")%>&nbsp;</td>
						<td><%=RS("lastlogintime")%>&nbsp;</td>
						<td><%=RS("logintimes")%>&nbsp;</td>
						
						<td><% If RS("account_level")=0 Then Response.Write "<span style=color=red>未审核</span>" Else Response.Write CokeShow.otherField("[CXBG_account_class]",RS("account_level"),"classid","classname",True,0) %>&nbsp;</td>
						<td><%=RS("cnname")%>&nbsp;</td>
						
						
						<td><% =RS("iis") %>&nbsp;</td>
						<td><% =RS("province") %>&nbsp;<% =RS("city") %></td>
						<td><% =RS("Birthday") %>&nbsp;</td>
						
						<td>
						<a href="?Action=Delete&id=<%=RS("id")%>" onClick="return confirm('确定要删除此<% =UnitName %>吗？');">删除</a>
						&nbsp;|&nbsp;
						<a href="?Action=Modify&id=<%=RS("id")%>">修改</a>
						&nbsp;|&nbsp;
						<a href="goto.asp?Action=GoTo&id=<%=RS("id")%>" target="_blank">进入会员后台</a>
						</td>
					  </tr>
					  <%
						  i=i+1
						  If i >= maxPerPage Then Exit Do
						  RS.MoveNext
					  Loop
					  %>
					  
					  
					  
					  </tbody>
					</table>
					
					<input type="hidden" name="Action"
					  value="Delete"
					  />
				</form>
			</p>
					
			<p><a href="#" onClick="deleteLot();">删除操作</a></p>
			
			<p>
			<%
			response.write CokeShow.ShowPage(strFileName,totalPut,maxPerPage,True,True,"个"& UnitName)
			%>
			</p>
		
		</div>
		<!--mainInfo1-->
		<!--mainInfo-->
<%
End Sub


Sub Add()
	strGuide=strGuide & "新增"& UnitName
%>
<script type="text/javascript" 
src="/script/ONCEFOREVER/__accountNameValidation.js" 
></script>
<script type="text/javascript" 
src="/script/ONCEFOREVER/__accountPasswordsValidation.js" 
></script>
<style type="text/css">
	table#listGo thead tr th, table#listGo tbody tr td {
		padding-left: 30px;
		padding-right: 20px;
	}
</style>				
			
			<!--mainInfo-->
			<!--mainInfo1-->
			<div class="mainInfo">
			<%
			'
			'自动检测显示屏宽度并处理：
			'1024*768: 	必须消除最右栏.
			'			消除最右栏，  style="width: 750px（auto）; margin-right: 0px;" .
			'1280*960: 	没有style，  不消除最右栏.
			'			消除最右栏，  style="width: 1000px（auto）; margin-right: 0px;".
			'1440*900: 	没有style，  不消除最右栏.
			'			消除最右栏，  style="width: 1120px（auto）; margin-right: 0px;".
			'>1440*900: 没有style，  不消除最右栏.
			'
			%>
				<h2><% =MenuName %> &#187; <% =TitleName %></h2>
					
				<p>
				<%
				Response.Write strGuide
				%>
				
				
				<form action="<% = CurrentPageNow %>" method="post" name="form1" id="form1"
				dojoType="dijit.form.Form"
				execute="processForm('form1')"
				>
					
					<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
					  <thead>
					  <tr>
					  	
						<th style="text-align: right;">名称</th>
						<th style="text-align: left;">填写数据</th>
						
						<th style="text-align: right;">名称</th>
						<th style="text-align: left;">填写数据</th>
						
					  </tr>
					  </thead>
					  <tbody>
					  
					  
					  <tr>
						<td style="text-align: right;">帐号</td>
						<td style="text-align: left;">
						<input type="text" id="username" name="username"
						dojoType="dijit.form.ValidationTextBox"
						required="true"
						promptMessage="帐号为必填项，格式为英文字母和数字以及-和下划线_."
						invalidMessage="帐号长度必须在6-30之内，例如：wangliang_6179"
						trim="true"
						lowercase="true"
						onChange="accountNameOnChange"
						regExp="[a-zA-Z0-9\_\-\.\@]{6,30}"
						 value=""
						/>
						</td>
						
						<td style="text-align: right;">中文姓名</td>
						<td style="text-align: left;">
						<input type="text" id="cnname" name="cnname" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="true"
						promptMessage=""
						invalidMessage=""
						trim="true"
						 value=""
						 style="width: 250px;"
						/>
						</td>
					  </tr>
					  
					  
					  <tr>
						<td style="text-align: right;">密码</td>
						<td style="text-align: left;">
						<input type="password" id="password" name="password" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="true"
						propercase="false"
						invalidMessage="密码不能为空！"
						trim="true"
						 value=""
						/>
						</td>
						
						<td style="text-align: right;">密码确认</td>
						<td style="text-align: left;">
						<input type="password" id="repassword" name="repassword" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="true"
						propercase="false"
						invalidMessage="确认密码不能为空！"
						trim="true"
						onChange="accountPasswordsOnChange"
						 value=""
						/>
						</td>
					  </tr>
					  
					  
					  <tr>
						<td style="text-align: right;">会员等级</td>
						<td style="text-align: left;">
						
						<select name="account_level" id="account_level"
							dojoType="dijit.form.FilteringSelect"
                            autoComplete="true"
                            forceValidOption="true"
                            queryExpr="*${0}*"
                            class="input_tell"
                            style="width:250px; height:24px;"

						  >
							<option value="0">未审核</option>
							<%
							Call CokeShow.Option_ID("[CXBG_account_class]","",8,0,"classid","classname",True)
							%>
						  </select>
						</td>
						
						<td style="text-align: right;"><!--余额--></td>
						<td style="text-align: left;" style="display:none;">
							<div name="money_WriteIn" id="money_WriteIn"
								dojoType="dijit.form.CurrencyTextBox"
								value=""
								constraints="{ currency:'RMB', places:2 }"
								style="width:100px;"
							  	>
									<script type="dojo/method" event="onChange" args="money_WriteIn">
										
									</script>
								</div>
								RMB
						</td>
					  </tr>
					  
					  <tr style="display:none;">
						<td style="text-align: right;">积分</td>
						<td style="text-align: left;">
							
							<div name="myjifen" id="myjifen"
								dojoType="dijit.form.NumberTextBox"
								value="0"
								constraints="{ pattern:'#,###+' }"
								
							  	>
								</div>
					
								
						</td>
						
						<td style="text-align: right;"></td>
						<td style="text-align: left;">
							
						</td>
					  </tr>
					  
					  
					  <tr>
						<td style="text-align: right;" colspan="4">
						  <input type="hidden" name="Action"
						  value="SaveAdd"
						  />
						  
						  
						      <button type="submit" id="submitbtn" 
							  dojoType="dijit.form.Button"
							  >
							  &nbsp;提交&nbsp;
							  </button>
						  	
							  <button type="button" id="backbtn" 
							  dojoType="dijit.form.Button"
							  onclick="history.back(-1);"
							  >
							  &nbsp;返回&nbsp;
							  </button>
						</td>
					  </tr>
					  
					  
					  </tbody>
					</table>
					
					
				</form>
			
			
			</p>
					
			
			
			<p>
			
			</p>
		
		</div>
		<!--mainInfo1-->
		<!--mainInfo-->
<%
	
End Sub



Sub Modify()
	strGuide=strGuide & "修改"& UnitName
	
	Dim intID
	intID=CokeShow.filtRequest(Request("id"))
	'处理id传值
	If intID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		Exit Sub
	Else
		intID=CokeShow.CokeClng(intID)
	End If
	
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT * FROM "& CurrentTableName &" WHERE deleted=0 AND id="& intID
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,1,3
	
	If RS.Bof And RS.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的"& UnitName &"！</li>"
		Exit Sub
	End If
%>
<script type="text/javascript" 
src="/script/ONCEFOREVER/__accountPasswordsValidation.js" 
></script>
<style type="text/css">
	table#listGo thead tr th, table#listGo tbody tr td {
		padding-left: 30px;
		padding-right: 20px;
	}
</style>


			<!--mainInfo-->
			<!--mainInfo1-->
			<div class="mainInfo" >
			<%
			'
			'自动检测显示屏宽度并处理：
			'1024*768: 	必须消除最右栏.
			'			消除最右栏，  style="width: 750px（auto）; margin-right: 0px;" .
			'1280*960: 	没有style，  不消除最右栏.
			'			消除最右栏，  style="width: 1000px（auto）; margin-right: 0px;".
			'1440*900: 	没有style，  不消除最右栏.
			'			消除最右栏，  style="width: 1120px（auto）; margin-right: 0px;".
			'>1440*900: 没有style，  不消除最右栏.
			'
			%>
				<h2><% =MenuName %> &#187; <% =TitleName %></h2>
					
				<p>
				<%
				Response.Write strGuide
				%>
				
				
				
				<form action="<% = CurrentPageNow %>" method="post" name="form1" id="form1"
				dojoType="dijit.form.Form"
				execute="processForm('form1')"
				>
					
					<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
					  <thead>
					  <tr>
					  	
						<th style="text-align: right;">名称</th>
						<th style="text-align: left;">填写数据</th>
						
						<th style="text-align: right;">名称</th>
						<th style="text-align: left;">填写数据</th>
						
					  </tr>
					  </thead>
					  <tbody>
					  
					  
					  <tr>
						<td style="text-align: right;">会员帐号</td>
						<td style="text-align: left;">
						<% =RS("username") %>
						</td>
						
						<td style="text-align: right;">
						中文姓名
						
						</td>
						<td style="text-align: left;">
						<input type="text" id="cnname" name="cnname" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="true"
						promptMessage="请输入您的中文姓名或者公司的中文姓名."
						invalidMessage="中文姓名出错..."
						trim="true"
						 value="<% =RS("cnname") %>"
						 style="width: 250px;"
						/>
						
						</td>
					  </tr>
					  
					  
					  <tr>
						<td style="text-align: right;">密码</td>
						<td style="text-align: left;">
						<input type="password" id="password" name="password" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						invalidMessage="如果无需修改密码，请留空..."
						trim="true"
						onChange="accountPasswordsOnChange"
						 value=""
						/>
						</td>
						
						<td style="text-align: right;">密码确认</td>
						<td style="text-align: left;">
						<input type="password" id="repassword" name="repassword" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						invalidMessage="如果无需修改密码，请留空..."
						trim="true"
						onChange="accountPasswordsOnChange"
						 value=""
						/>
						</td>
					  </tr>
					  
					  
					  <tr>
						<td style="text-align: right;">会员等级</td>
						<td style="text-align: left;">

						  <select name="account_level" id="account_level"
							dojoType="dijit.form.FilteringSelect"
                            autoComplete="true"
                            forceValidOption="true"
                            queryExpr="*${0}*"
                            class="input_tell"
                            style="width:250px; height:24px;"

						  >
							<option value="0" <% If RS("account_level")=0 Then Response.Write "selected=""selected""" %>>未审核</option>
							<%
							Call CokeShow.Option_ID("[CXBG_account_class]","",8,RS("account_level"),"classid","classname",True)
							%>
						  </select>
						</td>
						
						<td style="text-align: right;"><!--余额--></td>
						<td style="text-align: left;" style="display:none;">
							<div name="money_WriteIn" id="money_WriteIn"
								dojoType="dijit.form.CurrencyTextBox"
								value="<% =RS("money_WriteIn") %>"
								constraints="{ currency:'RMB', places:2 }"
								style="width:100px;"
							  	>
									<script type="dojo/method" event="onChange" args="money_WriteIn">
										
									</script>
								</div>
								RMB
						</td>
					  </tr>
					  
					  <tr style="display:none;">
						<td style="text-align: right;">积分</td>
						<td style="text-align: left;">

						  <div name="myjifen" id="myjifen"
								dojoType="dijit.form.NumberTextBox"
								value="<% =RS("myjifen") %>"
								constraints="{ pattern:'#,###+' }"
								
							  	>
								</div>
								
						</td>
						
						<td style="text-align: right;"></td>
						<td style="text-align: left;">
							
						</td>
					  </tr>
					  
					  
					  <tr>
						<td style="text-align: right;" colspan="4">
						  <input type="hidden" name="Action"
						  value="SaveModify"
						  />
						  <input type="hidden" name="id"
						  value="<% =RS("id") %>"
						  />
						  
						  
						      <button type="submit" id="submitbtn" 
							  dojoType="dijit.form.Button"
							  >
							  &nbsp;提交&nbsp;
							  </button>
						  	
							  <button type="button" id="backbtn" 
							  dojoType="dijit.form.Button"
							  onclick="history.back(-1);"
							  >
							  &nbsp;返回&nbsp;
							  </button>
						</td>
					  </tr>
					  
					  
					  </tbody>
					</table>
					
					
				</form>
				
				
				
				
			</p>
					
			
			
			<p>
			
			</p>
		
		</div>
		<!--mainInfo1-->
		<!--mainInfo-->
				
<%
	'类会自动关闭RS.
End Sub


Sub SaveAdd()
	Dim username,password,repassword,cnname,account_level
	Dim money_WriteIn
	Dim myjifen
	
	'获取其它参数
	username	=CokeShow.filtPass(Request("username"))
	password	=CokeShow.filtPass(Request("password"))
	repassword	=CokeShow.filtPass(Request("repassword"))
	cnname		=CokeShow.filtRequest(Request("cnname"))
	account_level	=CokeShow.filtRequest(Request("account_level"))
	money_WriteIn	=CokeShow.filtRequest(Request("money_WriteIn"))
	
	myjifen			=CokeShow.filtRequest(Request("myjifen"))
	
	'验证
	If username="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>帐号不能为空！</li>"
	Else
		If CokeShow.strLength(username)>50 Or CokeShow.strLength(username)<10 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>帐号长度不能大于50个字符，也不能小于10个字符！</li>"
		Else
			username=username
		End If
	End If
	
	If password="" Or repassword="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>密码与确认密码均不能为空！</li>"
	Else
		If CokeShow.strLength(password)>20 Or CokeShow.strLength(password)<6 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>密码长度不能大于20个字符，也不能小于6个字符（至少六位）！</li>"
		Else
			password=password
		End If
	End If
	If password<>repassword Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>密码和确认密码不一致！</li>"
	End If
	
	If cnname<>"" Then
		If Len(cnname)>20 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>中文姓名只能20位字符之内！此项也可以不填。</li>"
		Else
			cnname=cnname
		End If
	End If
	
	If account_level="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>会员等级不能为空！</li>"
	Else
		If isNumeric(account_level) Then
			account_level=CokeShow.CokeClng(account_level)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>会员等级请用数字填写！</li>"
		End If
	End If
	
	If money_WriteIn="" Then
		money_WriteIn=0
	Else
		If isNumeric(money_WriteIn) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>余额不是数字(非法输入)！</li>"
		End If
	End If
	
	If myjifen="" Then
		myjifen=0
	Else
		If isNumeric(myjifen) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>积分不是数字(非法输入)！</li>"
		End If
	End If
	
'	If MSN<>"" Then
'		If CokeShow.IsValidEmail(MSN)=false Then
'			FoundErr=True
'			ErrMsg=ErrMsg & "<br><li>你的MSN格式不正确！</li>"
'		End If
'	End If
	
	'拦截错误，不然错误往下进行！
	If FoundErr=True Then Exit Sub
	
	
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT * FROM "& CurrentTableName &" WHERE deleted=0"
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	RS.AddNew
		
		RS("username")		=username
		'如果填写了密码，则进行修改.
		If password<>"" Then
			RS("password")	=md5(password)
		End If
		If cnname<>"" Then
			RS("cnname")	=cnname
		End If
		RS("account_level")	=account_level
		
		RS("money_WriteIn")	=money_WriteIn
		RS("myjifen")		=myjifen
		
	
	RS.Update
	RS.Close
	Set RS=Nothing
	
	
	CokeShow.ShowOK "添加"& UnitName &"成功!",CurrentPageNow
End Sub


Sub SaveModify()
	Dim password,repassword,cnname,account_level
	Dim isHaveWork_account
	Dim money_WriteIn
	Dim myjifen
	
	Dim intID
	intID	=CokeShow.filtRequest(Request("id"))
	'检测id参数.
	If intID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		Exit Sub
	Else
		intID=CokeShow.CokeClng(intID)
	End If
	
	'获取其它参数
	password	=CokeShow.filtPass(Request("password"))
	repassword	=CokeShow.filtPass(Request("repassword"))
	cnname		=CokeShow.filtRequest(Request("cnname"))
	account_level	=CokeShow.filtRequest(Request("account_level"))
	
	isHaveWork_account	=CokeShow.filtRequest(Request("isHaveWork_account"))
	
	money_WriteIn	=CokeShow.filtRequest(Request("money_WriteIn"))
	myjifen			=CokeShow.filtRequest(Request("myjifen"))
	
	
	'验证
	If password<>"" Then
		If CokeShow.strLength(Password)>20 Or CokeShow.strLength(Password)<6 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>密码不能大于20小于6！如果你不想修改密码,请保持为空。</li>"
		End If
	End If
	If password<>repassword Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>密码和确认密码不一致！</li>"
	End If
	
	If cnname<>"" Then
		If Len(cnname)>20 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>中文姓名只能20位字符之内！此项也可以不填。</li>"
		Else
			cnname=cnname
		End If
	End If
	
	If account_level="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>会员等级不能为空！</li>"
	Else
		If isNumeric(account_level) Then
			account_level=CokeShow.CokeClng(account_level)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>会员等级请用数字填写！</li>"
		End If
	End If
	
	If money_WriteIn="" Then
		money_WriteIn=0
	Else
		If isNumeric(money_WriteIn) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>余额不是数字(非法输入)！</li>"
		End If
	End If
	
	If myjifen="" Then
		myjifen=0
	Else
		If isNumeric(myjifen) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>积分不是数字(非法输入)！</li>"
		End If
	End If
	
'	If MSN<>"" Then
'		If CokeShow.IsValidEmail(MSN)=false Then
'			FoundErr=True
'			ErrMsg=ErrMsg & "<br><li>你的MSN格式不正确！</li>"
'		End If
'	End If
	
	'拦截错误，不然错误往下进行！
	If FoundErr=True Then Exit Sub
	
	
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT * FROM "& CurrentTableName &" WHERE deleted=0 AND id="& intID
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,1,3
	
	'拦截此记录的异常情况.
	If RS.Bof And RS.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的"& UnitName &"！</li>"
		Exit Sub
	End If
	
		'如果填写了密码，则进行修改.
		If password<>"" Then
			RS("password")	=md5(password)
		End If
		If cnname<>"" Then
			RS("cnname")	=cnname
		End If
		RS("account_level")	=account_level
		
		RS("money_WriteIn")	=money_WriteIn
		RS("myjifen")	=myjifen
	
	RS.Update
	RS.Close
	Set RS=Nothing
	
'记入日志.
Call CokeShow.AddLog("编辑操作：成功编辑了ID为"& intID &"的"& UnitName &"-"& cnname, sql)
	
	CokeShow.ShowOK "修改"& UnitName &"成功!",CurrentPageNow
End Sub


Sub Delete()
	Dim strID,i
	strID=CokeShow.filtRequest(Request("id"))
	If strID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要删除的"& UnitName &"</li>"
		Exit Sub
	End If
	If Instr(strID,",")>0 Then
		strID=Split(strID,",")
		For i=0 To Ubound(strID)
			DeleteOne(strID(i))
		Next
	Else
		DeleteOne(strID)
	End If
	
	'拦截错误，不然错误往下进行！
	If FoundErr=True Then Exit Sub
	
'记入日志.
Call CokeShow.AddLog("删除操作：成功删除了ID为"& CokeShow.filtRequest(Request("id")) &"的"& UnitName, "")
	
	CokeShow.ShowOK "删除操作成功!",CurrentPageNow
End Sub

Sub DeleteOne(strID)
	strID=CokeShow.CokeClng(strID)
	If Not IsObject(CONN) Then link_database
	Set RS=CONN.Execute("SELECT * FROM "& CurrentTableName &" WHERE deleted=0 AND id="& strID)
	
	If Not RS.Eof Then
		
		'//CokeShow.Execute("DELETE FROM "& CurrentTableName &" WHERE id="& strID)
		'//CokeShow.Execute("UPDATE "& CurrentTableName &" SET deleted=1 WHERE username='"& username &"'")
		CONN.Execute("UPDATE "& CurrentTableName &" SET deleted=1 WHERE id="& strID)
		
	Else
		'找不着记录，则
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>记录id为"& strID &"的"& UnitName &"删除操作没成功，此记录有可能早已丢失！</li>"
		Exit Sub
		
	End If
	
End Sub


'设置会员是否通过我站点审核...通过便可以在站点显示，成为公开会员用户.
'示例函数.
Sub change_isValidate()
	Dim strState		'本过程的状态，字符串.
	strState = "审核"
	
	Dim strID
	Dim isValidate
	
	strID=CokeShow.filtRequest(Request("id"))
	isValidate= CokeShow.CokeClng(CokeShow.filtRequest(Request("isValidate")))
	
	If strID="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>参数不足！</li>"
		Exit Sub
	Else
		strID=CokeShow.CokeClng(strID)
	End If
	

	If FoundErr=True Then
		Exit Sub
	End If

	If isValidate= 1 Then
		CONN.Execute("UPDATE "& CurrentTableName &" SET isValidate=0 WHERE id="& id )
	ElseIf isValidate= 0 Then
		CONN.Execute("UPDATE "& CurrentTableName &" SET isValidate=1 WHERE id="& id )
	End If

	CokeShow.ShowOK "更新"& UnitName & "的"& strState &"状态成功！",CurrentPageNow
	
End Sub


%>