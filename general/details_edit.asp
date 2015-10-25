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

<!--#include file="../fckeditor.asp" -->

<%
'变量定义区.
'(用来存储对象的变量，用全大写!)
Const maxPerPage=15							'当前模块分页设置.
Dim CurrentPageNow,TitleName,UnitName
CurrentPageNow 	= "details_edit.asp"			'当前页.
TitleName 		= "内容列表管理"				'此模块管理页的名字.
UnitName 		= "内容"					'此模块涉及记录的元素名.
'自定义设置.
'本地设置.
Dim CurrentTableName
CurrentTableName 	= "[CXBG_details]"		'此模块涉及的[表]名.
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
		dojo.require("dijit.form.NumberSpinner");
		dojo.require("dijit.form.Button");
		dojo.require("dijit.form.Form");
		//dojo.require("dijit.form.CurrencyTextBox");
		//dojo.require("dijit.form.NumberTextBox");
		dojo.require("dijit.form.FilteringSelect");
		//dojo.require("dijit.form.DateTextBox");
		
		//dojo.require("dijit.layout.ContentPane");
		//dojo.require("dijit.layout.StackContainer");
		
		dojo.require("dijit.Dialog");
		//上传图片.
		dojo.require("dojo.io.iframe");
	</script>
	
	<script type="text/javascript">
		//Table偶数行变色函数.
		dojo.addOnLoad(function() {
			//初始化列表.
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
	
	
	
	
	
	<script type="text/javascript" src ="../fckeditor.js"></script> 
	
	<script language="javascript" type="text/javascript">
	  dojo.addOnLoad(function() {
	      if (dojo.byId("logtext")) {
			  var  oFCKeditor = new  FCKeditor("logtext") ;
				   oFCKeditor.BasePath  = "/";
				   oFCKeditor.Width = "723";
				   oFCKeditor.Height = "380";
				   oFCKeditor.ToolbarSet = "Default";
				   oFCKeditor.Config["SkinPath"] = "skins/default/";
				   
				   //补充设置
					oFCKeditor.Config["ImageUpload"]	="true";
					oFCKeditor.Config["ImageBrowser"]	="true";
					oFCKeditor.Config["LinkUpload"]		="true";
					oFCKeditor.Config["LinkBrowser"]	="true";
					oFCKeditor.Config["FlashUpload"]	="true";
					oFCKeditor.Config["FlashBrowser"]	="true";
				   
				   oFCKeditor.ReplaceTextarea();
		   }
	  });
	</script>
	
</head>
<body class="<% =classname_dj_ThemesCss %>">


<!--main-->
	<!-- Begin mainleft-->
	<div class="newsContainer">
		
		<div class="news1">
			
			<h3>相关管理项目</h3>
				
				<p>
					<!--#include file="menu_details.asp"-->
				</p>
				
			<h3>当前操作</h3>
				
				<ul>
					<!--<li><a href="http://www.iw3c2.org/">&#187;IW<sup>3</sup>C<sup>2</sup></a></li>-->
					<li><a href="?Action=Add">&#187;+新增<% =UnitName %></a></li>
					<li><a href="<% = CurrentPageNow %>">&#187;返回列表</a></li>
					<li><a href="#" onClick="deleteLot();">&#187;删除操作</a></li>
				</ul>
				
		</div><!-- End news1-->
		<div class="news2">
        	<h3>快速筛选</h3>
            <form name="form1" action="<% =CurrentPageNow %>" method="get">
                <tr class="tdbg"> 
                    <td width="100" height="30"></td>
                    <td width="687" height="30">
                        <select size="1" name="details_class_id" onChange="javascript:submit()">
                            <option value=>请选择筛选条件</option>
                            <%
                            Call CokeShow.Option_ID("[CXBG_details_class]","",888,0,"classid","classname",True)
                            %>
                            
                        </select>
                    </td>
                </tr>
                <input type="hidden" name="ExecuteSearch" value="1" />
            </form>
        </div>
		<div class="news2">
		
			<h3>查询操作</h3>
			<form action="<% =CurrentPageNow %>" method="GET" name="custForm" id="custForm"
			dojoType="dijit.form.Form"
			>
			<p>
					
					<select name="TypeSearch" id="TypeSearch">
					    
					    <option value="id" selected>按ID查询</option>
					    <option value="topic" >按标题查询</option>
						<option value="logtext" >按内容查询</option>
						
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

<%' ="<br />UnitPrice_Sales:"& request("UnitPrice_Sales") %>
<%' ="<br />isSales_StartDate:"& request("isSales_StartDate") %>
<%' ="<br />isSales_StopDate:"& request("isSales_StopDate") %>

</body>
</html>
<%
Sub Main()
	
	Select Case ExecuteSearch
		Case 0
			'sql="SELECT TOP 500 * FROM "& CurrentTableName &" WHERE deleted=0 ORDER BY details_orderid DESC,id DESC"
			sql="SELECT TOP 500 * FROM "& CurrentTableName &" WHERE deleted=0 ORDER BY id DESC"
			strGuide=strGuide & "所有"& UnitName
		Case 1
			sql="SELECT TOP 500 * FROM "& CurrentTableName &" WHERE deleted=0 ORDER BY id DESC"
			If isNumeric(Request("details_class_id")) Then
				If Clng(Request("details_class_id"))>0 Then
					sql="SELECT TOP 500 * FROM "& CurrentTableName &" WHERE deleted=0 AND details_class_id="& Clng(Request("details_class_id")) &" ORDER BY id DESC"
				End If
			End If
			strGuide=strGuide & "快速筛选到的"& UnitName
		
		Case 10
			If Keyword="" Then
				sql="SELECT TOP 500 * FROM "& CurrentTableName &" WHERE deleted=0 ORDER BY details_orderid DESC,id DESC"
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
					Case "topic"
						sql="select * from "& CurrentTableName &" where deleted=0 and topic like '%"& Keyword &"%' ORDER BY details_orderid DESC,id DESC"
						strGuide=strGuide & "菜品名称中含有“ <font color=red>" & Keyword & "</font> ”的"& UnitName
					Case "logtext"
						sql="select * from "& CurrentTableName &" where deleted=0 and logtext like '%"& Keyword &"%' ORDER BY details_orderid DESC,id DESC"
						strGuide=strGuide & "菜品编号中含有“ <font color=red>" & Keyword & "</font> ”的"& UnitName
					
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
						<th>新闻图片</th>
						<th>标题</th>
						
						<th>所属分类</th>
						<th>作者</th>
						<th>访问量</th>
						<th>是否推荐</th>
						<th>添加日期</th>
						<th>排序优先</th>
						<th>是否发布</th>
						<th>操作</th>
					  </tr>
					  </thead>
					  <tbody>
					  
					  <%
					  If RS.EOF Then
					  %>
					  <tr>
						<td colspan="11" style="color:red;">对不起，没有记录...</td>
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
						<td style="text-align:center;">
							<%
							If RS("photo")<>"" And Len(RS("photo"))>3 Then
							%>
								
								<img id="src_photo<% =RS("id") %>" src="<% =RS("photo") %>" width="50" />
								<!--begin-->
									<span dojoType="dijit.Tooltip"
									connectId="src_photo<% =RS("id") %>"
									id="tmp<% =RS("id") %>"
									style="display:none;"
									>
										图片详细尺寸：<br /><img src="<% =RS("photo") %>" />
									</span>
								<!--end-->
							<%
							End If
							%>
							&nbsp;
						</td>
						
						<td style="font-weight:bold;"><%=RS("topic")%>&nbsp;</td>
						
						<td><% =CokeShow.otherField("[CXBG_details_class]",RS("details_class_id"),"classid","classname",True,0) %>&nbsp;</td>
						<td><%=RS("author")%><!--&nbsp;(<%=RS("authorid")%>)--></td>
						<td><%=RS("iis")%>&nbsp;</td>
						<td>
						<% If RS("isRecommend")=1 Then Response.Write "<img src=/images/yes.gif />推荐" Else Response.Write "<img src=/images/no_1.png />尚未推荐" %>
						&nbsp;
						</td>
						<td><%=RS("adddate")%>&nbsp;</td>
						<td><%=RS("details_orderid")%>&nbsp;</td>
						<td>
						<% If RS("isOnpublic")=1 Then Response.Write "<img src=/images/yes.gif />发布" Else Response.Write "<img src=/images/no.gif />不发布" %>
						&nbsp;
						</td>
						<td>
						<a href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( RS("id") ) %>" target="_blank">浏览</a>
						&nbsp;|&nbsp;
						
						<a href="?Action=Modify&id=<%=RS("id")%>">修改</a>
						&nbsp;|&nbsp;
						<a href="?Action=Delete&id=<%=RS("id")%>" onClick="return confirm('确定要删除此<% =UnitName %>吗？');">删除</a>
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
src="../script/__supervisorNameValidation.js" 
></script>
<script type="text/javascript" 
src="../script/__supervisorPasswordsValidation.js" 
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
				><!--dojoType="dijit.form.Form" 为了标签dijit，暂时去掉了这里。-->
				
				
				
				
				<!--
				容器Begin
				-->
				
					
						<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
						  <thead>
						  <tr>
							
							<th style="text-align: right;">名称</th>
							<th style="text-align: left;" colspan="3">填写数据</th>
							
							
						  </tr>
						  </thead>
						  <tbody>
						  
						  
						  <tr>
							<td style="text-align: right;">标题</td>
							<td style="text-align: left;" colspan="3">
							<input type="text" id="topic" name="topic"
							dojoType="dijit.form.ValidationTextBox"
							required="true"
							propercase="true"
							promptMessage="标题为必填项."
							invalidMessage="标题长度必须在1-50之内."
							trim="true"
							lowercase="false"
							regExp=".{0,50}"
							 value=""
							 style="width: 400px;"
                             class="input_tell"
							/>
							</td>
							
							
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">所属分类</td>
							<td style="text-align: left;" colspan="3">
							<select name="details_class_id" id="details_class_id"
								dojoType="dijit.form.FilteringSelect"
								autoComplete="true"
                                forceValidOption="true"
                                queryExpr="*${0}*"
                                class="input_tell"
                                style="width:250px; height:24px;"
	
							  >
								<%
								'Call CokeShow.Option_ID("[CXBG_details_class]","",888,0,"classid","classname",True)
								%>
                                <% Call CokeShow.ClassOption_classid("[CXBG_details_class]","",0,0) %>
							  </select>
							 <span style="color:#999999;">必须选择.</span>
							</td>
							
						  </tr>
						  
						  
						  
						  <tr>
							<td style="text-align: right;">是否推荐</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="checkbox" name="isRecommend" id="isRecommend" value="1" /><label for="isRecommend" style="color:#999999;"> 打勾表示为推荐内容，将显示在右侧推荐信息处.</label>
							</td>
							
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">排序优先</td>
							<td style="text-align: left;" colspan="3">
							
							<div name="details_orderid" id="details_orderid"
							dojoType="dijit.form.NumberSpinner"
							constraints="{ max:999, min:0 }"
							style="width:8em;"
							value="0"
							onClick="this.select();"
							>
							</div>
							<br />
							<span style="color:#999999;">数字越大，排序越优先.建议以10递增填写，以便于长期的微调.<br />(PageUp键和PageDown键可以以10递增/减操作.)</span>
							</td>
							
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">发布</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="checkbox" name="isOnpublic" id="isOnpublic" value="1" checked="checked" /><label for="isOnpublic" style="color:#999999;"> 打勾表示<% =UnitName %>将正常显示在网站上，否则不会在网站显示.</label>
							</td>
							
						  </tr>
						  
						  <tr>
							
							<td style="text-align: left;" colspan="4">
								正文:
								<br />
								<br />
								<textarea name="logtext" id="logtext" style="display:none;"></textarea>
								
							</td>
							
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;" valign="top">
								新闻图片：
								
							</td>
							
							<td style="text-align: left;" colspan="3">
								<span style="color:#999999;"> 尺寸要求:宽200左右*高不限 (长方形)</span>
								<br />
								<br />
								
								<a href="#" target="_blank" id="href_photo">
								<img style="border:5px #999999 solid;" src="/images/NoPic.png" onerror='this.src="/images/NoPic.png"'
								 id="src_photo"
								/></a>
								
								<button type="button" 
								dojoType="dijit.form.Button"
								onclick="ShowDialog('<img src=/images/up.gif />上传图片','../upload/index.asp?Action=Add','width:300px;height:200px;');"
								>
								&nbsp;点击开始上传&nbsp;
								</button>
								
								<div id="div_photo">
								<input type="text" id="value_photo" name="photo"
								dojoType="dijit.form.ValidationTextBox"
								required="false"
								promptMessage="请上传图片."
								invalidMessage="必须在250长度之内！例如：/uploadimages/cokeshow.com.cn20097131746850193.png"
								trim="true"
								lowercase="false"
								
								regExp=".{0,250}"
								 value=""
								 style="width: 500px;"
                                 class="input_tell"
								/>
								
								</div>
							</td>
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">图片旁显示的摘要简介</td>
							<td style="text-align: left;" colspan="3">
							<input type="text" id="photo_desc" name="photo_desc"
							dojoType="dijit.form.ValidationTextBox"
							required="false"
							propercase="true"
							
							invalidMessage="图片旁显示的摘要简介长度必须在1-50之内."
							trim="true"
							lowercase="false"
							regExp=".{0,50}"
							 value=""
							 style="width: 400px;"
							/> <span style="color:#999999;">可以留空不填.</span>
							</td>
							
							
						  </tr>
						  
						  
						  
						  <tr>
							<td style="text-align: right;" colspan="4">
							  <input type="hidden" name="Action"
							  value="SaveAdd"
							  />
							  
							  
								  <button type="submit" id="submitbtn" 
								  dojoType="dijit.form.Button"
                                  iconClass="dijitEditorIcon dijitEditorIconSave"
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
					
				<!--
				容器End
				-->
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
src="../script/__supervisorPasswordsValidation.js" 
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
				
				
				
				<!--dojoType="dijit.form.Form" 为了标签dijit，暂时去掉了这里。-->
				<form action="<% = CurrentPageNow %>" method="post" name="form1" id="form1"
				dojoType="dijit.form.Form"
				execute="processForm('form1')"
				>
				
				
				<!--
				容器Begin
				-->
				
				
					
						<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
						  <thead>
						  <tr>
							
							<th style="text-align: right;">名称</th>
							<th style="text-align: left;" colspan="3">填写数据</th>
							
							
						  </tr>
						  </thead>
						  <tbody>
						  
						  
						  <tr>
							<td style="text-align: right;">标题</td>
							<td style="text-align: left;" colspan="3">
							<input type="text" id="topic" name="topic"
							dojoType="dijit.form.ValidationTextBox"
							required="true"
							propercase="true"
							promptMessage="标题为必填项."
							invalidMessage="标题长度必须在1-50之内."
							trim="true"
							lowercase="false"
							regExp=".{0,50}"
							 value="<% =RS("topic") %>"
							 style="width: 400px;"
                             class="input_tell"
							/>
							</td>
							
							
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">所属分类</td>
							<td style="text-align: left;" colspan="3">
							
							<select name="details_class_id" id="details_class_id"
								dojoType="dijit.form.FilteringSelect"
								autoComplete="true"
                                forceValidOption="true"
                                queryExpr="*${0}*"
                                class="input_tell"
                                style="width:250px; height:24px;"
	
							  >
								<%
								'Call CokeShow.Option_ID("[CXBG_details_class]","",888,RS("details_class_id"),"classid","classname",True)
								%>
                                <% Call CokeShow.ClassOption_classid("[CXBG_details_class]","",0,RS("details_class_id")) %>
							  </select>
							 <span style="color:#999999;">必须选择.</span>
							</td>
							
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">是否推荐</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="checkbox" name="isRecommend" id="isRecommend" value="1" <% If RS("isRecommend")=1 Then Response.Write "checked=""checked""" %> /><label for="isRecommend" style="color:#999999;"> 打勾表示为推荐内容，将显示在右侧推荐信息处.</label>
							</td>
							
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">排序优先</td>
							<td style="text-align: left;" colspan="3">
							
							<div name="details_orderid" id="details_orderid"
							dojoType="dijit.form.NumberSpinner"
							constraints="{ max:999, min:0 }"
							style="width:8em;"
							value="<% =RS("details_orderid") %>"
							onClick="this.select();"
							>
							</div>
							<br />
							<span style="color:#999999;">数字越大，排序越优先.建议以10递增填写，以便于长期的微调.<br />(PageUp键和PageDown键可以以10递增/减操作.)</span>
							
							
							<!--<label for="isSales_StartDate">起始促销日期</label>
									<div id="isSales_StartDate" name="isSales_StartDate"
									dojoType="dijit.form.DateTextBox"
									required="false"
									constraints="{min:'<% =DateAdd("d",-1, CokeShow.filt_DateStr(Date())) %>', max:'2011-01', datePattern:'yyyy-MM-dd'}"
									promptMessage="请选择起始促销日期."
									invalidMessage="Invalid Service Date."
									style="width:100px;"
									>
									</div>-->
							</td>
							
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">发布</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="checkbox" name="isOnpublic" id="isOnpublic" value="1" <% If RS("isOnpublic")=1 Then Response.Write "checked=""checked""" %> /><label for="isOnpublic" style="color:#999999;"> 打勾表示<% =UnitName %>将正常显示在网站上，否则不会在网站显示.</label>
							</td>
							
						  </tr>
						  
						  
						  <tr>
							
							<td style="text-align: left;" colspan="4">
								正文:
								<br />
								<br />
								<textarea name="logtext" id="logtext" style="display:none;"><% =RS("logtext") %></textarea>
								
							</td>
							
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;" valign="top">
								新闻图片：
								
							</td>
							
							<td style="text-align: left;" colspan="3">
								<span style="color:#999999;"> 尺寸要求:宽200左右，高度不限 (长方形)</span>
								<br />
								<br />
								
								<a href="<% =RS("photo") %>" target="_blank" id="href_photo">
								<img style="border:5px #999999 solid;" src="<% If RS("photo")<>"" Then Response.Write RS("photo") Else Response.Write "/images/NoPic.png" %>" onerror='this.src="/images/NoPic.png"'
								 id="src_photo"
								/></a>
								
								<button type="button" 
								dojoType="dijit.form.Button"
								onclick="ShowDialog('<img src=/images/up.gif />修改图片','../upload/index.asp?Action=Add','width:300px;height:200px;');"
								>
								&nbsp;点击开始上传&nbsp;
								</button>
								
								<div id="div_photo">
								<input type="text" id="value_photo" name="photo"
								dojoType="dijit.form.ValidationTextBox"
								required="false"
								promptMessage="请上传图片."
								invalidMessage="必须在250长度之内！例如：/uploadimages/cokeshow.com.cn20097131746850193.png"
								trim="true"
								lowercase="false"
								
								regExp=".{0,250}"
								 value="<% =RS("photo") %>"
								 style="width: 500px;"
                                 class="input_tell"
								/>
								
								</div>
							</td>
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">图片旁显示的摘要简介</td>
							<td style="text-align: left;" colspan="3">
							<input type="text" id="photo_desc" name="photo_desc"
							dojoType="dijit.form.ValidationTextBox"
							required="false"
							propercase="true"
							
							invalidMessage="图片旁显示的摘要简介长度必须在1-50之内."
							trim="true"
							lowercase="false"
							regExp=".{0,50}"
							 value="<% =RS("photo_desc") %>"
							 style="width: 400px;"
							/> <span style="color:#999999;">可以留空不填.</span>
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
                                  iconClass="dijitEditorIcon dijitEditorIconSave"
								  >
								  &nbsp;提交保存&nbsp;
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
					
				
				
				<!--
				容器End
				-->
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
	Dim topic,logtext,details_class_id,photo,isOnpublic,details_orderid
	Dim photo_desc
	Dim isRecommend
		
	'获取其它参数
	topic				=CokeShow.filtRequest(Request("topic"))
	logtext				=CokeShow.filtRequestRich(Request("logtext"))
	details_class_id	=CokeShow.filtRequest(Request("details_class_id"))
	photo				=CokeShow.filtRequest(Request("photo"))
	isOnpublic			=CokeShow.filtRequest(Request("isOnpublic"))
	details_orderid		=CokeShow.filtRequest(Request("details_orderid"))
	
	photo_desc			=CokeShow.filtRequest(Request("photo_desc"))
	
	isRecommend			=CokeShow.filtRequest(Request("isRecommend"))
	
	'验证
	If topic="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>标题不能为空！</li>"
	Else
		If CokeShow.strLength(topic)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>标题长度不能大于50个字符！</li>"
		Else
			topic=topic
		End If
	End If
		
	If logtext="" Or isNull(logtext) Or isEmpty(logtext) Then
'		FoundErr=True
'		ErrMsg=ErrMsg &"<br><li>正文不能为空！</li>"
		logtext=""
	Else
		If CokeShow.strLength(logtext)>8000 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>正文长度不能大于8000个字符！</li>"
		Else
			logtext=logtext
		End If
	End If
	
	If details_class_id="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请选择所属分类！</li>"
	Else
		If isNumeric(details_class_id) Then
			details_class_id=CokeShow.CokeClng(details_class_id)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前所属分类参数不是数字(出现异常)！</li>"
		End If
	End If
	
	If photo="" Or isNull(photo) Or isEmpty(photo) Then
		photo=""
		
	Else
		If Len(photo)>10 Then
			photo=photo
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>上传图片获得的图片地址出错(出现异常)！</li>"
		End If
	End If
	
	If isOnpublic="" Or isNull(isOnpublic) Or isEmpty(isOnpublic) Then
		isOnpublic=0
		
	Else
		If isNumeric(isOnpublic) Then
			isOnpublic=CokeShow.CokeClng(isOnpublic)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前发布状态出错(出现异常)！</li>"
		End If
	End If
	
'	If MSN<>"" Then
'		If CokeShow.IsValidEmail(MSN)=false Then
'			FoundErr=True
'			ErrMsg=ErrMsg & "<br><li>你的MSN格式不正确！</li>"
'		End If
'	End If
	
	If photo_desc="" Or isNull(photo_desc) Or isEmpty(photo_desc) Then
		photo_desc=""
		
	Else
		If Len(photo_desc)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>图片文字描述字符长度不得超过50！</li>"
		Else
			photo_desc=photo_desc
		End If
	End If
	
	If isRecommend="" Or isNull(isRecommend) Or isEmpty(isRecommend) Then
		isRecommend=0
		
	Else
		If isNumeric(isRecommend) Then
			isRecommend=CokeShow.CokeClng(isRecommend)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>是否推荐出错(出现异常)！</li>"
		End If
	End If
	
	'拦截错误，不然错误往下进行！
	If FoundErr=True Then Exit Sub
	
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT * FROM "& CurrentTableName
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	RS.AddNew
		
		RS("topic")					=topic	'必填项
		RS("logtext")				=logtext
		RS("details_class_id")		=details_class_id
		RS("photo")					=photo
		RS("isOnpublic")			=isOnpublic
		RS("details_orderid")		=details_orderid
		
		RS("photo_desc")			=photo_desc
		
		
		RS("authorid")				=Session("enterName")
		RS("author")				=Session("enterCnName")
		
		RS("isRecommend")			=isRecommend
	
	RS.Update
	
	RS.Close
	Set RS=Nothing
	
'记入日志.
Call CokeShow.AddLog("添加操作：成功添加了"& UnitName &"-"& topic, sql)
	
	CokeShow.ShowOK "添加"& UnitName &"成功!",CurrentPageNow
End Sub


Sub SaveModify()
	Dim topic,logtext,details_class_id,photo,isOnpublic,details_orderid
	Dim photo_desc
	Dim isRecommend
	
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
	topic				=CokeShow.filtRequest(Request("topic"))
	logtext				=CokeShow.filtRequestRich(Request("logtext"))
	details_class_id	=CokeShow.filtRequest(Request("details_class_id"))
	photo				=CokeShow.filtRequest(Request("photo"))
	isOnpublic			=CokeShow.filtRequest(Request("isOnpublic"))
	details_orderid		=CokeShow.filtRequest(Request("details_orderid"))
	
	photo_desc			=CokeShow.filtRequest(Request("photo_desc"))
	
	isRecommend			=CokeShow.filtRequest(Request("isRecommend"))
	
	
	'验证
	If topic="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>标题不能为空！</li>"
	Else
		If CokeShow.strLength(topic)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>标题长度不能大于50个字符！</li>"
		Else
			topic=topic
		End If
	End If
		
	If logtext="" Or isNull(logtext) Or isEmpty(logtext) Then
'		FoundErr=True
'		ErrMsg=ErrMsg &"<br><li>正文不能为空！</li>"
		logtext=""
	Else
		If CokeShow.strLength(logtext)>8000 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>正文长度不能大于8000个字符！</li>"
		Else
			logtext=logtext
		End If
	End If
	
	If details_class_id="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请选择所属分类！</li>"
	Else
		If isNumeric(details_class_id) Then
			details_class_id=CokeShow.CokeClng(details_class_id)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前所属分类参数不是数字(出现异常)！</li>"
		End If
	End If
	
	If photo="" Or isNull(photo) Or isEmpty(photo) Then
		photo=""
		
	Else
		If Len(photo)>10 Then
			photo=photo
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>上传图片获得的图片地址出错(出现异常)！</li>"
		End If
	End If
	
	If isOnpublic="" Or isNull(isOnpublic) Or isEmpty(isOnpublic) Then
		isOnpublic=0
		
	Else
		If isNumeric(isOnpublic) Then
			isOnpublic=CokeShow.CokeClng(isOnpublic)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前发布状态出错(出现异常)！</li>"
		End If
	End If
	
	
	If photo_desc="" Or isNull(photo_desc) Or isEmpty(photo_desc) Then
		photo_desc=""
		
	Else
		If Len(photo_desc)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>图片文字描述字符长度不得超过50！</li>"
		Else
			photo_desc=photo_desc
		End If
	End If
	
	If isRecommend="" Or isNull(isRecommend) Or isEmpty(isRecommend) Then
		isRecommend=0
		
	Else
		If isNumeric(isRecommend) Then
			isRecommend=CokeShow.CokeClng(isRecommend)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>是否推荐出错(出现异常)！</li>"
		End If
	End If
	
	'拦截错误，不然错误往下进行！
	If FoundErr=True Then Exit Sub
	
	
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT * FROM "& CurrentTableName &" WHERE id="& intID
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,1,3
	
	'拦截此记录的异常情况.
	If RS.Bof And RS.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的"& UnitName &"！</li>"
		Exit Sub
	End If
	
		RS("topic")					=topic	'必填项
		RS("logtext")				=logtext
		RS("details_class_id")		=details_class_id
		RS("photo")					=photo
		RS("isOnpublic")			=isOnpublic
		RS("details_orderid")		=details_orderid
		
		RS("photo_desc")			=photo_desc
		
		
		RS("modifydate")			=Now()
		
		RS("authorid")				=Session("enterName")
		RS("author")				=Session("enterCnName")
		
		RS("authorid_last")			=Session("enterName")
		
		RS("isRecommend")			=isRecommend
	
	RS.Update
	RS.Close
	Set RS=Nothing
	
'记入日志.
Call CokeShow.AddLog("编辑操作：成功编辑了ID为"& intID &"的"& UnitName &"-"& topic, sql)
	
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


'例子:AdditionZero(RS("classid"), 8)
'给各个以逗号分隔的参数数字，都不上足够的零！
'参数:
'theNumbersStringNow:以逗号为分隔的id集合，或者是一个数字——扩展分类product_class_id_extend.
Public Function toProcessRequest(theNumbersStringNow)
	If theNumbersStringNow="" Or isNull(theNumbersStringNow) Or isEmpty(theNumbersStringNow) Then
		toProcessRequest=""
		Exit Function
	End If
	'如果只有一个（就是一个纯数字）.
	If isNumeric(theNumbersStringNow) Then
		toProcessRequest = CokeShow.AdditionZero( theNumbersStringNow, 8 )
		Exit Function
	End If
	
	'如果不是纯数字（就是一批扩展分类了），则循环加工字符串.前边加零不影响id的读取.
	toProcessRequest=""
	Dim i_N
	For i_N=0 To Ubound( Split(theNumbersStringNow,",") )
		toProcessRequest = toProcessRequest &","& CokeShow.AdditionZero( Trim(Split(theNumbersStringNow,",")(i_N)), 8 )
	Next
	toProcessRequest = Right(toProcessRequest, Len(toProcessRequest)-1)
	
End Function


%>