<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True	'缓冲数据，才向客户输出.
Response.Charset="utf-8"
Session.CodePage = 65001
%>
<%
'模块说明：后台管理人员帐号管理模块.
'日期说明：2009-7-9
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
Dim CurrentPageNow,TitleName,UnitName
CurrentPageNow 	= "product_chiliIndex.asp"	'当前页.
TitleName 		= "辣椒指数管理"	'此模块管理页的名字.
UnitName 		= "辣椒指数"					'此模块涉及记录的元素名.
'自定义设置.
'本地设置.
Dim CurrentTableName
CurrentTableName 	= "[CXBG_product_chiliIndex]"	'此模块涉及的[表]名.
'非本地设置.
Dim UpOther_Table,UpOther_TableField
UpOther_Table 		= "[user]"				'此模块合并操作时，需要同时更新的外关联[表]名信息.
UpOther_TableField 	= "xx_classid"			'此模块合并操作时，需要同时更新的外关联[表]的字段名信息.
%>





<%
Dim RS, sql						'查询列表记录用的变量.
Dim FoundErr,ErrMsg				'控制错误流程用的控制变量.
Dim ParentID					'用于区别是否为添加子项操作的控制变量.
Dim Action						'流程控制用的控制变量.
Dim strGuide					'导航文字.

Action		=CokeShow.filtRequest(Request("Action"))
ParentID	=CokeShow.filtRequest(Request("ParentID"))

'处理区别是否为添加子项操作的控制变量.
If ParentID="" Then
	ParentID=0
Else
	ParentID=CokeShow.CokeClng(ParentID)
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
//		dojo.require("dijit.form.CheckBox");
		dojo.require("dijit.form.Button");
		dojo.require("dijit.form.Form");
		dojo.require("dijit.form.Textarea");
		//临时不用.
		//dojo.require("dijit.Editor");
		dojo.require("dijit.Dialog");
		//上传图片.
		dojo.require("dojo.io.iframe");
		
		
	</script>
	
	<script type="text/javascript">
		//Table偶数行变色函数.
		dojo.addOnLoad(function() {
			stripeTables("listGo");
		});
	</script>
	
	
	<!--#include file="_screen_JS.asp"-->
	
	
	



<script type="text/javascript"
>

function Welcome() {
   	confirm('欢迎今天第一次使用上传组件，您喜欢此款产品吗！)');
}
</script>


</head>
<body class="<% =classname_dj_ThemesCss %>">


<!--main-->
	<!-- Begin mainleft-->
	<div class="newsContainer">
		
		<div class="news1">
			
			<h3>相关管理项目</h3>
				
				<p>
					<!--#include file="menu_product.asp"-->
				</p>
				
			<h3>当前操作</h3>
				
				<ul>
					<!--<li><a href="http://www.iw3c2.org/">&#187;IW<sup>3</sup>C<sup>2</sup></a></li>-->
					<li><a href="?Action=Add">&#187;+新增<% =UnitName %></a></li>
					<li><a href="<% = CurrentPageNow %>">&#187;返回列表</a></li>
					
					<li><a href="?Action=Order">&#187;排序操作</a></li>
					<!--<li><a href="?Action=OrderN">&#187;多级排序操作</a></li>
					<li><a href="?Action=Reset">&#187;复位所有<% =UnitName %></a></li>
					<li><a href="?Action=Unite">&#187;合并所有<% =UnitName %></a></li>-->
					
				</ul>
				
		</div><!-- End news1-->
		
		<!--<div class="news2">
		
			<h3>查询操作</h3>-->
			<!--<form action="<% =CurrentPageNow %>" method="GET" name="custForm" id="custForm"
			dojoType="dijit.form.Form"
			>
			<p>
					
					<select name="TypeSearch" id="TypeSearch">
					    
					    <option value="id" selected>按ID查询</option>
					    <option value="cnname" >按中文名查询</option>
						<option value="readme" >按备注查询</option>
						
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
                     class="input_tell"
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
			</form>-->
		<!--</div>--><!-- End news2-->
	
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
			Call AddClass()
		ElseIf Action="SaveAdd" Then
			Call SaveAdd()
		ElseIf Action="Modify" Then
			Call Modify()
		ElseIf Action="SaveModify" Then
			Call SaveModify()
		ElseIf Action="Move" Then
			Call MoveClass()
		ElseIf Action="SaveMove" Then
			Call SaveMove()
		ElseIf Action="Del" Then
			Call DeleteClass()
		ElseIf Action="UpOrder" Then 
			Call UpOrder() 
		ElseIf Action="DownOrder" Then 
			Call DownOrder() 
		ElseIf Action="Order" Then
			Call Order()
		ElseIf Action="UpOrderN" Then 
			Call UpOrderN() 
		ElseIf Action="DownOrderN" Then 
			Call DownOrderN() 
		ElseIf Action="OrderN" Then
			Call OrderN()
		ElseIf Action="Reset" Then
			Call Reset()
		ElseIf Action="SaveReset" Then
			Call SaveReset()
		ElseIf Action="Unite" Then
			Call Unite()
		ElseIf Action="SaveUnite" Then
			Call SaveUnite()
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
	strGuide=strGuide & TitleName
	
	Dim arrShowLine(10)
	Dim sqlClass,rsClass,i,iDepth
	
	'定义一个每元素全为False的数组.
	For i=0 To Ubound(arrShowLine)
		arrShowLine(i)=False
	Next
	
	sqlClass="SELECT * From "& CurrentTableName &" ORDER BY RootID,OrderID"
	Set rsClass=Server.CreateObject("Adodb.RecordSet")
	rsClass.Open sqlClass,CONN,1,1
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
						<th><% =UnitName %>名称</th>
						<th>图标</th>
						<th>操作</th>
					  </tr>
					  </thead>
					  <tbody>
					  
					  <%
					  If rsClass.EOF Then
					  %>
					  <tr>
						<td colspan="3" style="color:red;">对不起，没有记录...</td>
					  </tr>
					  <%
					  End If
					  %>
					  
					  
					  <%
					  Do While Not rsClass.EOF
					  %>
					  <tr>
						<td style="width: 280px;">
						<%
						'列表展示中的分类.(列表展示)
						iDepth=rsClass("Depth")
						
						'arrShowLine(10)全是False.
						If rsClass("NextID")>0 Then
							'如果有下一个分类的，则它有深度.
							arrShowLine(iDepth)=True
						Else
							'如果没有下一分类的，则它没有深度.
							arrShowLine(iDepth)=False
						End If
						
						'如果此分类深度‘深’，则检测究竟有多‘深’，然后按深度输出它的结构表现出来.
						If iDepth>0 Then
							'以深度输出字符串结构.
							For i=1 To iDepth 
								'如果到达了‘横向’深度最‘深’之处时，（从左边到最右边要输出类名的时候），准备输出‘最后一个符号’（图）.
								If i=iDepth Then 
									If rsClass("NextID")>0 Then 
									'├ 图标.(如果有下一个跟着,输出├ )
										Response.Write "<img src='"& system_dir &"images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>" 
									Else 
									'└ 图标.(如果没有下一个跟着,输出└ 结束本大类结构！ )
										'response.Write "&nbsp;&nbsp;├ "
										Response.Write "<img src='"& system_dir &"images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>" 
									End If 
								Else
								'如果未达到‘横向’最深，正在‘横向’的中间时， 准备输出
									If arrShowLine(i)=True Then
									'│ 竖线图标.(有一个标记,输出│ )
										Response.Write "<img src='"& system_dir &"images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>" 
									Else
									'   空图标.(其它,输出   )
										'response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "
										response.write "<img src='"& system_dir &"images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>" 
									End If
								End If 
							Next 
						  End If
						  
						  If rsClass("Child")>0 Then 
							'response.write "<img src='"& system_dir &"images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>" 
						  Else 
							'response.write "<img src='"& system_dir &"images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>" 
						  End If
						  
						  '如果深度为0，一级分类，粗体.
						  If rsClass("Depth")=0 Then 
							Response.Write "<b>" 
						  End If 
						  Response.Write "<a href='?Action=Modify&id="& rsClass("id") &"' title='"& rsClass("ReadMe") &"'>"& rsClass("classname") &"</a>"
						  '如果有子类.
						  If rsClass("Child")>0 Then 
							Response.Write "&nbsp;("& rsClass("Child") &")" 
						  End If
						  
						  
						  %>
							&nbsp;
						</td>
						<td style="text-align:center;">
							
							<%
							If rsClass("photo")<>"" And Len(rsClass("photo"))>3 Then
							%>
								
								<img id="src_photo<% =rsClass("classid") %>" src="<% =rsClass("photo") %>" width="30" height="18" />
								<!--begin-->
									<span dojoType="dijit.Tooltip"
									connectId="src_photo<% =rsClass("classid") %>"
									id="tmp<% =rsClass("classid") %>"
									style="display:none;"
									>
										图片详细尺寸：<br /><img src="<% =rsClass("photo") %>" />
									</span>
								<!--end-->
							<%
							End If
							%>
						</td>
						<td>
							<!--<a href="?Action=Add&ParentID=<%=rsClass("id")%>">+向此新增子<% =UnitName %></a>-->
							<!--&nbsp;|&nbsp;-->
							<a href="?Action=Modify&id=<%=rsClass("id")%>">修改</a>
							&nbsp;|&nbsp;
							<!--<a href="?Action=Move&id=<%=rsClass("id")%>">移动</a>
							&nbsp;|&nbsp;-->
							<a href="?Action=Del&id=<%=rsClass("id")%>" onClick="<% If rsClass("Child")>0 Then %>return ConfirmDel1();<% Else %>return ConfirmDel2();<% End If %>">删除</a>
					  </tr>
					  <%
						  rsClass.MoveNext
					  Loop
					  %>
					  
					  
					  
					  </tbody>
					</table>
					
					
				</form>
			</p>
					
			<p>&nbsp;</p>
			
		
		</div>
		<!--mainInfo1-->
		<!--mainInfo-->
		
		
		<script language="JavaScript" type="text/JavaScript">
			function ConfirmDel1()
			{
			   alert("此<% =UnitName %>下还有子<% =UnitName %>，必须先删除下属子<% =UnitName %>后才能删除此<% =UnitName %>！");
			   return false;
			}
			
			function ConfirmDel2()
			{
			   if(confirm("删除<% =UnitName %>将不能恢复！确定要删除此<% =UnitName %>吗？"))
				 return true;
			   else
				 return false;
				 
			}
		</script>

<%
End Sub


Sub AddClass()
	strGuide=strGuide &"新增"& UnitName
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
				<form action="<% = CurrentPageNow %>" method="post" name="form1" id="form1"
				dojoType="dijit.form.Form"
				execute="processForm('form1')"
				><!--execute=别用新版-->
					
					<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
					  <thead>
					  <tr>
						<th style="text-align: right;">名称</th>
						
						<th style="text-align: left;">填写数据</th>
					  </tr>
					  </thead>
					  <tbody>
					  
					  <!--<tr>
						<td style="text-align: right; width: 180px;">
							所属<% =UnitName %>：
						</td>
						
						<td style="text-align: left;">
							<select name="ParentID">
								<% 'Call CokeShow.ClassOption_id(CurrentTableName,"",0,ParentID) %>
							</select>
						</td>
					  </tr>-->
					  <input type="hidden" name="ParentID" value="0" />
					  
					  <tr>
						<td style="text-align: right;">
							<% =UnitName %>名称：
						</td>
						
						<td style="text-align: left;">
							<input type="text" id="classname" name="classname"
							dojoType="dijit.form.ValidationTextBox"
							required="true"
							promptMessage="<% =UnitName %>名称为必填项."
							invalidMessage="<% =UnitName %>长度必须在1-50字之内，例如：商务人士类"
							trim="true"
							lowercase="false"
							
							regExp=".{1,50}"
							 value=""
                             class="input_tell"
							/>
						</td>
					  </tr>
					  
					  <!--<tr>
						<td style="text-align: right;" valign="top">
							<% =UnitName %>说明：
							
						</td>
						
						<td style="text-align: left;">
							
							<textarea name="readme_hidden" id="readme_hidden"
							dojoType="dijit.form.Textarea"
							 style="width: 350px;"></textarea>
							
						</td>
					  </tr>-->
					  
					  
					  <tr>
						<td style="text-align: right;" valign="top">
							上传图片：
							<!--上传图片-->
							<script type="text/javascript">
							
							</script>
						</td>
						
						<td style="text-align: left;">
							<span style="color:#999999;"> 尺寸要求:宽30*高18 (长方形)</span>
							<br />
							<br />
							
							<a href="#" target="_blank" id="href_photo">
							<img style="border:5px #999999 solid;" src="/images/NoPic.png" width="30" height="18" onerror='this.src="/images/NoPic.png"'
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
                             style="width:500px;"
                             class="input_tell"
							/>
							
							</div>
						</td>
					  </tr>
					  
					  <tr>
						<td colspan="2">
						  
						  
						  
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
					
					
					<!--临时不用.-->
					<input type="hidden" name="readme_hidden" id="readme_hidden"
				    value=""
				    />
					
					<input type="hidden" name="Action"
					  value="SaveAdd"
					  />
				</form>
			</p>
					
			<p>&nbsp;</p>
			
		
		</div>
		<!--mainInfo1-->
		<!--mainInfo-->
		

<%
End Sub


Sub Modify()
	strGuide=strGuide &"修改 "& UnitName
	
	Dim id,rsClass,i
	id=CokeShow.filtRequest(Request("id"))
	
	'处理id传值
	If id="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>参数不足！</li>"
		Exit Sub
	Else
		id=CokeShow.CokeClng(id)
	End If
	
	sql="SELECT * FROM "& CurrentTableName &" WHERE id="& id
	Set rsClass=Server.CreateObject ("Adodb.RecordSet")
	If Not IsObject(CONN) Then link_database
	rsClass.Open sql,CONN,1,3
	
	If rsClass.Bof And rsClass.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的"& UnitName &"！</li>"
		Exit Sub
	Else
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
				<form action="<% = CurrentPageNow %>" method="post" name="form1" id="form1"
				dojoType="dijit.form.Form"
				execute="processForm('form1')"
				><!--execute=别用新版-->
					
					<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
					  <thead>
					  <tr>
						<th style="text-align: right;">名称</th>
						
						<th style="text-align: left;">填写数据</th>
					  </tr>
					  </thead>
					  <tbody>
					  
					  <!--<tr>
						<td style="text-align: right; width: 180px;">
							所属<% =UnitName %>：
							<br />
							<a href="?Action=Move&id=<% =id %>">&#187;&#187;移动<% =UnitName %></a>
						</td>
						
						<td style="text-align: left;">
							<%
							'ClassParent(sqlTable,sqlWhere,UnitName,ParentID,ParentPath)
							''Call CokeShow.ClassParent(CurrentTableName,"",UnitName,rsClass("ParentID"),rsClass("ParentPath"))
							%>
						</td>
					  </tr>-->
					  
					  <tr>
						<td style="text-align: right;">
							<% =UnitName %>名称：
						</td>
						
						<td style="text-align: left;">
							<input type="text" id="classname" name="classname"
							dojoType="dijit.form.ValidationTextBox"
							required="true"
							promptMessage="<% =UnitName %>名称为必填项."
							invalidMessage="<% =UnitName %>长度必须在1-50字之内，例如：商务人士类"
							trim="true"
							lowercase="false"
							
							regExp=".{1,50}"
							 value="<% =rsClass("classname") %>"
                             class="input_tell"
							/>
						</td>
					  </tr>
					  
					  <!--<tr>
						<td style="text-align: right;" valign="top">
							<% =UnitName %>说明：
							
						</td>
						
						<td style="text-align: left;">
							
							<textarea name="readme_hidden" id="readme_hidden"
							dojoType="dijit.form.Textarea"
							 style="width: 350px;"><% =rsClass("ReadMe") %></textarea>
							
						</td>
					  </tr>-->
					  
					  
					  
					  
					  <tr>
						<td style="text-align: right;" valign="top">
							上传图片：
							<!--上传图片-->
							<script type="text/javascript">
								
							</script>
						</td>
						
						<td style="text-align: left;">
							<span style="color:#999999;"> 尺寸要求:宽30*高18 (长方形)</span>
							<br />
							<br />
							
							<a href="<% =rsClass("photo") %>" target="_blank" id="href_photo">
							<img style="border:5px #999999 solid;" src="<% If rsClass("photo")<>"" Then Response.Write rsClass("photo") Else Response.Write "/images/NoPic.png" %>" width="30" height="18" onerror='this.src="/images/NoPic.png"'
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
							 value="<% =rsClass("photo") %>"
                             style="width:500px;"
                             class="input_tell"
							/>
							
							</div>
						</td>
					  </tr>
					  
					  
					  <tr>
						<td colspan="2">
						  
						  
						  
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
					
					
					<!--临时不用.-->
					<input type="hidden" name="readme_hidden" id="readme_hidden"
				    value=""
				    />
					
					<input type="hidden" name="id"
					  value="<% =rsClass("id") %>"
					  />
					  
					<input type="hidden" name="Action"
					  value="SaveModify"
					  />
				</form>
			</p>
					
			<p>&nbsp;</p>
			
		
		</div>
		<!--mainInfo1-->
		<!--mainInfo-->
 		
		
<%
	End If
	
	rsClass.Close
	Set rsClass=Nothing
	
End Sub


Sub MoveClass()
	strGuide=strGuide &"移动"& UnitName
	
	Dim id,rsClass,i
	id=CokeShow.filtRequest(Request("id"))
	
	'处理id传值
	If id="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>参数不足！</li>"
		Exit Sub
	Else
		id=CokeShow.CokeClng(id)
	End If
	
	sql="SELECT * FROM "& CurrentTableName &" WHERE id="& id
	Set rsClass=Server.CreateObject ("Adodb.RecordSet")
	rsClass.Open sql,CONN,1,3
	
	If rsClass.Bof And rsClass.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的"& UnitName &"！</li>"
		Exit Sub
	Else
	
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
				<form action="<% = CurrentPageNow %>" method="post" name="form1" id="form1"
				dojoType="dijit.form.Form"
				
				>
					
					<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
					  <thead>
					  <tr>
						<th style="text-align: right;">名称</th>
						
						<th style="text-align: left;">填写数据</th>
					  </tr>
					  </thead>
					  <tbody>
					  
					  
					  <tr>
						<td style="text-align: right;">
							当前<% =UnitName %>名称：
						</td>
						
						<td style="text-align: left;">
							<b><% =rsClass("classname") %></b>
							<input type="hidden" name="id" id="id"
							  value="<% =rsClass("id") %>"
							  />
						</td>
					  </tr>
					  
					  <tr>
						<td style="text-align: right;" valign="top">
							所属<% =UnitName %>：
						</td>
						
						<td style="text-align: left;">
							<%
							'ClassParent(sqlTable,sqlWhere,UnitName,ParentID,ParentPath)
							Call CokeShow.ClassParent(CurrentTableName,"",UnitName,rsClass("ParentID"),rsClass("ParentPath"))
							%>
						</td>
					  </tr>
					  
					  
					  <tr>
						<td style="text-align: right;" valign="top">
							移动到：
							<br />
							(不能为<% =rsClass("classname") %>自己的下属子类)
							
						</td>
						
						<td style="text-align: left;">
							<select name="ParentID" size="2" style="height:300px;width:500px;">
								<% Call CokeShow.ClassOption_id(CurrentTableName,"",0,rsClass("ParentID")) %>
							</select>
						</td>
					  </tr>
					  
					  
					  <tr>
						<td colspan="2">
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
					
					<!--
					临时不用.
					<input type="hidden" name="readme_hidden" id="readme_hidden"
				    value=""
				    />-->
					
					<input type="hidden" name="Action"
					  value="SaveMove"
					  />
				</form>
			</p>
					
			<p>&nbsp;</p>
			
		
		</div>
		<!--mainInfo1-->
		<!--mainInfo-->
		
		
<%
	End If
	
End Sub


Sub Order()
	strGuide=strGuide &"一级"& UnitName &"排序"
	
	'定义变量.
	Dim sqlClass,rsClass,i,iCount,j
	
	'查询父类为0的分类，即一级分类.
	sqlClass="SELECT * FROM "& CurrentTableName &" WHERE ParentID=0 ORDER BY RootID"
	Set rsClass=Server.CreateObject("Adodb.RecordSet")
	rsClass.Open sqlClass,CONN,1,1
	
	'获取记录总数.
	iCount=rsClass.RecordCount
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
					
					<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
					  <thead>
					  <tr>
						<th style="text-align: right;" width="200"><% =UnitName %>名称</th>
						<th style="text-align: center;" width="160">向上移动操作</th>
						<th style="text-align: center;" width="160">向下移动操作</th>
						
					  </tr>
					  </thead>
					  <tbody>
					  
					  <%
					  '初始化j.
					  j=1
					  
					  Do While Not rsClass.Eof
					  %>
					  <tr>
						<td style="text-align: left;">
							<b><% =rsClass("classname") %></b>
						</td>
						
						
						<% 
						'输出向上移和向下移按钮组合.
						'如果不是此记录的头一回输出，则输出 向上移 按钮（如果是头一回，则输出空单元格.）.
						If j>1 Then
							Response.Write "<td>"
							Response.Write "<form action='?Action=UpOrder' method='post' dojoType=""dijit.form.Form"">"
							Response.Write "<select name='MoveNum' size='1'><option value='0'>向上移</option>"
							'向上移的位数选择.
							For i=1 To j-1
								Response.Write "<option value="& i &">"& i &"</option>"
							Next
							Response.Write "</select>"
							'传递隐藏数据.
							Response.Write "<input type=""hidden"" name=""id"" value="""& rsClass("id") &""" />"
							Response.Write "<input type=""hidden"" name=""cRootID"" value="""& rsClass("RootID") &""" />&nbsp;"
							'提交按钮.
							Response.Write "<button type=""submit""  dojoType=""dijit.form.Button"" onclick=""return confirm('确实要移动此"& UnitName &"吗？');"">&nbsp;移&nbsp;动&nbsp;</button>"
							
							Response.Write "</form>"
							Response.Write "</td>"
						Else
						'否则输出空单元格.
							Response.Write "<td>&nbsp;</td>"
						End If
						
						'当没到最后一条记录的时候，后边继续输出 向下移 按钮.
						If iCount>j Then
							Response.Write "<td>"
							Response.Write "<form action=""?Action=DownOrder"" method=""post"" dojoType=""dijit.form.Form"">"
							Response.Write "<select name=""MoveNum"" size=""1""><option value='0'>向下移</option>"
							'向下移的位数选择.
							For i=1 To iCount-j
								Response.Write "<option value="""& i &""">"& i &"</option>"
							Next
							Response.Write "</select>"
							'传递隐藏数据.
							Response.Write "<input type=""hidden"" name=""id"" value="""& rsClass("id") &""" />"
							Response.Write "<input type=""hidden"" name=""cRootID"" value="""& rsClass("RootID") &""" />&nbsp;"
							'提交按钮.
							Response.Write "<button type=""submit""  dojoType=""dijit.form.Button"" onclick=""return confirm('确实要移动此"& UnitName &"吗？');"">&nbsp;移&nbsp;动&nbsp;</button>"
							
							Response.Write "</form>"
							Response.Write "</td>"
						Else
						'否则输出空单元格.
							Response.Write "<td>&nbsp;</td>"
						End If
						%>
						
					  </tr>
					  <%
						  j=j+1
						  rsClass.MoveNext
					  Loop
					  %>
					  
					  
					  </tbody>
					</table>
					
					
			</p>
					
			<p>&nbsp;</p>
			
		
		</div>
		<!--mainInfo1-->
		<!--mainInfo-->
		
<% 
	
End Sub


Sub OrderN()
	strGuide=strGuide &"多级"& UnitName &"排序"
	
	'定义变量.
	Dim sqlClass,rsClass,i,iCount,trs,UpMoveNum,DownMoveNum
	
	sqlClass="SELECT * FROM "& CurrentTableName &" ORDER BY RootID,OrderID"
	Set rsClass=Server.CreateObject("Adodb.RecordSet")
	rsClass.Open sqlClass,CONN,1,1
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
					
					<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
					  <thead>
					  <tr>
						<th style="text-align: right;" width="200"><% =UnitName %>名称</th>
						<th style="text-align: center;" width="160">向上移动操作</th>
						<th style="text-align: center;" width="160">向下移动操作</th>
						
					  </tr>
					  </thead>
					  <tbody>
					  
					  <%
					  Do While Not rsClass.Eof
					  %>
					  <tr>
						<td style="text-align: left;">
							<%
							'按照格式输出分类名.
							'按‘深度’输出空格，形成缩进.
							For i=1 To rsClass("Depth")
								Response.Write "&nbsp;&nbsp;&nbsp;"
							Next
							'如果有子分类，输出+图标.
							If rsClass("Child")>0 Then
								Response.Write "<img src='"& system_dir &"images/tree_folder4.gif' width='15' height='15' valign='abvmiddle' />"
							Else
							'如果没有子分类，输出-图标.
								Response.Write "<img src='"& system_dir &"images/tree_folder3.gif' width='15' height='15' valign='abvmiddle' />"
							End If
							'如果此类上边没有父级分类存在，即是一个根类，则使用粗体字表示.
							If rsClass("ParentID")=0 Then
								Response.Write "<b>"
							End If
							'输出类名.
							Response.Write rsClass("classname")
							'如果下边有子类，则输出包含子类个数情况.
							If rsClass("Child")>0 Then
								Response.Write "("& rsClass("Child") &")"
							End If
							%>
						</td>
						
						<%
						'输出向上移和向下移按钮组合.
						'如果上边有父类(如果不是一级分类).
						If rsClass("ParentID")>0 Then   '则算出相同‘深度’的分类数目，得到该分类在相同深度的分类中所处位置（之上或者之下的分类数）
							'查询和本类同一级别中，并且在本类之上的兄弟类个数（可向上移动数）.
							Set trs=CONN.Execute("SELECT COUNT(id) FROM "& CurrentTableName &" WHERE ParentID="& rsClass("ParentID") &" AND OrderID<"& rsClass("OrderID") &"")
							UpMoveNum=trs(0)	'可向上移动数.
							If isNull(UpMoveNum) Then UpMoveNum=0	'设置可向上移动数默认值为0.
							
							'如果可向上移动数大于0，则输出上移按钮.
							If UpMoveNum>0 Then
								Response.Write "<td>"
								Response.Write "<form action='?Action=UpOrderN' method='post' dojoType=""dijit.form.Form"">"
								Response.Write "<select name='MoveNum' size='1'><option value='0'>向上移</option>"
								
								'向上移的位数选择.
								For i=1 To UpMoveNum
									Response.Write "<option value='"& i &"'>"& i &"</option>"
								Next
								Response.Write "</select>"
								'传递隐藏数据.
								Response.Write "<input type='hidden' name='id' value='"&rsClass("id")&"' />&nbsp;"
								'提交按钮.
								Response.Write "<button type=""submit""  dojoType=""dijit.form.Button"" onclick=""return confirm('确实要移动此"& UnitName &"吗？');"">&nbsp;移&nbsp;动&nbsp;</button>"
								
								Response.Write "</form>"
								Response.Write "</td>"
							Else
								Response.Write "<td>&nbsp;</td>"
							End If
							trs.Close
							
							
							'查询和本类同一级别中，并且在本类之下的兄弟类个数（可向下移动数）.
							Set trs=CONN.Execute("SELECT COUNT(id) FROM "& CurrentTableName &" WHERE ParentID="& rsClass("ParentID") &" AND orderID>"& rsClass("orderID") &"")
							DownMoveNum=trs(0)		'可向下移动数.
							If isNull(DownMoveNum) Then DownMoveNum=0	'设置可向下移动数默认值为0.
							
							'如果可向下移动数大于0，则输出下移按钮.
							If DownMoveNum>0 Then
								Response.Write "<td>"
								Response.Write "<form action='?Action=DownOrderN' method='post' dojoType=""dijit.form.Form"">"
								Response.Write "<select name='MoveNum' size='1'><option value='0'>向下移</option>"
								
								'向下移的位数选择.
								For i=1 To DownMoveNum
									Response.Write "<option value='"& i &"'>"& i &"</option>"
								Next
								Response.Write "</select>"
								'传递隐藏数据.
								Response.Write "<input type='hidden' name='id' value='"& rsClass("id") &"' />&nbsp;" 
								'提交按钮.
								Response.Write "<button type=""submit""  dojoType=""dijit.form.Button"" onclick=""return confirm('确实要移动此"& UnitName &"吗？');"">&nbsp;移&nbsp;动&nbsp;</button>"
								
								Response.Write "</form>"
								Response.Write "</td>"
							Else
								Response.Write "<td>&nbsp;</td>"
							End If
							trs.Close
						Else
							Response.Write "<td>&nbsp;</td><td>&nbsp;</td>"
						End If 
						%>
							
					  </tr>
					  <% 
						  UpMoveNum=0
						  DownMoveNum=0
						  rsClass.MoveNext
					  Loop
					  %> 
					  
					  
					  </tbody>
					</table>
					
					
			</p>
					
			<p>&nbsp;</p>
			
		
		</div>
		<!--mainInfo1-->
		<!--mainInfo-->
<%
	
End Sub


Sub Reset()
	strGuide=strGuide &"复位所有"& UnitName
	
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
				<form action="<% = CurrentPageNow %>?Action=SaveReset" method="post" name="form1" id="form1"
				dojoType="dijit.form.Form"
				
				>
					
					<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
					  <thead>
					  <tr>
						<th style="text-align: center;">注意：</th>
					  </tr>
					  </thead>
					  <tbody>
					  
					  
					  <tr>
						<td valign="top">
							
							<br /><br />
							如果选择复位所有<% =UnitName %>，则所有<% =UnitName %>都将作为一级<% =UnitName %>，这时您需要重新对各个<% =UnitName %>进行归属的基本设置。
							<br />
							不要轻易使用该功能，仅在做出了错误的设置而无法复原<% =UnitName %>之间的关系和排序的时候使用。
							<br /><br />
						</td>
						
					  </tr>
					  
					  
					  <tr>
						<td>
						      <button type="submit" id="submitbtn" 
							  dojoType="dijit.form.Button"
							  onclick="return confirm('您确实要复位所有<% =UnitName %>吗？');"
							  >
							  &nbsp;复位所有<% =UnitName %>&nbsp;
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
					临时不用.
					<input type="hidden" name="readme_hidden" id="readme_hidden"
				    value=""
				    />-->
					
					
				</form>
			</p>
					
			<p>&nbsp;</p>
			
		
		</div>
		<!--mainInfo1-->
		<!--mainInfo-->
		
<%
End Sub



Sub Unite()
	strGuide=strGuide &"合并所有"& UnitName
	
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
				<form action="<% = CurrentPageNow %>" method="post" name="form1" id="form1" onSubmit="return ConfirmUnite();"
				dojoType="dijit.form.Form"
				
				>
					
					<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
					  <thead>
					  <tr>
						<th style="text-align: center;">注意事项：</th>
					  </tr>
					  </thead>
					  <tbody>
					  
					  
					  <tr>
						<td valign="top">
							
							<br />
							所有操作不可逆，请慎重操作！！！
							<br />
							不能在同一个<% =UnitName %>内进行操作，不能将一个<% =UnitName %>合并到其下属<% =UnitName %>中。目标<% =UnitName %>中不能含有子<% =UnitName %>。
							<br />
							合并后您所指定的<% =UnitName %>（或者包括其下属<% =UnitName %>）将被删除，所有用户将转移到目标<% =UnitName %>中。
							<br /><br />
						</td>
						
					  </tr>
					  
					  
					  <tr>
						<td valign="top">
							&nbsp;
							<b>将<% =UnitName %>:</b>
							<select name="id" id="id">
								<% Call CokeShow.ClassOption_id(CurrentTableName,"",1,0) %>
							</select>(将被合并,并被删除.)
							
							<font style="color:#FF0000; font-weight:bold;">&nbsp;——&gt;&nbsp;</font>
							
							<b>合并到:</b>
							<select name="Targetid" id="Targetid">
								<% Call CokeShow.ClassOption_id(CurrentTableName,"",4,0) %>
							</select>(将被保留.)
							&nbsp;
						</td>
						
					  </tr>
					  
					  
					  <tr>
						<td>
						      <button type="submit" id="submitbtn" 
							  dojoType="dijit.form.Button"
							  onclick="return confirm('您确实要合并<% =UnitName %>吗？');"
							  >
							  &nbsp;合并<% =UnitName %>&nbsp;
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
					
					
					
					<input type="hidden" name="Action" id="Action"
				    value="SaveUnite"
				    />
					
					
				</form>
			</p>
					
			<p>&nbsp;</p>
			
		
		</div>
		<!--mainInfo1-->
		<!--mainInfo-->
		
<%

End Sub


sub SaveAdd()
	dim id,classname,Readme,PrevOrderID
	dim sql,rs,trs
	dim RootID,ParentDepth,ParentPath,ParentStr,ParentName,Maxid,MaxRootID
	dim PrevID,NextID,Child
	
	Dim photo
	
	photo=CokeShow.filtRequest(Request("photo"))

	classname=CokeShow.filtRequest(Request("classname"))
	Readme=CokeShow.filtRequest(Request("readme_hidden"))
	if classname="" then
		response.Write "<br><li>"& UnitName &"名称不能为空！</li>"
		response.End()
	end if
	set rs = conn.execute("select Max(id) From "& CurrentTableName &"")
	Maxid=rs(0)
	if isnull(Maxid) then
		Maxid=0
	end if
	rs.close
	id=Maxid+1
	set rs=conn.execute("select max(rootid) From "& CurrentTableName &"")
	MaxRootID=rs(0)
	if isnull(MaxRootID) then
		MaxRootID=0
	end if
	rs.close
	RootID=MaxRootID+1
	
	if ParentID>0 then
		sql="select * From "& CurrentTableName &" where id=" & ParentID & ""
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>所属"& UnitName &"已经被删除！</li>"
		end if
		if FoundErr=True then
			rs.close
			set rs=nothing
			exit sub
		else	
			RootID=rs("RootID")
			ParentName=rs("classname")
			ParentDepth=rs("Depth")
			ParentPath=rs("ParentPath")
			Child=rs("Child")
			ParentPath=ParentPath & "," & ParentID     '得到此分类的父级分类路径
			PrevOrderID=rs("OrderID")
			if Child>0 then		
				dim rsPrevOrderID
				'得到与本分类同级的最后一个分类的OrderID
				set rsPrevOrderID=conn.execute("select Max(OrderID) From "& CurrentTableName &" where ParentID=" & ParentID)
				PrevOrderID=rsPrevOrderID(0)
				set trs=conn.execute("select id From "& CurrentTableName &" where ParentID=" & ParentID & " and OrderID=" & PrevOrderID)
				PrevID=trs(0)
				
				'得到同一父分类但比本分类级数大的子分类的最大OrderID，如果比前一个值大，则改用这个值。
				set rsPrevOrderID=conn.execute("select Max(OrderID) From "& CurrentTableName &" where ParentPath like '" & ParentPath & ",%'")
				if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
					if not IsNull(rsPrevOrderID(0))  then
				 		if rsPrevOrderID(0)>PrevOrderID then
							PrevOrderID=rsPrevOrderID(0)
						end if
					end if
				end if
			else
				PrevID=0
			end if

		end if
		rs.close
	else
		if MaxRootID>0 then
			set trs=conn.execute("select id From "& CurrentTableName &" where RootID=" & MaxRootID & " and Depth=0")
			PrevID=trs(0)
			trs.close
		else
			PrevID=0
		end if
		PrevOrderID=0
		ParentPath="0"
	end if

	sql="Select * From "& CurrentTableName &" Where ParentID=" & ParentID & " AND classname='" & classname & "'"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	if not(rs.bof and rs.eof) then
		FoundErr=True
		if ParentID=0 then
			ErrMsg=ErrMsg & "<br><li>已经存在一级"& UnitName &"：" & classname & "</li>"
		else
			ErrMsg=ErrMsg & "<br><li>“" & ParentName & "”中已经存在子"& UnitName &"“" & classname & "”！</li>"
		end if
		rs.close
		set rs=nothing
		exit sub
	end if
	rs.close
	
	sql="Select top 1 * From "& CurrentTableName &""
	rs.open sql,conn,2,2
    rs.addnew
	rs("id")=id
   	rs("classname")=classname
	rs("RootID")=RootID
	rs("ParentID")=ParentID
	if ParentID>0 then
		rs("Depth")=ParentDepth+1
	else
		rs("Depth")=0
	end if
	rs("ParentPath")=ParentPath
	rs("OrderID")=PrevOrderID
	rs("Child")=0
	rs("Readme")=Readme
	rs("PrevID")=PrevID
	rs("NextID")=0
	
	If photo<>"" Then rs("photo")=photo
	
	rs.update
	rs.Close
    set rs=Nothing
	
	'更新与本分类同一父分类的上一个分类的“NextID”字段值
	if PrevID>0 then
		conn.execute("update "& CurrentTableName &" set NextID=" & id & " where id=" & PrevID)
	end if
	
	if ParentID>0 then
		'更新其父类的子分类数
		conn.execute("update "& CurrentTableName &" set child=child+1 where id="&ParentID)
		
		'更新该分类排序以及大于本需要和同在本分类下的分类排序序号
		conn.execute("update "& CurrentTableName &" set OrderID=OrderID+1 where rootid=" & rootid & " and OrderID>" & PrevOrderID)
		conn.execute("update "& CurrentTableName &" set OrderID=" & PrevOrderID & "+1 where id=" & id)
	end if
	
    'call CloseConn()
	
'记入日志.
Call CokeShow.AddLog("添加操作：成功添加了id为"& id &"的"& UnitName &"-"& classname, sql)
	
	Response.Redirect CurrentPageNow
end sub

sub SaveModify()
	dim classname,Readme,IsElite,ShowOnTop,Setting,ClassMaster,ClassPicUrl,LinkUrl,SkinID,LayoutID,BrowsePurview,AddPurview
	dim trs,rs
	dim id,sql,rsClass,i
	dim SkinCount,LayoutCount
	
	Dim photo
	
	photo=CokeShow.filtRequest(Request("photo"))
	
	id=CokeShow.filtRequest(Request("id"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		id=CokeShow.CokeClng(id)
	end if
	classname=CokeShow.filtRequest(Request("classname"))
	Readme=CokeShow.filtRequest(Request("readme_hidden"))
	if classname="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>"& UnitName &"名称不能为空！</li>"
	end if
	
	if FoundErr=True then
		exit sub
	end if	
	sql="select * From "& CurrentTableName &" where id=" & id
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的"& UnitName &"！</li>"
		rsClass.close
		set rsClass=nothing
		exit sub
	end if

	if FoundErr=True then
		rsClass.close
		set rsClass=nothing
		exit sub
	end if
	
   	rsClass("classname")=classname
	rsClass("Readme")=Readme
	
	If photo<>"" Then rsClass("photo")=photo
	
	rsClass.update
	rsClass.close
	set rsClass=nothing
	
	set rs=nothing
	set trs=nothing	
    'call CloseConn()
	
'记入日志.
Call CokeShow.AddLog("编辑操作：成功编辑了ID为"& id &"的"& UnitName &"-"& classname, sql)
	
	Response.Redirect CurrentPageNow
end sub


sub DeleteClass()
	dim sql,rs,PrevID,NextID,id
	id=CokeShow.filtRequest(Request("id"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		id=CokeShow.CokeClng(id)
	end if
	
	sql="select id,RootID,Depth,ParentID,Child,PrevID,NextID From "& CurrentTableName &" where id="&id
	set rs=server.CreateObject ("Adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>"& UnitName &"不存在，或者已经被删除</li>"
	else
		if rs("Child")>0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>该"& UnitName &"含有子"& UnitName &"，请删除其子"& UnitName &"后再进行删除本"& UnitName &"的操作</li>"
		end if
	end if
	if FoundErr=True then
		rs.close
		set rs=nothing
		exit sub
	end if
	PrevID=rs("PrevID")
	NextID=rs("NextID")
	if rs("Depth")>0 then
		conn.execute("update "& CurrentTableName &" set child=child-1 where id=" & rs("ParentID"))
	end if
	rs.delete
	rs.update
	rs.close
	set rs=nothing
	'修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update "& CurrentTableName &" set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "& CurrentTableName &" set PrevID=" & PrevID & " where id=" & NextID
	end if
	'call CloseConn()
	
'记入日志.
Call CokeShow.AddLog("删除操作：成功删除了ID为"& id &"的"& UnitName, sql)
	
	response.redirect CurrentPageNow
		
end sub


sub SaveMove()
	dim id,sql,rsClass,i
	dim rParentID
	dim trs,rs
	dim ParentID,RootID,Depth,Child,ParentPath,ParentName,iParentID,iParentPath,PrevOrderID,PrevID,NextID
	id=CokeShow.filtRequest(Request("id"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		id=CokeShow.CokeClng(id)
	end if
	
	sql="select * From "& CurrentTableName &" where id=" & id
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的"& UnitName &"！</li>"
		rsClass.close
		set rsClass=nothing
		exit sub
	end if

	rParentID=CokeShow.filtRequest(Request("ParentID"))
	if rParentID="" then
		rParentID=0
	else
		rParentID=CokeShow.CokeClng(rParentID)
	end if
	
	if rsClass("ParentID")<>rParentID then   '更改了所属分类，则要做一系列检查
		if rParentID=rsClass("id") then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>所属"& UnitName &"不能为自己！</li>"
		end if
		'判断所指定的分类是否为外部分类或本分类的下属分类
		if rsClass("ParentID")=0 then
			if rParentID>0 then
				set trs=conn.execute("select rootid From "& CurrentTableName &" where id="&rParentID)
				if trs.bof and trs.eof then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>不能指定外部"& UnitName &"为所属"& UnitName &"</li>"
				else
					if rsClass("rootid")=trs(0) then
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>不能指定该"& UnitName &"的下属"& UnitName &"作为所属"& UnitName &"</li>"
					end if
				end if
				trs.close
				set trs=nothing
			end if
		else
			set trs=conn.execute("select id From "& CurrentTableName &" where ParentPath like '"&rsClass("ParentPath")&"," & rsClass("id") & "%' and id="&rParentID)
			if not (trs.eof and trs.bof) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>您不能指定该"& UnitName &"的下属"& UnitName &"作为所属"& UnitName &"</li>"
			end if
			trs.close
			set trs=nothing
		end if
		
	end if

	if FoundErr=True then
		rsClass.close
		set rsClass=nothing
		exit sub
	end if
	
	if rsClass("ParentID")=0 then
		ParentID=rsClass("id")
		iParentID=0
	else
		ParentID=rsClass("ParentID")
		iParentID=rsClass("ParentID")
	end if
	Depth=rsClass("Depth")
	Child=rsClass("Child")
	RootID=rsClass("RootID")
	ParentPath=rsClass("ParentPath")
	PrevID=rsClass("PrevID")
	NextID=rsClass("NextID")
	rsClass.close
	set rsClass=nothing
	
	
  '假如更改了所属分类
  '需要更新其原来所属分类信息，包括深度、父级ID、分类数、排序、继承版主等数据
  '需要更新当前所属分类信息
  '继承版主数据需要另写函数进行更新--取消，在前台可用id in ParentPath来获得
  dim mrs,MaxRootID
  set mrs=conn.execute("select max(rootid) From "& CurrentTableName &"")
  MaxRootID=mrs(0)
  set mrs=nothing
  if isnull(MaxRootID) then
	MaxRootID=0
  end if
  dim k,nParentPath,mParentPath
  dim ParentSql,ClassCount
  dim rsPrevOrderID
  if CokeShow.CokeClng(parentid)<>rParentID and not (iParentID=0 and rParentID=0) then  '假如更改了所属分类
	'更新原来同一父分类的上一个分类的NextID和下一个分类的PrevID
	if PrevID>0 then
		conn.execute "update "& CurrentTableName &" set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "& CurrentTableName &" set PrevID=" & PrevID & " where id=" & NextID
	end if
	
	if iParentID>0 and rParentID=0 then  	'如果原来不是一级分类改成一级分类
		'得到上一个一级分类分类
		sql="select id,NextID From "& CurrentTableName &" where RootID=" & MaxRootID & " and Depth=0"
		set rs=server.CreateObject("Adodb.recordset")
		rs.open sql,conn,1,3
		PrevID=rs(0)      '得到新的PrevID
		rs(1)=id     '更新上一个一级分类分类的NextID的值
		rs.update
		rs.close
		set rs=nothing
		
		MaxRootID=MaxRootID+1
		'更新当前分类数据
		conn.execute("update "& CurrentTableName &" set depth=0,OrderID=0,rootid="&maxrootid&",parentid=0,ParentPath='0',PrevID=" & PrevID & ",NextID=0 where id="&id)
		'如果有下属分类，则更新其下属分类数据。下属分类的排序不需考虑，只需更新下属分类深度和一级排序ID(rootid)数据
		if child>0 then
			i=0
			ParentPath=ParentPath & ","
			set rs=conn.execute("select * From "& CurrentTableName &" where ParentPath like '%"&ParentPath & id&"%'")
			do while not rs.eof
				i=i+1
				mParentPath=replace(rs("ParentPath"),ParentPath,"")
				conn.execute("update "& CurrentTableName &" set depth=depth-"&depth&",rootid="&maxrootid&",ParentPath='"&mParentPath&"' where id="&rs("id"))
				rs.movenext
			loop
			rs.close
			set rs=nothing
		end if
		
		'更新其原来所属分类的分类数，排序相当于剪枝而不需考虑
		conn.execute("update "& CurrentTableName &" set child=child-1 where id="&iParentID)
		
	elseif iParentID>0 and rParentID>0 then    '如果是将一个分分类移动到其他分分类下
		'得到当前分类的下属子分类数
		ParentPath=ParentPath & ","
		set rs=conn.execute("select count(*) From "& CurrentTableName &" where ParentPath like '%"&ParentPath & id&"%'")
		ClassCount=rs(0)
		if isnull(ClassCount) then
			ClassCount=1
		end if
		rs.close
		set rs=nothing
		
		'获得目标分类的相关信息		
		set trs=conn.execute("select * From "& CurrentTableName &" where id="&rParentID)
		if trs("Child")>0 then		
			'得到与本分类同级的最后一个分类的OrderID
			set rsPrevOrderID=conn.execute("select Max(OrderID) From "& CurrentTableName &" where ParentID=" & trs("id"))
			PrevOrderID=rsPrevOrderID(0)
			'得到与本分类同级的最后一个分类的id
			sql="select id,NextID From "& CurrentTableName &" where ParentID=" & trs("id") & " and OrderID=" & PrevOrderID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			PrevID=rs(0)    '得到新的PrevID
			rs(1)=id     '更新上一个分类的NextID的值
			rs.update
			rs.close
			set rs=nothing
			
			'得到同一父分类但比本分类级数大的子分类的最大OrderID，如果比前一个值大，则改用这个值。
			set rsPrevOrderID=conn.execute("select Max(OrderID) From "& CurrentTableName &" where ParentPath like '" & trs("ParentPath") & "," & trs("id") & ",%'")
			if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
				if not IsNull(rsPrevOrderID(0))  then
			 		if rsPrevOrderID(0)>PrevOrderID then
						PrevOrderID=rsPrevOrderID(0)
					end if
				end if
			end if
		else
			PrevID=0
			PrevOrderID=trs("OrderID")
		end if
		
		'在获得移动过来的分类数后更新排序在指定分类之后的分类排序数据
		conn.execute("update "& CurrentTableName &" set OrderID=OrderID+" & ClassCount & "+1 where rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID)
		
		'更新当前分类数据
		conn.execute("update "& CurrentTableName &" set depth="&trs("depth")&"+1,OrderID="&PrevOrderID&"+1,rootid="&trs("rootid")&",ParentID="&rParentID&",ParentPath='" & trs("ParentPath") & "," & trs("id") & "',PrevID=" & PrevID & ",NextID=0 where id="&id)
		
		'如果有子分类则更新子分类数据，深度为原来的相对深度加上当前所属分类的深度
		set rs=conn.execute("select * From "& CurrentTableName &" where ParentPath like '%"&ParentPath&id&"%' order by OrderID")
		i=1
		do while not rs.eof
			i=i+1
			iParentPath=trs("ParentPath") & "," & trs("id") & "," & replace(rs("ParentPath"),ParentPath,"")
			conn.execute("update "& CurrentTableName &" set depth=depth-"&depth&"+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&iParentPath&"' where id="&rs("id"))
			rs.movenext
		loop
		rs.close
		set rs=nothing
		trs.close
		set trs=nothing
		
		'更新所指向的上级分类的子分类数
		conn.execute("update "& CurrentTableName &" set child=child+1 where id="&rParentID)
		
		'更新其原父类的子分类数			
		conn.execute("update "& CurrentTableName &" set child=child-1 where id="&iParentID)
	else    '如果原来是一级分类改成其他分类的下属分类
		'得到移动的分类总数
		set rs=conn.execute("select count(*) From "& CurrentTableName &" where rootid="&rootid)
		ClassCount=rs(0)
		rs.close
		set rs=nothing
		
		'获得目标分类的相关信息		
		set trs=conn.execute("select * From "& CurrentTableName &" where id="&rParentID)
		if trs("Child")>0 then		
			'得到与本分类同级的最后一个分类的OrderID
			set rsPrevOrderID=conn.execute("select Max(OrderID) From "& CurrentTableName &" where ParentID=" & trs("id"))
			PrevOrderID=rsPrevOrderID(0)
			sql="select id,NextID From "& CurrentTableName &" where ParentID=" & trs("id") & " and OrderID=" & PrevOrderID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			PrevID=rs(0)
			rs(1)=id
			rs.update
			set rs=nothing
			
			'得到同一父分类但比本分类级数大的子分类的最大OrderID，如果比前一个值大，则改用这个值。
			set rsPrevOrderID=conn.execute("select Max(OrderID) From "& CurrentTableName &" where ParentPath like '" & trs("ParentPath") & "," & trs("id") & ",%'")
			if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
				if not IsNull(rsPrevOrderID(0))  then
			 		if rsPrevOrderID(0)>PrevOrderID then
						PrevOrderID=rsPrevOrderID(0)
					end if
				end if
			end if
		else
			PrevID=0
			PrevOrderID=trs("OrderID")
		end if
	
		'在获得移动过来的分类数后更新排序在指定分类之后的分类排序数据
		conn.execute("update "& CurrentTableName &" set OrderID=OrderID+" & ClassCount &"+1 where rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID)
		
		conn.execute("update "& CurrentTableName &" set PrevID=" & PrevID & ",NextID=0 where id=" & id)
		set rs=conn.execute("select * From "& CurrentTableName &" where rootid="&rootid&" order by OrderID")
		i=0
		do while not rs.eof
			i=i+1
			if rs("parentid")=0 then
				ParentPath=trs("ParentPath") & "," & trs("id")
				conn.execute("update "& CurrentTableName &" set depth=depth+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&ParentPath&"',parentid="&rParentID&" where id="&rs("id"))
			else
				ParentPath=trs("ParentPath") & "," & trs("id") & "," & replace(rs("ParentPath"),"0,","")
				conn.execute("update "& CurrentTableName &" set depth=depth+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&ParentPath&"' where id="&rs("id"))
			end if
			rs.movenext
		loop
		rs.close
		set rs=nothing
		trs.close
		set trs=nothing
		'更新所指向的上级分类分类数		
		conn.execute("update "& CurrentTableName &" set child=child+1 where id="&rParentID)

	end if
  end if
	
  'call CloseConn()
  Response.Redirect CurrentPageNow
end sub

sub UpOrder()
	dim id,sqlOrder,rsOrder,MoveNum,cRootID,tRootID,i,rs,PrevID,NextID
	id=CokeShow.filtRequest(Request("id"))
	cRootID=CokeShow.filtRequest(Request("cRootID"))
	MoveNum=CokeShow.filtRequest(Request("MoveNum"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		id=CokeShow.CokeClng(id)
	end if
	if cRootID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		cRootID=CokeShow.CokeCint(cRootID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		MoveNum=CokeShow.CokeCint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要提升的数字！</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'得到本分类的PrevID,NextID
	set rs=conn.execute("select PrevID,NextID From "& CurrentTableName &" where id=" & id)
	PrevID=rs(0)
	NextID=rs(1)
	rs.close
	set rs=nothing
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update "& CurrentTableName &" set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "& CurrentTableName &" set PrevID=" & PrevID & " where id=" & NextID
	end if

	dim mrs,MaxRootID
	set mrs=conn.execute("select max(rootid) From "& CurrentTableName &"")
	MaxRootID=mrs(0)+1
	'先将当前分类移至最后，包括子分类
	conn.execute("update "& CurrentTableName &" set RootID=" & MaxRootID & " where RootID=" & cRootID)
	
	'然后将位于当前分类以上的分类的RootID依次加一，范围为要提升的数字
	sqlOrder="select * From "& CurrentTableName &" where ParentID=0 and RootID<" & cRootID & " order by RootID desc"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '如果当前分类已经在最上面，则无需移动
	end if
	i=1
	do while not rsOrder.eof
		tRootID=rsOrder("RootID")       '得到要提升位置的RootID，包括子分类
		i=i+1
		if i>MoveNum then
			rsOrder("PrevID")=id
			rsOrder.update
			conn.execute("update "& CurrentTableName &" set NextID=" & rsOrder("id") & " where id=" & id)
			conn.execute("update "& CurrentTableName &" set RootID=RootID+1 where RootID=" & tRootID)
			exit do
		end if
		conn.execute("update "& CurrentTableName &" set RootID=RootID+1 where RootID=" & tRootID)
		rsOrder.movenext
	loop
	rsOrder.movenext
	if rsOrder.eof then
		conn.execute("update "& CurrentTableName &" set PrevID=0 where id=" & id)
	else
		rsOrder("NextID")=id
		rsOrder.update
		conn.execute("update "& CurrentTableName &" set PrevID=" & rsOrder("id") & " where id=" & id)
	end if	
	rsOrder.close
	set rsOrder=nothing
	
	'然后再将当前分类从最后移到相应位置，包括子分类
	conn.execute("update "& CurrentTableName &" set RootID=" & tRootID & " where RootID=" & MaxRootID)
	'call CloseConn()
	response.Redirect "?Action=Order"
end sub

sub DownOrder()
	dim id,sqlOrder,rsOrder,MoveNum,cRootID,tRootID,i,rs,PrevID,NextID
	id=CokeShow.filtRequest(Request("id"))
	cRootID=CokeShow.filtRequest(Request("cRootID"))
	MoveNum=CokeShow.filtRequest(Request("MoveNum"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		id=CokeShow.CokeClng(id)
	end if
	if cRootID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		cRootID=CokeShow.CokeCint(cRootID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		MoveNum=CokeShow.CokeCint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要提升的数字！</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'得到本分类的PrevID,NextID
	set rs=conn.execute("select PrevID,NextID From "& CurrentTableName &" where id=" & id)
	PrevID=rs(0)
	NextID=rs(1)
	rs.close
	set rs=nothing
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update "& CurrentTableName &" set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "& CurrentTableName &" set PrevID=" & PrevID & " where id=" & NextID
	end if

	dim mrs,MaxRootID
	set mrs=conn.execute("select max(rootid) From "& CurrentTableName &"")
	MaxRootID=mrs(0)+1
	'先将当前分类移至最后，包括子分类
	conn.execute("update "& CurrentTableName &" set RootID=" & MaxRootID & " where RootID=" & cRootID)
	
	'然后将位于当前分类以下的分类的RootID依次减一，范围为要下降的数字
	sqlOrder="select * From "& CurrentTableName &" where ParentID=0 and RootID>" & cRootID & " order by RootID"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,2,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '如果当前分类已经在最下面，则无需移动
	end if
	i=1
	do while not rsOrder.eof
		tRootID=rsOrder("RootID")       '得到要提升位置的RootID，包括子分类
		
		i=i+1
		if i>MoveNum then
			rsOrder("NextID")=id
			rsOrder.update
			conn.execute("update "& CurrentTableName &" set PrevID=" & rsOrder("id") & " where id=" & id)
			conn.execute("update "& CurrentTableName &" set RootID=RootID-1 where RootID=" & tRootID)
			exit do
		end if
		conn.execute("update "& CurrentTableName &" set RootID=RootID-1 where RootID=" & tRootID)
		rsOrder.movenext
	loop
	rsOrder.movenext
	if rsOrder.eof then
		conn.execute("update "& CurrentTableName &" set NextID=0 where id=" & id)
	else
		rsOrder("PrevID")=id
		rsOrder.update
		conn.execute("update "& CurrentTableName &" set NextID=" & rsOrder("id") & " where id=" & id)
	end if	
	rsOrder.close
	set rsOrder=nothing
	
	'然后再将当前分类从最后移到相应位置，包括子分类
	conn.execute("update "& CurrentTableName &" set RootID=" & tRootID & " where RootID=" & MaxRootID)
	'call CloseConn()
	response.Redirect "?Action=Order"
end sub

sub UpOrderN()
	dim sqlOrder,rsOrder,MoveNum,id,i
	dim ParentID,OrderID,ParentPath,Child,PrevID,NextID
	id=CokeShow.filtRequest(Request("id"))
	MoveNum=CokeShow.filtRequest(Request("MoveNum"))
	if id="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		id=CokeShow.CokeClng(id)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		MoveNum=CokeShow.CokeCint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要提升的数字！</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	dim sql,rs,oldorders,ii,trs,tOrderID
	'要移动的分类信息
	set rs=conn.execute("select ParentID,OrderID,ParentPath,child,PrevID,NextID From "& CurrentTableName &" where id="&id)
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & id
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.close
	set rs=nothing
	if child>0 then
		set rs=conn.execute("select count(*) From "& CurrentTableName &" where ParentPath like '%"&ParentPath&"%'")
		oldorders=rs(0)
		rs.close
		set rs=nothing
	else
		oldorders=0
	end if
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update "& CurrentTableName &" set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "& CurrentTableName &" set PrevID=" & PrevID & " where id=" & NextID
	end if
	
	'和该分类同级且排序在其之上的分类------更新其排序，范围为要提升的数字
	sql="select id,OrderID,child,ParentPath,PrevID,NextID From "& CurrentTableName &" where ParentID="&ParentID&" and OrderID<"&OrderID&" order by OrderID desc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	i=1
	do while not rs.eof
		tOrderID=rs(1)
		
		if rs(2)>0 then
			ii=i+1
			set trs=conn.execute("select id,OrderID From "& CurrentTableName &" where ParentPath like '%"&rs(3)&","&rs(0)&"%' order by OrderID")
			if not (trs.eof and trs.bof) then
				do while not trs.eof
					conn.execute("update "& CurrentTableName &" set OrderID="&tOrderID+oldorders+ii&" where id="&trs(0))
					ii=ii+1
					trs.movenext
				loop
			end if
			trs.close
			set trs=nothing
		end if
		i=i+1
		if i>MoveNum then
			rs(4)=id
			rs.update
			conn.execute("update "& CurrentTableName &" set NextID=" & rs(0) & " where id=" & id)
			conn.execute("update "& CurrentTableName &" set OrderID="&tOrderID+oldorders+i-1&" where id="&rs(0))		
			exit do
		end if
		conn.execute("update "& CurrentTableName &" set OrderID="&tOrderID+oldorders+i-1&" where id="&rs(0))
		rs.movenext
	loop
	if not rs.eof then
	rs.movenext
	end if
	if rs.eof then
		conn.execute("update "& CurrentTableName &" set PrevID=0 where id=" & id)
	else
		rs(5)=id
		rs.update
		conn.execute("update "& CurrentTableName &" set PrevID=" & rs(0) & " where id=" & id)
	end if	
	rs.close
	set rs=nothing
	
	'更新所要排序的分类的序号
	conn.execute("update "& CurrentTableName &" set OrderID="&tOrderID&" where id="&id)
	'如果有下属分类，则更新其下属分类排序
	if child>0 then
		i=1
		set rs=conn.execute("select id From "& CurrentTableName &" where ParentPath like '%"&ParentPath&"%' order by OrderID")
		do while not rs.eof
			conn.execute("update "& CurrentTableName &" set OrderID="&tOrderID+i&" where id="&rs(0))
			i=i+1
			rs.movenext
		loop
		rs.close
		set rs=nothing
	end if
	'call CloseConn()
	response.Redirect "?Action=OrderN"
end sub

sub DownOrderN()
	dim sqlOrder,rsOrder,MoveNum,id,i
	dim ParentID,OrderID,ParentPath,Child,PrevID,NextID
	id=CokeShow.filtRequest(Request("id"))
	MoveNum=CokeShow.filtRequest(Request("MoveNum"))
	if id="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
		exit sub
	else
		id=CokeShow.CokeCint(id)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
		exit sub
	else
		MoveNum=CokeShow.CokeCint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要下降的数字！</li>"
			exit sub
		end if
	end if

	dim sql,rs,oldorders,ii,trs,tOrderID
	'要移动的分类信息
	set rs=conn.execute("select ParentID,OrderID,ParentPath,child,PrevID,NextID From "& CurrentTableName &" where id="&id)
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & id
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.close
	set rs=nothing

	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update "& CurrentTableName &" set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "& CurrentTableName &" set PrevID=" & PrevID & " where id=" & NextID
	end if
	
	'和该分类同级且排序在其之下的分类------更新其排序，范围为要下降的数字
	sql="select id,OrderID,child,ParentPath,PrevID,NextID From "& CurrentTableName &" where ParentID="&ParentID&" and OrderID>"&OrderID&" order by OrderID"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	i=0      '同级分类
	ii=0     '同级分类和子分类
	do while not rs.eof
		'conn.execute("update "& CurrentTableName &" set OrderID="&OrderID+ii&" where id="&rs(0))
		if rs(2)>0 then
			set trs=conn.execute("select id,OrderID From "& CurrentTableName &" where ParentPath like '%"&rs(3)&","&rs(0)&"%' order by OrderID")
			if not (trs.eof and trs.bof) then
				do while not trs.eof
					ii=ii+1
					conn.execute("update "& CurrentTableName &" set OrderID="&OrderID+ii&" where id="&trs(0))
					trs.movenext
				loop
			end if
			trs.close
			set trs=nothing
		end if
		ii=ii+1
		i=i+1
		if i>=MoveNum then
			rs(5)=id
			rs.update
			conn.execute("update "& CurrentTableName &" set PrevID=" & rs(0) & " where id=" & id)
			conn.execute("update "& CurrentTableName &" set OrderID="&OrderID+ii-1&" where id="&rs(0))		
			exit do
		end if
		conn.execute("update "& CurrentTableName &" set OrderID="&OrderID+ii-1&" where id="&rs(0))
		rs.movenext
	loop
	rs.movenext
	if rs.eof then
		conn.execute("update "& CurrentTableName &" set NextID=0 where id=" & id)
	else
		rs(4)=id
		rs.update
		conn.execute("update "& CurrentTableName &" set NextID=" & rs(0) & " where id=" & id)
	end if	
	rs.close
	set rs=nothing
	
	'更新所要排序的分类的序号
	conn.execute("update "& CurrentTableName &" set OrderID="&OrderID+ii&" where id="&id)
	'如果有下属分类，则更新其下属分类排序
	if child>0 then
		i=1
		set rs=conn.execute("select id From "& CurrentTableName &" where ParentPath like '%"&ParentPath&"%' order by OrderID")
		do while not rs.eof
			conn.execute("update "& CurrentTableName &" set OrderID="&OrderID+ii+i&" where id="&rs(0))
			i=i+1
			rs.movenext
		loop
		rs.close
		set rs=nothing
	end if
	'call CloseConn()
	response.Redirect "?Action=OrderN"
end sub

sub SaveReset()
	dim i,sql,rs,SuccessMsg,iCount,PrevID,NextID
	sql="select id From "& CurrentTableName &" order by RootID,OrderID"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	iCount=rs.recordcount
	i=1
	PrevID=0
	do while not rs.eof
		rs.movenext
		if rs.eof then
			NextID=0
		else
			NextID=rs(0)
		end if
		rs.moveprevious
		conn.execute("update "& CurrentTableName &" set RootID=" & i & ",OrderID=0,ParentID=0,Child=0,ParentPath='0',Depth=0,PrevID=" & PrevID & ",NextID=" & NextID & " where id=" & rs(0))
		PrevID=rs(0)
		i=i+1
		rs.movenext
	loop
	rs.close
	set rs=nothing	
	
	response.Write "复位成功！请返回<a href='"& CurrentPageNow &"'>"& UnitName &"管理首页</a>做"& UnitName &"的归属设置。"
end sub

sub SaveUnite()
	dim id,Targetid,ParentPath,iParentPath,Depth,iParentID,Child,PrevID,NextID
	dim rs,trs,i
	id=CokeShow.filtRequest(Request("id"))
	Targetid=CokeShow.filtRequest(Request("Targetid"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要合并的"& UnitName &"！</li>"
	else
		id=CokeShow.CokeClng(id)
	end if
	if Targetid="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定目标"& UnitName &"！</li>"
	else
		Targetid=CokeShow.CokeClng(Targetid)
	end if
	if id=Targetid then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请不要在相同"& UnitName &"内进行操作</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	'判断目标分类是否有子分类，如果有，则报错。
	set rs=conn.execute("select Child From "& CurrentTableName &" where id=" & Targetid)
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>目标"& UnitName &"不存在，可能已经被删除！</li>"
	else
		if rs(0)>0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>目标"& UnitName &"中含有子"& UnitName &"，不能合并！</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'得到当前分类信息
	set rs=conn.execute("select id,ParentID,ParentPath,PrevID,NextID,Depth From "& CurrentTableName &" where id="&id)
	iParentID=rs(1)
	Depth=rs(5)
	if iParentID=0 then
		ParentPath=rs(0)
	else
		ParentPath=rs(2) & "," & rs(0)
	end if
	iParentPath=rs(0)
	PrevID=rs(3)
	NextID=rs(4)
	
	'判断是否是合并到其下属分类中
	set rs=conn.execute("select id From "& CurrentTableName &" where id="&Targetid&" and ParentPath like '"&ParentPath&"%'")
	if not (rs.eof and rs.bof) then
		response.Write "<br><li>不能将一个"& UnitName &"合并到其下属子"& UnitName &"中</li>"
		exit sub
	end if
	
	'得到当前分类的下属分类ID
	set rs=conn.execute("select id From "& CurrentTableName &" where ParentPath like '"&ParentPath&"%'")
	i=0
	if not (rs.eof and rs.bof) then
		do while not rs.eof
			iParentPath=iParentPath & "," & rs(0)
			i=i+1
			rs.movenext
		loop
	end if
	if i>0 then
		ParentPath=iParentPath
	else
		ParentPath=id
	end if
	
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update "& CurrentTableName &" set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update "& CurrentTableName &" set PrevID=" & PrevID & " where id=" & NextID
	end if
	
	'手动操作WL！
	'更新user所属分类
	'//		CONN.Execute("UPDATE [meiqee_user] SET user_classid="& Targetid &" WHERE user_classid IN ("& ParentPath &")")
	'手动操作WL！
	
	'删除被合并分类及其下属分类
	conn.execute("delete From "& CurrentTableName &" where id in ("&ParentPath&")")
	
	
	'更新其原来所属分类的子分类数，排序相当于剪枝而不需考虑
	if Depth>0 then
		conn.execute("update "& CurrentTableName &" set Child=Child-1 where id="&iParentID)
	end if
	
	response.Write ""& UnitName &"合并成功！已经将被合并"& UnitName &"及其下属子"& UnitName &"的所有数据转入目标"& UnitName &"中。<br><br>同时删除了被合并的"& UnitName &"及其子"& UnitName &"。"
	set rs=nothing
	set trs=nothing
end sub

%>


