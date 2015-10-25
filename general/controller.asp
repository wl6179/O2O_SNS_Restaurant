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

<%
'变量定义区.
'(用来存储对象的变量，用全大写!)
Const maxPerPage=15							'当前模块分页设置.
Dim CurrentPageNow,TitleName,UnitName
CurrentPageNow 	= "controller.asp"			'当前页.
TitleName 		= "网站资料设置"				'此模块管理页的名字.
UnitName 		= "网站资料"					'此模块涉及记录的元素名.
'自定义设置.
'本地设置.
Dim CurrentTableName
CurrentTableName 	= "[CXBG_controller]"		'此模块涉及的[表]名.
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
If Action<>"SaveModify" Then			Action="Modify"

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
		dojo.require("dijit.form.Textarea");
		//上传图片.
		dojo.require("dojo.io.iframe");
//		dojo.require("dijit.form.FilteringSelect");
		
		
	</script>
	
	<script type="text/javascript">
		//Table偶数行变色函数.
		dojo.addOnLoad(function() {
			stripeTables("listGo");
		});
	</script>
	
	<script type="text/javascript">
	//确认删除函数，仅针对Table的批量删除操作
	function save() {
		var conf = confirm("确定要保存此<% =UnitName %>吗？");
		if(conf == true){
			dojo.byId("form1").submit();
		}
	}
	</script>
	
	
	<!--#include file="_screen_JS.asp"-->
	
	
	
	<style type="text/css">
		.title {
			color: #000000;
			font-weight: bold;
		}
	</style>
</head>
<body class="<% =classname_dj_ThemesCss %>">


<!--main-->
	<!-- Begin mainleft-->
	<div class="newsContainer">
		
		<div class="news1">
			
			<h3>相关管理项目</h3>
				
				<p>
					<!--#include file="menu01_1.asp"-->
				</p>
				
			<h3>当前操作</h3>
				
				<ul>
					<!--<li><a href="http://www.iw3c2.org/">&#187;IW<sup>3</sup>C<sup>2</sup></a></li>-->
					
					<li><a href="#" onClick="save();">&#187;保存修改操作</a></li>
				</ul>
				
		</div><!-- End news1-->
		
		
	
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
		'Response.Write Action
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


Sub Modify()
	strGuide=strGuide & "修改"& UnitName
	
	Dim intID
	intID=CokeShow.filtRequest(Request("id"))
	intID=1
	'处理id传值
	If intID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		Exit Sub
	Else
		intID=CokeShow.CokeClng(intID)
	End If
	
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT * FROM "& CurrentTableName &" WHERE id="& intID
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
				
				
				
				<form action="<% = CurrentPageNow %>" method="post" name="form1" id="form1"
				dojoType="dijit.form.Form"
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
						<td colspan="4" style="color:#999999;">
							<div class="title">网站主要设置</div>
						</td>
					  </tr>
					  <tr>
						<td style="text-align: right;">网站名称</td>
						<td style="text-align: left;">
						
						<input type="text" id="site_name" name="site_name" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="true"
						propercase="false"
						promptMessage="请输入网站名称."
						invalidMessage="网站名称必填！网站名称长度须在1~50之内！"
						trim="true"
						regExp=".{1,50}"
						 value="<% =RS("site_name") %>"
                         style="width:300px;"
						 class="input_tell"
						/>
						</td>
						
						<td style="text-align: right;">网站标题</td>
						<td style="text-align: left;">
						<input type="text" id="site_title" name="site_title" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="true"
						propercase="false"
						promptMessage="请输入显示在您网站首页浏览器上边显示的标题文字."
						invalidMessage="网站标题必填！网站标题长度须在1~100之内！"
						trim="true"
						regExp=".{1,100}"
						 value="<% =RS("site_title") %>"
						 style="width:300px;"
                         class="input_tell"
						/>
						</td>
					  </tr>
					  
					  
					  <tr>
						<td style="text-align: right;">网站域名</td>
						<td style="text-align: left;">
						
						<input type="text" id="site_domain" name="site_domain" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="true"
						propercase="false"
						promptMessage="请输入网站域名."
						invalidMessage="网站域名必填！请填写正确的域名格式，并且网站域名长度须在6~50之内！"
						trim="true"
						regExp="^[a-zA-z]+://(\w+(-\w+)*)(\.(\w+(-\w+)*))*(\?\S*)?$"
						 value="<% =RS("site_domain") %>"
                         style="width:300px;"
						 class="input_tell"
						/>
                        <!--^(http|ftp|https):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&:/~\+#]*[\w\-\@?^=%&/~\+#])?$-->
						</td>
						
						<td style="text-align: right;">网站关键字</td>
						<td style="text-align: left;">
						<input type="text" id="site_keyword" name="site_keyword" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="true"
						propercase="true"
						promptMessage="请输入网站关键字，有助于搜索引擎收录."
						invalidMessage="网站关键字必填！网站关键字长度须在6~100之内！"
						trim="true"
						regExp=".{6,100}"
						 value="<% =RS("site_keyword") %>"
						 style="width:300px;"
                         class="input_tell"
						/>
						</td>
					  </tr>
					  
					  
					  
					  <tr>
						<td style="text-align: right;">官方电子邮件</td>
						<td style="text-align: left;">
						
						<input type="text" id="site_email" name="site_email" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="true"
						propercase="false"
						promptMessage="请输入官方电子邮件."
						invalidMessage="官方电子邮件必填！请填写正确的Email电子邮件格式，例如：yourname6179@microsoft.com，且网站关键字长度须在6~100之内！"
						trim="true"
						regExp="^[0-9a-zA-Z]+([0-9a-zA-Z]|_|\.|-)+[0-9a-zA-Z]+@(([0-9a-zA-Z]+\.)|([0-9a-zA-Z]+-))+[0-9a-zA-Z]+$"
						 value="<% =RS("site_email") %>"
                         style="width:300px;"
						 class="input_tell"
						/>
						</td>
						
						<td style="text-align: right;">网站描述文字</td>
						<td style="text-align: left;">
						<textarea id="site_description" name="site_description"
						dojoType="dijit.form.Textarea"
						 style="width: 300px;"
                         
						><% =RS("site_description") %></textarea>
						
						</td>
					  </tr>
					  
                      <tr>
						<td style="text-align: right;">餐厅地址</td>
						<td style="text-align: left;">
						
						<input type="text" id="Address" name="Address" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						promptMessage="请输入餐厅地址."
						invalidMessage=""
						trim="true"
						regExp=".{0,50}"
						 value="<% =RS("Address") %>"
                         style="width:300px;"
						 class="input_tell"
						/>
						</td>
						
						<td style="text-align: right;"></td>
						<td style="text-align: left;">
						
						
						</td>
					  </tr>
					  
					  <tr>
						<td colspan="4" style="color:#999999;">
							
						</td>
					  </tr>
					  <tr>
						<td colspan="4" style="color:#999999;">
							<div class="title">细节营销设置</div>
						</td>
					  </tr>
					  <tr style="display:none;">
						<td style="text-align: right;">显示注册用户数</td>
						<td style="text-align: left;">
						
						<input type="text" id="show_userlist_num" name="show_userlist_num" size="20"
						dojoType="dijit.form.NumberSpinner"
						required="false"
						propercase="false"
						
						
						trim="true"
						 value="<% =RS("show_userlist_num") %>"
                         style="width:100px;"
						
						/>
						</td>
						
						<td style="text-align: right;">进度显示II</td>
						<td style="text-align: left;">
						
						</td>
					  </tr>
                      
                      <tr>
						<td style="text-align: right;">点餐牌过度广告图</td>
						<td style="text-align: left;">
							
							<span style="color:#999999;"> 尺寸要求建议:宽130*高170 (长方形)，如果不需要设置为新品则不需要上传此图片！</span>
                                        
                            <br />
                            
                            <a href="<% =RS("Loading_Photo") %>" target="_blank" id="href_photoLoading">
                            <img style="border:5px #999999 solid;" src="<% If RS("Loading_Photo")<>"" Then Response.Write RS("Loading_Photo") Else Response.Write "/images/NoPic.png" %>" width="130" height="170" onerror='this.src="/images/NoPic.png"'
                             id="src_photoLoading"
                            /></a>
                            
                            <button type="button"
                            dojoType="dijit.form.Button"
                            onclick="ShowDialog('<img src=/images/up.gif />上传图片','../upload/index.asp?Action=Add&controlStr=Loading','width:300px;height:200px;');"
                            >
                            &nbsp;点击开始上传&nbsp;
                            </button>
                            
                            <div id="div_photo">
                            <input type="text" id="value_photoLoading" name="Loading_Photo"
                            dojoType="dijit.form.ValidationTextBox"
                            required="false"
                            promptMessage="请上传图片."
                            invalidMessage="必须在250长度之内！例如：/uploadimages/cokeshow.com.cn20097131746850193.png"
                            trim="true"
                            lowercase="false"
                            
                            regExp=".{0,250}"
                             value="<% =RS("Loading_Photo") %>"
                             style="width:500px;"
                             class="input_tell"
                            />
                            </div>
						</td>
						
						<td style="text-align: right;">点餐牌过度广告文字<br />(Loading......显示)</td>
						<td style="text-align: left;">
						<input type="text" id="loadingString" name="loadingString" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						promptMessage="进度显示，在网络较慢读取页面时显示的字."
						invalidMessage="进度显示长度须在1~50之内！"
						trim="true"
						regExp=".{1,50}"
						 value="<% =RS("loadingString") %>"
                         style="width:300px;"
						class="input_tell"
						/>
						</td>
					  </tr>
                      
                      <!--首页最新促销主题图片-->
                      <tr>
						<td style="text-align: right;">首页促销主题图</td>
						<td style="text-align: left;">
							
							<span style="color:#999999;"> 尺寸要求建议:宽206*高103 (长方形)！</span>
                                        
                            <br />
                            
                            <a href="<% =RS("SalesPromotion_Photo") %>" target="_blank" id="href_photoSalesPromotion">
                            <img style="border:5px #999999 solid;" src="<% If RS("SalesPromotion_Photo")<>"" Then Response.Write RS("SalesPromotion_Photo") Else Response.Write "/images/NoPic.png" %>" width="206" height="103" onerror='this.src="/images/NoPic.png"'
                             id="src_photoSalesPromotion"
                            /></a>
                            
                            <button type="button"
                            dojoType="dijit.form.Button"
                            onclick="ShowDialog('<img src=/images/up.gif />上传图片','../upload/index.asp?Action=Add&controlStr=SalesPromotion','width:300px;height:200px;');"
                            >
                            &nbsp;点击开始上传&nbsp;
                            </button>
                            
                            <div id="div_photo">
                            <input type="text" id="value_photoSalesPromotion" name="SalesPromotion_Photo"
                            dojoType="dijit.form.ValidationTextBox"
                            required="false"
                            promptMessage="请上传图片."
                            invalidMessage="必须在250长度之内！例如：/uploadimages/cokeshow.com.cn20097131746850193.png"
                            trim="true"
                            lowercase="false"
                            
                            regExp=".{0,250}"
                             value="<% =RS("SalesPromotion_Photo") %>"
                             style="width:500px;"
                             class="input_tell"
                            />
                            </div>
						</td>
						
						<td style="text-align: right;">&nbsp;</td>
						<td style="text-align: left;">&nbsp;
						
						</td>
					  </tr>
					  <!--首页最新促销主题图片-->
					  
					  
                      
                      <!--首页对联广告图片设置-->
                      <tr>
						<td style="text-align: right;">是否启用对联广告（一般节庆日启用）</td>
						<td style="text-align: left;">
						
						启用<input type="radio" name="isAdContentImageShow" value="1" <% If RS("isAdContentImageShow")=1 Then Response.Write "checked" %> />
                        &nbsp;&nbsp;&nbsp;
                        停用<input type="radio" name="isAdContentImageShow" value="0" <% If RS("isAdContentImageShow")=0 Then Response.Write "checked" %> />
						</td>
						
						<td style="text-align: right;"></td>
						<td style="text-align: left;">
						
						
						</td>
					  </tr>
                      
                      <tr>
						<td style="text-align: right;">首页对联广告图（左图）</td>
						<td style="text-align: left;">
							
							<span style="color:#999999;"> 尺寸要求建议:宽100*高300 (竖长方形)！</span>
                                        
                            <br />
                            
                            <a href="<% =RS("AdContentImageLeft") %>" target="_blank" id="href_photoAdContentImageLeft">
                            <img style="border:5px #999999 solid;" src="<% If RS("AdContentImageLeft")<>"" Then Response.Write RS("AdContentImageLeft") Else Response.Write "/images/NoPic.png" %>" width="100" height="300" onerror='this.src="/images/NoPic.png"'
                             id="src_photoAdContentImageLeft"
                            /></a>
                            
                            <button type="button"
                            dojoType="dijit.form.Button"
                            onclick="ShowDialog('<img src=/images/up.gif />上传图片','../upload/index.asp?Action=Add&controlStr=AdContentImageLeft','width:300px;height:200px;');"
                            >
                            &nbsp;点击开始上传&nbsp;
                            </button>
                            
                            <div id="div_photo">
                            <input type="text" id="value_photoAdContentImageLeft" name="AdContentImageLeft"
                            dojoType="dijit.form.ValidationTextBox"
                            required="false"
                            promptMessage="请上传图片."
                            invalidMessage="必须在250长度之内！例如：/uploadimages/cokeshow.com.cn20097131746850193.png"
                            trim="true"
                            lowercase="false"
                            
                            regExp=".{0,250}"
                             value="<% =RS("AdContentImageLeft") %>"
                             style="width:500px;"
                             class="input_tell"
                            />
                            </div>
                            
                            <br />
                            <br />
                            
                            URL:
                            <input type="text" id="AdContentImageLeftURL" name="AdContentImageLeftURL"
                             dojoType="dijit.form.ValidationTextBox"
                             required="false"
                             promptMessage="请固定链接地址，格式应该为 http://www.chixinbugai.me/xxx001.html"
                             invalidMessage="必须在250长度之内！例如：http://www.chixinbugai.me/uploadimages/xxx001.html"
                             trim="true"
                             lowercase="false"
                             
                             regExp=".{0,250}"
                             value="<% =RS("AdContentImageLeftURL") %>"
                             style="width:500px;"
                             class="input_tell"
                             />
						</td>
						
						<td style="text-align: right;">首页对联广告图（右图）</td>
						<td style="text-align: left;">
							
							<span style="color:#999999;"> 尺寸要求建议:宽100*高300 (竖长方形)！</span>
                                        
                            <br />
                            
                            <a href="<% =RS("AdContentImageRight") %>" target="_blank" id="href_photoAdContentImageRight">
                            <img style="border:5px #999999 solid;" src="<% If RS("AdContentImageRight")<>"" Then Response.Write RS("AdContentImageRight") Else Response.Write "/images/NoPic.png" %>" width="100" height="300" onerror='this.src="/images/NoPic.png"'
                             id="src_photoAdContentImageRight"
                            /></a>
                            
                            <button type="button"
                            dojoType="dijit.form.Button"
                            onclick="ShowDialog('<img src=/images/up.gif />上传图片','../upload/index.asp?Action=Add&controlStr=AdContentImageRight','width:300px;height:200px;');"
                            >
                            &nbsp;点击开始上传&nbsp;
                            </button>
                            
                            <div id="div_photo">
                            <input type="text" id="value_photoAdContentImageRight" name="AdContentImageRight"
                            dojoType="dijit.form.ValidationTextBox"
                            required="false"
                            promptMessage="请上传图片."
                            invalidMessage="必须在250长度之内！例如：/uploadimages/cokeshow.com.cn20097131746850193.png"
                            trim="true"
                            lowercase="false"
                            
                            regExp=".{0,250}"
                             value="<% =RS("AdContentImageRight") %>"
                             style="width:500px;"
                             class="input_tell"
                            />
                            </div>
                            
                            
                            <br />
                            <br />
                            
                            URL:
                            <input type="text" id="AdContentImageRightURL" name="AdContentImageRightURL"
                             dojoType="dijit.form.ValidationTextBox"
                             required="false"
                             promptMessage="请固定链接地址，格式应该为 http://www.chixinbugai.me/xxx001.html"
                             invalidMessage="必须在250长度之内！例如：http://www.chixinbugai.me/uploadimages/xxx001.html"
                             trim="true"
                             lowercase="false"
                             
                             regExp=".{0,250}"
                             value="<% =RS("AdContentImageRightURL") %>"
                             style="width:500px;"
                             class="input_tell"
                             />
						</td>
					  </tr>
					  <!--首页对联广告图片设置-->
					  
					  
					  <tr style="display:none;">
						<td colspan="4" style="color:#999999;">
							
						</td>
					  </tr>
					  <tr style="display:none;">
						<td colspan="4" style="color:#999999;">
							<div class="title">网站上传文件控制</div>
						</td>
					  </tr>
					  <tr style="display:none;">
						<td colspan="4" style="color:#999999;">
							普通用户上传控制
						</td>
					  </tr>
					  <tr style="display:none;">
						<td style="text-align: right;">文件类型限制</td>
						<td style="text-align: left;">
						
						<input type="text" id="upfiles_user_type" name="upfiles_user_type" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						promptMessage="请输入普通用户上传文件类型限制."
						
						trim="true"
						 value="<% =RS("upfiles_user_type") %>"
						 
                         class="input_tell"
						/>
						</td>
						
						<td style="text-align: right;">单个文件大小限制(k)</td>
						<td style="text-align: left;">
						<input type="text" id="upfiles_user_onesize" name="upfiles_user_onesize" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						promptMessage="请输入普通用户上传单个文件大小限制(k)."
						invalidMessage="普通用户上传单个文件大小限制必填！"
						trim="true"
						 value="<% =RS("upfiles_user_onesize") %>"
						class="input_tell"
						/>
						</td>
					  </tr>
					  
					  
					  <tr style="display:none;">
						<td style="text-align: right;">文件总大小限制(k)</td>
						<td style="text-align: left;">
						<input type="text" id="upfiles_user_maxsize" name="upfiles_user_maxsize" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						promptMessage="请输入普通用户上传文件总大小限制(k)."
						invalidMessage="普通用户上传文件总大小限制必填！"
						trim="true"
						 value="<% =RS("upfiles_user_maxsize") %>"
						class="input_tell"
						/>
						</td>
						
						<td style="text-align: right;"></td>
						<td style="text-align: left;">
						
						</td>
					  </tr>
					  
					  
					  <tr style="display:none;">
						<td colspan="4" style="color:#999999;">
							VIP用户上传控制
						</td>
					  </tr>
					  <tr style="display:none;">
						<td style="text-align: right;">文件类型限制</td>
						<td style="text-align: left;">
						
						<input type="text" id="upfiles_vipuser_type" name="upfiles_vipuser_type" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						promptMessage="请输入VIP用户上传文件类型限制."
						
						trim="true"
						 value="<% =RS("upfiles_vipuser_type") %>"
						 
                         class="input_tell"
						/>
						</td>
						
						<td style="text-align: right;">单个文件大小限制(k)</td>
						<td style="text-align: left;">
						<input type="text" id="upfiles_vipuser_onesize" name="upfiles_vipuser_onesize" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						promptMessage="请输入VIP用户上传单个文件大小限制(k)."
						invalidMessage="VIP用户上传单个文件大小限制必填！"
						trim="true"
						 value="<% =RS("upfiles_vipuser_onesize") %>"
						class="input_tell"
						/>
						</td>
					  </tr>
					  
					  
					  <tr style="display:none;">
						<td style="text-align: right;">文件总大小限制(k)</td>
						<td style="text-align: left;">
						<input type="text" id="upfiles_vipuser_maxsize" name="upfiles_vipuser_maxsize" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						promptMessage="请输入VIP用户上传文件总大小限制(k)."
						invalidMessage="VIP用户上传文件总大小限制必填！"
						trim="true"
						 value="<% =RS("upfiles_vipuser_maxsize") %>"
						class="input_tell"
						/>
						</td>
						
						<td style="text-align: right;"></td>
						<td style="text-align: left;">
						
						</td>
					  </tr>
					  
					  
					  <tr style="display:none;">
						<td colspan="4" style="color:#999999;">
							管理员上传控制
						</td>
					  </tr>
					  <tr style="display:none;">
						<td style="text-align: right;">上传文件类型限制</td>
						<td style="text-align: left;">
						
						<input type="text" id="upfiles_supervisor_type" name="upfiles_supervisor_type" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						promptMessage="请输入管理员用户上传文件类型限制."
						
						trim="true"
						 value="<% =RS("upfiles_supervisor_type") %>"
						 
                         class="input_tell"
						/>
						</td>
						
						<td style="text-align: right;">上传单个文件大小限制(k)</td>
						<td style="text-align: left;">
						<input type="text" id="upfiles_supervisor_onesize" name="upfiles_supervisor_onesize" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						promptMessage="请输入管理员用户上传单个文件大小限制(k)."
						invalidMessage="管理员用户上传单个文件大小限制必填！"
						trim="true"
						 value="<% =RS("upfiles_supervisor_onesize") %>"
						class="input_tell"
						/>
						</td>
					  </tr>
					  
					  
					  <tr style="display:none;">
						<td style="text-align: right;">上传文件总大小限制(k)</td>
						<td style="text-align: left;">
						<input type="text" id="upfiles_supervisor_maxsize" name="upfiles_supervisor_maxsize" size="20"
						dojoType="dijit.form.ValidationTextBox"
						required="false"
						propercase="false"
						promptMessage="请输入管理员用户上传文件总大小限制(k)."
						invalidMessage="管理员用户上传文件总大小限制必填！"
						trim="true"
						 value="<% =RS("upfiles_supervisor_maxsize") %>"
						class="input_tell"
						/>
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
						  value="<% =RS("ID") %>"
						  />
						  
						  
						      <button type="submit" id="submitbtn" 
							  dojoType="dijit.form.Button"
							   onclick="return confirm('确定要保存此<% =UnitName %>吗？');"
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



Sub SaveModify()
	Dim site_name,site_title,site_domain,site_keyword,site_description,site_email,show_userlist_num,loadingString,Loading_Photo
	Dim upfiles_user_type,upfiles_user_onesize,upfiles_user_maxsize,upfiles_vipuser_type,upfiles_vipuser_onesize,upfiles_vipuser_maxsize,upfiles_supervisor_type,upfiles_supervisor_onesize,upfiles_supervisor_maxsize
	Dim Address
	Dim SalesPromotion_Photo
	Dim isAdContentImageShow,AdContentImageLeft,AdContentImageRight
	Dim AdContentImageLeftURL,AdContentImageRightURL
	
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
	site_name	=CokeShow.filtRequest(Request("site_name"))
	site_title	=CokeShow.filtRequest(Request("site_title"))
	site_domain	=CokeShow.filtRequest(Request("site_domain"))
	site_keyword=CokeShow.filtRequest(Request("site_keyword"))
	site_description=CokeShow.filtRequest(Request("site_description"))
	site_email	=CokeShow.filtRequest(Request("site_email"))
	show_userlist_num	=CokeShow.filtRequest(Request("show_userlist_num"))
	loadingString	=CokeShow.filtRequestSimple(Request("loadingString"))
	Loading_Photo	=CokeShow.filtRequest(Request("Loading_Photo"))
	
	upfiles_user_type		=CokeShow.filtRequest(Request("upfiles_user_type"))
	upfiles_user_onesize	=CokeShow.filtRequest(Request("upfiles_user_onesize"))
	upfiles_user_maxsize	=CokeShow.filtRequest(Request("upfiles_user_maxsize"))
	upfiles_vipuser_type	=CokeShow.filtRequest(Request("upfiles_vipuser_type"))
	upfiles_vipuser_onesize	=CokeShow.filtRequest(Request("upfiles_vipuser_onesize"))
	upfiles_vipuser_maxsize	=CokeShow.filtRequest(Request("upfiles_vipuser_maxsize"))
	upfiles_supervisor_type	=CokeShow.filtRequest(Request("upfiles_supervisor_type"))
	upfiles_supervisor_onesize	=CokeShow.filtRequest(Request("upfiles_supervisor_onesize"))
	upfiles_supervisor_maxsize	=CokeShow.filtRequest(Request("upfiles_supervisor_maxsize"))
	
	Address					=CokeShow.filtRequest(Request("Address"))
	
	SalesPromotion_Photo	=CokeShow.filtRequest(Request("SalesPromotion_Photo"))
	
	isAdContentImageShow	=CokeShow.filtRequest(Request("isAdContentImageShow"))
	AdContentImageLeft	=CokeShow.filtRequest(Request("AdContentImageLeft"))
	AdContentImageRight	=CokeShow.filtRequest(Request("AdContentImageRight"))
	
	AdContentImageLeftURL	=CokeShow.filtRequest(Request("AdContentImageLeftURL"))
	AdContentImageRightURL	=CokeShow.filtRequest(Request("AdContentImageRightURL"))
	
	'验证
	If site_name="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>网站名称不能为空！</li>"
	Else
		If Len(site_name)>100 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>网站名称长度只能在100位之内！</li>"
		Else
			site_name=site_name
		End If
	End If
	
	If site_title="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>网站标题不能为空！</li>"
	Else
		If Len(site_title)>100 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>网站标题长度只能在100位之内！</li>"
		Else
			site_title=site_title
		End If
	End If
	
	If site_domain="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>网站域名不能为空！</li>"
	Else
		If Len(site_domain)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>网站域名长度只能在50位之内！</li>"
		Else
			site_domain=site_domain
		End If
	End If
	
	
	If site_keyword="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>网站关键字不能为空！</li>"
	Else
		If Len(site_keyword)>100 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>网站关键字长度只能在100位之内！</li>"
		Else
			site_keyword=site_keyword
		End If
	End If
	
	If site_description="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>网站描述文字不能为空！</li>"
	Else
		If Len(site_description)>255 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>网站描述文字长度只能在255位之内！</li>"
		Else
			site_description=site_description
		End If
	End If
	
	If Loading_Photo="" Then
		'FoundErr=True
		'ErrMsg=ErrMsg &"<br><li>点餐牌过度广告图不能为空！</li>"
		Loading_Photo=""
	Else
		If Len(Loading_Photo)>250 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>点餐牌过度广告图长度只能在250位之内！</li>"
		Else
			Loading_Photo=Loading_Photo
		End If
	End If
	
	If SalesPromotion_Photo="" Then
		'FoundErr=True
		'ErrMsg=ErrMsg &"<br><li>首页促销主题图不能为空！</li>"
		SalesPromotion_Photo=""
	Else
		If Len(SalesPromotion_Photo)>250 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>首页促销主题图长度只能在250位之内！</li>"
		Else
			SalesPromotion_Photo=SalesPromotion_Photo
		End If
	End If
	
	If Address="" Then
		'FoundErr=True
		'ErrMsg=ErrMsg &"<br><li>餐厅地址不能为空！</li>"
		Address=""
	Else
		If Len(Address)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>餐厅地址长度只能在50位之内！</li>"
		Else
			Address=Address
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
	sql="SELECT * FROM "& CurrentTableName &" WHERE id="& intID
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,1,3
	
	'拦截此记录的异常情况.
	If RS.Bof And RS.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的"& UnitName &"！</li>"
		Exit Sub
	End If
	
		'如果填写了密码，则进行修改.
		RS("site_name")		=site_name
		RS("site_title")	=site_title
		RS("site_domain")	=site_domain
		RS("site_keyword")	=site_keyword
		RS("site_description")	=site_description
		RS("site_email")	=site_email
		RS("show_userlist_num")	=show_userlist_num
		RS("loadingString")		=loadingString
		RS("Loading_Photo")		=Loading_Photo
		
		RS("upfiles_user_type")			=upfiles_user_type
		RS("upfiles_user_onesize")		=upfiles_user_onesize
		RS("upfiles_user_maxsize")		=upfiles_user_maxsize
		RS("upfiles_vipuser_type")		=upfiles_vipuser_type
		RS("upfiles_vipuser_onesize")		=upfiles_vipuser_onesize
		RS("upfiles_vipuser_maxsize")		=upfiles_vipuser_maxsize
		RS("upfiles_supervisor_type")		=upfiles_supervisor_type
		RS("upfiles_supervisor_onesize")	=upfiles_supervisor_onesize
		RS("upfiles_supervisor_maxsize")	=upfiles_supervisor_maxsize
		
		RS("Address")				=Address
		
		RS("SalesPromotion_Photo")	=SalesPromotion_Photo
		
		RS("isAdContentImageShow")	=isAdContentImageShow
		RS("AdContentImageLeft")	=AdContentImageLeft
		RS("AdContentImageRight")	=AdContentImageRight
		
		RS("AdContentImageLeftURL")	=AdContentImageLeftURL
		RS("AdContentImageRightURL")	=AdContentImageRightURL
		
		RS("modifydate")	=Now()
	
	RS.Update
	RS.Close
	Set RS=Nothing
	
'记入日志.
Call CokeShow.AddLog("编辑操作：成功编辑了ID为"& intID &"的"& UnitName &"设置", sql)
	
	CokeShow.ShowOK "修改"& UnitName &"成功!",""
End Sub

%>