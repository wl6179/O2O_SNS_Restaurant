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
CurrentPageNow 	= "tuijianpengyou.asp"			'当前页.
TitleName 		= "推荐朋友列表"				'此模块管理页的名字.
UnitName 		= "推荐朋友"					'此模块涉及记录的元素名.
'自定义设置.
'本地设置.
Dim CurrentTableName
CurrentTableName 	= "[CXBG_account_TuijianPengyou]"		'此模块涉及的[表]名.
%>



<%
Dim totalPut,totalPages,currentPage			'分页用的控制变量.
Dim RS, sql									'查询列表记录用的变量.
Dim FoundErr,ErrMsg							'控制错误流程用的控制变量.
Dim strFileName								'构建查询字符串用的控制变量.
Dim ExecuteSearch,Keyword,TypeSearch,Action	'构建查询字符串以及流程控制用的控制变量.
Dim strGuide		'导航文字.

'SearchNowResult
Dim SearchNowResult

currentPage		=CokeShow.filtRequest(Request("Page"))
ExecuteSearch	=CokeShow.filtRequest(Request("ExecuteSearch"))
Keyword			=CokeShow.filtRequest(Request("Keyword"))
TypeSearch		=CokeShow.filtRequest(Request("TypeSearch"))
Action			=CokeShow.filtRequest(Request("Action"))

'SearchNowResult
SearchNowResult			=CokeShow.filtRequest(Request("SearchNowResult"))

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


'SearchNowResult
If SearchNowResult<>"" Then
	strFileName=strFileName&"&SearchNowResult="& SearchNowResult
End If


'处理当前页码的控制变量，通过获取到的传值获取，默认为第一页1.
If currentPage<>"" Then
    If isNumeric(currentPage) Then currentPage=CokeShow.CokeClng(currentPage) Else currentPage=1
Else
	currentPage=1
End If




'高级查询专区 Begin
'1.定义构造sql的变量.
Dim ProductName_SQL,ProductNo_SQL,product_class_id_SQL,product_class_id_extend_SQL,product_brand_id_SQL,product_keywords_SQL,isOnsale_SQL,isOutOfStore_SQL,isSales_SQL
Dim All_SQL

'构造sql命令.
'ProductName.
If CokeShow.filtRequest(Request("ProductName"))="" Then
	ProductName_SQL=""
Else
	ProductName_SQL=" AND ProductName LIKE '%"& CokeShow.filtRequest(Request("ProductName")) &"%' "
End If
'ProductNo.
If CokeShow.filtRequest(Request("ProductNo"))="" OR CokeShow.filtRequest(Request("ProductNo"))="所有" Then
	ProductNo_SQL=""
Else
	ProductNo_SQL=" AND ProductNo LIKE '%"& CokeShow.filtRequest(Request("ProductNo")) &"%' "
End If
'product_class_id
If CokeShow.filtRequest(Request("product_class_id"))="" Then
	product_class_id_SQL=""
Else
	If CokeShow.filtRequest(Request("product_class_id"))="0" Then
		product_class_id_SQL=""
	Else
		product_class_id_SQL=" AND product_class_id="& CokeShow.CokeClng(CokeShow.filtRequest(Request("product_class_id"))) &" "
	End If
End If
'product_class_id_extend.
If CokeShow.filtRequest(Request("product_class_id_extend"))="" Then
	product_class_id_extend_SQL=""
Else
	If CokeShow.filtRequest(Request("product_class_id"))="0" Then
		product_class_id_extend_SQL=""
	Else
		product_class_id_extend_SQL=" AND ( product_class_id_extend LIKE '%"& CokeShow.AdditionZero( Request("product_class_id_extend"), 8 ) &"%' ) "
	End If
End If
'product_brand_id
If CokeShow.filtRequest(Request("product_brand_id"))="" Then
	product_brand_id_SQL=""
Else
	If CokeShow.filtRequest(Request("product_brand_id"))="0" Then
		product_brand_id_SQL=""
	Else
		product_brand_id_SQL=" AND product_brand_id="& CokeShow.CokeClng(CokeShow.filtRequest(Request("product_brand_id"))) &" "
	End If
End If
'product_keywords
If CokeShow.filtRequest(Request("product_keywords"))="" Then
	product_keywords_SQL=""
Else
	product_keywords_SQL=" AND product_keywords LIKE '%"& CokeShow.filtRequest(Request("product_keywords")) &"%' "
End If
'isOnsale
If CokeShow.filtRequest(Request("isOnsale"))="" Then
	isOnsale_SQL=""
Else
	If CokeShow.filtRequest(Request("isOnsale"))="0" Then
		isOnsale_SQL=" AND isOnsale=0 "
	ElseIf CokeShow.filtRequest(Request("isOnsale"))="1" Then
		isOnsale_SQL=" AND isOnsale=1 "
	Else
		isOnsale_SQL=""
	End If
End If
'isOutOfStore
If CokeShow.filtRequest(Request("isOutOfStore"))="" Then
	isOutOfStore_SQL=""
Else
	If CokeShow.filtRequest(Request("isOutOfStore"))="0" Then
		isOutOfStore_SQL=" AND isOutOfStore=0 "
	ElseIf CokeShow.filtRequest(Request("isOutOfStore"))="1" Then
		isOutOfStore_SQL=" AND isOutOfStore=1 "
	Else
		isOutOfStore_SQL=""
	End If
End If
'isSales
If CokeShow.filtRequest(Request("isSales"))="" Then
	isSales_SQL=""
Else
	If CokeShow.filtRequest(Request("isSales"))="0" Then
		isSales_SQL=" AND isSales=0 "
	ElseIf CokeShow.filtRequest(Request("isSales"))="1" Then
		isSales_SQL=" AND isSales=1 "
	Else
		isSales_SQL=""
	End If
	
End If


'整合所有构造sql语句.
All_SQL = ProductName_SQL&ProductNo_SQL&product_class_id_SQL&product_class_id_extend_SQL&product_brand_id_SQL&product_keywords_SQL&isOnsale_SQL&isOutOfStore_SQL&isSales_SQL

'处理修改后能跳转回原处.
Dim fromurl
If SearchNowResult="True" Then
	'如果当前为高级查询列表时，获取当前的所有高级查询URL的参数.
	fromurl=CokeShow.EncodeURL( CokeShow.GetAllUrlII,"" )
End If

'高级查询专区 End



%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title><% = TitleName %></title>
	
	<link type="text/css" rel="stylesheet" href="<% =filename_dj_MainCss %>" />
	<link type="text/css" rel="stylesheet" href="<% =filename_dj_ThemesCss %>" />
	
	<link type="text/css" rel="stylesheet" href="<% =dir_dj_system_foreground %>dojox/widget/Toaster/Toaster.css" />
	<link type="text/css" rel="stylesheet" href="../style/general_style.css" />
	<style type="text/css">
		input.dojoButton1 {
			border-color:#dedede;
			background: #f5f5f5 url(../images/buttonActive.png) top left repeat-x;
			/*background: none;*/
			padding: 1px;
			font-size: 12px;
			border: 1px black solid !important;
			/*
			color:#243C5F;
			background:#fcfcfc url("../images/buttonHover.png") repeat-x top left;
			
			border-color:#dedede;
			background: #f5f5f5 url("../images/buttonActive.png") top left repeat-x;*/
		}
	</style>
	
	<script type="text/javascript" src="../script/public.js"></script>
	
	<script type="text/javascript" src="<% =filename_dj %>" djConfig="parseOnLoad:<% =parseOnLoad_dj %>, isDebug:<% =isDebug_dj %>, debugAtAllCosts:<% =isDebug_dj %>"></script>
	<script type="text/javascript" src="<% =filenameWidgetsCompress_dj %>"></script>
	<script type="text/javascript">
		dojo.require("dojo.parser");
		dojo.require("dijit.form.ValidationTextBox");
		//dojo.require("dijit.form.NumberSpinner");
		dojo.require("dijit.form.Button");
		dojo.require("dijit.form.Form");
		dojo.require("dijit.form.CurrencyTextBox");
		dojo.require("dijit.form.NumberTextBox");
		dojo.require("dijit.form.FilteringSelect");
		dojo.require("dijit.form.DateTextBox");
		
		dojo.require("dijit.layout.ContentPane");
		dojo.require("dijit.layout.StackContainer");
		
		dojo.require("dijit.Dialog");
		//上传图片.
		dojo.require("dojo.io.iframe");
		
		dojo.require("dojox.widget.Toaster");
	</script>
	
	<script type="text/javascript">
		//Table偶数行变色函数.
		dojo.addOnLoad(function() {
			//初始化列表.
			stripeTables("listGo");
			
			//赋予菜品编号默认值.
			if (dojo.byId("ProductNo").value == '') {
				dojo.byId("ProductNo").value = 'LCH' + String(parseInt( Math.random()*99999999 + 12345678 ));
			}
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
	
	<script type="text/javascript">
		var Browser = new Object();
		
		Browser.isMozilla = (typeof document.implementation != 'undefined') && (typeof document.implementation.createDocument != 'undefined') && (typeof HTMLDocument != 'undefined');
		Browser.isIE = window.ActiveXObject ? true : false;
		Browser.isFirefox = (navigator.userAgent.toLowerCase().indexOf("firefox") != - 1);
		Browser.isSafari = (navigator.userAgent.toLowerCase().indexOf("safari") != - 1);
		Browser.isOpera = (navigator.userAgent.toLowerCase().indexOf("opera") != - 1);
		
		
		/**添加扩展分类**/
		function addOtherClass(conObj,copyElementName,nowElementName)
		{
		  var sel = document.createElement("SELECT");
		  //var selCat = document.forms['form1'].elements['product_class_id'];
		  var selCat = document.getElementById(copyElementName); 
			console.log("selCat.length: " + selCat.length);
		  for (i = 0; i < selCat.length; i++)
		  {
			  var opt = document.createElement("OPTION");
			  opt.text = selCat.options[i].text;
			  opt.value = selCat.options[i].value;
			  
			  if (Browser.isIE)
			  {
				  sel.add(opt);
			  }
			  else
			  {
				  sel.appendChild(opt);
			  }
		  }
		  conObj.appendChild(sel);
		  sel.name = nowElementName;
		  //sel.onChange = function() {checkIsLeaf(this);};
		}
	</script>
	
	
	<script type="text/javascript">
		//菜品相册上传功能专用函数add.构建上传对话框.(此页面专用的PhotosAdd函数)
		//编程法——不行！OK
		function ShowDialog_AddPhotos(titlename,hrefurl,styleStr) {
			console.log(titlename, hrefurl,styleStr);
			
			//构建新tr-td的DOM元素.
			createDOM_Photos(hrefurl);
			
			var d1 = new dijit.Dialog({
					title:	titlename,
					href:	hrefurl + "&CokeShow=" + Math.random(),
					style:'' + styleStr + ''
					});
			//dijit.byId("dialog2").show();
			d1.show();
			//WL自创法宝.
			dojo.connect(d1, "hide", d1, function(e){d1.destroy()});
			
		}
		
		//构建新tr-td的DOM元素.
		function createDOM_Photos(hrefurl) {
			var tr_td_All = dojo.byId("tr_td_All");
			var randomNumber = hrefurl.split("controlStr=")[1];
			
			//构建tr.
			//alert( hrefurl.split("controlStr=")[1] );
			var trHolder = document.createElement("tr");
			trHolder.setAttribute( "id","tr_td" + randomNumber );
			//将整个tr挂到DOM树中.
			tr_td_All.appendChild(trHolder);
			
			//构建td.
			var tdHolder1 = document.createElement("td");
			//tdHolder1.setAttribute( "style","text-align: right;" );
			tdHolder1.style.textAlign = "right";
			trHolder.appendChild(tdHolder1);
				
				//构建text.
//				var textIDHolder = document.createTextNode("ID");
//				tdHolder1.appendChild(textIDHolder);
				//构建input ID.
				var inputIDHolder = document.createElement("input");
				inputIDHolder.setAttribute( "type","hidden" );
				inputIDHolder.setAttribute( "maxlength","3" );
				inputIDHolder.setAttribute( "size","3" );
				inputIDHolder.setAttribute( "name","photos_id" );
				inputIDHolder.setAttribute( "value","" );
				inputIDHolder.onclick = function () { this.select(); };
				tdHolder1.appendChild(inputIDHolder);
				//构建a.
				var LinksPhotosHolder = document.createElement("a");
				LinksPhotosHolder.setAttribute( "href","/images/no_1.png" );
				LinksPhotosHolder.setAttribute( "target","_blank" );
				LinksPhotosHolder.setAttribute( "id","href_photo" + randomNumber );
				tdHolder1.appendChild(LinksPhotosHolder);
				//构建img.
				var ImgPhotosHolder = document.createElement("img");
				ImgPhotosHolder.setAttribute( "src","/images/no_1.png" );
				ImgPhotosHolder.setAttribute( "id","src_photo" + randomNumber );
				ImgPhotosHolder.setAttribute( "width","118" );
				ImgPhotosHolder.setAttribute( "height","118" );
				LinksPhotosHolder.appendChild(ImgPhotosHolder);
				//构建input photo value.
				var inputPhotosHolder = document.createElement("input");
				inputPhotosHolder.setAttribute( "type","hidden" );
				inputPhotosHolder.setAttribute( "maxlength","50" );
				inputPhotosHolder.setAttribute( "name","photos_src" );
				inputPhotosHolder.setAttribute( "id","value_photo" + randomNumber );
				inputPhotosHolder.setAttribute( "value","" );
				inputPhotosHolder.onclick = function () { this.select(); };
				tdHolder1.appendChild(inputPhotosHolder);
				
					//Tooltip.
					//构建将要被Tooltip替换的span元素.
					var tmpTooltip = document.createElement("span");
					tdHolder1.appendChild(tmpTooltip);
					//构建修改上传图片按钮dijit.
					var TooltipHolder = new dijit.Tooltip({
												label:"<img src='"+ImgPhotosHolder.src+"' id='tmp"+randomNumber+"' />",
												id:'tmp' + randomNumber ,
												connectId:'src_photo' + randomNumber
												}, tmpTooltip);
					//TooltipHolder.appendChild(ImgPhotosHolderII);
				
			//构建td.
			var tdHolder2 = document.createElement("td");
			tdHolder2.setAttribute( "style","text-align: left; width:150px;" );
			trHolder.appendChild(tdHolder2);
//				//1.tmpInput1.
//				//构建将要被NumberSpinner替换的input元素.
//				var tmpInput1 = document.createElement("input");
//				tdHolder2.appendChild(tmpInput1);
//				//构建input - NumberSpinner排序.
//				var inputOrderidHolder = new dijit.form.NumberSpinner({
//												constraints:{ max:999, min:0 },
//												style:'width:8em;',
//												name:'photos_orderid',
//												value:'0'
//												}, tmpInput1);
				//构建input 排序.
				var inputOrderidHolder = document.createElement("input");
				inputOrderidHolder.setAttribute( "type","text" );
				inputOrderidHolder.setAttribute( "maxlength","3" );
				inputOrderidHolder.setAttribute( "size","3" );
				inputOrderidHolder.setAttribute( "name","photos_orderid" );
				inputOrderidHolder.setAttribute( "value","0" );
				inputOrderidHolder.onclick = function () { this.select(); };
				tdHolder2.appendChild(inputOrderidHolder);
				
			
			//构建td.
			var tdHolder3 = document.createElement("td");
			tdHolder3.setAttribute( "style","text-align: left;" );
			trHolder.appendChild(tdHolder3);
				//1.Button.
				//构建将要被button替换的div元素.
				var tmpButton1 = document.createElement("button");
				tdHolder3.appendChild(tmpButton1);
				//构建修改上传图片按钮dijit.
				var inputModifyButtonHolder = new dijit.form.Button({
												label:"&nbsp;修改上传图片&nbsp;"
												}, tmpButton1);
				dojo.connect( inputModifyButtonHolder, "onClick", function(){ShowDialog('<img src=/images/up.gif />修改上传图片', '../upload/index.asp?Action=Add&controlStr=' + randomNumber, 'width:300px;height:200px;');} );		//唯一可能会出错的地方！WL.
				
				
				//2.Button.
				//构建将要被button替换的div元素.
				var tmpButton2 = document.createElement("button");
				tdHolder3.appendChild(tmpButton2);
				//构建删除上传图片按钮dijit.
				var inputDeleteButtonHolder = new dijit.form.Button({
												label:"&nbsp;删除&nbsp;"
												}, tmpButton2);
				dojo.connect( inputDeleteButtonHolder, "onClick", function(){deleteDOM_Photos( randomNumber );} );		//唯一可能会出错的地方！WL.
				
			
		}
		
		//删除指定的tr-td的DOM元素.
		function deleteDOM_Photos(trID) {
			//alert('tr_td' + trID);
			//parentNode.removeChild(Node);
			var currentTrID = 'tr_td' + trID;
			//WL
			//clearAllNode( dojo.byId(currentTrID) );
			//WL
			
			var confirm_del = confirm("确定要删除此图片吗？");
			if (confirm_del == true) {
				
				
				//先判断是否为以上传的图片，数据库以有记录？如果是则执行ajax删除相应记录，如果否则简单的删除一下DOM元素即可！
				//正整数
				//alert(isInteger(trID));
				if (isInteger(trID) == true) {
					//传入的是正整数（数据库有记录的图片），提交Ajax删除相应图片记录.
					//...Ajax...
					
					dojo.xhrGet({
						url:		"services/generalservices.asp",
						content:	{ query:'True',open:'True',Action:'DeleteProductPhotos',id: trID },
						timeout:	10000,
						handleAs:	"json",
						//handle:		supervisorNameValidationHandler,	//处理回调.
						load:		function(response) {
										if (response.valid) {
											//成功删除图片时.销毁相应DOM元素.
											dojo.byId("tr_td_All").removeChild( dojo.byId(currentTrID) );
											
											//发布成功消息！
											dojo.publish("xhrDeletePhotosScc", [{
												 message: "<img src=/images/png-0094.png width=32 /> " + response.message,	
												 type: "fatal",
												 duration: 0
											}]
											);
											
										}
										if (! response.valid) {
											//失败删除图片时.警报错误，要求重试操作.
											//发布失败提示消息！
											dojo.publish("xhrDeletePhotosError", [{
												 message: "<img src=/images/no.png width=32 /> " + response.message,	
												 type: "error",
												 duration: 0
											}]
											);
											
										}
										
						},
						
						error:		function(text) {
									//一个Toaster将捕获这个错误并显示它.
									dojo.publish("xhrDeletePhotosError", [{
										 message: "<img src=/images/no.png width=32 /> " + "删除操作失败，请重试！",		//将error对象 传给小部件的message发布参数里，进行传送!
										 type: "error",
										 duration: 0
									}]
									);
									
									return text;
						}
						
						
					});
						
					
						
						
					//...Ajax...
				} else {
					//传入的是小数（数据库没有记录的新操作图片），直接删除相应DOM元素即可.
					dojo.byId("tr_td_All").removeChild( dojo.byId(currentTrID) );
					
					//发布成功消息！
					dojo.publish("xhrDeletePhotosScc", [{
						 message: "<img src=/images/png-0094.png width=32 /> 您的删除操作成功!",	
						 type: "fatal",
						 duration: 0
					}]
					);
					
				}
				
				
				
			}
			
		}
		
		//WL
		function clearAllNode(parentNode){
			while (parentNode.firstChild) {
			   var oldNode = parentNode.removeChild(parentNode.firstChild);
			   oldNode = null;
			 }
		}
		//WL
		
		
		//判断是否为正整数
		function isInteger(num) {
			var patrn=/^[0-9]*[1-9][0-9]*$/;
			
			if (!patrn.exec(num)) {
				return false;
			}
			else {
				return true;
			} 
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
					<!--#include file="menu02_1.asp"-->
				</p>
				
			<h3>当前操作</h3>
				
				<ul>
					<!--<li><a href="http://www.iw3c2.org/">&#187;IW<sup>3</sup>C<sup>2</sup></a></li>-->
					
					
					<li><a href="<% = CurrentPageNow %>">&#187;返回列表</a></li>
					<li><a href="#" onClick="deleteLot();">&#187;删除操作</a></li>
				</ul>
				
		</div><!-- End news1-->
		
		<div class="news2" style="display:none;">
		
			<h3>查询操作</h3>
			<form action="<% =CurrentPageNow %>" method="GET" name="custForm" id="custForm"
			dojoType="dijit.form.Form"
			>
			<p>
					
					<select name="TypeSearch" id="TypeSearch">
					    
					    <option value="id" selected>按ID查询</option>
					    <option value="ProductName" >按菜品名称查询</option>
						<option value="ProductNo" >按货号查询</option>
						
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
		
		ElseIf Action="SearchNow" Then
			Call SearchNow()
		
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


<div dojoType="dojox.widget.Toaster"
duration="0"
messageTopic="xhrDeletePhotosScc"
positionDirection="tr-left"

/><!--"br-up","br-left","bl-up","bl-right","tr-down","tr-left","tl-down","tl-right"-->
<div dojoType="dojox.widget.Toaster"
duration="0"
messageTopic="xhrDeletePhotosError"
positionDirection="tr-left"
/>

</body>
</html>
<%
Sub Main()
	
	Select Case ExecuteSearch
		Case 0
			sql="SELECT TOP 500 * FROM "& CurrentTableName &" WHERE deleted=0 ORDER BY id DESC"
			strGuide=strGuide & "所有"& UnitName
		Case 1
			sql="SELECT TOP 500 * FROM "& CurrentTableName &" WHERE deleted=0 ORDER BY iis DESC"
			strGuide=strGuide & "登录次数最多的前500个"& UnitName
		Case 6
			sql="SELECT TOP 500 * FROM "& CurrentTableName &" WHERE deleted=0 "& All_SQL &" ORDER BY id DESC"
			strGuide=strGuide & "查询结果"& UnitName
			
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
					Case "ProductName"
						sql="select * from "& CurrentTableName &" where deleted=0 and ProductName like '%"& Keyword &"%' order by id desc"
						strGuide=strGuide & "菜品名称中含有“ <font color=red>" & Keyword & "</font> ”的"& UnitName
					Case "ProductNo"
						sql="select * from "& CurrentTableName &" where deleted=0 and ProductNo like '%"& Keyword &"%' order by id desc"
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
						
						<th>朋友名字</th>
						<th>朋友Email</th>
						
						<th>推荐的菜品</th>
						<th>价格</th>
						
						<th>日期</th>
						
                        <th>推荐人会员帐号</th>
                        <th>推荐人姓名</th>
						
						<!--<th>操作</th>-->
					  </tr>
					  </thead>
					  <tbody>
					  
					  <%
					  If RS.EOF Then
					  %>
					  <tr>
						<td colspan="9" style="color:red;">对不起，没有记录...</td>
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
						
						<td><%=RS("FName")%>&nbsp;</td>
						<td><%=RS("FEmail")%>&nbsp;</td>
						
						<td><a href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num(RS("product_id")) %>" target="_blank"><%=RS("PName")%></a>&nbsp;</td>
						<td><%=FormatCurrency(RS("PPrice"),2)%>&nbsp;</td>
						
						<td><%=RS("adddate")%>&nbsp;</td>
						
                        
                        <td><%=RS("Account_LoginID")%>&nbsp;</td>
                        <td><%=RS("Account_LoginCNNAME")%>&nbsp;</td>
						
						<!--<td>
						
						<a href="?Action=Delete&id=<%=RS("id")%>" onClick="return confirm('确定要删除此<% =UnitName %>吗？');">删除</a>
						</td>-->
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


Sub SearchNow()
	strGuide=strGuide & "查询"& UnitName
%>

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
				
				
				<form action="<% = CurrentPageNow %>" method="get" name="form1" id="form1"
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
						<td style="text-align: right;">菜品名称</td>
						<td style="text-align: left;" colspan="3">
						<input type="text" id="ProductName" name="ProductName"
							dojoType="dijit.form.ValidationTextBox"
							required="false"
							propercase="true"
							promptMessage=""
							invalidMessage="菜品名称长度必须在0-50之内."
							trim="false"
							lowercase="false"
							regExp=".{0,50}"
							 value=""
							 style="width: 300px;"
							/>
						</td>
						
					  </tr>
					  
					  
					  <tr>
						<td style="text-align: right;">菜品货号</td>
						<td style="text-align: left;" colspan="3">
						<input type="text" id="ProductNo" name="ProductNo" size="20"
							dojoType="dijit.form.ValidationTextBox"
							required="false"
							propercase="false"
							promptMessage=""
							invalidMessage="超过20字符."
							regExp=".{0,20}"
							trim="true"
							 value="所有"
							
							/>
						</td>
						
					  </tr>
					  
					  <tr>
						<td style="text-align: right;">类别</td>
						<td style="text-align: left;" colspan="3">
							<!--dojoType="dijit.form.FilteringSelect"
								autoComplete="true"
                                forceValidOption="true"
                                queryExpr="*${0}*"
                                class="input_tell"
                                style="width:250px; height:24px;"-->
							<select name="product_class_id" id="product_class_id"
								
							>
								<%
								Call CokeShow.Option_ID("[CXBG_product_class]","",0,0,"classid","classname",True)
								%>
							</select>
						</td>
						
					  </tr>
					  <tr>
						<td style="text-align: right;">扩展类别</td>
						<td style="text-align: left;" colspan="3">
						<!--dojoType="dijit.form.FilteringSelect"
								autoComplete="true"
                                forceValidOption="true"
                                queryExpr="*${0}*"
                                class="input_tell"
                                style="width:250px; height:24px;"-->
							<select name="product_class_id_extend" id="product_class_id_extend"
								
							>
								<%
								Call CokeShow.Option_ID("[CXBG_product_class]","",0,0,"classid","classname",True)
								%>
							</select>
						
						</td>
						
					  </tr>
					  
					  <tr>
						<td style="text-align: right;">品牌</td>
						<td style="text-align: left;" colspan="3">
							<!--dojoType="dijit.form.FilteringSelect"
								autoComplete="true"
                                forceValidOption="true"
                                queryExpr="*${0}*"
                                class="input_tell"
                                style="width:250px; height:24px;"-->
							<select name="product_brand_id" id="product_brand_id"
								
	
							  >
								<option value="0">请选择</option>
								<%
								Call CokeShow.Option_ID("[CXBG_product_brand]","",8,0,"classid","classname",True)
								%>
							  </select>
						</td>
						
					  </tr>
					  
					  <tr>
						<td style="text-align: right;">菜品关键词</td>
						<td style="text-align: left;" colspan="3">
						<input type="text" id="product_keywords" name="product_keywords"
								dojoType="dijit.form.ValidationTextBox"
								required="false"
								propercase="true"
								promptMessage="(提示：关键字主要用于专题活动的时候.)"
								invalidMessage="菜品关键字长度必须在0-50之内."
								trim="true"
								lowercase="false"
								regExp=".{0,50}"
								 value=""
								 style="width: 300px;"
								/>
						</td>
						
					  </tr>
					  
					  
					  
					  <tr>
						<td style="text-align: right;">促销</td>
						<td style="text-align: left;" colspan="3">
						<input type="radio" name="isSales" id="isSales" value="" checked="checked" />所有
						&nbsp;
						<input type="radio" name="isSales" id="isSales1" value="1" />促销
						&nbsp;
						<input type="radio" name="isSales" id="isSales0" value="0" />非促销
						
						</td>
						
					  </tr>
					  
					  
					  <tr>
						<td style="text-align: right;">上架</td>
						<td style="text-align: left;" colspan="3">
							<input type="radio" name="isOnsale" id="isOnsale" value="" checked="checked" />所有
							&nbsp;
							<input type="radio" name="isOnsale" id="isOnsale1" value="1" />上架
							&nbsp;
							<input type="radio" name="isOnsale" id="isOnsale0" value="0" />下架
							
							
						</td>
						
					  </tr>
					  
					  <tr>
						<td style="text-align: right;">是否缺货</td>
						<td style="text-align: left;" colspan="3">
							
							<input type="radio" name="isOutOfStore" id="isOutOfStore" value="" checked="checked" />所有
							&nbsp;
							<input type="radio" name="isOutOfStore" id="isOutOfStore1" value="1" />缺货
							&nbsp;
							<input type="radio" name="isOutOfStore" id="isOutOfStore0" value="0" />不缺货
							
						</td>
						
					  </tr>
					  
					  
					  
					  
					  <tr>
						<td style="text-align: right;" colspan="4">
						  <input type="hidden" name="Action"
						  value=""
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
					
					
					<input type="hidden" name="ExecuteSearch" value="6" />
					<input type="hidden" name="SearchNowResult" value="True" />
				</form>
			
			
			</p>
					
			
			
			<p>
			
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
				
				><!--dojoType="dijit.form.Form" 为了标签dijit，暂时去掉了这里。-->
				
				<!--
				控制器
				-->
				<span
				dojoType="dijit.layout.StackController" containerId="sc"
				>
				</span>
				
				
				<!--
				容器Begin
				-->
				<div id="sc" style="height: 100%;"
				dojoType="dijit.layout.StackContainer"
				>
					<div id="c1" title="&nbsp;基本信息&nbsp;"
					dojoType="dijit.layout.ContentPane"
					loadingMessage="读取中<img src='/images/loading.gif' />……"
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
							<td style="text-align: right;">菜品名称</td>
							<td style="text-align: left;">
							<input type="text" id="ProductName" name="ProductName"
							dojoType="dijit.form.ValidationTextBox"
							required="true"
							propercase="true"
							promptMessage="菜品名称为必填项."
							invalidMessage="菜品名称长度必须在1-50之内."
							trim="false"
							lowercase="false"
							
							 value=""
							 style="width: 300px;"
							/>
							</td>
							
							<td style="text-align: right;">菜品货号</td>
							<td style="text-align: left;">
							<div type="text" id="ProductNo" name="ProductNo" size="20"
							dojoType="dijit.form.ValidationTextBox"
							required="true"
							propercase="false"
							promptMessage=""
							invalidMessage="菜品货号长度必须至少在3位以上."
							regExp=".{3,20}"
							trim="true"
							 value=""
							
							>
								
							</div>
							</td>
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">类别</td>
							<td style="text-align: left;" colspan="3">
							<select name="product_class_id" id="product_class_id">
								<%
								Call CokeShow.Option_ID("[CXBG_product_class]","",0,0,"classid","classname",True)
								%>
							</select> <span style="color:#999999;">必须选择.</span>
							</td>
							
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">扩展类别</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="button" class="dojoButton1" value="&nbsp;添加&nbsp;" onClick="addOtherClass(this.parentNode,'product_class_id','product_class_id_extend')" />
							
							</td>
							
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">品牌</td>
							<td style="text-align: left;" colspan="3">
							<!--dojoType="dijit.form.FilteringSelect"
								autoComplete="true"
								forceValidOption="true"-->
							<select name="product_brand_id" id="product_brand_id"
								
	
							  >
								<%
								Call CokeShow.Option_ID("[CXBG_product_brand]","",8,0,"classid","classname",True)
								%>
							  </select>
							</td>
							
							
						  </tr>
						  
						  
						  
						  
						  
						  
						  <tr>
							<td style="text-align: right;">(8折)持卡会员价格</td>
							<td style="text-align: left;" colspan="3">
								<div name="UnitPrice" id="UnitPrice"
								dojoType="dijit.form.CurrencyTextBox"
								value="0"
								constraints="{ currency:'RMB', places:2 }"
								style="width:100px;"
							  	>
									<script type="dojo/method" event="onChange" args="UnitPrice">
										dojo.publish("form1/change/UnitPrice", [UnitPrice]);
									</script>
								</div>
								RMB
								
								&nbsp;
								
								<button type="button" id="toCountBtn" 
								  dojoType="dijit.form.Button"
								  
								  >
								 	&nbsp;计算器&nbsp;
									<script type="dojo/method" event="onClick">
										//var tmpCountNum = formatFloat( dojo.byId("UnitPrice").value/0.8 , 2 );
										//var tmpCountNum = dojo.byId("UnitPrice").value * 1.5 ;
										//alert('当前价格的实时计算结果为:' + tmpCountNum);
										window.open( '/jisuanqi.html?Action=Add&controlStr=' + Math.random() );
										
									</script>
								  </button>
							</td>
							
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">价格</td>
							<td style="text-align: left;" colspan="3">
								<div name="UnitPrice_Market" id="UnitPrice_Market"
								dojoType="dijit.form.CurrencyTextBox"
								value="0"
								constraints="{ currency:'RMB', places:2 }"
								style="width:100px;"
							  	>
									<script type="dojo/method">
										dojo.subscribe("form1/change/UnitPrice", this, function(UnitPrice) {
											//alert(UnitPrice);
											//alert(this.value);
											if (this.value == '0') {
												dojo.byId("UnitPrice_Market").value = UnitPrice * 1.5;
											} else {
												//alert('god!');
												dojo.byId("UnitPrice_Market").value = UnitPrice * 1.5;
												//dojo.byId("UnitPrice_Market__span").innerHTML = '当前价格的实时计算结果为:<b>' + UnitPrice * 1.5 +'</b>';
											}
										});
									</script>
								</div>
								RMB
								
								<span id="UnitPrice_Market__span" style=" color:#FF6600;"></span>
							</td>
							
						  </tr>
						  
						  
						  <!--<tr>
							<td style="text-align: right;">赠送积分数</td>
							<td style="text-align: left;" colspan="3">
								<div name="jifen" id="jifen"
								dojoType="dijit.form.NumberTextBox"
								value="0"
								constraints="{ pattern:'#,###+' }"
								
							  	>
								</div>
							</td>
							
						  </tr>-->
						  <input type="hidden" name="jifen" id="jifen" value="0" />
						  
						  
						  <tr>
							<td style="text-align: right;">新品</td>
							<td style="text-align: left;" colspan="3">
								
								<input type="checkbox" name="is_display_newproduct" value="1"
								
								/>
								<span style="color:#999999;"> 打勾表示为新品，将在首页等多地的区块进行优先展示.</span>
								
							</td>
							
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">
								<input type="checkbox" name="isSales" id="isSales" value="1" onClick="DisplayTheElement('chuxiao_area')" /><label for="isSales">促销</label>
								<br />
								<div style="color:#999999;">(此处不打勾，则<br />菜品不会有促销价.)</div>
							</td>
							<td style="text-align: left;" colspan="3">
								<!--chuxiao_area begin-->
								<div id="chuxiao_area" style="display: none;">
									<label for="UnitPrice_Sales">促销价</label>
									<div name="UnitPrice_Sales" id="UnitPrice_Sales"
									dojoType="dijit.form.CurrencyTextBox"
									value="0"
									constraints="{ currency:'RMB', places:2 }"
									style="width:100px;"
									>
									</div>
									RMB
									
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									
									<label for="isSales_StartDate">起始促销日期</label>
									<div id="isSales_StartDate" name="isSales_StartDate"
									dojoType="dijit.form.DateTextBox"
									required="false"
									constraints="{min:'<% =CokeShow.filt_DateStr( DateAdd("d",0, Date()) ) %>', max:'2011-01', datePattern:'yyyy-MM-dd'}"
									promptMessage="请选择起始促销日期."
									invalidMessage="Invalid Service Date."
									style="width:100px;"
									>
									</div>
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<label for="isSales_StopDate">结束促销日期</label>
									<div id="isSales_StopDate" name="isSales_StopDate"
									dojoType="dijit.form.DateTextBox"
									required="false"
									constraints="{min:'<% =CokeShow.filt_DateStr( DateAdd("d",1, Date()) ) %>', max:'2011-01', datePattern:'yyyy-MM-dd'}"
									promptMessage="请选择结束促销日期."
									invalidMessage="Invalid Service Date."
									style="width:100px;"
									>
									</div>
									
								
								</div>
								<!--chuxiao_area end-->
							</td>
							
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">上架</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="checkbox" name="isOnsale" value="1" checked="checked" /><span style="color:#999999;"> 打勾表示允许上架展示，否则不允许上架展示.</span>
							</td>
							
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">是否缺货</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="checkbox" name="isOutOfStore" value="1" /><span style="color:#999999;"> 打勾表示缺货.</span>
							</td>
							
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">菜品关键词</td>
							<td style="text-align: left;" colspan="3">
								<input type="text" id="product_keywords" name="product_keywords"
								dojoType="dijit.form.ValidationTextBox"
								required="false"
								propercase="true"
								promptMessage="菜品关键词并不是必填项，可以留空.<br />(提示：关键字主要用于专题活动的时候.)"
								invalidMessage="菜品关键字长度必须在1-50之内."
								trim="true"
								lowercase="false"
								regExp=".{0,50}"
								 value=""
								 style="width: 300px;"
								/> <span style="color:#999999;">请用,逗号隔开.可以留空</span>
								<br /><span style="color:#999999;">例如：元旦狂欢20100101,情人佳节20100214,...即可关联 元旦狂欢 和 情人佳节 等的专题或广告！</span>
							</td>
							
						  </tr>
						  
						  
						  
						  
						  
						  <!--<tr>
							<td style="text-align: right;">加入推荐</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="checkbox" name="xxx" value="" checked="checked" />新品&nbsp;&nbsp;&nbsp;
							<input type="checkbox" name="xxx" value="" />促销&nbsp;&nbsp;&nbsp;
							
							
							</td>
							
						  </tr>-->
						  
						  
						  
						  <tr>
							<td style="text-align: center; background-color:#FEFEF0; color:#999999;" colspan="4">
								菜品附加选项
							</td>
						  </tr>
						  <!--可选项类别 begin-->
						  <tr>
							<td style="text-align: right;">所属菜系</td>
							<td style="text-align: left;" colspan="3">
							<select name="product_businessUSE_id" id="product_businessUSE_id">
								<%
								Call CokeShow.Option_ID("[CXBG_product_businessUSE]","",0,0,"classid","classname",True)
								%>
							</select> <span style="color:#999999;">可不选.</span>
							</td>
						  </tr>
						  <tr>
							<td style="text-align: right;">扩展所属菜系</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="button" class="dojoButton1" value="&nbsp;添加&nbsp;" onClick="addOtherClass(this.parentNode,'product_businessUSE_id','product_businessUSE_id_extend')" />
							
							</td>
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">所属口味</td>
							<td style="text-align: left;" colspan="3">
							<select name="product_activityUSE_id" id="product_activityUSE_id">
								<%
								Call CokeShow.Option_ID("[CXBG_product_activityUSE]","",0,0,"classid","classname",True)
								%>
							</select> <span style="color:#999999;">可不选.</span>
							</td>
						  </tr>
						  <tr>
							<td style="text-align: right;">扩展所属口味</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="button" class="dojoButton1" value="&nbsp;添加&nbsp;" onClick="addOtherClass(this.parentNode,'product_activityUSE_id','product_activityUSE_id_extend')" />
							
							</td>
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">福利用途</td>
							<td style="text-align: left;" colspan="3">
							<select name="product_welfareUSE_id" id="product_welfareUSE_id">
								<%
								Call CokeShow.Option_ID("[CXBG_product_welfareUSE]","",0,0,"classid","classname",True)
								%>
							</select> <span style="color:#999999;">可不选.</span>
							</td>
						  </tr>
						  <tr>
							<td style="text-align: right;">扩展福利用途</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="button" class="dojoButton1" value="&nbsp;添加&nbsp;" onClick="addOtherClass(this.parentNode,'product_welfareUSE_id','product_welfareUSE_id_extend')" />
							
							</td>
						  </tr>
						  <!--可选项类别 end-->
						  
						  <tr>
							<td style="text-align: right;">主厨评价</td>
							<td style="text-align: left;" colspan="3">
							<!--dojoType="dijit.form.FilteringSelect"
								autoComplete="true"
                                forceValidOption="true"
                                queryExpr="*${0}*"
                                class="input_tell"
                                style="width:250px; height:24px;"-->
							<select name="UsersEvaluate" id="UsersEvaluate"
								
	
							  >
								<option value="0">暂无主厨评价</option>
								<%
								Dim i_010
								For i_010=1 To 5
								%>
									<option value="<% =i_010 %>"><% =i_010 %>星</option>
								<%
								Next
								%>
							  </select>
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
								  &nbsp;提交，并进入下一步上传菜品图片相册&nbsp;
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
					</div>
					
					<div id="c2" title="&nbsp;详细描述&nbsp;"
					dojoType="dijit.layout.ContentPane"
					loadingMessage="读取中<img src='/images/loading.gif' />……"
					 style=" display:none;"
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
								<td style="text-align: right;">详细描述</td>
								<td style="text-align: left;">
								<%
								'使用最新编辑器.
								'定义变量
								Dim oFCKeditor
								Set oFCKeditor = New FCKeditor
								oFCKeditor.BasePath="/"
								'定义工具条（默认为：Default），当用于前台危险时，使用Basic.
								oFCKeditor.ToolbarSet="Default"
								'定义宽度（默认宽度："100%"或者766）
								oFCKeditor.Width=766
								'定义高度（默认高度：200）
								oFCKeditor.Height=380
								'输入框的初始值
								oFCKeditor.Value=""
								
								'补充设置
								oFCKeditor.Config("ImageUpload")	="true"
								oFCKeditor.Config("ImageBrowser")	="true"
								oFCKeditor.Config("LinkUpload")		="true"
								oFCKeditor.Config("LinkBrowser")	="true"
								oFCKeditor.Config("FlashUpload")	="true"
								oFCKeditor.Config("FlashBrowser")	="true"
								
								
								'创建输入框名为：description
								oFCKeditor.Create "description"
								%>
								</td>
								
								
								
								</td>
							</tr>
							
						  </tbody>
						</table>
					</div>
					
					
				
				</div>
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
				
				>
				<!--
				控制器
				-->
				<span
				dojoType="dijit.layout.StackController" containerId="sc"
				>
				</span>
				
				
				<!--
				容器Begin
				-->
				<div id="sc" style="height: 100%;"
				dojoType="dijit.layout.StackContainer"
				>
				
					<div id="c1" title="&nbsp;基本信息&nbsp;"
					dojoType="dijit.layout.ContentPane"
					loadingMessage="读取中<img src='/images/loading.gif' />……"
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
							<td style="text-align: right;">菜品名称</td>
							<td style="text-align: left;">
							<input type="text" id="ProductName" name="ProductName"
							dojoType="dijit.form.ValidationTextBox"
							required="true"
							propercase="true"
							promptMessage="菜品名称为必填项."
							invalidMessage="菜品名称长度必须在1-50之内."
							trim="false"
							lowercase="false"
							
							 value="<% =RS("ProductName") %>"
							 style="width: 300px;"
							/>
							</td>
							
							<td style="text-align: right;">菜品货号</td>
							<td style="text-align: left;">
							<div type="text" id="ProductNo" name="ProductNo" size="20"
							dojoType="dijit.form.ValidationTextBox"
							required="true"
							propercase="false"
							promptMessage=""
							invalidMessage="菜品货号长度必须至少在3位以上."
							regExp=".{3,20}"
							trim="true"
							 value="<% =RS("ProductNo") %>"
							
							>
							</div>
							</td>
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">类别</td>
							<td style="text-align: left;" colspan="3">
							<select name="product_class_id" id="product_class_id">
								<%
								Call CokeShow.Option_ID("[CXBG_product_class]","",0,RS("product_class_id"),"classid","classname",True)
								%>
							</select>
							</td>
							
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">扩展类别</td>
							<td style="text-align: left;" colspan="3" id="product_class_id_extend_parentNode">
							
							<input type="button" class="dojoButton1" value="&nbsp;添加&nbsp;" onClick="addOtherClass(this.parentNode,'product_class_id','product_class_id_extend')" />
							<%
							Dim product_class_id_extend_Array
							product_class_id_extend_Array=Split(RS("product_class_id_extend"), ",")
							Dim i_001
							For i_001=0 To Ubound(product_class_id_extend_Array)
								
								'处理特殊零的情况.
								If CokeShow.CokeClng(product_class_id_extend_Array(i_001))<>0 Then
							%>
								<select name="product_class_id_extend">
									<%
									Call CokeShow.Option_ID("[CXBG_product_class]","",0,CokeShow.CokeClng(product_class_id_extend_Array(i_001)),"classid","classname",True)
									%>
								</select>
							<%
								End If
							Next
							%>
							
							</td>
							
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">品牌</td>
							<td style="text-align: left;" colspan="3">
							<!--dojoType="dijit.form.FilteringSelect"
								dojoType="dijit.form.FilteringSelect"
								autoComplete="true"
                                forceValidOption="true"
                                queryExpr="*${0}*"
                                class="input_tell"
                                style="width:250px; height:24px;"-->
							<select name="product_brand_id" id="product_brand_id"
								
	
							  >
								<%
								Call CokeShow.Option_ID("[CXBG_product_brand]","",8,RS("product_brand_id"),"classid","classname",True)
								%>
							  </select>
							</td>
							
							
						  </tr>
						  
						  
						  
						  
						  
						  
						  
						  
						  <tr>
							<td style="text-align: right;">(8折)持卡会员价格</td>
							<td style="text-align: left;" colspan="3">
								<div name="UnitPrice" id="UnitPrice"
								dojoType="dijit.form.CurrencyTextBox"
								value="<% =RS("UnitPrice") %>"
								constraints="{ currency:'RMB', places:2 }"
								style="width:100px;"
							  	>
									<script type="dojo/method" event="onChange" args="UnitPrice">
										dojo.publish("form1/change/UnitPrice", [UnitPrice]);
									</script>
								</div>
								RMB
								
								&nbsp;
								
								<button type="button" id="toCountBtn" 
								  dojoType="dijit.form.Button"
								  
								  >
								 	&nbsp;计算器&nbsp;
									<script type="dojo/method" event="onClick">
										//var tmpCountNum = formatFloat( dojo.byId("UnitPrice").value/0.8 , 2 );
										//var tmpCountNum = dojo.byId("UnitPrice").value * 1.5 ;
										//alert('当前价格的实时计算结果为:' + tmpCountNum);
										window.open( '/jisuanqi.html?Action=Add&controlStr=' + Math.random() );
										
									</script>
								  </button>
							</td>
							
						  </tr>
						  
						  
						  <tr>
							<td style="text-align: right;">价格</td>
							<td style="text-align: left;" colspan="3">
								<div name="UnitPrice_Market" id="UnitPrice_Market"
								dojoType="dijit.form.CurrencyTextBox"
								value="<% =RS("UnitPrice_Market") %>"
								constraints="{ currency:'RMB', places:2 }"
								style="width:100px;"
							  	>
									<script type="dojo/method">
										dojo.subscribe("form1/change/UnitPrice", this, function(UnitPrice) {
											//alert(UnitPrice);
											//alert(this.value);
											if (this.value == '0') {
												dojo.byId("UnitPrice_Market").value = UnitPrice * 1.5;
											} else {
												//alert('god!');
												dojo.byId("UnitPrice_Market").value = UnitPrice * 1.5;
												//dojo.byId("UnitPrice_Market__span").innerHTML = '当前价格的实时计算结果为:<b>' + UnitPrice * 1.5 +'</b>';
											}
										});
									</script>
								</div>
								RMB
								
								<span id="UnitPrice_Market__span" style=" color:#FF6600;"></span>
							</td>
							
						  </tr>
						  
						  
						  <!--<tr>
							<td style="text-align: right;">赠送积分数</td>
							<td style="text-align: left;" colspan="3">
								<div name="jifen" id="jifen"
								dojoType="dijit.form.NumberTextBox"
								value="<% =RS("jifen") %>"
								constraints="{ pattern:'#,###+' }"
								
							  	>
								</div>
							</td>
							
						  </tr>-->
						  <input type="hidden" name="jifen" id="jifen" value="<% =RS("jifen") %>" />
						  
						  
						  
						  <tr>
							<td style="text-align: right;">新品</td>
							<td style="text-align: left;" colspan="3">
								
								<input type="checkbox" name="is_display_newproduct" value="1"
								<% If RS("is_display_newproduct")=1 Then Response.Write " checked=""checked""" %>
								/>
								<span style="color:#999999;"> 打勾表示为新品，将在首页等多地的区块进行优先展示.</span>
								
							</td>
							
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">
								<input type="checkbox" name="isSales" id="isSales"
								value="1"
								onclick="DisplayTheElement('chuxiao_area')"
								<% If RS("isSales")=1 Then Response.Write " checked=""checked""" %>
								/>
								<label for="isSales">促销</label>
								
								<br />
								<div style="color:#999999;">(此处不打勾，则<br />菜品不会有促销价.)</div>
							</td>
							<td style="text-align: left;" colspan="3">
								<!--chuxiao_area begin-->
								<div id="chuxiao_area" style="display: <% If RS("isSales")=1 Then Response.Write "black" Else Response.Write "none" %>;">
									
									<label for="UnitPrice_Sales">促销价</label>
									<div name="UnitPrice_Sales" id="UnitPrice_Sales"
									dojoType="dijit.form.CurrencyTextBox"
									value="<% =RS("UnitPrice_Sales") %>"
									constraints="{ currency:'RMB', places:2 }"
									style="width:100px;"
									>
									</div>
									RMB
									
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									
									<label for="isSales_StartDate">起始促销日期</label>
									<div id="isSales_StartDate" name="isSales_StartDate"
									dojoType="dijit.form.DateTextBox"
									required="false"
									constraints="{min:'<% =CokeShow.filt_DateStr( DateAdd("d",0, Date()) ) %>', max:'2011-01', datePattern:'yyyy-MM-dd'}"
									promptMessage="请选择起始促销日期."
									invalidMessage="Invalid Service Date."
									style="width:100px;"
									value="<% =CokeShow.filt_DateStr(RS("isSales_StartDate")) %>"
									>
									</div>
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<label for="isSales_StopDate">结束促销日期</label>
									<div id="isSales_StopDate" name="isSales_StopDate"
									dojoType="dijit.form.DateTextBox"
									required="false"
									constraints="{min:'<% =CokeShow.filt_DateStr( DateAdd("d",1, Date()) ) %>', max:'2011-01', datePattern:'yyyy-MM-dd'}"
									promptMessage="请选择结束促销日期."
									invalidMessage="Invalid Service Date."
									style="width:100px;"
									value="<% =CokeShow.filt_DateStr(RS("isSales_StopDate")) %>"
									>
									</div>
									
								
								</div>
								<!--chuxiao_area end-->
							</td>
							
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">上架</td>
							<td style="text-align: left;" colspan="3">
								
								<input type="checkbox" name="isOnsale" value="1"
								<% If RS("isOnsale")=1 Then Response.Write " checked=""checked""" %>
								/>
								<span style="color:#999999;"> 打勾表示允许上架展示，否则不允许上架展示.</span>
								
							</td>
							
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">是否缺货</td>
							<td style="text-align: left;" colspan="3">
								
								<input type="checkbox" name="isOutOfStore" value="1"
								<% If RS("isOutOfStore")=1 Then Response.Write " checked=""checked""" %>
								/>
								<span style="color:#999999;"> 打勾表示缺货.</span>
								
							</td>
							
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">菜品关键词</td>
							<td style="text-align: left;" colspan="3">
								<input type="text" id="product_keywords" name="product_keywords"
								dojoType="dijit.form.ValidationTextBox"
								required="false"
								propercase="true"
								promptMessage="菜品关键词并不是必填项，可以留空.<br />(提示：关键字主要用于专题活动的时候.)"
								invalidMessage="菜品关键字长度必须在1-50之内."
								trim="true"
								lowercase="false"
								regExp=".{0,50}"
								 value="<% =RS("product_keywords") %>"
								 style="width: 300px;"
								/> <span style="color:#999999;">请用,逗号隔开.可以留空</span>
								<br /><span style="color:#999999;">例如：元旦狂欢20100101,情人佳节20100214,...即可关联 元旦狂欢 和 情人佳节 等的专题或广告！</span>
							</td>
							
						  </tr>
						  
						  
						  
						  
						  
						  <!--<tr>
							<td style="text-align: right;">加入推荐</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="checkbox" name="xxx" value="" checked="checked" />新品&nbsp;&nbsp;&nbsp;
							<input type="checkbox" name="xxx" value="" />促销&nbsp;&nbsp;&nbsp;
							
							
							</td>
							
						  </tr>-->
						  
						  
						  <tr>
							<td style="text-align: center; background-color:#FEFEF0; color:#999999;" colspan="4">
								菜品附加选项
							</td>
						  </tr>
						  <!--可选项类别 begin-->
						  <tr>
							<td style="text-align: right;">所属菜系</td>
							<td style="text-align: left;" colspan="3">
							<select name="product_businessUSE_id" id="product_businessUSE_id">
								<%
								Call CokeShow.Option_ID("[CXBG_product_businessUSE]","",0,RS("product_businessUSE_id"),"classid","classname",True)
								%>
							</select> <span style="color:#999999;">可不选.</span>
							</td>
						  </tr>
						  <tr>
							<td style="text-align: right;">扩展所属菜系</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="button" class="dojoButton1" value="&nbsp;添加&nbsp;" onClick="addOtherClass(this.parentNode,'product_businessUSE_id','product_businessUSE_id_extend')" />
							<%
							Dim product_businessUSE_id_extend_Array
							product_businessUSE_id_extend_Array=Split(RS("product_businessUSE_id_extend"), ",")
							''Dim i_001
							For i_001=0 To Ubound(product_businessUSE_id_extend_Array)
								
								'处理特殊零的情况.
								If CokeShow.CokeClng(product_businessUSE_id_extend_Array(i_001))<>0 Then
							%>
								<select name="product_businessUSE_id_extend">
									<%
									Call CokeShow.Option_ID("[CXBG_product_businessUSE]","",0,CokeShow.CokeClng(product_businessUSE_id_extend_Array(i_001)),"classid","classname",True)
									%>
								</select>
							<%
								End If
							Next
							%>
							
							</td>
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">所属口味</td>
							<td style="text-align: left;" colspan="3">
							<select name="product_activityUSE_id" id="product_activityUSE_id">
								<%
								Call CokeShow.Option_ID("[CXBG_product_activityUSE]","",0,RS("product_activityUSE_id"),"classid","classname",True)
								%>
							</select> <span style="color:#999999;">可不选.</span>
							</td>
						  </tr>
						  <tr>
							<td style="text-align: right;">扩展所属口味</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="button" class="dojoButton1" value="&nbsp;添加&nbsp;" onClick="addOtherClass(this.parentNode,'product_activityUSE_id','product_activityUSE_id_extend')" />
							<%
							Dim product_activityUSE_id_extend_Array
							product_activityUSE_id_extend_Array=Split(RS("product_activityUSE_id_extend"), ",")
							''Dim i_001
							For i_001=0 To Ubound(product_activityUSE_id_extend_Array)
								
								'处理特殊零的情况.
								If CokeShow.CokeClng(product_activityUSE_id_extend_Array(i_001))<>0 Then
							%>
								<select name="product_activityUSE_id_extend">
									<%
									Call CokeShow.Option_ID("[CXBG_product_activityUSE]","",0,CokeShow.CokeClng(product_activityUSE_id_extend_Array(i_001)),"classid","classname",True)
									%>
								</select>
							<%
								End If
							Next
							%>
							
							</td>
						  </tr>
						  
						  <tr>
							<td style="text-align: right;">福利用途</td>
							<td style="text-align: left;" colspan="3">
							<select name="product_welfareUSE_id" id="product_welfareUSE_id">
								<%
								Call CokeShow.Option_ID("[CXBG_product_welfareUSE]","",0,RS("product_welfareUSE_id"),"classid","classname",True)
								%>
							</select> <span style="color:#999999;">可不选.</span>
							</td>
						  </tr>
						  <tr>
							<td style="text-align: right;">扩展福利用途</td>
							<td style="text-align: left;" colspan="3">
							
							<input type="button" class="dojoButton1" value="&nbsp;添加&nbsp;" onClick="addOtherClass(this.parentNode,'product_welfareUSE_id','product_welfareUSE_id_extend')" />
							<%
							Dim product_welfareUSE_id_extend_Array
							product_welfareUSE_id_extend_Array=Split(RS("product_welfareUSE_id_extend"), ",")
							''Dim i_001
							For i_001=0 To Ubound(product_welfareUSE_id_extend_Array)
								
								'处理特殊零的情况.
								If CokeShow.CokeClng(product_welfareUSE_id_extend_Array(i_001))<>0 Then
							%>
								<select name="product_welfareUSE_id_extend">
									<%
									Call CokeShow.Option_ID("[CXBG_product_welfareUSE]","",0,CokeShow.CokeClng(product_welfareUSE_id_extend_Array(i_001)),"classid","classname",True)
									%>
								</select>
							<%
								End If
							Next
							%>
							
							</td>
						  </tr>
						  <!--可选项类别 end-->
						  
						  <tr>
							<td style="text-align: right;">主厨评价</td>
							<td style="text-align: left;" colspan="3">
							<!--dojoType="dijit.form.FilteringSelect"
								dojoType="dijit.form.FilteringSelect"
								autoComplete="true"
                                forceValidOption="true"
                                queryExpr="*${0}*"
                                class="input_tell"
                                style="width:250px; height:24px;"-->
							<select name="UsersEvaluate" id="UsersEvaluate"
								
	
							  >
								<option value="0" <% If RS("UsersEvaluate")=0 Then Response.Write " selected=""selected"" " %>>暂无主厨评价</option>
								<%
								Dim i_010
								For i_010=1 To 5
								%>
									<option value="<% =i_010 %>" <% If RS("UsersEvaluate")=i_010 Then Response.Write " selected=""selected"" " %>><% =i_010 %>星</option>
								<%
								Next
								%>
							  </select>
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
					</div>
					
					<div id="c2" title="&nbsp;详细描述&nbsp;"
					dojoType="dijit.layout.ContentPane"
					loadingMessage="读取中<img src='/images/loading.gif' />……"
					 style=" display:none;"
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
								<td style="text-align: right;">详细描述</td>
								<td style="text-align: left;">
								<%
								'使用最新编辑器.
								'定义变量
								Dim oFCKeditor
								Set oFCKeditor = New FCKeditor
								oFCKeditor.BasePath="/"
								'定义工具条（默认为：Default），当用于前台危险时，使用Basic.
								oFCKeditor.ToolbarSet="Default"
								'定义宽度（默认宽度："100%"或者766）
								oFCKeditor.Width=766
								'定义高度（默认高度：200）
								oFCKeditor.Height=380
								'输入框的初始值
								oFCKeditor.Value= RS("description")
								
								'补充设置
								oFCKeditor.Config("ImageUpload")	="true"
								oFCKeditor.Config("ImageBrowser")	="true"
								oFCKeditor.Config("LinkUpload")		="true"
								oFCKeditor.Config("LinkBrowser")	="true"
								oFCKeditor.Config("FlashUpload")	="true"
								oFCKeditor.Config("FlashBrowser")	="true"
								
								'创建输入框名为：description
								oFCKeditor.Create "description"
								%>
								</td>
								
								
								
								</td>
							</tr>
							
						  </tbody>
						</table>
					</div>
					
					<div id="c3" title="&nbsp;相册&nbsp;"
					dojoType="dijit.layout.ContentPane"
					loadingMessage="读取中<img src='/images/loading.gif' />……"
					 style=" display:none;"
					>
						<table width="auto" id="listGo" cellpadding="0" cellspacing="0">
						  <thead>
						  <tr>
							
							<th style="text-align: center;">图片</th>
							
							<th style="text-align: center;">排序</th>
							<th style="text-align: center;">操作</th>
							
						  </tr>
						  </thead>
						  <tbody id="tr_td_All">
						  	<%
							'循环此菜品的相册集.
							Dim RS_PHOTOS,SQL_PHOTOS
							Set RS_PHOTOS=Server.CreateObject("Adodb.RecordSet")
							SQL_PHOTOS="SELECT * FROM [CXBG_product__photos] WHERE product_id="& RS("id") &" ORDER BY photos_orderid DESC,id ASC"
							RS_PHOTOS.Open SQL_PHOTOS,CONN,1,1
							
							If RS_PHOTOS.Eof And RS_PHOTOS.Bof Then
							
								
							
							Else
							
							Do While Not RS_PHOTOS.Eof
							%>
							<tr id="tr_td<% =RS_PHOTOS("id") %>">
								
								<td style="text-align: right;">
									<span style="color:#999999;"> 上传尺寸要求: 宽300*高300 以上的正方形尺寸！</span>
									<br />
									
									
									<input type="hidden" maxlength="3" size="3" name="photos_id" value="<% =RS_PHOTOS("id") %>" onClick="this.select();" />
									<a href="<% =RS_PHOTOS("photos_src") %>" target="_blank" id="href_photo<% =RS_PHOTOS("id") %>"
									><img src="<% =RS_PHOTOS("photos_src") %>" id="src_photo<% =RS_PHOTOS("id") %>"
									width="118" height="118"
									/></a>
									<input type="hidden" maxlength="50" name="photos_src" id="value_photo<% =RS_PHOTOS("id") %>"
									value="<% =RS_PHOTOS("photos_src") %>"
									onClick="this.select();"
									/>
									<br />
									<span style="color:#999999;"> 当前尺寸: 宽118*高118 (预览前台菜品列表页面的菜品图尺寸)</span>
									
									
									<!--begin-->
									<span dojoType="dijit.Tooltip"
									connectId="src_photo<% =RS_PHOTOS("id") %>"
									id="tmp<% =RS_PHOTOS("id") %>"
									style="display:none;"
									>
										图片详细尺寸：<br /><img src="<% =RS_PHOTOS("photos_src") %>" />
									</span>
									<!--end-->
									
								</td>
								
								<td style="text-align: left;">
									<!--<div dojoType="dijit.form.NumberSpinner"
									constraints="{ max:999, min:0 }"
									style="width:8em;"
									name="photos_orderid"
									value="<% =RS_PHOTOS("photos_orderid") %>"
									onClick="this.select();"
									>-->
									<input type="text" maxlength="3" size="3" name="photos_orderid"
									onClick="this.select();"
									value="<% =RS_PHOTOS("photos_orderid") %>"
									/>
									<!--</div>-->
									
								</td>
								<td style="text-align: left;">
									
									<button type="button" 
									dojoType="dijit.form.Button"
									onclick="ShowDialog('<img src=/images/up.gif />修改上传图片','../upload/index.asp?Action=Add&controlStr=<% =RS_PHOTOS("id") %>','width:300px;height:200px;');"
									>
									&nbsp;修改上传图片&nbsp;
									</button>
									
									
									
									<button type="button" 
									dojoType="dijit.form.Button"
									onclick="deleteDOM_Photos('<% =RS_PHOTOS("id") %>');"
									>
									&nbsp;删除&nbsp;
									</button>
									
								</td>
								
							</tr>
							<%
								RS_PHOTOS.MoveNext
							Loop
							
							End If
							
							RS_PHOTOS.Close
							Set RS_PHOTOS=Nothing
							
							%>
							
							
						  </tbody>
						  <tfoot>
						  	
							<tr>
								
								<td colspan="3">
									
									<button type="button" id="addPhotos" 
									  dojoType="dijit.form.Button"
									  onclick="ShowDialog_AddPhotos('<img src=/images/up.gif />+新增图片','../upload/index.asp?Action=Add&controlStr=' + Math.random() ,'width:300px;height:200px;');"
									  >
									  &nbsp;+新增图片&nbsp;
									  </button>
									  <br />
									  <span style="color:#FF6600">( 菜品图片具体尺寸要求: 请符合宽300*高300像素以上的正方形图片要求，同时菜品图片应为抠好的图片——浅底色，以符合菜品页风格。)</span>
								</td>
								
							</tr>
							
						  </tfoot>
						</table>
					</div>
				
				
				</div>
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
	Dim ProductName,ProductNo,product_class_id,product_class_id_extend,product_brand_id,description
	Dim UnitPrice,UnitPrice_Market,isSales,UnitPrice_Sales,isSales_StartDate,isSales_StopDate,jifen,product_keywords,isOnsale,isOutOfStore
	Dim product_businessUSE_id,product_businessUSE_id_extend,product_activityUSE_id,product_activityUSE_id_extend,product_welfareUSE_id,product_welfareUSE_id_extend
	Dim UsersEvaluate
	Dim is_display_newproduct
	
	'获取其它参数
	ProductName					=CokeShow.filtRequest(Request("ProductName"))
	ProductNo					=CokeShow.filtRequest(Request("ProductNo"))
	product_class_id			=CokeShow.filtRequest(Request("product_class_id"))
	product_class_id_extend		=CokeShow.filtRequest(Request("product_class_id_extend"))
	'BEGIN
	product_businessUSE_id		=CokeShow.filtRequest(Request("product_businessUSE_id"))
	product_businessUSE_id_extend=CokeShow.filtRequest(Request("product_businessUSE_id_extend"))
	product_activityUSE_id		=CokeShow.filtRequest(Request("product_activityUSE_id"))
	product_activityUSE_id_extend=CokeShow.filtRequest(Request("product_activityUSE_id_extend"))
	product_welfareUSE_id		=CokeShow.filtRequest(Request("product_welfareUSE_id"))
	product_welfareUSE_id_extend=CokeShow.filtRequest(Request("product_welfareUSE_id_extend"))
	'END
	product_brand_id			=CokeShow.filtRequest(Request("product_brand_id"))
	description					=CokeShow.filtRequestRich(Request("description"))
	
	UnitPrice					=CokeShow.filtRequest(Request("UnitPrice"))
	UnitPrice_Market			=CokeShow.filtRequest(Request("UnitPrice_Market"))
	isSales						=CokeShow.filtRequest(Request("isSales"))
	UnitPrice_Sales				=CokeShow.filtRequest(Request("UnitPrice_Sales"))
	isSales_StartDate			=CokeShow.filtRequest(Request("isSales_StartDate"))
	isSales_StopDate			=CokeShow.filtRequest(Request("isSales_StopDate"))
	jifen						=CokeShow.filtRequest(Request("jifen"))
	product_keywords			=CokeShow.filtRequest(Request("product_keywords"))
	isOnsale					=CokeShow.filtRequest(Request("isOnsale"))
	isOutOfStore					=CokeShow.filtRequest(Request("isOutOfStore"))
	
	UsersEvaluate				=CokeShow.filtRequest(Request("UsersEvaluate"))		'主厨评价
	
	is_display_newproduct		=CokeShow.filtRequest(Request("is_display_newproduct"))
	
	
	'验证
	If ProductName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>菜品名称不能为空！</li>"
	Else
		If CokeShow.strLength(ProductName)>50 Or CokeShow.strLength(ProductName)<2 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品名称长度不能大于50个字符，也不能小于2个字符！</li>"
		Else
			ProductName=ProductName
		End If
	End If
	
	If ProductNo="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>菜品编号不能为空！</li>"
	Else
		If CokeShow.strLength(ProductNo)>20 Or CokeShow.strLength(ProductNo)<4 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品编号长度不能大于20个字符，也不能小于4个字符！</li>"
		Else
			ProductNo=ProductNo
		End If
	End If
	
'	If cnname<>"" Then
'		If Len(cnname)>20 Then
'			FoundErr=True
'			ErrMsg=ErrMsg &"<br><li>中文名只能20位字符之内！此项也可以不填。</li>"
'		Else
'			cnname=cnname
'		End If
'	End If
	
	If product_class_id="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请选择菜品分类！</li>"
	Else
		If isNumeric(product_class_id) Then
			product_class_id=CokeShow.CokeClng(product_class_id)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前菜品分类数字不是数字(出现异常)！</li>"
		End If
	End If
	
	If product_brand_id="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请选择菜品品牌！</li>"
	Else
		If isNumeric(product_brand_id) Then
			product_brand_id=CokeShow.CokeClng(product_brand_id)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前菜品品牌数字不是数字(出现异常)！</li>"
		End If
	End If
	
	If product_class_id_extend="" Or isNull(product_class_id_extend) Or isEmpty(product_class_id_extend) Then
'		FoundErr=True
'		ErrMsg=ErrMsg &"<br><li>菜品扩展分类不能为空！</li>"
		product_class_id_extend=""
	Else
		If CokeShow.strLength(product_class_id_extend)>255 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品扩展分类长度不能大于255个字符！</li>"
		Else
			product_class_id_extend=product_class_id_extend
		End If
	End If
	
	If description="" Or isNull(description) Or isEmpty(description) Then
'		FoundErr=True
'		ErrMsg=ErrMsg &"<br><li>菜品扩展分类不能为空！</li>"
		description=""
	Else
		If CokeShow.strLength(description)>8000 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品描述长度不能大于8000个字符！</li>"
		Else
			description=description
		End If
	End If
	
'	If MSN<>"" Then
'		If CokeShow.IsValidEmail(MSN)=false Then
'			FoundErr=True
'			ErrMsg=ErrMsg & "<br><li>你的MSN格式不正确！</li>"
'		End If
'	End If
	
	
	If UnitPrice="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请填写持卡会员价格！</li>"
	Else
		If isNumeric(UnitPrice) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前持卡会员价格不是数字(非法输入)！</li>"
		End If
	End If
	
	If UnitPrice_Market="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请填写价格！</li>"
	Else
		If isNumeric(UnitPrice_Market) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前价格不是数字(非法输入)！</li>"
		End If
	End If
	
	If isSales="" Then
		isSales=0
		
	Else
		If isNumeric(isSales) Then
			isSales=CokeShow.CokeClng(isSales)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前促销选项的参数不是数字(非法输入)！</li>"
		End If
	End If
	
	If is_display_newproduct="" Then
		is_display_newproduct=0
		
	Else
		If isNumeric(is_display_newproduct) Then
			is_display_newproduct=CokeShow.CokeClng(is_display_newproduct)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>新品选项的参数不是数字(非法输入)！</li>"
		End If
	End If
	
	If UnitPrice_Sales="" Then
		
		
	Else
		If isNumeric(UnitPrice_Sales) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前促销价不是数字(非法输入)！</li>"
		End If
	End If
	
	If isSales_StartDate="" Then
		
		
	Else
		If isDate(isSales_StartDate) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前起始促销日期格式不对！</li>"
		End If
	End If
	
	If isSales_StopDate="" Then
		
		
	Else
		If isDate(isSales_StopDate) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前结束促销日期格式不对！</li>"
		End If
	End If
	
	'jifen,product_keywords,isOnsale
	If jifen="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请填写积分，如果无积分请填写为0！</li>"
	Else
		If isNumeric(jifen) Then
			jifen=CokeShow.CokeClng(jifen)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前积分不是数字(非法输入)！</li>"
		End If
	End If
	
	If product_keywords="" Or isNull(product_keywords) Or isEmpty(product_keywords) Then
		product_keywords=""
		'关键.
	Else
		If CokeShow.strLength(product_keywords)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前菜品关键词超过了50字符！</li>"
		Else
			product_keywords=Trim(product_keywords)
			
		End If
	End If
	
	If isOnsale="" Then
		isOnsale=0
		
	Else
		If isNumeric(isOnsale) Then
			isOnsale=CokeShow.CokeClng(isOnsale)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前上架选项的参数不是数字(非法输入)！</li>"
		End If
	End If
	
	If isOutOfStore="" Then
		isOutOfStore=0
		
	Else
		If isNumeric(isOutOfStore) Then
			isOutOfStore=CokeShow.CokeClng(isOutOfStore)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>是否缺货的参数不是数字(非法输入)！</li>"
		End If
	End If
	
	
	'拦截错误，不然错误往下进行！
	If FoundErr=True Then Exit Sub
	Dim newID
	
	Set RS=Server.CreateObject("Adodb.RecordSet")
	sql="SELECT * FROM "& CurrentTableName
	If Not IsObject(CONN) Then link_database
	RS.Open sql,CONN,2,2
	RS.AddNew
		
		RS("ProductName")					=ProductName	'必填项
		RS("ProductNo")						=ProductNo		'非必填项
		RS("product_class_id")				=product_class_id
		RS("product_class_id_extend")		=toProcessRequest(product_class_id_extend)
		'BEGIN
		RS("product_businessUSE_id")			=product_businessUSE_id
		RS("product_businessUSE_id_extend")		=toProcessRequest(product_businessUSE_id_extend)
		RS("product_activityUSE_id")			=product_activityUSE_id
		RS("product_activityUSE_id_extend")		=toProcessRequest(product_activityUSE_id_extend)
		RS("product_welfareUSE_id")				=product_welfareUSE_id
		RS("product_welfareUSE_id_extend")		=toProcessRequest(product_welfareUSE_id_extend)
		'END
		RS("product_brand_id")				=product_brand_id
		RS("description")					=description
		
		RS("UnitPrice")					=UnitPrice
		RS("UnitPrice_Market")			=UnitPrice_Market
		RS("isSales")					=isSales
		RS("UnitPrice_Sales")			=UnitPrice_Sales
		If isSales_StartDate<>"" Then RS("isSales_StartDate")	=isSales_StartDate
		If isSales_StopDate<>"" Then RS("isSales_StopDate")		=isSales_StopDate
		RS("jifen")						=jifen
		RS("product_keywords")			=product_keywords
		RS("isOnsale")					=isOnsale
		RS("isOutOfStore")					=isOutOfStore
		
		RS("UsersEvaluate")				=UsersEvaluate
		
		RS("is_display_newproduct")			=is_display_newproduct
	
	RS.Update
	RS.MoveLast
	newID = RS("id")
	
	RS.Close
	Set RS=Nothing

'记入日志.
Call CokeShow.AddLog("添加操作：成功添加了"& UnitName &"-"& ProductName &"["& ProductNo &"]", sql)
	
	CokeShow.ShowOK "添加"& UnitName &"成功!",CurrentPageNow &"?Action=Modify&id="& newID
End Sub


Sub SaveModify()
	Dim ProductName,ProductNo,product_class_id,product_class_id_extend,product_brand_id,description
	Dim UnitPrice,UnitPrice_Market,isSales,UnitPrice_Sales,isSales_StartDate,isSales_StopDate,jifen,product_keywords,isOnsale,isOutOfStore
	Dim product_businessUSE_id,product_businessUSE_id_extend,product_activityUSE_id,product_activityUSE_id_extend,product_welfareUSE_id,product_welfareUSE_id_extend
	Dim UsersEvaluate
	Dim is_display_newproduct
	
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
	ProductName					=CokeShow.filtRequest(Request("ProductName"))
	ProductNo					=CokeShow.filtRequest(Request("ProductNo"))
	product_class_id			=CokeShow.filtRequest(Request("product_class_id"))
	product_class_id_extend		=CokeShow.filtRequest(Request("product_class_id_extend"))
	'BEGIN
	product_businessUSE_id		=CokeShow.filtRequest(Request("product_businessUSE_id"))
	product_businessUSE_id_extend=CokeShow.filtRequest(Request("product_businessUSE_id_extend"))
	product_activityUSE_id		=CokeShow.filtRequest(Request("product_activityUSE_id"))
	product_activityUSE_id_extend=CokeShow.filtRequest(Request("product_activityUSE_id_extend"))
	product_welfareUSE_id		=CokeShow.filtRequest(Request("product_welfareUSE_id"))
	product_welfareUSE_id_extend=CokeShow.filtRequest(Request("product_welfareUSE_id_extend"))
	'END
	product_brand_id			=CokeShow.filtRequest(Request("product_brand_id"))
	description					=CokeShow.filtRequestRich(Request("description"))
	'获取相册参数
	Dim photos_id,photos_src,photos_orderid
	photos_id					=CokeShow.filtRequest(Request("photos_id"))
	photos_src					=CokeShow.filtRequest(Request("photos_src"))
	photos_orderid				=CokeShow.filtRequest(Request("photos_orderid"))
	
	UnitPrice					=CokeShow.filtRequest(Request("UnitPrice"))
	UnitPrice_Market			=CokeShow.filtRequest(Request("UnitPrice_Market"))
	isSales						=CokeShow.filtRequest(Request("isSales"))
	UnitPrice_Sales				=CokeShow.filtRequest(Request("UnitPrice_Sales"))
	isSales_StartDate			=CokeShow.filtRequest(Request("isSales_StartDate"))
	isSales_StopDate			=CokeShow.filtRequest(Request("isSales_StopDate"))
	jifen						=CokeShow.filtRequest(Request("jifen"))
	product_keywords			=CokeShow.filtRequest(Request("product_keywords"))
	isOnsale					=CokeShow.filtRequest(Request("isOnsale"))
	isOutOfStore					=CokeShow.filtRequest(Request("isOutOfStore"))
	
	UsersEvaluate				=CokeShow.filtRequest(Request("UsersEvaluate"))		'主厨评价
	
	is_display_newproduct		=CokeShow.filtRequest(Request("is_display_newproduct"))
	
	
	'验证
	If ProductName="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>菜品名称不能为空！</li>"
	Else
		If CokeShow.strLength(ProductName)>50 Or CokeShow.strLength(ProductName)<2 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品名称长度不能大于50个字符，也不能小于2个字符！</li>"
		Else
			ProductName=ProductName
		End If
	End If
	
	If ProductNo="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>菜品编号不能为空！</li>"
	Else
		If CokeShow.strLength(ProductNo)>20 Or CokeShow.strLength(ProductNo)<4 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品编号长度不能大于20个字符，也不能小于4个字符！</li>"
		Else
			ProductNo=ProductNo
		End If
	End If
	
'	If cnname<>"" Then
'		If Len(cnname)>20 Then
'			FoundErr=True
'			ErrMsg=ErrMsg &"<br><li>中文名只能20位字符之内！此项也可以不填。</li>"
'		Else
'			cnname=cnname
'		End If
'	End If
	
	If product_class_id="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请选择菜品分类！</li>"
	Else
		If isNumeric(product_class_id) Then
			product_class_id=CokeShow.CokeClng(product_class_id)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前菜品分类数字不是数字(出现异常)！</li>"
		End If
	End If
	
	If product_brand_id="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请选择菜品品牌！</li>"
	Else
		If isNumeric(product_brand_id) Then
			product_brand_id=CokeShow.CokeClng(product_brand_id)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前菜品品牌数字不是数字(出现异常)！</li>"
		End If
	End If
	
	If product_class_id_extend="" Or isNull(product_class_id_extend) Or isEmpty(product_class_id_extend) Then
'		FoundErr=True
'		ErrMsg=ErrMsg &"<br><li>菜品扩展分类不能为空！</li>"
		product_class_id_extend=""
	Else
		If CokeShow.strLength(product_class_id_extend)>255 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品扩展分类长度不能大于255个字符！</li>"
		Else
			product_class_id_extend=product_class_id_extend
		End If
	End If
	
	If description="" Or isNull(description) Or isEmpty(description) Then
'		FoundErr=True
'		ErrMsg=ErrMsg &"<br><li>菜品扩展分类不能为空！</li>"
		description=""
	Else
		If CokeShow.strLength(description)>8000 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>菜品描述长度不能大于8000个字符！</li>"
		Else
			description=description
		End If
	End If
		
'	If MSN<>"" Then
'		If CokeShow.IsValidEmail(MSN)=false Then
'			FoundErr=True
'			ErrMsg=ErrMsg & "<br><li>你的MSN格式不正确！</li>"
'		End If
'	End If
	
	
	If UnitPrice="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请填写持卡会员价格！</li>"
	Else
		If isNumeric(UnitPrice) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前持卡会员价格不是数字(非法输入)！</li>"
		End If
	End If
	
	If UnitPrice_Market="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请填写价格！</li>"
	Else
		If isNumeric(UnitPrice_Market) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前价格不是数字(非法输入)！</li>"
		End If
	End If
	
	If isSales="" Then
		isSales=0
		
	Else
		If isNumeric(isSales) Then
			isSales=CokeShow.CokeClng(isSales)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前促销选项的参数不是数字(非法输入)！</li>"
		End If
	End If
	
	If is_display_newproduct="" Then
		is_display_newproduct=0
		
	Else
		If isNumeric(is_display_newproduct) Then
			is_display_newproduct=CokeShow.CokeClng(is_display_newproduct)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>新品选项的参数不是数字(非法输入)！</li>"
		End If
	End If
	
	If UnitPrice_Sales="" Then
		
		
	Else
		If isNumeric(UnitPrice_Sales) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前促销价不是数字(非法输入)！</li>"
		End If
	End If
	
	If isSales_StartDate="" Then
		
		
	Else
		If isDate(isSales_StartDate) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前起始促销日期格式不对！</li>"
		End If
	End If
	
	If isSales_StopDate="" Then
		
		
	Else
		If isDate(isSales_StopDate) Then
			
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前结束促销日期格式不对！</li>"
		End If
	End If
	
	'jifen,product_keywords,isOnsale
	If jifen="" Then
		FoundErr=True
		ErrMsg=ErrMsg &"<br><li>请填写积分，如果无积分请填写为0！</li>"
	Else
		If isNumeric(jifen) Then
			jifen=CokeShow.CokeClng(jifen)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前积分不是数字(非法输入)！</li>"
		End If
	End If
	
	If product_keywords="" Or isNull(product_keywords) Or isEmpty(product_keywords) Then
		product_keywords=""
		'关键.
	Else
		If CokeShow.strLength(product_keywords)>50 Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前菜品关键词超过了50字符！</li>"
		Else
			product_keywords=Trim(product_keywords)
		End If
	End If
	
	If isOnsale="" Then
		isOnsale=0
		
	Else
		If isNumeric(isOnsale) Then
			isOnsale=CokeShow.CokeClng(isOnsale)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>当前上架选项的参数不是数字(非法输入)！</li>"
		End If
	End If
	
	If isOutOfStore="" Then
		isOutOfStore=0
		
	Else
		If isNumeric(isOutOfStore) Then
			isOutOfStore=CokeShow.CokeClng(isOutOfStore)
		Else
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>是否缺货的参数不是数字(非法输入)！</li>"
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
	
		RS("ProductName")					=ProductName	'必填项
		RS("ProductNo")						=ProductNo		'非必填项
		RS("product_class_id")				=product_class_id
		RS("product_class_id_extend")		=toProcessRequest(product_class_id_extend)
		'BEGIN
		RS("product_businessUSE_id")			=product_businessUSE_id
		RS("product_businessUSE_id_extend")		=toProcessRequest(product_businessUSE_id_extend)
		RS("product_activityUSE_id")			=product_activityUSE_id
		RS("product_activityUSE_id_extend")		=toProcessRequest(product_activityUSE_id_extend)
		RS("product_welfareUSE_id")				=product_welfareUSE_id
		RS("product_welfareUSE_id_extend")		=toProcessRequest(product_welfareUSE_id_extend)
		'END
		RS("product_brand_id")				=product_brand_id
		RS("description")					=description
		
		RS("modifydate")					=Now()
		
		
		RS("UnitPrice")					=UnitPrice
		RS("UnitPrice_Market")			=UnitPrice_Market
		RS("isSales")					=isSales
		RS("UnitPrice_Sales")			=UnitPrice_Sales
		If isSales_StartDate<>"" Then RS("isSales_StartDate")	=isSales_StartDate
		If isSales_StopDate<>"" Then RS("isSales_StopDate")		=isSales_StopDate
		RS("jifen")						=jifen
		RS("product_keywords")			=product_keywords
		RS("isOnsale")					=isOnsale
		RS("isOutOfStore")					=isOutOfStore
		
		RS("UsersEvaluate")				=UsersEvaluate
		
		RS("is_display_newproduct")			=is_display_newproduct
	
	RS.Update
	RS.Close
	Set RS=Nothing
	
	'更新菜品相册数据
	Call SaveModify__photos( photos_id, photos_src, photos_orderid, intID )
	
	
'记入日志.
Call CokeShow.AddLog("编辑操作：成功编辑了ID为"& intID &"的"& UnitName &"-"& ProductName &"["& ProductNo &"]", sql)
	
	
	'检测如果有hidden的fromurl传值过来，代表是源于高级查询的修改操作，此时记下上2步的fromurl.
	If Request("fromurl")<>"" And Len(Request("fromurl"))>10 Then
		CokeShow.ShowOK "修改"& UnitName &"成功!",CokeShow.DecodeURL( Request("fromurl") )
	Else
		CokeShow.ShowOK "修改"& UnitName &"成功!",CurrentPageNow
	End If
	
End Sub

Sub SaveModify__photos( photos_id, photos_src, photos_orderid, product_id )
	'(特殊处理)
	If Trim(photos_id)="" And Trim(photos_src)="" And Trim(photos_orderid)="" Then Exit Sub
	
	'分析数组
	Dim array_photos_id,array_photos_src,array_photos_orderid
	Dim arrayCount,i__photos
	Dim RS_PHOTOS,SQL_PHOTOS
	Dim theTableName
	
	'处理参数，适应只有一个图片的情况.(特殊处理)
	If Instr(photos_id,",")<=0 Then photos_id=photos_id &",0"
	If Instr(photos_src,",")<=0 Then photos_src=photos_src &",0"
	If Instr(photos_orderid,",")<=0 Then photos_orderid=photos_orderid &",0"
	
	theTableName			="[CXBG_product__photos]"
	array_photos_id			=Split(photos_id,",")
	array_photos_src		=Split(photos_src,",")
	array_photos_orderid	=Split(photos_orderid,",")
	arrayCount				=Ubound(array_photos_id)
	
	
	
	Set RS_PHOTOS=Server.CreateObject("Adodb.RecordSet")
	For i__photos=0 To arrayCount
		'验证
		'(特殊处理)
		If Trim(array_photos_id(i__photos))="0" And Trim(array_photos_src(i__photos))="0" And Trim(array_photos_orderid(i__photos))="0" Then Exit For
		
		'验证排序号
		If Trim(array_photos_orderid(i__photos))="" Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>相册问题:请填写排序号！</li>"
		Else
			If Not isNumeric(Trim(array_photos_orderid(i__photos))) Then
				FoundErr=True
				ErrMsg=ErrMsg &"<br><li>相册问题:当前排序号不是数字(出现异常)！</li>"& Trim(array_photos_orderid(i__photos))
			End If
		End If
		'验证菜品id号
		If product_id="" Then
			FoundErr=True
			ErrMsg=ErrMsg &"<br><li>相册问题:缺少菜品id号，相册将无法指定菜品！</li>"
		Else
			If isNumeric(product_id) Then
				product_id=CokeShow.CokeClng(product_id)
			Else
				FoundErr=True
				ErrMsg=ErrMsg &"<br><li>相册问题:当前菜品id号不是数字(出现异常)！</li>"
			End If
		End If
		
		If FoundErr=True Then Exit Sub
		
		
		'新增新的相片操作.
		If Trim(array_photos_id(i__photos))="" Then
			
			'判断是否啥图也没上传！
			If Trim(array_photos_src(i__photos))<>"" And Len(Trim(array_photos_src(i__photos)))>10 Then
				SQL_PHOTOS="SELECT * FROM "& theTableName
				If Not IsObject(CONN) Then link_database
				RS_PHOTOS.Open SQL_PHOTOS,CONN,2,2
				RS_PHOTOS.AddNew
					
					RS_PHOTOS("photos_src")		=Trim(array_photos_src(i__photos))
					RS_PHOTOS("photos_orderid")	=CokeShow.CokeClng(Trim(array_photos_orderid(i__photos)))
					RS_PHOTOS("product_id")		=product_id		'*
				
				RS_PHOTOS.Update
				RS_PHOTOS.Close
				
			End If
			
			
		'修改已有的相片数据操作
		ElseIf isNumeric(Trim(array_photos_id(i__photos))) Then
			SQL_PHOTOS="SELECT * FROM "& theTableName &" WHERE product_id="& product_id &" AND id="& CokeShow.CokeClng(Trim(array_photos_id(i__photos)))
			If Not IsObject(CONN) Then link_database
			RS_PHOTOS.Open SQL_PHOTOS,CONN,1,3
			
			'拦截此记录的异常情况.
			If RS_PHOTOS.Bof And RS_PHOTOS.Eof Then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>相册问题:找不到指定要修改的菜品相册的图片！</li>"
				Exit Sub
			End If
			
			'看看有没有数据变更，如果没有则不动！
			If Not( RS_PHOTOS("photos_src")=Trim(array_photos_src(i__photos)) And RS_PHOTOS("photos_orderid")=CokeShow.CokeClng(Trim(array_photos_orderid(i__photos))) ) Then
					
				RS_PHOTOS("photos_src")		=Trim(array_photos_src(i__photos))
				RS_PHOTOS("photos_orderid")	=CokeShow.CokeClng(Trim(array_photos_orderid(i__photos)))
				
				
				RS_PHOTOS("modifydate")		=Now()
				
			End If
			
			RS_PHOTOS.Update
			RS_PHOTOS.Close
				
			
			
		End If
		
		
	Next
	
	
	Set RS_PHOTOS=Nothing
	
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
'给各个以逗号分隔的参数数字，都补上足够的零！
'参数:
'theNumbersStringNow:以逗号为分隔的id集合，或者是一个数字——如扩展分类product_class_id_extend.
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