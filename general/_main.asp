<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True	'缓冲数据，才向客户输出；
Response.Charset="utf-8"
Session.CodePage = 65001
%>

<!--#include file="../system/system_conn.asp"-->
<!--#include file="../system/system_class.asp"-->

<%
'系统类实例化
Dim CokeShow
Set CokeShow = New SystemClass
CokeShow.Start
Call CokeShow.SQLWarningSys()	'预警.
%>

<%
'判断管理员身份.
Dim SupervisorLevelString
If isNumeric(Session("enterLevel")) Then
	If CokeShow.CokeClng(Session("enterLevel"))>0 Then
		SupervisorLevelString=CokeShow.otherField("[CXBG_supervisor_class]",CokeShow.CokeClng(Session("enterLevel")),"classid","classname",True,0)
	Else
		SupervisorLevelString=""
	End If
Else
	SupervisorLevelString=""
End If
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>痴心不改餐厅-大后台登录</title>
	
	<link type="text/css" rel="stylesheet" href="<% =filename_dj_MainCss %>" />
	<link type="text/css" rel="stylesheet" href="<% =filename_dj_ThemesCss %>" /><!--tundra-->
	
	<link type="text/css" rel="stylesheet" href="../style/general_style.css" />
	
	<script type="text/javascript" src="../script/public.js"></script>
	
	<script type="text/javascript" src="<% =filename_dj %>" djConfig="parseOnLoad:<% =parseOnLoad_dj %>, isDebug:<% =isDebug_dj %>, debugAtAllCosts:<% =isDebug_dj %>"></script>
	<script type="text/javascript" src="<% =filenameWidgetsCompress_dj %>"></script>
	<script type="text/javascript">
		dojo.require("dojo.parser");
//		dojo.require("dijit.form.ValidationTextBox");
		dojo.require("dijit.form.Button");
//		dojo.require("dijit.form.Form");
//		
		dojo.require("dijit.layout.BorderContainer");
		dojo.require("dijit.layout.ContentPane");
		
		dojo.require("dijit.Menu");
   		dojo.require("dijit.Toolbar");
		
		dojo.require("dijit.TitlePane");
		//dojo.require("dojox.widget.Toaster");
		dojo.require("dijit.Dialog");
		
		
		dojo.addOnLoad(function(){
		//	setTopDisplay();
		});
		
		function setTopDisplay() {
		//	dojo.addClass(dojo.byId("top"), "DisplayNone");
		}
	</script>
	
	<script type="text/javascript" src="../script/imgpngPrecc.js"></script>
	
	<style type="text/css">
		a:hover {
		color: #ffffff;
		}
		.DisplayNone {
		display: none;
		}
	</style>
</head>
<body class="<% =classname_dj_ThemesCss %>">
	
	<div dojoType="dijit.layout.BorderContainer" design="headline"
	style="width:100%; height:100%; margin:0px; padding:0px;" liveSizing="true"
	><!--design=headline标题行样式.liveSizing即时重绘.-->
		
		<!--top b-->
		<div dojoType="dijit.layout.ContentPane" region="top" id="top"
		style="height:; margin:0px; padding:0px;"
		>
			
			<div style=" margin:0px; padding:0px; font-size:12px; text-align:left; background:url(../images/logo_bg.jpg)">
				
				<img src="/images/logo.jpg" />
				&nbsp;
				
				
			</div>
			
			
			
			<div id="top_toolbar"
			dojoType="dijit.Toolbar"
			style="height:33px; line-height:33px;"
			>
			<%
			'检测是否已经登录.
			If Session("enterName")<>"" And Len(Session("enterName"))>3 Then
			%>
				
				<span style=" font-size:14px; font-weight:bold;">
					&nbsp;
					您好，<% =Session("enterName") %>！&nbsp;身份验证:<% =SupervisorLevelString %>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				</span>
				<div dojoType="dijit.form.DropDownButton">
					
					<span>菜品管理</span>
					<div dojoType="dijit.Menu" style="display: none; font-size:12px; width:200px;">
						<div dojoType="dijit.MenuItem">
							菜品高级查询
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "product.asp?Action=SearchNow";
							</script>
						</div>
						<!--<div dojoType="dijit.MenuItem">
							<img src="/images/tree_folder4.gif" width="16" />&nbsp;新增菜品
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "product.asp?Action=Add";
							</script>
						</div>-->
						<div dojoType="dijit.MenuItem">
							菜品管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "product.asp";
							</script>
						</div>
						<div dojoType="dijit.MenuSeparator"></div>
						<!--<div dojoType="dijit.MenuItem">
							<img src="/images/tree_folder4.gif" width="16" />&nbsp;新增菜品分类
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "product_class.asp?Action=Add";
							</script>
						</div>-->
						<div dojoType="dijit.MenuItem">
							菜品分类管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "product_class.asp";
							</script>
						</div>
						
						
						<!--可选项类别 begin-->
						<div dojoType="dijit.MenuSeparator"></div>
						<!--<div dojoType="dijit.MenuItem">
							<img src="/images/tree_folder4.gif" width="16" />&nbsp;新增所属菜系
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "product_businessUSE.asp?Action=Add";
							</script>
						</div>-->
						<div dojoType="dijit.MenuItem">
							菜品菜系管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "product_businessUSE.asp";
							</script>
						</div>
						<!--<div dojoType="dijit.MenuItem">
							<img src="/images/tree_folder4.gif" width="16" />&nbsp;新增所属口味
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "product_activityUSE.asp?Action=Add";
							</script>
						</div>-->
						<div dojoType="dijit.MenuItem">
							菜品口味管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "product_activityUSE.asp";
							</script>
						</div>
						<div dojoType="dijit.MenuItem">
							辣椒指数管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "product_chiliIndex.asp";
							</script>
						</div>
						<!--可选项类别 end-->
						
						
					</div>
					
				</div>
				<div dojoType="dijit.form.DropDownButton">
					
					<span>内容管理</span>
					<div dojoType="dijit.Menu" style="display: none; font-size:12px; width:200px;">
						
						<!--<div dojoType="dijit.MenuItem">
							<img src="/images/tree_folder4.gif" width="16" />&nbsp;新增内容
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "details_edit.asp?Action=Add";
							</script>
						</div>-->
						<div dojoType="dijit.MenuItem">
							内容管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "details_edit.asp";
							</script>
						</div>
						<div dojoType="dijit.MenuSeparator"></div>
						<!--<div dojoType="dijit.MenuItem">
							<img src="/images/tree_folder4.gif" width="16" />&nbsp;新增内容分类
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "details_class.asp?Action=Add";
							</script>
						</div>-->
						<div dojoType="dijit.MenuItem">
							内容分类管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "details_class.asp";
							</script>
						</div>
						
						<!--可选项类别 end-->
					</div>
					
				</div>
                <div dojoType="dijit.form.DropDownButton">
					
					<span>餐厅环境管理</span>
					<div dojoType="dijit.Menu" style="display: none; font-size:12px; width:200px;">
						
						<div dojoType="dijit.MenuItem">
							餐厅环境图片管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "DiningArea_edit.asp";
							</script>
						</div>
						<div dojoType="dijit.MenuSeparator"></div>
						<div dojoType="dijit.MenuItem">
							餐厅环境图片分类管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "DiningArea_class.asp";
							</script>
						</div>
						
						<!--可选项类别 end-->
					</div>
					
				</div>
				<div dojoType="dijit.form.DropDownButton">
					
					<span>广告管理</span>
					<div dojoType="dijit.Menu" style="display: none; font-size:12px;">
						
						<div dojoType="dijit.MenuItem">
							首页广告发布管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "advertisement_main.asp";
							</script>
						</div>
                        <div dojoType="dijit.MenuItem">
							会员Club广告发布管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "advertisement_club.asp";
							</script>
						</div>
						<!--<div dojoType="dijit.MenuItem">
							首页专题广告管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "advertisement_featured.asp";
							</script>
						</div>
						<div dojoType="dijit.MenuItem">
							热门搜索管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "HotKeywords.asp";
							</script>
						</div>-->
						
						
					</div>
					
				</div>
				<div dojoType="dijit.form.DropDownButton">
					
					<span>娱乐项目发布管理</span>
					<div dojoType="dijit.Menu" style="display: none; font-size:12px;">
						
						<div dojoType="dijit.MenuItem">
							游戏发布管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "Game_edit.asp";
							</script>
						</div>
						<div dojoType="dijit.MenuItem">
							礼品券发布管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "GiftCertificated_edit.asp";
							</script>
						</div>
						<div dojoType="dijit.MenuItem">
							调查问卷管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "Questionnaire_edit.asp";
							</script>
						</div>
						
					</div>
					
				</div>
				<div dojoType="dijit.form.DropDownButton">
					
					<span>(业务)互动管理</span>
					<div dojoType="dijit.Menu" style="display: none; font-size:12px;">
						
                        
						<div dojoType="dijit.MenuItem">
							注册会员帐号管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "account.asp";
							</script>
						</div>
						<div dojoType="dijit.MenuItem">
							会员留言管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "Message.asp";
							</script>
						</div>
						<div dojoType="dijit.MenuItem">
							推荐朋友列表
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "tuijianpengyou.asp";
							</script>
						</div>
                        <div dojoType="dijit.MenuItem">
							最新会员点评
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "Account_RemarkOns.asp";
							</script>
						</div>
                        <div dojoType="dijit.MenuItem">
							最新会员菜品收藏
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "Account_Favorites.asp";
							</script>
						</div>
                        <div dojoType="dijit.MenuItem">
							已兑换礼品券
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "Account_GiftCertificateds.asp";
							</script>
						</div>
                        <div dojoType="dijit.MenuItem">
							问卷调查结果
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "Questionnaire_ResultView.asp";
							</script>
						</div>
                        
                        <div dojoType="dijit.MenuSeparator"></div>
                        
                        <div dojoType="dijit.MenuItem">
							友情链接管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "FriendlyLink.asp";
							</script>
						</div>
                        <div dojoType="dijit.MenuItem">
							VIP卡卡号录入管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "VIPcardList.asp";
							</script>
						</div>
					</div>
					
				</div>
				<div dojoType="dijit.form.DropDownButton">
					<span>系统管理<!--menu01_1--></span>
					<div dojoType="dijit.Menu" style="display: none; font-size:12px;">
						
						<div dojoType="dijit.MenuItem">
							<img src="/images/database_exclamation.png" width="16" />&nbsp;网站资料设置
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "controller.asp";
							</script>
						</div>
						<div dojoType="dijit.MenuItem">
							管理员帐号管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "supervisor.asp";
							</script>
						</div>
						<div dojoType="dijit.MenuItem">
							管理员帐号分类设置
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "supervisor_class.asp";
							</script>
						</div>
						
						<!--<div dojoType="dijit.MenuItem">
							订单状态设置
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "Order_Status_Set.asp";
							</script>
						</div>-->
						<div dojoType="dijit.MenuItem">
							会员帐号等级设置
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "account_class.asp";
							</script>
						</div>
                        <div dojoType="dijit.MenuItem">
							属性——职业管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "attribute_work.asp";
							</script>
						</div>
                        <div dojoType="dijit.MenuItem">
							属性——学历管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "attribute_schooling.asp";
							</script>
						</div>
                        <div dojoType="dijit.MenuItem">
							属性——收入管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "attribute_income.asp";
							</script>
						</div>
                        
                        
                        <div dojoType="dijit.MenuItem">
							积分名目管理(禁)
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "attribute_jifensystem.asp";
							</script>
						</div>
                        
                        
						<% If Session("enterName")="coke" Then %>
						<div dojoType="dijit.MenuItem">
							日志管理
							<script type="dojo/method" event="onClick">
								dojo.byId("_mainFrame").src = "log.asp";
							</script>
						</div>
						<% End If %>
                        
                        
                        <div dojoType="dijit.MenuItem">
							<img src="/images/no_1.png" width="16" />&nbsp;退出系统
							<script type="dojo/method" event="onClick">
								if(confirm("您确定安全退出吗?")) {
									window.location.replace("enter.asp?Action=Logout");
								}
							</script>
						</div>
					</div>
					
				</div>
				<div dojoType="dijit.form.DropDownButton">
					
					<span>CokeShow 帮助</span>
					<div dojoType="dijit.Menu" style="display: none; font-size:12px;">
						
						
						<div dojoType="dijit.MenuItem">
							<img src="/images/favicon.ico" width="16" />&nbsp;访问可乐秀CokeShow首页
							<script type="dojo/method" event="onClick">
								//window.open("http://www.cokeshow.com.cn/");
								dojo.byId("_mainFrame").src = "http://www.cokeshow.com.cn/";
							</script>
						</div>
						<div dojoType="dijit.MenuItem">
							<img src="/images/umbrella_2.png" width="16" />&nbsp;操作系统版本：Version 3.20100501
							<script type="dojo/method" event="onClick">
								ShowDialog('<span style=color:black;>感谢您的支持</span>','','width:300px;height:200px;','<div style=background-color:#ffffff;width:300px;height:300px;>当前操作系统版本：<br />Version 3.20100501<br /><br />系统说明：<br />3.x系列属于最高端的会员级营销系统<br />2.x系列属于功能型的商城销售类系统<br />1.x属于高端的多国语公司展示系统</div>');
							</script>
						</div>
						<div dojoType="dijit.MenuItem">
							<img src="/images/QQII.png" width="16" />&nbsp;技术支持电话：010-67659219
							<script type="dojo/method" event="onClick">
								ShowDialog('<span style=color:black;>感谢您的支持</span>','','width:300px;height:200px;','<div style=background-color:#ffffff;width:300px;height:300px;>技术支持电话：010-67659219</div>');
							</script>
						</div>
						<div dojoType="dijit.MenuItem">
							<img src="/images/QQII.png" width="16" />&nbsp;技术支持QQ：595574668
							<script type="dojo/method" event="onClick">
								ShowDialog('<span style=color:black;>感谢您的支持</span>','','width:300px;height:200px;','<div style=background-color:#ffffff;width:300px;height:300px;>技术支持QQ：595574668</div>');
							</script>
						</div>
						
						
						
					</div>
					
				</div>
				<!--<div dojoType="dijit.form.DropDownButton">
					
					<span>退出系统</span>
					<div dojoType="dijit.Menu" style="display: none; font-size:12px;">
						
						<div dojoType="dijit.MenuItem">
							<img src="/images/no_1.png" width="16" />&nbsp;退出系统
							<script type="dojo/method" event="onClick">
								if(confirm("您确定退出吗?")) {
									window.location.replace("enter.asp?Action=Logout");
								}
							</script>
						</div>
						
						
					</div>
					
				</div>-->
				
				
				
				
			<%
			End If
			%>
			</div>
			
			
		</div>
		<!--top e-->
		
		
		
		<!--center b-->
		<div dojoType="dijit.layout.ContentPane" region="center"
		style=" margin:0px; padding:0px; border:0px;"
		>
			<iframe name="_mainFrame" id="_mainFrame" height="100%" width="100%" border="0" frameborder="0" src="controller.asp" scrolling="auto"></iframe>
		</div>
		<!--center e-->
		
		
		
		<!--bottom b-->
		<div dojoType="dijit.layout.ContentPane" region="bottom"
		style="height:; margin:0px; padding:0px; color:#999999; font-size: 12px;"
		>
			
			
			<div id="bottom_toolbar"
			dojoType="dijit.Toolbar"
			style="height:33px; line-height:33px;"
			>
					
					
					
					<div id="bottom_toolbar.a"
					dojoType="dijit.form.Button"
					>
						版权所有:痴心不改餐厅chixinbugai.me
						<script type="dojo/method" event="onClick">
							window.open("http://www.chixinbugai.me/");
						</script>
					</div>
					&nbsp;&nbsp;|&nbsp;&nbsp;公司电话:010-88888888
					&nbsp;&nbsp;|&nbsp;&nbsp;传真:010-66666666
					
					
					
			</div>
			
			
		</div>
		<!--bottom e-->
		
	</div>
</body>
</html>