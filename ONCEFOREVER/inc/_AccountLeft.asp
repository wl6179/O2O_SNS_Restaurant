	  <div class="leftclub_ht">
      
      
      
	    <div class="club_bt01"><span class="font16fff">我的帐号</span> <span class="font14000">资料</span><a class="more_01" href="/ONCEFOREVER/Account_PersonalInformation.Welcome"></a></div>
        <div class="hyxx">
		  <div class="yhtxxx">
            <img src="<% =Coke.ShowMemberSexPicURL(RS("id")) %>" width="100" height="100" />
			<div class="yhtxxx_right">
			  <p class="font_club14red"><% =RS("cnname") %></p>
			  <p>点评了：<span class="font_club12red"><% =RS("countRemarkOnTotal_Now") %></span></p>
			  <p>积分：<span class="font_club12red"><% =CokeShow.ChkAccountUserNameAllJifen(RS("username")) %></span></p>
			</div>	  
		  </div>
		  <ul class="hyxxul">
		  <li>
          	会员状态：&nbsp;<% If RS("account_level")=0 Then Response.Write "<span style=color:red;>未审核</span>" Else Response.Write CokeShow.otherField("[CXBG_account_class]",RS("account_level"),"classid","classname",True,0) %>
          </li>
          <li>
          	绑定餐厅会员卡状态：
		  	<br />
		  	<% If RS("isBindingVIPCardNumber")=1 Then %>
            	<img src="/images/hytx/card_01.jpg" /> 已绑定 卡号：<% =RS("BindingVIPCardNumber") %>
                <br />
                (<img src="/images/ico/small/coins_add.png" /> * 2 <span style="color:#FF5383;">双倍积分</span>)
			<% Else %>
            	<img src="/images/hytx/card_02.jpg" /> 未绑定会员卡号
                <br />
                (<img src="/images/ico/small/coins_add.png" /> * 1 <span style="color:#FF5383;">单倍积分</span>)
			<% End If %>
          </li>
          <li>性别：<% 'If isNumeric(RS("Sex")) Then Response.Write "ddd" %><%
		  Select Case RS("Sex")
		  Case 0
		  Response.Write "保密"
		  Case 1
		  Response.Write "女士"
		  Case 2
		  Response.Write "男士"
		  End Select
		  %></li>
		  <li>生日：<% =RS("Birthday") %><br />( <% If RS("now_day_num")<0 Then Response.Write "还有很久才到生日呢" Else Response.Write "还有<span style=""color:red;"">"& RS("now_day_num") &"</span>天就当寿星哦" %> )</li>
		  <li>注册日期：<br /><% =RS("adddate") %></li>
		  <li>上次登录时间：<br /><% =RS("lastlogintime") %></li>
		  <li>被浏览次数：<% =RS("iis") %> 次</li>
		  <li class="clubline"></li>
		  <li><a href="/ONCEFOREVER/Account_PersonalInformation.Welcome">完善个人资料</a>&nbsp;&nbsp;&nbsp;<a href="/ONCEFOREVER/Account_ModifyPassword.Welcome">修改密码</a></li>
		  </ul>
		</div>

	    <div class="club_bt"><span class="font16fff">我的帐号</span> <span class="font14000">管理项</span><!--<a class="more_01" href=""></a>--></div>
		  <ul class="zjfk">
            
            <li>
			  <span class="zjllimg"><img src="/images/ico/pencil_add.png" width="32" height="32" /></span>
			  <a <% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"account_index.welcome")>0 Then Response.Write "class=""xxxxxxx""" %> href="/ONCEFOREVER/Account_index.Welcome">我的点评</a>
			  <span class="zjname font_goy">我对菜品的点评记录</span>			
			</li>
            
            <li>
			  <span class="zjllimg"><img src="/images/ico/cart_add.png" width="32" height="32" /></span>
			  <a <% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"account_favorite.welcome")>0 Then Response.Write "class=""xxxxxxx""" %> href="/ONCEFOREVER/Account_Favorite.Welcome">我的收藏菜</a>
			  <span class="zjname font_goy">我平时的收藏记录</span>			
			</li>
            
            <li>
			  <span class="zjllimg"><img src="/images/ico/printer.png" width="32" height="32" /></span>
			  <a <% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"account_giftcertificated.welcome")>0 Then Response.Write "class=""xxxxxxx""" %> href="/ONCEFOREVER/Account_GiftCertificated.Welcome">我的礼品券</a>
			  <span class="zjname font_goy">我消费积分兑换来的礼品券</span>			
			</li>
            
            <li>
			  <span class="zjllimg"><img src="/images/ico/email_edit.png" width="32" height="32" /></span>
			  <a <% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"account_message.welcome")>0 Then Response.Write "class=""xxxxxxx""" %> href="/ONCEFOREVER/Account_Message.Welcome">我的站内信
			  <% If Coke.ShowReplyStatus( RS("username") )=True Then %>
              (<img src="/images/mailmanage.gif" />新)
			  <% End If %></a>
			  <span class="zjname font_goy">我和痴心不改的通讯</span>			
			</li>
            
            <li>
			  <span class="zjllimg"><img src="/images/ico/coins.png" width="32" height="32" /></span>
			  <a <% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"account_jifenhistory.welcome")>0 Then Response.Write "class=""xxxxxxx""" %> href="/ONCEFOREVER/Account_JifenHistory.Welcome">我的积分历史</a>
			  <span class="zjname font_goy">我的积分历史记录</span>			
			</li>
            
            <li>
			  <span class="zjllimg"><img src="/images/ico/creditcards.png" width="32" height="32" /></span>
			  <a <% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"account_bindingmyvipcard.welcome")>0 Then Response.Write "class=""xxxxxxx""" %> href="/ONCEFOREVER/Account_BindingMyVIPCard.Welcome">绑定会员卡卡号<% If RS("isBindingVIPCardNumber")=0 Then %><img src="/images/ico/hot.gif" /><% End If %></a>
			  <span class="zjname font_goy">我有餐厅会员卡赢双倍积分</span>			
			</li>
            
            <li>
			  <span class="zjllimg"><img src="/images/ico/door_out.png" width="32" height="32" /></span>
			  <a href="/ONCEFOREVER/LogOn.Welcome?Action=Logout" onclick="return confirm('您确定现在退出吗？');">安全退出</a>
			  <span class="zjname font_goy">正常退出登录状态</span>			
			</li>
		    
		  </ul>

	    <div class="club_bt"><span class="font16fff">我的</span> <span class="font14000">最近访客</span><a class="more_01" href="javascript:return false;" onClick="CheckDisplayAll('CheckDisplayAll1');"></a></div>
		  <ul class="zjfk">
		    <!--最近访客-->
		    <!--#include virtual="/_include/Club/_include_MemberInformation_Visited_Code.asp"-->
            <!--最近访客-->
		  </ul>
          
          
          
	   </div>
       
       
    <script type="text/javascript">
	//显示隐藏操作函数
	function CheckDisplayAll(elementIdName) {
		//
		alert("当访客超过8人时，将会列出更多访客列表");
		var checkbox_input_name = "www.cokeshow.com.cn";		//设置需要控制的选择框的id.
		if (dojo.byId(elementIdName).checked) {
			dojo.forEach(dojo.query("li[CokeShow='" + checkbox_input_name + "']"), function(x) {
				//x.setAttribute('display', '');
				dojo.style(x, {display:""});
				console.log(dojo.style(x,"display"));
			});
			dojo.byId(elementIdName).checked = false;
		}
		else {
			dojo.forEach(dojo.query("li[CokeShow='" + checkbox_input_name + "']"), function(x) {
				//x.setAttribute('display', 'none');
				dojo.style(x, {display:"none"});
				console.log(dojo.style(x,"display"));
			});
			dojo.byId(elementIdName).checked = true;
		}
	}
	</script>
