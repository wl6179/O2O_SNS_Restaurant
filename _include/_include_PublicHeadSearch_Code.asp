
    <ul class="nav_munu">
	  <li <% If ShowNavigationNo=1 Then Response.Write "class=""navhover""" %>><a href="/">餐厅首页</a></li>
	  <li <% If ShowNavigationNo=2 Then Response.Write "class=""navhover""" %>><a href="/ChineseDish/ChineseDish.Welcome">点餐牌</a></li>
	  <li <% If ShowNavigationNo=3 Then Response.Write "class=""navhover""" %>><a href="/ChineseDish/ChineseDish.Welcome?StarRating=True">星级菜</a></li>
	  <li <% If ShowNavigationNo=4 Then Response.Write "class=""navhover""" %>><a href="/Details/Details.Welcome?classid=25&isHot=1">娱乐秀</a></li><!--<a href="/ChineseDish/ChineseDish.Welcome?classid=26&RoomService=True">娱乐秀</a>-->
	  <li <% If ShowNavigationNo=5 Then Response.Write "class=""navhover""" %>><a href="/DiningArea/DiningArea.Welcome">餐厅环境</a></li>
	  <li <% If ShowNavigationNo=6 Then Response.Write "class=""clubhover""" %>><a href="/Club/index.Welcome">会员Club</a></li>
	  <li <% If ShowNavigationNo=7 Then Response.Write "class=""navhover""" %>><a href="/Details/Details.Welcome?classid=4">联盟商家</a></li>
	</ul>
	<ul class="top_munu">
	  
      
      <!--<li><a class="more_fw" href=""></a></li>-->
      <li style="padding:6px 2px 0;">
      		<div dojoType="dijit.form.DropDownButton">
              <span>更多免费体验</span>
              <div dojoType="dijit.TooltipDialog" id="tooltipDlg" title="痴心不改餐厅欢迎您" style="display:none;">
                
                <table>
                  <tr>
                    <td><a href="/Details/DetailsInformation.Welcome?CokeMark=JPKNDNH" target="_blank" title="注册网站会员 - 永久免费的享受积分优惠、新品品尝机会哦">注册网站会员 - 永久免费的享受积分...</a> <img src="/images/ico/small/coins_add.png" /></td>
                  </tr>
                  <tr>
                    <td><a href="/Details/DetailsInformation.Welcome?CokeMark=JPKMENJ" target="_blank" title="绑定会员卡 - 永久享受双倍积分特权和评星级特权">绑定会员卡 - 永久享受双倍积分特权...</a> <img src="/images/ico/small/creditcards.png" /></td>
                  </tr>
                  <tr>
                    <td><a href="/Details/DetailsInformation.Welcome?CokeMark=JPJQKOD" target="_blank" title="兑换网站礼品券 - 永久免费独享新品尝鲜机会">兑换网站礼品券 - 永久免费独享新品...</a> <img src="/images/ico/small/printer_add.png" /></td>
                  </tr>
                </table>
                
              </div>
            </div>
      </li>
	  
      
      <li class="font10">|</li>
      <li><iframe width="63" height="24" frameborder="0" allowtransparency="true" marginwidth="0" marginheight="0" scrolling="no" frameborder="No" border="0" src="http://widget.weibo.com/relationship/followbutton.php?width=63&height=24&uid=1750355351&style=1&btn=red&dpc=1"></iframe></li>
      
	  <li class="font10">|</li>
	  <li><a href="#" onclick="addBookmark('<% =CokeShow.GetAllUrlII() %>','<% =PageTitleWords %>')" title="浏览器收藏">收藏</a></li>
      
      <li class="font10">|</li>
      <li>
      <!--<a href="/Details/DetailsInformation.Welcome?CokeMark=JPIQGNR" target="_blank" title="餐厅360°全景图展览">360°全景</a>-->
      <a href="javascript:return false;" onClick="ShowDialog_Search('<% ="http://"& Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL") &"?"& Request.ServerVariables("QUERY_STRING") %>');" style=" display: inline-block">搜索</a><img src="/images/new1.gif" />
      </li>
      
      <li class="font10">|</li>
      <li><a href="/ONCEFOREVER/Account_BindingMyVIPCard.Welcome" title="绑定会员卡卡号">绑定</a></li>
      
      <% If CokeShow.CheckUserLogined()=True And isNumeric(Session("id")) And Len(Session("username"))>=10 Then %>
      <li class="font10">|</li>
	  <li><a href="/ONCEFOREVER/LogOn.Welcome?Action=Logout&fromurl=<% =CokeShow.EncodeURL( CokeShow.GetAllUrlII,"" ) %>" onclick="return confirm('您确定现在退出吗？');" title="安全退出">退出</a></li>
      
      <li class="font10">|</li>
      <li><a class="more_fw" href="/ONCEFOREVER/"></a></li>
      <li><a href="/ONCEFOREVER/"><span style="color: #F30; font-weight:bold;"><% =Session("cnname") %></span></a></li>
      
      <% Else %>
      <li class="font10">|</li>
	  <li><a href="/ONCEFOREVER/AccedeToRegiste.Welcome?fromurl=<% =CokeShow.EncodeURL( CokeShow.GetAllUrlII,"" ) %>">免费注册</a></li>
      
      <li class="font10">|</li>
      <li><a href="/ONCEFOREVER/LogOn.Welcome?fromurl=<% =CokeShow.EncodeURL( CokeShow.GetAllUrlII,"" ) %>">登陆</a></li>
      <% End If %>
      
      
	  
	</ul>

  <%' =CokeShow.SayHello %>