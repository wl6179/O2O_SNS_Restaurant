
		  <%
		  '最近访客.
		  Dim rsMemberInformation_Visited_Code,sqlMemberInformation_Visited_Code,countMemberInformation_Visited_Code,numMemberInformation_Visited_Code
		  	sqlMemberInformation_Visited_Code="SELECT TOP 38 *,DATEDIFF(ss,adddate,GETDATE()) AS now_second_num FROM [View_CXBG_account_IIS_AccountInformation] Where deleted=0 AND account_deleted=0 AND byvisit_Account_LoginID='"& RS("username") &"' AND Account_LoginID<>'"& RS("username") &"' ORDER BY id DESC"
			'被访问者是自己，而访问者不能出现自己.
			Set rsMemberInformation_Visited_Code=Server.CreateObject("Adodb.RecordSet")
			rsMemberInformation_Visited_Code.Open sqlMemberInformation_Visited_Code,CONN,1,1
			countMemberInformation_Visited_Code=rsMemberInformation_Visited_Code.RecordCount
			numMemberInformation_Visited_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsMemberInformation_Visited_Code.EOF Then
		  %>
		  <!--<li>
		    <p class="txtnoml lineheight28"><span class="fontred22">01</span> <a href="">欢迎光临</a></p>
		    <a class="abg01 display" href="#"><img src="/images/NoPic.png" width="110" height="110" /></a>
			<p class="xjcplxx">
            <img src="images/xx.gif" width="15" height="16" />
            <img src="images/xx.gif" width="15" height="16" />
            <img src="images/xx.gif" width="15" height="16" />
            <img src="images/xx.gif" width="15" height="16" />
            <img src="images/xx.gif" width="15" height="16" />
            </p>
		  </li>-->
		  <%
		  End If
		  %>
          
          <%
		  Do While Not rsMemberInformation_Visited_Code.EOF
		  %>
          	<li <% If numMemberInformation_Visited_Code>8 Then Response.Write "CokeShow=""www.cokeshow.com.cn"" style=""display:none;""" %>>
			  <span class="zjllimg"><img src="<% =Coke.ShowMemberSexPicURL(rsMemberInformation_Visited_Code("account_id")) %>" width="36" height="36" /><% If rsMemberInformation_Visited_Code("isBindingVIPCardNumber")=1 Then %><img src="/images/hytx/card_01.gif" width="10" height="7" /><% End If %></span>
			  <a class="zjname" href="/Club/MembersInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsMemberInformation_Visited_Code("account_id") ) %>" target="_blank"><% =rsMemberInformation_Visited_Code("cnname") %></a>
			  <span class="zjname font_goy"><% =CokeShow.SplitTime( rsMemberInformation_Visited_Code("now_second_num") ) %>到此访问</span>			
			</li>
          <%
			  numMemberInformation_Visited_Code=numMemberInformation_Visited_Code+1
			  rsMemberInformation_Visited_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsMemberInformation_Visited_Code.Close
		  Set rsMemberInformation_Visited_Code=Nothing
		  %>