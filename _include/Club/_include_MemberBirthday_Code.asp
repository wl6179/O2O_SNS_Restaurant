		  <%
		  '今日寿星.
		  Dim rsMemberBirthday_Code,sqlMemberBirthday_Code,countMemberBirthday_Code,numMemberBirthday_Code
		  	sqlMemberBirthday_Code="SELECT TOP 8 *,DATEDIFF(day,GETDATE(),dateadd(year,year(GETDATE())-year(Birthday),Birthday)) AS now_day_num FROM [CXBG_account] Where deleted=0 AND (dateadd(year,year(GETDATE())-year(Birthday),Birthday) between GETDATE() and DATEADD(month,1,GETDATE()) or dateadd(year,year(DATEADD(month,1,GETDATE()))-year(Birthday),Birthday) between GETDATE() and DATEADD(month,1,GETDATE())) ORDER BY now_day_num ASC,id DESC"
			Set rsMemberBirthday_Code=Server.CreateObject("Adodb.RecordSet")
			rsMemberBirthday_Code.Open sqlMemberBirthday_Code,CONN,1,1
			countMemberBirthday_Code=rsMemberBirthday_Code.RecordCount
			numMemberBirthday_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsMemberBirthday_Code.EOF Then
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
		  Do While Not rsMemberBirthday_Code.EOF
		  %>
          	<li>
			  <span class="tximg">
               
               	  <% If rsMemberBirthday_Code("now_day_num")=0 Then %>
                  		<img src="/images/jrsx_img01.jpg" width="22" height="22" /><% If rsMemberBirthday_Code("isBindingVIPCardNumber")=1 Then %><img src="/images/hytx/card_01.gif" width="10" height="7" /><% End If %>
                  <% Else %>
                  		<img src="<% =Coke.ShowMemberSexPicURL(rsMemberBirthday_Code("id")) %>" width="20" height="20" /><% If rsMemberBirthday_Code("isBindingVIPCardNumber")=1 Then %><img src="/images/hytx/card_01.gif" width="10" height="7" /><% End If %>
                  <% End If %>
               
              </span>
			  <span class="txt_jrsx">
                  <% If rsMemberBirthday_Code("now_day_num")=0 Then %>
                  今日
                  <% Else %>
                  <% =rsMemberBirthday_Code("now_day_num") %>天后
                  <% End If %><a class="name" href="/Club/MembersInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsMemberBirthday_Code("id") ) %>" target="_blank"><% =rsMemberBirthday_Code("cnname") %></a>
                  生日哦！
              </span>
			</li>
          <%
			  numMemberBirthday_Code=numMemberBirthday_Code+1
			  rsMemberBirthday_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsMemberBirthday_Code.Close
		  Set rsMemberBirthday_Code=Nothing
		  %>