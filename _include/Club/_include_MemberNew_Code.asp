		  <%
		  '欢迎 俱乐部新会员.
		  Dim rsMemberNew_Code,sqlMemberNew_Code,countMemberNew_Code,numMemberNew_Code
		  	sqlMemberNew_Code="SELECT TOP 8 * FROM [CXBG_account] Where deleted=0 ORDER BY id DESC"
			Set rsMemberNew_Code=Server.CreateObject("Adodb.RecordSet")
			rsMemberNew_Code.Open sqlMemberNew_Code,CONN,1,1
			countMemberNew_Code=rsMemberNew_Code.RecordCount
			numMemberNew_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsMemberNew_Code.EOF Then
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
		  Do While Not rsMemberNew_Code.EOF
		  %>
          	<li>
			  <span class="nubmber"><img src="/images/<%
			  Select Case numMemberNew_Code
			  Case 1
			  Response.Write "one"
			  Case 2
			  Response.Write "two"
			  Case 3
			  Response.Write "three"
			  Case 4
			  Response.Write "four"
			  Case 5
			  Response.Write "five"
			  Case 6
			  Response.Write "six"
			  Case 7
			  Response.Write "seven"
			  Case 8
			  Response.Write "eight"
			  End Select
			  %>.jpg" width="18" height="14" /></span>
			  <a class="name" href="/Club/MembersInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsMemberNew_Code("id") ) %>" target="_blank"><% =rsMemberNew_Code("cnname") %></a>
			  <span class="tximg"><img src="<% =Coke.ShowMemberSexPicURL(rsMemberNew_Code("id")) %>" width="20" height="20" /><% If rsMemberNew_Code("isBindingVIPCardNumber")=1 Then %><img src="/images/hytx/card_01.gif" width="10" height="7" /><% End If %></span>			
			</li>
          <%
			  numMemberNew_Code=numMemberNew_Code+1
			  rsMemberNew_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsMemberNew_Code.Close
		  Set rsMemberNew_Code=Nothing
		  %>