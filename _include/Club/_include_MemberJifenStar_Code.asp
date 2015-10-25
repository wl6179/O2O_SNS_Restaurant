		  <%
		  '积分明星.
		  Dim rsMemberJifenStar_Code,sqlMemberJifenStar_Code,countMemberJifenStar_Code,numMemberJifenStar_Code
		  	sqlMemberJifenStar_Code="SELECT TOP 8 *,(select distinct sum(Jifen) over(partition by Account_LoginID) as sumJifen from [CXBG_account_JifenSystem] where deleted=0 and JifenWhichOperationRule='+' and Account_LoginID=[CXBG_account].username) as sumJifen_Now FROM [CXBG_account] Where deleted=0 ORDER BY sumJifen_Now DESC,id desc"
			Set rsMemberJifenStar_Code=Server.CreateObject("Adodb.RecordSet")
			rsMemberJifenStar_Code.Open sqlMemberJifenStar_Code,CONN,1,1
			countMemberJifenStar_Code=rsMemberJifenStar_Code.RecordCount
			numMemberJifenStar_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsMemberJifenStar_Code.EOF Then
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
		  Do While Not rsMemberJifenStar_Code.EOF
		  %>
          	<li>
			  <span class="nubmber"><img src="/images/<%
			  Select Case numMemberJifenStar_Code
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
			  <a class="name" href="/Club/MembersInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsMemberJifenStar_Code("id") ) %>" target="_blank"><% =rsMemberJifenStar_Code("cnname") %></a>
			  <span class="tximg"><img src="<% =Coke.ShowMemberSexPicURL(rsMemberJifenStar_Code("id")) %>" width="20" height="20" /><% If rsMemberJifenStar_Code("isBindingVIPCardNumber")=1 Then %><img src="/images/hytx/card_01.gif" width="10" height="7" /><% End If %></span>			
			</li>
          <%
			  numMemberJifenStar_Code=numMemberJifenStar_Code+1
			  rsMemberJifenStar_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsMemberJifenStar_Code.Close
		  Set rsMemberJifenStar_Code=Nothing
		  %>