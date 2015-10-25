		  <%
		  '会员评论.
		  Dim rsMemberCommentsArea_Code,sqlMemberCommentsArea_Code,countMemberCommentsArea_Code,numMemberCommentsArea_Code
		  	sqlMemberCommentsArea_Code="select top 5 * from [View_RemarkOn_AccountInfor] where deleted=0 AND account_deleted=0 ORDER BY id DESC"
			Set rsMemberCommentsArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsMemberCommentsArea_Code.Open sqlMemberCommentsArea_Code,CONN,1,1
			countMemberCommentsArea_Code=rsMemberCommentsArea_Code.RecordCount
			numMemberCommentsArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsMemberCommentsArea_Code.EOF Then
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
		  Dim i_starrating
		  
		  Do While Not rsMemberCommentsArea_Code.EOF
		  %>
            <li>
		     <a class="hypl_linetxt" href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsMemberCommentsArea_Code("product_id") ) %>#RemarkOnArea_start" target="_blank"><% =rsMemberCommentsArea_Code("logtext") %></a>
			 <p class="hypl_hyp">
			   <span class="waith40">
               <img src="<% =Coke.ShowMemberSexPicURL(rsMemberCommentsArea_Code("account_id")) %>" width="20" height="20" />
               <% If rsMemberCommentsArea_Code("isBindingVIPCardNumber")=1 Then %><img src="/images/hytx/card_01.gif" width="10" height="7" /><% End If %>
               </span>
			   <span class="waith80 linheight22"><a class="fontred_a" href="/Club/MembersInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsMemberCommentsArea_Code("account_id") ) %>" target="_blank"><% =rsMemberCommentsArea_Code("cnname") %></a></span>
			   <span class="waith100"><%
				'输出星级.
				'Dim i_starrating
				For i_starrating=1 To rsMemberCommentsArea_Code("theStarRatingForChineseDishInformation")
				%><img src="/images/xx.gif" width="16" height="16" /><%
				Next
				%></span>
			   <span class="waith50"><a class="button_img47" href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsMemberCommentsArea_Code("product_id") ) %>#RemarkOnArea_start" target="_blank">查看</a></span>			 				<!--<span class="waith50"><a class="button_img47" href="javascript:return false;" onclick="ShowDialog('<span style=color:black;>点评</span>','/_include/Public/View_RemarkOnAndProductInfor.Welcome?id=<% =rsMemberCommentsArea_Code("id") %>','width:500px;height:400px;','');">查看</a></span>	-->
			 </p>
		   </li>
          <%
			  numMemberCommentsArea_Code=numMemberCommentsArea_Code+1
			  rsMemberCommentsArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsMemberCommentsArea_Code.Close
		  Set rsMemberCommentsArea_Code=Nothing
		  %>
		   