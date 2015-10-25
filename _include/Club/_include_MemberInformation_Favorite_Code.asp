
		  <%
		  '最近访客.
		  Dim rsMemberInformation_Favorite_Code,sqlMemberInformation_Favorite_Code,countMemberInformation_Favorite_Code,numMemberInformation_Favorite_Code
		  	sqlMemberInformation_Favorite_Code="SELECT TOP 20 * FROM [View_CXBG_Favorite_ProductInformation] Where deleted=0 AND Account_LoginID='"& RS("username") &"' ORDER BY id DESC"
			'被访问者是自己，而访问者不能出现自己.
			Set rsMemberInformation_Favorite_Code=Server.CreateObject("Adodb.RecordSet")
			rsMemberInformation_Favorite_Code.Open sqlMemberInformation_Favorite_Code,CONN,1,1
			countMemberInformation_Favorite_Code=rsMemberInformation_Favorite_Code.RecordCount
			numMemberInformation_Favorite_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsMemberInformation_Favorite_Code.EOF Then
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
		  Do While Not rsMemberInformation_Favorite_Code.EOF
		  %>
          	<li>
            	<span class="zjscimg"><img src="/images/jrscico.jpg" width="16" height="16" /></span>
                <span class="zjsc_txt">收藏了<a class="fontred_a" href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsMemberInformation_Favorite_Code("product_id") ) %>" target="_blank" title="<% =rsMemberInformation_Favorite_Code("ProductName") %>"><% =rsMemberInformation_Favorite_Code("ProductName") %></a></span>
                <span class="zjscrq"><% =Month( rsMemberInformation_Favorite_Code("adddate") ) %>月<% =Day( rsMemberInformation_Favorite_Code("adddate") ) %></span>
            </li>
          <%
			  numMemberInformation_Favorite_Code=numMemberInformation_Favorite_Code+1
			  rsMemberInformation_Favorite_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsMemberInformation_Favorite_Code.Close
		  Set rsMemberInformation_Favorite_Code=Nothing
		  %>