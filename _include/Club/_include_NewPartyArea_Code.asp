		  <%
		  'club最新Party区域.
		  Dim rsNewPartyArea_Code,sqlNewPartyArea_Code,countNewPartyArea_Code,numNewPartyArea_Code
		  	sqlNewPartyArea_Code="SELECT TOP 3 * FROM [CXBG_details] WHERE deleted=0 AND isOnpublic=1 AND details_class_id=11 ORDER BY isRecommend desc,details_orderid DESC,id DESC"
			Set rsNewPartyArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsNewPartyArea_Code.Open sqlNewPartyArea_Code,CONN,1,1
			countNewPartyArea_Code=rsNewPartyArea_Code.RecordCount
			numNewPartyArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsNewPartyArea_Code.EOF Then
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
		  Do While Not rsNewPartyArea_Code.EOF
		  %>
          <li>
		    <p class="txtnoml lineheight28"><span class="fontred22">0<% =numNewPartyArea_Code %></span> <a href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsNewPartyArea_Code("id") ) %>" target="_blank" title="<% =rsNewPartyArea_Code("topic") %>"><% =rsNewPartyArea_Code("topic") %></a></p>
		    <a class="abg01 display" href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsNewPartyArea_Code("id") ) %>" target="_blank" title="<% =rsNewPartyArea_Code("topic") %>"><img src="<% If rsNewPartyArea_Code("photo")<>"" Then Response.Write rsNewPartyArea_Code("photo") Else Response.Write "/images/NoPic.png" %>" width="110" height="110" /></a>
			<!--<p class="xjcplxx">
            
            <ul class="rating" style="margin-top:8px; margin-left:0px;">
            <li class="current-rating" style="width:<% =Coke.ShowProductStarRating_Num(rsNewPartyArea_Code("id")) * 20 + 1 %>px;"></li>
            
            </ul>
            
            </p>-->
		  </li>
          <%
			  numNewPartyArea_Code=numNewPartyArea_Code+1
			  rsNewPartyArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsNewPartyArea_Code.Close
		  Set rsNewPartyArea_Code=Nothing
		  %>