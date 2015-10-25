		  <%
		  'club娱乐秀区域.
		  Dim rsYuLeXiuArea_Code,sqlYuLeXiuArea_Code,countYuLeXiuArea_Code,numYuLeXiuArea_Code
		  	sqlYuLeXiuArea_Code="SELECT TOP 3 * FROM [CXBG_details] WHERE deleted=0 AND isOnpublic=1 AND details_class_id=25 ORDER BY isRecommend desc,details_orderid DESC,id DESC"
			Set rsYuLeXiuArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsYuLeXiuArea_Code.Open sqlYuLeXiuArea_Code,CONN,1,1
			countYuLeXiuArea_Code=rsYuLeXiuArea_Code.RecordCount
			numYuLeXiuArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsYuLeXiuArea_Code.EOF Then
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
		  Do While Not rsYuLeXiuArea_Code.EOF
		  %>
          <li>
		    <p class="txtnoml lineheight28"><span class="fontred22">0<% =numYuLeXiuArea_Code %></span> <a href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsYuLeXiuArea_Code("id") ) %>" target="_blank" title="<% =rsYuLeXiuArea_Code("topic") %>"><% =rsYuLeXiuArea_Code("topic") %></a></p>
		    <a class="abg01 display" href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsYuLeXiuArea_Code("id") ) %>" target="_blank" title="<% =rsYuLeXiuArea_Code("topic") %>"><img src="<% If rsYuLeXiuArea_Code("photo")<>"" Then Response.Write rsYuLeXiuArea_Code("photo") Else Response.Write "/images/NoPic.png" %>" width="110" height="110" /></a>
			<!--<p class="xjcplxx">
            
            <ul class="rating" style="margin-top:8px; margin-left:0px;">
            <li class="current-rating" style="width:<% =Coke.ShowProductStarRating_Num(rsYuLeXiuArea_Code("id")) * 20 + 1 %>px;"></li>
            
            </ul>
            
            </p>-->
		  </li>
          <%
			  numYuLeXiuArea_Code=numYuLeXiuArea_Code+1
			  rsYuLeXiuArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsYuLeXiuArea_Code.Close
		  Set rsYuLeXiuArea_Code=Nothing
		  %>