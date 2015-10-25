		  <%
		  '欢迎 俱乐部新会员.
		  Dim rsGames_Code,sqlGames_Code,countGames_Code,numGames_Code
		  	sqlGames_Code="SELECT TOP 6 * FROM [CXBG_Game] Where deleted=0 AND isOnpublic=1 ORDER BY details_orderid DESC,id desc"
			Set rsGames_Code=Server.CreateObject("Adodb.RecordSet")
			rsGames_Code.Open sqlGames_Code,CONN,1,1
			countGames_Code=rsGames_Code.RecordCount
			numGames_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsGames_Code.EOF Then
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
		  Do While Not rsGames_Code.EOF
		  %>
        <li>
        	<a class="hyyx_txt" href="/Club/GamesInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsGames_Code("id") ) %>" target="_blank"><% =rsGames_Code("topic") %></a><a class="hyyx_go" href="/Club/GamesInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsGames_Code("id") ) %>" target="_blank"></a>
        </li>
          <%
			  numGames_Code=numGames_Code+1
			  rsGames_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsGames_Code.Close
		  Set rsGames_Code=Nothing
		  %>