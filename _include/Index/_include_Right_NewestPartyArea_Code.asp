	    <%
		  '最新Party.
		  Dim rsRight_NewestPartyArea_Code,sqlRight_NewestPartyArea_Code,countRight_NewestPartyArea_Code,numRight_NewestPartyArea_Code
		  	sqlRight_NewestPartyArea_Code="select top 1 * from [CXBG_details] where deleted=0 and isOnpublic=1 and details_class_id=11 order by isRecommend desc,details_orderid desc,id desc"
			Set rsRight_NewestPartyArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsRight_NewestPartyArea_Code.Open sqlRight_NewestPartyArea_Code,CONN,1,1
			countRight_NewestPartyArea_Code=rsRight_NewestPartyArea_Code.RecordCount
			numRight_NewestPartyArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsRight_NewestPartyArea_Code.EOF Then
		  %>
		  <!--<li><a href="">热门信息张维迎：痴心不改餐厅隆重开业</a></li>-->
		  <%
		  End If
		  %>
          
          <%
		  Do While Not rsRight_NewestPartyArea_Code.EOF
		  %>
        
        
        <a href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsRight_NewestPartyArea_Code("id") ) %>" target="_blank"><img src="<% If rsRight_NewestPartyArea_Code("photo")<>"" Then Response.Write rsRight_NewestPartyArea_Code("photo") Else Response.Write "/images/NoPic.png" %>" width="200" height="130" /></a>
		<p><a class="fontred_a linheight22" href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsRight_NewestPartyArea_Code("id") ) %>" target="_blank" title="<% =rsRight_NewestPartyArea_Code("topic") %>"><% =rsRight_NewestPartyArea_Code("topic") %></a></p>
		<p><a class="linheight22" href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsRight_NewestPartyArea_Code("id") ) %>" target="_blank"><% =Left(rsRight_NewestPartyArea_Code("photo_desc"),48) %>...</a></p>
        <%
			  numRight_NewestPartyArea_Code=numRight_NewestPartyArea_Code+1
			  rsRight_NewestPartyArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsRight_NewestPartyArea_Code.Close
		  Set rsRight_NewestPartyArea_Code=Nothing
		  %>