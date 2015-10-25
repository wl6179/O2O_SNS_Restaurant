	    <%
		  '最新动态.
		  Dim rsRight_NewestNewsArea_Code,sqlRight_NewestNewsArea_Code,countRight_NewestNewsArea_Code,numRight_NewestNewsArea_Code
		  	sqlRight_NewestNewsArea_Code="select top 7 * from [CXBG_details] where deleted=0 and isOnpublic=1 and details_class_id<>999999999 order by details_orderid desc,id desc"
			Set rsRight_NewestNewsArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsRight_NewestNewsArea_Code.Open sqlRight_NewestNewsArea_Code,CONN,1,1
			countRight_NewestNewsArea_Code=rsRight_NewestNewsArea_Code.RecordCount
			numRight_NewestNewsArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsRight_NewestNewsArea_Code.EOF Then
		  %>
		  <!--<li><a href="">热门信息张维迎：痴心不改餐厅隆重开业</a></li>-->
		  <%
		  End If
		  %>
          
          <%
	 			'显示是否有最新HOT信息（近期，即七天之内发布的信息！）.WL
				Dim rsRight_isNewestNewsArea_Code,sqlRight_isNewestNewsArea_Code
		  
		  Do While Not rsRight_NewestNewsArea_Code.EOF
		  %>
        <li><a href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsRight_NewestNewsArea_Code("id") ) %>" target="_blank" title="<% =rsRight_NewestNewsArea_Code("topic") %>">
		
		<% =Left(rsRight_NewestNewsArea_Code("topic"),11) %><% If Len(rsRight_NewestNewsArea_Code("topic"))>11 Then Response.Write "..." %>


				<%
                '显示是否有最新HOT信息（近期，即七天之内发布的信息！）.WL
                'Dim rsRight_isNewestNewsArea_Code,sqlRight_isNewestNewsArea_Code
                sqlRight_isNewestNewsArea_Code="select top 1 * from [CXBG_details] where deleted=0 and isOnpublic=1 AND id="& rsRight_NewestNewsArea_Code("id") &" and datediff(day,adddate,GETDATE())<=7"
                Set rsRight_isNewestNewsArea_Code=Server.CreateObject("Adodb.RecordSet")
                rsRight_isNewestNewsArea_Code.Open sqlRight_isNewestNewsArea_Code,CONN,1,1
                
                '如果记录不为空，显示Hot提示.
                If Not rsRight_isNewestNewsArea_Code.EOF Then
                %>
                &nbsp;<img src="/images/ico/hot.gif" />
                <%
                End If
                
                '关闭记录集.
                rsRight_isNewestNewsArea_Code.Close
                'Set rsRight_isNewestNewsArea_Code=Nothing
                %>
                
        
        </a></li>
        <%
			  numRight_NewestNewsArea_Code=numRight_NewestNewsArea_Code+1
			  rsRight_NewestNewsArea_Code.MoveNext
		  Loop
		  
				'显示是否有最新HOT信息（近期，即七天之内发布的信息！）.WL
				Set rsRight_isNewestNewsArea_Code=Nothing
		  
		  '关闭记录集.
		  rsRight_NewestNewsArea_Code.Close
		  Set rsRight_NewestNewsArea_Code=Nothing
		  %>