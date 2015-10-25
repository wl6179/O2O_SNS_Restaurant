	    
        <li><img src="<% If CokeShow.Setup(32,0)<>"" Then Response.Write CokeShow.Setup(32,0) Else Response.Write "/images/NoPic.png" %>" width="206" height="103" /></li>
	    <%
		  '最新促销.
		  Dim rsRight_NewestSalesPromotionArea_Code,sqlRight_NewestSalesPromotionArea_Code,countRight_NewestSalesPromotionArea_Code,numRight_NewestSalesPromotionArea_Code
		  	sqlRight_NewestSalesPromotionArea_Code="select top 3 * from [CXBG_details] where deleted=0 and isOnpublic=1 and details_class_id=12 order by isRecommend desc,details_orderid desc,id desc"
			Set rsRight_NewestSalesPromotionArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsRight_NewestSalesPromotionArea_Code.Open sqlRight_NewestSalesPromotionArea_Code,CONN,1,1
			countRight_NewestSalesPromotionArea_Code=rsRight_NewestSalesPromotionArea_Code.RecordCount
			numRight_NewestSalesPromotionArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsRight_NewestSalesPromotionArea_Code.EOF Then
		  %>
		  <!--<li><a href="">热门信息张维迎：痴心不改餐厅隆重开业</a></li>-->
		  <%
		  End If
		  %>
          
          <%
		  Do While Not rsRight_NewestSalesPromotionArea_Code.EOF
		  %>
        <li><a href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsRight_NewestSalesPromotionArea_Code("id") ) %>" target="_blank" title="<% =rsRight_NewestSalesPromotionArea_Code("topic") %>"><% =rsRight_NewestSalesPromotionArea_Code("topic") %></a></li>
        <%
			  numRight_NewestSalesPromotionArea_Code=numRight_NewestSalesPromotionArea_Code+1
			  rsRight_NewestSalesPromotionArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsRight_NewestSalesPromotionArea_Code.Close
		  Set rsRight_NewestSalesPromotionArea_Code=Nothing
		  %>