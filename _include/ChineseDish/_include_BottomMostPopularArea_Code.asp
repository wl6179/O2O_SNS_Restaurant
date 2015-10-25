		  <%
		  '最受欢迎菜.
		  Dim rsBottomMostPopularArea_Code,sqlBottomMostPopularArea_Code,countBottomMostPopularArea_Code,numBottomMostPopularArea_Code
			sqlBottomMostPopularArea_Code="select top 7 *,(select distinct cast(sumChineseDish_Taste as decimal)/TotalChineseDish_Taste as avgChineseDish_Taste from (select product_id,sum(ChineseDish_Taste) over() as sumChineseDish_Taste,count(ChineseDish_Taste)over() as TotalChineseDish_Taste from [CXBG_account_RemarkOn] where product_id=[CXBG_product].id and deleted=0 and ChineseDish_Taste>0) as x) as avgChineseDish_TasteNow from [CXBG_product] where deleted=0 and isOnsale=1 order by avgChineseDish_TasteNow desc,OrderID desc,id desc"
			'sqlBottomMostPopularArea_Code="select top 7 * from [CXBG_product] where deleted=0 and isOnsale=1 and is_display_newproduct=1 order by newid()"
			Set rsBottomMostPopularArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsBottomMostPopularArea_Code.Open sqlBottomMostPopularArea_Code,CONN,1,1
			countBottomMostPopularArea_Code=rsBottomMostPopularArea_Code.RecordCount
			numBottomMostPopularArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsBottomMostPopularArea_Code.EOF Then
		  %>
		  欢迎光临，餐厅最受欢迎菜尚未产生：）
		  <%
		  End If
		  %>
          
          <%
		  Do While Not rsBottomMostPopularArea_Code.EOF
		  %>
      <li>
		<p class="txtnoml lineheight28"><span class="fontred22">0<% =numBottomMostPopularArea_Code %></span> <a href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsBottomMostPopularArea_Code("id") ) %>" target="_blank" title="<% =rsBottomMostPopularArea_Code("ProductName") %>"><% =rsBottomMostPopularArea_Code("ProductName") %></a></p>
		<a class="abg01 display" href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsBottomMostPopularArea_Code("id") ) %>" target="_blank" title="<% =rsBottomMostPopularArea_Code("ProductName") %>"><img src="<% If rsBottomMostPopularArea_Code("Photo")<>"" Then Response.Write rsBottomMostPopularArea_Code("Photo") Else Response.Write "/images/NoPic.png" %>" width="110" height="110" /></a>
		<p class="xjcplxx">
            <ul class="rating" style="margin-top:8px; margin-left:0px;">
            <li class="current-rating" style="width:<% =Coke.ShowProductStarRating_Num(rsBottomMostPopularArea_Code("id")) * 20 + 1 %>px;"></li>
            
            </ul>
        </p>
	  </li>
          <%
			  numBottomMostPopularArea_Code=numBottomMostPopularArea_Code+1
			  rsBottomMostPopularArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsBottomMostPopularArea_Code.Close
		  Set rsBottomMostPopularArea_Code=Nothing
		  %>