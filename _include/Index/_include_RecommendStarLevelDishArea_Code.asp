		  <%
		  '星级菜.
		  Dim rsRecommendStarLevelDishArea_Code,sqlRecommendStarLevelDishArea_Code,countRecommendStarLevelDishArea_Code,numRecommendStarLevelDishArea_Code
		  	sqlRecommendStarLevelDishArea_Code="select top 3 *,(select distinct cast(sumStarRating as decimal)/TotalStarRating as avgStarRating from (select product_id,sum(theStarRatingForChineseDishInformation) over() as sumStarRating,count(theStarRatingForChineseDishInformation)over() as TotalStarRating from [CXBG_account_RemarkOn] where product_id=[CXBG_product].id and deleted=0 and theStarRatingForChineseDishInformation>0) as x) as avgStarRatingNow from [CXBG_product] where deleted=0 and isOnsale=1 and isSetMeals=0 and 1=1 order by avgStarRatingNow desc,OrderID desc,id desc"
			Set rsRecommendStarLevelDishArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsRecommendStarLevelDishArea_Code.Open sqlRecommendStarLevelDishArea_Code,CONN,1,1
			countRecommendStarLevelDishArea_Code=rsRecommendStarLevelDishArea_Code.RecordCount
			numRecommendStarLevelDishArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsRecommendStarLevelDishArea_Code.EOF Then
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
		  Do While Not rsRecommendStarLevelDishArea_Code.EOF
		  %>
          <li>
		    <p class="txtnoml lineheight28"><span class="fontred22">0<% =numRecommendStarLevelDishArea_Code %></span> <a href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsRecommendStarLevelDishArea_Code("id") ) %>" target="_blank" title="<% =rsRecommendStarLevelDishArea_Code("ProductName") %>"><% =rsRecommendStarLevelDishArea_Code("ProductName") %></a></p>
		    <a class="abg01 display" href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsRecommendStarLevelDishArea_Code("id") ) %>" target="_blank" title="<% =rsRecommendStarLevelDishArea_Code("ProductName") %>"><img src="<% If rsRecommendStarLevelDishArea_Code("photo")<>"" Then Response.Write rsRecommendStarLevelDishArea_Code("photo") Else Response.Write "/images/NoPic.png" %>" width="110" height="110" /></a>
			<p class="xjcplxx">
            
            <ul class="rating" style="margin-top:8px; margin-left:0px;">
            <li class="current-rating" style="width:<% =Coke.ShowProductStarRating_Num(rsRecommendStarLevelDishArea_Code("id")) * 20 + 1 %>px;"></li>
            
            </ul>
            
            </p>
		  </li>
          <%
			  numRecommendStarLevelDishArea_Code=numRecommendStarLevelDishArea_Code+1
			  rsRecommendStarLevelDishArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsRecommendStarLevelDishArea_Code.Close
		  Set rsRecommendStarLevelDishArea_Code=Nothing
		  %>