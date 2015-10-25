		  <%
		  '首页新品推荐[套餐].
		  Dim rsRecommendSetMealsArea_Code,sqlRecommendSetMealsArea_Code,countRecommendSetMealsArea_Code,numRecommendSetMealsArea_Code
		  	sqlRecommendSetMealsArea_Code="select top 3 *,(select distinct cast(sumStarRating as decimal)/TotalStarRating as avgStarRating from (select product_id,sum(theStarRatingForChineseDishInformation) over() as sumStarRating,count(theStarRatingForChineseDishInformation)over() as TotalStarRating from [CXBG_account_RemarkOn] where product_id=[CXBG_product].id and deleted=0 and theStarRatingForChineseDishInformation>0) as x) as avgStarRatingNow from [CXBG_product] where deleted=0 and isOnsale=1 and isSetMeals=1 and 1=1 order by avgStarRatingNow desc,OrderID desc,id desc"
			Set rsRecommendSetMealsArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsRecommendSetMealsArea_Code.Open sqlRecommendSetMealsArea_Code,CONN,1,1
			countRecommendSetMealsArea_Code=rsRecommendSetMealsArea_Code.RecordCount
			numRecommendSetMealsArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsRecommendSetMealsArea_Code.EOF Then
		  %>
		  <!--<li>
		    <a class="abg04" href=""><img src="images/cpimg/tjtc_01.jpg" width="210" height="75" /></a>
			<p class="txtnoml linheight22"><span class="fontred22">01</span> <a href="">欢迎光临</a></p>
			<p class="xjcplxx"><img src="images/xx.gif" width="15" height="16" /><img src="images/xx.gif" width="15" height="16" /><img src="images/xx.gif" width="15" height="16" /><img src="images/xx.gif" width="15" height="16" /><img src="images/xx.gif" width="15" height="16" /></p>
		  </li>-->
		  <%
		  End If
		  %>
          
          <%
		  Dim i_starrating
		  
		  Do While Not rsRecommendSetMealsArea_Code.EOF
		  %>
          <li>
		    <a class="abg04" href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsRecommendSetMealsArea_Code("id") ) %>" target="_blank" title="<% =rsRecommendSetMealsArea_Code("ProductName") %>"><img src="<% If rsRecommendSetMealsArea_Code("SetMeals_Photo")<>"" Then Response.Write rsRecommendSetMealsArea_Code("SetMeals_Photo") Else Response.Write "/images/NoPic.png" %>" width="210" height="75" /></a>
			<p class="txtnoml linheight22">
            	<span class="fontred22">0<% =numRecommendSetMealsArea_Code %></span> <a href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsRecommendSetMealsArea_Code("id") ) %>" target="_blank" title="<% =rsRecommendSetMealsArea_Code("ProductName") %>"><% =rsRecommendSetMealsArea_Code("ProductName") %></a>
            </p>
			<p class="xjcplxx">
                <!--<ul class="rating" style="margin-top:8px; margin-left:0px;">
                <li class="current-rating" style="width:<% =Coke.ShowProductStarRating_Num(rsRecommendSetMealsArea_Code("id")) * 20 + 1 %>px;"></li>
                
                </ul>-->
                
                <%
				'输出星级.
				'Dim i_starrating
				For i_starrating=1 To CokeShow.QuZheng( Coke.ShowProductStarRating_Num(rsRecommendSetMealsArea_Code("id")) )
				%>
					<img src="/images/xx.gif" width="16" height="16" />
				<%
				Next
				%>
				<%
				'输出灰补星级.
				For i_starrating=(5-1) To CokeShow.QuZheng( Coke.ShowProductStarRating_Num(rsRecommendSetMealsArea_Code("id")) ) Step -1
				%>
					<img src="/images/xx_goy.gif" width="16" height="16" />
				<%
				Next
				%>
            </p>
		  </li>
          
          
          	
          <%
			  numRecommendSetMealsArea_Code=numRecommendSetMealsArea_Code+1
			  rsRecommendSetMealsArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsRecommendSetMealsArea_Code.Close
		  Set rsRecommendSetMealsArea_Code=Nothing
		  %>