		  <%
		  '好评菜.
		  Dim rsHighPointStarRatingDishArea_Code,sqlHighPointStarRatingDishArea_Code,countHighPointStarRatingDishArea_Code,numHighPointStarRatingDishArea_Code
		  	sqlHighPointStarRatingDishArea_Code="SELECT TOP 3 * FROM (select TOP 3000 *,(select distinct cast(sumStarRating as decimal)/TotalStarRating as avgStarRating from (select product_id,sum(theStarRatingForChineseDishInformation) over() as sumStarRating,count(theStarRatingForChineseDishInformation)over() as TotalStarRating from [CXBG_account_RemarkOn] where product_id=[CXBG_product].id and deleted=0 and theStarRatingForChineseDishInformation>0) as x) as avgStarRatingNow from [CXBG_product] where deleted=0 and isOnsale=1 and isSetMeals=0 and 1=1 order by avgStarRatingNow desc,OrderID desc,id desc) AS xx WHERE avgStarRatingNow>3.8 AND avgStarRatingNow<=5"
			Set rsHighPointStarRatingDishArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsHighPointStarRatingDishArea_Code.Open sqlHighPointStarRatingDishArea_Code,CONN,1,1
			countHighPointStarRatingDishArea_Code=rsHighPointStarRatingDishArea_Code.RecordCount
			numHighPointStarRatingDishArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsHighPointStarRatingDishArea_Code.EOF Then
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
		  Do While Not rsHighPointStarRatingDishArea_Code.EOF
		  %>
          <li>
		    <p class="txtnoml lineheight28"><span class="fontred22">0<% =numHighPointStarRatingDishArea_Code %></span> <a href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsHighPointStarRatingDishArea_Code("id") ) %>" target="_blank" title="<% =rsHighPointStarRatingDishArea_Code("ProductName") %>"><% =rsHighPointStarRatingDishArea_Code("ProductName") %></a></p>
		    <a class="abg01 display" href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsHighPointStarRatingDishArea_Code("id") ) %>" target="_blank" title="<% =rsHighPointStarRatingDishArea_Code("ProductName") %>"><img src="<% If rsHighPointStarRatingDishArea_Code("photo")<>"" Then Response.Write rsHighPointStarRatingDishArea_Code("photo") Else Response.Write "/images/NoPic.png" %>" width="110" height="110" /></a>
			<!--<p class="xjcplxx">
            
            <ul class="rating" style="margin-top:8px; margin-left:0px;">
            <li class="current-rating" style="width:<% =Coke.ShowProductStarRating_Num(rsHighPointStarRatingDishArea_Code("id")) * 20 + 1 %>px;"></li>
            
            </ul>
            
            </p>-->
		  </li>
          <%
			  numHighPointStarRatingDishArea_Code=numHighPointStarRatingDishArea_Code+1
			  rsHighPointStarRatingDishArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsHighPointStarRatingDishArea_Code.Close
		  Set rsHighPointStarRatingDishArea_Code=Nothing
		  %>