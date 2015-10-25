		  <%
		  '好评菜.
		  Dim rsMiddleendStarRatingDishArea_Code,sqlMiddleendStarRatingDishArea_Code,countMiddleendStarRatingDishArea_Code,numMiddleendStarRatingDishArea_Code
		  	sqlMiddleendStarRatingDishArea_Code="SELECT TOP 3 * FROM (select TOP 3000 *,(select distinct cast(sumStarRating as decimal)/TotalStarRating as avgStarRating from (select product_id,sum(theStarRatingForChineseDishInformation) over() as sumStarRating,count(theStarRatingForChineseDishInformation)over() as TotalStarRating from [CXBG_account_RemarkOn] where product_id=[CXBG_product].id and deleted=0 and theStarRatingForChineseDishInformation>0) as x) as avgStarRatingNow from [CXBG_product] where deleted=0 and isOnsale=1 and isSetMeals=0 and 1=1 order by avgStarRatingNow desc,OrderID desc,id desc) AS xx WHERE avgStarRatingNow>2.6 AND avgStarRatingNow<=3.8"
			Set rsMiddleendStarRatingDishArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsMiddleendStarRatingDishArea_Code.Open sqlMiddleendStarRatingDishArea_Code,CONN,1,1
			countMiddleendStarRatingDishArea_Code=rsMiddleendStarRatingDishArea_Code.RecordCount
			numMiddleendStarRatingDishArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsMiddleendStarRatingDishArea_Code.EOF Then
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
		  Do While Not rsMiddleendStarRatingDishArea_Code.EOF
		  %>
          <li>
		    <p class="txtnoml lineheight28"><span class="fontred22">0<% =numMiddleendStarRatingDishArea_Code %></span> <a href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsMiddleendStarRatingDishArea_Code("id") ) %>" target="_blank" title="<% =rsMiddleendStarRatingDishArea_Code("ProductName") %>"><% =rsMiddleendStarRatingDishArea_Code("ProductName") %></a></p>
		    <a class="abg01 display" href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsMiddleendStarRatingDishArea_Code("id") ) %>" target="_blank" title="<% =rsMiddleendStarRatingDishArea_Code("ProductName") %>"><img src="<% If rsMiddleendStarRatingDishArea_Code("photo")<>"" Then Response.Write rsMiddleendStarRatingDishArea_Code("photo") Else Response.Write "/images/NoPic.png" %>" width="110" height="110" /></a>
			<!--<p class="xjcplxx">
            
            <ul class="rating" style="margin-top:8px; margin-left:0px;">
            <li class="current-rating" style="width:<% =Coke.ShowProductStarRating_Num(rsMiddleendStarRatingDishArea_Code("id")) * 20 + 1 %>px;"></li>
            
            </ul>
            
            </p>-->
		  </li>
          <%
			  numMiddleendStarRatingDishArea_Code=numMiddleendStarRatingDishArea_Code+1
			  rsMiddleendStarRatingDishArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsMiddleendStarRatingDishArea_Code.Close
		  Set rsMiddleendStarRatingDishArea_Code=Nothing
		  %>