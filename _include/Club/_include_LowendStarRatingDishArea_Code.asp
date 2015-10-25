		  <%
		  '好评菜.
		  Dim rsLowendStarRatingDishArea_Code,sqlLowendStarRatingDishArea_Code,countLowendStarRatingDishArea_Code,numLowendStarRatingDishArea_Code
		  	sqlLowendStarRatingDishArea_Code="SELECT TOP 3 * FROM (select TOP 3000 *,(select distinct cast(sumStarRating as decimal)/TotalStarRating as avgStarRating from (select product_id,sum(theStarRatingForChineseDishInformation) over() as sumStarRating,count(theStarRatingForChineseDishInformation)over() as TotalStarRating from [CXBG_account_RemarkOn] where product_id=[CXBG_product].id and deleted=0 and theStarRatingForChineseDishInformation>0) as x) as avgStarRatingNow from [CXBG_product] where deleted=0 and isOnsale=1 and isSetMeals=0 and 1=1 order by avgStarRatingNow desc,OrderID desc,id desc) AS xx WHERE avgStarRatingNow>0 AND avgStarRatingNow<=2.6"
			Set rsLowendStarRatingDishArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsLowendStarRatingDishArea_Code.Open sqlLowendStarRatingDishArea_Code,CONN,1,1
			countLowendStarRatingDishArea_Code=rsLowendStarRatingDishArea_Code.RecordCount
			numLowendStarRatingDishArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsLowendStarRatingDishArea_Code.EOF Then
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
		  Do While Not rsLowendStarRatingDishArea_Code.EOF
		  %>
          <li>
		    <p class="txtnoml lineheight28"><span class="fontred22">0<% =numLowendStarRatingDishArea_Code %></span> <a href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsLowendStarRatingDishArea_Code("id") ) %>" target="_blank" title="<% =rsLowendStarRatingDishArea_Code("ProductName") %>"><% =rsLowendStarRatingDishArea_Code("ProductName") %></a></p>
		    <a class="abg01 display" href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsLowendStarRatingDishArea_Code("id") ) %>" target="_blank" title="<% =rsLowendStarRatingDishArea_Code("ProductName") %>"><img src="<% If rsLowendStarRatingDishArea_Code("photo")<>"" Then Response.Write rsLowendStarRatingDishArea_Code("photo") Else Response.Write "/images/NoPic.png" %>" width="110" height="110" /></a>
			<!--<p class="xjcplxx">
            
            <ul class="rating" style="margin-top:8px; margin-left:0px;">
            <li class="current-rating" style="width:<% =Coke.ShowProductStarRating_Num(rsLowendStarRatingDishArea_Code("id")) * 20 + 1 %>px;"></li>
            
            </ul>
            
            </p>-->
		  </li>
          <%
			  numLowendStarRatingDishArea_Code=numLowendStarRatingDishArea_Code+1
			  rsLowendStarRatingDishArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsLowendStarRatingDishArea_Code.Close
		  Set rsLowendStarRatingDishArea_Code=Nothing
		  %>