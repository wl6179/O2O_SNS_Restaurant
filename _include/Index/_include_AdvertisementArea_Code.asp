  <div id="cxbg_showcs">
    <div class="cxposition">
    <div class="scbutton_left"><a onMouseDown="ISL_GoUp_1()" onMouseUp="ISL_StopUp_1()" onMouseOut="ISL_StopUp_1()" href="javascript:void(0);" target="_self"></a></div>
	<div class="scmid" id="ISL_Cont_1">
	  <div class="scrcont">
        <div id="List1_1">
 	      <%
		  '广告轮换.
		  Dim rsAdvertisementArea_Code,sqlAdvertisementArea_Code,countAdvertisementArea_Code,numAdvertisementArea_Code
		  	sqlAdvertisementArea_Code="select top 15 * from [CXBG_advertisement_main] where isOnpublic=1 order by RootID,OrderID"
			Set rsAdvertisementArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsAdvertisementArea_Code.Open sqlAdvertisementArea_Code,CONN,1,1
			countAdvertisementArea_Code=rsAdvertisementArea_Code.RecordCount
			numAdvertisementArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsAdvertisementArea_Code.EOF Then
		  %>
		  <div class="scmid_txtimg">
	        <div class="scmid_txt">
			  <p class="font22">欢迎光临，餐厅尚未上传推荐</p>
			  <div class="fontred12 height110 linheight22">
			   <p>Welcome</p>
			   <p>：）</p>			  
			  </div>
			  <p><a class="font14b" href="#">更多详情</a></p>
			</div>
	        <div class="scmid_img">
			  <a href="#"><img src="/images/NoPic.png" width="400" height="200" /></a>			</div>
	      </div>
		  <%
		  End If
		  %>
          
          <%
		  Do While Not rsAdvertisementArea_Code.EOF
		  %>
          <div class="scmid_txtimg">
	        <div class="scmid_txt">
			  <p class="font22"><% =rsAdvertisementArea_Code("classname") %></p>
			  <div class="fontred12 height110 linheight22">
			   <p><% =CokeShow.filtResponseWriteHTML(rsAdvertisementArea_Code("readme")) %></p>
			   			  
			  </div>
			  <p><a class="font14b" href="<% If rsAdvertisementArea_Code("isDirectingLink")=1 Then Response.Write rsAdvertisementArea_Code("DirectingLink") Else Response.Write "/ChineseDish/ChineseDish.Welcome?ExecuteSearch=10&Keyword="& rsAdvertisementArea_Code("keywords") %>" target="_blank">更多详情</a></p>
			</div>
	        <div class="scmid_img">
			  <a href="<% If rsAdvertisementArea_Code("isDirectingLink")=1 Then Response.Write rsAdvertisementArea_Code("DirectingLink") Else Response.Write "/ChineseDish/ChineseDish.Welcome?ExecuteSearch=10&Keyword="& rsAdvertisementArea_Code("keywords") %>" target="_blank"><img src="<% If rsAdvertisementArea_Code("photo")<>"" Then Response.Write rsAdvertisementArea_Code("photo") Else Response.Write "/images/NoPic.png" %>" width="400" height="200" /></a>			</div>
	      </div>
          <%
			  numAdvertisementArea_Code=numAdvertisementArea_Code+1
			  rsAdvertisementArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsAdvertisementArea_Code.Close
		  Set rsAdvertisementArea_Code=Nothing
		  %>
 	      
          
		</div>
	    <div id="List2_1"></div> 
	  </div>
	</div>        
    <div class="scbutton_right"><a onMouseDown="ISL_GoDown_1()" onMouseUp="ISL_StopDown_1()" onMouseOut="ISL_StopDown_1()" href="javascript:void(0);" target="_self"></a></div>
	</div>
	</div>