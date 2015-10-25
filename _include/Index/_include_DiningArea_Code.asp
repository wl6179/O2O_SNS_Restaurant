		  <%
		  '餐厅环境.
		  Dim rsDiningArea_Code,sqlDiningArea_Code,CountDiningArea_Code,Num_DiningArea_Code
		  	sqlDiningArea_Code="select top 6 * from [CXBG_DiningArea_class] where isShow=1 order by RootID,OrderID"
			Set rsDiningArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsDiningArea_Code.Open sqlDiningArea_Code,CONN,1,1
			CountDiningArea_Code=rsDiningArea_Code.RecordCount
		  %>
          <%
		  '如果记录为空.
		  If rsDiningArea_Code.EOF Then
		  %>
		  <div class="incthj_conten" id="in22c_txt1">
		    <a class="incthj_conimg_left" href="#"><img src="/images/NoPic.png" width="230" height="115" /></a>
			<div class="incthj_contxt">
            <span class="fontred14">欢迎光临</span>
            <a href="#">Welcome</a>
            </div>
			<a class="incthj_conimg_right" href="#"><img src="/images/NoPic.png" width="70" height="70" /></a>	
			<a class="incthj_conimg_right" href="#"><img src="/images/NoPic.png" width="70" height="70" /></a>		  
		  </div>
		  <%
		  End If
		  %>
          
          <%
		  Dim rsDiningArea_Code_2,sqlDiningArea_Code_2,CountDiningArea_Code_2
		  '初始化计数器.
		  Num_DiningArea_Code=1
		  '循环.
		  Do While Not rsDiningArea_Code.EOF
		  %>
          <div class="incthj_conten<% If Num_DiningArea_Code>1 Then Response.Write "_none" %>" id="in22c_txt<% =Num_DiningArea_Code %>">
		    <%
			  '餐厅环境图片记录.
			  'Dim rsDiningArea_Code_2,sqlDiningArea_Code_2,CountDiningArea_Code_2
				sqlDiningArea_Code_2="select top 3 * from [CXBG_DiningArea] where deleted=0 and isOnpublic=1 and details_class_id="& rsDiningArea_Code("classid") &" order by details_orderid desc,id asc"
				Set rsDiningArea_Code_2=Server.CreateObject("Adodb.RecordSet")
				rsDiningArea_Code_2.Open sqlDiningArea_Code_2,CONN,1,1
				CountDiningArea_Code_2=rsDiningArea_Code_2.RecordCount
				
				'初始化计数器.
				  'Num_DiningArea_Code=1
				  
			  %>
            <!--第一张大图-->
            <% If Not rsDiningArea_Code_2.Eof Then %><a class="incthj_conimg_left" href="<% If rsDiningArea_Code_2("Photo")<>"" Then Response.Write rsDiningArea_Code_2("Photo") Else Response.Write "/images/NoPic.png" %>" dojoType="dojox.image.Lightbox" group="group<% =rsDiningArea_Code("classid") %>" title="<% =rsDiningArea_Code("classname") %>--<% =rsDiningArea_Code_2("topic") %>"><% Else %><a class="incthj_conimg_left" href="###"><% End If %><img src="<%
			If Not rsDiningArea_Code_2.Eof Then
			Response.Write rsDiningArea_Code_2("photo"):rsDiningArea_Code_2.MoveNext
			Else
			Response.Write "/images/NoPic.png"
			End If
			%>" width="230" height="115" /></a>
			
            <div class="incthj_contxt">
                <span class="fontred14"><% =rsDiningArea_Code("classname") %></span>
                <!--<% If Not rsDiningArea_Code_2.Eof Then %><a href="<% If rsDiningArea_Code_2("Photo")<>"" Then Response.Write rsDiningArea_Code_2("Photo") Else Response.Write "/images/NoPic.png" %>" dojoType="dojox.image.Lightbox" group="group<% =rsDiningArea_Code("classid") %>" title="<% =rsDiningArea_Code("classname") %>--<% =rsDiningArea_Code_2("topic") %>"><% Else %><a href="###"><% End If %>--><% =rsDiningArea_Code("readme") %><!--</a>-->
            </div>
            
            <!--第二张小图-->
			<% If Not rsDiningArea_Code_2.Eof Then %><a class="incthj_conimg_right" href="<% If rsDiningArea_Code_2("Photo")<>"" Then Response.Write rsDiningArea_Code_2("Photo") Else Response.Write "/images/NoPic.png" %>" dojoType="dojox.image.Lightbox" group="group<% =rsDiningArea_Code("classid") %>" title="<% =rsDiningArea_Code("classname") %>--<% =rsDiningArea_Code_2("topic") %>"><% Else %><a class="incthj_conimg_right" href="###"><% End If %><img src="<%
			If Not rsDiningArea_Code_2.Eof Then
			Response.Write rsDiningArea_Code_2("photo"):rsDiningArea_Code_2.MoveNext
			Else
			Response.Write "/images/NoPic.png"
			End If
			%>" width="70" height="70" /></a>
            
			<% If Not rsDiningArea_Code_2.Eof Then %><a class="incthj_conimg_right" href="<% If rsDiningArea_Code_2("Photo")<>"" Then Response.Write rsDiningArea_Code_2("Photo") Else Response.Write "/images/NoPic.png" %>" dojoType="dojox.image.Lightbox" group="group<% =rsDiningArea_Code("classid") %>" title="<% =rsDiningArea_Code("classname") %>--<% =rsDiningArea_Code_2("topic") %>"><% Else %><a class="incthj_conimg_right" href="###"><% End If %><img src="<%
			If Not rsDiningArea_Code_2.Eof Then
			Response.Write rsDiningArea_Code_2("photo"):rsDiningArea_Code_2.MoveNext
			Else
			Response.Write "/images/NoPic.png"
			End If
			%>" width="70" height="70" /></a>
            
		  </div>
          <%
			  Num_DiningArea_Code=Num_DiningArea_Code+1
			  rsDiningArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  'rsDiningArea_Code.Close
		  'Set rsDiningArea_Code=Nothing
		  %>
		  
          
		  <ul class="incthj_button">
		    <%
			  '先判断是否有环境分类记录.
			  If CountDiningArea_Code>0 Then
			  
			  '跳到第一个记录.
			  rsDiningArea_Code.MoveFirst
			  '初始化计数器.
			  Num_DiningArea_Code=1
			  '循环.
			  Do While Not rsDiningArea_Code.EOF
			  %>
            <li class="tjimg_menu<% If Num_DiningArea_Code=1 Then Response.Write "On" Else Response.Write "No" %>" id="tjimg<% =Num_DiningArea_Code %>" onMouseOver="switnewspbq('tjimg','in22c_txt','<% =Num_DiningArea_Code %>','tjimg_menuOn');this.blur();"><% =rsDiningArea_Code("classname") %></li>
            <%
				  Num_DiningArea_Code=Num_DiningArea_Code+1
				  rsDiningArea_Code.MoveNext
			  Loop
			  
			  End If
			  
			  '关闭记录集.
			  rsDiningArea_Code.Close
			  Set rsDiningArea_Code=Nothing
			  %>
		  </ul>