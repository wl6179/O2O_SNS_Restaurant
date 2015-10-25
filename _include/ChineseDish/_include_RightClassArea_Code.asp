		<%
		  '分类.
		  
		  'WL
		  Dim ExecuteSearchTemp01
		  ExecuteSearchTemp01=0
		  If Instr( Lcase(Trim(strFileName)) ,"executesearch=1")>0 Then strFileName=Replace(strFileName,"ExecuteSearch=1","ExecuteSearch=0"):ExecuteSearchTemp01=1		'过滤掉首页新品推荐筛选的条件携带！(待会儿在下边又改回来ExecuteSearch的值.)
		  
		  
		  '有限特殊处理参数：处理一下传参中的原classid，因为将会替换、并使用本函数中定义的新classid.
		  Dim strFileNameRightClassArea_Code
		  strFileNameRightClassArea_Code=strFileName
		  If Instr( Lcase(Trim(strFileNameRightClassArea_Code)) ,"classid=")>0 Then strFileNameRightClassArea_Code=Replace(strFileNameRightClassArea_Code,"classid=","classid$$$=")
		  
		  Dim rsRightClassArea_Code,sqlRightClassArea_Code,countRightClassArea_Code,numRightClassArea_Code
		  Dim rsRightClassArea_Code2,sqlRightClassArea_Code2,countRightClassArea_Code2
		  	sqlRightClassArea_Code="select * from [CXBG_product_class] where isNavigation=0 AND isShow=1 and Depth=0 ORDER BY RootID,OrderID"
			Set rsRightClassArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsRightClassArea_Code.Open sqlRightClassArea_Code,CONN,1,1
			countRightClassArea_Code=rsRightClassArea_Code.RecordCount
			numRightClassArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsRightClassArea_Code.EOF Then
		  %>
		  <!--www.CokeShow.com.cn-->
		  <%
		  End If
		  %>
          
          <%
		  Do While Not rsRightClassArea_Code.EOF
		  %>
          
          <div class="flxx_rightbt"><span class="fontredbt14"><% =rsRightClassArea_Code("classname") %></span><span class="font16000">分类</span></div>
          
          <ul class="flxx_ullist">
            <!--选中状态default="default"-->
            <li id="mod<% =rsRightClassArea_Code("classid") %>" tabcontentid="div<% =rsRightClassArea_Code("classid") %>" activeclass="flxx_ullist_vis" deactiveclass="flxx_ullist_link" groupname="m1" <% If CokeShow.CokeClng(Request("classid"))=rsRightClassArea_Code("classid") Then Response.Write "default=""default""" %> class="flxx_ullist_link" hoverclass="flxx_ullist_hov"><a href="<% =strFileNameRightClassArea_Code %>&classid=<% =rsRightClassArea_Code("classid") %>">全部</a><span class="font10fff"> (20xx)</span></li>
            <%
            '开始计数.
            numRightClassArea_Code=numRightClassArea_Code+1
            %>
                
                
                <!--二级分类 Begin-->
                <%
                  '二级分类.
                  'Dim rsRightClassArea_Code2,sqlRightClassArea_Code2,countRightClassArea_Code2
                    sqlRightClassArea_Code2="select * from [CXBG_product_class] where isNavigation=0 AND isShow=1 and Depth=1 and parentid="& rsRightClassArea_Code("id") &" ORDER BY RootID,OrderID"
                    Set rsRightClassArea_Code2=Server.CreateObject("Adodb.RecordSet")
                    rsRightClassArea_Code2.Open sqlRightClassArea_Code2,CONN,1,1
                    countRightClassArea_Code2=rsRightClassArea_Code2.RecordCount
                    'numRightClassArea_Code2=1
                  %>
                  <%
				  Do While Not rsRightClassArea_Code2.EOF
				  %>
                    <li id="mod<% =rsRightClassArea_Code2("classid") %>" tabcontentid="div<% =rsRightClassArea_Code2("classid") %>" activeclass="flxx_ullist_vis" deactiveclass="flxx_ullist_link" groupname="m1" <% If CokeShow.CokeClng(Request("classid"))=rsRightClassArea_Code2("classid") Then Response.Write "default=""default""" %> class="flxx_ullist_link" hoverclass="flxx_ullist_hov"><a href="<% =strFileNameRightClassArea_Code %>&classid=<% =rsRightClassArea_Code2("classid") %>"><% =rsRightClassArea_Code2("classname") %></a><span class="font10fff"> <!--(20x)--></span></li>
                  <%
					  numRightClassArea_Code=numRightClassArea_Code+1
					  rsRightClassArea_Code2.MoveNext
				  Loop
				  
				  '关闭记录集.
				  rsRightClassArea_Code2.Close
				  'Set rsRightClassArea_Code2=Nothing
				  %>
                <!--二级分类 End-->
                
                
            
            <div class="clear"></div>
          </ul>
          
          <div class="flxx_rightbt_bot"></div>
      
          <%
			  numRightClassArea_Code=numRightClassArea_Code+1
			  rsRightClassArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsRightClassArea_Code.Close
		  Set rsRightClassArea_Code=Nothing
		  Set rsRightClassArea_Code2=Nothing
		  
		  
		  'WL
		  If ExecuteSearchTemp01=1 Then strFileName=Replace(strFileName,"ExecuteSearch=0","ExecuteSearch=1")		'改回来！过滤掉首页新品推荐筛选的条件携带！(待会儿在下边又改回来ExecuteSearch的值.)
		  %>