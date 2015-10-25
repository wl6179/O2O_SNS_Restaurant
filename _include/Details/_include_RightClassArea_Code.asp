		<%
		  '信息分类.
		  '有限特殊处理参数：处理一下传参中的原classid，因为将会替换、并使用本函数中定义的新classid.
		  Dim strFileNameRightClassArea_Code
		  strFileNameRightClassArea_Code=strFileName
		  If Instr( Lcase(Trim(strFileNameRightClassArea_Code)) ,"classid=")>0 Then strFileNameRightClassArea_Code=Replace(strFileNameRightClassArea_Code,"classid=","classid$$$=")
		  
		  Dim rsRightClassArea_Code,sqlRightClassArea_Code,countRightClassArea_Code,numRightClassArea_Code
		  	sqlRightClassArea_Code="select * from [CXBG_details_class] where isShow=1 and Depth=0 ORDER BY RootID,OrderID"
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
          
          
          <div class="flxx_rightbt"><span class="fontredbt14"><% =rsRightClassArea_Code("classname") %></span><span class="font16000">分类</span></div>
          
          <ul class="flxx_ullist">
            <!--选中状态default="default"-->
            <li id="mod0" tabcontentid="div0" activeclass="flxx_ullist_vis" deactiveclass="flxx_ullist_link" groupname="m1" <% If CokeShow.CokeClng(Request("classid"))=0 Then Response.Write "default=""default""" %> class="flxx_ullist_link" hoverclass="flxx_ullist_hov"><a href="/Details/Details.Welcome">全部信息</a><span class="font10fff"> (20xx)</span></li>
                
                
                <!--所有分类 Begin-->
                  <%
				  Do While Not rsRightClassArea_Code.EOF
				  %>
                    <li id="mod<% =rsRightClassArea_Code("classid") %>" tabcontentid="div<% =rsRightClassArea_Code("classid") %>" activeclass="flxx_ullist_vis" deactiveclass="flxx_ullist_link" groupname="m1" <% If CokeShow.CokeClng(Request("classid"))=rsRightClassArea_Code("classid") Then Response.Write "default=""default""" %> class="flxx_ullist_link" hoverclass="flxx_ullist_hov"><a href="<% =strFileNameRightClassArea_Code %>&classid=<% =rsRightClassArea_Code("classid") %><% If rsRightClassArea_Code("isHot")=1 Then %>&isHot=1<% End If %>"><% =rsRightClassArea_Code("classname") %><% If rsRightClassArea_Code("isHot")=1 Then %>&nbsp;&nbsp;&nbsp;<img src="/images/ico/hot.gif" alt="热度" /><% End If %></a><span class="font10fff"> <!--(20x)--></span></li>
                  <%
					  numRightClassArea_Code=numRightClassArea_Code+1
					  rsRightClassArea_Code.MoveNext
				  Loop
				  
				  '关闭记录集.
				  rsRightClassArea_Code.Close
				  Set rsRightClassArea_Code=Nothing
				  %>
                <!--所有分类 End-->
                
                
            
            <div class="clear"></div>
          </ul>
          
          <div class="flxx_rightbt_bot"></div>