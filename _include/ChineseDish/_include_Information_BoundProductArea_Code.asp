		  <%
		  '推荐最佳搭配.
		  Dim rsInformation_BoundProductArea_Code,sqlInformation_BoundProductArea_Code,countInformation_BoundProductArea_Code,numInformation_BoundProductArea_Code
			sqlInformation_BoundProductArea_Code="select top 5 * from [CXBG_product__BoundProduct] where product_id="& RS("id") &" order by Product_Orderid desc,id desc"
			'sqlInformation_BoundProductArea_Code="select top 5 * from [CXBG_product__BoundProduct] where isOnsale=1 and is_display_newproduct=1 order by newid()"
			Set rsInformation_BoundProductArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsInformation_BoundProductArea_Code.Open sqlInformation_BoundProductArea_Code,CONN,1,1
			countInformation_BoundProductArea_Code=rsInformation_BoundProductArea_Code.RecordCount
			numInformation_BoundProductArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsInformation_BoundProductArea_Code.EOF Then
		  %>
		  欢迎光临，餐厅尚未上传相关的推荐最佳搭配菜：）
		  <%
		  End If
		  %>
          
          <%
		  Do While Not rsInformation_BoundProductArea_Code.EOF
		  %>
         <a style="height:110px" href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsInformation_BoundProductArea_Code("CurrentProductID") ) %>" target="_blank" title="<% =CokeShow.otherField("[CXBG_product]",rsInformation_BoundProductArea_Code("CurrentProductID"),"id","ProductName",True,0) %>"><img src="<% If CokeShow.otherField("[CXBG_product]",rsInformation_BoundProductArea_Code("CurrentProductID"),"id","Photo",True,0)<>"" Then Response.Write CokeShow.otherField("[CXBG_product]",rsInformation_BoundProductArea_Code("CurrentProductID"),"id","Photo",True,0) Else Response.Write "/images/NoPic.png" %>" width="80" height="80" /> <% =CokeShow.otherField("[CXBG_product]",rsInformation_BoundProductArea_Code("CurrentProductID"),"id","ProductName",True,0) %></a>
          <%
			  numInformation_BoundProductArea_Code=numInformation_BoundProductArea_Code+1
			  rsInformation_BoundProductArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsInformation_BoundProductArea_Code.Close
		  Set rsInformation_BoundProductArea_Code=Nothing
		  %>