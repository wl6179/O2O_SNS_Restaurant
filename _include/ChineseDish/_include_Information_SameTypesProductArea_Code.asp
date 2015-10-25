		  <%
		  '同菜系推荐.
		  Dim rsInformation_SameTypesProductArea_Code,sqlInformation_SameTypesProductArea_Code,countInformation_SameTypesProductArea_Code,numInformation_SameTypesProductArea_Code
		  Dim tmpString03471,i777
		  tmpString03471=""
		  '构建扩展分类匹配分类SQL.
		  If RS("product_class_id_extend")<>"" Then
		  	For i777=0 To Ubound(Split(RS("product_class_id_extend"),","))
				If isNumeric(Split(RS("product_class_id_extend"),",")(i777)) Then
					tmpString03471=tmpString03471 &" OR "& Coke.strSQL_ProductClassALL(CokeShow.CokeClng( Split(RS("product_class_id_extend"),",")(i777) ))
				End If
			Next
		  End If
		   
			sqlInformation_SameTypesProductArea_Code="select top 5 * from [CXBG_product] where deleted=0 and isOnsale=1 AND ( "& Coke.strSQL_ProductClassALL(RS("product_class_id")) & tmpString03471 &" ) and id<>"& RS("id") &" order by OrderID desc,id desc"
'//response.Write sqlInformation_SameTypesProductArea_Code
			'sqlInformation_SameTypesProductArea_Code="select top 5 * from [CXBG_product__BoundProduct] where isOnsale=1 and is_display_newproduct=1 order by newid()"
			Set rsInformation_SameTypesProductArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsInformation_SameTypesProductArea_Code.Open sqlInformation_SameTypesProductArea_Code,CONN,1,1
			countInformation_SameTypesProductArea_Code=rsInformation_SameTypesProductArea_Code.RecordCount
			numInformation_SameTypesProductArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsInformation_SameTypesProductArea_Code.EOF Then
		  %>
		  欢迎光临，餐厅尚无相关的同菜系菜品推荐：）
		  <%
		  End If
		  %>
          
          <%
		  Do While Not rsInformation_SameTypesProductArea_Code.EOF
		  %>
         <a style="height:110px" href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsInformation_SameTypesProductArea_Code("id") ) %>" target="_blank" title="<% =rsInformation_SameTypesProductArea_Code("ProductName") %>"><img src="<% If rsInformation_SameTypesProductArea_Code("Photo")<>"" Then Response.Write rsInformation_SameTypesProductArea_Code("photo") Else Response.Write "/images/NoPic.png" %>" width="80" height="80" /> <% =rsInformation_SameTypesProductArea_Code("ProductName") %></a>
          <%
			  numInformation_SameTypesProductArea_Code=numInformation_SameTypesProductArea_Code+1
			  rsInformation_SameTypesProductArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsInformation_SameTypesProductArea_Code.Close
		  Set rsInformation_SameTypesProductArea_Code=Nothing
		  %>