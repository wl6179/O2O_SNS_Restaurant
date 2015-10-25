		  <%
		  '今日特价菜.
		  Dim rsTopOnSpecialTodayArea_Code,sqlTopOnSpecialTodayArea_Code,countTopOnSpecialTodayArea_Code,numTopOnSpecialTodayArea_Code
			'sqlTopOnSpecialTodayArea_Code="select top 6 * from [CXBG_product] where deleted=0 and isOnsale=1 and is_display_newproduct=1 order by OrderID desc,id desc"
			sqlTopOnSpecialTodayArea_Code="select top 6 * from [CXBG_product] where deleted=0 and isOnsale=1 and is_display_newproduct=1 order by newid()"
			Set rsTopOnSpecialTodayArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsTopOnSpecialTodayArea_Code.Open sqlTopOnSpecialTodayArea_Code,CONN,1,1
			countTopOnSpecialTodayArea_Code=rsTopOnSpecialTodayArea_Code.RecordCount
			numTopOnSpecialTodayArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsTopOnSpecialTodayArea_Code.EOF Then
		  %>
		  欢迎光临，餐厅尚未上传今日特价菜：）
		  <%
		  End If
		  %>
          
          <%
		  Do While Not rsTopOnSpecialTodayArea_Code.EOF
		  %>
   <a href="/ChineseDish/ChineseDishInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsTopOnSpecialTodayArea_Code("id") ) %>" target="_blank" title="<% =rsTopOnSpecialTodayArea_Code("ProductName") %>"><img src="<% If rsTopOnSpecialTodayArea_Code("Photo")<>"" Then Response.Write rsTopOnSpecialTodayArea_Code("NewProduct_Photo") Else Response.Write "/images/NoPic.png" %>" width="130" height="170" /><% =rsTopOnSpecialTodayArea_Code("ProductName") %></a>
          <%
			  numTopOnSpecialTodayArea_Code=numTopOnSpecialTodayArea_Code+1
			  rsTopOnSpecialTodayArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsTopOnSpecialTodayArea_Code.Close
		  Set rsTopOnSpecialTodayArea_Code=Nothing
		  %>