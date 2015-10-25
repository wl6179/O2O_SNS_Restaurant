		<%
		  '招兵买马.
		  Dim rsRightHotInformationArea_Code,sqlRightHotInformationArea_Code,countRightHotInformationArea_Code,numRightHotInformationArea_Code
		  	sqlRightHotInformationArea_Code="select top 8 * from [CXBG_details] where deleted=0 AND isOnpublic=1 AND details_class_id=26 ORDER BY details_orderid desc,id desc"
			Set rsRightHotInformationArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsRightHotInformationArea_Code.Open sqlRightHotInformationArea_Code,CONN,1,1
			countRightHotInformationArea_Code=rsRightHotInformationArea_Code.RecordCount
			numRightHotInformationArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsRightHotInformationArea_Code.EOF Then
		  %>
		  	欢迎光临，餐厅尚未有招兵买马信息：）
		  <%
		  End If
		  %>

		  <%
          Do While Not rsRightHotInformationArea_Code.EOF
          %>
            <li><a href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsRightHotInformationArea_Code("id") ) %>" target="_blank" title="<% =rsRightHotInformationArea_Code("topic") %>"><% =CokeShow.InterceptStr(rsRightHotInformationArea_Code("topic"),30) %></a></li>
            </a>
          <%
              numRightHotInformationArea_Code=numRightHotInformationArea_Code+1
              rsRightHotInformationArea_Code.MoveNext
          Loop
          
          '关闭记录集.
          rsRightHotInformationArea_Code.Close
          Set rsRightHotInformationArea_Code=Nothing
          %>