		<%
		  '月度明星.
		  Dim rsRightUnionCafeInformationArea_Code,sqlRightUnionCafeInformationArea_Code,countRightUnionCafeInformationArea_Code,numRightUnionCafeInformationArea_Code
		  	sqlRightUnionCafeInformationArea_Code="select top 3 * from [CXBG_details] where deleted=0 AND isOnpublic=1 AND details_class_id=8 ORDER BY isRecommend desc,details_orderid DESC,id DESC"
			Set rsRightUnionCafeInformationArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsRightUnionCafeInformationArea_Code.Open sqlRightUnionCafeInformationArea_Code,CONN,1,1
			countRightUnionCafeInformationArea_Code=rsRightUnionCafeInformationArea_Code.RecordCount
			numRightUnionCafeInformationArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsRightUnionCafeInformationArea_Code.EOF Then
		  %>
		  	欢迎光临，餐厅尚未上传联合餐厅优惠活动：）
		  <%
		  End If
		  %>

		  <%
          Do While Not rsRightUnionCafeInformationArea_Code.EOF
          %>
            <a href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsRightUnionCafeInformationArea_Code("id") ) %>" target="_blank" title="<% =rsRightUnionCafeInformationArea_Code("topic") %>">
              <span class="zjllimg"><img src="<% If rsRightUnionCafeInformationArea_Code("Photo")<>"" Then Response.Write rsRightUnionCafeInformationArea_Code("Photo") Else Response.Write "/images/NoPic.png" %>" width="45" height="45" /></span>
              <span class="zjlltitle" style="overflow: visible; line-height:normal;"><font class="fontred_a" style="white-space:normal; line-height:normal; line-height:18px;"><% =CokeShow.InterceptStr(rsRightUnionCafeInformationArea_Code("topic"),30) %></font></span>
            </a>
          <%
              numRightUnionCafeInformationArea_Code=numRightUnionCafeInformationArea_Code+1
              rsRightUnionCafeInformationArea_Code.MoveNext
          Loop
          
          '关闭记录集.
          rsRightUnionCafeInformationArea_Code.Close
          Set rsRightUnionCafeInformationArea_Code=Nothing
          %>