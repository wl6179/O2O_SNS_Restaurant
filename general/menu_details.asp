					<%
					'
					'私有的内嵌文件，属于菜单02——帐号&会员管理
					'
					Dim MenuName
					MenuName = "内容管理中心"
					%>
					<a href="details_edit.asp">内容管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"details_edit.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					<a href="details_class.asp">内容分类</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"details_class.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					<br />
					
					