					<%
					'
					'私有的内嵌文件，属于菜单02——帐号&会员管理
					'
					Dim MenuName
					MenuName = "餐厅环境管理"
					%>
					<a href="DiningArea_edit.asp">餐厅环境图片管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"diningarea_edit.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					<a href="DiningArea_class.asp">餐厅环境图片分类管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"diningarea_class.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					<br />
					
					