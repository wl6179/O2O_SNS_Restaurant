					<%
					'
					'私有的内嵌文件，属于菜单02——帐号&会员管理
					'
					Dim MenuName
					MenuName = "菜品管理"
					%>
					<a href="vip_client.asp">VIP客户列表</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"vip_client.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					
					<br />
					
					