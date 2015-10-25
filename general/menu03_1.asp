					<%
					'
					'私有的内嵌文件，属于菜单02——帐号&会员管理
					'
					Dim MenuName
					MenuName = "栏目&amp;内容发布"
					%>
					<a href="column.asp">栏目结构管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"column.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %>
					&nbsp;<a href="column.asp?isDisplayAdvanceMenu=gogogo">高级操作</a><% If REQUEST("isDisplayAdvanceMenu")="gogogo" Then Response.Write " <img src=""/images/ok.gif"" />" %>
					<br />
					<a href="#">内容发布</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"content.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					
					<br />
					
					