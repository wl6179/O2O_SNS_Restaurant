					<%
					'
					'私有的内嵌文件，属于菜单01——系统管理
					'
					Dim MenuName
					MenuName = "系统管理"
					%>
					<a href="controller.asp">网站资料设置</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"controller.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					
					<a href="supervisor.asp">管理员帐号管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"supervisor.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					<a href="supervisor_class.asp">管理员帐号分类设置</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"supervisor_class.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					
					
					<a href="account_class.asp">会员帐号等级设置</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"account_class.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					<a href="attribute_work.asp">属性——职业管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"attribute_work.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                    <a href="attribute_schooling.asp">属性——学历管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"attribute_schooling.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                    <a href="attribute_income.asp">属性——收入管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"attribute_income.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                    
                    <a href="attribute_jifensystem.asp">积分名目管理(禁)</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"attribute_jifensystem.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                    
                    <% If Session("enterName")="coke" Then %>
                    <a href="log.asp">日志管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"log.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                    <% End If %>
					<br />
					<br />