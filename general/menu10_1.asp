					<%
					'
					'私有的内嵌文件，属于菜单02——帐号&会员管理
					'
					Dim MenuName
					MenuName = "娱乐项目发布管理"
					%>
					<a href="Game_edit.asp">游戏发布管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"game_edit.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %>
					<br />
					<a href="GiftCertificated_edit.asp">礼品券发布管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"giftcertificated_edit.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                    <a href="Questionnaire_edit.asp">调查问卷管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"questionnaire_edit.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					
					<br />
					
					