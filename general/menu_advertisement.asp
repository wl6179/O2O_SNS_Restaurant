					<%
					'
					'私有的内嵌文件，属于菜单02——帐号&会员管理
					'
					Dim MenuName
					MenuName = "广告管理"
					%>
					<a href="advertisement_main.asp">首页广告发布管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"advertisement_main.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                    <a href="advertisement_club.asp">会员Club广告发布管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"advertisement_club.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					<!--<a href="advertisement_featured.asp">首页专题广告管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"advertisement_featured.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					<a href="HotKeywords.asp">热门搜索管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"hotkeywords.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					<br />-->
					
					