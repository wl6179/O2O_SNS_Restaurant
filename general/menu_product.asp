					<%
					'
					'私有的内嵌文件，属于菜单02——帐号&会员管理
					'
					Dim MenuName
					MenuName = "菜品管理"
					%>
					<a href="product.asp?Action=SearchNow">菜品高级查询</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"product.asp")>0 And Instr( Lcase(Trim(Request.ServerVariables("HTTP_url"))) ,"action=searchnow")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					<a href="product.asp">菜品列表</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"product.asp")>0 And Instr( Lcase(Trim(Request.ServerVariables("HTTP_url"))) ,"action=searchnow")<=0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					<a href="product_class.asp">菜品分类</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"product_class.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					
					
					<a href="product_businessUSE.asp">菜品所属菜系</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"product_businessuse.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					<a href="product_activityUSE.asp">菜品所属口味</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"product_activityuse.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                    <!--<a href="product_chiliIndex.asp">菜品辣椒指数</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"product_chiliindex.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />-->
					
					<br />
					
					