					<%
					'
					'私有的内嵌文件，属于菜单02——帐号&会员管理
					'
					Dim MenuName
					MenuName = "(业务)互动管理"
					%>
					
                    
					<a href="account.asp">注册会员帐号管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"account.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					
					<a href="Message.asp">会员留言管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"message.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					
					<a href="tuijianpengyou.asp">推荐朋友列表</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"tuijianpengyou.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
					
                    <a href="Account_RemarkOns.asp">最新会员点评</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"account_remarkons.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                    <a href="Account_Favorites.asp">最新会员菜品收藏</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"account_favorites.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                    <a href="Account_GiftCertificateds.asp">已兑换礼品券</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"account_giftcertificateds.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                                        
                    <a href="Questionnaire_ResultView.asp">问卷调查结果</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"questionnaire_resultview.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                    
                    
					<br />
					
					<a href="FriendlyLink.asp">友情链接管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"friendlylink.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />
                    
                    <a href="VIPcardList.asp">VIP卡卡号录入管理</a><% If Instr( Lcase(Trim(Request.ServerVariables("SCRIPT_NAME"))) ,"vipcardlist.asp")>0 Then Response.Write " <img src=""/images/ok.gif"" />" %><br />