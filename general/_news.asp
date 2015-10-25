		<%
		'
		'公有的内嵌文件，最右栏的最新提示模块.
		'
		%>
		<!--rightInfo-->
		<%
		'
		'自动检测显示屏宽度并处理：
		'当屏幕太小小于等于1024、还有新提示都完成时，自动处理消除此最右栏.
		'
		%>
		
		
		<%
		'如果有事务需要处理，则显示！否则消失其右栏.
		If isHaveWork=True Then
		%>
		<div class="rightInfo">
			
			
			<%
			'先统计一下.
			Dim RSCOUNT_Application_News,COUNT_Application_News
			Set RSCOUNT_Application_News=CONN.Execute("SELECT COUNT(*) FROM [CokeShow_Application]")
			COUNT_Application_News=RSCOUNT_Application_News(0)
			%>
			<h2>&nbsp;申请预约提示(<% =COUNT_Application_News %>)</h2>
			<!--
			今天的所需处理记录.
			-->
				<p>
					<%
					'咨询订阅记录
					Dim sql_Application_News,RS_Application_News
					sql_Application_News="SELECT TOP 6 * FROM [CokeShow_Application] ORDER BY id DESC"
					Set RS_Application_News=Server.CreateObject("Adodb.RecordSet")
					RS_Application_News.Open sql_Application_News,CONN,1,1
					
					'记录集
					If Not RS_Application_News.Eof Then
					Do While Not RS_Application_News.Eof
					%>
						
						<% =RS_Application_News("addDate") %>您有预约:<br /><a href="#" title="EmailState:<% =RS_Application_News("EmailState") %>,isDelete:<% =RS_Application_News("isDelete") %>,modifyDate:<% =RS_Application_News("modifyDate") %><br />clienttelephone:<% =RS_Application_News("clienttelephone") %>,clientoptimaltelephonetime:<% =RS_Application_News("clientoptimaltelephonetime") %>,clientaddress:<% =RS_Application_News("clientaddress") %>,clientoptimaladdresstime:<% =RS_Application_News("clientoptimaladdresstime") %>,clientoptimalcontactway:<% =RS_Application_News("clientoptimalcontactway") %>,clientfavoritecolor:<% =RS_Application_News("clientfavoritecolor") %>,clienthowtoknow:<% =RS_Application_News("clienthowtoknow") %>,clientobjective:<% =RS_Application_News("clientobjective") %>,clientadditional:<% =RS_Application_News("clientadditional") %>,province:<% =RS_Application_News("province") %>,city:<% =RS_Application_News("city") %>"><% =RS_Application_News("clientname")&RS_Application_News("clientgender") %>(<% =RS_Application_News("email") %>)</a><br />
						
					<%
						RS_Application_News.MoveNext
					Loop
					End If
					
					RS_Application_News.Close
					Set RS_Application_News=Nothing
					%>
					
					
				</p>
				<p><a href="#">&#187;&#187;更多...</a></p>
			
			
			
			
			<%
			'先统计一下.
			Dim RSCOUNT_Report_News,COUNT_Report_News
			Set RSCOUNT_Report_News=CONN.Execute("SELECT COUNT(*) FROM [CokeShow_SubscriptionReport]")
			COUNT_Report_News=RSCOUNT_Report_News(0)
			%>
			<h2>&nbsp;订阅资讯播报提示(<% =COUNT_Report_News %>)</h2>
				
				<p>
					<%
					'咨询订阅记录
					Dim sql_Report_News,RS_Report_News
					sql_Report_News="SELECT TOP 6 * FROM [CokeShow_SubscriptionReport] ORDER BY id DESC"
					Set RS_Report_News=Server.CreateObject("Adodb.RecordSet")
					RS_Report_News.Open sql_Report_News,CONN,1,1
					
					'记录集
					If Not RS_Report_News.Eof Then
					Do While Not RS_Report_News.Eof
					%>
						
						<% =RS_Report_News("addDate") %><br /><a href="#" title="EmailState:<% =RS_Report_News("EmailState") %>,isDelete:<% =RS_Report_News("isDelete") %>"><% =RS_Report_News("email") %></a><br />
						
					<%
						RS_Report_News.MoveNext
					Loop
					End If
					
					RS_Report_News.Close
					Set RS_Report_News=Nothing
					%>
				</p>
				<p><a href="#">&#187;&#187;更多...</a></p>
			
			
			
			
			<%
			'先统计一下.
			Dim RSCOUNT_TelltoFriend_News,COUNT_TelltoFriend_News
			Set RSCOUNT_TelltoFriend_News=CONN.Execute("SELECT COUNT(*) FROM [CokeShow_SubscriptionTelltoFriend]")
			COUNT_TelltoFriend_News=RSCOUNT_TelltoFriend_News(0)
			%>
			<h2>&nbsp;推荐好友提示(<% =COUNT_TelltoFriend_News %>)</h2>
				
				<p>
					<%
					'推荐好友记录
					Dim sql_TelltoFriend_News,RS_TelltoFriend_News
					sql_TelltoFriend_News="SELECT TOP 6 * FROM [CokeShow_SubscriptionTelltoFriend] ORDER BY id DESC"
					Set RS_TelltoFriend_News=Server.CreateObject("Adodb.RecordSet")
					RS_TelltoFriend_News.Open sql_TelltoFriend_News,CONN,1,1
					
					'记录集
					If Not RS_TelltoFriend_News.Eof Then
					Do While Not RS_TelltoFriend_News.Eof
					%>
						
						<% =RS_TelltoFriend_News("addDate") %><br /><a href="#" title="EmailState:<% =RS_TelltoFriend_News("EmailState") %>,isDelete:<% =RS_TelltoFriend_News("isDelete") %>"><% =RS_TelltoFriend_News("email") %></a><br />
						
					<%
						RS_TelltoFriend_News.MoveNext
					Loop
					End If
					
					RS_TelltoFriend_News.Close
					Set RS_TelltoFriend_News=Nothing
					%>
				</p>
				<p><a href="#">&#187;&#187;更多...</a></p>
			
			
		</div>
		<!--rightInfo-->
		
		<%
		End If
		%>