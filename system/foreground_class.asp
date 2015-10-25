<%
'模块说明：系统类库.
'日期说明：2009-8-26
'版权说明：www.cokeshow.com.cn	可乐秀(北京)技术有限公司产品
'友情提示：本产品已经申请专利，请勿用于商业用途，如果需要帮助请联系可乐秀。
'联系电话：(010)-67659219
'电子邮件：cokeshow@qq.com
%>
<%
Class ForegroundClass
	Public reloadtime,userip,comeurl
	Private ABC
	
	
	Private Sub Class_Initialize()
        reloadtime = 10000
        userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")		'透过代理服务器取得客户端的真实IP地址.
        If userip = "" Then
			userip = Request.ServerVariables("REMOTE_ADDR")				'获取客户端IP.但如果客户端是使用代理服务器来访问，那取到的就是代理服务器的IP地址.
		End If
        comeurl = Trim(Request.ServerVariables("HTTP_REFERER"))			'可获得链接到此页面的来源地址，HTTP_REFERER可用于防外链.
		If Not IsObject(CONN) Then Call link_database()
		ABC="abc"
    End Sub
	
	Private Sub Class_Terminate()
    	On Error Resume Next
        	If IsObject(CONN) Then CONN.Close: Set CONN=Nothing
			If IsObject(RS) Then RS.Close: Set RS=Nothing
			If IsObject(rsClass) Then rsClass.Close: Set rsClass=Nothing
    End Sub
	
	
	Public Sub Start()
        
		'system_copyright = "<div style=""text-align:center;""><a href="""& system_owner_domain &""" target=""_blank""><img src=""/images/power.gif"" border=""0"" alt=""Powered By CokeShow.com.cn"" /></a></div>"
		'Response.Write system_copyright
    End Sub
	
	
	
	
'获取模块数据
'*************************************************************************
	'获取左侧栏目列表数据.（主要用于 首页左下侧、其它页的左下侧。）
	'参数：
	'1.CurrentClassid:当前页面的当前项的classid，如果匹配上了就要给当前项链接加上class="hover".
	'2.language:当前页面的language，筛选当前语言.
	Public Function ShowMenu(CurrentClassid,language)
		Dim ShowString
		ShowString=""
		Dim CurrentClassValue		'WL
		
		If isNull(language) Or language="" Then Exit Function
		
		'Begin
		'输出菜单结构.
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim sqlM3,RSM3
		Dim columnType_NowURL
		columnType_NowURL=""
		
		sqlM = "SELECT _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language "&_
				" FROM [CokeShow_column] AS _column "&_
					" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid) "&_
					
				" WHERE _column.Depth=0 "&_
					" AND _column_linkedName.languageCode='"& language &"' "&_
					" AND _column.isDisable=0 AND _column.isHidden=0 "&_
					
				" ORDER BY _column.RootID,_column.OrderID"
		Set RSM = CONN.Execute(sqlM)
		'//'是否第一次.
		Dim firstOfTimes,firstOfTimes_String
		firstOfTimes=True
		firstOfTimes_String=""
		
		If Not RSM.Eof Then
			'第一循环根项.
			Do While Not RSM.Eof
				'判断模式.(仅在第一次循环根项时进行判断，子项要全部按根项设定获得模式！WL)
				Select Case RSM("columnType")
				Case "standard"
				columnType_NowURL="CokeShow-Issue.asp"
				Case "list"
				columnType_NowURL="CokeShow-IssueList.asp"
				Case "photolist"
				columnType_NowURL="CokeShow-PhotoList.asp"
				Case Else
				columnType_NowURL="xxx.asp"
				End Select
				
				
				'如果是第一次，则消除上间距margin.
				If firstOfTimes=True Then firstOfTimes_String=" style=""margin:0px;"" " Else firstOfTimes_String=""
				'看看是否需要链接.
				If RSM("child")>0 Then
					ShowString=ShowString & "<h2 class=""service"""& firstOfTimes_String &">"& RSM("theName") &"</h2><ul>"
				Else
					
					CurrentClassValue=""		'WL
					If CurrentClassid=RSM("classid") Then CurrentClassValue=" class=""hover"" "		'WL
					ShowString=ShowString & "<h2 class=""service"""& firstOfTimes_String &"><a href="""& columnType_NowURL &"?classid="& RSM("classid") &"&language="& RSM("language") &""" "& CurrentClassValue &">"& RSM("theName") &"</a></h2><ul>"

If CokeShow.c10to2(0)=False Then CurrentClassid=""		''''''''--classid在此处消失了！！！！！！前台系统报废！！！！可用于版权检测环节！！！
				End If
				
				
				
				'第二循环相应的二级子项.
				sqlM2 = "SELECT _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language "&_
				" FROM [CokeShow_column] AS _column "&_
					" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid) "&_
					
				" WHERE _column.Depth=1 "&_
					" AND _column_linkedName.languageCode='"& language &"' "&_
					" AND _column.parentid='"& RSM("id") &"' "&_
					" AND _column.isDisable=0 AND _column.isHidden=0 "&_
					
				" ORDER BY _column.RootID,_column.OrderID"
				Set RSM2 = CONN.Execute(sqlM2)
				If Not RSM2.Eof Then
					Do While Not RSM2.Eof
						
						'如果有三级的子项，则当前依然为不可点击的(二级)栏目.
						If RSM2("child")>0 Then
							ShowString=ShowString & "<h3>"& RSM2("theName") &"</h3>"
							
							
							
							
							'只要当前(二级)项内还存在子项，那么还要继续循环出这些(三级)子项.
							'第三循环相应的三级子项.
							sqlM3 = "SELECT _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language "&_
							" FROM [CokeShow_column] AS _column "&_
								" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid) "&_
								
							" WHERE _column.Depth=2 "&_
								" AND _column_linkedName.languageCode='"& language &"' "&_
								" AND _column.parentid='"& RSM2("id") &"' "&_
								" AND _column.isDisable=0 AND _column.isHidden=0 "&_
								
							" ORDER BY _column.RootID,_column.OrderID"
							Set RSM3 = CONN.Execute(sqlM3)
							If Not RSM3.Eof Then
								Do While Not RSM3.Eof
									
									
									CurrentClassValue=""		'WL
									If CurrentClassid=RSM3("classid") Then CurrentClassValue=" class=""hover"" "		'WL
									'输出三级子项列表.
									ShowString=ShowString & "<li><a href="""& columnType_NowURL &"?classid="& RSM3("classid") &"&language="& RSM3("language") &""" "& CurrentClassValue &">"& RSM3("theName") &"</a></li>"
									CurrentClassValue=""		'WL
									
									RSM3.MoveNext
								Loop
							End If
							RSM3.Close
							
							
							
						'如果没有三级子项，则当前为可点击的栏目.
						Else
							
							CurrentClassValue=""		'WL
							If CurrentClassid=RSM2("classid") Then CurrentClassValue=" class=""hover"" "		'WL
							ShowString=ShowString & "<li><a href="""& columnType_NowURL &"?classid="& RSM2("classid") &"&language="& RSM2("language") &""" "& CurrentClassValue &">"& RSM2("theName") &"</a></li>"
							CurrentClassValue=""		'WL
						End If
						
						RSM2.MoveNext
					Loop
				End If
				RSM2.Close
				
				
				
				ShowString=ShowString & "</ul>"
				'第一次标记销毁.
				firstOfTimes=False
				RSM.MoveNext
				'到达最后时，结束ul标记
'//				If RSM.Eof Then Response.Write "</ul>"
			Loop
		End If
		
		RSM.Close

		Set RSM = Nothing
		
		Set RSM2 = Nothing
		Set RSM3 = Nothing
		'End
		
		ShowMenu=ShowString
	End Function
	
	
	'获取多国语言的 国旗图片和链接列表.（主要用于首页的导航处。）
	'参数：
	'1.language:当前页面的language，筛选当前语言，并显示选中状态.
	Public Function ShowLanguage_Index(language)
		Dim sql,RS
		Dim ShowString
		ShowString=""
		Dim TmpStr
		
		sql = "SELECT * FROM [CokeShow_language] ORDER BY RootID,OrderID"
		Set RS = CONN.Execute(sql)
		
		If Not RS.Eof Then
			Do While Not RS.Eof
				
				ShowString=ShowString & "<a href=""/index.asp?language="& CokeShow.ENDecode(RS("code")) &"""><img src="""& CokeShow.ENDecode(RS("photo")) &""""
				If language=CokeShow.ENDecode(RS("code")) Then ShowString=ShowString & " style=""border:#999999 3px solid; border-left:#cccccc 6px solid; border-right:#666666 8px solid;"""
				ShowString=ShowString & "  />"
				ShowString=ShowString & RS("classname")
				ShowString=ShowString & "</a>"
				ShowString=ShowString & "&nbsp;&nbsp;&nbsp;"
				
				
				RS.MoveNext
			Loop
		End If
		
		RS.Close
		Set RS = Nothing
		
		ShowLanguage_Index=ShowString
	End Function
	
	
	'获取多国语言的 国旗图片和链接列表.（主要用于详细页的导航处。）
	'参数：
	'1.CurrentClassid:当前页面的classid，如果匹配上了就要给链接加上class="hover".
	'2.language:当前页面的language，筛选当前语言，并显示选中状态.
	'3.当内容是列表中的内容时，光有classid和language两参数是不能唯一识别了，所以必须要有重要的随机数ContentMark来帮助识别众多相似的列表内容，只求唯一准确的识别.
	Public Function ShowLanguage_Issue(CurrentClassid,language,ContentMark)
		Dim sql,RS
		Dim ShowString
		ShowString=""
		Dim TmpStr
		
		'函数自我处理分页参数.
		Dim currentPage		'分页变量.
		currentPage		=CokeShow.filtRequest(Request("Page"))
		If currentPage<>"" Then
			If isNumeric(currentPage) Then currentPage=CokeShow.CokeClng(currentPage) Else currentPage=1
		Else
			currentPage=1
		End If
		
		'接收查询关键字.
		Dim keyword,requestKeyword
		keyword		=CokeShow.filtRequestSimple(Request("keyword"))
		'处理查询字符，决定是否传递此参数.
		If keyword<>"" Then
			requestKeyword	="&keyword="& Trim(keyword)
		Else
			requestKeyword	=""
		End If
		
		sql = "SELECT * FROM [CokeShow_language] ORDER BY RootID,OrderID"
		Set RS = CONN.Execute(sql)
		
		If Not RS.Eof Then
			Do While Not RS.Eof
				
				ShowString=ShowString & "<a href=""?page="& currentPage &"&language="& CokeShow.ENDecode(RS("code")) &"&classid="& CurrentClassid &"&content_mark="& ContentMark & requestKeyword &"""><img src="""& CokeShow.ENDecode(RS("photo")) &""""
				If language=CokeShow.ENDecode(RS("code")) Then ShowString=ShowString & " style=""border:#999999 3px solid; border-left:#cccccc 6px solid; border-right:#666666 8px solid;"""
				ShowString=ShowString & "  />"
				'显示文字.
				ShowString=ShowString & RS("classname")
				ShowString=ShowString & "</a>"
				ShowString=ShowString & "&nbsp;&nbsp;&nbsp;"
				
				
				RS.MoveNext
			Loop
		End If
		
		RS.Close
		Set RS = Nothing
		
		ShowLanguage_Issue=ShowString
	End Function
	
	
	'获取多国语言的 国旗图片和链接列表.（主要用于首页的导航处。）
	'参数：
	'1.language:当前页面的language，筛选当前语言，并显示选中状态.
	'2.CurrentURL:当前URL.
	Public Function ShowLanguage_Background(language,CurrentURL)
		Dim sql,RS
		Dim ShowString
		ShowString=""
		Dim TmpStr
		
		'替换掉URL里存在的language= ，不让它重复！
		Dim CurrentURL_A
		CurrentURL_A=CurrentURL
		CurrentURL_A=Replace(CurrentURL_A,"language=","languageCOKESHOW.COM.CN=")
		
		sql = "SELECT * FROM [CokeShow_language] ORDER BY RootID,OrderID"
		Set RS = CONN.Execute(sql)
		
		If Not RS.Eof Then
			Do While Not RS.Eof
				
				ShowString=ShowString & "<a href="""& CurrentURL_A &"&language="& CokeShow.ENDecode(RS("code")) &"""><img src="""& CokeShow.ENDecode(RS("photo")) &""""
				If language=CokeShow.ENDecode(RS("code")) Then ShowString=ShowString & " style=""border:#999999 3px solid; border-left:#cccccc 6px solid; border-right:#666666 8px solid;"""
				ShowString=ShowString & "  /></a>"
				'"& RS("classname") &"
				ShowString=ShowString & "&nbsp;&nbsp;&nbsp;"
				
				
				RS.MoveNext
			Loop
		End If
		
		RS.Close
		Set RS = Nothing
		
		ShowLanguage_Background=ShowString
	End Function
	
	
	
	'获取栏目列表数据.（主要用于借助参数，筛选出某一项其下的所有子项列表集合，运用于各处皆可。）
	'参数：
	'1.CurrentClassid:当前页面的classid，如果匹配上了就要给链接加上class="hover".
	'2.language:当前页面的language，筛选当前语言.
	'3.classid:当前项classid，用于循环其子项用的.
	'4.TopNum:设置显示最多几条记录.
	Public Function ShowSubMenu(CurrentClassid,language,classid,TopNum)
		Dim ShowString
		ShowString=""
		Dim CurrentClassValue		'WL
		
		If isNull(language) Or language="" Then Exit Function
		If Not isNumeric(classid) Then
			Exit Function
		Else
			If CokeShow.CokeClng(classid)=0 Then Exit Function
		End If
		
		'操作的栏目项，可以不是根项！
		Dim sqlTmp,RSTmp
		sqlTmp = "SELECT TOP "& TopNum &" id FROM [CokeShow_column] WHERE classid="& classid
		Set RSTmp = CONN.Execute(sqlTmp)
		
		
		Dim columnType_NowURL
		columnType_NowURL=""
		'判断模式.
		Dim RootColumnType
		RootColumnType=Coke.ReturnRootColumnType(classid)
		
		'判断模式.(仅在第一次循环根项时进行判断，子项要全部按根项设定获得模式！WL)
		Select Case RootColumnType
		Case "standard"
		columnType_NowURL="CokeShow-Issue.asp"
		Case "list"
		columnType_NowURL="CokeShow-IssueList.asp"
		Case "photolist"
		columnType_NowURL="CokeShow-PhotoList.asp"
		Case Else
		columnType_NowURL="xxx.asp"
		End Select
					
					
		'循环出子项集.
		If Not RSTmp.Eof Then
			'id用于匹配其子项中的PrarentID.
			'根据参数id输出其下级子菜单集合.
			Dim sqlM,RSM
			
			sqlM = "SELECT _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language "&_
					" FROM [CokeShow_column] AS _column "&_
						" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid) "&_
						
					" WHERE _column.Depth>0 "&_
						" AND _column_linkedName.languageCode='"& language &"' "&_
						" AND _column.parentid="& RSTmp("id") &_
						" AND _column.isDisable=0 AND _column.isHidden=0 "&_
						
					" ORDER BY _column.RootID,_column.OrderID"
			Set RSM = CONN.Execute(sqlM)
			
			If Not RSM.Eof Then
				'循环子项集.
				Do While Not RSM.Eof
					
					
					'看看是否需要链接.
					If RSM("child")>0 Then
						ShowString=ShowString & "<li><a href=""#"">"& RSM("theName") &"</a></li>"
					Else
						
						CurrentClassValue=""		'WL
						If CurrentClassid=RSM("classid") Then CurrentClassValue=" class=""hover"" "		'WL
						ShowString=ShowString & "<li><a href="""& columnType_NowURL &"?classid="& RSM("classid") &"&language="& RSM("language") &""" "& CurrentClassValue &">"& RSM("theName") &"</a></li>"
					End If
					
					RSM.MoveNext
				Loop
			End If
			
		End If
		
		RSM.Close
		Set RSM = Nothing
		
		RSTmp.Close
		Set RSTmp = Nothing
		
		
'		ShowSubMenu= sqlTmp
'		ShowSubMenu= sqlM&"555"
		ShowSubMenu=ShowString
	End Function
	
	
	'获取某单一栏目数据.（主要用于借助参数，筛选出某一个栏目，运用于各处皆可。）
	'参数：
	'1.language:当前语言.
	'2.classid:当前项的classid.
	'3.isLink:当前项是否需要链接，还是只显示文字而已.
	'4.LinkClassValue:设置a链接的class样式值.
	'5.LinkTarget:设置a链接的打开方式.
	'6.isShowTitle:设置是否显示a链接的title信息.
	Public Function ShowOnlyOneMenu(language,classid,isLink,LinkClassValue,LinkTarget,isShowTitle)
		Dim ShowString
		ShowString=""
		If isNull(language) Or language="" Then Exit Function
		If Not isNumeric(classid) Then
			Exit Function
		Else
			If CokeShow.CokeClng(classid)=0 Then Exit Function
		End If
		'处理参数
		If LinkClassValue="" Then
			LinkClassValue=""
		Else
			LinkClassValue=" class="""& LinkClassValue &""" "
		End If
		If LinkTarget="" Then
			LinkTarget=""
		Else
			LinkTarget=" target="""& LinkTarget &""" "
		End If
		
		
		
		'根据参数classid操作某一栏目项！
		Dim sql,RS
		sql = "SELECT _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language "&_
					" FROM [CokeShow_column] AS _column "&_
						" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid) "&_
						
					" WHERE 0=0 "&_
						" AND _column_linkedName.languageCode='"& language &"' "&_
						" AND _column.classid="& classid
						
		Set RS = CONN.Execute(sql)
		
		'如果有该国语言记录，则输出.
		If Not RS.Eof Then
			
			Dim columnType_NowURL
			columnType_NowURL=""
			'判断模式.
			Dim RootColumnType
			RootColumnType=Coke.ReturnRootColumnType(classid)
			
			Select Case RootColumnType
			Case "standard"
			columnType_NowURL="CokeShow-Issue.asp"
			Case "list"
			columnType_NowURL="CokeShow-IssueList.asp"
			Case "photolist"
			columnType_NowURL="CokeShow-PhotoList.asp"
			Case Else
			columnType_NowURL="xxx.asp"
			End Select
					
					
			'看看是否需要链接.
			If RS("child")>0 Then
				ShowString=ShowString & ""& RS("theName") &""
			Else
				'如果不是有子项的，那么就有链接，此时再次根据isLink参数决定要不要链接！
				If isLink=True Then
					Dim isShowTitle_go
					'是否为链接a标签显示title属性.
					If isShowTitle=True Then isShowTitle_go=" title="""& CokeShow.filt_astr(RS("theName"),1000) &""" " Else isShowTitle_go=""
					
					ShowString=ShowString & "<a href="""& columnType_NowURL &"?classid="& RS("classid") &"&language="& RS("language") &""" "& LinkClassValue & LinkTarget & isShowTitle_go &">"& RS("theName") &"</a>"
				Else
					ShowString=ShowString & ""& RS("theName") &""
				End If
			End If
			
			
		End If
		
		RS.Close
		Set RS = Nothing
		
		ShowOnlyOneMenu=ShowString
	End Function
	
	
	
	'获取当前详细页的记录的各项字段值.（主要用于借助参数，筛选出某一个具体发布内容记录的某一个字段，运用于详细页。）
	'参数：
	'0.SHOW_ContentField:指定要显示的Content发布内容记录的Field某一字段.
	'1.language:当前语言.
	'2.classid:当前项的classid.
	'3.isLink:当前项是否需要链接，还是只显示文字而已.
	'4.LinkClassValue:设置a链接的class样式值.
	'5.LinkTarget:设置a链接的打开方式.
	'6.isShowTitle:设置是否显示a链接的title信息（内容标题topic）.
	'7.ContentMark: 如果是coke，则是标准模式传来的请求，此时不理会ContentMark，选择Top 1记录就行了； 如果不是默认值coke，则是列表模式项产来的，此时要严格筛选参数ContentMark得到正确的详细记录信息；
	Public Function ShowContent(SHOW_ContentField,ContentMark,language,classid,isLink,LinkClassValue,LinkTarget,isShowTitle)
		Dim ShowString
		ShowString=""
		If isNull(language) Or language="" Then Exit Function
		If Not isNumeric(classid) Then
			Exit Function
		Else
			If CokeShow.CokeClng(classid)=0 Then Exit Function
		End If
		'处理参数
		If LinkClassValue="" Then
			LinkClassValue=""
		Else
			LinkClassValue=" class="""& LinkClassValue &""" "
		End If
		If LinkTarget="" Then
			LinkTarget=""
		Else
			LinkTarget=" target="""& LinkTarget &""" "
		End If
		'处理随机数ContentMark
		Dim KMMDS
		KMMDS=""
		
		'第一种情况，是标准模式传来的请求，只取TOP 1记录即可.
		If ContentMark="coke" Then
			KMMDS=" "
			
		'第二种情况，是列表模式传来的请求，由于相同记录太多，所以需要严格按随机数筛选记录.
		Else
			KMMDS=" AND _content.content_mark='"& ContentMark &"' "
			
		End If
		
		
		
		'根据参数classid，定位所有Content记录集！
		Dim sql,RS
		sql = "SELECT _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language, _content.* "&_
					" FROM [CokeShow_content] AS _content "&_
						" LEFT JOIN [CokeShow_column] AS _column ON (_content.classid=_column.classid) "&_
						" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid AND '"& language &"'=_column_linkedName.languageCode) "&_
						
					" WHERE 0=0 "&_
						" AND _content.languageCode='"& language &"' "&_
						" AND _content.classid="& classid &_
						KMMDS &_
					
					" ORDER BY _content.logid DESC "
		'response.Write sql
		Set RS = CONN.Execute(sql)
		
		'如果有该国语言记录，则输出.
		If Not RS.Eof Then
			
			Dim columnType_NowURL
			columnType_NowURL=""
			'只用standard模式.
			columnType_NowURL="CokeShow-Issue.asp"
					
			
			'如果不是有子项的，那么就有链接，此时再次根据isLink参数决定要不要链接！
			If isLink=True Then
				Dim isShowTitle_go
				If isShowTitle=True Then isShowTitle_go=" title="""& CokeShow.filt_astr(RS("topic"),1000) &""" " Else isShowTitle_go=""	'是否为链接a标签显示title属性.
				ShowString=ShowString & "<a href="""& columnType_NowURL &"?classid="& RS("classid") &"&language="& RS("language") &""" "& LinkClassValue & LinkTarget & isShowTitle_go &">"& RS(SHOW_ContentField) &"</a>"
			Else
				ShowString=ShowString & ""& RS(SHOW_ContentField) &""
			End If
			
		Else
			ShowString=ShowString & "Sorry,该语言下暂时没有发布任何内容."
		End If
		
		RS.Close
		Set RS = Nothing
		
		ShowContent=ShowString
	End Function
	
	
	'内容列表------------------------------------------------------------BEGIN
	'获取当前(列表)项的记录集，并且列出html代码列表出来，包括分页.
	'参数：
	'1.language:当前语言.
	'2.classid:要列出该classid的项中的所有内容，以列表形式展现出来！.
	'3.isLink:当前项是否需要链接，还是只显示文字而已.
	'4.LinkClassValue:设置a链接的class样式值.
	'5.LinkTarget:设置a链接的打开方式.
	'6.isShowTitle:设置是否显示a链接的title信息（内容标题topic）.
	'7.maxPerPage:分页参数.
	'8.UnitName:分页参数.
	Public Function ShowContentList(language,classid,isLink,LinkClassValue,LinkTarget,isShowTitle,maxPerPage,UnitName)
		Dim ShowString
		ShowString=""
		If isNull(language) Or language="" Then Exit Function
		If Not isNumeric(classid) Then
			Exit Function
		Else
			If CokeShow.CokeClng(classid)=0 Then Exit Function
		End If
		'处理参数
		If LinkClassValue="" Then
			LinkClassValue=""
		Else
			LinkClassValue=" class="""& LinkClassValue &""" "
		End If
		If LinkTarget="" Then
			LinkTarget=""
		Else
			LinkTarget=" target="""& LinkTarget &""" "
		End If
		
		'分页所需参数
		'strFileName,totalPut,maxPerPage,UnitName
		Dim strFileName,totalPut
		strFileName="CokeShow-IssueList.asp?language="& language &"&classid="& classid &""
		Dim currentPage
		currentPage		=CokeShow.filtRequest(Request("Page"))
		'接收传递参数，处理当前页码的控制变量，通过获取到的传值获取，默认为第一页1.
		If currentPage<>"" Then
			If isNumeric(currentPage) Then currentPage=CokeShow.CokeClng(currentPage) Else currentPage=1
		Else
			currentPage=1
		End If
		
		'分页所需参数
		
		
		
		'先为内容列表排序做准备——一切要按中文cn为准.
		Dim sql,RS
		Dim RSstr_content_mark
		
		'wl	初始化排序手段代码之计数器.
		Dim RSstr_content_mark2,nowCounter
		nowCounter=1
		'wl
		
		sql = "SELECT _content.content_mark "&_
					" FROM [CokeShow_content] AS _content "&_
						" LEFT JOIN [CokeShow_column] AS _column ON (_content.classid=_column.classid) "&_
						" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid AND 'cn'=_column_linkedName.languageCode) "&_
						
					" WHERE 0=0 "&_
						" AND _content.languageCode='cn' "&_
						" AND _content.classid="& classid &_
					
					" ORDER BY _content.logid DESC "
		'Response.Write sql
		Set RS = CONN.Execute(sql)
		If Not RS.Eof Then
			'记录下在中文处的顺序
			
			'wl	开始拼凑排序手段代码.
			RSstr_content_mark2=" Case _content.content_mark "
			'wl
			
			Do While Not RS.Eof
				
				'wl
				RSstr_content_mark2=RSstr_content_mark2 &" When '"& RS(0) &"' Then "& nowCounter &" "		'nowCounter从1开始计数.
				nowCounter=nowCounter+1
				'wl
				
				RSstr_content_mark=RSstr_content_mark &"'"& RS(0) &"'"
				
				RS.MoveNext
				If Not RS.Eof Then RSstr_content_mark=RSstr_content_mark & ","
			Loop
			
			'wl	结束拼凑排序手段代码.
			RSstr_content_mark2=RSstr_content_mark2 &" End As LogOrderID, "
			'wl
			
		End If
		RS.Close
		
		
		'根据参数classid，定位所有Content记录集！
		'//Dim sql,RS
		'因为排序有问题，所以区别对待语言是否为cn的2种情况吧.WL
'//		If Not language="cn" Then
		sql = "SELECT "& RSstr_content_mark2 &" _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language, _content.* "&_
					" FROM [CokeShow_content] AS _content "&_
						" LEFT JOIN [CokeShow_column] AS _column ON (_content.classid=_column.classid) "&_
						" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid AND '"& language &"'=_column_linkedName.languageCode) "&_
						
					" WHERE 0=0 "&_
						" AND _content.languageCode='"& language &"' "&_
						" AND _content.classid="& classid &_
						" AND _content.content_mark IN("& RSstr_content_mark &") "&_
						
					" ORDER BY LogOrderID ASC "
					'" ORDER BY _content.logid DESC "
'//		Else
'//		sql = "SELECT _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language, _content.* "&_
'//					" FROM [CokeShow_content] AS _content "&_
'//						" LEFT JOIN [CokeShow_column] AS _column ON (_content.classid=_column.classid) "&_
'//						" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid AND '"& language &"'=_column_linkedName.languageCode) "&_
'//						
'//					" WHERE 0=0 "&_
'//						" AND _content.languageCode='"& language &"' "&_
'//						" AND _content.classid="& classid &_
'//						
'//						
'//					" ORDER BY _content.logid DESC "
'//		End If
		'Response.Write sql
		'Set RS = CONN.Execute(sql)
		Set RS=Server.CreateObject("Adodb.RecordSet")
		RS.Open sql,CONN,1,1
		
		
		
		'如果有该国语言记录，则输出.
		If Not RS.Eof Then
			'获得记录总数(分页).
			totalPut=RS.RecordCount		
			
			Dim columnType_NowURL
			columnType_NowURL=""
			'只用standard模式.
			columnType_NowURL="CokeShow-Issue.asp"
			
			
			'输出分页
			'ShowString=ShowString & "<p class=""rightTxt3"">分页</p>"
			ShowString=ShowString &"<p class=""rightTxt3"" style=""text-align: right;  font-weight:bold;"">"& CokeShow.ShowPage(strFileName,totalPut,maxPerPage,True,True,UnitName) &"</p>"
			
			'------------------------------------------------处理页码+列表内容------------------------------------------
			If currentPage<1 Then
				currentPage=1
			End If
			'如果传递过来的Page当前页值很大，超过了应有的页数时，进行处理.
			If (currentPage-1) * maxPerPage > totalPut Then
				If (totalPut Mod maxPerPage)=0 Then
					'如果整好够页数，赋予当前页最大页.
					currentPage= totalPut \ maxPerPage
				Else
					'如果不整好，最有一页只有零散几条记录（不丰满的多余页），赋予当前页最大页.（不能整除情况下计算）
					currentPage= totalPut \ maxPerPage + 1
				End If
	
			End If
			If currentPage=1 Then
				
				'1
				'Call showMain
				ShowString=ShowString & ShowContentList_getLoopRS(RS,maxPerPage,isLink,isShowTitle,LinkClassValue,LinkTarget,columnType_NowURL)
				'1
				
			Else
				'如果传递过来的Page当前页值不大，在应有的页数范围之内时，理应(currentPage-1) * maxPerPage < totalPut，此时进行一些处理.
				if (currentPage-1) * maxPerPage < totalPut then
					'指针指到(currentPage-1)页（前一页）的最后一个记录处.
					RS.Move  (currentPage-1) * maxPerPage
					'RS.BookMark？
					Dim bookMark
					bookMark = RS.BookMark
					
					'2
					'Call showMain
					ShowString=ShowString & ShowContentList_getLoopRS(RS,maxPerPage,isLink,isShowTitle,LinkClassValue,LinkTarget,columnType_NowURL)
					'2
					
				else
				'如果传递过来的Page当前页值很大，超过了应有的页数时.打开第一页.
					currentPage=1
					
					'3
					'Call showMain
					ShowString=ShowString & ShowContentList_getLoopRS(RS,maxPerPage,isLink,isShowTitle,LinkClassValue,LinkTarget,columnType_NowURL)
					'3
					
				end if
			End If
			'------------------------------------------------处理页码+列表内容------------------------------------------
			
			
			
			
			'输出分页
			'ShowString=ShowString & "<p class=""rightTxt3"">分页</p>"
			ShowString=ShowString &"<p class=""rightTxt3"" style=""text-align: right; font-weight:bold;"">"& CokeShow.ShowPage(strFileName,totalPut,maxPerPage,True,True,UnitName) &"</p>"
			
		End If
		
		RS.Close
		Set RS = Nothing
		
		ShowContentList=ShowString
	End Function
	
	
	'附属于Function ShowContentList()的内容显示函数.
	'参数：
	'1.xxx:
	Public Function ShowContentList_getLoopRS(RS,maxPerPage,isLink,isShowTitle,LinkClassValue,LinkTarget,columnType_NowURL)
		Dim ShowString
		ShowString=""
		Dim i
		i=0
		
		'3
		Do While Not RS.Eof
			
			
			 ShowString=ShowString & "<h2 class=""rightTop"">"
				'--------------b
				'如果不是有子项的，那么就有链接，此时再次根据isLink参数决定要不要链接！
				If isLink=True Then
					Dim isShowTitle_go
					If isShowTitle=True Then isShowTitle_go=" title="""& CokeShow.filt_astr(RS("topic"),1000) &""" " Else isShowTitle_go=""	'是否为链接a标签显示title属性.
					ShowString=ShowString & "<a href="""& columnType_NowURL &"?classid="& RS("classid") &"&language="& RS("language") &"&content_mark="& RS("content_mark") &""" "& LinkClassValue & LinkTarget & isShowTitle_go &"><span style=""color:#ffffff;"">"& RS("topic") &"</span></a>"
				Else
					ShowString=ShowString & ""& RS("topic") &""
				End If
				'--------------e
			ShowString=ShowString & "</h2>"
			'控制显示字数.
			Dim theWordNum
			'如果是中、日、韩文，建议用较少的文字数限制.
			If Instr("cn,jp",language)>0 Then
				theWordNum=150
			'如果是英、美、法、德等欧美国家文字，建议用较多的文字数限制.
			ElseIf Instr("en",language) Then
				theWordNum=380
			'其余用标准而中等的字数限制.
			Else
				theWordNum=200
			End If
			ShowString=ShowString & "<p align=""justify"" class=""rightTxt3"" id=""field"& RS("logid") &"""><b>"& CokeShow.filt_astr(RS("logtext"),theWordNum) &"</b></p>"
			'ShowString=ShowString & "<p class=""rightTxt3""></p>"
			ShowString=ShowString & "<!--b-->"
			ShowString=ShowString & "<span dojoType=""dijit.Tooltip"" "
			ShowString=ShowString & "connectId=""field"& RS("logid") &""" "
			ShowString=ShowString & " style=display:none;>"
			'ShowString=ShowString & CokeShow.otherField_Function(("[CokeShow_language]",RS("languageCode"),"code","classname",False,1)
			ShowString=ShowString & "<img src="""& RS("photo") &""" onerror=""this.src='"& system_dir &"images/webPage2.jpg';"" onload=""resizepic(this,580);"" />"
			ShowString=ShowString & "</span>"
			ShowString=ShowString & "<!--e-->"
			
			
			i=i+1
			If i >= maxPerPage Then Exit Do
			RS.MoveNext
		Loop
		'3
		
		
		ShowContentList_getLoopRS=ShowString
	End Function
	'内容列表------------------------------------------------------------END
	
	
	'相册列表------------------------------------------------------------BEGIN
	'获取当前(列表)项的记录集，并且列出html代码列表出来，包括分页.
	'参数：
	'1.language:当前语言.
	'2.classid:要列出该classid的项中的所有内容，以列表形式展现出来！.
	'3.isLink:当前项是否需要链接，还是只显示文字而已.
	'4.LinkClassValue:设置a链接的class样式值.
	'5.LinkTarget:设置a链接的打开方式.
	'6.isShowTitle:设置是否显示a链接的title信息（内容标题topic）.
	'7.maxPerPage:分页参数.
	'8.UnitName:分页参数.
	Public Function ShowPhotoList(language,classid,isLink,LinkClassValue,LinkTarget,isShowTitle,maxPerPage,UnitName)
		Dim ShowString
		ShowString=""
		If isNull(language) Or language="" Then Exit Function
		If Not isNumeric(classid) Then
			Exit Function
		Else
			If CokeShow.CokeClng(classid)=0 Then Exit Function
		End If
		'处理参数
		If LinkClassValue="" Then
			LinkClassValue=""
		Else
			LinkClassValue=" class="""& LinkClassValue &""" "
		End If
		If LinkTarget="" Then
			LinkTarget=""
		Else
			LinkTarget=" target="""& LinkTarget &""" "
		End If
		
		'分页所需参数
		'strFileName,totalPut,maxPerPage,UnitName
		Dim strFileName,totalPut
		strFileName="CokeShow-PhotoList.asp?language="& language &"&classid="& classid &""
		Dim currentPage
		currentPage		=CokeShow.filtRequest(Request("Page"))
		'接收传递参数，处理当前页码的控制变量，通过获取到的传值获取，默认为第一页1.
		If currentPage<>"" Then
			If isNumeric(currentPage) Then currentPage=CokeShow.CokeClng(currentPage) Else currentPage=1
		Else
			currentPage=1
		End If
		
		'分页所需参数
		
		
		
		'先为内容列表排序做准备——一切要按中文cn为准.
		Dim sql,RS
		Dim RSstr_content_mark
		
		'wl	初始化排序手段代码之计数器.
		Dim RSstr_content_mark2,nowCounter
		nowCounter=1
		'wl
		
		sql = "SELECT _content.content_mark "&_
					" FROM [CokeShow_content] AS _content "&_
						" LEFT JOIN [CokeShow_column] AS _column ON (_content.classid=_column.classid) "&_
						" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid AND 'cn'=_column_linkedName.languageCode) "&_
						
					" WHERE 0=0 "&_
						" AND _content.languageCode='cn' "&_
						" AND _content.classid="& classid &_
					
					" ORDER BY _content.logid DESC "
		'Response.Write sql
		Set RS = CONN.Execute(sql)
		If Not RS.Eof Then
			'记录下在中文处的顺序
			
			'wl	开始拼凑排序手段代码.
			RSstr_content_mark2=" Case _content.content_mark "
			'wl
			
			Do While Not RS.Eof
				
				'wl
				RSstr_content_mark2=RSstr_content_mark2 &" When '"& RS(0) &"' Then "& nowCounter &" "		'nowCounter从1开始计数.
				nowCounter=nowCounter+1
				'wl
				
				RSstr_content_mark=RSstr_content_mark &"'"& RS(0) &"'"
				
				RS.MoveNext
				If Not RS.Eof Then RSstr_content_mark=RSstr_content_mark & ","
			Loop
			
			'wl	结束拼凑排序手段代码.
			RSstr_content_mark2=RSstr_content_mark2 &" End As LogOrderID, "
			'wl
			
		End If
		RS.Close
		
		
		'根据参数classid，定位所有Content记录集！
		'//Dim sql,RS
		'因为排序有问题，所以区别对待语言是否为cn的2种情况吧.WL
'//		If Not language="cn" Then
		sql = "SELECT "& RSstr_content_mark2 &" _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language, _content.* "&_
					" FROM [CokeShow_content] AS _content "&_
						" LEFT JOIN [CokeShow_column] AS _column ON (_content.classid=_column.classid) "&_
						" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid AND '"& language &"'=_column_linkedName.languageCode) "&_
						
					" WHERE 0=0 "&_
						" AND _content.languageCode='"& language &"' "&_
						" AND _content.classid="& classid &_
						" AND _content.content_mark IN("& RSstr_content_mark &") "&_
						
					" ORDER BY LogOrderID ASC "
					'" ORDER BY _content.logid DESC "
'//		Else
'//		sql = "SELECT _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language, _content.* "&_
'//					" FROM [CokeShow_content] AS _content "&_
'//						" LEFT JOIN [CokeShow_column] AS _column ON (_content.classid=_column.classid) "&_
'//						" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid AND '"& language &"'=_column_linkedName.languageCode) "&_
'//						
'//					" WHERE 0=0 "&_
'//						" AND _content.languageCode='"& language &"' "&_
'//						" AND _content.classid="& classid &_
'//						
'//						
'//					" ORDER BY _content.logid DESC "
'//		End If
		'//Response.Write sql
		'Set RS = CONN.Execute(sql)
		Set RS=Server.CreateObject("Adodb.RecordSet")
		RS.Open sql,CONN,1,1
		
		
		
		'如果有该国语言记录，则输出.
		If Not RS.Eof Then
			'获得记录总数(分页).
			totalPut=RS.RecordCount		
			
			Dim columnType_NowURL
			columnType_NowURL=""
			'只用standard模式.
			columnType_NowURL="CokeShow-Photo.asp"
			
			
			'输出分页
			'ShowString=ShowString & "<p class=""rightTxt3"">分页</p>"
			ShowString=ShowString &"<div><p class=""rightTxt3"" style=""text-align:right;""><b>"& CokeShow.ShowPage(strFileName,totalPut,maxPerPage,True,True,UnitName) &"</b></p></div>"
			ShowString=ShowString &"<br style=""clear:both;"" />"
			
			'------------------------------------------------处理页码+列表内容------------------------------------------
			If currentPage<1 Then
				currentPage=1
			End If
			'如果传递过来的Page当前页值很大，超过了应有的页数时，进行处理.
			If (currentPage-1) * maxPerPage > totalPut Then
				If (totalPut Mod maxPerPage)=0 Then
					'如果整好够页数，赋予当前页最大页.
					currentPage= totalPut \ maxPerPage
				Else
					'如果不整好，最有一页只有零散几条记录（不丰满的多余页），赋予当前页最大页.（不能整除情况下计算）
					currentPage= totalPut \ maxPerPage + 1
				End If
	
			End If
			If currentPage=1 Then
				
				'1
				'Call showMain
				ShowString=ShowString & ShowPhotoList_getLoopRS(RS,maxPerPage,isLink,isShowTitle,LinkClassValue,LinkTarget,columnType_NowURL)
				'1
				
			Else
				'如果传递过来的Page当前页值不大，在应有的页数范围之内时，理应(currentPage-1) * maxPerPage < totalPut，此时进行一些处理.
				if (currentPage-1) * maxPerPage < totalPut then
					'指针指到(currentPage-1)页（前一页）的最后一个记录处.
					RS.Move  (currentPage-1) * maxPerPage
					'RS.BookMark？
					Dim bookMark
					bookMark = RS.BookMark
					
					'2
					'Call showMain
					ShowString=ShowString & ShowPhotoList_getLoopRS(RS,maxPerPage,isLink,isShowTitle,LinkClassValue,LinkTarget,columnType_NowURL)
					'2
					
				else
				'如果传递过来的Page当前页值很大，超过了应有的页数时.打开第一页.
					currentPage=1
					
					'3
					'Call showMain
					ShowString=ShowString & ShowPhotoList_getLoopRS(RS,maxPerPage,isLink,isShowTitle,LinkClassValue,LinkTarget,columnType_NowURL)
					'3
					
				end if
			End If
			'------------------------------------------------处理页码+列表内容------------------------------------------
			
			
			
			
			'输出分页
			'ShowString=ShowString & "<p class=""rightTxt3"">分页</p>"
			ShowString=ShowString &"<div><p class=""rightTxt3"" style=""text-align: right; font-weight:bold;"">"& CokeShow.ShowPage(strFileName,totalPut,maxPerPage,True,True,UnitName) &"</p></div>"
			
		End If
		
		RS.Close
		Set RS = Nothing
		
		ShowPhotoList=ShowString
	End Function
	
	
	'附属于Function ShowContentList()的内容显示函数.
	'参数：
	'1.xxx:
	Public Function ShowPhotoList_getLoopRS(RS,maxPerPage,isLink,isShowTitle,LinkClassValue,LinkTarget,columnType_NowURL)
		Dim ShowString
		ShowString=""
		Dim i
		i=0
		
		ShowString=ShowString & "<div><ul class=""PhotoListUL"">"
		'3
		Do While Not RS.Eof
			 
			'控制显示字数.
			 Dim theWordNum
			'如果是中、日、韩文，建议用较少的文字数限制.
			 If Instr("cn,jp",language)>0 Then
			 	theWordNum=14
			'如果是英、美、法、德等欧美国家文字，建议用较多的文字数限制.
			 ElseIf Instr("en",language) Then
			 	theWordNum=23
			'其余用标准而中等的字数限制.
			 Else
			 	theWordNum=30
			 End If
			 ShowString=ShowString & "<li>"
			 	
				'ShowString=ShowString & CokeShow.filt_astr(RS("logtext"),theWordNum)
			 	ShowString=ShowString & "<img id=""photofield"& RS("logid") &""" src="""& RS("photo") &""" width=""300""   onerror=""this.src='"& system_dir &"images/webPage2.jpg';""  />"
				ShowString=ShowString & "<br />"
			 	'--------------b
				'如果不是有子项的，那么就有链接，此时再次根据isLink参数决定要不要链接！
				If isLink=True Then
					Dim isShowTitle_go
					If isShowTitle=True Then isShowTitle_go=" title="""& CokeShow.filt_astr(RS("topic"),1000) &""" " Else isShowTitle_go=""	'是否为链接a标签显示title属性.
					ShowString=ShowString & "<a href="""& columnType_NowURL &"?classid="& RS("classid") &"&language="& RS("language") &"&content_mark="& RS("content_mark") &""" "& LinkClassValue & LinkTarget & isShowTitle_go &">"& CokeShow.filt_astr(RS("topic"),theWordNum) &"</a>"
				Else
					ShowString=ShowString & ""& RS("topic") &""
				End If
				'--------------e
			 	
			 
				'ShowString=ShowString & "<p class=""rightTxt3""></p>"
				 ShowString=ShowString & "<!--b-->"
				 ShowString=ShowString & "<span dojoType=""dijit.Tooltip"" "
				 ShowString=ShowString & "connectId=""photofield"& RS("logid") &""" "
				 ShowString=ShowString & " style=display:none;>"
				'ShowString=ShowString & CokeShow.otherField_Function(("[CokeShow_language]",RS("languageCode"),"code","classname",False,1)
				 ShowString=ShowString & "<img src="""& RS("photo") &""" onerror=""this.src='"& system_dir &"images/webPage2.jpg';"" onload=""resizepic(this,680);"" />"
				 ShowString=ShowString & "</span>"
				 ShowString=ShowString & "<!--e-->"
			
			 ShowString=ShowString & "</li>"
			
			i=i+1
			If i >= maxPerPage Then Exit Do
			RS.MoveNext
		Loop
		'3
		ShowString=ShowString & "</ul></div>"
		ShowString=ShowString &"<br style=""clear:both;"" />"
		
		
		ShowPhotoList_getLoopRS=ShowString
	End Function
	'相册列表------------------------------------------------------------END
	
	
	
	
	
	'全文搜索内容列表------------------------------------------------------------BEGIN
	'获取当前(列表)项的记录集，并且列出html代码列表出来，包括分页.
	'参数：
	'1.language:当前语言.
	'2.classid:要列出该classid的项中的所有内容，以列表形式展现出来！.
	'3.isLink:当前项是否需要链接，还是只显示文字而已.
	'4.LinkClassValue:设置a链接的class样式值.
	'5.LinkTarget:设置a链接的打开方式.
	'6.isShowTitle:设置是否显示a链接的title信息（内容标题topic）.
	'7.maxPerPage:分页参数.
	'8.UnitName:分页参数.
	Public Function ShowSearchingContentList(language,classid,isLink,LinkClassValue,LinkTarget,isShowTitle,maxPerPage,UnitName)
		Dim ShowString
		ShowString=""
		'//If isNull(language) Or language="" Then Exit Function
'//		If Not isNumeric(classid) Then
'//			Exit Function
'//		Else
'//			If CokeShow.CokeClng(classid)=0 Then Exit Function
'//		End If
		'处理参数
		If LinkClassValue="" Then
			LinkClassValue=""
		Else
			LinkClassValue=" class="""& LinkClassValue &""" "
		End If
		If LinkTarget="" Then
			LinkTarget=""
		Else
			LinkTarget=" target="""& LinkTarget &""" "
		End If
		
		'分页所需参数
		'strFileName,totalPut,maxPerPage,UnitName
		Dim strFileName,totalPut
		strFileName="CokeShow-SearchIssueList.asp?language="& language &"&classid="& classid &""
		Dim currentPage
		currentPage		=CokeShow.filtRequest(Request("Page"))
		'接收传递参数，处理当前页码的控制变量，通过获取到的传值获取，默认为第一页1.
		If currentPage<>"" Then
			If isNumeric(currentPage) Then currentPage=CokeShow.CokeClng(currentPage) Else currentPage=1
		Else
			currentPage=1
		End If
		
		'分页所需参数
		Dim keyword
		keyword			=CokeShow.filtRequestSimple(Request("keyword"))
		strFileName=strFileName &"&keyword="& keyword
		
		Dim Action
		Action			=CokeShow.filtRequest(Request("Action"))
		strFileName=strFileName &"&Action="& Action
		
		
		'先为内容列表排序做准备——一切要按中文cn为准.
		Dim sql,RS
		Dim RSstr_content_mark
		
		'wl	初始化排序手段代码之计数器.
		Dim RSstr_content_mark2,nowCounter
		nowCounter=1
		'wl
		sql = "SELECT _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language, _content.* "&_
					" FROM [CokeShow_content] AS _content "&_
						" LEFT JOIN [CokeShow_column] AS _column ON (_content.classid=_column.classid) "&_
						" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid AND _content.languageCode=_column_linkedName.languageCode) "&_
						
					" WHERE 0=0 AND _content.deleted=0 "&_
						
						
						" AND (_content.logtext LIKE '%"& keyword &"%' OR _content.topic LIKE '%"& keyword &"%') "&_
					" ORDER BY _content.logid DESC "
					
					
					'" AND _content.languageCode='"& language &"' "&_
					'" AND _content.classid="& classid &_
					'" AND _content.content_mark IN("& RSstr_content_mark &") "&_
		'Response.Write sql
		'ShowSearchingContentList=sql
		'exit function
		
		Set RS=Server.CreateObject("Adodb.RecordSet")
		RS.Open sql,CONN,1,1
		
		
		
		'如果有该国语言记录，则输出.
		If Not RS.Eof Then
			'获得记录总数(分页).
			totalPut=RS.RecordCount		
			
			Dim columnType_NowURL
			columnType_NowURL=""
			'只用standard模式.
			columnType_NowURL="CokeShow-Issue.asp"
			
			
			'输出分页
			'ShowString=ShowString & "<p class=""rightTxt3"">分页</p>"
			ShowString=ShowString &"<p class=""rightTxt3"" style=""text-align: right;  font-weight:bold;"">"& CokeShow.ShowPage(strFileName,totalPut,maxPerPage,True,True,UnitName) &"</p>"
			
			'------------------------------------------------处理页码+列表内容------------------------------------------
			If currentPage<1 Then
				currentPage=1
			End If
			'如果传递过来的Page当前页值很大，超过了应有的页数时，进行处理.
			If (currentPage-1) * maxPerPage > totalPut Then
				If (totalPut Mod maxPerPage)=0 Then
					'如果整好够页数，赋予当前页最大页.
					currentPage= totalPut \ maxPerPage
				Else
					'如果不整好，最有一页只有零散几条记录（不丰满的多余页），赋予当前页最大页.（不能整除情况下计算）
					currentPage= totalPut \ maxPerPage + 1
				End If
	
			End If
			If currentPage=1 Then
				
				'1
				'Call showMain
				ShowString=ShowString & ShowSearchingContentList_getLoopRS(RS,maxPerPage,isLink,isShowTitle,LinkClassValue,LinkTarget,columnType_NowURL)
				'1
				
			Else
				'如果传递过来的Page当前页值不大，在应有的页数范围之内时，理应(currentPage-1) * maxPerPage < totalPut，此时进行一些处理.
				if (currentPage-1) * maxPerPage < totalPut then
					'指针指到(currentPage-1)页（前一页）的最后一个记录处.
					RS.Move  (currentPage-1) * maxPerPage
					'RS.BookMark？
					Dim bookMark
					bookMark = RS.BookMark
					
					'2
					'Call showMain
					ShowString=ShowString & ShowSearchingContentList_getLoopRS(RS,maxPerPage,isLink,isShowTitle,LinkClassValue,LinkTarget,columnType_NowURL)
					'2
					
				else
				'如果传递过来的Page当前页值很大，超过了应有的页数时.打开第一页.
					currentPage=1
					
					'3
					'Call showMain
					ShowString=ShowString & ShowSearchingContentList_getLoopRS(RS,maxPerPage,isLink,isShowTitle,LinkClassValue,LinkTarget,columnType_NowURL)
					'3
					
				end if
			End If
			'------------------------------------------------处理页码+列表内容------------------------------------------
			
			
			
			
			'输出分页
			'ShowString=ShowString & "<p class=""rightTxt3"">分页</p>"
			ShowString=ShowString &"<p class=""rightTxt3"" style=""text-align: right; font-weight:bold;"">"& CokeShow.ShowPage(strFileName,totalPut,maxPerPage,True,True,UnitName) &"</p>"
			
		End If
		
		RS.Close
		Set RS = Nothing
		
		ShowSearchingContentList=ShowString
	End Function
	
	
	'附属于Function ShowSearchingContentList()的内容显示函数.
	'参数：
	'1.xxx:
	Public Function ShowSearchingContentList_getLoopRS(RS,maxPerPage,isLink,isShowTitle,LinkClassValue,LinkTarget,columnType_NowURL)
		Dim ShowString
		ShowString=""
		Dim i
		i=0
		
		'3
		Do While Not RS.Eof
			
			
			 ShowString=ShowString & "<h2 class=""rightTop"">"
				'--------------b
				'如果不是有子项的，那么就有链接，此时再次根据isLink参数决定要不要链接！
				If isLink=True Then
					Dim isShowTitle_go
					If isShowTitle=True Then isShowTitle_go=" title="""& CokeShow.filt_astr(RS("topic"),1000) &""" " Else isShowTitle_go=""	'是否为链接a标签显示title属性.
					ShowString=ShowString & "<a href="""& columnType_NowURL &"?classid="& RS("classid") &"&language="& RS("language") &"&content_mark="& RS("content_mark") &""" "& LinkClassValue & LinkTarget & isShowTitle_go &"><span style=""color:#ffffff;"">"& RS("topic") &"</span></a>"
				Else
					ShowString=ShowString & ""& RS("topic") &""
				End If
				'--------------e
			 ShowString=ShowString & "</h2>"
			'控制显示字数.
			 Dim theWordNum
			'如果是中、日、韩文，建议用较少的文字数限制.
			 If Instr("cn,jp",language)>0 Then
			 	theWordNum=150
			'如果是英、美、法、德等欧美国家文字，建议用较多的文字数限制.
			 ElseIf Instr("en",language) Then
			 	theWordNum=380
			'其余用标准而中等的字数限制.
			 Else
			 	theWordNum=200
			 End If
			  
			 
			 
			 ShowString=ShowString & "<p align=""justify"" class=""rightTxt3"" id=""field"& RS("logid") &"""><b>"
			 ShowString=ShowString & CokeShow.filt_astr(RS("logtext"),theWordNum)
			'输出当前所属语言.
			 ShowString=ShowString &"&nbsp;"& CokeShow.otherField("[CokeShow_language]",CokeShow.ENEncode(RS("language")),"code","photo",False,2)
			'输出当前所属语言.
			 ShowString=ShowString & "</b></p>"
			 
			'ShowString=ShowString & "<p class=""rightTxt3""></p>"
			 ShowString=ShowString & "<!--b-->"
			 ShowString=ShowString & "<span dojoType=""dijit.Tooltip"" "
			 ShowString=ShowString & "connectId=""field"& RS("logid") &""" "
			 ShowString=ShowString & " style=display:none;>"
			'ShowString=ShowString & CokeShow.otherField_Function(("[CokeShow_language]",RS("languageCode"),"code","classname",False,1)
			 ShowString=ShowString & "<img src="""& RS("photo") &""" onerror=""this.src='"& system_dir &"images/webPage2.jpg';"" onload=""resizepic(this,580);"" />"
			 ShowString=ShowString & "</span>"
			 ShowString=ShowString & "<!--e-->"
			
			
			i=i+1
			If i >= maxPerPage Then Exit Do
			RS.MoveNext
		Loop
		'3
		
		
		ShowSearchingContentList_getLoopRS=ShowString
	End Function
	'全文搜索内容列表------------------------------------------------------------END
	
'*************************************************************************
	

	
	
'辅助函数专区
'*************************************************************************	
	'根据当前项classid，获取其根项的id值，最后得出根项所属的ColumnType值.
	'参数：
	'1.classid:当前项的标识classid，以此获得其根项的ColumnType值(模式).
	Public Function ReturnRootColumnType(classid)
		Dim ShowString
		ShowString=""
		
		'查询当前项根项记录.
		Dim sql,RS
		sql = "SELECT parentpath,Depth,columnType FROM [CokeShow_column] WHERE classid="& classid
		Set RS = CONN.Execute(sql)
		
		'如果有记录，开始处理各种可能情况.
		If Not RS.Eof Then
			'根据Depth判断当前是否为根项.
			'如果为根项，则直接获取ColumnType值.
			If RS("Depth")=0 Then
				ShowString=RS("columnType")
				
			'如果不是根项，则做进一步处理。
			Else
				'首先先根据当前项查询出来的parentpath结构，推算出根项id；然后根据此id查询出根项记录，并且获得其ColumnType值.
				Dim Root_TheIDValue
				Root_TheIDValue=Split(RS("parentpath"),",")(1)
				
				
				
				'查询当前项根项记录.
				Dim sql2,RS2
				sql2 = "SELECT columnType FROM [CokeShow_column] WHERE id="& Root_TheIDValue
				Set RS2 = CONN.Execute(sql2)
				
				'如果有记录，则获取其ColumnType值.
				If Not RS2.Eof Then
					ShowString=RS2("columnType")
				End If
				
				RS2.Close
				Set RS2=Nothing
				
				
			End If
		End If
		
		RS.Close
		Set RS=Nothing
		
		
		ReturnRootColumnType=ShowString
	End Function
	
	
	
	'内容之查看‘同类’的photo值。根据当前内容的language+classid+随机数参数，查看相对应的其它语言的相应记录，并看其是否有已经上传好的图片photo，如果有传回地址photo，如果没有传回空.
	'参数：
	'1.classid:当前内容的classid.
	'2.language:当前内容的languageCode.
	'3.content_mark:当前内容的content_mark随机数 (要不然是默认值coke《表明是标准模式的文章内容》，要不然就是list和photolist《表明是列表和相册模式下的某一内容》).
	'4.CurrentLogID:当前内容ID.
	'正在做........WL
	Public Function ReturnOtherLanguagePhoto_URLValue(classid,language,content_mark,CurrentLogID)
		Dim ShowString
		ShowString=""
		
		'查询当前内容的其余语言相应记录，如果有记录，就看看他们有没有photo.
		'有photo的话返回photo的URL，没有的话返回空.
		Dim sql,RS
		sql = "SELECT _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language, _content.* "&_
					" FROM [CokeShow_content] AS _content "&_
						" LEFT JOIN [CokeShow_column] AS _column ON (_content.classid=_column.classid) "&_
						" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid AND '"& language &"'=_column_linkedName.languageCode) "&_
						
					" WHERE 0=0 "&_
						
						" AND _content.classid="& classid &_
						
						" AND _content.content_mark='"& content_mark &"' "&_
						" AND _content.logid<>"& CurrentLogID &_
					" ORDER BY _content.logid DESC "
		Set RS = CONN.Execute(sql)
		
		Dim theHavePhotoString
		theHavePhotoString=""
		'如果有记录，开始处理各种可能情况.
		If Not RS.Eof Then
			Do While Not RS.Eof
				
				'如果传递过来的CurrentLogID=0，证明是添加操作中的全新记录添加操作.此时不用返回什么相应photo了，给它一个空即可.即不处理.
				'如果传递过来的CurrentLogID=9999999999，证明是已有记录的相应其他语言记录的添加操作，此时应该为其搜索photo.
				If CurrentLogID=0 Then
					
				Else
					'判断如果有记录有photo的话，返回photo的URL；没有的话返回空.
					If RS("photo")<>"" And Len(RS("photo"))>3 AND Instr(RS("photo"),".")>0 Then	'有后缀.符号证明有记录.
						theHavePhotoString=RS("photo")
					End If
					
				End If
				RS.MoveNext
			Loop
		End If
		
		RS.Close
		Set RS=Nothing
		
		ShowString=theHavePhotoString
		ReturnOtherLanguagePhoto_URLValue=ShowString
	End Function
	'(上边的姊妹篇)同步其它语言相应记录的photo字段的值，有一个相应记录有photo更新，则所有相应记录都会更新.
	'参数：
	'1.classid:当前内容的classid.
	'2.language:当前内容的languageCode.
	'3.content_mark:当前内容的content_mark随机数 (要不然是默认值coke《表明是标准模式的文章内容》，要不然就是list和photolist《表明是列表和相册模式下的某一内容》).
	'4.CurrentLogID:当前内容ID.
	'thePhotoURLValue_New:传入刚刚更新好的新photo值，此处将要把新值更新到其它相应记录中的photo中去.
	'正在做........WL
	Public Sub UpdateOtherLanguagePhoto_URLValue(classid,language,content_mark,CurrentLogID,thePhotoURLValue_New)
		Dim ShowString
		ShowString=""
		
		'查询当前内容的其余语言相应记录，如果有记录，就看看他们有没有photo.
		'有photo的话返回photo的URL，没有的话返回空.
		Dim sql,RS
		sql = "SELECT _column.*, _column_linkedName.displayname AS theName, _column_linkedName.languageCode AS language, _content.* "&_
					" FROM [CokeShow_content] AS _content "&_
						" LEFT JOIN [CokeShow_column] AS _column ON (_content.classid=_column.classid) "&_
						" LEFT JOIN [CokeShow_column_linkedName] AS _column_linkedName ON (_column.classid=_column_linkedName.linked_classid AND '"& language &"'=_column_linkedName.languageCode) "&_
						
					" WHERE 0=0 "&_
						
						" AND _content.classid="& classid &_
						
						" AND _content.content_mark='"& content_mark &"' "&_
						" AND _content.logid<>"& CurrentLogID &_
					" ORDER BY _content.logid DESC "
		Set RS = CONN.Execute(sql)
		
		'如果有记录，开始处理各种可能情况.
		If Not RS.Eof Then
			Do While Not RS.Eof
				
				'用于明确用delete命令时，为其清空所有相关记录的photo值.
				'正常改图时，更新相应记录的photo.
				If thePhotoURLValue_New="delete" Then thePhotoURLValue_New=""
				'更新相应记录中的pohto值为当前最新值thePhotoURLValue_New！
				CONN.Execute( "UPDATE [CokeShow_content] SET photo='"& thePhotoURLValue_New &"' WHERE logid="& RS("logid") )
				
				RS.MoveNext
			Loop
		End If
		
		RS.Close
		Set RS=Nothing
		
		
		
	End Sub
'*************************************************************************
	
	
	
	
	
	
	
	
'痴心不改餐厅前台模块调用类，方法集合
'*************************************************************************	
	
	
	
	
	'获取导航列表数据.（页头）
	'参数：
	'1.CurrentClassid	:当前页面所属的classid，如果匹配上了就要给链接加上class="mmvis".
	'2.TopNum			:设置最多显示几条记录.
	Public Function ShowNavigation(CurrentClassid,TopNum)
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim CurrentClassValue
		Dim sqlM,RSM
		Dim i_tmp
		Dim strTmpIndexString
		
		'初始化赋值.
		ShowString=""
		i_tmp=1
		
		'判断有各种效性.
		'//If isNull(language) Or isEmpty(language) Or language="" Then Exit Function
		If CurrentClassid=0 Then
			'CurrentClassid=0
			strTmpIndexString="<li><a class=""mmvis"" href=""/"" target=""_self"">商城首页</a></li>"
		Else
			CurrentClassid=CokeShow.CokeClng(CurrentClassid)
			strTmpIndexString="<li><a class=""menua"" href=""/"" target=""_self"">商城首页</a></li>"
		End If
		
		'-----------------Go Begin
		'列出导航记录
		sqlM = "SELECT TOP "& TopNum &" * "&_
				" FROM [CXBG_product_class] "&_
					
				" WHERE isNavigation=1 " &_
					
				" ORDER BY RootID,OrderID"
		Set RSM = CONN.Execute(sqlM)
		
		If Not RSM.Eof Then
			'循环记录集.
			Do While Not RSM.Eof
				'优先级1.
				If i_tmp=7 Then CurrentClassValue=" class=""menulpk"" "
				'优先级2.
				If CurrentClassid=RSM("classid") Then
					CurrentClassValue=" class=""mmvis"" "
				Else
					CurrentClassValue=" class=""menua"" "
				End If
				ShowString=ShowString & "<li><a "& CurrentClassValue &" href=""/ChineseDish/ChineseDish.Welcome?classid="& RSM("classid") &""" target=""_self"">"& RSM("classname") &"</a></li>"
				
				i_tmp=i_tmp+1
				RSM.MoveNext
			Loop
		End If
		
		If i_tmp<=7 Then
			Dim CurrNavNum,i_7
			CurrNavNum=7 + 1
			
			For i_7=1 To (CurrNavNum - i_tmp)
				If i_tmp=7 Then CurrentClassValue=" class=""menulpk"" "
				ShowString=ShowString & "<li><a "& CurrentClassValue &" href=""#"" target=""_self""></a></li>"
				
				i_tmp=i_tmp+1
			Next
		End If
		'-----------------Go End
		
		'终结化操作.
		RSM.Close
		Set RSM = Nothing
		ShowNavigation = strTmpIndexString & ShowString
		
	End Function
	
	'获取当前导航——所属各级分类.
	'参数：
	'CurrentClassid:当前菜品或者分类所属的分类classid.
	Public Function ShowNavigation_ForOnlyClass(CurrentClassid,CurrentStrFileName)
		ShowNavigation_ForOnlyClass=False
		'处理一下传参中的原classid，因为将会替换、并使用本函数中定义的新classid.
		If Instr( Lcase(Trim(CurrentStrFileName)) ,"classid=")>0 Then CurrentStrFileName=Replace(CurrentStrFileName,"classid=","classid$$$=")
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		
		'判断有各种效性.
		If isNumeric(CurrentClassid) Then
			CurrentClassid=CokeShow.CokeClng(CurrentClassid)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowNavigation_ForOnlyClass的参数ID不正确，无法获取当前导航的操作！"
			Exit Function
		End If
		'If isNull(language) Or isEmpty(language) Or language="" Then Exit Function
		
		'-----------------Go Begin
		'列出导航记录
		sqlM = "SELECT TOP 1 * "&_
				" FROM [CXBG_product_class] "&_
					
				" WHERE isShow=1 AND classid="& CurrentClassid &" " &_
					
				" ORDER BY RootID,OrderID"
		Set RSM = CONN.Execute(sqlM)
		
		'如果连当前分类都不存在，则输出空字符.
		If RSM.Eof Then
			'输出空字符串.
			ShowNavigation_ForOnlyClass=""
			Exit Function
		Else
			'如果有，则记录下当前分类字符，并作为待会儿的最后一级导航链接.
			strTmpString=strTmpString &" <span class=""font12000"">-</span> <a href="""& CurrentStrFileName &"&classid="& RSM("classid") &""" target=""_self"" class=""f00014"">"& RSM("classname") &"</a>"
		End If
		
		Dim tmpClassnameAddString
		'如果存在当前分类，则获取它的ParentPath字段值，并且循环ParentPath内的所有分类.
		For i_tmp=0 To Ubound( Split(RSM("ParentPath"),",") )
			'Split(RSM("ParentPath"),",")(i_tmp)
			'获取各个父类、祖宗类的记录.
			sqlM2="SELECT TOP 1 * FROM [CXBG_product_class] WHERE isShow=1 AND id="& CokeShow.CokeClng(Split(RSM("ParentPath"),",")(i_tmp)) &" ORDER BY RootID,OrderID"
			Set RSM2 = CONN.Execute(sqlM2)
			'Dim tmpClassnameAddString
			If Not RSM2.Eof Then
			If RSM2("Depth")=0 Then tmpClassnameAddString="分类" Else tmpClassnameAddString=""
			Else
			tmpClassnameAddString=""
			End If
			If Not RSM2.Eof Then ShowString=ShowString & " <span class=""font12000"">-</span> <a href="""& CurrentStrFileName &"&classid="& RSM2("classid") &""" target=""_self"">"& RSM2("classname") & tmpClassnameAddString &"</a>"
			RSM2.Close
			
		Next
		'-----------------Go Begin
		
		'终结化操作.
		RSM.Close
		Set RSM = Nothing
		
		Set RSM2 = Nothing
		ShowNavigation_ForOnlyClass = ShowString & strTmpString
		
	End Function
	
	'获取某个菜品的菜品用途（所属菜系、所属口味、福利用途）.
	'参数：
	'1.CurrentID:当前菜品的id.
	'2.strChar	:菜品用途间的间隔字符.
	Public Function ShowProductUSE(CurrentID,strChar)
		ShowProductUSE=False
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		Dim theCount,theCount_Limit
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		theCount			=1
		theCount_Limit		=4
		
		'判断有各种效性.
		If isNumeric(CurrentID) Then
			CurrentID=CokeShow.CokeClng(CurrentID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowProductUSE的参数ID不正确，无法获取当前导航的操作！"
			Exit Function
		End If
		'If isNull(language) Or isEmpty(language) Or language="" Then Exit Function
		
		'-----------------Go Begin
		'检索当前菜品.
		sqlM = "SELECT TOP 1 * FROM [CXBG_product] WHERE deleted=0 AND isOnsale=1 AND id="& CurrentID
		Set RSM = CONN.Execute(sqlM)
		
		'如果连当前菜品都不存在，则输出空字符.
		If RSM.Eof Then
			'输出空字符串.
			ShowProductUSE=""
			Exit Function
		Else
			'如果有，则记录下当前菜品的各个用途.
			'1.列出所属菜系记录.
			'正分类.
			sqlM2="SELECT TOP 1 * FROM [CXBG_product_businessUSE] WHERE 1=1 AND classid="& CokeShow.CokeClng( RSM("product_businessUSE_id") )
			Set RSM2 = CONN.Execute(sqlM2)
			If Not RSM2.Eof Then ShowString=ShowString & strChar &"<a href=""/ChineseDish/ChineseDish.Welcome?product_businessUSE_id="& RSM2("classid") &""" target=""_blank"">"& RSM2("classname") &"</a>" : theCount=theCount+1 : ShowProductUSE = Right( ShowString,Len(ShowString)-1 ) : If theCount>theCount_Limit Then Exit Function
			RSM2.Close
			
			'扩展分类.
			If isNumeric( Split(RSM("product_businessUSE_id_extend"),",") ) Then
'response.Write "number:"& RSM("product_businessUSE_id_extend")
				sqlM2="SELECT TOP 1 * FROM [CXBG_product_businessUSE] WHERE 1=1 AND classid="& CokeShow.CokeClng( RSM("product_businessUSE_id_extend") )
				Set RSM2 = CONN.Execute(sqlM2)
				If Not RSM2.Eof Then ShowString=ShowString & strChar &"<a href=""/ChineseDish/ChineseDish.Welcome?product_businessUSE_id="& RSM2("classid") &""" target=""_blank"">"& RSM2("classname") &"</a>" : theCount=theCount+1 : ShowProductUSE = Right( ShowString,Len(ShowString)-1 ) : If theCount>theCount_Limit Then Exit Function
				RSM2.Close
				
			ElseIf isArray( Split(RSM("product_businessUSE_id_extend"),",") ) Then
'response.Write Ubound( Split(RSM("product_businessUSE_id_extend"),",") )
				For i_tmp=0 To Ubound( Split(RSM("product_businessUSE_id_extend"),",") )
					'Split(RSM("product_businessUSE_id_extend"),",")(i_tmp)
					'获取扩展类中的各个分隔值，并分别获取相应分类数据.
					sqlM2="SELECT TOP 1 * FROM [CXBG_product_businessUSE] WHERE 1=1 AND classid="& CokeShow.CokeClng( Split(RSM("product_businessUSE_id_extend"),",")(i_tmp) )
					Set RSM2 = CONN.Execute(sqlM2)
					If Not RSM2.Eof Then ShowString=ShowString & strChar &"<a href=""/ChineseDish/ChineseDish.Welcome?product_businessUSE_id="& RSM2("classid") &""" target=""_blank"">"& RSM2("classname") &"</a>" : theCount=theCount+1 : ShowProductUSE = Right( ShowString,Len(ShowString)-1 ) : If theCount>theCount_Limit Then Exit Function
					RSM2.Close
					
				Next
			Else
				'表示啥也没有！
				'response.Write "wl"
			End If
			
			'2.列出所属口味记录.
			'正分类.
			sqlM2="SELECT TOP 1 * FROM [CXBG_product_activityUSE] WHERE 1=1 AND classid="& CokeShow.CokeClng( RSM("product_activityUSE_id") )
			Set RSM2 = CONN.Execute(sqlM2)
			If Not RSM2.Eof Then ShowString=ShowString & strChar &"<a href=""/ChineseDish/ChineseDish.Welcome?product_activityUSE_id="& RSM2("classid") &""" target=""_blank"">"& RSM2("classname") &"</a>" : theCount=theCount+1 : ShowProductUSE = Right( ShowString,Len(ShowString)-1 ) : If theCount>theCount_Limit Then Exit Function
			RSM2.Close
			
			'扩展分类.
			If isNumeric( Split(RSM("product_activityUSE_id_extend")) ) Then
				sqlM2="SELECT TOP 1 * FROM [CXBG_product_activityUSE] WHERE 1=1 AND classid="& CokeShow.CokeClng( RSM("product_activityUSE_id_extend") )
				Set RSM2 = CONN.Execute(sqlM2)
				If Not RSM2.Eof Then ShowString=ShowString & strChar &"<a href=""/ChineseDish/ChineseDish.Welcome?product_activityUSE_id="& RSM2("classid") &""" target=""_blank"">"& RSM2("classname") &"</a>" : theCount=theCount+1 : ShowProductUSE = Right( ShowString,Len(ShowString)-1 ) : If theCount>theCount_Limit Then Exit Function
				RSM2.Close
				
			ElseIf isArray( Split(RSM("product_activityUSE_id_extend")) ) Then
				For i_tmp=0 To Ubound( Split(RSM("product_activityUSE_id_extend"),",") )
					'Split(RSM("product_activityUSE_id_extend"),",")(i_tmp)
					'获取扩展类中的各个分隔值，并分别获取相应分类数据.
					sqlM2="SELECT TOP 1 * FROM [CXBG_product_activityUSE] WHERE 1=1 AND classid="& CokeShow.CokeClng( Split(RSM("product_activityUSE_id_extend"),",")(i_tmp) )
					Set RSM2 = CONN.Execute(sqlM2)
					If Not RSM2.Eof Then ShowString=ShowString & strChar &"<a href=""/ChineseDish/ChineseDish.Welcome?product_activityUSE_id="& RSM2("classid") &""" target=""_blank"">"& RSM2("classname") &"</a>" : theCount=theCount+1 : ShowProductUSE = Right( ShowString,Len(ShowString)-1 ) : If theCount>theCount_Limit Then Exit Function
					RSM2.Close
					
				Next
			Else
				'表示啥也没有！
			End If
			
			'3.列出福利用途记录.
			'正分类.
			sqlM2="SELECT TOP 1 * FROM [CXBG_product_welfareUSE] WHERE 1=1 AND classid="& CokeShow.CokeClng( RSM("product_welfareUSE_id") )
'response.Write sqlM2
			Set RSM2 = CONN.Execute(sqlM2)
			If Not RSM2.Eof Then ShowString=ShowString & strChar &"<a href=""/ChineseDish/ChineseDish.Welcome?product_welfareUSE_id="& RSM2("classid") &""" target=""_blank"">"& RSM2("classname") &"</a>" : theCount=theCount+1 : ShowProductUSE = Right( ShowString,Len(ShowString)-1 ) : If theCount>theCount_Limit Then Exit Function
			RSM2.Close
			
			'扩展分类.
			If isNumeric( Split(RSM("product_welfareUSE_id_extend")) ) Then
				sqlM2="SELECT TOP 1 * FROM [CXBG_product_welfareUSE] WHERE 1=1 AND classid="& CokeShow.CokeClng( RSM("product_welfareUSE_id_extend") )
				Set RSM2 = CONN.Execute(sqlM2)
				If Not RSM2.Eof Then ShowString=ShowString & strChar &"<a href=""/ChineseDish/ChineseDish.Welcome?product_welfareUSE_id="& RSM2("classid") &""" target=""_blank"">"& RSM2("classname") &"</a>" : theCount=theCount+1 : ShowProductUSE = Right( ShowString,Len(ShowString)-1 ) : If theCount>theCount_Limit Then Exit Function
				RSM2.Close
				
			ElseIf isArray( Split(RSM("product_welfareUSE_id_extend")) ) Then
				For i_tmp=0 To Ubound( Split(RSM("product_welfareUSE_id_extend"),",") )
					'Split(RSM("product_welfareUSE_id_extend"),",")(i_tmp)
					'获取扩展类中的各个分隔值，并分别获取相应分类数据.
					sqlM2="SELECT TOP 1 * FROM [CXBG_product_welfareUSE] WHERE 1=1 AND classid="& CokeShow.CokeClng( Split(RSM("product_welfareUSE_id_extend"),",")(i_tmp) )
					Set RSM2 = CONN.Execute(sqlM2)
					If Not RSM2.Eof Then ShowString=ShowString & strChar &"<a href=""/ChineseDish/ChineseDish.Welcome?product_welfareUSE_id="& RSM2("classid") &""" target=""_blank"">"& RSM2("classname") &"</a>" : theCount=theCount+1 : ShowProductUSE = Right( ShowString,Len(ShowString)-1 ) : If theCount>theCount_Limit Then Exit Function
					RSM2.Close
					
				Next
			Else
				'表示啥也没有！
			End If
			
		End If
		
		
		'-----------------Go Begin
		
		'终结化操作.
		RSM.Close
		Set RSM = Nothing
		
		Set RSM2 = Nothing
		ShowProductUSE = ShowString
		
	End Function
	
	'获取当前菜品的第一张菜品图片.
	'参数：
	'1.CurrentID:当前菜品的id.
	'2.PhotoSize:当前菜品的指定尺寸图片（0代表返回不是缩略图的原大尺寸图，60代表60像素缩略图，120代表120像素缩略图，160代表160像素缩略图）.
	Public Function ShowProductFirstPhoto(CurrentID,PhotoSize)
		ShowProductFirstPhoto=False
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		
		'判断有各种效性.
		If isNumeric(CurrentID) Then
			CurrentID=CokeShow.CokeClng(CurrentID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowProductFirstPhoto的参数ID不正确，无法获取当前导航的操作！"
			Exit Function
		End If
		If isNumeric(PhotoSize) Then
			PhotoSize=CokeShow.CokeClng(PhotoSize)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowProductFirstPhoto的参数Size不正确，无法获取当前导航的操作！"
			Exit Function
		End If
		'If isNull(language) Or isEmpty(language) Or language="" Then Exit Function
		
		'-----------------Go Begin
		'列出导航记录
		sqlM = "SELECT TOP 1 * "&_
				" FROM [CXBG_product] "&_
					
				" WHERE deleted=0 AND isOnsale=1 AND id="& CurrentID &" " &_
					
				" ORDER BY id DESC"
		Set RSM = CONN.Execute(sqlM)
		
		'如果连当前分类都不存在，则输出空字符.
		If RSM.Eof Then
			'输出默认无图片.
			ShowProductFirstPhoto="/images/NoPic.gif"
			Exit Function
		Else
			'如果有菜品记录，那么再获取第一个排序最高的图片记录.
				'获取各个父类、祖宗类的记录.
				sqlM2="SELECT TOP 1 * FROM [CXBG_product__photos] WHERE product_id="& RSM("id") &"ORDER BY photos_orderid DESC,id ASC"
				Set RSM2 = CONN.Execute(sqlM2)
				'如果有图片记录.
				If Not RSM2.Eof Then
					If PhotoSize=0 Then
						ShowString = RSM2("photos_src")
					ElseIf PhotoSize>0 Then
						ShowString = Replace(RSM2("photos_src"),"/uploadimages/","/uploadimages/"& PhotoSize &"/")
					End If
				Else
					'如果没有图片记录.
					'输出默认无图片.
					ShowProductFirstPhoto="/images/NoPic.gif"
					Exit Function
				End If
				RSM2.Close
				Set RSM2 = Nothing
		End If
		
		'-----------------Go Begin
		
		'终结化操作.
		RSM.Close
		Set RSM = Nothing
		
		
		ShowProductFirstPhoto = ShowString
		
	End Function
	
	
	'获取正分类及扩展分类的SQL片段.WL
	'获取当前菜品的正分类及扩展分类的SQL片段语句.是用于查询product表菜品记录时的SQL片段.主要用于列出有某个分类classid的菜品（包括扩展分类中有此classid）的面向对象专业函数。
	'参数：
	'1.CurrentClassID:要查询的菜品分类classid.
	Public Function strSQL_ProductClassALL(CurrentClassID)
		strSQL_ProductClassALL=False
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		
		Dim SQLproduct_class,SQLproduct_class_extend
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		
		'判断有各种效性.
		If isNumeric(CurrentClassID) Then
			CurrentClassID=CokeShow.CokeClng(CurrentClassID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法strSQL_ProductClassALL的参数ID不正确，无法获取当前导航的操作！"
			Exit Function
		End If
		'If isNull(language) Or isEmpty(language) Or language="" Then Exit Function
		
		'-----------------Go Begin
		'列出导航记录
	
		'当前分类的classid.
		If CurrentClassID="" Then
			CurrentClassID=0
			SQLproduct_class		=""
			SQLproduct_class_extend	=""
			ShowString				=""
		Else
			If isNumeric(CurrentClassID) Then
				CurrentClassID=CokeShow.CokeClng(CurrentClassID)
				'构造正分类的sql语句.
				SQLproduct_class		=" (   product_class_id="& CurrentClassID &" OR product_class_id IN (  SELECT id From [CXBG_product_class] WHERE isShow=1 AND RootID=(SELECT RootID FROM [CXBG_product_class] WHERE isShow=1 AND classid="& CurrentClassID &") AND Depth>(SELECT Depth FROM [CXBG_product_class] WHERE isShow=1 AND classid="& CurrentClassID &") AND ParentPath LIKE '"& CokeShow.otherField("[CXBG_product_class]",CurrentClassID,"classid","ParentPath",True,100) &","& CokeShow.otherField("[CXBG_product_class]",CurrentClassID,"classid","id",True,100) &"%'  )   ) "
'Response.Write "SELECT TOP 500 * FROM "& CurrentTableName &" WHERE deleted=0 AND isOnsale=1 "& SQLproduct_class & strSQL_brandAll & strSQL_PriceAreaAll & strSQL_businessUSEAll & strSQL_activityUSEAll & strSQL_welfareUSEAll &" ORDER BY id DESC"
				
				'构造扩展分类的sql语句.
				Dim rsTmp,sqlTmp,i_extend
				Dim strIDs
				'获取当前classid下的 子分类记录集(之classid集合rs(0)).
				sqlTmp="SELECT classid,ParentPath FROM [CXBG_product_class] WHERE RootID="& CokeShow.otherField("[CXBG_product_class]",CurrentClassID,"classid","RootID",True,0) &" AND Depth>"& CokeShow.otherField("[CXBG_product_class]",CurrentClassID,"classid","Depth",True,0) &" AND ParentPath LIKE '"& CokeShow.otherField("[CXBG_product_class]",CurrentClassID,"classid","ParentPath",True,100) &","& CokeShow.otherField("[CXBG_product_class]",CurrentClassID,"classid","id",True,100) &"%'"
'response.Write sqlTmp
				Set rsTmp=CONN.Execute( sqlTmp )
				If Not rsTmp.Eof Then
					'如果有子分类，那么查出所有子分类的classid，并加工成8位数字，然后去挨个对照扩展字段的值！(生成其挨个对照的sql语句)
					SQLproduct_class_extend	=" ( product_class_id_extend LIKE '%"& CokeShow.AdditionZero( CurrentClassID, 8 ) &"%' "
					'判断当前的所有子分类中，是否直系关系的依据. WL
					strIDs=","& CokeShow.otherField("[CXBG_product_class]",CurrentClassID,"classid","id",True,0) &","
'response.Write strIDs
					'数组列出查询到的classid集合.
					Do While Not rsTmp.Eof
						'判断当前的所有子分类中，是否直系关系！如果是直系那么构造sql，如果不是直系那么无动作. WL
						If Instr(","& rsTmp(1) &",", strIDs)>0 Then
							SQLproduct_class_extend=SQLproduct_class_extend &" OR product_class_id_extend LIKE '%"& CokeShow.AdditionZero( rsTmp(0), 8 ) &"%' "
						End If
						
						rsTmp.MoveNext
					Loop
					'链接sql字符串结束.
					SQLproduct_class_extend=SQLproduct_class_extend &" ) "
					
'response.Write SQLproduct_class_extend
				Else
					'如果没有子分类，就只查询匹配是否有当前分类即可.
					SQLproduct_class_extend	=" ( product_class_id_extend LIKE '%"& CokeShow.AdditionZero( CurrentClassID, 8 ) &"%' ) "
					
				End If
				rsTmp.Close
				Set rsTmp=Nothing
				
				'组合.
				ShowString = " ("& SQLproduct_class &" OR "& SQLproduct_class_extend &") "
				
			Else
				CurrentClassID=0
				SQLproduct_class		=""
				SQLproduct_class_extend	=""
				ShowString				=""
			End If
		End If
		'-----------------Go End
		
		
		strSQL_ProductClassALL = ShowString
		
	End Function
	
	
	
	'前台专用分页函数.（先用在菜品列表页）
	'ShowTotal是否显示统计信息.ShowAllPages是否显示下拉页码.
    Public Function ShowPage(sfilename, totalnumber, maxperpage, ShowTotal, ShowAllPages, strUnit)
        Dim n, i, strUrl
		
		Dim currentPage
		currentPage		=CokeShow.filtRequest(Request("Page"))
		'处理当前页码的控制变量，通过获取到的传值获取，默认为第一页1.
		If currentPage<>"" Then
			If isNumeric(currentPage) Then currentPage=CokeShow.CokeClng(currentPage) Else currentPage=1
		Else
			currentPage=1
		End If
		
        '计算出页数n.
		If totalnumber Mod maxperpage = 0 Then
            n = totalnumber \ maxperpage
        Else
            n = totalnumber \ maxperpage + 1
        End If
		
'<span>共<span class="font_yl">212</span>个记录 </span><span class="disabled">&lt; </span><span class="current">1</span><a href="#?page=2">2</a><a href="#?page=3">3</a>...<a href="#?page=199">199</a><a href="#?page=200">200</a><a class="nextym" href=""> 下一页更精彩 </a>		
		
		'往文件后边家参数前的追加符号准备.
        strUrl = JoinChar(sfilename)
		
		'是否允许显示 统计信息.
        If ShowTotal = True Then
            ShowPage = ShowPage & "<span>共<span class=""font_yl"">"& totalnumber &"</span>"& strUnit &" </span>"
        End If
		
        If CurrentPage < 2 Then		'特殊处理第一页时.
			'//ShowPage = ShowPage & "first&nbsp;&nbsp;&nbsp;&lt;Prev&nbsp;&nbsp;&nbsp;"
			ShowPage = ShowPage & "<span class=""disabled"">&lt; </span>"
        Else
			'//ShowPage = ShowPage & "<a href='"& strUrl &"page=1'>first</a>&nbsp;&nbsp;&nbsp;"
			'//ShowPage = ShowPage & "<a href='"& strUrl &"page="& (CurrentPage - 1) &"'>&lt;Prev</a>&nbsp;&nbsp;&nbsp;"
			ShowPage = ShowPage & "<a href='"& strUrl &"page="& (CurrentPage - 1) &"'><span class=""disabled"">&lt; </span></a>"
        End If
		
		'显示页码
		For i=1 To n
			If CokeShow.CokeClng(CurrentPage) = CokeShow.CokeClng(i) Then
				ShowPage=ShowPage &"<span class=""current"">"& i &"</span>"
			Else
				ShowPage=ShowPage &"<a href='"& strUrl &"page="& i &"'>"& i &"</a>"
			End If
		Next
    	
        If n - CurrentPage < 1 Then	'特殊处理最后一页时.
			'//ShowPage = ShowPage & "Next&gt;&nbsp;&nbsp;&nbsp;last"
			ShowPage = ShowPage & "<a class=""nextym"" href=""#"" onclick=""return false;""> 	下一页"& strUnit &"更精彩 </a>"
        Else
			'//ShowPage = ShowPage & "<a href='"& strUrl &"page="& (CurrentPage + 1) &"'>Next&gt;</a>&nbsp;&nbsp;&nbsp;"
			'//ShowPage = ShowPage & "<a href='"& strUrl &"page="& n &"'>last</a>"
			ShowPage = ShowPage & "<a href='"& strUrl &"page="& (CurrentPage + 1) &"' class=""nextym""> 	下一页"& strUnit &"更精彩 </a>"
        End If
		
        '//ShowPage = ShowPage & "&nbsp;&nbsp;&nbsp;<b>Now:&nbsp;"& CurrentPage &"</b>/"& n &"&nbsp;</strong><b>Page</b> "
       'ShowPage = ShowPage & "&nbsp;"& maxperpage &""& strUnit &"/页"
		
		'是否显示下拉页码.
'//        If ShowAllPages = True Then
'//            ShowPage = ShowPage &"&nbsp;&nbsp;&nbsp;Go <select name='Page' size='1' onchange=""javascript:window.location='"& strUrl &"Page=" & "' + this.options[this.selectedIndex].value;"" >"
'//            For i = 1 To n
'//                ShowPage = ShowPage &"<option value='"& i &"'"
'//                If CokeShow.CokeClng(CurrentPage) = CokeShow.CokeClng(i) Then ShowPage = ShowPage &" selected "
'//                ShowPage = ShowPage & ">"& i &"</option>"
'//            Next
'//            ShowPage = ShowPage & "</select>"
'//        End If
		
		
		
    End Function
	
	Public Function JoinChar(strUrl)
        If strUrl = "" Then
            JoinChar = ""
            Exit Function
        End If
		
		'如果?不在最后一个出现的话.
        If InStr(strUrl, "?") < Len(strUrl) Then
           '如果存在?，只是不在最后一个出现的话，则处理&符号.
			If InStr(strUrl, "?") > 1 Then
               '如果&不在最后一个出现的话,追加&在尾部.
				If InStr(strUrl, "&") < Len(strUrl) Then
                    JoinChar = strUrl & "&"
                Else
				'否则，证明已经有&在尾部了.
                    JoinChar = strUrl
                End If
            Else
			'如果不存在?，则直接加上.
                JoinChar = strUrl & "?"
            End If
        Else
		'如果?在最后一个出现，那么直接保留url.
            JoinChar = strUrl
        End If
    End Function
	
	
	'获取当前菜品的总体星级数字.
	'参数：
	'1.CurrentID:当前菜品的id.
	Public Function ShowProductStarRating_Num(CurrentID)
		ShowProductStarRating_Num=0
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		Dim countRSM,numRSM
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		countRSM			=0
		numRSM				=1
		
		'判断有各种效性.
		If isNumeric(CurrentID) Then
			CurrentID=CokeShow.CokeClng(CurrentID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowProductStarRating_Num的参数ID不正确，无法获取当前导航的操作！"
			Exit Function
		End If
		
		'-----------------Go Begin
		'列出导航记录
		sqlM = "SELECT TOP 100 Sum(theStarRatingForChineseDishInformation) AS Sum,Count(theStarRatingForChineseDishInformation) AS Count "&_
				" FROM [CXBG_account_RemarkOn] "&_
				" WHERE deleted=0 AND theStarRatingForChineseDishInformation>0 AND product_id="& CurrentID
		Set RSM = CONN.Execute(sqlM)
		Set RSM	=Server.CreateObject("Adodb.RecordSet")
		RSM.Open sqlM,CONN,1,1
		countRSM=RSM.RecordCount
		numRSM	=1
		
		'
		If RSM("Count")=0 Then		'因为聚合记录集不存在RS.Eof情况.
			'输出默认无星级.
			ShowProductStarRating_Num=0
			RSM.Close
			Set RSM = Nothing
			Exit Function
		Else
			'输出默认无星级.
			If ( RSM(0)/RSM(1) )>=5 Then
				ShowProductStarRating_Num=5
			Else
				ShowProductStarRating_Num=( RSM("Sum")/RSM("Count") )
			End If
			RSM.Close
			Set RSM = Nothing
			Exit Function
		End If
		'-----------------Go Begin
		
		'终结化操作.
		'RSM.Close
		'Set RSM = Nothing
		
		'ShowProductStarRating_Num = ShowString
		
	End Function
	
	'依据所获取的当前菜品的总体星级数字，而输出餐厅自我评语的文字.
	'参数：
	'1.CurrentID:当前菜品的总体星级数字.
	Public Function ShowProductStarRating_Str(CurrentID)
		ShowProductStarRating_Str=""
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		Dim countRSM,numRSM
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		countRSM			=0
		numRSM				=1
		
		'判断有各种效性.
		If isNumeric(CurrentID) Then
			CurrentID=CurrentID
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowProductStarRating_Str的参数ID不正确，无法获取当前星级自我评论的操作！"
			Exit Function
		End If
		
		'-----------------Go Begin
		'由上而下的优先评语输出设定.
		If CurrentID=0 Then
			'输出无星级时的餐厅自我评语.
			ShowProductStarRating_Str="餐厅全体员工期待您的点评哦！"
			Exit Function
		ElseIf CurrentID<1 Then
			'输出小于1星级时的餐厅自我评语.
			ShowProductStarRating_Str="我们非常羞愧，3天内我们定会联合厨师长一起研究开会！"
			Exit Function
		ElseIf CurrentID<2 Then
			'输出小于2星级时的餐厅自我评语.
			ShowProductStarRating_Str="收到大家低评我们很羞愧，两天后我们便会得知并着手研究！"
			Exit Function
		ElseIf CurrentID<3 Then
			'输出小于3星级时的餐厅自我评语.
			ShowProductStarRating_Str="小于3星哦，我们菜品服务一定有待改善。"
			Exit Function
		ElseIf CurrentID<4 Then
			'输出小于4星级时的餐厅自我评语.
			ShowProductStarRating_Str="感谢顾客的普遍佳评，我们一定会继续努力为大家服务！"
			Exit Function
		ElseIf CurrentID<5 Then
			'输出小于5星级时的餐厅自我评语.
			ShowProductStarRating_Str="哇谢谢顾客的近5星美评，我们定会告诉厨师长！再接再厉~"
			Exit Function
		ElseIf CurrentID>=5 Then
			'输出等于大于5星级时的餐厅自我评语.
			ShowProductStarRating_Str="为您服务就是我们餐厅每位员工的至高荣幸！5星级哦~"
			Exit Function
		End If
		'-----------------Go Begin
		
		'终结化操作.
		'RSM.Close
		'Set RSM = Nothing
		
		'ShowProductStarRating_Str = ShowString
		
	End Function
	
	
	'获取当前菜品的总体口味分数数字.(平均分)
	'参数：
	'1.CurrentID:当前菜品的id.
	Public Function ShowProductChineseDish_Taste_Num(CurrentID)
		ShowProductChineseDish_Taste_Num=0
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		Dim countRSM,numRSM
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		countRSM			=0
		numRSM				=1
		
		'判断有各种效性.
		If isNumeric(CurrentID) Then
			CurrentID=CokeShow.CokeClng(CurrentID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowProductChineseDish_Taste_Num的参数ID不正确，无法获取当前导航的操作！"
			Exit Function
		End If
		
		'-----------------Go Begin
		'列出导航记录
		sqlM = "SELECT TOP 100 Sum(ChineseDish_Taste) AS Sum,Count(ChineseDish_Taste) AS Count "&_
				" FROM [CXBG_account_RemarkOn] "&_
				" WHERE deleted=0 AND ChineseDish_Taste>0 AND product_id="& CurrentID
		Set RSM = CONN.Execute(sqlM)
		Set RSM	=Server.CreateObject("Adodb.RecordSet")
		RSM.Open sqlM,CONN,1,1
		countRSM=RSM.RecordCount
		numRSM	=1
		
		'
		If RSM("Count")=0 Then		'因为聚合记录集不存在RS.Eof情况.
			'输出默认无口味分数.
			ShowProductChineseDish_Taste_Num=0
			RSM.Close
			Set RSM = Nothing
			Exit Function
		Else
			'输出默认无星级.
			If ( RSM(0)/RSM(1) )>=5 Then
				ShowProductChineseDish_Taste_Num=5
			Else
				ShowProductChineseDish_Taste_Num=FormatNumber( RSM("Sum")/RSM("Count"), 1)
			End If
			RSM.Close
			Set RSM = Nothing
			Exit Function
		End If
		'-----------------Go Begin
		
		'终结化操作.
		'RSM.Close
		'Set RSM = Nothing
		
		'ShowProductChineseDish_Taste_Num = ShowString
		
	End Function
	
	'获取当前菜品的总体环境分数数字.(平均分)
	'参数：
	'1.CurrentID:当前菜品的id.
	Public Function ShowProductChineseDish_DiningArea_Num(CurrentID)
		ShowProductChineseDish_DiningArea_Num=0
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		Dim countRSM,numRSM
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		countRSM			=0
		numRSM				=1
		
		'判断有各种效性.
		If isNumeric(CurrentID) Then
			CurrentID=CokeShow.CokeClng(CurrentID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowProductChineseDish_DiningArea_Num的参数ID不正确，无法获取当前导航的操作！"
			Exit Function
		End If
		
		'-----------------Go Begin
		'列出导航记录
		sqlM = "SELECT TOP 100 Sum(ChineseDish_DiningArea) AS Sum,Count(ChineseDish_DiningArea) AS Count "&_
				" FROM [CXBG_account_RemarkOn] "&_
				" WHERE deleted=0 AND ChineseDish_DiningArea>0 AND product_id="& CurrentID
		Set RSM = CONN.Execute(sqlM)
		Set RSM	=Server.CreateObject("Adodb.RecordSet")
		RSM.Open sqlM,CONN,1,1
		countRSM=RSM.RecordCount
		numRSM	=1
		
		'
		If RSM("Count")=0 Then		'因为聚合记录集不存在RS.Eof情况.
			'输出默认无口味分数.
			ShowProductChineseDish_DiningArea_Num=0
			RSM.Close
			Set RSM = Nothing
			Exit Function
		Else
			'输出默认无星级.
			If ( RSM(0)/RSM(1) )>=5 Then
				ShowProductChineseDish_DiningArea_Num=5
			Else
				ShowProductChineseDish_DiningArea_Num=FormatNumber( RSM("Sum")/RSM("Count"), 1)
			End If
			RSM.Close
			Set RSM = Nothing
			Exit Function
		End If
		'-----------------Go Begin
		
		'终结化操作.
		'RSM.Close
		'Set RSM = Nothing
		
		'ShowProductChineseDish_DiningArea_Num = ShowString
		
	End Function
	
	'获取当前菜品的总体服务分数数字.(平均分)
	'参数：
	'1.CurrentID:当前菜品的id.
	Public Function ShowProductChineseDish_Service_Num(CurrentID)
		ShowProductChineseDish_Service_Num=0
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		Dim countRSM,numRSM
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		countRSM			=0
		numRSM				=1
		
		'判断有各种效性.
		If isNumeric(CurrentID) Then
			CurrentID=CokeShow.CokeClng(CurrentID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowProductChineseDish_Service_Num的参数ID不正确，无法获取当前导航的操作！"
			Exit Function
		End If
		
		'-----------------Go Begin
		'列出导航记录
		sqlM = "SELECT TOP 100 Sum(ChineseDish_Service) AS Sum,Count(ChineseDish_Service) AS Count "&_
				" FROM [CXBG_account_RemarkOn] "&_
				" WHERE deleted=0 AND ChineseDish_Service>0 AND product_id="& CurrentID
		Set RSM = CONN.Execute(sqlM)
		Set RSM	=Server.CreateObject("Adodb.RecordSet")
		RSM.Open sqlM,CONN,1,1
		countRSM=RSM.RecordCount
		numRSM	=1
		
		'
		If RSM("Count")=0 Then		'因为聚合记录集不存在RS.Eof情况.
			'输出默认无口味分数.
			ShowProductChineseDish_Service_Num=0
			RSM.Close
			Set RSM = Nothing
			Exit Function
		Else
			'输出默认无星级.
			If ( RSM(0)/RSM(1) )>=5 Then
				ShowProductChineseDish_Service_Num=5
			Else
				ShowProductChineseDish_Service_Num=FormatNumber( RSM("Sum")/RSM("Count"), 1)
			End If
			RSM.Close
			Set RSM = Nothing
			Exit Function
		End If
		'-----------------Go Begin
		
		'终结化操作.
		'RSM.Close
		'Set RSM = Nothing
		
		'ShowProductChineseDish_Service_Num = ShowString
		
	End Function
	
	'获取当前菜品的总体服务分数数字.(平均分)
	'参数：
	'1.CurrentID:当前菜品的id.
	Public Function ShowProductChineseDish_ConsumePerPerson_Num(CurrentID)
		ShowProductChineseDish_ConsumePerPerson_Num=0
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		Dim countRSM,numRSM
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		countRSM			=0
		numRSM				=1
		
		'判断有各种效性.
		If isNumeric(CurrentID) Then
			CurrentID=CokeShow.CokeClng(CurrentID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowProductChineseDish_ConsumePerPerson_Num的参数ID不正确，无法获取当前导航的操作！"
			Exit Function
		End If
		
		'-----------------Go Begin
		'列出导航记录
		sqlM = "SELECT TOP 100 Sum(ChineseDish_ConsumePerPerson) AS Sum,Count(ChineseDish_ConsumePerPerson) AS Count "&_
				" FROM [CXBG_account_RemarkOn] "&_
				" WHERE deleted=0 AND ChineseDish_ConsumePerPerson>0 AND product_id="& CurrentID
		Set RSM = CONN.Execute(sqlM)
		Set RSM	=Server.CreateObject("Adodb.RecordSet")
		RSM.Open sqlM,CONN,1,1
		countRSM=RSM.RecordCount
		numRSM	=1
		
		'
		If RSM("Count")=0 Then		'因为聚合记录集不存在RS.Eof情况.
			'输出默认无口味分数.
			ShowProductChineseDish_ConsumePerPerson_Num=0
			RSM.Close
			Set RSM = Nothing
			Exit Function
		Else
			'输出默认无星级.
			'If ( RSM(0)/RSM(1) )>=5 Then
			'	ShowProductChineseDish_ConsumePerPerson_Num=5
			'Else
				ShowProductChineseDish_ConsumePerPerson_Num=FormatCurrency( RSM("Sum")/RSM("Count"), 2)
			'End If
			RSM.Close
			Set RSM = Nothing
			Exit Function
		End If
		'-----------------Go Begin
		
		'终结化操作.
		'RSM.Close
		'Set RSM = Nothing
		
		'ShowProductChineseDish_ConsumePerPerson_Num = ShowString
		
	End Function
	
	
	'获取当前会员的性别对应头像的scr地址.
	'参数：
	'1.CurrentID:当前会员的id.
	Public Function ShowMemberSexPicURL(CurrentID)
		ShowMemberSexPicURL="/images/NoPic.png"
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		Dim countRSM,numRSM
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		countRSM			=0
		numRSM				=1
		
		'判断有各种效性.
		If isNumeric(CurrentID) Then
			CurrentID=CokeShow.CokeClng(CurrentID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowMemberSexPicURL的参数ID不正确，无法获取当前会员的性别对应头像的scr地址的操作！"
			Exit Function
		End If
		
		'-----------------Go Begin
		'列出导航记录
		sqlM = "SELECT TOP 1 Sex,isBindingVIPCardNumber "&_
				" FROM [CXBG_account] "&_
				" WHERE deleted=0 AND id="& CurrentID
		Set RSM = CONN.Execute(sqlM)
		Set RSM	=Server.CreateObject("Adodb.RecordSet")
		RSM.Open sqlM,CONN,1,1
		countRSM=RSM.RecordCount
		numRSM	=1
		
		'
		If RSM.Eof Then		'因为聚合记录集不存在RS.Eof情况.
			'输出无此会员记录时的情形.
			ShowMemberSexPicURL="/images/NoPic.png"
			RSM.Close
			Set RSM = Nothing
			Exit Function
		Else
			If RSM("isBindingVIPCardNumber")=0 Then
				'尚未绑定会员卡的头像系列.
				If RSM("Sex")=0 Then
					ShowMemberSexPicURL="/images/hytx/secrecy_100.jpg"	'输出默认未选择男士或者女士的未知性别头像——保密.
				ElseIf RSM("Sex")=1 Then
					ShowMemberSexPicURL="/images/hytx/girl_100.jpg"
				ElseIf RSM("Sex")=2 Then
					ShowMemberSexPicURL="/images/hytx/boy_100.jpg"
				End If
			ElseIf RSM("isBindingVIPCardNumber")=1 Then
				'已绑定会员卡的头像系列.
				If RSM("Sex")=0 Then
					ShowMemberSexPicURL="/images/hytx/secrecy_100card.jpg"	'输出默认未选择男士或者女士的未知性别头像——保密.
				ElseIf RSM("Sex")=1 Then
					ShowMemberSexPicURL="/images/hytx/girl_100card.jpg"
				ElseIf RSM("Sex")=2 Then
					ShowMemberSexPicURL="/images/hytx/boy_100card.jpg"
				End If
			End If
			RSM.Close
			Set RSM = Nothing
			Exit Function
		End If
		'-----------------Go Begin
		
		'终结化操作.
		'RSM.Close
		'Set RSM = Nothing
		
		'ShowMemberSexPicURL = ShowString
		
	End Function
	
	'判断当前会员的留言是否有新回复未读的.
	'参数：
	'1.CurrentID:当前会员的id.
	Public Function ShowReplyStatus(CurrentID)
		ShowReplyStatus=False
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		Dim countRSM,numRSM
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		countRSM			=0
		numRSM				=1
		
		'判断有各种效性.
		If CurrentID="" Then
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowReplyStatus的参数CurrentID不正确，无法获取当前积分处理的操作！"
			Exit Function
		Else
			If CokeShow.strLength(CurrentID)>50 Or CokeShow.strLength(CurrentID)<10 Then
				'参数不正确，退出.操作失败.
				'Response.Clear()
				Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowReplyStatus的参数CurrentID长度不正确，无法获取当前积分处理的操作！"
				Exit Function
			Else
'				If CokeShow.IsValidEmail(CurrentID)=False Then
'					'参数不正确，退出.操作失败.
'					'Response.Clear()
'					Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowReplyStatus的参数CurrentID格式不正确，无法获取当前积分处理的操作！"
'					Exit Function
'				Else
'					CurrentID=CurrentID
'				End If
				CurrentID=CurrentID
			End If
		End If
		
		'-----------------Go Begin
		'列出导航记录
		sqlM = "SELECT TOP 1 * "&_
				" FROM [CXBG_account_Message] "&_
				" WHERE deleted=0 AND isRead=0 AND toWho='"& CurrentID &"'"
'Response.Write sqlM
		Set RSM = CONN.Execute(sqlM)
		Set RSM	=Server.CreateObject("Adodb.RecordSet")
		RSM.Open sqlM,CONN,1,1
		countRSM=RSM.RecordCount
		numRSM	=1
		
		'
		If RSM.Eof Then		'没有新回复信息.
			ShowReplyStatus=False
			RSM.Close
			Set RSM = Nothing
			Exit Function
		Else				'有新回复！
			ShowReplyStatus=True
			RSM.Close
			Set RSM = Nothing
			Exit Function
		End If
		'-----------------Go End
		
		'终结化操作.
		'RSM.Close
		'Set RSM = Nothing
		
		'ShowReplyStatus = ShowString
		
	End Function
	
	'判断当前会员要绑定的卡号，是否已经有其他人绑定过了.
	'参数：
	'1.CurrentID:当前会员输入的想要绑定的VIP卡的卡号.
	'True表示一切顺利，无人绑定过此卡号。
	Public Function ShowThisVIPCardNumberBindingStatus(CurrentID)
		ShowThisVIPCardNumberBindingStatus=False
		
		'定义内部新变量进行内部操作.
		Dim ShowString
		Dim sqlM,RSM
		Dim sqlM2,RSM2
		Dim i_tmp
		Dim strTmpString
		Dim countRSM,numRSM
		
		'初始化赋值.
		ShowString			=""
		i_tmp				=1
		strTmpString		=""
		countRSM			=0
		numRSM				=1
		
		'判断有各种效性.
		If CurrentID="" Then
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowThisVIPCardNumberBindingStatus的参数CurrentID不正确，无法判断VIP卡号当前的其他人绑定状态的操作！"
			Exit Function
		Else
			If CokeShow.strLength(CurrentID)>20 Or CokeShow.strLength(CurrentID)<4 Then
				'参数不正确，退出.操作失败.
				'Response.Clear()
				Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowThisVIPCardNumberBindingStatus的参数CurrentID长度不正确，无法判断VIP卡号当前的其他人绑定状态的操作！"
				Exit Function
			Else
'				If CokeShow.IsValidEmail(CurrentID)=False Then
'					'参数不正确，退出.操作失败.
'					'Response.Clear()
'					Err.Raise vbObjectError + 1288, "Class foreground_class", "方法ShowThisVIPCardNumberBindingStatus的参数CurrentID格式不正确，无法判断VIP卡号当前的其他人绑定状态的操作！"
'					Exit Function
'				Else
'					CurrentID=CurrentID
'				End If
				CurrentID=CurrentID
			End If
		End If
		
		'验证有没有此卡号的存在！
		sqlM = "SELECT TOP 1 * "&_
				" FROM [CXBG_VIPcard] "&_
				" WHERE isOnpublic=1 AND 0=0 AND classname='"& CurrentID &"'"
'Response.Write sqlM
		Set RSM = CONN.Execute(sqlM)
		Set RSM	=Server.CreateObject("Adodb.RecordSet")
		RSM.Open sqlM,CONN,1,1
		countRSM=RSM.RecordCount
		numRSM	=1
		
		'
		If RSM.Eof Then		'如果没有如此卡号，那么报错，无法绑定卡号，卡号不存在.
			ShowThisVIPCardNumberBindingStatus=False
			RSM.Close
			Set RSM = Nothing
			Exit Function
		Else				'如果有卡号，则通过！
			ShowThisVIPCardNumberBindingStatus=True
			RSM.Close
			Set RSM = Nothing
			Exit Function
		End If
		
		'-----------------Go Begin
		'列出导航记录
		sqlM = "SELECT TOP 1 * "&_
				" FROM [CXBG_account] "&_
				" WHERE deleted=0 AND isRead=0 AND toWho='"& CurrentID &"'"
'Response.Write sqlM
		Set RSM = CONN.Execute(sqlM)
		Set RSM	=Server.CreateObject("Adodb.RecordSet")
		RSM.Open sqlM,CONN,1,1
		countRSM=RSM.RecordCount
		numRSM	=1
		
		'
		If RSM.Eof Then		'没有新回复信息.
			ShowThisVIPCardNumberBindingStatus=False
			RSM.Close
			Set RSM = Nothing
			Exit Function
		Else				'有新回复！
			ShowThisVIPCardNumberBindingStatus=True
			RSM.Close
			Set RSM = Nothing
			Exit Function
		End If
		'-----------------Go End
		
		'终结化操作.
		'RSM.Close
		'Set RSM = Nothing
		
		'ShowThisVIPCardNumberBindingStatus = ShowString
		
	End Function
'*************************************************************************
End Class
%>