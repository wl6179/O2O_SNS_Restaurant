		  <%
		  '点评列表.
'初始化赋值.
'变量定义.
Dim COKE_maxPerPage					'设置当前模块分页设置.
Dim COKE_CurrentTableName			'设置当前模块所涉及的[表]名.
Dim COKE_CurrentPageNow				'设置当前模块所在页面的文件名.
Dim COKE_UnitName					'此主要列表展示中，所涉及的记录的单位名称.
Dim COKE_totalPut,COKE_totalPages,COKE_currentPage			'分页用的控制变量.
Dim COKE_RS, COKE_sql									'查询列表记录用的变量.
Dim COKE_FoundErr,COKE_ErrMsg							'控制错误流程用的控制变量.
Dim COKE_strFileName								'构建查询字符串用的控制变量.
Dim COKE_ExecuteSearch,COKE_Keyword,COKE_TypeSearch,COKE_Action	'构建查询字符串以及流程控制用的控制变量.
Dim COKE_strGuide								'导航文字.



'接收参数.
COKE_maxPerPage			=18
COKE_CurrentTableName 	="[CXBG_account_RemarkOn]"		'此模块涉及的[表]名.
COKE_CurrentPageNow 	="/ChineseDish/ChineseDishInformation.Welcome?CokeMark="& CokeShow.DecodeURL(Request("CokeMark"))
COKE_UnitName			="点评"
COKE_currentPage		=CokeShow.filtRequest(Request("page"))
COKE_ExecuteSearch	=CokeShow.filtRequest(Request("COKE_ExecuteSearch"))
COKE_Keyword		=CokeShow.filtRequest(Request("COKE_Keyword"))
COKE_TypeSearch		=CokeShow.filtRequest(Request("COKE_TypeSearch"))
COKE_Action			=CokeShow.filtRequest(Request("COKE_Action"))


'处理参数.
'处理查询执行 控制变量.
If COKE_ExecuteSearch="" Then
	COKE_ExecuteSearch=0
Else
	If isNumeric(COKE_ExecuteSearch) Then COKE_ExecuteSearch=CokeShow.CokeClng(COKE_ExecuteSearch) Else COKE_ExecuteSearch=0
End If
'构造查询字符串.
If Instr(CokeShow.GetAllUrlII, "?")>0 Then COKE_strFileName=COKE_CurrentPageNow &"&COKE_ExecuteSearch="& COKE_ExecuteSearch Else COKE_strFileName=COKE_CurrentPageNow &"?COKE_ExecuteSearch="& COKE_ExecuteSearch
If COKE_Keyword<>"" Then
	COKE_strFileName=COKE_strFileName&"&COKE_Keyword="& COKE_Keyword
End If
If COKE_TypeSearch<>"" Then
	COKE_strFileName=COKE_strFileName&"&COKE_TypeSearch="& COKE_TypeSearch
End If


'处理当前页码的控制变量，默认为第一页.
If COKE_currentPage<>"" Then
    If isNumeric(COKE_currentPage) Then COKE_currentPage=CokeShow.CokeClng(COKE_currentPage) Else COKE_currentPage=1
Else
	COKE_currentPage=1
End If


'主体控制部分 Begin
Select Case COKE_ExecuteSearch
	Case 0		
		COKE_sql="SELECT TOP 500 * FROM "& COKE_CurrentTableName &" WHERE deleted=0 AND product_id="& CokeMark &" ORDER BY theStarRatingForChineseDishInformation DESC,id DESC"
		COKE_sql="SELECT TOP 100 * FROM "& COKE_CurrentTableName &" WHERE deleted=0 AND product_id="& CokeMark &" ORDER BY id DESC"
		COKE_strGuide=COKE_strGuide & "所有"& COKE_UnitName
'Response.Write sql
	Case 1
		COKE_sql="SELECT TOP 500 * FROM "& COKE_CurrentTableName &" WHERE deleted=0 AND isOnsale=1 AND is_display_newproduct=1 ORDER BY id DESC"
		COKE_strGuide=COKE_strGuide & "痴心不改餐厅最新上架"& COKE_UnitName
	Case 2
		COKE_sql="SELECT TOP 500 * FROM "& COKE_CurrentTableName &" WHERE deleted=0 AND isOnsale=1 AND isSales=1 ORDER BY isSales_StopDate DESC"
		COKE_strGuide=COKE_strGuide & "本月促销"& COKE_UnitName
	Case 10
		If COKE_Keyword="" Then
			COKE_sql="SELECT TOP 500 * FROM "& COKE_CurrentTableName &" WHERE deleted=0 AND isOnsale=1 ORDER BY id DESC"
			COKE_strGuide=COKE_strGuide & "搜索所有"& COKE_UnitName
		Else
			Select Case COKE_TypeSearch
				Case "Brand"
					If IsNumeric(COKE_Keyword)=False Then
						COKE_FoundErr=True
						COKE_ErrMsg=COKE_ErrMsg &"<br /><li>"& COKE_UnitName &"您必须选择一个品牌！</li>"
					Else
						COKE_sql="select TOP 500 * from "& COKE_CurrentTableName &" where deleted=0 AND isOnsale=1 and product_brand_id="& CokeShow.CokeClng(COKE_Keyword) &" order by id desc"
						COKE_strGuide=COKE_strGuide & COKE_UnitName &"品牌为:<font color=red> " & CokeShow.otherField("[CXBG_product_brand]",CokeShow.CokeClng(COKE_Keyword),"classid","classname",True,0) & " </font>的"& COKE_UnitName
					End If
					
				Case "ProductName"
					COKE_sql="select TOP 500 * from "& COKE_CurrentTableName &" where deleted=0 AND isOnsale=1 and (ProductName like '%"& COKE_Keyword &"%' OR product_keywords like '%"& COKE_Keyword &"%') order by id desc"
					COKE_strGuide=COKE_strGuide & "菜品中含有“ <font color=red>" & COKE_Keyword & "</font> ”的"& COKE_UnitName
'response.Write Keyword					
			End Select
			
		End If
		
		
	Case Else
		COKE_FoundErr=True
		COKE_ErrMsg=COKE_ErrMsg & "<br /><li>错误的参数！</li>"
	
End Select

'拦截错误.
If COKE_FoundErr=True Then
	'Response.Clear()
	Err.Raise vbObjectError + 6666, "列表查询出现异常", "如下异常："& COKE_ErrMsg
	Response.End()
End If

If Not IsObject(CONN) Then link_database
Set COKE_RS=Server.CreateObject("Adodb.RecordSet")
'	Response.Write "<br />"& COKE_sql
'	Response.End 
COKE_RS.Open COKE_sql,CONN,1,1

'主体控制部分 End
%>

<!--列表 Begin-->
<%
'主体需要控制的部分.
If COKE_RS.Eof And COKE_RS.Bof Then
	COKE_strGuide=COKE_strGuide & " &#187; 共找到 <font color=red>0</font> 个"& COKE_UnitName
	Call COKE_showMain
Else
	COKE_totalPut=COKE_RS.RecordCount		'记录总数.
	COKE_strGuide=COKE_strGuide & " &#187; 共找到 <font color=red>" & COKE_totalPut & "</font> 个"& COKE_UnitName
	
	
	'处理页码
	If COKE_currentPage<1 Then
		COKE_currentPage=1
	End If
	'如果传递过来的Page当前页值很大，超过了应有的页数时，进行处理.
	If (COKE_currentPage-1) * COKE_maxPerPage > COKE_totalPut Then
		If (COKE_totalPut Mod COKE_maxPerPage)=0 Then
			'如果整好够页数，赋予当前页最大页.
			COKE_currentPage= COKE_totalPut \ COKE_maxPerPage
		Else
			'如果不整好，最有一页只有零散几条记录（不丰满的多余页），赋予当前页最大页.（不能整除情况下计算）
			COKE_currentPage= COKE_totalPut \ COKE_maxPerPage + 1
		End If

	End If
	If COKE_currentPage=1 Then
		
		Call COKE_showMain
		
	Else
		'如果传递过来的Page当前页值不大，在应有的页数范围之内时，理应(COKE_currentPage-1) * COKE_maxPerPage < COKE_totalPut，此时进行一些处理.
		if (COKE_currentPage-1) * COKE_maxPerPage < COKE_totalPut then
			'指针指到(COKE_currentPage-1)页（前一页）的最后一个记录处.
			COKE_RS.Move  (COKE_currentPage-1) * COKE_maxPerPage
			'COKE_RS.BookMark？
			Dim COKE_bookMark
			COKE_bookMark = COKE_RS.BookMark
			
			Call COKE_showMain
			
		else
		'如果传递过来的Page当前页值很大，超过了应有的页数时.打开第一页.
			COKE_currentPage=1
			
			Call COKE_showMain
			
		end if
	End If
End If
'主体需要控制的部分.

%>
<!--列表 End-->






<%
'菜品列表输出.
'针对性的列表内容部分.
Sub COKE_showMain()
   	Dim COKE_i
    COKE_i=0
%>
		<% '=strClassNameDIV %>
		<%
        If COKE_RS.EOF Then
        %>
            <ul class="tjzjdp_sjym_hypl"><li><img src="/images/ico/small/emotion_unhappy.png" /> 暂时还没有会员点评此菜品哦 ... ... <a href="#RemarkOnArea">立刻参与点评</a></li></ul>
        <%
        End If
        %>
       <%
       If Not COKE_RS.EOF Then
       %>
       <ul class="tjzjdp_sjym_hypl">
       <%
       End If
       %>
        <%
        Dim COKE_rsTmp_account_information,COKE_strTmp_Product__Photos
        COKE_strTmp_Product__Photos=""
        
        Do While Not COKE_RS.EOF
            '获取第一张菜品图片.
            Set COKE_rsTmp_account_information=CONN.Execute("SELECT TOP 1 * FROM [CXBG_account] WHERE username='"& COKE_RS("Account_LoginID") &"'")
            If Not COKE_rsTmp_account_information.Eof Then
                'COKE_strTmp_Product__Photos=Replace(COKE_rsTmp_account_information("photos_src"),"/uploadimages/","/uploadimages/120/")
            Else
                'COKE_strTmp_Product__Photos="/images/NoPic.gif"
				COKE_rsTmp_account_information.Close
				COKE_RS.Close
				Set COKE_RS=Nothing
				Exit Sub
            End If
            
        %>
         <li>
	       <table width="100%" border="0" cellspacing="0" cellpadding="0">
             <tr >
               <td width="11%" rowspan="3" valign="top" <% If COKE_RS("Account_LoginID")=Session("username") Then Response.Write "style=""background-color:#; border:3px ORANGE solid;""" Else Response.Write "" %>>
               <img style="border:1px #CCC solid;" src="<% =Coke.ShowMemberSexPicURL(COKE_rsTmp_account_information("id")) %>" width="45" height="45" />
               
               </td>
               <td width="66%" class="coloreee">
               	<% If COKE_rsTmp_account_information("isBindingVIPCardNumber")=1 Then %><img src="/images/hytx/card_01.gif" width="10" height="7" /><% End If %>
                <a class="fontred_a" href="/Club/MembersInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( COKE_rsTmp_account_information("id") ) %>" target="_blank"><% =COKE_rsTmp_account_information("cnname") %></a>
               	&nbsp;&nbsp;&nbsp;
                <span style="color:#999;">点评日期：<% =COKE_RS("adddate") %></span>
               </td>
               <td width="23%" class="coloreee">
               	<ul class="rating">
                <li class="current-rating" style="width:<% =COKE_RS("theStarRatingForChineseDishInformation") * 20 %>px"></li>
                
                </ul>
               </td>
             </tr>
             <tr>
               <td colspan="2"><span class="fontred">口味：<strong style="font-size:14px; font-family: Georgia, 'Times New Roman', Times, serif;"><% =COKE_RS("ChineseDish_Taste") %></strong> 分 　  </span>环境：<% =COKE_RS("ChineseDish_DiningArea") %>分 　服务：<% =COKE_RS("ChineseDish_Service") %> 　人均消费：<% =FormatCurrency(COKE_RS("ChineseDish_ConsumePerPerson"),2) %></td>
               </tr>
             <tr>
               <td colspan="2"><% =COKE_RS("logtext") %></td>
               </tr>
           </table>
		 </li>
        <%
            COKE_rsTmp_account_information.Close	'临时销毁临时记录集.
			
			COKE_i=COKE_i+1
            If COKE_i >= COKE_maxPerPage Then Exit Do
            COKE_RS.MoveNext
        Loop
        
        '销毁临时对象.
        Set COKE_rsTmp_account_information = Nothing
        %>
	   <%
       If Not COKE_RS.EOF Then
       %>
       </ul>
       <%
       End If
       %>
	
<% If COKE_i>0 Then %>
<!--翻页-->
		<div class="sjfy">
		<%
		Response.Write Coke.ShowPage(COKE_strFileName,COKE_totalPut,COKE_maxPerPage,True,True,COKE_UnitName)
		%>
		</div>
<!--翻页-->
<% End If %>

<%
	COKE_RS.Close
	Set COKE_RS=Nothing
End Sub
%>