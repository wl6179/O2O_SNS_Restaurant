﻿		  <%
		  '帐号管理中心之收藏菜列表.
'初始化赋值.
'变量定义.
Dim MemberInformation_JifenHistory_Code_maxPerPage					'设置当前模块分页设置.
Dim MemberInformation_JifenHistory_Code_CurrentTableName			'设置当前模块所涉及的[表]名.
Dim MemberInformation_JifenHistory_Code_CurrentPageNow				'设置当前模块所在页面的文件名.
Dim MemberInformation_JifenHistory_Code_UnitName					'此主要列表展示中，所涉及的记录的单位名称.
Dim MemberInformation_JifenHistory_Code_totalPut,MemberInformation_JifenHistory_Code_totalPages,MemberInformation_JifenHistory_Code_currentPage			'分页用的控制变量.
Dim MemberInformation_JifenHistory_Code_RS, MemberInformation_JifenHistory_Code_sql									'查询列表记录用的变量.
Dim MemberInformation_JifenHistory_Code_FoundErr,MemberInformation_JifenHistory_Code_ErrMsg							'控制错误流程用的控制变量.
Dim MemberInformation_JifenHistory_Code_strFileName								'构建查询字符串用的控制变量.
Dim MemberInformation_JifenHistory_Code_ExecuteSearch,MemberInformation_JifenHistory_Code_Keyword,MemberInformation_JifenHistory_Code_TypeSearch,MemberInformation_JifenHistory_Code_Action	'构建查询字符串以及流程控制用的控制变量.
Dim MemberInformation_JifenHistory_Code_strGuide								'导航文字.



'接收参数.
MemberInformation_JifenHistory_Code_maxPerPage			=30
MemberInformation_JifenHistory_Code_CurrentTableName 	="[CXBG_account_JifenSystem]"		'此模块涉及的[表]名.
MemberInformation_JifenHistory_Code_CurrentPageNow 	="/ONCEFOREVER/Account_JifenHistory.Welcome?Welcome=1"
MemberInformation_JifenHistory_Code_UnitName			="我的积分历史"
MemberInformation_JifenHistory_Code_currentPage		=CokeShow.filtRequest(Request("page"))
MemberInformation_JifenHistory_Code_ExecuteSearch	=CokeShow.filtRequest(Request("MemberInformation_JifenHistory_Code_ExecuteSearch"))
MemberInformation_JifenHistory_Code_Keyword		=CokeShow.filtRequest(Request("MemberInformation_JifenHistory_Code_Keyword"))
MemberInformation_JifenHistory_Code_TypeSearch		=CokeShow.filtRequest(Request("MemberInformation_JifenHistory_Code_TypeSearch"))
MemberInformation_JifenHistory_Code_Action			=CokeShow.filtRequest(Request("MemberInformation_JifenHistory_Code_Action"))


'处理参数.
'处理查询执行 控制变量.
If MemberInformation_JifenHistory_Code_ExecuteSearch="" Then
	MemberInformation_JifenHistory_Code_ExecuteSearch=0
Else
	If isNumeric(MemberInformation_JifenHistory_Code_ExecuteSearch) Then MemberInformation_JifenHistory_Code_ExecuteSearch=CokeShow.CokeClng(MemberInformation_JifenHistory_Code_ExecuteSearch) Else MemberInformation_JifenHistory_Code_ExecuteSearch=0
End If
'构造查询字符串.
If Instr(MemberInformation_JifenHistory_Code_CurrentPageNow, "?")>0 Then MemberInformation_JifenHistory_Code_strFileName=MemberInformation_JifenHistory_Code_CurrentPageNow &"&MemberInformation_JifenHistory_Code_ExecuteSearch="& MemberInformation_JifenHistory_Code_ExecuteSearch Else MemberInformation_JifenHistory_Code_strFileName=MemberInformation_JifenHistory_Code_CurrentPageNow &"?MemberInformation_JifenHistory_Code_ExecuteSearch="& MemberInformation_JifenHistory_Code_ExecuteSearch
If MemberInformation_JifenHistory_Code_Keyword<>"" Then
	MemberInformation_JifenHistory_Code_strFileName=MemberInformation_JifenHistory_Code_strFileName&"&MemberInformation_JifenHistory_Code_Keyword="& MemberInformation_JifenHistory_Code_Keyword
End If
If MemberInformation_JifenHistory_Code_TypeSearch<>"" Then
	MemberInformation_JifenHistory_Code_strFileName=MemberInformation_JifenHistory_Code_strFileName&"&MemberInformation_JifenHistory_Code_TypeSearch="& MemberInformation_JifenHistory_Code_TypeSearch
End If


'处理当前页码的控制变量，默认为第一页.
If MemberInformation_JifenHistory_Code_currentPage<>"" Then
    If isNumeric(MemberInformation_JifenHistory_Code_currentPage) Then MemberInformation_JifenHistory_Code_currentPage=CokeShow.CokeClng(MemberInformation_JifenHistory_Code_currentPage) Else MemberInformation_JifenHistory_Code_currentPage=1
Else
	MemberInformation_JifenHistory_Code_currentPage=1
End If


'主体控制部分 Begin
Select Case MemberInformation_JifenHistory_Code_ExecuteSearch
	Case 0		
		MemberInformation_JifenHistory_Code_sql="SELECT TOP 500 * FROM "& MemberInformation_JifenHistory_Code_CurrentTableName &" WHERE deleted=0 AND Account_LoginID='"& RS("username") &"' ORDER BY id DESC"
		MemberInformation_JifenHistory_Code_strGuide=MemberInformation_JifenHistory_Code_strGuide & "所有"& MemberInformation_JifenHistory_Code_UnitName
'Response.Write sql
'	Case 1
'		MemberInformation_JifenHistory_Code_sql="SELECT TOP 500 * FROM "& MemberInformation_JifenHistory_Code_CurrentTableName &" WHERE deleted=0 AND isOnsale=1 AND is_display_newproduct=1 ORDER BY id DESC"
'		MemberInformation_JifenHistory_Code_strGuide=MemberInformation_JifenHistory_Code_strGuide & "痴心不改餐厅最新上架"& MemberInformation_JifenHistory_Code_UnitName
'	Case 2
'		MemberInformation_JifenHistory_Code_sql="SELECT TOP 500 * FROM "& MemberInformation_JifenHistory_Code_CurrentTableName &" WHERE deleted=0 AND isOnsale=1 AND isSales=1 ORDER BY isSales_StopDate DESC"
'		MemberInformation_JifenHistory_Code_strGuide=MemberInformation_JifenHistory_Code_strGuide & "本月促销"& MemberInformation_JifenHistory_Code_UnitName
'	Case 10
'		If MemberInformation_JifenHistory_Code_Keyword="" Then
'			MemberInformation_JifenHistory_Code_sql="SELECT TOP 500 * FROM "& MemberInformation_JifenHistory_Code_CurrentTableName &" WHERE deleted=0 AND isOnsale=1 ORDER BY id DESC"
'			MemberInformation_JifenHistory_Code_strGuide=MemberInformation_JifenHistory_Code_strGuide & "搜索所有"& MemberInformation_JifenHistory_Code_UnitName
'		Else
'			Select Case MemberInformation_JifenHistory_Code_TypeSearch
'				Case "Brand"
'					If IsNumeric(MemberInformation_JifenHistory_Code_Keyword)=False Then
'						MemberInformation_JifenHistory_Code_FoundErr=True
'						MemberInformation_JifenHistory_Code_ErrMsg=MemberInformation_JifenHistory_Code_ErrMsg &"<br /><li>"& MemberInformation_JifenHistory_Code_UnitName &"您必须选择一个品牌！</li>"
'					Else
'						MemberInformation_JifenHistory_Code_sql="select TOP 500 * from "& MemberInformation_JifenHistory_Code_CurrentTableName &" where deleted=0 AND isOnsale=1 and product_brand_id="& CokeShow.CokeClng(MemberInformation_JifenHistory_Code_Keyword) &" order by id desc"
'						MemberInformation_JifenHistory_Code_strGuide=MemberInformation_JifenHistory_Code_strGuide & MemberInformation_JifenHistory_Code_UnitName &"品牌为:<font color=red> " & CokeShow.otherField("[CXBG_product_brand]",CokeShow.CokeClng(MemberInformation_JifenHistory_Code_Keyword),"classid","classname",True,0) & " </font>的"& MemberInformation_JifenHistory_Code_UnitName
'					End If
'					
'				Case "ProductName"
'					MemberInformation_JifenHistory_Code_sql="select TOP 500 * from "& MemberInformation_JifenHistory_Code_CurrentTableName &" where deleted=0 AND isOnsale=1 and (ProductName like '%"& MemberInformation_JifenHistory_Code_Keyword &"%' OR product_keywords like '%"& MemberInformation_JifenHistory_Code_Keyword &"%') order by id desc"
'					MemberInformation_JifenHistory_Code_strGuide=MemberInformation_JifenHistory_Code_strGuide & "菜品中含有“ <font color=red>" & MemberInformation_JifenHistory_Code_Keyword & "</font> ”的"& MemberInformation_JifenHistory_Code_UnitName
''response.Write Keyword					
'			End Select
'			
'		End If
		
		
	Case Else
		MemberInformation_JifenHistory_Code_FoundErr=True
		MemberInformation_JifenHistory_Code_ErrMsg=MemberInformation_JifenHistory_Code_ErrMsg & "<br /><li>错误的参数！</li>"
	
End Select

'拦截错误.
If MemberInformation_JifenHistory_Code_FoundErr=True Then
	'Response.Clear()
	Err.Raise vbObjectError + 6666, "列表查询出现异常", "如下异常："& MemberInformation_JifenHistory_Code_ErrMsg
	Response.End()
End If

If Not IsObject(CONN) Then link_database
Set MemberInformation_JifenHistory_Code_RS=Server.CreateObject("Adodb.RecordSet")
'	Response.Write "<br />"& MemberInformation_JifenHistory_Code_sql
'	Response.End 
MemberInformation_JifenHistory_Code_RS.Open MemberInformation_JifenHistory_Code_sql,CONN,1,1

'主体控制部分 End
%>

<!--列表 Begin-->
<%
'主体需要控制的部分.
If MemberInformation_JifenHistory_Code_RS.Eof And MemberInformation_JifenHistory_Code_RS.Bof Then
	MemberInformation_JifenHistory_Code_strGuide=MemberInformation_JifenHistory_Code_strGuide & " &#187; 共找到 <font color=red>0</font> 个"& MemberInformation_JifenHistory_Code_UnitName
	Call MemberInformation_JifenHistory_Code_showMain
Else
	MemberInformation_JifenHistory_Code_totalPut=MemberInformation_JifenHistory_Code_RS.RecordCount		'记录总数.
	MemberInformation_JifenHistory_Code_strGuide=MemberInformation_JifenHistory_Code_strGuide & " &#187; 共找到 <font color=red>" & MemberInformation_JifenHistory_Code_totalPut & "</font> 个"& MemberInformation_JifenHistory_Code_UnitName
	
	
	'处理页码
	If MemberInformation_JifenHistory_Code_currentPage<1 Then
		MemberInformation_JifenHistory_Code_currentPage=1
	End If
	'如果传递过来的Page当前页值很大，超过了应有的页数时，进行处理.
	If (MemberInformation_JifenHistory_Code_currentPage-1) * MemberInformation_JifenHistory_Code_maxPerPage > MemberInformation_JifenHistory_Code_totalPut Then
		If (MemberInformation_JifenHistory_Code_totalPut Mod MemberInformation_JifenHistory_Code_maxPerPage)=0 Then
			'如果整好够页数，赋予当前页最大页.
			MemberInformation_JifenHistory_Code_currentPage= MemberInformation_JifenHistory_Code_totalPut \ MemberInformation_JifenHistory_Code_maxPerPage
		Else
			'如果不整好，最有一页只有零散几条记录（不丰满的多余页），赋予当前页最大页.（不能整除情况下计算）
			MemberInformation_JifenHistory_Code_currentPage= MemberInformation_JifenHistory_Code_totalPut \ MemberInformation_JifenHistory_Code_maxPerPage + 1
		End If

	End If
	If MemberInformation_JifenHistory_Code_currentPage=1 Then
		
		Call MemberInformation_JifenHistory_Code_showMain
		
	Else
		'如果传递过来的Page当前页值不大，在应有的页数范围之内时，理应(MemberInformation_JifenHistory_Code_currentPage-1) * MemberInformation_JifenHistory_Code_maxPerPage < MemberInformation_JifenHistory_Code_totalPut，此时进行一些处理.
		if (MemberInformation_JifenHistory_Code_currentPage-1) * MemberInformation_JifenHistory_Code_maxPerPage < MemberInformation_JifenHistory_Code_totalPut then
			'指针指到(MemberInformation_JifenHistory_Code_currentPage-1)页（前一页）的最后一个记录处.
			MemberInformation_JifenHistory_Code_RS.Move  (MemberInformation_JifenHistory_Code_currentPage-1) * MemberInformation_JifenHistory_Code_maxPerPage
			'MemberInformation_JifenHistory_Code_RS.BookMark？
			Dim MemberInformation_JifenHistory_Code_bookMark
			MemberInformation_JifenHistory_Code_bookMark = MemberInformation_JifenHistory_Code_RS.BookMark
			
			Call MemberInformation_JifenHistory_Code_showMain
			
		else
		'如果传递过来的Page当前页值很大，超过了应有的页数时.打开第一页.
			MemberInformation_JifenHistory_Code_currentPage=1
			
			Call MemberInformation_JifenHistory_Code_showMain
			
		end if
	End If
End If
'主体需要控制的部分.

%>
<!--列表 End-->






<%
'菜品列表输出.
'针对性的列表内容部分.
Sub MemberInformation_JifenHistory_Code_showMain()
   	Dim MemberInformation_JifenHistory_Code_i
    MemberInformation_JifenHistory_Code_i=0
%>
		<% '=strClassNameDIV %>
		<%
        If MemberInformation_JifenHistory_Code_RS.EOF Then
        %>
            <ul class="hyzs_hypl"><li><img src="/images/ico/small/emotion_unhappy.png" /> 暂时还没有<% =MemberInformation_JifenHistory_Code_UnitName %>哦 ... ...</li></ul>
        <%
        End If
        %>
       <%
       If Not MemberInformation_JifenHistory_Code_RS.EOF Then
       %>
       <ul class="hyzs_hypl">
       <%
       End If
       %>
        <%
'        Dim MemberInformation_JifenHistory_Code_rsTmp_account_information,MemberInformation_JifenHistory_Code_strTmp_Product__Photos
'        MemberInformation_JifenHistory_Code_strTmp_Product__Photos=""
        
        Do While Not MemberInformation_JifenHistory_Code_RS.EOF
            '获取第一张菜品图片.
'            Set MemberInformation_JifenHistory_Code_rsTmp_account_information=CONN.Execute("SELECT TOP 1 * FROM [CXBG_product] WHERE id="& MemberInformation_JifenHistory_Code_RS("product_id") &"")
'            If Not MemberInformation_JifenHistory_Code_rsTmp_account_information.Eof Then
'                'MemberInformation_JifenHistory_Code_strTmp_Product__Photos=Replace(MemberInformation_JifenHistory_Code_rsTmp_account_information("photos_src"),"/uploadimages/","/uploadimages/120/")
'            Else
'                'MemberInformation_JifenHistory_Code_strTmp_Product__Photos="/images/NoPic.gif"
'				MemberInformation_JifenHistory_Code_rsTmp_account_information.Close
'				MemberInformation_JifenHistory_Code_RS.Close
'				Set MemberInformation_JifenHistory_Code_RS=Nothing
'				Exit Sub
'            End If
            
        %>
         <li>
	       <table width="100%" border="0" cellspacing="0" cellpadding="0">
             <tr>
               <td width="17%" rowspan="3" valign="top"><img src="<% If MemberInformation_JifenHistory_Code_RS("JifenWhichOperationRule")="+" Then Response.Write "/images/ico/coins_add.png" Else Response.Write "/images/ico/coins_delete.png" %>" /></td>
               <td width="57%" class="coloreee">
               		
                    &nbsp;积分历史所属类型： <% =MemberInformation_JifenHistory_Code_RS("JifenWhichClassname") %>
                    
               </td>
               <td width="26%" class="coloreee">
               		<span class="xxright">
                    	积分数量：<% If MemberInformation_JifenHistory_Code_RS("JifenWhichOperationRule")="+" Then Response.Write "<span style=color:green>+ " Else Response.Write "<span style=color:red>- " %><% =MemberInformation_JifenHistory_Code_RS("jifen") %></span>
                    </span>
               </td>
             </tr>
             <tr>
               <td colspan="2">
               		<span class="fontred">
                    	积分历史描述：
                        
                    </span>
               </td>
               </tr>
             <tr>
               <td colspan="2">
               		<% =MemberInformation_JifenHistory_Code_RS("JifenDescription") %>
                    <br />
                    <span style="color:#999; font-size:;">日期：<% =MemberInformation_JifenHistory_Code_RS("adddate") %></span>
               </td>
               </tr>
           </table>
		 </li>
        <%
            'MemberInformation_JifenHistory_Code_rsTmp_account_information.Close	'临时销毁临时记录集.
			
			MemberInformation_JifenHistory_Code_i=MemberInformation_JifenHistory_Code_i+1
            If MemberInformation_JifenHistory_Code_i >= MemberInformation_JifenHistory_Code_maxPerPage Then Exit Do
            MemberInformation_JifenHistory_Code_RS.MoveNext
        Loop
        
        '销毁临时对象.
        'Set MemberInformation_JifenHistory_Code_rsTmp_account_information = Nothing
        %>
	   <%
       'If Not MemberInformation_JifenHistory_Code_RS.EOF Then
       %>
       </ul>
       <%
       'End If
       %>
	
<% If MemberInformation_JifenHistory_Code_i>0 Then %>
<!--翻页-->
		<div class="clubzs">
		<%
		Response.Write Coke.ShowPage(MemberInformation_JifenHistory_Code_strFileName,MemberInformation_JifenHistory_Code_totalPut,MemberInformation_JifenHistory_Code_maxPerPage,True,True,MemberInformation_JifenHistory_Code_UnitName)
		%>
		</div>
<!--翻页-->
<% End If %>

<%
	MemberInformation_JifenHistory_Code_RS.Close
	Set MemberInformation_JifenHistory_Code_RS=Nothing
End Sub
%>