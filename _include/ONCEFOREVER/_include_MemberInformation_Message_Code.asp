		  <%
		  '帐号管理中心之收藏菜列表.
'初始化赋值.
'变量定义.
Dim MemberInformation_Message_Code_maxPerPage					'设置当前模块分页设置.
Dim MemberInformation_Message_Code_CurrentTableName			'设置当前模块所涉及的[表]名.
Dim MemberInformation_Message_Code_CurrentPageNow				'设置当前模块所在页面的文件名.
Dim MemberInformation_Message_Code_UnitName					'此主要列表展示中，所涉及的记录的单位名称.
Dim MemberInformation_Message_Code_totalPut,MemberInformation_Message_Code_totalPages,MemberInformation_Message_Code_currentPage			'分页用的控制变量.
Dim MemberInformation_Message_Code_RS, MemberInformation_Message_Code_sql									'查询列表记录用的变量.
Dim MemberInformation_Message_Code_FoundErr,MemberInformation_Message_Code_ErrMsg							'控制错误流程用的控制变量.
Dim MemberInformation_Message_Code_strFileName								'构建查询字符串用的控制变量.
Dim MemberInformation_Message_Code_ExecuteSearch,MemberInformation_Message_Code_Keyword,MemberInformation_Message_Code_TypeSearch,MemberInformation_Message_Code_Action	'构建查询字符串以及流程控制用的控制变量.
Dim MemberInformation_Message_Code_strGuide								'导航文字.



'接收参数.
MemberInformation_Message_Code_maxPerPage			=6
MemberInformation_Message_Code_CurrentTableName 	="[CXBG_account_Message]"		'此模块涉及的[表]名.
MemberInformation_Message_Code_CurrentPageNow 	="/ONCEFOREVER/Account_Message.Welcome?Welcome=1"
MemberInformation_Message_Code_UnitName			="我的留言"
MemberInformation_Message_Code_currentPage		=CokeShow.filtRequest(Request("page"))
MemberInformation_Message_Code_ExecuteSearch	=CokeShow.filtRequest(Request("MemberInformation_Message_Code_ExecuteSearch"))
MemberInformation_Message_Code_Keyword		=CokeShow.filtRequest(Request("MemberInformation_Message_Code_Keyword"))
MemberInformation_Message_Code_TypeSearch		=CokeShow.filtRequest(Request("MemberInformation_Message_Code_TypeSearch"))
MemberInformation_Message_Code_Action			=CokeShow.filtRequest(Request("MemberInformation_Message_Code_Action"))


'处理参数.
'处理查询执行 控制变量.
If MemberInformation_Message_Code_ExecuteSearch="" Then
	MemberInformation_Message_Code_ExecuteSearch=0
Else
	If isNumeric(MemberInformation_Message_Code_ExecuteSearch) Then MemberInformation_Message_Code_ExecuteSearch=CokeShow.CokeClng(MemberInformation_Message_Code_ExecuteSearch) Else MemberInformation_Message_Code_ExecuteSearch=0
End If
'构造查询字符串.
If Instr(MemberInformation_Message_Code_CurrentPageNow, "?")>0 Then MemberInformation_Message_Code_strFileName=MemberInformation_Message_Code_CurrentPageNow &"&MemberInformation_Message_Code_ExecuteSearch="& MemberInformation_Message_Code_ExecuteSearch Else MemberInformation_Message_Code_strFileName=MemberInformation_Message_Code_CurrentPageNow &"?MemberInformation_Message_Code_ExecuteSearch="& MemberInformation_Message_Code_ExecuteSearch
If MemberInformation_Message_Code_Keyword<>"" Then
	MemberInformation_Message_Code_strFileName=MemberInformation_Message_Code_strFileName&"&MemberInformation_Message_Code_Keyword="& MemberInformation_Message_Code_Keyword
End If
If MemberInformation_Message_Code_TypeSearch<>"" Then
	MemberInformation_Message_Code_strFileName=MemberInformation_Message_Code_strFileName&"&MemberInformation_Message_Code_TypeSearch="& MemberInformation_Message_Code_TypeSearch
End If


'处理当前页码的控制变量，默认为第一页.
If MemberInformation_Message_Code_currentPage<>"" Then
    If isNumeric(MemberInformation_Message_Code_currentPage) Then MemberInformation_Message_Code_currentPage=CokeShow.CokeClng(MemberInformation_Message_Code_currentPage) Else MemberInformation_Message_Code_currentPage=1
Else
	MemberInformation_Message_Code_currentPage=1
End If


'主体控制部分 Begin
Select Case MemberInformation_Message_Code_ExecuteSearch
	Case 0		
		MemberInformation_Message_Code_sql="SELECT TOP 500 * FROM "& MemberInformation_Message_Code_CurrentTableName &" WHERE deleted=0 AND ReplyID=0 AND Account_LoginID='"& RS("username") &"' ORDER BY id DESC"
		MemberInformation_Message_Code_strGuide=MemberInformation_Message_Code_strGuide & "所有"& MemberInformation_Message_Code_UnitName
'Response.Write MemberInformation_Message_Code_sql
'	Case 1
'		MemberInformation_Message_Code_sql="SELECT TOP 500 * FROM "& MemberInformation_Message_Code_CurrentTableName &" WHERE deleted=0 AND isOnsale=1 AND is_display_newproduct=1 ORDER BY id DESC"
'		MemberInformation_Message_Code_strGuide=MemberInformation_Message_Code_strGuide & "痴心不改餐厅最新上架"& MemberInformation_Message_Code_UnitName
'	Case 2
'		MemberInformation_Message_Code_sql="SELECT TOP 500 * FROM "& MemberInformation_Message_Code_CurrentTableName &" WHERE deleted=0 AND isOnsale=1 AND isSales=1 ORDER BY isSales_StopDate DESC"
'		MemberInformation_Message_Code_strGuide=MemberInformation_Message_Code_strGuide & "本月促销"& MemberInformation_Message_Code_UnitName
'	Case 10
'		If MemberInformation_Message_Code_Keyword="" Then
'			MemberInformation_Message_Code_sql="SELECT TOP 500 * FROM "& MemberInformation_Message_Code_CurrentTableName &" WHERE deleted=0 AND isOnsale=1 ORDER BY id DESC"
'			MemberInformation_Message_Code_strGuide=MemberInformation_Message_Code_strGuide & "搜索所有"& MemberInformation_Message_Code_UnitName
'		Else
'			Select Case MemberInformation_Message_Code_TypeSearch
'				Case "Brand"
'					If IsNumeric(MemberInformation_Message_Code_Keyword)=False Then
'						MemberInformation_Message_Code_FoundErr=True
'						MemberInformation_Message_Code_ErrMsg=MemberInformation_Message_Code_ErrMsg &"<br /><li>"& MemberInformation_Message_Code_UnitName &"您必须选择一个品牌！</li>"
'					Else
'						MemberInformation_Message_Code_sql="select TOP 500 * from "& MemberInformation_Message_Code_CurrentTableName &" where deleted=0 AND isOnsale=1 and product_brand_id="& CokeShow.CokeClng(MemberInformation_Message_Code_Keyword) &" order by id desc"
'						MemberInformation_Message_Code_strGuide=MemberInformation_Message_Code_strGuide & MemberInformation_Message_Code_UnitName &"品牌为:<font color=red> " & CokeShow.otherField("[CXBG_product_brand]",CokeShow.CokeClng(MemberInformation_Message_Code_Keyword),"classid","classname",True,0) & " </font>的"& MemberInformation_Message_Code_UnitName
'					End If
'					
'				Case "ProductName"
'					MemberInformation_Message_Code_sql="select TOP 500 * from "& MemberInformation_Message_Code_CurrentTableName &" where deleted=0 AND isOnsale=1 and (ProductName like '%"& MemberInformation_Message_Code_Keyword &"%' OR product_keywords like '%"& MemberInformation_Message_Code_Keyword &"%') order by id desc"
'					MemberInformation_Message_Code_strGuide=MemberInformation_Message_Code_strGuide & "菜品中含有“ <font color=red>" & MemberInformation_Message_Code_Keyword & "</font> ”的"& MemberInformation_Message_Code_UnitName
''response.Write Keyword					
'			End Select
'			
'		End If
		
		
	Case Else
		MemberInformation_Message_Code_FoundErr=True
		MemberInformation_Message_Code_ErrMsg=MemberInformation_Message_Code_ErrMsg & "<br /><li>错误的参数！</li>"
	
End Select

'拦截错误.
If MemberInformation_Message_Code_FoundErr=True Then
	'Response.Clear()
	Err.Raise vbObjectError + 6666, "列表查询出现异常", "如下异常："& MemberInformation_Message_Code_ErrMsg
	Response.End()
End If

If Not IsObject(CONN) Then link_database
Set MemberInformation_Message_Code_RS=Server.CreateObject("Adodb.RecordSet")
'	Response.Write "<br />"& MemberInformation_Message_Code_sql
'	Response.End 
MemberInformation_Message_Code_RS.Open MemberInformation_Message_Code_sql,CONN,1,1

'主体控制部分 End
%>

<!--列表 Begin-->
<%
'主体需要控制的部分.
If MemberInformation_Message_Code_RS.Eof And MemberInformation_Message_Code_RS.Bof Then
	MemberInformation_Message_Code_strGuide=MemberInformation_Message_Code_strGuide & " &#187; 共找到 <font color=red>0</font> 个"& MemberInformation_Message_Code_UnitName
	Call MemberInformation_Message_Code_showMain
Else
	MemberInformation_Message_Code_totalPut=MemberInformation_Message_Code_RS.RecordCount		'记录总数.
	MemberInformation_Message_Code_strGuide=MemberInformation_Message_Code_strGuide & " &#187; 共找到 <font color=red>" & MemberInformation_Message_Code_totalPut & "</font> 个"& MemberInformation_Message_Code_UnitName
	
	
	'处理页码
	If MemberInformation_Message_Code_currentPage<1 Then
		MemberInformation_Message_Code_currentPage=1
	End If
	'如果传递过来的Page当前页值很大，超过了应有的页数时，进行处理.
	If (MemberInformation_Message_Code_currentPage-1) * MemberInformation_Message_Code_maxPerPage > MemberInformation_Message_Code_totalPut Then
		If (MemberInformation_Message_Code_totalPut Mod MemberInformation_Message_Code_maxPerPage)=0 Then
			'如果整好够页数，赋予当前页最大页.
			MemberInformation_Message_Code_currentPage= MemberInformation_Message_Code_totalPut \ MemberInformation_Message_Code_maxPerPage
		Else
			'如果不整好，最有一页只有零散几条记录（不丰满的多余页），赋予当前页最大页.（不能整除情况下计算）
			MemberInformation_Message_Code_currentPage= MemberInformation_Message_Code_totalPut \ MemberInformation_Message_Code_maxPerPage + 1
		End If

	End If
	If MemberInformation_Message_Code_currentPage=1 Then
		
		Call MemberInformation_Message_Code_showMain
		
	Else
		'如果传递过来的Page当前页值不大，在应有的页数范围之内时，理应(MemberInformation_Message_Code_currentPage-1) * MemberInformation_Message_Code_maxPerPage < MemberInformation_Message_Code_totalPut，此时进行一些处理.
		if (MemberInformation_Message_Code_currentPage-1) * MemberInformation_Message_Code_maxPerPage < MemberInformation_Message_Code_totalPut then
			'指针指到(MemberInformation_Message_Code_currentPage-1)页（前一页）的最后一个记录处.
			MemberInformation_Message_Code_RS.Move  (MemberInformation_Message_Code_currentPage-1) * MemberInformation_Message_Code_maxPerPage
			'MemberInformation_Message_Code_RS.BookMark？
			Dim MemberInformation_Message_Code_bookMark
			MemberInformation_Message_Code_bookMark = MemberInformation_Message_Code_RS.BookMark
			
			Call MemberInformation_Message_Code_showMain
			
		else
		'如果传递过来的Page当前页值很大，超过了应有的页数时.打开第一页.
			MemberInformation_Message_Code_currentPage=1
			
			Call MemberInformation_Message_Code_showMain
			
		end if
	End If
End If
'主体需要控制的部分.

%>
<!--列表 End-->






<%
'菜品列表输出.
'针对性的列表内容部分.
Sub MemberInformation_Message_Code_showMain()
   	Dim MemberInformation_Message_Code_i
    MemberInformation_Message_Code_i=0
%>
		<% '=strClassNameDIV %>
		<%
        If MemberInformation_Message_Code_RS.EOF Then
        %>
            <ul class="hyzs_hypl"><li><img src="/images/ico/small/emotion_unhappy.png" /> 暂时还没有<% =MemberInformation_Message_Code_UnitName %>哦 ... ...</li></ul>
        <%
        End If
        %>
       <%
       If Not MemberInformation_Message_Code_RS.EOF Then
       %>
       <ul class="hyzs_hypl">
       <%
       End If
       %>
        <%
'        Dim MemberInformation_Message_Code_rsTmp_account_information,MemberInformation_Message_Code_strTmp_Product__Photos
'        MemberInformation_Message_Code_strTmp_Product__Photos=""
        
		Dim rs222,sql222
		
        Do While Not MemberInformation_Message_Code_RS.EOF
            '获取第一张菜品图片.
'            Set MemberInformation_Message_Code_rsTmp_account_information=CONN.Execute("SELECT TOP 1 * FROM [CXBG_product] WHERE id="& MemberInformation_Message_Code_RS("product_id") &"")
'            If Not MemberInformation_Message_Code_rsTmp_account_information.Eof Then
'                'MemberInformation_Message_Code_strTmp_Product__Photos=Replace(MemberInformation_Message_Code_rsTmp_account_information("photos_src"),"/uploadimages/","/uploadimages/120/")
'            Else
'                'MemberInformation_Message_Code_strTmp_Product__Photos="/images/NoPic.gif"
'				MemberInformation_Message_Code_rsTmp_account_information.Close
'				MemberInformation_Message_Code_RS.Close
'				Set MemberInformation_Message_Code_RS=Nothing
'				Exit Sub
'            End If
            
        %>
         <li>
	       <table width="100%" border="0" cellspacing="0" cellpadding="0">
             <tr>
               <td width="17%" rowspan="3" valign="top"><img src="<% =Coke.ShowMemberSexPicURL(RS("id")) %>" width="36" /></td>
               <td width="57%" class="coloreee">
               		
                    &nbsp;我的留言标题： <strong><% =MemberInformation_Message_Code_RS("title") %></strong>
                    
               </td>
               <td width="26%" class="coloreee">
               		<span class="xxright">
                    	
                    </span>
               </td>
             </tr>
             <tr>
               <td colspan="2">
               		<span class="fontred">
                    	我的留言主要内容：
                        
                    </span>
               </td>
               </tr>
             <tr>
               <td colspan="2">
               		<% =MemberInformation_Message_Code_RS("content") %>
                    <br />
                    <span style="color:#999; font-size:;">日期：<% =MemberInformation_Message_Code_RS("adddate") %></span>
               </td>
               </tr>
           </table>
           		<%
				'循环出相应的回复记录.
				'Dim rs222,sql222
				sql222="SELECT TOP 10 * FROM "& MemberInformation_Message_Code_CurrentTableName &" WHERE deleted=0 AND Account_LoginID='Coke' AND ReplyID="& MemberInformation_Message_Code_RS("id") &" ORDER BY id ASC"
				Set rs222=CONN.Execute(sql222)
'Response.Write sql222
				'如果有，则循环出来.
				Do While Not rs222.Eof
				%>
				
                <!--回复-->
                <table width="100%" border="0" cellspacing="0" cellpadding="0" style="width:580px; float:right;">
                 <tr>
                   <td width="17%" rowspan="3" valign="top">
                   <img src="<% If rs222("Account_LoginID")="Coke" Then Response.Write "/images/ico/email_go.png" Else Response.Write Coke.ShowMemberSexPicURL( CokeShow.otherField("[CXBG_account]",rs222("Account_LoginID"),"username","id",False,0) ) %>" />
                   <br />
                   回复敬爱的顾客
                   </td>
                   <td width="57%" class="coloreee">
                        
                        &nbsp;餐厅回执标题： <% =rs222("title") %>
                        
                   </td>
                   <td width="26%" class="coloreee">
                        <span class="xxright">
                            餐厅客服人员回复您
                        </span>
                   </td>
                 </tr>
                 <tr>
                   <td colspan="2">
                        <span class="fontred">
                            餐厅回执主要内容如下：
                            
                        </span>
                   </td>
                   </tr>
                 <tr>
                   <td colspan="2">
                        <% =rs222("content") %>
                        <br />
                        <span style="color:#999; font-size:;">日期：<% =rs222("adddate") %></span>
                   </td>
                   </tr>
               </table>
               <!--回复-->
               		<%
					'如果有未读的回复，则更新为已读isRead=1.
					'只有当会员点击打开留言信息管理页时，才会更新已读状态.
					If rs222("isRead")=0 And rs222("toWho")=MemberInformation_Message_Code_RS("Account_LoginID") Then
					CONN.Execute("UPDATE "& MemberInformation_Message_Code_CurrentTableName &" SET isRead=1 WHERE deleted=0 AND id="& rs222("id"))	'更新为已读状态.
					End If
					%>
               
               <%
					rs222.MoveNext
				Loop
				
				rs222.Close
				'
				%>
                
                
                
                
               
		 </li>
		
			
        
        <%
            'MemberInformation_Message_Code_rsTmp_account_information.Close	'临时销毁临时记录集.
			
			MemberInformation_Message_Code_i=MemberInformation_Message_Code_i+1
            If MemberInformation_Message_Code_i >= MemberInformation_Message_Code_maxPerPage Then Exit Do
            MemberInformation_Message_Code_RS.MoveNext
        Loop
        
        '销毁临时对象.
        'Set MemberInformation_Message_Code_rsTmp_account_information = Nothing
		Set rs222=Nothing
        %>
	   <%
       'If Not MemberInformation_Message_Code_RS.EOF Then
       %>
       </ul>
       <%
       'End If
       %>
	
<% If MemberInformation_Message_Code_i>0 Then %>
<!--翻页-->
		<div class="clubzs" style="clear:both;">
		<%
		Response.Write Coke.ShowPage(MemberInformation_Message_Code_strFileName,MemberInformation_Message_Code_totalPut,MemberInformation_Message_Code_maxPerPage,True,True,MemberInformation_Message_Code_UnitName)
		%>
		</div>
<!--翻页-->


<!--留言框-->
<form action="/ONCEFOREVER/Account.Services.Private.asp" method="post" name="SendMessageForm" id="SendMessageForm"


>
<ul class="hyzs_hypl">
<li>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="width:580px;">
	<tr>
        <td width="17%" rowspan="3" valign="top">
        <img src="<% =Coke.ShowMemberSexPicURL(RS("id")) %>" width="36" />
        <br />
        我要留言
		</td>
		<td width="57%" class="coloreee">
        
        &nbsp;如您有很好的想法，可以随时告诉痴心不改餐厅哦！
        
		</td>
		<td width="26%" class="coloreee">
        <span class="xxright">
           
        </span>
		</td>
	</tr>
	<tr>
		<td colspan="2">
        <span class="fontred">
            
            
            填写留言标题： <span class="fontred">*</span>
            <input type="text" id="title" name="title"
                dojoType="dijit.form.ValidationTextBox"
                required="true"
                propercase="false"
                promptMessage="您的想法对我们来说很重要，我们将心存感激聆听您的想法！"
                invalidMessage="请填写50字以内"
                trim="true"
                lowercase="false"
                value=""
                regExp=".{1,50}"
                style="width:200px;"
                class="input_200"
                />
            <br />
            填写留言内容： <span class="fontred">*</span> ( 188字内)
        </span>
		</td>
   </tr>
   <tr>
		<td colspan="2">
        <textarea class="textarea_400" name="content" id="testBarInput" style="width:480px;"></textarea>
        
        <div id="testBar" style='width:400px; height:10px; font-size:10px; line-height:10px;' dojoType="dijit.ProgressBar" width="400"
          annotate="true" maximum="188" duration="2000">
          <script type="dojo/method" event="report">
            //return dojo.string.substitute("${0} /${1}", [this.progress, this.maximum]);
            return this.progress + "   Word  , Welcome To chixinbugai.me";
          </script>
          <script type="dojo/method">
            dojo.connect(dojo.byId("testBarInput"), "onkeyup", 
              dojo.hitch(this, function(e){
                if(e.target.value.length > Number(this.maximum)){          
                  e.target.value = e.target.value.substring(0, this.maximum);
                }
                this.update({progress:e.target.value.length});
              })
            );
          </script>
        </div>
		</td>
   </tr>
   
   <tr>
		<td>
        	验证码： <span class="fontred">*</span>
        </td>
        <td colspan="2">
        
        <input type="text" id="CodeStr" name="CodeStr" size="4"
        dojoType="dijit.form.ValidationTextBox"
        required="true"
        propercase="false"
        invalidMessage="请填写4位数字！"
        trim="true"
        lowercase="false"
        value=""
        regExp="\d{4}"
        style="width:80px;"
        class="input_150"
        maxlength="4"
        />
        &nbsp;
        <img id="GetCode" src="/public/code.asp" style="cursor:hand; float:none;" onClick="this.src='/public/code.asp?c='+Math.random()" alt="点击更换验证码" />
        <!--&nbsp;
        <a href="javascript:return false;" onClick="dojo.byId('GetCode').src='/public/code.asp?c='+Math.random()" class="fontgreen">重新刷新验证码</a>-->
        
		</td>
   </tr>
   <tr>
       <td colspan="3" style="text-align:center; width:100%;">
        <button type="submit" id="theSubmitButton_SendMessageForm" 
        dojoType="dijit.form.Button"
        class="button"
        >
        &nbsp;&nbsp;&nbsp;发送留言&nbsp;
        </button>
        
        <br />
        <span id="response" style="color:#F30;">&nbsp;</span>
        
        <!--<br />
        <span style="color:#999; font-size:;">日期：<% =Now() %></span>-->
       </td>
   </tr>
</table>
</li>
</ul>
<input type="hidden" id="ReplyID" name="ReplyID" value="0" />
<input type="hidden" id="ServicesAction" name="ServicesAction" value="addAccount_SendMessage" />
</form>
<!--留言框-->

<% End If %>

<%
	MemberInformation_Message_Code_RS.Close
	Set MemberInformation_Message_Code_RS=Nothing
End Sub
%>