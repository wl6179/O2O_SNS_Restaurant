
 	      <%
		  '俱乐部广告.
		  Dim rsClubQuestionnaires_Code,sqlClubQuestionnaires_Code,countClubQuestionnaires_Code,numClubQuestionnaires_Code
		  	sqlClubQuestionnaires_Code="select top 10 * from [CXBG_Questionnaire] where 1=1 order by RootID,OrderID"
			Set rsClubQuestionnaires_Code=Server.CreateObject("Adodb.RecordSet")
			rsClubQuestionnaires_Code.Open sqlClubQuestionnaires_Code,CONN,1,1
			countClubQuestionnaires_Code=rsClubQuestionnaires_Code.RecordCount
			numClubQuestionnaires_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsClubQuestionnaires_Code.EOF Then
		  %>
		  	PImgPlayer.addItem( "", "", "/images/NoPic.png"); 
		  <%
		  End If
		  %>
      <form action="/ONCEFOREVER/Account.Services.Public.asp" method="post" name="QuestionnairesForm" id="QuestionnairesForm"
      
      
      ><!--dojoType="dijit.form.Form"execute="processFormAjax"-->
	  <ul class="club_wjlist">
          <%
		  '记下当前是第几次出现depth=0一级分类的状态.
		  Dim DisplayDepthTimes,intTmp123
		  DisplayDepthTimes=0
		  intTmp123=0
		  
		  Do While Not rsClubQuestionnaires_Code.EOF
		  	If rsClubQuestionnaires_Code("depth")=0 Then DisplayDepthTimes=DisplayDepthTimes+1 : intTmp123=rsClubQuestionnaires_Code("classid")
			'如果第二次出现一级分类的调查项目时，跳出循环.
			If DisplayDepthTimes=2 Then Exit Do
		  %>
	<% If rsClubQuestionnaires_Code("depth")=0 Then %>
        <p class="fontred14_txt"><% =rsClubQuestionnaires_Code("classname") %></p>
    <% Else %>
		<p>
            <input type="radio" name="QuestionnairesID" id="QuestionnairesID<% =rsClubQuestionnaires_Code("classid") %>"
            dojoType="dijit.form.RadioButton"
            value="<% =rsClubQuestionnaires_Code("classid") %>"
            >
            </input>
            &nbsp;
            <label for="QuestionnairesID<% =rsClubQuestionnaires_Code("classid") %>"><% =rsClubQuestionnaires_Code("classname") %></label>	
        </p>
    <% End If %>
          <%
			  numClubQuestionnaires_Code=numClubQuestionnaires_Code+1
			  rsClubQuestionnaires_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsClubQuestionnaires_Code.Close
		  Set rsClubQuestionnaires_Code=Nothing
		  %>
		<p class="ancent" style=" text-align:center; padding:18px 0px;">
        <button type="submit" id="theSubmitButton_QuestionnairesForm" 
        dojoType="dijit.form.Button"
        class=""
        >
        &nbsp;提交问卷&nbsp;
        </button>
        </p>
        <p class="ancent" style=" text-align:center; padding:0px 0px;">
      	<span id="response" style="color:#F30;">&nbsp;</span>
        </p>
	  </ul>
      <input type="hidden" name="id"
        value="<% =intTmp123 %>"
        >
      <input type="hidden" name="ServicesAction"
        value="SubmitQuestionnaires"
        >
      </form>
      