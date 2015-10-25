
 	      <%
		  '俱乐部广告.
		  Dim rsClubAdvertisement_Code,sqlClubAdvertisement_Code,countClubAdvertisement_Code,numClubAdvertisement_Code
		  	sqlClubAdvertisement_Code="select top 6 * from [CXBG_advertisement_club] where isOnpublic=1 order by RootID,OrderID"
			Set rsClubAdvertisement_Code=Server.CreateObject("Adodb.RecordSet")
			rsClubAdvertisement_Code.Open sqlClubAdvertisement_Code,CONN,1,1
			countClubAdvertisement_Code=rsClubAdvertisement_Code.RecordCount
			numClubAdvertisement_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsClubAdvertisement_Code.EOF Then
		  %>
		  	PImgPlayer.addItem( "", "", "/images/NoPic.png"); 
		  <%
		  End If
		  %>
          
          <%
		  Do While Not rsClubAdvertisement_Code.EOF
		  %>
          	PImgPlayer.addItem( "<% =rsClubAdvertisement_Code("classname") %>", "<% If rsClubAdvertisement_Code("isDirectingLink")=1 Then Response.Write rsClubAdvertisement_Code("DirectingLink") Else Response.Write "/ChineseDish/ChineseDish.Welcome?keywords="& rsClubAdvertisement_Code("keywords") %>", "<% If rsClubAdvertisement_Code("photo")<>"" Then Response.Write rsClubAdvertisement_Code("photo") Else Response.Write "/images/NoPic.png" %>"); 
          <%
			  numClubAdvertisement_Code=numClubAdvertisement_Code+1
			  rsClubAdvertisement_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsClubAdvertisement_Code.Close
		  Set rsClubAdvertisement_Code=Nothing
		  %>