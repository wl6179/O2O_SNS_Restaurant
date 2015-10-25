		   <%
		  '礼品券 会员乐享.
		  Dim rsGiftCertificatedArea_Code,sqlGiftCertificatedArea_Code,countGiftCertificatedArea_Code,numGiftCertificatedArea_Code
		  	sqlGiftCertificatedArea_Code="select top 4 * from [CXBG_GiftCertificated] where deleted=0 and isOnpublic=1 order by details_orderid desc,id desc"
			Set rsGiftCertificatedArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsGiftCertificatedArea_Code.Open sqlGiftCertificatedArea_Code,CONN,1,1
			countGiftCertificatedArea_Code=rsGiftCertificatedArea_Code.RecordCount
			numGiftCertificatedArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsGiftCertificatedArea_Code.EOF Then
		  %>
		  <!--<li>
		    <a class="abg04" href=""><img src="images/cpimg/tjtc_01.jpg" width="210" height="75" /></a>
			<p class="txtnoml linheight22"><span class="fontred22">01</span> <a href="">欢迎光临</a></p>
			<p class="xjcplxx"><img src="images/xx.gif" width="15" height="16" /><img src="images/xx.gif" width="15" height="16" /><img src="images/xx.gif" width="15" height="16" /><img src="images/xx.gif" width="15" height="16" /><img src="images/xx.gif" width="15" height="16" /></p>
		  </li>-->
		  <%
		  End If
		  %>
          
          <%
		  Do While Not rsGiftCertificatedArea_Code.EOF
		  %>
           <li>
		     <img src="<% If rsGiftCertificatedArea_Code("photo")<>"" Then Response.Write rsGiftCertificatedArea_Code("photo") Else Response.Write "/images/NoPic.png" %>" />
			 <div class="dhqtxt_mid">
			   <p>礼品券名：<span class="fontred"><% =rsGiftCertificatedArea_Code("topic") %></span></p>
			   <p>兑换积分：<span class="fontred"><% =rsGiftCertificatedArea_Code("jifen") %>分</span><span style="font-size:10px;">&nbsp;&nbsp;<!--<a href="/Details/DetailsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsGiftCertificatedArea_Code("id") ) %>" target="_blank">什么积分?</a>--></span></p>
			   <p>今年有效期：<span class="fontred">
			   <% =Right(CokeShow.filt_DateStr(rsGiftCertificatedArea_Code("StartDateValid")),5) %>&nbsp;~&nbsp;<% =Right(CokeShow.filt_DateStr(rsGiftCertificatedArea_Code("StopDateValid")),5) %>
               </span></p>
			 </div>
			 <div class="dhqbutton"><a class="button_img47" href="/Club/GiftCertificatedsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsGiftCertificatedArea_Code("id") ) %>" target="_blank">兑换</a></div>
		   </li>
           <%
			  numGiftCertificatedArea_Code=numGiftCertificatedArea_Code+1
			  rsGiftCertificatedArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsGiftCertificatedArea_Code.Close
		  Set rsGiftCertificatedArea_Code=Nothing
		  %>
		   