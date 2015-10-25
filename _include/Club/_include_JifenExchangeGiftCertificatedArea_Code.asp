		  <%
		  '礼品券.
		  Dim rsJifenExchangeGiftCertificatedArea_Code,sqlJifenExchangeGiftCertificatedArea_Code,countJifenExchangeGiftCertificatedArea_Code,numJifenExchangeGiftCertificatedArea_Code
		  	sqlJifenExchangeGiftCertificatedArea_Code="select top 4 * from [CXBG_GiftCertificated] where deleted=0 and isOnpublic=1 order by details_orderid desc"
			Set rsJifenExchangeGiftCertificatedArea_Code=Server.CreateObject("Adodb.RecordSet")
			rsJifenExchangeGiftCertificatedArea_Code.Open sqlJifenExchangeGiftCertificatedArea_Code,CONN,1,1
			countJifenExchangeGiftCertificatedArea_Code=rsJifenExchangeGiftCertificatedArea_Code.RecordCount
			numJifenExchangeGiftCertificatedArea_Code=1
		  %>
          <%
		  '如果记录为空.
		  If rsJifenExchangeGiftCertificatedArea_Code.EOF Then
		  %>
		  <!--<li>
		     <img src="/images/cpimg/lpq_01.jpg" />
			 <div class="dhqtxt_mid">
			   <p>礼品券名：<span class="fontred">礼品名称礼品</span></p>
			   <p>兑换积分：<span class="fontred">100分</span></p>
			   <p>有效日期：<span class="fontred">2010/05-2010/08</span></p>
			 </div>
			 <div class="dhqbutton"><a class="button_img47" href="">兑换</a></div>
		   </li>-->
		  <%
		  End If
		  %>
          
          <%
		  Do While Not rsJifenExchangeGiftCertificatedArea_Code.EOF
		  %>
          <li>
		     <img src="<% If rsJifenExchangeGiftCertificatedArea_Code("photo")<>"" Then Response.Write rsJifenExchangeGiftCertificatedArea_Code("photo") Else Response.Write "/images/NoPic.png" %>" width="55" height="55" />
			 <div class="dhqtxt_mid">
			   <p>礼品券名：<span class="fontred"><% =rsJifenExchangeGiftCertificatedArea_Code("topic") %></span></p>
			   <p>兑换积分：<span class="fontred"><% =rsJifenExchangeGiftCertificatedArea_Code("jifen") %>分</span></p>
			   <p>有效日期：<span class="fontred">
			   <% '=Year(rsJifenExchangeGiftCertificatedArea_Code("StartDateValid")) %><!--/--><% '=Month(rsJifenExchangeGiftCertificatedArea_Code("StartDateValid")) %><!-----><% '=Year(rsJifenExchangeGiftCertificatedArea_Code("StopDateValid")) %><!--/--><% '=Month(rsJifenExchangeGiftCertificatedArea_Code("StopDateValid")) %>
               <% =Right(CokeShow.filt_DateStr(rsJifenExchangeGiftCertificatedArea_Code("StartDateValid")),5) %>&nbsp;~&nbsp;<% =Right(CokeShow.filt_DateStr(rsJifenExchangeGiftCertificatedArea_Code("StopDateValid")),5) %>
               </span></p>
			 </div>
			 <div class="dhqbutton"><a class="button_img47" href="/Club/GiftCertificatedsInformation.Welcome?CokeMark=<% =CokeShow.AddCode_Num( rsJifenExchangeGiftCertificatedArea_Code("id") ) %>" target="_blank">兑换</a></div>
		   </li>
          <%
			  numJifenExchangeGiftCertificatedArea_Code=numJifenExchangeGiftCertificatedArea_Code+1
			  rsJifenExchangeGiftCertificatedArea_Code.MoveNext
		  Loop
		  
		  '关闭记录集.
		  rsJifenExchangeGiftCertificatedArea_Code.Close
		  Set rsJifenExchangeGiftCertificatedArea_Code=Nothing
		  %>