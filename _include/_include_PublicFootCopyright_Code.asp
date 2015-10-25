
        <p style="width:900px;">
            <span style="padding-bottom:18px; margin-bottom:18px;">
            	
                <span style="color:#CD4400;">联盟商家 &amp; 兄弟连</span>
                <br />
                (当您在痴心不改餐厅享用完餐点之后，您还可以享受到以下本餐厅的兄弟连...为您提供的更多优惠活动哦~ 请咨询前台索取优惠券，祝您享受快乐时光！)
                
                <br />
            </span>
            <span>
            <%
			'广告轮换.
			Dim rsFriendlyLink,sqlFriendlyLink,countFriendlyLink,numFriendlyLink
			sqlFriendlyLink="select top 20 * from [CXBG_FriendlyLink] where isOnpublic=1 order by RootID,OrderID"
			Set rsFriendlyLink=Server.CreateObject("Adodb.RecordSet")
			rsFriendlyLink.Open sqlFriendlyLink,CONN,1,1
			countFriendlyLink=rsFriendlyLink.RecordCount
			numFriendlyLink=1
			%>
			<%
			'如果记录为空.
			If rsFriendlyLink.EOF Then
			%>
			<!--<a href="http://www.cokeshow.com.cn/" target="_blank">可乐秀CokeShow</a><span>|</span>-->
			<%
			End If
			%>
			
			<%
			Do While Not rsFriendlyLink.EOF
			%>
            <a href="<% =rsFriendlyLink("DirectingLink") %>" target="_blank"><img src="<% If rsFriendlyLink("photo")<>"" Then Response.Write rsFriendlyLink("photo") Else Response.Write "/images/NoPic.png" %>" width="128" height="50" alt="<% =rsFriendlyLink("classname") %>" /></a><!--<a href="<% =rsFriendlyLink("DirectingLink") %>" target="_blank"><% =rsFriendlyLink("classname") %></a><span>|</span>-->
            <%
				numFriendlyLink=numFriendlyLink+1
				rsFriendlyLink.MoveNext
			Loop
			
			'关闭记录集.
			rsFriendlyLink.Close
			Set rsFriendlyLink=Nothing
			%>
            </span>
        </p>
        
        <hr />
        

        <p>
            <a href="/Details/DetailsInformation.Welcome?CokeMark=JPHRQPP" target="_blank">餐厅介绍</a><span>|</span>
            <a href="/Details/DetailsInformation.Welcome?CokeMark=JPISDRP" target="_blank">社区介绍</a><span>|</span>
            <!--<a href="/Details/DetailsInformation.Welcome?CokeMark=JPJMOQP" target="_blank">团体订餐服务介绍</a><span>|</span>-->
            <a href="/Details/DetailsInformation.Welcome?CokeMark=JPJNNPR" target="_blank">联系我们</a><span>|</span>
            <a href="/Details/DetailsInformation.Welcome?CokeMark=JPHQRRP" target="_blank">乘车路线&amp;地图</a><span>|</span>
            <a href="/Details/DetailsInformation.Welcome?CokeMark=JPJOMPF" target="_blank">餐位预订介绍</a><span>|</span>
            <!--<a href="/Details/DetailsInformation.Welcome?CokeMark=JPJPLOL" target="_blank">菜点预订介绍</a><span>|</span>-->
            <a href="/Details/DetailsInformation.Welcome?CokeMark=JPIRFMN" target="_blank">会员卡介绍</a><span>|</span>
        </p>
        
        
        <p>
        	座位预订：010-64930888,64937666
            &nbsp;&nbsp;&nbsp;
            合作电话：13146991091
            &nbsp;&nbsp;&nbsp;
            其它合作：010-64977537-16
            &nbsp;&nbsp;&nbsp;
            传真：010-64899279
        </p>
        <p>餐厅地址：<% =CokeShow.Setup(31,0) %>   <a class="fontgreen" href="/Details/DetailsInformation.Welcome?CokeMark=JPHPEPR" target="_blank">&gt; 查看餐厅官方地图 <img src="/images/ico/small/map_magnify.png" /></a> <a class="fontgreen" href="/Details/DetailsInformation.Welcome?CokeMark=JPHQRRP" target="_blank">&gt; 查看餐厅乘车路线  <img src="/images/ico/small/map_magnify.png" /></a></p>
        <p>北京祝丰融餐饮有限责任公司(<a href="http://www.chixinbugai.me/">痴心不改餐厅</a>) 版权所有.&nbsp;<a href="http://www.miibeian.gov.cn/" target="_blank">京备号:10052760</a>&nbsp;&nbsp;<a href="http://www.miitbeian.gov.cn" target="_blank">京备号:10052760</a>&nbsp;餐厅官方电邮：<a href="mailto:supper@chixinbugai.me">supper@chixinbugai.me</a></p>
        <p>COPYRIGHT&copy; 2007-2010 BY 痴心不改餐厅 PR CONSULTING CN.LTD(C) ALL RIGHTS RESERVED.&nbsp;&nbsp;&nbsp;<span>支持伙伴:<a href="http://可乐秀.中国/" target="_blank">可乐秀.中国</a></span></p>
        <p style="color: #CCC;">我们将持之以恒，让最新菜品第一时间出现在痴心不改餐厅的官方网站，更完美的服务亚运村社区.</p>
		
<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-17069773-1']);
  _gaq.push(['_setDomainName', 'none']);
  _gaq.push(['_setAllowLinker', true]);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
<p style="display:none;">
<script language="javascript" type="text/javascript" src="http://js.users.51.la/3914408.js"></script>
</p>