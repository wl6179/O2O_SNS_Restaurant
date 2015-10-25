	  
	    <!--<a href="">
		  <span class="zjllimg"><img src="/images/cpimg/zjllimg01.jpg" width="45" height="45" /></span>
		  <span class="zjlltitle">菜品名称：<font class="fontred_a">痴心不改菜品</font></span>
	    <span class="zjllxjpl">星级评论：<img src="/images/xx_01.gif" width="13" height="12" /><img src="/images/xx_01.gif" width="13" height="12" /><img src="/images/xx_01.gif" width="13" height="12" /><img src="/images/xx_01.gif" width="13" height="12" /><img src="/images/xx_01.gif" width="13" height="12" /></span>		
		</a>-->	  
	    
        		
			  

      
<%
'遍历Top7个菜品历史浏览
'注:需要嵌入前台浏览历史类文件.< !--#include file="system/foreground_ToBrowseRecordsClass.asp"-- >
'定义内部新变量进行内部操作.
Dim objToBrowseRecords
Dim strToBrowseRecords

'初始化赋值.
strToBrowseRecords=""
Set objToBrowseRecords = New ToBrowseRecordsClass
'获得Session("ToBrowseRecords")字典.

'判断有各种效性.
'''//objToBrowseRecords.AddProduct 18, "2微软键盘套装MS-DSN$$125.00$$/images/logo.jpg"

'-----------------Go Begin
'列出菜品浏览记录列表.
Dim key_ToBrowseRecords,arrayItem_ToBrowseRecords
For Each key_ToBrowseRecords In Session("ToBrowseRecords").Keys
	
	'创建字典的模式 = ToBrowseRecords.AddProduct(8, "微软键盘套装MS-DSN$$125.00$$/images/NoPic.gif").
	strToBrowseRecords=	"<a href=""/ChineseDish/ChineseDishInformation.Welcome?CokeMark="& CokeShow.AddCode_Num( key_ToBrowseRecords ) &""" target=""_blank""><span class=""zjllimg""><img src="""& objToBrowseRecords.GetProductValue(CokeShow.CokeClng(key_ToBrowseRecords),"Photo") &""" width=""45"" height=""45"" onerror='this.src=""/images/NoPic.png""' /></span>"&_
											"<span class=""zjlltitle"">菜名：<font class=""fontred_a"" title="""& objToBrowseRecords.GetProductValue(CokeShow.CokeClng(key_ToBrowseRecords),"ProductName") &""">"& objToBrowseRecords.GetProductValue(CokeShow.CokeClng(key_ToBrowseRecords),"ProductName") &"</font></span>"&_
											"<span class=""zjllxjpl""><ul class=""rating"" style=""margin-top:3px; margin-left:0px;""><li class=""current-rating"" style=""width:"& Coke.ShowProductStarRating_Num(key_ToBrowseRecords) * 20 + 1 &"px;""></li></ul>"&_
											"</span></a>" & strToBrowseRecords
											'"& FormatCurrency( objToBrowseRecords.GetProductValue(CokeShow.CokeClng(key_ToBrowseRecords),"UnitPrice"),2 ) &"
	
	
Next
'-----------------Go End
		
'终结化操作.
Set objToBrowseRecords = Nothing

Response.Write strToBrowseRecords
%>
