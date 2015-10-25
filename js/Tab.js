/*
	对象有tabhead,tabcontent
	tabhead属性有 tabcontentid(必须),activeclass(必须),deactiveclass(必须), groupname(必须),default=default(可选),hoverclass(可选)
	tabcontentid 表示tabhead所要关联的内容Div或其它容器元素
	activeclass 表示未激活时的样式
	deactiveclass 表示激活时的样式
	groupname 表示组名
	default表示页面加载后激活签
	hoverclass 鼠标悬停样式
*/
$(document).ready(function(){
	$("*[@tabcontentid]").click(function(){
		var ctl = $("#"+$(this).attr("tabcontentid"));
		var groupname = $(this).attr("groupname");
		$("*[@groupname="+ groupname +"]").each(function(){
			var ctls = $("#"+$(this).attr("tabcontentid"));
			if(ctls!=ctl)
			{
				ctls.css("display","none");			
				$(this).attr("class",$(this).attr("deactiveclass"));
			}
		});
		if (ctl)
		{
			ctl.css("display","block");			
			$(this).attr("class",$(this).attr("activeclass"));
		}
	});
	
	$("*[@tabcontentid]").each(function(){
		if($(this).attr("hoverclass"))
		{
			$(this).mouseover(function(){
				$(this).addClass($(this).attr("hoverclass"));
			});
			$(this).mouseout(function(){
				$(this).removeClass($(this).attr("hoverclass"));
			});
		}
		if($(this).attr("default"))
		{
			$(this).click();
		}
	});
});
