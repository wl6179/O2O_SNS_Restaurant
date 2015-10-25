/*
	������tabhead,tabcontent
	tabhead������ tabcontentid(����),activeclass(����),deactiveclass(����), groupname(����),default=default(��ѡ),hoverclass(��ѡ)
	tabcontentid ��ʾtabhead��Ҫ����������Div����������Ԫ��
	activeclass ��ʾδ����ʱ����ʽ
	deactiveclass ��ʾ����ʱ����ʽ
	groupname ��ʾ����
	default��ʾҳ����غ󼤻�ǩ
	hoverclass �����ͣ��ʽ
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
