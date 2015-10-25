		<%
		'
		'公有的内嵌文件，最右栏的最新提示模块.
		'
		%>
		
		<%
		'
		'自动检测显示屏宽度并处理：
		'当屏幕太小小于等于1024、还有新提示都完成时，自动处理消除此最右栏.
		'
		Dim rsWork,sqlWork
		Dim isHaveWork
		'默认为无事务需要处理.
		isHaveWork = False
		
		'暂时省略处理过程...
		
		'暂时设置为有事务需要处理，需要显示右栏.
		If isNumeric(Session("isHaveWork_supervisor")) Then
			If CokeShow.CokeClng(Session("isHaveWork_supervisor"))=1 Then
				isHaveWork = False'True
			ElseIf CokeShow.CokeClng(Session("isHaveWork_supervisor"))=0 Then
				isHaveWork = False
			End If
			
		Else
			isHaveWork = False'True
		End If
		%>