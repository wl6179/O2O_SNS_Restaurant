	<script type="text/javascript">
	//dojo针对显示器宽度进行控制.
	function setScreen() {
		//如果右栏任务已全部完成，那么用程序将此处设为1（坚决消掉右栏）;
	<%
	'如果有事务需要处理，则显示！
	If isHaveWork=True Then
	%>
		var noRight = 0;
	<%
	Else
	'否则消失其右栏.
	%>
		var noRight = 1;
	<%
	End If
	%>
		var W = Number(screen.width);
		var H = Number(screen.height);
		console.log(W, H);
		//获取要控制的元素class="mainInfo"的div.
		var mainInfo = dojo.query(".mainInfo");
		var rightInfo = dojo.query(".rightInfo");
		
		if (typeof(mainInfo)=='object') {
			console.log("ok! typeof(mainInfo)=='object'");
			//1.
			if (W <= 1024) {
					//设置中栏.加宽auto.
					dojo.style(mainInfo,{
						width:"auto",
						marginRight:"0px"
					});
					//如果出现右栏，消除最右栏.
					if (typeof(rightInfo)=='object') {
						dojo.style(rightInfo,{
							display:"none"
						});
					}
			}
			//2.
			if (W > 1024) {
				//是否坚决消除右栏.
				if (noRight == 1) {
					console.log("ok! noRight == 1");
					//设置中栏.加宽auto.
					mainInfo.style({
						"width" : "auto",
						
						"marginRight" : "0px"
					});
					

					//如果出现右栏，消除最右栏.
					if (typeof(rightInfo)=='object') {
						rightInfo.style({
							"display":"none"
						});
					}
				}
			}

		}
		else {
			console.log("alert! typeof(mainInfo)!=='object'");
		}
	}
	
	
	//运行屏幕设置适应函数
   dojo.addOnLoad(function(){
	   setScreen();
	});
	</script>