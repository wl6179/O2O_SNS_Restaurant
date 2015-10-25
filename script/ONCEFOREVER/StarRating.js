var xmlHttp;    //用于保存XMLHttpRequest对象的全局变量

//用于创建XMLHttpRequest对象
function createXmlHttp() {
    //根据window.XMLHttpRequest对象是否存在使用不同的创建方式
    if (window.XMLHttpRequest) {
       xmlHttp = new XMLHttpRequest();                  //FireFox、Opera等浏览器支持的创建方式
    } else {
       xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");//IE浏览器支持的创建方式
    }
}

//获取产品参数
function vote(id, diamonds) {
    //先减去当前的星级.
	dojo.style(dojo.byId("CurrentRating"), {width: "0"});
	//在此处调用源页面定义的函数：
	//alert(dojo.byId("yuanyemian_opinionvalue").value);
	if ( typeof(dojo.byId("yuanyemian_opinionvalue"))=="object" ) {	//如果存在这个DOM，那么代表源页面定义了控制节点，用以控制当前函数要如何显示！
		//alert(dojo.byId("yuanyemian_opinionvalue").value);	ok
		if (dojo.byId("yuanyemian_opinionvalue").value == '1') {
			//控制当前点击星星函数要弹出提示框，说明一些信息给顾客.
			//创建ShowDialog的widght.
			showRatingInformation();
			
		} else {
			
		}
	}
	
//	createXmlHttp();                            //创建XmlHttpRequest对象
//    xmlHttp.onreadystatechange = showRating;    //设置回调函数
//    xmlHttp.open("GET", "/fcktemplates.xml?id=" + id + "&diamonds=" + diamonds, true);
//    xmlHttp.send(null);
	
	//定位所选择的星级.
	//dojo.style(dojo.byId("CurrentRating"),"class") = "diamond" + diamonds + "_hover";
	//dojo.style(c1, {display:"block"});
	dojo.style(dojo.byId("CurrentRating"), {width: ""+ (20 * diamonds) +"px"});		//显示顾客点击的星级(固定).
	var strDiamonds="" + diamonds + "";		//点击的星级的数字.并转化为字符串格式，好做操作判断.
	//操作判断点击的星级，是属于什么提示文字，如：1星级-不好吃.
	var strResultForClick="";
	if (strDiamonds == "1") {
		strResultForClick = "难吃";
	} else if (strDiamonds == "2") {
		strResultForClick = "不好吃";
	} else if (strDiamonds == "3") {
		strResultForClick = "不错哦";
	} else if (strDiamonds == "4") {
		strResultForClick = "挺好吃";
	} else if (strDiamonds == "5") {
		strResultForClick = "超级好吃哇，荐";
	}
	console.log("diamond" + diamonds + "_hover," + strResultForClick);
	//记下hidden参数.
	//dojo.byId("theStarRatingForChineseDishInformation").value == "" + diamonds + "";
	dojo.byId("theStarRatingForChineseDishInformation").setAttribute("value", "" + diamonds + "");
	//输出提示.
	dojo.byId("showRatingMessage").innerHTML = "<span style='color:orange;'>您当前为菜品点评的星级级别是：<strong style='color:#FF3300;'>" + diamonds + " 星级-" + strResultForClick + "。</strong> </span><br />";
}

//显示投票结果
function showRating() {
    if (xmlHttp.readyState == 4) {
        //var rating = eval("("+xmlHttp.responseText+")");    //解析服务器反馈信息（JSON格式）
        //将信息写入页面
        //document.getElementById("rating-" + rating.id).innerHTML = "投票人数：" + rating.totaltimes + "，钻石总数：" + rating.totaldiamonds;
		
		//dojo.byId("showRatingMessage").innerHTML = "投票人数：XX，钻石总数：XXX";
    }
}

//显示弹出对话框——评星权限说明：
function showRatingInformation() {
    ShowDialog('<span style=color:black;>提示信息</span>','','width:300px;height:360px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><ul class="rating"><li class="current-rating" style="width:100px;" id="CurrentRating"></li></ul>点亮评星级权限说明：<span style="color:red;"></span><br /><br />↑ 评星级选项，星级已经开发给所有的网站会员参与评星哦！<br /><br />↑ 如果您已经是痴心不改餐厅的老顾客，并且持有会员卡，请将会员卡上的卡号记下并绑定到您的网站会员帐号中，即可享受乐趣多多哦！<br /><br />详情如下：<br /><br /><a class="fontgreen" href="/Details/DetailsInformation.Welcome?CokeMark=JPKOCNH" target="_blank"> 如何绑定会员卡到我的网站会员帐号下?</a><br /><br /><a class="fontgreen" href="/Details/DetailsInformation.Welcome?CokeMark=JPIRFMN" target="_blank"> 如何办理餐厅会员卡?</a><br /><br /><a class="fontgreen" href="/Details/DetailsInformation.Welcome?CokeMark=JPKORNJ" target="_blank"> 为何绑定不上我的已有卡号?</a></span></div>');
}