//当户名填写完毕时，调用此函数检验用户名是否唯一.
function accountNameOnChange() {
	//getElementById也很OK，现在退役。var userName = document.getElementById("userName").value.
	//var userName = dijit.byId("userName").getValue().
	var userName = dijit.byId("username").attr('value');	//新版本Dojo中，用attr('value')取代了getValue()。
	if (userName == "") {
		console.log("username is empty.");
		return;		//截断.
	} else {
		console.log(userName);
		//return;	它会从此处截断整个函数.
	}
	
	//开始向服务器发送数据，使用Dojo.xhrGet函数.
	dojo.xhrGet({
				url:		"/ONCEFOREVER/Account.Services.Public.asp?ServicesAction=CheckAccountName&username=" + userName + "&CokeShow=" + Math.random(),
				handleAs:	"json",
				handle:		accountNameValidationHandler		//处理回调.
				});
				
}


//向服务器发送数据的回调函数.
function accountNameValidationHandler(response) {
	//此处的变量response直接接收上一个function中的变量response过来用！这应该就是闭包！
	var userName = dijit.byId("username").attr('value');
	//Clear any error messages that may have been displayed.
	//令dijit控件中的username的专用属性displayMessage提示一空消息，相当于清除消息！
	dijit.byId("username").displayMessage();
	
	//反馈环节一定要全权放在Handler函数中来处决！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！！Wangliang
	var patrn=/^[0-9a-zA-Z]+([0-9a-zA-Z]|_|\.|-)+[0-9a-zA-Z]+@(([0-9a-zA-Z]+\.)|([0-9a-zA-Z]+-))+[0-9a-zA-Z]+$/;
	if (!patrn.test(userName)) {
		//return false.
		//先判断是否符合正则表达式，如果不符合，则连Ajax都不用发送了.直接提示错误信息.
		var errorMessage = "请填写正确的Email电子邮件格式，例如：yourname6179@qq.com<img src='/images/ico/small/emotion_suprised.png' />";
		//Display error message as tooltip next to field.
		//令dijit控件中的username的专用属性displayMessage提示消息！
		dijit.byId("username").displayMessage(errorMessage);
	} else {
		//return true.
		
		
		//response就是返回的JSON对象.
		if (response.theResult_true_false == "false") {	//读取response引用获取到的JSON对象，的valid属性
			//当判断为false时
			var errorMessage = response.theAllInformation;
			//Display error message as tooltip next to field.
			//令dijit控件中的username的专用属性displayMessage提示消息！
			dijit.byId("username").displayMessage(errorMessage);
		} else {	
			//判断字数不会在控件dijit中默认报错的情况下，再提示恭喜可以注册.
			var errorMessage = response.theAllInformation;
			dijit.byId("username").displayMessage(errorMessage);
		}
		
		//return true.
	}
}
