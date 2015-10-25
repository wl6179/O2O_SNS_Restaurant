//define function to be called when username is entered.
//当户名填写完毕时，调用此函数检验用户名是否唯一.
function supervisorNameOnChange() {
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
				url:		"__supervisorNameValidation.asp?username=" + userName,
				handleAs:	"json",
				handle:		supervisorNameValidationHandler		//处理回调.
				});
				
}


//向服务器发送数据的回调函数.
function supervisorNameValidationHandler(response) {
	//此处的变量response直接接收上一个function中的变量response过来用！这应该就是闭包！
	
	//Clear any error messages that may have been displayed.
	//令dijit控件中的username的专用属性displayMessage提示一空消息，相当于清除消息！
	dijit.byId("username").displayMessage();
	
	if (! response.valid) {	//读取response引用获取到的JSON对象，的valid属性
		//当判断为false时
		var errorMessage = "<img src='/images/del.gif' />&nbsp;此帐号名已存在...请输入其它帐号名字！";
		//Display error message as tooltip next to field.
		//令dijit控件中的username的专用属性displayMessage提示消息！
		dijit.byId("username").displayMessage(errorMessage);
	}
}
