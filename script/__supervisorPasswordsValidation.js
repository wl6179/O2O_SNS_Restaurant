//当户名填写完毕第二个确认密码时，调用此函数检验密码和确认密码是否一致.
function supervisorPasswordsOnChange() {
	var _password = dijit.byId("password").attr('value');
	var _repassword = dijit.byId("repassword").attr('value');
	if (_password == "" || _repassword == "") {
		console.log("password or repassword is empty.");
		return;		//截断.
	} else {
		console.log(_password + ',' + _repassword);
		//return;	它会从此处截断整个函数.
		
		//dijit.byId("username").displayMessage('aaa');
	}
	
	//验证对比.
	dijit.byId("password").displayMessage();
	dijit.byId("repassword").displayMessage();
	
	if (_password != _repassword) {
		var errorMessage = "<img src='/images/del.gif' />&nbsp;密码和确认密码不匹配，请重新输入！";
		dijit.byId("password").displayMessage(errorMessage);
	}
}
