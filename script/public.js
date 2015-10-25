function addfavorite(theURL,theTitle) {
	if (document.all) {
		window.external.addFavorite(theURL);
	}
	else if (window.sidebar) {
		window.sidebar.addPanel(theTitle, theURL, "aaa");
	}
} 
function addBookmark(theURL,theTitle) {
	if (window.sidebar) { 
		window.sidebar.addPanel(theTitle, theURL,""); 
	} else if( document.all ) {
		window.external.AddFavorite( theURL, theTitle);
	} else if( window.opera && window.print ) {
		return true;
	}
}

function processForm(theFormElement) {
//Begin	使用上Editor时再用上.
//	var readme = dijit.byId("readme");
//	console.log(readme.attr("value"));
//	//alert(comments.attr("value"));
//	dojo.byId("readme_hidden").value = readme.attr("value");
//End	使用上Editor时再用上.
	
	//Re-validate form fields.	检测一遍表单中的所有元素的合法性。
	var custForm = dijit.byId(theFormElement);
	var firstInvalidWidget = null;
	//遍历表单的(Descendant子孙)，变量widget就是一个一个的子孙元素。
	//*对数组的每个元素,callback 都返回true, 则dojo.every=true. 也就是说只要有一个false，则dojo.every=false！*
	dojo.every(custForm.getDescendants(),function(widget){
		if (widget.isValid()==false) {
			firstInvalidWidget = widget;	//如果dojo.every=false，此句就执行？！
		}
		//firstInvalidWidget = widget;
		
		console.log("widget: ", widget);
		console.log("isValid: ", widget.isValid);
		console.log("isValid(): ", widget.isValid());
		console.log("firstInvalidWidget: ", firstInvalidWidget);
		return !widget.isValid || widget.isValid();	//判断依据isValid()！有false时，立刻停止循环！那么为true时，继续执行循环，直到找到错误输入框位置！
//		if (typeof(firstInvalidWidget)=="object") {
//			return false;
//		} else {
//			return true;
//		}
	});
	
	
	//alert("firstInvalidWidget:" + firstInvalidWidget);
	//Place focus on first invalid field.	将焦点放在无效的区域上。
	//如果第一个子孙元素不为空，
	if (firstInvalidWidget != null) {
		//set focus to first field with an error.
		firstInvalidWidget.focus();
		firstInvalidWidget.displayMessage('您的输入可能尚有错误，请检查！');
		//return false;		当用最新版的onSubmit方法时，才能多此一举用return false;来拦截表单继续提交动作！用老版execute就没事了。
	} else {
		custForm.submit();	//这个不是简单的JavaScript的submit方法，是Dojo的方法。
		//return true;
	}
	//If all fields are valid then submit form.	如果都合法，则提交表单(上边)。
}

//菜品点评Ajax提交函数POST.
//例子：execute="processFormAjax('RemarkOnForm','handleFunctionName1','errorFunctionName1','loadingFunctionName1','theSubmitButton')"
//然后立刻接着定义这3个处理函数分别处理不同的具体ajax情况.
function processFormAjax() {
	
	//扩展使用——用于Ajax表单提交并且反馈.Begin
	var button = dijit.byId("theSubmitButton");
	dojo.connect(button, "onClick", function(event) {
		//Stop the submit event since we want to control form submission.
		//dojo.stopEvent(event);//备用
		event.preventDefault();
		event.stopPropagation();
		//The parameters to pass to xhrPost, the form, how to handle it, and the callbacks.
		//Note that there isn't a url passed.  xhrPost will extract the url to call from the form's
		//'action' attribute.  You could also leave off the action attribute and set the url of the xhrPost object
		//either should work.
		var xhrArgs = {
                form: dojo.byId("RemarkOnForm"),
                handleAs: "json",
                load: function(data) {
                    dojo.byId("response").innerHTML = "反馈消息：" + data.theAllInformation;
					
					//在成功添加记录时，弹出提示框！！
					if (data.theResult_true_false == 'true') {
						if (data.intJifenFollowUp != '0') {
							var strJifenDesc = "<br /><img src=/images/ico/small/coins_add.png />您已成功获得了点评菜品送出的"+ data.strJifenFollowUp + data.intJifenFollowUp +"积分哦~";
						} else {
							var strJifenDesc = "<br /><img src=/images/ico/small/coins_add.png />您的第一次点评已得到过积分了，此次点评不计入积分哦:)";
						}
						ShowDialog('<span style=color:black;>成功点评提示信息</span>','','width:300px;height:200px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><img src=/images/ico/emotion_happy.png />恭喜您，点评成功！'+ strJifenDesc +'<br /><br />刷新查看您的新点评：<a class="button_img77" style="display:inline-block; color:green;" href="javascript:return false;" onclick="location.reload();">刷新查看&gt;&gt;</a></span></div>');
						
					}
					//在成功添加记录时，弹出提示框！！
                },
                error: function(error) {
                    //We'll 404 in the demo, but that's okay.  We don't have a 'postIt' service on the
                    //docs server.
                    dojo.byId("response").innerHTML = "<img src='/images/ico/small/emotion_suprised.png' style='float:none;' />通讯失效，请您重新访问本页面并重新尝试操作！" + error.message;		//可能是网络太慢、service反馈为空白无字页（反馈无效）等原因.
                }
		}
		//Call the asynchronous xhrPost
		dojo.byId("response").innerHTML = "<img src='/images/ajax-loader.gif' style='float:none;' /> Loading ..."
		//发送.
		var deferred = dojo.xhrPost(xhrArgs);
	});
	//扩展使用——用于Ajax表单提交并且反馈.End

}
dojo.addOnLoad(processFormAjax);


//菜品加入收藏夹Ajax提交函数GET.
function addFavoriteAjax(trID) {
	//thisElementObject.preventDefault();
	//thisElementObject.stopPropagation();

	//请求服务.
	dojo.xhrGet({
		url: "/ONCEFOREVER/Account.Services.Private.asp",
		content: { query:'True', ServicesAction:'addFavorite', id:trID, CokeShow:Math.random() },
		timeout: 10000,
		handleAs: "json",
		//handle: supervisorNameValidationHandler,	//处理回调.
		load: function(data) {
			if (data.theResult_true_false == 'true') {
				//dojo.byId("response").innerHTML = "反馈消息：" + data.theAllInformation;
				//在成功添加记录时，弹出提示框.Begin
				//操作成功时，提示！
				ShowDialog('<span style=color:black;>成功收藏提示信息</span>','','width:300px;height:200px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><img src=/images/ico/emotion_happy.png />'+ data.theAllInformation +'<br /><img src=/images/ico/small/coins_add.png />同时您也成功获得网站主办的鼓励小活动-收藏菜品的1积分哦~<br /></span></div>');
				
				
				//在成功添加记录时，弹出提示框.End
				
				//发布成功消息！
				//dojo.publish("xhrAddFavoriteScc", [{
				//	 message: "<span style='font-size:12px;'><img src=/images/png-0094.png width='23' /> 痴心不改提示您:<br /><br />" + response.message + "</span>",	
				//	 type: "fatal",
				//	 duration: 0
				//}]
				//);
			} else {
				//失败时.警报错误，要求重试操作.
				//发布失败提示消息！
				//dojo.publish("xhrAddFavoriteError", [{
				//	 message: "<span style='font-size:12px;'><img src=/images/no.png width='23' /> 痴心不改提示您:<br /><br />" + response.message + "</span>",	
				//	 type: "error",
				//	 duration: 0
				//}]
				//);
				
				//有其它尚未完成信息时，弹出提示.
				ShowDialog('<span style=color:black;>提示信息</span>','','width:300px;height:200px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><img src=/images/ico/emotion_wink.png />'+ data.theAllInformation +'<br /></span></div>');
				
			}
						
		},
		
		error: function(errorObject) {
			//一个Toaster将捕获这个错误并显示它.
			//dojo.publish("xhrAddFavoriteError", [{
			//	 message: "<span style='font-size:12px;'><img src=/images/no.png width='23' /> 痴心不改提示您:<br /><br />" + "由于您尚未登录，所以此次加入收藏失败，请您先登录会员中心！</span>",		//将error对象 传给小部件的message发布参数里，进行传送!
			//	 type: "error",
			//	 duration: 0
			//}]
			//);
			//return error;
			
			//有失误时，弹出提示.
			ShowDialog('<span style=color:black;>提示信息</span>','','width:300px;height:200px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><img src=/images/ico/emotion_unhappy.png />通讯失效，请您重新访问本页面并重新尝试操作！<br /><br />具体技术描述信息：' + errorObject.message +'<br /></span></div>');
			
		}
		
	});

}

//推荐朋友连环招01——GET传送弹出对话框检测请求，通过后，并访问推荐朋友表单页.[Ajax I]
//会员将菜品推荐给朋友之Ajax弹出对话框函数GET.（未登录提示先登陆注册对话框，完成后才回来继续点击操作；已登录则直接弹出推荐朋友对话框，完成发送。）
function addAccount_TuijianPengyou(trID) {
	//请求服务.
	dojo.xhrGet({
		url: "/ONCEFOREVER/Account.Services.Private.asp",
		content: { query:'True', ServicesAction:'addAccount_TuijianPengyou_CheckReady', id:trID, CokeShow:Math.random() },
		timeout: 10000,
		handleAs: "json",
		load: function(data) {
			if (data.theResult_true_false == 'true') {
				//操作成功时，提示！
				ShowDialog('<img src=/images/ico/small/group.png /> <span style=color:black;>请填写您推荐朋友的Email信息</span>','/ONCEFOREVER/addAccount_TuijianPengyou.Welcome?id=' + trID, 'width:460px;height:500px;','');
				//processFormAjax_addAccount_TuijianPengyou();	//关联对话框中的提交按钮事件为Ajax的POST提交函数.***
			} else {
				//有其它尚未完成信息时，弹出提示.
				ShowDialog('<span style=color:black;>提示信息</span>','','width:300px;height:200px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><img src=/images/ico/emotion_wink.png />'+ data.theAllInformation +'<br /></span></div>');
			}
					
		},
		
		error: function(errorObject) {
			//通讯失败时，弹出提示.
			ShowDialog('<span style=color:black;>提示信息</span>','','width:300px;height:200px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><img src=/images/ico/emotion_unhappy.png />通讯失效，请您重新访问本页面并重新尝试操作！<br /><br />具体技术描述信息：' + errorObject.message +'<br /></span></div>');
		}
		
	});

}
//推荐朋友连环招02——通过后，并访问推荐朋友表单页时.提交表单，开始POST提交Ajax到WebServices处理.[Ajax II]
//推荐朋友Ajax提交函数POST.
function processFormAjax_addAccount_TuijianPengyou() {
	
//	var button = dijit.byId("submit886");
//	dojo.connect(button, "onClick", function(event) {
//		event.preventDefault();
//		event.stopPropagation();
		var xhrArgs = {
                form: dojo.byId("tuijianpengyou"),
                handleAs: "json",
                load: function(data) {
                    
					if (data.theResult_true_false == 'true') {
						dojo.byId("responseDialog").innerHTML = "" + data.theAllInformation + data.strJifenFollowUp;
					} else {
						dojo.byId("responseDialog").innerHTML = "" + data.theAllInformation;
					}
					
                },
                error: function(error) {
                    dojo.byId("responseDialog").innerHTML = "<img src='/images/ico/small/emotion_suprised.png' style='float:none;' />通讯失效，请您重新访问本页面并重新尝试操作！";	//+ error.message可能是网络太慢、service反馈为空白无字页（反馈无效）等原因.
                }
		}
		dojo.byId("responseDialog").innerHTML = "<img src='/images/ajax-loader.gif' style='float:none;' /> Loading ..."
		//发送.
		var deferred = dojo.xhrPost(xhrArgs);
//	});

}
//dojo.addOnLoad(processFormAjax_addAccount_TuijianPengyou);//




//调查问卷Ajax提交函数POST.
//然后立刻接着定义这3个处理函数分别处理不同的具体ajax情况.
function processFormAjax_QuestionnairesForm() {
	
	var button = dijit.byId("theSubmitButton_QuestionnairesForm");
	dojo.connect(button, "onClick", function(event) {
		event.preventDefault();
		event.stopPropagation();
		
		var xhrArgs = {
                form: dojo.byId("QuestionnairesForm"),
                handleAs: "json",
                load: function(data) {
                    dojo.byId("response").innerHTML = "" + data.theAllInformation;
					
					//在成功添加记录时，弹出提示框！！
					//if (data.theResult_true_false == 'true') {
						//if (data.xxxxxx != '0') {
						//	var strDesc = "<br /><img src=/images/ico/small/coins_add.png />您已成功提交调查问卷";
						//} else {
						//	var strDesc = "<br /><img src=/images/ico/small/coins_add.png />您已经成功提交过调查问卷了";
						//}
						//ShowDialog('<span style=color:black;>成功提交调查问卷反馈信息</span>','','width:300px;height:200px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><img src=/images/ico/emotion_happy.png />感谢您的支持！'+ strDesc +'<br /></span></div>');
						
					//}
					//在成功添加记录时，弹出提示框！！
                },
                error: function(error) {
                    dojo.byId("response").innerHTML = "<img src='/images/ico/small/emotion_suprised.png' style='float:none;' />通讯失效，请您重新访问本页面并重新尝试操作！" + error.message;		//可能是网络太慢、service反馈为空白无字页（反馈无效）等原因.
                }
		}
		dojo.byId("response").innerHTML = "<img src='/images/ajax-loader.gif' style='float:none;' /> Loading ..."
		//发送.
		var deferred = dojo.xhrPost(xhrArgs);
	});


}
dojo.addOnLoad(processFormAjax_QuestionnairesForm);



//GET发送请求并弹出对话框，通过后登录等验证后，提交数据给services.
//会员兑换礼品券之Ajax弹出对话框函数GET.（未登录提示先登陆注册对话框，完成后才回来继续点击操作；已登录则直接传送请求并弹出信息提示对话框，完成发送。）
function addAccount_GiftCertificated(trID) {
var confirm_go = confirm( "确定要兑换此礼品券吗？( 领取礼品券后将扣除您的相应积分数 )" );
	if (confirm_go == true) {	//IF GO
		
	//请求服务.
	dojo.xhrGet({
		url: "/ONCEFOREVER/Account.Services.Private.asp",
		content: { query:'True', ServicesAction:'addAccount_GiftCertificated', id:trID, CokeShow:Math.random() },
		timeout: 10000,
		handleAs: "json",
		load: function(data) {
			if (data.theResult_true_false == 'true') {
				//操作成功时，提示！
				ShowDialog('<img src=/images/ico/small/group.png /> <span style=color:black;>成功领取礼品券提示信息</span>','','width:300px;height:230px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><img src=/images/ico/emotion_happy.png />'+ data.theAllInformation +'<br /></span></div>');
				
			} else {
				//有其它尚未完成信息时，弹出提示.
				ShowDialog('<span style=color:black;>提示信息</span>','','width:300px;height:230px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><img src=/images/ico/emotion_wink.png />'+ data.theAllInformation +'<br /></span></div>');
				
			}
					
		},
		
		error: function(errorObject) {
			//通讯失败时，弹出提示.
			ShowDialog('<span style=color:black;>提示信息</span>','','width:300px;height:230px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><img src=/images/ico/emotion_unhappy.png />通讯失效，请您重新访问本页面并重新尝试操作！<br /><br />具体技术描述信息：' + errorObject.message +'<br /></span></div>');
		}
		
	});
}								//IF GO
}



//提交站内信的Ajax提交函数POST.
function processFormAjax_SendMessageForm() {
	
	var button = dijit.byId("theSubmitButton_SendMessageForm");
	dojo.connect(button, "onClick", function(event) {
		event.preventDefault();
		event.stopPropagation();
		
		var xhrArgs = {
				form: dojo.byId("SendMessageForm"),
				handleAs: "json",
				load: function(data) {
					dojo.byId("response").innerHTML = "" + data.theAllInformation;
					
					//在成功添加记录时，弹出提示框！！
					//if (data.theResult_true_false == 'true') {
						//if (data.xxxxxx != '0') {
						//	var strDesc = "<br /><img src=/images/ico/small/coins_add.png />您已成功提交调查问卷";
						//} else {
						//	var strDesc = "<br /><img src=/images/ico/small/coins_add.png />您已经成功提交过调查问卷了";
						//}
						//ShowDialog('<span style=color:black;>成功提交调查问卷反馈信息</span>','','width:300px;height:200px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><img src=/images/ico/emotion_happy.png />感谢您的支持！'+ strDesc +'<br /></span></div>');
						
					//}
					//在成功添加记录时，弹出提示框！！
				},
				error: function(error) {
					dojo.byId("response").innerHTML = "<img src='/images/ico/small/emotion_suprised.png' style='float:none;' />通讯失效，请您重新访问本页面并重新尝试操作！" + error.message;		//可能是网络太慢、service反馈为空白无字页（反馈无效）等原因.
				}
		}
		dojo.byId("response").innerHTML = "<img src='/images/ajax-loader.gif' style='float:none;' /> Loading ..."
		//发送.
		var deferred = dojo.xhrPost(xhrArgs);
	});


}
dojo.addOnLoad(processFormAjax_SendMessageForm);


//提交申请VIP卡号的Ajax提交函数POST.
function processFormAjax_BindingMyVIPCardForm() {
	
	var button = dijit.byId("theSubmitButton_BindingMyVIPCardForm");
	dojo.connect(button, "onClick", function(event) {
		event.preventDefault();
		event.stopPropagation();
		
		var xhrArgs = {
				form: dojo.byId("BindingMyVIPCardForm"),
				handleAs: "json",
				load: function(data) {
					dojo.byId("response").innerHTML = "" + data.theAllInformation;
				},
				error: function(error) {
					dojo.byId("response").innerHTML = "<img src='/images/ico/small/emotion_suprised.png' style='float:none;' />通讯失效，请您重新访问本页面并重新尝试操作！" + error.message;		//可能是网络太慢、service反馈为空白无字页（反馈无效）等原因.
				}
		}
		dojo.byId("response").innerHTML = "<img src='/images/ajax-loader.gif' style='float:none;' /> Loading ..."
		//发送.
		var deferred = dojo.xhrPost(xhrArgs);
	});


}
dojo.addOnLoad(processFormAjax_BindingMyVIPCardForm);



//修改密码的Ajax提交函数POST.
function processFormAjax_ChangePassword() {
	
	var button = dijit.byId("theSubmitButton_ChangePassword");
	dojo.connect(button, "onClick", function(event) {
		event.preventDefault();
		event.stopPropagation();
		
		var xhrArgs = {
				form: dojo.byId("ChangePassword"),
				handleAs: "json",
				load: function(data) {
					dojo.byId("response").innerHTML = "" + data.theAllInformation;
				},
				error: function(error) {
					dojo.byId("response").innerHTML = "<img src='/images/ico/small/emotion_suprised.png' style='float:none;' />通讯失效，请您重新访问本页面并重新尝试操作！" + error.message;		//可能是网络太慢、service反馈为空白无字页（反馈无效）等原因.
				}
		}
		dojo.byId("response").innerHTML = "<img src='/images/ajax-loader.gif' style='float:none;' /> Loading ..."
		//发送.
		var deferred = dojo.xhrPost(xhrArgs);
	});


}
dojo.addOnLoad(processFormAjax_ChangePassword);



//修改个人资料的Ajax提交函数POST.
function processFormAjax_PersonalInformationForm() {
	
	var button = dijit.byId("theSubmitButton_PersonalInformationForm");
	dojo.connect(button, "onClick", function(event) {
		event.preventDefault();
		event.stopPropagation();
		
		var xhrArgs = {
				form: dojo.byId("PersonalInformationForm"),
				handleAs: "json",
				load: function(data) {
					dojo.byId("response").innerHTML = "" + data.theAllInformation;
				},
				error: function(error) {
					dojo.byId("response").innerHTML = "<img src='/images/ico/small/emotion_suprised.png' style='float:none;' />通讯失效，请您重新访问本页面并重新尝试操作！" + error.message;		//可能是网络太慢、service反馈为空白无字页（反馈无效）等原因.
				}
		}
		dojo.byId("response").innerHTML = "<img src='/images/ajax-loader.gif' style='float:none;' /> Loading ..."
		//发送.
		var deferred = dojo.xhrPost(xhrArgs);
	});


}
dojo.addOnLoad(processFormAjax_PersonalInformationForm);

//将DOM新元素加入到目标元素的后边e函数.
/*
此函数用到了以下DOM方法和属性.
parentNode属性
lastChild属性
appendChild()方法
insertBefore()方法
nextSibling属性

备用方法和属性.
createElement()方法
setAttribute()方法
createTextNode()方法
*/
function insertAfter(newElement,targetElement) {
	var parent = targetElement.parentNode;
	if (parent.lastChild == targetElement) {
		parent.appendChild(newElement);
	} else {
		parent.insertBefore(newElement,targetElement.nextSibling);
	}
}

//dojo的dom用法
/*
dojo.dom.isNode(dojo.byId('edtTitle'));
dojo.dom.firstElement(parentNode, 'SPAN');
dojo.dom.lastElement(parentNode, 'SPAN');

dojo.dom.nextElement(node, 'SPAN');
dojo.dom.prevElement(node, 'SPAN');

dojo.dom.moveChildren
把指定节点下的所有子节点移动到目标节点下，并返回移动的节点数
Usage Example:
dojo.dom.moveChildren(srcNode, destNode, true); //仅移动子节点，srcNode中的文字将被丢弃
dojo.dom.moveChildren(srcNode, destNode, false);//包括文字和子节点都将被移动到目标节点下


dojo.dom.copyChildren
把指定节点下的所有子节点复制到目标节点下，并返回复制的节点数
Usage Example:
dojo.dom.moveChildren(srcNode, destNode, true); //仅复制子节点，srcNode中的文字将被忽略
dojo.dom.moveChildren(srcNode, destNode, false);//包括文字和子节点都将被复制到目标节点下

dojo.dom.removeChildren
删除指定节点下的所有子节点，并返回删除的节点数
Usage Example:
dojo.dom.moveChildren(node);

dojo.dom.replaceChildren
用指定的新节点替换父节点下的所有子节点
Usage Example:
dojo.dom.replaceChildren(node, newChild); //目前还不支持newChild为数组形式


dojo.dom.removeNode
删除指定的节点
Usage Example:
dojo.dom.removeNode(node);


dojo.dom.getAncestors
返回指定节点的父节点集合
Usage Example:
dojo.dom.getAncestors(node, null, false); //返回所有的父节点集合（包括指定的节点node）
dojo.dom.getAncestors(node, null, true); //返回最近的一个父节点
dojo.dom.getAncestors(node, function(el){// 此处增加过滤条件 return true}, false); //返回所有满足条件的父节点集合

dojo.dom.getAncestorsByTag
返回所有符合指定Tag的指定节点的父节点集合
Usage Example:
dojo.dom.getAncestorsByTag(node, 'span', false); //返回所有的类型为SPAN的父节点集合
dojo.dom.getAncestorsByTag(node, 'span', true);  //返回最近的一个类型为SPAN的父节点


dojo.dom.getFirstAncestorByTag
返回最近的一个符合指定Tag的指定节点的父节点
Usage Example:
dojo.dom.getFirstAncestorByTag(node, 'span'); //返回最近的一个类型为SPAN的父节点


dojo.dom.isDescendantOf
判断指定的节点是否为另一个节点的子孙
Usage Example:
dojo.dom.isDescendantOf(node, ancestor, true); //判断node是否为ancestor的子孙
dojo.dom.isDescendantOf(node, node, false); //will return true
dojo.dom.isDescendantOf(node, node, true); //will return false


dojo.dom.innerXML
返回指定节点的XML
Usage Example:
dojo.dom.innerXML(node);


dojo.dom.createDocument
创建一个空的文档对象
Usage Example:
dojo.dom.createDocument();


dojo.dom.createDocumentFromText
根据文字创建一个文档对象
Usage Example:
dojo.dom.createDocumentFromText('<?xml version="1.0" encoding="gb2312" ?><a>1</a>','text/xml');

doc.load
根据文件装在XML
Usage Example:
var doc = dojo.dom.createDocument();
doc.load('http://server/dojo.xml');

 
dojo.dom.prependChild
将指定的节点插入到父节点的最前面
Usage Example:
dojo.dom.prependChild(node, parent);


dojo.dom.insertBefore
将指定的节点插入到参考节点的前面
Usage Example:
dojo.dom.insertBefore(node, ref, false); //如果满足要求的话就直接退出
dojo.dom.insertBefore(node, ref, true);


dojo.dom.insertAfter
将指定的节点插入到参考节点的后面
Usage Example:
dojo.dom.insertAfter(node, ref, false); //如果满足要求的话就直接退出
dojo.dom.insertAfter(node, ref, true);


dojo.dom.insertAtPosition
将指定的节点插入到参考节点的指定位置
Usage Example:
dojo.dom.insertAtPosition(node, ref, "before");//参考节点之前
dojo.dom.insertAtPosition(node, ref, "after"); //参考节点之后
dojo.dom.insertAtPosition(node, ref, "first"); //参考节点的第一个子节点
dojo.dom.insertAtPosition(node, ref, "last");  //参考节点的最后一个子节点
dojo.dom.insertAtPosition(node, ref); //默认位置为"last"


dojo.dom.insertAtIndex
将指定的节点插入到参考节点的子节点中的指定索引的位置
Usage Example:
dojo.dom.insertAtIndex(node, containingNode, 3);  //把node插入到containingNode的子节点中，使其成为第3个子节点


dojo.dom.textContent
设置或获取指定节点的文本
Usage Example:
dojo.dom.textContent(node, 'text'); //设置node的文本为'text'
dojo.dom.textContent(node); //返回node的文本


dojo.dom.hasParent
判断指定节点是否有父节点
Usage Example:
dojo.dom.hasParent(node);


dojo.dom.isTag
判断节点是否具有指定的tag
Usage Example:
var el = document.createElement("SPAN");
dojo.dom.isTag(el, "SPAN"); //will return "SPAN"
dojo.dom.isTag(el, "span"); //will return ""
dojo.dom.isTag(el, "INPUT", "SPAN", "IMG"); //will return "SPAN"

*/

//菜品相册上传功能专用函数.上传图片并且更新DOM的数据.
function Upload(controlStr) {
	//先令dijit.Dialog上传对话框产生loading...
	dojo.byId("uploading").innerHTML = '<img src=/images/loading.gif />';
	
	dojo.io.iframe.send({
		form : "foo",
		method : "post",
		//html
		handleAs : "json",
		url : "/upload/upfile.asp",
		load : function(response, ioArgs) {
			console.log(response, ioArgs);
			//alert(response);
			//get photo
			//var responseJson = dojo.fromJson(response);
			if (! response.valid) {	//上传失败.
				
				dojo.byId("uploading").innerHTML = '上传失败！<img src=/images/no.png />';
			} else {					//上传成功.
				//wl
				if (controlStr=="null" || controlStr==null) { controlStr=''; }
				
				dojo.byId("value_photo" + controlStr).value 	= response.src;
				dojo.byId("src_photo" + controlStr).src 		= response.src;
				dojo.byId("href_photo" + controlStr).href 		= response.src;
				//alert(typeof dijit.byId("tmp" + controlStr));
				if (typeof dijit.byId("tmp" + controlStr) == 'object' ) {
					dijit.byId("tmp" + controlStr).label 		= "图片详细尺寸：<br /><img src='" + response.src + "' />";
				}
				
				
				//上传结束，令dijit.Dialog上传对话框产生的loading停止.
				dojo.byId("uploading").innerHTML = '上传成功<img src=/images/ok.gif />';
			}
			
			
			return response;
		},
		error : function(response, ioArgs) {
			console.log("Error");
			console.log(response, ioArgs);
			
			//上传结束，令dijit.Dialog上传对话框产生的loading停止.
			dojo.byId("uploading").innerHTML = '网络出错，上传失败！<img src=/images/no.png />';
			return response;
		}
	});
}

//菜品相册上传功能专用函数2.构建上传对话框.
//编程法——不行！OK
//例1:	ShowDialog('<img src=/images/up.gif />修改上传图片', '../upload/index.asp?Action=Add&controlStr=' + randomNumber, 'width:300px;height:200px;','');
//例2:	ShowDialog('<span style=color:black;>感谢您的支持</span>','','width:300px;height:600px;','<div style=background-color:#ffffff;width:300px;height:300px;>技术支持QQ：595574668</div>');
function ShowDialog(titlename,hrefurl,styleStr,contentStr) {
	console.log('titlename:',titlename, 'hrefurl:',hrefurl, 'styleStr:',styleStr, 'contentStr:',contentStr);
	if (hrefurl=='') {
		var d1 = new dijit.Dialog({
			title:	titlename,
			content:contentStr,
			style:'' + styleStr + ''
		});
	} else {
		var d1 = new dijit.Dialog({
			title:	titlename,
			href:	hrefurl + "&CokeShow=" + Math.random(),
			style:'' + styleStr + ''
		});
	}
	
	//dijit.byId("dialog2").show();
	d1.show();
	//WL自创法宝.
	dojo.connect(d1, "hide", d1, function(e){d1.destroy()});
	//WL私有消除Widget管理中的id的bug.[重要补充处理！][针对菜品详情页-推荐朋友对话框中的form表单个Widget][01 Begin]
	//if (dijit.byId("tuijianpengyou")) { dojo.connect(d1, "hide", d1, function(e){dijit.byId("tuijianpengyou").destroy()}) };	//+ widget.attr("value") +
	//dojo.connect(d1, "hide", d1, function(e){dijit.byId("tuijianpengyou").destroy()})
	if (typeof(dojo.byId("FName"))=="object") { dojo.connect(d1, "hide", d1, function(e){dijit.byId("FName").destroy()}) };
	if (typeof(dojo.byId("FEmail"))=="object") { dojo.connect(d1, "hide", d1, function(e){dijit.byId("FEmail").destroy()}) };
	if (typeof(dojo.byId("CodeStr_TuiJianPengYou"))=="object") { dojo.connect(d1, "hide", d1, function(e){dijit.byId("CodeStr_TuiJianPengYou").destroy()}) };
	if (typeof(dojo.byId("submit886"))=="object") { dojo.connect(d1, "hide", d1, function(e){dijit.byId("submit886").destroy()}) };
	//WL私有消除Widget管理中的id的bug.[重要补充处理！][针对菜品详情页-推荐朋友对话框中的form表单个Widget][01 End]
	
}

//显示隐藏元素函数.
//dojo.style(this.domNode,{opacity:0,visibility:""});
function DisplayTheElement(ElementID) {
	var theElement = dojo.byId(ElementID);
	console.log(ElementID);
	console.log(theElement);
	console.log(dojo.style(theElement,"display"));
	
	//dojo.style(ElementID,{"display":"block"});
	if ( dojo.style(theElement,"display") == "none" ) {
		dojo.style(theElement, {display:"block"});
	} else {
		dojo.style(theElement, {display:"none"});
	}
	
}


//后台专用的表格样式控制代码.
//Table偶数行变色函数
function stripeTables(theTableIdName) {
	//鼠标掠过时的颜色.
	var highlightcolor = '#EAE8E3';		//#FFDFBF#FEFEF0鼠标掠过加亮显示的颜色.
	
	if (!document.getElementsByTagName) return false;
	var tables = document.getElementsByTagName("table");
	for (var i=0; i<tables.length; i++) {
		if (tables[i].getAttribute("id",1) == theTableIdName) {
			//令Table偶数行变色
			var odd = false;
			var rows = tables[i].getElementsByTagName("tr");
			for (var j=0; j<rows.length; j++) {
				
				//1.令Table偶数行变色
				if (odd == true) {
					rows[j].style.backgroundColor = '#FFFFFF';	//#FFFFEC偶数行颜色.
					rows[j].style.height = '23';				//'23'偶数行的高度.
					odd = false;
					
				} else {
					rows[j].style.backgroundColor = '#FFFFFF';	//#FFFFFF奇数行颜色.
					rows[j].style.height = '23';				//'23'奇数行的高度.
					odd = true;
				}
				
				//2.针对Table鼠标掠过加亮显示，绑定匿名函数
				rows[j].onmouseover = function() {
					//this.style.fontWeight = "bold";
					this.style.backgroundColor = highlightcolor;
				}
				rows[j].onmouseout = function() {
					//this.style.fontWeight = "normal";
					this.style.backgroundColor = "#FFFFFF";
				}
				
			}
			
			
		}
	}
}
//在需要的地方的文件中再去调用！例如后台页面中.
//Table偶数行变色函数.
//dojo.addOnLoad(function() {
//	stripeTables("listGo");
//});

//判断是否为正整数
function isPatrn(STRING,Patrn) {
	//var patrn=/^[0-9]*[1-9][0-9]*$/;
	
	if (!Patrn.exec(STRING)) {
		return false;
	}
	else {
		return true;
	} 
}


//搜索对话框.
//直接弹出对话框.
function ShowDialog_Search(strFromURL) {
	//操作成功时，提示！
	ShowDialog('<img src=/images/ico/small/group.png /> <span style=color:black;>请填写您要搜索的关键字(菜品)</span>','/Club/ShowDialog_Search.Welcome?strFromURL=' + strFromURL, 'width:460px;height:200px;','');
	//有其它尚未完成信息时，弹出提示.
	//ShowDialog('<span style=color:black;>提示信息</span>','','width:300px;height:200px;','<div style=background-color:#ffffff;width:auto;height:auto;><span style="color: #666;"><img src=/images/ico/emotion_wink.png />'+ data.theAllInformation +'<br /></span></div>');

}