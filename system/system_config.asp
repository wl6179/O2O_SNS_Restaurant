<%
Const system_owner_domain	="http://www.cokeshow.com.cn"	'系统开发维护版权所有者链接域名
Const system_user_domain	="http://localhost:80"	'拥有者域名
Const system_dir			="/"						'系统所在目录

Const cookies_name		="CokeShowComCn" 				'cookies名前缀
Const cookies_domain	="http://www.cokeshow.com.cn" 		'cookeies域名根
Const cache_name		="CokeShow" 					'系统缓存名前缀
Const SYSFOLDER_ADMIN	="general"						'该目录名称将被作为系统禁止外部访问使用，专供内部使用的后台目录文件夹


Const upload_dir		="uploadimages"					'默认上传目录，为空则使用用户所在目录，若需要修改为其他目录，请手工建目录
Const upload_file_fxt	="ChixinBugai"						'默认上传文件前缀.

Const dir_dj_system			="/scriptcokeshowcomcndojov100/"										'dj的文件夹所在.
Const filename_dj_MainCss	="/scriptcokeshowcomcndojov100/dojo/resources/dojo.css"				'dj的主要CSS文件名.
Const filename_dj_ThemesCss	="/scriptcokeshowcomcndojov100/dijit/themes/tundra/tundra.css"		'dj的风格样式CSS文件名.
Const classname_dj_ThemesCss="tundra"																	'dj的风格样式CSS样式的总类名-body区的定义.
Const filename_dj				="/scriptcokeshowcomcndojov100/dojo/dojo.js"						'dj的文件名.
Const filenameWidgetsCompress_dj="/scriptcokeshowcomcndojov100/dojo/dojo.CXBG.CokeShow.com.cn.js"'dj的压缩好的所有Widgets集的文件名.
Const isDebug_dj				="false" 																'dj的debug功能是否开启.
Const parseOnLoad_dj			="true"																	'dj的parseOnLoad功能是否开启.

Const dir_dj_system_foreground			="/scriptcokeshowcomcndojov110/"										'dj的文件夹所在.
Const filename_dj_MainCss_foreground	="/scriptcokeshowcomcndojov110/dojo/resources/dojo.css"				'dj的主要CSS文件名.
Const filename_dj_ThemesCss_foreground	="/scriptcokeshowcomcndojov110/dijit/themes/tundra/tundra.css"		'dj的风格样式CSS文件名.
Const classname_dj_ThemesCss_foreground	="tundra"																	'dj的风格样式CSS样式的总类名-body区的定义.
Const filename_dj_foreground				="/scriptcokeshowcomcndojov110/dojo/dojo.js"						'dj的文件名.
Const filenameWidgetsCompress_dj_foreground	="/scriptcokeshowcomcndojov110/dojo/dojo.CXBG.CokeShow.com.cn.js"'dj的压缩好的所有Widgets集的文件名.
Const isDebug_dj_foreground					="false" 																'dj的debug功能是否开启.
Const parseOnLoad_dj_foreground				="true"																	'dj的parseOnLoad功能是否开启.

Const is_password_cookies=1


Const system_JMailFrom					="services@chixinbugai.me" 		'系统邮件的完整email地址.
Const system_JMailSMTP					="smtp.qq.com"				'系统邮件的smtp服务器地址.
Const system_JMailMailServerUserName	="1419591768" 				'系统邮件的账号.
Const system_JMailMailServerPassWord	="YAHOO1982@000OOO"				'系统邮件的密码.

Const system_ReplyEmailAddress			="supper@chixinbugai.me"				'官方使用的显示可以进行回复的官方电子邮件帐号.


Const is_cokeshow_404ErrorAlert_system	=1 							'404报错系统是否开启.
Const is_cokeshow_warning_system		=1 							'预警系统是否开启.
%>