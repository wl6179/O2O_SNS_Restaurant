<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True	'缓冲数据，才向客户输出；
Response.Charset="utf-8"
Session.CodePage = 65001
%>
<%
'模块说明：上传文件模块，接收二进制数据及处理模块.
'日期说明：2009-7-12
'版权说明：www.cokeshow.com.cn	可乐秀(北京)技术有限公司产品
'友情提示：本产品已经申请专利，请勿用于商业用途，如果需要帮助请联系可乐秀技术有限公司。
'联系电话：(010)-67659219
'电子邮件：cokeshow@qq.com
%>

<%
'可乐秀.中国CokeShow保护盾2010
'记录访问者IP也有点必要RS("LastLoginIP")	=Request.ServerVariables("REMOTE_ADDR")
If Session("enterLevel")="" Or isNull(Session("enterLevel")) Or isEmpty(Session("enterLevel")) Then
	'阻止执行
	Response.Write "<script type=""text/javascript"">alert('"& "未授权情况下使用编辑器！" &"');</script>"
	Response.End()
	Response.Redirect "/"
Else
	If Not isNumeric(Session("enterLevel")) Then
		'阻止执行
		Response.Write "<script type=""text/javascript"">alert('"& "未授权情况下使用编辑器！" &"');</script>"
		Response.End()
		Response.Redirect "/"
	End If
End If
'WL新增防御.
If Len(Session("enterName"))>30 Or Len(Session("enterName"))<4 Then
	Response.Write "<script type=""text/javascript"">alert('"& "未授权情况下使用编辑器！" &"');</script>"
	Response.End()
	Response.Redirect "/"
Else
	'通行.
End If


Public Function CheckPostSafe()
	Dim server_v1, server_v2
	CheckPostSafe = False
	
	server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))	'http://localhost:45233/test.asp
	server_v2 = CStr(Request.ServerVariables("SERVER_NAME"))	'localhost	(www.myhomestay.com.cn)
	If Mid(server_v1, 8, Len(server_v2))=server_v2 Then CheckPostSafe=True	'截取字符数Len(server_v2)；
	CheckPostSafe = True	'''WL;强制避免限制；
End Function


If CheckPostSafe()=False Then
	Response.Write "<br />对不起，为了系统安全，本操作已经记入日志。"
	Response.End
	Response.Redirect "/"
End If

%>

<!--#include file="../system/system_conn.asp"-->
<!--#include file="../system/system_class.asp"-->
<!--#include file="../system/photos_class.asp"-->
<%
'类实例化
Dim CokeShowPhotosClass
Set CokeShowPhotosClass = New PhotosClass
%>

<!--#include file="_upload.asp"-->
<%
'类实例化
Dim CokeShow
Set CokeShow = New SystemClass
CokeShow.Start		'调用类的Start方法，初始化类里的ReloadSetup()函数，并得到二维数组Setup
'Call CokeShow.SQLWarningSys()	'预警.

Dim CurrentTableName
CurrentTableName = "[CXBG_upfiles]"
%>



<%
'上传类实例化
Dim o
Set o = New GetPostImg

'设置接收的文件的字段名称.
Dim theRequestFileName
'"file1"
theRequestFileName	=	"file1"

'取得文件信息.获取传递文本参数.
Dim Action,FoundErr,ErrMsg
Dim cnname,enname
Dim file_name_old,file_name,file_ext,file_system_dir,file_upload_dir,file_webpath,file_type,file_size
Dim classid
cnname			=CokeShow.filtRequest(o.RetFieldText("cnname"))			'接收文本参数2
enname			=CokeShow.filtRequest(o.RetFieldText("enname"))			'接收文本参数3

file_name_old	=o.GetFilePath(theRequestFileName)						'原始文件名
file_name		=upload_file_fxt & o.getrandStr() & o.GetExtendName(theRequestFileName)	'生成文件名
file_ext		=o.GetExtendName(theRequestFileName)					'文件扩展名
file_system_dir	=system_dir									
file_upload_dir	=upload_dir									
file_webpath	=system_dir & upload_dir &"/"& file_name	'最终web完整路径
file_type		=o.GetFileType(theRequestFileName)			'文件类型
file_size		=CokeShow.CokeClng(o.GetFileSize(theRequestFileName))	'文件大小

Action			=CokeShow.filtRequest(o.RetFieldText("Action"))			'接收隐藏文本参数1

classid			=CokeShow.filtRequest(o.RetFieldText("classid"))		'接收隐藏文本参数10



'处理参数
'处理cnname
If cnname="" Then
	cnname=""
Else
	If Len(cnname)>0 Then cnname=Trim(cnname) Else cnname=""
End If
'处理enname
If enname="" Then
	enname=""
Else
	If Len(enname)>0 Then enname=Trim(enname) Else enname=""
End If
'处理classid
If classid="" Then
	classid=0
Else
	If isNumeric(classid) Then classid=CokeShow.CokeClng(classid) Else classid=0
End If


'限制上传类型--扩展名限制！
Dim allowExtend,arrayAllowExtend
allowExtend="gif|jpg|png|bmp"
arrayAllowExtend=Split(allowExtend,"|")
'如果不满足限制条件，则销毁一些重要参数，来保证处理不正常.
Dim ii,isSafe
isSafe=False
For ii=0 To Ubound(arrayAllowExtend)
	'如果有匹配的扩展名，标记其为合法！
	If "."& Lcase(arrayAllowExtend(ii))=Lcase(file_ext) Then
		
		isSafe= True
		
	End If
	
Next

'如果没有找到匹配的扩展名，标记其为危险！
If isSafe=False Then
	Action=""
	theRequestFileName=""
	Response.Write "文件不安全！"
	Response.End()
	'结束运行.
End If
'//'上传流程处理.
'//If Action="Add" Then
'//	Call SaveAdd()
'//ElseIf Action="Modify" Then
'//	Call SaveModify()
'//End If
'//
'//'错误处理
'//If FoundErr=True Then
'//	CokeShow.AlertErrMsg_general( ErrMsg )
'//End If


'定义变量
Dim sql,RS
Dim filename
Dim file1name
'新增.
If Action="Add" Then
	'记录上传后的完整本地路径.
	'file2必要.
	filename=Server.Mappath(system_dir) &"\"& upload_dir &"\"& file_name	'上传后的虚拟主机磁盘位置
'	Response.Write filename
'	response.End()
	'保存文件.并记录上传后的文件名.
	
	file1name=o.SaveToFile(theRequestFileName,filename)			'将file1的上传文件，保存到虚拟主机磁盘位置上.
	'file1name2=o.SaveToFile("file2",filename)
	Response.Write ("<textarea>{valid:true, msg:'文件"& file_webpath &"上传成功！',src:'"& file_webpath &"'}</textarea>")
	
	Set RS = Server.CreateObject("Adodb.RecordSet")
	sql = "SELECT * FROM "& CurrentTableName &""
	RS.Open sql,CONN,1,3
	
	RS.Addnew
		If cnname<>"" Then RS("cnname") = cnname
		If enname<>"" Then RS("enname") = enname
		
		RS("file_name_old")		= file_name_old
		RS("file_name")			= file_name
		RS("file_ext") 			= file_ext
		RS("file_system_dir") 	= file_system_dir
		RS("file_upload_dir") 	= file_upload_dir
		RS("file_webpath") 		= file_webpath
		RS("file_type") 		= file_type
		RS("file_size") 		= file_size
		
		If classid>0 Then RS("classid") = classid
		'RS("file2").AppendChunk o.GetFile("file2")	'把file2上传的文件直接写到数据库里
	RS.Update
	
	RS.Close
	Set RS=Nothing
	
	'生成缩略图.
	 Call CokeShowPhotosClass.shortPhoto(system_dir & upload_dir &"/"& file_name, 300, system_dir & upload_dir &"/300/"& file_name)
	 'Call CokeShowPhotosClass.shortPhoto(system_dir & upload_dir &"/"& file_name, 118, system_dir & upload_dir &"/120/"& file_name)
	 'Call CokeShowPhotosClass.shortPhoto(system_dir & upload_dir &"/"& file_name, 300, system_dir & upload_dir &"/300/"& file_name)

'修改.
ElseIf Action="Modify" Then
	Dim fileid
	'接收参数.
	fileid =CokeShow.filtRequest(o.RetFieldText("fileid"))
	
	'判断参数.
	If fileid="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	Else
		fileid=CokeShow.CokeClng(fileid)
	End If
	
	'拦截错误.
	'//If FoundErr=True Then Exit Sub
	
	'查询记录.
	sql="SELECT * FROM "& CurrentTableName &" WHERE fileid="& fileid
	Set RS=Server.CreateObject ("Adodb.RecordSet")
	RS.Open sql,CONN,1,3
	
	If RS.Bof And RS.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的记录！</li>"
		RS.Close
		Set RS=Nothing
		'//Exit Sub
	End If
	
	'往下正常执行修改操作.
	
	'file2必要.
	filename=Server.Mappath( RS("file_system_dir") ) &"\"& RS("file_upload_dir") &"\"& RS("file_name")	'此文件在虚拟主机的原来磁盘位置
	
	
	'保存文件.
	'如果修改操作中，没有需要上传新图片，则保留原有数据，且不新覆盖新图操作，保留原图.
	If Not(file_name_old="" And file_ext="" And file_size=0) Then
		file1name=o.SaveToFile(theRequestFileName,filename)			'仍然保存在原来位置.
	End If
	
	'file1name2=o.SaveToFile("file2",filename)
	file_webpath	=system_dir & RS("file_upload_dir") &"/"& RS("file_name")	'仅更新一下当前系统位置而已，它可能有变化,文件目录保持不变，文件名也是.
	
	Response.Write ("<textarea>{valid:true,msg:'文件"& file_webpath &"修改操作 上传成功！',src:'"& file_webpath &"'}</textarea>")		'仍然保存在原来位置.
	
	
	
		'如果修改操作中，没有需要上传新图片，则保留原有数据，仅更新classid字段.
		If Not(file_name_old="" And file_ext="" And file_size=0) Then
			If cnname<>"" Then RS("cnname") = cnname
			If enname<>"" Then RS("enname") = enname
			
			RS("file_name_old")		= file_name_old
			'//RS("file_name")			= file_name
			RS("file_ext") 			= file_ext
			RS("file_system_dir") 	= file_system_dir
			'//RS("file_upload_dir") 	= file_upload_dir
			RS("file_webpath") 		= file_webpath
			RS("file_type") 		= file_type
			RS("file_size") 		= file_size
		End If
		
		If classid>0 Then RS("classid") = classid
		
		RS("modifydate") 		= Now()
		'RS("file2").AppendChunk o.GetFile("file2")	'把file2上传的文件直接写到数据库里
	RS.Update
	
	RS.Close
	Set RS=Nothing

End If


'错误处理
If FoundErr=True Then
	CokeShow.AlertErrMsg_general( ErrMsg )
End If
Response.End()
%>