<%@ CodePage=65001 Language="VBScript"%>
<%
Option Explicit
Response.Buffer = True
%>

<%
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

<%
 ' FCKeditor - The text editor for Internet - http://www.fckeditor.net
 ' Copyright (C) 2003-2009 Frederico Caldeira Knabben
 '
 ' == BEGIN LICENSE ==
 '
 ' Licensed under the terms of any of the following licenses at your
 ' choice:
 '
 '  - GNU General Public License Version 2 or later (the "GPL")
 '    http://www.gnu.org/licenses/gpl.html
 '
 '  - GNU Lesser General Public License Version 2.1 or later (the "LGPL")
 '    http://www.gnu.org/licenses/lgpl.html
 '
 '  - Mozilla Public License Version 1.1 or later (the "MPL")
 '    http://www.mozilla.org/MPL/MPL-1.1.html
 '
 ' == END LICENSE ==
 '
 ' This is the File Manager Connector for ASP.
%>
<!--#include file="config.asp"-->
<!--#include file="util.asp"-->
<!--#include file="io.asp"-->
<!--#include file="basexml.asp"-->
<!--#include file="commands.asp"-->
<!--#include file="class_upload.asp"-->
<%

If ( ConfigIsEnabled = False ) Then
	SendError 1, "This connector is disabled. Please check the ""editor/filemanager/connectors/asp/config.asp"" file"
End If

DoResponse

Sub DoResponse()
	Dim sCommand, sResourceType, sCurrentFolder

	' Get the main request information.
	sCommand = Request.QueryString("Command")

	sResourceType = Request.QueryString("Type")
	If ( sResourceType = "" ) Then sResourceType = "File"

	sCurrentFolder = GetCurrentFolder()

	' Check if it is an allowed command
	if ( Not IsAllowedCommand( sCommand ) ) then
		SendError 1, "The """ & sCommand & """ command isn't allowed"
	end if

	' Check if it is an allowed resource type.
	if ( Not IsAllowedType( sResourceType ) ) Then
		SendError 1, "Invalid type specified"
	end if

	' File Upload doesn't have to Return XML, so it must be intercepted before anything.
	If ( sCommand = "FileUpload" ) Then
		FileUpload sResourceType, sCurrentFolder, sCommand
		Exit Sub
	End If

	SetXmlHeaders

	CreateXmlHeader sCommand, sResourceType, sCurrentFolder, GetUrlFromPath( sResourceType, sCurrentFolder, sCommand)

	' Execute the required command.
	Select Case sCommand
		Case "GetFolders"
			GetFolders sResourceType, sCurrentFolder
		Case "GetFoldersAndFiles"
			GetFoldersAndFiles sResourceType, sCurrentFolder
		Case "CreateFolder"
			CreateFolder sResourceType, sCurrentFolder
	End Select

	CreateXmlFooter

	Response.End
End Sub

%>
