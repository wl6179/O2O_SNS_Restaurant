<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True	'缂撳啿鏁版嵁锛屾墠鍚戝鎴疯緭鍑猴紱
Response.Charset="utf-8"
Session.CodePage = 65001
%>

<%
Response.Write "{valid: false, message: '您需要登录！'}"
%>