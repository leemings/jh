<%@ LANGUAGE=VBScript%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
nickname=info(0)
if info(0)<>Application("ljjh_admin") then Response.Redirect "manerr.asp?id=255"
pass=request.form("pass")
if pass="" then
session("ljjh_adminok")=false
%>
<html>
<head>
<title><%=Application("ljjh_chatroomname")%></title>
<style></style>
<link rel="stylesheet" href="../chat/READONLY/Style.css">
</head>
<body bgcolor=#FFFFFF background="../bgcheetah.gif">
<center>
<font color="#000000"><b><font size="+2"><%=Application("ljjh_chatroomname")%></font></b></font><br>
<br>
管理登陆窗口
<form action=login.asp method=post>
请输入超级管理密码:<input type=password size=20 name=pass><input type=submit value='确认'>
</form>
</center>
</body>
</html>
<%elseif pass=Application("ljjh_adminkey") then
session("sjjh_adminok")=true
Response.Redirect "admin.asp"
else
Response.Redirect "../exit.asp"
end if%>