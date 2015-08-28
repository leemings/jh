
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
if username<>Application("yx8_mhjh_admin") then Response.Redirect "../exit.asp"
%>
<html>
<head>
<title>管理</title>
<style></style>
<link rel="stylesheet" href="../Style.css">
</head>
<body oncontextmenu=self.event.returnValue=false background="../chatroom/bg1.gif">
<p></p>
<p>
<center>
<br>
管理登陆窗口</p>
<form action=do.asp method=post>
请输入超级管理帐号:<input type=text size=20 name=user><br>请输入超级管理密码:<input type=password size=20 name=pass><br><input type=submit value='确认'>
</form>
</center>
<p align="center">游戏吧 2002-7-25</p>
</body>
</html>