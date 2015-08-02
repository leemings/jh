<%
username=session("Ba_jxqy_username")
if username="" then Response.redirect "../error.asp?id=016"
%>
<html>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
</head>
<frameset cols="90,*" border=0>
	<frame name="menufrm" src="menu.asp"  scrolling=auto>
	<frame name="resultfrm" src="welcome.asp">
</frameset>	
</html>