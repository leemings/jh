<%
username=session("yx8_mhjh_username")
if session("jxqy_adminbok")="" then Response.Redirect "../error.asp?id=016"
if username<>Application("yx8_mhjh_admin") then Response.redirect "../exit.asp"
%>
<html>
<head>
<title><%=Application("yx8_mhjh_systemname")%></title>
</head>
<frameset cols="90,*" border=0 frameborder="0">
<frame name="menufrm" src="mee.asp"  scrolling=auto>
<frame name="resultfrm" src="welcome.asp">
</frameset>
</html>