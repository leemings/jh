<%
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
if session("Ba_jxqy_userposr")="" then
	session("Ba_jxqy_userposr")=1
	session("Ba_jxqy_userposc")=0
	session("Ba_jxqy_userfight")="none"
end if
if Application("Ba_jxqy_fellow")=true then msg="(仅对会员开放 )"
%>
<head>
<title><%=Application("Ba_jxqy_systemname")&msg%></title>
</head>
<frameset cols="108,*,120" border=0>
	<frameset rows="*,108"  border=0>
		<frame name=confrm src="about:blank" scrolling=no>
		<frame name=mapfrm src="map.asp"  scrolling=no>
	</frameset>	
	<frame name=msgfrm src='standby.htm' scrolling=auto>
	<frameset rows="200,*,0"  border=0>
		<frame name=behfrm src="about:blank" scrolling=no>
		<frame name=optfrm src="welcome.asp" scrolling=auto>
		<frame name=actfrm src="about:blank">
	</frameset>
</frameset>
