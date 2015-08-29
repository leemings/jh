<%
	if Session("IsAdmin")<>true then response.redirect "admin_login.asp"
%>