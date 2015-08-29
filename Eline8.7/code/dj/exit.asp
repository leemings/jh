<%
	if request.cookies("okerer")="" then
	response.write "<script>alert('дЦипн╢╣гб╫ё╛нчпК╣Ц╩ВмкЁЖё║');history.back();</Script>"
	else
	response.cookies("okerer")=""
	response.cookies("username")=""
	response.redirect "index.asp"
	end if
%>