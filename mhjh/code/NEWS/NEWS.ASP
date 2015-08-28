
<!-- #include file="connect.asp" -->
<%rs.Open "Select top 3 * From data Order By time DESC",conn,3
'可将5修改为你所需要在首页显示的信息条数%>
<%
if not rs.eof then
rs.movefirst
dis_mes=dis_mes & ""
do while not rs.eof
    topic=rs("name")
	if len(topic)>20 then
		s_topic=left(topic,20)+"……"
	else
		s_topic=topic
	end if
dis_mes=dis_mes &"&nbsp;<img src=./image/lastpost.gif>&nbsp;<a href=./news/details.asp?id="&rs("id")&" target=_blank>"&s_topic&"</a><br>"
rs.movenext
loop
dis_mes=dis_mes & ""
end if
rs.close
dis_mes="document.write('" & dis_mes & "');"
response.write dis_mes
%>

