<!--#include file="conn.asp"-->
<%
id1=request("id")
set rs=server.createobject("adodb.recordset")
	if request("id")="" then
		'sql="select * from Music where checkup=true order by id desc"
		sql="select * from Music where checkup=true and playorder>0 order by playorder asc"
	else
		sql="select * from Music where id>"&id1&" or id="&id1&""
	end if
rs.open sql,conn,1,3

	do while not rs.eof
%>
		mkList("<%=rs("file")%>","<%=rs("musicname")%> - <%=rs("singer")%>");
<% 
		rs.movenext
	loop
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
