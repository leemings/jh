<!-- #include file="conn.asp" -->
<%
sql="select username from online where username<>'' and eremite<>1"
Set Rs=Conn.Execute(sql)
onlinelist="|"
Do While Not RS.EOF
onlinelist=""&onlinelist&""&rs("username")&"|"
rs.movenext
loop
Set Rs = Nothing
%>
var onlinelist="<%=onlinelist%>";
<%
conn.close
set conn=nothing
%>