<!--#include file="function.asp"-->
<%CheckAdmin2%>
<!--#include file="conn.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="delete from Poll where ID="&request.QueryString("id")
rs.open sql,conn,1,3
conn.close
set conn=nothing
response.redirect "Polllist.asp"
%>


