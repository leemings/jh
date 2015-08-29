<!--#include file="session.asp"-->
<!--#include file="conn.asp"-->
<%
ID=request("ID")
set rs=server.createobject("adodb.recordset")
sql="delete * from Flash where id="&id
rs.open sql,conn,1,3
response.redirect "manage_Flash.asp"
rs.close
set rs=nothing  
conn.close
set conn=nothing
%>