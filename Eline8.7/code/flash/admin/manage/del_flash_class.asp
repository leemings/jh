<!--#include file="session.asp"-->
<!--#include file="conn.asp"-->
<%
ClassID=request("ClassID")
set rs=server.createobject("adodb.recordset")
sql="delete * from flashClass where Classid="&Classid
rs.open sql,conn,1,3
response.redirect "manage_flash_class.asp"
rs.close
set rs=nothing  
conn.close
set conn=nothing
%>