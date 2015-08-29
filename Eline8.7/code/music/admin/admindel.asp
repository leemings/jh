<!--#include file="function.asp"-->
<%CheckAdmin3%>
<!--#include file="conn.asp"-->
<%
set rs=server.CreateObject("ADODB.RecordSet")
rs.open "delete * from admin where id="&request.QueryString("id"),conn,1
set rst=nothing
response.redirect "adminlist.asp"
%>


