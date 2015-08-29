<!--#include file="session.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/char.inc"-->
<%
Password=replace(request.form("Password"),"'","''")
set rs=server.createobject("adodb.recordset")
sql="select * from Admin where username='"&session("username")&"'"
rs.open sql,conn,1,3

rs("Password")=Password
rs.update
rs.close
set rs=nothing
response.write"<Script>alert('个人密码修改成功');location.href='personal.asp';</Script>"
response.end
%>