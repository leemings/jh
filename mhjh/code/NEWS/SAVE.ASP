<% @language="vbscript" %>

<%
if session("yx8_mhjh_username")<>Application("yx8_mhjh_admin") then Response.Redirect "../error.asp?id=046"
response.buffer=true
Response.Expires=0
%>
<!-- #include file="connect.asp" -->
<%
name=request("name")
neirong=request("neirong")
sql="select * from data "
rs.open sql,conn,3,2
rs.addnew
rs("name")=name
rs("neirong")=neirong
rs("time")=now()
rs.update
rs.close
response.redirect "add.asp"
%>
