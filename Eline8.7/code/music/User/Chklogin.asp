<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="conn.asp"-->
<!--#include file="function.asp"-->
<%
founderr=false
UserIP=Request.ServerVariables("REMOTE_ADDR")
if request("username")="" then
	errmsg=errmsg+""+"请输入您的用户名。"
	founderr=true
else
	username=Checkin(trim(request("username")))
end if
if request("password")="" then
	errmsg=errmsg+""+"请输入您的密码。"
	founderr=true
else
	password=Checkin(trim(request("password")))
end if
set rs=server.createobject("adodb.recordset")
sql="select username,password,loginIP from [user] where username='"&username&"' and lockuser=0"
rs.open sql,conn,1,3
if rs.bof and rs.eof then
	errmsg=errmsg+""+"您输入的用户名并不存在。<br><br><a href='reg.asp'>请先注册</a>"
	founderr=true
else
	if rs("password")<>password then
		errmsg=errmsg+""+"您输入的密码不正确。"
		founderr=true
	else
		rs("loginIP")=Request.ServerVariables("REMOTE_ADDR")
		rs.update
	       	session("UserName")=UserName
		session("PassWord")=PassWord
		session.timeout=200
	end if
end if

if founderr=true then
	call error()
else
	response.redirect("../index.asp")
end if
%>