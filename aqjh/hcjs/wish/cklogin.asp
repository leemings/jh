<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="config.asp"-->
<%
varadm=request.form("username")
varpwd=request.form("password")
if varadm="" then
      response.redirect "error.asp?msg=请输入管理员名字!"
end if
if varpwd="" then
      response.redirect "error.asp?msg=请输入管理员密码!"
end if
if varadm=adm and varpwd=pwd then
      session("admin")="ture"
      response.redirect "manage.asp"
else
      response.redirect "error.asp?msg=你的用户名和密码有问题!"
end if
%>