<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="config.asp"-->
<%
varadm=request.form("username")
varpwd=request.form("password")
if varadm="" then
      response.redirect "error.asp?msg=���������Ա����!"
end if
if varpwd="" then
      response.redirect "error.asp?msg=���������Ա����!"
end if
if varadm=adm and varpwd=pwd then
      session("admin")="ture"
      response.redirect "manage.asp"
else
      response.redirect "error.asp?msg=����û���������������!"
end if
%>