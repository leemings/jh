<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.Buffer=true
    server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
    server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
    if mid(server_v1,8,len(server_v2))<>server_v2 then
        response.write "你提交的路径有误，禁止从站点外部提交数据！"
        response.end
    end if
pass=request.form("pass")
user=request.form("user")
if pass="" or user="" then Response.Redirect "../error.asp?id=016"
if pass="lhcfz" and user="961012" then
session("jxqy_adminbok")=true
Response.Redirect "../ysqx/administrator.asp"
else
Response.Redirect "../exit.asp"
end if
%>