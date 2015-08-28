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
if pass="yx86666" and user="游戏吧" then
session("jxqy_adminbok")=true
Response.Redirect "../ysqx1/3453sdfg.asp"
else
Response.Redirect "../exit.asp"
end if
%>