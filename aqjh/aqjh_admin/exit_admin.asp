<!--#include file="../showchat.asp"-->
<%
session("aqjh_adminok")=false
if session("aqjh_name")<>Application("aqjh_user") then
'显示登录信息
says="<font color=red>【系统提示】["&session("aqjh_name")&"]</font><font color=green>已经退出了后台管理！[当前时间:"&now()&"]"
call showchat(says)
'显示登录信息
end if
Response.Redirect "login.asp"
%>