<!--#include file="../showchat.asp"-->
<%
session("aqjh_adminok")=false
if session("aqjh_name")<>Application("aqjh_user") then
'��ʾ��¼��Ϣ
says="<font color=red>��ϵͳ��ʾ��["&session("aqjh_name")&"]</font><font color=green>�Ѿ��˳��˺�̨����[��ǰʱ��:"&now()&"]"
call showchat(says)
'��ʾ��¼��Ϣ
end if
Response.Redirect "login.asp"
%>