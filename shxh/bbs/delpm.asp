<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
%>
<!--#include file="set.asp"-->
<%
id=Request.QueryString("id")
username=session("Ba_jxqy_username")
if not isnumeric(id) then Response.Redirect "../error.asp?id=220"
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
conn.execute "delete * from �ż� where ������='"&username&"' and id="&id
conn.close
set conn=nothing
Response.write "<html><head><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>setTimeout('window.close();',3000);</script></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" topmargin=100><table align=center border=0 width=80% bgcolor=cccccc><tr align=center><td><font color=FF0000 size=4>��Ϣ�ɹ����ͣ������Ӻ��Զ��ر�</font></td></tr><tr align=center><td><a href='#' onclick='javascript:window.close();'>�ر�</a></td></tr></table></body></html>"
%>