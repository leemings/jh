<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1%>
<!--#include file="set.asp"-->
<%
username=session("Ba_jxqy_username")
bid=Request.QueryString("bid")
id=Request.QueryString("id")
if not (isnumeric(bid) and isnumeric(id))then Response.Redirect "../error.asp?id=220"
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ��� where id="&bid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=220"
hostlist=";"&rst("����")&";"
rst.Close
rst.Open "select ����,���� from bbs where ɾ��=false and id="&id&" and ��̳="&bid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=223"
author=rst("����")
fid=rst("����")
rst.Close
set rst=nothing
if not (author=username or instr(hostlist," "&username&" ")<>0 or instr(Ba_hostname,";"&username&";")<>0) then Response.Redirect "../error.asp?id=027"
conn.execute "update bbs set ɾ��=True where id="&id&" and ��̳="&bid
if fid<>0 then
	conn.execute "update bbs set �ظ�=�ظ�-1 where id="&fid&" and ��̳="&bid
end if	
conn.close
set conn=nothing
%>
<html><head><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>setTimeout('window.close();',3000);</script></head><body bgcolor="<%=Ba_bgcolor%>" background="<%=Ba_bgimage%>" topmargin=100><table align=center border=0 width=80% bgcolor=cccccc><tr align=center><td><font color=FF0000 size=4>��Ϣ�ɹ����ͣ������Ӻ��Զ��ر�</font></td></tr><tr align=center><td><a href='#' onclick='javascript:window.close();'>�ر�</a></td></tr></table></body></html>