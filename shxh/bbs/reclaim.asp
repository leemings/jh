<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1%>
<!--#include file="set.asp"-->
<%
username=session("Ba_jxqy_username")
if instr(Ba_hostname,";"&username&";")=0 then Response.Redirect "../error.asp?id=233"
bid=Request.QueryString("bid")
id=Request.QueryString("id")
if not (isnumeric(bid) and isnumeric(id))then Response.Redirect "../error.asp?id=220"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ���� from bbs where ɾ��=True and ��̳="&bid&" and id="&id,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=238"
fid=rst("����")
rst.Close
set rst=nothing
conn.execute "update bbs set ɾ��=False where id="&id&" and ��̳="&bid
if fid<>0 then
	conn.execute "update bbs set �ظ�=�ظ�+1 where id="&fid&" and ��̳="&bid
end if	
conn.close
set conn=nothing
%>
<html><head><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="<%=Ba_bgcolor%>" background="<%=Ba_bgimage%>" topmargin=100><table align=center border=0 width=80% bgcolor=cccccc><tr align=center><td><font color=FF0000 size=4>��Ϣ�ɹ����ͣ������Ӻ��Զ��ر�</font></td></tr><tr align=center><td><a href='#' onclick='javascript:window.close();'>�ر�</a></td></tr></table></body></html>