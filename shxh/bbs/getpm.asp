<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
%>
<!--#include file="set.asp"-->
<%
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �ż� where ������='"&username&"' and �ۿ�=false order by д��ʱ�� desc",conn
do while not (rst.EOF or rst.BOF)
	content=rst("����")
	msg=msg&"<table width=100% bgcolor=cccccc border=3><tr><td>�����ˣ�"&rst("д����")&"</td><td>����ʱ��:"&rst("д��ʱ��")&"</td><td align=right>"&autopro(rst("д����"))&autopm(rst("д����"))&"<a href='delpm.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='ɾ��������';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"ɾ��������"&chr(34)&"><img border=0 src='../images/del.gif'></a></td></tr><tr bgcolor=f7f7f7><td colspan=3>���⣺"&rst("����")&"</td><tr align=center><td colspan=3>����</td></tr><tr bgcolor=f7f7f7><td colspan=3>"&Autolink(content)&"</td></tr></table><br>"
	rst.MoveNext
loop
rst.Close
rst.Open "select * from �ż� where ������='"&username&"' and �ۿ�=true order by д��ʱ�� desc",conn
do while not (rst.EOF or rst.BOF)
	content=rst("����")
	msg=msg&"<table width=100% bgcolor=dddddd border=1><tr><td>�����ˣ�"&rst("д����")&"</td><td>����ʱ��:"&rst("д��ʱ��")&"</td><td align=right>"&autopro(rst("д����"))&autopm(rst("д����"))&"<a href='delpm.asp?id="&rst("id")&"'onmouseover="&chr(34)&"window.status='ɾ��������';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"ɾ��������"&chr(34)&"><img border=0 src='../images/del.gif'></a></td></tr><tr bgcolor=d7d7d7><td colspan=3>���⣺"&rst("����")&"</td><tr align=center><td colspan=3>����</td></tr><tr bgcolor=d7d7d7><td colspan=3>"&Autolink(content)&"</td></tr></table><br>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.execute "update �ż� set �ۿ�=True where ������='"&username&"'"
conn.close
set conn=nothing
if msg="" then msg="<p align=center><font color=0000FF size=4>��û������</font><input type=button value='��������ȥף��' onclick=javascript:location.href='../bbs/pm.asp?st="&username&"';></p>"
msg="<html><title>"&Application("Ba_jxqy_systemname")&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" oncontextmenu='self.event.returnValue=false' topmargin=20 leftmargin=0>"&msg&"</body></html>"
Response.Write msg
%>