<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
stock=Request.QueryString("stock")
if stock="" then stock="�������"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=Session("Ba_jxqy_username")
msg="<head><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>setTimeout('location.reload();',300000);</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><center><font color=red><h3>ʵʱ����</h3></font><center><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffff00><tr align=center bgcolor=FFFF00><td>����</td><td>����</td></tr><tr align=center bgcolor=00FF00><td colspan=2><a href='sale.asp?stock="&stock&"' onmouseover="&chr(34)&"window.status='���������е�"&stock&"�ɷ�';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">����</a></td></tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select top 10 * from �ֹ� where ����='"&stock&"' and ����=True and ����>0  order by  ��� desc,ʱ��",conn
do while not (rst.EOF or rst.BOF) 
	msg=msg&"<tr align=right><td>"&rst("���")&"</td><td>"&rst("����")&"</td></tr>"
	rst.MoveNext
loop
rst.Close
msg=msg&"<tr align=center bgcolor=00FF00><td colspan=2><a href='buy.asp?stock="&stock&"' onmouseover="&chr(34)&"window.status='�չ�"&stock&"�ɷ�';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">�չ�</a></td></tr>"
rst.Open "select top 10 * from �ֹ� where ����='"&stock&"' and ����=False and ����>0  order by  ���,ʱ��",conn
do while not (rst.EOF or rst.BOF) 
	msg=msg&"<tr align=right><td>"&rst("���")&"</td><td>"&rst("����")&"</td></tr>"
	rst.MoveNext 
loop
rst.Close
rst.Open "select * from �ֹ� where ����='"&stock&"' and �ֹ���='"&username&"' and ����>0",conn
if not rst.EOF then
	if rst("����")=True then 
		opt="����"
	else
		opt="�չ�"
	end if	
	msg=msg&"<tr bgcolor=00ff00><td>����</td><td>"&opt&"</td></tr><tr><td>"&rst("���")&"</td><td>"&rst("����")&"</td><tr><td colspan=2 align=center><a href='cancel.asp?stock="&stock&"' onmouseover="&chr(34)&"window.status='�������"&stock&"�Ĵ˴�"&opt&"Ҫ��';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��������</a></td></tr></tr>"
end if
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg&"</table></body>"
%>