<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("Ba_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�" then opt="<td><a href='modcorp2.asp?opt=����&corpid=-1'>��&nbsp;��&nbsp;��&nbsp;��</a></td>"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>���ɹ���</p><hr><table border=3 width=100% align=center><tr bgcolor=FFFF00 align=center><td>����</td><td>����</td><td width='50%'>���</td><td>������</td><td>����</td>"&opt&"</tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
set rst2=server.CreateObject("adodb.recordset")
rst.Open "����",conn
do while not (rst.EOF or rst.BOF)
	corpid=rst("id")
	corp=rst("����")
	if usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�" then opt="<TD><a href='modcorp2.asp?opt=�޸�&corpid="&corpid&"'>�޸�</a> | <a href='modcorp2.asp?opt=ɾ��&corpid="&corpid&"'>ɾ��</a></td>"
	if usercorp=corp then 
		msg=msg&"<tr><td><a href='showcorp.asp'>"&corp&"</a></td>"
	else
		msg=msg&"<tr><td>"&corp&"</td>"
	end if		
	rst2.Open "select count(*) as cnumber from �û� where ����='"&corp&"'",conn
	cnumber=rst2("cnumber")
	rst2.Close
	rst2.Open "select ���� from �û� where ����='"&corp&"' and ���='����'"
	if rst2.EOF or rst2.BOF then 
		cname="&nbsp;"
	else
		cname=rst2("����")
	end if
	msg=msg&"<td align=right>"&rst("����ϵ��")&"</td><td>"&rst("���")&"</td><td>"&cnumber&"</td><td>"&cname&"</td>"&opt&"</tr>"
	rst2.Close
	rst.MoveNext
loop
rst.Close
set rst2=nothing
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>