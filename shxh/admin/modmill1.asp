<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("Ba_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�" then opt="<td><a href='modmill2.asp?opt=����&id=-1'>��������Ŀ</a></td>"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>�򹤹���</p><hr><table border=3 width=80% align=center><tr bgcolor=FFFF00 align=center><td>����</td><td>˵��</td><td>����</td>"&opt&"</tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �� order by ����",conn
do while not (rst.EOF or rst.BOF)
	id=rst("id")
	if usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�" then opt="<TD><a href='modmill2.asp?opt=�޸�&id="&id&"'>�޸�</a> | <a href='modmill2.asp?opt=ɾ��&id="&id&"'>ɾ��</a></td>"
	msg=msg&"<tr><td>"&rst("����")&"</td><td>"&rst("˵��")&"</td><td>"&rst("����")&"</td>"&opt&"</tr>"
	rst.MoveNext
loop
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>