<%
username=session("yx8_mhjh_username")
usergrade=session("yx8_mhjh_usergrade")
usercorp=session("yx8_mhjh_usercorp")
bgcolor=Application("yx8_mhjh_backgroundcolor")
bgimage=Application("yx8_mhjh_backgroundimage")
if usergrade=Application("yx8_mhjh_allright") and usercorp="�ٸ�" then opt="<td><a href='modti.asp?opt=����&id=-1'>��&nbsp;��&nbsp;��&nbsp;��</a></td>"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center>�������</p><hr><table border=3 width=90% align=center><tr bgcolor=FFFF00 align=center><td>����</td><td>��</td><td>�ṩ��</td><td>����</td><td>����</td>"&opt&"</tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ����",conn
do while not (rst.EOF or rst.BOF)
	id=rst("id")
	if usergrade=Application("yx8_mhjh_allright") and usercorp="�ٸ�" then opt="<TD><a href='modti.asp?opt=�޸�&id="&id&"'>�޸�</a> | <a href='modti.asp?opt=ɾ��&id="&id&"'>ɾ��</a></td>"
	msg=msg&"<tr><td>"&rst("����")&"</td><td>"&rst("��")&"</td><td>"&rst("�ṩ��")&"</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td>"&opt&"</tr>"
	rst.MoveNext
loop
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>