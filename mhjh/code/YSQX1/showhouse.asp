<%
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"�ٸ�" then Response.Redirect "../exit.asp"
if username=adminname  then opt="<td><a href='modhouse.asp?opt=����&id=-1'>��&nbsp;��&nbsp;��&nbsp;��</a></td>"
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'><p align=center>���ݹ���</p><hr><table border=3 width=90% align=center><tr bgcolor=FFFF00 align=center><td>����</td><td>����</td><td>ͼƬ</td><td>����</td><td>�ۼ�</td><td>λ��</td><td>�ֵ�</td><td>״̬</td><td>����</td><td>��������</td>"&opt&"</tr>"
%>
<!--#include file="../fw/data1.asp"--><%
sql="select * from ���� order by �ۼ�"
set rst=conn.execute(sql)
 do while not (rst.EOF or rst.BOF)
	id=rst("id")
	if username=adminname then opt="<TD><a href='modhouse.asp?opt=�޸�&id="&id&"'>�޸�</a> | <a href='modhouse.asp?opt=ɾ��&id="&id&"'>ɾ��</a></td>"
	msg=msg&"<tr><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("ͼƬ")&"</td><td>"&rst("����")&"</td><td>"&rst("�ۼ�")&"</td><td>"&rst("λ��")&"</td><td>"&rst("�ֵ�")&"</td><td>"&rst("״̬")&"</td><td>"&rst("����")&"</td><td>"&rst("��������")&"</td>"&opt&"</tr>"
	rst.MoveNext
loop
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>