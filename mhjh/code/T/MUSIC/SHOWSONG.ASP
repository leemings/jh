<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username<>Application("yx8_mhjh_admin") then Response.Redirect "../../error.asp?id=046"
if username=Application("yx8_mhjh_admin")  then opt="<td><a href='modsong.asp?opt=����&id=-1'>��&nbsp;��&nbsp;��&nbsp;��</a></td>"
msg="<head><link rel='stylesheet' href='../../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='../../chatroom/bg1.gif'><p align=center>��������</p><hr><table border=3 width=90% align=center><tr bgcolor=FFFF00 align=center><td>������ַ</td><td>��������</td><td>����</td><td>������</td><td>���</td>"&opt&"</tr>"
%>
<!--#include file="data.asp"-->
<%
rs.Open"SELECT * FROM ����  order by ��� desc",conn,3
 do while not (rs.EOF or rs.BOF)
	id=rs("id")
	if username=Application("yx8_mhjh_admin") then opt="<TD><a href='modsong.asp?opt=�޸�&id="&id&"'>�޸�</a> | <a href='modsong.asp?opt=ɾ��&id="&id&"'>ɾ��</a></td>"
	msg=msg&"<tr><td>"&rs("��ַ")&"</td><td>"&rs("����")&"</td><td>"&rs("����")&"</td><td>"&rs("������")&"</td><td>"&rs("���")&"</td>"&opt&"</tr>"
	rs.MoveNext
loop
rs.Close 
set rs=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>




