<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
selcorp=Request.QueryString("selcorp")
if selcorp="" then selcorp="��"
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
msg="<head><link rel='stylesheet' href='../chatroom/css.css'>.</head><body oncontextmenu=self.event.returnValue=false background='"&bgimage&"' bgcolor='"&bgcolor&"'><p align=center><form name=form1><select name=sele1 onchange="&chr(34)&"javascript:location.replace('modatt1.asp?selcorp='+document.form1.sele1.value);"&chr(34)&"><option value='��'>��������</option>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "����",conn
do while not rst.EOF
	corp=rst("����")
	if corp=selcorp then
		msg=msg&"<option value='"&corp&"' selected>"&corp&"</option>"
	else
		msg=msg&"<option value='"&corp&"'>"&corp&"</option>"
	end if
	rst.MoveNext
loop
if usercorp="�ٸ�" and usergrade=Application("Ba_jxqy_allright") then opt="<td><a href='modatt2.asp?opt=����&corp="&selcorp&"&attid=-1'>��&nbsp;��&nbsp;��&nbsp;ʽ</a></td>"
msg=msg&"</select></form></p><hr><table border=3 width=80% align=center><tr align=center bgcolor=FFFF00><td>��ʽ</td><td>����</td><td>˵��</td>"&opt&"</tr>"
rst.Close
rst.Open "select * from ��ʽ where ����='"&selcorp&"' order by ��������",conn
do while not (rst.EOF or rst.BOF)
	attid=rst("id")
	if usercorp="�ٸ�" and usergrade=Application("Ba_jxqy_allright") then opt="<td><a href='modatt2.asp?opt=�޸�&corp="&selcorp&"&attid="&attid&"'>�޸�</a> | <a href='modatt2.asp?opt=ɾ��&corp="&selcorp&"&attid="&attid&"'>ɾ��</a></td>"
	msg=msg&"<tr><td>"&rst("��ʽ")&"</td><td align=right>"&rst("��������")&"</td><td>"&rst("˵��")&"</td>"&opt&"</tr>"
	rst.MoveNext
loop
msg=msg&"</table></body>"
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg
%>