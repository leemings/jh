<%
bgcolor=Application("yx8_mhjh_backgroundcolor")
bgimage=Application("yx8_mhjh_backgroundimage")
username=session("yx8_mhjh_username")
usergrade=session("yx8_mhjh_usergrade")
usercorp=session("yx8_mhjh_usercorp")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=120 and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
id=Request.QueryString("id")
opt=Request.QueryString("opt")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"

set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ����  where id="&id,conn
if opt="����" then
	question=""
	answer=""
	ti=""
	money=""
	qiang=false
else
	if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
	question=rst("����")
	answer=rst("��")
	ti=rst("�ṩ��")
	money=rst("����")
	qiang=rst("����")
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if qiang=true then
	qiang="<select name=qiang><option value='true' selected>������</option><option value='false' >������</option></select>"
else
	qiang="<select name=qiang><option value='true' >������</option><option value='false'  selected>������</option></select>"
end if
msg=msg&"<head><title>"&Application("yx8_mhjh_systemname")&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false  background=../chatroom/bg1.gif bgcolor='"&bgcolor&"'><p align=center>�������</p><hr><form action=updateti.asp method=post><input type=hidden name='schid' value='"&id&"'><table border=3 width=90% align=center><tr><td>����</td><td><input type=text name=question size=50 maxlength=255 value='"&question&"'></td></tr><tr><td>��</td><td><input type=text value='"&answer&"'size=20 maxlength=30 name=answer></td></tr><tr><td>�ṩ��</td><td><input type=text name=ti value="&chr(34)&ti&chr(34)&" size=10 maxlength=20></td></tr><tr><td>����</td><td><input type=text name=money value="&chr(34)&money&chr(34)&" size=10 maxlength=10></td></tr><tr><td>����</td><td>"&qiang&"</td></tr><tr align=center><td colspan=2><input type=submit name=submit value='"&opt&"'> <input type=button onclick=javascript:history.back(); value='����'></td></tr></table></body>"
Response.Write msg
%>