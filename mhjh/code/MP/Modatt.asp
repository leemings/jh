<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
selcorp=Request.QueryString("selcorp")
if selcorp="" then selcorp="��"
username=session("yx8_mhjh_username")
usercorp=session("yx8_mhjh_usercorp")
if username="" then Response.Redirect "../error.asp?id=016"
msg="<head><link rel='stylesheet' href='../style.css'><title>�����书</title></head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'><p align=center><form name=form1><select name=sele1 onchange="&chr(34)&"javascript:location.replace('modatt.asp?selcorp='+document.form1.sele1.value);"&chr(34)&"><option value='��'>��������</option>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
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
msg=msg&"</select></form></p><hr><table border=1 cellspacing=0 cellpadding=2  bordercolorlight=#993300 bordercolordark=#FFFFFF  align=center><tr align=center bgcolor=ffffff><td>��ʽ</td><td>����</td><td>˵��</td>"&opt&"</tr>"
rst.Close
rst.Open "select * from ��ʽ where ����='"&selcorp&"' order by ��������",conn
do while not (rst.EOF or rst.BOF)
attid=rst("id")
msg=msg&"<tr><td>"&rst("��ʽ")&"</td><td align=right>"&rst("��������")&"</td><td>"&rst("˵��")&"</td>"&opt&"</tr>"
rst.MoveNext
loop
msg=msg&"</table></body>"
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg
%>
