<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("Ba_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"'><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442><tr align=center>"
pattern=Request.QueryString("pattern")
if pattern="" then pattern="����"
patternarr=array("ͷ��","����","����","����","��Ʒ","����")
for i=0 to 5
	if pattern=patternarr(i) then
		msg=msg&"<td bgcolor=ffFF00><a href='modarm1.asp?pattern="&patternarr(i)&"' onmouseover=""window.status='"&patternarr(i)&"��';return true;"" onmouseout=""window.status='';return true;"">"&patternarr(i)&"</a></td>"
	else
		msg=msg&"<td><a href='modarm1.asp?pattern="&patternarr(i)&"' onmouseover=""window.status='"&patternarr(i)&"��';return true;"" onmouseout=""window.status='';return true;"">"&patternarr(i)&"</a></td>"
	end if	
next
if usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�" then opt="<td><a href='modarm2.asp?opt=����&id=-1&pattern="&pattern&"'>������"&pattern&"</a></td>"
msg=msg&"</tr></table><p align=center>��������</p><hr><table border=3 width=80% align=center><tr bgcolor=FFFF00 align=center><td>����</td><td>����</td><td>����</td><td>��Ч</td><td>�۸�</td>"&opt&"</tr>"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ID,����,����,����,��Ч,�۸� from �̵� where ����='"&pattern&"' order by �۸�",conn
do while not (rst.EOF or rst.BOF)
	id=rst("id")
	if usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�" then opt="<TD><a href='modarm2.asp?opt=�޸�&id="&id&"'>�޸�</a> | <a href='modarm2.asp?opt=ɾ��&id="&id&"'>ɾ��</a></td>"
	msg=msg&"<tr><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("��Ч")&"</td><td>"&rst("�۸�")&"</td>"&opt&"</tr>"
	rst.MoveNext
loop
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
msg=msg&"</table></body>"
Response.Write msg
%>