<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
msg=Request.QueryString("msg")
if msg="" then msg="����"
msgarr=array("ͷ��","����","����","����","��Ʒ","����")
msgtmp="<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442><tr align=center>"
for i=0 to 5
	if msg=msgarr(i) then
		msgtmp=msgtmp&"<td bgcolor=ffFF00><a href='armorshop.asp?msg="&msgarr(i)&"' onmouseover=""window.status='"&msgarr(i)&"��';return true;"" onmouseout=""window.status='';return true;"">"&msgarr(i)&"</a></td>"
	else
		msgtmp=msgtmp&"<td><a href='armorshop.asp?msg="&msgarr(i)&"' onmouseover=""window.status='"&msgarr(i)&"��';return true;"" onmouseout=""window.status='';return true;"">"&msgarr(i)&"</a></td>"
	end if	
next
msgtmp=msgtmp&"</tr></table>"
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select ID,����,����,����,��Ч,�۸� from �̵� where ����='"&msg&"' order by �۸� "
rst.open sqlstr,conn
msg="<table width='100%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF><tr bgcolor=FFFF00 align=center><td>����</td><td>����</td><td>����</td><td>��Ч</td><td>�۸�</td></TR>"
do while not(rst.EOF or rst.BOF)
	msg=msg&"<tr><td width=100><a href='buy.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='����"&rst("����")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&rst("����")&"</td><td> "&rst("����")&"</td><td> "&rst("����")&"</td><td> "&rst("��Ч")&"</td><td> "&rst("�۸�")&"</td></tr>"
	rst.MoveNext
loop
msg=msg+"</table>"
rst.Close
set rst=nothing 	
conn.Close 
set conn=nothing
%>
<head><LINK href="../chatroom/css.css" rel=stylesheet><title>������</title></head>
<body bgcolor='<%=bgcolor%>' background='../images/bg.gif' oncontextmenu=self.event.returnValue=false>
<div align=center>
<%=msgtmp%>
  <p align=center><b><font color=CC0000 face="��Բ" size="4"><%=Application("Ba_jxqy_systemname")%>������</font></b></p>
<%=msg%></div>
<p align=center>
  <input type="button" value=" �� �� " onClick="javascript:location.href='street.asp'" name="button"> 
  <input type="button" value=" �� �� " onClick="javascript:window.close();" name="button">
</p>
</body>