<%
Response.Expires=-1
chatroomsn=Session("Ba_jxqy_userchatroomsn")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
topmsg=Request.QueryString("topmsg")
if topmsg="" then topmsg="����"
toparr=array("����","����","����","����","����","��ʿ","��ħ","�ֽ�","���")
for i=0 to ubound(toparr)
	if topmsg=toparr(i) then
		toparrmsg=toparrmsg+"<td bgcolor=FFFF00><a href="&chr(34)&"top.asp?topmsg="&toparr(i)&chr(34)&" onmouseover=""window.status='TOP20��';return true;"" onmouseout=""window.status='';return true;"">"&toparr(i)&"</a></td>"
	else
		toparrmsg=toparrmsg+"<td><a href="&chr(34)&"top.asp?topmsg="&toparr(i)&chr(34)&" onmouseover=""window.status='TOP20��';return true;"" onmouseout=""window.status='';return true;"">"&toparr(i)&"</a></td>"
	end if
next	
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
if topmsg="��ħ" then
	sqlstr="select top 20 ����,����,����,����,����,����,����,��� from �û� order by ����"
	topmsg="����"	
else
	if topmsg="��ʿ" then topmsg="����"
	if topmsg="�ֽ�" then topmsg="����"
	sqlstr="select top 20 ����,����,����,����,����,����,����,����,��� from �û� order by "&topmsg&" desc"
end if	
set rst=server.CreateObject("adodb.recordset")
rst.Open sqlstr,conn
msg=msg+"<tr><td><font color=ff0000>����</font></td>"
for i=1 to rst.Fields.Count-1
		if rst.Fields(i).name=topmsg then
			msg=msg+"<td align=center bgcolor=FFFF00>"&rst.Fields(i).name&"</td>"
		else
			msg=msg+"<td align=center>"&rst.Fields(i).name&"</td>"
		end if	
next
msg=msg+"</tr>"
do while not (rst.EOF or rst.BOF)
	msg=msg+"<tr><td><font color=ff0000>"&rst("����")&"</font></td>"
	for i=1 to rst.Fields.Count-1
	if rst.Fields(i).name=topmsg then
			msg=msg+"<td bgcolor=FFFF00>"&rst.Fields(i).value&"</td>"
		else
			msg=msg+"<td>"&rst.Fields(i).value&"</td>"
		end if	
	next
	msg=msg+"</tr>"
	rst.Movenext
loop
rst.Close
set rst=nothing
conn=close
set conn=nothing
%>
<head><title><%=Application("Ba_jxqy_systemname")%>Ӣ�����а�</title><LINK href="../chatroom/css.css" rel=stylesheet></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=3 bordercolordark=FFFFFF bgcolor=FFB442><tr align=center><%=toparrmsg%></tr></table>
<p align=center><font color="#FF0000" size="4" face="��Բ"><b>Ӣ�����а�</b></font></p>
<table width=100% border="1" bordercolorlight=000000 cellspacing=0 cellpadding=3 bordercolordark=FFFFFF>
<%=msg%>
</table>
<p align=center>
  <input type="button" value=" �� �� " onClick="javascript:top.window.close();" name="button">
</p>
</body>