<!--#include file="../config.asp"--><%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Request.QueryString("chatroomsn")
if chatroomsn="" then chatroomsn=1
if not isnumeric(chatroomsn) then chatroomsn=1
chatroomsn=clng(chatroomsn)
chatroomname=Application("yx8_mhjh_systemname"&chatroomsn)
adminname=Application("yx8_mhjh_admin")
chatroomnum=Application("yx8_mhjh_chatroomnum")
lockiplist=Application("yx8_mhjh_lockip")
lockip=split(lockiplist,";")
lockipubd=ubound(lockip)
username=session("yx8_mhjh_username")
if username="" then Response.redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select �ȼ�,���� from �û� where ����='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
mygrade=rst("�ȼ�")
mycorp=rst("����")
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg="<head><title>"&Application("yx8_mhjh_systemname")&"</title><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>function init(){var chatsn='"&chatroomsn&"';var chatnum='"&chatroomnum&"';var i;for(i=0;i<chatnum;i++){if(document.form1.sele1.options[i].value==chatsn){document.form1.sele1.options[i].selected=true;}}}</script></head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif' onload=javascript:init();><p align=center><form name=form1><select name=sele1 onchange='javascript:location.replace("&chr(34)&"onlinelist.asp?chatroomsn="&chr(34)&"+document.form1.sele1.value);'>"
for i=1 to chatroomnum
msg=msg&"<option value='"&i&"'>"&Application("yx8_mhjh_systemname"&i)&"</option>"
next
msg=msg&"</select></form></p>������ַ�б�"
if (mygrade>lockipright and mycorp="�ٸ�") then
for i=1 to lockipubd-1
msg=msg&"&nbsp;<a href='unlockip.asp?ip="&lockip(i)&"&chatroomsn="&chatroomsn&"' onmouseover="&chr(34)&"window.status='������IP';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&lockip(i)&"</a>&nbsp;"
next
else
for i=1 to lockipubd-1
msg=msg&"&nbsp;"&lockip(i)&"&nbsp;"
next
end if
msg=msg&"<hr><table width=100% align=center border=3><tr align=center bgcolor=FFFF00><td>����</td><td>IP</td><td>�Ա�</td><td>����</td><td>����</td></tr>"
onlinelist=Application("yx8_mhjh_onlinelist")
onlinelistubd=ubound(onlinelist)
for i=1 to onlinelistubd step 6
if chatroomsn=onlinelist(i+5) then
opt="����"
if mygrade>=105 and mycorp="�ٸ�"  then opt="<a href='aberrant.asp?chatroomsn="&chatroomsn&"&opt=����&touser="&onlinelist(i)&"'>����</a>"
if mygrade>=106 and mycorp="�ٸ�" then opt=opt&" | <a href='aberrant.asp?chatroomsn="&chatroomsn&"&opt=����&touser="&onlinelist(i)&"'>����</a>"
if mygrade>=107 and mycorp="�ٸ�" then opt=opt&" | <a href='aberrant.asp?chatroomsn="&chatroomsn&"&opt=����&touser="&onlinelist(i)&"'>����</a>"
if mygrade>=108 and mycorp="�ٸ�" then opt=opt&" | <a href='aberrant.asp?chatroomsn="&chatroomsn&"&opt=����&touser="&onlinelist(i)&"'>����</a>"
if mygrade>=109 and mycorp="�ٸ�" then opt=opt&" | <a href='aberrant.asp?chatroomsn="&chatroomsn&"&opt=ը��&touser="&onlinelist(i)&"'>ը��</a>"
if adminname<>onlinelist(i) then
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
	sql="select ����¼IP from �û� where ����='"&onlinelist(i)&"'"
	set rst=conn.execute(sql)
	if not(rst.BOF or rst.EOF) then 
	do while not (rst.bof or rst.eof)
msg=msg&"<tr><td>"&onlinelist(i)&"</td><td>"&rst("����¼IP")&"</td><td>"&onlinelist(i+1)&"</td><td>"&onlinelist(i+2)&"</td><td>"&opt&"</td></tr>"
	rst.movenext
	loop
	end if
	rst.Close
set rst=nothing
conn.Close
set conn=nothing
	end if
end if
next
msg=msg&"</table></body>"
Response.Write msg
%>
