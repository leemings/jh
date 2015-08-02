<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Request.QueryString("chatroomsn")
if chatroomsn="" then chatroomsn=1
if not isnumeric(chatroomsn) then chatroomsn=1
chatroomsn=clng(chatroomsn)
chatroomname=Application("Ba_jxqy_systemname"&chatroomsn)
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
arrestright=Application("Ba_jxqy_arrestright")
gaolright=Application("Ba_jxqy_gaolright")
lockipright=Application("Ba_jxqy_lockipright")
unlockipright=Application("Ba_jxqy_unlockipright")
bombright=application("Ba_jxqy_bombright")
exaltgraderight=Application("Ba_jxqy_exaltgraderight")
declinegraderight=Application("Ba_jxqy_declinegraderight")
chatroomnum=Application("Ba_jxqy_chatroomnum")
lockiplist=Application("Ba_jxqy_lockip")
lockip=split(lockiplist,";")
lockipubd=ubound(lockip)
username=session("Ba_jxqy_username")
if username="" then Response.redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
mygrade=rst("等级")
mycorp=rst("门派")
myidentity=rst("身份")
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg="<head><title>"&Application("Ba_jxqy_systemname")&"</title><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>function init(){var chatsn='"&chatroomsn&"';var chatnum='"&chatroomnum&"';var i;for(i=0;i<chatnum;i++){if(document.form1.sele1.options[i].value==chatsn){document.form1.sele1.options[i].selected=true;}}}</script></head><body oncontextmenu=self.event.returnValue=false background='"&bgimage&"' bgcolor='"&bgcolor&"' onload=javascript:init();><p align=center><form name=form1><select name=sele1 onchange='javascript:location.replace("&chr(34)&"onlinelist.asp?chatroomsn="&chr(34)&"+document.form1.sele1.value);'>"
for i=1 to chatroomnum
	msg=msg&"<option value='"&i&"'>"&Application("Ba_jxqy_systemname"&i)&"</option>"
next
msg=msg&"</select></form></p>锁定地址列表："
if (mygrade>lockipright and mycorp="官府") then
	for i=1 to lockipubd-1
		msg=msg&"&nbsp;<a href='unlockip.asp?ip="&lockip(i)&"&chatroomsn="&chatroomsn&"' onmouseover="&chr(34)&"window.status='解锁此IP';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&lockip(i)&"</a>&nbsp;"
	next
else
	for i=1 to lockipubd-1
		msg=msg&"&nbsp;"&lockip(i)&"&nbsp;"
	next
end if
msg=msg&"<hr><table width=100% align=center border=3><tr align=center bgcolor=FFFF00><td>姓名</td><td>性别</td><td>门派</td><td>IP</td><td>操作</td></tr>"
onlinelist=Application("Ba_jxqy_onlinelist")
onlinelistubd=ubound(onlinelist)
for i=1 to onlinelistubd step 8
	if chatroomsn=onlinelist(i+5) then
	ip=split(onlinelist(i+3),".")
	if mygrade>=lockipright and mycorp="官府" then
		uip=ip(0)&"."&ip(1)&"."&ip(2)&"."&ip(3)
	elseif mygrade>=10 and mycorp="官府"  then
		uip=ip(0)&"."&ip(1)&"."&ip(2)&".XXX"
	elseif mygrade>=8 then
		uip=ip(0)&"."&ip(1)&".XXX.XXX"
	elseif mygrade>=5 then
		uip=ip(0)&".XXX.XXX.XXX"
	else
		uip="XXX.XXX.XXX.XXX"
	end if
	opt="操作"
	if (mygrade>10 and mycorp="官府") or myidentity="掌门" then opt="<a href='aberrant.asp?chatroomsn="&chatroomsn&"&opt=驱逐&touser="&onlinelist(i)&"'>驱逐</a>"
	if mygrade>=arrestright and mycorp="官府" then opt=opt&" | <a href='aberrant.asp?chatroomsn="&chatroomsn&"&opt=逮捕&touser="&onlinelist(i)&"'>逮捕</a>"
	if mygrade>=gaolright and mycorp="官府" then opt=opt&" | <a href='aberrant.asp?chatroomsn="&chatroomsn&"&opt=入狱&touser="&onlinelist(i)&"'>入狱</a>"
	if (mygrade>=lockipright and mycorp="官府") then opt=opt&" | <a href='aberrant.asp?chatroomsn="&chatroomsn&"&opt=锁定&touser="&onlinelist(i)&"'>锁定</a>"
	if (mygrade>=bombright and mycorp="官府") then opt=opt&" | <a href='aberrant.asp?chatroomsn="&chatroomsn&"&opt=炸弹&touser="&onlinelist(i)&"'>炸弹</a>"
	msg=msg&"<tr><td>"&onlinelist(i)&"</td><td>"&onlinelist(i+1)&"</td><td>"&onlinelist(i+2)&"</td><td>"&uip&"</td><td>"&opt&"</td></tr>"
	end if
next
msg=msg&"</table></body>"
Response.Write msg
%>