<%
username=session("yx8_mhjh_username")
posr=session("yx8_mhjh_userposr")
posc=session("yx8_mhjh_userposc")
direction=Request.QueryString("direction")
if session("yx8_mhjh_userfight")<>"none" then
Response.Write "<script language=javascript>alert('遭到怪物拦截,无法移动,选择战斗或者逃跑吧');parent.confrm.document.form1.move.disabled=true;;parent.behfrm.location.replace('fight.asp');</script>"
Response.End
end if
if direction="" then Response.End
if not(isnumeric(posr) and isnumeric(posc)) then Response.End
select case direction
case "tl"
posr=posr-1
posc=posc-1
case "t"
posr=posr-1
case "tr"
posr=posr-1
posc=posc+1
case "l"
posc=posc-1
case "r"
posc=posc+1
case "bl"
posr=posr+1
posc=posc-1
case "b"
posr=posr+1
case "br"
posr=posr+1
posc=posc+1
case else
Response.End
end select
session("yx8_mhjh_userfight")="none"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from map where posr="&posr&" and posc="&posc&" and displace=True",conn
if rst.EOF or rst.BOF then
msg="<script language=javascript>parent.msgfrm.document.writeln('<FONT color=#ff0000>【错误】</FONT>对不起，此处无法通过<br>');</script>"
else
session("yx8_mhjh_userposr")=posr
session("yx8_mhjh_userposc")=posc
Response.Write "<script language=javascript>parent.mapfrm.gotorc("&posr&","&posc&");</script>"
action=split(rst("affair")," ")
rst.close 
rst.Open "select * from biological ",conn
randomize          
r=20 
id=int((r-1)*rnd)+1 
rst.close 
set rst=nothing 
set rst = server.createobject("adodb.recordset") 
rst.open "select * from biological where id="&id,conn
if rst.EOF or rst.BOF then
msg="<script language=javascript>location.href='action.asp'</script>"
else
session("yx8_mhjh_userfight")=rst("biological")
rst.Close
select case action(0)
case "fight"
session("yx8_mhjh_userfight")=action(1)
msg="<script language=javascript>alert('遭到"&action(1)&"骚扰');parent.confrm.document.form1.move.disabled=true;parent.behfrm.location.replace('fight.asp');</script>"
case "exit"
msg="<script language=javascript>alert('遭到骚扰');parent.confrm.document.form1.move.disabled=true;parent.behfrm.location.replace('fight.asp');</script>"
case "msg"
msg="<script language=javascript>alert('【冒险消息】"&action(1)&"<br>');parent.msgfrm.location.replace('0.htm');</script>"
case "gd"
msg="<script language=javascript>alert('【冒险消息】"&action(1)&"<br>');parent.msgfrm.location.replace('2.htm');</script>"
case "wb"
msg="<script language=javascript>alert('【冒险消息】"&action(1)&"<br>');parent.msgfrm.location.replace('1.htm');</script>"
case "cb"
msg="<script language=javascript>alert('【冒险消息】"&action(1)&"<br>');parent.msgfrm.location.replace('3.htm');</script>"
case "ck"
msg="<script language=javascript>alert('【冒险消息】"&action(1)&"<br>');parent.msgfrm.location.replace('4.htm');</script>"
case "fb"
msg="<script language=javascript>alert('【冒险消息】"&action(1)&"<br>');parent.msgfrm.location.replace('5.htm');</script>"
case "dy"
msg="<script language=javascript>alert('【冒险消息】"&action(1)&"<br>');parent.msgfrm.location.replace('6.htm');</script>"
case "tx"
msg="<script language=javascript>alert('【冒险消息】"&action(1)&"<br>');parent.msgfrm.location.replace('7.htm');</script>"
case "xz"
msg="<script language=javascript>alert('【冒险消息】"&action(1)&"<br>');parent.msgfrm.location.replace('8.htm');</script>"
end select
end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg
%>
