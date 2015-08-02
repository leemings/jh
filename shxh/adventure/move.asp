<%
username=session("Ba_jxqy_username")
posr=session("Ba_jxqy_userposr")
posc=session("Ba_jxqy_userposc")
direction=Request.QueryString("direction")
if username="" then 
	Response.Write "<script language=javascript>top.location.replace('../error.asp?id=016');</script>"
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
session("Ba_jxqy_userfight")="none"				
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from map where posr="&posr&" and posc="&posc&" and displace=True",conn
if rst.EOF or rst.BOF then 
	msg="<script language=javascript>parent.msgfrm.document.writeln('<FONT color=#ff0000>【错误】</FONT>对不起，此地形无法移动<br>');</script>"
else 
	session("Ba_jxqy_userposr")=posr
	session("Ba_jxqy_userposc")=posc
	Response.Write "<script language=javascript>parent.mapfrm.gotorc("&posr&","&posc&");</script>"
	action=split(rst("affair")," ")
	select case action(0)
		case "fight"
			session("Ba_jxqy_userfight")=action(1)
			msg="<script language=javascript>parent.msgfrm.document.writeln('<FONT color=#ff0000>【战斗】</FONT>这儿由"&action(1)&"守护，是否与之战斗?<br>');parent.behfrm.location.replace('fight.asp');</script>"
		case "exit"
			msg="<script language=javascript>alert('探险结束！');top.window.close();</script>"
			session("Ba_jxqy_userposr")=""
		case "msg"
			msg="<script language=javascript>parent.msgfrm.document.writeln('<FONT color=#ff0000>【消息】</FONT>"&action(1)&"<br>');</script>"
	end select	
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg
%>