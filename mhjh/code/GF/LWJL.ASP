<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Request.QueryString("chatroomsn")
if not isnumeric(chatroomsn) then chatroomsn=1
chatroomname=Application("yx8_mhjh_systemname"&chatroomsn)
chatroomnum=Application("yx8_mhjh_chatroomnum")
username=session("yx8_mhjh_username")
if username="" then Response.redirect "../error.asp?id=016"
msg1="<head><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><title>聊务公开</title><link rel='stylesheet' href='../style.css'><script language=javascript>function init(){var chatsn='"&chatroomsn&"';var chatnum='"&chatroomnum&"';var i;for(i=0;i<chatnum;i++){if(document.form1.sele1.options[i].value==chatsn){document.form1.sele1.options[i].selected=true;}}}</script></head><body oncontextmenu=self.event.returnValue=false  background='../chatroom/bg1.gif' onload=javascript:init();><p align=center><form name=form1><select name=sele1 onchange='javascript:location.replace("&chr(34)&"lwjl.asp?topmsg=驱逐&chatroomsn="&chr(34)&"+document.form1.sele1.value);'>"
for i=1 to chatroomnum
msg1=msg1&"<option value='"&i&"'>"&Application("yx8_mhjh_systemname"&i)&"</option>"
next
msg1=msg1&"</table></body>"
Response.Write msg1
topmsg=Request.QueryString("topmsg")
toparr=array("驱逐","逮捕","入狱","斩首","罚款","刑罚","点穴","封咒","发钱","锁定","炸弹")
for i=0 to ubound(toparr)
if topmsg=toparr(i) then
toparrmsg=toparrmsg+"<td bgcolor=FFFF00><a href="&chr(34)&"lwjl.asp?chatroomsn="&chatroomsn&"&topmsg="&toparr(i)&chr(34)&" onmouseover=""window.status='保留最新50条';return true;"" onmouseout=""window.status='';return true;"">"&toparr(i)&"</a></td>"
else
toparrmsg=toparrmsg+"<td><a href="&chr(34)&"lwjl.asp?chatroomsn="&chatroomsn&"&topmsg="&toparr(i)&chr(34)&" onmouseover=""window.status='保留最新50条';return true;"" onmouseout=""window.status='';return true;"">"&toparr(i)&"</a></td>"
end if
next
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
sqlstr="select top 50 管理人员,被管理人,管理原因,管理时间 from 聊务记录 where 管理房间='"&chatroomname&"' and 管理类型='"&topmsg&"' order by 管理时间 desc"
set rst=server.CreateObject("adodb.recordset")
rst.Open sqlstr,conn
msg=msg+"<tr><td><font color=ff0000>管理人员</font></td>"
for i=1 to rst.Fields.Count-1
if rst.Fields(i).name=topmsg then
msg=msg+"<td align=center bgcolor=FFFFff>"&rst.Fields(i).name&"</td>"
else
msg=msg+"<td align=center>"&rst.Fields(i).name&"</td>"
end if
next
msg=msg+"</tr>"
do while not (rst.EOF or rst.BOF)
msg=msg+"<tr><td><font color=ff0000>"&rst("管理人员")&"</font></td>"
for i=1 to rst.Fields.Count-1
if rst.Fields(i).name=topmsg then
msg=msg+"<td bgcolor=FFFFff>"&rst.Fields(i).value&"</td>"
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
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=3 bordercolordark=FFFFFF bgcolor=FFFFff><tr align=center><%=toparrmsg%></tr></table>
<p align=center><font color="#FF0000" size="4" face="幼圆"><b><%=chatroomname%><%=topmsg%>记录</b></font></p>
<table width=100% border="1" bordercolorlight=000000 cellspacing=0 cellpadding=3 bordercolordark=FFFFFF>
<%=msg%>
</table>
<p align=center>
　
</p>

