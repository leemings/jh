<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Session("yx8_mhjh_userchatroomsn")
topmsg=Request.QueryString("topmsg")
if topmsg="" then topmsg="体力"
toparr=array("积分","体力","内力","攻击","防御","侠士","恶魔","杀人","被杀","现金","存款")
for i=0 to ubound(toparr)
if topmsg=toparr(i) then
toparrmsg=toparrmsg+"<td bgcolor=FFFF00><a href="&chr(34)&"top.asp?topmsg="&toparr(i)&chr(34)&" onmouseover=""window.status='TOP20榜';return true;"" onmouseout=""window.status='';return true;"">"&toparr(i)&"</a></td>"
else
toparrmsg=toparrmsg+"<td><a href="&chr(34)&"top.asp?topmsg="&toparr(i)&chr(34)&" onmouseover=""window.status='TOP20榜';return true;"" onmouseout=""window.status='';return true;"">"&toparr(i)&"</a></td>"
end if
next
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
if topmsg="恶魔" then
sqlstr="select top 20 头像,姓名,体力,内力,攻击,防御,杀人,被杀,道德,银两,存款 from 用户 order by 道德"

topmsg="道德"
else
if topmsg="侠士" then topmsg="道德"
if topmsg="现金" then topmsg="银两"
sqlstr="select top 20 头像,姓名,积分,体力,内力,攻击,防御,杀人,被杀,道德,银两,存款 from 用户 order by "&topmsg&" desc"
end if
set rst=server.CreateObject("adodb.recordset")
rst.Open sqlstr,conn
msg=msg+"<tr><td><font color=ff0000>头像</font></td>"
for i=1 to rst.Fields.Count-1
if rst.Fields(i).name=topmsg then
msg=msg+"<td align=center bgcolor=FFFF00>"&rst.Fields(i).name&"</td>"
else
msg=msg+"<td align=center>"&rst.Fields(i).name&"</td>"
end if
next
msg=msg+"</tr>"
do while not (rst.EOF or rst.BOF)
msg=msg+"<tr><td><font color=ff0000><img border='0' src='"&rst("头像")&"' width='21' height='21'></font></td>"
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
<head><title>英雄排行榜</title><LINK href="../chatroom/css.css" rel=stylesheet>
<base target="main">
</head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=3 bordercolordark=FFFFFF bgcolor=FFB442><tr align=center><%=toparrmsg%></tr></table>
<p align=center><img border="0" src="../image/top3.gif"></p>
<table width=100% border="1" bordercolorlight=000000 cellspacing=0 cellpadding=3 bordercolordark=FFFFFF>
<%=msg%>
</table>
<p align=center>
<input type="button" value=" 关 闭 " onClick="javascript:top.window.close();" name="button">
</p>
</body>
