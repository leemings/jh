<%
Response.Expires=-1
chatroomsn=Session("yx8_mhjh_userchatroomsn")
topmsg=Request.QueryString("topmsg")
if topmsg="" then topmsg="介绍人"
toparr=array("积分","介绍人","内力","攻击","防御","侠士","恶魔","杀人","被杀","现金","存款")

set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
sqlstr="select top 20 头像,姓名,积分,介绍人,内力,攻击,防御,杀人,被杀,道德,银两,存款 from 用户 where order by 介绍人 desc"
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
<head><title><%=Application("yx8_mhjh_systemname")%>英雄排行榜</title><LINK href="../chatroom/css.css" rel=stylesheet></head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=3 bordercolordark=FFFFFF bgcolor=FFB442><tr align=center><%=toparrmsg%></tr></table>
<p align=center><b><font size="3" face="华文仿宋" color="#FF0000">拉人排行榜</font></b></p>
<table width=100% border="1" bordercolorlight=000000 cellspacing=0 cellpadding=3 bordercolordark=FFFFFF>
<%=msg%>
</table>
<p align=center>
<input type="button" value=" 关 闭 " onClick="javascript:top.window.close();" name="button">
</p>
</body>
