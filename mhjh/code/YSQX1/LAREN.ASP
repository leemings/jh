<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
uname=Request("uname")
if uname="" then Response.Redirect "../error.asp?id=024"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>查看你的拉人情况</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body text="#FFFFFF" oncontextmenu=self.event.returnValue=false background="../chatroom/bg1.gif">
<div align="center">
<table border="0" width="390" cellspacing="0" cellpadding="0" align="center" height="22"  >
<tr><center>
<td width="388"><b><font color="#000000">说明</font></b></center></tr>
<tr>
<td width="388"></tr>
</table>
<%
dim sql
dim rs
dim filename
dim hx
sql="select * from 用户 where 积分>=10000 and 介绍人='"&uname&"'"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
sj=rs.recordcount
sql="select * from 用户 where  介绍人='"&uname&"' order by 积分 desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof or  rs.bof then
response.write "还没拉一个人"
else
%>
<table border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" height="1">
<tr>
<td align="center" height="1" width="188"><font color="#808080">姓 名</font></td> 
<td align="center" height="1" width="188"><font color="#808080">泡 分</font></td> 
</tr> 
<%do while not rs.eof%> 
<tr> 
<td align="center" width="188" height="1"><%=rs("姓名")%> 
</td> 
<td align="center" width="188" height="1"><%=rs("积分")%> 
</td> 
</tr> 
<%rs.movenext 
loop 
tj=rs.recordcount 
end if 
rs.Close 
set rs=nothing 
tjy=5-sj 
if sj<5 then 
response.write "当前共介绍"&tj&"人加入江湖,其中泡分大于10000的有"&sj&"人,还差"&tjy&"人就能申请拉人奖励200万银两了和够条件加入官府了!" 
else 
response.write "当前共介绍"&tj&"人加入的江湖,其中泡分大于10000的有"&sj&"人!恭喜你,你已经可以申请拉人奖励200万银两了和够条件加入官府了!你现在要做的事就是" 
end if%> 
</table> 
　 
</div> 
<%
Response.Write "<p align=center><a href="&chr(34)&"javascript:location.replace('showuser.asp?search="&uname&"');"&chr(34)&">返回</a></p>"
%>
</body> 
</html> 
