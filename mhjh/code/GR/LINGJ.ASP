<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
myname=session("yx8_mhjh_username")
if myname="" then Response.Redirect "error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="select * from 用户 where 积分>=20000 and 介绍人='"&myname&"'"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
sj=rs.recordcount
if sj<5 then
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Redirect "../error.asp?id=509"
else
dim name(20)
for i = 1 to 5
name(i)=rs("姓名")
rs.MoveNext
conn.execute("Update 用户 set 介绍人='无' where 姓名='"&name(i)&"'")
next
conn.execute("Update 用户 set 银两=银两+2000000,精力=精力+500000 where 姓名='"&myname&"'")
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>【江湖通报】</font><b><font color=red>"&myname&"已经拉了5个积分在20000以上的人进江湖了，他获得官府特别奖励：200万的银两，50万精力!大家向他学习哦，努力拉人。为江湖贡献自己的一份力，也给自己带来好处多多！</font></b>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>领取拉人奖金</title>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<style type="text/css">.lh22px {
FONT-SIZE: 12px; LINE-HEIGHT: 22px
}
BODY {
FONT-SIZE: 9pt; LINE-HEIGHT: 18px
}
TD {
FONT-SIZE: 9pt; LINE-HEIGHT: 18px
}
A:visited {
COLOR: #0000a0; TEXT-DECORATION: none
}
A:link {
COLOR: #0000a0; TEXT-DECORATION: none
}
A:hover {
COLOR: #ff0000; TEXT-DECORATION: underline
}
.haody A:visited {
COLOR: #e9edfe; TEXT-DECORATION: none
}
.haody A:link {
COLOR: #e9edfe; TEXT-DECORATION: none
}
.haody A:hover {
COLOR: #ffffff; TEXT-DECORATION: underline
}
</style>
</head>
<body bgcolor=#000000 oncontextmenu=self.event.returnValue=false>
<table border=1 bgcolor="#948754" align=center width=287 cellpadding="10" cellspacing="13" height="182">
<tr><td bgcolor=#C6BD9B height="146" width="235">
<table height="123">
<tr><td height="37">
<font color="#000000"><strong>奖励!</strong></font>
<tr>
<td height="45" valign="top">由于你为剑侠拉来5位泡分超过20000的用户,特此奖励你200万的银两，50万精力!<Br><Br><center>
<p></p>
</center>
</td>
</tr>
</table>
</table>
</body>