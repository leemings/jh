<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
myname=session("yx8_mhjh_username")
if myname="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>查看你的拉人情况</title>
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
<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false>
<div align="center">
<table border="0" width="390" cellspacing="0" cellpadding="0" align="center" height="22" bgcolor="#000000" >
<tr><center>
<td width="388"><b><font color="#FFFF00">说明</font></b></center></tr>
<tr>
<td width="388">1.你要介绍至少5个人加入我们的江湖,并且介绍的每个人的泡分必须大于20000</tr>
<tr>
<td width="388">2.如何拉人:你介绍一个人加入我们的剑侠时,让对方在用户注册时中,介绍人一栏填上你的名字才有效!否则一律不算数!</tr>
<tr>
<td width="388">3.拉人要点:为了防止自己泡号作弊,只有所拉的人泡分要大于20000分才有效</tr>
</table>
<%
dim sql
dim rs
dim filename
dim hx
sql="select * from 用户 where 积分>=20000 and 介绍人='"&myname&"'"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
sj=rs.recordcount
sql="select * from 用户 where  介绍人='"&myname&"' order by 积分 desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof or  rs.bof then
response.write "你还没拉一个人呀,还需努力呀!"
else
%>
<table border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" height="1">
<tr bgcolor="#9C3000">
<td align="center" height="1" width="188"><font  color=white>姓 名</font></td>
<td align="center" height="1" width="188"><font  color=white>泡 分</font></td>
</tr>
<%do while not rs.eof%>
<tr>
<td align="center" width="188" height="1"><font color=ff0000><%=rs("姓名")%></font>
</td>
<td align="center" width="188" height="1"><font color=#ffff00><%=rs("积分")%></font>
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
response.write "当前你共介绍<font color=#ffff00>"&tj&"</font>人加入我们的江湖,其中泡分大于20000的有<font color=#ffff00>"&sj&"</font>人,还差<font color=#ffff00>"&tjy&"</font>人就能申请拉人奖励200万银两和50万精力了!"
else
response.write "当前你共介绍<font color=#ffff00>"&tj&"</font>人加入我们的江湖,其中泡分大于20000的有<font color=#ffff00>"&sj&"</font>人!恭喜你,你已经可以申请拉人奖励200万银两和50万精力了!你现在要做的事就是<a href='lingj.asp'><font color=#ffff00>领取奖金</font></a>"
end if%>
</table>
　
</div>
</body>
</html>