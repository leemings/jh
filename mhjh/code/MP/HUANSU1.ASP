<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="select 银两,门派,道德 from 用户 where 姓名='"&username&"'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then Response.Redirect "../error.asp?id=016"
yl=rs("银两")
mp=rs("门派")
if yl<1000000 then
mess="你的钱不够，老实的出家吧，自己的路可是自己选的哦！"
elseif mp<>"出家" then
mess="你不是剑侠的和尚，哪里来的流氓，我们不欢迎你哦，下山吧！"
else
yl1=yl-1000000
dd=rs("道德")-300
mess="恭喜，你已经还俗了，自己的路自己走吧<br><br>请退出后重新登陆"
conn.execute "update 用户 set 银两='"&yl1&"',道德='"&dd&"',门派='无' where 姓名='"&username&"'"
end if
rs.Close             
set rs=nothing             
conn.Close             
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
		newtalkarr(599)="<font color=red>【和尚还俗】</font><font color=blue>"&username&"过不惯寺院里枯燥的生活，花费100万银子在买通了方丈,还俗了，看来江湖江湖中又要有一番腥风血雨了!</font>" 
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
<title>还俗</title>
<link rel="stylesheet" href="../style.css">
</head>
<body bgcolor=#000000 oncontextmenu=self.event.returnValue=false text="#FFFFFF">
<p align="center">　</p>
<table width="382" border="0" cellspacing="0" cellpadding="0" align="CENTER">
<tr>
<td height="249" width="472">
<table width="387" border="0" cellspacing="0" cellpadding="0">
<tr>
<td  width="1"></td>
<td width="391" align="center"><br>
<table width="367" border="0" cellspacing="0" cellpadding="0" align="center"  height="138">
<tr align="center">
<td height="67" width="365"><%=mess%></td>
</tr>
<tr align="center">
<td width="365"><a href="huansu.asp">返回</a></td>
</tr>
</table>
</td>
<td width="1" ></td>
</tr>
</table>
</td>
</tr>
</table>
