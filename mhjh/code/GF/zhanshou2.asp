<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
usercorp=session("yx8_mhjh_usercorp")
usergrade=session("yx8_mhjh_usergrade")
mess="官府的人喊到：吉时已到，人犯"&Request("name")&"斩立决!!!!"
if usercorp="官府" and usergrade>=lockipright then
chatroomsn=session("yx8_mhjh_userchatroomsn")
shi=0.0416*1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="update 用户 set 最后登录时间=now()+"&shi&",状态='死亡',等级=等级-2,体力=体力-50000,攻击=攻击-20000,防御=防御-10000,被杀=被杀+1 WHERE 姓名='"&Request("name")&"'"
conn.execute sql
sql="insert into 英烈堂(死者,时间,凶手,死因) values ('"&Request("name")&"',now(),'"&username&"','牢房被斩')"
conn.execute sql
sql="insert into 聊务记录(管理类型,管理人员,被管理人,管理原因,管理时间,管理房间) values ('斩首','"&username&"','"&Request("name")&"','牢房被斩',now(),'魔幻江湖')"
conn.execute sql
talkarr=Application("yx8_mhjh_talkarr")
talkpoint=Application("yx8_mhjh_talkpoint")
dim newtalkarr(600)
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
newtalkarr(599)="<font color=red>【消息】</font><b><font color=red>"&mess&"</font></b>"
newtalkarr(600)=chatroomsn
Application("yx8_mhjh_talkarr")=newtalkarr
end if
conn.Close             
set conn=nothing  
%>
<head>
<LINK
href="../style.css" rel=stylesheet>
</head>
<body background='../chatroom/bg1.gif'>
<div align="center">
  <center>
<table width=338 bordercolorlight="#000000" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF">
<tr><td width="330">
<p align=center style='font-size:14;color:red'><%=mess%></p>
<p align=center><a href="../STREET/PRISON.ASP">返回</a></p>
</td></tr>
</table>
  </center>
</div>
</body>
