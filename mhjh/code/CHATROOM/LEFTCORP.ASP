<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
if un="" then Response.Redirect "../error.asp?id=016"
chatroomsn=Session("yx8_mhjh_userchatroomsn")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&un&"'",conn
nowcorp=rst("门派")
if nowcorp="无" then
leftcorp="虽然您勇气可嘉，但是实在无帮可叛"
elseif rst("体力")<=1000 or rst("内力")<=1000 or rst("等级")<2 then
leftcorp="虽然您勇气可嘉,但是这样可能要被乱棒打死哟"
else
conn.Execute "update 用户 set 体力=体力*0.8,内力=内力*0.8,精力=精力*0.8,积分=积分*0.8,等级=0,攻击=攻击*0.8,防御=防御*0.8,门派='无',身份='无' where 姓名='"&un&"'"
onlinelist=Application("yx8_mhjh_onlinelist")
onlinelistubd=ubound(onlinelist)
for i=1 to onlinelistubd step 8
if onlinelist(i)=un then
onlinelist(i+2)="无"
exit for
end if
next
Application.Lock
Application("yx8_mhjh_onlinelist")=onlinelist
Application.UnLock
erase onlinelist
session("yx8_mhjh_usercorp")="无"
session("yx8_mhjh_usergrade")=0
leftcorp="<font color=FF0000>【叛帮】</font>##离开了"&nowcorp&",等级降为0级，其余各项属性下降了两成！"
talkarr=Application("yx8_mhjh_talkarr")
talkpoint=clng(Application("yx8_mhjh_talkpoint"))
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
newtalkarr(592)=2
newtalkarr(593)=0
newtalkarr(594)=un
newtalkarr(595)="大家"
newtalkarr(596)=""
newtalkarr(597)="#660099"
newtalkarr(598)="#660099"
newtalkarr(599)="<font color=FF0000>【叛帮】</font>##离开了"&nowcorp&",等级降为0级，其余各项属性下降了两成！<font class=timsty>（"&time()&"）<\/font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
end if
rst.Close
set rst=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title></title>
<link rel=stylesheet href='css3.css'>
<script language=javascript>
setTimeout("location.href='onlinelist.asp'",3000);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false background='bg1.gif'>
<div align=center>
<font color=0000ff size=5></font>
<hr noshade size="1" color=red>
三秒钟自动返回<br>
<input type=button value='返回' onclick="javascript:location.href='onlinelist.asp'" id=button1 name=button1>
</div>
<%=leftcorp%>
</body>
</html>
