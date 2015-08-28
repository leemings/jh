<%
username=session("yx8_mhjh_username")
mygrade=Session("yx8_mhjh_usergrade")
biological=session("yx8_mhjh_userfight")
if biological="none" then
msg="<FONT color=#ff0000>【操作错误】</FONT>没有东西阻击你,你慌什么,胆小鬼!<br>"
elseif username="" then
msg="<FONT color=#ff0000>【操作错误】</FONT>你没有登录或超时断开连接,请重新进入<br>"
else
randomize()
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 体力,防御 from 用户 where 姓名='"&username&"' and 体力>=100000",conn
if rst.EOF or rst.BOF then
msg="<FONT color=#ff0000>【操作错误】</FONT>你的体力眼看要消耗光了,没有力气逃跑,赶快使用千里传音到魔幻江湖去喊人来救你吧!<br>"
else
fhp=rst("体力")
fdefence=rst("防御")
rst.Close
rst.Open "select * from biological where biological='"&biological&"'",conn
maxsinew=rst("hp")
tattack=rst("attack")
tdefence=rst("defence")
encourage=rst("encourage")
encourage_m=Mid(encourage,1,2)
encourage_c=Mid(encourage,4)
defence=tattack-fdefence
Maxattack=Application("yx8_mhjh_Maxattack")*mygrade\50
randomize()
if defence>0 then
defence=Maxattack-clng(rnd()*1000)
elseif defence<=0 then
defence=clng(rnd()*10000)
end if
fhp=fhp-defence
if fhp<=0 then
conn.Execute "update 用户 set 状态='死亡',最后登录时间="&nowtimetype&" where 姓名='"&username&"'"
session.Abandon
Response.Write "<script language=javascript>top.location.replace('../error.asp?id=054');</script>"
msg="<FONT color=#ff0000>【发生惨案】</FONT>你被"&biological&"打死了!<br>"
else
conn.Execute "update 用户 set 体力="&fhp&",防御=防御+10 where 姓名='"&username&"'"
rndcatch=rnd()*99+1
if rndcatch<25 then
rst.Close
rst.Open "select * from pet where username='"&username&"' and exist=true"
if rst.EOF or rst.BOF then
msg="<FONT color=#ff0000>【逃跑失败】</FONT>你笨首笨脚,摔了一个跟头,"&biological&"的反击使您的生命下降"&defence&"<br>"
else
session("yx8_mhjh_userfight")="none"
msg="<FONT color=#ff0000>【逃跑成功】</FONT>你的随身宠物帮助了你,但"&biological&"的反击使您的生命下降"&defence&"<br>"
end if
else
msg="<FONT color=#ff0000>【逃跑失败】</FONT>"&biological&"的阻击使您的生命下降"&defence&"<br>"
end if
end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
end if
Response.Write "<script language=javascript>parent.msgfrm.document.writeln('"&msg&"');parent.confrm.document.form1.move.disabled=false;parent.behfrm.location.replace('Action.asp');</script>"
%>
