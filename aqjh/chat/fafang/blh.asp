<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/func.asp"-->
<!--#include file="../../mywp.asp"-->
<!--#include file="../../chk.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
call chkpost()
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');</script>"
	Response.End
end if
gwm="白老虎"
gwtp="<img src=img/blh.gif>"
if Application("aqjh_blh")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子来晚了，"&gwm&"被别人杀死了！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_blh"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	response.end
end if
Application.Lock
Application("aqjh_blh")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update 用户 set 体力=体力-"&tempjs&" where 姓名='" & aqjh_name &"'"
rs.open "select 体力,w1,w2,w3,w4,w5,w6,w7,w8 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
if rs("体力")<-100 then
	conn.execute "update 用户 set 状态='死',事件原因='"&gwm&"|"&fn1&"' where 姓名='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'"&gwm&"','因勇打"&gwm&"所牺牲','人命')"
	kl=gwtp&"["&aqjh_name&"]去打"&gwm&"，谁知武功太差，结果成了"&gwm&"的美餐………"
	call boot(aqjh_name,gwm&"，操作者："&gwm&"，江湖小辈敢打我，自寻死路，好好练级吧！")
else
dim js(10)
js(0) ="骨头"
js(1) ="蓝宝石"
js(2) ="红宝石"
js(3) ="僵尸血"
js(4) ="僵尸牙"
js(5) ="水晶石"
js(6) ="绿宝石"
js(7) ="硫磺"
js(8) ="硝酸"
js(9)="木炭"
randomize()
myxy = Int(Rnd*10)
zstemp=add(rs("w6"),js(myxy),1)
conn.execute "update 用户 set 银两=银两+"& tempjs*1 &",w6='"&zstemp&"' where 姓名='" & aqjh_name & "'"
kl="["&aqjh_name&"]好生历害，三下五除二便将"&gwm&gwtp&"杀死，真乃当代武松。官府奖励其<img src='img/251.gif'>"&tempjs*30&"两！并得到了[<b><font color=red>"&js(myxy)&"</font></b>]1个……"
end if
kl="<font color=red><b>【江湖消息】</b></font>"&kl
says=kl
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="CCAA00"
saycolor="CCAA00"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>