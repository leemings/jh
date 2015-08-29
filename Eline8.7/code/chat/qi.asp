<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Application("sjjh_qi")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不是没事找事呀，哪里有抢？！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("sjjh_qi"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	response.end
end if
Application.Lock
Application("sjjh_qi")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
conn.execute "update 用户 set 体力=体力-"&tempjs&" where 姓名='" & sjjh_name &"'"
rs.open "select 体力,w9 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn
if rs("体力")<-100 then
	conn.execute "update 用户 set 状态='死',事件原因='手枪|"&fn1&"' where 姓名='" & sjjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'手枪','因勇打精制手枪所牺牲','人命')"
	kl="<img src='img/tank.gif'>太不幸了，["&sjjh_name&"]为了保护大家的安全，舍身取义，被手枪扫射的七零八落，死翘翘了………"
	call boot(sjjh_name,"手枪，操作者：手枪，打我呀找死！")
else
dim js(13)
js(0) ="子弹"
js(1) ="魔幻水晶"
js(2) ="狼牙棒"
js(3) ="破天锥"
js(4) ="抢劫令"
js(5) ="血滴子"
js(6) ="魔力钻石"
js(7) ="生日蛋糕"
js(8) ="紧箍咒"
js(9)="天梦魂"
js(10)="精制手枪"
js(11)="弹药"
js(12)="无影水晶"
randomize()
myxy = Int(Rnd*13)
zstemp=add(rs("w9"),js(myxy),1)
conn.execute "update 用户 set 银两=银两+"& tempjs*10 &",w9='"&zstemp&"' where 姓名='" & sjjh_name & "'"
kl="我call,英雄，真是英雄，["&sjjh_name&"]与手枪为敌抗击手枪……手枪终于敌不过中原武功，而他自己也伤的不轻，官府奖励<img src='img/251.gif'>"&tempjs*15&"两！在打败手枪后，"&sjjh_name&"得到了[<b><font color=red>"&js(myxy)&"</font></b>]1个……"

end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>【江湖消息】</b></font>"&kl

says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
