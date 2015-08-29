<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/sjfunc.asp"-->

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
if Application("sjjh_kl")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不是没事找事呀，哪里有鬼？！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("sjjh_kl"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	response.end
end if
Application.Lock
Application("sjjh_kl")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
conn.execute "update 用户 set 体力=体力-"&tempjs&" where 姓名='" & sjjh_name &"'"
rs.open "select 体力,w6 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn
if rs("体力")<-60000 then
	conn.execute "update 用户 set 状态='僵尸王',身份='僵尸王',事件原因='僵尸|"&fn1&"' where 姓名='" & sjjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'僵尸','因勇打僵尸所牺牲','人命')"
	kl="<bgsound src=wav/diao.wav loop=1><img src='img/kl.gif'>太不幸了，["&sjjh_name&"]为了保护大家的安全，舍身取义，被僵尸咬的七零八落，死翘翘了变成了一只僵尸王从此三国江湖多事了.各位大虾还是小心点………"
	call boot(sjjh_name,"僵尸，操作者：僵尸，打我呀找死！")
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
conn.execute "update 用户 set 银两=银两+"& tempjs*1 &",w6='"&zstemp&"' where 姓名='" & sjjh_name & "'"
kl="我call,英雄，真是英雄，["&sjjh_name&"]与僵尸大打出手……僵尸终于倒下了，而他自己也伤的不轻，官府奖励<img src='img/251.gif'>"&tempjs*1&"两！在打败僵尸后，"&sjjh_name&"得到了[<b><font color=red>"&js(myxy)&"</font></b>]1个……"

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
