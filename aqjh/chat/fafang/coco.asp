<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/func.asp"-->
<!--#include file="../../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if aqjh_grade>5 then
Response.Write "<script Language=Javascript>alert('提示：你是官府中人！别玩了！');</script>"
	response.end
end if
if Application("aqjh_coco")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不是没事找事呀，哪里有人妖？！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_coco"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 身份 from 用户 where 姓名='"&aqjh_name&"'",conn
if rs("身份")="掌门" then 
           rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你是掌门没有资格打妖怪掌门只有资格放妖怪');}</script>"
	Response.End
end if
Application.Lock
Application("aqjh_coco")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update 用户 set 内力=内力-"&tempjs&" where 姓名='" & aqjh_name &"'"
rs.open "select 内力,w6 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
if rs("内力")<-5000 then
	conn.execute "update 用户 set 状态='死',事件原因='人妖|"&fn1&"' where 姓名='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'人妖','因勇打人妖所牺牲','人命')"
	kl="<img src='pic/kl.gif'>太不幸了，["&aqjh_name&"]为了保护江湖人的安全，舍身取义，被人妖打的七零八落，死翘翘了………"
	call boot(aqjh_name,"人妖，操作者：人妖，打我呀找死！")
else
dim js(10)
js(0) ="骨头"
js(1) ="蓝宝石"
js(2) ="红宝石"
js(3) ="蓝宝石"
js(4) ="绿宝石"
js(5) ="水晶石"
js(6) ="绿宝石"
js(7) ="硫磺"
js(8) ="硝酸"
js(9)="木炭"
randomize()
myxy = Int(Rnd*10)
zstemp=add(rs("w6"),js(myxy),1)
conn.execute "update 用户 set 法力=法力+"& tempjs*1 &",w6='"&zstemp&"' where 姓名='" & aqjh_name & "'"
kl="英雄降世，超级英雄，["&aqjh_name&"]与人妖抗击……桶妖终于敌不过中原武功，而他自己也伤的不轻，官府奖励"&tempjs*1&"法力！在打败人妖后，"&aqjh_name&"得到了[<b><font color=red>"&js(myxy)&"</font></b>]1个……"
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
