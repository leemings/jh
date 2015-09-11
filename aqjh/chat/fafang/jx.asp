<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
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
if Application("aqjh_jx")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不是没事找事呀，哪里有巨蝎？！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_jx"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	response.end
end if
Application.Lock
Application("aqjh_jx")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update 用户 set 体力=体力-"&tempjs&" where 姓名='" & aqjh_name &"'"
rs.open "select 体力,w6 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
if rs("体力")<-100 then
	conn.execute "update 用户 set 状态='死',事件原因='巨蝎|"&fn1&"' where 姓名='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'僵尸','因勇打巨蝎所牺牲','人命')"
	kl="<img src='img/jx.gif'>太不幸了，["&aqjh_name&"]为了保护大家的安全，舍身取义，被巨蝎尾巴扎了一下壮烈牺牲了………"
	call boot(aqjh_name,"巨蝎，操作者：巨蝎，打我呀找死！")
else
dim js(7)
js(0) ="骨头"
js(1) ="蓝宝石"
js(2) ="红宝石"
js(3) ="僵尸血"
js(4) ="僵尸牙"
js(5) ="水晶石"
js(6) ="绿宝石"
randomize()
myxy = Int(Rnd*7)
zstemp=add(rs("w6"),js(myxy),1)
conn.execute "update 用户 set 银两=银两+"& tempjs*1 &",w6='"&zstemp&"' where 姓名='" & aqjh_name & "'"
kl="["&aqjh_name&"]为众降魔伏妖，三下五除二便将"&gwm&gwtp&"杀死，真乃豪杰是也。官府奖励其<img src='img/251.gif'>"&tempjs*30&"两！并得到了[<b><font color=red>"&js(myxy)&"</font></b>]1个……"
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