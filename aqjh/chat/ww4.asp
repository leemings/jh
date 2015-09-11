<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Application("aqjh_ww4")=0 then
	Response.Write "<script Language=Javascript>alert('提示：此物品已被领取，你只抢到了一把臭狗屎！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww4"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('提示：非法操作，拒绝执行！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww4"))))

Application.Lock
Application("aqjh_ww4")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update 用户 set 体力=体力-"&tempjs&" where 姓名='" & aqjh_name &"'"
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
if rs("体力")<-1000 then
	conn.execute "update 用户 set 状态='死',事件原因='大家|抢物品被打死' where 姓名='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'大家','抢物品被打死','人命')"
	kl="<img src='pic/kl.gif'>太不幸了，["&aqjh_name&"]在跟大家抢物品时被大家打死………"
	call boot(aqjh_name,"大家，操作者：大家，跟我抢东西，你找死啊！")
else
dim js(25)
js(0) ="吃豆卡"
js(1) ="补血卡"
js(2) ="练功卡"
js(3) ="离婚卡"
js(4) ="暴豆卡"
js(5) ="陷害卡"
js(6) ="清纯卡"
js(7) ="免税卡"
js(8) ="非礼卡"
js(9) ="抢亲卡"
js(10)="免罪卡"
js(11)="加点卡"
js(12)="真爱卡"
js(13)="会员卡"
js(14)="福神卡"
js(15)="财神卡"
js(16)="穷神卡"
js(17)="变性卡"
js(18)="衰神卡"
js(19)="送神卡"
js(20)="请神卡"
js(21)="归来卡"
js(22)="分手卡"
js(23)="清醒卡"
js(24)="羞羞卡"
randomize()
myxy = Int(Rnd*25)
s=Int(Rnd*1000)
zstemp=add(rs("w5"),js(myxy),1)
conn.execute "update 用户 set w5='"&zstemp&"' where 姓名='" & aqjh_name & "'"
kl="别看<b>"&aqjh_name&"</b>平时做事慢腾腾的，有好处的时候他最快！你看，一下子抢得会员卡片[<b>"&js(myxy)&"</font></b>1张，已揣进怀里！</font>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="【抢夺卡片】"&kl

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
