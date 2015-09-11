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
if Application("aqjh_ww2")=0 then
	Response.Write "<script Language=Javascript>alert('提示：此物品已被领取，你只抢到了一把臭狗屎！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww2"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('提示：非法操作，拒绝执行！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww2"))))

Application.Lock
Application("aqjh_ww2")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update 用户 set 体力=体力-"&tempjs&" where 姓名='" & aqjh_name &"'"
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
if rs("体力")<100 then
	conn.execute "update 用户 set 状态='死',事件原因='大家|抢物品被打死' where 姓名='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'大家','抢物品被打死','人命')"
	kl="<img src='pic/kl.gif'>太不幸了，["&aqjh_name&"]在跟大家抢物品时被大家打死………"
	call boot(aqjh_name,"大家，操作者：大家，跟我抢东西，你找死啊！")
else
dim js(20)
js(0) ="鹤顶红"
js(1) ="泻药"
js(2) ="含笑半步癫"
js(3) ="迷香"
js(4) ="五毒丸"
js(5) ="七步倒"
js(6) ="五步倒"
js(7) ="一步倒"
js(8) ="酷鱼胆"
js(9)="砒霜"
js(10)="蛇毒"
js(11)="迷药"
js(12)="鹤顶红"
js(13)="迷药"
js(14)="蛇毒"
js(15)="一步倒"
js(16) ="砒霜"
js(17) ="五毒丸"
js(18) ="含笑半步癫"
js(19) ="泻药"
randomize()
myxy = Int(Rnd*18)
s=Int(Rnd*1000)
zstemp=add(rs("w2"),js(myxy),1)
conn.execute "update 用户 set w2='"&zstemp&"' where 姓名='" & aqjh_name & "'"
kl="<font color=green>别看<b>"&aqjh_name&"</b>平时做事慢腾腾的，有好处的时候他最快！毒药[<b>"&js(myxy)&"</b>1个，已被他揣进怀里！</font>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>【抢夺毒药】</b></font>"&kl

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
