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
if Application("aqjh_ww5")=0 then
	Response.Write "<script Language=Javascript>alert('提示：此物品已被领取，你只抢到了一把臭狗屎！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww5"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('提示：非法操作，拒绝执行！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww5"))))

Application.Lock
Application("aqjh_ww5")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update 用户 set 体力=体力-"&tempjs&" where 姓名='" & aqjh_name &"'"
rs.open "select * FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
if rs("体力")<-100 then
	conn.execute "update 用户 set 状态='死',事件原因='大家|抢物品被打死' where 姓名='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'大家','抢物品被打死','人命')"
	kl="<img src='pic/kl.gif'>太不幸了，["&aqjh_name&"]在跟大家抢物品时被大家打死………"
	call boot(aqjh_name,"大家，操作者：大家，跟我抢东西，你找死啊！")
else
dim js(18)
js(0) ="玫瑰衷情"
js(1) ="一生有你"
js(2) ="比翼双飞"
js(3) ="热情奔放"
js(4) ="我爱我妻"
js(5) ="片片柔情"
js(6) ="永结同心"
js(7) ="恋爱热吻"
js(8) ="我的牵挂"
js(9) ="美满生活"
js(10)="浓情似火"
js(11)="千里相思"
js(12)="摇头的玫瑰"
js(13)="情深款款"
js(14)="可人儿"
js(15)="想你念你"
js(16)="爱的诉说"
js(17)="跳舞的玫瑰"
randomize()
myxy = Int(Rnd*18)
s=Int(Rnd*1000)
zstemp=add(rs("w7"),js(myxy),10)
conn.execute "update 用户 set w7='"&zstemp&"' where 姓名='" & aqjh_name & "'"
kl="别看<b>"&aqjh_name&"</b>平时做事慢腾腾的，有好处的时候他最快！鲜花<b>"&js(myxy)&"</b>10朵，已被他揣进怀里！"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<b>【捡取鲜花】</b>"&kl

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
