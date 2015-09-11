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
if Application("aqjh_ww6")=0 then
	Response.Write "<script Language=Javascript>alert('提示：此物品已被领取，你只抢到了一把臭狗屎！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww6"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('提示：非法操作，拒绝执行！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_ww6"))))

Application.Lock
Application("aqjh_ww6")=0
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
dim js(21)
js(0) ="节日献礼"
js(1) ="健康永远"
js(2) ="情意浓浓"
js(3) ="幸福泉源"
js(4) ="美满幸福"
js(5) ="浓情似火"
js(6) ="高风亮节"
js(7) ="百年好合"
js(8) ="此情不渝"
js(9) ="梦中情人"
js(10)="温馨的爱"
js(11)="灿烂人生"
js(12)="此情可待"
js(13)="真情歉意"
js(14)="幸福安康"
js(15)="粉色舞蹈"
js(16)="水晶之恋"
js(17)="爱的诉说"
js(18)="风情玫瑰"
js(19)="快乐时光"
js(20)="情定一生"
js(21)="两情相悦"
randomize()
myxy = Int(Rnd*21)
s=Int(Rnd*1000)
s1=Int(Rnd*3)+1
zstemp=add(rs("w7"),js(myxy),s1)
conn.execute "update 用户 set w7='"&zstemp&"' where 姓名='" & aqjh_name & "'"
kl="哇！怎么这么走运啊？<b>"&aqjh_name&"</b>这花都能捡到，有了这花，泡MM帅哥，不信没有不上钩的：）捡了<b>"&js(myxy)&"</b>"&s1&"朵，真是幸运！"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<b>【动感鲜花】</b>"&kl

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
