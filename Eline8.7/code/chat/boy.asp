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
if Application("sjjh_kl")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子来晚了，衰哥被别人打跑了！');</script>"
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
rs.open "select 体力,w6 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("体力")<-100 then
	conn.execute "update 用户 set 状态='死',事件原因='衰哥|"&fn1&"' where 姓名='" & sjjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'衰哥','因勇打衰哥所牺牲','人命')"
	kl="<bgsound src=wav/diao.wav loop=1><img src='img/kl.gif'>哈哈，["&sjjh_name&"]拿着棒子来打衰哥，谁知道体力不如人家，["&sjjh_name&"]挨了衰哥一招晴空霹雳，死翘翘了………"
	call boot(sjjh_name,"衰哥，操作者：衰哥，敢打我，找死！")
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
conn.execute "update 用户 set 银两=银两+"& tempjs*30 &",w6='"&zstemp&"' where 姓名='" & sjjh_name & "'"
kl="哈哈，["&sjjh_name&"]拿着棒子打跑了来江湖泡美女的衰哥，得到了官府奖励<img src='img/251.gif'>"&tempjs*30&"两！在打败僵尸后，"&sjjh_name&"得到了[<b><font color=red>"&js(myxy)&"</font></b>]1个……"
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
