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
if Application("sjjh_lw")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子来晚了，卡片被别人捡走了！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("sjjh_lw"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	response.end
end if
Application.Lock
Application("sjjh_lw")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
conn.execute "update 用户 set 体力=体力+"&tempjs&" where 姓名='" & sjjh_name &"'"
rs.open "select 体力,w5 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
if rs("体力")<-10000 then
	conn.execute "update 用户 set 状态='死',事件原因='抢礼物|"&fn1&"' where 姓名='" & sjjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'站长发礼物','抢礼物不成被人踩死','人命')"
	kl="<img src='img/kl.gif'>哈哈，["&sjjh_name&"]急着抢礼物，但是不小心，结果被人给踩死了………"
	call boot(sjjh_name,"急着抢礼物，但是不小心，结果被人给踩死了！")
else
dim js(10)
js(0) ="清醒卡"
js(1) ="补血卡"
js(2) ="解除卡"
js(3) ="陷害卡"
js(4) ="清纯卡"
js(5) ="请神卡"
js(6) ="送神卡"
js(7) ="吃豆卡"
js(8) ="练功卡"
js(9) ="暴豆卡"
randomize()
myxy = Int(Rnd*10)
zstemp=add(rs("w5"),js(myxy),1)
conn.execute "update 用户 set 银两=银两+"& tempjs*30 &",w5='"&zstemp&"' where 姓名='" & sjjh_name &"'"
kl="站长发礼物来了["&sjjh_name&"]跑的最快，抢在大家前面<img src='img/lw.gif'>把礼物领走了，"&sjjh_name&"得到站长送出的[<b><font color=red>"&js(myxy)&"</font></b>]1张……"
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