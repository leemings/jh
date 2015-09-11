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
if Application("aqjh_kl")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不是没事找事呀，哪里有舞狮？！');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("aqjh_kl"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	response.end
end if
Application.Lock
Application("aqjh_kl")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update 用户 set 体力=体力-"&tempjs&" where 姓名='" & aqjh_name &"'"
rs.open "select 体力,w6 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
if rs("体力")<-100 then
	conn.execute "update 用户 set 状态='死',事件原因='狮子|"&fn1&"' where 姓名='" & aqjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'舞女','泡舞女被打死','人命')"
	kl="<img src='img/mm.gif'>哼！，就["&aqjh_name&"]你这体力也想狮子？没门！"
	call boot(aqjh_name,"狮子，操作者：狮子，没力气？没门！")
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
conn.execute "update 用户 set 魅力=魅力+1000,银两=银两+"& tempjs*40 &",w6='"&zstemp&"' where 姓名='" & aqjh_name & "'"
kl="恭喜["&aqjh_name&"]狮子给你舞好了，增加魅力<img src='img/heart.gif'>1000点！银两<img src='img/251.gif'>"&tempjs*40&"两！得到了[<b><font color=red>"&js(myxy)&"</font></b>]1个……"
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
