<%@ LANGUAGE=VBScript%>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w6,操作时间 from 用户 where 姓名='"&sjjh_name&"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<10 then
	ss=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('请等上"&ss&"分再来挖宝吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
randomize timer
r=int(rnd*6)+3
money=0
tl=0
nl=0
zstemp=rs("w6")
select case r
case 1
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
	mess="恭喜"& sjjh_name &"挖到一[<font color=blue>"&js(myxy)&"</font>]一块，这可是个好东西呀！"
	erase js
case 2
	mess="恭喜"& sjjh_name &"挖到一包金器，到集市卖得200万银子"
	money=2000000
case 3
	mess="恭喜"& sjjh_name &"挖到一把E线神剑，变卖得到银两500万两"
	money=5000000
case 4
	mess="恭喜"& sjjh_name &"挖到一个魔法戒指，变卖得到银子100万两"
	money=1000000
case 5
	mess="倒霉的"& sjjh_name &"宝藏没找到，而且不小心踩到陷阱,体力减少50000，内力减少20000"
	tl=50000
	nl=20000
case 6
	mess="强盗又来抢劫,"& sjjh_name &"反抗遭到毒打,体力下降50000、内力下降50000"
	tl=50000
	nl=50000
case 7
	rs.close
	rs.open "select 姓名 from 用户 where 宝物='" & Application("sjjh_baowuname") & "'",conn
	if rs.eof or rs.bof  then
	    mess=""& sjjh_name &"运气真是好，居然挖出了江湖至宝"&Application("sjjh_baowuname")&"！大家还不快抢！"
		conn.execute  "update 用户 set 保护=false,宝物修练=0,宝物='"& Application("sjjh_baowuname") &"' where 姓名='" & sjjh_name &"'"
	else
		mess="强盗又来抢劫,"& sjjh_name &"把强盗打跑了，得到银两100万!"
		money=1000000
	end if
end select
conn.execute "update 用户 set 操作时间=now(),体力=体力-"&tl&",内力=内力-"&nl&",银两=银两+"&money&",w6='"&zstemp&"'  where 姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>消息</font>"& sjjh_name &"在事山挖宝："&mess			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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

<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link href="../dg/setup.css" rel=stylesheet type="text/css">
<title>挖宝贝OK</title></head>

<body oncontextmenu=self.event.returnValue=false bgcolor="#000000">
<div align="center"> <br>
<br>
<table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="186" width="300">
<tr>
<td bgcolor=#C6BD9B>
<table align="center" width="248">
<tr>
<td valign="top">
<div align="center">
<p><%=mess%></p>
</div>
</table>
<div align="center"><br>
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</div>
</td>
</tr>
</table>
</div>
</body>
</html>