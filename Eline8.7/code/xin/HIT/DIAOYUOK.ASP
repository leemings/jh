<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 操作时间 from 用户 where 姓名='"&sjjh_name&"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<10 then
	ss=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('请等上"&ss&"分再来吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 操作时间=now()  where 姓名='"&sjjh_name&"'"
randomize timer
r=int(rnd*21)+1
select case r
case 1
	mess= sjjh_name &"驾驶破飞机成功，得到奖励法力300点"
	conn.execute "update 用户 set 法力=法力+300 where 姓名='"&sjjh_name&"'"
case 2
	mess= sjjh_name &"驾驶40年代的破飞机，忽然电闪雷鸣，飞机被击中,体力下降50万、内力下降28万"
	conn.execute "update 用户 set 体力=体力-500000,内力=内力-280000 where 姓名='"&sjjh_name&"'"
case 3
	mess= sjjh_name &"驾驶破飞机失事，堕落太平洋，跑去见上帝了"
	conn.execute "update 用户 set 状态='死',登录=now(),事件原因='上帝|"&fn1&"' where 姓名='" & sjjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'上帝','飞机失事','人命')"
case 4
	mess= sjjh_name &"驾驶40年代的破飞机成功，得到奖励法力200点"
	conn.execute "update 用户 set 法力=法力+200 where 姓名='"&sjjh_name&"'"
case 5
	mess= sjjh_name &"驾驶40年代的破飞机，发现太平洋新岛屿，得到奖金价值法力250个"
	conn.execute "update 用户 set 法力=法力+250 where 姓名='"&sjjh_name&"'"
case 6
	mess= sjjh_name &"驾驶40年代的破飞机，不小心掉下河，体力下降49万、内力下降24万"
	conn.execute "update 用户 set 体力=体力-490000,内力=内力-240000 where 姓名='"&sjjh_name&"'"
case 7
	mess="各位各位!!"& sjjh_name &"把他的破飞机卖掉了，卖得银子1000万"
	conn.execute "update 用户 set 银两=银两+10000000 where 姓名='"&sjjh_name&"'" 
case 8
	mess= sjjh_name &"驾驶破飞机想环球旅行，不料飞机失事，死得好惨啊"
	conn.execute "update 用户 set 状态='死',登录=now(),事件原因='上帝|"&fn1&"' where 姓名='" & sjjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'上帝','飞机失事','人命')"
case 9
	mess= sjjh_name &"驾驶40年代的破飞机，不小心撞中飞鸟，失事了！体力下降38万、内力下降24万"
	conn.execute "update 用户 set 体力=体力-380000,内力=内力-240000 where 姓名='"&sjjh_name&"'"
case 10
	mess= sjjh_name &"驾驶40年代的破飞机想飞越美国领空，不料被防空导弹击中，体力下降37万、内力下降32万"
	conn.execute "update 用户 set 体力=体力-3700,内力=内力-3200 where 姓名='"&sjjh_name&"'"
case 11
	mess= sjjh_name &"驾驶40年代的破飞机，谁知飞机撞入森林，体力下降28万、内力下降10万"
	conn.execute "update 用户 set 体力=体力-280000,内力=内力-100000 where 姓名='"&sjjh_name&"'"
case 12
	mess= sjjh_name &"驾驶40年代的破飞机，发现太平洋新岛屿，得到奖金价值法力250点"
	conn.execute "update 用户 set 法力=法力+250 where 姓名='"&sjjh_name&"'"
case 13
	mess= sjjh_name &"驾驶40年代的破飞机，发现太平洋新岛屿，得到奖金价值法力150点"
	conn.execute "update 用户 set 法力=法力+150 where 姓名='"&sjjh_name&"'"
case 14
	mess= sjjh_name &"驾驶破飞机失事，联合国下半旗哀悼"
	conn.execute "update 用户 set 状态='死',登录=now(),事件原因='上帝|"&fn1&"' where 姓名='" & sjjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'上帝','飞机失事','人命')"
case 15
	mess="各位各位!!"& sjjh_name &"把他的破飞机卖掉了，卖得银子800万"
	conn.execute "update 用户 set 银两=银两+8000000 where 姓名='"&sjjh_name&"'"
case 16
	mess="各位各位!!"& sjjh_name &"把他的破飞机卖掉了，卖得银子1000万"
	conn.execute "update 用户 set 银两=银两+10000000 where 姓名='"&sjjh_name&"'"  
case 17
	mess= sjjh_name &"驾驶破飞机失事，惨啦，现在只有一口气了"
	conn.execute "update 用户 set 状态='死',登录=now(),事件原因='上帝|"&fn1&"' where 姓名='" & sjjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'上帝','飞机失事','人命')"
case 18
	mess= sjjh_name &"驾驶40年代的破飞机,可惜飞机发动机坏了，修理费要银子1000万"
	conn.execute "update 用户 set 银子=银子-10000000 where 姓名='"&sjjh_name&"'"
case 19
	mess= sjjh_name &"驾驶40年代的破飞机,可惜飞机发动机坏了，修理费要银子1000万"
	conn.execute "update 用户 set 银子=银子-10000000 where 姓名='"&sjjh_name&"'"
case 20
	mess= sjjh_name &"把自己的破飞机卖给博物馆，得到金币1个，真是开心"
	conn.execute "update 用户 set 金币=金币+1 where 姓名='"&sjjh_name&"'"
case 21
	mess= sjjh_name &"驾驶破飞机失事，各位大虾默哀一分钟吧"
	conn.execute "update 用户 set 状态='死',登录=now(),事件原因='上帝|"&fn1&"' where 姓名='" & sjjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'上帝','飞机失事','人命')"
end select
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>消息</font>"& sjjh_name &"开飞机："&mess			'聊天数据
says=replace(says,"'","\'")
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
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link href="../dg/setup.css" rel=stylesheet type="text/css">
<title>飞行OK!</title></head>

<body oncontextmenu=self.event.returnValue=false background="../jhimg/bk_hc3w.gif">
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

