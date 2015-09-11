<%@ LANGUAGE=VBScript%>
<!--#include file="../../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 金币,会员,会员金卡,w6,操作时间 from 用户 where 姓名='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<30 then
	ss=30-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('还没有出丑够？快去练武"&ss&"分钟再来！');window.close();</script>"
	Response.End
end if
hy=rs("会员")
if hy>5 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('["&aqjh_name &"]你还以为你是五级以上会员吗？！');window.close();</script>"
 response.end
end if
randomize timer
r=int(rnd*15)+1
jinbi=0
jinka=0
zstemp=rs("w6")
select case r
case 1
	dim js(10)
	js(0) ="骨头"
	js(1) ="蓝宝石"
	js(2) ="红宝石"
	js(3) ="僵尸血"
	js(4) ="僵尸牙"
	js(5) ="水晶石"
	js(6) ="绿宝石"
	js(7) ="硫磺"
	js(8) ="硝酸"
	js(9)="木炭"
	randomize()
	myxy = Int(Rnd*10)
	zstemp=add(rs("w6"),js(myxy),1)
	mess="衰哦"& aqjh_name &"武功不如曹操就与人家比武，内力失去1000、武功下降666！"
	sql="update 用户 set 操作时间=now(),内力=内力-1000,武功=武功-666 where 姓名='"&aqjh_name&"'"
	erase js
case 2
	mess="在一个漆黑的夜晚，三国吕布来到了爱情江湖，打伤"& aqjh_name &"体力下降20000<img src='img/kkkkk.gif'>"
	sql="update 用户 set 操作时间=now(),金币=体力-20000 where 姓名='"&aqjh_name&"'"
case 3
	mess="在爱情江湖大道上闲逛，"& aqjh_name &"误入董卓的圈套，武功大失9999点<img src='img/kkkkk.gif'>"
	sql="update 用户 set 操作时间=now(),武功=武功-9999 where 姓名='"&aqjh_name&"'"
case 4
	mess=""& aqjh_name &"在后山与刘备大战三百回合，体力大失<img src='img/kkkkk.gif'>30000点"
	sql="update 用户 set 操作时间=now(),体力=体力-30000 where 姓名='"&aqjh_name&"'"
case 5
	mess=""& aqjh_name &"碰到了永不放弃，永不放弃带他与张飞切磋，武功大增2000点<img src='img/kkkkk.gif'>"
	sql="update 用户 set 操作时间=now(),武功=武功+2000 where 姓名='"&aqjh_name&"'"
case 6
	mess=""& aqjh_name &"在一处溪流处饮水，意外的学习到关羽的健身拳，使体力增加2000点，真是幸运哦"
	sql="update 用户 set 操作时间=now(),体力=体力+2000 where 姓名='"&aqjh_name&"'"
case 7
	mess=""& aqjh_name &"好厉害，好厉害，居然打败了周瑜，银两增加10000！"
	sql="update 用户 set 操作时间=now(),银两=银两+10000 where 姓名='"&aqjh_name&"'"
case 8
	mess=""& aqjh_name &"在翠花楼，碰到了三国人物袁绍，平日里袁绍这小子最一毛不拔了，今天不知道遇什么喜事了，心情特别好，就送他金币<img src='img/kkkkk.gif'>4个"
	sql="update 用户 set 操作时间=now(),金币=金币+4 where 姓名='"&aqjh_name&"'"
case 9
	mess=""& aqjh_name &"在一处险道见一剑客练剑，近前一看才知道是诸葛亮在习武，你悄悄偷学了几招，攻击增加200"
	sql="update 用户 set 操作时间=now(),攻击=攻击+200 where 姓名='"&aqjh_name&"'"
case 10
	mess=""& aqjh_name &"看到一剑客练剑，一打听原来是小将关平，悄悄偷学了几招，武功增加了100!~~"
        sql="update 用户 set 操作时间=now(),武功=武功+100 where 姓名='"&aqjh_name&"'"
case 11
	mess=""& aqjh_name &"跑到武当山去旅游,遇到诸葛亮,受到指点,于是武功上限增加200点"
	sql="update 用户 set 操作时间=now(),武功加=武功加+200 where 姓名='"&aqjh_name&"'"
	case 12
	mess=""& aqjh_name &"运气差差,碰到虎将张飞内力减少2000"
	sql="update 用户 set 操作时间=now(),内力=内力-2000 where 姓名='"&aqjh_name&"'"
	case 13
	mess=""& aqjh_name &"敢和董卓去玩反恐精英,防御减少200点!"
	sql="update 用户 set 操作时间=now(),防御=防御-200 where 姓名='"&aqjh_name&"'"
	case 14
	mess=""& aqjh_name &"受到诸葛亮指点,轻功增加2000点!"
	sql="update 用户 set 操作时间=now(),轻功=轻功+2000 where 姓名='"&aqjh_name&"'"
        case 15
        rs.close
	rs.open "select 姓名 from 用户 where 姓名='"&aqjh_name&"' and 宝物='" & Application("aqjh_baowuname") & "'",conn
	if rs.eof or rs.bof  then
	    mess=""& aqjh_name &"马超大战三百回合没有分出胜负，马超念其武功高强送江湖至宝"&Application("aqjh_baowuname")&"！大家还不快抢！"
            sql="update 用户 set 保护=false,宝物修练=0,宝物='"& Application("aqjh_baowuname") &"' where 姓名='"&aqjh_name &"'"
        else
            mess=""& aqjh_name &"为了得到江湖至宝"&Application("aqjh_baowuname")&"！与曹操密谋，正准备去拿！由于太贪心，自己的宝物已经被一江湖高手抢走。"
            sql="update 用户 set 宝物修练=0,宝物='无' where 姓名='"&aqjh_name &"'"
	end if
end select
conn.execute ""&sql&""
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<marquee width=100% scrollamount=8><font color=reg>【<font color=reg>现场</font>直播】</font></marquee>"&mess				'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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
<link href="jh.css" rel=stylesheet type="text/css">
<title>造化三国</title></head>
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
</body></html>