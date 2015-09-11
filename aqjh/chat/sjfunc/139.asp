<%@ LANGUAGE=VBScript%>
<!--#include file="../../mywp.asp"-->
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
rs.open "select 金币,会员金卡,w6,操作时间 from 用户 where 姓名='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<3 then
	ss=3-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('请等上"&ss&"分再来走运气！');window.close();</script>"
	Response.End
end if
randomize timer
r=int(rnd*24)+1
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
	mess="恭喜"& aqjh_name &"捡到一[<font color=blue>"&js(myxy)&"</font>]一块，可作为配药系统必备的好东西 ！"
	sql="update 用户 set 操作时间=now(),w6='"&zstemp&"' where 姓名='"&aqjh_name&"'"
	erase js
case 2
	mess=""& aqjh_name &"运气真好，捡到别人不小心掉的金币<img src='img/jinbi.gif'>1个"
	sql="update 用户 set 操作时间=now(),金币=金币+1 where 姓名='"&aqjh_name&"'"
case 3
	mess=""& aqjh_name &"运气真好，捡到别人不小心掉的金币<img src='img/jinbi.gif'>2个"
	sql="update 用户 set 操作时间=now(),金币=金币+2 where 姓名='"&aqjh_name&"'"
case 4
	mess=""& aqjh_name &"运气真好，捡到别人不小心掉的金币<img src='img/jinbi.gif'>3个"
	sql="update 用户 set 操作时间=now(),金币=金币+3 where 姓名='"&aqjh_name&"'"
case 5
	mess=""& aqjh_name &"运气真好，碰到了江湖大好人别人，别人送他金币<img src='img/jinbi.gif'>2个"
	sql="update 用户 set 操作时间=now(),金币=金币+2 where 姓名='"&aqjh_name&"'"
case 6
	mess=""& aqjh_name &"运气差差,碰到一美女，对他有所不轨,结果被众人殴打,道德减少1000"
	sql="update 用户 set 操作时间=now(),道德=道德-1000 where 姓名='"&aqjh_name&"'"
case 7
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖疯子几个哥们，被狂殴了一顿,体力减少1500"
	sql="update 用户 set 操作时间=now(),体力=体力-1500 where 姓名='"&aqjh_name&"'"
case 8
	mess=""& aqjh_name &"运气真好，碰到了江湖首富周扒皮，平日里周扒皮这小子最一毛不拔了，今天不知道遇什么喜事了，心情特别好，就送他金币<img src='img/jinbi.gif'>4张"
	sql="update 用户 set 操作时间=now(),金币=金币+4 where 姓名='"&aqjh_name&"'"
case 9
	mess=""& aqjh_name &"运气差差,碰到大恶人神剑追命这小子，被狂殴了一顿,回去纠集一帮兄弟报仇,结果2败俱伤,体力减少1000点"
	sql="update 用户 set 操作时间=now(),体力=体力-1000 where 姓名='"&aqjh_name&"'"
case 10
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，被狂殴了一顿,灰溜溜的跑回去哭了!~~"
        sql="update 用户 set 操作时间=now() where 姓名='"&aqjh_name&"'"
case 11
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，被狂殴了一顿,体力下降100点"
	sql="update 用户 set 操作时间=now(),体力=体力-100 where 姓名='"&aqjh_name&"'"
	case 12
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，被狂殴了一顿,武功损失1000点"
	sql="update 用户 set 操作时间=now(),武功=武功-1000 where 姓名='"&aqjh_name&"'"
	case 13
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，被狂殴了一顿,豆点被抢走了10点"
	sql="update 用户 set 操作时间=now(),泡豆点数=泡豆点数-10 where 姓名='"&aqjh_name&"'"
	case 14
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，被狂殴了一顿,内力损失200点"
	sql="update 用户 set 操作时间=now(),内力=内力-200 where 姓名='"&aqjh_name&"'"
	case 15
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，被狂殴了一顿,得到医疗费用2000元"
        sql="update 用户 set 操作时间=now(),存款=存款+2000 where 姓名='"&aqjh_name&"'"
	case 16
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，被狂殴了一顿,正好站长路过,救回一条命,什么也没有损失"
	sql="update 用户 set 操作时间=now() where 姓名='"&aqjh_name&"'"
	case 17
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，被狂殴了一顿,被一美眉看见,魅力大减300点"
	sql="update 用户 set 操作时间=now(),魅力=魅力-300  where 姓名='"&aqjh_name&"'"
	case 18
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，被狂殴了一顿,被一美眉遇见,碍与面子,和他们拼了,结果体力损失3000点,"
	sql="update 用户 set 操作时间=now(),体力=体力-3000 where 姓名='"&aqjh_name&"'"
	case 19
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，被狂殴了一顿,自己跑回家,哭了3天,觉得还是做坏人好,于是加入黑社会,道德损失300点"
	sql="update 用户 set 操作时间=now(),道德=道德+2000 where 姓名='"&aqjh_name&"'"
	case 20
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，被狂殴了一顿,回家后发奋图强,苦练武功,终于成为一代大侠,武功暴涨2000点"
	sql="update 用户 set 操作时间=now(),武功=武功+2000 where 姓名='"&aqjh_name&"'"
	case 21
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，被狂殴了一顿,打成残废,体力减少10000"
	sql="update 用户 set 操作时间=now(),体力=体力-1000 where 姓名='"&aqjh_name&"'"
        case 22
	mess=""& aqjh_name &"运气真好，碰到了江湖大好人别人，别人送他金卡<img src='img/jk.gif'>1元"
	sql="update 用户 set 操作时间=now(),会员金卡=会员金卡+1 where 姓名='"&aqjh_name&"'"
	case 23
	mess=""& aqjh_name &"运气差差,碰到大恶人和江湖恐龙几个哥们，差点被狂殴了一顿,好在他武功高强,把坏人打的落化流水.魅力大涨500点"
	sql="update 用户 set 操作时间=now(),魅力=魅力+500 where 姓名='"&aqjh_name&"'"
        case 24
        rs.close
	rs.open "select 姓名 from 用户 where 姓名='"&aqjh_name&"' and 宝物='" & Application("aqjh_baowuname") & "'",conn
	if rs.eof or rs.bof  then
	    mess=""& aqjh_name &"运气真是好，居然捡到了江湖至宝"&Application("aqjh_baowuname")&"！大家还不快抢！"
            sql="update 用户 set 保护=false,宝物修练=0,宝物='"& Application("aqjh_baowuname") &"' where 姓名='"&aqjh_name &"'"
        else
            mess=""& aqjh_name &"忽然眼前一亮，居然看到前面不远出现江湖至宝"&Application("aqjh_baowuname")&"！正准备去拿！由于太贪心，自己的宝物已经被一个大汉抢走。"
            sql="update 用户 set 宝物修练=0,宝物='无' where 姓名='"&aqjh_name &"'"
	end if
end select
conn.execute ""&sql&""
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>【运气消息】</font>"&mess			'聊天数据
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
<title>走运气OK</title></head>
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