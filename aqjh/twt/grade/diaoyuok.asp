<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间 from 用户 where 姓名='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<2 then
	ss=2-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('请等上"&ss&"分再来寻宝吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 操作时间=now()  where 姓名='"&aqjh_name&"'"
randomize timer
r=int(rnd*18)+1
select case r
case 1
	mess="各位帅哥美女!!"& aqjh_name &"在黄金岛寻宝，宝没寻到跑去钓鱼，呵呵！收获真不错，得到大鲨鱼五条"
	rs.open "SELECT w6 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"大鲨鱼",5)
	conn.execute "update 用户 set  w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 2
	mess= aqjh_name &"在黄金岛寻宝，忽然电闪雷鸣，被淋得象个落汤鸡，体力下降5000、内力下降2800"
	conn.execute "update 用户 set 体力=体力-5000,内力=内力-2800 where 姓名='"&aqjh_name&"'"
case 3
	mess="各位帅哥美女!!"& aqjh_name &"在黄金岛寻宝，得到药农相助，宝虽没寻到，却获赠精制当归99包"
	rs.open "SELECT w1 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w1"),"精制当归",99)
	conn.execute "update 用户 set  w1='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 4
	mess="各位帅哥美女!!"& aqjh_name &"在黄金岛寻宝，得到药农相助，宝虽没寻到，却获赠女儿茶99包"
	rs.open "SELECT w1 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w1"),"女儿茶",99)
	conn.execute "update 用户 set  w1='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 5
	mess= aqjh_name &"在黄金岛寻宝，找到萨达姆埋在那里的金库，得到金币1个"
	conn.execute "update 用户 set 金币=金币+1 where 姓名='"&aqjh_name&"'"
case 6
	mess="各位帅哥美女!!"& aqjh_name &"在黄金岛寻宝，得到药农相助，宝虽没寻到，却获赠金创药99包"
	rs.open "SELECT w1 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w1"),"金创药",99)
	conn.execute "update 用户 set  w1='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 7
	mess= aqjh_name &"在黄金岛寻宝，不小心掉下河，体力下降4900、内力下降2400"
	conn.execute "update 用户 set 体力=体力-4900,内力=内力-2400 where 姓名='"&aqjh_name&"'"
case 8
	mess="各位帅哥美女!!"& aqjh_name &"在黄金岛寻宝，在海边拾到2个金币"
	conn.execute "update 用户 set 金币=金币+2 where 姓名='"&aqjh_name&"'"
case 9
	mess= aqjh_name &"在黄金岛寻宝，被老虎袭击，体力下降3800、内力下降2400"
	conn.execute "update 用户 set 体力=体力-3800,内力=内力-2400 where 姓名='"&aqjh_name&"'"
case 10
	mess="各位帅哥美女!!"& aqjh_name &"在黄金岛寻宝，由于寻不到宝便跳到河里捉鱼，捉到一条大草鱼"
	rs.open "SELECT w6 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"大草鱼",1)
	conn.execute "update 用户 set  w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 11
	mess= aqjh_name &"在黄金岛寻宝,没想到被强盗抢劫，一顿暴打，体力下降3700、内力下降3200"
	conn.execute "update 用户 set 体力=体力-3700,内力=内力-3200 where 姓名='"&aqjh_name&"'"
case 12
	mess= aqjh_name &"在黄金岛寻宝，谁知碰见恶人，一翻拳打脚踢，体力下降2800、内力下降1000"
	conn.execute "update 用户 set 体力=体力-2800,内力=内力-1000 where 姓名='"&aqjh_name&"'"
case 13
	mess="各位帅哥美女!!"& aqjh_name &"在黄金岛寻宝，得到花农相助，宝虽没寻到，却获赠跳舞的玫瑰99朵<img src='../hcjs/jhjs/images/flower390.gif'>"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w7"),"跳舞的玫瑰",99)
	conn.execute "update 用户 set  w7='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 14
	mess="各位帅哥美女!!"& aqjh_name &"在黄金岛寻宝，得到花农相助，宝虽没寻到，却获赠美满生活99朵<img src='../hcjs/jhjs/images/f24.jpg'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w7"),"美满生活",99)
	conn.execute "update 用户 set  w7='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 15
	mess="各位帅哥美女!!"& aqjh_name &"在黄金岛寻宝，得到花农相助，宝虽没寻到，却获赠摇头的玫瑰99朵<img src='../hcjs/jhjs/images/flower370.gif'>"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w7"),"摇头的玫瑰",99)
	conn.execute "update 用户 set  w7='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 16
	mess="各位帅哥美女!!"& aqjh_name &"在黄金岛寻宝，得到花农相助，宝虽没寻到，却获赠恋爱热吻99朵<img src='../hcjs/jhjs/images/flowerr12.GIF'>。"
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w7"),"恋爱热吻",99)
	conn.execute "update 用户 set  w7='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 17
	mess= aqjh_name &"在黄金岛寻宝，谁知碰见恶人，一翻拳打脚踢，体力下降2800、内力下降1000"
	conn.execute "update 用户 set 体力=体力-2800,内力=内力-1000 where 姓名='"&aqjh_name&"'"
case 18
	mess= aqjh_name &"在黄金岛寻宝，忽然电闪雷鸣，被淋得象个落汤鸡，体力下降5000、内力下降2800"
	conn.execute "update 用户 set 体力=体力-5000,内力=内力-2800 where 姓名='"&aqjh_name&"'"
end select
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>【黄金岛寻宝】</font>"& aqjh_name &"寻宝："&mess			'聊天数据
says=replace(says,"'","\'")
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
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link href="../../css.css" rel=stylesheet type="text/css">
<title>寻宝OK!</title></head>
<body oncontextmenu=self.event.returnValue=false background="../../bg.gif">
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
</table></div></body></html>