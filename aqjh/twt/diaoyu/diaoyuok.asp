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
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"twt/diaoyu/diao.asp")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间 from 用户 where 姓名='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<8 then
	ss=8-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('请等上"&ss&"分再来钓鱼吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 操作时间=now()  where 姓名='"&aqjh_name&"'"
randomize timer
r=int(rnd*6)+1
select case r
case 1
	mess="恭喜"& aqjh_name &"钓到一条大鲨鱼，可作为暗器使用，杀伤体力1200、内力800"
	rs.open "SELECT w6 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"大鲨鱼",1)
	conn.execute "update 用户 set  w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 2
	mess="恭喜"& aqjh_name &"钓到一条娃娃鱼，到集市卖得5万银子"
	conn.execute "update 用户 set 银两=银两+50000 where 姓名='"&aqjh_name&"'" 
case 3
	mess="恭喜"& aqjh_name &"钓到一只大草鱼，吃下可以增加体力300、内力50"
	rs.open "SELECT w6 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"大草鱼",1)
	conn.execute "update 用户 set  w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 4
	mess="恭喜"& aqjh_name &"钓到一条小鲤鱼，吃下可以增加体力100、内力30"
	rs.open "SELECT w6 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"小鲤鱼",1)
	conn.execute "update 用户 set  w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 5
	mess="唉！！"& aqjh_name &"鱼没钓到，摔到河里体力减少300，内力减少100"
	conn.execute "update 用户 set 体力=体力-300,内力=内力-100 where 姓名='"&aqjh_name&"'"
case 6
	mess= aqjh_name &"偷钓鱼塘的鱼塘被主人发现，一阵殴打，体力下降500、内力下降200"
	conn.execute "update 用户 set 体力=体力-500,内力=内力-200 where 姓名='"&aqjh_name&"'"
case 7
	mess="恭喜"& aqjh_name &"运气真是好的BiangBiang声呀！钓到大鲨鱼、大草鱼、小鲤鱼各一条！"
	rs.open "SELECT w6 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"大鲨鱼",1)
	duyao=add(duyao,"大草鱼",1)
	duyao=add(duyao,"小鲤鱼",1)
	conn.execute "update 用户 set  w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
end select
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>【河边垂钓】</font>"& aqjh_name &"在河边钓鱼："&mess			'聊天数据
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
<title>钓鱼OK!</title></head>
<body oncontextmenu=self.event.returnValue=false background="../../bg.gif">
<div align="center"> <br>
<br><table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="186" width="300">
<tr><td bgcolor=#C6BD9B>
<table align="center" width="248">
<tr><td valign="top">
<div align="center">
<p><%=mess%></p>
</div></table>
<div align="center"><br>
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</div></td></tr></table></div></body></html>