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
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"diaoyu/diao.asp")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 操作时间 from 用户 where 姓名='"&sjjh_name&"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<15 then
	ss=15-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('请等上"&ss&"分再来钓鱼吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 操作时间=now()  where 姓名='"&sjjh_name&"'"
randomize timer
r=int(rnd*6)+1
select case r
case 1
	mess="恭喜"& sjjh_name &"钓到三只<font color=red>大鲨鱼</font>，杀伤体力1200、内力800"
	rs.open "SELECT w1 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"大鲨鱼",3)
	conn.execute "update 用户 set  w6='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
case 2
	mess="恭喜"& sjjh_name &"钓到一条美人鱼，到集市卖得<font color=red>5</font>个金币"
	conn.execute "update 用户 set 金币=金币+5 where 姓名='"&sjjh_name&"'" 
case 3
	mess="恭喜"& sjjh_name &"钓到一只<font color=red>大补丸</font>，增加内力100，体力3000"
	rs.open "SELECT w1 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
	duyao=add(rs("w1"),"大补丸",50)
	conn.execute "update 用户 set  w1='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
case 4
	mess="恭喜"& sjjh_name &"钓到一条吃人鱼，到集市卖得<font color=red>20</font>个银币"
	conn.execute "update 用户 set 银币=银币+20 where 姓名='"&sjjh_name&"'" 
case 5
	mess="唉！！"& sjjh_name &"鱼没钓到，摔到河里体力减少30000，内力减少10000"
	conn.execute "update 用户 set 体力=体力-30000,内力=内力-10000 where 姓名='"&sjjh_name&"'"
case 6
	mess= sjjh_name &"偷钓鱼塘的鱼塘被主人发现，一阵殴打，体力下降50000、内力下降20000"
	conn.execute "update 用户 set 体力=体力-50000,内力=内力-20000 where 姓名='"&sjjh_name&"'"
case 7
	mess="恭喜"& sjjh_name &"运气真是好的BiangBiang声呀！钓到大鲨鱼、大草鱼、小鲤鱼各一条！"
	rs.open "SELECT w6 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"大鲨鱼",1)
	duyao=add(duyao,"大草鱼",1)
	duyao=add(duyao,"小鲤鱼",1)
	conn.execute "update 用户 set  w6='"&duyao&"' where 姓名='"&sjjh_name&"'"
	rs.close
end select
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>消息</font>江湖会员："& sjjh_name &"在河边钓鱼："&mess			'聊天数据
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
<title>钓鱼OK!</title></head>

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
