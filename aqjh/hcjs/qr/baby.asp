<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 性别,保留,hw from 用户 where 姓名='"&aqjh_name&"'",conn
if rs("性别")="男" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你一个大男人进产房来做什么，出去！');window.close();}</script>"
	Response.End
end if
if rs("hw")<>"" then
	rs.close
	conn.execute "update 用户 set 保留='保留' where 姓名='"&aqjh_name&"'"
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('要想富，先修路，少生孩子多种树。你已经有孩子了，不能再生！');window.close();}</script>"
	Response.End
end if
if rs("保留")="保留" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('还没怀孕就想生孩子？！');window.close();}</script>"
	Response.End
end if
yun=rs("保留")
rs.close
hdata=split(yun,"|")
baby=hdata(0)
babysj=hdata(1)
babytl=clng(hdata(2))
babyzt=hdata(3)
erase hdata
sysj=DateDiff("d",babysj,now())
if sysj<20 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('怀孕不到20天，未到生育时间。');window.close();}</script>"
	Response.End
end if
if sysj>21 then
	conn.execute "update 用户 set 保留='保留' where 姓名='"&aqjh_name&"'"
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你怀孕时间超出21天，孩子早已死在腹中了！');window.close();}</script>"
	Response.End
end if
newbl="保留"
randomize timer
r=int(rnd*5)+1
if sysj=20 or (sysj=21 and r<>5 and r<>6) then
	n=int(rnd*3)+1
	if n=1 or n=3 then
		xhxb="男"
		h1="大胖小子"
		t="boy.gif"
	else
		xhxb="女"
		h1="漂亮公主"
		t="girl.gif"
	end if
	xhtl=babytl+1000
	xhnl=int(xhtl*3/4)
	xhz="小孩|"&xhxb&"|"&xhtl&"|"&xhnl&"|200|200|"&now()&"|乳|0|"&now()&"|0"
	hh="<img src=images/"&t&"><BR><BR><img src='img/004.gif'><br><marquee width=100% behavior=alternate scrollamount=5><font color=red size=+1>喜喜</font></marquee><BR>恭喜你，你顺利生下一个"& h1 &"，孩子体力："& xhtl &"，内力："&xhnl&"，攻击：200，防御：200，快去告诉孩子的爸爸，和他一起给孩子起个名。这是你们爱情的结晶。"
	mess="恭喜<font color=blue>"& aqjh_name &"</font>生下一个<font color=red><b>"& h1 &"</b></font><BR><marquee width=100% behavior=alternate scrollamount=5><font color=red size=+1>喜喜</font></marquee>"
else
	hh="由于你在孕期第20天的时候未能及时来医院，延误了一天的时间。虽然经过医生的努力，但你的孩子还是死掉了。对不起，请节哀顺便。"
	xhz=""
	mess="由于"& aqjh_name &"怀孕时间超过20天，在生育小孩时由于难产，小孩死掉了，默哀……"
end if
conn.execute "update 用户 set 保留='"&newbl&"',hw='"&xhz&"' where 姓名='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000><b>【医院消息】</b></font>"&mess			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session(nowinroom) & ");<"&"/script>"
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
<title>江湖产房</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<div align="center">
  <p> <br>
  </p>
  <p><%=hh%> </p>
  <p>
    <input type=button value=关闭窗口 onClick='window.close()' name="button">
  </p>
</div>
</body></html>