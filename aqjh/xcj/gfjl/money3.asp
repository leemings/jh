<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
nowinroom=session("nowinroom")
id=LCase(trim(request.querystring("id")))
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
n=Year(date())
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r
rs.open "SELECT * FROM 用户 where 姓名='"&aqjh_name&"'",conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Redirect "error.asp?id=453"
	response.end
end if
dlsj=DateDiff("n",rs("登录"),now())
if dlsj<5 and aqjh_grade<10 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('为防止作弊，刚上线5分钟的人领取不了月薪水!');location.href = 'javascript:history.back()';}</script>"
		response.end

end if
if rs("姓名")<>"小傻" then
                rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你不是最佳玩家之一!别捣乱!');location.href = 'javascript:history.back()';}</script>"
		response.end
end if
if rs("mvalue")<100000 then
                rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你月积分小于10万，不能领奖励!');location.href = 'javascript:history.back()';}</script>"
		response.end
end if
if aqjh_jhdj>30 then Response.Write "<script language=JavaScript>{alert('你等级太低了，领不了，等下次吧！');location.href = 'javascript:history.back()';}</script>"
pai=rs("门派")
rs.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
tmprs=conn.execute("Select count(*) As 数量 from 用户 where mvalue>200000 and 介绍人='"&aqjh_name&"'")
jsren=tmprs("数量")
rs.open "Select * from 用户 where 姓名='"&aqjh_name&"'",conn
mdate=rs("字段1")
'if CDate(mdate)=now() then
'if day(mdate)>=day(now()) and month(mdate)=month(now()) and year(mdate)=year(now()) then 
if  DateDiff("d",rs("字段1"),now())=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('天呀，["&aqjh_name &"]今天你来领过奖励的！忘记了？');window.close();</script>"
	response.end
end if
yin=rs("会员金卡")
ying=rs("金币")
if yin>40000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('天呀，["&aqjh_name &"]你现在有4000多的金卡你还想要啊！！');window.close();</script>"
	response.end
end if
if ying>40000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('天呀，["&aqjh_name &"]你现在有4000多的金币你还想要啊！！');window.close();</script>"
	response.end
end if
dj=rs("等级")
zs=rs("转生")
gl=rs("grade")
if trim(rs("门派"))<>"出家" then
	'个人按会员乘10万
	money=(5)*2
	conn.execute("update 用户 set 金币=金币+"&money&",字段1=now() where 姓名='"&aqjh_name&"'")
end if
%>
<html>
<head>
<title>奖励领取处</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
TD {FONT-FAMILY: "宋体"; FONT-SIZE: 16pt}
BODY {FONT-FAMILY: "宋体"; FONT-SIZE: 16pt}
SELECT {FONT-FAMILY: "宋体"; FONT-SIZE: 16pt}
A {COLOR: #FFC106; FONT-FAMILY: "宋体"; FONT-SIZE: 16pt; TEXT-DECORATION: none}
A:hover {COLOR: #cc0033; FONT-FAMILY: "宋体"; FONT-SIZE: 16pt; TEXT-DECORATION: underline}
</STYLE>
</head>
<body bgcolor="#CCCCCC" text="#000000" leftmargin="0" background="../jhimg/bk_hc3w.gif">
<div align="center">
<p><font size="2"><font color="#000000"><b>爱情奖励领取处</b></font> <br>
<br>
今天你从爱情领取到奖励<b><font color="#FF0000"><%=money%>个金币</font></b>，小心保存，不要乱花！        
<% 
rs.close 
conn.close 
set rs=nothing 
set conn=nothing 
fn1="<font size=2 color=red>【奖励颁发】☆"&aqjh_name&"☆</font><font size=2 color=blue>贵为爱情十大最佳玩家之一，站长奖励<font color=red>"&money&"</font>个<font color=brown>金币</font>，要再多努力做到更好哦~~</font>" 
says=nuhou(fn1) 
function nuhou(fn1) 
nuhou="<marquee height=80 behavior=alternate loop=100 direction=left >" & fn1 & "" & "</marquee>" 
end function 
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
<br> 
</font> 
</p> 
<p align="center"><font size="2"><br>
<font color="#FF0000" size="+1"><b>奖励领取说明方法</b></font><br>
<br>
<font color="#0000FF">最佳玩家奖：每天可以领取奖励，只有月积分达到10万的才可以领取！</font></font></p>
<p align="center"><font color="#0000FF" size="2">无会员金卡</font></p>
<p align="center"><b><font color="#FF00FF" size="2">无会员金卡领取！</font></b></p>
<p align="center">&nbsp;</p>
<p align="center">
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</p>
</div>
</body>
</html>