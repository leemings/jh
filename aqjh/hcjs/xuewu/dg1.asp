<%@ LANGUAGE=VBScript codepage ="936" %>
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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
my=aqjh_name
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "你不是江湖中人，不能进入江湖武馆"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<1 then
	ss=1-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('请等上"&ss&"分再来学武吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
tl=rs("体力")
nl=rs("内力")
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body background="by.gif" bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<center>
<%
if tl<100 or nl<50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "错误！你目前的体力或内力值不够，回去多加练练马步蹲裆再来吧！"
location.href = "javascript:history.back()"
</script>
<%
else
conn.execute "update 用户 set 体力=体力-100,内力=内力-50,武功=武功+35,攻击=攻击+5,操作时间=now() where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
message="" & my & "辛辛苦苦学了一天，你武功+35攻击+5,消耗体力100内力50！"
says="<font color=red>【江湖武馆】</font><font color=green>["&aqjh_name&"]辛辛苦苦学了一天[六脉神剑]，武功+35攻击+5,消耗体力100内力50</font>"
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
end if
%>
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr><td height="37">
<font color="#000000"><strong>江湖武馆:</strong></font>
<tr>
<td height="182" valign="top"><%=message%><Br><Br><center>
</td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="返 回" onclick="location.href='xuetang.htm'">
</div>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>
