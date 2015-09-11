<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name=""  then
Response.Redirect "../error.asp?id=440"
else
randomize
m=int(6*rnd+1)
if m>3 then
if request.form("select")="big" then
Randomize
m1 = Int(6 * Rnd + 1)
Randomize
m3 = Int(6 * Rnd + 1)
Randomize
m2 = Int(5 * Rnd + 1)
else
Randomize
m1 = Int(6 * Rnd + 1)
Randomize
m3 = Int(6 * Rnd + 1)
Randomize
m2 = Int(7 * Rnd + 1)
if m2>6 then m3=6 end if
end if
else
Randomize
m1 = Int(6 * Rnd + 1)
Randomize
m3 = Int(6 * Rnd + 1)
Randomize
m2 = Int(6 * Rnd + 1)
end if



if request.form("select")="big" then
if m1+m2+m3>9 then
mm=int(6*rnd+1)
if mm=1 or mm=2 or mm=3 or mm=4 then
do while m1+m2+m3>=10
m1 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
m2 = Int(6 * Rnd + 1)
loop
end if
end if
else
if m1+m2+m3<9 then
mm=int(6*rnd+1)
if mm=1 or mm=2 or mm=3 or mm=4 then
do while m1+m2+m3<9
m1 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
m2 = Int(6 * Rnd + 1)
loop
end if
end if
end if


dim betcash,nowcash
nowcash=0
betcash=0
betcash=clng(request.form("splosh"))
if betcash<10 then
%>
<script language=vbscript>
MsgBox "开什么玩笑？您下这么少的注赌个什么劲？"
location.href = "javascript:history.back()"
</script>
<%
elseif betcash>3000 then
%>
<script language=vbscript>
MsgBox "开什么玩笑？您赌这么多想叫我破产啊！"
location.href = "javascript:history.back()"
</script>
<%
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("aqjh_usermdb")
sql="Select  * from 用户 where 姓名='"&aqjh_name&"'"
set rs=conn.Execute(sql)
nowcash=rs("银两")
if betcash>nowcash then
%>
<script language=vbscript>
MsgBox "有没有搞错？您老荷包里哪有这么多银子？（嘿嘿嘿，快去挣点去吧，穷鬼！）"
location.href = "javascript:history.back()"
</script>
<%
else

nowmeili=rs("魅力")
if nowmeili<10 then
%>
<script language=vbscript>
MsgBox "大侠，您目前的魅力低于10了，面子上看不过去呀，以后再来吧！"
location.href = "javascript:history.back()"
</script>
<%
else
%>
<html>
<head>
<title>江湖赌馆 - 赌大小</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
--></style>
</head>

<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
<p><font size="3" class="c" color="#000000"><b>赌 场 - 赌 大 小<br>
</b></font><font size="2" class="c"><b><font color="#FF0033">你押的是</font><font color="#CC0000">
<%if request.form("select")="big" then%>
</font><font color="#CC0000"> </font></b></font><img src="../jhimg/bbs/big.gif" width="46" height="40"><font size="2" class="c"><b><font color="#CC0000">
<%else%>
<img src="../jhimg/bbs/small.gif" width="46" height="40"></font><font size="2" class="c" color="#CC0000">
<%end if%>
</font></b></font>
<%if (m1+m2+m3)>9 then%>
</p>
<table border="1" cellspacing="0" cellpadding="3" align="center" width="400" bordercolordark="#FFFFFF" bordercolor="#000000">
<tr bgcolor="#336633">
<td colspan="3" align="center"><font class="c" size="2" color="#FFFFFF"><b>结果:</b></font><font size="2" color="#FFFFFF"><b><%=(m1+m2+m3)%>点，大</b></font></td>
</tr>
<tr>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m1%>.gif" width="32" height="31"></font></td>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m2%>.gif"></font></td>
<td width="34%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m3%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"> <font size="2" color="#FFFFFF">
<%if request.form("select")="small" then%>
</font><font size="2" color="#FFFFFF"><b><%=betcash%></b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right"><font size="2" color="#000099"><b>庄家：</b></font></td>
<td colspan="2"><font size="2">呵呵还要来吗？ 有赌未必输哦~！</font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>我说：</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">再来！为什么不？，我呸！我就不相信这么倒霉~！</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td colspan="2"><font size="2"><a href="../jh.asp">算了算了，认倒霉吧，留着点银子救命吧~！</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%
conn.execute("Update 用户 set 银两=银两-"&betcash&",赌次数=赌次数+1,魅力=魅力-1,赢钱=赢钱-"&betcash&" where 姓名='"&aqjh_name&"'")
%>
</font><font size="2" color="#FFFFFF"><b><%=(nowcash-betcash)%> </b><%=(nowmeili-1)%>
<%else%>
</font><font size="2" color="#FFFFFF"><b><%=betcash%> 两</b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right"><font size="2" color="#000099"><b>庄家：</b></font></td>
<td colspan="2"><font size="2">厉害呀~！客官您还要来吗？ </font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>我说：</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">当然再来！我今天运气不错~！</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
            <td colspan="2"><font size="2"><a href="../welcome.asp">见好就收，你以为我傻？嘿嘿！</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%conn.execute("Update 用户 set 银两=银两+"&betcash&",赢次数=赢次数+1,赢钱=赢钱+"&betcash&",赌次数=赌次数+1,魅力=魅力-1 where 姓名='"&aqjh_name&"'")
%>
</font><font size="2" color="#FFFFFF"><b><%=(betcash+nowcash)%></b><%=(nowmeili-1)%></font><font size="2" color="#FFFFFF">
<%end if%>
</font> </td>
</tr>
<tr bgcolor="#FFCCCC">
<td align="right" colspan="3" height="2"><a href="betindex.asp"><font size="2">赌</font><font size="2">场首页</font></a></td>
</tr>
</table>
<font size="3"><%else%></font>
<table border="1" cellspacing="0" cellpadding="3" align="center" width="400" bordercolordark="#FFFFFF" height="160">
<tr bgcolor="#336633">
<td colspan="3" align="center"><font class="c" size="2" color="#FFFFFF"><b>结果:</b></font><font size="2" color="#FFFFFF"><b><%=(m1+m2+m3)%>点，小</b></font></td>
</tr>
<tr>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m1%>.gif"></font></td>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m2%>.gif"></font></td>
<td width="34%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m3%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"><font size="2" color="#FFFFFF">
<%if request.form("select")="big" then%>
</font> <font size="2" color="#FFFFFF"><b><%=betcash%></b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right"><font size="2" color="#000099"><b>庄家：</b></font></td>
<td colspan="2"><font size="2">呵呵还要来吗？ 有赌未为输哦~！</font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>我说：</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">再来！为什么不？，我呸！我就不相信这么倒霉~！</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td colspan="2"><font size="2"><a href="../jh.asp">算了算了，认倒霉吧，留着点银子救命吧~！</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%conn.execute("Update 用户 set 银两=银两-"&betcash&",赌次数=赌次数+1,魅力=魅力-1,赢钱=赢钱-"&betcash&" where 姓名='"&aqjh_name&"'")
%>
</font> <font size="2" color="#FFFFFF"><b><%=(nowcash-betcash)%></b><%=(nowmeili-1)%>
<%else%>
</font><font size="2" color="#FFFFFF"><b><%=betcash%> 两</b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC">
<tr>
<td align="right"><font size="2" color="#000099"><b>庄家：</b></font></td>
<td colspan="2"><font size="2">厉害呀~！客官您还要来吗？ </font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>我说：</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">当然再来！我今天运气不错~！</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td colspan="2"><font size="2"><a href="../jh.asp">见好就收，你以为我傻？嘿嘿！</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%conn.execute("Update 用户 set 银两=银两+"&betcash&",赢次数=赢次数+1,赢钱=赢钱+"&betcash&",赌次数=赌次数+1,魅力=魅力-1 where 姓名='"&aqjh_name&"'")
%>
</font> <font size="2" color="#FFFFFF"><b><%=(betcash+nowcash)%></b><%=(nowmeili-1)%></font><font size="2" color="#FFFFFF">
<%end if%>
</font>
<p>&nbsp;</p>
</td>
</tr>
<tr bgcolor="#FFCCCC">
<td align="right" colspan="3"><a href="betindex.asp"><font size="2">赌场首页</font></a></td>
</tr>
</table>
<%end if%>
<p>版权所有『爱情江湖总站』</p>
</div>
</body>
</html>
<%rs.close
set rs=nothing
set username=nothing
set betcash=nothing
set nowcash=nothing
end if
end if
end if
end if
%>
