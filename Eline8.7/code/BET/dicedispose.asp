<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
dim betcash,nowcash
nowcash=0
betcash=0
betcash=clng(request.form("splosh"))
if betcash<10 then
%>
<script language=vbscript>
MsgBox "开什么玩笑？您下这这么少赌个什么劲呀？"
location.href = "javascript:history.back()"
</script>
<%
elseif betcash>3000 then
%>
<script language=vbscript>
MsgBox "开什么玩笑？您赌这么多想叫我破产啊!"
location.href = "javascript:history.back()"
</script>
<%
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="Select * From 用户 Where 姓名='"&sjjh_name&"'"
set rs=conn.Execute(sql)

Randomize
Randomize
m1 = Int(6 * Rnd + 1)
m4 = Int(6 * Rnd + 1)

nowcash=rs("银两")
if betcash>nowcash then
%>
<script language=vbscript>
MsgBox "有没有搞错？您老荷包里哪有这么多银子？（嘿嘿嘿，快去挣点去吧，穷鬼！）"
location.href = "javascript:history.back()"
</script>
<%
else

Randomize
m2 = Int(6 * Rnd + 1)
m5 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
m6 = Int(6 * Rnd + 1)

nowmeili=rs("魅力")
if nowmeili<10 then
%>
<script language=vbscript>
MsgBox "大侠，您目前的魅力低于10了，面子上看不过去呀，以后再来吧！"
location.href = "javascript:history.back()"
</script>
<%
else

m=Second(time())
m=right(m,1)
if m="0" or m="1" or m="6" or m="7" or m="8" then
do while m1+m2+m3<=m4+m5+m6

m4 = Int(6 * Rnd + 1)
m5 = Int(6 * Rnd + 1)
m6 = Int(6 * Rnd + 1)
loop
else
do while m1+m2+m3>=m4+m5+m6

m4 = Int(6 * Rnd + 1)
m5 = Int(6 * Rnd + 1)
m6 = Int(6 * Rnd + 1)
loop
end if
if m1+m2+m3<m4+m5+m6 then
mm=int(rnd*6+1)
if mm=1 or mm=2   then
do while m1+m2+m3<m4+m5+m6

m4 = Int(6 * Rnd + 1)
m5 = Int(6 * Rnd + 1)
m6 = Int(6 * Rnd + 1)
loop
end if
end if




%>

<html>
<head>

<title>赌骰子♀一线网络→wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
BODY {font-size: 9pt; font-family: 宋体;
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
table {font-size: 9pt; font-family: 宋体}
input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}
.c {  font-family: 宋体; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>

</head>

<body text="#000000" background="../jhimg/bk_hc3w.gif" topmargin="0">
<font color="#FFFFFF"><br>
</font>
<div align="center">
<p><font size="2" class="c" color="#FFFFFF"><font size="3"><b><font color="#000000">赌
场 - 赌 骰 子</font></b></font></font></p>


<table border="1" cellspacing="0" cellpadding="3" align="center" bordercolordark="#FFFFFF" bordercolorlight="#000000" bgcolor="#CCCCCC">
<tr>
<td bgcolor="#336633" colspan="3" height="23"><font size="2" class="c"><b>&nbsp;&nbsp;<font color="#FFFFFF">赌
局 结 果</font></b></font></td>
</tr>
<tr>
<td colspan="3" align="center"><font size="2">□-庄家骰子：<%=(m1+m2+m3)%> 点 </font></td>
</tr>
<tr>
<td width="100" align="center"><font size="2"><img src="../jhimg/bbs/<%=m1%>.gif" width="31" height="32"></font></td>
<td width="100" align="center"><font size="2"><img src="../jhimg/bbs/<%=m2%>.gif"width="31" height="32"></font></td>
<td width="139" align="center"><font size="2"><img src="../jhimg/bbs/<%=m3%>.gif" width="32" height="32"></font></td>
</tr>
<tr>
<td colspan="3" align="center"><font size="2">□-你的骰子：<%=(m4+m5+m6)%> 点 </font></td>
</tr>
<tr>
<td width="100" align="center" height="39"><font size="2"><img src="../jhimg/bbs/<%=m4%>.gif"></font></td>
<td width="100" align="center" height="39"><font size="2"><img src="../jhimg/bbs/<%=m5%>.gif"></font></td>
<td width="139" align="center" height="39"><font size="2"><img src="../jhimg/bbs/<%=m6%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"> <font size="2">
<%if (m1+m2+m3)>=(m4+m5+m6) then%>
</font>
<table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td height="49" width="14%">&nbsp;</td>
<td colspan="2" height="49" width="86%">
<p><font size="2"><img src="../jhimg/bbs/face13.gif" width="20" height="20" align="absmiddle">
真倒霉，输了： <b><font color="#CC0000"><%=betcash%> </font>两</b></font></p>
</td>
</tr>
<tr>
<td align="right" width="14%"><font size="2"><b> 庄家：</b></font></td>
<td colspan="2" width="86%"><font size="2"><b><font color="#000099"><img src="../jhimg/bbs/face18.gif" width="20" height="20" align="absmiddle"></font></b>
呵呵~还要来吗？ 有赌未必输哦~！</font></td>
</tr>
<tr>
<td align="right" rowspan="2" width="14%"><font size="2"><b> 我要：</b></font></td>
<td colspan="2" width="86%"><font size="2"><b><font color="#CC0000"><img src="../jhimg/bbs/face8.gif" width="20" height="20" align="absmiddle"></font></b>
<a href="dice.asp">再来！为什么不？，我呸！我就不相信这么倒霉~！</a></font></td>
</tr>
<tr>
<td colspan="2" width="86%"><font size="2"><img src="../jhimg/bbs/face11.gif" width="20" height="20" align="absmiddle">
<a href="../jh.asp">算了算了，认倒霉吧，留着点银子救命吧~！</a></font></td>
</tr>
</table>
<font size="2"><%
conn.execute("Update 用户 set 银两=银两-"&betcash&",赌次数=赌次数+1,魅力=魅力-1,赢钱=赢钱-"&betcash&" where 姓名='"&sjjh_name&"'")
%> </font>
<hr size="1" width="250">
<font size="2"> 你现在有银子：<font color="#CC0000"><b><font color="#CC0000"><%=(nowcash-betcash)%></font>
</b></font>两 魅力：<%=(nowmeili-1)%><%else%> </font>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right" height="50" width="14%">&nbsp;</td>
<td colspan="2" height="50" width="86%">
<p><font size="2"><img src="../jhimg/bbs/face14.gif" width="20" height="20" align="absmiddle">
嘻嘻~，赢了：<b><font color="#CC0000"><%=betcash%> </font>两</b></font></p>
</td>
</tr>
<tr>
<td align="right" width="14%"><font size="2"><b> 庄家：</b></font></td>
<td colspan="2" width="86%"><font size="2"><b><font color="#000099"><img src="../jhimg/bbs/face17.gif" width="20" height="20" align="absmiddle"></font></b>
厉害呀~！客官您还要来吗？ </font></td>
</tr>
<tr>
<td align="right" width="14%"><font size="2"><b>我要：</b></font></td>
<td colspan="2" width="86%"><font size="2"><img src="../jhimg/bbs/face10.gif" width="20" height="20" align="absmiddle">
<a href="dice.asp">当然再来！我今天运气不错~！</a></font></td>
</tr>
<tr>
<td align="right" width="14%">&nbsp;</td>
<td colspan="2" width="86%"><font size="2"><img src="../jhimg/bbs/face2.gif" width="20" height="20" align="absmiddle">
<a href="../jh.asp">见好就收，你以为我傻？嘿嘿！</a></font></td>
</tr>
</table>
<font size="2"><%conn.execute("Update 用户 set 银两=银两+"&betcash&",赢次数=赢次数+1,赢钱=赢钱+"&betcash&",赌次数=赌次数+1,魅力=魅力-1 where 姓名='"&sjjh_name&"'")%>
</font>
<hr size="1" width="250">
<font size="2"> 你现在有银子：<b><font color="#CC0000"><b><%=(betcash+nowcash)%></b>
</font>两 </b>魅力: <%=(nowmeili-1)%><%end if%> </font>
<hr size="1" width="250">
</td>
</tr>
<tr>
<td align="right" colspan="3"><font size="2"><a href="betindex.asp">赌场首页</a></font></td>
</tr>
</table>
<p><%=Application("sjjh_chatroomname")%>版权所有</p>
</div>
</body>
</html>
<%rs.close
conn.close
set conn=nothing
set betcash=nothing
set nowcash=nothing
end if
end if
end if
%>
