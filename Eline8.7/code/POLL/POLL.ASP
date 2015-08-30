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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 官府公告 from 投票",conn
gg=rs("官府公告")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>投票系统♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../css.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="10" valign="bottom"><img src="../images/tp1.gif" width="96" height="23"></td>
<td width="768"><img src="../images/tp4.jpg" width="158" height="100"></td>
<td width="768">&nbsp;</td>
</tr>
<tr bgcolor="#000000">
<td width="10">&nbsp;</td>
<td width="768"><img src="../images/tp3.jpg" width="158" height="60"></td>
<td width="768" valign="bottom" align="right"><img src="../images/tp2.gif" width="80" height="18"></td>
</tr>
</table>
<br>
【<font color="#FF6600">投票系统</font>】
<table width="100%" border="0" align="center" bgcolor="#33CCFF">
  <tr>
<td>
      <blockquote><b><font color="#FF0000">[官府掌门公告]</font></b><font size="2"><br>
        <%=gg%></font>
</blockquote>
</td>
</tr>
</table>

<blockquote>
<p>● <a href="pollview.asp">投票 / 查看投票结果</a><br> 
<% if sjjh_grade>=5 then %>
● <a href="polluser.asp">查阅已投票者</a><br>
<%end if%> 
<% if sjjh_grade>=10 and sjjh_name=application("sjjh_user") then %>
● <a href="pollchangenotify.asp">更改官府掌门公告</a><br> 
● <a href="polltime.asp">设置投票起止时间</a><br> 
● <a href="pollvalue.asp">设置用户投票最低战斗等级</a><br> 
● <a href="pollcand.asp">候选人管理</a><br> 
● <a href="pollclear.asp">清除所有候选人的得票数，重新投票</a><br> 
● <a href="pollreset.asp">复位投票系统，清除数据</a></p> 
<%end if%>
</blockquote>
</body>
</html>