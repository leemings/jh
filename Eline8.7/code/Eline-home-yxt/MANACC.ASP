<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
sjjh_name=sjjh_name
grade=Int(sjjh_grade)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim tmprs
tmprs=conn.execute("Select count(*) As 数量 from 用户 where times<3 and CDate(lasttime)<now()-10")
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
user1=musers
tmprs=conn.execute("Select count(*) As 数量 from 用户 where allvalue<=5 and CDate(lasttime)<now()-10")
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
user2=musers
tmprs=conn.execute("Select count(*) As 数量 from 用户 where CDate(lasttime)<now()-90")
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
user3=musers
%>
<html>
<head>
<title>帐号管理♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body bgcolor="#FFFFFF" class=p150 background="../jhimg/bk_hc3w.gif">
<div align="center">『E线江湖』账号管理 <br>
</div>
<table width="390" border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="3" align="center">
<tr bgcolor="#0099FF">
<td width="388"><font face="Wingdings" color="#FFFFFF">1</font><font color="#FFFFFF">清理帐号</font></td>
</tr>
<tr>
<td class=p150 width="388" height="34"> ◇ <a href="manaccdel7.asp?page=1">清理60天内只用过一次的帐号</a><br>
◇ <a href="maingl.asp">更新所有用户管理等级为1(掌门5级管理员不变)</a><br>
◇ <a href="maingl1.asp">更新所有的人配偶为无</a><br>
◇ <a href="maingl2.asp">更新所有现金/银两为负用户为0</a></td>
</tr>
</table>
<br>
<br>
<table width="390" border="1" cellspacing="0" cellpadding="3" align="center">
  <tr bgcolor="#0099FF"> 
    <td><font face="Wingdings" color="#FFFFFF">1</font><font color="#FFFFFF">数据整理</font></td>
  </tr>
  <tr> 
    <td height="26">◇<a href="gl1.asp">删除10天前,仅登陆3次用户：<font color="#0000FF"><b><%=user1%></b>个</font></a> 
      <br>
      ◇<a href="gl2.asp">删除10天前泡点allvalue小于5用户：<b><font color="#0000FF"><%=user2%></font></b><font color="#0000FF">个</font></a> 
	<br>
     ◇<a href="gl3.asp">删除90天从未登陆的用户有：<font color="#0000FF"><b><%=user3%></b>个</font></a> 
     <br>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<hr noshade size="1" color=009900>
<div align=center class=cp><%Response.Write "序列号：<font color=blue>" & Application("sjjh_sn") & "</font>，授权给：<font color=blue>" & Application("sjjh_user") & "</font><br><font color=999999>" & Application("sjjh_copyright") & "</font>"%></div>
</body>
</html>