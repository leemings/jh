<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_name=aqjh_name
grade=Int(aqjh_grade)
%>
<!--#include file="config.asp"-->
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
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
tmprs=conn.execute("Select count(*) As 数量 from 用户 where CDate(lasttime)<now()-60 and 会员等级=0")
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
user3=musers
%>
<html>
<head>
<title>帐号管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<div align="center">账号管理 <br>
</div>
<table width="390" border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="3" align="center">
<tr bgcolor="#0099FF">
<td width="388"><font face="Wingdings" color="#FFFFFF">1</font><font color="#FFFFFF">清理帐号</font></td>
</tr>
<tr>
<td class=p150 width="388" height="34"> ◇ <a href="manaccdel7.asp?page=1">清理７天内只用过一次的帐号</a><br>
◇ <a href="maingl3.asp">查看１个月内未登录的会员</a><br>
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
     ◇<a href="gl3.asp">删除二个月从未登陆的用户有：<font color="#0000FF"><b><%=user3%></b>个</font></a> 
     <br>
    </td>
  </tr>
</table>
</body>
</html>