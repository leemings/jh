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
rs.open "select 会员金卡 from 用户 where 姓名='"&sjjh_name&"'",conn
hyfei=rs("会员金卡")
rs.close
rs.open "select * from b where b='卡片' and h>0 order by h"
%>
<html>
<head>
<title>会员卡片♀wWw.happyjh.com♀</title>
<link rel="stylesheet" href="../../chat/READONLY/Style.css">
</head>
<body bgcolor=#FFFFFF background="../bgcheetah.gif" text="#000000">
<BR>
<p align="center"><font size="6" face="隶书">江湖卡片屋</font></p>
<div align="center">江湖卡片只有江湖特殊会员才可以购买<br>
<font color=blue><%=sjjh_name%></font>现在有会员金卡:<font color=red><b><%=hyfei%></b></font>元整
</div>
<table border=1 align=center width=600 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#3399FF">
<tr bgcolor="#CC9900">
<td colspan="4" height="32">
<div align="center"><font color="#000000">会员卡片</font></div>
</td>
</tr>
<tr>
    <td width="12%" height="11"> 
      <div align="center">卡片名称</div>
</td>
    <td width="67%" height="11"> 
      <div align="center">功 能</div>
</td>
    <td width="7%" height="11"> 
      <div align="center">会费</div>
</td>
    <td height="11" width="14%"> 
      <div align="center">操 作</div>
</td>
</tr>
</table>
<table border=1 align=center width=600 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<%
do while not rs.bof and not rs.eof
%>
<tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';">
    <td width='12%' >
<div align="center"><%=rs("a")%>
</div>
</td>
    <td width='67%' ><%=rs("c")%></td>
    <td width='8%' >
<div align="center"><%=rs("h")%>元</div>
</td>
    <td  align="center" width="13%"><a href="buycard.asp?id=<%=rs("id")%>">购买</a></td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</body>
</html>
