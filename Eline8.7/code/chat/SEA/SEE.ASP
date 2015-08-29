<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
name=sjjh_name
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Eline_global.asa")
Conn.Open Connstr
set rs=conn.execute("select * from 航海物品 where 拥有者='"& name &"'")
%>
<html>
<head>
<title>商品一缆♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 12px}
input {  font-family: "宋体"; font-size: 12px; border: thin #333333 dotted; color: #660000; background-color: #FFFFCC}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#660000" background="../../images/8.jpg">
<br>
<br>
<table width="60%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr align="center"> 
    <td height="22"><big><strong></strong></big>你现有商品</td>
  </tr>
</table>
<table width="60%" border="1" cellspacing="4" cellpadding="4" align="center">
  <tr align="center" bgcolor="#660000"> 
    <td width="25%"><font color="#FFFFFF">丝绸</font></td>
    <td width="25%"><font color="#FFFFFF">人参</font></td>
    <td width="25%"><font color="#FFFFFF">珠宝</font></td>
    <td width="25%"><font color="#FFFFFF">瓷器</font></td>
  </tr>
<% do while not rs.eof %>
  <tr align="center"> 
    <td width="25%"><%=rs("丝绸")%></td>
    <td width="25%"><%=rs("人参")%></td>
    <td width="25%"><%=rs("珠宝")%></td>
    <td width="25%"><%=rs("瓷器")%></td>
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
<table width="60%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr align="center"> 
    <td height="32"><br>『E线江湖』<br>
    </td>
  </tr>
</table>
</body>
</html>