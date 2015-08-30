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
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>--我的徒弟</title>
<style type="text/css">
<!--
p            { line-height: 20px; font-size: 9pt }
table        { font-size: 9pt }
a:link       { color: #FF9900; text-decoration: none }
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" vlink="#FF9900" topmargin="0"
leftmargin="0" background="../jhimg/bk_hc3w.gif">
<p align="center">
<%
id=request("id")
if instr(id,"官")<>0 then
		Response.Write "<script language=JavaScript>{alert('严重警告，不要搞乱');window.close();}</script>"
		Response.End
end if
my=request("my")
%>
<font color="#CC0000" face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font>
<table border="0" width="500" cellspacing="0" cellpadding="0"
background="bg.gif" align="center">
<tr align="center">
<td background="top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">我 的 徒 弟</font></b></font></td>
</tr>
<tr align="center">
<td>
<table width="497" border="1" cellpadding="0" cellspacing="0" bordercolor="#FF9900" height="13">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 where 师傅='" & sjjh_name &"'",conn
%>
<td width="54">
<div align="center"><font color=#ffff00>姓名</font></div>
</td>
<td width="27">
<div align="center"><font color=#ffff00>性别</font></div>
</td>
<td width="71">
<div align="center"><font color=#ffff00>身份</font></div>
</td>
<td width="83">
<div align="center"><font color=#ffff00>注册时间</font></div>
</td>
<td width="44">
<div align="center"><font color=#ffff00>登录</font></div>
</td>
<td width="52">
<div align="center"><font color=#ffff00>总积分</font></div>
</td>
<td width="49">
<div align="center"><font color=#ffff00>师傅发钱</font></div>
</td>
<td width="49">
<div align="center"><font color=#ffff00>师傅指导</font></div>
</td>
</tr>
<%
do while not rs.bof and not rs.eof
%>
<tr>
<td width="54" height="30">
<div align="center"><a href="../yamen/mt.asp?action=<%=rs("姓名")%>" target="_blank"><font color="#FFFFFF"><%=rs("姓名")%></font></a></div>
</td>
<td width="27" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("性别")%></font></div>
</td>
<td width="71" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("身份")%></font></div>
</td>
<td width="83" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("regtime")%><br><%=rs("lasttime")%></font></div>
</td>
<td width="44" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("times")%></font></div>
</td>
<td width="52" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("allvalue")%></font></div>
</td>
<td width="49" height="30">
<div align="center"><a href="mp6.asp" target="_blank"><font color="#FFFFFF">师傅操作</font></a></div>
</td>
<td width="49" height="30">
<div align="center"><a href="mp7.asp" target="_blank"><font color="#FFFFFF">师傅操作</font></a></div>
</td>
</tr>
<%
Application.Lock
Application("sjjh_bais_sf")=to1
Application("sjjh_bais_td")=sjjh_name
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</td>
</tr>
<tr align="center">
<td background="top3.gif" width="500" height="28"><font color="#FFFFFF">『快乐江湖』</font></td>
</tr>
</table>
</body>
</html>

