<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>牢房重地，闲人莫入</title>
<link rel="stylesheet" type="text/css" href="style.css">
<style type="text/css">td           { font-family: 宋体; font-size: 9pt }
body         { font-family: 宋体; font-size: 9pt }
select       { font-family: 宋体; font-size: 9pt }
a            { color: #FFC106; font-family: 宋体; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 宋体; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF">
<br>
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="318">&nbsp;</td>
<td width="460"><font color="#FF3300">官府牢房在押人员名单</font><br>
作了坏事，官府决不轻饶！ <br>
3天内必须救出，否则斩首示众 </td>
</tr>
</table>
<br>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
<td bgcolor="#000000">&nbsp;</td>
</tr>
<tr bgcolor="#990000">
<td><img src="../images/lf1.gif"><img src="../images/lf2.gif"><img src="../images/lf3.gif" height="80"></td>
</tr>
<tr>
<td bgcolor="#000000">&nbsp;</td>
</tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="778">
<tr>
<td>
<div align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
align="center" >
<tr>
<td colspan="2" valign="top" width="100%" height="110">
<div align="center"> <br>
<table border="1" cellspacing="1" cellpadding="0" width="100%"
bordercolor="#000000" align="center">
<tr>
<td width="15%" nowrap>
<div align="center"> 犯人 </div>
<td width="15%" nowrap>
<div align="center"> 门派 </div>
</td>
<td width="11%">
<div align="center"> 身份 </div>
</td>
<td width="27%">
<div align="center"> 请 求 </div>
</td>
<td width="23%">&nbsp;</td>
<td width="9%" nowrap>　</td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 where 状态='狱' or 状态='牢'",conn
do while not rs.bof and not rs.eof
%>
<tr>
<td width="15%" height="31"><%=rs("姓名")%>
<td width="15%" height="31"><%=rs("门派")%></td>
<td width="11%" nowrap height="31"><%=rs("身份")%></td>
<%if rs("状态")="狱" then %>
<td width="27%" height="31">
<div align="center"> <a href="shifang.asp?name=<%=rs("姓名")%>">交钱5000万两释放</a>|<a href="jieyu.asp?name=<%=rs("姓名")%>"
title="成功率只有30%"><font color="#FF0000">劫狱</font></a>
<%if sjjh_grade>=9 then%>
|<a href="zhanshou.asp?name=<%=rs("姓名")%>"><font color="#FF0000">斩首</font></a>
<%end if%> <%if sjjh_grade>=6 then%>
|<a href="shuchu.asp?id=<%=rs("id")%>"><font color="#FF0000">释放</font></a>
<%end if%> </div>
</td>
<% else
if rs("登录")>now() then%>
<td width="23%" height="31">
	<div align="center"> <font color="#FF0033">不能保释</font>|释放时间：<%=rs("登录")%>
	</div>
	</td>
<%else%>
	<td width="9%" height="31">
	<div align="center"> <font color="#FF0033">已释放</font>
	</div>
	</td>
	<% conn.execute("update 用户 set 状态='正常',登录=now() where 姓名='"&rs("姓名")&"'")
end if%>
<%end if%>
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
<p>少干点坏事！ ！：（ </p>
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
<div align="center">
</div>

</body>

</html>

<html></html>