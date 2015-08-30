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

%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>色魔排行♀wWw.happyjh.com♀</title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<style type="text/css">
<!--
body {  font-size: 9pt; color=000000}
td {  font-size: 9pt; color=000000}
A {text-decoration: none; color: #00ffff; font-family: "宋体"; font-size: 9pt}
a:hover      { color: ffffFF; font-family: 宋体; font-size: 9pt; text-decoration:
underline }
A:active {  font-family: "宋体"; font-style: normal;  color: ffffFF; text-decoration: underline; font-size: 9pt}
br {font-size:6pt; line-height:80%;}
-->
</style>
</head>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
const MaxPerPage=20
dim totalPut
dim CurrentPage
dim TotalPages
dim i,j
%>

<body text="#000000" background="../jhimg/bk_hc3w.gif">
<div align="center">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" bgcolor="#336633" height="16">
<tr>
<td width="100%"><font color="#ffffff" size="2">20大色魔排行</font></td>
</tr>
</table>
<%
dim sql
dim rs
dim filename
sql="select * from 用户 where 状态='正常' order by 结婚次数 desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<p align='center'>没有可排行的对象 </p>"
else
%>
<table border="1" cellspacing="1" width="100%" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="0" bordercolor="#000000">
<tr>
<td align="center" height="16" bgcolor="#336633" width="15%"><font color="#FFFFFF">姓名</font></td>
<td align="center" height="16" bgcolor="#336633" width="15%"><font color="#FFFFFF">性别</font></td>
<td align="center" height="16" bgcolor="#336633" width="15%"><font color="#FFFFFF">门派</font></td>
<td align="center" height="16" bgcolor="#336633" width="15%"><font color="#FFFFFF">身份</font></td>
<td align="center" height="16" bgcolor="#336633" width="7%"><font color="#FFFFFF">结婚次数</font></td>
<td align="center" height="16" bgcolor="#336633" width="8%"><font color="#FFFFFF">登录次数</font></td>
<td align="center" height="16" bgcolor="#336633" width="11%"><font color="#FFFFFF">总积分</font></td>
<td align="center" height="16" bgcolor="#336633" width="14%"><font color="#FFFFFF">月积分</font></td>
</tr>
<%do while not rs.eof%>
<tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';">
<td align="center" width="15%"><font color="#0000FF"><%=rs("姓名")%> </td>
<td align="center" width="15%"><font color="#0000FF"><%=rs("性别")%></font></td>
<td align="center" width="15%"><font color="#0000FF"><%=rs("门派")%> </td>
<td align="center" width="15%"><font color="#0000FF"><%=rs("身份")%> </td>
<td align="center" width="7%"><font color=ff0000><%=rs("结婚次数")%></font> </td>
<td align="center" width="8%"><font color="#0000FF"><%=rs("times")%></font></td>
<td align="center" width="11%"><font color="#0000FF"><%=rs("allvalue")%></font></td>
<td align="center" width="14%"><font color="#0000FF"><%=rs("mvalue")%> </td>
</tr>
<%
rs.movenext
filename=filename+1
if filename>19 then Exit Do
loop
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</div>
<div align="center"><font color=#999999 size=2><%Response.Write "序列号：<font color=blue>" & Application("sjjh_sn") & "</font>，授权给：<font color=blue>" & Application("sjjh_user") & ""%></font>
</div>
</body>
</html>
