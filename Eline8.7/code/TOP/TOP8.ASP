<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>门派基金排行♀wWw.51eline.com♀</title>
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
<td width="100%"><font color="#ffffff" size="2">这里可以显示出各门派的基金情况(基金来源为弟子自愿捐献)</font></td>
</tr>
</table>
<%
dim sql
dim rs
dim filename
sql="select * from p where a<>'官府' and a<>'游侠' order by h desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<p align='center'>没有可排行的对象 </p>"
else
%>
<table border="1" cellspacing="1" width="100%" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="0" bordercolor="#000000">
<tr>
<td align="center" height="16" bgcolor="#336633" width="13%"><font color="#FFFFFF">门派名</font></td>
<td align="center" height="16" bgcolor="#336633" width="7%"><font color="#FFFFFF">掌门</font></td>
<td align="center" height="16" bgcolor="#336633" width="9%"><font color="#FFFFFF">弟子数</font></td>
<td align="center" height="16" bgcolor="#336633" width="9%"><font color="#FFFFFF">门派基金数</font></td>
</tr>
<%do while not rs.eof%>
<tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';">
<td align="center" width="13%"><%=rs("a")%> </td>
<td align="center" width="7%"><font color="#0000FF"><%=rs("b")%></font></td>
<td align="center" width="9%"><%=rs("c")%> </td>
<td align="center" width="9%"><font color="#0000FF"><b><%=rs("h")%></b> </font> </td>
</tr>
<%
rs.movenext
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
