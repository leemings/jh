<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title></title>
<link rel="stylesheet" href="setup.css">
</head>
<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<br>
<table border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000" width="400">
<tr>
<td height="19" colspan="3">
<div align="center">
<div align="center"><a href="addyao.asp"><font color="#0000ff">添加药品</font></a></div>
</div>
</td>
<td height="19" width="124">
<div align="center"><a href="yaopu.asp"><font color="#0000ff">补药</font></a></div>
</td>
<td height="19" colspan="2" width="133">
<div align="center"><a href="yaopu2.asp"><font color="#0000ff">毒药</font></a></div>
</td>
</tr>
</table>
<p align="center"> &nbsp;
<script src="../hcjs/2.js"></script>
<center>
<table border=1 bgcolor="#006699" align=center width=550 cellpadding="0" cellspacing="1" bordercolor="#000000">
<tr bgcolor="#336633">
<td width="538" colspan="4">
<p align="center"><span style="letter-spacing: 1">现有药品</span></p>
</td>
</tr>
<tr bgcolor=#C4DEFF>
<td width="117"><font color="#000000"><span style="letter-spacing: 1">药品名</span></font></td>
<td width="224"><font color="#000000"><span style="letter-spacing: 1">功能</span></font></td>
<td width="62"><font color="#000000"><span style="letter-spacing: 1">售价</span></font></td>
<td width="135"><font color="#000000"><span style="letter-spacing: 1">操作</span></font></td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("aqjh_usermdb")
sql="SELECT * FROM b where  b='毒药'"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr>
<td width="117"><font color="#FFFFFF" size="-1"><span style="letter-spacing: 1"><%=rs("a")%>
</span> </font>
<td width="224"><span style="letter-spacing: 1"><font color="#FFFFFF" size="-1">补内力<%=rs("d")%>，补生命<%=rs("e")%></font></span></td>
<td width="62"><span style="letter-spacing: 1"><font color="#FFFFFF" size="-1"><%=rs("h")%>两</font></span></td>
<td width="135">
<div align="center"><font size="-1"><span style="letter-spacing: 1"><a href="modifyyao.asp?wupinid=<%=rs("id")%>"><font color="#FFFFFF">修改</font></a>
<a href="del.asp?id=<%=rs("id")%>"><font color="#FFFFFF">删除</font></a>
</span></font></div>
</td>
</tr>
<%
rs.movenext
loop
%>
</table>
</center>
</body>
</html>

<html></html>