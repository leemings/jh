<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>兵器铺♀wWw.51eline.com♀</title>
<style></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
</head>

<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table border="1" cellspacing="1" cellpadding="0" align="center" width="80%" bordercolor="#000000">
<tr>
<td height="9" width="14%">
<div align="center"><a href="binqi.asp"><font color="#0000ff">手持刀剑</font></a></div>
</td>
<td height="19" width="14%" rowspan="2">
<div align="center"><a href="head.asp"><font color="#0000ff">头盔</font></a></div>
</td>
<td height="19" width="14%" rowspan="2">
<div align="center"><a href="body.asp"><font color="#0000ff">盔甲</font></a></div>
</td>
<td height="19" width="14%" rowspan="2">
<div align="center"><a href="foot.asp"><font color="#0000ff">双脚</font></a></div>
</td>
<td height="19" width="14%" rowspan="2">
<div align="center"></div>
<div align="center"><a href="anqi.asp"></a><a href="other.asp"><font color="#0000ff">装饰</font></a></div>
</td>
<td height="19" width="14%" rowspan="2">
<div align="center"><a href="anqi.asp"><font color="#0000ff">暗器</font></a></div>
</td>
<td height="19" width="14%" rowspan="2">
<div align="center"><a href="addbq.asp"><font color="#0000ff">添加装备</font></a></div>
</td>
</tr>
<tr>
<td height="9" width="14%">
<div align="center"><a href="dunpai.asp"><font color="#0000ff">手持护具</font></a></div>
</td>
</tr>
</table>
<p align="center">暗器管理</p>
<br>
<table border="1" bgcolor="#006699" align="center" width="600" cellpadding="0"
cellspacing="1" bordercolor="#000000">
<tr>
<td height="17" colspan="4" bgcolor="#336633">
<p align="center">现 有 暗 器</p>
<tr>
<td height="18" bgcolor="#C4DEFF">
<div align="center"> <font size="2" color="#000000">暗器名</font> </div>
</td>
<td bgcolor="#C4DEFF" height="18">
<div align="center"> <font size="2" color="#000000">功 能</font> </div>
</td>
<td height="18" bgcolor="#C4DEFF">
<div align="center"> <font size="2" color="#000000">售 价</font> </div>
</td>
<td height="18" bgcolor="#C4DEFF">
<div align="center"> <font size="2" color="#000000">操 作</font> </div>
</td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("sjjh_usermdb")
sql="SELECT * FROM b where b='暗器'"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr>
<td height="11"><font color="#FFFFFF" size="2"><%=rs("a")%> </font>
<td height="11"><font color="#FFFFFF" size="2">伤内力<%=abs(rs("d"))%>，伤体力<%=abs(rs("e"))%></font></td>
<td height="11"><font color="#FFFFFF" size="2"><%=rs("h")%>两</font></td>
<td height="11">
<div align="center"><a href="modifyyao.asp?wupinid=<%=rs("id")%>"><font color="#FFFFFF" size="-1">修改</font></a>
<font size="-1"><a href="del.asp?id=<%=rs("id")%>"><font color="#FFFFFF">删除</font></a></font></div>
</td>
</tr>
<%
rs.movenext
loop
%>
</table>
</body>
</html>
