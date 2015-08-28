<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
%>
<html>
<head>
<title>管理员名单</title>
<link rel=stylesheet href='../style.css'>
</head>
<body background='../chatroom/bg1.gif' leftmargin="0">
<div align="center">
<center>
<table width="450" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="363" valign="top">
<table width="1" border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td width="579">
<table width="317" align="center" cellspacing="1" border="0"
cellpadding="0" bordercolor="#000000">
<tr>
<td colspan="7" width="571">
<div align="center"><font color="#FF6600">官 府 网 管 资 料<br>
</font></div>
</td>
<tr bgcolor="#FFCC33">
<td align="center" width="78" height="21"> <font color="#0000FF">姓 名</font> </td>
<td align="center" width="76" height="21"> <font color="#0000FF">等 级</font> </td>
<td align="center" width="78" height="21" bgcolor="#FFCC00">
<font color="#0000FF">
状 态</font> </td>
<td align="center" width="78" height="21"> <font color="#0000FF">杀 人</font> </td>
<td align="center" width="79" height="21"> <font color="#0000FF">被 杀</font> </td>
<td align="center" width="79" height="21"> <font color="#0000FF">积 分</font> </td>
<td align="center" width="79" height="21"> <font color="#0000FF">道 德</font> </td>
</tr>
<%
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT 姓名,状态,等级,总杀人,被杀,积分,道德 FROM 用户 where 门派='官府' order by 等级 DESC"
Set Rs=conn.Execute(sql)
do while not rs.eof and not rs.bof
%>
<tr>
<td width="78" nowrap>
<div align="center">
<%=rs("姓名")%>  </div>
</td>
<td width="76" nowrap>
<div align="center">
<%=rs("等级")%>  </div>
</td>
<td width="78" nowrap>
<div align="center">
<%=rs("状态")%> </div>
</td>
<td width="78" nowrap>
<p align="center"><%=rs("总杀人")%>
</td>
<td width="79" nowrap>
<p align="center"><%=rs("被杀")%>
</td>
<td width="79" nowrap>
<p align="center"><%=rs("积分")%>
</td>
<td width="79" nowrap>
<p align="center"><%=rs("道德")%>
</td>
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
</td>
</tr>
<tr>
<td width="579">
<ul>
<li>管理员只管理刷屏、骂人，及违反中国有关互联网规定的对象！</li>
<li>发现管理员有乱用权力的，应当及时举报投诉，站长将在查清事实后进行严肃的处理！</li>
</ul>
</td>
</tr>
</table>
</td>
</tr>
</table>
</center>
</div>
</body>
</html>
