<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
%>
<html>
<head>
<title>会员名单</title>
<link rel=stylesheet href='../style.css'>
</head>
<body background='../chatroom/bg1.gif' leftmargin="0" oncontextmenu=self.event.returnValue=false>
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
<div align="center"><font color="#FF6600">魔幻江湖会员列表<br>

</font></div>
</td>
<tr bgcolor="#FFCC33">
<td align="center" width="78" height="21"> <font color="#0000FF"> 姓 名</font> </td>
<td align="center" width="76" height="21"> <font color="#0000FF"> 等 级</font> </td>
<td align="center" width="78" height="21" bgcolor="#FFCC00">
<font color="#0000FF">
状 态</font> </td>
<td align="center" width="78" height="21"> <font color="#0000FF">杀人</font> </td>
<td align="center" width="79" height="21"> <font color="#0000FF">被杀</font> </td>
<td align="center" width="79" height="21"> <font color="#0000FF">积分</font> </td>
<td align="center" width="79" height="21"> <font color="#0000FF">会员时间</font> </td>
</tr>
<%
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT 姓名,等级,状态,总杀人,被杀,积分,会员时间 FROM 用户 where 会员=True order by 等级 DESC"
Set Rs=conn.Execute(sql)
do while not rs.eof and not rs.bof
sql="Select count(*) from 用户 where 会员=True"
set rst=conn.execute(sql)
renshu=rst(0)
%>
</center>
<tr>
<td width="78" nowrap>
<div align="center">
<p align="center">
<%=rs("姓名")%> </div>
</td>
<center>
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
</center>
<td width="79" nowrap>
<p align="left"><%=rs("积分")%>
</td>
<td width="79" nowrap>
<p align="left"><%=rs("会员时间")%>
</td>
</tr>
<center>
<%
rs.movenext
loop
%>
</table>
</center>
</td>
</tr>
<tr>
<td width="579">
<div align="center">　共<%=renshu%>人 </div>
<p>&nbsp;&nbsp;&nbsp;
无论是交费会员还是其它网管、掌门，如有违反规定，将永久不得录用，因被开除的会员我们不退回购买时交纳的费用，请大家注意！<br>
</td>
</tr>
</table>
</td>
</tr>
</table>
</div>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>