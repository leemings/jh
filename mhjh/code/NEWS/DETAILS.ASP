
<html>
<head>
<title>[魔幻江湖]之[官府公告]</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet>
</head>
<%
id=request.querystring("id")
%><!-- #include file="connect.asp" --><%
rs.Open "Select * From data where id="&id, conn,3,3%>
<body bgcolor="#FFFFFF">
<div align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolordark="#FFFFFF" bordercolorlight="#000000">
<tr>
<td height="16" width="470"><font color="#CC0000">[公告标题]：<%=rs("name")%></font>[<%=rs("time")%>]
<hr noshade size="1">
</td>
</tr>
<tr>
<td width="470">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td>[公告信息]：<br>
<%=rs("neirong")%>
<hr noshade size="1">
</td>
</tr>
</table>
</td>
</tr>
</table>

</div>
</body>
</html>
