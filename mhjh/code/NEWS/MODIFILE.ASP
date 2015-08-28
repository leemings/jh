

<html>
<head>
<title>官府公告</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet>
</head>
<!-- #include file="connect.asp" -->
<%id=request.querystring("id")
rs.Open "Select * From data where id="&id, conn,3,3%>
<body bgcolor="#FFFFFF">
<div align="center">
<form method="POST" action="changexx.asp">
<table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolordark="#FFFFFF" bordercolorlight="#000000"><tr><td bgcolor="#33CCFF" colspan=2 align=center>信息修改</td></tr>
<tr>
<td width="10%">标题</td><td>
<input type="text" name="name" size="60" value="<%=rs("name")%>">
<input type=hidden name=id value="<%=rs("id")%>">
</td></tr><tr><td>
内容</td><td>
<textarea name="neirong" rows="10" cols="80"><%=rs("neirong")%></textarea>
</td>
</tr>
<tr><td colspan=2 align=center><input type="submit" name="Submit" value="提交">
<input type="reset" name="Submit2" value="清除"></td>
</tr>
</table>
</div>
</body>
</html>
