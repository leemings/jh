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
if sjjh_grade<10 or sjjh_name<>application("sjjh_user") then
	Response.Write "<b>[����ʧ��]</b><p>��û��Ȩ�ޣ�"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="Select * from ͶƱ"
set rs=conn.Execute(sql)
dj=rs("�ȼ�")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>����ͶƱ��;���ֵ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
<!--
.p9 {line-height: 150%; font-size: 9pt;}
.p12 {line-height: 150%; font-size: 12pt;}
body {line-height: 150%;font-size : 12pt;}
td {line-height: 150%;font-size : 12pt;}
A  {text-decoration: none;}
A:Hover  {text-decoration : none;}
a:visited {  color: #0000FF}
-->
</style>
</head>
<body bgcolor="#FFFFEC" topmargin="0" leftmargin="0">
<table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#C0C0C0" height="1">
<tr>
<td width="100%" height="1"><font size="2">����λ��&gt;&gt;<a href="../jh.asp">�������</a>&gt;&gt;<font color="#0000FF"><a href="poll.asp">����ͶƱ</a></font>&gt;&gt;<a href="pollvalue.asp">����ͶƱ��;���ֵ</a></font></td>
</tr>
</table>
<h1 align="center"><font size="3" color="#0000FF">������ͶƱ���ս���ȼ���</font></h1>
<hr noshade size="1" color=009900>
<blockquote>
<p><b><font color="#FF0000">[����˵��]</font></b></p>
<p><font size="2">ֻ��ս���ȼ��ﵽ��ֵ���ϵ��û����������ͶƱ��</font></p>
</blockquote>
<table border="1" align="center" bgcolor="#E0E0E0" cellpadding="4" bordercolorlight="#000000" cellspacing="0" bordercolordark="#FFFFFF" width="250">
<form method="post" action="pollvalueok.asp">
<tr>
<td width="98"><font size="2">ͶƱ��͵ȼ���</font></td>
<td width="115"> <font size="2">
<input type="text" name="downvalue" size="10" style="font-size:12pt" value="<%=dj%>" maxlength="10">
</font>
</td>
</tr>
<tr>
<td align="right" width="98"><font size="2">ѡ�������</font></td>
<td align="center" width="115"> <font size="2">
<input type="submit" name="Submit" value="����" style="font-size:12pt">
<input type="button" value="����" style="font-size:12pt" onclick=javascript:history.go(-1)>
</font>
</td>
</tr>
</form>
</table>
<br>
<hr noshade size="1" color=009900>
</body>
</html>