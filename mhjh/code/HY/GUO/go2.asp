<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if Session("usepro")<>"1"  then
Session("usepro")=""
response.redirect "index.asp"
end if
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
%>
<html>
<head>
<title>�ڶ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet></head>
<body bgcolor="#000000" oncontextmenu=self.event.returnValue=false >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="595">
<div align="left"><b><font color="#FFFFFF">�ڶ�����һλPPMMҮ~������</font></b></div>
</td>
<td width="165" rowspan="2" valign="top"><font color="#FFFFFF"><img src="jh0.jpg" width="153" height="172"></font></td>
</tr>
<tr>
<td width="595" valign="top">
<p style="line-height: 200%; margin: 20"> <font color="#FFFFFF">PPMM���ˣ�С˧�磬�Ҹ����Ƕ��³ǵ����ҡ��������ҵĸ��װ��ֵڶ��أ���Ҫ��ѽ����������ʮ�����⣬ȫ������˾Ϳ��Թ�ȥ��Ҫ��Ҫ����ѽ��</font></p>
</td>
</tr>
<tr valign="top">
<td colspan="2">
<p><font color=red><%=username%>:</font><a href="answer.asp">�ðɣ��ҵ���ſ��ǽ�����������š����Ѿ��Ӹ��ҡ�</a><br>
</p>
</td>
</tr>
</table>
</body>
</html>




