<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if Session("usepro")<>"3" then
Session("usepro")=""
response.redirect "index.asp"
end if
%>
<HTML>
<HEAD>
<title>���Ĺأ��Ĵ�С</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet></HEAD>
<body bgcolor="#000000" oncontextmenu=self.event.returnValue=false text="#FFFFFF" >
<div align="left"></div>

<div align=center>
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td><font size="2" class=c><font size="3"><b> </b></font></font>
<div align=center><b>˾ͽ���</b> <font color="#FF0033">vs</font> <b><%=username%></b></div>

</td>
</tr>
</table>
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td>
<div align="center">һ�ֶ���Ӯ�ǣ����������</div>
</td>
</tr>
</table>
<form method="POST" action="ggg.asp">
<table border=1 cellspacing=0 cellpadding=0 align="center" width="350" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr align="center" bgcolor="#FFFFFF">
<td width="33%"><font size="2" color="#000000"><img src="run.gif" width="38" height="36"></font></td>
<td width="33%"><font size="2" color="#000000"><img src="run.gif" width="38" height="36"></font></td>
<td width="33%"><font size="2" color="#000000"><img src="run.gif" width="38" height="36"></font></td>
</tr>
<tr bgcolor="#336633">
<td width="960" colspan="3" height="29"><font size="2" class="c" color="#000000"><b>&nbsp;&nbsp;<font color="#FFFFFF">��ѡ��</font></b></font></td>
</tr>
<tr bgcolor="#FFFFFF">
<td align=center colspan="3">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
<tr align="center">
<td width="50%"><img src="big.gif" width="46" height="40"></td>
<td width="50%"><img src="small.gif" width="46" height="40"></td>
</tr>
<tr align="center">
<td width="50%">
<input type="radio" name="select" value="big" checked>
</td>
<td width="50%">
<input type="radio" name="select" value="small">
</td>
</tr>
</table>
</td>
</tr>
<tr>
<td align=center colspan="3">&nbsp;</td>
</tr>
<tr>
<td align=center colspan="3"> <font size="2" color="#000000">
<input type="submit" value="���ɣ���˭���ˡ�" name="B1" >
<input type="reset" value="��Ҫ����һ�£���" name="B2" >
</font></td>
</tr>
</table>
</form>
<p align="center">&nbsp;</p>
</div>
</BODY>
</HTML>

