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
randomize
m=int(6*rnd+1)
if m>3 then
if request.form("select")="big" then
Randomize
m1 = Int(6 * Rnd + 1)
Randomize
m3 = Int(6 * Rnd + 1)
Randomize
m2 = Int(5 * Rnd + 1)
else
Randomize
m1 = Int(6 * Rnd + 1)
Randomize
m3 = Int(6 * Rnd + 1)
Randomize
m2 = Int(7 * Rnd + 1)
if m2>6 then m3=6 end if
end if
else
Randomize
m1 = Int(6 * Rnd + 1)
Randomize
m3 = Int(6 * Rnd + 1)
Randomize
m2 = Int(6 * Rnd + 1)
end if



if request.form("select")="big" then
if m1+m2+m3>9 then
mm=int(6*rnd+1)
if mm=1 or mm=2 or mm=3 or mm=4 then
do while m1+m2+m3>=10
m1 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
m2 = Int(6 * Rnd + 1)
loop
end if
end if
else
if m1+m2+m3<9 then
mm=int(6*rnd+1)
if mm=1 or mm=2 or mm=3 or mm=4 then
do while m1+m2+m3<9
m1 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
m2 = Int(6 * Rnd + 1)
loop
end if
end if
end if
dim betcash,nowcash,username
Session("ds")=""
%>

<head>
<title>���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet>
</head>

<body bgcolor="#000000" text="#FFFFFF">
<font color="#FFFFFF"><br>
</font>
<div align="center">
<p><font size="3" class="c" color="#FFFFFF"><b><font color="#000000">�����</font></b></font></p>
<p><font size="2" class="c"><b><font color="#FF0033">��Ѻ����</font><br>
<font color="#CC0000"><%if request.form("select")="big" then%></font><font color="#CC0000">
</font></b></font><img src="big.gif" width="46" height="40"><font size="2" class="c"><b><font color="#CC0000"><%else%><img src="small.gif" width="46" height="40"></font><font size="2" class="c" color="#CC0000"><%end if%></font></b></font><%if (m1+m2+m3)>9 then%>
<table border="1" cellspacing="0" cellpadding="3" align="center" width="400" bordercolordark="#FFFFFF">
<tr bgcolor="#336633">
<td colspan="3" align="center"><font class="c" size="2" color="#FFFFFF"><b>���:</b></font><font size="2" color="#FFFFFF"><b><%=(m1+m2+m3)%>�㣬��</b></font></td>
</tr>
<tr>
<td width="33%" align="center"><font size="2"><img src="images/<%=m1%>.gif" width="32" height="31"></font></td>
<td width="33%" align="center"><font size="2"><img src="images/<%=m2%>.gif"></font></td>
<td width="34%" align="center"><font size="2"><img src="images/<%=m3%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"> <font size="2" color="#FFFFFF"><br>
<%if request.form("select")="small" then%></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right" width="22%"><font size="2" color="#000099"><b>˾ͽ��⣺</b></font></td>
<td colspan="2" width="78%"><font size="2"><a href="javascript:self.close()"><font color="#FFFFFF">��Ӯ�ˣ��㻹���Ȼ�ȥ���������ɡ�</font></a></font></td>
</tr>
<tr>
<td align="right" width="22%"><font size="2" color="#CC0000"><b>��˵��</b></font></td>
<td colspan="2" width="78%"><font size="2">�浹ù!!</font></td>
</tr>
</table>
<%
Session("usepro")=""
%>
<font size="2" color="#FFFFFF"><%else%>
</font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right" width="22%"><font size="2" color="#000099"><b>˾ͽ��⣺</b></font></td>
<td colspan="2" width="78%"><font size="2"><a href="go8.asp"><font color="#FFFFFF">�����ˣ��㰴����������ذɡ�</font>
</a></font></td>
</tr>
<tr>
<td align="right" width="22%"><font size="2" color="#CC0000"><b>��˵��</b></font></td>
<td colspan="2" width="78%"><font size="2">��Ȼ���ҽ�����������~��</font></td>
</tr>
</table>
<%
Session("usepro")="4"
%>
<%end if%>
<p>&nbsp;</p>
</td>
</tr>
<tr bgcolor="#FFCCCC">
<td align="right" colspan="3" height="9">&nbsp;</td>
</tr>
</table>
<font size="3"><%else%></font>
<table border="1" cellspacing="0" cellpadding="3" align="center" width="400" bordercolordark="#FFFFFF">
<tr bgcolor="#336633">
<td colspan="3" align="center"><font class="c" size="2" color="#FFFFFF"><b>���:</b></font><font size="2" color="#FFFFFF"><b><%=(m1+m2+m3)%>�㣬С</b></font></td>
</tr>
<tr>
<td width="33%" align="center"><font size="2"><img src="images/<%=m1%>.gif" width="32" height="31"></font></td>
<td width="33%" align="center"><font size="2"><img src="images/<%=m2%>.gif"></font></td>
<td width="34%" align="center"><font size="2"><img src="images/<%=m3%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"> <font size="2" color="#FFFFFF"><br>
<%if request.form("select")="big" then%></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right" width="22%"><font size="2" color="#000099"><b>˾ͽ��⣺</b></font></td>
<td colspan="2" width="78%"><font size="2"><a href="javascript:self.close()"><font color="#FFFFFF">��Ӯ�ˣ��㻹���Ȼ�ȥ���������ɡ�</font></a></font></td>
</tr>
<tr>
<td align="right" width="22%"><font size="2" color="#CC0000"><b>��˵��</b></font></td>
<td colspan="2" width="78%"><font size="2">�浹ù!!</font></td>
</tr>
</table>
<%
Session("usepro")=""
%>
<font size="2" color="#FFFFFF"><%else%>
</font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right" width="22%"><font size="2" color="#000099"><b>˾ͽ��⣺</b></font></td>
<td colspan="2" width="78%"><font size="2"><a href="go8.asp"><font color="#FFFFFF">�����ˣ��㰴����������ذɡ�</font>
</a></font></td>
</tr>
<tr>
<td align="right" width="22%"><font size="2" color="#CC0000"><b>��˵��</b></font></td>
<td colspan="2" width="78%"><font size="2">��Ȼ���ҽ�����������~��</font></td>
</tr>
</table>
<%
Session("usepro")="4"
%>
<%end if%>
</td>
</tr>
<tr bgcolor="#FFCCCC">
<td align="right" colspan="3" height="9">&nbsp;</td>
</tr>
</table>
<%end if%></div>

</body>

