<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name=""  then
Response.Redirect "../error.asp?id=440"
else
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


dim betcash,nowcash
nowcash=0
betcash=0
betcash=clng(request.form("splosh"))
if betcash<10 then
%>
<script language=vbscript>
MsgBox "��ʲô��Ц��������ô�ٵ�ע�ĸ�ʲô����"
location.href = "javascript:history.back()"
</script>
<%
elseif betcash>3000 then
%>
<script language=vbscript>
MsgBox "��ʲô��Ц��������ô��������Ʋ�����"
location.href = "javascript:history.back()"
</script>
<%
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("aqjh_usermdb")
sql="Select  * from �û� where ����='"&aqjh_name&"'"
set rs=conn.Execute(sql)
nowcash=rs("����")
if betcash>nowcash then
%>
<script language=vbscript>
MsgBox "��û�и�����Ϻɰ���������ô�����ӣ����ٺٺ٣���ȥ����ȥ�ɣ������"
location.href = "javascript:history.back()"
</script>
<%
else

nowmeili=rs("����")
if nowmeili<10 then
%>
<script language=vbscript>
MsgBox "��������Ŀǰ����������10�ˣ������Ͽ�����ȥѽ���Ժ������ɣ�"
location.href = "javascript:history.back()"
</script>
<%
else
%>
<html>
<head>
<title>�����Ĺ� - �Ĵ�С</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
--></style>
</head>

<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
<p><font size="3" class="c" color="#000000"><b>�� �� - �� �� С<br>
</b></font><font size="2" class="c"><b><font color="#FF0033">��Ѻ����</font><font color="#CC0000">
<%if request.form("select")="big" then%>
</font><font color="#CC0000"> </font></b></font><img src="../jhimg/bbs/big.gif" width="46" height="40"><font size="2" class="c"><b><font color="#CC0000">
<%else%>
<img src="../jhimg/bbs/small.gif" width="46" height="40"></font><font size="2" class="c" color="#CC0000">
<%end if%>
</font></b></font>
<%if (m1+m2+m3)>9 then%>
</p>
<table border="1" cellspacing="0" cellpadding="3" align="center" width="400" bordercolordark="#FFFFFF" bordercolor="#000000">
<tr bgcolor="#336633">
<td colspan="3" align="center"><font class="c" size="2" color="#FFFFFF"><b>���:</b></font><font size="2" color="#FFFFFF"><b><%=(m1+m2+m3)%>�㣬��</b></font></td>
</tr>
<tr>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m1%>.gif" width="32" height="31"></font></td>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m2%>.gif"></font></td>
<td width="34%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m3%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"> <font size="2" color="#FFFFFF">
<%if request.form("select")="small" then%>
</font><font size="2" color="#FFFFFF"><b><%=betcash%></b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right"><font size="2" color="#000099"><b>ׯ�ң�</b></font></td>
<td colspan="2"><font size="2">�Ǻǻ�Ҫ���� �ж�δ����Ŷ~��</font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>��˵��</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">������Ϊʲô���������ޣ��ҾͲ�������ô��ù~��</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td colspan="2"><font size="2"><a href="../jh.asp">�������ˣ��ϵ�ù�ɣ����ŵ����Ӿ�����~��</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%
conn.execute("Update �û� set ����=����-"&betcash&",�Ĵ���=�Ĵ���+1,����=����-1,ӮǮ=ӮǮ-"&betcash&" where ����='"&aqjh_name&"'")
%>
</font><font size="2" color="#FFFFFF"><b><%=(nowcash-betcash)%> </b><%=(nowmeili-1)%>
<%else%>
</font><font size="2" color="#FFFFFF"><b><%=betcash%> ��</b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right"><font size="2" color="#000099"><b>ׯ�ң�</b></font></td>
<td colspan="2"><font size="2">����ѽ~���͹�����Ҫ���� </font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>��˵��</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">��Ȼ�������ҽ�����������~��</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
            <td colspan="2"><font size="2"><a href="../welcome.asp">���þ��գ�����Ϊ��ɵ���ٺ٣�</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%conn.execute("Update �û� set ����=����+"&betcash&",Ӯ����=Ӯ����+1,ӮǮ=ӮǮ+"&betcash&",�Ĵ���=�Ĵ���+1,����=����-1 where ����='"&aqjh_name&"'")
%>
</font><font size="2" color="#FFFFFF"><b><%=(betcash+nowcash)%></b><%=(nowmeili-1)%></font><font size="2" color="#FFFFFF">
<%end if%>
</font> </td>
</tr>
<tr bgcolor="#FFCCCC">
<td align="right" colspan="3" height="2"><a href="betindex.asp"><font size="2">��</font><font size="2">����ҳ</font></a></td>
</tr>
</table>
<font size="3"><%else%></font>
<table border="1" cellspacing="0" cellpadding="3" align="center" width="400" bordercolordark="#FFFFFF" height="160">
<tr bgcolor="#336633">
<td colspan="3" align="center"><font class="c" size="2" color="#FFFFFF"><b>���:</b></font><font size="2" color="#FFFFFF"><b><%=(m1+m2+m3)%>�㣬С</b></font></td>
</tr>
<tr>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m1%>.gif"></font></td>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m2%>.gif"></font></td>
<td width="34%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m3%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"><font size="2" color="#FFFFFF">
<%if request.form("select")="big" then%>
</font> <font size="2" color="#FFFFFF"><b><%=betcash%></b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right"><font size="2" color="#000099"><b>ׯ�ң�</b></font></td>
<td colspan="2"><font size="2">�Ǻǻ�Ҫ���� �ж�δΪ��Ŷ~��</font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>��˵��</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">������Ϊʲô���������ޣ��ҾͲ�������ô��ù~��</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td colspan="2"><font size="2"><a href="../jh.asp">�������ˣ��ϵ�ù�ɣ����ŵ����Ӿ�����~��</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%conn.execute("Update �û� set ����=����-"&betcash&",�Ĵ���=�Ĵ���+1,����=����-1,ӮǮ=ӮǮ-"&betcash&" where ����='"&aqjh_name&"'")
%>
</font> <font size="2" color="#FFFFFF"><b><%=(nowcash-betcash)%></b><%=(nowmeili-1)%>
<%else%>
</font><font size="2" color="#FFFFFF"><b><%=betcash%> ��</b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC">
<tr>
<td align="right"><font size="2" color="#000099"><b>ׯ�ң�</b></font></td>
<td colspan="2"><font size="2">����ѽ~���͹�����Ҫ���� </font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>��˵��</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">��Ȼ�������ҽ�����������~��</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td colspan="2"><font size="2"><a href="../jh.asp">���þ��գ�����Ϊ��ɵ���ٺ٣�</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%conn.execute("Update �û� set ����=����+"&betcash&",Ӯ����=Ӯ����+1,ӮǮ=ӮǮ+"&betcash&",�Ĵ���=�Ĵ���+1,����=����-1 where ����='"&aqjh_name&"'")
%>
</font> <font size="2" color="#FFFFFF"><b><%=(betcash+nowcash)%></b><%=(nowmeili-1)%></font><font size="2" color="#FFFFFF">
<%end if%>
</font>
<p>&nbsp;</p>
</td>
</tr>
<tr bgcolor="#FFCCCC">
<td align="right" colspan="3"><a href="betindex.asp"><font size="2">�ĳ���ҳ</font></a></td>
</tr>
</table>
<%end if%>
<p>��Ȩ���С����齭����վ��</p>
</div>
</body>
</html>
<%rs.close
set rs=nothing
set username=nothing
set betcash=nothing
set nowcash=nothing
end if
end if
end if
end if
%>
