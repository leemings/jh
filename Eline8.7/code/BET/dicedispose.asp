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
dim betcash,nowcash
nowcash=0
betcash=0
betcash=clng(request.form("splosh"))
if betcash<10 then
%>
<script language=vbscript>
MsgBox "��ʲô��Ц����������ô�ٶĸ�ʲô��ѽ��"
location.href = "javascript:history.back()"
</script>
<%
elseif betcash>3000 then
%>
<script language=vbscript>
MsgBox "��ʲô��Ц��������ô��������Ʋ���!"
location.href = "javascript:history.back()"
</script>
<%
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="Select * From �û� Where ����='"&sjjh_name&"'"
set rs=conn.Execute(sql)

Randomize
Randomize
m1 = Int(6 * Rnd + 1)
m4 = Int(6 * Rnd + 1)

nowcash=rs("����")
if betcash>nowcash then
%>
<script language=vbscript>
MsgBox "��û�и�����Ϻɰ���������ô�����ӣ����ٺٺ٣���ȥ����ȥ�ɣ������"
location.href = "javascript:history.back()"
</script>
<%
else

Randomize
m2 = Int(6 * Rnd + 1)
m5 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
m6 = Int(6 * Rnd + 1)

nowmeili=rs("����")
if nowmeili<10 then
%>
<script language=vbscript>
MsgBox "��������Ŀǰ����������10�ˣ������Ͽ�����ȥѽ���Ժ������ɣ�"
location.href = "javascript:history.back()"
</script>
<%
else

m=Second(time())
m=right(m,1)
if m="0" or m="1" or m="6" or m="7" or m="8" then
do while m1+m2+m3<=m4+m5+m6

m4 = Int(6 * Rnd + 1)
m5 = Int(6 * Rnd + 1)
m6 = Int(6 * Rnd + 1)
loop
else
do while m1+m2+m3>=m4+m5+m6

m4 = Int(6 * Rnd + 1)
m5 = Int(6 * Rnd + 1)
m6 = Int(6 * Rnd + 1)
loop
end if
if m1+m2+m3<m4+m5+m6 then
mm=int(rnd*6+1)
if mm=1 or mm=2   then
do while m1+m2+m3<m4+m5+m6

m4 = Int(6 * Rnd + 1)
m5 = Int(6 * Rnd + 1)
m6 = Int(6 * Rnd + 1)
loop
end if
end if




%>

<html>
<head>

<title>�����ӡ�һ�������wWw.happyjh.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
BODY {font-size: 9pt; font-family: ����;
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
table {font-size: 9pt; font-family: ����}
input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}
.c {  font-family: ����; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>

</head>

<body text="#000000" background="../jhimg/bk_hc3w.gif" topmargin="0">
<font color="#FFFFFF"><br>
</font>
<div align="center">
<p><font size="2" class="c" color="#FFFFFF"><font size="3"><b><font color="#000000">��
�� - �� �� ��</font></b></font></font></p>


<table border="1" cellspacing="0" cellpadding="3" align="center" bordercolordark="#FFFFFF" bordercolorlight="#000000" bgcolor="#CCCCCC">
<tr>
<td bgcolor="#336633" colspan="3" height="23"><font size="2" class="c"><b>&nbsp;&nbsp;<font color="#FFFFFF">��
�� �� ��</font></b></font></td>
</tr>
<tr>
<td colspan="3" align="center"><font size="2">��-ׯ�����ӣ�<%=(m1+m2+m3)%> �� </font></td>
</tr>
<tr>
<td width="100" align="center"><font size="2"><img src="../jhimg/bbs/<%=m1%>.gif" width="31" height="32"></font></td>
<td width="100" align="center"><font size="2"><img src="../jhimg/bbs/<%=m2%>.gif"width="31" height="32"></font></td>
<td width="139" align="center"><font size="2"><img src="../jhimg/bbs/<%=m3%>.gif" width="32" height="32"></font></td>
</tr>
<tr>
<td colspan="3" align="center"><font size="2">��-������ӣ�<%=(m4+m5+m6)%> �� </font></td>
</tr>
<tr>
<td width="100" align="center" height="39"><font size="2"><img src="../jhimg/bbs/<%=m4%>.gif"></font></td>
<td width="100" align="center" height="39"><font size="2"><img src="../jhimg/bbs/<%=m5%>.gif"></font></td>
<td width="139" align="center" height="39"><font size="2"><img src="../jhimg/bbs/<%=m6%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"> <font size="2">
<%if (m1+m2+m3)>=(m4+m5+m6) then%>
</font>
<table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td height="49" width="14%">&nbsp;</td>
<td colspan="2" height="49" width="86%">
<p><font size="2"><img src="../jhimg/bbs/face13.gif" width="20" height="20" align="absmiddle">
�浹ù�����ˣ� <b><font color="#CC0000"><%=betcash%> </font>��</b></font></p>
</td>
</tr>
<tr>
<td align="right" width="14%"><font size="2"><b> ׯ�ң�</b></font></td>
<td colspan="2" width="86%"><font size="2"><b><font color="#000099"><img src="../jhimg/bbs/face18.gif" width="20" height="20" align="absmiddle"></font></b>
�Ǻ�~��Ҫ���� �ж�δ����Ŷ~��</font></td>
</tr>
<tr>
<td align="right" rowspan="2" width="14%"><font size="2"><b> ��Ҫ��</b></font></td>
<td colspan="2" width="86%"><font size="2"><b><font color="#CC0000"><img src="../jhimg/bbs/face8.gif" width="20" height="20" align="absmiddle"></font></b>
<a href="dice.asp">������Ϊʲô���������ޣ��ҾͲ�������ô��ù~��</a></font></td>
</tr>
<tr>
<td colspan="2" width="86%"><font size="2"><img src="../jhimg/bbs/face11.gif" width="20" height="20" align="absmiddle">
<a href="../jh.asp">�������ˣ��ϵ�ù�ɣ����ŵ����Ӿ�����~��</a></font></td>
</tr>
</table>
<font size="2"><%
conn.execute("Update �û� set ����=����-"&betcash&",�Ĵ���=�Ĵ���+1,����=����-1,ӮǮ=ӮǮ-"&betcash&" where ����='"&sjjh_name&"'")
%> </font>
<hr size="1" width="250">
<font size="2"> �����������ӣ�<font color="#CC0000"><b><font color="#CC0000"><%=(nowcash-betcash)%></font>
</b></font>�� ������<%=(nowmeili-1)%><%else%> </font>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right" height="50" width="14%">&nbsp;</td>
<td colspan="2" height="50" width="86%">
<p><font size="2"><img src="../jhimg/bbs/face14.gif" width="20" height="20" align="absmiddle">
����~��Ӯ�ˣ�<b><font color="#CC0000"><%=betcash%> </font>��</b></font></p>
</td>
</tr>
<tr>
<td align="right" width="14%"><font size="2"><b> ׯ�ң�</b></font></td>
<td colspan="2" width="86%"><font size="2"><b><font color="#000099"><img src="../jhimg/bbs/face17.gif" width="20" height="20" align="absmiddle"></font></b>
����ѽ~���͹�����Ҫ���� </font></td>
</tr>
<tr>
<td align="right" width="14%"><font size="2"><b>��Ҫ��</b></font></td>
<td colspan="2" width="86%"><font size="2"><img src="../jhimg/bbs/face10.gif" width="20" height="20" align="absmiddle">
<a href="dice.asp">��Ȼ�������ҽ�����������~��</a></font></td>
</tr>
<tr>
<td align="right" width="14%">&nbsp;</td>
<td colspan="2" width="86%"><font size="2"><img src="../jhimg/bbs/face2.gif" width="20" height="20" align="absmiddle">
<a href="../jh.asp">���þ��գ�����Ϊ��ɵ���ٺ٣�</a></font></td>
</tr>
</table>
<font size="2"><%conn.execute("Update �û� set ����=����+"&betcash&",Ӯ����=Ӯ����+1,ӮǮ=ӮǮ+"&betcash&",�Ĵ���=�Ĵ���+1,����=����-1 where ����='"&sjjh_name&"'")%>
</font>
<hr size="1" width="250">
<font size="2"> �����������ӣ�<b><font color="#CC0000"><b><%=(betcash+nowcash)%></b>
</font>�� </b>����: <%=(nowmeili-1)%><%end if%> </font>
<hr size="1" width="250">
</td>
</tr>
<tr>
<td align="right" colspan="3"><font size="2"><a href="betindex.asp">�ĳ���ҳ</a></font></td>
</tr>
</table>
<p><%=Application("sjjh_chatroomname")%>��Ȩ����</p>
</div>
</body>
</html>
<%rs.close
conn.close
set conn=nothing
set betcash=nothing
set nowcash=nothing
end if
end if
end if
%>
