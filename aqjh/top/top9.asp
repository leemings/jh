<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")

%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������TOP20</title>
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<style type="text/css">
<!--
body {  font-size: 9pt; color=000000}
td {  font-size: 9pt; color=000000}
A {text-decoration: none; color: #00ffff; font-family: "����"; font-size: 9pt}
a:hover      { color: ffffFF; font-family: ����; font-size: 9pt; text-decoration:
underline }
A:active {  font-family: "����"; font-style: normal;  color: ffffFF; text-decoration: underline; font-size: 9pt}
br {font-size:6pt; line-height:80%;}
-->
</style>
</head>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���� from �û� where ����='"&aqjh_name&"'",conn
mp=rs("����")
rs.close
const MaxPerPage=20
dim totalPut
dim CurrentPage
dim TotalPages
dim i,j
%>

<body text="#000000" background="../jhimg/bk_hc3w.gif">
<div align="center">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" bgcolor="#336633" height="16">
<tr>
<td width="100%"><font size="2" color="#ffffff">���������ɹ��׻���ǰ10λ����</font></td>
</tr>
</table>
<%
dim sql
dim rs
dim filename
sql="select * from �û� where ����='" & mp &"'order by ���ɻ��� desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<p align='center'>û�п����еĶ��� </p>"
else
%>
<table border="1" cellspacing="1" width="100%" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="0" bordercolor="#000000">
<tr>
<td align="center" height="16" bgcolor="#336633" width="9%"><font color="#FFFFFF">����</font></td>
<td align="center" height="16" bgcolor="#336633" width="6%"><font color="#FFFFFF">�Ա�</font></td>
<td align="center" height="16" bgcolor="#336633" width="16%"><font color="#FFFFFF">����</font></td>
<td align="center" height="16" bgcolor="#336633" width="22%"><font color="#FFFFFF">���</font></td>
<td align="center" height="16" bgcolor="#336633" width="11%"><font color="#FFFFFF">�ܻ���</font></td>
<td align="center" height="16" bgcolor="#336633" width="8%"><font color="#FFFFFF">�»���</font></td>
<td align="center" height="16" bgcolor="#336633" width="8%"><font color="#FFFFFF">��¼����</font></td>
<td align="center" height="16" bgcolor="#336633" width="13%"><font color="#FFFFFF">���ɻ���</font></td>
<td align="center" height="16" bgcolor="#336633" width="11%"><font color="#FFFFFF">����</font></td>
</tr>
<%do while not rs.eof%>
<tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';">
<td align="center" width="9%"><%=rs("����")%> </td>
<td align="center" width="6%"><font color="#0000FF"><%=rs("�Ա�")%></font></td>
<td align="center" width="16%"><%=rs("����")%> </td>
<td align="center" width="22%"><%=rs("���")%> </td>
<td align="center" width="11%"><%=rs("allvalue")%></td>
<td align="center" width="8%"><font color="#0000FF"><%=rs("mvalue")%></font></td>
<td align="center" width="8%"><font color=ff0000><%=rs("times")%></font></td>
<td align="center" width="13%"><font color="#0000FF"><b><%=rs("���ɻ���")%> </b>
<td align="center" width="22%"><%=rs("����")%> </td>
</font></td>
</tr>
<%
rs.movenext
filename=filename+1
if filename>9 then Exit Do
loop
end if %>
</table>
<br>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" bgcolor="#336633" height="16">
<tr>
<td width="100%"><font size="2" color="#ffffff">���������ɹ��׻����10λ����</font></td>
</tr>
</table>
<%
filename=0
rs.close
sql="select * from �û� where ����='" & mp &"'order by -���ɻ��� desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<p align='center'>û�п����еĶ��� </p>"
else
%>
<table border="1" cellspacing="1" width="100%" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="0" bordercolor="#000000">
<tr>
<td align="center" height="16" bgcolor="#336633" width="9%"><font color="#FFFFFF">����</font></td>
<td align="center" height="16" bgcolor="#336633" width="6%"><font color="#FFFFFF">�Ա�</font></td>
<td align="center" height="16" bgcolor="#336633" width="16%"><font color="#FFFFFF">����</font></td>
<td align="center" height="16" bgcolor="#336633" width="22%"><font color="#FFFFFF">���</font></td>
<td align="center" height="16" bgcolor="#336633" width="11%"><font color="#FFFFFF">�ܻ���</font></td>
<td align="center" height="16" bgcolor="#336633" width="8%"><font color="#FFFFFF">�»���</font></td>
<td align="center" height="16" bgcolor="#336633" width="8%"><font color="#FFFFFF">��¼����</font></td>
<td align="center" height="16" bgcolor="#336633" width="13%"><font color="#FFFFFF">���ɻ���</font></td>
</tr>
<%do while not rs.eof%>
<tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';">
<td align="center" width="9%"><%=rs("����")%> </td>
<td align="center" width="6%"><font color="#0000FF"><%=rs("�Ա�")%></font></td>
<td align="center" width="16%"><%=rs("����")%> </td>
<td align="center" width="22%"><%=rs("���")%> </td>
<td align="center" width="11%"><%=rs("allvalue")%> </td>
<td align="center" width="8%"><font color="#0000FF"><%=rs("mvalue")%></font></td>
<td align="center" width="8%"><font color=ff0000><%=rs("times")%></font></td>
<td align="center" width="13%"><font color="#0000FF"><b><%=rs("���ɻ���")%> </b>
</font></td>
</tr>
<%
rs.movenext
filename=filename+1
if filename>9 then Exit Do
loop
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</div>
<div align="center"><font color=#999999 size=2><br>
<%Response.Write "���кţ�<font color=blue>" & Application("aqjh_sn") & "</font>����Ȩ����<font color=blue>" & Application("aqjh_user") & "</font><br><font color=999999>" & Application("aqjh_copyright") & "</font>"%>
</font> </div>
</body>
</html>
