<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������TOP10</title>
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
const MaxPerPage=10
dim totalPut
dim CurrentPage
dim TotalPages
dim i,j
%>
<body text="#000000" background="../bg.gif">
<div align="center">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" height="22" bgcolor="#336633">
<tr>
<td width="100%"><font color=#ffffff>��������TOP10</font></td>
</tr>
</table>
<%
dim sql
dim rs
dim filename
sql="select * from �û� where ���='����' and ����<>'�ٸ�' order by �书+����+���� desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<p align='center'>û�п����еĶ��� </p>"
else
%>
<table border="1" cellspacing="1" width="100%" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="0" bordercolor="#000000">
<tr>
<td align="center" height="17" bgcolor="#336633" width="12%"><font size="2" color="#FFFFFF">��
��</font></td>
<td align="center" height="17" bgcolor="#336633" width="12%"><font size="2" color="#FFFFFF">��
��</font></td>
<td align="center" height="17" bgcolor="#336633" width="12%"><font size="2" color="#FFFFFF">��
��</font></td>
<td align="center" height="17" bgcolor="#336633" width="12%"><font size="2" color="#FFFFFF">�书</font></td>
<td align="center" height="17" bgcolor="#336633" width="9%"><font size="2" color="#FFFFFF">����</font></td>
<td align="center" height="17" bgcolor="#336633" width="10%"><font size="2" color="#FFFFFF">����</font></td>
<td align="center" height="17" bgcolor="#336633" width="14%"><font color="#FFFFFF">�ϼ�</font></td>
<td align="center" height="17" bgcolor="#336633" width="19%"><font size="2" color="#FFFFFF">��
��</font></td>
</tr>
<%do while not rs.eof%>
<tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';">
<td align="center" width="12%"><font color=ff0000><%=rs("����")%></font> </td>
<td align="center" width="12%"><font color=#6666FF><%=rs("����")%></font>
</td>
<td align="center" width="12%"><%=rs("���")%> </td>
<td align="center" width="12%"><font color=ff0000><%=rs("�书")%></font> </td>
<td align="center" width="9%"><font color=ff0000><%=rs("����")%></font></td>
<td align="center" width="10%"><font color=ff0000><%=rs("����")%></font></td>
<td align="center" width="14%"><font color="#0000FF"><%=rs("�书")+rs("����")+rs("����")%></font></td>
<td align="center" width="19%"><%=rs("����")%> </td>
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
<div align="center"><br><br>�ڰ��齭����վ��
</div>
</body>
</html>