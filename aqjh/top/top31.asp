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
if aqjh_grade<>10 then Response.Redirect "manerr.asp?id=255"

%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ݶ���������TOP50</title>
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
<td width="100%"><font color="#ffffff" size="2">50λ�ݶ�������������(�ǹٸ�)</font></td>
</tr>
</table>
<%
dim sql
dim rs
dim filename
sql="select * from �û� where ����<>'�ٸ�' and DateDiff('m',cdate(lasttime),now())=0 order by �ݶ����� desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<p align='center'>û�п����еĶ��� </p>"
else
%>
  <table border="1" cellspacing="1" width="100%" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="0" bordercolor="#000000">
    <tr> 
      <td align="center" height="16" bgcolor="#336633" width="10%"><font color="#FFFFFF">����</font></td>
      <td align="center" height="16" bgcolor="#336633" width="7%"><font color="#FFFFFF">�Ա�</font></td>
      <td align="center" height="16" bgcolor="#336633" width="9%"><font color="#FFFFFF">����</font></td>
      <td align="center" height="16" bgcolor="#336633" width="19%"><font color="#FFFFFF">����</font></td>
      <td align="center" height="16" bgcolor="#336633" width="9%"><font color="#FFFFFF">�ݶ�����</font></td>
    </tr>
    <%do while not rs.eof%>
    <tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td align="center" width="10%"><%=rs("����")%> </td>
      <td align="center" width="7%"><font color="#0000FF"><%=rs("�Ա�")%></font></td>
      <td align="center" width="9%"><%=rs("����")%> </td>
      <td align="center" width="19%"><%=rs("����")%> </td>
      <td align="center" width="9%"><font color=ff0000><%=rs("�ݶ�����")%></font> 
      </td>
    </tr>
    <%
rs.movenext
filename=filename+1
if filename>50 then Exit Do
loop
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
  </table>
</div>
<div align="center"><font color=#999999 size=2><%Response.Write "���кţ�<font color=blue>" & Application("aqjh_sn") & "</font>����Ȩ����<font color=blue>" & Application("aqjh_user") & "</font><br><font color=999999>" & Application("aqjh_copyright") & "</font>"%></font>
</div>
</body>
</html>