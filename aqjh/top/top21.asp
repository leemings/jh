<%@ LANGUAGE=VBScript codepage ="936" %>
<%
response.expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"%><html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<%
const MaxPerPage=10
dim totalPut
dim CurrentPage
dim TotalPages
dim i,j
%>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" leftmargin="0">
<table border="1" width="125" cellspacing="0" cellpadding="0" bgcolor="#336633" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td width="100%" height="28">
<div align="center"><font color="#ffffff" size="2">��ʮ�����</font></div>
</td>
</tr>
</table>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dim sql
dim rs
dim filename
rs.open "select top 20 * from �û� where ����<0 and ״̬<>'��' and ����<>'�ٸ�' order by ����",conn,1,1
if rs.eof and rs.bof then
	response.write "<p align='center'>û�п����еĶ��� </p>"
else
%>
<table border="1" cellspacing="0" width="125" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="4" align="center">
<tr bgcolor="#336633">
<td align="center" width="62"><font color="#FFFFFF">��������</font></td>
<td align="center" width="36"><font color="#FFFFFF">�� ��</font></td>
</tr>
<%do while not rs.eof%>
<tr>
<td align="center" bgcolor="#F7F7F7" width="62"><a href='../yamen/mt.asp?action=<%=rs("����")%>' target='_blank'><%=rs("����")%></a></td>
<td align="center" bgcolor="#F7F7F7" width="36"><%=rs("����")%> </td>
</tr>
<%
rs.movenext
filename=filename+1
if filename>19 then Exit Do
loop
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</body>
</html>