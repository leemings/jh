<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_name="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="Select  * from �û� where ����='"&aqjh_name&"'"
set rs=conn.Execute(sql)
%>
<html>
<head>
<title>��������</title>
<LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<table width="428" border="0" cellspacing="15" cellpadding="0" background="../images/b2.gif" height="323">
<tr>
<td width="415" height="310">
<table border="1"
width="400" cellspacing="1" cellpadding="1" background="../bg.gif">
<tr>
<td colspan="5" height="33">
<div align="center"><%=rs("����")%>�����Ľ������� </div>
</td>
</tr>
<tr>
<td width="86">�� ��</td>
<td colspan="2"><%=rs("����")%></td>
<td colspan="2"> �������ţ� <font color="#0000FF" size="2">
<%if rs("�ȼ�")<5  then response.write("����է��")
if rs("�ȼ�")>=5 and rs("�ȼ�")<10  then response.write("��������")
if rs("�ȼ�")>=10 and rs("�ȼ�")<15  then response.write("С�гɾ�")
if rs("�ȼ�")>=15 and rs("�ȼ�")<20  then response.write("�����Ժ�")
if rs("�ȼ�")>=20 and rs("�ȼ�")<35  then response.write("�д�����")
if rs("�ȼ�")>=35 and rs("�ȼ�")<45  then response.write("һ������")
if rs("�ȼ�")>=45 and rs("�ȼ�")<55  then response.write("��������")
if rs("�ȼ�")>=55 and rs("�ȼ�")<65  then response.write("��������")
if rs("�ȼ�")>=65 and rs("�ȼ�")<75  then response.write("������ʤ")
if rs("�ȼ�")>=100 then response.write("����")
%>
</font></td>
</tr>
<tr>
<td width="86" height="23">�� ��</td>
<td height="23" width="84"><%=rs("�Ա�")%> </td>
<td height="23" width="35"><font size="2">״ ̬</font></td>
<td height="23" width="57"><%=rs("״̬")%></td>
<td height="23" width="110">��Ա�ѣ�<font color="#0000FF" size="2"><%=rs("��Ա��")%>Ԫ</font></td>
</tr>
<tr>
<td width="86">Email��ַ:</td>
<td width="84"><%=rs("����")%> </td>
<td width="35">OIcq:</td>
<td colspan="2"><%=rs("oicq")%></td>
</tr>
</table>
<div align="left"></div>
<table border="1"
width="400" cellspacing="1" cellpadding="1" background="../bgcheetah.gif">
<tr>
<td colspan="5" height="20">
<div align="center"> <font size="2">�� �� �� ��</font> </div>
</td>
</tr>
<tr>
<td width="66" height="25">�� ��</td>
<td width="110" height="25"><%=rs("����")%> ��</td>
<td width="63" height="25">�����ˣ�</td>
<td height="25" colspan="2"><%=rs("������")%></td>
</tr>
<tr>
<td width="66" height="20">�� �</td>
<td width="110" height="24"><%=rs("���")%> ��</td>
<td width="63" height="24">�� �£�</td>
<td height="24" colspan="2"><%=rs("����")%></td>
</tr>
<tr>
<td width="66" height="20">�� ����</td>
<td width="110" height="24"><%=rs("�书")%></td>
<td width="63" height="24">�� ����</td>
<td height="24" colspan="2"><%=rs("����")%></td>
</tr>
<tr>
<td width="66" height="20">�� ����</td>
<td width="110"><%=rs("����")%></td>
<td width="63">�� ����</td>
<td colspan="2"><%=rs("����")%></td>
</tr>
<tr>
<td width="66" height="20">�� ����</td>
<td width="110"><%=rs("����")%></td>
<td width="63">�� �ɣ�</td>
<td colspan="2"><%=rs("����")%> </td>
</tr>
<tr>
<td width="66" height="20">�� ����</td>
<td width="110"><%=rs("����")%></td>
<td width="63">�� �ݣ�</td>
<td colspan="2"><%=rs("����")%></td>
</tr>
<tr>
<td width="66" height="20">�� ż��</td>
<td width="110"><%=rs("��ż")%></td>
<td width="63">Ӯ / �ģ�</td>
<td colspan="2"><%=rs("Ӯ����")%> / <%=rs("�Ĵ���")%></td>
</tr>
<tr>
<td width="66" height="20">�ĳ��ȼ���</td>
<td width="110">
<%if rs("ӮǮ")<100  then response.write("�ĳ�����")
if rs("ӮǮ")>=2000 and rs("ӮǮ")<10000  then response.write("�ĳ�����")
if rs("ӮǮ")>=10000 and rs("ӮǮ")<50000  then response.write("������ͽ")
if rs("ӮǮ")>=50000 and rs("ӮǮ")<100000  then response.write("�ĳ�����")
if rs("ӮǮ")>=100000 and rs("ӮǮ")<300000  then response.write("ְҵ����")
if rs("ӮǮ")>=300000 and rs("ӮǮ")<800000  then response.write("����")
if rs("ӮǮ")>=800000 and rs("ӮǮ")<1500000  then response.write("����")
if rs("ӮǮ")>=1000000 then response.write("����")
%>
</td>
<td width="63">ս������</td>
<td width="61"><%=rs("�ȼ�")%>��</td>
<td width="84">������:<%=rs("grade")%>��</td>
</tr>
</table>
</td>
</tr>
</table>
</body>

</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>