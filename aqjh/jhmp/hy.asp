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


%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>��Ա��ѯ����</title>
<style type="text/css">
<!--
p            { line-height: 20px; font-size: 9pt }
table        { font-size: 9pt }
a:link       { color: #FF9900; text-decoration: none }
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" vlink="#FF9900" topmargin="0"
leftmargin="0" background="../jhimg/bk_hc3w.gif">
<p align="center"> <font color="#CC0000" face="��Բ"><a href="javascript:this.location.reload()">ˢ��</a></font>
<br>
��л��Щ���Ѷ����ǽ����Ĵ���֧�֣������Ա��<a href="../chat/hy.asp" target="_blank">����</a><br>

<table border="0" width="500" cellspacing="0" cellpadding="0"
background="bg.gif" align="center">
<tr align="center">
<td background="top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">�� �� �� Ա</font></b></font></td>
</tr>
<tr align="center">
<td>
<table width="470" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" height="13">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� where ״̬='����' and ��Ա�ȼ�>0 order by ��Ա���� ",conn
%>
<td width="28" height="18">
<div align="center"><font color="#FFFFFF">ID</font></div>
</td>
<td width="47" height="18">
<div align="center"><font color="#FFFFFF">����</font></div>
</td>
<td width="25" height="18">
<div align="center"><font color="#FFFFFF">�Ա�</font></div>
</td>
<td width="63" height="18">
<div align="center"><font color="#FFFFFF">����</font></div>
</td>
<td width="86" height="18">
<div align="center"><font color="#FFFFFF">���</font></div>
</td>
<td width="40" height="18">
<div align="center"><font color="#FFFFFF">��Ա��</font></div>
</td>
<td width="75" height="18">
<div align="center"><font color="#FFFFFF">��Ա��������</font></div>
</td>
<td width="51" height="18">
<div align="center"><font color="#FFFFFF">�����ȼ�</font></div>
</td>
<td width="35" height="18">
<div align="center"><font color="#FFFFFF">��¼</font></div>
</td>
</tr>
<%
dengji1=0
dengji2=0
dengji3=0
dengji4=0
dengji5=0
dengji6=0
dengji7=0
dengji8=0
do while not rs.bof and not rs.eof
Select Case rs("��Ա�ȼ�")
Case 1
	dengji1=dengji1+1
case 2
	dengji2=dengji2+1
case 3
	dengji3=dengji3+1
case 4
	dengji4=dengji4+1
case 5
	dengji5=dengji5+1
case 6
	dengji6=dengji6+1
case 7
	dengji7=dengji7+1
case 8
	dengji8=dengji8+1

end select
%>
<tr>
<td width="28" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
</td>
<td width="47" height="30">
<div align="center"><a href="../yamen/mt.asp?action=<%=rs("����")%>"><font color="#FF9900"><%=rs("����")%></font></a></div>
</td>
<td width="25" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("�Ա�")%></font></div>
</td>
<td width="63" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("����")%></font></div>
</td>
<td width="86" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("���")%></font></div>
</td>
<td width="40" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("��Ա�ȼ�")%></font></div>
</td>
<td width="75" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("��Ա����")%></font></div>
</td>
<td width="51" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("�ȼ�")%></font></div>
</td>
<td width="35" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("times")%></font></div>
</td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</td>
</tr>
<tr align="center">
<td background="top3.gif" width="500" height="2"><font color="#FFFFFF"><%=Application("aqjh_chatroomname")%>��Ȩ����</font></td>
</tr>
</table>
<div align="center"><font color="#000000">һ���û���</font><b><font color="#0000FF"><%=dengji1%></font></b>
<font color="#000000">������Ա:</font><b><font color="#0000FF"><%=dengji2%></font></b><font color="#000000">��
������Ա��</font><font color="#0000FF"><b><%=dengji3%></b></font><font color="#000000">��
�ļ���Ա��</font><font color="#0000FF"><b><%=dengji4%></b></font><font color="#000000">��
�弶��Ա��</font><font color="#0000FF"><b><%=dengji5%></b></font><font color="#000000">��
������Ա��</font><font color="#0000FF"><b><%=dengji6%></b></font><font color="#000000">��
�߼���Ա��</font><font color="#0000FF"><b><%=dengji7%></b></font><font color="#000000">��
�˼���Ա��</font><font color="#0000FF"><b><%=dengji8%></b></font><font color="#000000">��<br>
��Ա������</font><b><font color="#0000FF"><%=(dengji1+dengji2+dengji3+dengji4+dengji5+dengji6+dengji7+dengji8)%></font></b><font color="#000000">��</font>
</div>
</body>
</html>