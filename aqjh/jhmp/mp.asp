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
<title><%=Application("aqjh_chatroomname")%>���ɹ���</title>
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
<p align="center">
<%
id=request("id")
if instr(id,"��")<>0 then
		Response.Write "<script language=JavaScript>{alert('���ؾ��棬��Ҫ����');window.close();}</script>"
		Response.End
end if
my=request("my")
%>
<font color="#CC0000" face="��Բ"><a href="javascript:this.location.reload()">ˢ��</a></font>
<table border="0" width="500" cellspacing="0" cellpadding="0"
background="bg.gif" align="center">
<tr align="center">
<td background="top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">�� �� �� ��</font></b></font></td>
</tr>
<tr align="center">
<td>
<table width="470" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" height="13">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� where ����='"&id&"' order by mvalue desc",conn
%>
<td width="31">
<div align="center"><font color="#FFFFFF">ID</font></div>
</td>
<td width="54">
<div align="center"><font color="#FFFFFF">����</font></div>
</td>
<td width="27">
<div align="center"><font color="#FFFFFF">�Ա�</font></div>
</td>
<td width="74">
<div align="center"><font color="#FFFFFF">���</font></div>
</td>
<td width="75">
<div align="center"><font color="#FFFFFF">�»���</font></div>
</td>
<td width="66">
<div align="center"><font color="#FFFFFF">ע��ip</font></div>
</td>
<td width="65">
<div align="center"><font color="#FFFFFF">ע��ʱ��</font></div>
</td>
<td width="32">
<div align="center"><font color="#FFFFFF">��¼</font></div>
</td>
<td width="73">
<div align="center"><font color="#FFFFFF">���Ų���</font></div>
</td>
<td width="28">
<div align="center"><font color="#FFFFFF">ɾ��</font></div>
</td>
</tr>
<%
do while not rs.bof and not rs.eof
%>
<tr>
<td width="31" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
</td>
<td width="54" height="30">
<div align="center"><a href="../yamen/mt.asp?action=<%=rs("����")%>"><font color="#FFFFFF"><%=rs("����")%></font></a></div>
</td>
<td width="27" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("�Ա�")%></font></div>
</td>
<td width="74" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("���")%></font></div>
</td>
<td width="75" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("mvalue")%></font></div>
</td>
<td width="66" height="30">
<div align="center"><font color="#FFFFFF">
<%if aqjh_grade>=9 then%>
	<%=rs("regip")%>
<%else%>
	��Ȩ�鿴
<%end if%>
</font></div>
</td>
<td width="65" height="30">
<div align="center"><font color="#FFFFFF">
<%if aqjh_grade>=9 then%>
	<%=rs("regtime")%><br>
	<%=rs("lasttime")%>
<%else%>
	<%=rs("regtime")%>
<%end if%>
</font></div>
</td>
<td width="32" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("times")%></font></div>
</td>
<td width="73" height="30">
<div align="center"><a href="mp3.asp?you=<%=rs("����")%>&amp;id=<%=rs("����")%>"><font color="#FFFFFF">���ʦ��</font></a></div>
</td>
<td width="28" height="30">
<div align="center"><a href="mp4.asp?you=<%=rs("����")%>&id=<%=rs("����")%>"><font color="#FFFFFF">
<%if aqjh_grade>=9 then%>
	ɾ��
<%else%>
	��Ȩ
<%end if%>
</font></a></div>
</td>
</tr>
<%
rs.movenext
filename=filename+1 
if filename>20 then Exit Do 
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
<td background="top3.gif" width="500" height="28"><font color="#FFFFFF">��Ȩ���С����ֽ�����վ��</font></td>
</tr></table></body></html>