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
<title><%=Application("aqjh_chatroomname")%>--�ҵ�ͽ��</title>
<link rel="stylesheet" href="../css.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body text="#000000" vlink="#FF9900" topmargin="0" leftmargin="0" background="../../bg.gif">
<p align="center">
<%
id=request("id")
if instr(id,"��")<>0 then
		Response.Write "<script language=JavaScript>{alert('���ؾ��棬��Ҫ����');window.close();}</script>"
		Response.End
end if
my=request("my")
%>
<font color="#CC0000"><a href="javascript:this.location.reload()">ˢ��</a></font>
<table border="0" width="500" cellspacing="0" cellpadding="0" align="center">
<tr align="center">
<td width="500" height="26"><font
color="#FF6600"><b>�� �� ͽ ��</b></font></td>
</tr>
<tr align="center">
<td>
<table width="497" border="1" cellpadding="0" cellspacing="0" bordercolor="#FF9900" height="13">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� where ʦ��='" & aqjh_name &"'",conn
%>
<td width="54">
<div align="center">����</div>
</td>
<td width="27">
<div align="center">�Ա�</div>
</td>
<td width="71">
<div align="center">���</div>
</td>
<td width="83">
<div align="center">ע��ʱ��</div>
</td>
<td width="44">
<div align="center">��¼</div>
</td>
<td width="52">
<div align="center">�ܻ���</div>
</td>
<td width="49">
<div align="center">ʦ����Ǯ</div>
</td>
<td width="49">
<div align="center">ʦ��ָ��</div>
</td>
</tr>
<%
do while not rs.bof and not rs.eof
%>
<tr>
<td width="54" height="30">
<div align="center"><a href="../yamen/mt.asp?action=<%=rs("����")%>" target="_blank"><font color="#FFFFFF"><%=rs("����")%></font></a></div>
</td>
<td width="27" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("�Ա�")%></font></div>
</td>
<td width="71" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("���")%></font></div>
</td>
<td width="83" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("regtime")%><br><%=rs("lasttime")%></font></div>
</td>
<td width="44" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("times")%></font></div>
</td>
<td width="52" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("allvalue")%></font></div>
</td>
<td width="49" height="30">
<div align="center"><a href="mp6.asp" target="_blank"><font color="#FFFFFF">ʦ������</font></a></div>
</td>
<td width="49" height="30">
<div align="center"><a href="mp7.asp" target="_blank"><font color="#FFFFFF">ʦ������</font></a></div>
</td>
</tr>
<%
Application.Lock
Application("aqjh_bais_sf")=to1
Application("aqjh_bais_td")=aqjh_name
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
<td width="500" height="28"><FONT color=#0000ff>&copy; ��Ȩ���� 2004-2005 </FONT><A href="http://www.7758530.com/" target=_blank><FONT color=#0000ff>���齭����</FONT></A></td>
</tr></table></body></html>