<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=210"
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>ҽԺ</title>
<link rel="stylesheet" type="text/css" href="style.css">
<style type="text/css">td           { font-family: ����; font-size: 9pt }
body         { font-family: ����; font-size: 9pt }
select       { font-family: ����; font-size: 9pt }
a            { color: #FFC106; font-family: ����; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: ����; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF">
<br>
<table width="700" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="318">&nbsp;</td>
<td width="460"><font color="#FF3300"><b>������Ա����</b></font> </td>
</tr>
</table>
<br>
<table width="700" border="0" cellpadding="0" cellspacing="0">
<tr>
<td bgcolor="#000000">&nbsp;</td>
</tr>
<tr bgcolor="#990000">
<td><img src="../images/lf2.gif"><img src="../images/lf3.gif" height="80"></td>
</tr>
<tr>
<td bgcolor="#000000">&nbsp;</td>
</tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="700">
<tr>
<td>
<div align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
align="center" >
<tr>
<td colspan="2" valign="top" width="100%" height="110">
<div align="center"> <br>
<table border="1" cellspacing="1" cellpadding="0" width="100%"
bordercolor="#000000" align="center">
<tr>
<td width="15%" nowrap>
<div align="center"> ����ײ����� </div>
<td width="15%" nowrap>
<div align="center"> ���� </div>
</td>
<td width="11%">
<div align="center"> ��� </div>
</td>
<td width="29%">
<div align="center"> �� �� </div>
</td>
<td width="21%">&nbsp;</td>
<td width="9%" nowrap>��</td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� where ״̬='�����'",conn
do while not rs.bof and not rs.eof
%>
<tr>
<td width="15%" height="31"><%=rs("����")%>
<td width="15%" height="31"><%=rs("����")%></td>
<td width="11%" nowrap height="31"><%=rs("���")%></td>
<%if rs("״̬")="�����" then %>
<td width="29%" height="31">
<div align="center"> <a href="yiyuan4.asp?name=<%=rs("����")%>"><font color="#008000">��Ǯ1000��������</font></a>  
<%if aqjh_grade=10 then%> 
 <a href="zhanshou.asp?name=<%=rs("����")%>"><font color="#FF0000"></font></a>  
<%end if%> <%if aqjh_grade>=6 then%>   
  |<a href="yiyuan5.asp?id=<%=rs("id")%>"><font color="#FF0000">�ٸ�����</font></a> 
<%end if%> </div>   
</td>
<% else
if rs("��¼")>now() then%>
<td width="21%" height="31">
	<div align="center"> <font color="#FF0033"></font>
	</div>
	</td>
<%else%>
	<td width="9%" height="31">
	<div align="center"> <font color="#FF0033"></font>
	</div>
	</td>
	<% conn.execute("update �û� set ״̬='����',��¼=now() where ����='"&rs("����")&"'")
end if%>
<%end if%>
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
<p>�Ժ�С��ע�����壡 ������ </p> 
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
<div align="center"></div></body></html>