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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>����Ա����</title>
<link rel="stylesheet" href="../css.css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0">
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="190"><img src="../images/gf1.jpg" width="57" height="323"><img src="../images/gf2.jpg" width="58" height="323"><img src="../images/gf3.jpg" width="65" height="323"></td>
<td width="588" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td>
<table width="97%" align="center" cellspacing="1" border="0"
cellpadding="0" bordercolor="#000000">
<tr>
<td colspan="5">
<div align="center"><font color="#FF6600">�� �� �� �� �� ��<br>
</font></div>
</td>
<tr bgcolor="#FFCC33">
<td align="center" width="19%" height="21"> �� �� </td>
<td align="center" width="11%" height="21"> �� �� </td>
<td align="center" width="14%" height="21" bgcolor="#FFCC00">
״ ̬ </td>
<td align="center" width="56%" height="21"> վ������ </td>
</tr>
<%
rs.open "SELECT * FROM �û� where ����='�ٸ�' order by grade DESC",conn
do while not rs.eof and not rs.bof
%>
<tr>
<td width="19%" nowrap>
<div align="center"> <%if aqjh_grade>=6 then%><%=rs("����")%> <%end if%> </div>
</td>
<td width="11%" nowrap>
<div align="center"> <%if aqjh_grade>=9 then%><%=rs("grade")%> <%end if%> </div>
</td>
<td width="14%" nowrap>
<div align="center"> <%if aqjh_grade>=6 then%><%=rs("״̬")%> <%end if%> </div>
</td>
<td width="56%" nowrap>
                  <div align="center">��̨����</div>
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
<tr>
<td height="16">
            <div align="center"></div>
</td>
</tr>
<tr>
<td>
<ul>
<li><font color="#000000" size="2">����Աֻ����ˢ�������ˣ���Υ���й��йػ������涨�Ķ���</font></li>
<li>�ٸ���ݷ�Ϊ6--10����������ݵĸߵ�ӵ�в�ͬ��Ȩ�������У�<br>
</li>
<li>����������Ϊ����Ȩ��<br>
</li>
<li>6��������ݿ��Դ���<br>
</li>
<li>7�����Ͽ���ץ�����Σ�10���Ӳ��ɽ�����(����)<br>
</li>
<li>8�����Ͽ��Է�Ǯ����ʱ���������ܣ����ң�<br>
</li>
<li>9�����Ͽ��Զ�����ķ��˽���ն��<br>
</li>
<li>10����������Ա<br>
</li>
<li>���ֹ���Ա������Ȩ���ģ�Ӧ����ʱ�ٱ�Ͷ�ߣ�վ�����ڲ�����ʵ���������Ĵ���</li>
</ul>
</td>
</tr>
</table>
</td>
</tr>
</table>
<div align="center"><br>
<font color="#000000" size="-1"><br>
����Ȩ:<%=Application("aqjh_user")%> ���к�:<%=Application("aqjh_sn")%>�� </font><br><br>�ڿ��ֽ�����վ��</div>
</body></html>