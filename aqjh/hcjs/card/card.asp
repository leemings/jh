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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ��Ա�� from �û� where ����='"&aqjh_name&"'",conn
hyfei=rs("��Ա��")
rs.close
rs.open "select * from b where b='��Ƭ' and h>0 order by h"
%>
<!--#include file="cardset.asp"-->
<html>
<head>
<title>��Ա��Ƭ</title>
<link rel="stylesheet" href="../../chat/READONLY/Style.css">
</head>
<body bgcolor=#FFFFFF background="../../bg.gif" text="#000000">
<BR>
<p align="center"><font size="6" face="����">������Ƭ��</font></p>
<div align="center">Ϊ�˷�ֹ������Ƭ���ң�ͬһ��Ū�Ϳ�Ƭֻ�ܹ���<font color=red><%=max%></font>��<br>
<font color=blue><%=aqjh_name%></font>�����л�Ա��:<font color=red><b><%=hyfei%></b></font>Ԫ��
</div>
<table border=1 align=center width=600 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#3399FF">
<tr bgcolor="#CC9900">
<td colspan="4" height="32">
<div align="center"><font color="#000000">��Ա��Ƭ</font></div>
</td>
</tr>
<tr>
    <td width="12%" height="11"> 
      <div align="center">��Ƭ����</div>
</td>
    <td width="67%" height="11"> 
      <div align="center">�� ��</div>
</td>
    <td width="7%" height="11"> 
      <div align="center">���</div>
</td>
    <td height="11" width="14%"> 
      <div align="center">�� ��</div>
</td>
</tr>
</table>
<table border=1 align=center width=600 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<%
do while not rs.bof and not rs.eof
%>
<tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';">
    <td width='12%' >
<div align="center"><%=rs("a")%>
</div>
</td>
    <td width='67%' ><%=rs("c")%></td>
    <td width='8%' >
<div align="center"><%=rs("h")%>Ԫ</div>
</td>
    <td  align="center" width="13%"><a href="buycard.asp?id=<%=rs("id")%>">����</a></td>
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
</body>
</html>