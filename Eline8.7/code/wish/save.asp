<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="config.asp"-->
<!--#include file="conn.asp"-->
<%
varname= request.form("name")
varsex= request.form("sex")
varemail= request.form("email")
varhomepage= request.form("homepage")
varwishtype= request.form("wishtype")
varaddress= request.form("address")
varinfo= request.form("info")

if varname="" then
    response.redirect "error.asp?msg=�����������Ϊ��"
end if

if varemail="" then
    response.redirect "error.asp?msg=��ĵ����ʼ���ַ����!!"
end if

if varaddress="" then
    response.redirect "error.asp?msg=��ľ�ס�ط����ܲ�ѡ!"
end if

if varinfo="" then
    response.redirect "error.asp?msg=���Ը����������д!"
end if

function HTMLEncode(fString)
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode = fString
end function

Set rs = Server.CreateObject("ADODB.Recordset")
sql ="Select * From wish"
rs.Open sql, conn,1,3
rs.addnew
rs("name")=varname
rs("sex")=varsex
rs("email")=varemail
rs("homepage")=varhomepage
rs("wishtype")=varwishtype
rs("address")=varaddress
rs("info")=htmlencode(varinfo)
rs("ip")=Request.ServerVariables("REMOTE_ADDR")
rs("date")=date()
rs.update
rs.close
%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<head>
<style type="text/css">
<!--
a {  text-decoration: none}  
a:hover {  text-decoration: underline} 
table {  font-size: 9pt}
body,table,p,td,input {  font-size: 9pt} 
-->
</style>
<title><%=title%></title>
</head>
<body bgcolor="<%=bgcolor%>" text="<%=textcolor%>" link="<%=linkcolor%>">
<form method="post" action="wish.asp">
<div align="center"><center>

<table border="1" cellspacing="1" width="550" bordercolor="<%=titlelightcolor%>">
  <tr>
    <td width="100%" height="18" bgcolor=<%=titledarkcolor%>><p align="center"><font color=<%=titlelightcolor%>>[<%=title%>] ����Ը��</font></td>
  </tr>
        <tr>
            <td colspan="2" height="23"><p align="center">&nbsp;<font
            color="#c0c0c0">������ĵ�¼���</font></p>
            </td>
        </tr>

    <tr align="center"><td><table align=center width=90% border=0><tr>
         <td>�ա�����: </td><td><%=varname%></td></tr><tr><td>
             �ա�����: </td><td><%=varsex%></td></tr><tr><td>
             �š�����: </td><td><%=varemail%></td></tr><tr><td>
             ��    ��: </td><td><%=varhomepage%></td></tr><tr><td>
             Ը�����: </td><td><%=varwishtype%></td></tr><tr><td>
             ��ס�ط�: </td><td><%=varaddress%></td></tr><tr><td>
             ����Ը��: </td><td><%=varinfo%></td></tr></table></td>
    </tr>
<tr>
        <tr>
            <td align="center" height="25"><input type="submit" style="border:1 solid <%=titlelightcolor%>;color:<%=titlelightcolor%>;background-color:white" value="������Ը��"></td>
        </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>