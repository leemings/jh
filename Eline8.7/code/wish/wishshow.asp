<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="config.asp"-->
<!--#include file="conn.asp"-->
<%
if request.QueryString("id")="" then
     response.redirect "error.asp?msg=û�д���Ը!"
else

Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from wish where id="&request.QueryString("id")
rs.open sql,conn,3,3
rs("counter")=rs("counter")+1
rs.update
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
<title><%=rs("name")%> ����Ը�ơ�wWw.51eline.com��</title>
</head>
<body bgcolor="<%=bgcolor%>" text="<%=textcolor%>" link="<%=linkcolor%>">
<form method="post" action="wish.asp">
  <div align="center">
  <center>    
  <div align="center">
  <center>
  <table border="1" cellspacing="1" width="400" bordercolor="<%=titlelightcolor%>">
        <tr>
            
          <td align="center" bgcolor="<%=titledarkcolor%>"><font color="#FFFFFF">��<%=title%>����Ը��</font></td>
        </tr>
        <tr>
            <td align="center" style=padding:20>
			<div align="center">
			<center>
			<table border="0" cellpadding="0" cellspacing="0" width="310">
                <tr>
                    <td align="right" valign="bottom"><img src="img/<%=rs("wishtype")%>/top.gif" width="300" height="50"></td>
                    <td>��</td>
                </tr>
                <tr>
                    <td align="center" bgcolor="#FFE8BC" style="padding:5" height="100"><blockquote><p align="left"><font color="#800040"><%=rs("info")%></font></p></blockquote>
                      <p align="right"><font color="#800040"><%=rs("name")%> ���ɣ�<%=rs("homepage")%>&nbsp;<%=rs("sex")%>&nbsp;<%=rs("address")%></font> 
                        <a href="mailto:<%=rs("email")%>"><img src="img/mail.gif" align="absmiddle" border="0" width="15" height="15"></a> 
                      </p>
                    </td>
                    <td valign="bottom" rowspan="2" width="10"><img src="img/edge.gif" width="10" height="100%"></td>
                </tr>
                <tr>
                    <td align="right" valign="top"><img src="img/base.gif" width="300" height="25"></td>
                </tr>
            </table>
            </center></div></td>
        </tr>
        <tr>
            
          <td align="center" bgcolor="<%=titledarkcolor%>"><font color="#FFFFFF">��Ը����ף����λ��Ը��...</font></td>
        </tr>
        <tr>
            <td align="center" height="25"><input
            type="submit" style="border:1 solid <%=titlelightcolor%>;color:<%=titlelightcolor%>;background-color:white" value="������Ը��"></td>
        </tr>
    </table>
    </center></div>
</form>
<!--#include file="copyright.asp"-->
</body>
</html>
<%end if%>