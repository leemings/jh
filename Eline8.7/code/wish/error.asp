<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="config.asp"-->
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
<div align="center">

<table border="1" cellspacing="1" width="550" bordercolor="<%=titlelightcolor%>">
  <tr>
    <td width="100%" height="18" bgcolor=<%=titledarkcolor%> align="center"><font color=<%=titlelightcolor%>>���ִ���</font></td>
  </tr>
  <tr>
    <td width="100%" height="92" align="center">����ԭ�� : <font color="#FF0000"><%=request.QueryString("msg")%></font></td>
  </tr>
</table>
</center>
</div>
<br><center><a href="javascript:history.go(-1);">������һҳ��</a>
<hr width=550 color=<%=titlelightcolor%> size=1></center>
<!--#include file="copyright.asp"-->
</body>
</html>