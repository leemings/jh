<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="config.asp"-->
<!--#include file="conn.asp"-->
<%
if session("admin")="ture" then

   if not isempty(request("did")) then
      cdid=cint(request("did"))
   else
      response.write "没有您要删除的内容!"
   end if
set rs=server.createobject("adodb.recordset")
sql="delete from wish where id="&cdid
conn.execute(sql)
conn.close
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
<div align="center"><center>
<table border="1" cellspacing="1" width="550" bordercolor="<%=titlelightcolor%>">
  <tr>
    <td width="100%" height="18" bgcolor=<%=titledarkcolor%>><p align="center">你已经删除了第 <font color=<%=titlelightcolor%>><%=cdid%></font> 条记录</p></td>
  </tr>
  <tr>
    <td width="100%" height="92"><p align="center">请返回上一页继续删除用户,返回后请RELOAD该页!!</td>
  </tr>
</table><br>
<div align="center"><a href="javascript:history.go(-1);">【回上一页】</a></p>
<hr width=550 color=<%=titlelightcolor%> size=1></div>
<!--#include file="copyright.asp"-->
</body>
</html>
<%
else
    response.redirect "error.asp?msg=你没权利查看本页"
end if
%>