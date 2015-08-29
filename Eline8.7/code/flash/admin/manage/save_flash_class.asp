<!--#include file="session.asp"-->
<!--#include file="conn.asp"-->

<%
action=request.form("action")
%>
<% if action="edit" then %>
<%
if request.form("ClassName")="" then
response.write "<p align='center'>错误：类别不能为空，请输入类别名称！</p>"
response.write "<a href='javascript:history.back(1)'>返回重新输入</a>"
response.end
end if

dim ClassName
dim classid

ClassID=request.form("classid")
ClassName=htmlencode(request.form("ClassName"))

set rs=server.createobject("adodb.recordset")
sql="select * from FlashClass where Classid="&classid 

rs.open sql,conn,1,3

rs("ClassName")=ClassName
rs.update

rs.close
set rs=nothing

%>
<% else %>
<%
if request.form("ClassName")="" then
response.write "<p align='center'>错误：类别不能为空，请输入类别名称！</p>"
response.write "<a href='javascript:history.back(1)'>返回重新输入</a>"
response.end
end if

ClassName=htmlencode(request.form("ClassName"))

set rs=server.createobject("adodb.recordset")
sql="select * from FlashClass where (classid is null)" 

rs.open sql,conn,1,3
rs.addnew
rs("ClassName")=ClassName
rs.update

rs.close
set rs=nothing
conn.close
set conn=nothing

%>
<% end if %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="author" content="Thirdsnow;Email:qb169@sohu.com;QQ:34608582">
<title>添加成功</title>
<link rel="stylesheet" href="../style.css" type="text/css">
</head>
<body>
<table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center" height="22" class="font">恭喜!添加类别成功!1秒后自动返回到类别添加页面!</td>
  </tr>
</table>
<meta http-equiv=refresh content="1; url=add_Flash_class.asp">
</body>
</html>