<!--#include file="conn.asp"-->
<!--#include file="char.inc"-->
<%
if request.form("title")="" then
response.write "<p align='center'>错误：动画标题不能为空，请输入动画标题！</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>返回重新输入</a>"
response.end
end if
if request.form("class")="" then
response.write "<p align='center'>错误：动画类别没有选择！</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>返回重新输入</a>"
response.end
end if
if request.form("content")="" then
response.write "<p align='center'>错误：动画简介不能为空，请输入动画简介！</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>返回重新输入</a>"
response.end
end if
if request.form("url")="" then
response.write "<p align='center'>错误：动画地址不能为空，请输入动画地址！</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>返回重新输入</a>"
response.end
end if
if request.form("commend")="" then
response.write "<p align='center'>错误：动画推荐没有选择！</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>返回重新输入</a>"
response.end
end if
if request.form("KB")="" then
response.write "<p align='center'>错误：动画大小不能为空，请输入大小！</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>返回重新输入</a>"
response.end
end if
if request.form("img")="" then
response.write "<p align='center'>错误：动画图片不能为空，请输入图片地址！</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>返回重新输入</a>"
response.end
end if
%>
<%
dim title
dim content
dim classid
dim img
dim url
dim email
dim KB
dim commend
dim sql
dim rs

title=request.form("title")
content=UBBCode(request.form("content"))
ClassID=request.form("class")
img=request.form("img")
url=request.form("url")
email=request.form("email")
KB=request.form("KB")
commend=request.form("commend")

set rs=server.createobject("adodb.recordset")
sql="select * from flash" 

rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("content")=content
rs("ClassID")=ClassID
rs("img")=img
rs("url")=url
rs("email")=email
rs("KB")=KB
rs("commend")=commend
rs("time")=date()
rs.update

rs.close
set rs=nothing
conn.close
set conn=nothing

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="author" content="Thirdsnow;Email:qb169@sohu.com;QQ:34608582">
<title>添加成功</title>
<link rel="stylesheet" type="text/css" href="style.css">
<meta http-equiv=refresh content="1; url=add_flash.asp">
</head>
<body>
<table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center" height="22" class="font">恭喜!添加动画成功!1秒后自动返回到动画添加页面!</td>
  </tr>
</table>
</body>
</html>