<!--#include file="conn.asp"-->
<!--#include file="char.inc"-->
<%
if request.form("title")="" then
response.write "<p align='center'>���󣺶������ⲻ��Ϊ�գ������붯�����⣡</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>������������</a>"
response.end
end if
if request.form("class")="" then
response.write "<p align='center'>���󣺶������û��ѡ��</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>������������</a>"
response.end
end if
if request.form("content")="" then
response.write "<p align='center'>���󣺶�����鲻��Ϊ�գ������붯����飡</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>������������</a>"
response.end
end if
if request.form("url")="" then
response.write "<p align='center'>���󣺶�����ַ����Ϊ�գ������붯����ַ��</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>������������</a>"
response.end
end if
if request.form("commend")="" then
response.write "<p align='center'>���󣺶����Ƽ�û��ѡ��</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>������������</a>"
response.end
end if
if request.form("KB")="" then
response.write "<p align='center'>���󣺶�����С����Ϊ�գ��������С��</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>������������</a>"
response.end
end if
if request.form("img")="" then
response.write "<p align='center'>���󣺶���ͼƬ����Ϊ�գ�������ͼƬ��ַ��</p>"
response.write "<p align='center'><a href='javascript:history.back(1)'>������������</a>"
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
<title>��ӳɹ�</title>
<link rel="stylesheet" type="text/css" href="style.css">
<meta http-equiv=refresh content="1; url=add_flash.asp">
</head>
<body>
<table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center" height="22" class="font">��ϲ!��Ӷ����ɹ�!1����Զ����ص��������ҳ��!</td>
  </tr>
</table>
</body>
</html>