<!--#include file="conn.asp"-->
<%
dim rs
dim sql
username=replace(request("username"),"'"," ")
password=replace(Request("password"),"'"," ")
if password="" or username="" then
response.write("�û��������벻��Ϊ�գ�")
response.end
end if

sql="select * from user where username='"&username&"'and password='"&password&"'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if not(rs.bof and rs.eof) then
'if request.cookies("username")<>"" or request.cookies("password")<>"" then
'response.cookies("okerer")="yesok"
'else
     
     sql="update user set lastlogin=now() where id="&rs("ID")
     conn.Execute(sql)
     response.cookies("username")=""
     response.cookies("password")=""
     response.cookies("username")=rs("username")
     response.cookies("password")=rs("password")
     response.cookies("okerer")="yesok"
' end if
     
else
  response.write"<script>alert('�û�����������󣬲����ܣ�');history.back();</Script>"
  response.end
end if
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  response.redirect Request.ServerVariables("HTTP_REFERER")
 %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>E�߽�����վ___DJ���___wWw.51eline.com</title>
</head>

<body>

</body>

</html>