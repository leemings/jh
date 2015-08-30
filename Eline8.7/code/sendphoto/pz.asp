<!--#include file="connpic.asp"-->
<% 
id=request("id")
set rs=server.createobject("ADODB.recordset") 
sql="select * from pic where id=" & id 
rs.open sql,conn,1,1 

If Session("sjjh_grade")>=10 Then
	sql = "update pic set 批准=1 where id=" & id
	conn.Execute sql
else

Response.Redirect "error.asp?id=000"

end if

rs.close 
set rs=nothing 
set conn=nothing 
%> 
<html><head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title>提示信息♀wWw.happyjh.com♀</title>
</head>
<style type="text/css">
.p9,body {font-size: 14}
.p16{font-size:16}
A {text-decoration: none; color: #000000;}
A:hover {text-decoration: underline; color: #FF0000;}
</style>
<body background=../bgcheetah.gif>
<br><br><br><br><br><br><p class=p16 align=center>你已经成功的批准了此张照片！<p>
<p class=p9 align=center><a href='photo.asp'>按此直接返回</a></p><br>
</body></html>