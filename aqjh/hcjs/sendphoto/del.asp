<!--#include file="connpic.asp"-->
<% 
id=request("id")
set rs=server.createobject("ADODB.recordset") 
sql="select * from pic where id=" & id 
rs.open sql,conn,1,1 
If Session("aqjh_grade")>=10 or Session("aqjh_name")=rs("name") Then
	sql = "delete from pic where id=" & id
	conn.Execute sql
else
Response.write "<br><br><br><br><br>�Բ����㲻�ǹٸ�������Ա�����Ƭ�����ˣ���Ȩɾ��������Ƭ"
response.end
end if
rs.close 
set rs=nothing 
set conn=nothing 
%> 
<html><head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<title>��ʾ��Ϣ</title>
</head>
<link rel="stylesheet" href="../../css.CSS" type="text/css">
<body background=../../bg.gif>
<br><br><br><br><br><br><p class=p16 align=center>���Ѿ��ɹ���ɾ���˴�����Ƭ��<p>
<p class=p9 align=center><a href='photo.asp'>����ֱ�ӷ���</a></p><br><br><br><br><br><br><br><br><br>
</body></html>