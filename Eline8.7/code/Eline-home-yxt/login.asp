<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0 then Response.Redirect "manerr.asp?id=255"
pass=request.form("pass")
if pass="" then
session("sjjh_adminok")=false
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>��wWw.happyjh.com��</title>
<style></style>
<link rel="stylesheet" href="../chat/READONLY/Style.css">
</head>
<body bgcolor=#FFFFFF background="../bgcheetah.gif">
<center>
<font color="#000000"><b><font size="+2">���ֽ��� V8.7���ް�</font></b></font><br>
<br>
�����½��
<form action=login.asp method=post>
�����볬����������:<input type=password size=20 name=pass><input type=submit value='ȷ��'>
</form>
</center>
</body>
</html>
<%elseif pass=Application("sjjh_adminkey") then
session("sjjh_adminok")=true
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','�ɹ���¼��')"
conn.close
set conn=nothing
Response.Redirect "admin.asp"
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','��¼ʧ�ܡ�')"
conn.close
set conn=nothing
Response.Redirect "../exit.asp"
end if%>