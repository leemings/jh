<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from �û� where ����='"&sjjh_name&"'",conn
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body background="by.gif" bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<center>
<%
conn.execute "update �û� set ����=����-5,����=����+5,����=����+50,����=����+50 where ����='"&sjjh_name&"'"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "��æ��һ�죬�����������������5�㣬��������5����������50�㣬��������50�㣡"
location.href = "javascript:history.back()"
</script>
</body>
</html>