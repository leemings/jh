<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="data1.asp"--><%
sql="select * from ���� where ����='" & aqjh_name & "' or ����='" & aqjh_name & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('����û�������أ�');location.href = 'fangwu.asp';}</script>"
Response.End
end if
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="select * from �û� where ����='"&aqjh_name&"'"
set rs=conn.execute(sql)
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
conn.execute "update �û� set ����=����+100 where ����='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�㿴��ħ���飬һ��С�ľͱ����һ���ǮŶ�����Ҹ���,�����������100����');location.href = 'moshu.asp';}</script>"
Response.End
%>
</body>
</html>
<html><script language="JavaScript">                                                                  </script></html>

