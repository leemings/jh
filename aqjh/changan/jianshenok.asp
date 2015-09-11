<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="data1.asp"--><%
sql="select * from 房屋 where 户主='" & aqjh_name & "' or 伴侣='" & aqjh_name & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('您还没有买房子呢！');location.href = 'fangwu.asp';}</script>"
Response.End
end if
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间 from 用户 where 姓名='"&aqjh_name&"'",conn
sj=DateDiff("s",rs("操作时间"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒,再操作！');location.href = 'fangniu.asp';}</script>"
	Response.End
end if
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
conn.execute "update 用户 set 体力=体力+10,操作时间=now() where 姓名='"&aqjh_name&"'"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你进行身体锻练，体力上升10点！');location.href = 'jianshen.asp';}</script>"
Response.End
%>
</body>
</html>