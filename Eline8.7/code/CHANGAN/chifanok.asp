<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("sjjh_useronlinename"))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="data1.asp"--><%
sql="select * from 房屋 where 户主='" & sjjh_name & "' or 伴侣='" & sjjh_name & "'"
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
conn.open Application("sjjh_usermdb")
rs.open "select 操作时间 from 用户 where 姓名='"&sjjh_name&"'",conn
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
conn.execute "update 用户 set 内力=内力+20,操作时间=now() where 姓名='"&sjjh_name&"'"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你细嚼慢咽的吃完了汉堡，精神好多了，内力提升20点！');location.href = 'chifan.asp';}</script>"
Response.End
%>
</body>
</html>