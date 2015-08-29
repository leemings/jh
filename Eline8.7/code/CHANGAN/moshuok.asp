<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
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
sql="select * from 用户 where 姓名='"&sjjh_name&"'"
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
conn.execute "update 用户 set 银两=银两+100 where 姓名='"&sjjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('你看过魔法书，一不小心就变出来一大堆钱哦，好幸福啊,你的银两增加100两！');location.href = 'moshu.asp';}</script>"
Response.End
%>
</body>
</html>
<html><script language="JavaScript">                                                                  </script></html>

