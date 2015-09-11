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
my=aqjh_name%>
<!--#include file="data1.asp"-->
<%sql="select * from 房屋 where 户主='" & my & "' or 伴侣='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>
MsgBox "您还没有买房子呢。"
location.href = "fangwu.asp"
</script>
<%else
if day(rs("游泳"))<day(now()) and month(rs("游泳"))<month(now()) and year(rs("游泳"))<year(now()) and isnull(rs("游泳")) then
sql="update 房屋 set 游泳=now() where 户主='"& my &"' or 伴侣='" & my & "'"
set rs=conn.execute(sql)
set rs=nothing	
	conn.close
	set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
sql="update 用户 set 魅力=魅力+100,体力=体力+10000 where 姓名='"&aqjh_name&"'"
set rs=conn.execute(sql)
Response.Redirect "../ok.asp?id=602"
else
%><script language=vbscript>
MsgBox "您今天游过泳了啊。"
location.href = "xiaowu7.asp"
</script>
<%end if
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>