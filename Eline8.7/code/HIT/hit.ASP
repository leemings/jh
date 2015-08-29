<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 操作时间 from 用户 where 姓名='"&sjjh_name&"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<10 then
ss=10-sj
	Response.Write "<script Language=Javascript>alert('请等上"&ss&"分再来试飞吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>挑战极限</title>
<meta http-equiv="refresh" content="10">
<style></style>
</head>
<BODY oncontextmenu=self.event.returnValue=false BGCOLOR="#fefefe" text="#FFFFFF" background="../gif/Clouds.jpg">
<DIV ID="Layer1" STYLE="position:absolute; left:46px; top:139px; width:192px; height:165px; z-index:3">
<%
if Session("diaoyu")=true then
	Session("diaoyu")=now()
end if 
abc=DateDiff("s",Session("diaoyu"),now())
if abc>=85 then
%>
<div align="center">
<script language=vbscript>
location.href = "pao.asp"
</script>
<%
response.end
end if
if abc<=60 then
%>
<IMG SRC="diaoyu.gif" width="167" height="115"><br>
<font color="#0000FF">我是大侠我怕谁！</font> 
<%  
else
%>
<a href="hitok.ASP"><IMG SRC="diaoyuok.gif" border="0" width="98" height="77"></a>   
 &nbsp;<font color="#0000FF"><br> 
真主保佑！真主保佑！</font> 
<%  
end if
%>
</div>
</DIV>
<span id="liveclock" style="position: absolute; left: 328px; top: 275px; width: 76px; height: 23px">
<div align="center"><font color=red>飞行时间:<br>
<%=abc%>秒</font> </div>
</span>
</body>
</html>