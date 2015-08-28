
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if Session("usepro")<>"4" then
Session("usepro")=""
response.redirect "index.asp"
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT 内力 FROM 用户 where 姓名='" & username & "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "你不是江湖中人"
conn.close
response.end
else
at=rs("内力")
Randomize
at=int(at*rnd)
Randomize
at1=int(100000000*rnd)
if at>at1 then
Session("usepro")="5"
msg="经过一番苦斗后，你以"&at&"点内力赢了林平之的"&at1&"点内力，<a href='go10.asp'>按这里用内力震开宝藏的门吧。</a>"
else
Session("usepro")=""
msg="经过一番苦斗后，你以"&at&"点内力不敌林平之的"&at1&"点内力"
Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
end if
%>
<html>

<head>
<title>比武结果</title><LINK href="../style.css" rel=stylesheet></head>
<body background="../images/back2.gif" oncontextmenu=self.event.returnValue=false >
<div align="center"><%=msg%></div>
</html>
<%
end if
%>

