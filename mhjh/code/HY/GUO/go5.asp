<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if Session("usepro")<>"2" then
Session("usepro")=""
response.redirect "index.asp"
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT 攻击 FROM 用户 where 姓名='" & username & "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "你不是江湖中人"
conn.close
response.end
else
at=rs("攻击")
randomize timer
r=int(rnd*10)
s=int(rnd*50)
at=at*s
randomize timer
r=int(rnd*10)
s=int(rnd*50000)
at1=100*s
if at>at1 then
Session("usepro")="3"
msg="经过一番苦斗后，你以"&at&"点攻击赢了打不死的"&at1&"点攻击，<a href='go6.asp'>按这里闯下一关吧。</a>"
else
Session("usepro")=""
msg="经过一番苦斗后，你以"&at&"点攻击不敌打不死的"&at1&"点攻击，<a href='javascript:self.close()'>按这里回去吧吧。</a>"
end if
%>
<html>

<head>
<title>比武结果</title>
<LINK href="../../style.css" rel=stylesheet>
</head><body bgcolor="#000000" oncontextmenu=self.event.returnValue=false text="#FFFFFF" >
<div align="center"><b><%=msg%></b></div></body>
</html>
<%
end if
conn.close
%>
