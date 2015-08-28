<!--#include file="../config.asp"--><%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if session("advtime")<>"" then
if session("advtime")>now()-0.0005 then  Response.Redirect "../error.asp?id=490"
end if
session("advtime")=now()
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT 会员 FROM 用户 where 姓名='" &username& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "你不是江湖中人或者连接超时"
conn.close
response.end
else
hy=rs("会员")
if yx8_mhjh_fellow=false then
%>
<script language=vbscript>
MsgBox "错误！会员功能尚未开放！"
location.href = "javascript:history.back()"
</script>
<%
else
if hy=false then
%>
<script language=vbscript>
MsgBox "错误！你不是会员，请回吧！"
location.href = "javascript:history.back()"
</script>
<%
else
randomize timer
r=int(rnd*50)
s=1
nl=0
sm=0
nu=int(rnd*18)+1
s=int(rnd*300)
sql="update 用户 set 精力=精力+'"& s &"'  where 姓名='" &username& "'"
conn.execute sql
conn.close
%>
<script language=vbscript>
MsgBox "<%=my%>"& "恭喜您!体力店的老板给了你一本武学秘笈，使你获得精力"&"<%=s%>"&"！！"
location.href = "tili.asp"
</script>
<%
end if
end if
end if
%>
