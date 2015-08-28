<!--#include file="../config.asp"--><%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if session("advtime")<>"" then
if session("advtime")>now()-0.0003 then  Response.Redirect "../error.asp?id=490"
end if
session("advtime")=now()
username=session("yx8_mhjh_username")
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT 会员,体力,内力,攻击,防御 FROM 用户 where 姓名='" &username& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "你不是江湖中人或者连接超时"
conn.close
set conn=nothing
response.end
else
hy=rs("会员")
tl=rs("体力")
nl=rs("内力")
gj=rs("攻击")
fy=rs("防御")
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
cid=Request.QueryString("id")
select case cid
Case "1"
if tl<1000 then
Response.write "<script language='javascript'>alert ('你的体力低于1000');location.href='hxie.asp';</script>"
else
sql="update 用户 set 内力=内力+1000,体力=体力-1000  where 姓名='" &username& "'"
conn.execute sql
Response.write "<script language='javascript'>alert ('你用1000的体力换了1000的内力');location.href='hxie.asp';</script>"
end if
Case "2"
if nl<1000 then
Response.write "<script language='javascript'>alert ('你用的内力低于1000');location.href='hxie.asp';</script>"
else
sql="update 用户 set 内力=内力-1000,体力=体力+1000  where 姓名='" &username& "'"
conn.execute sql
Response.write "<script language='javascript'>alert ('你用1000的内力换了1000的体力');location.href='hxie.asp';</script>"
end if
Case "3"
if gj<1000 then
Response.write "<script language='javascript'>alert ('你的攻击低于1000');location.href='hxie.asp';</script>"
else
sql="update 用户 set 攻击=攻击-1000,防御=防御+1000  where 姓名='" &username& "'"
conn.execute sql
Response.write "<script language='javascript'>alert ('你用1000的攻击换了1000的防御');location.href='hxie.asp';</script>"
end if
Case "4"
if fy<1000 then
Response.write "<script language='javascript'>alert ('你的防御低于1000');location.href='hxie.asp';</script>"
else
sql="update 用户 set 防御=防御-1000,攻击=攻击+1000  where 姓名='" &username& "'"
conn.execute sql
Response.write "<script language='javascript'>alert ('你用1000的防御换了1000的攻击');location.href='hxie.asp';</script>"
end if
Case else
Response.write "<script language='javascript'>alert ('错误类型);location.href='hxie.asp';</script>"
end select
end if
end if
end if
conn.close
set conn=nothing
%>
