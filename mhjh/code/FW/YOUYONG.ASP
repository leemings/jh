<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
my=session("yx8_mhjh_username")
if my="" then response.redirect "../error.asp?id=016"
%>
<!--#include file="data1.asp"-->
<%
sql="select * from 房屋 where 户主='" & my & "' or 伴侣='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%>
<script language=vbscript>
   MsgBox "您还没有买房子呢。"
   location.href = "../11.htm"
</script>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
if day(rs("游泳"))<day(now()) or month(rs("游泳"))<month(now()) or year(rs("游泳"))<year(now()) or isnull(rs("游泳")) then
sql="update 房屋 set 游泳=now() where 户主='"& my &"' or 伴侣='" & my & "'"
conn.execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="data.asp"-->
<%
sql="update 用户 set 精力=精力+50000,体力=体力+1000000 where 姓名='"& my &"'"
conn.execute(sql)
conn.close
set conn=nothing
Response.Write "<head><title>消息</title><link rel=stylesheet href='../chatroom/css.css'><script lanjuage=javascript>setTimeout('history.back();',3000);</script></script></head><body  bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu='self.event.returnValue=false;' topmargin=100><p align=center>游一下泳，果然发现自己精神也好多了，也许该到家里睡一觉会更加开心吧<br>精力加50000，体力加1000000！<br>三秒钟后自动返回<br><a href='javascript:history.back();'>返回</a></p></body>"
else
conn.close
set conn=nothing
Response.Write "<head><title>消息</title><link rel=stylesheet href='../chatroom/css.css'><script lanjuage=javascript>setTimeout('history.back();',3000);</script></script></head><body  bgcolor='"&bgcolor&"' background='"&bgimage&"' oncontextmenu='self.event.returnValue=false;' topmargin=100><p align=center>一天只能游泳一次！<br>三秒钟后自动返回<br><a href='javascript:history.back();'>返回</a></p></body>"
end if
%>