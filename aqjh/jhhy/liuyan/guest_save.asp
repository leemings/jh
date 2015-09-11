<!--#include file="conn.asp"-->
<%
user=session("aqjh_name")
if user="" then
response.write"<script>alert('对不起，只有江湖玩家才可以在这里发布信息给站长，请您去登陆江湖之后再来！'); history.back();</script>"
response.end()
end if
Dim user,title,book
title=request("title")
book=request("book")
if title="" or book="" then
response.write"<script>alert('很抱歉，请您写好标题和内容再发布！');history.go(-1);</script>"
response.end()
end if
if len(book)>1000 then
response.write"<script>alert('对不起，发布内容不能大于 1000 个字符');history.go(-1);</script>"
response.end()
end if

set rs=server.createobject("adodb.recordset")
sql="select * from book"
rs.open sql,conn,1,3
rs.addnew
rs("user")=user
rs("title")=title
rs("book")=book
rs.update
rs.close
set rs=nothing
response.write"<script>alert('恭喜您，成功给站长发送了一条留言，只有您自己才能看到您发布的有关信息！请注意来查看站长给您的回复！'); location.replace('chakan.asp');</script>"
response.end()
%>