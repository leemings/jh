<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if session("yx8_mhjh_usercorp")<>"官府" or session("yx8_mhjh_username")<>Application("yx8_mhjh_admin") then Response.Redirect "../error.asp?id=046"
result=request("result")
id=request("id")
bg=request("bg")
my=request("my")
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
if result="罚款一万" then
sql="update 用户 set 银两=银两-10000 where 姓名='"&bg&"'"
conn.Execute(sql)
sql="update 用户 set 银两=银两+10000 where 姓名='"&my&"'"
conn.Execute(sql)
end if
if result="罚款十万" then
sql="update 用户 set 银两=银两-100000 where 姓名='"&bg&"'"
conn.Execute(sql)
sql="update 用户 set 银两=银两+100000 where 姓名='"&my&"'"
conn.Execute(sql)
end if
if result="坐牢" then
sql="update 用户 set 状态='入狱',道德=道德-1000,最后登录时间=now()+1/144 where 姓名='"&bg&"'"
conn.execute sql
end if
conn.Close             
set conn=nothing   
%>
<!--#include file="data.asp"-->
<%
sql="update bbs set 结果='1' where id=" & id & ""
conn.execute sql
conn.Close             
set conn=nothing   
Response.Redirect "manage.asp"
%>
