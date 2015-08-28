<!--#include file="data.asp"-->
<%
if session("yx8_mhjh_usercorp")<>"官府" or session("yx8_mhjh_username")<>Application("yx8_mhjh_admin") then Response.Redirect "../../error.asp?id=046"
id=request("id")
sql="update bbs set 结果='0' where id=" & id & ""
conn.execute sql
conn.close
set conn=nothing
%>
<p><a href="manage.asp">返回含冤告状</a></p>
