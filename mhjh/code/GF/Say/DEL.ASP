<!--#include file="data.asp"-->
<%
if session("yx8_mhjh_usercorp")<>"�ٸ�" or session("yx8_mhjh_username")<>Application("yx8_mhjh_admin") then Response.Redirect "../../error.asp?id=046"
id=request("id")
sql="delete from bbs where id=" & id & ""
conn.execute sql
conn.close
set conn=nothing
%>
<p><a href="manage.asp">����������</a></p>
