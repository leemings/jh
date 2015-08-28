<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then
Response.Write "<script language=javascript>"
Response.Write "parent.game1.location.href='/p';"
Response.Write "</script>"
Response.End
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
conn.Execute ("update 用户 set 银两=银两+10000 where 姓名='"&username&"'")
conn.Close
set conn=nothing
%><!--#include file="data.asp"--><%
cn.Execute ("update game set win=win+1 where username='"&username&"'")
cn.Close
set cn=nothing
Response.Write "<script language=javascript>"
Response.Write "parent.fb.location.href='button.asp';"
Response.Write "parent.t.location.href='start.asp';"
Response.Write "</script>"
Response.End
%>