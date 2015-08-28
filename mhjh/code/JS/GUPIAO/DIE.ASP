<%@ LANGUAGE=VBScript%>
<!--#include file="jhb.asp"-->
<%Response.Expires=0
if session("yx8_mhjh_inchat")<>"1" then %>
<script language="vbscript">
alert "请进入聊天室再进入股票！"
window.close()
</script>
<%Response.End
end if
if username=""  then
  Response.Redirect "../../error.asp?id=016"
elseif username<>Application("yx8_mhjh_admin") then
	Response.Redirect "../../error.asp?id=439"
else%>
<%
sid=Request.Form ("sid")
sql="update 股票 set 当前价格=当前价格*0.9 where sid="&sid
conn.execute sql
conn.close
Response.Redirect "exchange.asp"
set conn=nothing
end if
%>
