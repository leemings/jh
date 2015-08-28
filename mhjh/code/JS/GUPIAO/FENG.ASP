<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
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
	Response.Redirect "../../error.asp?id=046"
else%>
<!--#include file="jhb.asp"-->
<%
for i=1 to 20
sql="update 股票 set 状态='封' where sid="&i
conn.execute sql
next
end if
conn.close
Response.Redirect "index.asp"
%>

